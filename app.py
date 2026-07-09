import streamlit as st
import pdfplumber
import docx
from docx import Document
import io
import os
import re
import json
import requests
import xml.etree.ElementTree as ET
import zipfile
from github import Github

# ==========================================
# ИНИЦИАЛИЗАЦИЯ И НАСТРОЙКИ СТРАНИЦЫ
# ==========================================
st.set_page_config(page_title="MSDS Translator v3", layout="wide", page_icon="📝")

# Чтение конфигураций из secrets
yandex_secrets = st.secrets.get("yandex", {})
folder_id = yandex_secrets.get("folder_id", "")
api_key = yandex_secrets.get("api_key", "")

github_secrets = st.secrets.get("github", {})
GITHUB_TOKEN = github_secrets.get("token", "")
GITHUB_REPO = github_secrets.get("repo", "")
TARGET_BRANCH = github_secrets.get("branch", "global-glossary-02")

# Инициализация сессий для стабильной работы редактора текста
if "raw_text" not in st.session_state:
    st.session_state.raw_text = ""
if "original_raw_text" not in st.session_state:
    st.session_state.original_raw_text = ""
if "loaded_file_id" not in st.session_state:
    st.session_state.loaded_file_id = None
if "file_name_output" not in st.session_state:
    st.session_state.file_name_output = "translated_document"
if "translated_text" not in st.session_state:
    st.session_state.translated_text = ""
if "current_glossary_cache" not in st.session_state:
    st.session_state.current_glossary_cache = {}

def reset_state():
    st.session_state.raw_text = ""
    st.session_state.original_raw_text = ""
    st.session_state.loaded_file_id = None
    st.session_state.translated_text = ""

# ==========================================
# ШАГ 1: ГЛУБОКИЙ И ЧИСТЫЙ ПАРСИНГ ИЗВЛЕЧЕНИЯ
# ==========================================
def extract_raw_xml_text_from_zip(docx_bytes):
    """
    Глубокий XML-парсер DOCX. Вытаскивает текст из абсолютно всех тегов <w:t>,
    включая плавающие надписи, фигуры и текстовые поля.
    """
    try:
        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
            xml_content = z.read('word/document.xml')
            root = ET.fromstring(xml_content)
            
            paragraphs = []
            # Ищем абсолютно все абзацы в структуре документа
            for p in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):
                p_text = []
                for t in p.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
                    if t.text:
                        p_text.append(t.text)
                if p_text:
                    paragraphs.append("".join(p_text))
            
            # Базовая дедупликация идущих подряд одинаковых строк, вызванных рендерингом фигур
            seen = set()
            unique_paragraphs = []
            for para in paragraphs:
                cleaned = para.strip()
                if cleaned and cleaned not in seen:
                    seen.add(cleaned)
                    unique_paragraphs.append(para)
            
            return "\n".join(unique_paragraphs)
    except Exception as e:
        st.error(f"Ошибка глубокого разбора DOCX через ZIP: {e}")
        try:
            doc = Document(io.BytesIO(docx_bytes))
            return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        except Exception as e2:
            st.error(f"Ошибка стандартного разбора DOCX: {e2}")
            return ""

def parse_uploaded_file(uploaded_file):
    """
    Определяет тип файла и извлекает из него «сырой» текст в чистом виде.
    """
    file_bytes = uploaded_file.read()
    if uploaded_file.name.endswith('.pdf'):
        try:
            with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                full_text = []
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        full_text.append(text)
                return "\n".join(full_text)
        except Exception as e:
            st.error(f"Ошибка парсинга PDF: {e}")
            return ""
    elif uploaded_file.name.endswith('.docx') or uploaded_file.name.endswith('.doc'):
        return extract_raw_xml_text_from_zip(file_bytes)
    else:
        try:
            return file_bytes.decode('utf-8', errors='ignore')
        except Exception as e:
            st.error(f"Ошибка парсинга текстового файла: {e}")
            return ""

# ==========================================
# ИНТЕГРАЦИЯ GITHUB ДЛЯ ГЛОССАРИЯ
# ==========================================
def load_glossary_from_github():
    """Загружает файл глоссария в формате JSON напрямую из репозитория GitHub."""
    if not GITHUB_TOKEN or not GITHUB_REPO:
        return {}
    try:
        g = Github(GITHUB_TOKEN)
        repo = g.get_repo(GITHUB_REPO)
        contents = repo.get_contents("glossary.json", ref=TARGET_BRANCH)
        glossary_data = json.loads(contents.decoded_content.decode('utf-8'))
        # Превращаем в плоский словарь и сортируем cache
        sorted_glossary = dict(sorted(glossary_data.items(), key=lambda item: item[0].lower()))
        st.session_state.current_glossary_cache = sorted_glossary
        return sorted_glossary
    except Exception as e:
        st.warning(f"Не удалось загрузить глоссарий из GitHub (используем пустой): {e}")
        return {}

def save_glossary_to_github(updated_glossary):
    """Сортирует и перезаписывает глоссарий в репозитории на GitHub."""
    if not GITHUB_TOKEN or not GITHUB_REPO:
        st.error("Данные GitHub аутентификации не заполнены в secrets.")
        return False
    try:
        g = Github(GITHUB_TOKEN)
        repo = g.get_repo(GITHUB_REPO)
        
        # Сортируем словарь по алфавиту перед сохранением
        sorted_glossary = dict(sorted(updated_glossary.items(), key=lambda item: item[0].lower()))
        
        contents = repo.get_contents("glossary.json", ref=TARGET_BRANCH)
        updated_json_str = json.dumps(sorted_glossary, ensure_ascii=False, indent=2)
        
        repo.update_file(
            contents.path,
            "style(glossary): авто-сортировка и обновление словаря v3",
            updated_json_str,
            contents.sha,
            branch=TARGET_BRANCH
        )
        st.session_state.current_glossary_cache = sorted_glossary
        return True
    except Exception as e:
        st.error(f"Ошибка сохранения словаря на GitHub: {e}")
        return False

# ==========================================
# РЕНДЕРИНГ ВКЛАДОК ПРИЛОЖЕНИЯ
# ==========================================
tab_main, tab_glossary = st.tabs(["Переводчик (v3)", "Редактор глоссария"])

# --- ВКЛАДКА 1: ОСНОВНОЙ ПАЙПЛАЙН ---
with tab_main:
    st.title("📝 Переводчик MSDS — Версия v3")
    st.subheader("Пайплайн: Чистая загрузка исходника ➡️ Интерактивное редактирование")
    
    # --- ШАГ 1: Загрузка файла ---
    st.header("Шаг 1: Загрузка исходного MSDS")
    uploaded_file = st.file_uploader("Перетащите сюда исходный файл MSDS (PDF, DOCX, TXT)", type=["pdf", "docx", "doc", "txt"])
    
    if uploaded_file is not None:
        current_file_id = f"{uploaded_file.name}_{uploaded_file.size}"
        
        # Запускаем парсинг только один раз при смене файла
        if st.session_state.loaded_file_id != current_file_id:
            with st.spinner("Извлекаем чистый текст из документа..."):
                extracted = parse_uploaded_file(uploaded_file)
                st.session_state.raw_text = extracted
                st.session_state.original_raw_text = extracted # Храним резервную копию оригинала
                st.session_state.loaded_file_id = current_file_id
                st.session_state.file_name_output = os.path.splitext(uploaded_file.name)[0]
                st.rerun()

    # --- ШАГ 2: Примитивный интерактивный редактор ---
    st.header("Шаг 2: Редактор исходного текста")
    
    if st.session_state.raw_text:
        st.caption("Вы можете свободно изменять, дополнять, удалять артефакты, выстраивать нумерацию или группировать строки прямо в поле ниже:")
        
        # Текстовое поле отображает текущее состояние из сессии и обновляет его при вводе
        user_edited_text = st.text_area(
            label="Редактируемый текст MSDS перед отправкой на перевод:",
            value=st.session_state.raw_text,
            height=500,
            key="msds_main_editor"
        )
        st.session_state.raw_text = user_edited_text
        
        # Кнопки управления состоянием текста
        col_btn1, col_btn2, _ = st.columns([1, 1, 4])
        with col_btn1:
            if st.button("Сбросить изменения к оригиналу"):
                st.session_state.raw_text = st.session_state.original_raw_text
                st.rerun()
        with col_btn2:
            if st.button("Очистить редактор полностью"):
                reset_state()
                st.rerun()
    else:
        st.info("Пожалуйста, загрузите файл на Шаге 1, чтобы открыть интерактивный редактор текста.")
        
    # --- ШАГ 3: Заглушка для будущего идеального перевода нейросетью ---
    st.divider()
    st.header("Шаг 3: Генерация идеального перевода v3 (В разработке)")
    st.caption("На следующем этапе здесь будет настроен промпт для YandexGPT, который переведет отредактированный текст единым красивым блоком.")
    st.button("Запустить нейросетевой перевод v3", disabled=True)

# --- ВКЛАДКА 2: РЕДАКТОР ГЛОССАРИЯ ---
with tab_glossary:
    st.title("🗂️ Синхронизация глоссария с GitHub")
    
    if st.button("🔄 Загрузить/Обновить словарь из GitHub"):
        with st.spinner("Получаем актуальный глоссарий..."):
            load_glossary_from_github()
            st.success("Глоссарий успешно загружен и отсортирован по алфавиту!")

    # Работа со словарем, если он загружен в кэш сессии
    if st.session_state.current_glossary_cache:
        glossary_dict = st.session_state.current_glossary_cache
        
        # Преобразуем словарь в список строк для удобного отображения в data_editor
        data_list = [{"Исходный термин (ENG)": k, "Перевод (RUS)": v} for k, v in glossary_dict.items()]
        
        st.subheader("Редактирование терминов словаря")
        st.caption("Вы можете изменять значения, добавлять новые строки внизу таблицы или удалять ненужные (выделив строку и нажав Delete).")
        
        # Интерактивная таблица для редактирования
        edited_data_list = st.data_editor(
            data_list,
            num_rows="dynamic",
            use_container_width=True,
            key="glossary_table_editor"
        )
        
        # Кнопка сохранения изменений на GitHub
        if st.button("💾 Сохранить изменения на GitHub"):
            with st.spinner("Сортируем данные и отправляем коммит в репозиторий..."):
                # Собираем данные обратно в формат словаря
                new_glossary = {}
                for row in edited_data_list:
                    key = row.get("Исходный термин (ENG)")
                    val = row.get("Перевод (RUS)")
                    if key and str(key).strip():
                        new_glossary[str(key).strip()] = str(val).strip() if val else ""
                
                if save_glossary_to_github(new_glossary):
                    st.success("Изменения успешно сохранены на GitHub и выстроены по алфавиту!")
                    st.rerun()
    else:
        st.info("Нажмите кнопку выше, чтобы загрузить текущую базу данных глоссария с GitHub.")
