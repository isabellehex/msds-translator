import streamlit as st
import pdfplumber
import io
import os
import re
import json
import requests
from github import Github

# ==========================================
# ИНИЦИАЛИЗАЦИЯ И НАСТРОЙКИ СТРАНИЦЫ
# ==========================================
st.set_page_config(page_title="MSDS Translator v3 (PDF Focus)", layout="wide", page_icon="📝")

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
# АЛГОРИТМ УМНОГО РАЗДЕЛЕНИЯ PDF ПО СТРУКТУРЕ
# ==========================================
def format_pdf_text_by_sections(text):
    """
    Разбирает сплошной текст из PDF по строкам.
    Находит разделы (SECTION/РАЗДЕЛ) и подразделы (1.1, 1.1.1) и отделяет их пустой строкой.
    """
    if not text:
        return ""
        
    lines = text.split("\n")
    formatted_lines = []
    
    # Объединенное регулярное выражение:
    # Группа 1 (Разделы): Начинается со слова SECTION или РАЗДЕЛ + номер
    # Группа 2 (Подразделы): Начинается с цифр формата X.X или X.X.X (после которых идет буква или конец строки, чтобы не брать физические величины вроде 1.5 %)
    pattern = re.compile(
        r'^\s*(?:(SECTION|РАЗДЕЛ)\s+(\d+)\b|(?:(\d+\.\d+\.\d+)|\b(\d+\.\d+))\s*(?=[A-Za-zА-Яа-я]|$))', 
        re.IGNORECASE
    )
    
    for line in lines:
        cleaned_line = line.strip()
        
        # Если строка является разделом или подразделом 1.1 / 1.1.1
        if pattern.match(cleaned_line):
            # Если предыдущая строка в списке не пустая — принудительно делаем отступ в одну пустую строку
            if formatted_lines and formatted_lines[-1] != "":
                formatted_lines.append("")
            formatted_lines.append(line)
        else:
            formatted_lines.append(line)
            
    # Собираем обратно в текст
    result_text = "\n".join(formatted_lines)
    
    # Схлопываем случайные тройные и более переносы строк до одной пустой строки (\n\n)
    result_text = re.sub(r'\n{3,}', '\n\n', result_text)
    
    return result_text.strip()

def parse_uploaded_file(uploaded_file):
    """
    Извлекает текст из файлов. Сфокусирован на глубоком и чистом разборе PDF.
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
                
                raw_extracted = "\n".join(full_text)
                
                # Применяем умную разметку структуры (Разделы + Подразделы)
                structured_text = format_pdf_text_by_sections(raw_extracted)
                return structured_text
                
        except Exception as e:
            st.error(f"Ошибка парсинга PDF: {e}")
            return ""
            
    elif uploaded_file.name.endswith('.docx') or uploaded_file.name.endswith('.doc'):
        st.warning("⚠️ Ветка парсинга DOCX отключена в v3. Фокус смещен на PDF.")
        return "Ветка алгоритма для DOCX отключена. Пожалуйста, загрузите PDF файл."
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
        sorted_glossary = dict(sorted(glossary_data.items(), key=lambda item: item[0].lower()))
        st.session_state.current_glossary_cache = sorted_glossary
        return sorted_glossary
    except Exception as e:
        st.warning(f"Не удалось загрузить глоссарий из GitHub: {e}")
        return {}

def save_glossary_to_github(updated_glossary):
    """Сортирует и перезаписывает глоссарий в репозитории на GitHub."""
    if not GITHUB_TOKEN or not GITHUB_REPO:
        st.error("Данные GitHub аутентификации не заполнены в secrets.")
        return False
    try:
        g = Github(GITHUB_TOKEN)
        repo = g.get_repo(GITHUB_REPO)
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
        st.error(f"Ошибка保存словаря на GitHub: {e}")
        return False

# ==========================================
# РЕНДЕРИНГ ВКЛАДОК ПРИЛОЖЕНИЯ
# ==========================================
tab_main, tab_glossary = st.tabs(["Переводчик (v3 PDF)", "Редактор глоссария"])

# --- ВКЛАДКА 1: ОСНОВНОЙ ПАЙПЛАЙН ---
with tab_main:
    st.title("📝 Переводчик MSDS — Версия v3")
    st.subheader("Фокус: Структурированный парсинг PDF ➡️ Интерактивное редактирование")
    
    # --- ШАГ 1: Загрузка файла ---
    st.header("Шаг 1: Загрузка исходного MSDS")
    uploaded_file = st.file_uploader("Перетащите сюда исходный файл MSDS в формате PDF", type=["pdf", "txt"])
    
    if uploaded_file is not None:
        current_file_id = f"{uploaded_file.name}_{uploaded_file.size}"
        
        # Запускаем парсинг только один раз при смене файла
        if st.session_state.loaded_file_id != current_file_id:
            with st.spinner("Извлекаем и структурируем текст из PDF по подразделам..."):
                extracted = parse_uploaded_file(uploaded_file)
                st.session_state.raw_text = extracted
                st.session_state.original_raw_text = extracted 
                st.session_state.loaded_file_id = current_file_id
                st.session_state.file_name_output = os.path.splitext(uploaded_file.name)[0]
                st.rerun()

    # --- ШАГ 2: Примитивный интерактивный редактор ---
    st.header("Шаг 2: Редактор исходного текста (с разметкой структуры)")
    
    if st.session_state.raw_text:
        st.caption("Автоматически выделены пустой строкой: заголовки SECTION, а также подразделы типа 1.1 и 1.1.1.")
        
        user_edited_text = st.text_area(
            label="Текст MSDS, готовый к проверке:",
            value=st.session_state.raw_text,
            height=550,
            key="msds_main_editor"
        )
        st.session_state.raw_text = user_edited_text
        
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
        st.info("Пожалуйста, загрузите PDF-файл на Шаге 1, чтобы увидеть разбитый по подразделам текст.")
        
    # --- ШАГ 3: Заглушка для будущего идеального перевода нейросетью ---
    st.divider()
    st.header("Шаг 3: Генерация идеального перевода v3 (В разработке)")
    st.caption("На следующем этапе мы передадим этот структурированный текст в YandexGPT.")
    st.button("Запустить нейросетевой перевод v3", disabled=True)

# --- ВКЛАДКА 2: РЕДАКТОР ГЛОССАРИЯ ---
with tab_glossary:
    st.title("🗂️ Синхронизация глоссария с GitHub")
    
    if st.button("🔄 Загрузить/Обновить словарь из GitHub"):
        with st.spinner("Получаем актуальный глоссарий..."):
            load_glossary_from_github()
            st.success("Глоссарий успешно загружен и отсортирован по алфавиту!")

    if st.session_state.current_glossary_cache:
        glossary_dict = st.session_state.current_glossary_cache
        data_list = [{"Исходный термин (ENG)": k, "Перевод (RUS)": v} for k, v in glossary_dict.items()]
        
        st.subheader("Редактирование терминов словаря")
        
        edited_data_list = st.data_editor(
            data_list,
            num_rows="dynamic",
            use_container_width=True,
            key="glossary_table_editor"
        )
        
        if st.button("💾 Сохранить изменения на GitHub"):
            with st.spinner("Сортируем данные и отправляем коммит в репозиторий..."):
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
