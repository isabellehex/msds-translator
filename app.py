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

# Инициализация сессий
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

def normalize_section_2_1(text_block: str) -> str:
    """
    Очищает и пересобирает хаотичный текст из пункта 2.1,
    распознавая как новый формат (H-коды), так и старый (R-фразы).
    """
    lines = text_block.split('\n')
    clean_lines = []
    
    garbage_patterns = [
        r'www\.\S+', 
        r'\d{2}/\d{2}/\d{4}', 
        r'\b\d+\s*/\s*\d+\b', 
        r'(?i)safety\s+data\s+sheet',
        r'(?i)formaldehyde\s+solution\s+ar/acs' # при необходимости расширяй список мусора
    ]
    
    for line in lines:
        line_strip = line.strip()
        if not line_strip:
            continue
        if any(re.search(p, line_strip) for p in garbage_patterns):
            continue
        clean_lines.append(line_strip)
        
    single_flow = " ".join(clean_lines)
    result_rows = []
    
    has_h_codes = bool(re.search(r'\bH\d{3}\b', single_flow))
    has_r_phrases = bool(re.search(r'\bR\d{2,}', single_flow))
    
    if has_h_codes:
        pattern = r'(.*?\bH\d{3}\b.*?Category\s+\d+\w*(?:\s*,\s*[^A-Z]*)?)'
        alt_pattern = r'(.*?Category\s+\d+\w*.*?\bH\d{3}\b)'
        matches = re.findall(f'{pattern}|{alt_pattern}', single_flow)
        for match in matches:
            row = match[0] if match[0] else match[1]
            if row:
                result_rows.append(row.strip())
    elif has_r_phrases:
        pattern = r'(.*?\bR\d{2,}(?:\/\d{2,})*\b)'
        matches = re.findall(pattern, single_flow)
        for match in matches:
            if match:
                result_rows.append(match.strip())
        full_text_msg = re.search(r'(Full text of.*)', single_flow)
        if full_text_msg:
            result_rows.append(full_text_msg.group(1).strip())

    if not result_rows:
        return "\n".join(clean_lines)
        
    return "\n".join(result_rows)

def reset_state():
    st.session_state.raw_text = ""
    st.session_state.original_raw_text = ""
    st.session_state.loaded_file_id = None
    st.session_state.translated_text = ""

# ==========================================
# АЛГОРИТМ УМНОГО РАЗДЕЛЕНИЯ И ОБРЕЗКИ PDF
# ==========================================
def format_pdf_text_by_sections(text):
    """
    1. Находит SECTION 1 / РАЗДЕЛ 1 и отсекает весь текст до него.
    2. Разбирает оставшийся текст из PDF по строкам.
    3. Находит разделы (SECTION) и подразделы (1.1, 1.1.1) и отделяет их пустой строкой.
    """
    if not text:
        return ""
        
    lines = text.split("\n")
    
    # --- Шаг А: Поиск Первого Раздела для отсечения «шапки» ---
    first_section_pattern = re.compile(r'^\s*(SECTION|РАЗДЕЛ)\s+1\b', re.IGNORECASE)
    start_index = 0
    
    for idx, line in enumerate(lines):
        if first_section_pattern.match(line.strip()):
            start_index = idx
            break
            
    # Отсекаем всё, что было до найденного индекса первой секции
    meaningful_lines = lines[start_index:]
    
    # --- Шаг Б: Форматирование и расстановка пустых строк ---
    formatted_lines = []
    
    # Регулярка для всех главных разделов (1-16)
    section_pattern = re.compile(r'^\s*(SECTION|РАЗДЕЛ)\s+(\d+)\b', re.IGNORECASE)
    
    # Регулярка для подразделов (ищет X.X.X или X.X в начале строки)
    subsection_pattern = re.compile(r'^\s*(\d+\.\d+\.\d+|\d+\.\d+)\b')
    
    # Черный список параметров, чтобы не путать подразделы с физическими величинами
    measurement_units = ["%", "mg", "g/", "ppm", "cst", "°", "linc", "v/v", "w/w", "min", "max", "hpa", "kpa"]
    
    for line in meaningful_lines:
        cleaned_line = line.strip()
        is_structure_element = False
        
        # Проверяем на главный раздел
        if section_pattern.match(cleaned_line):
            is_structure_element = True
            
        # Проверяем на подраздел 1.1 / 1.1.1
        elif subsection_pattern.match(cleaned_line):
            lower_line = cleaned_line.lower()
            if not any(unit in lower_line for unit in measurement_units):
                is_structure_element = True
        
        # Если это элемент структуры (раздел или подраздел)
        if is_structure_element:
            # Если предыдущая строка не пустая — принудительно вставляем пустую строку перед элементом
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
                
                # Применяем новую разметку с отсечением шапки и разделением пустой строкой
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
        st.error(f"Ошибка сохранения словаря на GitHub: {e}")
        return False

# ==========================================
# РЕНДЕРИНГ ВКЛАДОК ПРИЛОЖЕНИЯ
# ==========================================
tab_main, tab_glossary = st.tabs(["Переводчик (v3 PDF)", "Редактор глоссария"])

# --- ВКЛАДКА 1: ОСНОВНОЙ ПАЙПЛАЙН ---
with tab_main:
    st.title("📝 Переводчик MSDS — Версия v3")
    st.subheader("Фокус: Структурированный парсинг PDF ➡️ Интерактивное редактирование")
    
    st.header("Шаг 1: Загрузка исходного MSDS")
    uploaded_file = st.file_uploader("Перетащите сюда исходный файл MSDS в формате PDF", type=["pdf", "txt"])
    
    if uploaded_file is not None:
        current_file_id = f"{uploaded_file.name}_{uploaded_file.size}"
        
        if st.session_state.loaded_file_id != current_file_id:
            with st.spinner("Извлекаем текст, отрезаем лишнее и структурируем по разделам..."):
                extracted = parse_uploaded_file(uploaded_file)
                
                # === ВОТ СЮДА МЫ ВСТАВЛЯЕМ НАШУ ФУНКЦИЮ НОРМАЛИЗАЦИИ ===
                # Пропускаем извлеченный текст через очистку перед сохранением в стейт
                extracted = normalize_section_2_1(extracted)
                # ====================================================
                
                st.session_state.raw_text = extracted
                st.session_state.original_raw_text = extracted 
                st.session_state.loaded_file_id = current_file_id
                st.session_state.file_name_output = os.path.splitext(uploaded_file.name)[0]
                st.rerun()

    st.header("Шаг 2: Редактор исходного текста (чистый контент с SECTION 1)")
    
    if st.session_state.raw_text:
        st.caption("Всё, что шло до SECTION 1 / РАЗДЕЛ 1 автоматически отрезано. Элементы структуры разделены пустой строкой.")
        
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
        st.info("Пожалуйста, загрузите PDF-файл на Шаге 1, чтобы увидеть чистый структурированный текст.")
        
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
