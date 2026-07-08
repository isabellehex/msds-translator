import streamlit as st
import openai
import pdfplumber
import io
import re
import zipfile
import json
import xml.etree.ElementTree as ET
from github import Github
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- Настройка страницы ---
st.set_page_config(
    page_title="MSDS Yandex AI Studio Pro",
    page_icon="🧪",
    layout="wide"
)

# --- Автоматическое чтение конфигураций из Streamlit Secrets ---
yandex_secrets = st.secrets.get("yandex", {})
FOLDER_ID = yandex_secrets.get("folder_id", "")
API_KEY = yandex_secrets.get("api_key", "")

github_secrets = st.secrets.get("github", {})
GITHUB_TOKEN = github_secrets.get("token", "")
GITHUB_REPO = github_secrets.get("repo", "")


def clean_inline_duplicate(text: str) -> str:
    """Удаляет дублирование фраз, склеенных внутри одной строки (например, 'HelloHello')"""
    text = text.strip()
    if not text:
        return ""
    mid = len(text) // 2
    if len(text) % 2 == 0 and text[:mid] == text[mid:]:
        return text[:mid].strip()
    return text

def extract_raw_xml_text_from_zip(file_bytes) -> str:
    """Глубокое извлечение текста из DOCX, включая текстовые поля и фигуры."""
    WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    PARA_TAG = f'{WORD_NAMESPACE}p'
    TEXT_TAG = f'{WORD_NAMESPACE}t'
    NUM_PR_TAG = f'.//{WORD_NAMESPACE}numPr'

    all_extracted_lines = []
    
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
            xml_files = [f for f in z.namelist() if f.endswith('.xml') and f.startswith('word/')]
            xml_files.sort(key=lambda x: ('document' in x, x))

            for xml_file in xml_files:
                with z.open(xml_file) as f:
                    root = ET.fromstring(f.read())
                    for p in root.iter(PARA_TAG):
                        prefix = "• " if p.find(NUM_PR_TAG) is not None else ""
                        text_pieces = [node.text for node in p.iter(TEXT_TAG) if node.text]
                        p_text = "".join(text_pieces).strip()
                        p_text = clean_inline_duplicate(p_text)
                        
                        if p_text:
                            # Избегаем дублирования идущих подряд одинаковых строк
                            if not all_extracted_lines or all_extracted_lines[-1] != f"{prefix}{p_text}":
                                all_extracted_lines.append(f"{prefix}{p_text}")
    except Exception as e:
        st.error(f"Ошибка при XML-парсинге DOCX: {e}")
        return ""

    return "\n".join(all_extracted_lines)

def get_and_update_glossary(raw_text: str, folder_id: str, api_key: str, github_token: str, github_repo: str) -> dict:
    """Умное обновление словаря: сверка с GitHub и перевод только новых заголовков."""
    raw_sections = set(re.findall(r'(?im)^[ \t]*(?:section|раздел)\s*\d+.*$', raw_text))
    if not raw_sections:
        return {}

    if not github_token or not github_repo:
        st.error("GitHub токен или репозиторий не настроены в Secrets!")
        return {}

    g = Github(github_token)
    repo = g.get_repo(github_repo)
    file_path = "glossary.json"
    
    try:
        contents = repo.get_contents(file_path)
        current_glossary = json.loads(contents.decoded_content.decode("utf-8"))
    except Exception as e:
        st.warning(f"Не удалось загрузить словарь с GitHub. Создаем новый. Ошибка: {e}")
        current_glossary = {}
        contents = None

    new_sections = [sec for sec in raw_sections if sec not in current_glossary]
    
    if not new_sections:
        return current_glossary

    st.info(f"Найдено {len(new_sections)} новых заголовков. Отправляем в YandexGPT и коммитим в Git...")
    sections_list = "\n".join(new_sections)
    
    client = openai.OpenAI(
        api_key=api_key,
        base_url="https://ai.api.cloud.yandex.net/v1",
        project=folder_id
    )
    
    prompt = (
        "Ты — AI-модуль normalisation технической документации. Тебе дан список новых заголовков разделов из MSDS.\n"
        "Переведи их на русский язык (ГОСТ 30333-2022).\n\n"
        "ОБЯЗАТЕЛЬНОЕ ТРЕБОВАНИЕ: Верни ответ СТРОГО в формате валидного JSON-объекта, "
        "где КЛЮЧ — это оригинальная строка из списка (без изменений), а ЗНАЧЕНИЕ — её эталонный перевод.\n"
        "Не пиши никаких вступлений, комментариев или markdown-разметки. Только чистый JSON.\n\n"
        f"Список заголовков для перевода:\n{sections_list}"
    )
    
    try:
        response = client.responses.create(
            model=f"gpt://{folder_id}/yandexgpt",
            input=[{"role": "user", "content": prompt}],
            temperature=0.1
        )
        answer = response.output[0].content[0].text.strip()
        
        # Безопасное регулярное выражение с шестнадцатеричными кодами обратных кавычек
        answer = re.sub(r'\x60\x60\x60(?:json)?\s*|\s*\x60\x60\x60', '', answer)
        new_translations = json.loads(answer)
        
        current_glossary.update(new_translations)
        updated_content = json.dumps(current_glossary, ensure_ascii=False, indent=4)
        commit_message = f"Авто-обновление глоссария: добавлено {len(new_translations)} терминов"
        
        if contents:
            repo.update_file(contents.path, commit_message, updated_content, contents.sha)
        else:
            repo.create_file(file_path, commit_message, updated_content)
            
        st.success("Глоссарий успешно обновлен на GitHub!")
        
    except json.JSONDecodeError:
        st.error("YandexGPT вернул невалидный JSON. Обновление словаря отменено.")
    except Exception as e:
        st.error(f"Ошибка при обновлении глоссария: {e}")

    return current_glossary

def normalize_msds_with_glossary(text: str, glossary: dict) -> str:
    """Заменяет термины по словарю и аккуратно выравнивает структуру (ключ-значение)"""
    if not text.strip():
        return ""
    
    # 1. Сортируем словарь по убыванию длины. 
    # Это важно: сначала заменяем длинные предложения до точки, потом короткие фразы
    sorted_glossary_keys = sorted(glossary.keys(), key=len, reverse=True)
    
    for orig_term in sorted_glossary_keys:
        ru_term = glossary[orig_term]
        
        # Проверяем, является ли ключ из словаря ГЛАВНЫМ заголовком (SECTION)
        if bool(re.match(r'(?im)^[ \t]*(?:section|раздел)\s*\d+', orig_term)):
            cleaned_ru = ru_term.strip().lstrip('#').strip()
            if not cleaned_ru.lower().startswith('раздел'):
                num_match = re.search(r'\d+', orig_term)
                if num_match:
                    cleaned_ru = f"РАЗДЕЛ {num_match.group(0)}: {cleaned_ru}"
                    
            formatted_term = f"\n# {cleaned_ru}\n"
            text = text.replace(orig_term, formatted_term)
        else:
            # Если это обычная фраза, предложение до точки или подраздел — заменяем прямо в тексте (inline)
            text = text.replace(orig_term, ru_term)
            
    # Убираем лишние пробелы и чиним разорванные номера подразделов
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'^(\d+\.\d+\.?)([A-Za-zА-Яа-я])', r'\1 \2', text, flags=re.MULTILINE)
    
    lines = text.split('\n')
    cleaned_lines = []
    seen_sections = set()
    
    stop_patterns = [
        r'www\.spanlab\.in',
        r'Safety Data Sheet',
        r'MATERIAL SAFETY DATA SHEET'
    ]
    
    for line in lines:
        line_str = line.strip()
        if not line_str:
            continue
            
        if any(re.search(pat, line_str, re.IGNORECASE) for pat in stop_patterns):
            if len(line_str) < 50:
                continue
        
        line_str = re.sub(r':\s*:', ':', line_str)
        
        # ОПРЕДЕЛЯЕМ ТИП СТРОКИ
        is_main_section = line_str.startswith('# РАЗДЕЛ') or line_str.startswith('# SECTION')
        is_sub_section = bool(re.match(r'^(\d+\.\d+|\d+\.)', line_str)) or line_str.startswith('•') or line_str.startswith('-')
        
        # НОВОЕ: Детектор параметров "Ключ: Значение" (например "Form: Solid" или "pH : 7")
        is_key_value = bool(re.match(r'^[^:]+:', line_str)) and len(line_str.split(':')[0]) < 50
        
        if is_main_section:
            section_marker = " ".join(line_str.split()[:3]) 
            if section_marker in seen_sections:
                continue 
            seen_sections.add(section_marker)
            cleaned_lines.append('\n' + line_str)
            continue

        if cleaned_lines and line_str == cleaned_lines[-1].strip():
            continue
            
        # НОВОЕ: Если это подраздел, список ИЛИ параметр с двоеточием — строго с новой строки!
        if is_sub_section or is_key_value or line_str.endswith(':'):
            cleaned_lines.append('\n' + line_str)
        else:
            # Попытка склеить разорванный текст
            if cleaned_lines and not cleaned_lines[-1].startswith('\n# ') and not cleaned_lines[-1].endswith(':'):
                prev = cleaned_lines[-1]
                if line_str not in prev:
                    # Если строка короткая, не склеиваем ее сплошным текстом
                    if len(line_str) < 30:
                        cleaned_lines[-1] = prev + "\n" + line_str
                    else:
                        cleaned_lines[-1] = prev + " " + line_str
            else:
                cleaned_lines.append(line_str)
                
    normalized = '\n'.join(cleaned_lines)
    return re.sub(r'\n\s*\n+', '\n\n', normalized).strip()

def translate_msds_with_studio(text: str, folder_id: str, api_key: str, product_name_ru: str) -> str:
    """Перевод фрагментов через Yandex GPT на базе истинных SECTION"""
    if not text.strip():
        return ""
    
    if not folder_id or not api_key:
        st.error("Yandex Folder ID или API Key не настроены в Secrets!")
        return "Ошибка: Отсутствуют ключи авторизации Yandex."
        
    client = openai.OpenAI(
        api_key=api_key,
        base_url="https://ai.api.cloud.yandex.net/v1",
        project=folder_id
    )
    
    lines = text.split('\n')
    blocks = []
    current_block = []
    current_length = 0
    
    for line in lines:
        cleaned_line = line.strip()
        is_new_section = bool(re.match(r'^SECTION\s+\d+:', cleaned_line))
        
        if (is_new_section and current_block) or current_length > 2500:
            blocks.append('\n'.join(current_block))
            current_block = []
            current_length = 0
            
        current_block.append(line)
        current_length += len(line)
        
    if current_block:
        blocks.append('\n'.join(current_block))
            
    blocks = [b.strip() for b in blocks if b.strip()]
    translated_blocks = []
    
    system_instruction = (
        "Ты — высококлассный технический переводчик и эксперт по химической безопасности. "
        "Твоя задача — перевести фрагмент MSDS на русский язык (ГОСТ 30333-2022) и ОФОРМИТЬ ЕГО В СТРОГОМ MARKDOWN.\n\n"
        f"КРИТИЧЕСКИ ВАЖНОЕ ТРЕБОВАНИЕ: Везде, где в тексте упоминается название продукта, "
        f"ты ОБЯЗАН использовать исключительно название '{product_name_ru}'. Не склоняй его и не изменяй.\n\n"
        "ПРАВИЛА ФОРМАТИРОВАНИЯ:\n"
        "1. Главные разделы (SECTION / РАЗДЕЛ) выделяй одной решеткой: `# РАЗДЕЛ X: Название`.\n"
        "2. Подразделы (1.1, 14.2) выделяй двумя решетками: `## 1.1 Название`.\n"
        "3. Разделяй параметры и значения! Оформляй параметры жирным: `**Цвет:** Белые чешуйки`.\n"
        "4. Списки оформляй через дефис `- `.\n"
        "Убирай пустые строки. Выдавай ТОЛЬКО чистый перевод без комментариев."
    )
    
    progress_bar = st.progress(0)
    total_blocks = len(blocks)
    
    for i, block in enumerate(blocks):
        if not block.strip():
            continue
        try:
            response = client.responses.create(
                model=f"gpt://{folder_id}/yandexgpt",
                instructions=system_instruction,
                input=[{"role": "user", "content": block}],
                temperature=0.1,
                max_output_tokens=4000
            )
            if response.output and response.output[0].content and response.output[0].content[0].text:
                translated_blocks.append(response.output[0].content[0].text)
            else:
                translated_blocks.append(block)
        except Exception as e:
            st.warning(f"Ошибка на блоке {i+1}: {str(e)}")
            translated_blocks.append(block)
            
        progress_bar.progress(min((i + 1) / total_blocks, 1.0))
        
    progress_bar.empty()
    return '\n\n'.join(translated_blocks)

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def make_formatted_docx(markdown_text: str, product_name_ru: str, product_cas: str):
    """Сборщик Word-документа по ГОСТ-стилистике"""
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.39)
        section.bottom_margin = Inches(0.39)
        section.left_margin = Inches(0.39)
        section.right_margin = Inches(0.39)
        
        hp = section.header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run(f"{product_name_ru} | Паспорт безопасности химической продукции")
        hrun.font.name = 'Arial'
        hrun.font.size = Pt(8.5)
        hrun.font.italic = True
        hrun.font.color.rgb = RGBColor(128, 128, 128)
        
        fp = section.footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        frun = fp.add_run("www.spanlab.in                                                                                  ")
        frun.font.name = 'Arial'
        frun.font.size = Pt(8.5)
        frun.font.color.rgb = RGBColor(128, 128, 128)
        
    p_header1 = doc.add_paragraph()
    p_header1.paragraph_format.space_before = Pt(0)
    p_header1.paragraph_format.space_after = Pt(2)
    run_h1 = p_header1.add_run("Паспорт безопасности материала")
    run_h1.bold = True
    run_h1.font.name = 'Arial'
    run_h1.font.size = Pt(16)
    
    p_header2 = doc.add_paragraph()
    p_header2.paragraph_format.space_before = Pt(0)
    p_header2.paragraph_format.space_after = Pt(2)
    run_h2 = p_header2.add_run(product_name_ru)
    run_h2.bold = True
    run_h2.font.name = 'Arial'
    run_h2.font.size = Pt(16)
    
    p_header3 = doc.add_paragraph()
    p_header3.paragraph_format.space_before = Pt(0)
    p_header3.paragraph_format.space_after = Pt(18)
    run_h3 = p_header3.add_run(f"CAS № {product_cas.strip()}")
    run_h3.bold = True
    run_h3.font.name = 'Arial'
    run_h3.font.size = Pt(16)
    
    pPr = p_header3._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')       
    bottom.set(qn('w:space'), '8')    
    bottom.set(qn('w:color'), '000000')
    pBdr.append(bottom)
    pPr.append(pBdr)

    DARK_BLUE = RGBColor(0, 51, 102)
    lines = markdown_text.split('\n')
    
    for line in lines:
        cleaned_line = line.strip()
        if not cleaned_line:
            continue
            
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        
        if cleaned_line.startswith('# '):
            text_content = cleaned_line.replace('# ', '').strip()
            run = p.add_run(text_content)
            run.bold = True
            run.font.name = 'Arial'
            run.font.size = Pt(12)
            run.font.color.rgb = DARK_BLUE
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(6)
        elif cleaned_line.startswith('## '):
            text_content = cleaned_line.replace('## ', '').strip()
            run = p.add_run(text_content)
            run.font.name = 'Arial'
            run.font.size = Pt(11)
            run.font.color.rgb = DARK_BLUE
            p.paragraph_format.space_before = Pt(6)
        else:
            if cleaned_line.startswith('- '):
                cleaned_line = cleaned_line.replace('- ', '', 1)
                p.paragraph_format.left_indent = Inches(0.25)
            
            parts = re.split(r'(\*\*.*?\*\*)', cleaned_line)
            for part in parts:
                if part.startswith('**') and part.endswith('**'):
                    bold_text = part.replace('**', '')
                    run = p.add_run(bold_text)
                    run.bold = True
                else:
                    run = p.add_run(part)
                run.font.name = 'Arial'
                run.font.size = Pt(9)
                
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def render_glossary_tab():
    """Вкладка с интерактивной таблицей для управления глоссарием на GitHub"""
    st.header("Управление глоссарием")
    st.caption("Здесь вы можете просматривать, изменять и удалять записи словаря. Изменения автоматически улетят на GitHub.")
    
    if not GITHUB_TOKEN or not GITHUB_REPO:
        st.error("GitHub конфигурации не найдены в Streamlit Secrets!")
        return

    try:
        g = Github(GITHUB_TOKEN)
        repo = g.get_repo(GITHUB_REPO)
        contents = repo.get_contents("glossary.json")
        glossary_data = json.loads(contents.decoded_content.decode("utf-8"))
        
        data_list = [{"Оригинал (English)": k, "Перевод (Russian)": v} for k, v in glossary_data.items()]
        
        edited_df = st.data_editor(data_list, use_container_width=True, num_rows="dynamic")
        
        if st.button("💾 Сохранить изменения в словаре", type="primary"):
            updated_dict = {row["Оригинал (English)"]: row["Перевод (Russian)"] for row in edited_df if row["Оригинал (English)"]}
            new_content = json.dumps(updated_dict, ensure_ascii=False, indent=4)
            
            repo.update_file(
                contents.path, 
                "Ручное редактирование глоссария через интерфейс", 
                new_content, 
                contents.sha
            )
            st.success("Словарь успешно обновлен на GitHub! Изменения применятся к следующим переводам.")
            
    except Exception as e:
        st.error(f"Не удалось загрузить данные с GitHub. Ошибка: {e}")

# --- Инициализация состояния ---
if "raw_text" not in st.session_state:
    st.session_state.raw_text = ""
if "normalized_text" not in st.session_state:
    st.session_state.normalized_text = ""
if "translated_text" not in st.session_state:
    st.session_state.translated_text = ""
if "file_name_output" not in st.session_state:
    st.session_state.file_name_output = "MSDS_RU_Translated"

def reset_state():
    st.session_state.raw_text = ""
    st.session_state.normalized_text = ""
    st.session_state.translated_text = ""

# --- Основной Интерфейс ---
st.title("🧪 MSDS Translator — Premium AI Studio")

# Разбиваем на две вкладки
tab_main, tab_glossary = st.tabs(["🔄 Переводчик MSDS", "📚 Редактор глоссария"])

with tab_main:
    st.sidebar.header("Параметры продукта")
    product_name_ru = st.sidebar.text_input("Официальное название продукта (RU):", value="ТРИМЕТИЛОЛПРОПАН")
    product_cas = st.sidebar.text_input("Номер CAS:", value="77-99-6")

    # --- ШАГ 1 ---
    st.header("Шаг 1: Загрузка исходного MSDS (EN)")
    input_method = st.radio("Способ загрузки:", ("Загрузить файл (DOCX / PDF / TXT)", "Вставить текст вручную"), on_change=reset_state)

    if input_method == "Вставить текст вручную":
        inserted_text = st.text_area("Вставьте текст MSDS на английском языке:", height=250, placeholder="SECTION 1: Identification...")
        if inserted_text:
            st.session_state.raw_text = inserted_text
    else:
        uploaded_file = st.file_uploader("Выберите файл", type=["docx", "pdf", "txt"])
        if uploaded_file is not None:
            st.session_state.file_name_output = f"{uploaded_file.name}_RU"
            if uploaded_file.name.endswith(".txt"):
                st.session_state.raw_text = str(uploaded_file.read(), "utf-8")
            elif uploaded_file.name.endswith(".docx"):
                st.session_state.raw_text = extract_raw_xml_text_from_zip(uploaded_file.read())
            elif uploaded_file.name.endswith(".pdf"):
                with pdfplumber.open(io.BytesIO(uploaded_file.read())) as pdf:
                    st.session_state.raw_text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

    if st.session_state.raw_text:
        with st.expander("Просмотр извлеченного оригинального текста (Шаг 1)", expanded=False):
            st.text_area("Оригинал без изменений:", value=st.session_state.raw_text, height=200, disabled=True, key="raw_preview")

    st.divider()

    # --- ШАГ 2 ---
    st.header("Шаг 2: Выравнивание и нормализация табличной структуры")
    st.caption("Автоматически сверяется с глоссарием на GitHub, добавляет новые заголовки и выравнивает текст.")

    if st.button("🔧 Запустить нормализацию текста", type="secondary", use_container_width=True):
        if st.session_state.raw_text:
            with st.spinner("Сверка словаря с GitHub и анализ документа..."):
                glossary = get_and_update_glossary(
                    st.session_state.raw_text, 
                    FOLDER_ID, 
                    API_KEY, 
                    GITHUB_TOKEN,
                    GITHUB_REPO
                )
                
            with st.spinner("Выравнивание структуры и удаление дубликатов..."):
                st.session_state.normalized_text = normalize_msds_with_glossary(st.session_state.raw_text, glossary)
                
            st.success("Успех! Текст нормализован.")
        else:
            st.warning("Сначала загрузите или вставьте исходный текст на Шаге 1.")

    st.divider()

    # --- ШАГ 3 ---
    st.header("Шаг 3: Перевод через YandexGPT")

    if st.button("Выполнить перевод", type="primary", use_container_width=True):
        if st.session_state.normalized_text:
            with st.spinner("YandexGPT переводит документ..."):
                st.session_state.translated_text = translate_msds_with_studio(
                    st.session_state.normalized_text, FOLDER_ID, API_KEY, product_name_ru
                )
            st.success("Перевод завершен!")
        else:
            st.warning("Нечего переводить. Сначала выполните Шаг 2.")

    if st.session_state.translated_text:
        with st.expander("Предпросмотр готового перевода Markdown (Шаг 3)", expanded=True):
            st.markdown(st.session_state.translated_text)

    st.divider()

    # --- ШАГ 4 ---
    st.header("Шаг 4: Экспорт в Word")

    if st.session_state.translated_text:
        if "Ошибка" not in st.session_state.translated_text:
            docx_data = make_formatted_docx(st.session_state.translated_text, product_name_ru, product_cas)
            st.download_button(
                label="Скачать отформатированный файл WORD (.docx)",
                data=docx_data,
                file_name=st.session_state.file_name_output if st.session_state.file_name_output.endswith(".docx") else f"{st.session_state.file_name_output}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
    else:
        st.info("Кнопка скачивания появится здесь, когда Шаг 3 будет успешно выполнен.")

with tab_glossary:
    render_glossary_tab()
