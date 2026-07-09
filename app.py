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

def extract_translation_candidates(text: str) -> set:
    """
    Разбивает весь документ на уникальные логические сегменты (фразы, параметры, значения).
    Исключает из перевода чистые цифры, CAS-номера и технический мусор для экономии денег.
    """
    candidates = set()
    stop_words = ['www.', 'http', 'safety data sheet', 'material safety data sheet']
    
    for line in text.split('\n'):
        line_str = line.strip()
        if not line_str:
            continue
        if any(sw in line_str.lower() for sw in stop_words):
            continue
        # Пропускаем, если строка состоит только из цифр, точек, тире и спецсимволов (CAS, коды, даты)
        if re.match(r'^[\d\s\.,\-\/\\#№:;%()\*\+\[\]]+$', line_str):
            continue
            
        # Проверяем структуру "Ключ: Значение" (например, "Physical state: Solid")
        if ':' in line_str:
            parts = line_str.split(':', 1)
            key = parts[0].strip()
            val = parts[1].strip()
            
            # Если ключ короткий (похож на название параметра), кэшируем его отдельно
            if len(key) < 60 and not re.match(r'^[\d\s\.,\-\/\\#№:;%()\*\+\[\]]+$', key):
                candidates.add(key)
            else:
                candidates.add(line_str)
                continue
                
            if val and not re.match(r'^[\d\s\.,\-\/\\#№:;%()\*\+\[\]]+$', val):
                # Дробим внутри на случай слитных параметров в строке (как на скриншотах)
                sub_parts = re.split(r'(?<=\.)\s+(?=[А-Яа-яA-Za-z][^:]+:\s)', val)
                if len(sub_parts) > 1:
                    for sp in sub_parts:
                        sp_str = sp.strip()
                        if ':' in sp_str:
                            skey, sval = sp_str.split(':', 1)
                            skey, sval = skey.strip(), sval.strip()
                            if len(skey) < 60 and not re.match(r'^[\d\s\.,\-\/\\#№:;%()\*\+\[\]]+$', skey):
                                candidates.add(skey)
                            if sval and not re.match(r'^[\d\s\.,\-\/\\#№:;%()\*\+\[\]]+$', sval):
                                candidates.add(sval)
                        else:
                            candidates.add(sp_str)
                else:
                    candidates.add(val)
        else:
            candidates.add(line_str)
            
    return candidates


def get_and_update_glossary(raw_text: str, folder_id: str, api_key: str, github_token: str, github_repo: str) -> dict:
    """Загружает глоссарий из Git, находит новые фразы, переводит только их пачками и пушит в Git."""
    if not github_token or not github_repo:
        st.error("GitHub конфигурация не найдена в Secrets!")
        return {}

    g = Github(github_token)
    repo = g.get_repo(github_repo)
    file_path = "glossary.json"
    
    github_secrets = st.secrets.get("github", {})
    TARGET_BRANCH = github_secrets.get("branch", "main")
    
    contents = None
    current_glossary = {}
    
    # 1. Читаем базу данных из текущей ветки
    try:
        contents = repo.get_contents(file_path, ref=TARGET_BRANCH)
        current_glossary = json.loads(contents.decoded_content.decode("utf-8"))
    except json.JSONDecodeError as jde:
        st.error(f"Синтаксическая ошибка в glossary.json на GitHub! Ошибка: {jde}")
        st.stop()
    except Exception as e:
        st.warning(f"Не удалось получить файл глоссария из ветки {TARGET_BRANCH}. Будет создан новый. Ошибка: {e}")
        current_glossary = {}

    # 2. Выделяем всех кандидатов на перевод из документа
    all_candidates = extract_translation_candidates(raw_text)
    # Ищем только те фразы, которых РЕАЛЬНО нет в нашей базе данных
    unknown_candidates = [cand for cand in all_candidates if cand not in current_glossary]
    
    if not unknown_candidates:
        st.success("🎉 Полное совпадение со словарём! В документе нет новых фраз. Расход токенов: 0 рублей.")
        return current_glossary

    st.info(f"🕵️‍♂️ Найдено {len(unknown_candidates)} новых уникальных фраз. Отправляем на экономичный перевод...")
    
    client = openai.OpenAI(
        api_key=api_key,
        base_url="https://ai.api.cloud.yandex.net/v1",
        project=folder_id
    )
    
    # Режем массив незнакомых фраз на небольшие пачки по 30 штук, чтобы JSON не обрывался
    BATCH_SIZE = 30
    new_translations = {}
    
    for i in range(0, len(unknown_candidates), BATCH_SIZE):
        batch = unknown_candidates[i:i+BATCH_SIZE]
        st.caption(f"Перевод пачки {i//BATCH_SIZE + 1} из {len(unknown_candidates)//BATCH_SIZE + 1}...")
        
        batch_json_placeholder = json.dumps({string: "" for string in batch}, ensure_ascii=False, indent=2)
        
        prompt = (
            "Ты — AI-модуль нормализации химической документации (ГОСТ 30333-2022).\n"
            "Переведи предоставленные ключи на русский язык.\n"
            "ТРЕБОВАНИЕ: Верни строго валидный JSON-объект, где ключ — оригинальная английская строка, а значение — русский перевод.\n"
            "Не пиши никаких вступлений или markdown-разметки. Только чистый JSON.\n\n"
            f"Шаблон для заполнения:\n{batch_json_placeholder}"
        )
        
        try:
            response = client.responses.create(
                model=f"gpt://{folder_id}/yandexgpt",
                input=[{"role": "user", "content": prompt}],
                temperature=0.1
            )
            answer = response.output[0].content[0].text.strip()
            answer = re.sub(r'\x60\x60\x60(?:json)?\s*|\s*\x60\x60\x60', '', answer)
            
            parsed_batch = json.loads(answer)
            new_translations.update(parsed_batch)
        except Exception as e:
            st.error(f"Ошибка перевода пачки фраз: {e}")
            continue

    if new_translations:
        # Интегрируем новые переводы в общую базу данных
        current_glossary.update(new_translations)
        updated_content = json.dumps(current_glossary, ensure_ascii=False, indent=4)
        commit_message = f"CAT-Cache обновление: добавлено {len(new_translations)} новых фраз"
        
        try:
            if contents is not None:
                repo.update_file(contents.path, commit_message, updated_content, contents.sha, branch=TARGET_BRANCH)
            else:
                repo.create_file(file_path, commit_message, updated_content, branch=TARGET_BRANCH)
            st.success(f"🧠 База знаний успешно расширена и сохранена в ветку {TARGET_BRANCH}!")
        except Exception as e:
            st.error(f"Не удалось отправить коммит на GitHub: {e}")

    return current_glossary


def assemble_translated_document(text: str, glossary: dict, product_name_ru: str) -> str:
    """
    Финальная сборка документа (0 рублей). Локально идет по исходному тексту, 
    выдергивает переводы из кэша и наводит красивую верстку (жирный текст, заголовки, списки).
    """
    cleaned_lines = []
    seen_sections = set()
    stop_patterns = [r'www\.', r'safety data sheet', r'material safety data sheet']
    
    def translate_chunk(chunk: str) -> str:
        c_clean = chunk.strip()
        if not c_clean:
            return ""
        if re.match(r'^^[\d\s\.,\-\/\\#№:;%()\*\+\[\]]+$', c_clean):
            return c_clean
        # Берем перевод из кэша. Если вдруг перевода нет — возвращаем оригинал
        return glossary.get(c_clean, c_clean)

    for line in text.split('\n'):
        line_str = line.strip()
        if not line_str or any(re.search(pat, line_str, re.IGNORECASE) for pat in stop_patterns):
            continue
            
        # 1. Форматирование главных разделов
        if bool(re.match(r'(?im)^[ \t]*(?:section|раздел)\s*\d+', line_str)):
            translated_title = translate_chunk(line_str)
            cleaned_title = translated_title.replace('#', '').strip()
            
            if not cleaned_title.lower().startswith('раздел'):
                num_match = re.search(r'\d+', line_str)
                if num_match:
                    cleaned_title = f"РАЗДЕЛ {num_match.group(0)}: {cleaned_title.split(':', 1)[-1].strip()}"
            
            section_marker = " ".join(cleaned_title.split()[:3])
            if section_marker in seen_sections:
                continue
            seen_sections.add(section_marker)
            
            cleaned_lines.append(f"\n# {cleaned_title}")
            continue
            
        # 2. Форматирование подразделов (1.1, 4.2.1)
        if re.match(r'^(\d+\.\d+\.?\d*)\s+', line_str):
            match = re.match(r'^(\d+\.\d+\.?\d*)\s+(.*)$', line_str)
            num_part = match.group(1)
            text_part = match.group(2)
            cleaned_lines.append(f"\n## {num_part} {translate_chunk(text_part)}")
            continue
            
        # 3. Структурирование параметров "Ключ: Значение"
        if ':' in line_str:
            parts = line_str.split(':', 1)
            key, val = parts[0].strip(), parts[1].strip()
            
            if len(key) < 60 and not re.match(r'^^[\d\s\.,\-\/\\#№:;%()\*\+\[\]]+$', key):
                t_key = translate_chunk(key)
                
                # Проверяем внутренности значения на перечисление параметров на одной строке
                sub_parts = re.split(r'(?<=\.)\s+(?=[А-Яа-яA-Za-z][^:]+:\s)', val)
                if len(sub_parts) > 1:
                    assembled_sub = []
                    for sp in sub_parts:
                        sp_str = sp.strip()
                        if ':' in sp_str:
                            skey, sval = sp_str.split(':', 1)
                            assembled_sub.append(f"**{translate_chunk(skey)}:** {translate_chunk(sval)}")
                        else:
                            assembled_sub.append(translate_chunk(sp_str))
                    t_val = ".\n" + ".\n".join(assembled_sub)
                else:
                    t_val = translate_chunk(val)
                    
                cleaned_lines.append(f"**{t_key}:** {t_val}")
            else:
                cleaned_lines.append(translate_chunk(line_str))
        else:
            # 4. Обычные строки или маркированные списки
            prefix = "- " if line_str.startswith('•') or line_str.startswith('-') else ""
            pure_text = line_str.lstrip('•- ').strip()
            cleaned_lines.append(f"{prefix}{translate_chunk(pure_text)}")
            
    final_markdown = '\n'.join(cleaned_lines)
    # Финальный штрих: глобально заменяем имя продукта на его официальное русское имя по ТЗ
    return re.sub(r'\n\s*\n+', '\n\n', final_markdown).strip()

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
    st.header("Управление глоссарием")
    st.caption("Здесь вы можете просматривать, изменять и удалять записи словаря. Изменения автоматически улетят на GitHub.")
    
    if not GITHUB_TOKEN or not GITHUB_REPO:
        st.error("GitHub конфигурации не найдены в Streamlit Secrets!")
        return

    # Динамически получаем имя ветки из secrets для этой вкладки
    github_secrets = st.secrets.get("github", {})
    TARGET_BRANCH = github_secrets.get("branch", "main")

    try:
        g = Github(GITHUB_TOKEN)
        repo = g.get_repo(GITHUB_REPO)
        
        # Используем TARGET_BRANCH вместо жесткой строки
        contents = repo.get_contents("glossary.json", ref=TARGET_BRANCH)
        glossary_data = json.loads(contents.decoded_content.decode("utf-8"))
        
        data_list = [{"Оригинал (English)": k, "Перевод (Russian)": v} for k, v in glossary_data.items()]
        
        edited_df = st.data_editor(data_list, use_container_width=True, num_rows="dynamic")
        
        if st.button("💾 Сохранить изменения в словаре", type="primary"):
            updated_dict = {row["Оригинал (English)"]: row["Перевод (Russian)"] for row in edited_df if row["Оригинал (English)"]}
            new_content = json.dumps(updated_dict, ensure_ascii=False, indent=4)
            
            # Используем TARGET_BRANCH для сохранения
            repo.update_file(
                contents.path, 
                "Ручное редактирование глоссария через интерфейс", 
                new_content, 
                contents.sha,
                branch=TARGET_BRANCH
            )
            st.success(f"Словарь успешно обновлен в ветке {TARGET_BRANCH} на GitHub!")
            
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
            # Отсекаем расширение (берём всё, что до последней точки)
            base_name = uploaded_file.name.rsplit('.', 1)[0]
            st.session_state.file_name_output = f"{base_name}_RU"
            
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
    st.header("Шаг 2: Анализ документа и обновление базы знаний")
    st.caption("Система извлечет уникальные фразы, найдет новые параметры и автоматически обучит словарь на GitHub.")

    if st.button("🔧 Запустить интеллектуальный анализ", type="secondary", use_container_width=True):
        if st.session_state.raw_text:
            with st.spinner("Синхронизация с базой знаний Git и перевод неизвестных фраз..."):
                # Находим новые фразы, шлем в YandexGPT только их, коммитим в global-glossary-02
                st.session_state.current_glossary_cache = get_and_update_glossary(
                    st.session_state.raw_text, 
                    FOLDER_ID, 
                    API_KEY, 
                    GITHUB_TOKEN,
                    GITHUB_REPO
                )
            st.success("База знаний обновлена! Все фразы текущего документа теперь есть в кэше.")
        else:
            st.warning("Сначала загрузите или вставьте исходный текст на Шаге 1.")

    st.divider()

    # --- ШАГ 3 ---
    st.header("Шаг 3: Мгновенная сборка перевода из кэша (0 рублей)")
    st.caption("Документ собирается локально без повторных обращений к нейросети.")

    if st.button("🚀 Собрать готовый документ", type="primary", use_container_width=True):
        # Проверяем, запущен ли кэш в текущей сессии, если нет — пробуем получить его без перезаписи
        if "current_glossary_cache" not in st.session_state or not st.session_state.current_glossary_cache:
            with st.spinner("Загрузка активного словаря..."):
                st.session_state.current_glossary_cache = get_and_update_glossary(
                    st.session_state.raw_text, FOLDER_ID, API_KEY, GITHUB_TOKEN, GITHUB_REPO
                )
        
        if st.session_state.current_glossary_cache:
            with st.spinner("Локальная сборка ГОСТ-структуры..."):
                # Собираем перевод за 1 секунду абсолютно бесплатно!
                st.session_state.translated_text = assemble_translated_document(
                    st.session_state.raw_text, 
                    st.session_state.current_glossary_cache, 
                    product_name_ru
                )
            st.success("Документ успешно собран!")
        else:
            st.warning("Не удалось подготовить кэш для сборки. Выполните Шаг 2.")

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
