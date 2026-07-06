import streamlit as st
import openai
import pdfplumber
import io
import re
import os
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- Настройка страницы ---
st.set_page_config(
    page_title="MSDS Yandex AI Studio Pro",
    page_icon="🧪",
    layout="wide"
)

def normalize_msds_text(text: str) -> str:
    """Шаг 2: Нормализация текста с защитой от ложных разделов (ссылок вида see section 16)"""
    if not text.strip():
        return ""
    
    # 1. Заменяем множественные пробелы и горизонтальные табы на один пробел
    text = re.sub(r'[ \t]+', ' ', text)
    
    # 2. Стандартизируем ТОЛЬКО реальные заголовки разделов (когда SECTION стоит в начале строки)
    # Игнорируем упоминания внутри предложений (например, "see section 13")
    text = re.sub(r'(?im)^[ \t]*(section|раздел)\s*(\d+)\b', r'\nSECTION \2: ', text)
    
    lines = text.split('\n')
    cleaned_lines = []
    
    for line in lines:
        line_str = line.strip()
        if not line_str:
            continue
            
        # Строка признается заголовком раздела, только если она начинается строго с SECTION [цифра]:
        # и она относительно короткая (заголовок не бывает длиной во весь абзац)
        is_main_section = bool(re.match(r'^SECTION\s+\d+:', line_str)) and len(line_str) < 120
        
        # Проверка на подразделы (1.1, 2.3 и т.д.) или пункты списков
        is_sub_section = bool(re.match(r'^(\d+\.\d+|\d+\.)', line_str)) or line_str.startswith('-')
        
        is_structural = is_main_section or is_sub_section or line_str.endswith(':')
        
        if is_structural:
            cleaned_lines.append('\n' + line_str)  # Выделяем структуру
        else:
            # Если это обычный текст, и предыдущая строка не была главным заголовком раздела,
            # склеиваем её с предыдущей (лечим разрывы таблиц)
            if cleaned_lines and not bool(re.match(r'^\nSECTION\s+\d+:', cleaned_lines[-1])) and not cleaned_lines[-1].endswith(':'):
                cleaned_lines[-1] = cleaned_lines[-1] + " " + line_str
            else:
                cleaned_lines.append(line_str)
                
    normalized = '\n'.join(cleaned_lines)
    normalized = re.sub(r'\n\s*\n+', '\n\n', normalized)
    
    return normalized.strip()

def translate_msds_with_studio(text: str, folder_id: str, api_key: str, product_name_ru: str) -> str:
    """Шаг 3: Перевод MSDS с жестким контролем деления блоков по истинным SECTION"""
    if not text.strip():
        return ""
    
    client = openai.OpenAI(
        api_key=api_key,
        base_url="https://ai.api.cloud.yandex.net/v1",
        project=folder_id
    )
    
    YANDEX_MODEL = "yandexgpt" 
    
    lines = text.split('\n')
    blocks = []
    current_block = []
    current_length = 0
    
    for line in lines:
        cleaned_line = line.strip()
        
        # Блок закрывается ТОЛЬКО если мы встретили истинный заголовок "SECTION [цифра]:"
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
        f"КРИТИЧЕСКИ ВАЖНОЕ ТРЕБОВАНИЕ: Везде, где в тексте упоминается название продукта (в заголовках, свойствах, синонимах), "
        f"ты ОБЯЗАН использовать исключительно название '{product_name_ru}'. "
        f"Не склоняй его, не переводи дословно, не изменяй и не редактируй. Пиши ровно так: {product_name_ru}.\n\n"
        "ПРАВИЛА ЖЕЛЕЗНОГО ФОРМАТИРОВАНИЯ:\n"
        "1. Главные разделы (SECTION / РАЗДЕЛ) выделяй одной решеткой: `# РАЗДЕЛ X: Название`.\n"
        "2. Подразделы (1.1, 14.2 и т.д.) выделяй двумя решетками: `## 1.1 Название подраздела`.\n"
        "3. Разделяй параметры и значения! Если строка содержит технический параметр и его значение "
        "(например: 'Colour: White scales' или 'Flash point: 172 °C'), ты ОБЯЗАН оформить параметр жирным, "
        "а значение оставить обычным. Пример: `**Цвет:** Белые чешуйки`.\n"
        "4. Списки оформляй через дефис `- `.\n"
        "5. КРИТИЧЕСКИ ВАЖНО: Сохраняй оригинальную нумерацию пунктов и подпунктов (1., 1.1, a), b)) в точности.\n"
        "Убирай пустые строки. Выдавай ТОЛЬКО чистый Markdown перевод без своих комментариев."
    )
    
    progress_bar = st.progress(0)
    total_blocks = len(blocks)
    
    for i, block in enumerate(blocks):
        if not block.strip():
            continue
            
        try:
            response = client.responses.create(
                model=f"gpt://{folder_id}/{YANDEX_MODEL}",
                instructions=system_instruction,
                input=[{"role": "user", "content": block}],
                temperature=0.1, 
                max_output_tokens=4000  
            )
            
            if response.output and response.output[0].content and response.output[0].content[0].text:
                translated_text = response.output[0].content[0].text
                translated_blocks.append(translated_text)
            else:
                translated_blocks.append(block)
                
        except Exception as e:
            st.warning(f"Ошибка на блоке {i+1}: {str(e)}")
            translated_blocks.append(block)
            
        progress_bar.progress(min((i + 1) / total_blocks, 1.0))
        
    progress_bar.empty()
    return '\n\n'.join(translated_blocks)

def make_formatted_docx(markdown_text: str, product_name_ru: str):
    """Шаг 4: Сборщик Word-документа с интеграцией официального названия в колонтитулы"""
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
        
        fpage = fp.add_run("Страница [Page]")
        fpage.font.name = 'Arial'
        fpage.font.size = Pt(8.5)
        fpage.font.color.rgb = RGBColor(128, 128, 128)

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


# --- Интерфейс Streamlit ---
st.title("🧪 MSDS Translator — Premium AI Studio")
st.caption("Пошаговый конвейер с фильтрацией ложных разделов и перекрестных ссылок.")

st.divider()

st.sidebar.header("🔑 Доступ к Yandex AI Studio")
folder_id = st.sidebar.text_input("Yandex Folder ID", type="password")
api_key = st.sidebar.text_input("Yandex API Key", type="password")

st.sidebar.markdown("---")
st.sidebar.header("📦 Химическая номенклатура")
product_name_ru = st.sidebar.text_input("Официальное название продукта (RU):", value="ТРИМЕТИЛОЛПРОПАН")

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

# --- ШАГ 1 ---
st.header("Шаг 1: Загрузка исходного MSDS (EN)")
input_method = st.radio("Способ загрузки:", ("Загрузить файл (DOCX / TXT)", "Вставить текст вручную"), on_change=reset_state)

if input_method == "Вставить текст вручную":
    inserted_text = st.text_area("Вставьте текст MSDS на английском языке:", height=250, placeholder="SECTION 1: Identification...")
    if inserted_text:
        st.session_state.raw_text = inserted_text
else:
    uploaded_file = st.file_uploader("Выберите файл", type=["docx", "txt"])
    if uploaded_file is not None:
        st.session_state.file_name_output = f"Translated_{uploaded_file.name}"
        if uploaded_file.name.endswith(".txt"):
            st.session_state.raw_text = str(uploaded_file.read(), "utf-8")
        elif uploaded_file.name.endswith(".docx"):
            doc = Document(io.BytesIO(uploaded_file.read()))
            st.session_state.raw_text = "\n".join([para.text for para in doc.paragraphs])

if st.session_state.raw_text:
    with st.expander("🔍 Просмотр оригинального текста (Шаг 1)", expanded=True):
        st.text_area("Оригинал без изменений:", value=st.session_state.raw_text, height=200, disabled=True, key="raw_preview")

st.divider()

# --- ШАГ 2 ---
st.header("Шаг 2: Выравнивание и нормализация табличной структуры")
st.caption("Склеивает таблицы, защищая ссылки вида 'see section 16' от превращения в новые разделы.")

if st.button("🔧 Запустить нормализацию текста", type="secondary", use_container_width=True):
    if st.session_state.raw_text:
        st.session_state.normalized_text = normalize_msds_text(st.session_state.raw_text)
        st.success("Текст успешно нормализован и очищен!")
    else:
        st.warning("Сначала загрузите или вставьте исходный текст на Шаге 1.")

if st.session_state.normalized_text:
    with st.expander("🛠️ Просмотр нормализованного текста (Шаг 2)", expanded=True):
        st.text_area("Текст, готовый к отправке в нейросеть (проверьте отсутствие ложных SECTION):", value=st.session_state.normalized_text, height=250, disabled=True, key="norm_preview")

st.divider()

# --- ШАГ 3 ---
st.header("Шаг 3: Перевод через YandexGPT")

if st.button("🔄 Выполнить перевод (строго по секциям)", type="primary", use_container_width=True):
    if not folder_id or not api_key:
        st.warning("Пожалуйста, введите Yandex Folder ID и API Key в боковой панели.")
    elif st.session_state.normalized_text:
        with st.spinner("YandexGPT переводит документ..."):
            st.session_state.translated_text = translate_msds_with_studio(
                st.session_state.normalized_text, folder_id, api_key, product_name_ru
            )
        st.success("Перевод завершен!")
    else:
        st.warning("Нечего переводить. Сначала выполните Шаг 2.")

if st.session_state.translated_text:
    with st.expander("📄 Предпросмотр готового перевода Markdown (Шаг 3)", expanded=True):
        st.markdown(st.session_state.translated_text)

st.divider()

# --- ШАГ 4 ---
st.header("Шаг 4: Экспорт в Word")

if st.session_state.translated_text:
    if "Ошибка" not in st.session_state.translated_text:
        docx_data = make_formatted_docx(st.session_state.translated_text, product_name_ru)
        st.download_button(
            label="💾 Скачать отформатированный файл WORD (.docx)",
            data=docx_data,
            file_name=st.session_state.file_name_output if st.session_state.file_name_output.endswith(".docx") else f"{st.session_state.file_name_output}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
else:
    st.info("Кнопка скачивания появится здесь, когда Шаг 3 будет успешно выполнен.")
