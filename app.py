import streamlit as st
import openai
import pdfplumber
import io
import re
import os
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Настройка страницы
st.set_page_config(
    page_title="MSDS Yandex AI Studio Pro",
    page_icon="🧪",
    layout="wide"
)

def translate_msds_with_studio(text: str, folder_id: str, api_key: str, product_name_ru: str) -> str:
    """Перевод MSDS с жестким требованием разметки Markdown и фиксированным именем продукта"""
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
        current_block.append(line)
        current_length += len(line)
        if current_length > 3000:
            blocks.append('\n'.join(current_block))
            current_block = []
            current_length = 0
    if current_block:
        blocks.append('\n'.join(current_block))
            
    translated_blocks = []
    
    # ЖЕСТКИЙ ПРОМПТ С ДИРЕКТИВОЙ ПО НАЗВАНИЮ ПРОДУКТА
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
                max_output_tokens=3000
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
    """Сборщик Word-документа с интеграцией официального названия в колонтитулы"""
    doc = Document()
    
    # Конфигурация страницы (Узкие поля 1 см)
    for section in doc.sections:
        section.top_margin = Inches(0.39)
        section.bottom_margin = Inches(0.39)
        section.left_margin = Inches(0.39)
        section.right_margin = Inches(0.39)
        
        # Верхний колонтитул
        hp = section.header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run(f"{product_name_ru} | Паспорт безопасности химической продукции")
        hrun.font.name = 'Arial'
        hrun.font.size = Pt(8.5)
        hrun.font.italic = True
        hrun.font.color.rgb = RGBColor(128, 128, 128)
        
        # Нижний колонтитул
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
st.caption("Высокоструктурированный перевод паспортов химической безопасности с фиксацией номенклатуры.")

st.divider()

st.sidebar.header("🔑 Доступ к Yandex AI Studio")
folder_id = st.sidebar.text_input("Yandex Folder ID", type="password")
api_key = st.sidebar.text_input("Yandex API Key", type="password")

st.sidebar.markdown("---")
st.sidebar.header("📦 Химическая номенклатура")
# Изменили поле по твоему запросу
product_name_ru = st.sidebar.text_input("Официальное название продукта (RU):", value="ТРИМЕТИЛОЛПРОПАН")

col1, col2 = st.columns(2)
source_text = ""
file_name_output = "MSDS_RU_Translated"

with col1:
    st.subheader("Исходный документ (EN)")
    input_method = st.radio("Способ загрузки:", ("Загрузить файл (DOCX / PDF / TXT)", "Вставить текст вручную"))
    
    if input_method == "Вставить текст вручную":
        source_text = st.text_area("Вставьте текст MSDS на английском языке:", height=450, placeholder="SECTION 1: Identification...")
    else:
        uploaded_file = st.file_uploader("Выберите файл", type=["docx", "pdf", "txt"])
        if uploaded_file is not None:
            file_name_output = f"Translated_{uploaded_file.name}"
            
            if uploaded_file.name.endswith(".txt"):
                source_text = str(uploaded_file.read(), "utf-8")
            elif uploaded_file.name.endswith(".docx"):
                doc = Document(io.BytesIO(uploaded_file.read()))
                source_text = "\n".join([para.text for para in doc.paragraphs])
            elif uploaded_file.name.endswith(".pdf"):
                with pdfplumber.open(io.BytesIO(uploaded_file.read())) as pdf:
                    source_text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
            
            st.text_area("Предпросмотр оригинального текста:", value=source_text, height=300, disabled=True)

if st.button("🔄 Выполнить интеллектуальный перевод и форматирование", type="primary", use_container_width=True):
    if not folder_id or not api_key:
        st.warning("Пожалуйста, введите Yandex Folder ID и API Key.")
    elif source_text.strip():
        with col2:
            st.subheader("Осмысленный перевод (Markdown Preview)")
            
            # Передаем зафиксированное название в функцию перевода
            translated_result = translate_msds_with_studio(source_text, folder_id, api_key, product_name_ru)
            
            st.markdown(translated_result)
            
            if "Ошибка" not in translated_result:
                # Генерируем Word с тем же зафиксированным названием
                docx_data = make_formatted_docx(translated_result, product_name_ru)
                
                st.download_button(
                    label="💾 Скачать отформатированный файл WORD (.docx)",
                    data=docx_data,
                    file_name=file_name_output if file_name_output.endswith(".docx") else f"{file_name_output}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
    else:
        st.warning("Пожалуйста, добавьте текст или загрузите файл перед переводом.")
