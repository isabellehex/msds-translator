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

def translate_msds_with_studio(full_text: str, folder_id: str, api_key: str, product_name_ru: str) -> str:
    """
    Разбивает любой текст MSDS на аккуратные смысловые куски по объему,
    переводит их по очереди и склеивает обратно.
    """
    # Разбиваем весь текст на отдельные строчки
    lines = full_text.split('\n')
    
    chunks = []
    current_chunk = []
    current_length = 0
    
    # Задаем лимит: примерно 3000 символов на один запрос (это безопасно для вывода)
    max_chunk_size = 3000 
    
    for line in lines:
        current_chunk.append(line)
        current_length += len(line) + 1  # +1 для символа переноса строки
        
        # Если набрали критическую массу — сохраняем кусок и начинаем новый
        if current_length > max_chunk_size:
            chunks.append("\n".join(current_chunk))
            current_chunk = []
            current_length = 0
            
    # Не забываем забрать остаток текста, если он есть
    if current_chunk:
        chunks.append("\n".join(current_chunk))
        
    translated_parts = []
    total_chunks = len(chunks)
    
    st.info(f"📋 Документ успешно разделен на {total_chunks} частей для гарантированного перевода.")
    
    # Переводим каждый кусок по очереди
    for i, chunk in enumerate(chunks, 1):
        st.write(f"⏳ Переводим часть {i} из {total_chunks}...")
        
        prompt = f"""Ты — профессиональный химик-технолог и переводчик. 
Переведи предоставленный фрагмент паспорта безопасности (MSDS) вещества {product_name_ru} на русский язык.
Переводи строго, сохраняй структуру, числовые данные, таблицы, аббревиатуры и оригинальные CAS-номера.
Не сокращай текст, не убирай технические данные и не пиши никаких вступлений от себя — только чистый перевод текста.

ФРАГМЕНТ ДЛЯ ПЕРЕВОДА:
{chunk}"""

        try:
            # Твой стандартный запрос к API (замени на свой рабочий вызов, если он отличается)
            response = openai.chat.completions.create(
                model="yandexgpt/latest", 
                messages=[{"role": "user", "content": prompt}]
            )
            chunk_translation = response.choices[0].message.content
            translated_parts.append(chunk_translation)
        except Exception as e:
            st.error(f"Ошибка при переводе части {i}: {e}")
            translated_parts.append(f"\n[Ошибка перевода части {i}]\n")
            
    # Соединяем все части в единый итоговый текст
    full_translation = "\n\n".join(translated_parts)
    return full_translation

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
            file_name_output = f"{uploaded_file.name}_RU"
            
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
