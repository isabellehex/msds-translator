import streamlit as st
import openai
from openai import OpenAI
import pdfplumber
import io
import re
import time
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 1. Настройка страницы интерфейса
st.set_page_config(
    page_title="MSDS Yandex AI Studio Pro",
    page_icon="🧪",
    layout="wide"
)

# Красивый корпоративный заголовок
st.title("🧪 MSDS Yandex AI Studio Pro")
st.caption("Профессиональный облачный перевод паспортов безопасности химической продукции")

# 2. Боковая панель для ввода настроек и ключей
st.sidebar.header("🔑 Настройки авторизации")
folder_id = st.sidebar.text_input("Yandex Folder ID", type="password", help="Введите идентификатор вашего каталога в Yandex Cloud")
api_key = st.sidebar.text_input("Yandex API Key", type="password", help="Введите ваш секретный API-ключ")

st.sidebar.markdown("---")
st.sidebar.header("📝 Параметры перевода")
product_name_ru = st.sidebar.text_input("Название продукта на русском", value="Триметилолпропан", help="Как называть вещество в итоговом переводе")

# 3. Функция умного перевода по частям (решает проблему пропуска разделов)
def translate_msds_by_chunks(full_text: str, folder_id: str, api_key: str, product_name_ru: str) -> str:
    import requests
    from requests.adapters import HTTPAdapter
    from urllib3.util import Retry
    
    # Разбиваем текст на строки
    lines = full_text.split('\n')
    chunks = []
    current_chunk = []
    current_length = 0
    max_chunk_size = 3000 # Немного уменьшим размер куска для надежности
    
    for line in lines:
        current_chunk.append(line)
        current_length += len(line) + 1
        if current_length > max_chunk_size:
            chunks.append("\n".join(current_chunk))
            current_chunk = []
            current_length = 0
    if current_chunk:
        chunks.append("\n".join(current_chunk))
        
    translated_parts = []
    total_chunks = len(chunks)
    
    st.info(f"📋 Документ успешно разделен на {total_chunks} части для гарантированного перевода.")
    
    # --- НАСТРОЙКА БРОНЕБОЙНОЙ СЕТЕВОЙ СЕССИИ ---
    session = requests.Session()
    # Если соединение оборвется, код автоматически повторит попытку до 5 раз с увеличивающейся паузой
    retries = Retry(total=5, backoff_factor=1, status_forcelist=[500, 502, 503, 504])
    session.mount('https://', HTTPAdapter(max_retries=retries))
    
    url = "https://llm.api.cloud.yandex.net/foundationModels/v1/completion"
    
    headers = {
        "Authorization": f"Api-Key {api_key}",
        "x-folder-id": folder_id,
        "Content-Type": "application/json"
    }
    # --------------------------------------------
    
    for i, chunk in enumerate(chunks, 1):
        st.write(f"⏳ Переводим часть {i} из {total_chunks}...")
        
        prompt = f"""Ты — профессиональный химик-технолог, эксперт по техническому регулированию и переводчик. 
Переведи предоставленный фрагмент паспорта безопасности (MSDS) вещества {product_name_ru} на русский язык.
Переводи строго, сохраняй структуру, числовые данные, таблицы, аббревиатуры и оригинальные CAS-номера.
Не сокращай текст, не убирай технические данные, показатели и не пиши никаких вступлений от себя — выдай только чистый перевод текста фрагмента.

ФРАГМЕНТ ДЛЯ ПЕРЕВОДА:
{chunk}"""

        # Структура запроса напрямую к Yandex GPT
        data = {
            "modelUri": f"gpt://{folder_id}/yandexgpt/latest",
            "completionOptions": {
                "stream": False,
                "temperature": 0.1, # Низкая температура для максимальной строгости перевода
                "maxTokens": "4000"
            },
            "messages": [
                {
                    "role": "user",
                    "text": prompt
                }
            ]
        }

        try:
            # Отправляем прямой POST-запрос с таймаутом ожидания в 60 секунд
            response = session.post(url, headers=headers, json=data, timeout=60)
            
            # Если Яндекс ответил ошибкой (например, 401 или 403) — сработает исключение
            response.raise_for_status()
            
            # Вытаскиваем текст перевода из ответа Яндекса
            result_json = response.json()
            chunk_translation = result_json["result"]["alternatives"][0]["message"]["text"]
            
            translated_parts.append(chunk_translation)
            
            # Обязательная пауза между частями
            time.sleep(2.0)
            
        except Exception as e:
            st.error(f"Ошибка при переводе части {i}: {e}")
            translated_parts.append(f"\n[Ошибка перевода части {i}. Технические детали: {str(e)}]\n")
            
    return "\n\n".join(translated_parts)

# 4. Функция генерации красивого Word-документа (.docx)
def create_word_document(translated_text: str) -> io.BytesIO:
    doc = Document()
    
    # Настройки стилей страницы (шрифт Arial, размер 11)
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(11)
    
    # Читаем текст построчно для красивой верстки заголовков
    paragraphs = translated_text.split('\n')
    
    for p_text in paragraphs:
        p_text = p_text.strip()
        if not p_text:
            continue
            
        p = doc.add_paragraph()
        
        # Если строка выглядит как заголовок раздела (например, РАЗДЕЛ, SECTION или Номер раздела)
        if re.match(r'^(SECTION|РАЗДЕЛ|\d+\.\s+)', p_text, re.IGNORECASE):
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(6)
            run = p.add_run(p_text)
            run.bold = True
            run.font.size = Pt(13)
            run.font.color.rgb = RGBColor(0, 102, 51) # Красивый строгий темно-зеленый цвет для структуры
        else:
            # Обычный текст
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.line_spacing = 1.15
            p.add_run(p_text)
            
    # Сохраняем файл в виртуальную память, чтобы выдать пользователю
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# 5. Главный экран приложения
# ТЕПЕРЬ ПРИНИМАЕМ И PDF, И DOCX!
uploaded_file = st.file_uploader(
    "Шаг 1: Загрузите оригинальный паспорт безопасности (PDF или DOCX)", 
    type=["pdf", "docx"]
)

if uploaded_file is not None:
    st.success("Файл успешно загружен!")
    
    # Кнопка запуска процесса
    if st.button("Шаг 2: Запустить перевод", type="primary"):
        if not folder_id or not api_key:
            st.error("❌ Пожалуйста, заполните Yandex Folder ID и API Key в левой боковой панели!")
        else:
            full_text = ""
            
            # --- УМНЫЙ ОБРАБОТЧИК ФОРМАТОВ ---
            if uploaded_file.name.endswith('.pdf'):
                with st.spinner("Считываем текст из PDF..."):
                    try:
                        with pdfplumber.open(uploaded_file) as pdf:
                            for page in pdf.pages:
                                text = page.extract_text()
                                if text:
                                    full_text += text + "\n"
                    except Exception as e:
                        st.error(f"Не удалось прочитать PDF-файл: {e}")
                        full_text = None
                        
            elif uploaded_file.name.endswith('.docx'):
                with st.spinner("Считываем текст из Word (.docx)..."):
                    try:
                        # Читаем вордовский файл построчно
                        doc_in = Document(uploaded_file)
                        for paragraph in doc_in.paragraphs:
                            if paragraph.text.strip():
                                full_text += paragraph.text + "\n"
                    except Exception as e:
                        st.error(f"Не удалось прочитать DOCX-файл: {e}")
                        full_text = None
            # ---------------------------------
            
            if full_text and full_text.strip():
                # Запускаем наш разделенный перевод (он работает одинаково для любого текста)
                translated_result = translate_msds_by_chunks(full_text, folder_id, api_key, product_name_ru)
                
                st.markdown("---")
                st.success("🎉 Перевод успешно завершен!")
                
                # Показываем превью перевода на экране
                with st.expander("👀 Посмотреть превью перевода прямо на сайте"):
                    st.markdown(translated_result)
                
                # Создаем Word файл
                with st.spinner("Формируем документ Word..."):
                    word_file = create_word_document(translated_result)
                
                # Кнопка скачивания готового документа
                st.download_button(
                    label="📥 Скачать готовый перевод (.docx)",
                    data=word_file,
                    file_name=f"MSDS_{product_name_ru}_RU.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
