import os
import pandas as pd
import streamlit as st
import random
from openpyxl import Workbook

# Функция для инициализации таблицы-шаблона с заголовками и подзаголовками
def initialize_template_table(headers, subheaders):
    template_df = pd.DataFrame({
        "Заголовки": headers,
        "Подзаголовки": subheaders,
        "Выбранные товары": [[] for _ in range(len(headers))]  # Пустые списки для товаров
    })
    return template_df

# Функция для маппинга данных в итоговую таблицу и сохранения в Excel
def save_to_excel(mapped_data, output_filename):
    wb = Workbook()
    ws = wb.active

    # Записываем данные в Excel по группам
    row_num = 1

    for category, items in mapped_data.items():
        # Вставляем название категории
        ws[f'A{row_num}'] = category
        row_num += 1

        # Вставляем заголовки
        ws[f'A{row_num}'] = "Наименование"
        ws[f'B{row_num}'] = "Цена, руб"
        ws[f'C{row_num}'] = "Количество"
        row_num += 1

        # Вставляем данные
        for item in items:
            ws[f'A{row_num}'] = item["Наименование"]
            ws[f'B{row_num}'] = item.get("Цена", "")  # Цена может быть не указана
            ws[f'C{row_num}'] = item.get("Количество", "")  # Количество может быть не указано
            row_num += 1

        # Пустая строка между категориями
        row_num += 1

    # Сохранение файла
    wb.save(output_filename)
    print(f"Файл успешно сохранен как {output_filename}")

# Загрузка данных для второй таблицы (например, текущий файл Excel)
file_path = 'Каталог_Чинт.xlsx'  # Укажи путь к файлу Excel
df2 = pd.read_excel(file_path, header=14)  # Чтение с заголовками из 15-й строки

# Очищаем заголовки второй таблицы от лишних пробелов и спецсимволов
df2.columns = df2.columns.str.strip()
df2.columns = df2.columns.str.replace(r'[\n\r\t]', ' ', regex=True)
df2.columns = df2.columns.str.replace(r'\s+', ' ', regex=True)

# Фильтруем корректные столбцы
valid_columns = [col for col in df2.columns if not col.startswith("Unnamed")]
df2 = df2[valid_columns]

# Столбцы по умолчанию
default_columns = [col for col in ["Наименование", "Тариф с НДС, руб"] if col in valid_columns]

# Установка кастомных стилей для центрирования и уменьшения отступов
st.markdown(
    """
    <style>
    .block-container {
        padding-top: 1.7rem;
        padding-bottom: 0rem;
        padding-left: 1rem;
        padding-right: 1rem;
        text-align: center;  /* Центрирование текста */
    }
    .css-1p05t01 {
        padding: 0;
    }
    .st-dataframe {
        width: 100%;
    }
    .stSelectbox label, .stTextInput label {
        text-align: center;  /* Центрирование текста в полях ввода */
    }
    .stSelectbox div, .stTextInput div {
        margin-left: auto;
        margin-right: auto;
        text-align: center;
    }
    .stCheckbox div {
        margin-left: auto;
        margin-right: auto;
        text-align: center;  /* Центрирование чекбоксов */
    }
    </style>
    """, 
    unsafe_allow_html=True
)

# Интерфейс Streamlit
st.title("Табличный интерфейс")

# Первый блок: данные из файлов в папке clients
st.subheader("Данные клиентов")
client_data = []  # Заглушка для функции
if client_data:
    selected_client = st.selectbox("Выберите файл клиента", options=[f for f in os.listdir('clients') if f.endswith('.xlsx')])
    client_df = pd.read_excel(os.path.join('clients', selected_client))

    # Выбор необходимых столбцов для отображения в клиентских данных
    client_columns = ["Наименование", "Тариф с НДС, руб", "Выбрать"]
    if all(col in client_df.columns for col in client_columns):
        client_df_filtered = client_df[client_columns]
    else:
        client_df_filtered = client_df  # Показать все, если нужные столбцы отсутствуют

    st.dataframe(client_df_filtered, use_container_width=True, hide_index=True)  # Таблица на всю ширину контейнера
else:
    st.write("Нет данных для отображения")

# Второй блок: таблица с файла из Каталог_Чинт.xlsx с чекбоксами
st.subheader("Таблица товаров")

# Фиксированные переменные для раздела
categories = ["Корпус", "Отсек высоковольтного выключателя", "Отсек РЗА", "Прочее"]

# Инициализация таблицы-шаблона с фиксированными категориями и подкатегорией "Оборудование"
template_df = initialize_template_table(categories, ["Оборудование"] * len(categories))

# Добавление поля для поиска и фильтров на одной строке
with st.container():
    col1, col2, col3 = st.columns([3, 1, 1])
