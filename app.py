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

# Второй блок: таблица с файла из Каталог_Чинт.xlsx с чекбоксами
st.subheader("Таблица товаров")

# Фиксированные переменные для раздела
categories = ["Корпус", "Отсек высоковольтного выключателя", "Отсек РЗА", "Прочее"]

# Инициализация таблицы-шаблона с фиксированными категориями и подкатегорией "Оборудование"
template_df = initialize_template_table(categories, ["Оборудование"] * len(categories))

# Добавление поля для поиска и фильтров на одной строке
with st.container():
    col1, col2, col3 = st.columns([3, 1, 1])
    with col1:
        search_query = st.text_input("Поиск товаров", "")
    with col2:
        selected_header = st.selectbox("Раздел", categories)  # Используем фиксированные переменные
    with col3:
        selected_subheader = st.selectbox("Подраздел", ["Оборудование"])  # Один вариант для подраздела

# Поиск по таблице. Если ничего не введено, отображается пустая таблица
if search_query.strip():
    filtered_df2 = df2[df2.apply(lambda row: row.astype(str).str.contains(search_query, case=False).any(), axis=1)]
else:
    filtered_df2 = pd.DataFrame(columns=df2.columns)  # Пустая таблица

# Отображение чекбоксов для выбора товаров
selected_items = []
for index, row in filtered_df2.iterrows():
    # Визуализируем каждую строку с чекбоксом
    is_selected = st.checkbox(f"{row['Наименование']} - {row.get('Тариф с НДС, руб', 'Цена не указана')}", key=f"item_{index}")
    if is_selected:
        selected_items.append(row["Наименование"])

# Перемещение выбранных товаров в шаблон
if selected_items:
    template_df.at[categories.index(selected_header), "Выбранные товары"].extend(selected_items)

# Третий блок: таблица-шаблон для выбранных товаров на всю ширину
st.subheader("Выбранные товары по категориям")

# Добавляем заголовки перед товарами
output_df = pd.DataFrame({
    "Заголовки": template_df["Заголовки"],  # Столбец заголовков
    "Выбранные товары": template_df["Выбранные товары"]
})

# Удаляем строки, где нет выбранных товаров, чтобы не показывать пустые строки
output_df = output_df[output_df["Выбранные товары"].apply(len) > 0]

# Разворачиваем список выбранных товаров в каждой строке на несколько строк
expanded_output_df = output_df.explode("Выбранные товары")

# Отображаем таблицу с заголовками и выбранными товарами
st.dataframe(expanded_output_df, use_container_width=True, hide_index=True)

# Четвертый блок: Сохранение в Excel
if st.button("Сохранить в Excel"):
    # Маппинг выбранных данных в итоговую таблицу
    mapped_data = {}
    for index, row in template_df.iterrows():
        header = row["Заголовки"]
        selected_items = row["Выбранные товары"]
        if selected_items:
            mapped_data[header] = [{"Наименование": item} for item in selected_items]

    # Сохранение в Excel файл
    save_to_excel(mapped_data, 'mapped_data.xlsx')
    st.success("Файл сохранен как mapped_data.xlsx!")
