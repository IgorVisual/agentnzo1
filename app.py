import os
import pandas as pd
import streamlit as st
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
st.title("Агент NZO")

# Фиксированные переменные для раздела
categories = ["Корпус", "Отсек высоковольтного выключателя", "Отсек РЗА", "Прочее"]

# Инициализация таблицы-шаблона с фиксированными категориями и подкатегорией "Оборудование"
template_df = initialize_template_table(categories, ["Оборудование"] * len(categories))

# Используем session_state для сохранения выбранных товаров между разделами
if 'selected_items' not in st.session_state:
    st.session_state.selected_items = {category: [] for category in categories}

# Второй блок: таблица с файла из Каталог_Чинт.xlsx с чекбоксами
st.subheader("Поиск оборудования")

# Добавление поля для поиска и фильтра по разделам на одной строке
with st.container():
    col1, col2 = st.columns([3, 1])
    with col1:
        if 'search_query' not in st.session_state:
            st.session_state.search_query = ""
        search_query = st.text_input("Поиск товаров", st.session_state.search_query)
    with col2:
        selected_header = st.selectbox("Раздел", categories)  # Используем фиксированные переменные

# Очищаем строку поиска и результаты при переключении раздела
if 'last_selected_header' in st.session_state and st.session_state.last_selected_header != selected_header:
    # Перед переключением категории сохраняем выбор
    if 'current_selected_rows' in st.session_state:
        st.session_state.selected_items[st.session_state.last_selected_header].extend(
            st.session_state.current_selected_rows
        )

    # Очищаем выбранные товары и строку поиска для нового поиска
    st.session_state.current_selected_rows = []
    st.session_state.last_selected_header = selected_header
    st.session_state.search_query = ""
    search_query = ""  # Сброс поля поиска
else:
    st.session_state.last_selected_header = selected_header
    st.session_state.current_selected_rows = []

# Поиск по таблице (обновляется в реальном времени при изменении текста)
if search_query.strip():
    filtered_df2 = df2[df2.apply(lambda row: row.astype(str).str.contains(search_query, case=False).any(), axis=1)]
    show_table = True
else:
    filtered_df2 = pd.DataFrame(columns=df2.columns)  # Пустая таблица
    show_table = False

# Используем текстовое представление данных с чекбоксами
if show_table:
    st.write(f"Результаты поиска для раздела: {selected_header}")
    
    # Для каждой строки оборудования отображаем текст с ценой и чекбокс
    for idx, row in filtered_df2.iterrows():
        item_name = row["Наименование"]
        item_price = row.get("Тариф с НДС, руб", "Цена не указана")
        selected = st.checkbox(f"{item_name} — {item_price} руб", key=f"{item_name}_{selected_header}")
        
        # Сохраняем выбор в текущем разделе
        if selected and item_name not in st.session_state.current_selected_rows:
            st.session_state.current_selected_rows.append(item_name)

# Третий блок: таблица-шаблон для выбранных товаров на всю ширину
st.subheader("Итоговый файл для просчета")

# Отображаем таблицу с заголовками и выбранными товарами
selected_df = pd.DataFrame({
    "Заголовки": categories,
    "Выбранные товары": [", ".join(st.session_state.selected_items[header]) for header in categories]
})
st.dataframe(selected_df, use_container_width=True, hide_index=True)

# Четвертый блок: Сохранение в Excel
if st.button("Сохранить в Excel"):
    # Маппинг выбранных данных в итоговую таблицу
    mapped_data = {}
    for category, items in st.session_state.selected_items.items():
        if items:
            mapped_data[category] = [{"Наименование": item} for item in items]
    
    # Сохранение в файл
    output_filename = "Итоговый_файл.xlsx"
    save_to_excel(mapped_data, output_filename)
    st.success(f"Файл успешно сохранен как {output_filename}")
