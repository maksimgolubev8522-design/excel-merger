import pandas as pd
import os
from pathlib import Path


def merge_excel_reports_universal(input_folder, output_file):
    """
    Универсальный скрипт для объединения Excel-файлов.
    Работает с любыми названиями колонок.
    """

    # Получаем список всех Excel-файлов
    excel_files = list(Path(input_folder).glob('*.xlsx'))

    if not excel_files:
        print(f"❌ В папке {input_folder} не найдено Excel-файлов")
        return

    print(f"📁 Найдено файлов: {len(excel_files)}")

    # Читаем первый файл, чтобы узнать структуру данных
    first_file = excel_files[0]
    first_df = pd.read_excel(first_file)
    columns = list(first_df.columns)

    print(f"\n📋 Найдены колонки: {columns}")
    print(f"💡 Всего колонок: {len(columns)}")

    # Автоматически определяем ключевые колонки
    date_col = None
    product_col = None
    quantity_col = None
    price_col = None

    # Ищем колонку с датой
    for col in columns:
        if 'дата' in col.lower() or 'date' in col.lower():
            date_col = col
            break

    # Ищем колонку с товаром
    for col in columns:
        if any(word in col.lower() for word in ['товар', 'продукт', 'наименование', 'product', 'item']):
            product_col = col
            break

    # Ищем колонку с количеством
    for col in columns:
        if any(word in col.lower() for word in ['кол', 'количество', 'qty', 'quantity', 'шт']):
            quantity_col = col
            break

    # Ищем колонку с ценой
    for col in columns:
        if any(word in col.lower() for word in ['цена', 'стоимость', 'price', 'cost']):
            price_col = col
            break

    print(f"\n🔍 Автоматически определены колонки:")
    print(f"   📅 Дата: {date_col}")
    print(f"   📦 Товар: {product_col}")
    print(f"   🔢 Количество: {quantity_col}")
    print(f"   💰 Цена: {price_col}")

    # Читаем и объединяем все файлы
    all_data = []

    for file in excel_files:
        print(f"\n📄 Читаю файл: {file.name}")
        df = pd.read_excel(file)
        df['Источник'] = file.stem  # Добавляем имя файла как источник
        all_data.append(df)

    # Объединяем все данные
    merged_df = pd.concat(all_data, ignore_index=True)

    print(f"\n🔧 Очищаю данные...")

    # Базовая очистка (работает с любыми колонками)
    merged_df = merged_df.dropna(how='all')  # Удаляем полностью пустые строки
    merged_df = merged_df.drop_duplicates()  # Удаляем дубликаты

    # Преобразуем дату (если нашли колонку)
    if date_col:
        try:
            merged_df[date_col] = pd.to_datetime(merged_df[date_col], errors='coerce')
            merged_df = merged_df.sort_values(date_col)
        except:
            print(f"   ⚠️ Не удалось преобразовать колонку '{date_col}' в дату")

    # Преобразуем числовые колонки (если нашли их)
    if quantity_col:
        merged_df[quantity_col] = pd.to_numeric(merged_df[quantity_col], errors='coerce')

    if price_col:
        merged_df[price_col] = pd.to_numeric(merged_df[price_col], errors='coerce')

    # Создаем колонку "Сумма" (если есть количество и цена)
    sum_col = None
    if quantity_col and price_col:
        sum_col = 'Сумма'
        merged_df[sum_col] = merged_df[quantity_col] * merged_df[price_col]
        print(f"   ✅ Добавлена колонка '{sum_col}'")

    # Создаем сводную таблицу (если есть товар)
    summary = None
    if product_col:
        agg_dict = {}

        if quantity_col:
            agg_dict[quantity_col] = 'sum'

        if sum_col:
            agg_dict[sum_col] = 'sum'

        if agg_dict:
            summary = merged_df.groupby(product_col).agg(agg_dict).reset_index()

            # Переименовываем колонки для красоты
            rename_dict = {}
            if quantity_col:
                rename_dict[quantity_col] = f'{quantity_col} (итого)'
            if sum_col:
                rename_dict[sum_col] = f'{sum_col} (итого)'

            summary = summary.rename(columns=rename_dict)

            print(f"   ✅ Создана сводная таблица по колонке '{product_col}'")

    # Сохраняем результат
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Первый лист - все данные
        merged_df.to_excel(writer, sheet_name='Все данные', index=False)

        # Второй лист - сводная таблица (если удалось создать)
        if summary is not None:
            summary.to_excel(writer, sheet_name='Итоги', index=False)

    print(f"\n✅ Готово! Результат сохранен в: {output_file}")
    print(f"📊 Всего записей: {len(merged_df)}")

    if sum_col:
        total_sum = merged_df[sum_col].sum()
        print(f"💰 Общая сумма: {total_sum:,.0f} руб.")

    return merged_df, summary


# Запуск
if __name__ == "__main__":
    input_folder = "test_data"
    output_file = "merged_report_universal.xlsx"

    merge_excel_reports_universal(input_folder, output_file)