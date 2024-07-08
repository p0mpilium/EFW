import pandas as pd

def filter_and_clean_data(original_file_path, reference_file_path, output_file_path):
    """
    Фильтрует данные из оригинального файла Excel, используя структуру ссылочного файла,
    удаляет определённые слова из текста в ячейках, сохраняет отфильтрованные данные в новый файл.

    Args:
        original_file_path (str): Путь к оригинальному файлу данных.
        reference_file_path (str): Путь к файлу, на основе которого нужно отфильтровать данные.
        output_file_path (str): Путь для сохранения отфильтрованного файла.
    """
    # Загрузка данных из оригинального и ссылочного файлов
    original_df = pd.read_excel(original_file_path)
    reference_df = pd.read_excel(reference_file_path)

    # Определение индексов ненулевых строк в ссылочном файле
    non_null_indices = reference_df.dropna().index

    # Применение этих индексов к оригинальному файлу
    filtered_df = original_df.iloc[non_null_indices]

    # Удаление определённых слов из текста в ячейках
    keywords_to_remove = ["Раздел 1.", "Раздел 2.", "Раздел 3.", "Раздел 4."]  # Добавьте или удалите ключевые слова по необходимости
    for keyword in keywords_to_remove:
        filtered_df['Data'] = filtered_df['Data'].astype(str).str.replace(keyword, "", regex=False)

    # Удаление пустых строк из отфильтрованных данных
    filtered_cleaned_df = filtered_df.dropna()

    # Сохранение отфильтрованных и очищенных данных в новый файл
    filtered_cleaned_df.to_excel(output_file_path, index=False, engine='openpyxl')

# Пути к файлам
original_file_path = 'path_to_your_ОбработанныеДанные_20240329_201729.xlsx'
reference_file_path = 'path_to_your_modified_file.xlsx'
output_file_path = 'path_to_your_output_filtered_file.xlsx'

# Выполнение функции фильтрации и очистки
filter_and_clean_data(original_file_path, reference_file_path, output_file_path)
