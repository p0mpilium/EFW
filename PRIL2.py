import pandas as pd
import re

# Загрузка данных из файла
file_path = 'C:/Users/naver/Downloads/ОбработанныеДанные_20240329_201729_mod.xlsx'  # Замените на ваш путь к файлу
df = pd.read_excel(file_path)

# Функция для замены выражений
def replace_section(text):
    if pd.isnull(text):
        return text  # Возвращаем как есть, если текста нет
    return re.sub(r'Раздел \d+\.', '', text)  # Заменяем "Раздел X." на ничего

# Применяем функцию замены ко всем текстовым колонкам
for column in df.columns:
    df[column] = df[column].apply(replace_section)

# Сохраняем измененные данные обратно в XLSX файл
output_path = 'D:/mame/modified_file0.xlsx'  # Замените на ваш путь к выходному файлу
df.to_excel(output_path, index=False)

print(f'Изменения сохранены в файл: {output_path}')
