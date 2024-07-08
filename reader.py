import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
import numpy as np

def choose_file():
    Tk().withdraw()  # Мы не хотим видеть основное окно Tk.
    filename = askopenfilename()  # Показывает диалоговое окно и возвращает выбранный путь к файлу
    print(f"Выбранный файл: {filename}")
    return filename

def process_and_save(filename):
    if filename:
        df = pd.read_excel(filename)
        
        # Преобразование всего DataFrame в строки
        df_str = df.astype(str)
        
        символы = input("Введите искомые символы для фильтрации: ").strip()
        
        # Создание маски для фильтрации
        mask = np.column_stack([df_str[col].str.contains(символы, na=False, case=False) for col in df_str])
        filtered_df = df.loc[mask.any(axis=1)]

        new_file = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if new_file:
            filtered_df.to_excel(new_file, index=False)
            print(f"Данные сохранены в {new_file}")
            return new_file
        else:
            print("Сохранение отменено.")
            return None
    else:
        print("Файл не был выбран.")
        return None

файл = choose_file()
if файл:  # Убедитесь, что файл был выбран перед продолжением
    newFile_excel = process_and_save(файл)
