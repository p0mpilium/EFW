import asyncio
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
from concurrent.futures import ThreadPoolExecutor

def choose_file():
    Tk().withdraw()
    filename = askopenfilename()
    print(f"Выбранный файл: {filename}")
    return filename

async def async_process_and_save(filename):
    symbols = input("Введите искомые символы для фильтрации: ").strip()

    loop = asyncio.get_running_loop()

    with ThreadPoolExecutor() as pool:
        df = await loop.run_in_executor(pool, pd.read_excel, filename, 0)
        mask = await loop.run_in_executor(pool, lambda: df.apply(lambda row: row.astype(str).str.contains(symbols).any(), axis=1))
        filtered_df = df[mask]

    if not filtered_df.empty:
        new_file = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if new_file:
            filtered_df.to_excel(new_file, index=False)
            print(f"Данные сохранены в {new_file}")
        else:
            print("Сохранение отменено.")
    else:
        print("Не найдено строк, удовлетворяющих критериям.")

def run_async_process_and_save(filename):
    asyncio.run(async_process_and_save(filename))

файл = choose_file()
if файл:
    run_async_process_and_save(файл)
