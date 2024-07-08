import pandas as pd
import asyncio
from concurrent.futures import ThreadPoolExecutor
import datetime

# Глобальный пул потоков
executor = ThreadPoolExecutor()

async def read_excel_async(file_path):
    print(f"Чтение файла {file_path}...")
    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(executor, pd.read_excel, file_path)

async def filter_and_save_excel(filename, symbols, save_path):
    df = await read_excel_async(filename)
    print("Фильтрация данных...")
    mask = df.apply(lambda x: x.astype(str).str.contains(symbols).any(), axis=1)
    filtered_df = df[mask]

    if not filtered_df.empty:
        print(f"Сохранение отфильтрованных данных в {save_path}...")
        filtered_df.to_excel(save_path, index=False)
        print("Данные успешно сохранены.")
    else:
        print("Не найдено строк, удовлетворяющих критериям.")

async def main():
    # Предполагается, что file_path и save_path уже известны или заданы заранее
    file_path = input("Введите полный путь к файлу для чтения: ").strip()
    save_path = f"C:/Users/naver/Downloads/filtered_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    if file_path:
        symbols = input("Введите искомые символы для фильтрации: ").strip()
        await filter_and_save_excel(file_path, symbols, save_path)
    else:
        print("Путь к файлу не задан.")

if __name__ == "__main__":
    asyncio.run(main())
