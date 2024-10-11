import pandas as pd

# Пути к файлу
old_file_path = r"C:\Users\kruts\PycharmProjects\Exl\exlfile\old.xlsx"

# Чтение файла
old_data = pd.read_excel(old_file_path)

# Вывод заголовков столбцов
print("Заголовки столбцов в файле old.xlsx:", old_data.columns)
