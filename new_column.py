import pandas as pd

# Путь к исходному файлу
old_file_path = r"C:\Users\kruts\PycharmProjects\Exl\exlfile\old.xlsx"
new_file_path = r"C:\Users\kruts\PycharmProjects\Exl\exlfile\old_updated.xlsx"

# Чтение файла Excel
df = pd.read_excel(old_file_path)

# Создание новых столбцов с префиксом new_
for column in df.columns:
    new_column_name = f"new_{column.replace(' ', '_')}"  # Новый столбец с подчеркиваниями
    df[new_column_name] = ''  # Создание пустого столбца

# Сохранение обновленного DataFrame в новый Excel файл
df.to_excel(new_file_path, index=False)

print(f"Файл сохранен по адресу: {new_file_path}")
