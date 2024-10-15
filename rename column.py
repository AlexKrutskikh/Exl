import os
import pandas as pd

# Путь к исходному файлу
old_file_path = r"C:\Users\kruts\PycharmProjects\Exl\exlfile\old.xlsx"
# Путь к новой папке, где будут обрабатываться файлы
new_folder_path = r"C:\Users\kruts\PycharmProjects\Exl\exlfile\new"

# Чтение исходного файла
df_old = pd.read_excel(old_file_path)

# Список всех столбцов с суффиксом "_new"
new_columns = [col for col in df_old.columns if col.endswith('_new')]

# Проход по каждому файлу в новой папке
for file_name in os.listdir(new_folder_path):
    # Проверка, что файл является Excel файлом
    if file_name.endswith('.xlsx'):
        file_path = os.path.join(new_folder_path, file_name)

        # Чтение файла Excel в новой папке
        df_new = pd.read_excel(file_path)

        # Поочередное вставление значений в первую строку
        for index, column in enumerate(new_columns):
            if index < len(df_old):
                # Вставляем значение из df_old в первую строку df_new поочередно
                df_new.loc[0, df_new.columns[index]] = df_old[column][0]

        # Сохранение обновленного файла
        df_new.to_excel(file_path, index=False)

print(f"Все файлы в папке '{new_folder_path}' обновлены.")
