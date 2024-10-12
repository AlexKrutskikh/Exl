import pandas as pd
import os

# Пути к папкам
new_folder_path = r"C:\Users\kruts\PycharmProjects\Exl\exlfile\new"
old_file_path = r"C:\Users\kruts\PycharmProjects\Exl\exlfile\old.xlsx"

# Чтение файла old
old_data = pd.read_excel(old_file_path)

# Удаление лишних пробелов в заголовках
old_data.columns = old_data.columns.str.strip()

# Приведение всех числовых столбцов в old_data к целым числам
for column in old_data.select_dtypes(include='number').columns:
    old_data[column] = pd.to_numeric(old_data[column], errors='coerce').fillna(0).astype(int)

# Статический список полей для обновления
fields_to_update = [
    'Node_name_new', 'IP_Address_new', 'ObjectSubType_new', 'Address_new', 'Interface_Name_new',
    'InterfaceID_new', 'Contact_Person_new', 'Support_Time_Start_new', 'support_time_End_new',
    'MailTo_new', 'ContactPerson_RG_new', 'Business_service_new', 'Operational_service_new',
    'Escalation_Error_Warning_new', 'Escalation_Error_Critical_new', 'Escalation_availability_new',
    'Escalation_Disposal_Warning_new', 'Escalation_Disposal_Critical_new'
]

# Получаем список всех файлов в папке new
new_files = [f for f in os.listdir(new_folder_path) if f.endswith('.xlsx')]

# Проверка наличия файлов в папке new
if not new_files:
    print("Нет файлов в папке new.")
else:
    # DataFrame для хранения строк без совпадений
    not_found_data = pd.DataFrame()

    # Цикл по всем файлам в папке new
    for new_file in new_files:
        new_file_path = os.path.join(new_folder_path, new_file)
        print(f"Обработка файла: {new_file}")

        # Чтение данных из текущего файла
        new_data = pd.read_excel(new_file_path)
        new_data.columns = new_data.columns.str.strip()  # Удаление лишних пробелов в заголовках

        # Приведение всех числовых столбцов в new_data к целым числам
        for column in new_data.select_dtypes(include='number').columns:
            new_data[column] = pd.to_numeric(new_data[column], errors='coerce').fillna(0).astype(int)

        # Цикл по всем значениям InterfaceID_new в текущем файле new
        for index, row in new_data.iterrows():
            interface_id_new = row['InterfaceID_new']

            # Поиск строки в old_data с совпадающим значением InterfaceID
            old_row_index = old_data[old_data['InterfaceID'] == interface_id_new].index

            # Вывод информации о найденных индексах
            if old_row_index.empty:
                print(f"Не найдено совпадение для InterfaceID: {interface_id_new}")
                # Добавляем строку без совпадений в DataFrame
                not_found_data = pd.concat([not_found_data, row.to_frame().T], ignore_index=True)
            else:
                print(f"Найдено совпадение для InterfaceID: {interface_id_new} на индексе {old_row_index[0]}")

                # Если совпадение найдено, обновляем поля в old_data
                for field in fields_to_update:
                    # Проверка наличия поля в new_data
                    if field in new_data.columns:
                        value_to_update = row[field]
                        if pd.notna(value_to_update):  # Проверка на наличие значения (не NaN)
                            # Приведение типа к str для обновления
                            old_data.loc[old_row_index[0], field] = str(value_to_update)
                            print(f"Обновлено поле: {field} для {interface_id_new} на {value_to_update}")

    # Сохранение изменений в файле old
    old_data.to_excel(r'C:\Users\kruts\PycharmProjects\Exl\exlfile\old_updated.xlsx', index=False)

    # Сохранение строк без совпадений в отдельный файл
    if not not_found_data.empty:
        not_found_data.to_excel(r'C:\Users\kruts\PycharmProjects\Exl\exlfile\not_found_interface_id.xlsx', index=False)
        print("Строки без совпадений сохранены в not_found_interface_id.xlsx")

    print("Обновление завершено.")
