import pandas as pd

# Пути к файлам
new_file_path = r'./exlfile/new.xlsx'
old_file_path = r'./exlfile/old.xlsx'



# Чтение файлов Excel
new_data = pd.read_excel(new_file_path)
old_data = pd.read_excel(old_file_path)

# Поля, которые нужно обновлять
fields_to_update = [
    'Node Name_new', 'IP_Address_new', 'ObjectSubType_new', 'Address_new', 'Volume_Name_new',
    'Disk_size_new', 'VolumeID_new', 'Polling_frequency_new', 'Threshold_Warning_new',
    'Escalation_Warning_new', 'Threshold_Critical_new', 'Escalation_Critical_new',
    'MailTo_new', 'Business_services_new', 'Contact_Person_new',
    'Support_time_beginning_new', 'Support_time_over_new', 'Contact _ERSON_RG_new'
]

# Цикл по всем значениям Node_Name_new в файле new
for index, row in new_data.iterrows():
    node_name_new = row['Node_Name_new']

    # Поиск строки в old_data с совпадающим значением Node_Name
    old_row_index = old_data[old_data['Node_Name'] == node_name_new].index

    if not old_row_index.empty:
        # Если совпадение найдено, обновляем поля в old_data
        for field in fields_to_update:
            if field in new_data.columns and field in old_data.columns:
                old_data.at[old_row_index, field.replace('_new', '')] = row[field]

# Сохранение изменений в файле old
old_data.to_excel('./old/old_updated.xlsx', index=False)

print("Обновление завершено.")
