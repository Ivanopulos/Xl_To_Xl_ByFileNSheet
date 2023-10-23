import os
import pandas as pd

# Укажите путь к папке с файлами Excel
path_to_folder = 'C:\\Users\\IMatveev\\Desktop\\НВИ\\голосовалка\\для базы'

# Список для хранения всех данных из всех файлов
all_data = []

# Обходим все файлы и подпапки в указанной папке
for subdir, _, files in os.walk(path_to_folder):
    for file in files:
        # Проверяем расширение файла
        if file.endswith('.xlsx') or file.endswith('.xlsm'):
            full_path = os.path.join(subdir, file)
            # Читаем все листы из файла
            xls = pd.ExcelFile(full_path)
            for sheet_name in xls.sheet_names:
                df = xls.parse(sheet_name)
                df['Имя файла'] = full_path
                df['Имя листа'] = sheet_name
                all_data.append(df)

# Объединяем все данные в один DataFrame
final_df = pd.concat(all_data, ignore_index=True)

# Сохраняем в новый файл Excel
final_df.to_excel('combined_data.xlsx', index=False, engine='openpyxl')