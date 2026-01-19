import pandas as pd
from datetime import datetime as date

# проверка на уникальность колонок -> дата
def analysisi_excel_file(path_file: str) -> list:
    return pd.ExcelFile(path_file).sheet_names


# получение данных с листа excel -> выбор пользователя
def get_data_sheet_list(path_file: str, sheet_name: str) -> None:
    df = pd.read_excel(path_file, sheet_name=sheet_name, header=None) # читаем документ как есть, без заголовков
    index_row_sheet = 0
    
    while index_row_sheet < len(df):
        for element_row in  df.iloc[index_row_sheet].values:
            if isinstance(element_row, date):
                print(element_row)
        index_row_sheet += 1
        

path_file = "C:\\Users\\ilyam\\Desktop\\Проекты\\python_проекты\\read_excel_table\\Пр-во+экспрессы ноябр-дек21г.xlsx"
list_sheet_excel = analysisi_excel_file(path_file=path_file)

for index in list_sheet_excel:
    print(index)

user_opinion = input()

get_data_sheet_list(path_file=path_file, sheet_name=user_opinion)