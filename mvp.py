import pandas as pd
from datetime import datetime as date
import re

# чтение excel-файла
def analysisi_excel_file(path_file: str) -> list:
    return pd.ExcelFile(path_file).sheet_names


def search_date(text: str) -> bool:
    return re.search(r'\b\d{2}\.\d{4}\b', text)
    
        
# проверка на дату в excel
def get_data_sheet_list(path_file: str, sheet_name: str) -> list:
    list_id_record = []
    df = pd.read_excel(path_file, sheet_name=sheet_name, header=None) # читаем документ как есть, без заголовков
    index_row_sheet = 0
    
    while index_row_sheet < len(df):
        for element_row in  df.iloc[index_row_sheet].values:
            if isinstance(element_row, date):
                list_id_record.append(element_row)
            else:
                if search_date(str(element_row)):
                    list_id_record.append(element_row)        
                
        index_row_sheet += 1
    
    return list_id_record

# TODO: научиться получать данные с таблицы по датам(диапазон в плоскости Х)
# TODO: получить значения под колонками с датами
# TODO: разобраться почему при чтении excel файла я получаю дубликаты


def format_date(date_list: list) -> None: # -> функция только для вывода отформтированных данных(дата)
    for element_list in date_list:
        if isinstance(element_list, date):
            print(date.strftime(element_list, '%d.%m.%Y'))
        else:
            str_element = str(element_list)
            match_date = re.search(r'\d{2}.\d{2}.\d{4}', str_element)
            print(date.strftime(date.strptime(match_date.group(), '%d.%m.%Y').date(), '%d.%m.%Y'))
    
                    
    return None

# получение данных
def get_column_value(path: str, sheet: str, list_date: list) -> None:
    df = pd.read_excel(path, sheet_name=sheet, header=None)
    
    # print(df.iloc[0, 0]) # -> df.iloc[x, y]
    format_date(list_date) # -> вывод отформатированных данных для выборки
    
    return None


path_file = "C:\\Users\\ilyam\\Desktop\\Пр-во+экспрессы ноябр-дек21г.xlsx"
list_sheet_excel = analysisi_excel_file(path_file=path_file)

for index in list_sheet_excel:
    print(index)

user_opinion = input()

list_record = get_data_sheet_list(path_file=path_file, sheet_name=user_opinion)

format_date(date_list=list_record)

get_column_value(path=path_file, sheet=user_opinion, list_date=list_record)