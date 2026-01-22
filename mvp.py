import pandas as pd
from datetime import datetime as date
import re

# чтение excel-файла
def analysisi_excel_file(path_file: str) -> list:
    return pd.ExcelFile(path_file).sheet_names


def search_date(text: str) -> bool:
    return re.search(r'\b\d{2}\.\d{4}\b', text)

# TODO: научиться получать данные с таблицы по датам(диапазон в плоскости Х)
# TODO: получить значения под колонками с датами


def format_date(date_list: list) -> None: # -> функция только для вывода отформтированных данных(дата)
    for element_list in date_list:
        if isinstance(element_list, date):
            print(date.strftime(element_list, '%d.%m.%Y'))
        else:
            str_element = str(element_list)
            match_date = re.search(r'\d{2}.\d{2}.\d{4}', str_element)
            print(date.strftime(date.strptime(match_date.group(), '%d.%m.%Y').date(), '%d.%m.%Y'))
    
                    
    return None

# -> получение даты вместе с ее координатами
def get_date_excel(path: str, sheet_name: str) -> dict:
    df = pd.read_excel(path, sheet_name=sheet_name, header=None)
    date_database = {}
    
    max_size_row = df.shape[0]
    max_size_column = df.shape[1]
    
    count_iteration_column = int(max_size_column / max_size_row)
    count_operation = 0    
    delimetr_column_size = max_size_row
        
    while count_operation < count_iteration_column:
        for index_row in range(max_size_row):
            for index_column in range(delimetr_column_size):
                cell_value = df.iloc[index_row, index_column]
                if isinstance(cell_value, date):
                    if cell_value not in date_database:
                        date_database[cell_value] = (index_row, index_column)
                else:
                    string_data_cell = str(cell_value)
                    if search_date(string_data_cell):
                        date_database[string_data_cell] = (index_row, index_column) 
            
        delimetr_column_size += max_size_row    
        count_operation += 1
    
    return date_database


# -> создание диапазона(последний элемент не входит в диапазон)
def create_diaposon(path: str, sheet: str, database_for_date: dict, list_date_user: list) -> None:
    df = pd.read_excel(path, sheet_name=sheet, header=None)
    get_date = date.strptime(list_date_user[0], '%Y-%m-%d %H:%M:%S')
    get_date1 = date.strptime(list_date_user[1], '%Y-%m-%d %H:%M:%S')
    
    if get_date in database_for_date and get_date1 in database_for_date:
        print(database_for_date[get_date])
        print(database_for_date[get_date1])
        
        matrix_table = []
        max_count_row = df.shape[0]
        count_element_column = database_for_date[get_date1][1]
        index_row = 1
        
        while index_row < max_count_row:
            row = []
            for index_column in range(count_element_column):
                row.append(df.iloc[index_row, index_column])
            matrix_table.append(row)
            index_row += 1
            
        
        for i in range(len(matrix_table)):
            for j in range(len(matrix_table[i])):
                print(matrix_table[i][j], end =  ' ')
            print()
        
    else:
        print('Данные не содержаться в словаре')
        
    return None


# -> создание выборки данных
def create_selector_data(path: str, sheet: str, database_for_date: dict, list_date_user: list) -> None:
    df = pd.read_excel(path, sheet_name=sheet, header=None)
    
    
    return None


path_file = "C:\\Users\\ilyam\\Desktop\\Пр-во+экспрессы ноябр-дек21г.xlsx"
list_sheet_excel = analysisi_excel_file(path_file=path_file)

for index in list_sheet_excel:
    print(index)

user_opinion = input()

# format_date(date_list=list_record)

# get_column_value(path=path_file, sheet=user_opinion, list_date=list_record)

dict_date_coordinate = get_date_excel(path=path_file, sheet_name=user_opinion)

list = ['2021-11-01 00:00:00', '2021-11-05 00:00:00']
create_diaposon(path_file, user_opinion, dict_date_coordinate, list)