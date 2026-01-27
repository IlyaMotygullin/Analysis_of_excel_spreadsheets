import pandas as pd
from datetime import datetime as date
import re
import numpy as num

# чтение excel-файла
def analysisi_excel_file(path_file: str) -> list:
    return pd.ExcelFile(path_file).sheet_names


def search_date(text: str) -> bool:
    return re.search(r'\b\d{2}\.\d{4}\b', text)

# TODO: реализовать аналогичный алгоритм для получения данных из листов где нет четкой даты
# TODO: форматирование данных


def format_date(date_list: list) -> list: # -> функция только для вывода отформтированных данных(дата)
    list_date = []
    
    for element_list in date_list:
        if isinstance(element_list, date):
            list_date.append(date.strftime(element_list, '%d.%m.%Y'))
        else:
            str_element = str(element_list)
            match_date = re.search(r'\d{2}.\d{2}.\d{4}', str_element)
            list_date.append(date.strftime(date.strptime(match_date.group(), '%d.%m.%Y').date(), '%d.%m.%Y'))
    
    return list_date 

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

dict_date_coordinate = get_date_excel(path=path_file, sheet_name=user_opinion)

for key, value in dict_date_coordinate.items():
    print(f'{key} - {value}')

# получение отфоратированных дат из excel
def get_date_list(dict_date: dict) -> list:
    return format_date(dict_date.keys())

list_date = get_date_list(dict_date_coordinate) # -> список отформатированных дат

# функция для проверки типа(не несет никакого смысла)
def print_type_element(list_date: list) -> None:
    for element in list_date:
        print(type(element))
    return None

print()

for element in list_date:
    print(f'{element} -> {type(element)}')
    
def choose_user_date(date_user_one: str, date_user_two: str) -> list:
    list_user_date = []
    list_user_date.append(date_user_one)
    list_user_date.append(date_user_two)
    return list_user_date

print('Введите дату 1: ')
user_date_one = input()
print('Введите  дату 2:')
user_date_two = input()

# -> функция для проверки принадлежности всех элементов к одному типу данных
def check_type_element(list_data: list) -> bool:
    result_check = all(type(x) == type(list_data[0]) and isinstance(x, date) for x in list_data)
    return result_check

def get_value_diapason(path_file: str, sheet_name: str, dict_date: dict, list_date_user: list) -> None:
    date_list_excel = list(dict_date.keys())
    
    if check_type_element(date_list_excel):
        date_user_one = date.strptime(list_date_user[0], '%d.%m.%Y')
        date_user_two = date.strptime(list_date_user[1], '%d.%m.%Y')
        
        list_date_user = []
        list_date_user.append(str(date_user_one))
        list_date_user.append(str(date_user_two))
        
        create_diaposon(path_file, sheet_name, dict_date, list_date_user)
    else:
        find_user_date1 = ''
        find_user_date2 = ''
        
        for key in dict_date:
            if list_date_user[0] in key:
                find_user_date1 = list_date_user[0]
                break
             
        for key in dict_date:
            if list_date_user[1] in key:
                find_user_date2 = list_date_user[1]
                break
                
        
        list_user_date = []
        
        print(find_user_date1)
        print(find_user_date2)
        
        create_diaposon(path_file, sheet_name, dict_date, list_user_date)

list_user_dates = choose_user_date(user_date_one, user_date_two)
get_value_diapason(path_file, user_opinion, dict_date_coordinate, list_date_user=list_user_dates)

# list = ['2021-11-01 00:00:00', '2021-11-05 00:00:00']
# create_diaposon(path_file, user_opinion, dict_date_coordinate, list)