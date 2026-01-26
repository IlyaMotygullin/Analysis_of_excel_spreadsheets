import pandas as pd
from datetime import datetime, date
import re
import numpy as np
from typing import Tuple, List, Dict
from openpyxl import load_workbook


# чтение excel-файла
def get_excel_info(path_file: str):
    excel_obj = pd.ExcelFile(path_file)  # Создаем объект ExcelFile ОДИН РАЗ
    return excel_obj, excel_obj.sheet_names


def search_date(text: str) -> bool:
    return re.search(r'\b\d{2}\.\d{2}\.\d{4}\b', text) # ищет полную дату в таблице пример 21.01.2026


# TODO: научиться получать данные с таблицы по датам(диапазон в плоскости Х)
# TODO: получить значения под колонками с датами


def format_date(date_list: list) -> None:  # -> функция только для вывода отформтированных данных(дата)
    for element_list in date_list:
        if isinstance(element_list, date):
            print(date.strftime(element_list, '%d.%m.%Y'))
        else:
            str_element = str(element_list)
            match_date = re.search(r'\b\d{2}\.\d{2}\.\d{4}\b', str_element)
            print(date.strftime(date.strptime(match_date.group(), '%d.%m.%Y').date(), '%d.%m.%Y'))

    return None


# -> получение даты вместе с ее координатами
def get_date_excel(excel_obj: pd.ExcelFile, sheet_name: str) -> dict:
    df = excel_obj.parse(sheet_name=sheet_name, header=None)  # ВМЕСТО: df = pd.read_excel(path, sheet_name=sheet_name, header=None) -> ИСПОЛЬЗУЕМ: excel_obj.parse()
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
def create_diaposon(excel_obj: pd.ExcelFile, sheet: str, database_for_date: dict, list_date_user: list) -> None:
    df = excel_obj.parse(sheet_name=sheet, header=None)  # ВМЕСТО: df = pd.read_excel(path, sheet_name=sheet_name, header=None) -> ИСПОЛЬЗУЕМ: excel_obj.parse()

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
                print(matrix_table[i][j], end=' ')
            print()

    else:
        print('Данные не содержаться в словаре')

    return None


# выводи просто ключи
# get_date_excel()


# -> создание выборки данных
def create_selector_data(excel_obj: pd.ExcelFile, sheet: str, database_for_date: dict, list_date_user: list) -> None:
    df = excel_obj.parse(sheet_name=sheet, header=None)  # ВМЕСТО: df = pd.read_excel(path, sheet_name=sheet_name, header=None) -> ИСПОЛЬЗУЕМ: excel_obj.parse()

    get_date = date.strptime(list_date_user[0], '%Y-%m-%d %H:%M:%S')
    get_date1 = date.strptime(list_date_user[1], '%Y-%m-%d %H:%M:%S')

    return None


path_file = "C:\\Users\\user\\Desktop\\Картотека\\Работа\\Прог проекты\\1 проект\\Таблица\\Декабрь 21 МГОК.xlsx"
excel_file, list_sheet_excel = get_excel_info(path_file=path_file)


for index in list_sheet_excel:
    print(index)

user_opinion = input()

# format_date(date_list=list_record)

# get_column_value(path=path_file, sheet=user_opinion, list_date=list_record)

dict_date_coordinate = get_date_excel(excel_file, sheet_name=user_opinion)

date_list = dict_date_coordinate.keys()

date_mapping = {}
for key in dict_date_coordinate.keys():
    if isinstance(key, (datetime, pd.Timestamp)):
        normalized = datetime.combine(key.date(), datetime.min.time())
        date_str = normalized.strftime('%d.%m.%Y')
        date_mapping[date_str] = key
    else:
        try:
            dt_str = str(key).split()[0]  # Берем только дату
            dt = datetime.strptime(dt_str, '%Y-%m-%d')
            date_str = dt.strftime('%d.%m.%Y')
            date_mapping[date_str] = key
        except:
            continue

sorted_dates = sorted(date_mapping.keys(),
                      key=lambda x: datetime.strptime(x, '%d.%m.%Y'))

for i, date_str in enumerate(sorted_dates, 1):
    print(date_str)

# for key in date_list:
#     new_date = str(key)
#     new_date = new_date.replace('-', '.')
#
#     print(new_date)
#
# write_date_info_1 = input('Введите первую дату из списка: ')
# write_date_info_2 = input('Введите вторую дату из списка: ')
#
# found_dates = []
#
# found1 = None
# found2 = None
#
# for date_info in date_list:
#     new_date = date_info
#
#     date_str = str(new_date)
#
#     if date_str == write_date_info_1:
#         found1 = new_date
#     if date_str == write_date_info_2:
#         found2 = new_date
#
# if found1 is not None and found2 is not None:
#     print(f"Найдены даты: {found1} и {found2}")
#
#     if isinstance(found1, date):
#         # date1_str = found1.strftime('%d-%m-%Y', )
#         print(type(found1))
#     else:
#         date1_str = str(found1)
#
#     if isinstance(found2, date):
#         date2_str = found2.strftime('%Y-%m-%d %H:%M:%S')
#     else:
#         date2_str = str(found2)
#
#     list_for_create_diaposon = [date1_str, date2_str]
#
#     # Вызываем функцию с правильными параметрами
#     create_diaposon(excel_file, user_opinion, dict_date_coordinate, list_for_create_diaposon)
#
# list = date_list
# create_diaposon(excel_file, user_opinion, dict_date_coordinate, list)