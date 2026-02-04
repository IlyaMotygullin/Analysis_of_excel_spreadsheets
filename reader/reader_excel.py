import pandas as pd
from datetime import datetime as date
import re

class ReaderExcle:
    def __init__(self, path_file: str, data_frame: pd.DataFrame):
        self.path_file = path_file
        self.data_frame = data_frame
        
    def get_sheet_list(self) -> list:
        return pd.ExcelFile(self.path_file).sheet_names
    
    # -> данная функция проверяет строку на садержание даты
    def __check_date_in_string(self, date_string: str) -> bool:
        return re.search(r'\b\d{2}.\d{4}\b', date_string)
    
    # -> функция должна находить даты или даты в строках, а также находить их координаты
    # -> возвращаемое значение: dict - словарь из ключей(даты) и значений(координаты)
    def __find_date(self, data_frame: pd.DataFrame) -> dict:
        date_map = {}
        all_column = data_frame.shape[1] # -> все количество колонок в excel-файле
        all_row = data_frame.shape[0] # -> все количетсов строк в excel-файле
        
        for index_row in range(all_row):
            for index_column in range(all_column):
                if not pd.isna(data_frame.iloc[index_row, index_column]) and isinstance(data_frame.iloc[index_row, index_column], date):
                    date_map[data_frame.iloc[index_row, index_column]] = (index_row, index_column)
                
                if not pd.isna(data_frame.iloc[index_row, index_column]) and isinstance(data_frame.iloc[index_row, index_column], str):
                    if self.__check_date_in_string(data_frame.iloc[index_row, index_column]):
                        date_map[data_frame.iloc[index_row, index_column]] = (index_row, index_column)
            
        return date_map
    
    # -> функция для форматирования даты    
    def __format_date(self, element: date) -> str:
        return element.strftime("%d.%m.%Y")
        
    # -> функция, которая предназначена для пользователя(выводит все даты)
    def print_date(self) -> list:
        format_date_list = []
        
        date_map = self.__find_date(self.data_frame)        
        list_date = date_map.keys()
        
        for element in list_date:
            format_date_list.append(self.__format_date(element))
        
        return format_date_list
    
    # -> получение координат дат
    def __get_coordinate_date(self, date_map: dict, user_date_list: list) -> list:
        list_coordinate = []
        list_date = date_map.keys()
        
        if all(isinstance(x, date) for x in list_date):
            list_coordinate.append(date_map.get(date.strptime(user_date_list[0], '%d.%m.%Y')))
            list_coordinate.append(date_map.get(date.strptime(user_date_list[1], '%d.%m.%Y')))
            
            return list_coordinate
        
        for element in list_date:
            if user_date_list[0] in element:
                list_coordinate.append(date_map.get(element))  
                break 
        
        for element in list_date:
            if user_date_list[1] in element:
                list_coordinate.append(date_map.get(element))
                break
        
        return list_coordinate
    
    # -> функция должна создавать диапазоны между указанными значения
    def get_data_list(self, sheet_name: str, user_date_list: list) -> pd.DataFrame:
        self.data_frame = pd.read_excel(self.path_file, sheet_name=sheet_name, header=None)
        date_map = self.__find_date(self.data_frame) # -> return dict
        
        # -> зная, что во втором виде файла у нас в ключах содержаться строки, 
        # а в значениях, как и в первом случае у нас кортеж, можно составить алгоритм,
        # который будет делать проверку перед тем как создать диапазон
        
        # print(self.data_frame.iloc[1:self.data_frame.shape[0], 19:36]) # -> решение для диапазонов (возвращаемый тип: DataFrame)
        list_coordinate = self.__get_coordinate_date(date_map, user_date_list) # -> два кортежа, в каждом первый элемент совпадает
        
        print(self.data_frame.iloc[list_coordinate[0][0]:self.data_frame.shape[0], list_coordinate[0][1]:list_coordinate[1][1]])
        
        return self.data_frame.iloc[
            list_coordinate[0][0]:self.data_frame.shape[0],
            list_coordinate[0][1]:list_coordinate[1][1]
        ]    