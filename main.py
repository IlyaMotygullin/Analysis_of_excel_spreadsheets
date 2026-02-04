from reader.reader_excel import ReaderExcle
from formater.format_excel import FormatExcel
import pandas as pd

path_file = "C:\\Users\\ilyam\\Desktop\\Пр-во+экспрессы ноябр-дек21г.xlsx"
data_frame = pd.read_excel(path_file)

# -> чтение excel файла
reader = ReaderExcle(path_file=path_file, data_frame=data_frame) 
# list_sheet = reader.get_sheet_list()

# for element in list_sheet:
#     print(element)
    

# print('Введите лист для получения данных: ')
# user_input = input()

# print('Введите даты, по которым Вы хотите составить диапазон:')
# list_date = reader.print_date()

# for element in list_date:
#     print(element)

# print('Введите первую дату: ')
# user_date_one = input()

# print('Введите вторую дату: ')
# user_date_two = input()

# user_date_list = []
# user_date_list.append(user_date_one)
# user_date_list.append(user_date_two)

# reader.get_data_list(user_input, user_date_list)

# -> форматирование извлеченной части excel
formater = FormatExcel()
formater.print_console()
