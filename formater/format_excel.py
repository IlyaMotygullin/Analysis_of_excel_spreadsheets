import pandas as pd
from reader.reader_excel import ReaderExcle

class FormatExcel:
    
    def __init__(self, data_frame: pd.DataFrame, reader: ReaderExcle):
        self.data_frame = data_frame
        self.reader = reader
    
    def format_diapason_excel(self) -> None:
        self.data_frame = self.reader.get_data_list() # -> получения диапазона на основе выбора пользователя 
        return None
    
    