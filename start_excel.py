from openpyxl import load_workbook


wb = load_workbook("D:\Работа\Авито парсер локальный сайт\Учет Уведомлений.xlsx")
ws = wb['Главный_лист']

def get_number():
    numbers_excel = ws['C']
    count = len(numbers_excel)
for col in ws.iter_cols(min_row=2, min_col=3, max_col=3, max_row=count, values_only=True):
    for cell in col:
        number = cell

