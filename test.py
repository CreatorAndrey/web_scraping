from openpyxl import load_workbook

list_excel = ['5', '3', '4', '1']
wb = load_workbook('D:\Работа\Авито парсер локальный сайт\Учет Уведомлений.xlsx')
ws = wb['Главный_лист']
numbers_excel = ws['C']
count = len(numbers_excel)
for col in ws.iter_cols(min_row=2, min_col=3, max_col=3, max_row=count):
    for cell in col:
        number = cell.value
        r = cell.row
        c = cell.column + 1
        for i in list_excel:
            ws.cell(row=r, column=c, value = i)
            c += 1
wb.save('Auto.xlsx')
wb.close()