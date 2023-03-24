# https://istories.media/workshops/2021/09/20/parsing-s-pomoshchyu-python-urok-2/
# https://questu.ru/articles/81673/

from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from time import sleep
from tkinter import *
from tkinter.messagebox import showerror, showinfo
from openpyxl import load_workbook

xpath_login = '//*[@id="login"]'
xpath_password = '//*[@id="password"]'
xpath_button = '/html/body/div[2]/div/section/div[2]/form/div[4]/input'
url_entry = 'https://roboparts.ru/'
s = Service("chromedriver.exe")
browser = Chrome(service=s)
browser.get(url_entry)
#D:\Работа\Авито парсер локальный сайт\Учет Уведомлений.xlsx

def get_number(folder):
    try:
        wb = load_workbook(folder)
    except:
        showerror('Ошибка', 'Не удается подключиться к книге Excel')
        return
    try:
        ws = wb['Главный_лист']
    except:
        showerror('Ошибка', 'Не найден лист')
        return
    try:
        br_login = browser.find_element(By.XPATH, xpath_login)
    except:
        showerror('Ошибка', 'Не удается найти number по xpath')
        return
    try:
        br_button = browser.find_element(By.XPATH, xpath_button)
    except:
        showerror('Ошибка', 'Не удается найти apply по xpath')
        return
    numbers_excel = ws['C']
    count = len(numbers_excel)
    for col in ws.iter_cols(min_row=2, min_col=3, max_col=3, max_row=count, values_only=True):
        for cell in col:
            number = cell
            br_login.send_keys(number)
            br_button.click()
            sleep(2)
            # try:
            #     br_open = browser.find_element(By.XPATH, xpath_open)           # нажимаем на забавный файлик
            # except:
            #     showerror('Ошибка', 'Не найден элемент')
            #     #код с переходом на следующий элемент
            #     return
            # br_open.click()
            showinfo('Ок','поиск по номеру завершен')

def start():
    folder_xl = txt_folder_xl.get()
    try:
        get_number(folder_xl)
    except:
        showerror('Ошибка', 'Ошибка входа в функцию')

window = Tk()
window.title('Программа')
window.geometry('400x250')

lbl_demo = Label(window, text = 'ДЕМО', font=('Arial',18,'bold'))
lbl_folder = Label(window, text='Расположение файла Excel')

txt_folder_xl = Entry(window, width = 20)

btn_start = Button(window, text='Начать заполнение', command=start)

lbl_demo.grid(column=2, row=0)
lbl_folder.grid(column=0, row=3)
txt_folder_xl.grid(column=1, row=3)
btn_start.grid(column=1, row=4)
window.mainloop()






