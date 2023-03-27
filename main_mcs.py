# https://istories.media/workshops/2021/09/20/parsing-s-pomoshchyu-python-urok-2/
# https://questu.ru/articles/81673/
# https://stackoverflow.com/questions/29858752/error-message-chromedriver-executable-needs-to-be-available-in-the-path

from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
from tkinter import *
from tkinter.messagebox import showerror, showinfo
from openpyxl import load_workbook
import logging

logging.basicConfig(level=logging.DEBUG, filename='log.log', filemode='w', format="%(asctime)s %(levelname)s %(message)s")

xpath_login = '//*[@id="login"]'
xpath_password = '//*[@id="password"]'
xpath_button = '/html/body/div/div/div/div/div/div/div/div[3]/form/button'          #кнопка авторизации
xpath_go = '/html/body/div/div/div/div/div/div[4]/div/div/div/div[2]/div/div[2]/div/div/div/a'           #ссылка второй страницы, "перейти"
xpath_notice = '//*[@id="block-system-main-menu"]/ul/li/ul/li[3]/ul/li[4]/a'        #уведомление
xpath_filter_open = '/html/body/div[2]/div/section/h1/span[2]'
xpath_apply = '//*[@id="edit-submit-fgpn-license-notification-record"]'         # кнопка применить
xpath_number = '//*[@id="edit-field-gl-registry-num-value"]'           # ссылка на поле с вводом номера регистрации
xpath_open = '//*[@id="block-system-main"]/div/div[2]/div/table/tbody/tr/td[12]/div/a/span'     # забавный файлик

url_entry = 'https://passport.cgu.mchs.ru/oauth/login?login_challenge=9043dbc4bed146c3ae16ef4e6c39fa7e'
s = Service("chromedriver.exe")
browser = Chrome(service=s)

def entry():
    try:
        browser.get(url_entry)
        logging.info(f'успешное открытие страницы регистрации {url_entry}')
    except:
        logging.exception(f'Не удается открыть {url_entry}')
        showerror('Ошибка', 'Не удается перейти на страницу входа')
        return

    try:
        br_login = browser.find_element(By.XPATH, xpath_login)
        br_password = browser.find_element(By.XPATH, xpath_password)
        br_button = browser.find_element(By.XPATH, xpath_button)
        logging.info(f'Элементы br_button, br_login, br_password успешно найдены')
    except:
        showerror('Ошибка','Не найден элемент')
        logging.exception(f'Один из элементов br_button, br_login, br_password не найден')
        return
    login = txt_login.get()
    password = txt_password.get()
    br_login.send_keys(login)
    br_password.send_keys(password)
    br_button.click()
    #browser.refresh()
    #sleep(2)
    br_login = browser.find_elements(By.XPATH, xpath_login)
    if len(br_login) == 0:
        showinfo('Вход','Успешный вход')
        btn_entry.configure(state='disabled')
        try:
            br_go = browser.find_element(By.XPATH, xpath_go)            # ищем ссылку "перейти"
        except:
            showerror('Ошибка', 'Не найдена ссылка')
            return
        br_go.click()           # нажимаем на ссылку перейти
        sleep(2)
        try:
            browser.switch_to.window(browser.window_handles[1])         # переключаемся на вторую вкладку
        except:
            showerror('Ошибка', 'Ожидание второй вкладки, а в браузере всего лишь одна')
        try:
            br_notice = browser.find_element(By.XPATH, xpath_notice)        # ищем ссылку под названием "уведомление"
        except:
            showerror('Ошибка', 'Не найдена ссылка')
            return
        br_notice.click()           # кликаем и переходим на страницу с поиском
        txt_folder_xl.configure(state='normal')
        btn_start.configure(state='normal')
        sleep(2)
    else:
        showerror('Ошибка', 'Ошибка входа')
        logging.exception('Ошибка входа')

def get_number(folder):
    logging.info('Успешно вошли в функцию get_number')
    try:
        wb = load_workbook(folder)
        logging.info(f'Книга успешно открыта по адресу {folder}')
    except:
        showerror('Ошибка', 'Не удается подключиться к книге Excel')
        logging.exception(f'Не удается подключиться к книге по пути {folder}')
        return
    try:
        ws = wb['Главный_лист']
        logging.info('Лист Excel успешно открыт')
    except:
        showerror('Ошибка', 'Не найден лист')
        logging.exception(f'Лист не найден')
        return
    try:
        br_filter_open = browser.find_element(By.XPATH, xpath_filter_open)      # находим кнопку фильтра
        logging.info("The filter is opening successfully")
    except:
        logging.exception("The filter isn't opening successfully")
        return
    br_filter_open.click()          # кликаем на кнопку фильтра
    try:
        br_number = browser.find_element(By.XPATH, xpath_number)        # находим поле с вводом номера регистрации
        logging.info('Элемент br_number упешно найден')
    except:
        showerror('Ошибка', 'Не удается найти элемент')
        logging.exception('Элемент br_number не найден')
        return
    try:
        br_apply = browser.find_element(By.XPATH, xpath_apply)          # находим поле с кнопкой "применить"
        logging.info('Элемент br_apply упешно найден')
    except:
        showerror('Ошибка', 'Не удается найти элемент')
        logging.exception('Элемент br_apply не найден')
        return
    numbers_excel = ws['C']
    count = len(numbers_excel)
    for col in ws.iter_cols(min_row=2, min_col=3, max_col=3, max_row=count, values_only=True):
        for cell in col:
            number = cell
            logging.info(f"Take number {number} from worklist")
            try:
                br_number.send_keys(number)         # отправляем номер в поле регистрации
                logging.info("the number from excel is sended in br_number complitely ")
            except:
                logging.exception("the number from excel is not sended in br_number")
                return
            try:
                br_apply.click()            # нажимаем на кнопку "применить"
                logging.info("Click on button_apply is successful")
            except:
                logging.exception("Click on button_apply isn't successful")
                return
            sleep(2)
            try:
                br_open = browser.find_element(By.XPATH, xpath_open)           # находим забавный файлик
                logging.info('Элемент br_open найден')
            except:
                showerror('Ошибка', 'Не найден элемент')
                logging.exception('Не найден элемент br_open')
                #код с переходом на следующий элемент
                return
            br_open.click()         # нажимаем на кнопку с файликом
            showinfo('Ок','поиск по номеру завершен')

def start():
    folder_xl = txt_folder_xl.get()
    get_number(folder_xl)

window = Tk()
window.title('Программа')
window.geometry('400x250')
lbl_login = Label(window, text="Логин:")
lbl_password = Label(window, text="Пароль:")
lbl_demo = Label(window, text = 'ДЕМО', font=('Arial',18,'bold'))
lbl_folder = Label(window, text='Расположение файла Excel')
txt_login = Entry(window, width=20)
txt_password = Entry(window, width=20)
txt_folder_xl = Entry(window, width = 20, state='disabled')
btn_entry = Button(window, text='Войти', width=17, command=entry)
btn_start = Button(window, text='Начать заполнение', state='disabled', command=start)
lbl_login.grid(column=0, row=0)
lbl_demo.grid(column=2, row=0)
txt_login.grid(column=1, row=0)
lbl_password.grid(column=0, row=1)
txt_password.grid(column=1, row=1)
btn_entry.grid(column=1, row=2)
lbl_folder.grid(column=0, row=3)
txt_folder_xl.grid(column=1, row=3)
btn_start.grid(column=1, row=4)
window.mainloop()
