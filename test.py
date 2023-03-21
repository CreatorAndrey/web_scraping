# https://istories.media/workshops/2021/09/20/parsing-s-pomoshchyu-python-urok-2/
# https://questu.ru/articles/81673/

from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
from tkinter import *
from tkinter.messagebox import showerror, showinfo

xpath_login = '//*[@id="login"]'
xpath_password = '//*[@id="password"]'
xpath_button = '/html/body/div/div/div/div/div/div/div/div[3]/form/button'
xpath_notice = '//*[@id="block-system-main-menu"]/ul/li/ul/li[3]/ul/li[4]/a'
url_entry = 'https://passport.cgu.mchs.ru/oauth/login?login_challenge=9043dbc4bed146c3ae16ef4e6c39fa7e'
#s = Service("chromedriver.exe")
browser = Chrome(ChromeDriverManager().install())

def entry():
    try:
        browser.get(url_entry)
    except:
        showerror('Ошибка', 'Не удается перейти на страницу входа')
        return

    try:
        br_login = browser.find_element(By.XPATH, xpath_login)
        br_password = browser.find_element(By.XPATH, xpath_password)
        br_button = browser.find_element(By.XPATH, xpath_button)
    except:
        showerror('Ошибка','Не найден элемент')
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
        br_notice = browser.find_element(By.XPATH, xpath_notice)
        br_notice.click()
        btn_start.configure(state='normal')
    else:
        showerror('Ошибка', 'Ошибка входа')


window = Tk()
window.title('Программа')
window.geometry('400x250')
lbl_login = Label(window, text="Логин:")
lbl_password = Label(window, text="Пароль:")
lbl_demo = Label(window, text = 'ДЕМО', font=('Arial',18,'bold'))
txt_login = Entry(window, width=20)
txt_password = Entry(window, width=20)
btn_entry = Button(window, text='Войти', width=17, command=entry)
btn_start = Button(window, text='Начать заполнение', state='disabled')
lbl_login.grid(column=0, row=0)
lbl_demo.grid(column=2, row=0)
txt_login.grid(column=1, row=0)
lbl_password.grid(column=0, row=1)
txt_password.grid(column=1, row=1)
btn_entry.grid(column=1, row=2)
btn_start.grid(column=2, row=2)
window.mainloop()
