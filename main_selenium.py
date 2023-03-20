# https://istories.media/workshops/2021/09/20/parsing-s-pomoshchyu-python-urok-2/
# https://questu.ru/articles/81673/

from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from time import sleep
from tkinter import *
from tkinter.messagebox import showerror, showinfo

xpath_login = '//*[@id="login"]'
xpath_password = '//*[@id="password"]'
xpath_button = '/html/body/div[2]/div/section/div[2]/form/div[4]/input'
url_entry = 'https://roboparts.ru/auth/'
s = Service("\chromedriver.exe")
browser = Chrome(service=s)

def entry():
    browser.get(url_entry)
    br_login = browser.find_element(By.XPATH, xpath_login)
    br_password = browser.find_element(By.XPATH, xpath_password)
    br_button = browser.find_element(By.XPATH, xpath_button)
    login = txt_login.get()
    password = txt_password.get()
    br_login.send_keys(login)
    br_password.send_keys(password)
    br_button.click()
    browser.refresh()
    #sleep(2)
    br_login = browser.find_elements(By.XPATH, xpath_login)
    if len(br_login) == 0:
        showinfo('Вход','Успешная авторизация')
        btn_entry.configure(state='disabled')
    else:
        showerror('Ошибка', 'Ошибка авторизации')


window = Tk()
window.title('Программа')
window.geometry('400x250')
lbl_login = Label(window, text="Логин:")
lbl_password = Label(window, text="Пароль:")
txt_login = Entry(window, width=10)
txt_password = Entry(window, width=10)
btn_entry = Button(window, text='Авторизоваться', command=entry)
lbl_login.grid(column=0, row=0)
txt_login.grid(column=1, row=0)
lbl_password.grid(column=0, row=1)
txt_password.grid(column=1, row=1)
btn_entry.grid(column=0, row=2)
window.mainloop()






