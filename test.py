from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from time import sleep
from tkinter import *
from tkinter.messagebox import showerror,showinfo

xpath_login = '//*[@id="login"]'
xpath_password = '//*[@id="password"]'
xpath_button = '/html/body/div[2]/div/section/div[2]/form/div[4]/input'
xpath_other = '/html/body/div[2]/div/section/div[2]/form/noindex/div[1]/a'
url_entry = 'https://roboparts.ru/auth/'
s = Service('chromedriver.exe')
browser = Chrome(service=s)
browser.get(url_entry)
# br_login = browser.find_element(By.XPATH, xpath_login)
# br_password = browser.find_element(By.XPATH, xpath_password)
br_button = browser.find_element(By.XPATH, xpath_other)
# br_login.send_keys('simandor@yandex.ru')
# br_password.send_keys('12345678')
br_button.click()
browser.refresh()
sleep(5)