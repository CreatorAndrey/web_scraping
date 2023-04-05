def parser 2(html, number):
    logging.info('In parser function')
    list_excel = []
    try:
        soup = BeautifulSoup(html, 'html.parser')
        soup2 = BeautifulSoup(html, 'lxml')
        logging.info("The soup is creating successful")
    except Exception:
        logging.exception("Problem with creating the soup")
        return

    # парсинг даты завершения работ E
    pr_date_end = soup.find_all('div', class_=class_date_end)
    logging.info('parsing date of ending')
    if len(pr_date_end) == 0:
        text_log.insert(END, f"Номер {number}, не найдена дата окончания работ (E)\n")
        list_excel.append("")
    else:
        span_end = pr_date_registration[0]
        date_end = span_end.find_next('span').string
        list_excel.append(date_end)
        logging.info(f'append in list date of end {date_end}')

    # парсинг уведомления о завершении работ F
    pr_notify_end = soup.find_all('div', class_='field-name-field-fgpn-notify-end-rel')
    if len(pr_notify_end) == 0:
        text_log.insert(END, f"Номер {number}, не найдено уведомление о завершении работ (F)\n")
        list_excel.append("")
    else:
        a = pr_notify_end[0]
        notify_end = a.find_next('a').string
        m = re.search('..-..-....-......', notify_end)
        if m is None:
            list_excel.append("")
            text_log.insert(END, f"Номер {number}, не найдено уведомление о завершении работ (F)\n")
            logging.info('не найдена подстрока в строке')
        else:
            list_excel.append(m.group())
        logging.info(f'append in list the {m.group()}')

numbers_excel = ws['C']         # берем все ячейки C
count = len(numbers_excel)      # считаем кол-во ячеек непустых
for col in ws.iter_cols(min_row=2, min_col=3, max_col=3, max_row=count):
    for cell in col:
        number = cell.value
        logging.info(f"Take number {number} from worklist")
        if number is None:          # если в excel пустоя ячейка (отсутствует номер), исключаем None. Иначе ошибка TypeError в send_keys()
            logging.info(f"Take number number from worklist, number is {number}, break")
            break
        try:
            br_number = browser.find_element(By.XPATH, xpath_number)  # находим поле с вводом номера регистрации
            logging.info('Элемент br_number упешно найден')
        except Exception:
            showerror('Ошибка', 'Не удается найти элемент')
            logging.exception('Элемент br_number не найден')
            return
        try:
            br_apply = browser.find_element(By.XPATH, xpath_apply)  # находим поле с кнопкой "применить"
            logging.info('Элемент br_apply упешно найден')
        except Exception:
            showerror('Ошибка', 'Не удается найти элемент')
            logging.exception('Элемент br_apply не найден')
            return
        logging.info('Clear the input textarea')
        br_number.clear()
        try:
            br_number.send_keys(number)         # отправляем номер в поле регистрации
            logging.info("the number from excel is sended in the br_number complitely ")
        except Exception:
            logging.exception("the number from excel is not sended in br_number")
            return
        try:
            br_apply.click()            # нажимаем на кнопку "применить"
            logging.info("Click on the button_apply is successful")
        except Exception:
            logging.exception("Click on the button_apply isn't successful")
            return
        sleep(2)
        try:
            br_open = browser.find_element(By.XPATH, xpath_open)           # находим забавный файлик
            logging.info('Элемент br_open найден')
        except Exception:
            # showerror('Ошибка', 'Не найден элемент')
            logging.exception('Не найден элемент br_open')
            # код с переходом на следующий элемент
            text_log.insert(END, f"Номер {number}, не найден 'файлик' (D)\n")
            break
        br_open.click()         # нажимаем на кнопку с файликом
        logging.info('Click on the img_file')
        html = browser.page_source          # берем html страницы
        logging.info('Take the html')
        #logging.info(html)
        list_excel = parser(html, number)            # передаем html в парсер и создаем список list_excel
        logging.info('start insert data into excel')