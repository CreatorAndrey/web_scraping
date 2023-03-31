from bs4 import BeautifulSoup
import lxml

with open('D:\Работа\Авито парсер локальный сайт\html lxml.txt', 'rb') as f:
    soup = BeautifulSoup(f, 'lxml')
    # print(soup.prettify())
    # pr_address_work = soup.find('table', class_='tableheader-processed')
    # print(pr_address_work)

    data = []
    table = soup.find('table', class_='tableheader-processed')
    table_body = table.find('tbody')
    rows = table_body.find_all('tr')
    for row in rows:

        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        data.append([ele for ele in cols if ele])
    print(data)


    # div_list = pr_address_work.find_all('div', class_='field-item')
    # addresses = ""
    # for i in div_list:
    #     addresses += i.text + '\n'
    # print(addresses)

    # div = pr_address_work[0]
    # address_work = div.find_next('div', class_='field-item even').string
    # print(address_work)
