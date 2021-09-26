import os

import url_constract
from bs4 import BeautifulSoup as Soup
import excel
from mail import send_excel

currencies = ["USD/RUB", "EUR/RUB"]
recipient = "stepan26062001@gmail.com"
data_len = 1


# parse xml to array of dicts
def parse(path):
    rows = []
    with open(path, 'r') as f:
        soup = Soup(f.read(), 'lxml')

    for row in soup.find_all('rate'):
        rows.append([f'{row.get("moment")}', f'{row.get("value")}'])
    return rows


# form right form of word
def form_text():
    res = f'{data_len} '
    if data_len % 10 == 1:
        res += "строка"
    elif 10 < data_len < 15:
        res += "строк"
    elif 1 < data_len % 10 < 5:
        res += "строки"
    elif 4 < data_len % 10:
        res += "строк"
    return res


for i in range(len(currencies)):
    request = url_constract.get_request(currencies[i])
    path = f'{currencies[i].replace("/", "2")}.xml'
    with open(path, 'w+') as f:
        f.write(request)
    data = parse(path)
    data_len = len(data) + 1
    excel.add_to_excel(data, i)
    os.remove(path)

excel.fill_G_column()
excel.optimize_columns_width()
text = form_text()
send_excel(recipient, text)
os.remove("exchange.xlsx")
