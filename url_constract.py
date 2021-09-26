import datetime
from dateutil import relativedelta
import requests

url = "https://www.moex.com/export/derivatives/currency-rate.aspx"

today = datetime.date.today()
prev_month = (datetime.date.today() - relativedelta.relativedelta(months=1))

url_args = {
    "language": "ru",
    "currency": "",
    "moment_start": str(prev_month),
    "moment_end": str(today)
}


def get_request(currency):
    url_args['currency'] = currency
    return requests.get(url, params=url_args).text
