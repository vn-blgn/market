import os
import glob
import yfinance as yf
import xlwings as xw
from openpyxl import load_workbook
from datetime import date
from src.ticker_symbols import ticker_symbols as ts, tickers as t


def get_today(*args):
    today_file = date.today().strftime(f'%d{args[0]}%m{args[0]}%Y')
    today_cell = date.today().strftime(f'%d{args[1]}%m{args[1]}%Y')
    return today_file, today_cell


def change_file_name(today_file):
    current_file = glob.glob('*.xlsx')[0]
    new_file_name = f'Portfolio_{today_file}.xlsx'
    os.rename(current_file, new_file_name)
    desktop = os.path.expanduser(f'~/Desktop/{new_file_name}')
    return new_file_name, desktop


def get_cell(file_name):
    app = xw.App(visible=False)
    xlsx_book = app.books.open(file_name)
    sheet = xlsx_book.sheets['Portfolio']
    file_date = sheet['A1'].value
    total_before = round(sheet['C439'].value, 2)
    total_before_cash = round(sheet['C473'].value, 2)
    xlsx_book.close()
    return file_date, total_before, total_before_cash


def upload_data():
    tickers = yf.Tickers(t)
    data = tickers.tickers.items()
    return data


def get_values(data, ws):
    counter = 0
    report = []
    for key, value in data:
        stock_info = value.info
        symbol = stock_info['symbol']
        current_price = stock_info['currentPrice']
        currency = stock_info['currency']
        previous_close = stock_info['previousClose']
        counter += 1
        cell = ts.get(symbol)
        ws[cell] = current_price
        symbols = {'symbol': symbol, 'current_price': current_price,
                   'previous_close': previous_close, 'currency': currency}
        report.append(symbols)
    return report, counter


def write_data(new_file_name, today_cell, data):
    wb = load_workbook(new_file_name)
    ws = wb['Portfolio']
    ws['A1'] = today_cell
    report, counter = get_values(data, ws)
    wb.save(new_file_name)
    wb.close()
    return report, counter


def get_margin(cur_close, pre_close):
    if cur_close > pre_close:
        margin = ((cur_close - pre_close) / pre_close) * 100
        return round(margin, 2)
    elif cur_close < pre_close:
        margin = ((cur_close - pre_close) / pre_close) * 100
        return round(margin, 2)
    else:
        return '0'


def get_analytics(report_list):
    for elem in report_list:
        elem['margin'] = get_margin(elem['current_price'], elem['previous_close'])
        if int(elem['margin']) >= 5:
            print(f"{elem['symbol']}: {round(elem['current_price'], 2)}{elem['currency']} "
                  f"({round(elem['previous_close'], 2)}{elem['currency']}) ↑ рост +{elem['margin']}%")
        elif int(elem['margin']) <= -5:
            print(f"{elem['symbol']}: {round(elem['current_price'], 2)}{elem['currency']} "
                  f"({round(elem['previous_close'], 2)}{elem['currency']}) ↓ падение {elem['margin']}%")


def get_total(file_name, cell):
    app = xw.App(visible=False)
    xlsx_book = app.books.open(file_name)
    sheet = xlsx_book.sheets['Portfolio']
    total = sheet[cell].value
    xlsx_book.close()
    return round(total, 2)


def data_print(file_date, total_before, today_cell, total_after,
               total_difference, total_before_cash):
    print(f'Total ({file_date}) —> {total_before}')
    print(f'Total ({today_cell}) —> {total_after}')
    print(f'Разница —> {round(total_difference, 2)}')
    print(f'Total + Cash ({file_date}) —> {round(total_before_cash, 2)}')
