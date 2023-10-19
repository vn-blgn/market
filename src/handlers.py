import os
import shutil
import requests
from tkinter import filedialog as fd
from src import helpers


def update_quotes(txt_edit, END):
    try:
        today_file, today_cell = helpers.get_today('_', '.')
        new_file_name = f'Portfolio_{today_file}.xlsx'
        helpers.change_file_name(new_file_name)
        desktop = os.path.expanduser(f'~/Desktop/{new_file_name}')
        file_date, total_before, total_before_cash = helpers.get_cell(new_file_name)

        data = helpers.upload_data()
        report, counter = helpers.write_data(new_file_name, today_cell, data)
        txt_edit.insert(END, f'Обновлено {counter} котировок!\n')

        txt_edit.insert(END, 'Ценовые изменения на пять и более процентов:\n')
        helpers.get_analytics(report, txt_edit, END)

        txt_edit.insert(END, 'Изменения в Total:\n')
        total_after = helpers.get_total(new_file_name, 'C439')
        total_difference = total_after - total_before
        helpers.data_print(file_date, total_before, today_cell, total_after,
                           total_difference, total_before_cash, txt_edit, END)

        shutil.copy(new_file_name, desktop)
        txt_edit.insert(END, 'Обновленный файл находится на рабочем столе!\n')
    except requests.ConnectionError:
        txt_edit.insert(END, 'Проверьте подключение к интернету!\n')


def upload_file(txt_edit, END):
    src_file = fd.askopenfilename()
    file_name = src_file.split('/')[-1]
    helpers.change_file_name(file_name)
    destination = os.getcwd()
    dest_file = os.path.expanduser(f'{destination}\\{file_name}')
    if src_file:
        shutil.copy(src_file, dest_file)
        txt_edit.insert(END, 'Файл успешно загружен!\n')
