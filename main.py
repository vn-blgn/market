import shutil
from src import helpers


def main():
    print('Подготовка процесса...')
    today_file, today_cell = helpers.get_today('_', '.')
    new_file_name, desktop = helpers.change_file_name(today_file)
    file_date, total_before, total_before_cash = helpers.get_cell(new_file_name)

    print('Обновление и запись котировок в файл...')
    data = helpers.upload_data()
    report, counter = helpers.write_data(new_file_name, today_cell, data)
    print(f'Обновлено {counter} котировок!')

    print('Ценовые изменения на пять и более процентов:')
    helpers.get_analytics(report)

    print('Изменения в Total:')
    total_after = helpers.get_total(new_file_name, 'C439')
    total_difference = total_after - total_before
    helpers.data_print(file_date, total_before, today_cell, total_after,
                       total_difference, total_before_cash)

    shutil.copy(new_file_name, desktop)
    print('Обновленный файл находится на рабочем столе!')


if __name__ == '__main__':
    main()
