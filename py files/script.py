'''
Проект создан в рамках хакатона Skillfactory.
В данном файле хранится основной скрипт проекта.
'''

import gspread as gs
import pandas as pd

from datetime import date

import functions as f


account = gs.service_account(filename='table-report-0e5d772bb3fa.json')
today = date.today().strftime('%d.%m.%y')

while True:
    try:
        table = account.open_by_url(
            input('Введите URL-адрес таблицы и нажмите "Enter": ')
        )
        ws_titles = [i.title for i in table.worksheets()]
        month = table.worksheet('Месяц учета оказания услуг').row_values(1)
        date = table.worksheet('Дата учета оказания услуг').row_values(1)

        if ((('Месяц учета оказания услуг' or 'Дата учета оказания услуг')
             in ws_titles) and (today in (month or date))):
            print('''
В отчётах уже есть колонки с сегодняшней датой.
Их необходимо удалить для корректного анализа Вашей таблицы.
''')
        else:
            break

    except gs.exceptions.NoValidUrlKeyFound:
        print('Введён некорректный URL-адрес.')

    except gs.exceptions.APIError:
        print('Необходимо открыть доступ к таблице с ролью "Редактор" '
              'для пользователя "tablereport.sf@gmail.com".')

ws_0 = table.get_worksheet(0)
ws_0_df = pd.DataFrame(ws_0.get_all_records())
ws_0_df_month = f.get_data(ws_0_df, 'Месяц учета оказания услуг')
ws_0_df_date = f.get_data(ws_0_df, 'Дата учета оказания услуг')

if (('Месяц учета оказания услуг' and 'Дата учета оказания услуг')
        not in ws_titles):
    f.create_new_worksheet(table, 'Месяц учета оказания услуг', ws_0_df_month)
    f.create_new_worksheet(table, 'Дата учета оказания услуг', ws_0_df_date)
    print('Отчёты обновлены успешно!')

else:
    f.update_worksheet(table.worksheet('Месяц учета оказания услуг'),
                       ws_0_df_month)
    f.update_worksheet(table.worksheet('Дата учета оказания услуг'),
                       ws_0_df_date)
    print('Отчёты обновлены успешно!')
