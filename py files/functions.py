'''
Проект создан в рамках хакатона Skillfactory.
В данном файле хранятся все исполняемые функции.
'''

import gspread as gs
import pandas as pd

from datetime import date


def get_data(
        dataframe: pd.DataFrame,
        needed_col_title: str
) -> pd.DataFrame:
    '''Получение информации из основной таблицы'''
    return dataframe[[
        'ФИО/Название\nподрядчика',
        'Уникальный номер размещения',
        needed_col_title
    ]]


def get_upd_data(
    upd_data_df: pd.DataFrame
) -> list:
    '''Получение последнего заполненного столбца из отчёта'''
    result = []
    for row in upd_data_df.values.tolist():
        for cell in row[::-1]:
            if cell == '':
                continue
            result.append(cell)
            break
    return result


def create_new_worksheet(
        table: gs.Spreadsheet,
        worksheet_title: str,
        data: pd.DataFrame
) -> None:
    '''Создание отчёта'''
    ws = table.add_worksheet(title=worksheet_title, rows=50, cols=20)
    first_row = [[
        'ФИО/Название\nподрядчика',
        'Уникальный номер размещения'
    ]]
    ws.insert_rows(first_row + data.values.tolist(), 1)
    ws.update_cell(1, 3, date.today().strftime('%d.%m.%y'))


def update_worksheet(
        upd_ws: gs.Worksheet,
        old_data_df: pd.DataFrame,
        today_date=date.today().strftime('%d.%m.%y')
) -> None:
    '''Обновление отчёта'''
    upd_data_df = pd.concat(
        [
            pd.DataFrame(upd_ws.get_all_records()),
            pd.DataFrame(get_upd_data(pd.DataFrame(upd_ws.get_all_records())))
        ],
        axis=1
    )
    temp_data = [today_date]
    old_data = old_data_df[upd_ws.title].values.tolist()

    if old_data_df.shape[0] != upd_data_df.shape[0]:
        temp_list = [[], []]

        for cell in old_data_df['ФИО/Название\nподрядчика'].values.tolist():
            if (cell not in
                    upd_data_df['ФИО/Название\nподрядчика'].values.tolist()):
                temp_list[0].append(cell)

        for cell in old_data_df['Уникальный номер размещения'].values.tolist():
            if (cell not in
                    upd_data_df['Уникальный номер размещения'].values.tolist()):
                temp_list[1].append(cell)

        value = list()
        for i in range(len(temp_list[0])):
            temp = list()
            temp.append(temp_list[0][i])
            temp.append(temp_list[1][i])
            value.append(temp)

        upd_ws.insert_rows(value, len(upd_ws.col_values(1)) + 1)
        upd_data_df = pd.DataFrame(upd_ws.get_all_records())

    upd_data = get_upd_data(upd_data_df)

    if old_data != upd_data:
        for i in range(len(old_data)):
            temp_data.append('') if old_data[i] == upd_data[i] \
                else temp_data.append(old_data[i])

    upd_ws.insert_cols([temp_data], len(upd_ws.row_values(1)) + 1)
