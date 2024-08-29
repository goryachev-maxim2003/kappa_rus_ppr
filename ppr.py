# Создать exe файл:
# pyinstaller --onefile --windowed ppr.py

import tkinter as tk
from tkinter import ttk
import pandas as pd
import numpy as np
import win32com.client
from win32com.client import constants
import os
import pythoncom

# Заменяем NaN значение в series последним значением, которое не NaN
def replace_nan_to_not_nan_before(sr):
    last_not_nan = np.nan
    for i in range(len(sr)):
        el = sr.iloc[i]
        if not pd.isnull(el):
            last_not_nan = el
        else:
            sr.iloc[i] = last_not_nan

def execute(f, par = []): #Добавляет проверки перед выполнением функции
    global errors
    errors.delete('1.0', 'end')
    try:
        try:
            f(*par) #запускаем функцию с переданными параметрами
        except PermissionError:
            raise Exception(f'Файл не может быть обновлён пока он открыт. Закройте файл "{graph_file_name}"')
        except ValueError as err:
            if str(err) == 'Value must be either numerical or a string containing a wildcard':
                raise Exception(f'Из-за применённых в файле фильтров, файл "{graph_file_name}" не может быть считан. \
Выставите все фильтры на (Выделить всё) и сохраните файл')
            else:
                raise ValueError(err)
        except pythoncom.com_error as error:
            if "Нет доступа к" in error.args[2][2]:
                raise Exception(f"{error.args[2][2]} Пожалуйста закройте этот файл")
            else:
                raise Exception(error)
    except Exception as e:
        errors.insert(1.0, str(e))

def open_task_file():
    global ppr, get_file_name, graph_file_name, graph_sheet_name, entry_machine, after_open_new_ttk_el, entry_machine_text
    #Получаем станок => формируем имя файла
    entry_machine_text = entry_machine.get()
    if (entry_machine_text):
        graph_file_name = get_file_name.loc[entry_machine_text, 'Файл']
    else:
        raise Exception(f"Не найден файл {graph_file_name}")
    frame = pd.read_excel(graph_file_name, graph_sheet_name, skiprows=[0])
    ppr = frame
    # Переводим столбец 'Кто' в нижний регистр (игнорируем np.nan)
    ppr.loc[~ppr['Кто'].isna(), 'Кто'] = ppr.loc[~ppr['Кто'].isna(), 'Кто'].map(lambda x : x.lower())
    # Изменяем null на последнее значение переодичности (потому что в исходном файле ячейки объеденины )
    replace_nan_to_not_nan_before(ppr.loc[:, "ПЕРИОДИЧНОСТЬ"])
    after_open_new_ttk_el()  # Добавляет новые элементы на интерфейс если файл != none


def format_sheet(worksheet):
    worksheet.Range("A1:B1").EntireColumn.AutoFit()
    worksheet.Range("D1").EntireColumn.AutoFit()

    worksheet.Range("E1:I1").EntireColumn.ColumnWidth = 25

    worksheet.Range("C1").EntireColumn.ColumnWidth = 50
    worksheet.Range("C1").EntireColumn.WrapText = True

def to_excel(worksheet, df, start_row = 1, start_col = 1, header = False, first_column_bold = False, border = False):
    if header:
        for i in range(len(df.columns)):
            cell = worksheet.Cells(start_row, i+1)
            cell.Value = df.columns[i]
            cell.Font.Bold = True
            if border:
                cell.Borders.Weight = 3
    for row in range(start_row, start_row+len(df)):
        for col in range(start_col, start_col + len(df.columns)):
            value = df.iloc[row - start_row, col - start_col]
            cell = worksheet.Cells(row+header, col)
            if not pd.isnull(value):
                cell.Value = value
                if first_column_bold and col == start_col:
                    cell.Font.Bold = True
            if border:
                cell.Borders.Weight = 3

                
def create_task_file():
    global graph_file_name, graph_sheet_name, ppr, who, week, entry_machine, entry_week
    # получаем week, who
    if (entry_week.get()):
        week = int(entry_week.get())
    else:
        raise Exception('Введите значение номера недели')
    if (entry_who.get()):
        who = entry_who.get()
    else:
        raise Exception('Введите значение плановой группы')
    
    # Формирование заданий
    task_cols_names = ['Узел', 'Название работы', 'Периодичность', 'Плановое время', 'Фактическое время', 'Исполнитель', 'Дата выполнения', 'Комментарии']
    # Добавляем новые столбцы в исходную таблицу
    new_cols = ['Плановое время', 'Фактическое время', 'Исполнитель', 'Дата выполнения', 'Комментарии']
    for col in new_cols:
        ppr[col] = np.nan
    # Формируем таблицы с еженедельным рабочим заданием
    # Прихотливым индексированием выбираем только те строчки где есть галочка в соответствующей неделе и где соответствующий столбец 'Кто'
    week_who_df = ppr.loc[  
        (ppr.loc[:, week] == '✓') & (ppr.loc[:, 'Кто'] == who),
        ['Узел', 'ПУНКТ ТО / НЕДЕЛЯ', 'ПЕРИОДИЧНОСТЬ', 'Плановое время', 'Фактическое время', 'Исполнитель', 'Дата выполнения', 'Комментарии', 'Кто']
    ]
    # Переименовываем колонки (столбец 'Кто' в итоговой таблице не нужен, он нужен только по группировке)
    week_who_df.columns = task_cols_names + ['Кто']
    rz_file_name = f'РЗ нед {week} {who} {entry_machine_text}.xlsx'.replace('/', ' ')
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    workbook = excel.Workbooks.Add()
    # with pd.ExcelWriter(rz_file_name) as writer: # По умолчанию используется движок openpyxl
    # Группируем по узлам
    first_iter = True
    for node in week_who_df['Узел'].dropna().unique():
        week_who_node_df = week_who_df[week_who_df['Узел'] == node]
        info_df = pd.DataFrame([
            ["Наименование оборудования", node],
            ["N недели", week],
            ["Плановая группа", who],
            ["Подтверждение ответственного", np.nan],
            ["Подтверждение руководителя", np.nan]
        ])
        worksheet = workbook.Worksheets(1) if first_iter else workbook.Worksheets.Add()
        worksheet_name = f'нед {week} {who} {node}'
        worksheet.Name = worksheet_name if len(worksheet_name) < 32 else worksheet_name[:31]
        to_excel(worksheet, info_df, header=False, first_column_bold=True)
        week_who_node_df = week_who_node_df.reset_index()
        week_who_node_df = week_who_node_df[task_cols_names] #Убираем лишние столбцы(столбец 'Кто')
        week_who_node_df.insert(0, 'N пп', week_who_node_df.index + 1)
        to_excel(worksheet, week_who_node_df, start_row = len(info_df)+2, header=True, border=True)
        # Форматируем лист
        format_sheet(worksheet)
        first_iter = False
    workbook.SaveAs(os.getcwd()+'/'+rz_file_name)
    # workbook.Close()
    # excel.Quit()
root = tk.Tk()
root.title('Формирование еженедельного задания')
root.geometry('750x500')
errors_label = ttk.Label(text='Ошибки:')
errors_label.place(x=50, y=210)
errors = tk.Text(root, width=80, height=10, foreground='red')
errors.place(x=50, y=240)

graph_sheet_name = 'ГОД'

get_file_name = pd.read_excel('Файл для ppr exe.xlsx', 'Файлы', index_col = 0)

# Добавляет новые элементы на интерфейс если файл успешно открылся
def after_open_new_ttk_el():
    global ppr, week_label, entry_week, who_label, entry_who, bt_generate_task  
    if ppr is not None:
        file_label = ttk.Label(text=f'Считан файл {graph_file_name}')
        file_label.place(x=50, y=85)
        week_label = ttk.Label(text='Ввод номера недели')
        week_label.place(x=50, y=115)
        entry_week = ttk.Combobox(values=[i for i in range(1, 52+1)], state='readonly')
        entry_week.place(x=200, y=115)

        who_label = ttk.Label(text='Ввод плановой группы')
        who_label.place(x=50, y=145)
        entry_who = ttk.Combobox(values=list(ppr['Кто'].dropna().unique()), state='readonly')
        entry_who.place(x=200, y=145)

        bt_generate_task = tk.Button(root, text='Сформировать задание', width=30, height=1, command=lambda: execute(create_task_file))
        bt_generate_task.place(x=50, y=175)

ppr = None
machine_label = ttk.Label(text='Выбор станка')
machine_label.place(x=50, y=25)
entry_machine = ttk.Combobox(values=list(get_file_name.index), state='readonly', width=50)
entry_machine.place(x=200, y=25)
bt_generate_task = tk.Button(root, text='Считать файл', width=30, height=1, command=lambda: execute(open_task_file))
bt_generate_task.place(x=50, y=55)

root.mainloop()