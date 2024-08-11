# pyinstaller --onefile --windowed ppr.py //Для того, чтобы сделать exe файл без консоли

import tkinter as tk
from tkinter import ttk
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Border, Side
import xlwings as xw

def execute(f): #Добавляет проверки перед выполнением функции
    errors.delete('1.0', 'end')
    try:
        try:
            f()
        except PermissionError:
            raise Exception(f'Файл не может быть обновлён пока он открыт. Закройте файл "{graph_file_name}"')
        except ValueError as err:
            if str(err) == 'Value must be either numerical or a string containing a wildcard':
                raise Exception(apply_filter_error_message)
            else:
                raise ValueError(err)
    except Exception as e:
        errors.insert(1.0, str(e))

def open_task_file():
    global ppr
    # Открываем файл
    # try:
    #     try:
    frame = pd.read_excel(graph_file_name, graph_sheet_name, skiprows=[0])
    ppr = frame
        # except ValueError as err:
            # if str(err) == 'Value must be either numerical or a string containing a wildcard':
                # raise Exception(apply_filter_error_message)
            # else:
                # raise ValueError(err)
        # Переводим столбец 'Кто' в нижний регистр (игнорируем np.nan)
    ppr.loc[~ppr['Кто'].isna(), 'Кто'] = ppr.loc[~ppr['Кто'].isna(), 'Кто'].map(lambda x : x.lower())  
    # except Exception as err:
        # raise ValueError(err) 


def create_task_file():
    global graph_file_name, graph_sheet_name, ppr, who, week
    # получаем week и who
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
    # Группируем по узлам
    with pd.ExcelWriter(f'рабочее_задание_неделя_{week}_{who}.xlsx') as writer:
        for node in week_who_df['Узел'].dropna().unique():
            week_who_node_df = week_who_df[week_who_df['Узел'] == node]
            info_df = pd.DataFrame([
                ["Наименование оборудования", node],
                ["N недели", week],
                ["Плановая группа", who],
                ["Подтверждение ответственного", np.nan],
                ["Подтверждение руководителя", np.nan]
            ])
            info_df.to_excel(writer, sheet_name=f'нед_{week}_{who}_{node}', header=False, index=False)
            week_who_node_df = week_who_node_df.reset_index()
            week_who_node_df = week_who_node_df[task_cols_names] #Убираем лишние столбцы('Кто')
            week_who_node_df.insert(0, 'N пп', week_who_node_df.index + 1)
            week_who_node_df.to_excel(writer, sheet_name=f'нед_{week}_{who}_{node}', startrow = len(info_df)+1, index=False)  # записываем без колонки кто
            #Форматируем лист по ширине
            worksheet = writer.sheets[f'нед_{week}_{who}_{node}']
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter
                max_length = max([len(str(cell.value)) for cell in col if cell.value != None])
                # перенос текста
                if (max_length > 100):
                    worksheet.column_dimensions[column].width = 50
                    for cell in col:
                        cell.alignment = openpyxl.styles.Alignment(wrap_text=True)
                # автоподбор ширины столбца
                else:
                    worksheet.column_dimensions[column].width = max_length+2
                # for cell in col:
                #     cell.border = Border(left = Side(border_style = 'double'), right = Side(border_style = 'double'), top = Side(border_style = 'double'), bottom = Side(border_style = 'double'))
    # Открываем файл
    xw.Book(f'рабочее_задание_неделя_{week}_{who}.xlsx')

root = tk.Tk()
root.title('Формирование еженедельного задания')
root.geometry('750x500')
errors_label = ttk.Label(text='Ошибки:')
errors_label.place(x=50, y=130)
errors = tk.Text(root, width=80, height=10, foreground='red')
errors.place(x=50, y=160)

graph_file_name = 'График ТО ГА 2024.xlsx'
graph_sheet_name = 'ГОД'

apply_filter_error_message = f'Из-за применённых в файле фильтров, файл "{graph_file_name}" не может быть считан.  \
    Выставите все фильтры на (Выделить всё) и сохраните файл'

#Открываем файл
ppr = None
execute(open_task_file) #появляется переменная ppr
# open_task_file()

if ppr is not None:
    week_label = ttk.Label(text='Ввод номера недели')
    week_label.place(x=50, y=25)
    entry_week = ttk.Combobox(values=[i for i in range(1, 52+1)], state='readonly')
    entry_week.place(x=200, y=25)
    who_label = ttk.Label(text='Ввод плановой группы')
    who_label.place(x=50, y=55)
    entry_who = ttk.Combobox(values=list(ppr['Кто'].dropna().unique()), state='readonly')
    entry_who.place(x=200, y=55)
    bt_generate_task = tk.Button(root, text='Сформировать задание', width=30, height=1, command=lambda: execute(create_task_file))
    # bt_generate_task = tk.Button(root, text='Сформировать задание', width=30, height=1, command=create_task_file)
    bt_generate_task.place(x=50, y=85)

root.mainloop()