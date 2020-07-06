"""
© Ivan Chanke, 2020
"""

import io
import tkinter as tk
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import tkinter.ttk as ttk
from tkinter.filedialog import askopenfilename, asksaveasfilename
from tkinter.messagebox import showerror, askyesno
import seaborn as sns
import pickle

sns.set_style('darkgrid')

app = tk.Tk()
app.geometry('640x480')
app.title('AnalyticsApp')
# Глобальные переменные
"""
database - главная глобальная переменная, в которой хранится загруженная
в программу база, состоящая из 4 таблиц:
станции, линии, метрополитены и соединение.
Все функции, изменяющие базу, а также вид базы (сортировка, удаление),
работают с этой переменной, изменяя её
"""
database = { # Пустая база - пустые датафреймы со столбцами
    'stations': pd.DataFrame(data = [], columns = [
            'Название', 'Код линии', 'Глубина заложения', 'Тип', 'Дата открытия'
        ]),
    'lines': pd.DataFrame(data = [], columns = [
            'Код линии', 'Название', 'Метрополитен', 'Цвет', 'Год открытия линии'
        ]),
    'metros': pd.DataFrame(data = [], columns = [
            'Метрополитен', 'Город', 'Год основания метро'
        ]),
    'merged': pd.DataFrame(data = [], columns = [
            'Название станции', 'Код линии', 'Глубина заложения', 'Тип',
            'Дата открытия станции',
            'Название линии', 'Метрополитен', 'Цвет линии', 'Год открытия линии'
        ])
}
current_filtered = None # Хранит текущее фильтрованное отображение для возможного экспорта
current_file_name = None # Имя загруженной в данный момент базы
open_tab = None
# Хранят длины датафреймов и количество выбранных в данный момент строк
len_1, sel_1 = 0, 0
len_2, sel_2 = 0, 0
len_3, sel_3 = 0, 0
len_4, sel_4 = 0, 0
number_of_rows = tk.StringVar()
number_selected = tk.StringVar()

# События и управляющие функции

def configure_f(current):
    """
    Открывает окно настройки фильтров
    """
    def clear_all_entries():
        """
        Очищает все строки окна настройки фильтров
        """
        sn_entry.delete(0, tk.END)
        lc_entry.delete(0, tk.END)
        depth_entry.delete(0, tk.END)
        type_entry.delete(0, tk.END)
        open_date_entry.delete(0, tk.END)
        line_name_entry.delete(0, tk.END)
        metro_name_entry.delete(0, tk.END)
        line_color_entry.delete(0, tk.END)
        line_year_entry.delete(0, tk.END)
        city_metro_entry.delete(0, tk.END)
        metro_year_entry.delete(0, tk.END)

    def apply(current):
        """
        Берет значения фильтров из полей ввода,
        составляет query-запрос, показывает отфильтрованную таблицу
        НЕ меняет саму таблицу в базе данных, показывает отображение
        """
        global current_filtered
        if current == 0: # Фильтруем таблицу "станции"
            global len_1
            try:
                total_filter = ''
                name_filter = sn_entry.get()
                if name_filter != '':
                    total_filter += 'Название {}'.format(name_filter)
                else:
                    total_filter += '(Название == Название)'
                for item in filters.get_children():
                    if filters.index(item) == 0:
                        filters.item(item, values = ["Название станции", name_filter])
                code_filter = lc_entry.get()
                if code_filter != '':
                    total_filter += ' and (`Код линии` {})'.format(code_filter)
                for item in filters.get_children():
                    if filters.index(item) == 1:
                        filters.item(item, values = ["Код линии", code_filter])
                depth_filter = depth_entry.get()
                if depth_filter != '':
                    total_filter += ' and (`Глубина заложения` {})'.format(depth_filter)
                for item in filters.get_children():
                    if filters.index(item) == 2:
                        filters.item(item, values = ["Глубина заложения", depth_filter])
                type_filter = type_entry.get()
                if type_filter != '':
                    total_filter += ' and (Тип {})'.format(type_filter)
                for item in filters.get_children():
                    if filters.index(item) == 3:
                        filters.item(item, values = ["Тип станции", type_filter])
                date_filter = open_date_entry.get()
                if date_filter != '':
                    for item in filters.get_children():
                        if filters.index(item) == 4:
                            filters.item(item, values = ["Дата открытия станции", date_filter])
                    #print(date_filter)
                    date_filter = date_filter.split(" ")
                    date_filter[1] = pd.to_datetime(date_filter[1], format ='%d-%m-%Y')
                    #print(date_filter)
                    total_filter += ' and (`Дата открытия` {} '.format(date_filter[0])
                    total_filter += '@date_filter[1])'
                else:
                    for item in filters.get_children():
                        if filters.index(item) == 4:
                            filters.item(item, values = ["Дата открытия станции", date_filter])
                filtered = database["stations"].query(total_filter)
                current_filtered = filtered
                for i in table_1.get_children():
                    table_1.delete(i)
                populate_table(0, table_1, filtered)
                len_1 = len(filtered)
                manage_open(None)
                configuration_page.destroy()
                edit_item = filters.identify_element(0, 1)
                filters.item(edit_item, text = name_filter)
                return
            except:
                showerror('Ошибка фильтра', 'Фильтр был заполнен неверно')
                for item in filters.get_children():
                    filters.delete(item)
                filters.insert('', 0, values = ['Год основания метро', ''])
                filters.insert('', 0, values = ['Город', ''])
                filters.insert('', 0, values = ['Метрополитен', ''])
                filters.insert('', 0, values = ['Цвет', ''])
                filters.insert('', 0, values = ['Название линии', ''])
                filters.insert('', 0, values = ['Год открытия линии', ''])
                filters.insert('', 0, values = ['Дата открытия станции', ''])
                filters.insert('', 0, values = ['Тип станции', ''])
                filters.insert('', 0, values = ['Глубина заложения', ''])
                filters.insert('', 0, values = ['Код линии', ''])
                filters.insert('', 0, values = ['Название станции', ''])
                configuration_page.destroy()
                return
        if current == 1: # Фильтруем таблицу "Линии"
            global len_2
            try:
                total_filter = ''
                code_filter = lc_entry.get()
                if code_filter != '':
                    total_filter += '`Код линии` {}'.format(code_filter)
                else:
                    total_filter += '`Код линии` == `Код линии`'
                for item in filters.get_children():
                    if filters.index(item) == 1:
                        filters.item(item, values = ["Код линии", code_filter])
                name_filter = line_name_entry.get()
                if name_filter != '':
                    total_filter += ' and Название {}'.format(name_filter)
                for item in filters.get_children():
                    if filters.index(item) == 6:
                        filters.item(item, values = ["Название линии", name_filter])
                metro_filter = metro_name_entry.get()
                if metro_filter != '':
                    total_filter += ' and Метрополитен {}'.format(metro_filter)
                for item in filters.get_children():
                    if filters.index(item) == 8:
                        filters.item(item, values = ["Метрополитен", metro_filter])
                color_filter = line_color_entry.get()
                if color_filter != '':
                    total_filter += ' and Цвет {}'.format(color_filter)
                for item in filters.get_children():
                    if filters.index(item) == 7:
                        filters.item(item, values = ["Цвет", color_filter])
                year_filter = line_year_entry.get()
                if year_filter != '':
                    total_filter += 'and `Год открытия линии` {}'.format(year_filter)
                for item in filters.get_children():
                    if filters.index(item) == 5:
                        filters.item(item, values = ["Год открытия линии", year_filter])
                filtered = database["lines"].query(total_filter)
                current_filtered = filtered
                for i in table_2.get_children():
                    table_2.delete(i)
                populate_table(0, table_2, filtered)
                len_2 = len(filtered)
                manage_open(None)
                configuration_page.destroy()
                return
            except:
                showerror('Ошибка фильтра', 'Фильтр был заполнен неверно')
                for item in filters.get_children():
                    filters.delete(item)
                filters.insert('', 0, values = ['Год основания метро', ''])
                filters.insert('', 0, values = ['Город', ''])
                filters.insert('', 0, values = ['Метрополитен', ''])
                filters.insert('', 0, values = ['Цвет', ''])
                filters.insert('', 0, values = ['Название линии', ''])
                filters.insert('', 0, values = ['Год открытия линии', ''])
                filters.insert('', 0, values = ['Дата открытия станции', ''])
                filters.insert('', 0, values = ['Тип станции', ''])
                filters.insert('', 0, values = ['Глубина заложения', ''])
                filters.insert('', 0, values = ['Код линии', ''])
                filters.insert('', 0, values = ['Название станции', ''])
                configuration_page.destroy()
                return
        if current == 2: # Фильтруем таблицу "Метрополитены"
            global len_3
            try:
                total_filter = ''
                metro_filter = metro_name_entry.get()
                if metro_filter != '':
                    total_filter += 'Метрополитен {}'.format(metro_filter)
                else:
                    total_filter += 'Метрополитен == Метрополитен'
                for item in filters.get_children():
                    if filters.index(item) == 8:
                        filters.item(item, values = ["Метрополитен", metro_filter])
                city_filter = city_metro_entry.get()
                if city_filter != '':
                    total_filter += ' and Город {}'.format(city_filter)
                for item in filters.get_children():
                    if filters.index(item) == 9:
                        filters.item(item, values = ["Город", city_filter])
                year_filter = metro_year_entry.get()
                if year_filter != '':
                    total_filter += ' and `Год основания метро` {}'.format(year_filter)
                for item in filters.get_children():
                    if filters.index(item) == 10:
                        filters.item(item, values = ["Год основания метро", year_filter])
                filtered = database['metros'].query(total_filter)
                current_filtered = filtered
                for i in table_3.get_children():
                    table_3.delete(i)
                populate_table(0, table_3, filtered)
                len_3 = len(filtered)
                manage_open(None)
                configuration_page.destroy()
                return
            except:
                showerror('Ошибка фильтра', 'Фильтр был заполнен неверно')
                for item in filters.get_children():
                    filters.delete(item)
                filters.insert('', 0, values = ['Год основания метро', ''])
                filters.insert('', 0, values = ['Город', ''])
                filters.insert('', 0, values = ['Метрополитен', ''])
                filters.insert('', 0, values = ['Цвет', ''])
                filters.insert('', 0, values = ['Название линии', ''])
                filters.insert('', 0, values = ['Год открытия линии', ''])
                filters.insert('', 0, values = ['Дата открытия станции', ''])
                filters.insert('', 0, values = ['Тип станции', ''])
                filters.insert('', 0, values = ['Глубина заложения', ''])
                filters.insert('', 0, values = ['Код линии', ''])
                filters.insert('', 0, values = ['Название станции', ''])
                configuration_page.destroy()
                return
        if current == 3: # Фильтруем таблицу "Соединение таблиц"
            global len_4
            try:
                total_filter = ''
                name_filter = sn_entry.get()
                if name_filter != '':
                    total_filter += '`Название станции` {}'.format(name_filter)
                else:
                    total_filter += '`Название станции` == `Название станции`'
                for item in filters.get_children():
                    if filters.index(item) == 0:
                        filters.item(item, values = ["Название станции", name_filter])
                code_filter = lc_entry.get()
                if code_filter != '':
                    total_filter += ' and `Код линии` {}'.format(code_filter)
                for item in filters.get_children():
                    if filters.index(item) == 1:
                        filters.item(item, values = ["Код линии", code_filter])
                depth_filter = depth_entry.get()
                if depth_filter != '':
                    total_filter += ' and `Глубина заложения` {}'.format(depth_filter)
                for item in filters.get_children():
                    if filters.index(item) == 2:
                        filters.item(item, values = ["Глубина заложения", depth_filter])
                type_filter = type_entry.get()
                if type_filter != '':
                    total_filter += ' and Тип {}'.format(depth_filter)
                for item in filters.get_children():
                    if filters.index(item) == 3:
                        filters.item(item, values = ["Тип станции", type_filter])
                station_open_date_filter = open_date_entry.get()
                if station_open_date_filter != '':
                    for item in filters.get_children():
                        if filters.index(item) == 4:
                            filters.item(item, values = ["Дата открытия станции", station_open_date_filter])
                    station_open_date_filter = station_open_date_filter.split(" ")
                    station_open_date_filter[1] = pd.to_datetime(
                        station_open_date_filter[1], format ='%d-%m-%Y'
                    )
                    #print(station_open_date_filter)
                    total_filter += ' and (`Дата открытия станции` {} '.format(
                        station_open_date_filter[0]
                    )
                    total_filter += '@station_open_date_filter[1])'
                else:
                    for item in filters.get_children():
                        if filters.index(item) == 4:
                            filters.item(item, values = ["Дата открытия станции", station_open_date_filter])
                line_name_filter = line_name_entry.get()
                if line_name_filter != '':
                    total_filter += ' and `Название линии` {}'.format(line_name_filter)
                for item in filters.get_children():
                    if filters.index(item) == 6:
                        filters.item(item, values = ["Название линии", line_name_filter])
                metro_filter = metro_name_entry.get()
                if metro_filter != '':
                    total_filter += ' and Метрополитен {}'.format(metro_filter)
                for item in filters.get_children():
                    if filters.index(item) == 8:
                        filters.item(item, values = ["Метрополитен", metro_filter])
                color_filter = line_color_entry.get()
                if color_filter != '':
                    total_filter += ' and `Цвет линии` {}'.format(color_filter)
                for item in filters.get_children():
                    if filters.index(item) == 7:
                        filters.item(item, values = ["Цвет", color_filter])
                line_year_filter = line_year_entry.get()
                if line_year_filter != '':
                    total_filter += 'and `Год открытия линии` {}'.format(line_year_filter)
                for item in filters.get_children():
                    if filters.index(item) == 5:
                        filters.item(item, values = ["Год открытия линии", line_year_filter])
                filtered = database['merged'].query(total_filter)
                current_filtered = filtered
                for i in table_4.get_children():
                    table_4.delete(i)
                populate_table(0, table_4, filtered)
                len_4 = len(filtered)
                manage_open(None)
                configuration_page.destroy()
                return
            except:
                showerror('Ошибка фильтра', 'Фильтр был заполнен неверно')
                for item in filters.get_children():
                    filters.delete(item)
                filters.insert('', 0, values = ['Год основания метро', ''])
                filters.insert('', 0, values = ['Город', ''])
                filters.insert('', 0, values = ['Метрополитен', ''])
                filters.insert('', 0, values = ['Цвет', ''])
                filters.insert('', 0, values = ['Название линии', ''])
                filters.insert('', 0, values = ['Год открытия линии', ''])
                filters.insert('', 0, values = ['Дата открытия станции', ''])
                filters.insert('', 0, values = ['Тип станции', ''])
                filters.insert('', 0, values = ['Глубина заложения', ''])
                filters.insert('', 0, values = ['Код линии', ''])
                filters.insert('', 0, values = ['Название станции', ''])
                configuration_page.destroy()
                return


    configuration_page = tk.Toplevel(app)
    configuration_page.geometry('300x400')
    configuration_page.resizable(False, False)
    configuration_page.title('Настроить фильтры')

    parameters = tk.LabelFrame(configuration_page, text = 'Фильтры')
    parameters.pack(anchor = 'nw', expand = True, fill = tk.BOTH)

    station_name = tk.Label(parameters, text = '{: ^25}'.format('Название станции'))
    station_name.grid(row = 0, column = 0)
    sn_entry = tk.Entry(parameters, width = 22)
    sn_entry.grid(row = 0, column = 1, pady = 5)

    line_code = tk.Label(parameters, text = '{: ^25}'.format('Код линии'))
    line_code.grid(row = 2, column = 0)
    lc_entry = tk.Entry(parameters, width = 22)
    lc_entry.grid(row = 2, column = 1, pady = 5)

    depth = tk.Label(parameters, text = '{: ^25}'.format('Глубина заложения'))
    depth.grid(row = 3, column = 0)
    depth_entry = tk.Entry(parameters, width = 22)
    depth_entry.grid(row = 3, column = 1, pady = 5)

    type = tk.Label(parameters, text = '{: ^25}'.format('Тип станции'))
    type.grid(row = 4, column = 0)
    type_entry = tk.Entry(parameters, width = 22)
    type_entry.grid(row = 4, column = 1, pady = 5)

    open_date = tk.Label(parameters, text = '{: ^25}'.format('Дата открытия станции'))
    open_date.grid(row = 5, column = 0)
    open_date_entry = tk.Entry(parameters, width = 22)
    open_date_entry.grid(row = 5, column = 1, pady = 5)

    line_name = tk.Label(parameters, text = '{: ^25}'.format('Название линии'))
    line_name.grid(row = 7, column = 0)
    line_name_entry = tk.Entry(parameters, width = 22)
    line_name_entry.grid(row = 7, column = 1, pady = 5)

    metro_name = tk.Label(parameters, text = '{: ^25}'.format('Метрополитен'))
    metro_name.grid(row = 8, column = 0)
    metro_name_entry = tk.Entry(parameters, width = 22)
    metro_name_entry.grid(row = 8, column = 1, pady = 5)

    line_color = tk.Label(parameters, text = '{: ^25}'.format('Цвет линии'))
    line_color.grid(row = 9, column = 0)
    line_color_entry = tk.Entry(parameters, width = 22)
    line_color_entry.grid(row = 9, column = 1, pady = 5)

    line_year = tk.Label(parameters, text = '{: ^25}'.format('Год открытия линии'))
    line_year.grid(row = 10, column = 0)
    line_year_entry = tk.Entry(parameters, width = 22)
    line_year_entry.grid(row = 10, column = 1, pady = 5)

    city_metro = tk.Label(parameters, text = '{: ^25}'.format('Город'))
    city_metro.grid(row = 11, column = 0)
    city_metro_entry = tk.Entry(parameters, width = 22)
    city_metro_entry.grid(row = 11, column = 1, pady = 5)

    metro_year = tk.Label(parameters, text = '{: ^25}'.format('Год основания метро'))
    metro_year.grid(row = 12, column = 0)
    metro_year_entry = tk.Entry(parameters, width = 22)
    metro_year_entry.grid(row = 12, column = 1, pady = 5)

    clear_entries = tk.Button(
        parameters, text = 'Очистить поля', width = 15,
        command = clear_all_entries
    )
    clear_entries.grid(row = 13, column = 0, pady = 15, padx = 10)

    create_instance = tk.Button(
        parameters, text = 'Применить', width = 15,
        command = lambda: apply(current)
    )
    create_instance.grid(row = 13, column = 1, pady = 15, padx = 10)

    """
    Активность/неактивность полей зависит от открытой в данный момент таблицы
    """
    if current == 0:
        station_name["state"] = "normal"
        sn_entry["state"] = "normal"
        line_code["state"] = "normal"
        lc_entry["state"] = "normal"
        depth["state"] = "normal"
        depth_entry["state"] = "normal"
        type["state"] = "normal"
        type_entry["state"] = "normal"
        open_date["state"] = "normal"
        open_date_entry["state"] = "normal"
        line_name["state"] = "disabled"
        line_name_entry["state"] = "disabled"
        metro_name["state"] = "disabled"
        metro_name_entry["state"] = "disabled"
        line_color["state"] = "disabled"
        line_color_entry["state"] = "disabled"
        line_year["state"] = "disabled"
        line_year_entry["state"] = "disabled"
        city_metro["state"] = "disabled"
        city_metro_entry["state"] = "disabled"
        metro_year["state"] = "disabled"
        metro_year_entry["state"] = "disabled"
    if current == 1:
        station_name["state"] = "disabled"
        sn_entry["state"] = "disabled"
        line_code["state"] = "normal"
        lc_entry["state"] = "normal"
        depth["state"] = "disabled"
        depth_entry["state"] = "disabled"
        type["state"] = "disabled"
        type_entry["state"] = "disabled"
        open_date["state"] = "disabled"
        open_date_entry["state"] = "disabled"
        line_name["state"] = "normal"
        line_name_entry["state"] = "normal"
        metro_name["state"] = "normal"
        metro_name_entry["state"] = "normal"
        line_color["state"] = "normal"
        line_color_entry["state"] = "normal"
        line_year["state"] = "normal"
        line_year_entry["state"] = "normal"
        city_metro["state"] = "disabled"
        city_metro_entry["state"] = "disabled"
        metro_year["state"] = "disabled"
        metro_year_entry["state"] = "disabled"
    if current == 2:
        station_name["state"] = "disabled"
        sn_entry["state"] = "disabled"
        line_code["state"] = "disabled"
        lc_entry["state"] = "disabled"
        depth["state"] = "disabled"
        depth_entry["state"] = "disabled"
        type["state"] = "disabled"
        type_entry["state"] = "disabled"
        open_date["state"] = "disabled"
        open_date_entry["state"] = "disabled"
        line_name["state"] = "disabled"
        line_name_entry["state"] = "disabled"
        metro_name["state"] = "normal"
        metro_name_entry["state"] = "normal"
        line_color["state"] = "disabled"
        line_color_entry["state"] = "disabled"
        line_year["state"] = "disabled"
        line_year_entry["state"] = "disabled"
        city_metro["state"] = "normal"
        city_metro_entry["state"] = "normal"
        metro_year["state"] = "normal"
        metro_year_entry["state"] = "normal"
    if current == 3:
        city_metro["state"] = "disabled"
        city_metro_entry["state"] = "disabled"
        metro_year["state"] = "disabled"
        metro_year_entry["state"] = "disabled"


def analysis_function(current):
    """
    Вызывается кнопкой "Показать". Отвечает за инструменты анализа
    Параметр функции - индекс открытой таблицы
    """
    tables = {0: "stations", 1: "lines", 2: "metros", 3: "merged"}

    def check_possibility(list_of_cols, possible):
        """
        Проверяет, что все значения из списка list_of_cols есть в множестве
        possible
        """
        for element in list_of_cols:
            if element not in possible:
                return False
        return True

    def insert_columns(current):
        """
        Вставляет в меню выбора колонок варианты в зависимости от открытой
        таблицы
        """
        if current == 0:
            col_selection.insert('', 0, values = ['Название'])
            col_selection.insert('', 0, values = ['Код линии'])
            col_selection.insert('', 0, values = ['Глубина заложения'])
            col_selection.insert('', 0, values = ['Тип'])
        if current == 1:
            col_selection.insert('', 0, values = ['Код линии'])
            col_selection.insert('', 0, values = ['Название'])
            col_selection.insert('', 0, values = ['Метрополитен'])
            col_selection.insert('', 0, values = ['Цвет'])
            col_selection.insert('', 0, values = ['Год открытия линии'])
        if current == 2:
            col_selection.insert('', 0, values = ['Метрополитен'])
            col_selection.insert('', 0, values = ['Город'])
            col_selection.insert('', 0, values = ['Год основания метро'])
        if current == 3:
            col_selection.insert('', 0, values = ['Название станции'])
            col_selection.insert('', 0, values = ['Код линии'])
            col_selection.insert('', 0, values = ['Глубина заложения'])
            col_selection.insert('', 0, values = ['Тип'])
            col_selection.insert('', 0, values = ['Название линии'])
            col_selection.insert('', 0, values = ['Метрополитен'])
            col_selection.insert('', 0, values = ['Цвет'])
            col_selection.insert('', 0, values = ['Год открытия линии'])

    def select_cols_to_process():
        """
        Выбирает колонки для работы. Вызывается кнопкой "Выбрать".
        Также выполняет работу выбранного инструмента анализа
        """
        selected_cols = []
        for element in col_selection.selection():
            selected_cols.append(col_selection.item(element)["values"])
        select_column.destroy()
        if selected[0] == 'I005': # Строим диаграмму размаха
            selected_cols_2 = [] # Список названий выбранных колонок
            for element in selected_cols:
                selected_cols_2.append(element[0])
            if len(selected_cols_2) not in [1, 2]:
                showerror(
                    'Ошибка выбора',
                    'Для диаграммы размаха необходимо выбрать один столбец c числовыми данными\
(допускается также столбец с данными о годе), либо столбец с числовыми и столбец с категориальными для группировки'
                )
                return
            elif len(selected_cols_2) == 1 and selected_cols_2[0] not in ['Глубина заложения', 'Год открытия линии', 'Год основания метро']:
                showerror(
                    'Ошибка выбора',
                    'Для диаграммы размаха необходимо выбрать один столбец c числовыми данными \
(допускается также столбец с данными о годе), либо столбец с числовыми и столбец с категориальными для группировки'
                )
                return
            elif len(selected_cols_2) == 1:
                fig, axis = plt.subplots(1, 1)
                sns.boxplot(
                    x = database[tables[current]][selected_cols_2[0]], ax = axis
                )
                plt.show()
            else:
                try:
                    fig, axis = plt.subplots(1, 1)
                    sns.boxplot(
                        data = database[tables[current]], ax = axis,
                        y = selected_cols_2[0],
                        x = selected_cols_2[1]
                    )
                    plt.show()
                except:
                    showerror(
                        'Ошибка выбора',
                        'Для диаграммы размаха необходимо выбрать один столбец c числовыми данными \
(допускается также столбец с данными о годе), либо столбец с числовыми и столбец с категориальными для группировки'
                    )
                    return
        if selected[0] == 'I003': # Строим столбчатую диаграмму.
            selected_cols_2 = [] # Список названий выбранных колонок
            for element in selected_cols:
                selected_cols_2.append(element[0])
            if len(selected_cols_2) not in [1, 2] or \
            not check_possibility(selected_cols_2, ['Код линии', 'Тип', 'Цвет', 'Метрополитен', 'Город']):
                showerror(
                    'Ошибка выбора',
                    'Для столбчатой диаграммы необходимо выбрать от одного до двух столбцов c категориальными данными'
                )
                return
            elif len(selected_cols_2) == 1:
                fig, axis = plt.subplots(1, 1)
                sns.countplot(
                    x = selected_cols_2[0], data = database[tables[current]],
                    ax = axis
                )
                axis.set_ylabel('Количество записей')
                plt.show()
            elif len(selected_cols_2) == 2:
                fig, axis = plt.subplots(1, 1)
                sns.countplot(
                    x = selected_cols_2[1], data = database[tables[current]],
                    ax = axis, hue = selected_cols_2[0]
                )
                plt.legend(loc = 'upper right')
                plt.show()
        if selected[0] == 'I002': # Строим график плотности распределения
            selected_col = selected_cols[0][0]
            if len(selected_cols) != 1 or \
            selected_col not in ['Глубина заложения', 'Год открытия линии', 'Год основания метро']:
                showerror(
                    'Ошибка выбора',
                    'Для графика плотности необходимо выбрать строго один столбец c числовыми данными'
                )
                return
            else:
                fig, axis = plt.subplots(1, 1)
                sns.distplot(
                    database[tables[current]][selected_col],
                    ax = axis
                )
                plt.show()
        if selected[0] == 'I004': # Строим гистограмму
            selected_col = selected_cols[0][0]
            if len(selected_cols) != 1 or \
            selected_col not in ['Глубина заложения', 'Год открытия линии', 'Год основания метро']:
                showerror(
                    'Ошибка выбора',
                    'Для гистограммы доступны только численные столбцы'
                )
                return
            else:
                fig, axis = plt.subplots(1, 1)
                database[tables[current]].hist(column = selected_col, ax = axis)
                plt.show()
        if selected[0] == 'I001': # Строим сводную таблицу
            selected_cols_2 = [] # Список названий выбранных колонок
            for element in selected_cols:
                selected_cols_2.append(element[0])
            if len(selected_cols_2) != 2:
                showerror(
                    'Ошибка выбора',
                    'Для сводной таблицы необходимо выбрать один численный и один категориальный столбец'
                )
                return
            try:
                pivot = pd.pivot_table(
                    database[tables[current]], index = selected_cols_2[1],
                    values = selected_cols_2[0],
                    aggfunc = ['mean', 'median', 'max', 'min'],
                    margins = True, margins_name = 'По всем группам'
                )
                pivot.columns = [
                    'Среднее значение',
                    'Медианное',
                    'Максимальное',
                    'Минимальное'
                ]
                pivot_window = tk.Toplevel(app)
                pivot_window.resizable(False, False)
                pivot_window.geometry('650x375')
                pivot_window.title('Сводная таблица')
                pivot_tree = ttk.Treeview(pivot_window, height = 10)
                pivot_tree['columns'] = [
                    'Категория',
                    'Среднее значение',
                    'Медианное',
                    'Максимальное',
                    'Минимальное'
                ]
                pivot_tree.heading("#0", text = "", anchor = 'w')
                pivot_tree.column("#0", anchor = 'w', width = 2, stretch = tk.NO)
                for col in pivot_tree['columns']:
                    pivot_tree.heading(col, text = col)
                    pivot_tree.column(col, anchor = 'w', width = 100)
                pivot = pivot.reset_index()
                pivot['Среднее значение'] = np.around(
                    pivot['Среднее значение'], decimals = 4
                )
                for i in range(len(pivot)):
                    vals = pivot.iloc[i, :].tolist()
                    pivot_tree.insert('', i, values = vals)
                vertical_scrollbar_pivot_tree = tk.Scrollbar(
                    pivot_tree, command = pivot_tree.yview
                )
                vertical_scrollbar_pivot_tree.pack(side = tk.RIGHT, fill = tk.Y)
                horizontal_scrollbar_pivot_tree = tk.Scrollbar(
                    pivot_tree, command = pivot_tree.xview,
                    orient = 'horizontal'
                )
                horizontal_scrollbar_pivot_tree.pack(side = tk.BOTTOM, fill = tk.X)
                export_shown = tk.Button(
                    pivot_window, text = 'Экспортировать', width = 20,
                    command = lambda: export_pivot(pivot)
                )
                export_shown.pack(side = tk.BOTTOM, padx = 10, pady = 10, fill = tk.X)
                pivot_tree.pack(fill = tk.BOTH, padx = 5, pady = 5, expand = True)
            except:
                try:
                    pivot = pd.pivot_table(
                        database[tables[current]], index = selected_cols_2[0],
                        values = selected_cols_2[1],
                        aggfunc = ['mean', 'median', 'max', 'min'],
                        margins = True, margins_name = 'По всем группам'
                        )
                    pivot.columns = [
                        'Среднее значение',
                        'Медианное',
                        'Максимальное',
                        'Минимальное'
                    ]
                    pivot_window = tk.Toplevel(app)
                    pivot_window.resizable(False, False)
                    pivot_window.geometry('650x250')
                    pivot_window.title('Сводная таблица')
                    pivot_tree = ttk.Treeview(pivot_window, height = 10)
                    pivot_tree['columns'] = [
                        'Категория',
                        'Среднее значение',
                        'Медианное',
                        'Максимальное',
                        'Минимальное'
                    ]
                    pivot_tree.heading("#0", text = "", anchor = 'w')
                    pivot_tree.column("#0", anchor = 'w', width = 2, stretch = tk.NO)
                    for col in pivot_tree['columns']:
                        pivot_tree.heading(col, text = col)
                        pivot_tree.column(col, anchor = 'w', width = 100)
                    pivot['Среднее значение'] = np.around(
                        pivot['Среднее значение'], decimals = 4
                    )
                    pivot = pivot.reset_index()
                    #print(pivot)
                    for i in range(len(pivot)):
                        vals = pivot.iloc[i, :].tolist()
                        pivot_tree.insert('', i, values = vals)
                    vertical_scrollbar_pivot_tree = tk.Scrollbar(
                        pivot_tree, command = pivot_tree.yview
                    )
                    vertical_scrollbar_pivot_tree.pack(side = tk.RIGHT, fill = tk.Y)
                    horizontal_scrollbar_pivot_tree = tk.Scrollbar(
                        pivot_tree, command = pivot_tree.xview,
                        orient = 'horizontal'
                    )
                    horizontal_scrollbar_pivot_tree.pack(side = tk.BOTTOM, fill = tk.X)
                    export_shown = tk.Button(
                        pivot_window, text = 'Экспортировать', width = 20,
                        command = lambda: export_pivot(pivot)
                    )
                    export_shown.pack(side = tk.BOTTOM, padx = 10, pady = 10, fill = tk.X)
                    pivot_tree.pack(fill = tk.BOTH, padx = 5, pady = 5, expand = True)
                except:
                    showerror(
                        'Ошибка выбора',
                        'Для сводной таблицы необходимо выбрать один численный и один категориальный столбец'
                    )
                    return

    selected = functions_treeview.selection()
    # Окно выбора столбцов для анализа
    select_column = tk.Toplevel(app)
    select_column.resizable(False, False)
    select_column.geometry('400x170')
    select_column.title('Выберете столбец/столбцы для работы')
    col_selection = ttk.Treeview(select_column, height = 4)
    col_selection['columns'] = ('Столбец')
    col_selection.heading("#0", text = "", anchor = 'w')
    col_selection.column("#0", anchor = 'w', width = 2, stretch = tk.NO)
    col_selection.heading('Столбец', text = 'Столбец', anchor = 'w')
    insert_columns(current)
    vertical_scrollbar_selcol = tk.Scrollbar(
        col_selection, command = col_selection.yview
    )
    vertical_scrollbar_selcol.pack(side = tk.RIGHT, fill = tk.Y)

    select = tk.Button(
        select_column, text = 'Выбрать',
        width = 23,
        command = lambda: select_cols_to_process()
    )

    col_selection.pack(fill = tk.BOTH, padx = 12, pady = 5, expand = True)
    select.pack(padx = 5, pady = 5)


def print_info():
    """
    Выводит в специальное поле информацию об открытой таблице
    """
    info_message['state'] = 'normal'
    current = tablayout.index((tablayout.select()))
    trees = {0: "stations", 1: "lines", 2: "metros", 3: "merged"}
    buf = io.StringIO()
    database[trees[current]].info(buf = buf)
    information = buf.getvalue()
    info_message.delete('1.0', tk.END)
    info_message.insert(tk.END, str(information))
    info_message['state'] = 'disabled'

def export_pivot(pivot):
    """
    Экспортирует сводную таблицу. Вызывается кнопкой в окне сводной таблицы
    """
    file_name = asksaveasfilename(filetypes = (('Excel Files (*.xlsx)', "*.xlsx*"), ))
    file_name += '.xlsx'
    pivot.to_excel(file_name, index = False)

def export_representation(current_filtered):
    """
    Экспортирует в формате Excel текущее отображение на экране.
    Вызывается кнопкой "Экспорт отображения"
    """
    file_name = asksaveasfilename(filetypes = (('Excel Files (*.xlsx)', "*.xlsx*"), ))
    file_name += '.xlsx'
    current_filtered.to_excel(file_name, index = False)

def export_selected(current):
    """
    Экспортирует выбранные строки в формате Excel
    в выбранную пользователем директорию. Вызывается кнопкой "Экспорт"
    """
    trees = {
        0: ('stations', table_1),
        1: ('lines', table_2),
        2: ('metros', table_3),
        3: ('merged', table_4)
    }
    file_name = asksaveasfilename(filetypes = (('Excel Files (*.xlsx)', "*.xlsx*"), ))
    file_name += '.xlsx'
    data = []
    for element in trees[current][1].selection():
        data.append(trees[current][1].item(element)["values"])
    dataframe = pd.DataFrame(data = data, columns = database[trees[current][0]].columns)
    dataframe.to_excel(file_name, index = False)

def delete_selected(tree):
    """
    Удаляет выбранные строки из базы
    Если строка содержит внешний ключ, на который ссылаются строки других таблиц,
    те тоже удаляются. Вызывается кнопкой "Удалить"
    """
    global database
    global current_filtered
    trees = {"stations": table_1, "lines": table_2, "metros": table_3}
    selected = []
    for element in trees[tree].selection(): # Сохраняем значения выбранных строк
        selected.append(trees[tree].item(element)["values"])
    if tree == 'stations': # Удаление станции из таблицы 1 (как следствие и из таблицы 4)
        global sel_1
        global sel_4
        for element in selected:
            ind = database[tree][
                (database[tree]['Название'] == element[0]) & (database[tree]['Код линии'] == element[1])
            ].index # Ищем индекс выбранной строки в датафрейме
            database[tree].drop(index = ind, inplace = True)
            database[tree].reset_index(drop = True, inplace = True)
            ind_merged = database['merged'][
                (database['merged']['Название станции'] == element[0]) & (database['merged']['Код линии'] == element[1])
            ].index # Ищем индекс в соединенном датафрейме
            database['merged'].drop(index = ind_merged, inplace = True)
            database['merged'].reset_index(drop = True, inplace = True)
        for i in table_1.get_children():
            table_1.delete(i)
        for i in table_4.get_children():
            table_4.delete(i)
        populate_table(0, table_1, database[tree])
        populate_table(3, table_4, database['merged'])
        sel_1 = 0
        sel_4 = 0
        current_filtered = database[tree]
        manage_open(None)
    if tree == 'lines': # Удаление линии и всех связанных станций
        global sel_2
        proceed = askyesno(
            title = 'Удаление данных',
            message = 'Удаление линии приведет к удалению всех связанных станций. Продолжить?'
        )
        if not proceed:
            return
        for element in selected:
            ind = database[tree][
                (database[tree]['Код линии'] == element[0])
            ].index # Ищем индекс выбранной записи в таблице о линиях
            database[tree].drop(index = ind, inplace = True)
            database[tree].reset_index(drop = True, inplace = True)
            ind_st =  database['stations'][
                (database['stations']['Код линии'] == element[0])
            ].index # Ищем индексы станций на удаляемой линии
            database['stations'].drop(index = ind_st, inplace = True)
            database['stations'].reset_index(drop = True, inplace = True)
            ind_merged = database['merged'][
                (database['merged']['Код линии'] == element[0])
            ].index # Ищем индексы удаляемых строк из соединенной таблицы
            database['merged'].drop(index = ind_merged, inplace = True)
            database['merged'].reset_index(drop = True, inplace = True)
        for i in table_1.get_children():
            table_1.delete(i)
        for i in table_2.get_children():
            table_2.delete(i)
        for i in table_4.get_children():
            table_4.delete(i)
        populate_table(0, table_1, database['stations'])
        populate_table(1, table_2, database[tree])
        populate_table(3, table_4, database['merged'])
        sel_2 = 0
        current_filtered = database[tree]
        manage_open(None)
    if tree == 'metros': # Удаление целых метрополитенов и всей связанной информации
        global sel_3
        proceed = askyesno(
            title = 'Удаление данных',
            message = 'Удаление метро приведет к удалению всех связанных линий и станций. Продолжить?'
        )
        if not proceed:
            return
        for element in selected:
            ind = database[tree][
                (database[tree]['Метрополитен'] == element[0])
            ].index # Ищем индекс выбранной строки в датафрейме
            database[tree].drop(index = ind, inplace = True)
            database[tree].reset_index(drop = True, inplace = True)
            ind_st = database['stations'][
                    database['stations']['Код линии'].isin(
                            database['lines'][
                                    database['lines']['Метрополитен'] == element[0]
                                ]['Код линии']
                        )
                ].index # Ищем индексы всех станций метрополитена
            database['stations'].drop(index = ind_st, inplace = True)
            database['stations'].reset_index(drop = True, inplace = True)
            ind_lines = database['lines'][database['lines']['Метрополитен'] == element[0]].index
            database['lines'].drop(index = ind_lines, inplace = True)
            database['lines'].reset_index(drop = True, inplace = True)
            ind_merged = database['merged'][database['merged']['Метрополитен'] == element[0]].index
            database['merged'].drop(index = ind_merged, inplace = True)
            database['merged'].reset_index(drop = True, inplace = True)
        for i in table_1.get_children():
            table_1.delete(i)
        for i in table_2.get_children():
            table_2.delete(i)
        for i in table_3.get_children():
            table_3.delete(i)
        for i in table_4.get_children():
            table_4.delete(i)
        populate_table(0, table_1, database['stations'])
        populate_table(1, table_2, database['lines'])
        populate_table(2, table_3, database[tree])
        populate_table(3, table_4, database['merged'])
        sel_3 = 0
        current_filtered = database[tree]
        manage_open(None)

def add_instance(tree):
    """
    Добавляет строку в таблицу. Вызывается кнопкой "Добавить"
    """

    global database
    current = tree

    def clear_all_entries():
        """
        Очищает все строки окна редактирования записи
        """
        sn_entry.delete(0, tk.END)
        lc_entry.delete(0, tk.END)
        depth_entry.delete(0, tk.END)
        type_entry.delete(0, tk.END)
        open_date_entry.delete(0, tk.END)
        line_name_entry.delete(0, tk.END)
        metro_name_entry.delete(0, tk.END)
        line_color_entry.delete(0, tk.END)
        line_year_entry.delete(0, tk.END)
        city_metro_entry.delete(0, tk.END)
        metro_year_entry.delete(0, tk.END)

    def create_new_instance(current):
        """
        Собирает строку датафрейма из введенных
        в окно редактирования строка. Добавляет строку в базу.
        Все правила целостности соблюдаются, в случае нарушения
        функция не добавляет запись, а выводит ошибку и завершает работу
        """
        global current_filtered
        if current == 2: # Добавление нового метрополитена
            new_metro_name = metro_name_entry.get()
            new_metro_city = city_metro_entry.get()
            try:
                new_metro_year = int(metro_year_entry.get())
            except:
                showerror('Ошибка записи', 'Год должен был целым числом')
                new_instance_window.destroy()
                return
            if new_metro_name in list(database['metros']['Метрополитен']):
                showerror('Ошибка записи', 'Метрополитен с таким именем уже существует')
                new_instance_window.destroy()
            else:
                database['metros'].loc[-1] = [new_metro_name, new_metro_city, new_metro_year]
                database['metros'].index = database['metros'].index + 1
                database['metros'] = database['metros'].sort_index()
                for i in table_3.get_children():
                    table_3.delete(i)
                populate_table(2, table_3, database['metros'])
        if current == 1: # Добавление новой линии
            new_line_code = lc_entry.get()
            new_line_name = line_name_entry.get()
            new_line_metro = metro_name_entry.get()
            new_line_color = line_color_entry.get()
            try:
                new_line_year = int(line_year_entry.get())
            except:
                showerror('Ошибка записи', 'Год должен был целым числом')
                new_instance_window.destroy()
                return
            if new_line_code in list(database['lines']['Код линии']):
                showerror('Ошибка записи', 'Линия с таким кодом уже существует')
                new_instance_window.destroy()
                return
            if new_line_metro not in list(database['metros']['Метрополитен']):
                showerror('Ошибка записи', 'Указанного метрополитена не существует')
                new_instance_window.destroy()
            else:
                database['lines'].loc[-1] = [
                    new_line_code, new_line_name, new_line_metro, new_line_color,
                    new_line_year
                ]
                database['lines'].index = database['lines'].index + 1
                database['lines'] = database['lines'].sort_index()
                for i in table_2.get_children():
                    table_2.delete(i)
                populate_table(1, table_2, database['lines'])
        if current == 0: # Добавление новой станции
            new_station_name = sn_entry.get()
            new_station_line = lc_entry.get()
            try:
                new_station_depth = float(depth_entry.get())
            except:
                showerror('Ошибка записи', 'Глубина должна быть числом')
                new_instance_window.destroy()
                return
            if new_station_depth < 0:
                showerror('Ошибка записи', 'Глубина должна быть неотрицательной')
                new_instance_window.destroy()
                return
            new_station_type = type_entry.get()
            try:
                new_station_date = pd.to_datetime(
                    open_date_entry.get(), format ='%d-%m-%Y'
                )
                new_station_date = new_station_date.date()
            except:
                showerror(
                    'Ошибка записи',
                    'Введите корректную дату в формате дд-мм-гг'
                )
                new_instance_window.destroy()
                return
            if new_station_line not in list(database['lines']['Код линии']):
                showerror(
                    'Ошибка записи',
                    'Указанной линии не существует'
                )
                new_instance_window.destroy()
                return
            if new_station_name in list(database['stations']['Название']) \
            and new_station_line in list(database['stations']['Код линии']):
                showerror(
                    'Ошибка записи',
                    'Станция уже существует'
                )
                new_instance_window.destroy()
                return
            else:
                database['stations'].loc[-1] = [
                    new_station_name, new_station_line, new_station_depth,
                    new_station_type, new_station_date
                ]
                database['stations'].index = database['stations'].index + 1
                database['stations'] = database['stations'].sort_index()
                for i in table_1.get_children():
                    table_1.delete(i)
                populate_table(0, table_1, database['stations'])
                columns = database['merged'].columns
                database['merged'] = database['stations'].merge(
                    database['lines'], on = 'Код линии'
                ).merge(
                    database['metros'], on = 'Метрополитен'
                    ).drop(
                        ['Город', 'Год основания метро'], axis = 1
                        )
                database['merged'].columns = columns
                for i in table_4.get_children():
                    table_4.delete(i)
                populate_table(3, table_4, database['merged'])

                manage_open(None)

    """
    Окно редактирования записи
    """
    new_instance_window = tk.Toplevel(app)
    new_instance_window.resizable(False, False)
    new_instance_window.geometry('300x400')
    new_instance_window.title('Добавить запись')

    parameters = tk.LabelFrame(new_instance_window, text = 'Параметры')
    parameters.pack(anchor = 'nw', expand = True, fill = tk.BOTH)

    station_name = tk.Label(parameters, text = '{: ^25}'.format('Название станции'))
    station_name.grid(row = 0, column = 0)
    sn_entry = tk.Entry(parameters, width = 22)
    sn_entry.grid(row = 0, column = 1, pady = 5)

    line_code = tk.Label(parameters, text = '{: ^25}'.format('Код линии'))
    line_code.grid(row = 2, column = 0)
    lc_entry = tk.Entry(parameters, width = 22)
    lc_entry.grid(row = 2, column = 1, pady = 5)

    depth = tk.Label(parameters, text = '{: ^25}'.format('Глубина заложения'))
    depth.grid(row = 3, column = 0)
    depth_entry = tk.Entry(parameters, width = 22)
    depth_entry.grid(row = 3, column = 1, pady = 5)

    type = tk.Label(parameters, text = '{: ^25}'.format('Тип станции'))
    type.grid(row = 4, column = 0)
    type_entry = tk.Entry(parameters, width = 22)
    type_entry.grid(row = 4, column = 1, pady = 5)

    open_date = tk.Label(parameters, text = '{: ^25}'.format('Дата открытия станции'))
    open_date.grid(row = 5, column = 0)
    open_date_entry = tk.Entry(parameters, width = 22)
    open_date_entry.grid(row = 5, column = 1, pady = 5)

    line_name = tk.Label(parameters, text = '{: ^25}'.format('Название линии'))
    line_name.grid(row = 7, column = 0)
    line_name_entry = tk.Entry(parameters, width = 22)
    line_name_entry.grid(row = 7, column = 1, pady = 5)

    metro_name = tk.Label(parameters, text = '{: ^25}'.format('Метрополитен'))
    metro_name.grid(row = 8, column = 0)
    metro_name_entry = tk.Entry(parameters, width = 22)
    metro_name_entry.grid(row = 8, column = 1, pady = 5)

    line_color = tk.Label(parameters, text = '{: ^25}'.format('Цвет линии'))
    line_color.grid(row = 9, column = 0)
    line_color_entry = tk.Entry(parameters, width = 22)
    line_color_entry.grid(row = 9, column = 1, pady = 5)

    line_year = tk.Label(parameters, text = '{: ^25}'.format('Год открытия линии'))
    line_year.grid(row = 10, column = 0)
    line_year_entry = tk.Entry(parameters, width = 22)
    line_year_entry.grid(row = 10, column = 1, pady = 5)

    city_metro = tk.Label(parameters, text = '{: ^25}'.format('Город'))
    city_metro.grid(row = 11, column = 0)
    city_metro_entry = tk.Entry(parameters, width = 22)
    city_metro_entry.grid(row = 11, column = 1, pady = 5)

    metro_year = tk.Label(parameters, text = '{: ^25}'.format('Год основания метро'))
    metro_year.grid(row = 12, column = 0)
    metro_year_entry = tk.Entry(parameters, width = 22)
    metro_year_entry.grid(row = 12, column = 1, pady = 5)

    clear_entries = tk.Button(
        parameters, text = 'Очистить поля', width = 15,
        command = clear_all_entries
    )
    clear_entries.grid(row = 13, column = 0, pady = 15, padx = 10)

    create_instance = tk.Button(
        parameters, text = 'Создать запись', width = 15,
        command = lambda: create_new_instance(current)
    )
    create_instance.grid(row = 13, column = 1, pady = 15, padx = 10)

    """
    Активность/неактивность полей зависит от открытой в данный момент таблицы
    """
    if current == 0:
        station_name["state"] = "normal"
        sn_entry["state"] = "normal"
        line_code["state"] = "normal"
        lc_entry["state"] = "normal"
        depth["state"] = "normal"
        depth_entry["state"] = "normal"
        type["state"] = "normal"
        type_entry["state"] = "normal"
        open_date["state"] = "normal"
        open_date_entry["state"] = "normal"
        line_name["state"] = "disabled"
        line_name_entry["state"] = "disabled"
        metro_name["state"] = "disabled"
        metro_name_entry["state"] = "disabled"
        line_color["state"] = "disabled"
        line_color_entry["state"] = "disabled"
        line_year["state"] = "disabled"
        line_year_entry["state"] = "disabled"
        city_metro["state"] = "disabled"
        city_metro_entry["state"] = "disabled"
        metro_year["state"] = "disabled"
        metro_year_entry["state"] = "disabled"
    if current == 1:
        station_name["state"] = "disabled"
        sn_entry["state"] = "disabled"
        line_code["state"] = "normal"
        lc_entry["state"] = "normal"
        depth["state"] = "disabled"
        depth_entry["state"] = "disabled"
        type["state"] = "disabled"
        type_entry["state"] = "disabled"
        open_date["state"] = "disabled"
        open_date_entry["state"] = "disabled"
        line_name["state"] = "normal"
        line_name_entry["state"] = "normal"
        metro_name["state"] = "normal"
        metro_name_entry["state"] = "normal"
        line_color["state"] = "normal"
        line_color_entry["state"] = "normal"
        line_year["state"] = "normal"
        line_year_entry["state"] = "normal"
        city_metro["state"] = "disabled"
        city_metro_entry["state"] = "disabled"
        metro_year["state"] = "disabled"
        metro_year_entry["state"] = "disabled"
    if current == 2:
        station_name["state"] = "disabled"
        sn_entry["state"] = "disabled"
        line_code["state"] = "disabled"
        lc_entry["state"] = "disabled"
        depth["state"] = "disabled"
        depth_entry["state"] = "disabled"
        type["state"] = "disabled"
        type_entry["state"] = "disabled"
        open_date["state"] = "disabled"
        open_date_entry["state"] = "disabled"
        line_name["state"] = "disabled"
        line_name_entry["state"] = "disabled"
        metro_name["state"] = "normal"
        metro_name_entry["state"] = "normal"
        line_color["state"] = "disabled"
        line_color_entry["state"] = "disabled"
        line_year["state"] = "disabled"
        line_year_entry["state"] = "disabled"
        city_metro["state"] = "normal"
        city_metro_entry["state"] = "normal"
        metro_year["state"] = "normal"
        metro_year_entry["state"] = "normal"

def edit_instance(current):
    """
    Редактирует строку в базе. Вызывается кнопкой "Править".
    Правила целостности соблюдены, в случае нарушения функция досрочно
    завершается сообщением об ошибке.
    При изменении ключевого атрибута он также меняется во всех таблицах,
    на этот атрибут ссылающихся
    """
    global database

    trees = {0: table_1, 1: table_2, 2: table_3}
    selected = trees[current].item(trees[current].selection())["values"]
    #print(selected)
    if current == 0: # Редактируем запись о станции
        selected_index = database['stations'][
                (database['stations']['Название'] == selected[0]) & \
                (database['stations']['Код линии'] == selected[1])
            ].index # Сохраняем индекс выбранной строки в датафрейме
    if current == 1: # Редактируем запись о линии
        selected_index = database['lines'][
                database['lines']['Код линии'] == selected[0]
            ].index # Сохраняем индекс выбранной строки в датафрейме
    if current == 2: # Редактируем запись о метрополитене
        selected_index = database['metros'][
                database['metros']['Метрополитен'] == selected[0]
            ].index

    def clear_all_entries():
        """
        Очищает все строки в полях окна редактирования
        """
        sn_entry.delete(0, tk.END)
        lc_entry.delete(0, tk.END)
        depth_entry.delete(0, tk.END)
        type_entry.delete(0, tk.END)
        open_date_entry.delete(0, tk.END)
        line_name_entry.delete(0, tk.END)
        metro_name_entry.delete(0, tk.END)
        line_color_entry.delete(0, tk.END)
        line_year_entry.delete(0, tk.END)
        city_metro_entry.delete(0, tk.END)
        metro_year_entry.delete(0, tk.END)

    def alter_row(current, selected_index):
        """
        Изменяет строку в базе. Вызывается кнопкой "Редактировать"
        окна редактировать. Проверяет легитимность правки.
        """
        if current == 0: # Редактируем запись о станции
            blacklist_names = set(database['stations']['Название'])
            blacklist_names.remove(selected[0])
            blacklist_lines = set(database['stations']['Код линии'])
            blacklist_lines.remove(selected[1])
            new_station_name = sn_entry.get()
            new_station_line = lc_entry.get()
            try:
                new_station_depth = float(depth_entry.get())
            except:
                showerror('Ошибка записи', 'Глубина должна быть числом')
                new_instance_window.destroy()
                return
            if new_station_depth < 0:
                showerror('Ошибка записи', 'Глубина должна быть неотрицательной')
                new_instance_window.destroy()
                return
            new_station_type = type_entry.get()
            try:
                new_station_date = pd.to_datetime(
                    open_date_entry.get(), format ='%d-%m-%Y'
                )
                new_station_date = new_station_date.date()
            except:
                try:
                    new_station_date = pd.to_datetime(
                        open_date_entry.get(), format ='%Y-%m-%d'
                    )
                    new_station_date = new_station_date.date()
                except:
                    showerror(
                        'Ошибка записи',
                        'Введите корректную дату в формате дд-мм-гг'
                    )
                    new_instance_window.destroy()
                    return
            if new_station_line not in list(database['lines']['Код линии']):
                showerror(
                    'Ошибка записи',
                    'Указанной линии не существует'
                )
                new_instance_window.destroy()
                return
            # Если поменяли линию и это имя уже есть на данной линии или если поменяли имя и на этой линии уже есть данное имя
            if \
            (new_station_line != selected[1] \
                and new_station_name in set(database['stations'][database['stations']['Код линии'] == new_station_line]['Название'])) \
            or \
            (new_station_name != selected[0] \
                and new_station_line in set(database['stations'][database['stations']['Название'] == new_station_name]['Код линии'])):
                showerror(
                    'Ошибка записи',
                    'Данная станция уже существует'
                    )
                new_instance_window.destroy()
                return
            else:
                database['stations'].loc[selected_index] = [
                    new_station_name, new_station_line, new_station_depth,
                    new_station_type, new_station_date
                    ] # Собирается новая строка
                for i in table_1.get_children():
                    table_1.delete(i)
                populate_table(0, table_1, database['stations'])
                columns = database['merged'].columns
                database['merged'] = database['stations'].merge(
                    database['lines'], on = 'Код линии'
                ).merge(
                    database['metros'], on = 'Метрополитен'
                    ).drop(
                        ['Город', 'Год основания метро'], axis = 1
                    )
                database['merged'].columns = columns
                for i in table_4.get_children():
                    table_4.delete(i)
                populate_table(3, table_4, database['merged'])
                manage_open(None)
                new_instance_window.destroy()
        if current == 1: # Редактируем запись о линии
            new_line_code = lc_entry.get()
            new_line_name = line_name_entry.get()
            new_line_metro = metro_name_entry.get()
            new_line_color = line_color_entry.get()
            try:
                new_line_year = int(line_year_entry.get())
            except:
                showerror('Ошибка записи', 'Год должен был целым числом')
                new_instance_window.destroy()
                return
            if new_line_metro not in list(database['metros']['Метрополитен']):
                showerror('Ошибка записи', 'Указанного метрополитена не существует')
                new_instance_window.destroy()
                return
            if new_line_code != selected[0] and new_line_code in set(database['lines']['Код линии']):
                showerror('Ошибка записи', 'Линия с таким кодом уже существует')
                new_instance_window.destroy()
                return
            if new_line_code != selected[0]: # Если хотим поменять код линии
                proceed = askyesno(
                    'Изменение кода',
                    'Вы собираетесь изменить ключевой атрибут. \
                    При изменении кода линии все станции, расположенные на ней, также поменяют код линии. \
                    Продолжить?'
                )
                if not proceed:
                    new_instance_window.destroy()
                    return
                else:
                    database['stations']['Код линии'].where(
                        cond = database['stations']['Код линии'] != selected[0],
                        other = new_line_code,
                        inplace = True
                    ) # Изменяем код линии в таблице о станциях
                    for i in table_1.get_children():
                        table_1.delete(i)
                    populate_table(0, table_1, database['stations'])
            database['lines'].loc[selected_index] = [
                new_line_code, new_line_name, new_line_metro, new_line_color,
                new_line_year
            ] # Собираем новую строку
            for i in table_2.get_children():
                table_2.delete(i)
            populate_table(1, table_2, database['lines'])
            columns = database['merged'].columns
            database['merged'] = database['stations'].merge(
                database['lines'], on = 'Код линии'
            ).merge(
                database['metros'], on = 'Метрополитен'
                ).drop(
                    ['Город', 'Год основания метро'], axis = 1
                )
            database['merged'].columns = columns
            for i in table_4.get_children():
                table_4.delete(i)
            populate_table(3, table_4, database['merged'])
            manage_open(None)
            new_instance_window.destroy()
        if current == 2: # Если редактируем запись о метро
            new_metro_name = metro_name_entry.get()
            new_metro_city = city_metro_entry.get()
            try:
                new_metro_year = int(metro_year_entry.get())
            except:
                showerror('Ошибка записи', 'Год должен был целым числом')
                new_instance_window.destroy()
                return
            if new_metro_name != selected[0] and new_metro_name in set(database['metros']['Метрополитен']):
                showerror('Ошибка записи', 'Метрополитен с таким именем уже существует')
                new_instance_window.destroy()
                return
            if new_metro_name != selected[0]:
                proceed = askyesno(
                    'Изменение названия',
                    'Вы собираетесь изменить ключевой атрибут. \
                    Имя метрополитена также поменяется в таблице "Линии". \
Продолжить?'
                )
                if not proceed:
                    new_instance_window.destroy()
                    return
            database['lines']['Метрополитен'].where(
                cond = database['lines']['Метрополитен'] != selected[0],
                other = new_metro_name, inplace = True
            ) # Изменяем имя метрополитена в таблице о линиях
            for i in table_2.get_children():
                table_2.delete(i)
            populate_table(2, table_2, database['lines'])
            database['metros'].loc[selected_index] = [
                new_metro_name, new_metro_city, new_metro_year
            ] # Собираем новую строчку из полей ввода
            for i in table_3.get_children():
                table_3.delete(i)
            populate_table(2, table_3, database['metros'])
            columns = database['merged'].columns
            database['merged'] = database['stations'].merge(
                database['lines'], on = 'Код линии'
            ).merge(
                database['metros'], on = 'Метрополитен'
                ).drop(
                    ['Город', 'Год основания метро'], axis = 1
                )
            database['merged'].columns = columns
            for i in table_4.get_children():
                table_4.delete(i)
            populate_table(3, table_4, database['merged'])
            manage_open(None)
            new_instance_window.destroy()
    """
    Окно редактирования записи. Аналогично окну добавления записи
    с небольшими изменениями, которые потребовали копировать код
    """
    new_instance_window = tk.Toplevel(app)
    new_instance_window.resizable(False, False)
    new_instance_window.geometry('300x400')
    new_instance_window.title('Редактировать запись')

    parameters = tk.LabelFrame(new_instance_window, text = 'Параметры')
    parameters.pack(anchor = 'nw', expand = True, fill = tk.BOTH)

    station_name = tk.Label(parameters, text = '{: ^25}'.format('Название станции'))
    station_name.grid(row = 0, column = 0)
    sn_entry = tk.Entry(parameters, width = 22)
    sn_entry.grid(row = 0, column = 1, pady = 5)

    line_code = tk.Label(parameters, text = '{: ^25}'.format('Код линии'))
    line_code.grid(row = 2, column = 0)
    lc_entry = tk.Entry(parameters, width = 22)
    lc_entry.grid(row = 2, column = 1, pady = 5)

    depth = tk.Label(parameters, text = '{: ^25}'.format('Глубина заложения'))
    depth.grid(row = 3, column = 0)
    depth_entry = tk.Entry(parameters, width = 22)
    depth_entry.grid(row = 3, column = 1, pady = 5)

    type = tk.Label(parameters, text = '{: ^25}'.format('Тип станции'))
    type.grid(row = 4, column = 0)
    type_entry = tk.Entry(parameters, width = 22)
    type_entry.grid(row = 4, column = 1, pady = 5)

    open_date = tk.Label(parameters, text = '{: ^25}'.format('Дата открытия станции'))
    open_date.grid(row = 5, column = 0)
    open_date_entry = tk.Entry(parameters, width = 22)
    open_date_entry.grid(row = 5, column = 1, pady = 5)

    line_name = tk.Label(parameters, text = '{: ^25}'.format('Название линии'))
    line_name.grid(row = 7, column = 0)
    line_name_entry = tk.Entry(parameters, width = 22)
    line_name_entry.grid(row = 7, column = 1, pady = 5)

    metro_name = tk.Label(parameters, text = '{: ^25}'.format('Метрополитен'))
    metro_name.grid(row = 8, column = 0)
    metro_name_entry = tk.Entry(parameters, width = 22)
    metro_name_entry.grid(row = 8, column = 1, pady = 5)

    line_color = tk.Label(parameters, text = '{: ^25}'.format('Цвет линии'))
    line_color.grid(row = 9, column = 0)
    line_color_entry = tk.Entry(parameters, width = 22)
    line_color_entry.grid(row = 9, column = 1, pady = 5)

    line_year = tk.Label(parameters, text = '{: ^25}'.format('Год открытия линии'))
    line_year.grid(row = 10, column = 0)
    line_year_entry = tk.Entry(parameters, width = 22)
    line_year_entry.grid(row = 10, column = 1, pady = 5)

    city_metro = tk.Label(parameters, text = '{: ^25}'.format('Город'))
    city_metro.grid(row = 11, column = 0)
    city_metro_entry = tk.Entry(parameters, width = 22)
    city_metro_entry.grid(row = 11, column = 1, pady = 5)

    metro_year = tk.Label(parameters, text = '{: ^25}'.format('Год основания метро'))
    metro_year.grid(row = 12, column = 0)
    metro_year_entry = tk.Entry(parameters, width = 22)
    metro_year_entry.grid(row = 12, column = 1, pady = 5)

    clear_entries = tk.Button(
        parameters, text = 'Очистить поля', width = 15,
        command = clear_all_entries
    )
    clear_entries.grid(row = 13, column = 0, pady = 15, padx = 10)

    create_instance = tk.Button(
        parameters, text = 'Редактировать', width = 15,
        command = lambda: alter_row(current, selected_index)
    )
    create_instance.grid(row = 13, column = 1, pady = 15, padx = 10)
    """
    Акнивность/неактивность элемента зависит от открытой таблицы
    """
    if current == 0:
        station_name["state"] = "normal"
        sn_entry["state"] = "normal"
        sn_entry.insert(tk.END, selected[0])
        line_code["state"] = "normal"
        lc_entry["state"] = "normal"
        lc_entry.insert(tk.END, selected[1])
        depth["state"] = "normal"
        depth_entry["state"] = "normal"
        depth_entry.insert(tk.END, selected[2])
        type["state"] = "normal"
        type_entry["state"] = "normal"
        type_entry.insert(tk.END, selected[3])
        open_date["state"] = "normal"
        open_date_entry["state"] = "normal"
        open_date_entry.insert(tk.END, selected[4])
        line_name["state"] = "disabled"
        line_name_entry["state"] = "disabled"
        metro_name["state"] = "disabled"
        metro_name_entry["state"] = "disabled"
        line_color["state"] = "disabled"
        line_color_entry["state"] = "disabled"
        line_year["state"] = "disabled"
        line_year_entry["state"] = "disabled"
        city_metro["state"] = "disabled"
        city_metro_entry["state"] = "disabled"
        metro_year["state"] = "disabled"
        metro_year_entry["state"] = "disabled"
    if current == 1:
        station_name["state"] = "disabled"
        sn_entry["state"] = "disabled"
        line_code["state"] = "normal"
        lc_entry["state"] = "normal"
        lc_entry.insert(tk.END, selected[0])
        depth["state"] = "disabled"
        depth_entry["state"] = "disabled"
        type["state"] = "disabled"
        type_entry["state"] = "disabled"
        open_date["state"] = "disabled"
        open_date_entry["state"] = "disabled"
        line_name["state"] = "normal"
        line_name_entry["state"] = "normal"
        line_name_entry.insert(tk.END, selected[1])
        metro_name["state"] = "normal"
        metro_name_entry["state"] = "normal"
        metro_name_entry.insert(tk.END, selected[2])
        line_color["state"] = "normal"
        line_color_entry["state"] = "normal"
        line_color_entry.insert(tk.END, selected[3])
        line_year["state"] = "normal"
        line_year_entry["state"] = "normal"
        line_year_entry.insert(tk.END, selected[4])
        city_metro["state"] = "disabled"
        city_metro_entry["state"] = "disabled"
        metro_year["state"] = "disabled"
        metro_year_entry["state"] = "disabled"
    if current == 2:
        station_name["state"] = "disabled"
        sn_entry["state"] = "disabled"
        line_code["state"] = "disabled"
        lc_entry["state"] = "disabled"
        depth["state"] = "disabled"
        depth_entry["state"] = "disabled"
        type["state"] = "disabled"
        type_entry["state"] = "disabled"
        open_date["state"] = "disabled"
        open_date_entry["state"] = "disabled"
        line_name["state"] = "disabled"
        line_name_entry["state"] = "disabled"
        metro_name["state"] = "normal"
        metro_name_entry["state"] = "normal"
        metro_name_entry.insert(tk.END, selected[0])
        line_color["state"] = "disabled"
        line_color_entry["state"] = "disabled"
        line_year["state"] = "disabled"
        line_year_entry["state"] = "disabled"
        city_metro["state"] = "normal"
        city_metro_entry["state"] = "normal"
        city_metro_entry.insert(tk.END, selected[1])
        metro_year["state"] = "normal"
        metro_year_entry["state"] = "normal"
        metro_year_entry.insert(tk.END, selected[2])

def buttons_state():
    """
    Управляет состоянием кнопок приложения
    Вызывается каждый раз при смене таблицы или выборе функции анализа
    в соответствующей таблице.
    Также вызывается при выборе строк основной таблицы.
    """
    current = tablayout.index((tablayout.select()))
    if current == 3:
        delete["state"] = "disabled"
        add["state"] = "disabled"
        edit["state"] = "disabled"
        if sel_4 == 0:
            export["state"] = "disabled"
        else:
            export["state"] = "active"
    elif (current == 0 and sel_1 == 0) or (current == 1 and sel_2 == 0) or \
    (current == 2 and sel_3 == 0):
        delete["state"] = "disabled"
        add["state"] = "active"
        edit["state"] = "disabled"
        export["state"] = "disabled"
    else:
        delete["state"] = "active"
        export["state"] = "active"
        if (current == 0 and sel_1 == 1) or (current == 1 and sel_2 == 1) or \
        (current == 2 and sel_3 == 1):
            edit["state"] = "active"
        else:
            edit["state"] = "disabled"
        add["state"] = "active"
    if (len(functions_treeview.selection()) != 1) or \
    (current == 0 and len_1 == 0) or \
    (current == 1 and len_2 == 0) or \
    (current == 2 and len_3 == 0) or \
    (current == 3 and len_4 == 0) : # Управление кнопкой "Показать"
        show["state"] = "disabled"
    else:
        show["state"] = "active"
    if (current == 0 and len_1 == 0) or \
    (current == 1 and len_2 == 0) or \
    (current == 2 and len_3 == 0) or \
    (current == 3 and len_4 == 0) : # Управление кнопками фильтрации
        configure_filter["state"] = "disabled"
    else:
        configure_filter["state"] = "active"

def manage_open(event):
    """
    Управляет текущей информацией о длине таблицы и количестве выбранных строк
    Также ссылается на buttons_state для управления состоянием кнопок
    Вызывается событием "Смена таблицы"
    """
    current = tablayout.index((tablayout.select()))
    global open_tab
    if current == 0:
        open_tab = "stations"
        number_of_rows.set("{: ^21}{}\n".format("Количество строк:", len_1))
        number_selected.set("{: ^20}{}\n".format("Выбранных строк:", sel_1))
    if current == 1:
        open_tab = "lines"
        number_of_rows.set("{: ^21}{}\n".format("Количество строк:", len_2))
        number_selected.set("{: ^20}{}\n".format("Выбранных строк:", sel_2))
    if current == 2:
        open_tab = "metros"
        number_of_rows.set("{: ^21}{}\n".format("Количество строк:", len_3))
        number_selected.set("{: ^20}{}\n".format("Выбранных строк:", sel_3))
    if current == 3:
        open_tab = "merged"
        number_of_rows.set("{: ^21}{}\n".format("Количество строк:", len_4))
        number_selected.set("{: ^20}{}\n".format("Выбранных строк:", sel_4))
    selection_message['state'] = 'normal'
    selection_message.delete('1.0', tk.END)
    selection_message.insert(tk.END, '\n\n')
    selection_message.insert(tk.END, number_of_rows.get())
    selection_message.insert(tk.END, '\n')
    selection_message.insert(tk.END, number_selected.get())
    selection_message['state'] = 'disabled'
    buttons_state()
    print_info()

def select(table, num):
    """
    Управляет счетчиком выбранных в данный момент строк таблицы
    Управляет отображением информации о количестве выбранных строк
    """
    selected = table.selection()
    sel = len(selected)
    number_selected.set("{: ^20}{}\n".format("Выбранных строк:", sel))
    selection_message['state'] = 'normal'
    selection_message.delete('1.0', tk.END)
    selection_message.insert(tk.END, '\n\n')
    selection_message.insert(tk.END, number_of_rows.get())
    selection_message.insert(tk.END, '\n')
    selection_message.insert(tk.END, number_selected.get())
    selection_message['state'] = 'disabled'
    if num == 0:
        global sel_1
        sel_1 = sel
    if num == 1:
        global sel_2
        sel_2 = sel
    if num == 2:
        global sel_3
        sel_3 = sel
    if num == 3:
        global sel_4
        sel_4 = sel
    buttons_state()

def sort(table, col):
    """
    Сортирует выбранный столбец текущей таблицы по возрастанию
    Вызывается кликом по заголовку столбца
    НЕ изменяет саму базу, а лишь управляет отображением таблицы в приложении
    """
    sorted = database[table].sort_values(by = col)
    global current_filtered
    current_filtered = sorted
    if table == 'stations':
        for i in table_1.get_children():
            table_1.delete(i)
        populate_table(0, table_1, sorted)
        global sel_1
        global len_1
        sel_1 = 0
        len_1 = len(sorted)
    if table == 'lines':
        for i in table_2.get_children():
            table_2.delete(i)
        populate_table(0, table_2, sorted)
        global sel_2
        global len_2
        sel_2 = 0
        len_2 = len(sorted)
    if table == 'metros':
        for i in table_3.get_children():
            table_3.delete(i)
        populate_table(0, table_3, sorted)
        global sel_3
        global len_3
        sel_3 = 0
        len_3 = len(sorted)
    if table == 'merged':
        for i in table_4.get_children():
            table_4.delete(i)
        populate_table(0, table_4, sorted)
        global sel_4
        global len_4
        sel_4 = 0
        len_4 = len(sorted)
    manage_open(None)

def unsort_f():
    current = tablayout.index((tablayout.select()))
    """
    Убирает сортировку - стирает отсортированное отображение
    и выводит датафрейм из базы
    """
    global current_filtered
    if current == 0:
        current_filtered = database['stations']
        for i in table_1.get_children():
            table_1.delete(i)
        populate_table(0, table_1, database['stations'])
        global sel_1
        global len_1
        sel_1 = 0
        len_1 = len(database['stations'])
    if current == 1:
        current_filtered = database['lines']
        for i in table_2.get_children():
            table_2.delete(i)
        populate_table(0, table_2, database['lines'])
        global sel_2
        global len_2
        sel_2 = 0
        len_2 = len(database['lines'])
    if current == 2:
        current_filtered = database['metros']
        for i in table_3.get_children():
            table_3.delete(i)
        populate_table(0, table_3, database['metros'])
        global sel_3
        global len_3
        sel_3 = 0
        len_3 = len(database['metros'])
    if current == 3:
        current_filtered = database['merged']
        for i in table_4.get_children():
            table_4.delete(i)
        populate_table(0, table_4, database['merged'])
        global sel_4
        global len_4
        sel_4 = 0
        len_4 = len(database['merged'])
    manage_open(None)
    for item in filters.get_children():
        filters.delete(item)
    filters.insert('', 0, values = ['Год основания метро', ''])
    filters.insert('', 0, values = ['Город', ''])
    filters.insert('', 0, values = ['Метрополитен', ''])
    filters.insert('', 0, values = ['Цвет', ''])
    filters.insert('', 0, values = ['Название линии', ''])
    filters.insert('', 0, values = ['Год открытия линии', ''])
    filters.insert('', 0, values = ['Дата открытия станции', ''])
    filters.insert('', 0, values = ['Тип станции', ''])
    filters.insert('', 0, values = ['Глубина заложения', ''])
    filters.insert('', 0, values = ['Код линии', ''])
    filters.insert('', 0, values = ['Название станции', ''])

# Меню
"""
Функции, вызываемые кнопками выпадающего меню
"""
def open_file():
    """
    Позволяет открыть базу из файла.
    Загружает базу в глобальную переменную-словарь database
    """
    try:
        global current_file_name
        global database
        file = askopenfilename()
        current_file_name = file
        f = open(file, 'rb')
        base = pickle.load(f)
        f.close()
        for i in table_1.get_children():
            table_1.delete(i)
        for i in table_2.get_children():
            table_2.delete(i)
        for i in table_3.get_children():
            table_3.delete(i)
        for i in table_4.get_children():
            table_4.delete(i)
        populate_table(0, table_1, base[0])
        database['stations'] = base[0]
        populate_table(1, table_2, base[1])
        database['lines'] = base[1]
        populate_table(2, table_3, base[2])
        database['metros'] = base[2]
        populate_table(3, table_4, base[3])
        database['merged'] = base[3]
        database['merged'].columns = (
            'Название станции', 'Код линии', 'Глубина заложения', 'Тип',
            'Дата открытия станции',
            'Название линии', 'Метрополитен', 'Цвет линии',
            'Год открытия линии'
        )
        manage_open(None)
        unsort_f()
    except:
        showerror('Ошибка загрузки', 'Что-то пошло не так')

def new_base():
    """
    Инициализирует пустую базу
    """
    global len_1
    global len_2
    global len_3
    global len_4
    global database
    database = { # Пустая база - пустые датафреймы со столбцами
        'stations': pd.DataFrame(data = [], columns = [
                'Название', 'Код линии', 'Глубина заложения', 'Тип', 'Дата открытия'
            ]),
        'lines': pd.DataFrame(data = [], columns = [
                'Код линии', 'Название', 'Метрополитен', 'Цвет', 'Год открытия линии'
            ]),
        'metros': pd.DataFrame(data = [], columns = [
                'Метрополитен', 'Город', 'Год основания метро'
            ]),
        'merged': pd.DataFrame(data = [], columns = [
                'Название станции', 'Код линии', 'Глубина заложения', 'Тип',
                'Дата открытия станции',
                'Название линии', 'Метрополитен', 'Цвет линии', 'Год открытия линии'
            ])
    }
    global current_file_name
    current_file_name = None
    print('New Base')
    for i in table_1.get_children():
        table_1.delete(i)
    for i in table_2.get_children():
        table_2.delete(i)
    for i in table_3.get_children():
        table_3.delete(i)
    for i in table_4.get_children():
        table_4.delete(i)
    len_1, len_2, len_3, len_4 = 0, 0, 0, 0
    manage_open(None)

def base_from_csv():
    """
    Позволяет создать pickle базу
    из трех csv таблиц
    """
    def select_file(number):
        """
        Отвечает за выбор файла с csv таблицей
        """
        file = askopenfilename()
        try:
            df = pd.read_csv(file, encoding = 'utf-16')
            if number == 0:
                text_one.delete(0, tk.END)
                text_one.insert(tk.END, file)
            if number == 1:
                text_two.delete(0, tk.END)
                text_two.insert(tk.END, file)
            if number == 2:
                text_three.delete(0, tk.END)
                text_three.insert(tk.END, file)
            if number == 3:
                text_four.delete(0, tk.END)
                text_four.insert(tk.END, file)
        except:
            showerror('Ошибка загрузки', 'Что-то пошло не так')

    def create_database():
        """
        Создает базу из трех таблиц и сохраняет
        в формате pickle
        """
        try:
            df1 = pd.read_csv(text_one.get(), encoding = 'utf-16')
            df1['Дата открытия'] = pd.to_datetime(df1['Дата открытия'], format = '%d-%m-%Y').dt.date
            df2 = pd.read_csv(text_two.get(), encoding = 'utf-16')
            df3 = pd.read_csv(text_three.get(), encoding = 'utf-16')
            merged = df1.merge(
                df2, on = 'Код линии'
            ).merge(
                df3, on = 'Метрополитен'
                ).drop(
                    ['Город', 'Год основания метро'], axis = 1
                    )
            data_base = [df1, df2, df3, merged]
            f = open(text_name.get(), 'wb')
            pickle.dump(data_base, f)
            f.close()
            bft.destroy()
        except:
            showerror('Ошибка', 'Что-то пошло не так')

    """
    Окно выбора трех csv файлов
    """
    bft = tk.Toplevel(app)
    bft.resizable(False, False)
    bft.geometry('250x250')
    bft.title('База из CSV')

    one = tk.LabelFrame(bft, text = 'Таблица "Станции"')
    one.pack(anchor = 'nw', expand = True, fill = tk.BOTH)
    text_one = tk.Entry(one)
    text_one.place(anchor = 'c', relx = 0.65, rely = 0.5)
    text_one.insert(tk.END, 'Файл не выбран')
    select_one = tk.Button(one, text = 'Выбрать', command = lambda: select_file(0))
    select_one.place(anchor = 'c', relx = 0.25, rely = 0.5)

    two = tk.LabelFrame(bft, text = 'Таблица "Линии"')
    two.pack(anchor = 'nw', expand = True, fill = tk.BOTH)
    text_two = tk.Entry(two)
    text_two.place(anchor = 'c', relx = 0.65, rely = 0.5)
    text_two.insert(tk.END, 'Файл не выбран')
    select_two = tk.Button(two, text = 'Выбрать', command = lambda: select_file(1))
    select_two.place(anchor = 'c', relx = 0.25, rely = 0.5)

    three = tk.LabelFrame(bft, text = 'Таблица "Метрополитены"')
    three.pack(anchor = 'nw', expand = True, fill = tk.BOTH)
    text_three = tk.Entry(three)
    text_three.place(anchor = 'c', relx = 0.65, rely = 0.5)
    text_three.insert(tk.END, 'Файл не выбран')
    select_three = tk.Button(three, text = 'Выбрать', command = lambda: select_file(2))
    select_three.place(anchor = 'c', relx = 0.25, rely = 0.5)

    ready = tk.LabelFrame(bft, text = 'Имя новой базы')
    ready.pack(anchor = 'nw', expand = True, fill = tk.BOTH)
    create = tk.Button(ready, text = 'Создать', command = create_database)
    create.place(anchor = 'c', relx = 0.25, rely = 0.5)
    text_name = tk.Entry(ready)
    text_name.place(anchor = 'c', relx = 0.65, rely = 0.5)
    text_name.insert(tk.END, 'Введите имя базы')

def save():
    """
    Сохраняет текущую базу
    """
    print("Save")
    proceed = askyesno(
        'Сохранить', 'Вы действительно хотите перезаписать базу?'
    )
    if not proceed:
        return
    try:
        file_name = current_file_name
        f = open(file_name, 'wb')
        base = []
        base.append(database['stations'])
        base.append(database['lines'])
        base.append(database['metros'])
        base.append(database['merged'])
        pickle.dump(base, f)
        f.close()
    except:
        save_as()

def save_as():
    """
    Функция "Сохранить как""
    """
    file_name = asksaveasfilename(
        filetypes = (('All Files', "*.*"), ),
    )
    f = open(file_name, 'wb')
    base = []
    base.append(database['stations'])
    base.append(database['lines'])
    base.append(database['metros'])
    base.append(database['merged'])
    pickle.dump(base, f)
    f.close
"""
Код выпадающего меню
"""
menu_frame = tk.Frame(app, relief = tk.RAISED, borderwidth = 1)
menu_frame.pack(side = tk.TOP, anchor = 'n', fill = tk.X)
menu_button = tk.Menubutton(menu_frame, text = 'Управление базой')
menu_button.pack(side = tk.LEFT)
menu_button.menu = tk.Menu(menu_button, tearoff = False)
menu_button.menu.add_command(label = 'Новая база', command = new_base)
menu_button.menu.add_command(label = 'Открыть базу', command = open_file)
menu_button.menu.add_command(label = 'Новая база из CSV', command = base_from_csv)
menu_button.menu.add("separator")
menu_button.menu.add_command(label = 'Сохранить', command = save)
menu_button.menu.add_command(label = 'Сохранить как', command = save_as)
menu_button['menu'] = menu_button.menu

# Вкладки для таблицы

def populate_table(number, tree, data_frame):
    """
    Функция, заполняющая указанную таблицу Treeview
    датафреймом из указанной переменной
    """
    if number == 0:
        global len_1
        len_1 = len(data_frame)
    if number == 1:
        global len_2
        len_2 = len(data_frame)
    if number == 2:
        global len_3
        len_3 = len(data_frame)
    if number == 3:
        global len_4
        len_4 = len(data_frame)
    for i in range(len(data_frame)):
        vals = data_frame.iloc[i, :].tolist()
        tree.insert('', i, values = vals, tags = ('even',) if i % 2 == 0 else ('odd',))


table_frame = tk.Frame(app, borderwidth = 1, relief = tk.RAISED)
table_frame.pack(anchor = 'n', fill = tk.X, side = tk.TOP)
# Окно-тетрадка, переключатель таблиц
tablayout = ttk.Notebook(table_frame, height = 350)

# Таблица 1
tab1 = tk.Label(tablayout)
tab1.pack(anchor = 'n', side = tk.TOP)
tablayout.add(tab1, text = '{: ^25}'.format('Станции'))
tablayout.bind('<<NotebookTabChanged>>', manage_open) # Событие смены вкладки

table_1 = ttk.Treeview(tab1)
table_1.bind("<<TreeviewSelect>>", lambda x: select(table_1, 0)) # Событие выбора строки
table_1.tag_configure('even', background = 'blue')
table_1.tag_configure('odd', background = 'red')
table_1['columns'] = (
    'Название', 'Код линии', 'Глубина заложения', 'Тип', 'Дата открытия',
)
table_1.heading("#0", text = "", anchor = 'w')
table_1.column("#0", anchor = 'e', width = 2, stretch = tk.NO)

table_1.heading(
    'Название', text = '{: <25}'.format('Название'), anchor = 'w',
    command = lambda: sort('stations', 'Название')
)
table_1.heading(
    'Код линии', text = '{: <25}'.format('Код линии'), anchor = 'w',
    command = lambda: sort('stations', 'Код линии')
)
table_1.heading(
    'Глубина заложения', text = '{: <25}'.format('Глубина заложения'),
    anchor = 'w', command = lambda: sort('stations', 'Глубина заложения')
)
table_1.heading(
    'Тип', text = '{: <25}'.format('Тип'), anchor = 'w',
    command = lambda: sort('stations', 'Тип')
)
table_1.heading(
    'Дата открытия', text = '{: <25}'.format('Дата открытия'),
    anchor = 'w', command = lambda: sort('stations', 'Дата открытия')
)
for col in table_1['columns']:
    table_1.column(col, anchor = 'w', width = 150, stretch = tk.NO)
table_1.pack(anchor = 'n', expand = True, fill = tk.BOTH)
# Скроллбары таблицы
vertical_scrollbar_1 = tk.Scrollbar(table_1, command = table_1.yview)
vertical_scrollbar_1.pack(side = tk.RIGHT, fill = tk.Y)
horizontal_scrollbar_1 = tk.Scrollbar(table_1, command = table_1.xview, orient = 'horizontal')
horizontal_scrollbar_1.pack(side = tk.BOTTOM, fill = tk.X)
table_1.configure(yscroll = vertical_scrollbar_1.set, xscroll = horizontal_scrollbar_1.set)

# Таблица 2
tab2 = tk.Label(tablayout)
tab2.pack(anchor = 'n', side = tk.LEFT, fill = tk.X)

table_2 = ttk.Treeview(tab2)
table_2.bind("<<TreeviewSelect>>", lambda x: select(table_2, 1)) # Событие выбора строки
table_2['columns'] = (
    'Код линии', 'Название', 'Метрополитен', 'Цвет', 'Год открытия линии',
)
table_2.heading("#0", text = "", anchor = 'w')
table_2.column("#0", anchor = 'w', width = 2, stretch = tk.NO)
table_2.heading(
    'Код линии', text = 'Код линии', anchor = 'w',
    command = lambda: sort('lines', 'Код линии')
)
table_2.heading(
    'Название', text = 'Название', anchor = 'w',
    command = lambda: sort('lines', 'Название')
)
table_2.heading(
    'Метрополитен', text = 'Метрополитен', anchor = 'w',
    command = lambda: sort('lines', 'Метрополитен')
)
table_2.heading(
    'Цвет', text = 'Цвет', anchor = 'w',
    command = lambda: sort('lines', 'Цвет')
)
table_2.heading(
    'Год открытия линии', text = 'Год открытия линии', anchor = 'w',
    command = lambda: sort('lines', 'Год открытия линии')
)
for column in table_2['columns']:
    table_2.column(column, anchor = 'w', width = 170, stretch = tk.NO)
table_2.pack(anchor = 'n', expand = True, fill = tk.BOTH)
tablayout.add(tab2, text = '{: ^25}'.format('Линии'))
# Скроллбары таблицы
vertical_scrollbar_2 = tk.Scrollbar(table_2, command = table_2.yview)
vertical_scrollbar_2.pack(side = tk.RIGHT, fill = tk.Y)
horizontal_scrollbar_2 = tk.Scrollbar(table_2, command = table_2.xview, orient = 'horizontal')
horizontal_scrollbar_2.pack(side = tk.BOTTOM, fill = tk.X)
table_2.configure(yscroll = vertical_scrollbar_2.set, xscroll = horizontal_scrollbar_2.set)

# Таблица 3
tab3 = tk.Label(tablayout)
tab3.pack(anchor = 'n', side = tk.LEFT, fill = tk.X)

table_3 = ttk.Treeview(tab3)
table_3.bind("<<TreeviewSelect>>", lambda x: select(table_3, 2)) # Событие выбора строк
table_3['columns'] = (
    'Метрополитен', 'Город', 'Год основания метро',
)
table_3.heading("#0", text = "", anchor = 'w')
table_3.column("#0", anchor = 'w', width = 2, stretch = tk.NO)
table_3.heading(
    'Метрополитен', text = 'Метрополитен', anchor = 'w',
    command = lambda: sort('metros', 'Метрополитен')
)
table_3.heading(
    'Город', text = 'Город', anchor = 'w',
    command = lambda: sort('metros', 'Город')
)
table_3.heading(
    'Год основания метро', text = 'Год основания метро', anchor = 'w',
    command = lambda: sort('metros', 'Год основания метро')
)
for column in table_3['columns']:
    table_3.column(column, anchor = 'w', width = 170, stretch = tk.NO)
table_3.pack(anchor = 'n', expand = True, fill = tk.BOTH)

tablayout.add(tab3, text = '{: ^25}'.format('Метрополитены'))
# Скроллбары таблицы
vertical_scrollbar_3 = tk.Scrollbar(table_3, command = table_3.yview)
vertical_scrollbar_3.pack(side = tk.RIGHT, fill = tk.Y)
horizontal_scrollbar_3 = tk.Scrollbar(table_3, command = table_3.xview, orient = 'horizontal')
horizontal_scrollbar_3.pack(side = tk.BOTTOM, fill = tk.X)
table_3.configure(yscroll = vertical_scrollbar_3.set, xscroll = horizontal_scrollbar_3.set)

# Таблица 4
tab4 = tk.Label(tablayout)
tab4.pack(anchor = 'n', side = tk.LEFT, fill = tk.X)

table_4 = ttk.Treeview(tab4)
table_4.bind("<<TreeviewSelect>>", lambda x: select(table_4, 3)) # Событие выбора строк
table_4['columns'] = (
    'Название станции', 'Код линии', 'Глубина заложения', 'Тип', 'Дата открытия станции',
    'Название линии', 'Метрополитен', 'Цвет линии', 'Год открытия линии',
)
table_4.heading("#0", text = "", anchor = 'w')
table_4.column("#0", anchor = 'w', width = 2, stretch = tk.NO)
table_4.heading(
    'Название станции', text = 'Название станции',
    anchor = 'w', command = lambda: sort('merged', 'Название станции')
)
table_4.heading(
    'Код линии', text = 'Код линии',
    anchor = 'w', command = lambda: sort('merged', 'Код линии')
)
table_4.heading(
    'Глубина заложения', text = 'Глубина заложения',
    anchor = 'w', command = lambda: sort('merged', 'Глубина заложения')
)
table_4.heading(
    'Тип', text = 'Тип',
    anchor = 'w', command = lambda: sort('merged', 'Тип')
)
table_4.heading(
    'Дата открытия станции', text = 'Дата открытия станции',
    anchor = 'w', command = lambda: sort('merged', 'Дата открытия станции')
)
table_4.heading(
    'Название линии', text = 'Название линии',
    anchor = 'w', command = lambda: sort('merged', 'Название линии')
)
table_4.heading(
    'Метрополитен', text = 'Метрополитен',
    anchor = 'w', command = lambda: sort('merged', 'Метрополитен')
)
table_4.heading(
    'Цвет линии', text = 'Цвет линии',
    anchor = 'w', command = lambda: sort('merged', 'Цвет линии')
)
table_4.heading(
    'Год открытия линии', text = 'Год открытия линии',
    anchor = 'w', command = lambda: sort('merged', 'Год открытия линии')
)
for column in table_4['columns']:
    table_4.column(column, anchor = 'w', width = 170, stretch = tk.NO)
table_4.pack(anchor = 'n', expand = True, fill = tk.BOTH)

tablayout.add(tab4, text = '{: ^25}'.format('Соединение таблиц'))
# Скроллбары таблицы
vertical_scrollbar_4 = tk.Scrollbar(table_4, command = table_4.yview)
vertical_scrollbar_4.pack(side = tk.RIGHT, fill = tk.Y)
horizontal_scrollbar_4 = tk.Scrollbar(table_4, command = table_4.xview, orient = 'horizontal')
horizontal_scrollbar_4.pack(side = tk.BOTTOM, fill = tk.X)
table_4.configure(yscroll = vertical_scrollbar_4.set, xscroll = horizontal_scrollbar_4.set)

tablayout.pack(side = tk.LEFT, anchor = 's', fill = tk.BOTH, expand = True)

# Управление таблицей

"""
Код четырех блоков, расположенных под таблицей
Блоки включают в себя управление записями в базе,
средства анализа (графики и сводные таблицы),
а также фильтр отображаемых таблиц
Также есть блок с информацией о таблице
"""
"""
Блок управления записями в базе
"""
control_frame = tk.LabelFrame(
    app, text = 'Управление таблицей',
)
selection_message = tk.Text(control_frame, height = 7, width = 28)
selection_message.configure(font = 'calibri 11')
selection_message.insert(tk.END, '\n\n')
selection_message.insert(tk.END, number_of_rows.get())
selection_message.insert(tk.END, '\n')
selection_message.insert(tk.END, number_selected.get())
selection_message.grid(columnspan = 3, pady = 7)
add = tk.Button(
    control_frame, text = 'Добавить', width = 11,
    command = lambda: add_instance(tablayout.index((tablayout.select())))
) # Добавить запись
delete = tk.Button(
    control_frame, text = 'Удалить', width = 11,
    command = lambda: delete_selected(open_tab)
) # Удалить запись
edit = tk.Button(
    control_frame, text = 'Править', width = 11,
    command = lambda: edit_instance(tablayout.index((tablayout.select())))
) # Редактировать запись
export = tk.Button(
    control_frame, text = 'Экспорт', width = 11,
    command = lambda: export_selected(tablayout.index((tablayout.select())))
) # Экспортировать запись
add.grid(column = 0, row = 4, rowspan = 2, pady = 6)
delete.grid(column = 0, row = 7, rowspan = 2, pady = 6)
edit.grid(column = 2, row = 4, rowspan = 2, pady = 6)
export.grid(column = 2, row = 7, rowspan = 2, pady = 6)

control_frame.pack(side = tk.LEFT, fill = tk.BOTH)
"""
Блок просмотра информации о таблице
"""
info = tk.LabelFrame(app, text = 'Информация о таблице')
info_message = tk.Text(info, height = 8, width = 33)
info_message.configure(font = 'calibri 10')
info_message.grid(columnspan = 3, pady = 7)
unsort = tk.Button(
    info, text = 'Сбросить сортировку',
    width = 25,
    command = unsort_f
)
some_button = tk.Button(
    info, text = 'Экспорт отображения',
    width = 25,
    command = lambda: export_representation(current_filtered)
)
unsort.grid(row = 4, padx = 15, pady = 11)
some_button.grid(row = 5, padx = 15, pady = 2)
info.pack(side = tk.LEFT, fill = tk.BOTH)
"""
Блок инструментов анализа данных
"""
analysis_frame = tk.LabelFrame(
    app, text = 'Инструменты анализа',
)
# Таблица выбора инструмента анализа
functions_treeview = ttk.Treeview(analysis_frame)
functions_treeview.bind("<<TreeviewSelect>>", lambda x: buttons_state()) # Событие выбора строки
functions_treeview['columns'] = ('Функция')
functions_treeview.heading("#0", text = "", anchor = 'w')
functions_treeview.column("#0", anchor = 'w', width = 2, stretch = tk.NO)
functions_treeview.heading('Функция', text = 'Функция', anchor = 'w')
functions_treeview.insert('', 0, values = ['Сводная таблица'])
functions_treeview.insert('', 0, values = ['График плотности распределения'])
functions_treeview.insert('', 0, values = ['Столбчатая диаграмма'])
functions_treeview.insert('', 0, values = ['Гистограмма'])
functions_treeview.insert('', 0, values = ['Диаграмма размаха'])
# Скроллбар таблицы
vertical_scrollbar_functions = tk.Scrollbar(
    functions_treeview, command = functions_treeview.yview
)
vertical_scrollbar_functions.pack(side = tk.RIGHT, fill = tk.Y)
functions_treeview.pack(expand = True, fill = tk.BOTH, padx = 2, pady = 8)

show = tk.Button(
    analysis_frame, text = 'Показать', width = 42,
    command = lambda: analysis_function(tablayout.index((tablayout.select())))
) # Кнопка, подтверждающая выбор инструмента анализа и открывающая окно выбора столбцов
show.pack(side = tk.BOTTOM, pady = 20, padx = 10, fill = tk.X)

analysis_frame.pack(side = tk.LEFT, fill = tk.Y)

"""
Блок фильтрации отображаемой таблицы.
Фильтрация также влияет на графики и прочие инструменты анализа.
"""
filter_frame = tk.LabelFrame(
    app, text = 'Фильтры',
)
filters = ttk.Treeview(filter_frame, height = 4)
filters['columns'] = ('Столбец', 'Фильтр')
filters.heading("#0", text = "", anchor = 'w')
filters.column("#0", anchor = 'w', width = 2, stretch = tk.NO)
filters.heading('Столбец', text = 'Столбец', anchor = 'w')
filters.heading('Фильтр', text = 'Фильтр', anchor = 'w')

filters.insert('', 0, values = ['Год основания метро', ''])
filters.insert('', 0, values = ['Город', ''])
filters.insert('', 0, values = ['Метрополитен', ''])
filters.insert('', 0, values = ['Цвет', ''])
filters.insert('', 0, values = ['Название линии', ''])
filters.insert('', 0, values = ['Год открытия линии', ''])
filters.insert('', 0, values = ['Дата открытия станции', ''])
filters.insert('', 0, values = ['Тип станции', ''])
filters.insert('', 0, values = ['Глубина заложения', ''])
filters.insert('', 0, values = ['Код линии', ''])
filters.insert('', 0, values = ['Название станции', ''])

# Скроллбар таблицы
vertical_scrollbar_filters = tk.Scrollbar(
    filters, command = filters.yview
)
vertical_scrollbar_filters.pack(side = tk.RIGHT, fill = tk.Y)
filters.pack(expand = True, fill = tk.BOTH, padx = 5, pady = 7)
configure_filter = tk.Button(
    filter_frame, text = 'Настроить фильтры', width = 30,
    command = lambda: configure_f(tablayout.index((tablayout.select())))
)
configure_filter.pack(side = tk.BOTTOM, padx = 15, pady = 21, fill = tk.X)
filter_frame.pack(side = tk.LEFT, expand = True, fill = tk.BOTH)


app.mainloop()
