import subprocess

import xlrd
from tkinter import *
from tkinter import ttk
import os
import sys

class Glavn:
    _tt = None
    _selected_template = None
    _foo = None
    _params_visibility = None
    _comboboxes = []
    _selected_params_static = None
    _selected_params_static_prev = []

    def __init__(self):

        def compatible_option(option_index, selected_options, selected_options_prev, db):
            for i in range(len(selected_options_prev)): #Последний измененный пункт не должен становиться неактивным
                if (selected_options_prev[i] != selected_options[i]) and (i == option_index):
                    self._selected_params_static_prev = selected_options
                    return TRUE

            y_count_in_so = 0
            for option in selected_options:
                if option == 'Да':
                    y_count_in_so += 1

            if y_count_in_so == 0:
                return TRUE

            for col_index in range(db.ncols):
                if col_index > 0:
                    y_counter = 0
                    for so_index in range(len(selected_options)):
                        so = selected_options[so_index]
                        cv = db.cell_value(rowx=so_index+1, colx=col_index)
                        if so == 'Да' and cv != '':
                            y_counter += 1
                    if y_counter == y_count_in_so:
                        cv = db.cell_value(rowx=option_index + 1, colx=col_index)
                        if cv != '':
                            return TRUE

        def refresh_options():

            lista = root.grid_slaves()
            for ls in lista:
                ls.destroy()
            self._comboboxes.clear()
            for rx1 in range(self._foo.nrows):
                if rx1 > 0:
                    self._comboboxes.append(ttk.Combobox(values=['Да', 'Нет']))
                    self._comboboxes[rx1 - 1].grid(row=rx1 - 1, column=1)
                    self._comboboxes[rx1 - 1].set(self._selected_params_static[rx1 - 1])
                    self._comboboxes[rx1 - 1].bind("<<ComboboxSelected>>", param_selected)
                    label1 = ttk.Label(text=self._foo.cell_value(rowx=rx1, colx=0))
                    label1.grid(row=rx1 - 1, column=0)
                    if not compatible_option(rx1 - 1, self._selected_params_static, self._selected_params_static_prev,
                                         self._foo):
                        self._comboboxes[rx1 - 1].state(['disabled'])

        def open_file_click(event):
            file_path = os.path.join(self._tt, str(int(self._selected_template)))

            print(sys.platform)
            if os.path.exists(file_path + '.dwg'):
                if sys.platform == 'linux':
                    subprocess.call(["xdg-open", file_path + '.dwg'])
                else:
                    os.startfile(file_path + '.dwg')

            if os.path.exists(file_path + '.txt'):
                if sys.platform == 'linux':
                    subprocess.call(["xdg-open", file_path + '.txt'])
                else:
                    os.startfile(file_path + '.txt')



        # ШАГ 1. Выбор типа схемы
        def type_selected(event):
            selected_type = type_names.index(typ_select_combobox.get())
            book2 = xlrd.open_workbook("DB{0}.xls".format(str(selected_type + 2)))
            self._foo = book2.sheet_by_index(0)
            self._tt = str(selected_type + 2)
            for rx1 in range(self._foo.nrows):
                if rx1 > 0:
                    self._comboboxes.append(ttk.Combobox(values=['Да', 'Нет']))
                    self._comboboxes[rx1 - 1].set('Нет')
                    self._comboboxes[rx1 - 1].grid(row=rx1 - 1, column=1)
                    self._comboboxes[rx1 - 1].bind("<<ComboboxSelected>>", param_selected)
                    label1 = ttk.Label(text=self._foo.cell_value(rowx=rx1, colx=0))
                    label1.grid(row=rx1 - 1, column=0)

        # ШАГ 2. Проверка параметров

        def param_selected(event):

            selected_params = []
            [selected_params.append(combobox.get()) for combobox in self._comboboxes]
            self._selected_params_static = selected_params
            if self._selected_params_static_prev == []:
                for s in range(len(selected_params)):
                    self._selected_params_static_prev.append('Нет')
            refresh_options()
            for col_index in range(self._foo.ncols):
                template_params = []
                for row_index in range(self._foo.nrows):
                    if col_index > 0:
                        template_params.append('Да') if self._foo.cell_value(rowx=row_index, colx=col_index) != '' else template_params.append('Нет')
                if len(template_params) > 0:
                    template_params.pop(0)
                if template_params == selected_params:
                    status_label = ttk.Label(text="Обнаружен подходящий шаблон")
                    status_label.grid(row=len(selected_params) + 1, column=0)
                    btn = ttk.Button(text="Открыть")
                    btn.bind("<ButtonPress-1>", open_file_click)
                    btn.grid(row=len(selected_params) + 1, column=1)
                    self._selected_template = self._foo.cell_value(rowx=0, colx=col_index)
                    break
                else:
                    status_label = ttk.Label(text="Подходящий шаблон не найден")
                    status_label.grid(row=len(selected_params) + 1, column=0)
                    btn = ttk.Button(text="Открыть")
                    btn.grid(row=len(selected_params) + 1, column=1)
                    btn.state(['disabled'])







        root = Tk()
        root.title("RAY")

        w = root.winfo_screenwidth()
        h = root.winfo_screenheight()
        w = w // 2  # середина экрана
        h = h // 2
        w = w - 250  # смещение от середины
        h = h - 250
        root.geometry('500x500+{}+{}'.format(w, h))

        type_names = []


        book = xlrd.open_workbook("DB1.xls")
        sh = book.sheet_by_index(0)
        [type_names.append(sh.cell_value(rowx=rx, colx=0)) for rx in range(sh.nrows)]
        label = ttk.Label(text="Тип схемы:")
        label.grid(row=0, column=0)
        type_names.pop(0)
        typ_select_combobox = ttk.Combobox(values=type_names)
        typ_select_combobox.grid(row=0, column=1)
        typ_select_combobox.bind("<<ComboboxSelected>>", type_selected)

        root.mainloop()





a = Glavn()


