import xlrd
from tkinter import *
from tkinter import ttk
import os


class Glavn:

    _tt = None
    _selected_template = None
    _foo = None
    def __init__(self):

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
        comboboxes = []
        selected_params = []

        book = xlrd.open_workbook("DB1.xls")
        sh = book.sheet_by_index(0)
        [type_names.append(sh.cell_value(rowx=rx, colx=0)) for rx in range(sh.nrows)]

        def single_click(event):
            file_path = self._tt + "\\" + str(int(self._selected_template))
            if os.path.exists(file_path + '.dwg'):
                os.startfile(file_path + '.dwg', 'open')
            if os.path.exists(file_path + '.txt'):
                os.startfile(file_path + '.txt', 'open')

        # ШАГ 1. Выбор типа схемы
        def type_selected(event):
            selected_type = type_names.index(typ_select_combobox.get())
            book2 = xlrd.open_workbook("DB{0}.xls".format(str(selected_type + 2)))
            self._foo = book2.sheet_by_index(0)
            self._tt = str(selected_type + 2)
            for rx1 in range(self._foo.nrows):
                if rx1 > 0:
                    comboboxes.append(ttk.Combobox(values=['Да', 'Нет']))
                    comboboxes[rx1 - 1].grid(row=rx1 - 1, column=1)
                    comboboxes[rx1 - 1].bind("<<ComboboxSelected>>", param_selected)
                    label1 = ttk.Label(text=self._foo.cell_value(rowx=rx1, colx=0))
                    label1.grid(row=rx1 - 1, column=0)

        label = ttk.Label(text="Тип схемы:")
        label.grid(row=0, column=0)
        type_names.pop(0)
        typ_select_combobox = ttk.Combobox(values=type_names)
        typ_select_combobox.grid(row=0, column=1)
        typ_select_combobox.bind("<<ComboboxSelected>>", type_selected)

        # ШАГ 2. Проверка параметров
        def param_selected(event, ):
            selected_params.clear()
            match_found = False
            [selected_params.append(combobox.get()) for combobox in comboboxes]
            if selected_params.count('') == 0:
                status_label = ttk.Label(text="Все поля заполнены")
                status_label.grid(row=len(selected_params) + 1, column=0)
                for col_index in range(self._foo.ncols):
                    if col_index > 0 and not match_found:
                        for row_index in range(self._foo.nrows):
                            if row_index > 0:
                                if selected_params[row_index-1]:
                                    cell = self._foo.cell_value(rowx=row_index, colx=col_index)
                                    param = selected_params[row_index-1]
                                    if (param == 'Да' and cell != '') or (param == 'Нет' and cell == ''):
                                        if row_index == len(selected_params):
                                            status_label = ttk.Label(text="Обнаружен подходящий блок")
                                            status_label.grid(row=len(selected_params) + 1, column=0)
                                            btn = ttk.Button(text="Открыть")
                                            btn.bind("<ButtonPress-1>", single_click)
                                            btn.grid(row=len(selected_params) + 1, column=1)
                                            self._selected_template = self._foo.cell_value(rowx=0, colx=col_index)
                                            match_found = True
                                    else:
                                        break
            else:
                status_label = ttk.Label(text="Не все поля заполнены")
                status_label.grid(row=len(selected_params) + 1, column=0)

            if not match_found:
                status_label = ttk.Label(text="Подходящий шаблон не найден")
                status_label.grid(row=len(selected_params) + 1, column=0)
                btn = ttk.Button(text="Открыть")
                btn.grid(row=len(selected_params) + 1, column=1)
                btn.state(['disabled'])

        root.mainloop()


a = Glavn()
