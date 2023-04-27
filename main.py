import xlrd
from tkinter import *
from tkinter import ttk
import os


class Glavn:
    _tt = None
    _selected_template = None
    _foo = None
    _params_visibility = None
    _comboboxes = []
    _selected_params_static = None



    def __init__(self):




        def refresh_options():
            self._comboboxes.clear()
            for rx1 in range(self._foo.nrows):
                if rx1 > 0:
                    self._comboboxes.append(ttk.Combobox(values=['Да', 'Нет']))
                    self._comboboxes[rx1 - 1].grid(row=rx1 - 1, column=1)
                    self._comboboxes[rx1 - 1].set(self._selected_params_static[rx1 - 1])
                    self._comboboxes[rx1 - 1].bind("<<ComboboxSelected>>", param_selected)
                    label1 = ttk.Label(text=self._foo.cell_value(rowx=rx1, colx=0))
                    label1.grid(row=rx1 - 1, column=0)



        def open_file_click(event):
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
            for col_index in range(self._foo.ncols):
                template_params = []
                for row_index in range(self._foo.nrows):
                    if col_index > 0:
                        template_params.append('Да') if self._foo.cell_value(rowx=row_index, colx=col_index) != '' else template_params.append('Нет')
                if len(template_params) > 0:
                    template_params.pop(0)
                if template_params == selected_params:
                    status_label = ttk.Label(text="Обнаружен подходящий блок")
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
                    lista = root.grid_slaves()
                    for ls in lista:
                        ls.destroy()
                    refresh_options()





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


