import xlrd
from tkinter import *
from tkinter import ttk


class Glavn:
    def __init__(self, foo=None):
        self._foo = foo

    def Run (self):



        root = Tk()
        root.title("RAY")

        type_names = []
        comboboxes = []
        selected_params = []
        selected_type = None

        book = xlrd.open_workbook("DB1.xls")
        print("The number of worksheets is {0}".format(book.nsheets))
        print("Worksheet name(s): {0}".format(book.sheet_names()))
        sh = book.sheet_by_index(0)
        print("Лист {0} имеет  {1} строк и {2} столбцов".format(sh.name, sh.nrows, sh.ncols))
        print("Cell A1 is {0}".format(sh.cell_value(rowx=0, colx=0)))
        for rx in range(sh.nrows):
            type_names.append(sh.cell_value(rowx=rx, colx=0))

        # ШАГ 1. Выбор тиа схемы

        def type_selected(event):
            selected_type = type_names.index(typ_select_combobox.get())
            book2 = xlrd.open_workbook("DB{0}.xls".format(str(selected_type + 2)))
            self._foo = book2.sheet_by_index(0)
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
            for combobox in comboboxes:
                selected_params.append(combobox.get())
            if selected_params.count('') == 0:
                status_label = ttk.Label(text="Все поля заполнены")
                status_label.grid(row=len(selected_params) + 1, column=0)
                for col_index in range(self._foo.ncols):
                    if col_index > 0:
                        for row_index in range(self._foo.nrows):
                            if row_index > 0:
                                # no = (self._foo.cell_value(rowx=row_index, colx=col_index) == '')
                                # yes = (self._foo.cell_value(rowx=row_index, colx=col_index) != '')
                                if selected_params[row_index-1]:
                                    if (selected_params[row_index-1] == 'Да' and self._foo.cell_value(rowx=row_index, colx=col_index) != '') or (selected_params[row_index-1] == 'Нет' and self._foo.cell_value(rowx=row_index, colx=col_index) == ''):
                                        if row_index == len(selected_params)-1:
                                            print('Match found')
                                            status_label = ttk.Label(text="Обнаружен подходящий блок")
                                            status_label.grid(row=len(selected_params) + 1, column=0)
                                    else:
                                        print('Прервано')
                                        break
            else:
                status_label = ttk.Label(text="Не все поля заполнены")
                status_label.grid(row=len(selected_params) + 1, column=0)

        root.mainloop()


a = Glavn()
a.Run()















