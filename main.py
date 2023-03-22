import xlrd
from tkinter import *
from tkinter import ttk

root = Tk()
root.title("RAY")

type_names = []
comboboxes = []
selections = []

book = xlrd.open_workbook("DB1.xls")
print("The number of worksheets is {0}".format(book.nsheets))
print("Worksheet name(s): {0}".format(book.sheet_names()))
sh = book.sheet_by_index(0)
print("Лист {0} имеет  {1} строк и {2} столбцов".format(sh.name, sh.nrows, sh.ncols))
print("Cell A1 is {0}".format(sh.cell_value(rowx=0, colx=0)))
for rx in range(sh.nrows):
    # print(sh.cell_value(rowx=0, colx=rx))
    type_names.append(sh.cell_value(rowx=rx, colx=0))


def selected(event):
    selected_type = type_names.index(typ_select_combobox.get())

    book2 = xlrd.open_workbook("DB{0}.xls". format(str(selected_type+1)))
    sh1 = book2.sheet_by_index(0)
    for rx1 in range(sh1.nrows):
        comboboxes.append(ttk.Combobox(values=['Да', 'Нет']))
        comboboxes[rx1].grid(row=rx1, column=1)
        label = ttk.Label(text=sh1.cell_value(rowx=rx1, colx=0))
        label.grid(row=rx1, column=0)



label = ttk.Label(text="Тип схемы:")
label.grid(row=0, column=0)

typ_select_combobox = ttk.Combobox(values=type_names)
typ_select_combobox.grid(row=0, column=1)
typ_select_combobox.bind("<<ComboboxSelected>>", selected)











#
# def selected(event):
#     i1 = 0
#     for combobox in comboboxes:
#         if combobox.get() != selections[i1]:
#             print('Изменен комбобокс ' + str(i1))
#             selections[i1] = combobox.get()
#         i1 += 1
#
#
# i = 0
# for val in type_names:
#     label = ttk.Label(text=val)
#     label.grid(row=i, column=0)
#     comboboxes.append(ttk.Combobox(values=['Да', 'Нет']))
#     comboboxes[i].grid(row=i, column=1)
#     comboboxes[i].bind("<<ComboboxSelected>>", selected)
#     selections.append(comboboxes[i].get())
#     print(str(i))
#     i += 1
#
#
#
#
#
#
#
















root.mainloop()
