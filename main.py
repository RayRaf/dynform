import xlrd
from tkinter import *
from tkinter import ttk

labelsvalues = []
comboboxes = []

book = xlrd.open_workbook("DB.xls")
print("The number of worksheets is {0}".format(book.nsheets))
print("Worksheet name(s): {0}".format(book.sheet_names()))
sh = book.sheet_by_index(0)
print("Лист {0} имеет  {1} строк и {2} столбцов".format(sh.name, sh.nrows, sh.ncols))
print("Cell A1 is {0}".format(sh.cell_value(rowx=0, colx=0)))
for rx in range(sh.ncols):
    print(sh.cell_value(rowx=0, colx=rx))
    labelsvalues.append(sh.cell_value(rowx=0, colx=rx))

root = Tk()
root.title("RAY")
# root.geometry("250x200")
# btn = ttk.Button(text="Button")
# btn = ttk.Button(text="Button", command=click)

i = 0
for val in labelsvalues:
    label = ttk.Label(text=val)
    label.grid(row=i, column=0)
    combobox = ttk.Combobox(values=['Да', 'Нет'])
    combobox.grid(row=i, column=1)
    print(str(i))
    i += 1


# def single_click(event):
#     combobox = ttk.Combobox(values=languages)
#     combobox.pack(anchor=NW, padx=6, pady=6)
#
#
# def double_click(event):
#     btn["text"] = "Double Click"


# btn = ttk.Button(text="Click")
# btn.pack(anchor=CENTER, expand=1)
#
# btn.bind("<ButtonPress-1>", single_click)
# btn.bind("<Double-ButtonPress-1>", double_click)

root.mainloop()

