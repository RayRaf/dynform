import xlrd
from tkinter import *
from tkinter import ttk

labels_values = []
comboboxes = []
selections = []

book = xlrd.open_workbook("DB.xls")
print("The number of worksheets is {0}".format(book.nsheets))
print("Worksheet name(s): {0}".format(book.sheet_names()))
sh = book.sheet_by_index(0)
print("Лист {0} имеет  {1} строк и {2} столбцов".format(sh.name, sh.nrows, sh.ncols))
print("Cell A1 is {0}".format(sh.cell_value(rowx=0, colx=0)))
for rx in range(sh.ncols):
    print(sh.cell_value(rowx=0, colx=rx))
    labels_values.append(sh.cell_value(rowx=0, colx=rx))

root = Tk()
root.title("RAY")
# root.geometry("250x200")
# btn = ttk.Button(text="Button")
# btn = ttk.Button(text="Button", command=click)

# i = 0
# for val in labelsvalues:
#     label = ttk.Label(text=val)
#     label.grid(row=i, column=0)
#     combobox = ttk.Combobox(values=['Да', 'Нет'])
#     combobox.grid(row=i, column=1)
#     print(str(i))
#     i += 1


def selected(event):
    i1 = 0
    for combobox in comboboxes:
        if combobox.get() != selections[i1]:
            print('Изменен комбобокс ' + str(i1))
            selections[i1] = combobox.get()
        i1 += 1



    # for selection in selections:
    #     if selection != selections[i1]:
    #         label["text"] = f"вы выбрали: {selections[i1]} в комбобоксе"
    #         # label.grid(row=5, column=1)
    #         selections[i1] = selection


    # label["text"] = f"вы выбрали: {selections} в комбобоксе"
    # label.grid(row=1, column=1)


i = 0


for val in labels_values:
    label = ttk.Label(text=val)
    label.grid(row=i, column=0)
    comboboxes.append(ttk.Combobox(values=['Да', 'Нет']))
    comboboxes[i].grid(row=i, column=1)
    comboboxes[i].bind("<<ComboboxSelected>>", selected)
    selections.append(comboboxes[i].get())
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

