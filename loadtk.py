import os
from tkinter import *
import pandas as pd

master = Tk()

tkvar1 = StringVar()
  
values = ['Load All Data', 'Last 3 Rows'] #
tkvar1.set("Load All Data")
popMenu = OptionMenu(master, tkvar1, *values) # define popup menu

def load_all():
        # remove restrictions to allow returning of all data
        
        pd.set_option('display.max_rows', None)
        pd.set_option('display.max_columns', None)
        pd.set_option('display.width', None)
        pd.set_option('display.max_colwidth', -1)
        
        df = pd.ExcelFile(file.get()).parse(e2.get())

        df.to_excel(e1.get(), sheet_name=e2.get())
        
        return 'process finished!'
    # function to load into more than one sheet
def load_to_many_sheets(sheet_num):
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    pd.set_option('display.max_colwidth', -1)
    
    df = pd.ExcelFile(file.get()).parse(e2.get())
    name = 'Sheet'
    for i in range(sheet_num):  
        d = df.copy()
        d.to_excel(pd.ExcelWriter('output.xlsx', name+str(i+1)))
    return 'process finished'

    # function to load only last 3 rows of data
def load_last_3_rows():
    pd.set_option('display.max_rows', 3)
    df = pd.ExcelFile(file.get()).parse(e2.get())
    last3rows = df.tail(3)

    
    last3rows.to_excel(e1.get(), sheet_name=e2.get())
    return 'process finished !'

def check():
    if tkvar1.get() == 'Load All Data':
        load_all()
    elif tkvar1.get() == 'Future Value':
        load_last_3_rows()

file = Entry(master)
file.grid(row=0, column=1)


l = Label(master, text='Insert Excel File')
l.grid(row=0, column=0)
 
l1 = Label(master, text='Output File')
l1.grid(row=1, column=0)

e1 = Entry(master)
e1.grid(row=1, column=1)

popMenu.grid(row=0, column=2)

e2 = Entry(master)
e2.grid(row=2, column=1)

l2 = Label(master, text='Sheet Name')
l2.grid(row=2, column=0)

b1 = Button(master, text="Run", command=check)
b1.grid(row=2, column=2)

master.geometry("330x130")

mainloop()
