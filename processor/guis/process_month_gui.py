import os
from processor.month.process_month import process_month
from tkinter import *

WINDOW_WIDTH = 400
WINDOW_HEIGHT = 200
FILE_DIR = os.path.abspath('..')

YEARS = []
import datetime
now_year = datetime.datetime.now().year
for year in range(-2,6):
    YEARS.append(now_year + year)

root = Tk()

# window size
pos_right = int(root.winfo_screenwidth()/2 - WINDOW_WIDTH/2)
pos_down = int(root.winfo_screenheight()/2 - WINDOW_HEIGHT)
root.geometry('{}x{}+{}+{}'.format(WINDOW_WIDTH,WINDOW_HEIGHT,pos_right,pos_down))

# window title and icon
root.title('One Stop Excel Processor ~ Monthly')
root.iconbitmap(FILE_DIR + '\\resources\\scu.ico')

# general tabling
root.grid_columnconfigure(0,weight=1)
root.grid_columnconfigure(3,weight=1)
root.grid_rowconfigure(2,weight=1)
root.grid_rowconfigure(4,weight=1)

# header
Label(root,text='Select Which Month and Year to Process').grid(row=0,column=1,pady=30,columnspan=2)

# month dropdown
month = StringVar(root)
month.set('January')
month_menu = OptionMenu(root,month,'January', 'February', 'March',
                                   'April', 'May', 'June',
                                   'July', 'August', 'September',
                                   'October', 'November', 'December').grid(row=1,column=1)

# year dropdown
year = StringVar(root)
year.set('2019')
year_menu = OptionMenu(root,year,'2019', '2020', '2021',
                                 '2022', '2023', '2024',
                                 '2025', '2026', '2027',
                                 '2028', '2029', '2030').grid(row=1,column=2)

# close button
Button(root,text='Close',command=root.destroy).grid(row=3,column=1,sticky=W)

def run_process_month (month,year):
    process_month(month,year)

# process button
Button(root,text='Process',command=lambda:run_process_month(month.get(),year.get())).grid(row=3,column=2,sticky=E) #removed developer_mode.get()


root.mainloop()
