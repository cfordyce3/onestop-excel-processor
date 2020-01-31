import os
from tkinter import *

from process_quarter import process_quarter
from resources import (YEARS, MONTHS, DAYS,
                       now_year, now_month, now_day,
                       get_stats_folder)

FILE_DIR = get_stats_folder()

WINDOW_WIDTH = 450
WINDOW_HEIGHT = 400

root = Tk()

# window size
pos_right = int(root.winfo_screenwidth()/2 - WINDOW_WIDTH/2)
pos_down = int(root.winfo_screenheight()/2 - WINDOW_HEIGHT)
root.geometry('{}x{}+{}+{}'.format(WINDOW_WIDTH,WINDOW_HEIGHT,pos_right,pos_down))

# window title and icon
root.title('One Stop Excel Processor')
if ('scu.ico' in os.listdir()): root.iconbitmap(os.getcwd() + '\\scu.ico')

# general tabling
root.grid_columnconfigure(0,weight=1)
root.grid_columnconfigure(4,weight=1)
root.grid_rowconfigure(2,weight=1)
root.grid_rowconfigure(7,weight=1)

# labels
Label(root,text='Select Quarter and Year').grid(row=0,column=1,pady=30,columnspan=3)
Label(root,text='Select Starting and Ending Dates for the Quarter').grid(row=2,column=1,pady=20,columnspan=3)
Label(root,text='Start Date').grid(row=3,column=1,pady=10,columnspan=3)
Label(root,text='End Date').grid(row=5,column=1,pady=10,columnspan=3)

# quarter dropdown
quarter = StringVar(root)
quarter.set('Fall')
quarter_menu = OptionMenu(root,quarter,'Fall','Winter','Spring','Summer').grid(row=1,column=1,sticky=W,columnspan=2)

# quarter year dropdown
quarter_year = StringVar(root)
quarter_year.set(now_year)
quarter_year_menu = OptionMenu(root,quarter_year,*YEARS).grid(row=1,column=2,sticky=W,columnspan=2)

# start month dropdown
start_month = StringVar(root)
start_month.set(MONTHS[now_month-1])
start_month_menu = OptionMenu(root,start_month,*MONTHS).grid(row=4,column=1,columnspan=3)

# start year dropdown
start_year = StringVar(root)
start_year.set(now_year)
start_year_menu = OptionMenu(root,start_year,*YEARS).grid(row=4,column=1,sticky=E,columnspan=3)

# start week dropdown
start_week = StringVar(root)
start_week.set(now_day)
start_week_menu = OptionMenu(root,start_week,*DAYS).grid(row=4,column=1,sticky=W,columnspan=3)

# end month dropdown
end_month = StringVar(root)
end_month.set(MONTHS[now_month-1])
end_month_menu = OptionMenu(root,end_month,*MONTHS).grid(row=6,column=1,columnspan=3)

# end year dropdown
end_year = StringVar(root)
end_year.set(now_year)
end_year_menu = OptionMenu(root,end_year,*YEARS).grid(row=6,column=1,sticky=E,columnspan=3)

# end week dropdown
end_week = StringVar(root)
end_week.set(now_day)
end_week_menu = OptionMenu(root,end_week,*DAYS).grid(row=6,column=1,sticky=W,columnspan=3)


# close button
Button(root,text='Close',command=root.destroy).grid(row=7,column=1,pady=30,sticky=W)

def run_process_quarter (quarter,start_week,start_month,start_year,end_week,end_month,end_year):
    process_quarter(quarter,start_week,start_month,start_year,end_week,end_month,end_year)

# process button
Button(root,text='Process',command=lambda:run_process_quarter(quarter.get(),start_week.get(),start_month.get(),start_year.get(),end_week.get(),end_month.get(),end_year.get())).grid(row=7,column=3,pady=30,sticky=E)

def run_gui (root):
    root.mainloop()

run_gui(root)
