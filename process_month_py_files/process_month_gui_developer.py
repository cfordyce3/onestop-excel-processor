from process_month import process_month
from tkinter import *

WINDOW_WIDTH = 500
WINDOW_HEIGHT = 300

YEARS = ['2019', '2020', '2021',
         '2022', '2023', '2024',
         '2025', '2026', '2027',
         '2028', '2029', '2030']

root = Tk()

# window size
pos_right = int(root.winfo_screenwidth()/2 - WINDOW_WIDTH/2)
pos_down = int(root.winfo_screenheight()/2 - WINDOW_HEIGHT)
root.geometry('{}x{}+{}+{}'.format(WINDOW_WIDTH,WINDOW_HEIGHT,pos_right,pos_down))

# window title and icon
root.title('One Stop Excel Processor ~ Monthly')
root.iconbitmap('scu.ico')

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

# developer mode check
developer_mode = IntVar(root)
developer_mode.set(1)
developer_mode_check = Checkbutton(root,text='Developer mode',variable=developer_mode).grid(row=2,column=1)

# close button
Button(root,text='Close',command=root.destroy).grid(row=3,column=1,sticky=W)

def run_process_month (month,year,developer):
    process_month(month,year,developer)
    if (developer==1):
        print('Done')

# process button
Button(root,text='Process',command=lambda:run_process_month(month.get(),year.get(),developer_mode.get())).grid(row=3,column=2,sticky=E)


root.mainloop()
