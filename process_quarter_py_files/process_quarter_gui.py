from process_quarter import process_quarter
from tkinter import *

WINDOW_WIDTH = 500
WINDOW_HEIGHT = 400

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
root.grid_columnconfigure(4,weight=1)
root.grid_rowconfigure(2,weight=1)
root.grid_rowconfigure(4,weight=1)

# labels
Label(root,text='Select Starting and End Dates and the Quarter').grid(row=0,column=1,pady=30,columnspan=3)
Label(root,text='Start Date').grid(row=2,column=1,pady=10,columnspan=3)
Label(root,text='End Date').grid(row=4,column=1,pady=10,columnspan=3)

# quarter dropdown
quarter = StringVar(root)
quarter.set('Fall')
quarter_menu = OptionMenu(root,quarter,'Fall','Winter','Spring','Summer').grid(row=1,column=1,sticky=W)

# start week dropdown
start_week = StringVar(root)
start_week.set('1')
start_week_menu = OptionMenu(root,start_week,'2','3','4','5').grid(row=3,column=1,sticky=W)

# start month dropdown
start_month = StringVar(root)
start_month.set('January')
start_month_menu = OptionMenu(root,start_month,'January', 'February', 'March',
                                   'April', 'May', 'June',
                                   'July', 'August', 'September',
                                   'October', 'November', 'December').grid(row=3,column=2)

# start year dropdown
start_year = StringVar(root)
start_year.set('2019')
start_year_menu = OptionMenu(root,start_year,'2019', '2020', '2021',
                                 '2022', '2023', '2024',
                                 '2025', '2026', '2027',
                                 '2028', '2029', '2030').grid(row=3,column=3,sticky=E)

# end week dropdown
start_week = StringVar(root)
start_week.set('1')
start_week_menu = OptionMenu(root,start_week,'2','3','4','5').grid(row=5,column=1,sticky=W)

# end month dropdown
start_month = StringVar(root)
start_month.set('January')
start_month_menu = OptionMenu(root,start_month,'January', 'February', 'March',
                                   'April', 'May', 'June',
                                   'July', 'August', 'September',
                                   'October', 'November', 'December').grid(row=5,column=2)

# end year dropdown
start_year = StringVar(root)
start_year.set('2019')
start_year_menu = OptionMenu(root,start_year,'2019', '2020', '2021',
                                 '2022', '2023', '2024',
                                 '2025', '2026', '2027',
                                 '2028', '2029', '2030').grid(row=5,column=3,sticky=E)

# developer mode check
#developer_mode = IntVar(root)
#developer_mode.set(0)
#developer_mode_check = Checkbutton(root,text='Developer mode',variable=developer_mode).grid(row=2,column=1)

# close button
Button(root,text='Close',command=root.destroy).grid(row=6,column=1,pady=30,sticky=W)

def run_process_month (month,year):
    pass
    #process_quarter(month,year)
    #if (developer==1):
    #    print('Done')

# process button
Button(root,text='Process',command=lambda:run_process_month(month.get(),year.get())).grid(row=6,column=3,pady=30,sticky=E) #removed developer_mode.get()


root.mainloop()
