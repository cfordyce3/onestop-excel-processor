from process_quarter import process_quarter
from tkinter import *

WINDOW_WIDTH = 450
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
root.title('One Stop Excel Processor ~ Quarterly')
root.iconbitmap('scu.ico')

# general tabling
root.grid_columnconfigure(0,weight=1)
root.grid_columnconfigure(4,weight=1)
root.grid_rowconfigure(2,weight=1)
root.grid_rowconfigure(7,weight=1)

# labels
Label(root,text='Select Starting and End Dates and the Quarter').grid(row=0,column=1,pady=30,columnspan=3)
Label(root,text='Start Date').grid(row=2,column=1,pady=10,columnspan=3)
Label(root,text='End Date').grid(row=4,column=1,pady=10,columnspan=3)

# quarter dropdown
quarter = StringVar(root)
quarter.set('Fall_1')
quarter_menu = OptionMenu(root,quarter,'Fall_1','Winter_1','Spring_1','Summer_1').grid(row=1,column=1,sticky=W,columnspan=2)

# start week dropdown
start_week = StringVar(root)
start_week.set('1')
start_week_menu = OptionMenu(root,start_week,'1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30','31').grid(row=3,column=1,sticky=W,columnspan=3)

# start month dropdown
start_month = StringVar(root)
start_month.set('January')
start_month_menu = OptionMenu(root,start_month,'January', 'February', 'March',
                                               'April', 'May', 'June',
                                               'July', 'August', 'September',
                                               'October', 'November', 'December').grid(row=3,column=1,columnspan=3)

# start year dropdown
start_year = StringVar(root)
start_year.set('2019')
start_year_menu = OptionMenu(root,start_year,'2019', '2020', '2021',
                                             '2022', '2023', '2024',
                                             '2025', '2026', '2027',
                                             '2028', '2029', '2030').grid(row=3,column=1,sticky=E,columnspan=3)

# end week dropdown
end_week = StringVar(root)
end_week.set('1')
end_week_menu = OptionMenu(root,end_week,'1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30','31').grid(row=5,column=1,sticky=W,columnspan=3)

# end month dropdown
end_month = StringVar(root)
end_month.set('January')
end_month_menu = OptionMenu(root,end_month,'January', 'February', 'March',
                                           'April', 'May', 'June',
                                           'July', 'August', 'September',
                                           'October', 'November', 'December').grid(row=5,column=1,columnspan=3)

# end year dropdown
end_year = StringVar(root)
end_year.set('2019')
end_year_menu = OptionMenu(root,end_year,'2019', '2020', '2021',
                                         '2022', '2023', '2024',
                                         '2025', '2026', '2027',
                                         '2028', '2029', '2030').grid(row=5,column=1,sticky=E,columnspan=3)


# close button
Button(root,text='Close',command=root.destroy).grid(row=6,column=1,pady=30,sticky=W)

def run_process_quarter (quarter,start_week,start_month,start_year,end_week,end_month,end_year):
    process_quarter(quarter,start_week,start_month,start_year,end_week,end_month,end_year)

# process button
Button(root,text='Process',command=lambda:run_process_quarter(quarter.get(),start_week.get(),start_month.get(),start_year.get(),end_week.get(),end_month.get(),end_year.get())).grid(row=6,column=3,pady=30,sticky=E)

root.mainloop()
