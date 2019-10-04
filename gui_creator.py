from tkinter import *

WINDOW_WIDTH = 500
WINDOW_HEIGHT = 300

def next_button (which_next,weekly,monthly,quarterly):
    second_window(weekly,monthly,quarterly)



def second_window (weekly,monthly,quarterly):
    root = Tk()

    # title and icon
    root.title('One Stop Excel Processor')

    # screen size and position formatting
    pos_right = int(root.winfo_screenwidth()/2 - WINDOW_WIDTH/2)
    pos_down = int(root.winfo_screenheight()/2 - WINDOW_HEIGHT)
    root.geometry('{}x{}+{}+{}'.format(WINDOW_WIDTH,WINDOW_HEIGHT,pos_right,pos_down))

    # cancel and next buttons
    cancel_button = Button(root,text='Cancel',command=root.destroy)
    cancel_button.grid(row=4,column=0,ipadx=25)

    next_button = Button(root,text='Next')
    next_button.grid(row=4,column=2,ipadx=25)

root = Tk()

# title and icon
root.title('One Stop Excel Processor')

# screen size and position formatting
pos_right = int(root.winfo_screenwidth()/2 - WINDOW_WIDTH/2)
pos_down = int(root.winfo_screenheight()/2 - WINDOW_HEIGHT)
root.geometry('{}x{}+{}+{}'.format(WINDOW_WIDTH,WINDOW_HEIGHT,pos_right,pos_down))

# checkboxes for weekly, monthly, and quarterly processing
weekly_onoff = IntVar()
weekly_box = Checkbutton(root,text='Weekly',variable=weekly_onoff,onvalue=1,offvalue=0)
weekly_box.grid(row=1,column=1,sticky=W)

monthly_onoff = IntVar()
monthly_box = Checkbutton(root,text='Montly',variable=monthly_onoff,onvalue=1,offvalue=0)
monthly_box.grid(row=2,column=1,sticky=W)

quarterly_onoff = IntVar()
quarterly_box = Checkbutton(root,text='Quarterly',variable=quarterly_onoff,onvalue=1,offvalue=0)
quarterly_box.grid(row=3,column=1,sticky=W)

# label for the question boxes
which_processing = Label(root,text='Which would you like to process?')
which_processing.grid(row=0,column=1)

# cancel and next buttons
cancel_button = Button(root,text='Cancel',command=root.destroy)
cancel_button.grid(row=4,column=0,ipadx=25)

next_button = Button(root,text='Next',command=lambda:[root.destroy,next_button(root,weekly_onoff,monthly_onoff,quarterly_onoff)])
next_button.grid(row=4,column=2,ipadx=25)

# overall formatting
root.grid_rowconfigure(0,weight=1)
root.grid_rowconfigure(4,weight=1)
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(2, weight=1)

root.mainloop()
