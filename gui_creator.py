import tkinter as tk
from process_month import process_month

WINDOW_WIDTH = 500
WINDOW_HEIGHT = 300

def startup ():
    startup_text.pack()
    weekly_box.pack()
    monthly_box.pack()
    quarterly_box.pack()
    next_button.pack()
    cancel_button.pack()


def startup_page():
    cancel_button.pack_forget()
    page2text.pack_forget()
    back_button.pack_forget()

    startup_text.pack()
    weekly_box.pack()
    monthly_box.pack()
    quarterly_box.pack()
    next_button.pack()
    cancel_button.pack()

def page2():
    cancel_button.pack_forget()
    startup_text.pack_forget()
    next_button.pack_forget()
    weekly_box.pack_forget()
    monthly_box.pack_forget()
    quarterly_box.pack_forget()

    page2text.pack()
    back_button.pack()
    cancel_button.pack()

root = tk.Tk()

# window size
pos_right = int(root.winfo_screenwidth()/2 - WINDOW_WIDTH/2)
pos_down = int(root.winfo_screenheight()/2 - WINDOW_HEIGHT)
root.geometry('{}x{}+{}+{}'.format(WINDOW_WIDTH,WINDOW_HEIGHT,pos_right,pos_down))

# window title and icon
root.title('One Stop Excel Processor')
root.iconbitmap('scu.ico')

# next, back, cancel buttons
cancel_button = tk.Button(root,text='Cancel',command=root.destroy)
next_button = tk.Button(root,text='Next',command=page2)
back_button = tk.Button(root,text='Back',command=startup_page)

# labels
startup_text = tk.Label(root, text='Which would you like to process?')
page2text = tk.Label(root, text="This is page 2")

# first page checkboxes
# SHOULD NOT BE CHECKBOXES, SHOULD ONLY SELECT ONE
weekly_onoff = tk.IntVar()
weekly_box = tk.Checkbutton(root,text='Weekly',variable=weekly_onoff,onvalue=1,offvalue=0)

monthly_onoff = tk.IntVar()
monthly_box = tk.Checkbutton(root,text='Montly',variable=monthly_onoff,onvalue=1,offvalue=0)

quarterly_onoff = tk.IntVar()
quarterly_box = tk.Checkbutton(root,text='Quarterly',variable=quarterly_onoff,onvalue=1,offvalue=0)

startup()
root.mainloop()
