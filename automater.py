from starter_mail import automater
import tkinter.filedialog
import tkinter as tk

def selector():
    """Empty function, returns nothing but helps select the Excel file containing the data """
    global root
    root.filename = tkinter.filedialog.askopenfilename(title="Select Excel File")
    print(root.filename)
def sender():
    """Empty function, returns nothing but calls the automater() function after the fields 
    of the window have been filled"""
    automater(email.get(), password.get(), var.get(), date.get(), subj.get(), root.filename)
    root.destroy()

root = tk.Tk(); root.wm_title('AutoMailer in Python')
root.minsize(width=400, height=400)
w = tk.Label(root, text="\n\n"); w.pack()
sel_button = tk.Button(root, text='Select Excel File', command=selector); sel_button.pack()
w = tk.Label(root, text='Email:'); w.pack()
email = tk.Entry(root); email.pack()
w = tk.Label(root, text='Password:'); w.pack()
password = tk.Entry(root, show='*'); password.pack()
w = tk.Label(root, text='Current Month:'); w.pack()
var = tk.StringVar(root); var.set("January")
w = tk.Label(root, text='Subject'); w.pack()
subj = tk.Entry(root); subj.pack()
current_month = tk.OptionMenu(root, var, "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"); current_month.pack()
w = tk.Label(root, text='Date:'); w.pack()
date = tk.Entry(root); date.pack()
root.filename=None
send_button = tk.Button(root, text='Send', command=sender); send_button.pack()
root.mainloop()