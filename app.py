import tkinter as tk
from tkinter import ttk
import openpyxl
import time
from PIL import Image, ImageTk

def load_data():
    path = "Peoples.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    list_values = list(sheet.values)
    for idx, col in enumerate(list_values[0]):
        treeview.heading('#' + str(idx+1), text=col)
    for value_tuple in (list_values[1:]):
        treeview.insert('',tk.END,values=value_tuple)


def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")


def insert_new():
    name = name_entry.get()
    age = int(age_spinbox.get())
    gender = status_combo.get()
    employment_status = "Employed" if a.get() else "Unemployed"
    time.sleep(0.5)
    path = "Peoples.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [name, age, gender, employment_status]
    sheet.append(row_values)
    workbook.save(path)
    treeview.insert('', tk.END, values=row_values)

    feedback_label.config(text="Data inserted successfully!", foreground="green")
    root.after(2000, clear_feedback)
    name_entry.delete(0, tk.END)
    age_spinbox.delete(0, tk.END)
    status_combo.set(el[0])
    a.set(False)

def clear_feedback():
    feedback_label.config(text="")


root = tk.Tk()
style = ttk.Style(root)
root.tk.call("source","forest-light.tcl")
root.tk.call("source","forest-dark.tcl")
root.title("Data Entry")
style.theme_use("forest-dark")

logo_img = Image.open("logo.png")
logo_img = logo_img.resize((200, 200))
logo_img = ImageTk.PhotoImage(logo_img)

# Combine logo with app title
title_frame = ttk.Frame(root)
title_frame.pack(pady=10)
logo_label = ttk.Label(title_frame, image=logo_img)
logo_label.grid(row=0, column=0)
# title_label = ttk.Label(title_frame, text="", font=("Helvetica", 18, "bold"))
# title_label.grid(row=0, column=1, padx=10)



frame = ttk.Frame(root)
frame.pack()

widget_frame = ttk.LabelFrame(frame,text = "Insert Row",padding=(0,10))
widget_frame.grid(row=0,column=0,padx=20,pady=10)

name_entry = ttk.Entry(widget_frame)
name_entry.grid(row=0,column=0,sticky='ew',padx=10,pady=(0,5))
name_entry.insert(0,"Name")
name_entry.bind("<FocusIn>",lambda e:name_entry.delete('0','end'))


age_spinbox = ttk.Spinbox(widget_frame,from_=18,to=100)
age_spinbox.grid(row=1,column=0,sticky='ew',padx=10,pady=10)
age_spinbox.insert(0,"Age")
age_spinbox.bind("<FocusIn>",lambda e:age_spinbox.delete('0','end'))

el = ['']
combo_list = ["Male","Female","Other"]
status_combo = ttk.Combobox(widget_frame,values=combo_list)
status_combo.current(0)
status_combo.grid(row=2,column=0,sticky='ew',padx=10,pady=10)

a = tk.BooleanVar()
check_button = ttk.Checkbutton(widget_frame,text = "Employed",variable=a)
check_button.grid(row=3,column=0,sticky="nsew",padx=10,pady=10)

button = ttk.Button(widget_frame,text= "Insert",command=insert_new)
button.grid(row=4,column=0,sticky='nsew',padx=10,pady=10)

seperator = ttk.Separator(widget_frame)
seperator.grid(row=5,column=0,padx=(20,15),pady=10,sticky='ew')

mode_switch = ttk.Checkbutton(widget_frame,text="Mode",style="Switch",command=toggle_mode)
mode_switch.grid(row=6,column=0,padx=20,sticky='nsew')

feedback_label = ttk.Label(widget_frame, text="", foreground="green")
feedback_label.grid(row=7, column=0, padx=10, pady=5, sticky='ew')


treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0,column=1,pady=20)
treescroll = ttk.Scrollbar(treeFrame)
treescroll.pack(side="right",fill="y")

cols = ("Name","Age","Gender","Employment")
treeview = ttk.Treeview(treeFrame,show="headings",columns=cols,height=13,
                       yscrollcommand=treescroll.set)
treeview.column("Name",width=150)
treeview.column("Age",width=30)
treeview.column("Gender",width=50)
treeview.column("Employment",width=90)
treeview.pack()
treescroll.config(command=treeview.yview)
load_data()


root.mainloop()