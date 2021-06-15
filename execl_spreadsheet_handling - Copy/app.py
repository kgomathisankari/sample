from appending_data_to_excel_file import appending_data_func
from tkinter import messagebox
from tkinter import *
import os

frame = Tk()
frame.title("Aravind Excel App")

data_var = StringVar()
subject_var = StringVar()
chapter_var = StringVar()
duration_var = StringVar()


def action():
    appending_data_func(
        data=data_var.get(),
        subject=subject_var.get(),
        chapter=chapter_var.get(),
        duration=duration_var.get()
    )
    messagebox.showinfo("Success", "Your study details have been updated successfully !")


def submit_and_open_func():
    appending_data_func(
        data=data_var.get(),
        subject=subject_var.get(),
        chapter=chapter_var.get(),
        duration=duration_var.get()
    )
    file = "G:\Work\execl_spreadsheet_handling\Aravind_Studies.xlsx"
    os.startfile(file)


welcome_label = Label(frame, text="Enter the below details to update your Excel Sheet !", fg="red").grid(row=0, column=0)

data_info_label = Label(frame, text="Enter the date in this format : dd/mm/yyyy", fg="green").grid(row=1, column=0)
data_label = Label(frame, text="Enter the Date : ").grid(row=2, column=0)
data_entry = Entry(frame, textvariable=data_var).grid(row=2, column=1)

subject_label = Label(frame, text="Enter the Subject : ").grid(row=3, column=0)
subject_entry = Entry(frame, textvariable=subject_var).grid(row=3, column=1)

chapter_label = Label(frame, text="Enter the Chapter : ").grid(row=4, column=0)
chapter_entry = Entry(frame, textvariable=chapter_var).grid(row=4, column=1)

duration_label = Label(frame, text="Enter the Duration : ").grid(row=5, column=0)
duration_entry = Entry(frame, textvariable=duration_var).grid(row=5, column=1)

submit_button = Button(frame, text="Submit", fg="blue", command=action).grid(row=6, column=0)
submit_and_open_button = Button(frame, text="Submit and Open", fg="blue", command=submit_and_open_func).grid(row=6, column=1)

frame.mainloop()
