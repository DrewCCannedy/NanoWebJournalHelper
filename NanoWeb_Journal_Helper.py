"""Python app that helps with NanoWeb Journal Tasks."""
"""1) Removes excel sheets with zero as the cost ammount."""
"""2) Sends emails to people who used NanoWeb facilities."""
"""3) Creates a file of all usage of NanoWeb for a certain month."""

import tkinter as tk
from tkinter import Label, Entry, messagebox
import remove_zeros
import automatic_mailer
import fill_jrnl1

gui = tk.Tk()
gui.title("NanoWeb Journal Helper")

def remove_zeros_call():
    folder = folder_field.get()
    if folder == "":
        messagebox.showinfo("Error!", "Enter a folder name first!")
    else:
        result = remove_zeros.remove_zeros(folder)
        messagebox.showinfo("Results", result)

def automatic_mailer_call():
    folder = folder_field.get()
    if folder == "":
        messagebox.showinfo("Error!", "Enter a folder name first!")
    else:
        a = automatic_mailer.Automatic_mailer(folder)
        a.run()
        result = a.output_string
        messagebox.showinfo("Results", result)

def fill_jrnl1_call():
    folder = folder_field.get()
    if folder == "":
        messagebox.showinfo("Error!", "Enter a folder name first!")
    else:
        result = fill_jrnl1.write_template(folder)
        messagebox.showinfo("Results", result)

remove_zeros_button = tk.Button(gui, text = "Remove Zeroes",command = remove_zeros_call)
mailer_button = tk.Button(gui, text = "Send Emails",command = automatic_mailer_call)
fill_button = tk.Button(gui, text = "Create Monthly Report",command = fill_jrnl1_call)
folder_label = Label(gui, text="Folder Name", fg="black")
folder_field = Entry(gui)

remove_zeros_button.grid(column=0, row = 1)
mailer_button.grid(column=1, row = 1)
fill_button.grid(column=2, row = 1)
folder_label.grid(column=0, row=0)
folder_field.grid(column=1, row=0)
gui.mainloop()