# Import packages

from tkinter import *
from tkinter import messagebox
from Main import *
import os

import pandas as pd

# Define functions

def button_Pressed(days_var: int):
    days_var = days_var.get()
    Run = Main(days_var)
    Run.main_processor()


def Run_GUI():
    # Initialise GUI

    GUI = Tk()

    GUI.geometry("400x150")
    GUI.title(" customers Automation Tool")

    # Set heading

    heading = Label(text = '''Enter number of days worth of  customers you wish to capture''',
                     bg = "black", fg = "white", height = "3", width = "600")

    heading.pack()

    # Create number of days input

    days_var = StringVar()

    days_func_text = Label(text = "Number of Days:")

    days_entry = Entry(textvariable = days_var)

    days_func_text.place(x = "40", y = "65")

    days_entry.place(x="140", y = "66")


    # Create and place button to generate customers

    gen_ = Button(text = "Generate  customer Files", command = lambda: button_Pressed(days_var))

    gen_.place(x= "4", y = "120", width = "392")

    GUI.mainloop()

Run_GUI()
