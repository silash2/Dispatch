import openpyxl 
import xlrd 
import subprocess
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from pathlib import *
from datetime import datetime


background = "#06283D"

def interface_1():
    root.destroy()
    subprocess.run(["python", "interface_1.py"])

def dispatch():
    root.destroy()
    subprocess.run(["python", "Dispatch/telecharge.py"])

#interface graphique
root = Tk()
root.title("Celero Express")
root.geometry("650x600+150+100")
root.config(bg=background)
root.resizable(width= False, height=False)


#entete de l'application
Label(root, text="CELERO EXPRESS FACILITY SOLUTION",width=10,height=3,bg="white",fg='blue',font="arial 20 bold").pack(side= TOP, fill=X)
#Bottom application
Label(root, text="E-mail: wzafitsara@gmail.com", width=10,height=2, anchor='e').pack(side=BOTTOM, fill=X)

#dispatch button
Button(root, text="Dispatch parcel", bd=2, width=30, font='arial 16 bold',command=dispatch).place(x=120, y=200)

#cession doc
Button(root, text="Cession document", bd=2, width=30, font='arial 16 bold', command=interface_1).place(x=120, y=260)

#Update shipment
Button(root, text="Mis a jour", bd=2, width=30, font='arial 16 bold').place(x=120, y=320)



root.mainloop()