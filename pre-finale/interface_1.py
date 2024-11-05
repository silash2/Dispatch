import openpyxl
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from tkinter.ttk import Combobox, Scrollbar, Treeview
import subprocess
background = "#06283D"


root =Tk()
root.title("Cession document")
root.geometry("700x350+700+100")
root.config(bg=background)
root.resizable(width= False, height=False)

def cession_autres():
    root.destroy()
    subprocess.run(["python", "cession/cession_autres.py"])

def cession_aramex():
    root.destroy()
    subprocess.run(["python", "cession/cession_aramex.py"])

#cession document aramex 
Button(root, text="Cession Aramex",width=30, bd=2, font='arial 16 bold',bg='white', command=cession_aramex).place(x=150, y=120)

#cession document autres
Button(root, text="Cession Autres", width=30, bd=2, font='arial 16 bold', bg='green', command=cession_autres).place(x=150, y=220)


#entete de l'application
Label(root, text="CELERO EXPRESS CESSION DOCUMENT",width=10,height=3,bg="white",fg='blue',font="arial 20 bold").pack(side= TOP, fill=X)

#Bottom application
Label(root, text="E-mail: wzafitsara@gmail.com", width=10,height=2, anchor='e').pack(side=BOTTOM, fill=X)



root.mainloop()