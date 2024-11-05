from tkinter import *
from tkinter import ttk
from tkinter.ttk import Combobox, Scrollbar, Treeview
import pandas
import os
from tkinter import Tk, Label, Button, filedialog, messagebox
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import shutil
import subprocess
from datetime import datetime

background = "#06283D"
def select_and_process():

# Demander à l'utilisateur de sélectionner un fichier Excel
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx")])

    # Vérifier si l'utilisateur a sélectionné un fichier
    if file_path:
        # Extraire le nom de fichier à partir du chemin complet
        nom_fichier = os.path.basename(file_path)
        # Charger le fichier Excel dans un DataFrame
        df = pd.read_excel(file_path)

        # Sélectionner uniquement les colonnes spécifiques
        columns_of_interest = ['AWB', 'PCS', 'Weight', 'ConsigneeName', 'ConsigneeAddress',
                               'ConsigneeTel', 'Receiver-email', 'DestCity']
        
        # Vérifier si toutes les colonnes spécifiées sont présentes dans le fichier
        for col in columns_of_interest:
            if col not in df.columns:
                print(f"La colonne '{col}' n'est pas présente dans le fichier.")
                return
        
        # Créer un nouveau DataFrame avec seulement les colonnes d'intérêt
        df_filtered = df[columns_of_interest]

        # Créer le dossier 'data/data_usable' s'il n'existe pas encore
        date_str = datetime.now().strftime("%d-%m-%Y")
        save_folder = f'Manifest/manifest{date_str}'
        if not os.path.exists(save_folder):
            os.makedirs(save_folder)


        # Nom de fichier de sauvegarde dans le dossier 'data/data_usable'
        save_path = os.path.join(save_folder, f"{nom_fichier}_donnees_filtrees.xlsx")

        # Enregistrer le DataFrame filtré dans un nouveau fichier Excel
        if save_path:
            df_filtered.to_excel(save_path, index=False)

        messagebox.showinfo("info", f"Données filtrées enregistrées avec succès dans '{save_path}'.")

        root.destroy()
        subprocess.run(["python", "./sican.py"])

    
def dispatch():
    root.destroy()
    subprocess.run(["python", "./sican.py"])
def retour():
    root.destroy()
    subprocess.run(["python", "interface_0.py"])
def quiter():
    root.destroy()

#interface grapthique
root = Tk()
root.title('telechargement de manifest')
root.geometry("500x450")
root.config(bg=background)
root.resizable(width= False, height=False)


#entete de l'application
Label(root, text="Telecharger fichier manifest",width=10,height=3,bg="white",fg='blue',font="arial 16 bold").pack(side= TOP, fill=X)
#Bottom application
Label(root, text="E-mail: wzafitsara@gmail.com", width=10,height=2, anchor='e').pack(side=BOTTOM, fill=X)

#button pour telecharger manifest
label = Label(root, text="Sélectionnez un fichier manifest", bg=background, width=30, font='arial 15 bold', foreground='white').place(x=70, y=100)
button = Button(root, text="Sélectionner le fichier", command=select_and_process, bg="green", width=30, font='arial 12 bold')
button.place(x=100, y=150)

#button pour aller directement vers test1
Button(root, text="Manifest exist deja",foreground="blue", command=dispatch, width=30, font='arial 12 bold').place(x=100, y=200)

#button retour
Button(root, text="Retour", width= 30, font='arial 12 bold', command=retour).place(x=100, y=250)

#button quiter
Button(root, text="Quiter", width=30, foreground="red", font='arial 12 bold', command=quiter).place(x=100, y=300)

root.mainloop()