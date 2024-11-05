import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from openpyxl import Workbook

# Charger les données Excel
def load_data():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        df = pd.read_excel(file_path)
        return df
    return pd.DataFrame()

# Afficher les données dans l'interface tkinter
def display_data(df):
    for widget in frame.winfo_children():
        widget.destroy()

    if not df.empty:
        tree = ttk.Treeview(frame, columns=list(df.columns), show='headings')
        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=100)

        for index, row in df.iterrows():
            tree.insert("", tk.END, values=list(row))

        tree.pack(fill=tk.BOTH, expand=True)

# Filtrer les données selon Kaominina et Distrika
def filter_data(df):
    kaominina = kaominina_entry.get()
    distrika = distrika_entry.get()
    
    filtered_df = df
    if kaominina:
        filtered_df = filtered_df[filtered_df['Kaomina'] == kaominina]
    if distrika:
        filtered_df = filtered_df[filtered_df['Distrika'] == distrika]
    
    display_data(filtered_df)
    return filtered_df

# Exporter les données filtrées vers un fichier Excel
def export_to_excel(df):
    if df.empty:
        messagebox.showwarning("Avertissement", "Aucune donnée à exporter.")
        return
    
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Succès", f"Données exportées avec succès vers {file_path}")

# Interface utilisateur
root = tk.Tk()
root.title("Filtrer les Fokontany")

frame = tk.Frame(root)
frame.pack(fill=tk.BOTH, expand=True)

kaominina_label = tk.Label(root, text="Kaominina")
kaominina_label.pack()
kaominina_entry = tk.Entry(root)
kaominina_entry.pack()

distrika_label = tk.Label(root, text="Distrika")
distrika_label.pack()
distrika_entry = tk.Entry(root)
distrika_entry.pack()

load_button = tk.Button(root, text="Charger les Données", command=lambda: display_data(load_data()))
load_button.pack()

filter_button = tk.Button(root, text="Filtrer", command=lambda: filter_data(load_data()))
filter_button.pack()

export_button = tk.Button(root, text="Exporter vers Excel", command=lambda: export_to_excel(filter_data(load_data())))
export_button.pack()

root.mainloop()
