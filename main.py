import tkinter as tk
from tkinter import ttk
from create_excel import CreateExcel
from copy_folder import CopyFolder

root = tk.Tk()
root.title("Casablanca Online")
root.geometry("350x250")
root.config(background = "white")

notebook = ttk.Notebook(root)

tab1 = CreateExcel(notebook)
notebook.add(tab1.tab, text="Gerar Excel")

tab2 = CopyFolder(notebook)
notebook.add(tab2.tab, text="Copiar Arquivos")

notebook.pack(fill="both", expand=True)

root.mainloop()
