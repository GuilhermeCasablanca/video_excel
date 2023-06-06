import tkinter as tk
from tkinter import ttk
from create_excel import CreateExcel
from copy_folder import CopyFolder

root = tk.Tk()
root.title("Casablanca Online")
root.geometry("500x400")
root.config(background = "white")

# Grid method is chosen for placing
# the widgets at respective positions
# in a table like structure by
# specifying rows and columns
#label_file_explorer.grid(column = 1, row = 1)
  
#button_explore.grid(column = 1, row = 2)
  
#button_exit.grid(column = 1,row = 3)

notebook = ttk.Notebook(root)

tab1 = CreateExcel(notebook)
notebook.add(tab1.tab, text="Gerar Excel")

tab2 = CopyFolder(notebook)
notebook.add(tab2.tab, text="Copiar Arquivos")

notebook.pack(fill="both", expand=True)

root.mainloop()
