import tkinter as tk
import sys
from tkinter import ttk
from create_excel import CreateExcel
from copy_folder import CopyFolder

def centralizar (root, largura_janela, altura_janela):
    # Obtém a largura e a altura da tela
    largura_tela = root.winfo_screenwidth()
    altura_tela = root.winfo_screenheight()
    # Calcula a posição x e y para centralizar a janela na tela
    pos_x = (largura_tela - largura_janela) // 2
    pos_y = (altura_tela - altura_janela) // 2
    # Define a geometria da janela para centralizá-la na tela
    return f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}"

if __name__ == '__main__':
    if getattr(sys, 'frozen', False):
        # Splash screen
        import pyi_splash

    root = tk.Tk()
    root.title("Casablanca Online")

    root.config(background = "white")

    root.geometry(centralizar(root, 450, 300))

    notebook = ttk.Notebook(root)

    tab1 = CreateExcel(notebook)
    notebook.add(tab1.tab, text="Gerar Excel")

    tab2 = CopyFolder(notebook)
    notebook.add(tab2.tab, text="Mover Arquivos")

    notebook.pack(fill="both", expand=True)

    if getattr(sys, 'frozen', False):
        # Splash screen close
        pyi_splash.close()

    root.mainloop()
