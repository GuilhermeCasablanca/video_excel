import tkinter as tk
from tkinter import filedialog
import os
import csv
import datetime
import time
from moviepy.editor import VideoFileClip
import pandas as pd
from tkinter import ttk

class CopyFolder:
    def __init__(self, parent):
        self.parent = parent
        self.tab = ttk.Frame(parent)

        # Add your GUI elements here
        # Arquivo Excel
        self.label_excel = ttk.Label(self.tab, text="Arquivo Excel:")
        self.entry_excel = ttk.Entry(self.tab, width=40)
        self.button_excel = ttk.Button(self.tab, text="Browse", command=self.browse_excel)

        # Diretório dos Arquivos
        self.label_files = ttk.Label(self.tab, text="Diretório dos Arquivos:")
        self.entry_files = ttk.Entry(self.tab, width=40)
        self.button_files = ttk.Button(self.tab, text="Browse", command=self.browse_files)

        # Diretório de Destino
        self.label_destination = ttk.Label(self.tab, text="Diretório de Destino:")
        self.entry_destination = ttk.Entry(self.tab, width=40)
        self.button_destination = ttk.Button(self.tab, text="Browse", command=self.browse_destination)

        # Copiar Arquivos
        self.save_button = ttk.Button(self.tab, text="Copiar Arquivos", command=self.save_directory)

        # Limpar
        self.clear_button = ttk.Button(self.tab, text="Limpar", command=self.clear)

        # Formatacao
        self.label_excel.grid(column = 1, row = 1, sticky=tk.W)
        self.entry_excel.grid(column = 1, row = 2, padx=5, pady=5)
        self.button_excel.grid(column = 2, row = 2, padx=5, pady=5)

        self.label_files.grid(column = 1, row = 3, sticky=tk.W)
        self.entry_files.grid(column = 1, row = 4, padx=5, pady=5)
        self.button_files.grid(column = 2, row = 4, padx=5, pady=5)
        
        self.label_destination.grid(column = 1, row = 5, sticky=tk.W)
        self.entry_destination.grid(column = 1, row = 6, padx=5, pady=5)
        self.button_destination.grid(column = 2, row = 6, padx=5, pady=5)

        self.save_button.grid(column = 1, row = 7, padx=5, pady=5)

        self.clear_button.grid(column = 2, row = 7, padx=5, pady=5)

        # Variaveis
        self.path_excel = None
        self.path_files = None
        self.path_destination = None

    def browse_excel(self):
        path_excel = filedialog.askdirectory()
        if path_excel:
            self.path_excel = path_excel
            self.entry_excel.delete(0, tk.END)
            self.entry_excel.insert(0, self.path_excel)

    def browse_files(self):
        path_files = filedialog.askdirectory()
        if path_files:
            self.path_files = path_files
            self.entry_files.delete(0, tk.END)
            self.entry_files.insert(0, self.path_files)

    def browse_destination(self):
        path_destination = filedialog.askdirectory()
        if path_destination:
            self.path_destination = path_destination
            self.entry_destination.delete(0, tk.END)
            self.entry_destination.insert(0, self.path_destination)

    def clear(self):
        self.path_excel = None
        self.entry_excel.delete(0, tk.END)
        self.path_files = None
        self.entry_files.delete(0, tk.END)
        self.path_destination = None
        self.entry_destination.delete(0, tk.END)

    def save_directory(self):
        if self.path_excel:
            print("Directory saved:", self.path_excel)
        else:
            print("No directory selected!")
