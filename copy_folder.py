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
        self.label = ttk.Label(self.tab, text="Arquivo Excel:")
        self.label.pack(padx=10, pady=10)

        self.entry = ttk.Entry(self.tab, width=40)
        self.entry.pack(padx=10, pady=5)

        self.browse_button = ttk.Button(self.tab, text="Browse", command=self.browse_directory)
        self.browse_button.pack(padx=10, pady=5)

        self.save_button = ttk.Button(self.tab, text="Copiar", command=self.save_directory)
        self.save_button.pack(padx=10, pady=5)

        self.directory = None

    def browse_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.directory = directory
            self.entry.delete(0, tk.END)
            self.entry.insert(0, self.directory)

    def save_directory(self):
        if self.directory:
            print("Directory saved:", self.directory)
        else:
            print("No directory selected!")
