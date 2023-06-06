import tkinter as tk
import os
import csv
import datetime
import time
import pandas as pd
from moviepy.editor import VideoFileClip
from tkinter import filedialog
from tkinter import ttk

class CreateExcel:
    def __init__(self, parent):
        self.parent = parent
        self.tab = ttk.Frame(parent)

        # Add your GUI elements here
        self.label = ttk.Label(self.tab, text="Diretório dos Arquivos:")
        self.label.grid(column = 1, row = 1)

        self.entry = ttk.Entry(self.tab, width=40)
        self.entry.grid(column = 1, row = 2)
        
        self.browse_button = ttk.Button(self.tab, text="Browse", command=self.browse_directory)
        self.browse_button.grid(column = 2, row = 2)

        self.label2 = ttk.Label(self.tab, text="Diretório do Excel:")
        self.label2.grid(column = 1, row = 3)

        self.entry2 = ttk.Entry(self.tab, width=40)
        self.entry2.grid(column = 1, row = 4)

        self.browse_button2 = ttk.Button(self.tab, text="Browse", command=self.browse_directory2)
        self.browse_button2.grid(column = 2, row = 4)

        self.save_button = ttk.Button(self.tab, text="Gerar Excel", command=self.save_directory)
        self.save_button.grid(column = 1, row = 5)

        self.clear_button = ttk.Button(self.tab, text="Limpar", command=self.clear)
        self.clear_button.grid(column = 2, row = 5)

        self.directory = None
        self.directory2 = None

    def browse_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.directory = directory
            self.entry.delete(0, tk.END)
            self.entry.insert(0, self.directory)

    def browse_directory2(self):
        directory2 = filedialog.askdirectory()
        if directory2:
            self.directory2 = directory2
            self.entry2.delete(0, tk.END)
            self.entry2.insert(0, self.directory2)

    def clear(self):
        self.directory = None
        self.entry.delete(0, tk.END)
        self.directory2 = None
        self.entry2.delete(0, tk.END)

    def get_video_duration(self, file_path):
        try:
            video = VideoFileClip(file_path)
            duration = video.duration
            video.close()
            return duration
        except Exception as e:
            print(f"Error occurred while getting duration for file '{file_path}': {str(e)}")
            return 0

    def calculate_video_bitrate(self, file_size, duration):
        # Convert file size to bits
        file_size_bits = file_size*8

        # Calculate video bit rate
        video_bitrate = file_size_bits / duration

        return video_bitrate

    def save_directory(self):
        directory = self.directory
        directory2 = self.directory2
        if directory and directory2:
            # Get all the files in the directory
            files = os.listdir(directory)

            # Create a CSV file
            csv_file = open(directory2 + '/' + 'file_data.csv', 'w', newline='', encoding='utf-8')
            print(directory2)
            csv_writer = csv.writer(csv_file)

            # Write header row
            csv_writer.writerow(["Name", "Extension", "Duration", "Size", "CreationDate", "Bitrate"])

            # Loop through each file and get the details
            for file in files:
                file_path = os.path.join(directory, file)
                if os.path.isfile(file_path):
                    # Get file details
                    file_name = os.path.basename(file_path).split('.')[0]
                    extension = os.path.splitext(file_path)[1]
                    duration = self.get_video_duration(file_path)
                    file_size = os.path.getsize(file_path)
                    created_date = datetime.datetime.fromtimestamp(os.path.getctime(file_path)).isoformat(sep=" ", timespec="seconds")
                    bitrate = self.calculate_video_bitrate(file_size, duration)

                    # Formata os dados
                    duration = time.strftime("%Hh %Mmin %Ss", time.gmtime(duration))
                    if file_size < 1073741824:
                        file_size = str(round(file_size/(1024*1024), 2)) + " MB"
                    else:
                        file_size = str(round(file_size/(1024*1024*1024), 2)) + " GB"
                    bitrate = str(round(bitrate/1000)) + " kb/s"

                    # Write data to the CSV file
                    csv_writer.writerow([file_name, extension, duration, file_size, created_date, bitrate])

            csv_file.close()
            print("CSV file created successfully!")

            # Read the CSV file
            csv_file = directory2 + '/' + 'file_data.csv'
            df = pd.read_csv(csv_file)

            # Folder name
            folder_name = os.path.basename(directory)

            # Define the output Excel file
            created_date = str(created_date)
            created_date = created_date.replace(" ", "_")
            created_date = created_date.replace(":", "h", 1)
            created_date = created_date.split(":", 1)[0]
            created_date = created_date + 'min'
            excel_file = folder_name + '_' + created_date + '.xlsx'
            print(directory)

            # Write the DataFrame to an Excel file
            df.to_excel(directory2 + '/' + excel_file, index=False)
            print("Excel file created successfully!")

            # Delete csv
            os.remove(csv_file)
        else:
            print("Please select a directory.")