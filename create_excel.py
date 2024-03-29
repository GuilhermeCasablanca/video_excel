import tkinter as tk
import os
import datetime
import time
import cv2
import pandas as pd
from tkinter import filedialog, ttk, messagebox

class CreateExcel:
    def __init__(self, parent):
        self.parent = parent
        self.tab = ttk.Frame(parent)

        # Add your GUI elements here
        # Diretório dos Arquivos
        self.label = ttk.Label(self.tab, text="Diretório dos Arquivos:")
        self.entry = ttk.Entry(self.tab, width=40)
        self.browse_button = ttk.Button(self.tab, text="Browse", command=self.browse_directory)

        # Diretório do Excel
        self.label2 = ttk.Label(self.tab, text="Diretório do Excel:")
        self.entry2 = ttk.Entry(self.tab, width=40)
        self.browse_button2 = ttk.Button(self.tab, text="Browse", command=self.browse_directory2)

        # Gerar Excel
        self.save_button = ttk.Button(self.tab, text="Gerar Excel", command=self.save_directory)
        
        # Limpar
        self.clear_button = ttk.Button(self.tab, text="Limpar", command=self.clear)
        
        # Formatação
        self.label.grid(column = 1, row = 1, sticky=tk.W)
        self.entry.grid(column = 1, row = 2, padx=5, pady=5)
        self.browse_button.grid(column = 2, row = 2, padx=5, pady=5)

        self.label2.grid(column = 1, row = 3, sticky=tk.W)
        self.entry2.grid(column = 1, row = 4, padx=5, pady=5)
        self.browse_button2.grid(column = 2, row = 4, padx=5, pady=5)

        self.save_button.grid(column = 1, row = 5, padx=5, pady=5)
        
        self.clear_button.grid(column = 2, row = 5, padx=5, pady=5)

        # Variaveis
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
            video_capture = cv2.VideoCapture(file_path)
            frame_count = int(video_capture.get(cv2.CAP_PROP_FRAME_COUNT))
            fps = int(video_capture.get(cv2.CAP_PROP_FPS))
            duration = frame_count / fps
            video_capture.release()
            return duration
        except Exception as e:
            print(f"Error: {e}")
            return None

    def calculate_video_bitrate(self, file_size, duration):
        # Convert file size to bits
        file_size_bits = file_size*8

        # Calculate video bit rate
        video_bitrate = file_size_bits / duration

        return video_bitrate
    
    def is_video(self, file_path):
        try:
            # Open the file
            cap = cv2.VideoCapture(file_path)

            # Check if the file is opened successfully
            if not cap.isOpened():
                return False

            # Get the frames per second (fps) of the video
            fps = cap.get(cv2.CAP_PROP_FPS)

            # Check if the fps is valid
            if fps > 0:
                 return True
            else:
                  return False
        except Exception as e:
            print(f"Error: {str(e)}")
            return False
        finally:
            # Release the video capture object
            if cap is not None:
                cap.release()

    def save_directory(self):
        directory = self.directory
        directory2 = self.directory2
        if directory and directory2:
            # Get all the files in the directory
            files = os.listdir(directory)

            # Define a matriz dos dados
            excel_data = []

            # Loop through each file and get the details
            for file in files:
                file_path = os.path.join(directory, file)
                if (os.path.isfile(file_path)) and (self.is_video(file_path)):
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
                    bitrate = f"{bitrate/(1000*1000):.2f}" + " Mbps"
                    
                    # Write data to the CSV file
                    excel_data.append([file_name, extension, duration, file_size, created_date, bitrate])

            # Folder name
            folder_name = os.path.basename(directory)

            # Define the output Excel file

            created_date = str(created_date)
            created_date = created_date.replace(" ", "_")
            created_date = created_date.replace(":", "h", 1)
            created_date = created_date.split(":", 1)[0]
            created_date = created_date + 'min'
            excel_file = directory2 + "/" + folder_name + '_' + created_date + '.xlsx'

            # Create a DataFrame from the array list
            df = pd.DataFrame(excel_data, columns=["Name", "Extension", "Duration", "Size", "CreationDate", "Bitrate"])

            # Write the DataFrame to an Excel file
            df.to_excel(excel_file, index=False)

            messagebox.showinfo("Confirmação", "Arquivo Excel criado com sucesso")

        else:
            messagebox.showinfo("Erro", "Por favor selecione um diretório")