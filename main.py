import tkinter as tk
from tkinter import filedialog
import os
import csv
import datetime
import time
from moviepy.editor import VideoFileClip
import pandas as pd

def browse_directory():
    directory = filedialog.askdirectory()
    directory_entry.delete(0, tk.END)
    directory_entry.insert(0, directory)

def get_video_duration(file_path):
    try:
        video = VideoFileClip(file_path)
        duration = video.duration
        video.close()
        return duration
    except Exception as e:
        print(f"Error occurred while getting duration for file '{file_path}': {str(e)}")
        return 0

def calculate_video_bitrate(file_size, duration):
    # Convert file size to bits
    file_size_bits = file_size*8

    # Calculate video bit rate
    video_bitrate = file_size_bits / duration

    return video_bitrate

def save_directory():
    directory = directory_entry.get()
    if directory:
        # Get all the files in the directory
        files = os.listdir(directory)

        # Create a CSV file
        csv_file = open('file_data.csv', 'w', newline='', encoding='utf-8')
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
                duration = get_video_duration(file_path)
                file_size = os.path.getsize(file_path)
                created_date = datetime.datetime.fromtimestamp(os.path.getctime(file_path)).isoformat(sep=" ", timespec="seconds")
                bitrate = calculate_video_bitrate(file_size, duration)

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
        csv_file = 'file_data.csv'
        df = pd.read_csv(csv_file)

        # Define the output Excel file
        excel_file = str(directory) + '.xlsx'

        # Write the DataFrame to an Excel file
        df.to_excel(excel_file, index=False)
        print("Excel file created successfully!")
    else:
        print("Please select a directory.")

root = tk.Tk()
root.title("Directory Input")

# Create a label
directory_label = tk.Label(root, text="Directory:")
directory_label.pack()

# Create an entry field
directory_entry = tk.Entry(root, width=50)
directory_entry.pack()

# Create a button to browse the directory
browse_button = tk.Button(root, text="Browse", command=browse_directory)
browse_button.pack()

# Create a button to save the directory
save_button = tk.Button(root, text="Executar", command=save_directory)
save_button.pack()

root.mainloop()
