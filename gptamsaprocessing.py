import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import glob
import re
import configparser
import math
import xlsxwriter
import shutil

logo_image = "iVBORw0KGgoAAAANSUhwMIrEtXUadYaQAAAABJRU5ErkJggg=="

data_columns = {}
data_columns1 = {}
data_columns2 = {}

product_var = None
file_label = None

def create_ui():
    window = tk.Tk()
    window.geometry("1080x720")
    
    create_logo_label(window)
    create_description_labels(window)
    create_product_option(window)
    create_processAMSA_section(window)
    
    window.mainloop()

def create_logo_label(window):
    logo_photo = tk.PhotoImage(data=logo_image)
    logo_label = tk.Label(window, image=logo_photo)
    logo_label.image = logo_photo
    logo_label.pack(side="top", pady=5)

def create_description_labels(window):
    description_label = tk.Label(window, text="Infineon IFAP DC - AMSA File Processing", font=("Calibri", 16, "bold"))
    description_label.pack(pady=20)

    instructions_label = tk.Label(window, text="This program inputs .stdf-derived .XLSX files, DMC Code of Golden Devices and the CalibrationSetupFile\nto create a new ACC Folder with the Calibrated Data which can be used for calibration purposes.\n\n CURRENTLY: The generated files and directory will be generated in the same directory from where this program is located", font=("Calibri", 13))
    instructions_label.pack(pady=5)

def create_product_option(window):
    global product_var
    
    product_var = tk.StringVar(window)
    product_var.set('Select Product')

    frame = create_frame(window, 0.2, 0.15, 0.85)

    product_combobox = ttk.Combobox(frame, textvariable=product_var, values=['Akari', 'Fuji'], font=("Calibri", 12))
    product_combobox.set('Select Product')
    product_combobox.pack(side="bottom", padx=10, pady=5, fill="x")
    product_combobox.bind("<<ComboboxSelected>>", update_data_columns)

def update_data_columns(event=None):
    product = product_var.get()
    if product in ['Akari', 'Fuji']:
        update_data_columns_dict(product)
        
def update_data_columns_dict(product):
    global data_columns, data_columns1, data_columns2

    # Define and update data_columns dictionaries based on the selected product
    # ... Your dictionary definitions here ...
    
    if product == 'Akari':
        data_columns = {
            '3001': '20Hz',
            '3002': '35Hz',
            '3003': '80Hz',
            '3004': '300Hz',
            '3005': '900Hz',
            '3006': '1000Hz',
            '3007': '1100Hz',
            '3008': '3000Hz',
            '3009': '8000Hz',
            '3010': '10000Hz',
        }

        data_columns1 = {
            '4001': '75Hz',
            '4002': '1000Hz',
            '4003': '3000Hz',
            '4004': '10000Hz',
        }
        
        data_columns2 = {
            '5001': '94dBSPL',
            '5002': '100dBSPL',
            '5003': '106dBSPL',
            '5004': '112dBSPL',
            '5005': '118dBSPL',
            '5006': '124dBSPL',
            '5007': '127dBSPL',
            '5008': '130dBSPL',
        }
        
    elif product == 'Fuji':
        data_columns = {
            '3001': '20Hz',
            '3002': '35Hz',
            '3003': '80Hz',
            '3004': '300Hz',
            '3005': '900Hz',
            '3006': '1000Hz',
            '3007': '1100Hz',
            '3008': '3000Hz',
            '3009': '8000Hz',
            '3010': '10000Hz',
        }

        data_columns1 = {
            '4001': '75Hz',
            '4002': '1000Hz',
            '4003': '3000Hz',
            '4004': '10000Hz',
        }
        
        data_columns2 = {
            '5001': '94dBSPL',
            '5002': '100dBSPL',
            '5003': '106dBSPL',
            '5004': '112dBSPL',
            '5005': '118dBSPL',
            '5006': '124dBSPL',
            '5007': '127dBSPL',
            '5008': '130dBSPL',
        }

def create_processAMSA_section(window):
    global file_label
    
    frame = create_frame(window, 0.9, 0.15, 0.35)

    button_font = ("Calibri", 14)
    open_button = tk.Button(frame, text='Upload Raw Data (GD AMSA) file (.xlsx)', command=open_file, font=button_font)
    open_button.place(relx=0.05, rely=0.2)

    file_label = tk.Label(frame, text="")
    file_label.place(relx=0.05, rely=0.6)

    process_button = tk.Button(frame, text='Process Raw Data and GD Selection', command=split_file, font=button_font)
    process_button.place(relx=0.67, rely=0.5)

def open_file():
    # Function to open and process files
    pass

# Create similar functions for other sections like loadDMCCode and loadCalibSetupFile

def create_frame(window, width, height, rely):
    screen_width = window.winfo_screenwidth()
    frame_x = (screen_width - width * screen_width) / 2

    frame = tk.Frame(window, bd=0)
    frame.place(relx=frame_x / screen_width, rely=rely, relwidth=width, relheight=height)
    return frame

if __name__ == "__main__":
    create_ui()