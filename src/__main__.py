import os
import shutil
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import ipywidgets as widgets
from IPython.display import display
from src.generator import Generator
from IPython.display import clear_output

# inputs and variables

golbalparent_dir = input("Enter a path where you want to create folder :")
project_name = input("Enter your project_name here : ")
parent_dir = golbalparent_dir.replace("\\", "/")
src_dir = input("Enter a path to folder with documentation :")
src_dir = src_dir.replace("\\", "/")
path = os.path.join(parent_dir, project_name)
dst_dir = path + "/12_All Data"
dirs = [
    "00_Overview_Documentation",
    "01_GA Plans",
    "02_Typical Floor",
    "03_Typical Area",
    "04_Details",
    "05_Elevations & Sections",
    "06_Facades",
    "07_Struktural",
    "08_Schedule",
    "09_Schenatic",
    "10_Notes & Symbols",
    "11_Other",
    "12_All Data",
]  # adding standardized folder names
path_to_all_data = path + "/12_All Data/"
other_dir = path + "/11_Other/"
programe_dir = path + "/00_Programme/"
spec_dir = path + "/00_Overview_Documentation/"
excel_file_path = spec_dir + project_name + ".xlsx"

Generator_object = Generator(parent_dir, src_dir, dirs, project_name)

# os.makedirs(path, exist_ok=True)

# Creating Buttons


button_create_project = widgets.Button(description="Create Project", button_style="info")
button_fill_excel = widgets.Button(description="Fill Excel", button_style="info")
button_move_files = widgets.Button(description="Move Files", button_style="info")


display(widgets.HBox([button_create_project, button_fill_excel, button_move_files]))

log_output=widgets.Output()
display(log_output)

def call_move_function(_):
    with log_output:
        clear_output()              
        Generator_object.create_and_move_files()


def call_fill_excel_sheet(_):
    with log_output:
        clear_output()
        Generator_object.fill_excel_sheet()
    


def call_move_to_final_location(_):
    with log_output:
        clear_output()
        Generator_object.move_to_final_location()


button_create_project.on_click(call_move_function)
button_fill_excel.on_click(call_fill_excel_sheet)
button_move_files.on_click(call_move_to_final_location)

