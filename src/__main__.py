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
    "02_Typrical Floor",
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
button3 = widgets.Button(description="Rename Files", button_style="info")

display(widgets.HBox([button_create_project, button_fill_excel, button_move_files, button3]))

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
# part for checking / should be deleted after testing
# print(f'Path where you will create new folder is : {parent_dir}')
# print(f'{project_name} folder is created')
# print(f'{src_dir} is folder from where you want to move documentation')
# print (f'{path} : paath')
# print(f'{path} : excel file path' )


# Creating standardized folders inside folder project_name
# def create_standardized_folder(b):
#     for dir in dirs:
#         path1 = os.path.join(path, dir)
#         os.makedirs(path1, exist_ok=True)

#     # Moving documents in created folder
#     # def move_documents(b):
#     for root, dirs, files in os.walk(src_dir):
#         for file in files:
#             src_file = os.path.join(root, file)
#             dst_file = os.path.join(dst_dir, file)
#             shutil.move(src_file, dst_file)

#     # Sorting by type
#     # def sort_by_type(b):
#     names_of_doc = os.listdir(path + "/12_All Data")

#     for file in names_of_doc:
#         if not file.endswith(".pdf"):
#             shutil.move(path_to_all_data + file, other_dir)


# fill excel sheet with information
# def fill_excel_sheet():

#     files = [
#         f
#         for f in os.listdir(path_to_all_data)
#         if os.path.isfile(os.path.join(path_to_all_data, f))
#     ]
#     file_paths = [
#         os.path.dirname(os.path.join(path_to_all_data, f)) for f in files
#     ]  # filepath treba doraditi
#     df = pd.DataFrame({"File Name": files, "File Path": ""})
#     df["Renamed Files "] = ""
#     df["Move to "] = ""

#     df.to_excel(excel_file_path, index=False)


# Adding data validation to column Move to : - to fix


# def move_files():

#     df1 = pd.read_excel(spec_dir + project_name + ".xlsx")

#     # Iterate over rows in dataframe
#     for index, row in df1.iterrows():
#         file_name = row["File Name"]
#         source_path = os.path.join(dst_dir, file_name)
#         destination_folder = row["Move to "]
#         destination_path = os.path.join(destination_folder, file_name)

#         # Check if source file exists
#         if os.path.exists(source_path):
#             shutil.move(source_path, destination_path)
#         else:
#             print("File not found:", source_path)


# Calling Functions / Buttons


# print("Folders are created....")


# from openpyxl.worksheet.datavalidation import DataValidation
# import openpyxl
# wb = openpyxl.load_workbook("C:/Users/sanin/Desktop/Documentation sorting tool/test/withDataValid/00_Overview_Documentation/withDataValid.xlsx")
# sheet = wb['Sheet1']

# # Create a validation rule to limit the values in the "Move to" column to a list of paths
# dv = DataValidation(type="list", formula1='Dog,Cat,Bat', allow_blank=False,showDropDown=True)
# # Apply the validation rule to the "Move to" column
# for row in sheet['D']:
#     dv.add(row)

# # Save the changes to the Excel file
# wb.save('C:/Users/sanin/Desktop/Documentation sorting tool/test/withDataValid/00_Overview_Documentation/withDataValid.xlsx')


# # # # move files using excel - to fix
# from openpyxl import load_workbook

# book = load_workbook(spec_dir+project_name+'.xlsx')
# print(book[Sheet1])

# df = pd.read_excel('C:/Users/sanin/Desktop/Documentation sorting tool/test/withDataValid/00_Overview_Documentation/withDataValid.xlsx')

# path_to_all_data = os.path.dirname('C:/Users/sanin/Desktop/Documentation sorting tool/test/withDataValid/00_Overview_Documentation/withDataValid.xlsx')

# # Loop through each row in the dataframe
# for index, row in df.iterrows():
#     # Get the file name and destination path
#     file_name = row['File Name']
#     destination_path = row['Move to']

#     # Construct the full path of the file
#     file_path = os.path.join(path_to_all_data, file_name)

#     # Check if the file exists
#     if os.path.isfile(file_path):
#         # Move the file to the destination path
#         shutil.move(file_path, destination_path)


# excel_file = 'C:/Users/sanin/Desktop/Documentation sorting tool/MoveFiles.xlsx' ## should be adjusted for others, location of excel file
# src_dir = 'C:/Users/sanin/Desktop/Doc Sorting Programm/srcDic' ## should be adjusted for others, location of excel file


# #Write list of all files to Excel
# files = os.listdir(src_dir)
# df=pd.DataFrame({'File Name':files})


# for index, row in df.iterrows():
#     file_name = row['File Name']
#     dest_folder = row['Destination Folder']

# src_file = os.path.join(src_dir,file_name)

# if os.path.isfile(src_file):
#         dest_folder = os.path.join(src_dir, dest_folder)

#         if not os.path.isdir(dest_folder):
#             os.makedirs(dest_folder)

#         dest_file = os.path.join(dest_folder, file_name)
#         shutil.move(src_file, dest_file)
# else:
#         print(f'Error: {src_file} does not exist.')


# 01. Connect to server and download all data releted to project - potential PROBLEM because we have multiple server for dowloading documentation, different users and passwords,
# also in some cases we recive documentation from client by email.
# 02. Create a folder with project_name for a project that we need documentation for - DONE (Crated based on input)
# 03. Place data in one centralized place inside project name ( Is it better to have one standardized folder where we will store data after dowloading from servers/receiving from client,
# and then from this folder read documents and move them into right folders)
# 04. Analyze data and sort by type ( In this step I mean that after dowloading all data from server we should keep only pdfs and other documents we can send into folder Other
# and there sorty by file_type)


# 05. Create new Standardized folders - DONE - ( to discus with Vanessa and Ahmed about standardization of folders/ which folder names we need )
# 06. Rename files based on legend on drawings/ we can use title_name/name/title + number of drawing - (we can use AdobeReader to rename files and than use file names for sorting)
# 07. List out all the files inside your folder (This step to be discused is it better to create one separated folder for documentation in root folder a than to move files into
# right folder, or to create folder with documentation inside folder with project name)
#
# 08. Move those documents into right folders ( Option 1 : move automatically to folder based on filename ,
#                                               Option 2 : Sort data by Drawing num and ask user where to store all documents that similar drawing num.)
# 9. Delete empty folders

# 10. Rewiew Code by some senior member and handover to Docs_team.
