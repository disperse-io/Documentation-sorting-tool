# import os
# import shutil
# import pandas as pd
# import openpyxl

# class FolderCreator:
#     def __init__(self, parent_dir, project_name):
#         self.parent_dir = parent_dir.replace("\\","/")
#         self.project_name = project_name
#         self.path = os.path.join(parent_dir, project_name)
#         self.new_dir = os.path.join(parent_dir, project_name)
#         self.dirs = ['00_Overview_Documentation', '01_GA Plans', '02_Typrical Floor', '03_Typical Area', '04_Details',
#                      '05_Elevations & Sections', '06_Facades', '07_Struktural', '08_Schedule', '09_Schenatic',
#                      '10_Notes, Symbols', '11_Other', '12_All Data']
    
#     def create_folder(self):
#         os.makedirs(self.path, exist_ok=True)
#         print(f'Folder path is : {self.parent_dir}')
#         print(f'{self.project_name} folder is created')
    

    
#     def create_standardized_folders(self):
#         for dir in self.dirs:
#             path1 = os.path.join(self.new_dir, dir)
#             os.makedirs(path1, exist_ok=True)
        

# class DocumentMover:
#     def __init__(self, src_dir, dst_dir):
#         self.src_dir = src_dir.replace("\\", "/")
#         self.dst_dir = dst_dir

#     def move_docs(self):
#         for root, dirs, files in os.walk(self.src_dir):
#             for file in files:
#                 src_file = os.path.join(root, file)
#                 dst_file = os.path.join(self.dst_dir, file)
#                 shutil.move(src_file, dst_file)


# class DocumentSorter:
#     def __init__(self, path_to_all_data, other_dir):
#         self.path_to_all_data = path_to_all_data
#         self.other_dir = other_dir
#         self.names_of_doc = os.listdir(path_to_all_data)

#     def sort_by_type(self):
#         for file in self.names_of_doc:
#             if not file.endswith('.pdf'):
#                 shutil.move(os.path.join(self.path_to_all_data, file), self.other_dir)


# class InformationFiller:
#     def __init__(self, spec_dir, excel_file_path, file_paths, files):
#         self.spec_dir = spec_dir
#         self.excel_file_path = excel_file_path
#         self.file_paths = file_paths
#         self.files = files

#     def fill_excel_sheet(self):
#         df = pd.DataFrame({'File Name': self.files, 'File Path': self.file_paths})
#         df['Renamed Files '] = ""
#         df['Move to '] = ""

# # Create folder function

# folder_creator=FolderCreator(parent_dir=input('Enter a path where you want to create folder :'),project_name=input('Enter your project_name here : '))
# folder_creator.create_folder()

# # Crate standardizations folders

# folder = FolderCreator()
# folder.create_standardized_folders()