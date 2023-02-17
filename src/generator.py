import os
import shutil
import pandas as pd
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
import math


class Generator:
    # def __init__(self):
    def __init__(
        self, destination_dir_path, source_dir_path, folder_names, project_name
    ):

        self.destination_dir_path = destination_dir_path
        self.source_dir_path = source_dir_path
        self.folder_names = folder_names
        self.project_name = project_name

    def create_and_move_files(self):

        project_folders_path = os.path.join(
            self.destination_dir_path, self.project_name
        )
        if os.path.exists(project_folders_path):
            display("Error: Folder already exists.")
            return
        [
            os.makedirs(
                os.path.join(project_folders_path, x),
                exist_ok=True,
            )
            for x in self.folder_names
        ]
        for root, dirs, files in os.walk(self.source_dir_path):
            for file in files:
                src_file = os.path.join(root, file)
                if not file.endswith(".pdf"):
                    dst_file = os.path.join(
                        project_folders_path, self.folder_names[-2], file
                    )
                    shutil.copy(src_file, dst_file)
                else:
                    dst_file = os.path.join(
                        project_folders_path, self.folder_names[-1], file
                    )
                    shutil.copy(src_file, dst_file)

        if (
            len(os.listdir(os.path.join(project_folders_path, self.folder_names[-2])))
            > 0
        ):
            display(
                str(
                    len(
                        os.listdir(
                            os.path.join(project_folders_path, self.folder_names[-2])
                        )
                    )
                )
                + " non-pdf files has been moved to 11_Others folder"
            )
        display("Folder created to following path: " + str(project_folders_path))

    # def fill_excel_sheet(self, destination_dir_path, folder_names, file_name):

    def fill_excel_sheet(self):
        all_data_files = os.listdir(
            os.path.join(
                self.destination_dir_path, self.project_name, self.folder_names[-1]
            )
        )
        validation_rule = ""
        # data_to_paste = [[x] for x in folder_names if folder_names != "12_All_Data"]
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.append(["File Name", "New Name", "Destination"])
        for row in all_data_files:
            worksheet.append([row])

        for folders in self.folder_names:
            if folders == self.folder_names[-1]:
                validation_rule = validation_rule + folders
            else:
                validation_rule = validation_rule + folders + ","
        validation = DataValidation(
            type="list", formula1='"' + validation_rule + '"', allow_blank=True
        )
        validation.add("C2:C" + str(len(self.folder_names)))
        worksheet.add_data_validation(validation)
        workbook.save(
            os.path.join(
                self.destination_dir_path,
                self.project_name,
                self.folder_names[0],
                self.project_name,
            )
            + ".xlsx"
        )
        display(
            "Successfully created excel file to following path: "
            + str(
                os.path.join(
                    self.destination_dir_path,
                    self.project_name,
                    self.folder_names[1],
                    self.project_name,
                )
            )
        )

    # def move_to_final_location(self, project_folder, file_name):
    def move_to_final_location(self):
        try:
            excel_file = pd.read_excel(
                os.path.join(
                    self.destination_dir_path,
                    self.project_name,
                    self.folder_names[0],
                    self.project_name,
                )
                + ".xlsx",
                header=0,
            )

            for row in excel_file.itertuples():
                new_file_name = row[2]
                if len(str(row[3])) > 1:  # Checking if renamed from file

                    file_src_path = os.path.join(
                        self.destination_dir_path,
                        self.project_name,
                        self.folder_names[-1],
                        row[1],
                    )
                    if isinstance(new_file_name, str):
                        file_dest_path = os.path.join(
                            self.destination_dir_path,
                            self.project_name,
                            str(row[3]),
                            new_file_name,
                        )
                    else:
                        file_dest_path = os.path.join(
                            self.destination_dir_path,
                            self.project_name,
                            str(row[3]),
                            row[1],
                        )
                try:
                    shutil.move(file_src_path, file_dest_path)
                    # if len(new_file_name) > 1:
                    #     os.remove(file_src_path)

                except FileNotFoundError:
                    pass
            file_count = len(
                os.listdir(
                    os.path.join(
                        self.destination_dir_path,
                        self.project_name,
                        self.folder_names[-1],
                    )
                )
            )
            if file_count == 0:
                display("Congratulations you moved all files!")
            else:
                display("You have " + str(file_count) + " more PDFs to move !")
        except PermissionError:
            if "excel_file" in locals() and not excel_file.closed:
                excel_file.close()
            raise
