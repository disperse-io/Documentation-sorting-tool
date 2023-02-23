import os
import shutil
import pandas as pd
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import quote_sheetname
import subprocess
from pdf2image import convert_from_path


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

        # Creating Folders
        for directory_name in self.folder_names:
            if directory_name == "02_Typical Floor":
                for floor_no in range(-4, 51):
                    os.makedirs(
                        os.path.join(
                            project_folders_path,
                            directory_name,
                            "Floor_" + str(floor_no),
                        ),
                        exist_ok=True,
                    )

            else:
                os.makedirs(
                    os.path.join(project_folders_path, directory_name),
                    exist_ok=True,
                )

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

    def fill_excel_sheet(self):
        all_data_files = os.listdir(
            os.path.join(
                self.destination_dir_path, self.project_name, self.folder_names[-1]
            )
        )
        typical_floor_direcs = os.listdir(
            os.path.join(
                self.destination_dir_path, self.project_name, "02_Typical Floor"
            )
        )
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        validation_worksheet = workbook.create_sheet("validation_data")
        worksheet.append(["File Name", "New Name", "Destination"])
        for row in all_data_files:
            worksheet.append([row])

        for folders in self.folder_names:
            if folders == "02_Typical Floor":
                for floor in typical_floor_direcs:
                    validation_worksheet.append([folders + " <> " + floor])
            validation_worksheet.append([folders])
        validation = DataValidation(
            type="list",
            formula1="{0}!$A:$A".format(quote_sheetname("validation_data")),
            allow_blank=False,
        )
        validation.add("C2:C" + str(len(self.folder_names)))
        worksheet.add_data_validation(validation)
        validation_worksheet.sheet_state = "hidden"
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
        subprocess.call(
            os.path.join(
                self.destination_dir_path,
                self.project_name,
                self.folder_names[0],
                self.project_name,
            )
            + ".xlsx",
            shell=True,
        )

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
                if len(str(row[3])) > 1:
                    last_folder = os.path.join(
                        self.destination_dir_path,
                        self.project_name,
                        *str(row[3]).split(" <> "),
                    )
                    file_src_path = os.path.join(
                        self.destination_dir_path,
                        self.project_name,
                        self.folder_names[-1],
                        row[1],
                    )

                    if isinstance(new_file_name, str):
                        file_name = new_file_name
                        file_dest_path = os.path.join(last_folder, new_file_name)
                    else:
                        file_name = row[1]
                        file_dest_path = os.path.join(last_folder, row[1])
                try:
                    shutil.move(file_src_path, file_dest_path)
                    self.convert_pdf_to_jpeg(file_name, last_folder)

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

    def convert_pdf_to_jpeg(self, file_to_convert, folder_path):
        if not os.path.exists(os.path.join(folder_path, "JPEGS")):
            os.mkdir(os.path.join(folder_path, "JPEGS"))
        jpeg_counter = 0
        pdfImages = convert_from_path(os.path.join(folder_path, file_to_convert))
        setupPath = os.path.join(folder_path, "JPEGS", file_to_convert[:-4])
        for img in pdfImages:
            if jpeg_counter == 0:
                img.save(setupPath + ".jpg", "JPEG")
                jpeg_counter += 1
            else:
                img.save(setupPath + "_" + str(jpeg_counter) + ".jpg", "JPEG")
                jpeg_counter += 1
