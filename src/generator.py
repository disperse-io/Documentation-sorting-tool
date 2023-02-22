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
                for floor_no in range(00, 51):
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
            str(
                os.path.join(
                    self.destination_dir_path,
                    self.project_name,
                    self.folder_names[1],
                    self.project_name,
                )
            ),
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
                    last_folder = str(row[3]).split(" <> ")
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
                            *last_folder,
                            new_file_name,
                        )
                    else:
                        file_dest_path = os.path.join(
                            self.destination_dir_path,
                            self.project_name,
                            *last_folder,
                            row[1],
                        )
                try:
                    shutil.move(file_src_path, file_dest_path)

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


    def pdf_to_JPEG(self):

        path_to_typical_floor=os.path.join(self.destination_dir_path,self.project_name,"02_Typical Floor")       
        files_and_folders = os.listdir(os.path.join(self.destination_dir_path,self.project_name,"02_Typical Floor"))

                
        # display(files_and_folders)

        for item in files_and_folders:
            folder_path = os.path.join(path_to_typical_floor,item)
            folder_path = folder_path.replace("\\", "/")
            pdf_file_list=os.listdir(folder_path)                   

            # display("First Step on road !!")
            display(folder_path)
            # display(pdf_file_list)

            JPEG_Folder_path=os.path.join(folder_path,"JPEGs")
            JPEG_Folder_path=JPEG_Folder_path.replace("\\", "/")
            display(JPEG_Folder_path)
            if os.path.isdir(JPEG_Folder_path):
                display(f'Folder at {item} already exist')
                pass
            else:
                os.mkdir(JPEG_Folder_path)
                display(f'Folder created in {item} ')
 
            for file in pdf_file_list:
                if file[-3:] == "pdf":
                    filePath = os.path.join(folder_path,file)
                    display(f'converting {file}')
                    pdfImages = convert_from_path(filePath)                                      
                    counter= 0
                    for img in pdfImages:
                        setupPath = os.path.join(JPEG_Folder_path, file[:-4])
                        if counter==0:
                            img.save(setupPath + ".jpg", "JPEG")
                            counter +=1
                        else:
                            img.save(setupPath +"_" +str(counter) + ".jpg", "JPEG")
                            counter += 1
                    display(f'Converted to {file}.jpg') 
                display('All pdf are converted successfully')                  
                    
           
                
            # for pdf_file in pdf_file_list:                    
            #     pdf_path=os.path.join(item,pdf_file)
            #     try:
            #         with open(pdf_path,'rb') as f:
            #             pages = convert_from_path(pdf_path)
            #         for i, page in enumerate(pages):
            #             jpeg_path=os.path.join(item,f"{pdf_file}_{i}.jpg")
            #             page.save(jpeg_path,'JPEG')

            #     except IOError as e:
            #         print(f"I/O Error: {e}")
                    

                 
                    



                   

                    # for i, page in enumerate(pages):
                    #     jpeg_path = os.path.join(JPEG_Folder_path,f"{pdf_file}_{i}.jpg")
                    #     page.save(jpeg_path,"JPEG")



        #         print(folder_path)
        #         print(pdf_file_list)

        



        
        # for item in folder_path:
        #     print(item)
            # file_list=os.listdir(folder)
            # print(file_list)





            # pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
            # display('First Loop')
            # for pdf_file in pdf_files:
            #     pdf_path = os.path.join(folder, pdf_file)
            #     pages = convert_from_path(pdf_path)
            #     display("Second Loop")
            #     for i, page in enumerate(pages):
            #         jpeg_path = os.path.join(folder, f"{pdf_file}_{i}.jpg")
            #         page.save(jpeg_path, 'JPEG')
            #         display("3rd Loop")
        # with open(pdf_path, 'rb') as f:
        #     pdf = PyPDF2.PdfFileReader(f)
        #     for page_num in range(pdf.getNumPages()):
        #         page = pdf.getPage(page_num)
        #         _, temp_filename = tempfile.mkstemp()
        #         with open(temp_filename, 'wb') as f:
        #             f.write(page.contentStream.getData())
        #         img = Image.open(temp_filename)
        #         output_filename = os.path.splitext(os.path.basename(pdf_path))[0] + f'_page_{page_num+1}.jpg'
        #         img.save(os.path.join(output_folder, output_filename), 'JPEG')
        #         os.remove(temp_filename)

        # project_folders_path = os.path.join(
        #     self.destination_dir_path, self.project_name
        # )
        # folder_path = project_folders_path + "02_Typical Floor"
        # Display("This is folder path : ", folder_path)
        # output_folder