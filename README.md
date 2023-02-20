This program is designed to interact with you and help you manipulate documents.

To run this code, you will need Jupiter Notebook. Some packages must be installed before running the notebook. 
Follow the steps below to install those packages:


- Search for Anaconda Prompt in Windows, right-click on it, and select "Run as Administrator."
- It will ask for permission to open it. Press "Yes."
- Now run the following commands (without the braces) one-by-one:
    -conda update --all (enter "y" when it asks for permission) (This will take some time to install)
- pip install pandas, shutil, openpyxl, ipywidgets, pdf2image.


Usage:

1. Run the file named "DocSortingTool."
2. Fill in the input fields:
    - The first input will ask you to enter a path where you want to create a folder.
	- The second input will ask you to enter the name of the project that you want to create.
	- The third input will ask you to enter the path of the folder with documentation.
3. After entering the input fields, buttons should appear.
4. Click "Create Project." (This will create a main folder named the same as your entered project name, and inside it will automatically create standardized folders, move all documents from the folder that you chose in the third step, and sort documents by type. All PDFs from the chosen folder will be in "12_All Data," and other types of documents will be moved to "11_Other.")
5. After creating the folder for sorting documentation and moving all documents, you should use Adobe Acrobat to rename files.
6. When the documents are renamed, click "Fill Excel Button," and that function will create an Excel file inside "00_Overview_Documentation." Inside that Excel file,
you will have two columns:
	- File Name (list of all files stored in "12_All_Data")
	- Destination Folder (here you will have a dropdown where you can choose the folder where you want to move files)
	- You have comun new name, where you can add new document name if you need to rename some documents.
	- You should fill the "Destination Folder" column and hit "Ctrl + s" or save the file in another way.
7. After filling the Excel sheet, you should click "Move Files," and all files will be moved to the chosen folders, files with input in column
new name will be renamed and moved when you click to Move Files.

After you finish with steps above pdf documents will be stored in choosen directory, each directory will contain folder with JPEG's automatically converted from pdf's inside folder, and JPEG files you can use to upload files to figma.
In case that you have some problems or need further explanation contact Farrukh or Sanin from DAPS team. 

