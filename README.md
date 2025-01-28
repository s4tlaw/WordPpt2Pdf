# WordPpt2Pdf
A python script for converting Word and PPT files to PDF

Created with Copilot and make a minor adjustment in line 38-40.

What I prompt to copilot is 
1. i have a list of ppt and word document in my C:\Files folder. write me a python code to convert them to pdf with win32com.client
2. Base on the program you provided, add one more function to open a dialog box for selecting the folder path

**There is many ways to convert documents to pdf. 
And Copilot will provide different codes and method for this project.
I found win32com.client seems to be the more reliable method.**

However for some reason the selected folder cannot be read by Presentation.Open.
The original version is:
a = select_folder() # a= C:/xxxxxxx/xxxx/

Thus I made a replace for the string path
input_folder_path = input_folder_path_old.replace("/", "\\")

C:\xxxxx\xxxx can now be read by Presentation.Open.

-------------------------------------

# Can further package with pyinstaller to run on any computer without installing any environment.

In CMD:
Install PyInstaller
pip install pyinstaller

Navigate to your script's directory
cd path/to/your_script_directory

Create the executable
pyinstaller --onefile your_script.py

----------------------------------------
# To Use the application:

1. Unzip
2. Run the application
3. Select the folder that contain your documents

