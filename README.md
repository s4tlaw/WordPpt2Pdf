# WordPpt2Pdf
A python script for converting Word and PPT files to PDF

Created with Copilot and make a minor adjustment in line 38-40.

What I prompt to copilot is 
1. I have a list of ppt and word document in my C:\Files folder. write me a python code to convert them to pdf
2. Base on the program you provided, add one more function to open a dialog box for selecting the folder path

However for some reason the selected folder cannot be read by Presentation.Open.
The original version is:
a = select_folder() # a= C:/xxxxxxx/xxxx/

Thus I made a replace for the string path
input_folder_path = input_folder_path_old.replace("/", "\\")

C:\xxxxx\xxxx can now be read by Presentation.Open.
