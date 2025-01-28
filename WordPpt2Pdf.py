import os
import win32com.client
import tkinter as tk
from tkinter import filedialog

# Function to open a dialog box to select a folder
def select_folder():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    folder_selected = filedialog.askdirectory()
    return folder_selected

# Function to convert PowerPoint to PDF
def convert_pptx_to_pdf(input_file_path, output_file_path):
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1
        presentation = powerpoint.Presentations.Open(input_file_path, WithWindow=False)
        presentation.SaveAs(output_file_path, 32)  # 32 is the format for PDF
        presentation.Close()
        powerpoint.Quit()
    except Exception as e:
        print(f"Error converting {input_file_path}: {e}")

# Function to convert Word to PDF
def convert_docx_to_pdf(input_file_path, output_file_path):
    try:
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = 0  # Set to 0 to hide the window
        document = word.Documents.Open(input_file_path, Visible=False)
        document.SaveAs(output_file_path, FileFormat=17)  # 17 is the format for PDF
        document.Close()
        word.Quit()
    except Exception as e:
        print(f"Error converting {input_file_path}: {e}")

# Select the input folder
input_folder_path_old = select_folder()

input_folder_path = input_folder_path_old.replace("/", "\\")

if input_folder_path:
    output_folder_path = os.path.join(input_folder_path, "Converted")

    # Create output directory if it doesn't exist
    if not os.path.exists(output_folder_path):
        os.makedirs(output_folder_path)

    # Iterate over files in the input folder
    for file_name in os.listdir(input_folder_path):
        input_file_path = os.path.join(input_folder_path, file_name)
        file_base_name, file_extension = os.path.splitext(file_name)
        output_file_path = os.path.join(output_folder_path, file_base_name + ".pdf")

        print(f"Converting {input_file_path} to {output_file_path}...")

        if file_extension.lower() == ".pptx" or ".ppt":
            convert_pptx_to_pdf(input_file_path, output_file_path)
        elif file_extension.lower() == ".docx":
            convert_docx_to_pdf(input_file_path, output_file_path)
        elif not file_extension.lower() == ".docx" or ".pptx" or ".ppt":
            continue
    
    print("Conversion complete. Check the 'Converted' folder for PDF files.")
else:
    print("No folder selected.")

