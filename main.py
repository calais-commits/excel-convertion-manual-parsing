import openpyxl
import os
import xlwings as xw
import tkinter as tk
import threading
import time
import subprocess
from pathlib import Path
from tkinter import *
import customtkinter as ctk

# The function will evaluate the value of the 'Unicode' field of the file TechnicalProperties
def evaluate_file(file_path, file_name, main_window):
  if file_path:
    # Open the file using the selected path
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    unicode_value = sheet['B6'].value
        
    # Evaluate the 'Unicode' field value
    # If its value is 'yes', the script will continue without warnings
    if unicode_value.lower() in ('yes'):
      return 'The information in this file is Unicode'
    # if the value in the Unicode field is not 'yes', it will show a warning message 
    else:
      return 'The information in this file is not Unicode'
    
  else:
    return "Can't found files"

# This function will ask for all the other files and will add them as sheets
def merge_files_handler(script_file_path, folder_name, timestamp, label_get_files, label_copy_files, label_copy_files_result, label_merge_status):
  label_get_files.configure(text="Getting files...")
  # This function will get all the files in the folder and copy them
  def copy_files():
    with xw.App(visible=False) as app:
      # Get the excel files in the folder and open a Workbook
      excel_files_path = os.path.join(script_file_path, folder_name)
      excel_files = list(Path(excel_files_path).glob('*.xlsx'))
      combined_wb = xw.Book()
      # Procedures for UI messages
      def copy_files_status():
        label_copy_files.configure(text="Copying files...")

      def copy_files_result():
        label_copy_files_result.configure(text="Saving changes...")
      # If we have Excel files, then print messages
      if len(excel_files) > 0:
        thread2 = threading.Thread(target=copy_files_status)
        thread2.start()
        # For each Excel file, open it and copy each sheet to the combined workbook
        for excel_file in excel_files:
          wb = xw.Book(excel_file)
          for sheet in wb.sheets:
            # Get the file name without the extension .xlsx
            sheet_name = excel_file.stem
            # Copy the sheet and assign the file name as the sheet name
            sheet.api.Copy(After=combined_wb.sheets[0].api)
            combined_wb.sheets[1].name = sheet_name

        # Delete the first default sheet from the combined workbook
        combined_wb.sheets[0].delete()
        # Save the combined workbook with a timestamp
        result_file_name = f'merged{timestamp}.xlsx'
        # Save changes
        combined_wb.save(result_file_name)
        # Print result messages
        thread3 = threading.Thread(target=copy_files_result)
        thread3.start()
        label_merge_status.configure(text="Files merged successfully")
      
      else:
        label_copy_files.configure(text="Can't found files")
        
  # Start the copy_files function in a separate thread
  thread = threading.Thread(target=copy_files)
  thread.start()

# This function will close the main window and the confirmation window
def close_windows(main_window, confirmation_window):
  confirmation_window.destroy()
  main_window.destroy()

# This function will execute the merge_files_handler function and will close the confirmation window, after that, it will return a 'yes'
# If the result is 'yes', it will call the merge_files_handler function
# If the result is 'no', it will close the windows
def confirm_and_return_yes(script_file_path, confirmation_window, folder_name, timestamp, label_get_files, label_copy_files, label_copy_files_result, label_merge_status):
  # Close confirmation window
  confirmation_window.destroy()
  # Call the merge_files_handler function
  result_merge_files_handler = merge_files_handler(script_file_path, folder_name, timestamp, label_get_files, label_copy_files, label_copy_files_result, label_merge_status)
  return 'yes'

# This function will create a confirmation window
# If the result from the evaluate_file function is 'The information in this file is not Unicode', it will show a confirmation window
# If the result is 'yes', it will call the merge_files_handler function
# If the result is 'no', it will close the windows
# If the result from the evaluate_file function is 'The information in this file is Unicode', it will call the merge_files_handler function without showing a confirmation window
def open_confirmation_window(script_file_path, main_window, folder_name, timestamp, label_get_files, label_copy_files, label_copy_files_result, label_merge_status):
  # Create confirmation window
  confirmation_window = ctk.CTkToplevel()
  confirmation_window.focus_force()
  confirmation_window.title("Confirmation")
  confirmation_window.geometry("450x120")

  label = ctk.CTkLabel(confirmation_window, text="The information in the file is not Unicode. Are you sure you want to continue?")
  label.pack(pady=20)

  # If yes
  yes_button = ctk.CTkButton(confirmation_window, text="Yes", fg_color=("teal"), width=(100), command=lambda: confirm_and_return_yes(script_file_path, confirmation_window, folder_name, timestamp, label_get_files, label_copy_files, label_copy_files_result, label_merge_status))
  yes_button.pack(side=tk.LEFT, padx=50)

  # If no
  no_button = ctk.CTkButton(confirmation_window, text="No", fg_color=("teal"), width=(100), command=lambda: close_windows(main_window, confirmation_window))
  no_button.pack(side=tk.RIGHT, padx=50)

# This function will open the manual
def open_manual():
  # File name
  manual_file = 'manual.txt'
  # Operative System
  operative_system = os.name

  if operative_system == 'nt':  # Windows
    subprocess.Popen(f'start {manual_file}', shell=True)
  elif operative_system == 'posix':  # macOS, Linux
    subprocess.Popen(['open', manual_file], shell=False)

# This procedure is the main handler of the script
def main(main_window, label_get_files, label_copy_files, label_copy_files_result, label_merge_status):
  # Get current file path (The excel file should be in the same folder)
  script_file_path = os.path.dirname(os.path.abspath(__file__))
  folder_name = 'documents'
  file_name = "TechnicalProperties.xlsx"

  # Get the excel file path
  file_path = os.path.join(script_file_path, folder_name, file_name)

  # Create a timestamp
  t = time.localtime()
  timestamp = time.strftime('%Y-%m-%d_%H-%M-%S', t)

  # Call the first function to evaluate the TechnicalProperties file and show the result
  result_evaluate_file = evaluate_file(file_path, file_name, main_window)

  # If the result from the function returns 'The information in this file is Unicode', now it will call the function to merge all files
  if result_evaluate_file == 'The information in this file is Unicode':
    result_merge_files_handler = merge_files_handler(script_file_path, folder_name, timestamp, label_get_files, label_copy_files, label_copy_files_result, label_merge_status)
  # If not yes, it will show a confirmation window
  elif result_evaluate_file == 'The information in this file is not Unicode':
    open_confirmation_window(script_file_path, main_window, folder_name, timestamp, label_get_files, label_copy_files, label_copy_files_result, label_merge_status)

# Create UI
def ui():
  # Main window
  main_window = ctk.CTk()
  main_window.title('Excel Tools')
  main_window.minsize(width=300, height=300)
  main_window.resizable(False, False)
  main_window.config(padx=20, pady=20)

  # Title label
  label1 = ctk.CTkLabel(main_window, text="GLOBPAR EXCEL TOOLS", font=("Arial", 14, "bold"))
  label1.grid(column=1, row=1, pady=(20, 10))

  # Start button
  button1 = ctk.CTkButton(main_window, text="START", fg_color = 'teal', font=("Arial", 14, "bold"), command=lambda: main(main_window, label_get_files, label_copy_files, label_copy_files_result, label_merge_status))
  button1.grid(column=1, row=2)

  # Manual button
  button2 = ctk.CTkButton(main_window, text="Manual", fg_color = 'teal', width=(85), font=("Arial", 14, "bold"), command= lambda: open_manual())
  button2.grid(column=1, row=10, pady=(0, 20))

  # Labels for messages 
  label_get_files = ctk.CTkLabel(main_window, text="", font=("Arial", 14))
  label_get_files.grid(column=2, row=1, padx=(130))

  label_copy_files = ctk.CTkLabel(main_window, text="", font=("Arial", 14))
  label_copy_files.grid(column=2, row=2, padx=(110), pady=(0, 11))

  label_copy_files_result = ctk.CTkLabel(main_window, text="", font=("Arial", 14))
  label_copy_files_result.grid(column=2, row=3, padx=130, pady=(0, 11))

  label_merge_status = ctk.CTkLabel(main_window, text="", font=("Arial", 14))
  label_merge_status.grid(column=2, row=4, padx=(130), pady=(0, 11))

  main_window.mainloop()

if __name__ == '__main__':
  ui()