import openpyxl
import os
import xlwings as xw
import tkinter as tk
import threading
import time
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
def merge_files_handler(folder_name, timestamp, label_get_files, label_copy_files, label_copy_files_result, label_merge_status):
  # Get the excel files in the folder and open a Workbook
  print('Getting files')
  label_get_files.configure(text="Getting files...")
  
  def copy_files():
    with xw.App(visible=False) as app:
      excel_files = list(Path(folder_name).glob('*.xlsx'))
      combined_wb = xw.Book()
      
      def copy_files_status():
        print('Copying files')
        label_copy_files.configure(text="Copying files...")

      def copy_files_result():
        print('Saving changes')
        label_copy_files_result.configure(text="Saving changes...")

      if len(excel_files) > 0:
        thread2 = threading.Thread(target=copy_files_status)
        thread2.start()
        for excel_file in excel_files:
          wb = xw.Book(excel_file)
          for sheet in wb.sheets:
            # Get the file name without the extension .xlsx
            sheet_name = excel_file.stem
            # Copy the sheet and assign the file name as the sheet name
            sheet.api.Copy(After=combined_wb.sheets[0].api)
            combined_wb.sheets[1].name = sheet_name

          # Save the combined workbook with a timestamp
        combined_wb.sheets[0].delete()
        result_file_name = f'merged{timestamp}.xlsx'
        combined_wb.save(result_file_name)
        thread3 = threading.Thread(target=copy_files_result)
        thread3.start()
        label_merge_status.configure(text="Files merged successfully")

  thread = threading.Thread(target=copy_files)
  thread.start()


def close_windows(main_window, confirmation_window):
  confirmation_window.destroy()
  main_window.destroy()
  print('Confirmation result: no')

def confirm_and_return_yes(confirmation_window, folder_name, timestamp, label_get_files, label_copy_files, label_copy_files_result, label_merge_status):
  confirmation_window.destroy()
  print('Confirmation result: yes')
  result_merge_files_handler = merge_files_handler(folder_name, timestamp, label_get_files, label_copy_files, label_copy_files_result, label_merge_status)
  return 'yes'

def open_confirmation_window(main_window, folder_name, timestamp, label_get_files, label_copy_files, label_copy_files_result, label_merge_status):
  confirmation_window = ctk.CTkToplevel()
  confirmation_window.focus_force()
  confirmation_window.title("Confirmation")
  confirmation_window.geometry("450x120")

  label = ctk.CTkLabel(confirmation_window, text="The information in the file is not Unicode. Are you sure you want to continue?")
  label.pack(pady=20)

  yes_button = ctk.CTkButton(confirmation_window, text="Yes", fg_color=("teal"), width=(100), command=lambda: confirm_and_return_yes(confirmation_window, folder_name, timestamp, label_get_files, label_copy_files, label_copy_files_result, label_merge_status))
  yes_button.pack(side=tk.LEFT, padx=50)

  no_button = ctk.CTkButton(confirmation_window, text="No", fg_color=("teal"), width=(100), command=lambda: close_windows(main_window, confirmation_window))
  no_button.pack(side=tk.RIGHT, padx=50)


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
  print(result_evaluate_file)

  # If the result from the function returns 'The information in this file is Unicode', now it will call the function to merge all files
  if result_evaluate_file == 'The information in this file is Unicode':
    result_merge_files_handler = merge_files_handler(folder_name, timestamp, label_get_files, label_copy_files, label_copy_files_result, label_merge_status)
    print(result_merge_files_handler)
  # If not yes, it will show a confirmation window
  elif result_evaluate_file == 'The information in this file is not Unicode':
    open_confirmation_window(main_window, folder_name, timestamp, label_get_files, label_copy_files, label_copy_files_result, label_merge_status)

def ui():
  # Create UI
  main_window = ctk.CTk()

  main_window.title('Excel Tools')
  main_window.minsize(width=300, height=300)
  main_window.resizable(False, False)
  main_window.config(padx=20, pady=20)

  label1 = ctk.CTkLabel(main_window, text="GLOBPAR EXCEL TOOLS", font=("Arial", 14, "bold"))
  label1.grid(column=1, row=1, pady=(20, 10))

  button1 = ctk.CTkButton(main_window, text="Start", fg_color = 'teal', font=("Arial", 14, "bold"), command=lambda: main(main_window, label_get_files, label_copy_files, label_copy_files_result, label_merge_status))
  button1.grid(column=1, row=2)

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
  # Estimación de horas
  ui()