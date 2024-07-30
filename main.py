import openpyxl
import os
import xlwings as xw
import tkinter as tk
import time
from pathlib import Path
from tkinter import messagebox, filedialog

# The function will evaluate the value of the 'Unicode' field
def evaluate_file(file_path):

  if file_path:
    try:
      # Open the file using the selected path
      wb = openpyxl.load_workbook(file_path)
      sheet = wb.active
      unicode_value = sheet['B6'].value
        
      # Evaluate the 'Unicode' field value
      # If its value is 'yes', the script will continue without warnings
      if unicode_value.lower() in ('yes'):
        messagebox.showinfo('Info', 'The information in this file is Unicode')
        return 'The information in this file is Unicode'
      
      # if the value in the Unicode field is not 'yes', it will show a warning message 
      else:
        answer = messagebox.askquestion(
          'Confirmation',
          'The information in this file IS NOT Unicode, do you want to continue?',
          icon='warning'
        )
        
        if answer in ('yes', 'ok'):
          messagebox.showinfo('Info', 'The information in this file IS NOT Unicode. The process will continue')
          return 'The information in this file IS NOT Unicode. The process will continue'

        else:
          messagebox.showinfo('Info', 'Operation cancelled')
          return 'Operation cancelled'
        
    except:
      return f"Error: The file '{file_path}' was not found"
      
  else:
    return 'No file selected'

# Opens the specified Excel file, renames the first worksheet to the filename (without extension), and saves the changes
def rename_sheet(file_path):
  if not file_path:
    return 'No se encontró ningún archivo'

  try:
    # Open the workbook and access the worksheet
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    
    # Extract filename without extension
    filename_extraction, file_extension = os.path.splitext(os.path.basename(file_path))
    new_sheet_name = filename_extraction
    
    # Evaluate if the name was already changed
    if sheet.title == new_sheet_name:
      return f"The sheet name was already changed to '{new_sheet_name}'"
    
    else:
      # Rename the worksheet
      sheet.title = new_sheet_name
    
      # Save changes to the workbook
      wb.save(file_path)
      return f"Sheet renamed to '{new_sheet_name}'"
    
  except Exception as e:
    return f"Error: {str(e)}"

# This function will ask for all the other files and will add them as sheets
def merge_files(file_path, folder_name, timestamp):
  if file_path:
    with xw.App(visible=False) as app:
      excel_files = list(Path(folder_name).glob('*.xlsx'))
      combined_wb = xw.Book()
    
      for excel_file in excel_files:
        wb = xw.Book(excel_file)
        for sheet in wb.sheets:
          # Obtener el nombre del archivo sin la extensión .xlsx
          sheet_name = excel_file.stem
          # Copiar la hoja y asignarle el nombre del archivo
          sheet.api.Copy(After=combined_wb.sheets[0].api)
          combined_wb.sheets[1].name = sheet_name

      
      # Save the combined workbook with a timestamp
      combined_wb.sheets[0].delete()
      combined_wb.save(f'merged{timestamp}.xlsx')

      if len(combined_wb.app.books) == 1:
        combined_wb.app.quit()
        return 'All files merged successfully'
      else:
        combined_wb.close()
        return 'Error: Failed to merge files'

def main():
  root = tk.Tk()
  root.withdraw()

  # Get current file path (The excel file should be in the same)
  script_file_path = os.path.dirname(os.path.abspath(__file__))
  folder_name = 'documents'
  file_name = 'TechnicalProperties.xlsx'

  # Get the excel file path
  file_path = os.path.join(script_file_path, folder_name, file_name)

  t = time.localtime()
  timestamp = time.strftime('%Y-%m-%d_%H-%M-%S', t)

  # Call the first function and print the result
  evaluation_result = evaluate_file(file_path)
  print(evaluation_result)
  
  # If the operation was cancelled by the user, just return the result
  if evaluation_result == 'Operation cancelled':
    return evaluation_result
  
  # If the user wants to proceed, execute the second function
  else:
    rename_result = rename_sheet(file_path)
    print(rename_result)

  merge_files(file_path, folder_name, timestamp)

if __name__ == '__main__':
  main()