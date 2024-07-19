import openpyxl
import os
import tkinter as tk
from tkinter import messagebox, filedialog


# This function will open a file explorer and ask to the user to select a file
def get_file():
  file_path = tk.filedialog.askopenfilename(
    title='Seleccione el archivo TechnicalProperties',
    filetypes=[("Archivos de Excel", "*.xlsx *.xlsm *.xls")]
  )
  
  # Get result from first function and print it
  evaluation_result = evaluate_file(file_path)
  print(evaluation_result)
  
  # If the operation was cancelled by the user, just return the result
  if evaluation_result == 'Operation cancelled':
    return evaluation_result
  
  # If the user wants to proceed, execute the second function
  else:
    rename_result = rename_sheet(file_path)
    print(rename_result)

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
  
def main():
  root = tk.Tk()
  root.withdraw()
  get_file()

if __name__ == '__main__':
  main()