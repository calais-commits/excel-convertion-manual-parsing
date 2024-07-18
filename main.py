import openpyxl
import tkinter as tk
from tkinter import messagebox, filedialog

# The function will evaluate the value of the 'Unicode' field
def evaluate_file(file_path):

  if file_path:
    try:
      # Open the file using the selected path
      wb = openpyxl.load_workbook(file_path)
      sheet = wb['Doc1']
      unicode_value = sheet['B6'].value
        
      # Evaluate the 'Unicode' field value
      # If its value is 'yes', the script will continue without warnings
      if unicode_value.lower() in ('yes', 'y'):
        result = 'La información de este archivo es Unicode'
      
      # if the value in the Unicode field is not 'yes', it will show a warning message 
      else:
        answer = messagebox.askquestion(
          'Confirmation',
          'The information in this file IS NOT Unicode, do you want to continue?',
          icon='warning'
        )
        
        if answer in ('yes', 'ok'):
          return 'The information in this file IS NOT Unicode. The process will continue'

        else:
          return 'Operation cancelled'
        
    except:
      return f"Error: El archivo '{file_path}' no se encontró "
      
  else:
    return 'No se seleccionó ningun archivo'
  
# This function will open a file explorer and ask to the user to select a file
def get_file():
  file_path = tk.filedialog.askopenfilename(
    title='Seleccione el archivo TechnicalProperties',
    filetypes=[("Archivos de Excel", "*.xlsx *.xlsm *.xls")]
  )
  
  result = evaluate_file(file_path)
  print(result)
  
def main():
  root = tk.Tk()
  root.withdraw()
  get_file()

if __name__ == '__main__':
  main()