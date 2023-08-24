# Importación de las librerías necesarias
import pandas as pd
import glob
import openpyxl
from tkinter import filedialog

# Solicitar al usuario que seleccione una carpeta
folder_selected = filedialog.askdirectory()

# Verificar si no se seleccionó ninguna carpeta y salir del programa
if not folder_selected: 
    (exit) # Salir del programa

# Encontrar todos los archivos Excel (.xlsx) en la carpeta seleccionada
file_list = glob.glob(folder_selected + "/*.xlsx")

# Crear una lista vacía para almacenar DataFrames de archivos Excel
excel_list = []

# Iterar a través de la lista de archivos y leer cada archivo Excel en un DataFrame
for file in file_list:
    excel_list.append(pd.read_excel(file))
    
# Crear un DataFrame vacío para almacenar la consolidación    
excel_merged = pd.DataFrame()

# Iterar a través de la lista de DataFrames y concatenarlos en el DataFrame consolidado
for excel_file in excel_list:
    excel_merged = pd.concat([excel_merged, pd.DataFrame.from_records(excel_file)], ignore_index= False)

# Guardar el DataFrame consolidado en un archivo Excel
    excel_merged.to_excel("consolidated_file.xlsx")
