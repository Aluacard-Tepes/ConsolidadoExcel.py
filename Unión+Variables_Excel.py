# Importación de las librerías necesarias
import pandas as pd
import glob
import re
import openpyxl
import time
from tkinter import filedialog

print("Creado por : Agrega Tu Nombre )

# Solicitar al usuario que seleccione una carpeta
folder_selected = filedialog.askdirectory()

# Verificar si no se seleccionó ninguna carpeta y salir del programa
if not folder_selected:
    exit()  # Salir del programa

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
    excel_merged = pd.concat([excel_merged, pd.DataFrame.from_records(excel_file)], ignore_index=False)
    


# Definir una función para asignar valores de peso en las filas en función de las condiciones
def asignar_valor(row):
    if  
    else:
        return 0

# Definir una función para asignar valores de nombre de familia por producto en función de las condiciones
def asignar_familia(row):
    if 
    else:
        return 0

# Definir una función para asignar subgrupo por familia en función de las condiciones
def asignar_familia_Grupo(row):
    if 
    else:
        return 0

# Definir una función para asignar valores de COD = ATC Name en función de las condiciones  
def ATC_COD(row):
    if 
    else:
        return 0

#  Definir una función para asignar valores COD = Región en función de las condiciones  
def ATC_REG(row):
    if 
    else:
        return 0


# Filtrar y eliminar las filas con 'FLETES', 'FLETE ABONOS' y 'ANALISIS DE SUELOS  (ASU)' en la columna 'des_item'
excel_merged = excel_merged[~(excel_merged['des_item'].str.contains('FLETES'))]

# Filtrar y eliminar las filas con 'FLETES', 'FLETE ABONOS' y 'ANALISIS DE SUELOS  (ASU)' en la columna 'des_item'
excel_merged = excel_merged[~(excel_merged['des_item'].str.contains('FLETE ABONOS'))]

# Aplicar la función y agregar la columna "Nueva_Columna"
excel_merged["PesoItem"] = excel_merged.apply(asignar_valor, axis=1)

# Aplicar la función y agregar la columna "Familia"
excel_merged["Familia"] = excel_merged.apply(asignar_familia, axis=1)

# Aplicar la función y agregar la columna "Familia"
excel_merged["Grupo"] = excel_merged.apply(asignar_familia_Grupo, axis=1)

# Aplicar la función y agregar la columna "Familia"
excel_merged["ATC"] = excel_merged.apply(ATC_COD, axis=1)

# Aplicar la función y agregar la columna "Familia"
excel_merged["Región"] = excel_merged.apply(ATC_REG, axis=1)

# Aplicar condiciones y agregar la columna "Nueva_Columna"
excel_merged["CnTimac"] = excel_merged.apply(lambda row: -row["cantidad"] if "NC" in str(row["num_doc"]) else row["cantidad"], axis=1)

# Crear la nueva columna multiplicando CnTimac por PesoItem
excel_merged["Kg/Lt"] = excel_merged["CnTimac"] * excel_merged["PesoItem"]

#Crer la nueva columna el cual los fertilizantes se dividen en 1000
excel_merged["Tn/Lt"] = excel_merged.apply(lambda row: row["Kg/Lt"] / 1000 if "Fertilizantes" in str(row["Familia"]) else row["Kg/Lt"], axis=1)

#Crer la nueva columna el cual los Hidrosolubles y Fertilizantes estan / 10000
excel_merged["Tn"] = excel_merged.apply(lambda row: row["Tn/Lt"] / 1000 if "Líquidos" in str(row["Familia"]) else (row["Tn/Lt"] / 1000 if "Hidrosolubles" in str(row["Familia"]) else row["Tn/Lt"]), axis=1)

# Aplicar condiciones y agregar la columna "Nueva_Columna"
excel_merged["Valor_Def"] = excel_merged.apply(lambda row: -row["val_def"] if "NC" in str(row["num_doc"]) else row["val_def"], axis=1)

# Guardar el DataFrame consolidado en un archivo Excel
excel_merged.to_excel("consolidated_file.xlsx", index=False)


#pausar 1 min
time.sleep(60)


