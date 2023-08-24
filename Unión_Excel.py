import pandas as pd
import glob
import openpyxl
from tkinter import filedialog

folder_selected = filedialog.askdirectory()

if not folder_selected: 
    (exit)

file_list = glob.glob(folder_selected + "/*.xlsx")

excel_list = []

for file in file_list:
    excel_list.append(pd.read_excel(file))
    
    
excel_merged = pd.DataFrame()

for excel_file in excel_list:
    excel_merged = pd.concat([excel_merged, pd.DataFrame.from_records(excel_file)], ignore_index= False)
    
    excel_merged.to_excel("consolidated_file.xlsx")
