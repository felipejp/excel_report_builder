import pandas as pd
import csv
import openpyxl
import os
import json
from datetime import datetime

# Open file with excel files to work with
with open("excel_files.json", "r") as excel_files_file:

    excel_files = json.load(excel_files_file)

# Open file with sheets names    
with open("excel_sheets.json", "r") as excel_sheets_file:

    excel_sheets = json.load(excel_sheets_file)

directory = "C:/Users/fbova/Documents/Documentos IP/Reporte semanal trafico de Internet/06-05-24"
file_name = "Trafico de Internet 06-05-24.xlsx"
    
try:
    # Build path name
    complete_path = os.path.join(directory, file_name)
    
    sheets_index = 1
    
    # Open existing excel file
    workbook = openpyxl.load_workbook(complete_path)

    # Iterate files (TASA, GGC1, etc.) taking names from json file
    for file, file_name in excel_files.items():
        
        # Iterate sheets (TASA BNZ, GGC - BNZ - 1, etc.) taking sheet names from json file
        sheet = workbook[excel_sheets["sheet" + str(sheets_index)]]
        
        column_A = 'A'  # timestamp
        column_B = 'B'  # INBOUND
        column_C = 'C'  # OUTBOUND
        row_init = 2  # Fila inicial para empezar a insertar los valores
    
        # Open CSV file
        with open(file_name, newline='') as csv_file:
            # Creates a CSV reader
            csv_reader = csv.reader(csv_file)
        
            # Iterates on first 10 rows doing nothing to skip headers
            for index, _ in enumerate(csv_reader):
                if index >= 9:
                    break  
        
            # Iterates on each row of CSV file
            row_number = 1
            for row in csv_reader:
            
                date_data      = row[1]
                text_data_in   = row[2]
                text_data_out  = row[3]
                
                # Inserts values in FECHA column
                cell = column_A + str(row_init)
                sheet[cell] = date_data
            
                # Eliminates character "G", "M" or "K" of text data and converts into a Gbps number
                if 'G' in text_data_in:
                    number = float(text_data_in.replace('G', ''))
                elif 'M' in text_data_in:
                    number = float(text_data_in.replace('M', ''))/1000
                elif 'K' in text_data_in:
                    number = float(text_data_in.replace('K', ''))/1000000
                else:
                    number = float(text_data_in)/1000000000
                
                # Inserts values in INBOUND column
                cell = column_B + str(row_init)
                sheet[cell] = number
                
                if 'G' in text_data_out:
                   number = float(text_data_out.replace('G', ''))
                elif 'M' in text_data_out:
                   number = float(text_data_out.replace('M', ''))/1000
                elif 'K' in text_data_out:
                   number = float(text_data_out.replace('K', ''))/1000000
                else:
                   number = float(text_data_out)/1000000000

                # Inserts values in OUTBOUND column
                cell = column_C + str(row_init)
                sheet[cell] = number
            
                row_init += 1
                row_number += 1
                
            if(row_number < 168):
                print("Atención: el archivo " + str(file_name) + " contiene menos de 168 filas")
    
        sheets_index += 1   
        
    # Saves changes on a new file
    workbook.save('archivo_modificado.xlsx')

except Exception as e:
    print("Ocurrió un error al abrir el archivo:", e)