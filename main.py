import os
import pandas as pd
import numpy as np
import xlrd
from dotenv import load_dotenv
import re
from swifter import swifter

# Load environment variables
load_dotenv()

# Select server and download routes
server_route = os.getenv('server_route')
download_route = os.getenv('download_route')

# Name of the files with the data for suspension authorizations and the broadcasting stations
file_aut_sus = os.getenv('file_aut_sus')
file_aut_bp = os.getenv('file_aut_bp')

# Columns to be selected in the data files
columnasAUT = ['No. INGRESO ARCOTEL', 'FECHA INGRESO ARCOTEL', 'NOMBRE ESTACIÓN', 'M/R', 'FREC / CANAL',
               'CIUDAD PRINCIPAL COBERTURA', 'DIAS SOLICITADOS', 'DIAS AUTORIZADOS', 'No. OFICIO ARCOTEL',
               'FECHA OFICIO', 'FECHA INICIO SUSPENSION ', 'DIAS', 'FECHA FIN SUSPENSION', 'ZONAL']
columnasAUTBP = ['No. INGRESO', 'FECHA', 'OFICIO ARCOTEL', 'FECHA OFICIO', 'NOMBRE ESTACIÓN',
                 'CIUDAD PRINCIPAL COBERTURA', 'M/R', 'CANAL', 'UBICACIÓN TRANSMISOR', 'MODIFICACION TEMPORAL',
                 'PLAZO OTORGADO', 'FECHA INICIO PLAZO/NOTIFICACION', 'ZONAL']

# Load the Excel workbook with openpyxl
workbook = xlrd.open_workbook(server_route + file_aut_bp, formatting_info=True)

# Select the desired sheet (replace 'Sheet1' with your actual sheet name)
sheet = workbook.sheet_by_name('MTTEMP')

# Get the column names from the first row
column_names = sheet.row_values(0)

# Assuming "CIUDAD PRINCIPAL COBERTURA" and "CANAL" are the desired columns, you can manually set their indices
# based on the printed information or search for the indices dynamically:
desired_ciudad_column_name = "CIUDAD PRINCIPAL COBERTURA"
ciudad_column_index = column_names.index(desired_ciudad_column_name)

desired_canal_column_name = "CANAL"
canal_column_index = column_names.index(desired_canal_column_name)

# Get the values from the sheet and handle merged cells
data = []
for row_num in range(sheet.nrows):
    row_values = []
    for col_num in range(sheet.ncols):
        # Convert datetime values from serial numbers to datetime objects
        cell_value = sheet.cell_value(row_num, col_num)
        if sheet.cell_type(row_num, col_num) == xlrd.XL_CELL_DATE:
            cell_value = xlrd.xldate_as_datetime(cell_value, workbook.datemode)

        # Check if the cell is part of a merged range in the current column
        is_merged_in_column = any(
            start_row <= row_num < end_row and col_num == start_col
            for start_row, end_row, start_col, _ in sheet.merged_cells
        )

        # Use the value from the top-left cell of the merged range in the current column
        if is_merged_in_column:
            for start_row, end_row, start_col, _ in sheet.merged_cells:
                if start_row <= row_num < end_row and col_num == start_col:
                    cell_value = sheet.cell_value(start_row, start_col)
                    # Convert datetime values from serial numbers to datetime objects for merged cells
                    if sheet.cell_type(start_row, start_col) == xlrd.XL_CELL_DATE:
                        cell_value = xlrd.xldate_as_datetime(cell_value, workbook.datemode)
                    break

        row_values.append(cell_value)

    # Split values in the "CIUDAD PRINCIPAL" column based on newline character or whitespace
    ciudad_values = re.split(r'\n|\s{2,}', row_values[ciudad_column_index])

    # Convert "CANAL" values to strings and split based on newline character or whitespace
    canal_values = re.split(r'\n|\s{2,}', str(row_values[canal_column_index]))

    # Create new rows with the split values
    for ciudad_value, canal_value in zip(ciudad_values, canal_values):
        new_row = row_values.copy()
        new_row[ciudad_column_index] = ciudad_value
        new_row[canal_column_index] = canal_value
        data.append(new_row)

# Convert datetime values from serial numbers to datetime objects for specific columns
datetime_columns = ['FECHA OFICIO', 'FECHA INICIO PLAZO/NOTIFICACION']
for col_name in datetime_columns:
    col_index = columnasAUTBP.index(col_name)
    for row in data:
        if sheet.cell_type(row_num, col_index) == xlrd.XL_CELL_DATE:
            row[col_index] = xlrd.xldate_as_datetime(row[col_index], workbook.datemode)

# Create a new DataFrame with the modified data
df1 = pd.read_excel(server_route + file_aut_sus, engine='xlrd', header=1, usecols=columnasAUT,
                    sheet_name='SUSPENSIÓN EMISIONES 2021-2023')
df1['FECHA INGRESO ARCOTEL'] = pd.to_datetime(df1['FECHA INGRESO ARCOTEL'], errors='coerce')
df1['FECHA OFICIO'] = pd.to_datetime(df1['FECHA OFICIO'], errors='coerce')
df1['FECHA INICIO SUSPENSION '] = pd.to_datetime(df1['FECHA INICIO SUSPENSION '], errors='coerce')
df1['FECHA INICIO SUSPENSION '].replace(['', '-'], np.nan, inplace=True)
df1.dropna(subset=['FECHA INICIO SUSPENSION '], inplace=True)
df1['FREC / CANAL'] = pd.to_numeric(df1['FREC / CANAL'], errors='coerce')
df1['FREC / CANAL'].replace(['', '-'], np.nan, inplace=True)
df1.dropna(subset=['FREC / CANAL'], inplace=True)
df1['DIAS'] = pd.to_numeric(df1['DIAS'], errors='coerce')
# Remove rows with '-' or empty values in 'column_name'
df1['No. OFICIO ARCOTEL'] = df1['No. OFICIO ARCOTEL'].str.strip()
df1['No. OFICIO ARCOTEL'].replace(['', '-'], np.nan, inplace=True)
df1.dropna(subset=['No. OFICIO ARCOTEL'], inplace=True)
df1['Tipo'] = 'S'

df2 = pd.DataFrame(data)
df2.columns = df2.iloc[0]
df2 = df2[1:].reset_index(drop=True).rename_axis(None, axis=1)
df2 = df2[columnasAUTBP]
df2['OFICIO ARCOTEL'] = df2['OFICIO ARCOTEL'].str.strip()
df2['OFICIO ARCOTEL'].replace(['', '-'], np.nan, inplace=True)
df2.dropna(subset=['OFICIO ARCOTEL'], inplace=True)
df2['CANAL'] = pd.to_numeric(df2['CANAL'], errors='coerce')
df2['CANAL'].replace(['', '-'], np.nan, inplace=True)
df2.dropna(subset=['CANAL'], inplace=True)
df2['MODIFICACION TEMPORAL'] = df2['MODIFICACION TEMPORAL'].str.replace('\n', ' ')
df2['FECHA OFICIO'] = pd.to_datetime(df2['FECHA OFICIO'], errors='coerce')
df2['FECHA INICIO PLAZO/NOTIFICACION'] = pd.to_datetime(df2['FECHA INICIO PLAZO/NOTIFICACION'],
                                                        errors='coerce')
df2['FECHA INICIO PLAZO/NOTIFICACION'].replace(['', '-'], np.nan, inplace=True)
df2.dropna(subset=['FECHA INICIO PLAZO/NOTIFICACION'], inplace=True)
df2['PLAZO OTORGADO'] = df2['PLAZO OTORGADO'].swifter.apply(
    lambda x: pd.to_numeric(''.join(filter(str.isdigit, str(x))), errors='coerce') if isinstance(x,
                                                                                                 str) and 'días' in x else x)
df2['Tipo'] = 'BP'
df2['DIAS AUTORIZADOS'] = df2['PLAZO OTORGADO'].copy()
df2['DIAS'] = df2['PLAZO OTORGADO'].copy()
df2['FECHA FIN BAJA POTENCIA'] = df2['FECHA INICIO PLAZO/NOTIFICACION'] + pd.to_timedelta(df2['DIAS'] - 1, unit='D')
df2_columns = ['No. INGRESO', 'FECHA', 'NOMBRE ESTACIÓN', 'M/R', 'CANAL', 'CIUDAD PRINCIPAL COBERTURA',
               'PLAZO OTORGADO', 'DIAS AUTORIZADOS', 'OFICIO ARCOTEL', 'FECHA OFICIO',
               'FECHA INICIO PLAZO/NOTIFICACION', 'DIAS', 'FECHA FIN BAJA POTENCIA', 'ZONAL', 'UBICACIÓN TRANSMISOR',
               'MODIFICACION TEMPORAL', 'Tipo']
df2 = df2.reindex(columns=df2_columns)

# Create a Pandas Excel writer using XlsxWriter as the engine
"""Remove the previous file if already exist"""
excel_file_path = os.getenv('excel_file_path')
if os.path.exists(f'{excel_file_path}/autorizaciones.xlsx'):
    os.remove(f'{excel_file_path}/autorizaciones.xlsx')

with pd.ExcelWriter(f'{excel_file_path}/autorizaciones.xlsx') as writer:
    # Write each dataframe to a different sheet
    df1.to_excel(writer, sheet_name='Suspensión', index=False)
    df2.to_excel(writer, sheet_name='Baja Potencia', index=False)
    worksheet1 = writer.sheets['Suspensión']
    worksheet2 = writer.sheets['Baja Potencia']

    (max_row1, max_col1) = df1.shape
    (max_row2, max_col2) = df2.shape

    worksheet1.autofilter(0, 0, 0, int(max_col1 - 1))
    worksheet2.autofilter(0, 0, 0, int(max_col2 - 1))

# Save the Excel file
print(f"Excel file saved to: {excel_file_path}")
