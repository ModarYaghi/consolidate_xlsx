import pandas as pd

# Load the workbook
xl = pd.ExcelFile(r'C:\Users\myagh\OneDrive - SİVİL HUKUK DERNEĞi\Desktop\psc_src_files_test\TS_psc_Y-v03.xlsx')

# Create an empty list to hold dictionaries
sheets_columns = []

# Loop over all sheets
for sheet in xl.sheet_names:
    # Load the sheet into a dataframe
    df = xl.parse(sheet)

    # Create a dictionary for this sheet
    sheet_dict = {'sheet_name': sheet, 'columns': df.columns.tolist()}

    # Add the dictionary to the list
    sheets_columns.append(sheet_dict)

# Print the list of dictionaries
for sheet_dict in sheets_columns:
    print(sheet_dict)
