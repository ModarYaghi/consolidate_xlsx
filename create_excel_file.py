import json
import pandas as pd

def create_excel(json_file, output_filename):
    # Load the settings from the JSON file
    with open(json_file) as f:
        settings = json.load(f)

    # Create a new Excel writer object
    writer = pd.ExcelWriter(output_filename)

    # If the top-level JSON structure is a list, we iterate over the list instead
    for setting in settings:
        # Extract the sheet_name and columns from the current list item
        sheet_name = setting['name']
        columns = setting['columns']

        # Here you need to have your data ready. As an example, I'm using a DataFrame of random values.
        df = pd.DataFrame(columns=columns)
        
        # Write the DataFrame to an Excel sheet
        df.to_excel(writer, sheet_name=sheet_name, index=False)

    # Save the Excel file
    writer.save()



json_file = 'sheets_details.json'
output_filename = 'output.xlsx'

create_excel(json_file, output_filename) 
