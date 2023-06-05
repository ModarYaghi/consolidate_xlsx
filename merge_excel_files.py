import os
import pandas as pd
import logging as logger


def merge_excel_files(dir_path, output_file, sheet_details):
    """ Merge all Excel files in a directory into one Excel file."""
    data = {}  # A dictionary to hold data for each sheet
    first_file = True
    sheet_order = [sheet['name'] for sheet in sheet_details]  # Extract sheet names from sheet_details

    for file in os.listdir(dir_path):
        if file.endswith('.xlsx') and file != output_file:  # Exclude the merge file
            file_path = os.path.join(dir_path, file)
            try:
                xls = pd.ExcelFile(file_path, engine='openpyxl')  # specify the engine manually

                for sheet_name in xls.sheet_names:
                    if sheet_name in sheet_order:  # Only process sheets that are in sheet_order
                        df = pd.read_excel(xls, sheet_name)
                        if sheet_name in data:
                            data[sheet_name] = pd.concat([data[sheet_name], df], ignore_index=True)
                        else:
                            data[sheet_name] = df

            except Exception as e:
                logger.error("Error processing %s: %s", file_path, e)

    # Write data to 'consolidated.xlsx'
    with pd.ExcelWriter(os.path.join(dir_path, output_file), engine='openpyxl') as writer:
        for sheet_name in sheet_order:
            if sheet_name in data:
                data[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
