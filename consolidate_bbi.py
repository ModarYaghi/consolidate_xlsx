import pandas as pd
import tkinter as tk
from tkinter import filedialog


def select_file(prompt_message):
    """Open the file explorer to select an Excel file."""
    print(prompt_message)
    root = tk.Tk()
    root.withdraw()  # Hide the main window.

    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])  # Open the file dialog.

    return file_path


# Select input file
file_path = select_file("Please select the Excel file to process.")

# Load the datasets from the first and second sheets of the xlsx file
screening_sheet_name = 0
intake_sheet_name = 1

try:
    screening_df = pd.read_excel(file_path, sheet_name=screening_sheet_name, usecols=range(7))
    intake_df = pd.read_excel(file_path, sheet_name=intake_sheet_name, usecols=[0,1, 2, 3])
except FileNotFoundError:
    print(f"Error: File '{file_path}' not found")
    exit()
except KeyError:
    print(f"Error: Sheet '{screening_sheet_name}' or '{intake_sheet_name}' not found in file '{file_path}'")
    exit()

# Check that the 'Referral #' column exists in both dataframes
assert '02_rn' in screening_df.columns, "'Referral Number' column not found in the first sheet"
assert '02_rn' in intake_df.columns, "'Referral Number' column not found in the second sheet"

# Merge the dataframes on 'Referral #'
merged_df = pd.merge(screening_df, intake_df, on='02_rn', how='outer', suffixes=('_screening', '_intake'))

# Check that the merge was successful
assert not merged_df.empty, "The merged dataframe is empty"

# Remove duplicate columns
merged_df = merged_df.loc[:, ~merged_df.columns.duplicated()]

# Write the merged dataframe to a new sheet in the same Excel file
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
    # writer.book = pd.load_workbook(file_path)
    merged_df.to_excel(writer, sheet_name='bbi')

# Check that the merged dataframe is not empty
assert not merged_df.empty, "The merged dataframe is empty"

# --- End of file ---
