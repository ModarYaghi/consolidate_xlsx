import os
import passwords as ps
import time
import tkinter as tk
from tkinter import filedialog
import xlsx_processor as xp


initial_dir = r"C:\Users\mfyag\OneDrive - SİVİL HUKUK DERNEĞi\Documents\Family Center\Tracking_System_Collective"

def select_path(prompt_message, select_type="file", initial_dir=initial_dir):
    print(prompt_message)
    root = tk.Tk()
    root.withdraw()  # Hide the main window.

    if select_type == "file":
        path = filedialog.askopenfilenames(initialdir=initial_dir)  # Open the file dialog.
    elif select_type == "dir":
        path = filedialog.askdirectory(initialdir=initial_dir)  # Open the directory dialog.
    else:
        raise ValueError("Invalid selection type. Expected 'file' or 'dir'.")

    return path


def main():
    """Main function to execute the script."""

    # Select input files
    files = select_path("Please select the Excel files to process.", select_type="file")

    # Select output directory
    dir_path = select_path("Please select the directory to write the output file to.", select_type="dir")

    dir_path = os.path.join(dir_path, 'consolidated')
    
    # Get the current script directory
    script_dir = os.path.dirname(os.path.realpath(__file__))

    # Get the json file path
    json_file = os.path.join(script_dir, 'sheets_details.json')


    # Copy selected Excel files to 'consolidated_tracking_tools' directory
    for file in files:
        try:
            # First, try to open the file without a password
            xp.decrypt_and_copy_xlsx_file(file, dir_path)
            print(f"{file} does not require a password.")
        except Exception as e:
            print(f"Failed to decrypt {file} without password: {str(e)}")
            # If it fails, try to open the file with each password
            for password in ps.passwords:
                try:
                    xp.decrypt_and_copy_xlsx_file(file, dir_path, password)
                    print(f"Password for {file} is {password}.")
                    break  # If the password is correct, go to the next file
                except Exception as e:
                    print(f"Failed to decrypt {file} with password {password}: {str(e)}")
                    continue  # If the password is incorrect, try the next password 


    # Pause for 5 seconds to allow all write operations to complete
    time.sleep(5)

    # Specify columns and sheets to drop
    cols_to_drop = ["Unnamed: 0", "System ID"]
    sheets_to_drop = ["GZT_Service_Map", "Glossary", "Drop-down"]

    # Clean Excel files in the 'consolidated_tracking_tools' directory
    xp.clean_and_rename_excel_files(dir_path, cols_to_drop, sheets_to_drop, json_file)

    # Consolidate all the cleaned Excel files into one
    xp.merge_excel_files(dir_path, 'TS_psc_All.xlsx')

if __name__ == '__main__':
    main()
