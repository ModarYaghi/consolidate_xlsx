import os
import time
import xlsx_processor as xp

def main():
    """Main function to execute the script."""

    # Get the current working directory
    cwd = os.getcwd()

    # Get the 'TS_processed' directory
    dir_path = os.path.join(cwd, 'TS_processed')

    # Copy Excel files from cwd to 'consolidated_tracking_tools' directory
    xp.copy_xlsx_files(cwd)

    # Pause for 5 seconds to allow all write operations to complete
    time.sleep(5)

    # Specify columns and sheets to drop
    cols_to_drop = ["Unnamed: 0", "System ID"]
    sheets_to_drop = ["GZT_Service_Map", "Glossary", "Drop-down"]

    # Clean Excel files in the 'consolidated_tracking_tools' directory
    xp.clean_excel_files(dir_path, cols_to_drop, sheets_to_drop)

    # Consolidate all the cleaned Excel files into one
    xp.merge_excel_files(dir_path, 'TS_psc_All.xlsx')

if __name__ == '__main__':
    main()
