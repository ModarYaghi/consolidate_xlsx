from io import BytesIO
import json
import logging
import msoffcrypto
import os
import pandas as pd

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def decrypt_and_copy_xlsx_file(file, dst_dir, password=None):
    """Decrypts a password-protected Excel file (if a password is provided) and copies it to the destination
    directory."""

    os.makedirs(dst_dir, exist_ok=True)

    fname = os.path.basename(file)
    if fname.endswith('.xlsx'):
        logger.info('Found excel file: %s', fname)

        # If a password has been provided, decrypt the Excel file
        if password is not None:
            with open(file, "rb") as f:
                crypto = msoffcrypto.OfficeFile(f)
                crypto.load_key(password=password)
                decrypted_file = BytesIO() # f"{file}_decrypted.xlsx"
                crypto.decrypt(decrypted_file)
                decrypted_file.seek(0) # Reset the file pointer to the beginning of the file
        else:
            decrypted_file = open(file, "rb")

        # Read the (decrypted) Excel file
        xls = pd.ExcelFile(decrypted_file)
        with pd.ExcelWriter(os.path.join(dst_dir, fname)) as writer:
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        logger.info('Copied file to: %s', os.path.join(dst_dir, fname))

        # If a password has been provided, delete the decrypted file after copying it
        if password is not None:
            # os.remove(decrypted_file)
            decrypted_file.close()


def get_initials(filename):
    """Extracts the initials from a filename."""
    logger.info('Extracting initials from filename: %s', filename)
    parts = filename.split('_')
    if len(parts) >= 3:
        # return parts[2].split('-')[0]
        initials = parts[2].split('-')[0]
    else:
        raise ValueError(f"Unexpected filename format: {filename}")
    logger.info('Extracted initials: %s', initials)
    return initials


def add_service_provider_column(df, provider_initials):
    """Adds a service provider initials column to a DataFrame."""
    logger.info('Adding service provider initials column: %s', provider_initials)
    df.insert(0, 'sp', provider_initials)
    logger.info('Added service provider initials column.')
    return df


def remove_empty_rows(df):
    """Removes rows from a DataFrame where all cells (except in the '#' and 'Service Provider' columns) are NA."""
    logger.info('Removing empty rows.')
    df = df.dropna(axis=0, how='all', subset=df.columns[df.columns.isin(['sp', '#']) == False])
    logger.info('Removed empty rows.')
    return df


def load_and_clean_sheet(excel_file, sheet_name, cols_to_drop):
    """Loads and cleans a single sheet from an Excel file."""
    logger.info('Loading and cleaning sheet: %s', sheet_name)
    df = pd.read_excel(excel_file, sheet_name, header=None)
    if all(df.iloc[0].str.contains('Unnamed')):
        df = df.iloc[1:]
        df = df.iloc[:, 1:]
    df.columns = df.iloc[0]
    df = df.iloc[1:]
    df = df.reset_index(drop=True)
    df = df.drop(columns=[col for col in cols_to_drop if col in df.columns])
    logger.info('Loaded and cleaned sheet.')
    return df


def clean_and_rename_excel_files(dir_path, cols_to_drop, sheets_to_drop, json_file):
    """ Cleans Excel files in a directory by dropping specified columns and sheets, and renames sheets and columns."""

    logger.info('Cleaning and renaming Excel files in directory: %s', dir_path)

    # Read the JSON file
    try:
        with open(json_file) as f:
            name_mappings = json.load(f)
    except Exception as e:
        logger.error('Error reading JSON file %s: %s', json_file, e)
        return

    for file in os.listdir(dir_path):
        if file.endswith('.xlsx'):
            provider_initials = get_initials(file)
            file_path = os.path.join(dir_path, file)
            temp_file_path = os.path.join(dir_path, f'temp_{file}')
            try:
                with pd.ExcelFile(file_path) as excel_file:
                    with pd.ExcelWriter(temp_file_path, engine='openpyxl') as writer:
                        for i, sheet_name in enumerate(excel_file.sheet_names):
                            if sheet_name not in sheets_to_drop:
                                df = load_and_clean_sheet(excel_file, sheet_name, cols_to_drop)
                                df = add_service_provider_column(df, provider_initials)
                                df = remove_empty_rows(df)

                                # Get the name mapping for this sheet
                                name_mapping = name_mappings[i]

                                # Rename the columns
                                df.columns = name_mapping['columns']

                                # Write the DataFrame to the new Excel file
                                df.to_excel(writer, sheet_name=name_mapping['name'], index=False)

                os.remove(file_path)
                os.rename(temp_file_path, file_path)
            except Exception as e:
                logger.error("Error processing %s: %s", file_path, e)
    logger.info('Cleaned and renamed Excel files in directory.')


def merge_excel_files(dir_path, output_file):
    """ Merge all Excel files in a directory into one Excel file."""
    logger.info('Merging Excel files in directory: %s', dir_path)
    data = {}  # A dictionary to hold data for each sheet
    first_file = True

    for file in os.listdir(dir_path):
        logger.info('Processing file: %s', file)
        if file.endswith('.xlsx') and file != output_file:  # Exclude the merge file
            file_path = os.path.join(dir_path, file)
            try:
                xls = pd.ExcelFile(file_path, engine='openpyxl')  # specify the engine manually
                if first_file:  # For the first file, save the order of the sheets
                    sheet_order = xls.sheet_names  # Make a list of sheet names if it's the first file.
                    first_file = False

                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name)
                    if sheet_name in data:
                        data[sheet_name] = pd.concat([data[sheet_name], df], ignore_index=True)
                    else:
                        data[sheet_name] = df
            except Exception as e:
                logger.error("Error processing %s: %s", file_path, e)
        logger.info('Processed file: %s', file)

    # Write data to 'consolidated.xlsx'
    with pd.ExcelWriter(os.path.join(dir_path, output_file), engine='openpyxl') as writer:
        for sheet_name in sheet_order:
            logger.info('Writing sheet: %s', sheet_name)
            data[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
            logger.info('Wrote sheet: %s', sheet_name)
