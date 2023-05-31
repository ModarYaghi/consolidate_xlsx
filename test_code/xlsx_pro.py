import os
import pandas as pd
import shutil
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def get_initials(filename):
    """Extracts the initials from a filename."""
    parts = filename.split('_')
    if len(parts) >= 3:
        return parts[2].split('-')[0]
    else:
        raise ValueError(f"Unexpected filename format: {filename}")


def add_service_provider_column(df, provider_initials):
    """Adds a service provider initials column to a DataFrame."""
    df.insert(0, 'sp', provider_initials)
    return df


def remove_empty_rows(df):
    """Removes rows from a DataFrame where all cells (except in the '#' and 'Service Provider' columns) are NA."""
    df = df.dropna(axis=0, how='all', subset=df.columns[df.columns.isin(['sp', '#']) == False])
    return df


def rename_sheets_and_columns(xls, sheet_names, col_names):
    """Renames sheets and columns in an Excel file."""
    for old_sheet_name, new_sheet_name in sheet_names.items():
        if old_sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, old_sheet_name)
            df.rename(columns=col_names[old_sheet_name], inplace=True)
            df.to_excel(xls, sheet_name=new_sheet_name, index=False)


def load_and_clean_sheet(excel_file, sheet_name, cols_to_drop):
    """Loads and cleans a single sheet from an Excel file."""
    df = pd.read_excel(excel_file, sheet_name, header=None)
    if all(df.iloc[0].str.contains('Unnamed')):
        df = df.iloc[1:]
        df = df.iloc[:, 1:]
    df.columns = df.iloc[0]
    df = df.iloc[1:]
    df = df.reset_index(drop=True)
    df = df.drop(columns=[col for col in cols_to_drop if col in df.columns])
    return df


def clean_and_copy_files(src_dir, dst_dir_name='TS_processed', sheet_names={}, col_names={}, cols_to_drop=[]):
    """Cleans and copies Excel files from source directory to destination directory."""
    dst_dir = os.path.join(src_dir, dst_dir_name)
    os.makedirs(dst_dir, exist_ok=True)

    for fname in os.listdir(src_dir):
        if fname.endswith('.xlsx'):
            logger.info('Found excel file: %s', fname)
            src_file_path = os.path.join(src_dir, fname)
            dst_file_path = os.path.join(dst_dir, fname)

            # Copy file to destination directory
            shutil.copy(src_file_path, dst_file_path)

            # Clean and modify copied file
            xls = pd.ExcelFile(dst_file_path)
            provider_initials = get_initials(fname)
            rename_sheets_and_columns(xls, sheet_names, col_names)
            for sheet_name in xls.sheet_names:
                df = load_and_clean_sheet(xls, sheet_name, cols_to_drop)
                df = add_service_provider_column(df, provider_initials)
                df = remove_empty_rows(df)
                df.to_excel(dst_file_path, sheet_name=sheet_name, index=False)

            logger.info('Processed and copied file to: %s', dst_file_path)


def merge_excel_files(dir_path, output_file):
    """Merge all Excel files in a directory into one Excel file."""
    data = {}
    first_file = True

    for file in os.listdir(dir_path):
        if file.endswith('.xlsx') and file != output_file:
            file_path = os.path.join(dir_path, file)
            try:
                xls = pd.ExcelFile(file_path)
                if first_file:  # For the first file, save the order of the sheets
                    sheet_order = xls.sheet_names
                    first_file = False

                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name)
                    if sheet_name in data:
                        data[sheet_name] = pd.concat([data[sheet_name], df], ignore_index=True)
                    else:
                        data[sheet_name] = df

            except Exception as e:
                logger.error("Error processing %s: %s", file_path, e)

    # Write data to 'consolidated.xlsx'
    with pd.ExcelWriter(os.path.join(dir_path, output_file), engine='xlsxwriter') as writer:
        for sheet_name in sheet_order:
            data[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
