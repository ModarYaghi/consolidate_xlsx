import os
import pandas as pd
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def copy_xlsx_files(src_dir, dst_dir_name='TS_processed'):
    """ Copies all Excel files from the source directory to the destination directory."""
    dst_dir = os.path.join(src_dir, dst_dir_name)
    os.makedirs(dst_dir, exist_ok=True)

    for fname in os.listdir(src_dir):
        if fname.endswith('.xlsx'):
            logger.info('Found excel file: %s', fname)
            xls = pd.ExcelFile(os.path.join(src_dir, fname))
            with pd.ExcelWriter(os.path.join(dst_dir, fname)) as writer:
                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name)
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            logger.info('Copied file to: %s', os.path.join(dst_dir, fname))


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


def clean_excel_files(dir_path, cols_to_drop, sheets_to_drop):
    """ Cleans Excel files in a directory by dropping specified columns and sheets."""
    for file in os.listdir(dir_path):
        if file.endswith('.xlsx'):
            provider_initials = get_initials(file)
            file_path = os.path.join(dir_path, file)
            temp_file_path = os.path.join(dir_path, f'temp_{file}')
            try:
                with pd.ExcelFile(file_path) as excel_file:
                    with pd.ExcelWriter(temp_file_path) as writer:
                        for sheet_name in excel_file.sheet_names:
                            if sheet_name not in sheets_to_drop:
                                df = load_and_clean_sheet(excel_file, sheet_name, cols_to_drop)
                                df = add_service_provider_column(df, provider_initials)
                                df = remove_empty_rows(df)
                                df.to_excel(writer, sheet_name=sheet_name, index=False)
                os.remove(file_path)
                os.rename(temp_file_path, file_path)
            except Exception as e:
                logger.error("Error processing %s: %s", file_path, e)


def merge_excel_files(dir_path, output_file):
    """ Merge all Excel files in a directory into one Excel file."""
    data = {}  # A dictionary to hold data for each sheet
    first_file = True

    for file in os.listdir(dir_path):
        if file.endswith('.xlsx') and file != output_file:  # Exclude the merge file
            file_path = os.path.join(dir_path, file)
            try:
                xls = pd.ExcelFile(file_path, engine='openpyxl')  # specify the engine manually
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
