# Consolidate XLSX

Consolidate XLSX is a Python module aimed at automating the task of cleaning and formatting Excel files within a specified directory. This is particularly useful when dealing with large quantities of Excel files that require uniform formatting or certain data to be excluded.

## Functionality

The module provides two key functions:

1. **copy_xlsx_files**: This function copies all Excel files from a source directory to a new directory, while preserving all columns and sheet names. This can be used as a preparatory step for processing, allowing the original files to remain unchanged.

2. **clean_excel_files**: This function cleans the Excel files in a specified directory by removing undesired columns and sheets. It also handles unnamed header rows and ensures that the cleaned data is written back to the Excel files.

## Libraries Used

The module relies on the following Python libraries:

1. **os**: This library is used for handling file and directory paths, allowing the module to navigate the file system and perform operations like renaming and deleting files.

2. **pandas**: This is a powerful data processing library that provides the underlying functionality for handling Excel files. The module uses it to read and write Excel files, manipulate dataframes (the main data structure in pandas), and perform operations like dropping columns and renaming headers.

3. **openpyxl**: This is an optional dependency required for pandas to read and write Excel .xlsx files. It's not directly imported in the module but is essential for its functionality.

## Usage

The module can be used as a standalone script or integrated into a larger system. It's compatible with Python 3.6+ and requires the aforementioned libraries. To use the module, you need to specify the source directory for `copy_xlsx_files` and the directory, columns, and sheets for `clean_excel_files`.

## Future Development

The modular structure of Excel Data Cleaner allows for easy extension and customization. Possible future improvements include adding more configuration options, supporting more complex cleaning operations, or extending support to other file formats.

Remember to always keep a backup of your original files when performing operations that alter or delete data.
