def clean_excel_files(dir_path, cols_to_drop, sheets_to_drop, json_file):
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
                    with pd.ExcelWriter(temp_file_path, engine='xlsxwriter') as writer:
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
