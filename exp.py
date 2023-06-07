def decrypt_and_copy_xlsx_file(file, dst_dir, password=None):
    """Decrypts a password-protected Excel file (if a password is provided) and copies it to the destination directory."""

    os.makedirs(dst_dir, exist_ok=True)

    fname = os.path.basename(file)
    if fname.endswith('.xlsx'):
        logger.info('Found excel file: %s', fname)

        # If a password has been provided, decrypt the Excel file
        if password is not None:
            with open(file, "rb") as f:
                crypto = msoffcrypto.OfficeFile(f)
                crypto.load_key(password=password)
                decrypted_file = f"{file}_decrypted.xlsx"
                with open(decrypted_file, "wb") as df:
                    crypto.decrypt(df)
        else:
            decrypted_file = file

        # Read the (decrypted) Excel file
        xls = pd.ExcelFile(decrypted_file)
        with pd.ExcelWriter(os.path.join(dst_dir, fname)) as writer:
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        logger.info('Copied file to: %s', os.path.join(dst_dir, fname))

        # If a password has been provided, delete the decrypted file after copying it
        if password is not None:
            os.remove(decrypted_file)
