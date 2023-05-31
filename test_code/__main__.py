import os
import xlsx_pro as xp


def main():
    # Define your source directory and destination directory
    src_dir = r'C:\Users\myagh\OneDrive - SİVİL HUKUK DERNEĞi\Desktop\psc_src_files_test'

    # Define your sheet names dictionary and column names dictionary
    sheet_names = {
        'Screening': '0_scr',
        'MH Intake ': '1_int',
        'MH Intake': '1_int',
        'Counseling': '2_cnsl',
        'Follow-up Assessment': '3_fua',
        'TRW': '4_trw',
        'TD': '5_td',
        'Creative Workshop': '6_cws'
        # add more as needed
    }
    col_names = {
        'Screening': {
            'Unnamed: 0': 'Unnamed: 0',
            '#': '00_sn',
            'System ID': 'System ID',
            'Referral #': '01_rn',
            'Beneficiary name': '02_beneficiary_name',
            'Gender': '03_gender',
            'Age': '04_age',
            'Nationality': '05_natly',
            'Screening date': '06_scr_date',
            'Source of referral': '07_sor',
            "Referring organization \n(I/NGO, CBO, int'l agency)\nor (Outreach)": '08_rorg',
            'Notes': '09_note'
            # add more as needed"
            
        },
        'MH Intake ': {
            'Unnamed: 0': 'Unnamed: 0',
            '#': '00_sn',
            'System ID': 'System ID',
            'Referral #': '01_rn',
            'FC ID #': '02_fcid',
            'Beneficiary name': '03_beneficiary_name',
            'Gender': '04_gender',
            'Age': '05_age',
            'Nationality': '06_natly',
            'Missing\nfamily': '07_is_mf',
            'Former\ndetainee': '08_is_fd',
            'Primary\nTS': '09_is_pts',
            'Secondary\nTS': '10_is_sts',
            'WoV': '11_is_wov',
            'ST/GBV': '12_is_st/gbv',
            'LGBTI': '13_is_lgbti',
            'Other': '14_is_other',
            'Intake # 1': '15_int_fs_date',
            'Intake # 2': '16_int_ss_date',
            'Intake # 3': '17_int_ts_date',
            'Re-intake': '18_rint_date',
            'Notes': '19_note',
            # add more as needed
        },
        'MH Intake ': {
            'Unnamed: 0': 'Unnamed: 0',
            '#': '00_sn',
            'System ID': 'System ID',
            'Referral #': '01_rn',
            'FC ID #': '02_fcid',
            'Beneficiary name': '03_beneficiary_name',
            'Gender': '04_gender',
            'Age': '05_age',
            'Nationality': '06_natly',
            'Missing\nfamily': '07_is_mf',
            'Former\ndetainee': '08_is_fd',
            'Primary\nTS': '09_is_pts',
            'Secondary\nTS': '10_is_sts',
            'WoV': '11_is_wov',
            'ST/GBV': '12_is_st/gbv',
            'LGBTI': '13_is_lgbti',
            'Other': '14_is_other',
            'Intake # 1': '15_int_fs_date',
            'Intake # 2': '16_int_ss_date',
            'Intake # 3': '17_int_ts_date',
            'Re-intake': '18_rint_date',
            'Notes': '19_note',
        },
        'Counseling': {
            
        },
        'Follow-up Assessment': {
            
        },
        'TRW': {
            
        },
        'TD': {
            
        },
        'Creative Workshop': {
            
        },
        # add more as needed
    }

    # Clean and copy files
    xp.clean_and_copy_files(src_dir, 'TS_processed', sheet_names, col_names)

    # Merge all files in 'TS_processed' directory
    xp.merge_excel_files(os.path.join(src_dir, 'TS_processed'), 'consolidated.xlsx')


if __name__ == '__main__':
    main()
