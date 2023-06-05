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
    'MH Intake': {
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
        'Intake # 1': '15.1_int_s1_date',
        'Intake # 2': '15.2_int_s2_date',
        'Intake # 3': '15.3_int_s3_date',
        'Re-intake': '15.4_rint_date',
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
        'Intake # 1': '15.1_int_s1_date',
        'Intake # 2': '15.2_int_s2_date',
        'Intake # 3': '15.3_int_s3_date',
        'Re-intake': '15.4_rint_date',
        'Notes': '16_note',
    },
    'Counseling': { 
        'Unnamed: 0': 'Unnamed: 0',
        '#': '00_sn',
        'System ID': 'System ID',
        'Referral #': '01_rn',
        'FC ID #': '02_fcid',
        'Beneficiary name': '03_bebeficiary_name',
        'Gender': '04_gender',
        'Age': '05_age',
        'Nationality': '06_natly',
        'Counseling ': '07_ic/gc',
        'PT': '08_is_pt',
        'Group name': '09_gn',
        'Counseling session 1': '10.1_s1_date',
        'Counseling session 2': '10.2_s2_date',
        'Counseling session 3': '10.3_s3_date',
        'Counseling session 4': '10.4_s4_date',
        'Counseling session 5': '10.5_s5_date',
        'Counseling session 6': '10.6_s6_date',
        'Counseling session 7': '10.7_s7_date',
        'Counseling session 8': '10.8_s8_date',
        'Counseling session 9': '10.9_s9_date',
        'Counseling session 10': '10.10_s10_date',
        'Total # of counseling sessions': '11_tns',
        'Notes': '12_note'

    },
    'Follow-up Assessment': {
        'Unnamed: 0': 'Unnamed: 0', 
        '#': '00_sn', 
        'System ID': 'System ID', 
        'Referral #': '01_rn', 
        'FC ID #': '02_fcid', 
        'Beneficiary name': '03_beneficiary_name', 
        'Gender': '04_gender', 
        'Age': '05_age', 
        'Nationality': '06_natly', 
        '1M check-in': '07_1m_chin_date', 
        '3M FUA': '08.1_3m_fua_date', 
        '6M FUA': '08.2_6m_fua_date', 
        '12M FUA': '08.3_12m_fua_date', 
        'Status': '09_stus', 
        'Closure reason': '10_closure_reason',
        'Notes': 'note'

    },
    'TRW': {

    },
    'TD': {

    },
    'Creative Workshop': {

    },
    # add more as needed
}