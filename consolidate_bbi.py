import pandas as pd

# Load the datasets from the first and second sheets of the xlsx file
screening_df = pd.read_excel('TS_processed/TS_psc_All.xlsx', sheet_name=0, usecols=range(7))
intake_df = pd.read_excel('TS_processed/TS_psc_All.xlsx', sheet_name=1, usecols=range(8))

# Check that the 'Referral #' column exists in both dataframes
assert 'Referral #' in screening_df.columns, "'Referral #' column not found in the first sheet"
assert 'Referral #' in intake_df.columns, "'Referral #' column not found in the second sheet"

# Merge the dataframes on 'Referral #'
merged_df = pd.merge(screening_df, intake_df, on='Referral #', how='outer', suffixes=('_screening', '_intake'))

# Check that the merge was successful
assert not merged_df.empty, "The merged dataframe is empty"

# Remove duplicate columns
merged_df = merged_df.loc[:,~merged_df.columns.duplicated()]

# Write the merged dataframe to a new sheet in the same Excel file
with pd.ExcelWriter('TS_processed/TS_psc_All.xlsx', engine='openpyxl', mode='a') as writer: 
    merged_df.to_excel(writer, sheet_name='bbi')

assert merged_df