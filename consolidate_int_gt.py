import pandas as pd

# Load the datasets from the first and second sheets of the xlsx file
int_df = pd.read_excel('TS_processed/TS_psc_All.xlsx', sheet_name='1_int')
gt_df = pd.read_excel('TS_processed/TS_psc_All.xlsx', sheet_name='2_icgc')

# Check that the 'Referral #' column exists in both dataframes
assert '03_fcid' in gt_df.columns, "'03_fcid' column not found in the first sheet"
assert '03_fcid' in int_df.columns, "'03_fcid' column not found in the second sheet"

# Merge the dataframes on 'Referral #'
merged_df = pd.merge(gt_df, int_df, on='03_fcid', how='outer', suffixes=('_gt', '_int'))

# Check that the merge was successful
assert not merged_df.empty, "The merged dataframe is empty"

# Remove duplicate columns
merged_df = merged_df.loc[:,~merged_df.columns.duplicated()]

# Write the merged dataframe to a new sheet in the same Excel file
with pd.ExcelWriter('TS_processed/TS_psc_All.xlsx', engine='openpyxl', mode='a') as writer: 
    merged_df.to_excel(writer, sheet_name='gt_int')

assert merged_df