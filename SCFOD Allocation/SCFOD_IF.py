import pandas as pd
from datetime import datetime

file_path_Main = "C:\\Users\\pradsush\\Documents\\IF AUDIT\\IF Audit Work_Allocation\\Develope thing.xlsx"
file_path_Ref = "C:\\Users\\pradsush\\Documents\\IF AUDIT\\IF Audit Work_Allocation\\SCFOD-consolidation.xlsx"
final_file_path = "C:\\Users\\pradsush\\Documents\\IF AUDIT\\IF Audit Work_Allocation\\Final File.xlsx"

# Read
df = pd.read_excel(file_path_Main, sheet_name="Sheet1")
df_EPI = pd.read_excel(file_path_Main, sheet_name="EPI DATA")
df_SCFOD = pd.read_excel(file_path_Ref, sheet_name='Sheet1')

# Get minimum prices
df_nonzero = df[df['normalized_price'] != 0].copy()
step_1 = df_nonzero.groupby('asin')['normalized_price'].transform('min')
all_min = df_nonzero[df_nonzero['normalized_price'] == step_1].copy()

# Getting Referce col
all_min['Reference'] = all_min['asin'].astype(str) + all_min['product_url'].astype(str)

# connecting to SCFOD, lookup
unmatched = all_min[~all_min['Reference'].isin(df_SCFOD['R1'])].copy()

# connecting to epi getting the look up values only
df_EPI['EPI_Reference'] = df_EPI['Asin'].astype(str) + df_EPI['MappedUrl'].astype(str)
matched_epi = unmatched[unmatched['Reference'].isin(df_EPI['EPI_Reference'])].copy()

# ipq collection
ipq_dict = df_EPI.set_index('EPI_Reference')['IPQ'].to_dict()
matched_epi.loc[:, 'IPQ'] = matched_epi['Reference'].map(ipq_dict)

# getting mapped date from EPI
def convert_utc_date(date_string):
    try:
        if pd.isna(date_string):
            return None
        # Parse the UTC date string and return as datetime object
        return datetime.strptime(date_string, '%a %b %d %H:%M:%S UTC %Y').date()
    except:
        return None

# Create date mapping using LastMappedDate instead of LastMappedDateEpoch
date_mapping = df_EPI.set_index('EPI_Reference')['LastMappedDate'].to_dict()

# Convert dates using the new function
converted_dates = matched_epi['Reference'].map(date_mapping).apply(convert_utc_date)
matched_epi['Nor mapped date'] = pd.to_datetime(converted_dates)
matched_epi['Audit date'] = datetime.now().date()

# First get all records with score >= 0.7
high_score_records = matched_epi[matched_epi['score'] >= 0.7]

if len(high_score_records) >= 100:
    # If we have 100 or more records with score >= 0.7, take all of themhem
    final_records = high_score_records
else:
    # If we have less than 100 records with score >= 0.7
    # Calculate how many more records we need
    additional_records_needed = 100 - len(high_score_records)

    # Get the remaining records (where score < 0.7) and sort by GMS in ascending order
    remaining_records = matched_epi[matched_epi['score'] < 0.7].sort_values('gms_rank', ascending=True)

    # Take the required number of additional records
    additional_records = remaining_records.head(additional_records_needed)

    # Combine high score records with additional records
    final_records = pd.concat([high_score_records, additional_records])

# Write to Excel
with pd.ExcelWriter(final_file_path, engine='openpyxl', date_format='mm/dd/yyyy') as writer:
    final_records.to_excel(writer, index=False)

# Print summary
print(f"Number of records with score >= 0.7: {len(high_score_records)}")
print(f"Number of additional records from GMS sorting: {len(final_records) - len(high_score_records)}")
print(f"Total number of final records: {len(final_records)}")








