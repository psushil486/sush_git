import pandas as pd
import re

def extract_info_from_description(Description):
    try:
        results = {
            'sku_url': re.search(r'SKU URL: (https?://[^\s]+)', Description),
            'marketplace': re.search(r'Marketplace: ([^\n]+)', Description),
            'asin': re.search(r'ASIN: ([^\n]+)', Description),
            'price': re.search(r'Price: ([0-9.]+)', Description),
            'competitor': re.search(r'Competitor: ([^\n]+)', Description),
            'GL': re.search(r'GL \(if known\): ([^\n]+)', Description)
        }

        results = {k: v.group(1).strip() if v else '' for k, v in results.items()}
        return pd.Series(results)

    except Exception as e:
        print(f"Error processing description: {e}")
        return pd.Series({k: '' for k in ['sku_url', 'marketplace', 'asin', 'price', 'competitor', 'GL']})


df = pd.read_csv('INPUT SIM FILE.csv')

# Create normalized date-time columns
df['nor_CreateDate_time'] = pd.to_datetime(df['CreateDate']).dt.strftime('%m/%d/%Y %H:%M:%S')
df['nor_ResolvedDate_time'] = pd.to_datetime(df['ResolvedDate']).dt.strftime('%m/%d/%Y %H:%M:%S')

extracted_info = df.apply(lambda row: extract_info_from_description(row['Description']), axis=1)

final_df = pd.concat([
    df[['ShortId', 'IssueUrl', 'Title', 'Description', 'Status',
        'AssignedFolderLabel', 'RequesterIdentity', 'CreateDate',
        'ResolvedDate', 'AssigneeIdentity', 'RootCauses','ResolvedByIdentity']],
    extracted_info[['sku_url', 'marketplace', 'asin', 'price', 'competitor', 'GL']],
    df[['nor_CreateDate_time', 'nor_ResolvedDate_time']]  # Moved to the end
], axis=1)

final_df.to_csv('OUTPUT SIM FILE.csv', index=False)