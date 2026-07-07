
import pandas as pd
import glob
import os
import numpy as np
from datetime import datetime
import ast
import pick_listings as pl
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
#import RMINFData_No_of_search_module as sh

# -------------------------------------------------------

def extract_title(scraped_attrs_str):
    if pd.isna(scraped_attrs_str):
        return None

    scraped_attrs_str = str(scraped_attrs_str).strip()
    if not scraped_attrs_str or scraped_attrs_str.lower() == 'nan':
        return None

    try:
        data = ast.literal_eval(scraped_attrs_str)
        if not isinstance(data, dict):
            return None

        data_lower = {k.lower(): v for k, v in data.items()}
        title_fields = [
            'titleval',
            'title',
            'a2e_greedy_schema_org_name',
            'a2e_greedy_title',
            'item_name1'
        ]

        for field in title_fields:
            if field in data_lower and data_lower[field]:
                return data_lower[field]

        return None

    except Exception:
        return None


    # ------------------------------------------------------


def calculate_riskybit_scores(df):
    print("Starting riskybit score calculation...")

    df = df.copy()
    df['title'] = df['scraped_attrs'].apply(extract_title)

    IF_df = df[df['mapping_status'] == 'IF']
    INF_df = df[df['mapping_status'] == 'RMINF']

    if IF_df.empty or INF_df.empty:
        print("⚠ No IF or no RMINF records. Skipping RiskyBit.")
        return pd.DataFrame(columns=['product_url', 'riskybit_score'])

    comps = set(IF_df['competitor_name']).intersection(INF_df['competitor_name'])
    print(f"Found {len(comps)} common competitors between IF and RMINF")

    results = []

    for comp in comps:
        IF_c = IF_df[IF_df['competitor_name'] == comp].dropna(subset=['title'])
        INF_c = INF_df[INF_df['competitor_name'] == comp].dropna(subset=['title'])

        if IF_c.empty or INF_c.empty:
            continue

        vectorizer = TfidfVectorizer(stop_words='english')
        tfidf_IF = vectorizer.fit_transform(IF_c['title'])

        if not vectorizer.vocabulary_:
            continue

        tfidf_INF = vectorizer.transform(INF_c['title'])
        sims = cosine_similarity(tfidf_INF, tfidf_IF).max(axis=1)

        temp = INF_c[['product_url']].copy()
        temp['riskybit_score'] = sims
        results.append(temp)

    if results:
        scored = pd.concat(results, ignore_index=True)
        print(f"Calculated riskybit scores for {len(scored)} RMINF items")
        return scored

    return pd.DataFrame(columns=['product_url', 'riskybit_score'])


# ---------------------------------------------------------


def rm_sampling(audit_op, mapping_in, output_files):

    # ================= AUDIT OUTPUT =================
    all_op_files = glob.glob(audit_op + "/*.xlsx")
    audit_output = [pd.read_excel(f) for f in all_op_files]
    audit_output_base = pd.concat(audit_output, ignore_index=True)

    audit_output_base['Mapped By'] = audit_output_base['Mapped By'].str.replace('@amazon.com', '')

    MTD = audit_output_base[
        ['MKPL', 'Competitor', 'URL', 'Norm URL Length (250)', 'Mapped Status',
         'Reason code', 'Mapped By', 'Mapped Date', 'Manager ID', 'Auditor ID',
         'New Mapping Status', 'Reason Code', 'Audit Date', 'Audit Week',
         'Audit Status (Error/No Error)', 'ASIN', 'Search Technique',
         'No of search performed', 'Final Search Technique', 'Audit Type']
    ]

    MTD.set_index('MKPL', inplace=True)
    MTD.to_excel("MTD_NWTC_RMINF_Consolidated.xlsx")

    df_summary_mapper = MTD.pivot_table(
        index='Mapped By',
        columns='Audit Status (Error/No Error)',
        values='URL',
        aggfunc='count',
        fill_value=0
    ).reset_index()

    df_summary_mapper['#Audit'] = df_summary_mapper.get('Error', 0) + df_summary_mapper.get('No Error', 0)

    date_str = datetime.now().strftime("%Y-%m-%d")
    df_summary_mapper.to_excel(f"{output_files}/Mapper level audit count - MTD - {date_str}.xlsx")

    # ================= MAPPING FILES =================
    all_files = glob.glob(mapping_in + "/*.txt")

    number_of_files = len(all_files)
    print('Number of mapping reports = ', number_of_files)
    
    cont = input("Do you want to continue (Y/N):")
    if cont.upper() == 'N':
        return None

    mapping_base = [
        pd.read_csv(f, sep='\t', encoding='latin1', on_bad_lines='skip')
        for f in all_files
    ]
    df_base = pd.concat(mapping_base, ignore_index=True)
    df_base.drop_duplicates('product_url', inplace=True)
    df_base["Number of Search"] = df_base["mapper_history"].fillna("").str.count(r"s\?k")
    df_base_inf = df_base[df_base['mapping_status']=='RMINF']

    # CRITICAL FIX
    #if 'mapping_status_code' not in df_mapping.columns and 'mapping_status_code' in df_base.columns:
        #df_mapping['mapping_status_code'] = df_base['mapping_status_code']

    # ================= RISKYBIT =================
    risky_scores = calculate_riskybit_scores(df_base)
    df_base_inf = df_base_inf.merge(
        risky_scores, on='product_url', how='left'
    )
    df_base_inf['riskybit_score'] = df_base_inf['riskybit_score'].fillna(-1000)

    # ================= EANRM =================
    df_lookup = pd.read_csv('EANRM.csv')
    df_lookup.rename(columns={'product_url': 'URL'}, inplace=True)

    df_result = df_base_inf.merge(df_lookup[['URL']], how='left',
                                 left_on='product_url', right_on='URL')

    df_result['process'] = np.where(df_result['URL'].notna(), 'EANRM', 'Non EAN RM')

    # ================= PRIORITY =================
    df_result['priority'] = np.where(df_result['riskybit_score'] >= 0.7, 1, np.where(df_result['riskybit_score'] >= 0.5, 3, None))
    
    df_result.loc[(df_result['priority'].isna()) & (df_result['process'] == 'EANRM'), 'priority'] = 2
    
    df_result.loc[df_result['priority'].isna(), 'priority'] = 4

    # ================= OUT OF STOCK =================
    #if 'mapping_status_code' in df_result.columns:
     #   df_result['out_of_stock_flag'] = df_result['mapping_status_code'].isin(
      #      ['USER_ACTION_COMPETITOR_OUT_OF_STOCK', 'COMPETITOR_OUT_OF_STOCK']
       # ).map({True: 'Out of Stock', False: 'In Stock'})
    #else:
     #   df_result['out_of_stock_flag'] = 'Unknown'

    df_result = df_result.sort_values(by=['riskybit_score','priority'], ascending=[False, True],na_position='last')
    
    #df_result.drop_duplicates('product_url', inplace=True)

    #-------------- Backlog and pick samples--------------#

    final_df = pl.pick_listings(df_result,df_summary_mapper,date_str,output_files)
    
    # ================= FINAL FEED =================
    auditfeed = final_df[
        ['marketplace_id', 'activity_day', 'product_url', 'competitor_name',
         'mapper_id', 'event_time', 'mapping_status', 'mapping_status_code',
         'mapper_history', 'suggested_asin', 'scraped_attrs',
         'mapping_duration_seconds', 'process', 'Number of Search','Need to audit','riskybit_score','priority',]
    ]

    auditfeed.to_excel(f"{output_files}/Audit Feed - {date_str}.xlsx", index=False)
    auditfeed.to_csv(f"Audit Feed {date_str}.csv", index=False)

    print("✓ Final Audit feed generated successfully")
    return auditfeed

# ----------------------------------------------------






