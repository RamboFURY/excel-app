import pandas as pd

def process_file(path):
    df = pd.read_excel(path)
    df.columns = df.columns.str.strip()

    folio_col = "Folio Number"
    dpid_col = "DP Id-Client Id-Account Number"

    # Clean columns
    df[folio_col] = df[folio_col].astype(str).str.strip().replace(["", "nan", "None"], pd.NA)
    df[dpid_col] = df[dpid_col].astype(str).str.strip().replace(["", "nan", "None"], pd.NA)

    # Unique datasets
    df_folio_unique = df[df[folio_col].notna()].drop_duplicates(subset=[folio_col])
    df_dpid_unique = df[df[dpid_col].notna()].drop_duplicates(subset=[dpid_col])

    # Summary
    summary_df = pd.DataFrame({
        "Metric": ["Total Records", "Unique Folio Number", "Unique DP ID"],
        "Value": [len(df), df[folio_col].nunique(), df[dpid_col].nunique()]
    })

    return df, df_folio_unique, df_dpid_unique, summary_df