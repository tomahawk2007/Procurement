import re
import pandas as pd
import numpy as np
from typing import Tuple, Dict, Any, Optional

def clean_col_name(col: Any) -> str:
    """Clean and standardize column name string."""
    if pd.isna(col):
        return ""
    val = str(col).strip()
    # Normalize internal multiple whitespace
    val = re.sub(r'\s+', ' ', val)
    return val

def is_separator_column(col_name: str) -> bool:
    """Check if column name represents an empty/separator column like 'x' or 'X'."""
    clean = clean_col_name(col_name).lower()
    if clean in {'x', 'xx', 'xxx', 'sep', 'separator', ''}:
        return True
    if re.match(r'^x(\.\d+)?$', clean):
        return True
    if clean.startswith('unnamed:'):
        return True
    return False

def to_clean_numeric(series: pd.Series) -> pd.Series:
    """Safely convert strings with currency symbols, commas, or parentheses to numeric."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors='coerce').fillna(0.0)
    
    cleaned = series.astype(str).str.strip()
    # Replace negative accounting format (123.45) with -123.45
    cleaned = cleaned.str.replace(r'^\((.*)\)$', r'-\1', regex=True)
    # Remove currency signs, commas, spaces
    cleaned = cleaned.str.replace(r'[\$,€,£,¥]', '', regex=True)
    cleaned = cleaned.str.replace(',', '', regex=False)
    # Replace '-' alone with 0
    cleaned = cleaned.replace(['-', 'nan', 'None', '', 'N/A', 'null'], '0')
    return pd.to_numeric(cleaned, errors='coerce').fillna(0.0)

def parse_file_1(file_path_or_buffer) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Parse File 1: Spend, Volume & PPV Tracker.
    - Row 2 contains grouped section super-headers.
    - Row 3 contains actual column headers: Market, Brand, Category, Sub Category, CM, Source,
      Item Code, Item Description, UOM, Updated Category, Updated Subcategory, Status, monthly columns,
      and summary total columns.
    - Automatically detects the row containing 'Item Code'.
    - Combines super-headers with monthly/duplicate sub-headers.
    - Drops empty separator columns labeled 'x' or 'X'.
    """
    # Read raw data without header to inspect rows
    raw_df = pd.read_excel(file_path_or_buffer, header=None)
    
    # 1. Search for header row containing 'Item Code'
    header_row_idx = None
    item_code_col_idx = None
    
    for row_idx in range(min(15, len(raw_df))):
        row_vals = [clean_col_name(v).lower() for v in raw_df.iloc[row_idx].values]
        for col_idx, val in enumerate(row_vals):
            if val == 'item code' or val.startswith('item code'):
                header_row_idx = row_idx
                item_code_col_idx = col_idx
                break
        if header_row_idx is not None:
            break
            
    if header_row_idx is None:
        raise ValueError("Could not automatically locate 'Item Code' header row in File 1.")

    # 2. Extract super-header row if present (row immediately above header row)
    super_headers = []
    if header_row_idx > 0:
        raw_super = raw_df.iloc[header_row_idx - 1].tolist()
        # Forward fill super headers to handle merged cells across grouped sections
        current_super = ""
        for val in raw_super:
            c_val = clean_col_name(val)
            if c_val and not is_separator_column(c_val) and not c_val.lower().startswith('unnamed:'):
                current_super = c_val
            super_headers.append(current_super)
    else:
        super_headers = [""] * len(raw_df.columns)
        
    sub_headers = [clean_col_name(v) for v in raw_df.iloc[header_row_idx].tolist()]
    
    # Known core metadata columns that should stay simple and pristine
    core_metadata_cols = {
        'market', 'brand', 'category', 'sub category', 'subcategory', 'cm', 'source',
        'item code', 'item description', 'uom', 'updated category', 'updated subcategory', 'status'
    }
    
    # 3. Build unique combined column headers
    combined_headers = []
    seen_counts = {}
    
    for col_idx, (sup, sub) in enumerate(zip(super_headers, sub_headers)):
        # Check if separator
        if is_separator_column(sub) and not sup:
            combined_headers.append(f"__DROP_SEP_{col_idx}__")
            continue
            
        sub_lower = sub.lower()
        if sub_lower in core_metadata_cols:
            final_col = sub
        elif is_separator_column(sub):
            combined_headers.append(f"__DROP_SEP_{col_idx}__")
            continue
        elif sup and sub:
            # If sub already starts with super, don't repeat
            if sub.lower().startswith(sup.lower()):
                final_col = sub
            elif sup.lower() in sub_lower:
                final_col = sub
            elif sub.lower().startswith("total"):
                # Clear total summary columns keep their direct name
                final_col = sub
            else:
                final_col = f"{sup} - {sub}"
        elif sub:
            final_col = sub
        elif sup:
            final_col = f"{sup} (Col {col_idx})"
        else:
            combined_headers.append(f"__DROP_SEP_{col_idx}__")
            continue
            
        # Ensure complete uniqueness
        count = seen_counts.get(final_col, 0)
        seen_counts[final_col] = count + 1
        if count > 0:
            final_col = f"{final_col} ({count})"
        combined_headers.append(final_col)
        
    # 4. Construct DataFrame from rows below header_row_idx
    data_df = raw_df.iloc[header_row_idx + 1:].copy().reset_index(drop=True)
    data_df.columns = combined_headers
    
    # 5. Drop separator columns
    cols_to_keep = [c for c in data_df.columns if not c.startswith("__DROP_SEP_") and not is_separator_column(c)]
    data_df = data_df[cols_to_keep]
    
    # Drop rows where Item Code is NaN or empty
    # Find actual item code column name
    item_col = next((c for c in data_df.columns if clean_col_name(c).lower() == 'item code'), None)
    if item_col:
        data_df = data_df[data_df[item_col].notna()]
        data_df = data_df[data_df[item_col].astype(str).str.strip() != ""]
        data_df['Item Code'] = data_df[item_col].astype(str).str.strip()
    
    # 6. Detect and map key analytical columns
    col_map = identify_file_1_metrics(data_df.columns)
    
    # Clean numeric fields
    for key in ['spend_2026', 'spend_2025', 'spend_2026_lc', 'vol_2026', 'vol_2025', 'ppv', 'fx_impact']:
        target_col = col_map.get(key)
        if target_col and target_col in data_df.columns:
            data_df[target_col] = to_clean_numeric(data_df[target_col])
            
    # Calculate Actual Rate 2026 = Spend 2026 / Volume 2026 (if not directly provided)
    spend_26_col = col_map.get('spend_2026')
    vol_26_col = col_map.get('vol_2026')
    if spend_26_col and vol_26_col:
        data_df['Actual Rate 2026'] = np.where(
            data_df[vol_26_col] > 0,
            data_df[spend_26_col] / data_df[vol_26_col],
            0.0
        )
    else:
        data_df['Actual Rate 2026'] = 0.0

    return data_df, col_map

def identify_file_1_metrics(columns) -> Dict[str, Any]:
    """Map dynamic column names to standard keys."""
    col_map = {
        'item_code': 'Item Code',
        'item_desc': None,
        'market': None,
        'brand': None,
        'category': None,
        'sub_category': None,
        'cm': None,
        'status': None,
        'spend_2026': None,
        'spend_2025': None,
        'spend_2026_lc': None,
        'vol_2026': None,
        'vol_2025': None,
        'ppv': None,
        'fx_impact': None,
        'monthly_spend_2026': {},
        'monthly_vol_2026': {}
    }
    
    months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
    
    for c in columns:
        cl = c.lower()
        has_month = any(m in cl for m in months)
        
        # Item Description
        if ('item desc' in cl or 'description' in cl) and not col_map['item_desc']:
            col_map['item_desc'] = c
        elif cl == 'market':
            col_map['market'] = c
        elif cl == 'brand':
            col_map['brand'] = c
        elif cl in ['category', 'updated category'] and not col_map['category']:
            col_map['category'] = c
        elif cl in ['sub category', 'subcategory', 'updated subcategory'] and not col_map['sub_category']:
            col_map['sub_category'] = c
        elif cl in ['cm', 'contract manufacturer']:
            col_map['cm'] = c
        elif cl == 'status':
            col_map['status'] = c
            
        # Spend 2026 USD (Totals ONLY, excluding month columns)
        elif not has_month and 'total' in cl and '2026' in cl and ('usd' in cl or '$' in cl):
            col_map['spend_2026'] = c
        # Spend 2025 USD (Totals ONLY, excluding month columns)
        elif not has_month and 'total' in cl and '2025' in cl and ('usd' in cl or '$' in cl):
            col_map['spend_2025'] = c
        # Spend 2026 LC
        elif not has_month and 'total' in cl and '2026' in cl and 'lc' in cl:
            col_map['spend_2026_lc'] = c
        # Volume 2026 Total
        elif not has_month and 'total' in cl and ('volume' in cl or 'vol' in cl) and ('2026' in cl or 'total volume' in cl):
            # Check if this is 2025
            if '2025' not in cl and not col_map['vol_2026']:
                col_map['vol_2026'] = c
        # Volume 2025 Total
        elif not has_month and 'total' in cl and ('volume' in cl or 'vol' in cl) and '2025' in cl:
            col_map['vol_2025'] = c
        # PPV
        elif 'ppv' in cl and not col_map['ppv']:
            col_map['ppv'] = c
        # Fx Impact
        elif ('fx impact' in cl or 'fx' in cl) and not col_map['fx_impact']:
            col_map['fx_impact'] = c

    # Fallbacks for Spend and Volume totals if not yet matched
    for c in columns:
        cl = c.lower()
        has_month = any(m in cl for m in months)
        if not col_map['spend_2026'] and not has_month and 'spend' in cl and '2026' in cl:
            col_map['spend_2026'] = c
        if not col_map['spend_2025'] and not has_month and 'spend' in cl and '2025' in cl:
            col_map['spend_2025'] = c
        if not col_map['vol_2026'] and not has_month and 'vol' in cl and '2026' in cl:
            col_map['vol_2026'] = c
        if not col_map['vol_2026'] and not has_month and cl in ['total volume', 'actual volume - 2026 - total volume']:
            col_map['vol_2026'] = c

    # Detect monthly series for 2026 (Jan through Dec in chronological order)
    month_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    for m in month_names:
        m_lower = m.lower()
        for c in columns:
            cl = c.lower()
            if m_lower in cl and ('2026' in cl or '26' in cl) and 'total' not in cl:
                if 'spend' in cl and ('usd' in cl or '$' in cl):
                    if m not in col_map['monthly_spend_2026']:
                        col_map['monthly_spend_2026'][m] = c
                elif 'volume' in cl or 'vol' in cl:
                    if m not in col_map['monthly_vol_2026']:
                        col_map['monthly_vol_2026'][m] = c

    return col_map

def parse_file_2(file_path_or_buffer) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Parse File 2: Procurement Agreements & Supplier Pricing.
    - Rows 1-12 contain report metadata.
    - Row 13 contains actual column headers:
      Procurement Bu Name, Requisitioning Bu Name, Agreement Number, Supplier, Supplier Site Name,
      Currency Code, Item Name, Item Description, Price, Agreement UOM, Category Name,
      Line Status, Expired Blanket Agreement Line Flag, Start Date, End Date, Expiration Date.
    - Automatically detects the row containing 'Item Name'.
    """
    raw_df = pd.read_excel(file_path_or_buffer, header=None)
    
    header_row_idx = None
    for row_idx in range(min(25, len(raw_df))):
        row_vals = [clean_col_name(v).lower() for v in raw_df.iloc[row_idx].values]
        for val in row_vals:
            if val == 'item name' or val.startswith('item name'):
                header_row_idx = row_idx
                break
        if header_row_idx is not None:
            break
            
    if header_row_idx is None:
        raise ValueError("Could not automatically locate 'Item Name' header row in File 2.")
        
    headers = [clean_col_name(v) for v in raw_df.iloc[header_row_idx].tolist()]
    
    # Handle duplicates in headers if any
    unique_headers = []
    seen = {}
    for h in headers:
        if not h:
            h = "Unnamed"
        count = seen.get(h, 0)
        seen[h] = count + 1
        unique_headers.append(h if count == 0 else f"{h}_{count}")
        
    df = raw_df.iloc[header_row_idx + 1:].copy().reset_index(drop=True)
    df.columns = unique_headers
    
    # Find Item Name column
    item_col = next((c for c in df.columns if clean_col_name(c).lower() == 'item name'), None)
    if not item_col:
        raise ValueError("Header found but 'Item Name' column could not be mapped.")
        
    df = df[df[item_col].notna()]
    df = df[df[item_col].astype(str).str.strip() != ""]
    df['Item Name'] = df[item_col].astype(str).str.strip()
    
    # Standardize column names
    col_map = {}
    for c in df.columns:
        cl = clean_col_name(c).lower()
        if 'procurement bu' in cl: col_map['procurement_bu'] = c
        elif 'requisitioning bu' in cl: col_map['requisitioning_bu'] = c
        elif 'agreement number' in cl or 'agreement #' in cl: col_map['agreement_number'] = c
        elif cl == 'supplier' or 'supplier name' in cl: col_map['supplier'] = c
        elif 'supplier site' in cl: col_map['supplier_site'] = c
        elif 'currency' in cl: col_map['currency_code'] = c
        elif cl == 'item name': col_map['item_name'] = c
        elif 'item desc' in cl or (cl.startswith('item') and 'desc' in cl): col_map['item_description'] = c
        elif cl in ['price', 'unit price']: col_map['price'] = c
        elif 'agreement uom' in cl or cl == 'uom': col_map['agreement_uom'] = c
        elif 'category name' in cl: col_map['category_name'] = c
        elif 'line status' in cl: col_map['line_status'] = c
        elif 'expired blanket' in cl or 'expired flag' in cl: col_map['expired_flag'] = c
        elif 'start date' in cl: col_map['start_date'] = c
        elif 'end date' in cl: col_map['end_date'] = c
        elif 'expiration date' in cl: col_map['expiration_date'] = c
        
    # Clean Price to float
    price_col = col_map.get('price')
    if price_col and price_col in df.columns:
        df[price_col] = to_clean_numeric(df[price_col])
    else:
        df['Price'] = 0.0
        col_map['price'] = 'Price'
        
    # Convert dates safely
    for date_key in ['start_date', 'end_date', 'expiration_date']:
        d_col = col_map.get(date_key)
        if d_col and d_col in df.columns:
            df[d_col] = pd.to_datetime(df[d_col], errors='coerce')

    return df, col_map

def link_datasets(df1: pd.DataFrame, 
                  df2: pd.DataFrame, 
                  col_map1: Dict[str, Any], 
                  col_map2: Dict[str, Any],
                  join_type: str = "left",
                  active_only: bool = False) -> Dict[str, Any]:
    """
    Links File 1 and File 2 on Item Code == Item Name (trimmed, uppercase).
    
    Prevents double-counting:
    - Creates an aggregated Item-Level dataset (one row per File 1 item) for KPIs & charts.
    - Creates a detailed Agreement-Level dataset (1-to-many join) for the drilldown table & price variances.
    - Calculates matching statistics.
    """
    # 1. Normalize Join Keys
    item_code_col = col_map1.get('item_code', 'Item Code')
    item_name_col = col_map2.get('item_name', 'Item Name')
    
    df1_clean = df1.copy()
    df2_clean = df2.copy()
    
    df1_clean['_join_key'] = df1_clean[item_code_col].astype(str).str.strip().str.upper()
    df2_clean['_join_key'] = df2_clean[item_name_col].astype(str).str.strip().str.upper()
    
    # Set of unique items
    items_f1 = set(df1_clean['_join_key'].dropna().unique()) - {"", "NAN"}
    items_f2 = set(df2_clean['_join_key'].dropna().unique()) - {"", "NAN"}
    matched_items = items_f1.intersection(items_f2)
    unmatched_f1 = items_f1 - items_f2
    unmatched_f2 = items_f2 - items_f1
    
    match_rate_f1 = (len(matched_items) / len(items_f1) * 100) if len(items_f1) > 0 else 0.0
    
    # 2. Filter File 2 if active_only toggle is set
    df2_filtered = df2_clean.copy()
    status_col = col_map2.get('line_status')
    expired_col = col_map2.get('expired_flag')
    exp_date_col = col_map2.get('expiration_date')
    
    if active_only:
        conditions = []
        if status_col and status_col in df2_filtered.columns:
            conditions.append(df2_filtered[status_col].astype(str).str.upper().isin(['OPEN', 'ACTIVE']))
        if expired_col and expired_col in df2_filtered.columns:
            conditions.append(df2_filtered[expired_col].astype(str).str.upper().isin(['N', 'NO', 'FALSE', '0']))
        if exp_date_col and exp_date_col in df2_filtered.columns:
            now = pd.Timestamp.now()
            conditions.append(df2_filtered[exp_date_col].isna() | (df2_filtered[exp_date_col] >= now))
            
        if conditions:
            combined_cond = conditions[0]
            for c in conditions[1:]:
                combined_cond = combined_cond & c
            df2_filtered = df2_filtered[combined_cond]

    # 3. Aggregate File 2 at the Item Level (to avoid double-counting in KPIs and Spend/Volume charts)
    supplier_col = col_map2.get('supplier', 'Supplier')
    price_col = col_map2.get('price', 'Price')
    agr_num_col = col_map2.get('agreement_number', 'Agreement Number')
    curr_col = col_map2.get('currency_code', 'Currency Code')
    
    agg_dict = {}
    if agr_num_col in df2_filtered.columns:
        agg_dict['agreement_count'] = (agr_num_col, 'count')
    else:
        agg_dict['agreement_count'] = ('_join_key', 'count')
        
    def safe_join_unique(s):
        unique_vals = [str(x).strip() for x in s.dropna().unique() if str(x).strip() and str(x).strip().upper() != 'NAN']
        return ", ".join(unique_vals) if unique_vals else "N/A"
        
    def first_val(s):
        unique_vals = [str(x).strip() for x in s.dropna().unique() if str(x).strip() and str(x).strip().upper() != 'NAN']
        return unique_vals[0] if unique_vals else "N/A"
        
    agg_funcs = {
        'Agreements Count': (agr_num_col if agr_num_col in df2_filtered.columns else '_join_key', 'count'),
        'Primary Supplier': (supplier_col if supplier_col in df2_filtered.columns else '_join_key', first_val),
        'All Suppliers': (supplier_col if supplier_col in df2_filtered.columns else '_join_key', safe_join_unique),
        'Avg Agreement Price': (price_col if price_col in df2_filtered.columns else '_join_key', 'mean'),
        'Min Agreement Price': (price_col if price_col in df2_filtered.columns else '_join_key', 'min'),
        'Max Agreement Price': (price_col if price_col in df2_filtered.columns else '_join_key', 'max'),
        'Agreement Numbers': (agr_num_col if agr_num_col in df2_filtered.columns else '_join_key', safe_join_unique),
        'Currencies': (curr_col if curr_col in df2_filtered.columns else '_join_key', safe_join_unique),
    }
    
    if status_col and status_col in df2_filtered.columns:
        agg_funcs['Line Statuses'] = (status_col, safe_join_unique)
        
    if not df2_filtered.empty:
        df2_agg = df2_filtered.groupby('_join_key').agg(**agg_funcs).reset_index()
    else:
        # Empty df2
        df2_agg = pd.DataFrame(columns=['_join_key', 'Agreements Count', 'Primary Supplier', 'All Suppliers', 
                                       'Avg Agreement Price', 'Min Agreement Price', 'Max Agreement Price', 
                                       'Agreement Numbers', 'Currencies'])

    # 4. Create Deduped Item-Level Joined Dataset
    item_df = pd.merge(df1_clean, df2_agg, on='_join_key', how='left')
    item_df['Is Matched'] = item_df['_join_key'].isin(matched_items)
    item_df['Agreements Count'] = item_df['Agreements Count'].fillna(0).astype(int)
    item_df['Primary Supplier'] = item_df['Primary Supplier'].fillna('Unmatched / No Agreement')
    item_df['All Suppliers'] = item_df['All Suppliers'].fillna('Unmatched')
    item_df['Avg Agreement Price'] = item_df['Avg Agreement Price'].fillna(0.0)
    
    # Calculate Price Variance: Agreement Price vs Actual Rate 2026
    actual_rate = item_df['Actual Rate 2026']
    item_df['Price Variance ($)'] = np.where(
        item_df['Avg Agreement Price'] > 0,
        item_df['Avg Agreement Price'] - actual_rate,
        0.0
    )
    item_df['Price Variance (%)'] = np.where(
        (item_df['Avg Agreement Price'] > 0) & (actual_rate > 0),
        ((item_df['Avg Agreement Price'] - actual_rate) / actual_rate) * 100.0,
        0.0
    )
    
    # Filter by join_type for item_df
    if join_type.lower() == 'inner':
        item_df = item_df[item_df['Is Matched']].copy()
        
    # 5. Create Detailed Agreement-Level Dataset (1-to-many join)
    # Prefix or rename overlapping columns to keep traceability clean
    df2_detail = df2_filtered.copy()
    overlap_cols = set(df1_clean.columns).intersection(set(df2_detail.columns)) - {'_join_key'}
    rename_f2 = {c: f"F2_{c}" for c in overlap_cols}
    df2_detail = df2_detail.rename(columns=rename_f2)
    
    if join_type.lower() == 'inner':
        agreement_df = pd.merge(df1_clean[df1_clean['_join_key'].isin(matched_items)], 
                                df2_detail, on='_join_key', how='inner')
    else:
        agreement_df = pd.merge(df1_clean, df2_detail, on='_join_key', how='left')
        
    agreement_df['Is Matched'] = agreement_df['_join_key'].isin(matched_items)
    
    # Compute detailed line price variance
    agr_price_col = rename_f2.get(price_col, price_col)
    if agr_price_col in agreement_df.columns:
        agreement_df['Agreement Line Price'] = to_clean_numeric(agreement_df[agr_price_col])
        agreement_df['Line Price Variance ($)'] = np.where(
            agreement_df['Agreement Line Price'] > 0,
            agreement_df['Agreement Line Price'] - agreement_df['Actual Rate 2026'],
            0.0
        )
    else:
        agreement_df['Agreement Line Price'] = 0.0
        agreement_df['Line Price Variance ($)'] = 0.0

    stats = {
        'total_items_f1': len(items_f1),
        'total_items_f2': len(items_f2),
        'matched_items_count': len(matched_items),
        'unmatched_items_f1_count': len(unmatched_f1),
        'unmatched_items_f2_count': len(unmatched_f2),
        'match_rate_f1': match_rate_f1,
        'unmatched_items_f1_sample': list(unmatched_f1)[:10],
        'unmatched_items_f2_sample': list(unmatched_f2)[:10],
    }

    # Extract all unmatched records
    unmatched_f1_df = df1_clean[df1_clean['_join_key'].isin(unmatched_f1)].copy()
    unmatched_f2_df = df2_clean[df2_clean['_join_key'].isin(unmatched_f2)].copy()

    return {
        'item_df': item_df,
        'agreement_df': agreement_df,
        'unmatched_f1_df': unmatched_f1_df,
        'unmatched_f2_df': unmatched_f2_df,
        'stats': stats,
        'col_map1': col_map1,
        'col_map2': col_map2
    }
