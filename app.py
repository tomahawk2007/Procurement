import os
import io
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from data_loader import (
    parse_file_1, 
    parse_file_2, 
    link_datasets, 
    clean_col_name
)

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Spend & Procurement Contract Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -----------------------------------------------------------------------------
# CUSTOM CSS STYLING
# -----------------------------------------------------------------------------
st.markdown("""
<style>
    /* Global layout tweaks */
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2.5rem;
    }
    
    /* Executive Card Styling */
    .metric-card {
        background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 16px 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.04);
        transition: transform 0.15s ease, box-shadow 0.15s ease;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.08);
    }
    .metric-title {
        font-size: 0.82rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #64748b;
        margin-bottom: 6px;
    }
    .metric-value {
        font-size: 1.65rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 4px;
        line-height: 1.2;
    }
    .metric-subtitle {
        font-size: 0.8rem;
        color: #64748b;
    }
    .metric-delta-pos {
        color: #16a34a;
        font-weight: 600;
    }
    .metric-delta-neg {
        color: #dc2626;
        font-weight: 600;
    }
    
    /* Match Summary Banner */
    .match-banner {
        background: #f0fdf4;
        border: 1px solid #bbf7d0;
        border-left: 5px solid #16a34a;
        border-radius: 8px;
        padding: 14px 18px;
        margin-bottom: 1.2rem;
    }
    .match-banner-warn {
        background: #fffbeb;
        border: 1px solid #fde68a;
        border-left: 5px solid #d97706;
        border-radius: 8px;
        padding: 14px 18px;
        margin-bottom: 1.2rem;
    }
    
    /* Badges */
    .badge-matched {
        background-color: #dcfce7;
        color: #15803d;
        padding: 3px 8px;
        border-radius: 12px;
        font-size: 0.75rem;
        font-weight: 600;
    }
    .badge-unmatched {
        background-color: #fee2e2;
        color: #b91c1c;
        padding: 3px 8px;
        border-radius: 12px;
        font-size: 0.75rem;
        font-weight: 600;
    }
    
    /* Upload boxes header */
    .upload-label {
        font-weight: 600;
        color: #1e293b;
        font-size: 0.95rem;
        margin-bottom: 4px;
    }
</style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# HELPER FUNCTIONS
# -----------------------------------------------------------------------------
def format_currency(val: float) -> str:
    """Format float into compact or standard currency."""
    if pd.isna(val) or val == 0:
        return "$0"
    abs_val = abs(val)
    sign = "-" if val < 0 else ""
    if abs_val >= 1_000_000:
        return f"{sign}${abs_val / 1_000_000:.2f}M"
    elif abs_val >= 1_000:
        return f"{sign}${abs_val / 1_000:.1f}K"
    else:
        return f"{sign}${abs_val:.2f}"

def format_number(val: float) -> str:
    """Format integer or volume with thousands separator."""
    if pd.isna(val) or val == 0:
        return "0"
    abs_val = abs(val)
    sign = "-" if val < 0 else ""
    if abs_val >= 1_000_000:
        return f"{sign}{abs_val / 1_000_000:.2f}M"
    elif abs_val >= 1_000:
        return f"{sign}{abs_val / 1_000:.1f}K"
    else:
        return f"{sign}{abs_val:,.0f}"

def to_excel_download_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    """Convert dataframe to downloadable Excel bytes."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

# -----------------------------------------------------------------------------
# APPLICATION HEADER
# -----------------------------------------------------------------------------
col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.title("Spend & Procurement Agreement Analytics")
    st.caption("Automated ingestion, intelligent super-header parsing, multi-tier data linking & price variance intelligence.")
with col_h2:
    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
    if os.path.exists("1.xlsx") and os.path.exists("2.xlsx"):
        use_demo = st.button("📁 Load Default Demo Files (1.xlsx & 2.xlsx)", use_container_width=True, type="secondary")
    else:
        use_demo = False

# -----------------------------------------------------------------------------
# FILE UPLOAD & INGESTION
# -----------------------------------------------------------------------------
with st.expander("📂 File Ingestion & Parser Configuration", expanded=True):
    col_u1, col_u2 = st.columns(2)
    
    with col_u1:
        st.markdown("<div class='upload-label'>📄 File 1: Spend, Volume & PPV Tracker (.xlsx)</div>", unsafe_allow_html=True)
        st.caption("Row 2 super-headers + Row 3 column headers. Auto-detects `Item Code`, combines duplicates, drops `x` separators.")
        uploaded_file_1 = st.file_uploader(
            "Upload Spend Tracker", 
            type=["xlsx", "xls"], 
            key="file1_uploader",
            help="Excel workbook containing monthly volumes, spend, and PPV columns."
        )

    with col_u2:
        st.markdown("<div class='upload-label'>📑 File 2: Procurement Agreements & Pricing (.xlsx)</div>", unsafe_allow_html=True)
        st.caption("Rows 1–12 metadata filters + Row 13 column headers. Auto-detects `Item Name`, supplier BPA & contract prices.")
        uploaded_file_2 = st.file_uploader(
            "Upload Procurement Agreements", 
            type=["xlsx", "xls"], 
            key="file2_uploader",
            help="Blanket purchase agreements with supplier names, contract prices, line status, and validity dates."
        )

# Fallback to local sample files if user pressed button or no files uploaded yet
f1_source = uploaded_file_1
f2_source = uploaded_file_2

if f1_source is None and f2_source is None:
    if os.path.exists("1.xlsx") and os.path.exists("2.xlsx"):
        f1_source = "1.xlsx"
        f2_source = "2.xlsx"
        st.info("💡 Currently viewing preloaded demo files (`1.xlsx` and `2.xlsx`). You can upload your own Excel files above at any time.", icon="ℹ️")

if f1_source is None or f2_source is None:
    st.warning("Please upload both File 1 (Spend Tracker) and File 2 (Procurement Agreements) to begin analysis.")
    st.stop()

# -----------------------------------------------------------------------------
# DATA PARSING
# -----------------------------------------------------------------------------
@st.cache_data(show_spinner="Parsing and structuring Excel files...")
def load_and_parse(f1, f2):
    df1, map1 = parse_file_1(f1)
    df2, map2 = parse_file_2(f2)
    return df1, map1, df2, map2

try:
    df1_raw, col_map1, df2_raw, col_map2 = load_and_parse(f1_source, f2_source)
except Exception as e:
    st.error(f"❌ Error during file parsing: {str(e)}")
    st.exception(e)
    st.stop()

# -----------------------------------------------------------------------------
# SIDEBAR CONTROLS & FILTERS
# -----------------------------------------------------------------------------
st.sidebar.header("⚙️ Linking & Join Controls")

join_type = st.sidebar.radio(
    "Data Join Mode",
    options=["Left Join (All File 1 Items)", "Inner Join (Matched Items Only)"],
    index=0,
    help="Left Join displays all spend items and flags unmatched ones. Inner Join shows only items that have contract agreements in File 2."
)
join_mode = "left" if "Left" in join_type else "inner"

active_agreements_only = st.sidebar.checkbox(
    "Filter File 2 to Active Agreements Only",
    value=False,
    help="When checked, filters File 2 to open lines where expired flag is 'N' and expiration date has not passed before linking."
)

st.sidebar.markdown("---")
st.sidebar.header("🔍 File 1 Filters (Spend & Items)")

# File 1 Filters
# Market
market_col = col_map1.get('market')
selected_markets = []
if market_col and market_col in df1_raw.columns:
    markets_list = sorted([str(x) for x in df1_raw[market_col].dropna().unique() if str(x).strip()])
    selected_markets = st.sidebar.multiselect("Market", options=markets_list, default=[])

# Brand
brand_col = col_map1.get('brand')
selected_brands = []
if brand_col and brand_col in df1_raw.columns:
    brands_list = sorted([str(x) for x in df1_raw[brand_col].dropna().unique() if str(x).strip()])
    selected_brands = st.sidebar.multiselect("Brand", options=brands_list, default=[])

# Category
cat_col = col_map1.get('category')
selected_categories = []
if cat_col and cat_col in df1_raw.columns:
    cats_list = sorted([str(x) for x in df1_raw[cat_col].dropna().unique() if str(x).strip()])
    selected_categories = st.sidebar.multiselect("Category", options=cats_list, default=[])

# Sub Category
subcat_col = col_map1.get('sub_category')
selected_subcategories = []
if subcat_col and subcat_col in df1_raw.columns:
    subcats_list = sorted([str(x) for x in df1_raw[subcat_col].dropna().unique() if str(x).strip()])
    selected_subcategories = st.sidebar.multiselect("Sub Category", options=subcats_list, default=[])

# CM (Contract Manufacturer)
cm_col = col_map1.get('cm')
selected_cms = []
if cm_col and cm_col in df1_raw.columns:
    cms_list = sorted([str(x) for x in df1_raw[cm_col].dropna().unique() if str(x).strip()])
    selected_cms = st.sidebar.multiselect("Contract Manufacturer (CM)", options=cms_list, default=[])

# Status
status_col = col_map1.get('status')
selected_statuses = []
if status_col and status_col in df1_raw.columns:
    statuses_list = sorted([str(x) for x in df1_raw[status_col].dropna().unique() if str(x).strip()])
    selected_statuses = st.sidebar.multiselect("Item Status", options=statuses_list, default=[])

# Free text item search
item_search_query = st.sidebar.text_input("Item Code / Description Search", placeholder="e.g. ITM-1001 or Shampoo").strip()

st.sidebar.markdown("---")
st.sidebar.header("🏢 File 2 Filters (Agreements)")

# Supplier
supplier_col = col_map2.get('supplier')
selected_suppliers = []
if supplier_col and supplier_col in df2_raw.columns:
    suppliers_list = sorted([str(x) for x in df2_raw[supplier_col].dropna().unique() if str(x).strip()])
    selected_suppliers = st.sidebar.multiselect("Supplier", options=suppliers_list, default=[])

# Currency Code
curr_col = col_map2.get('currency_code')
selected_currencies = []
if curr_col and curr_col in df2_raw.columns:
    currencies_list = sorted([str(x) for x in df2_raw[curr_col].dropna().unique() if str(x).strip()])
    selected_currencies = st.sidebar.multiselect("Currency Code", options=currencies_list, default=[])

# Line Status
line_status_col = col_map2.get('line_status')
selected_line_statuses = []
if line_status_col and line_status_col in df2_raw.columns:
    line_statuses_list = sorted([str(x) for x in df2_raw[line_status_col].dropna().unique() if str(x).strip()])
    selected_line_statuses = st.sidebar.multiselect("Line Status", options=line_statuses_list, default=[])

# Expired Blanket Agreement Line Flag
expired_flag_col = col_map2.get('expired_flag')
selected_expired_flags = []
if expired_flag_col and expired_flag_col in df2_raw.columns:
    expired_flags_list = sorted([str(x) for x in df2_raw[expired_flag_col].dropna().unique() if str(x).strip()])
    selected_expired_flags = st.sidebar.multiselect("Expired Blanket Flag", options=expired_flags_list, default=[])

# Agreement Number
agr_num_col = col_map2.get('agreement_number')
selected_agreements = []
if agr_num_col and agr_num_col in df2_raw.columns:
    agrs_list = sorted([str(x) for x in df2_raw[agr_num_col].dropna().unique() if str(x).strip()])
    selected_agreements = st.sidebar.multiselect("Agreement Number", options=agrs_list, default=[])

# -----------------------------------------------------------------------------
# APPLY SIDEBAR FILTERS
# -----------------------------------------------------------------------------
df1_filtered = df1_raw.copy()
if selected_markets and market_col:
    df1_filtered = df1_filtered[df1_filtered[market_col].astype(str).isin(selected_markets)]
if selected_brands and brand_col:
    df1_filtered = df1_filtered[df1_filtered[brand_col].astype(str).isin(selected_brands)]
if selected_categories and cat_col:
    df1_filtered = df1_filtered[df1_filtered[cat_col].astype(str).isin(selected_categories)]
if selected_subcategories and subcat_col:
    df1_filtered = df1_filtered[df1_filtered[subcat_col].astype(str).isin(selected_subcategories)]
if selected_cms and cm_col:
    df1_filtered = df1_filtered[df1_filtered[cm_col].astype(str).isin(selected_cms)]
if selected_statuses and status_col:
    df1_filtered = df1_filtered[df1_filtered[status_col].astype(str).isin(selected_statuses)]
if item_search_query:
    q = item_search_query.lower()
    item_code_val = df1_filtered['Item Code'].astype(str).str.lower()
    desc_c = col_map1.get('item_desc')
    if desc_c and desc_c in df1_filtered.columns:
        item_desc_val = df1_filtered[desc_c].astype(str).str.lower()
        df1_filtered = df1_filtered[item_code_val.str.contains(q) | item_desc_val.str.contains(q)]
    else:
        df1_filtered = df1_filtered[item_code_val.str.contains(q)]

df2_filtered = df2_raw.copy()
if selected_suppliers and supplier_col:
    df2_filtered = df2_filtered[df2_filtered[supplier_col].astype(str).isin(selected_suppliers)]
if selected_currencies and curr_col:
    df2_filtered = df2_filtered[df2_filtered[curr_col].astype(str).isin(selected_currencies)]
if selected_line_statuses and line_status_col:
    df2_filtered = df2_filtered[df2_filtered[line_status_col].astype(str).isin(selected_line_statuses)]
if selected_expired_flags and expired_flag_col:
    df2_filtered = df2_filtered[df2_filtered[expired_flag_col].astype(str).isin(selected_expired_flags)]
if selected_agreements and agr_num_col:
    df2_filtered = df2_filtered[df2_filtered[agr_num_col].astype(str).isin(selected_agreements)]

# -----------------------------------------------------------------------------
# LINK DATASETS & PREVENT DOUBLE COUNTING
# -----------------------------------------------------------------------------
linked_data = link_datasets(
    df1=df1_filtered,
    df2=df2_filtered,
    col_map1=col_map1,
    col_map2=col_map2,
    join_type=join_mode,
    active_only=active_agreements_only
)

item_df = linked_data['item_df']
agreement_df = linked_data['agreement_df']
stats = linked_data['stats']

# -----------------------------------------------------------------------------
# MATCH SUMMARY BANNER
# -----------------------------------------------------------------------------
banner_class = "match-banner" if stats['match_rate_f1'] >= 80 else "match-banner-warn"
st.markdown(f"""
<div class='{banner_class}'>
    <div style='display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;'>
        <div>
            <strong style='font-size: 1.05rem;'>Data Match Status: {stats['matched_items_count']} of {stats['total_items_f1']} Items Matched</strong>
            <span style='margin-left: 8px;' class='{"badge-matched" if stats["match_rate_f1"] >= 80 else "badge-unmatched"}'>{stats['match_rate_f1']:.1f}% Match Rate</span>
            <div style='font-size: 0.85rem; color: #475569; margin-top: 4px;'>
                File 1 Unique Items: <b>{stats['total_items_f1']}</b> &nbsp;|&nbsp; 
                File 2 Unique Items: <b>{stats['total_items_f2']}</b> &nbsp;|&nbsp; 
                Unmatched File 1 Items: <b>{stats['unmatched_items_f1_count']}</b> &nbsp;|&nbsp; 
                Unmatched File 2 Items: <b>{stats['unmatched_items_f2_count']}</b>
            </div>
        </div>
        <div>
            <span style='font-size: 0.82rem; background: #e2e8f0; padding: 4px 10px; border-radius: 6px; font-weight: 500;'>
                Mode: {join_type.split(' ')[0]} Join | Aggregation: Single Item Deduped View
            </span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# KPI SUMMARY CARDS (Strictly deduped to prevent double-counting)
# -----------------------------------------------------------------------------
spend_26_col = col_map1.get('spend_2026')
spend_25_col = col_map1.get('spend_2025')
vol_26_col = col_map1.get('vol_2026')
ppv_col = col_map1.get('ppv')
fx_col = col_map1.get('fx_impact')

tot_spend_26 = item_df[spend_26_col].sum() if spend_26_col and spend_26_col in item_df.columns else 0.0
tot_spend_25 = item_df[spend_25_col].sum() if spend_25_col and spend_25_col in item_df.columns else 0.0
tot_vol_26 = item_df[vol_26_col].sum() if vol_26_col and vol_26_col in item_df.columns else 0.0
tot_ppv = item_df[ppv_col].sum() if ppv_col and ppv_col in item_df.columns else 0.0
tot_fx = item_df[fx_col].sum() if fx_col and fx_col in item_df.columns else 0.0

# Active suppliers count across matched items
active_suppliers = set()
if 'All Suppliers' in item_df.columns:
    for sup_str in item_df[item_df['Is Matched']]['All Suppliers']:
        for s in str(sup_str).split(','):
            s_clean = s.strip()
            if s_clean and s_clean not in ['Unmatched', 'N/A', 'nan', 'None']:
                active_suppliers.add(s_clean)
active_suppliers_count = len(active_suppliers)

# YoY Spend delta
spend_diff = tot_spend_26 - tot_spend_25
spend_pct = ((spend_diff) / tot_spend_25 * 100) if tot_spend_25 > 0 else 0.0
delta_class = "metric-delta-pos" if spend_pct <= 0 else "metric-delta-neg"
delta_sign = "+" if spend_diff > 0 else ""

k1, k2, k3, k4, k5 = st.columns(5)

with k1:
    st.markdown(f"""
    <div class='metric-card'>
        <div class='metric-title'>Matched Items & Rate</div>
        <div class='metric-value'>{stats['matched_items_count']} <span style='font-size: 1.1rem; color: #64748b;'>/ {stats['total_items_f1']}</span></div>
        <div class='metric-subtitle'>Match Rate: <b>{stats['match_rate_f1']:.1f}%</b></div>
    </div>
    """, unsafe_allow_html=True)

with k2:
    st.markdown(f"""
    <div class='metric-card'>
        <div class='metric-title'>Total 2026 Spend (USD)</div>
        <div class='metric-value'>{format_currency(tot_spend_26)}</div>
        <div class='metric-subtitle'>2025: {format_currency(tot_spend_25)} (<span class='{delta_class}'>{delta_sign}{spend_pct:.1f}%</span>)</div>
    </div>
    """, unsafe_allow_html=True)

with k3:
    st.markdown(f"""
    <div class='metric-card'>
        <div class='metric-title'>Total 2026 Volume</div>
        <div class='metric-value'>{format_number(tot_vol_26)}</div>
        <div class='metric-subtitle'>Realized Units Across Selection</div>
    </div>
    """, unsafe_allow_html=True)

with k4:
    ppv_class = "metric-delta-pos" if tot_ppv <= 0 else "metric-delta-neg"
    st.markdown(f"""
    <div class='metric-card'>
        <div class='metric-title'>YTD PPV & Fx Impact</div>
        <div class='metric-value'>{format_currency(tot_ppv)}</div>
        <div class='metric-subtitle'>Fx Impact: {format_currency(tot_fx)}</div>
    </div>
    """, unsafe_allow_html=True)

with k5:
    st.markdown(f"""
    <div class='metric-card'>
        <div class='metric-title'>Active Suppliers</div>
        <div class='metric-value'>{active_suppliers_count}</div>
        <div class='metric-subtitle'>Contracted Across Selection</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# INTERACTIVE CHARTS
# -----------------------------------------------------------------------------
st.subheader("📈 Interactive Analytics & Variance Diagnostics")

tab_charts, tab_tables, tab_unmatched = st.tabs([
    "📊 Analytical Dashboards", 
    "📑 Linked Data Tables & Exports", 
    "⚠️ Unmatched Items Diagnostics"
])

with tab_charts:
    # Row 1: Top 10 Spend Comparison & Supplier Allocation
    c1, c2 = st.columns(2)
    
    with c1:
        st.markdown("##### 🏆 2026 vs 2025 Spend Comparison")
        group_opt = st.radio("Group By", ["Item", "Category", "Brand"], horizontal=True, key="top_spend_group")
        
        if group_opt == "Item":
            group_col = 'Item Code'
            name_col = col_map1.get('item_desc', 'Item Code')
            plot_df = item_df.copy()
            plot_df['Display Label'] = plot_df['Item Code'] + " - " + plot_df[name_col].astype(str).str[:25]
            gb = plot_df.groupby('Display Label')[[spend_26_col, spend_25_col]].sum().reset_index()
            gb = gb.sort_values(by=spend_26_col, ascending=False).head(10)
            x_col = 'Display Label'
        elif group_opt == "Category":
            group_col = col_map1.get('category', 'Category')
            gb = item_df.groupby(group_col)[[spend_26_col, spend_25_col]].sum().reset_index()
            gb = gb.sort_values(by=spend_26_col, ascending=False).head(10)
            x_col = group_col
        else:
            group_col = col_map1.get('brand', 'Brand')
            gb = item_df.groupby(group_col)[[spend_26_col, spend_25_col]].sum().reset_index()
            gb = gb.sort_values(by=spend_26_col, ascending=False).head(10)
            x_col = group_col
            
        fig_spend = go.Figure()
        fig_spend.add_trace(go.Bar(
            x=gb[x_col],
            y=gb[spend_25_col],
            name='2025 Spend ($)',
            marker_color='#94a3b8'
        ))
        fig_spend.add_trace(go.Bar(
            x=gb[x_col],
            y=gb[spend_26_col],
            name='2026 Spend ($)',
            marker_color='#2563eb'
        ))
        fig_spend.update_layout(
            barmode='group',
            xaxis_tickangle=-35,
            margin=dict(l=40, r=20, t=20, b=80),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            yaxis_title="Spend in USD ($)",
            template="plotly_white",
            height=380
        )
        st.plotly_chart(fig_spend, use_container_width=True)

    with c2:
        st.markdown("##### 🏭 Spend & Allocation by Supplier")
        sup_metric = st.radio("Supplier Metric", ["Spend (USD)", "Items Count"], horizontal=True, key="sup_metric_toggle")
        
        # We compute supplier allocation using the matched items dataset
        sup_data = []
        for idx, row in item_df.iterrows():
            sups = [s.strip() for s in str(row['All Suppliers']).split(',') if s.strip() and s.strip() != 'Unmatched']
            if not sups:
                sup_data.append({'Supplier': 'Unmatched / No Contract', 'Spend': row[spend_26_col], 'Item Code': row['Item Code']})
            else:
                for s in sups:
                    sup_data.append({'Supplier': s, 'Spend': row[spend_26_col] / len(sups), 'Item Code': row['Item Code']})
        
        sup_df = pd.DataFrame(sup_data)
        if not sup_df.empty:
            if sup_metric == "Spend (USD)":
                sup_agg = sup_df.groupby('Supplier')['Spend'].sum().reset_index().sort_values(by='Spend', ascending=True).tail(10)
                fig_sup = px.bar(
                    sup_agg,
                    y='Supplier',
                    x='Spend',
                    orientation='h',
                    color='Spend',
                    color_continuous_scale='Blues',
                    labels={'Spend': '2026 Spend (USD)'}
                )
            else:
                sup_agg = sup_df.groupby('Supplier')['Item Code'].nunique().reset_index().sort_values(by='Item Code', ascending=True).tail(10)
                fig_sup = px.bar(
                    sup_agg,
                    y='Supplier',
                    x='Item Code',
                    orientation='h',
                    color='Item Code',
                    color_continuous_scale='Teal',
                    labels={'Item Code': 'Unique Items Count'}
                )
            fig_sup.update_layout(
                margin=dict(l=40, r=20, t=20, b=40),
                coloraxis_showscale=False,
                template="plotly_white",
                height=380
            )
            st.plotly_chart(fig_sup, use_container_width=True)
        else:
            st.info("No supplier data available for the current filter selection.")

    # Row 2: Monthly Trend (Spend & Volume) and Spend Distribution Donut
    c3, c4 = st.columns([3, 2])
    
    with c3:
        st.markdown("##### 📅 2026 Monthly Spend & Volume Run-Rate (Jan–Dec)")
        month_spend_cols = col_map1.get('monthly_spend_2026', {})
        month_vol_cols = col_map1.get('monthly_vol_2026', {})
        
        month_order = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        m_spend_vals = []
        m_vol_vals = []
        m_labels = []
        
        for m in month_order:
            if m in month_spend_cols and month_spend_cols[m] in item_df.columns:
                m_spend_vals.append(item_df[month_spend_cols[m]].sum())
                m_labels.append(m)
            if m in month_vol_cols and month_vol_cols[m] in item_df.columns:
                m_vol_vals.append(item_df[month_vol_cols[m]].sum())
                
        if m_labels:
            fig_trend = make_subplots(specs=[[{"secondary_y": True}]])
            
            fig_trend.add_trace(
                go.Bar(
                    x=m_labels, 
                    y=m_vol_vals, 
                    name="Volume (Units)", 
                    marker_color='rgba(203, 213, 225, 0.7)'
                ),
                secondary_y=True
            )
            
            fig_trend.add_trace(
                go.Scatter(
                    x=m_labels, 
                    y=m_spend_vals, 
                    name="Spend (USD)", 
                    mode="lines+markers", 
                    line=dict(color='#2563eb', width=3),
                    marker=dict(size=7)
                ),
                secondary_y=False
            )
            
            fig_trend.update_xaxes(title_text="Month (2026)")
            fig_trend.update_yaxes(title_text="Spend in USD ($)", secondary_y=False)
            fig_trend.update_yaxes(title_text="Volume (Units)", secondary_y=True, showgrid=False)
            fig_trend.update_layout(
                template="plotly_white",
                margin=dict(l=40, r=40, t=20, b=40),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                height=380
            )
            st.plotly_chart(fig_trend, use_container_width=True)
        else:
            st.info("Monthly spend columns not found or not mapped in File 1.")

    with c4:
        st.markdown("##### 🍩 Spend Distribution Breakdown")
        donut_opt = st.radio("Breakdown Dimension", ["Brand", "Market", "Category"], horizontal=True, key="donut_group")
        dim_col = col_map1.get(donut_opt.lower().replace(' ', '_'), donut_opt)
        
        if dim_col and dim_col in item_df.columns:
            donut_df = item_df.groupby(dim_col)[spend_26_col].sum().reset_index()
            fig_donut = px.pie(
                donut_df,
                names=dim_col,
                values=spend_26_col,
                hole=0.55,
                color_discrete_sequence=px.colors.qualitative.Prism
            )
            fig_donut.update_traces(textposition='inside', textinfo='percent+label')
            fig_donut.update_layout(
                margin=dict(l=20, r=20, t=20, b=40),
                showlegend=False,
                height=380
            )
            st.plotly_chart(fig_donut, use_container_width=True)
        else:
            st.info(f"Dimension '{donut_opt}' is not available in the dataset.")

    # Row 3: Scatter Variance Chart (Agreement Price vs Actual Rate)
    st.markdown("##### 🎯 Agreement Contract Price vs. Realized Actual Rate (Leakage & Variance Radar)")
    st.caption("Spot pricing variances across linked items: points above the diagonal parity line have higher agreement prices; points below indicate realized unit rate exceeds agreed contract terms.")
    
    scatter_df = item_df[item_df['Is Matched'] & (item_df['Avg Agreement Price'] > 0) & (item_df['Actual Rate 2026'] > 0)].copy()
    
    if not scatter_df.empty:
        desc_col = col_map1.get('item_desc', 'Item Code')
        cat_color_col = col_map1.get('category', 'Category')
        
        max_limit = max(scatter_df['Avg Agreement Price'].max(), scatter_df['Actual Rate 2026'].max()) * 1.08
        
        fig_scatter = px.scatter(
            scatter_df,
            x='Actual Rate 2026',
            y='Avg Agreement Price',
            color=cat_color_col if cat_color_col in scatter_df.columns else None,
            size=spend_26_col if spend_26_col in scatter_df.columns else None,
            hover_name='Item Code',
            hover_data={
                desc_col: True,
                'Primary Supplier': True,
                'Actual Rate 2026': ':.2f',
                'Avg Agreement Price': ':.2f',
                'Price Variance ($)': ':.2f',
                'Price Variance (%)': ':.1f',
                'Agreements Count': True
            },
            labels={
                'Actual Rate 2026': 'Realized Actual Rate ($/unit = Spend / Volume)',
                'Avg Agreement Price': 'Agreement Contract Price ($/unit)',
                cat_color_col: 'Category'
            }
        )
        
        # Add 45-degree parity line
        fig_scatter.add_trace(go.Scatter(
            x=[0, max_limit],
            y=[0, max_limit],
            mode='lines',
            name='Parity Line (Agr Price = Actual Rate)',
            line=dict(color='rgba(220, 38, 38, 0.6)', dash='dash', width=2)
        ))
        
        fig_scatter.update_layout(
            template="plotly_white",
            xaxis=dict(range=[0, max_limit]),
            yaxis=dict(range=[0, max_limit]),
            margin=dict(l=40, r=40, t=20, b=40),
            height=450
        )
        st.plotly_chart(fig_scatter, use_container_width=True)
    else:
        st.info("No matched items with both positive Agreement Price and Actual Rate in the current filter selection.")

# -----------------------------------------------------------------------------
# TAB 2: DATA TABLES & EXPORT
# -----------------------------------------------------------------------------
with tab_tables:
    st.markdown("#### 📋 Searchable Linked Datasets")
    table_view = st.radio("Select Table View", [
        "Deduped Item-Level Summary (Spend, Volume & Aggregated Contract Info)",
        "Detailed Agreement-Level Drill-Down (All Individual Contract Lines)"
    ], horizontal=True)
    
    if "Deduped" in table_view:
        st.caption(f"Showing {len(item_df)} items. Spend and volume metrics are strictly unique per item to guarantee no double-counting.")
        
        # Prepare display columns
        display_cols = ['Item Code']
        desc_col = col_map1.get('item_desc')
        if desc_col: display_cols.append(desc_col)
        for col_name in [col_map1.get('market'), col_map1.get('brand'), col_map1.get('category'), col_map1.get('status')]:
            if col_name and col_name in item_df.columns: display_cols.append(col_name)
            
        display_cols.extend([
            spend_25_col, spend_26_col, vol_26_col, 'Actual Rate 2026', 
            'Is Matched', 'Agreements Count', 'Primary Supplier', 'Avg Agreement Price', 'Price Variance ($)'
        ])
        display_cols = [c for c in display_cols if c and c in item_df.columns]
        
        show_df = item_df[display_cols].copy()
        
        # Format columns for display
        format_dict = {}
        if spend_25_col in show_df.columns: format_dict[spend_25_col] = "${:,.2f}"
        if spend_26_col in show_df.columns: format_dict[spend_26_col] = "${:,.2f}"
        if vol_26_col in show_df.columns: format_dict[vol_26_col] = "{:,.0f}"
        if 'Actual Rate 2026' in show_df.columns: format_dict['Actual Rate 2026'] = "${:,.2f}"
        if 'Avg Agreement Price' in show_df.columns: format_dict['Avg Agreement Price'] = "${:,.2f}"
        if 'Price Variance ($)' in show_df.columns: format_dict['Price Variance ($)'] = "${:,.2f}"
        
        st.dataframe(
            show_df.style.format(format_dict, na_rep="N/A"),
            use_container_width=True,
            height=420
        )
        
        col_exp1, col_exp2 = st.columns(2)
        with col_exp1:
            csv_bytes = show_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                "📥 Download Item Summary (.csv)",
                data=csv_bytes,
                file_name="linked_items_summary.csv",
                mime="text/csv",
                use_container_width=True
            )
        with col_exp2:
            xlsx_bytes = to_excel_download_bytes(show_df, "Item_Summary")
            st.download_button(
                "📥 Download Item Summary (.xlsx)",
                data=xlsx_bytes,
                file_name="linked_items_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
    else:
        st.caption(f"Showing {len(agreement_df)} agreement lines across linked records with detailed supplier contract numbers, expiration, and line pricing.")
        
        # Select key columns for agreement view
        agr_cols = ['Item Code']
        desc_col = col_map1.get('item_desc')
        if desc_col: agr_cols.append(desc_col)
        
        for k in ['agreement_number', 'supplier', 'supplier_site', 'currency_code', 'price', 'agreement_uom', 'line_status', 'expired_flag', 'expiration_date']:
            target = col_map2.get(k)
            # Check if prefixed with F2_
            if target and target in agreement_df.columns:
                agr_cols.append(target)
            elif target and f"F2_{target}" in agreement_df.columns:
                agr_cols.append(f"F2_{target}")
                
        agr_cols.extend(['Actual Rate 2026', 'Line Price Variance ($)'])
        agr_cols = [c for c in agr_cols if c in agreement_df.columns]
        
        show_agr = agreement_df[agr_cols].copy()
        
        st.dataframe(
            show_agr,
            use_container_width=True,
            height=420
        )
        
        col_exp1, col_exp2 = st.columns(2)
        with col_exp1:
            csv_bytes = show_agr.to_csv(index=False).encode('utf-8')
            st.download_button(
                "📥 Download Agreement Drill-Down (.csv)",
                data=csv_bytes,
                file_name="agreement_drilldown.csv",
                mime="text/csv",
                use_container_width=True
            )
        with col_exp2:
            xlsx_bytes = to_excel_download_bytes(show_agr, "Agreements")
            st.download_button(
                "📥 Download Agreement Drill-Down (.xlsx)",
                data=xlsx_bytes,
                file_name="agreement_drilldown.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

# -----------------------------------------------------------------------------
# TAB 3: UNMATCHED ITEMS DIAGNOSTICS
# -----------------------------------------------------------------------------
with tab_unmatched:
    st.markdown("#### ⚠️ Master Data Alignment & Unmatched Diagnostics")
    st.caption("Identifies governance gaps between commercial spend tracking (File 1) and procurement contractual coverage (File 2).")
    
    unmatched_f1_df = linked_data.get('unmatched_f1_df', pd.DataFrame())
    unmatched_f2_df = linked_data.get('unmatched_f2_df', pd.DataFrame())
    
    col_d1, col_d2 = st.columns(2)
    
    with col_d1:
        st.markdown("##### 🚨 File 1 Items Missing in File 2 (Spend Without Contract)")
        if not unmatched_f1_df.empty:
            f1_miss_cols = ['Item Code']
            if col_map1.get('item_desc'): f1_miss_cols.append(col_map1.get('item_desc'))
            if col_map1.get('brand'): f1_miss_cols.append(col_map1.get('brand'))
            if col_map1.get('category'): f1_miss_cols.append(col_map1.get('category'))
            if spend_26_col: f1_miss_cols.append(spend_26_col)
            f1_miss_cols = [c for c in f1_miss_cols if c in unmatched_f1_df.columns]
            
            st.dataframe(unmatched_f1_df[f1_miss_cols], use_container_width=True, height=280)
            st.warning(f"Total: **{len(unmatched_f1_df)}** items with spend have no corresponding procurement agreement.")
            
            c_csv = unmatched_f1_df[f1_miss_cols].to_csv(index=False).encode('utf-8')
            st.download_button("📥 Export Unmatched File 1 Items (.csv)", data=c_csv, file_name="unmatched_spend_items.csv", mime="text/csv", use_container_width=True)
        else:
            st.success("🎉 All File 1 items have matched procurement agreements in File 2!")
        
    with col_d2:
        st.markdown("##### 📁 File 2 Items Missing in File 1 (Contracts Without Realized Spend)")
        if not unmatched_f2_df.empty:
            f2_miss_cols = ['Item Name']
            if col_map2.get('item_description'): f2_miss_cols.append(col_map2.get('item_description'))
            if col_map2.get('supplier'): f2_miss_cols.append(col_map2.get('supplier'))
            if col_map2.get('price'): f2_miss_cols.append(col_map2.get('price'))
            if col_map2.get('agreement_number'): f2_miss_cols.append(col_map2.get('agreement_number'))
            f2_miss_cols = [c for c in f2_miss_cols if c in unmatched_f2_df.columns]
            
            st.dataframe(unmatched_f2_df[f2_miss_cols], use_container_width=True, height=280)
            st.info(f"Total: **{len(unmatched_f2_df)}** items in contracts have zero recorded spend in File 1.")
            
            c2_csv = unmatched_f2_df[f2_miss_cols].to_csv(index=False).encode('utf-8')
            st.download_button("📥 Export Unmatched File 2 Items (.csv)", data=c2_csv, file_name="unmatched_contract_items.csv", mime="text/csv", use_container_width=True)
        else:
            st.success("🎉 All File 2 contract items have corresponding spend in File 1!")

# -----------------------------------------------------------------------------
# FOOTER
# -----------------------------------------------------------------------------
st.markdown("---")
st.markdown("<div style='text-align: center; color: #94a3b8; font-size: 0.8rem;'>Spend & Procurement Agreement Analytics Engine • Built with Streamlit, Pandas & Plotly</div>", unsafe_allow_html=True)
