import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import random
from datetime import datetime, timedelta

def create_sample_files():
    # -------------------------------------------------------------
    # 1. GENERATE FILE 1: Spend, Volume & PPV Tracker (1.xlsx)
    # -------------------------------------------------------------
    wb1 = openpyxl.Workbook()
    ws1 = wb1.active
    ws1.title = "Spend & Volume Tracker"
    
    # Items seed data
    items_meta = [
        ("ITM-1001", "Hydrating Shampoo 500ml", "US", "LuxeCare", "Hair Care", "Shampoo", "ContractPack Inc", "Internal", "EA", "Active"),
        ("ITM-1002", "Clarifying Conditioner 500ml", "US", "LuxeCare", "Hair Care", "Conditioner", "ContractPack Inc", "Internal", "EA", "Active"),
        ("ITM-1003", "Botanical Body Wash 750ml", "UK", "PureGlow", "Personal Wash", "Body Wash", "EuroBottlers Ltd", "External", "EA", "Active"),
        ("ITM-1004", "Antibacterial Hand Soap 250ml", "UK", "PureGlow", "Personal Wash", "Liquid Soap", "EuroBottlers Ltd", "External", "EA", "Active"),
        ("ITM-1005", "Moisturizing Day Cream 50ml", "Germany", "DermaNova", "Skin Care", "Face Cream", "AlpsPharma GmbH", "Internal", "EA", "Active"),
        ("ITM-1006", "Night Repair Serum 30ml", "Germany", "DermaNova", "Skin Care", "Serum", "AlpsPharma GmbH", "Internal", "EA", "Active"),
        ("ITM-1007", "Mineral Sunscreen SPF50 100ml", "France", "SoleilCare", "Sun Care", "Lotion", "CosmoProd SAS", "External", "EA", "Active"),
        ("ITM-1008", "Aloe After-Sun Gel 200ml", "France", "SoleilCare", "Sun Care", "Gel", "CosmoProd SAS", "External", "EA", "Phase-out"),
        ("ITM-1009", "Ultra Concentrated Detergent 2L", "UAE", "CleanHome", "Home Care", "Laundry Detergent", "GulfPack FZE", "Internal", "EA", "Active"),
        ("ITM-1010", "Fabric Softener Spring Mist 1.5L", "UAE", "CleanHome", "Home Care", "Softener", "GulfPack FZE", "Internal", "EA", "Active"),
        ("ITM-1011", "Enzyme Stain Remover 500ml", "KSA", "CleanHome", "Home Care", "Specialty Cleaner", "RiyadhChem LLC", "External", "EA", "Active"),
        ("ITM-1012", "Citrus Dishwashing Liquid 1L", "KSA", "CleanHome", "Home Care", "Dish Care", "RiyadhChem LLC", "External", "EA", "Active"),
        ("ITM-1013", "Bamboo Toothbrush Soft 2pk", "US", "EcoDent", "Oral Care", "Manual Toothbrush", "SinoDental Corp", "External", "PK", "Active"),
        ("ITM-1014", "Whitening Fluoride Toothpaste 120g", "US", "EcoDent", "Oral Care", "Toothpaste", "ContractPack Inc", "Internal", "TU", "Active"),
        ("ITM-1015", "Alcohol-Free Mouthwash 500ml", "UK", "EcoDent", "Oral Care", "Mouthwash", "EuroBottlers Ltd", "External", "EA", "Active"),
        ("ITM-1016", "Cardboard Shipping Carton Medium", "US", "LuxeCare", "Packaging", "Corrugated Boxes", "GlobalBox Inc", "External", "EA", "Active"),
        ("ITM-1017", "HDPE White Bottle 500ml w/ Pump", "Germany", "DermaNova", "Packaging", "Bottles & Closures", "Plastix AG", "External", "EA", "Active"),
        ("ITM-1018", "Laminated Barrier Tube 100ml", "France", "SoleilCare", "Packaging", "Tubes", "Tubex France", "External", "EA", "Active"),
        ("ITM-1019", "Surfactant SLES 70% Industrial Drum", "UAE", "CleanHome", "Raw Materials", "Surfactants", "PetroChem Middle East", "Internal", "KG", "Active"),
        ("ITM-1020", "Organic Argan Oil Pure Grade", "France", "DermaNova", "Raw Materials", "Essential Oils", "Atlas Naturals SARL", "External", "KG", "Active"),
        ("ITM-1021", "Active Anti-Dandruff Zinc Pyrithione", "US", "LuxeCare", "Raw Materials", "Active Ingredients", "BioSynthetics USA", "External", "KG", "Active"),
        ("ITM-1022", "Fragrance Floral Elegance Premium", "France", "LuxeCare", "Raw Materials", "Fragrances", "Grasse Essences", "External", "KG", "Active"),
        ("ITM-1023", "Biodegradable Wet Wipes 80ct", "UK", "PureGlow", "Personal Wash", "Wipes", "EuroBottlers Ltd", "External", "PK", "Discontinued"),
        ("ITM-1024", "Natural Deodorant Stick 75g", "Germany", "DermaNova", "Personal Wash", "Deodorant", "AlpsPharma GmbH", "Internal", "EA", "Active"),
        ("ITM-1025", "Recycled PET Clear Bottle 250ml", "UK", "PureGlow", "Packaging", "Bottles & Closures", "EcoPlast UK", "External", "EA", "Active"),
    ]
    
    # Row 1: Optional Title / blank
    ws1.cell(row=1, column=1, value="ANNUAL SPEND, VOLUME & PPV TRACKER 2025-2026")
    ws1.cell(row=1, column=1).font = Font(name="Calibri", size=14, bold=True, color="1F497D")
    
    # Months for 2025 and 2026
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    
    # Define columns and their super-headers
    # Structure:
    # 1. Base Meta: Market, Brand, Category, Sub Category, CM, Source, Item Code, Item Description, UOM, Updated Category, Updated Subcategory, Status
    # 2. Actual Volumes 2025: Jan-25 .. Dec-25, TOTAL 2025 Volume
    # 3. Separator 'x'
    # 4. Actual Volume - 2026: Jan-26 .. Dec-26, TOTAL Volume
    # 5. Separator 'X'
    # 6. Actual 2025 Spend in USD: Jan-25 .. Dec-25, Total 2025 Spend $
    # 7. Separator 'x'
    # 8. Actual 2026 Spend in LC: Jan-26 .. Dec-26, Total 2026 Spend in LC
    # 9. Separator 'x'
    # 10. Actual 2026 Spend in USD: Jan-26 .. Dec-26, Total 2026 Spend in USD
    # 11. Separator 'x'
    # 12. PPV vs 2025 Avg Rate in 2025 Constant Fx
    # 13. Fx Impact
    # 14. PPV vs 2025 Avg Rate Actual Fx
    
    cols_def = [
        # (Super Header, Column Header, Type)
        ("", "Market", "meta"),
        ("", "Brand", "meta"),
        ("", "Category", "meta"),
        ("", "Sub Category", "meta"),
        ("", "CM", "meta"),
        ("", "Source", "meta"),
        ("", "Item Code", "meta"),
        ("", "Item Description", "meta"),
        ("", "UOM", "meta"),
        ("", "Updated Category", "meta"),
        ("", "Updated Subcategory", "meta"),
        ("", "Status", "meta"),
    ]
    
    # Volumes 2025
    for m in months:
        cols_def.append(("Actual Volumes 2025", f"{m}-25", "vol_25_m"))
    cols_def.append(("Actual Volumes 2025", "TOTAL 2025 Volume", "vol_25_tot"))
    
    # Sep
    cols_def.append(("", "x", "sep"))
    
    # Volumes 2026
    for m in months:
        cols_def.append(("Actual Volume - 2026", f"{m}-26", "vol_26_m"))
    cols_def.append(("Actual Volume - 2026", "TOTAL Volume", "vol_26_tot"))
    
    # Sep
    cols_def.append(("", "X", "sep"))
    
    # Spend 2025 USD
    for m in months:
        cols_def.append(("Actual 2025 Spend in USD", f"{m}-25", "spend_25_m"))
    cols_def.append(("Actual 2025 Spend in USD", "Total 2025 Spend $", "spend_25_tot"))
    
    # Sep
    cols_def.append(("", "x", "sep"))
    
    # Spend 2026 LC
    for m in months:
        cols_def.append(("Actual 2026 Spend in LC", f"{m}-26", "spend_26_lc_m"))
    cols_def.append(("Actual 2026 Spend in LC", "Total 2026 Spend in LC", "spend_26_lc_tot"))
    
    # Sep
    cols_def.append(("", "x", "sep"))
    
    # Spend 2026 USD
    for m in months:
        cols_def.append(("Actual 2026 Spend in USD", f"{m}-26", "spend_26_m"))
    cols_def.append(("Actual 2026 Spend in USD", "Total 2026 Spend in USD", "spend_26_tot"))
    
    # Sep
    cols_def.append(("", "X", "sep"))
    
    # PPV & Fx
    cols_def.append(("PPV vs 2025 Avg Rate in 2025 Constant Fx", "PPV vs 2025 Avg Rate in 2025 Constant Fx", "ppv_const"))
    cols_def.append(("Fx Impact", "Fx Impact", "fx_impact"))
    cols_def.append(("PPV vs 2025 Avg Rate Actual Fx", "PPV vs 2025 Avg Rate Actual Fx", "ppv_actual"))
    
    # Populate Row 2 (Super-headers) and Row 3 (Headers)
    super_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    super_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    hdr_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    hdr_font = Font(name="Calibri", size=10, bold=True, color="1F497D")
    
    # We want merged section super headers in Row 2
    current_sec = None
    sec_start_col = None
    
    for c_idx, (sup, sub, _) in enumerate(cols_def, start=1):
        ws1.cell(row=2, column=c_idx, value=sup if sup else None)
        ws1.cell(row=2, column=c_idx).fill = super_fill
        ws1.cell(row=2, column=c_idx).font = super_font
        ws1.cell(row=2, column=c_idx).alignment = Alignment(horizontal="center", vertical="center")
        
        ws1.cell(row=3, column=c_idx, value=sub)
        ws1.cell(row=3, column=c_idx).fill = hdr_fill
        ws1.cell(row=3, column=c_idx).font = hdr_font
        ws1.cell(row=3, column=c_idx).alignment = Alignment(horizontal="center", vertical="center")
        
    # Merge super header blocks where appropriate
    idx = 1
    while idx <= len(cols_def):
        sup_text = cols_def[idx-1][0]
        if sup_text:
            start_idx = idx
            while idx <= len(cols_def) and cols_def[idx-1][0] == sup_text:
                idx += 1
            end_idx = idx - 1
            if end_idx > start_idx:
                ws1.merge_cells(start_row=2, start_column=start_idx, end_row=2, end_column=end_idx)
        else:
            idx += 1

    # Populate Data rows starting at row 4
    random.seed(42)
    start_row = 4
    for r_offset, meta in enumerate(items_meta):
        row_num = start_row + r_offset
        item_code, desc, market, brand, cat, subcat, cm, source, uom, status = meta
        
        # Base unit prices and volume scales
        base_unit_price = round(random.uniform(2.5, 45.0), 2)
        base_monthly_vol = random.randint(1500, 25000)
        
        vol_25_list = [int(base_monthly_vol * random.uniform(0.85, 1.15)) for _ in range(12)]
        vol_26_list = [int(v * random.uniform(1.05, 1.25)) for v in vol_25_list]
        
        spend_25_list = [round(v * base_unit_price * random.uniform(0.98, 1.02), 2) for v in vol_25_list]
        price_26 = base_unit_price * random.uniform(0.92, 1.12)
        spend_26_usd_list = [round(v * price_26 * random.uniform(0.98, 1.02), 2) for v in vol_26_list]
        
        # Local currency exchange factor
        fx_rate = 1.0 if market == "US" else (0.92 if market in ["Germany", "France"] else (0.78 if market == "UK" else 3.67))
        spend_26_lc_list = [round(s * fx_rate, 2) for s in spend_26_usd_list]
        
        tot_vol_25 = sum(vol_25_list)
        tot_vol_26 = sum(vol_26_list)
        tot_spend_25 = sum(spend_25_list)
        tot_spend_26_usd = sum(spend_26_usd_list)
        tot_spend_26_lc = sum(spend_26_lc_list)
        
        ppv_const = round((price_26 - base_unit_price) * tot_vol_26 * 0.85, 2)
        fx_impact = round((price_26 - base_unit_price) * tot_vol_26 * 0.15, 2)
        ppv_actual = round(ppv_const + fx_impact, 2)
        
        vol_25_idx = 0
        vol_26_idx = 0
        spend_25_idx = 0
        spend_26_idx = 0
        spend_26_lc_idx = 0
        
        for c_idx, (_, sub, c_type) in enumerate(cols_def, start=1):
            val = None
            if c_type == "meta":
                if sub == "Market": val = market
                elif sub == "Brand": val = brand
                elif sub == "Category": val = cat
                elif sub == "Sub Category": val = subcat
                elif sub == "CM": val = cm
                elif sub == "Source": val = source
                elif sub == "Item Code": val = item_code
                elif sub == "Item Description": val = desc
                elif sub == "UOM": val = uom
                elif sub == "Updated Category": val = cat
                elif sub == "Updated Subcategory": val = subcat
                elif sub == "Status": val = status
            elif c_type == "vol_25_m":
                val = vol_25_list[vol_25_idx]
                vol_25_idx += 1
            elif c_type == "vol_25_tot":
                val = tot_vol_25
            elif c_type == "vol_26_m":
                val = vol_26_list[vol_26_idx]
                vol_26_idx += 1
            elif c_type == "vol_26_tot":
                val = tot_vol_26
            elif c_type == "spend_25_m":
                val = spend_25_list[spend_25_idx]
                spend_25_idx += 1
            elif c_type == "spend_25_tot":
                val = tot_spend_25
            elif c_type == "spend_26_lc_m":
                val = spend_26_lc_list[spend_26_lc_idx]
                spend_26_lc_idx += 1
            elif c_type == "spend_26_lc_tot":
                val = tot_spend_26_lc
            elif c_type == "spend_26_m":
                val = spend_26_usd_list[spend_26_idx]
                spend_26_idx += 1
            elif c_type == "spend_26_tot":
                val = tot_spend_26_usd
            elif c_type == "ppv_const":
                val = ppv_const
            elif c_type == "fx_impact":
                val = fx_impact
            elif c_type == "ppv_actual":
                val = ppv_actual
            elif c_type == "sep":
                val = None
                
            cell = ws1.cell(row=row_num, column=c_idx, value=val)
            if c_type in ["spend_25_tot", "spend_26_tot", "ppv_const", "fx_impact", "ppv_actual"]:
                cell.number_format = '$#,##0.00'
            elif c_type in ["vol_25_tot", "vol_26_tot"]:
                cell.number_format = '#,##0'

    # Save File 1
    wb1.save("1.xlsx")
    print("Successfully generated 1.xlsx")
    
    # -------------------------------------------------------------
    # 2. GENERATE FILE 2: Procurement Agreements (2.xlsx)
    # -------------------------------------------------------------
    wb2 = openpyxl.Workbook()
    ws2 = wb2.active
    ws2.title = "Supplier Agreements"
    
    # Rows 1-12: Report filter metadata
    metadata_rows = [
        ("REPORT NAME:", "Procurement Blanket Purchase Agreements & Active Pricing Master"),
        ("REPORT ID:", "BPA-EXTRACT-2026-Q4"),
        ("GENERATION TIMESTAMP:", "2026-11-20 08:30:15 UTC"),
        ("BUSINESS GROUP:", "Global Supply Chain & Procurement Services"),
        ("OPERATING UNITS:", "All Operating Units (US, EMEA, APAC, MEA)"),
        ("FILTER: LINE STATUS:", "OPEN, PENDING, EXPIRED"),
        ("FILTER: AGREEMENT TYPES:", "BPA, CPA, STANDARD CONTRACT"),
        ("EFFECTIVE PERIOD:", "2024-01-01 to 2027-12-31"),
        ("CONFIDENTIALITY:", "STRICTLY INTERNAL PROCUREMENT DATA"),
        ("EXTRACTED BY USER:", "Procurement.Analyst@enterprise.com"),
        ("", ""),
        ("----------------------------------------", "TABLE DATA COMMENCES AT ROW 13 BELOW")
    ]
    
    for r_idx, (k, v) in enumerate(metadata_rows, start=1):
        c1 = ws2.cell(row=r_idx, column=1, value=k)
        c2 = ws2.cell(row=r_idx, column=2, value=v)
        c1.font = Font(name="Calibri", size=9, bold=True, color="595959")
        c2.font = Font(name="Calibri", size=9, italic=True, color="595959")
        
    # Row 13: Actual Column Headers
    headers_f2 = [
        "Procurement Bu Name",
        "Requisitioning Bu Name",
        "Agreement Number",
        "Supplier",
        "Supplier Site Name",
        "Currency Code",
        "Item Name",
        "Item Description",
        "Price",
        "Agreement UOM",
        "Category Name",
        "Line Status",
        "Expired Blanket Agreement Line Flag",
        "Start Date",
        "End Date",
        "Expiration Date"
    ]
    
    hdr2_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    hdr2_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    
    for c_idx, h in enumerate(headers_f2, start=1):
        cell = ws2.cell(row=13, column=c_idx, value=h)
        cell.fill = hdr2_fill
        cell.font = hdr2_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    # Populate agreements
    # Include multiple agreements for some items, single agreements for others,
    # and unmatched items on both sides!
    suppliers_pool = [
        ("Apex Packaging Solutions Inc", "NEW YORK-HQ", "USD"),
        ("ChemCore International Ltd", "LONDON-PLANT", "GBP"),
        ("Bavaria Ingredients GmbH", "MUNICH-OPS", "EUR"),
        ("Arabian Chemical Ventures", "DUBAI-JAFZA", "AED"),
        ("Pacific BioSynthetics Pte", "SINGAPORE-MAIN", "USD"),
        ("Nordic Polymer Industries", "HELSINKI-DIST", "EUR"),
        ("Crown Closures & Bottles", "CHICAGO-MFG", "USD")
    ]
    
    # We will generate ~35 agreement lines
    # Match items from ITM-1001 to ITM-1022 (ITM-1023, 1024, 1025 will be unmatched in File 1)
    # Also add ITM-9901, ITM-9902 (unmatched in File 2)
    agreement_rows = []
    agr_counter = 50010
    
    # Matched items
    for meta in items_meta[:22]:
        item_code, desc, market, brand, cat, subcat, cm, source, uom, status = meta
        
        # Decide how many agreements for this item: 1 or 2
        num_agrs = 2 if item_code in ["ITM-1001", "ITM-1005", "ITM-1009", "ITM-1016", "ITM-1019"] else 1
        
        for a_idx in range(num_agrs):
            agr_num = f"BPA-{agr_counter}"
            agr_counter += 1
            
            sup_name, sup_site, sup_curr = random.choice(suppliers_pool)
            base_price = round(random.uniform(3.0, 48.0), 2)
            
            if a_idx == 0:
                line_status = "OPEN"
                expired_flag = "N"
                start_dt = "2025-01-01"
                end_dt = "2026-12-31"
                exp_dt = "2026-12-31"
            else:
                # Secondary or expired line
                line_status = "EXPIRED" if a_idx == 1 and random.random() > 0.4 else "OPEN"
                expired_flag = "Y" if line_status == "EXPIRED" else "N"
                start_dt = "2023-06-01"
                end_dt = "2024-12-31" if expired_flag == "Y" else "2026-12-31"
                exp_dt = end_dt
                
            bu_name = f"{market} Procurement BU"
            req_bu = f"{market} Manufacturing BU"
            
            agreement_rows.append([
                bu_name, req_bu, agr_num, sup_name, sup_site, sup_curr,
                item_code, desc, base_price, uom, cat,
                line_status, expired_flag, start_dt, end_dt, exp_dt
            ])
            
    # Add unmatched items in File 2 (items not in File 1)
    extra_f2_items = [
        ("ITM-9901", "Specialty Emulsifier Polysorbate 80", "Raw Materials", "KG", 18.50),
        ("ITM-9902", "Corrugated Shipper Heavy Duty 50L", "Packaging", "EA", 4.25),
        ("ITM-9903", "Aerosol Dispenser Valve Assembly", "Packaging", "EA", 0.85),
    ]
    for ext_code, ext_desc, ext_cat, ext_uom, ext_price in extra_f2_items:
        agr_num = f"BPA-{agr_counter}"
        agr_counter += 1
        sup_name, sup_site, sup_curr = random.choice(suppliers_pool)
        agreement_rows.append([
            "Global Sourcing BU", "Central Distribution BU", agr_num, sup_name, sup_site, sup_curr,
            ext_code, ext_desc, ext_price, ext_uom, ext_cat,
            "OPEN", "N", "2025-01-01", "2027-01-01", "2027-01-01"
        ])

    # Write to ws2
    for r_idx, row_data in enumerate(agreement_rows, start=14):
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws2.cell(row=r_idx, column=c_idx, value=val)
            if c_idx == 9: # Price
                cell.number_format = '$#,##0.00'

    wb2.save("2.xlsx")
    print("Successfully generated 2.xlsx")

if __name__ == "__main__":
    create_sample_files()

