# Spend & Procurement Agreement Analytics Dashboard

An interactive, high-performance web application built with **Python**, **Streamlit**, **Pandas**, and **Plotly** to ingest, parse, dynamically link, and visualize procurement spend versus supplier blanket purchase agreements (BPAs).

---

## 🚀 Quick Start Guide

### 1. Prerequisites
Ensure Python 3.10+ is installed on your system.

### 2. Install Dependencies
Run the following command in this directory:
```bash
pip install -r requirements.txt
```

### 3. Launch the Application
You can launch the dashboard using either of the following:

- **Option A (One-Click Windows Launcher):**
  Double-click `run.bat` in the project folder.

- **Option B (Terminal / Command Prompt):**
  ```bash
  streamlit run app.py
  ```

The application will launch automatically in your default browser at `http://localhost:8501`.

---

## 📁 Input Files & Intelligent Parsing Rules

The application includes resilient automatic header detection and sanitization for both input files:

### 1. File 1: Spend, Volume & PPV Tracker (`1.xlsx`)
- **Super-Header Parsing:** Row 2 contains grouped section super-headers (e.g., *Actual Volumes 2025*, *Actual Volume - 2026*, *Actual 2026 Spend in LC*, *Actual 2026 Spend in USD*, *Actual 2025 Spend in USD*, *PPV vs 2025 Avg Rate in 2025 Constant Fx*, *Fx Impact*, etc.).
- **Automatic Header Detection:** Automatically scans the top rows to locate the header row containing `Item Code`.
- **Merged Cell Forward-Fill & Collision Prevention:** Propagates super-headers across grouped sections and merges them with duplicate sub-headers (e.g., monthly columns such as `Jan-26`) so no columns collide.
- **Separator Elimination:** Strips empty and separator columns labeled `x` or `X` (case-insensitive).
- **Core Metadata Extraction:** Preserves clean attributes for `Market`, `Brand`, `Category`, `Sub Category`, `CM`, `Source`, `Item Code`, `Item Description`, `UOM`, `Updated Category`, `Updated Subcategory`, and `Status`.

### 2. File 2: Procurement Agreements & Supplier Pricing (`2.xlsx`)
- **Metadata Bypass:** Skips top metadata report filters (Rows 1–12) automatically.
- **Automatic Header Detection:** Automatically detects the header row containing `Item Name` (Row 13).
- **Contract Fields:** Ingests `Procurement Bu Name`, `Requisitioning Bu Name`, `Agreement Number`, `Supplier`, `Supplier Site Name`, `Currency Code`, `Item Name`, `Item Description`, `Price`, `Agreement UOM`, `Category Name`, `Line Status`, `Expired Blanket Agreement Line Flag`, `Start Date`, `End Date`, and `Expiration Date`.

---

## 🔗 Data Linking & Double-Counting Prevention Logic

### Normalized Key Matching
Data is linked using **`Item Code` (File 1) = `Item Name` (File 2)**:
- Both keys are trimmed of leading/trailing whitespace.
- Converted to uppercase string representations to ensure exact matching across disparate source systems.

### Join Modes
- **Left Join (Default):** Retains all items from File 1. Items with no active contract in File 2 are highlighted with "Unmatched / No Contract" status.
- **Inner Join:** Filters strictly to items present in both files.

### 🛡️ Double-Counting Prevention Architecture
In procurement datasets, a single item often has multiple supplier agreements (e.g., secondary sourcing, regional suppliers, or historical expired contracts). A naive flat join duplicates File 1's spend and volume records.

To prevent metric inflation:
1. **Item-Level Deduped Engine:** Aggregates File 2 contract attributes (agreement count, list of suppliers, primary supplier, average agreement price, min/max prices) to the unique item level before joining.
2. **KPIs & Spend Charts Integrity:** All KPI summary cards, Spend/Volume charts, Brand/Market distributions, and run-rates operate exclusively on the deduped item table. Every item's spend and volume is counted **exactly once**.
3. **Detailed Agreement Drill-Down:** Individual agreement lines are preserved in the agreement drill-down table and scatter variance plots.
4. **Active Agreement Filter:** Sidebar toggle allows isolating only active, non-expired agreements (`Line Status == OPEN` and `Expired Blanket Agreement Line Flag == N`).

---

## 📊 Dashboard Features

1. **File Upload & Live Match Banner:**
   - Drag-and-drop uploaders for File 1 and File 2.
   - Built-in sample generator button (`📁 Load Default Demo Files`) for instant testing.
   - Executive match banner displaying unique items per file, matched items count, and overall match rate percentage.

2. **Sidebar Filters:**
   - **File 1 Filters:** Market, Brand, Category, Sub Category, Contract Manufacturer (CM), Status, and Item Code / Description search.
   - **File 2 Filters:** Supplier, Currency Code, Line Status, Expired Blanket Flag, and Agreement Number.

3. **KPI Summary Cards:**
   - **Matched Items & Match Rate:** Count and coverage percentage.
   - **Total 2026 Spend (USD):** Realized spend with YoY comparison against 2025 and % change.
   - **Total 2026 Volume:** Realized volume across units.
   - **YTD PPV & Fx Impact:** Purchase price variance and foreign exchange breakdown.
   - **Active Suppliers:** Count of contracted suppliers in the active selection.

4. **Interactive Plotly Visualizations:**
   - **2026 vs 2025 Spend Comparison:** Grouped bar chart comparing 2025 vs 2026 spend by Item, Category, or Brand.
   - **Spend & Allocation by Supplier:** Horizontal bar chart displaying spend distribution or unique items contracted per supplier.
   - **Monthly Spend & Volume Run-Rate:** Dual-axis chart tracking Jan–Dec 2026 spend run-rate alongside volume delivery.
   - **Agreement Contract Price vs Realized Actual Rate:** Scatter variance radar with a 45-degree parity reference line ($y=x$) to instantly identify price leakage, off-contract spending, or contractual savings.
   - **Spend Distribution Donut:** Breakdown by Brand, Market, or Category.

5. **Data Tables & Exports:**
   - **Deduped Item-Level Summary:** Clean overview of spend, volume, realized unit rates, contract count, average contract price, and price variance ($).
   - **Detailed Agreement Drill-Down:** Complete contract lines with supplier site, currency, line status, expiration date, and agreement pricing.
   - **Unmatched Diagnostics:** Specific views for spend items lacking contracts and contracts with zero spend.
   - **Exports:** One-click downloads for filtered datasets in both `.xlsx` and `.csv` formats.

---

## 🛠️ Testing & Generating Sample Data

A sample data generator is included to create realistic files:
```bash
python generate_sample_data.py
```
This produces:
- `1.xlsx`: Spend & volume tracking 25 items across 6 markets and 5 brands, with 12 months of volumes, spend, LC spend, PPV, and Fx impact.
- `2.xlsx`: Procurement report with 12 metadata header rows, 30+ blanket purchase agreement lines, multiple suppliers, active/expired statuses, and intentional unmatched items for testing.

