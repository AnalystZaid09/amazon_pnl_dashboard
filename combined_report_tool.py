import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import warnings
from datetime import datetime

warnings.filterwarnings('ignore', category=FutureWarning)

# ==========================================
# PAGE CONFIGURATION
# ==========================================
st.set_page_config(
    page_title="Amazon Combined Support Tool",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# PROFESSIONAL STYLING
# ==========================================
st.markdown("""
<style>
    .main { background-color: #F8FAFC; }
    .metric-container {
        background: white;
        padding: 20px;
        border-radius: 12px;
        border: 1px solid #E5E7EB;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    div[data-testid="stMetricValue"] {
        font-size: 24px;
        font-weight: 700;
        color: #1F2937;
    }
    div[data-testid="stMetricLabel"] {
        font-size: 14px;
        color: #6B7280;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# UTILITY FUNCTIONS
# ==========================================

def format_currency(val):
    if pd.isna(val): return "₹ 0.00"
    return f"₹ {val:,.2f}"

@st.cache_data
def convert_to_excel(df, sheet_name="Combined Report"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

def find_brand_col(df):
    """Attempt to find a column name that matches 'Brand' case-insensitively"""
    for col in df.columns:
        if str(col).strip().lower() == "brand":
            return col
    # Fallback to anything containing 'brand' but not metrics
    for col in df.columns:
        lc = str(col).lower()
        if "brand" in lc:
            # Skip columns that look like metrics (e.g. "Brand Damage Resolve")
            if any(m in lc for m in ["interest", "damage", "resolve", "%"]):
                continue
            return col
    return None

def normalize_df(df):
    """Normalize brand column for merging"""
    brand_col = find_brand_col(df)
    if brand_col:
        df = df.rename(columns={brand_col: "Brand"})
        df["Brand"] = df["Brand"].astype(str).str.replace('\r', '').str.replace('\n', ' ').str.strip().str.title()
    return df

def to_clean_numeric(series):
    """Convert series to numeric, handling commas and spaces"""
    return pd.to_numeric(series.astype(str).str.replace(',', '', regex=False).str.replace(' ', '', regex=False).replace(['nan', 'None', '', 'nan '], '0'), errors='coerce').fillna(0)

# ==========================================
# SIDEBAR - PROCESSED REPORT UPLOADS
# ==========================================
st.sidebar.title("📤 Upload Processed Reports")
st.sidebar.info("Upload all required processed reports manually to generate the summary.")

with st.sidebar.expander("📊 Primary Sales Data", expanded=True):
    net_sale_result = st.file_uploader("Net Sale: brand_summary.xlsx", type=["xlsx", "csv"], key="net_sale_res")

with st.sidebar.expander("🏷️ Support Tool Results"):
    coupon_res = st.file_uploader("Coupon: coupon_pivot_table.xlsx", type=["xlsx", "csv"])
    ncemi_res = st.file_uploader("NCEMI: ncemi_brand_analysis.csv", type=["xlsx", "csv"])
    ads_res = st.file_uploader("Ads: ads_pivot_table.xlsx", type=["xlsx", "csv"])
    exchange_res = st.file_uploader("Exchange: exchange_pivot_all.xlsx", type=["xlsx", "csv"])
    freebies_res = st.file_uploader("Freebies: freebies_pivot_table.xlsx", type=["xlsx", "csv"])
    rl_res = st.file_uploader("Replacement: rl_brand_summary.xlsx", type=["xlsx", "csv"])
    dyson_res = st.file_uploader("Dyson: dyson_final_support.csv", type=["xlsx", "csv"])
    inbound_res = st.file_uploader("Inbound: inbound_pickup_pivot.csv", type=["xlsx", "csv"])

with st.sidebar.expander("🏭 Secondary Support Results"):
    sec_bergner = st.file_uploader("Bergner Secondary Support", type=["xlsx", "csv"], key="sec_bergner")
    sec_tramontina = st.file_uploader("Tramontina Secondary Support", type=["xlsx", "csv"], key="sec_tramontina")
    sec_hafele = st.file_uploader("Hafele Secondary Support", type=["xlsx", "csv"], key="sec_hafele")
    sec_wonderchef = st.file_uploader("Wonderchef Secondary Support", type=["xlsx", "csv"], key="sec_wonderchef")
    sec_panasonic = st.file_uploader("Panasonic Secondary Support", type=["xlsx", "csv"], key="sec_panasonic")
    sec_inalsa = st.file_uploader("Inalsa Secondary Support", type=["xlsx", "csv"], key="sec_inalsa")
    sec_victorinox = st.file_uploader("Victorinox Secondary Support", type=["xlsx", "csv"], key="sec_victorinox")

with st.sidebar.expander("📉 System Metrics (RLC, Reimb, Damage)"):
    rev_fba = st.file_uploader("RLC FBA: rlc_fba_pivot.csv", type=["xlsx", "csv"])
    rev_sel = st.file_uploader("RLC Seller: rlc_seller_pivot.csv", type=["xlsx", "csv"])
    reimb_fba = st.file_uploader("Reimb FBA: reimbursement_fba_pivot.csv", type=["xlsx", "csv"])
    reimb_sel = st.file_uploader("Reimb Seller: reimbursement_seller_pivot.csv", type=["xlsx", "csv"])
    loss_fba = st.file_uploader("Loss FBA: loss_fba_pivot.csv", type=["xlsx", "csv"])
    loss_sel = st.file_uploader("Loss Seller: loss_seller_pivot.csv", type=["xlsx", "csv"])
    inv_res = st.file_uploader("Current Inventory: inventory_brand_pivot.csv", type=["xlsx", "csv"], key="inv_res")
    storage_res = st.file_uploader("Storage: storage_brand_pivot.csv", type=["xlsx", "csv"])
    damage_res = st.file_uploader("Current Damage: current_damage_summary.csv", type=["xlsx", "csv"])

with st.sidebar.expander("⚙️ Optional Overrides"):
    interest_damage_file = st.file_uploader("Interest & Damage Resolve Override", type=["xlsx", "xls"], key="int_dam_res")



# ==========================================
# MAIN UI
# ==========================================
st.title("🚀 Amazon Combined Support Tool")
st.markdown("Consolidate multiple processed reports into a single brand-wise summary.")

if not net_sale_result:
    st.info("👋 Please upload the **Net Sale Analyzer Result** (`brand_summary.xlsx`) to begin.")
    st.stop()

# Helper to load and normalize
def load_and_norm(f, prefix=None):
    if not f: return None
    try:
        # Check if f is our LocalF class or streamlit UploadedFile
        if hasattr(f, 'path'):
            if f.path.endswith(".csv"): df = pd.read_csv(f.path)
            else: df = pd.read_excel(f.path)
        else:
            if f.name.endswith(".csv"): df = pd.read_csv(f)
            else: df = pd.read_excel(f)
        
        df.columns = [str(c).replace('\n', ' ').strip() for c in df.columns]
        df = normalize_df(df)
        
        # If Brand is missing and we have a secondary support prefix, assign it
        if df is not None and "Brand" not in df.columns and prefix:
            if prefix in ["Bergner", "Tramontina", "Hafele", "Wonderchef", "Panasonic", "Victorinox"]:
                df["Brand"] = prefix
        
        # Remove total rows
        # Remove total rows
        if df is not None and "Brand" in df.columns:
            df = df[~df["Brand"].astype(str).str.upper().isin(["GRAND TOTAL", "TOTAL"])]
            
            if prefix:
                # Rename all non-brand columns
                for col in df.columns:
                    if col != "Brand":
                        c_str = str(col)
                        if not c_str.lower().startswith(prefix.lower()):
                            df.rename(columns={col: f"{prefix} {c_str}"}, inplace=True)
            
            # CRITICAL: Aggregate by Brand BEFORE returning to prevent 1-to-many merge duplicates
            # Convert all non-brand columns to numeric first
            for col in df.columns:
                if col != "Brand":
                    df[col] = to_clean_numeric(df[col])
            
            df = df.groupby("Brand").sum().reset_index()
            return df
        return df
    except Exception as e:
        return None

# Load primary file
base_df = load_and_norm(net_sale_result)
if base_df is None or "Brand" not in base_df.columns:
    st.error("Could not find a 'Brand' column in the Net Sale report.")
    st.stop()

# Merge Logic with specific prefixes
merges = [
    (coupon_res, "Coupon"), (ncemi_res, "NCEMI"), (ads_res, "Ads"), 
    (exchange_res, "Exchange"), (freebies_res, "Freebies"), (rl_res, "Replacement"), 
    (dyson_res, "Dyson"), (rev_fba, "RLC FBA"), (rev_sel, "RLC Seller"), 
    (reimb_fba, "Reimb FBA"), (reimb_sel, "Reimb Seller"), (loss_fba, "Loss FBA"), 
    (loss_sel, "Loss Seller"),    (inv_res, "Inventory"), (storage_res, "Storage"),
    (damage_res, "Damage"), (inbound_res, "Inbound"),
    # Secondary Files
    (sec_bergner, "Bergner"), (sec_tramontina, "Tramontina"), (sec_hafele, "Hafele"),
    (sec_wonderchef, "Wonderchef"), (sec_panasonic, "Panasonic"), (sec_inalsa, "Inalsa"),
    (sec_victorinox, "Victorinox")
]

final_df = base_df.copy()
for f, pref in merges:
    df_add = load_and_norm(f, prefix=pref)
    if df_add is not None and "Brand" in df_add.columns:
        cols = [c for c in df_add.columns if c != "Brand"]
        final_df = pd.merge(final_df, df_add[["Brand"] + cols], on="Brand", how="outer")

final_df["Brand"] = final_df["Brand"].fillna("Unknown/Unmapped")

# Safety group by Brand to ensure no duplicates after merges
# Ensure all columns except Brand are robustly numeric to prevent data loss in groupby
for col in final_df.columns:
    if col != "Brand":
        final_df[col] = to_clean_numeric(final_df[col])
final_df = final_df.groupby("Brand").sum().reset_index()

# ==========================================
# CALCULATION ENGINE
# ==========================================

# Comprehensive Column Mapping (Updated for prefixes)
col_rename = {
    "net sales": "Net Sales", "quantity": "Net Sales",
    "turn over": "Turn Over", "sales amount (turn over)": "Turn Over",
    "payout": "Payout", "transferred price": "Payout", "total": "Payout",
    "cost of goods sold": "Cost of goods sold", "cp as per qty": "Cost of goods sold",
    "profit": "Net PnL", "p&l": "Net PnL", "net pnl": "Net PnL",
    "coupon discount": "Coupon Support", "coupon support": "Coupon Support", "coupon total": "Coupon Support", "coupon coupon discount": "Coupon Support", "coupon coupon support": "Coupon Support",
    "total seller funding": "Exchange Support", "exchange funding": "Exchange Support", "exchange support": "Exchange Support", "exchange exchange support": "Exchange Support", "exchange total seller funding": "Exchange Support",
    "freebies discount": "Freebies", "freebies": "Freebies (Raw)", "freebies support": "Freebies Support", "freebies disc count": "Freebies Disc Count", "freebies base amount": "Freebies Base Amount",
    "total amount (incl. gst)": "Advertising Support", "advertising support": "Advertising Support", "ads support": "Advertising Support", "ads total amount (incl. gst)": "Advertising Support", "ads total amount": "Advertising Support",
    "ncemi support": "NCEMI Support", "ncemi total": "NCEMI Support", "ncemi ncemi support": "NCEMI Support",
    "inbound total": "Inbound Pick Up Service", "inbound inbound pick up service": "Inbound Pick Up Service", "inbound pick up service": "Inbound Pick Up Service",
    "replacement total": "Replacement charges", "replacement charges": "Replacement charges", "replacement quantity": "Replacement Quantity", "replacement tot78": "Replacement charges",
    "dyson support": "Price Support", "dyson support as per net sale": "Price Support", "dyson dyson support": "Price Support",
    # All secondary brands map to Price Support
    "bergner support": "Price Support", "bergner p/l on orders qty": "Price Support", "bergner sec cn value": "Price Support", "bergner bergner support": "Price Support",
    "tramontina support": "Price Support", "tramontina sec support": "Price Support", "tramontina tramontina support": "Price Support",
    "wonderchef support": "Price Support", "wonderchef sec support": "Price Support",
    "hafele support": "Price Support", "hafele sec support": "Price Support",
    "panasonic support": "Price Support", "panasonic sec support": "Price Support",
    "victorinox support": "Price Support", "victorinox sec support": "Price Support", "victorinox sec cn value": "Price Support",
    "inalsasupport": "Price Support", "inalsa credit note amount": "Price Support", "inalsa support": "Price Support",
    "reimb fba total": "Reimbursement FBA", "reimb fba fba reimbursement": "Reimbursement FBA", "reimbursement fba": "Reimbursement FBA", "reimb fba fba reimbursement amount": "Reimbursement FBA", "reimb fba amount-total": "Reimbursement FBA",
    "reimb seller total": "Reimbursement Seller Flex (Safe T Claim)", "reimbursement seller flex (safe t claim)": "Reimbursement Seller Flex (Safe T Claim)", "reimb seller seller reimbursement": "Reimbursement Seller Flex (Safe T Claim)",
    "rlc fba total": "Reverse logistics FBA", "rlc fba reverse logistics fba": "Reverse logistics FBA", "reverse logistics fba": "Reverse logistics FBA",
    "rlc seller total": "Reverse logistics Seller Flex Reverse", "reverse logistics seller flex reverse": "Reverse logistics Seller Flex Reverse", "rlc seller reverse logistics seller flex reverse": "Reverse logistics Seller Flex Reverse",
    "loss fba total": "Loss in damages FBA", "loss fba fba loss": "Loss in damages FBA", "loss in damages fba": "Loss in damages FBA",
    "loss seller total": "Loss in damages Seller Flex", "loss in damages seller flex": "Loss in damages Seller Flex", "loss seller seller loss": "Loss in damages Seller Flex",
    "inventory total": "Current Inventory", "inventory cp inventory value": "Current Inventory", "current inventory": "Current Inventory", "inventory current inventory value": "Current Inventory", "cp inventory value": "Current Inventory",
    "storage total": "Storage Charges", "storage storage fee": "Storage Charges", "storage charges": "Storage Charges", "storage estimated-monthly-storage-fee": "Storage Charges", "storage fee": "Storage Charges",
    "damage current damages": "Current damages", "damage cp as per qty": "Current damages", "current damages": "Current damages", "current cp as per qty": "Current damages", "damage total": "Current damages"
}

# Standardize columns by summing those that map to the same name
standardized_data = {"Brand": final_df["Brand"]}
for col in final_df.columns:
    if col == "Brand": continue
    
    c_lower = str(col).lower().strip()
    std_name = col_rename.get(c_lower, col) # Default to original if not in map
    
    if std_name not in standardized_data:
        standardized_data[std_name] = pd.to_numeric(final_df[col], errors="coerce").fillna(0)
    else:
        # Sum with existing
        standardized_data[std_name] = standardized_data[std_name] + pd.to_numeric(final_df[col], errors="coerce").fillna(0)

# Re-create final_df without duplicates
final_df = pd.DataFrame(standardized_data)
# Normalize Brand case one last time before grouping
final_df["Brand"] = final_df["Brand"].astype(str).str.strip().str.title()
# Group by Brand in case merges created duplicate brand rows (or case-variant brands)
final_df = final_df.groupby("Brand").sum().reset_index()

# List of columns to ensure presence
requested_cols = [
    "Net Sales", "Turn Over", "Payout", "Cost of goods sold", "Gross PnL Level 1",
    "Price Support", "Coupon Support", "NCEMI Support", "Freebies", "Exchange Support",
    "Advertising Support", "Inbound Pick Up Service", "Gross PnL level 2", "P&L%", "Reimbursement FBA",
    "Reimbursement Seller Flex (Safe T Claim)", "Total Reimbursement", "Reverse logistics FBA",
    "Reverse logistics Seller Flex Reverse", "Total Reverse", "Replacement charges",
    "Storage Charges", "Admin @1%", "Gross PnL level 3", "Interest %", "Interest",
    "Loss in damages FBA", "Loss in damages Seller Flex", "Loss in damages Total",
    "Damage Resolve %", "Actual Loss of Damage", "Net PnL", "Net PnL%", "Current Inventory",
    "Cost Of Interest Rate On Good", "Current damages", "Net", "Profit in %"
]

# Ensure presence and order
for col in requested_cols:
    if col not in final_df.columns:
        final_df[col] = 0.0

final_df = final_df.fillna(0)

# Re-run calculations for consistency
final_df["Gross PnL Level 1"] = final_df["Payout"] - final_df["Cost of goods sold"]
final_df["Total Reimbursement"] = final_df["Reimbursement FBA"] + final_df["Reimbursement Seller Flex (Safe T Claim)"]
final_df["Total Reverse"] = final_df["Reverse logistics FBA"] + final_df["Reverse logistics Seller Flex Reverse"]

final_df["Gross PnL level 2"] = (
    final_df["Gross PnL Level 1"] + 
    final_df["Price Support"] + 
    final_df["NCEMI Support"] + 
    final_df["Freebies"] + 
    final_df["Exchange Support"] + 
    final_df["Advertising Support"] +
    final_df["Inbound Pick Up Service"]
)

mask_to = final_df["Turn Over"] != 0
final_df.loc[mask_to, "P&L%"] = (final_df.loc[mask_to, "Gross PnL level 2"] / final_df.loc[mask_to, "Turn Over"]) * 100

final_df["Admin @1%"] = -final_df["Turn Over"] * 0.01

final_df["Gross PnL level 3"] = (
    final_df["Gross PnL level 2"] + 
    final_df["Total Reimbursement"] + 
    final_df["Total Reverse"] + 
    final_df["Replacement charges"] + 
    final_df["Storage Charges"] +
    final_df["Admin @1%"]
)

final_df["Loss in damages Total"] = final_df["Loss in damages FBA"] + final_df["Loss in damages Seller Flex"]
mask_dmg = final_df["Loss in damages Total"] != 0
final_df.loc[mask_dmg, "Damage Resolve %"] = (final_df.loc[mask_dmg, "Total Reimbursement"] / final_df.loc[mask_dmg, "Loss in damages Total"]) * 100


# Optional: Override Damage Resolve if provided by file
if interest_damage_file:
    try:
        id_df = pd.read_excel(interest_damage_file)
        id_df.columns = [str(c).replace('\n', ' ').strip() for c in id_df.columns]
        id_df = normalize_df(id_df)
        if "Brand" in id_df.columns:
            # Explicitly find and rename Interest/Damage columns in ID file
            for c in id_df.columns:
                lc = str(c).lower()
                if "interest" in lc and ("%" in lc or "rate" in lc):
                    id_df.rename(columns={c: "Interest %"}, inplace=True)
                elif "damage" in lc and ("%" in lc or "resolve" in lc):
                    id_df.rename(columns={c: "Damage Resolve %"}, inplace=True)
            
            # Ensure numeric and scale if they are in decimal format (e.g. 0.01 -> 1.0)
            if "Interest %" in id_df.columns:
                id_df["Interest %"] = pd.to_numeric(id_df["Interest %"], errors="coerce").fillna(0)
                if id_df["Interest %"].max() <= 1.0 and id_df["Interest %"].any():
                    id_df["Interest %"] = id_df["Interest %"] * 100
            
            if "Damage Resolve %" in id_df.columns:
                id_df["Damage Resolve %"] = pd.to_numeric(id_df["Damage Resolve %"], errors="coerce").fillna(0)
                if id_df["Damage Resolve %"].max() <= 1.0 and id_df["Damage Resolve %"].any():
                    id_df["Damage Resolve %"] = id_df["Damage Resolve %"] * 100

            # Drop existing columns from final_df to ensure override
            for col_to_ovr in ["Interest %", "Damage Resolve %"]:
                if col_to_ovr in id_df.columns and col_to_ovr in final_df.columns:
                    final_df.drop(columns=[col_to_ovr], inplace=True)
            
            # Merge the override values
            id_sub = id_df[["Brand"] + [c for c in ["Interest %", "Damage Resolve %"] if c in id_df.columns]]
            final_df = pd.merge(final_df, id_sub, on="Brand", how="left").fillna(0)
    except Exception as e:
        st.error(f"Error merging Interest/Damage override: {e}")

# Re-calculate Actual Loss of Damage (accounts for overrides)
final_df["Actual Loss of Damage"] = final_df["Loss in damages Total"] - (final_df["Loss in damages Total"] * final_df["Damage Resolve %"] / 100)

final_df["Net PnL"] = final_df["Gross PnL level 3"] - final_df["Actual Loss of Damage"]
mask_cogs = final_df["Cost of goods sold"] != 0
final_df.loc[mask_cogs, "Net PnL%"] = (final_df.loc[mask_cogs, "Net PnL"] / final_df.loc[mask_cogs, "Cost of goods sold"]) * 100

final_df["Interest"] = final_df["Turn Over"] * (final_df["Interest %"] / 100)
final_df["Cost Of Interest Rate On Good"] = -final_df["Current Inventory"] * (final_df["Interest %"] / 100)
final_df["Net"] = final_df["Net PnL"] + final_df["Cost Of Interest Rate On Good"]

mask_final_to = final_df["Turn Over"] != 0
final_df.loc[mask_final_to, "Profit in %"] = (final_df.loc[mask_final_to, "Net"] / final_df.loc[mask_final_to, "Turn Over"]) * 100

# Cleanup and Table
final_df = final_df[["Brand"] + requested_cols].sort_values("Net PnL", ascending=False)

# Summary Row
sum_row = final_df.sum(numeric_only=True).to_frame().T
sum_row["Brand"] = "TOTAL"

# Recalculate summary percentages
t_to = sum_row["Turn Over"].iloc[0]
t_gp2 = sum_row["Gross PnL level 2"].iloc[0]
sum_row["P&L%"] = (t_gp2 / t_to * 100) if t_to != 0 else 0

t_reimb = sum_row["Total Reimbursement"].iloc[0]
t_loss = sum_row["Loss in damages Total"].iloc[0]
sum_row["Damage Resolve %"] = (t_reimb / t_loss * 100) if t_loss != 0 else 0

t_net_pnl = sum_row["Net PnL"].iloc[0]
t_cogs = sum_row["Cost of goods sold"].iloc[0]
sum_row["Net PnL%"] = (t_net_pnl / t_cogs * 100) if t_cogs != 0 else 0

t_net = sum_row["Net"].iloc[0]
sum_row["Profit in %"] = (t_net / t_to * 100) if t_to != 0 else 0

final_df = pd.concat([final_df, sum_row], ignore_index=True)

# Show Metrics
m1, m2 = st.columns(2)
m1.metric("Total Brands", len(final_df[final_df["Brand"] != "TOTAL"]))
m2.metric("Total Net PnL", format_currency(sum_row["Net PnL"].iloc[0]))

# Final Styling
def format_percent(val):
    try:
        return f"{float(val):.2f}%"
    except:
        return str(val)

st.dataframe(
    final_df.style.format({
        **{c: format_currency for c in requested_cols if ("%" not in c and "in %" not in c) or c == "Admin @1%"},
        **{c: format_percent for c in requested_cols if ("%" in c or "in %" in c) and c != "Admin @1%"}
    }),
    use_container_width=True,
    height=600
)

st.download_button(
    "📥 Download Combined Report (Excel)",
    data=convert_to_excel(final_df),
    file_name=f"Combined_Report_{datetime.now().strftime('%Y%m%d')}.xlsx"
)
