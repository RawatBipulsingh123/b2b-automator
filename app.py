import streamlit as st
import pandas as pd
import io
import re

# ==========================================
# 1. Page Configuration & Custom CSS
# ==========================================
st.set_page_config(page_title="B2B Invoice Automator", page_icon="⚡", layout="wide")

st.markdown("""
    <style>
    .main .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    h1 { color: #1F4E78; }
    .stMetric { background-color: #F8FAFC; padding: 15px; border-radius: 8px; border: 1px solid #E2E8F0; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. Robust Extraction Logic (Carry Bag & Date Fixed)
# ==========================================
@st.cache_data(show_spinner=False)
def extract_invoice_data(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    data = []
    
    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet, header=None)
            
            # 1. Transfer From
            transfer_from = df.iloc[13, 1] if pd.notna(df.iloc[13, 1]) else "FABRIKA FASHION & LIFESTYLE LLP"
            
            # 2. Smart Store Name Extraction (Transfer to)
            address_block = f"{df.iloc[24,1]} {df.iloc[25,1]} {df.iloc[26,1]} {df.iloc[27,1]} {df.iloc[28,1]} {df.iloc[29,1]}".upper()
            
            if 'ETHNIQ' in address_block: 
                transfer_to = 'ETHNIQ RETAIL PRIVATE LIMITED'
            elif 'LAKESHORE' in address_block or 'VIVIANA' in address_block: 
                transfer_to = 'Viviana Mall'
            elif 'ORION' in address_block: 
                transfer_to = 'Orion Mall'
            elif 'CAPITAL' in address_block: 
                transfer_to = 'The Capital Mall'
            elif 'INORBIT' in address_block: 
                transfer_to = 'Inorbit Mall'
            elif 'SEAWOODS' in address_block: 
                transfer_to = 'Nexus Seawoods Mall'
            elif 'BORIVALI' in address_block: 
                transfer_to = 'RM Borivali'
            elif 'UDHANA' in address_block: 
                transfer_to = 'RM Udhana'
            elif 'BARODA' in address_block or 'VADODARA' in address_block: 
                transfer_to = 'RM Vadodara'
            elif 'RELIANCE' in address_block: 
                transfer_to = 'Reliance Mall (Other)'
            else: 
                transfer_to = f"{df.iloc[25,1]} {df.iloc[26,1]}".strip()
            
            # 3. Invoice Number & Date
            inv_str = str(df.iloc[14, 6]) if pd.notna(df.iloc[14, 6]) else ""
            date_str = str(df.iloc[15, 6]) if pd.notna(df.iloc[15, 6]) else ""

            inv_match = re.search(r':\s*(.+)', inv_str)
            inv_num = inv_match.group(1).strip() if inv_match else inv_str
            
            date_match = re.search(r':\s*(.+)', date_str)
            date_val = date_match.group(1).strip() if date_match else date_str
            
            # 4. Strict Item Extraction (Ignoring Carry Bags & Unlabeled Totals)
            true_qty = 0
            taxable = 0
            igst5 = 0
            igst18 = 0
            
            start_row = 30
            for i in range(25, 40):
                if 'Qty' in str(df.iloc[i, 6]) or 'Description' in str(df.iloc[i, 4]):
                    start_row = i + 1
                    break
                    
            for i in range(start_row, min(55, len(df))):
                row_text = str(df.iloc[i, 0]).lower() + str(df.iloc[i, 1]).lower() + str(df.iloc[i, 2]).lower() + str(df.iloc[i, 3]).lower()
                
                if 'total' in row_text or 'total value' in row_text:
                    break
                if pd.isna(df.iloc[i, 4]) and pd.isna(df.iloc[i, 2]):
                    break 
                
                q = df.iloc[i, 6]
                tax = df.iloc[i, 8]
                
                if pd.notna(q) and pd.notna(tax):
                    try:
                        q_val = float(q)
                        tax_val = float(tax)
                        if tax_val > 0: 
                            true_qty += q_val
                            taxable += tax_val
                            igst5 += pd.to_numeric(df.iloc[i, 9], errors='coerce') if pd.notna(df.iloc[i, 9]) else 0
                            igst18 += pd.to_numeric(df.iloc[i, 10], errors='coerce') if pd.notna(df.iloc[i, 10]) else 0
                    except:
                        continue
            
            if true_qty == 0:
                continue 
                
            # 5. Round off and Total Amount
            round_off_idx = df[df[6].astype(str).str.contains("Round Off", case=False, na=False)].index
            round_off = df.iloc[round_off_idx[0], 11] if len(round_off_idx) > 0 else 0
            
            tot_amt_idx = df[df[6].astype(str).str.contains("Total Amount after Tax", case=False, na=False)].index
            tot_amt = df.iloc[tot_amt_idx[0], 11] if len(tot_amt_idx) > 0 else 0
            
            rate = taxable / true_qty if true_qty > 0 else 0
            
            data.append({
                "Transfer From": transfer_from,
                "Transfer to": transfer_to,
                "DATE": date_val,
                "Invoice Number": inv_num,
                "Qty": true_qty,
                "Rate Per Unit": round(rate, 2),
                "Taxable Value": taxable,
                "IGST 5%": igst5,
                "IGST 18%": igst18,
                "Round Off": round_off,
                "Total Amount": tot_amt
            })
        except Exception:
            pass 
            
    df_final = pd.DataFrame(data)
    
    if not df_final.empty:
        # Date Parsing for Filtering
        df_final['DATE_Parsed'] = pd.to_datetime(df_final['DATE'].str.replace('-', '/'), format='%d/%m/%Y', errors='coerce')
        df_final['Month_Year'] = df_final['DATE_Parsed'].dt.strftime('%B %Y')
    
    return df_final

# ==========================================
# 3. Main UI & Dashboard Layout
# ==========================================
st.title("⚡ Advanced B2B Invoice Automator")
st.markdown("Automate data extraction from Fabrika B2B invoices with precise carry-bag exclusion and smart store mapping.")
st.divider()

# --- SIDEBAR: Upload & Filters ---
st.sidebar.markdown("### 📥 Data Ingestion")
uploaded_file = st.sidebar.file_uploader("Upload Fabrika Raw File (.xlsx)", type=["xlsx"])

if uploaded_file:
    with st.spinner("🚀 Scanning & Extracting 200+ invoices..."):
        df_b2b = extract_invoice_data(uploaded_file.getvalue())
    
    if not df_b2b.empty:
        st.sidebar.success(f"✅ Extracted {len(df_b2b)} Invoices")
        st.sidebar.divider()
        
        # --- FILTERS ---
        st.sidebar.markdown("### 🔍 Filters")
        
        # Month Filter
        available_months = ["All Months"] + sorted(df_b2b['Month_Year'].dropna().unique().tolist())
        selected_month = st.sidebar.selectbox("📅 Filter by Month", available_months)
        
        if selected_month != "All Months":
            filtered_df = df_b2b[df_b2b['Month_Year'] == selected_month]
        else:
            filtered_df = df_b2b
            
        # Store Filter
        available_stores = sorted(filtered_df['Transfer to'].dropna().unique().tolist())
        selected_stores = st.sidebar.multiselect("🏬 Filter by Store", available_stores, default=available_stores)
        
        if selected_stores:
            filtered_df = filtered_df[filtered_df['Transfer to'].isin(selected_stores)]
            
        # Drop helper columns for display/download
        display_df = filtered_df.drop(columns=['DATE_Parsed', 'Month_Year'], errors='ignore')
        
        # --- KPI DASHBOARD ---
        st.markdown(f"#### 📊 Performance Snapshot: {selected_month if selected_month != 'All Months' else 'Overall'}")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("🧾 Total Invoices", f"{len(display_df)}")
        c2.metric("👕 Total Garments", f"{display_df['Qty'].sum():,.0f}")
        c3.metric("💰 Total Taxable Value", f"₹ {display_df['Taxable Value'].sum():,.0f}")
        c4.metric("💵 Total Amount (Net)", f"₹ {display_df['Total Amount'].sum():,.0f}")
        
        st.divider()
        
        # --- DATA PREVIEW ---
        st.markdown("#### 📋 Extracted B2B Ledger")
        st.dataframe(display_df, use_container_width=True, height=350)
        
        # --- DOWNLOAD BUTTON ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            display_df.to_excel(writer, index=False, sheet_name="B2B_Data")
            
            # Auto-adjust column widths in Excel
            worksheet = writer.sheets['B2B_Data']
            for i, col in enumerate(display_df.columns):
                max_len = max(display_df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, max_len)
                
        st.divider()
        col_empty, col_btn, col_empty2 = st.columns([1, 2, 1])
        with col_btn:
            st.download_button(
                label=f"📥 DOWNLOAD B2B FORMAT EXCEL",
                data=output.getvalue(),
                file_name=f"Sales_B2B_{selected_month.replace(' ', '_')}.xlsx" if selected_month != "All Months" else "Sales_B2B_All_Months.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )
    else:
        st.error("❌ No valid invoice data found in this file. Please check the format.")
else:
    st.info("👈 Please upload the Raw Fabrika Excel file in the sidebar to get started.")