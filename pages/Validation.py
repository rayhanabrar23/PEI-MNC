import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="PEI Validation Engine", layout="wide")

st.title("🛡️ PEI Transaction Validation Engine")
st.info("Closing Price otomatis diambil dari file Margin. Data Non-Client otomatis difilter.")

# --- 1. PARAMETER GLOBAL ---
with st.sidebar:
    st.header("⚙️ Parameter Global")
    # Input manual untuk limit partisipan (MNC Sekuritas)
    limit_mnc_input = st.number_input("Total Credit Limit MNC Sekuritas (IDR)", min_value=0.0, value=100000000000.0)

# --- 2. UPLOAD AREA (7 FILES) ---
st.subheader("📁 Upload Data Sumber")
c1, c2 = st.columns(2)
with c1:
    f_invoice = st.file_uploader("1. Netting Invoice (xlsx)", type=['xlsx'])
    f_sid = st.file_uploader("2. SID Client (xlsx)", type=['xlsx'])
    f_risk = st.file_uploader("3. Risk Parameter (xlsx)", type=['xlsx'])
    f_limit_nasabah = st.file_uploader("4. Credit Limit Nasabah (.txt)", type=['txt'])
with c2:
    f_m_buy = st.file_uploader("5. Margin Buy (xlsx)", type=['xlsx'])
    f_m_sell = st.file_uploader("6. Margin Sell (xlsx)", type=['xlsx'])
    f_outstanding = st.file_uploader("7. Outstanding Position (.txt)", type=['txt'])

# --- 3. FUNGSI PARSING & STANDARISASI ---
def find_and_rename(df):
    """Mencari kolom secara adaptif agar tidak error 'None of Index'"""
    mapping = {
        'stock_key': ['no_share', 'no_shares', 'STOCK CODE', 'Stockcode', 'SYMBOL'],
        'sid_key': ['SID', 'SID_No', 'Client_SID'],
        'cid_key': ['no_cust', 'CID', 'Client_ID'],
        'avail_risk': ['Available Quantity', 'availablequantity', 'Available Qty'],
        'closing_key': ['CLOSING PRICE', 'Closing Price', 'Price']
    }
    rename_dict = {}
    for official, aliases in mapping.items():
        for col in df.columns:
            if str(col).strip() in aliases:
                rename_dict[col] = official
                break
    return df.rename(columns=rename_dict)

def parse_txt_files(file, mode):
    """Parsing file .txt Pipe Delimited"""
    content = file.getvalue().decode("utf-8").splitlines()
    if mode == 'outstanding':
        data_nasabah = []
        data_porto = []
        current_sid = None
        for line in content:
            p = line.split('|')
            if p[0] == '0': # Header Nasabah
                current_sid = p[3]
                data_nasabah.append({'sid_key': p[3], 'loan_existing': float(p[5]), 'accrued_interest': float(p[6])})
            elif p[0] == '1' and current_sid: # Detail Saham
                data_porto.append({'sid_key': current_sid, 'stock_key': p[3], 'vol_existing': float(p[4])})
        return pd.DataFrame(data_nasabah), pd.DataFrame(data_porto)
    else: # mode == 'limit_nasabah'
        rows = [l.split('|') for l in content]
        df = pd.DataFrame(rows[1:], columns=rows[0]) # Baris 1 adalah header
        # Ambil SID (kolom 3) dan Available Limit (kolom 8)
        return df.iloc[:, [2, 7]].rename(columns={df.columns[2]: 'sid_key', df.columns[7]: 'avail_limit_nasabah'})

# --- 4. VALIDATION LOGIC ---
if all([f_invoice, f_sid, f_risk, f_limit_nasabah, f_m_buy, f_m_sell, f_outstanding]):
    try:
        with st.spinner('Memproses Validasi...'):
            # Load Data
            df_inv = find_and_rename(pd.read_excel(f_invoice, dtype=str))
            df_sid_master = find_and_rename(pd.read_excel(f_sid, dtype=str))
            df_risk = find_and_rename(pd.read_excel(f_risk, dtype=str))
            df_mbuy = find_and_rename(pd.read_excel(f_m_buy, dtype=str))
            df_msell = find_and_rename(pd.read_excel(f_m_sell, dtype=str))
            
            # Parsing TXT
            df_n_master, df_p_exist = parse_txt_files(f_outstanding, 'outstanding')
            df_limit_nasabah = parse_txt_files(f_limit_nasabah, 'limit_nasabah')
            
            # Mapping Closing Price dari file Margin
            price_map = pd.concat([df_mbuy, df_msell]).drop_duplicates('stock_key').set_index('stock_key')['closing_key'].to_dict()
            haircut_map = df_mbuy.set_index('stock_key')['HAIRCUT'].to_dict() # Haircut biasanya di file margin

            # Filter hanya nasabah yang ada di SID Client (Clean)
            df_main = df_inv.merge(df_sid_master[['cid_key', 'sid_key', 'Name']], on='cid_key', how='inner')
            
            # Gabungkan dengan data Outstanding & Limit
            df_main = df_main.merge(df_n_master, on='sid_key', how='left')
            df_main = df_main.merge(df_limit_nasabah, on='sid_key', how='left')
            df_main = df_main.merge(df_p_exist, on=['sid_key', 'stock_key'], how='left')
            df_main = df_main.merge(df_risk[['stock_key', 'avail_risk']], on='stock_key', how='left')

            # Pembersihan angka
            cols_num = ['loan_existing', 'accrued_interest', 'avail_limit_nasabah', 'vol_existing', 'avail_risk', 'tot_vol', 'amt_pay']
            for c in cols_num: df_main[c] = pd.to_numeric(df_main[c], errors='coerce').fillna(0)

            # --- ENGINE VALIDASI ---
            def run_check(row):
                side = row['bors']
                errors = []
                price = float(price_map.get(row['stock_key'], 0))
                haircut = float(haircut_map.get(row['stock_key'], 0))
                
                if side == 'B': # A. LOAN REQUEST
                    # 1. Volume vs Available Risk
                    if row['tot_vol'] > row['avail_risk']: errors.append("Vol > Avail Risk")
                    # 2. Ratio 65% (Buy)
                    collateral = (row['vol_existing'] + row['tot_vol']) * price * (1 - haircut)
                    total_loan = row['loan_existing'] + row['amt_pay'] + row['accrued_interest']
                    if collateral > 0 and (total_loan / collateral) > 0.65: errors.append("Ratio > 65%")
                    # 4. Limit Nasabah
                    if row['amt_pay'] > row['avail_limit_nasabah']: errors.append("Over Client Limit")
                
                else: # B. REPAYMENT CASH
                    # 1. Volume vs Inventory
                    if row['tot_vol'] > row['vol_existing']: errors.append("Vol > Inventory")
                    # 2. Repay Value vs Loan Existing
                    if row['amt_pay'] > row['loan_existing']: errors.append("Repay > Loan Exist")
                    # 3. Ratio 65% (Sell)
                    collateral = (row['vol_existing'] - row['tot_vol']) * price * (1 - haircut)
                    total_loan = row['loan_existing'] - row['amt_pay'] + row['accrued_interest']
                    if collateral > 0 and (total_loan / collateral) > 0.65: errors.append("Ratio > 65%")

                return "✅ PASSED" if not errors else f"❌ REJECTED: {', '.join(errors)}"

            df_main['Validation_Status'] = df_main.apply(run_check, axis=1)

            # --- OUTPUT ---
            st.success("Analisis Validasi Berhasil!")
            st.dataframe(df_main[['sid_key', 'stock_key', 'bors', 'tot_vol', 'amt_pay', 'Validation_Status']], use_container_width=True)
            
            # Download
            output = io.BytesIO()
            df_main.to_excel(output, index=False)
            st.download_button("📥 Download Hasil Validasi", output.getvalue(), "Validation_Report.xlsx")

    except Exception as e:
        st.error(f"Error pada sistem: {e}")
