import streamlit as st
import pandas as pd
import io

# 1. Konfigurasi Halaman
st.set_page_config(page_title="TRX PEI Details", layout="wide")

st.title("🚀 TRX PEI Details Generator")
st.info("Aplikasi ini otomatis mendeteksi kolom kode saham (no_share, STOCK CODE, dll) di semua file Excel.")

# --- 2. FUNGSI STANDARISASI KOLOM ---
def standardize_columns(df, type_context):
    """Mencari dan menyeragamkan nama kolom kunci agar konsisten"""
    # List kemungkinan nama kolom untuk Kode Saham
    stock_names = ['no_share', 'STOCK CODE', 'Stock', 'Stockcode', 'no_shares', 'SYMBOL']
    # List kemungkinan nama kolom untuk SID
    sid_names = ['SID', 'SID_No', 'Client_SID']
    # List kemungkinan nama kolom untuk CID/Client ID
    cid_names = ['no_cust', 'CID', 'Client_ID', 'Account_No']

    new_cols = {}
    for col in df.columns:
        if col in stock_names: new_cols[col] = 'stock_key'
        if col in sid_names: new_cols[col] = 'sid_key'
        if col in cid_names: new_cols[col] = 'cid_key'
    
    return df.rename(columns=new_cols)

# --- 3. UPLOAD AREA ---
col_u1, col_u2 = st.columns(2)
with col_u1:
    file_invoice = st.file_uploader("1. Netting Invoice (xlsx)", type=['xlsx'])
    file_sid_client = st.file_uploader("2. SID Client (xlsx)", type=['xlsx'])
    file_risk = st.file_uploader("3. Risk Parameter (xlsx)", type=['xlsx'])
with col_u2:
    file_m_buy = st.file_uploader("4. Margin Buy (xlsx)", type=['xlsx'])
    file_m_sell = st.file_uploader("5. Margin Sell (xlsx)", type=['xlsx'])

if all([file_invoice, file_sid_client, file_risk, file_m_buy, file_m_sell]):
    try:
        with st.spinner('Menyelaraskan kolom-kolom data...'):
            # Load & Standardize
            df_inv = standardize_columns(pd.read_excel(file_invoice, dtype=str), 'inv')
            df_sid = standardize_columns(pd.read_excel(file_sid_client, dtype=str), 'sid_master')
            df_risk = standardize_columns(pd.read_excel(file_risk, dtype=str), 'risk')
            df_mbuy = standardize_columns(pd.read_excel(file_m_buy, dtype=str), 'mbuy')
            df_msell = standardize_columns(pd.read_excel(file_m_sell, dtype=str), 'msell')

            # --- 4. DATA CLEANING & FORMULA ---
            # Pastikan angka terbaca benar
            for c in ['amt_pay', 'tot_vol']:
                if c in df_inv.columns:
                    df_inv[c] = pd.to_numeric(df_inv[c].str.replace(',', '').str.replace('"', ''), errors='coerce').fillna(0)
            
            # Hitung Volume_Formula
            df_inv['vol_net_total'] = df_inv.groupby(['cid_key', 'stock_key'])['tot_vol'].transform(
                lambda x: (df_inv.loc[x.index, 'tot_vol'] * df_inv.loc[x.index, 'bors'].map({'B': 1, 'S': -1})).sum()
            )

            def get_vol_formula(row):
                total = row['vol_net_total']
                if total < 0: return total if row['bors'] == 'S' else 0
                elif total > 0: return total if row['bors'] == 'B' else 0
                return 0
            df_inv['Volume_Formula'] = df_inv.apply(get_vol_formula, axis=1)

            # --- 5. PROCESSING SHEET BUY & SELL ---
            final_sheets = {}
            for side in ['B', 'S']:
                work_df = df_inv[df_inv['bors'] == side].copy()
                
                # Join SID Client (Mapping CID ke Name & SID)
                work_df = work_df.merge(df_sid, on='cid_key', how='left')
                
                # Pisahkan PEI dan Non-PEI
                pei_data = work_df[work_df['sid_key'].notna()].copy()
                not_pei = work_df[work_df['sid_key'].isna()].copy()

                # Join Margin & Risk Data
                if side == 'B':
                    pei_data = pei_data.merge(df_mbuy, on=['sid_key', 'stock_key'], how='left')
                    # Ambil availablequantity dari risk parameter
                    risk_subset = df_risk[['stock_key', 'availablequantity']].drop_duplicates('stock_key')
                    pei_data = pei_data.merge(risk_subset, on='stock_key', how='left')
                    val_col = 'AVAILABLE MARKET VALUE'
                else:
                    pei_data = pei_data.merge(df_msell, on=['sid_key', 'stock_key'], how='left')
                    pei_data['availablequantity'] = 0 
                    val_col = 'AVAILABLE SELL VALUE'

                # --- 6. MAPPING KE TEMPLATE ---
                template = pd.DataFrame()
                template['SID'] = pei_data['sid_key']
                template['STOCK CODE'] = pei_data['stock_key']
                
                # Daftar kolom dinamis
                cols_target = ['MARGIN BUY QUANTITY', 'LOAN QUANTITY', 'AVAILABLE QUANTITY', 'CLOSING PRICE', 
                               'AVAILABLE MARKET VALUE', 'HAIRCUT', 'AVAILABLE COLLATERAL VALUE',
                               'REGULAR SELL QUANTITY', 'REPAYMENT QUANTITY', 'AVAILABLE SELL QUANTITY', 'AVAILABLE SELL VALUE']
                
                for c in cols_target:
                    template[c] = pei_data[c] if c in pei_data else 0

                template['B/S'] = side
                template['CID'] = pei_data['cid_key']
                template['Name'] = pei_data['Name']
                template['Stock'] = pei_data['stock_key']
                
                # Logika Volume & Error Side
                def check_side(v):
                    if side == 'B' and float(v) < 0: return "ERROR: Wrong Side"
                    if side == 'S' and float(v) > 0: return "ERROR: Wrong Side"
                    return v
                template['Volume'] = pei_data['Volume_Formula'].apply(check_side)
                
                template['Value'] = pei_data[val_col] if val_col in pei_data else 0
                template['PEI (Risk/Porto)'] = pei_data['availablequantity'] if 'availablequantity' in pei_data else 0

                # Logika NETT
                def get_nett(row):
                    p = row['Volume']
                    s = pd.to_numeric(row['PEI (Risk/Porto)'], errors='coerce') or 0
                    if p == "ERROR: Wrong Side" or p == 0: return ""
                    p_num = float(p)
                    if p_num > 0: return "LOAN PEI" if s > p_num else "LOAN PARTIAL"
                    else: return "REPAY PEI" if s < p_num else "ALL STOCK REPAY"
                
                template['NETT'] = template.apply(get_nett, axis=1)

                # Tambahkan baris Not Client PEI
                not_pei_rows = pd.DataFrame({'CID': ['Not Client PEI'] * len(not_pei)})
                final_sheets[side] = pd.concat([template, not_pei_rows], ignore_index=True).fillna("")

            # --- 7. EXPORT ---
            st.success("✅ Berhasil memproses! Silakan download hasilnya.")
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_sheets['B'].to_excel(writer, index=False, sheet_name='Buy')
                final_sheets['S'].to_excel(writer, index=False, sheet_name='Sell')
            
            st.download_button("📥 Download Hasil_MNC.xlsx", output.getvalue(), "Hasil_MNC.xlsx")

    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
        st.info("Pastikan semua file memiliki header di baris pertama.")
