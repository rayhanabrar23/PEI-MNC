import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="TRX PEI Details", layout="wide")

st.title("🚀 TRX PEI Details Generator")
st.info("Upload 5 file raw data untuk menghasilkan template Hasil_MNC (Sheet Buy & Sell).")

# --- 1. UPLOAD AREA ---
col_u1, col_u2 = st.columns(2)
with col_u1:
    file_invoice = st.file_uploader("1. Netting Invoice (CSV)", type=['csv'])
    file_sid = st.file_uploader("2. SID Client (CSV/Excel)", type=['csv', 'xlsx'])
    file_risk = st.file_uploader("3. Risk Parameter (CSV/Excel)", type=['csv', 'xlsx'])
with col_u2:
    file_m_buy = st.file_uploader("4. Margin Buy (CSV/Excel)", type=['csv', 'xlsx'])
    file_m_sell = st.file_uploader("5. Margin Sell (CSV/Excel)", type=['csv', 'xlsx'])

def load_file(file_obj, dtype_dict=None):
    if file_obj is None: return None
    if file_obj.name.endswith('.csv'):
        return pd.read_csv(file_obj, dtype=dtype_dict)
    else:
        return pd.read_excel(file_obj, dtype=dtype_dict)

if all([file_invoice, file_sid, file_risk, file_m_buy, file_m_sell]):
    try:
        # --- 2. LOAD & CLEANING DATA ---
        df_inv = load_file(file_invoice, {'no_cust': str, 'no_share': str})
        df_sid = load_file(file_sid, {'SID': str, 'CID': str})
        df_risk = load_file(file_risk, {'Stockcode': str})
        df_mbuy = load_file(file_m_buy, {'SID': str, 'STOCK CODE': str})
        df_msell = load_file(file_m_sell, {'SID': str, 'STOCK CODE': str})

        # Bersihkan angka di Invoice (amt_pay & tot_vol diambil dari logika sebelumnya)
        for c in ['amt_pay', 'tot_vol']:
            df_inv[c] = pd.to_numeric(df_inv[c].astype(str).str.replace(',', '').str.replace('"', ''), errors='coerce').fillna(0)
        
        # Hitung Volume_Formula (Logika dominan B/S)
        df_inv['vol_net'] = df_inv.apply(lambda x: x['tot_vol'] if x['bors'] == 'B' else -x['tot_vol'], axis=1)
        net_calc = df_inv.groupby(['no_cust', 'no_share'])['vol_net'].sum().reset_index()
        df_inv = df_inv.merge(net_calc, on=['no_cust', 'no_share'], suffixes=('', '_total'))

        def get_vol_formula(row):
            total = row['vol_net_total']
            if total < 0: return total if row['bors'] == 'S' else 0
            elif total > 0: return total if row['bors'] == 'B' else 0
            return 0
        df_inv['Volume_Formula'] = df_inv.apply(get_vol_formula, axis=1)

        # --- 3. PROCESSING SHEET BUY & SELL ---
        results = {}
        for side in ['B', 'S']:
            # Filter side
            work_df = df_inv[df_inv['bors'] == side].copy()
            
            # Join SID Client (Mendapatkan Name & SID)
            work_df = work_df.merge(df_sid, left_on='no_cust', right_on='CID', how='left')
            
            # Pisahkan yang PEI dan Not PEI
            pei_data = work_df[work_df['SID'].notna()].copy()
            not_pei = work_df[work_df['SID'].isna()].copy()
            not_pei['CID'] = "Not Client PEI"

            # Join Margin Data & Risk Parameter
            if side == 'B':
                pei_data = pei_data.merge(df_mbuy, on=['SID', 'no_share'], how='left')
                pei_data['PEI (Risk/Porto)'] = pei_data.merge(df_risk, left_on='no_share', right_on='Stockcode', how='left')['availablequantity']
                val_col = 'AVAILABLE MARKET VALUE'
            else:
                # Untuk Sell, join dengan Margin Sell
                pei_data = pei_data.merge(df_msell, left_on=['SID', 'no_share'], right_on=['SID', 'STOCK CODE'], how='left')
                pei_data['PEI (Risk/Porto)'] = 0 # Sesuai instruksi fokus buy
                val_col = 'AVAILABLE SELL VALUE'

            # --- 4. MAPPING KE TEMPLATE 17 KOLOM ---
            final_cols = pd.DataFrame()
            final_cols['SID'] = pei_data['SID']
            final_cols['STOCK CODE'] = pei_data['no_share']
            
            # Kolom Spesifik Buy/Sell
            if side == 'B':
                for c in ['MARGIN BUY QUANTITY', 'LOAN QUANTITY', 'AVAILABLE QUANTITY', 'CLOSING PRICE', 'AVAILABLE MARKET VALUE', 'HAIRCUT', 'AVAILABLE COLLATERAL VALUE']:
                    final_cols[c] = pei_data[c] if c in pei_data else 0
            else:
                for c in ['REGULAR SELL QUANTITY', 'REPAYMENT QUANTITY', 'AVAILABLE SELL QUANTITY', 'CLOSING PRICE', 'AVAILABLE SELL VALUE']:
                    final_cols[c] = pei_data[c] if c in pei_data else 0

            final_cols['B/S'] = side
            final_cols['CID'] = pei_data['no_cust']
            final_cols['Name'] = pei_data['Name']
            final_cols['Stock'] = pei_data['no_share']
            
            # Logika Error Volume
            def validate_vol(v):
                if side == 'B' and v < 0: return "ERROR: Wrong Side"
                if side == 'S' and v > 0: return "ERROR: Wrong Side"
                return v
            final_cols['Volume'] = pei_data['Volume_Formula'].apply(validate_vol)
            
            final_cols['Value'] = pei_data[val_col] if val_col in pei_data else 0
            final_cols['PEI (Risk/Porto)'] = pei_data['PEI (Risk/Porto)']

            # Logika NETT (Formula Complex)
            def apply_nett_logic(row):
                # Merujuk pada logika P (Volume) dan S (PEI Risk)
                p = row['Volume']
                s = row['PEI (Risk/Porto)']
                if p == "ERROR: Wrong Side" or p == 0: return ""
                if p > 0: return "LOAN PEI" if s > p else "LOAN PARTIAL"
                else: return "REPAY PEI" if s < p else "ALL STOCK REPAY"
            
            final_cols['NETT'] = final_cols.apply(apply_nett_logic, axis=1)

            # Gabungkan dengan Not PEI di paling bawah
            final_df = pd.concat([final_cols, not_pei[['no_cust']].rename(columns={'no_cust': 'CID'})], ignore_index=True)
            results[side] = final_df

        # --- 5. DOWNLOAD ---
        st.success("✅ File Hasil_MNC Berhasil Dibuat!")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            results['B'].to_excel(writer, index=False, sheet_name='Buy')
            results['S'].to_excel(writer, index=False, sheet_name='Sell')
        
        st.download_button("📥 Download Hasil_MNC.xlsx", output.getvalue(), "Hasil_MNC.xlsx")

    except Exception as e:
        st.error(f"Terjadi kesalahan pemrosesan: {e}")
