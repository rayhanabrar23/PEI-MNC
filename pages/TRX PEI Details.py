import streamlit as st
import pandas as pd
import io

# 1. Konfigurasi Halaman
st.set_page_config(page_title="TRX PEI Details", layout="wide")

st.title("📑 TRX PEI Details Generator")
st.info("Sistem ini hanya akan memproses data nasabah yang terdaftar di SID Client (Bersih).")

# --- 2. FUNGSI STANDARISASI KOLOM ---
def find_and_rename(df):
    mapping = {
        'stock_key': ['no_share', 'no_shares', 'STOCK CODE', 'Stockcode', 'Stock', 'SYMBOL', 'Stock Code'],
        'sid_key': ['SID', 'SID_No', 'Client_SID'],
        'cid_key': ['no_cust', 'CID', 'Client_ID', 'Account_No'],
        'avail_risk': ['Available Quantity', 'availablequantity', 'Available Qty'],
        'name_key': ['Name', 'Client_Name', 'Nama']
    }
    rename_dict = {}
    for official_name, aliases in mapping.items():
        for col in df.columns:
            if str(col).strip() in aliases:
                rename_dict[col] = official_name
                break
    return df.rename(columns=rename_dict)

# --- 3. AREA UPLOAD ---
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
        with st.spinner('Menghubungkan seluruh data...'):
            df_inv = find_and_rename(pd.read_excel(file_invoice, dtype=str))
            df_sid = find_and_rename(pd.read_excel(file_sid_client, dtype=str))
            df_risk = find_and_rename(pd.read_excel(file_risk, dtype=str))
            df_mbuy = find_and_rename(pd.read_excel(file_m_buy, dtype=str))
            df_msell = find_and_rename(pd.read_excel(file_m_sell, dtype=str))

            def clean_num(df):
                for c in df.columns:
                    if any(key in str(c).lower() for key in ['amt', 'vol', 'qty', 'val', 'price', 'avail', 'haircut']):
                        df[c] = pd.to_numeric(df[c].astype(str).str.replace(',', '').str.replace('"', ''), errors='coerce').fillna(0)
                return df

            df_inv = clean_num(df_inv)
            df_mbuy = clean_num(df_mbuy)
            df_msell = clean_num(df_msell)
            df_risk = clean_num(df_risk)

            # Logika Volume Formula
            df_inv['vol_net_total'] = df_inv.groupby(['cid_key', 'stock_key'])['tot_vol'].transform(
                lambda x: (df_inv.loc[x.index, 'tot_vol'] * df_inv.loc[x.index, 'bors'].map({'B': 1, 'S': -1})).sum()
            )
            def calc_vol_formula(row):
                total = row['vol_net_total']
                if total < 0: return total if row['bors'] == 'S' else 0
                elif total > 0: return total if row['bors'] == 'B' else 0
                return 0
            df_inv['Volume_Formula'] = df_inv.apply(calc_vol_formula, axis=1)

            final_sheets = {}
            for side in ['B', 'S']:
                work_df = df_inv[df_inv['bors'] == side].copy()
                
                # Inner Join agar data Non-Client otomatis terhapus
                pei_data = work_df.merge(df_sid, on='cid_key', how='inner')

                if side == 'B':
                    pei_data = pei_data.merge(df_mbuy, on=['sid_key', 'stock_key'], how='left', suffixes=('', '_m'))
                    risk_sub = df_risk[['stock_key', 'avail_risk']].drop_duplicates('stock_key')
                    pei_data = pei_data.merge(risk_sub, on='stock_key', how='left')
                    val_src = 'AVAILABLE MARKET VALUE'
                else:
                    pei_data = pei_data.merge(df_msell, on=['sid_key', 'stock_key'], how='left', suffixes=('', '_m'))
                    pei_data['avail_risk'] = 0
                    val_src = 'AVAILABLE SELL VALUE'

                # --- BUILD TEMPLATE 17 KOLOM ---
                template = pd.DataFrame()
                template['SID'] = pei_data['sid_key']
                template['STOCK CODE'] = pei_data['stock_key']
                
                if side == 'B':
                    cols = ['MARGIN BUY QUANTITY', 'LOAN QUANTITY', 'AVAILABLE QUANTITY', 'CLOSING PRICE', 'AVAILABLE MARKET VALUE', 'HAIRCUT', 'AVAILABLE COLLATERAL VALUE']
                else:
                    cols = ['REGULAR SELL QUANTITY', 'REPAYMENT QUANTITY', 'AVAILABLE SELL QUANTITY', 'CLOSING PRICE', 'AVAILABLE SELL VALUE']
                
                for c in cols:
                    template[c] = pei_data[c] if c in pei_data.columns else 0

                template['B/S'] = side
                template['CID'] = pei_data['cid_key']
                template['Name'] = pei_data['name_key'] if 'name_key' in pei_data else ""
                template['Stock'] = pei_data['stock_key']
                
                def check_vol(v):
                    if side == 'B' and v < 0: return "ERROR: Wrong Side"
                    if side == 'S' and v > 0: return "ERROR: Wrong Side"
                    return v
                
                template['Volume'] = pei_data['Volume_Formula'].apply(check_vol)
                def calculate_real_value(row):
                # Jika volume 0 atau error, value harus 0
                vol = row['Volume_Formula']
                val = row[val_src] if val_src in row and pd.notnull(row[val_src]) else 0
    
                    if vol == 0:
                    return 0
                return val

                template['Value'] = pei_data.apply(calculate_real_value, axis=1)
                template['PEI (Risk/Porto)'] = pei_data['avail_risk']

                def get_nett_status(row):
                    p, s = row['Volume'], row['PEI (Risk/Porto)']
                    if p == "ERROR: Wrong Side" or p == 0: return ""
                    if side == 'B': return "LOAN PEI" if s > p else "LOAN PARTIAL"
                    else: return "REPAY PEI" if abs(s) < abs(p) else "ALL STOCK REPAY"
                
                template['NETT'] = template.apply(get_nett_status, axis=1)
                final_sheets[side] = template.fillna("")

            # --- 4. PREVIEW & DOWNLOAD ---
            st.success("✅ Data Berhasil Diproses (Data Non-Client Otomatis Dibuang)!")
            
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as wr:
                final_sheets['B'].to_excel(wr, index=False, sheet_name='Buy')
                final_sheets['S'].to_excel(wr, index=False, sheet_name='Sell')
            
            st.download_button("📥 Download Hasil_MNC.xlsx", out.getvalue(), "Hasil_MNC.xlsx")

            st.divider()
            tab1, tab2 = st.tabs(["📊 Preview BUY", "📊 Preview SELL"])
            with tab1: st.dataframe(final_sheets['B'], use_container_width=True)
            with tab2: st.dataframe(final_sheets['S'], use_container_width=True)

    except Exception as e:
        st.error(f"Gagal memproses data: {e}")
