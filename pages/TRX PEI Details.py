import streamlit as st
import pandas as pd
import io

# 1. Konfigurasi Halaman
st.set_page_config(page_title="TRX PEI Details", layout="wide")

st.title("TRX PEI Details Generator (Excel Only)")
st.info("Halaman ini mengolah 5 data sumber (.xlsx) menjadi file template Hasil_MNC.")

# --- 2. UPLOAD AREA (Khusus XLSX) ---
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
        with st.spinner('Sedang memproses data Excel...'):
            # --- 3. LOAD DATA ---
            # Memastikan ID Client dan SID dibaca sebagai string agar 0 di depan tidak hilang
            df_inv = pd.read_excel(file_invoice, dtype={'no_cust': str, 'no_share': str})
            df_sid = pd.read_excel(file_sid_client, dtype={'SID': str, 'CID': str})
            df_risk = pd.read_excel(file_risk, dtype={'Stockcode': str})
            df_mbuy = pd.read_excel(file_m_buy, dtype={'SID': str, 'STOCK CODE': str})
            df_msell = pd.read_excel(file_m_sell, dtype={'SID': str, 'STOCK CODE': str})

            # Pembersihan Angka di Invoice
            for c in ['amt_pay', 'tot_vol']:
                df_inv[c] = pd.to_numeric(df_inv[c].astype(str).str.replace(',', '').str.replace('"', ''), errors='coerce').fillna(0)
            
            # Hitung Volume_Formula (Logika Dominan B/S)
            df_inv['vol_net_total'] = df_inv.groupby(['no_cust', 'no_share'])['tot_vol'].transform(
                lambda x: (df_inv.loc[x.index, 'tot_vol'] * df_inv.loc[x.index, 'bors'].map({'B': 1, 'S': -1})).sum()
            )

            def get_vol_formula(row):
                total = row['vol_net_total']
                if total < 0: return total if row['bors'] == 'S' else 0
                elif total > 0: return total if row['bors'] == 'B' else 0
                return 0
            df_inv['Volume_Formula'] = df_inv.apply(get_vol_formula, axis=1)

            # --- 4. PROCESSING SHEET BUY & SELL ---
            final_sheets = {}
            for side in ['B', 'S']:
                work_df = df_inv[df_inv['bors'] == side].copy()
                
                # Join SID Client (CID ke no_cust)
                work_df = work_df.merge(df_sid, left_on='no_cust', right_on='CID', how='left')
                
                # Pisahkan PEI dan Non-PEI
                pei_data = work_df[work_df['SID'].notna()].copy()
                not_pei = work_df[work_df['SID'].isna()].copy()

                # Join Data Tambahan untuk PEI
                if side == 'B':
                    pei_data = pei_data.merge(df_mbuy, on=['SID', 'no_share'], how='left')
                    pei_data = pei_data.merge(df_risk[['Stockcode', 'availablequantity']], left_on='no_share', right_on='Stockcode', how='left')
                    val_col = 'AVAILABLE MARKET VALUE'
                else:
                    # Rename STOCK CODE di Margin Sell agar sinkron
                    df_msell_clean = df_msell.rename(columns={'STOCK CODE': 'no_share'})
                    pei_data = pei_data.merge(df_msell_clean, on=['SID', 'no_share'], how='left')
                    pei_data['availablequantity'] = 0 
                    val_col = 'AVAILABLE SELL VALUE'

                # --- 5. MAPPING KOLOM TEMPLATE (17 KOLOM) ---
                template = pd.DataFrame()
                template['SID'] = pei_data['SID']
                template['STOCK CODE'] = pei_data['no_share']
                
                if side == 'B':
                    cols_buy = ['MARGIN BUY QUANTITY', 'LOAN QUANTITY', 'AVAILABLE QUANTITY', 'CLOSING PRICE', 'AVAILABLE MARKET VALUE', 'HAIRCUT', 'AVAILABLE COLLATERAL VALUE']
                    for c in cols_buy: template[c] = pei_data[c] if c in pei_data else 0
                else:
                    cols_sell = ['REGULAR SELL QUANTITY', 'REPAYMENT QUANTITY', 'AVAILABLE SELL QUANTITY', 'CLOSING PRICE', 'AVAILABLE SELL VALUE']
                    for c in cols_sell: template[c] = pei_data[c] if c in pei_data else 0

                template['B/S'] = side
                template['CID'] = pei_data['no_cust']
                template['Name'] = pei_data['Name']
                template['Stock'] = pei_data['no_share']
                
                # Logika Volume & Error Side
                def check_side(v):
                    if side == 'B' and v < 0: return "ERROR: Wrong Side"
                    if side == 'S' and v > 0: return "ERROR: Wrong Side"
                    return v
                template['Volume'] = pei_data['Volume_Formula'].apply(check_side)
                
                template['Value'] = pei_data[val_col] if val_col in pei_data else 0
                template['PEI (Risk/Porto)'] = pei_data['availablequantity'] if 'availablequantity' in pei_data else 0

                # Logika NETT (Formula)
                def get_nett(row):
                    p, s = row['Volume'], row['PEI (Risk/Porto)']
                    if p == "ERROR: Wrong Side" or p == 0: return ""
                    if p > 0: return "LOAN PEI" if s > p else "LOAN PARTIAL"
                    else: return "REPAY PEI" if s < p else "ALL STOCK REPAY"
                template['NETT'] = template.apply(get_nett, axis=1)

                # Gabungkan dengan Not Client PEI
                not_pei_rows = pd.DataFrame({'CID': ['Not Client PEI'] * len(not_pei)})
                final_sheets[side] = pd.concat([template, not_pei_rows], ignore_index=True).fillna("")

            # --- 6. OUTPUT & PREVIEW ---
            st.success("✅ Berhasil memproses data!")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_sheets['B'].to_excel(writer, index=False, sheet_name='Buy')
                final_sheets['S'].to_excel(writer, index=False, sheet_name='Sell')
            
            st.download_button("📥 Download Hasil_MNC.xlsx", output.getvalue(), "Hasil_MNC.xlsx")

            st.divider()
            tab1, tab2 = st.tabs(["Preview Sheet Buy", "Preview Sheet Sell"])
            with tab1: st.dataframe(final_sheets['B'], use_container_width=True)
            with tab2: st.dataframe(final_sheets['S'], use_container_width=True)

    except Exception as e:
        st.error(f"Gagal memproses file Excel: {e}")
