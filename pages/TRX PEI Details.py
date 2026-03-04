import streamlit as st
import pandas as pd
import io

# 1. Konfigurasi Halaman
st.set_page_config(page_title="TRX PEI Details", layout="wide")

st.title("🚀 TRX PEI Details Generator")
st.info("Gunakan Sidebar di sebelah kiri untuk menyesuaikan nama kolom jika terjadi error.")

# --- 2. UPLOAD AREA ---
with st.sidebar:
    st.header("⚙️ Pengaturan Kolom")
    st.write("Jika kolom tidak terdeteksi, ubah nama di bawah ini sesuai file Excel-mu.")
    col_stock = st.text_input("Nama Kolom Kode Saham", value="no_share")
    col_sid = st.text_input("Nama Kolom SID", value="SID")
    col_cid = st.text_input("Nama Kolom CID (no_cust)", value="no_cust")
    col_avail = st.text_input("Nama Kolom Available Qty (Risk)", value="availablequantity")

st.subheader("📁 Upload Raw Data (Excel)")
col_u1, col_u2 = st.columns(2)
with col_u1:
    file_invoice = st.file_uploader("1. Netting Invoice", type=['xlsx'])
    file_sid_client = st.file_uploader("2. SID Client", type=['xlsx'])
    file_risk = st.file_uploader("3. Risk Parameter", type=['xlsx'])
with col_u2:
    file_m_buy = st.file_uploader("4. Margin Buy", type=['xlsx'])
    file_m_sell = st.file_uploader("5. Margin Sell", type=['xlsx'])

# --- 3. FUNGSI STANDARISASI ---
def unify_df(df, mapping):
    """Menyeragamkan kolom berdasarkan input user di sidebar"""
    inv_map = {v: k for k, v in mapping.items()}
    return df.rename(columns=inv_map)

if all([file_invoice, file_sid_client, file_risk, file_m_buy, file_m_sell]):
    try:
        # Mapping kunci
        mapping = {
            'stock_key': col_stock,
            'sid_key': col_sid,
            'cid_key': col_cid,
            'avail_key': col_avail
        }

        with st.spinner('Menghubungkan data...'):
            # Load Data
            df_inv = unify_df(pd.read_excel(file_invoice, dtype=str), mapping)
            df_sid = unify_df(pd.read_excel(file_sid_client, dtype=str), mapping)
            df_risk = unify_df(pd.read_excel(file_risk, dtype=str), mapping)
            df_mbuy = unify_df(pd.read_excel(file_m_buy, dtype=str), mapping)
            df_msell = unify_df(pd.read_excel(file_m_sell, dtype=str), mapping)

            # --- 4. DATA CLEANING ---
            # Konversi angka yang dibutuhkan
            for df in [df_inv, df_mbuy, df_msell, df_risk]:
                for c in df.columns:
                    if 'amt' in c.lower() or 'vol' in c.lower() or 'qty' in c.lower() or 'value' in c.lower() or 'price' in c.lower() or 'avail' in c.lower():
                        df[c] = pd.to_numeric(df[c].astype(str).str.replace(',', '').str.replace('"', ''), errors='coerce').fillna(0)

            # Hitung Volume Formula di Invoice
            df_inv['vol_net_total'] = df_inv.groupby(['cid_key', 'stock_key'])['tot_vol'].transform(
                lambda x: (df_inv.loc[x.index, 'tot_vol'] * df_inv.loc[x.index, 'bors'].map({'B': 1, 'S': -1})).sum()
            )
            
            def get_vol_formula(row):
                total = row['vol_net_total']
                if total < 0: return total if row['bors'] == 'S' else 0
                elif total > 0: return total if row['bors'] == 'B' else 0
                return 0
            df_inv['Volume_Formula'] = df_inv.apply(get_vol_formula, axis=1)

            # --- 5. PROCESSING ---
            final_sheets = {}
            for side in ['B', 'S']:
                work_df = df_inv[df_inv['bors'] == side].copy()
                work_df = work_df.merge(df_sid, on='cid_key', how='left', suffixes=('', '_sidmaster'))
                
                pei_data = work_df[work_df['sid_key'].notna()].copy()
                not_pei = work_df[work_df['sid_key'].isna()].copy()

                if side == 'B':
                    pei_data = pei_data.merge(df_mbuy, on=['sid_key', 'stock_key'], how='left', suffixes=('', '_mbuy'))
                    # Ambil Available Qty dari Risk Parameter
                    risk_sub = df_risk[['stock_key', 'avail_key']].drop_duplicates('stock_key')
                    pei_data = pei_data.merge(risk_sub, on='stock_key', how='left')
                    val_col = 'AVAILABLE MARKET VALUE'
                else:
                    # Sesuaikan nama kolom Stock Code di Margin Sell
                    msell_ready = df_msell.copy()
                    if 'stock_key' not in msell_ready.columns:
                        # Cari kolom Stock Code di file sell (biasanya STOCK CODE)
                        msell_ready = msell_ready.rename(columns={'STOCK CODE': 'stock_key'})
                    
                    pei_data = pei_data.merge(msell_ready, on=['sid_key', 'stock_key'], how='left', suffixes=('', '_msell'))
                    pei_data['avail_key'] = 0
                    val_col = 'AVAILABLE SELL VALUE'

                # --- 6. MAPPING 17 KOLOM ---
                template = pd.DataFrame()
                template['SID'] = pei_data['sid_key']
                template['STOCK CODE'] = pei_data['stock_key']
                
                target_cols = ['MARGIN BUY QUANTITY', 'LOAN QUANTITY', 'AVAILABLE QUANTITY', 'CLOSING PRICE', 
                               'AVAILABLE MARKET VALUE', 'HAIRCUT', 'AVAILABLE COLLATERAL VALUE',
                               'REGULAR SELL QUANTITY', 'REPAYMENT QUANTITY', 'AVAILABLE SELL QUANTITY', 'AVAILABLE SELL VALUE']
                
                for c in target_cols:
                    # Cari kolom yang mirip jika tidak pas namanya
                    if c in pei_data.columns:
                        template[c] = pei_data[c]
                    else:
                        template[c] = 0

                template['B/S'] = side
                template['CID'] = pei_data['cid_key']
                template['Name'] = pei_data['Name'] if 'Name' in pei_data else ""
                template['Stock'] = pei_data['stock_key']
                
                # Volume & Validation
                def check_side(v):
                    if side == 'B' and float(v) < 0: return "ERROR: Wrong Side"
                    if side == 'S' and float(v) > 0: return "ERROR: Wrong Side"
                    return v
                template['Volume'] = pei_data['Volume_Formula'].apply(check_side)
                
                template['Value'] = pei_data[val_col] if val_col in pei_data else 0
                template['PEI (Risk/Porto)'] = pei_data['avail_key']

                # Logika NETT
                def get_nett(row):
                    p = row['Volume']
                    s = pd.to_numeric(row['PEI (Risk/Porto)'], errors='coerce') or 0
                    if p == "ERROR: Wrong Side" or p == 0: return ""
                    p_num = float(p)
                    if side == 'B':
                        return "LOAN PEI" if s > p_num else "LOAN PARTIAL"
                    else:
                        return "REPAY PEI" if abs(s) < abs(p_num) else "ALL STOCK REPAY"
                
                template['NETT'] = template.apply(get_nett, axis=1)

                not_pei_rows = pd.DataFrame({'CID': ['Not Client PEI'] * len(not_pei)})
                final_sheets[side] = pd.concat([template, not_pei_rows], ignore_index=True).fillna("")

            # --- 7. EXPORT ---
            st.success("✅ Berhasil! Silakan download.")
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_sheets['B'].to_excel(writer, index=False, sheet_name='Buy')
                final_sheets['S'].to_excel(writer, index=False, sheet_name='Sell')
            
            st.download_button("📥 Download Hasil_MNC.xlsx", output.getvalue(), "Hasil_MNC.xlsx")

    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
        st.warning("Periksa apakah nama kolom di sidebar sudah sesuai dengan file Excel Anda.")
