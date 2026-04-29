import streamlit as st
import pandas as pd
import io

# 1. Konfigurasi Halaman
st.set_page_config(page_title="MNC - Invoice Netting", layout="wide")

st.title("📑 List of Invoice Netting")
st.info("Halaman ini memproses data Invoice. Pastikan file CSV memiliki kolom: no_cust, no_share, bors, amt_pay, dan tot_vol.")

# 2. Fitur Upload
uploaded_file = st.file_uploader("Upload Invoice CSV", type=['csv'])

if uploaded_file:
    try:
        # Load data
        df = pd.read_csv(uploaded_file, dtype={'no_cust': str, 'no_share': str, 'bors': str})
        
        # --- VALIDASI KOLOM ---
        required_columns = ['no_cust', 'no_share', 'bors', 'amt_pay', 'tot_vol']
        missing_cols = [col for col in required_columns if col not in df.columns]
        
        if missing_cols:
            st.error(f"❌ File salah! Kolom berikut tidak ditemukan: {', '.join(missing_cols)}")
            st.info("Tips: Pastikan kamu tidak mengupload file SOA di halaman ini.")
            st.stop()

        # 3. Pembersihan Angka
        for col in ['amt_pay', 'tot_vol']:
            df[col] = df[col].astype(str).str.replace(',', '').str.replace('"', '')
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # 4. Dasar Perhitungan Netting
        df['amt_pay_net'] = df.apply(lambda x: x['amt_pay'] if x['bors'] == 'B' else -x['amt_pay'], axis=1)
        df['vol_net'] = df.apply(lambda x: x['tot_vol'] if x['bors'] == 'B' else -x['tot_vol'], axis=1)

        # --- PROSES 4 SHEET ---
        
        # SHEET 1: DETAIL BS PER SAHAM
        stock_detail_bs = df.groupby(['no_cust', 'no_share', 'bors']).agg({'tot_vol': 'sum', 'amt_pay': 'sum'}).reset_index()

        # SHEET 2: TOTAL PER CLIENT
        client_final_net = df.groupby('no_cust').agg({'amt_pay_net': 'sum'}).reset_index()
        client_final_net.rename(columns={'amt_pay_net': 'Grand_Total_Net_IDR'}, inplace=True)

        # SHEET 3: REPLIKA FORMULA VOLUME (Logika Dominan B/S)
        net_vol_calc = df.groupby(['no_cust', 'no_share']).agg({'vol_net': 'sum'}).reset_index()
        sheet3 = stock_detail_bs.copy().merge(net_vol_calc, on=['no_cust', 'no_share'], how='left')
        
        def apply_formula_logic(row):
            total_net = row['vol_net']
            status = row['bors']
            if total_net < 0: return total_net if status == 'S' else 0
            elif total_net > 0: return total_net if status == 'B' else 0
            else: return 0

        sheet3['Volume_Formula'] = sheet3.apply(apply_formula_logic, axis=1)
        sheet3_final = sheet3[['no_cust', 'no_share', 'bors', 'tot_vol', 'Volume_Formula', 'amt_pay']]

        # SHEET 4: NETTING PER EMITEN (Murni Per Saham)
        net_emiten = df.groupby(['no_cust', 'no_share']).agg({'vol_net': 'sum', 'amt_pay_net': 'sum'}).reset_index()
        net_emiten.rename(columns={'vol_net': 'Net_Volume_Stock', 'amt_pay_net': 'Net_Amount_IDR'}, inplace=True)

        # --- OUTPUT & DOWNLOAD ---
        st.success("✅ Data Invoice Berhasil Diproses!")

        # Buffer 1: Complete (4 sheet)
        output_complete = io.BytesIO()
        with pd.ExcelWriter(output_complete, engine='openpyxl') as writer:
            stock_detail_bs.to_excel(writer, index=False, sheet_name='Detail_BS_per_Saham')
            client_final_net.to_excel(writer, index=False, sheet_name='Total_per_Client')
            sheet3_final.to_excel(writer, index=False, sheet_name='Netting_Volume')
            net_emiten.to_excel(writer, index=False, sheet_name='Netting_per_Saham')

        # Buffer 2: Khusus Sheet 3 saja
        output_netting_vol = io.BytesIO()
        with pd.ExcelWriter(output_netting_vol, engine='openpyxl') as writer:
            sheet3_final.to_excel(writer, index=False, sheet_name='Netting_Volume')

        dl_col1, dl_col2 = st.columns(2)
        with dl_col1:
            st.download_button(
                label="📥 Download Netting Complete Version.xlsx",
                data=output_complete.getvalue(),
                file_name="MNCN_Netting_Complete.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with dl_col2:
            st.download_button(
                label="📥 Download Netting Volume Only.xlsx",
                data=output_netting_vol.getvalue(),
                file_name="Netting Invoice.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        st.divider()
        
        # --- 5. PREVIEW DUA TABEL (Sheet 3 & Sheet 4) ---
        st.write("💡 *Gunakan fitur search (ikon kaca pembesar) pada tabel untuk memfilter data.*")
        
        col1, col2 = st.columns(2)

        with col1:
            st.subheader("Total Buy & Sell")
            s3_display = sheet3_final.copy()
            for col in ['tot_vol', 'Volume_Formula', 'amt_pay']:
                s3_display[col] = s3_display[col].map('{:,.2f}'.format)
            st.dataframe(s3_display, use_container_width=True, height=500)

        with col2:
            st.subheader("Netting Stock & Cash")
            net_display = net_emiten.copy()
            for col in ['Net_Volume_Stock', 'Net_Amount_IDR']:
                net_display[col] = net_display[col].map('{:,.2f}'.format)
            st.dataframe(net_display, use_container_width=True, height=500)

    except Exception as e:
        st.error(f"Terjadi kesalahan sistem: {e}")
