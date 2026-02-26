import streamlit as st
import pandas as pd
import io

st.title("📑 List of Invoice Netting - Formula Replica")

uploaded_file = st.file_uploader("Upload Invoice CSV", type=['csv'])

if uploaded_file:
    try:
        # 1. Load Data
        df = pd.read_csv(uploaded_file, dtype={'no_cust': str, 'no_share': str, 'bors': str})
        
        # 2. Pembersihan Angka
        for col in ['amt_pay', 'tot_vol']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace(',', '').str.replace('"', '')
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # 3. Dasar Perhitungan (Netting)
        # amt_pay_net: Buy (+), Sell (-)
        df['amt_pay_net'] = df.apply(lambda x: x['amt_pay'] if x['bors'] == 'B' else -x['amt_pay'], axis=1)
        # vol_net: Buy (+), Sell (-)
        df['vol_net'] = df.apply(lambda x: x['tot_vol'] if x['bors'] == 'B' else -x['tot_vol'], axis=1)

        # --- SHEET 1: DETAIL BS PER SAHAM ---
        stock_detail_bs = df.groupby(['no_cust', 'no_share', 'bors']).agg({
            'tot_vol': 'sum',
            'amt_pay': 'sum'
        }).reset_index()

        # --- SHEET 2: TOTAL NET PER CLIENT ---
        client_final_net = df.groupby('no_cust').agg({'amt_pay_net': 'sum'}).reset_index()
        client_final_net.rename(columns={'amt_pay_net': 'Grand_Total_Net_IDR'}, inplace=True)

        # --- SHEET 3: REPLIKA FORMULA EXCEL (LOGIKA BARU) ---
        # A. Hitung Net Volume per Saham per Client
        net_vol_calc = df.groupby(['no_cust', 'no_share']).agg({'vol_net': 'sum'}).reset_index()
        
        # B. Gabungkan kembali ke detail agar tahu mana baris B dan mana baris S
        sheet3 = stock_detail_bs.copy()
        sheet3 = sheet3.merge(net_vol_calc, on=['no_cust', 'no_share'], how='left')
        
        # C. Terapkan Formula: Angka muncul hanya di status yang dominan
        def apply_formula_logic(row):
            total_net = row['vol_net']
            status = row['bors']
            
            if total_net < 0: # Kondisi Net Sell
                return total_net if status == 'S' else 0
            elif total_net > 0: # Kondisi Net Buy
                return total_net if status == 'B' else 0
            else:
                return 0

        sheet3['Volume_Formula'] = sheet3.apply(apply_formula_logic, axis=1)
        
        # Rapikan kolom Sheet 3 sesuai permintaan
        sheet3_final = sheet3[['no_cust', 'no_share', 'bors', 'tot_vol', 'Volume_Formula', 'amt_pay']]

        st.success("✅ Data Berhasil Diproses! Sheet 3 dengan Replika Formula sudah siap.")

        # 4. Generate Excel dengan 3 Sheet
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            stock_detail_bs.to_excel(writer, index=False, sheet_name='Detail_BS_per_Saham')
            client_final_net.to_excel(writer, index=False, sheet_name='Total_Net_Client')
            sheet3_final.to_excel(writer, index=False, sheet_name='Replika_Formula_Volume')
        
        st.download_button(
            label="📥 Download Netting Formula Version.xlsx",
            data=output.getvalue(),
            file_name="MNCN_Netting_Formula.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
