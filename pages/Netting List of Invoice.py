import streamlit as st
import pandas as pd
import io

st.title("📑 List of Invoice Netting - Logic Fixed")

uploaded_file = st.file_uploader("Upload Invoice CSV", type=['csv'])

if uploaded_file:
    try:
        # 1. Load Data (no_cust sebagai ID Client)
        df = pd.read_csv(uploaded_file, dtype={'no_cust': str, 'no_share': str})
        
        # 2. Pembersihan Angka
        for col in ['amt_pay', 'tot_vol']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace(',', '').str.replace('"', '')
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # 3. PERBAIKAN LOGIKA: Buy (+) , Sell (-)
        # Sesuai revisi: Buy - Sell
        df['amt_pay_net'] = df.apply(lambda x: x['amt_pay'] if x['bors'] == 'B' else -x['amt_pay'], axis=1)
        df['tot_vol_net'] = df.apply(lambda x: x['tot_vol'] if x['bors'] == 'B' else -x['tot_vol'], axis=1)
        
        # 4. Grouping Detail (Tetap menampilkan status B/S asli untuk audit)
        detail_report = df.groupby(['no_cust', 'no_share', 'bors']).agg({
            'tot_vol': 'sum',
            'amt_pay': 'sum'
        }).reset_index()

        # 5. Summary Final per Client (Hasil Netting Buy - Sell)
        # Menghitung total IDR dan total Volume bersih
        client_summary = df.groupby(['no_cust', 'no_share']).agg({
            'tot_vol_net': 'sum',
            'amt_pay_net': 'sum'
        }).reset_index()
        
        client_summary.rename(columns={
            'tot_vol_net': 'Net_Volume_Stock', 
            'amt_pay_net': 'Net_Amount_IDR'
        }, inplace=True)

        st.success("✅ Logika diperbaiki: Buy (+) dan Sell (-)")

        # 6. Generate Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            detail_report.to_excel(writer, index=False, sheet_name='Detail_B_S')
            client_summary.to_excel(writer, index=False, sheet_name='Final_Netting_per_Client')
        
        st.download_button(
            label="📥 Download Fixed Netting.xlsx",
            data=output.getvalue(),
            file_name="Fixed_Netting_MNCN.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
