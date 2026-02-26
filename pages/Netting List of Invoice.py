import streamlit as st
import pandas as pd
import io

st.title("📑 List of Invoice Netting - Final Summary")

uploaded_file = st.file_uploader("Upload Invoice CSV", type=['csv'])

if uploaded_file:
    try:
        # 1. Load Data
        df = pd.read_csv(uploaded_file, dtype={'no_cust': str, 'no_share': str})
        
        # 2. Pembersihan Angka
        for col in ['amt_pay', 'tot_vol']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace(',', '').str.replace('"', '')
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # 3. Logika: Buy (+) , Sell (-)
        df['amt_pay_net'] = df.apply(lambda x: x['amt_pay'] if x['bors'] == 'B' else -x['amt_pay'], axis=1)
        df['tot_vol_net'] = df.apply(lambda x: x['tot_vol'] if x['bors'] == 'B' else -x['tot_vol'], axis=1)
        
        # 4. Sheet 1: Netting per Saham (seperti di gambar kamu)
        stock_summary = df.groupby(['no_cust', 'no_share']).agg({
            'tot_vol_net': 'sum',
            'amt_pay_net': 'sum'
        }).reset_index()
        stock_summary.rename(columns={'tot_vol_net': 'Net_Volume_Stock', 'amt_pay_net': 'Net_Amount_IDR'}, inplace=True)

        # 5. Sheet 2: TOTAL NET PER CLIENT (Ini yang kamu minta)
        # Menjumlahkan semua Net_Amount_IDR dari semua saham per client
        client_total = stock_summary.groupby('no_cust').agg({
            'Net_Amount_IDR': 'sum'
        }).reset_index()
        client_total.rename(columns={'Net_Amount_IDR': 'Grand_Total_Net_IDR'}, inplace=True)
        
        # Tambah status biar jelas
        client_total['Status'] = client_total['Grand_Total_Net_IDR'].apply(
            lambda x: 'Net BUY (Bayar)' if x >= 0 else 'Net SELL (Terima Uang)'
        )

        st.success("✅ Data berhasil diproses dengan Ringkasan Client!")

        # 6. Generate Excel dengan 2 Sheet
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            stock_summary.to_excel(writer, index=False, sheet_name='Netting_per_Saham')
            client_total.to_excel(writer, index=False, sheet_name='Total_per_Client')
        
        st.download_button(
            label="📥 Download Hasil Netting Final.xlsx",
            data=output.getvalue(),
            file_name="Final_Netting_MNCN.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
