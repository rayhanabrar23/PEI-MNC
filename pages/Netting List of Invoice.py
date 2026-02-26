import streamlit as st
import pandas as pd
import io

st.title("📑 List of Invoice Netting - Revised")

uploaded_file = st.file_uploader("Upload Invoice CSV", type=['csv'])

if uploaded_file:
    try:
        # 1. Load Data & Rename kolom sesuai permintaan
        # Menggunakan no_cust sebagai identifier utama
        df = pd.read_csv(uploaded_file, dtype={'no_cust': str, 'no_share': str})
        
        # 2. Pembersihan Angka
        for col in ['amt_pay', 'tot_vol']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace(',', '').str.replace('"', '')
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # 3. Logika Perhitungan Baru
        # amt_pay: Sell (+) , Buy (-) -> untuk cari sisa uang/balance
        df['amt_pay_net'] = df.apply(lambda x: x['amt_pay'] if x['bors'] == 'S' else -x['amt_pay'], axis=1)
        
        # tot_vol: Buy (+) , Sell (-) -> untuk cari sisa stok barang
        df['tot_vol_net'] = df.apply(lambda x: -x['tot_vol'] if x['bors'] == 'S' else x['tot_vol'], axis=1)
        
        # 4. Grouping per Client + per Emiten + per Status (B/S)
        # Agar di Excel terlihat detail B/S nya
        detail_report = df.groupby(['no_cust', 'no_share', 'bors']).agg({
            'tot_vol': 'sum',
            'amt_pay': 'sum',
            'amt_pay_net': 'sum',
            'tot_vol_net': 'sum'
        }).reset_index()

        # 5. Summary Final per Client (Total IDR saja)
        client_summary = df.groupby('no_cust').agg({
            'amt_pay_net': 'sum'
        }).reset_index()
        client_summary.rename(columns={'amt_pay_net': 'Total_IDR_Balance'}, inplace=True)

        st.success("✅ Revisi perhitungan berhasil diproses!")

        # 6. Generate Excel dengan 2 Sheet
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Sheet 1: Detail per transaksi yang sudah dikelompokkan
            detail_report.to_excel(writer, index=False, sheet_name='Detail_Netting')
            # Sheet 2: Ringkasan saldo uang per Client
            client_summary.to_excel(writer, index=False, sheet_name='Summary_Client_IDR')
        
        st.download_button(
            label="📥 Download Revised Netting.xlsx",
            data=output.getvalue(),
            file_name="Revised_Netting_MNCN.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
