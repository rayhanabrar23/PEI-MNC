import streamlit as st
import pandas as pd
import io

st.title("📑 List of Invoice Netting - Detailed B/S")

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

        # 3. Logika Perhitungan Netting untuk Summary Akhir
        # Buy (+) , Sell (-)
        df['amt_pay_net'] = df.apply(lambda x: x['amt_pay'] if x['bors'] == 'B' else -x['amt_pay'], axis=1)
        
        # 4. Sheet 1: Detail per Saham & per Status B/S (Permintaanmu)
        # Kita masukkan 'bors' ke dalam groupby agar B dan S muncul terpisah
        stock_detail_bs = df.groupby(['no_cust', 'no_share', 'bors']).agg({
            'tot_vol': 'sum',
            'amt_pay': 'sum'
        }).reset_index()
        
        # Tambahkan kolom penanda nilai net di sheet detail agar informatif
        stock_detail_bs['Net_Value_IDR'] = stock_detail_bs.apply(
            lambda x: x['amt_pay'] if x['bors'] == 'B' else -x['amt_pay'], axis=1
        )

        # 5. Sheet 2: Total Bersih per ID Client (Satu baris per Client)
        client_final_net = df.groupby('no_cust').agg({
            'amt_pay_net': 'sum'
        }).reset_index()
        
        client_final_net.rename(columns={'amt_pay_net': 'Grand_Total_Net_IDR'}, inplace=True)
        client_final_net['Status_Posisi'] = client_final_net['Grand_Total_Net_IDR'].apply(
            lambda x: 'NET BUY (Wajib Bayar)' if x >= 0 else 'NET SELL (Terima Dana)'
        )

        st.success("✅ Selesai! Detail B/S ada di Sheet 1 & Total Net Client di Sheet 2.")

        # 6. Generate Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Sheet 1: Tidak langsung netting emiten, tapi pisah B / S
            stock_detail_bs.to_excel(writer, index=False, sheet_name='Detail_BS_per_Saham')
            # Sheet 2: Total satu client satu baris
            client_final_net.to_excel(writer, index=False, sheet_name='Total_Net_Client')
        
        st.download_button(
            label="📥 Download Hasil Netting B-S Lengkap.xlsx",
            data=output.getvalue(),
            file_name="Netting_MNCN_Detailed_BS.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Terjadi kesalahan saat proses: {e}")
