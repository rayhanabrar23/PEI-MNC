import streamlit as st
import pandas as pd
import io

st.title("📊 TRX PEI Details")

# Instruksi untuk user
st.info("Halaman ini mengolah data Invoice untuk menghasilkan kolom khusus: Client, Efek, Volume (Net), dan Value.")

uploaded_file = st.file_uploader("Upload Invoice CSV yang sama", type=['csv'], key="pei_details")

if uploaded_file:
    try:
        # 1. Load Data
        df = pd.read_csv(uploaded_file, dtype={'no_cust': str, 'no_share': str, 'bors': str})
        
        # 2. Pembersihan Angka
        for col in ['amt_pay', 'tot_vol']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace(',', '').str.replace('"', '')
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # 3. Hitung Volume_Formula (Logika Netting yang sama dengan halaman sebelumnya)
        # vol_net: Buy (+), Sell (-)
        df['vol_net'] = df.apply(lambda x: x['tot_vol'] if x['bors'] == 'B' else -x['tot_vol'], axis=1)
        
        # Grouping untuk mendapatkan total net volume per saham per client
        net_vol_calc = df.groupby(['no_cust', 'no_share']).agg({'vol_net': 'sum'}).reset_index()
        
        # Grouping detail awal (per B/S)
        detail_bs = df.groupby(['no_cust', 'no_share', 'bors']).agg({
            'tot_vol': 'sum',
            'amt_pay': 'sum'
        }).reset_index()
        
        # Gabungkan data
        merged_data = detail_bs.merge(net_vol_calc, on=['no_cust', 'no_share'], how='left')
        
        # Terapkan Logika Volume_Formula
        def apply_volume_formula(row):
            total_net = row['vol_net']
            status = row['bors']
            if total_net < 0: # Net Sell
                return total_net if status == 'S' else 0
            elif total_net > 0: # Net Buy
                return total_net if status == 'B' else 0
            else:
                return 0

        merged_data['Volume_Formula'] = merged_data.apply(apply_volume_formula, axis=1)

        # 4. FILTER KOLOM SESUAI PERMINTAAN: 
        # Kolom: no_cust, no_share, Volume_Formula, amt_pay
        trx_pei_details = merged_data[['no_cust', 'no_share', 'Volume_Formula', 'amt_pay']].copy()
        
        # Rename agar judul kolom lebih rapi di Excel
        trx_pei_details.columns = ['Client Number', 'Kode Efek', 'Volume', 'Value']

        st.success("✅ Data TRX PEI Details berhasil dibuat!")

        # 5. Tampilkan Preview Singkat (10 data teratas saja biar gak panjang)
        st.subheader("Preview TRX PEI Details")
        st.dataframe(trx_pei_details.head(10), use_container_width=True)

        # 6. Tombol Download Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            trx_pei_details.to_excel(writer, index=False, sheet_name='TRX_PEI_Details')
        
        st.download_button(
            label="📥 Download TRX PEI Details.xlsx",
            data=output.getvalue(),
            file_name="TRX_PEI_Details.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error pada proses TRX PEI: {e}")
