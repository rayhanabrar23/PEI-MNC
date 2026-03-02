import streamlit as st
import pandas as pd
import io

# 1. Konfigurasi halaman agar tampilan lebar (Wide Mode)
st.set_page_config(page_title="List of Invoice Netting", layout="wide")

# 2. Header dan Penjelasan Halaman (Mirip SOA Generator)
st.title("📑 List of Invoice Netting")
st.info("Aplikasi ini akan mengolah data Invoice untuk menghitung netting per saham (B/S), total IDR per client, serta menerapkan formula volume khusus.")

# 3. Fitur Upload File
uploaded_file = st.file_uploader("Upload Invoice CSV", type=['csv'])

if uploaded_file:
    try:
        # 4. Load Data dengan paksa string pada ID agar 0 di depan tidak hilang
        df = pd.read_csv(uploaded_file, dtype={'no_cust': str, 'no_share': str, 'bors': str})
        
        # 5. Pembersihan Angka (Menghapus koma/kutip dan konversi ke angka)
        for col in ['amt_pay', 'tot_vol']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace(',', '').str.replace('"', '')
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # 6. Dasar Perhitungan Netting
        # amt_pay_net: Buy (+), Sell (-) untuk perhitungan uang
        df['amt_pay_net'] = df.apply(lambda x: x['amt_pay'] if x['bors'] == 'B' else -x['amt_pay'], axis=1)
        # vol_net: Buy (+), Sell (-) untuk perhitungan sisa barang
        df['vol_net'] = df.apply(lambda x: x['tot_vol'] if x['bors'] == 'B' else -x['tot_vol'], axis=1)

        # --- SHEET 1: DETAIL BS PER SAHAM ---
        stock_detail_bs = df.groupby(['no_cust', 'no_share', 'bors']).agg({
            'tot_vol': 'sum',
            'amt_pay': 'sum'
        }).reset_index()

        # --- SHEET 2: TOTAL NET PER CLIENT (Grand Total IDR) ---
        client_final_net = df.groupby('no_cust').agg({'amt_pay_net': 'sum'}).reset_index()
        client_final_net.rename(columns={'amt_pay_net': 'Grand_Total_Net_IDR'}, inplace=True)
        
        # Dataframe khusus untuk preview (dengan format ribuan)
        client_display = client_final_net.copy()
        client_display['Grand_Total_Net_IDR'] = client_display['Grand_Total_Net_IDR'].map('{:,.2f}'.format)

        # --- SHEET 3: REPLIKA FORMULA EXCEL ---
        # A. Hitung Net Volume per Saham per Client
        net_vol_calc = df.groupby(['no_cust', 'no_share']).agg({'vol_net': 'sum'}).reset_index()
        
        # B. Gabungkan kembali ke detail
        sheet3 = stock_detail_bs.copy()
        sheet3 = sheet3.merge(net_vol_calc, on=['no_cust', 'no_share'], how='left')
        
        # C. Fungsi Logika Formula Volume (Hanya muncul di sisi yang dominan)
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
        sheet3_final = sheet3[['no_cust', 'no_share', 'bors', 'tot_vol', 'Volume_Formula', 'amt_pay']]
        
        # Dataframe khusus untuk preview sheet 3 (dengan format ribuan)
        sheet3_display = sheet3_final.copy()
        for col in ['tot_vol', 'Volume_Formula', 'amt_pay']:
            sheet3_display[col] = sheet3_display[col].map('{:,.2f}'.format)

        st.success("✅ Data Berhasil Diproses!")

        # 7. GENERATE EXCEL (Download Button di taruh di ATAS preview)
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

        st.divider() # Garis pemisah

        # 8. BAGIAN PREVIEW (Dua kolom berdampingan)
        st.write("💡 *Gunakan fitur search (kaca pembesar) di pojok kanan tabel untuk mencari data.*")
        
        col1, col2 = st.columns([1, 2])

        with col1:
            st.subheader("Total Net per Client")
            st.dataframe(client_display, use_container_width=True, height=500)

        with col2:
            st.subheader("Recap All Data")
            st.dataframe(sheet3_display, use_container_width=True, height=500)

    except Exception as e:
        st.error(f"Terjadi kesalahan saat pemrosesan: {e}")
