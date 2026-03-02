import streamlit as st
import pandas as pd
import io

# 1. Konfigurasi Halaman
st.set_page_config(page_title="MNCN - Invoice Netting Complete", layout="wide")

st.title("📑 List of Invoice Netting")
st.info("Aplikasi ini mengolah data Invoice untuk netting per saham, total IDR client, dan ringkasan per emiten.")

uploaded_file = st.file_uploader("Upload Invoice CSV", type=['csv'])

if uploaded_file:
    try:
        # 2. Load Data
        df = pd.read_csv(uploaded_file, dtype={'no_cust': str, 'no_share': str, 'bors': str})
        
        # 3. Pembersihan Angka
        for col in ['amt_pay', 'tot_vol']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace(',', '').str.replace('"', '')
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # 4. Dasar Perhitungan Netting
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

        # --- SHEET 3: REPLIKA FORMULA EXCEL (Hanya di sisi dominan) ---
        net_vol_calc = df.groupby(['no_cust', 'no_share']).agg({'vol_net': 'sum'}).reset_index()
        sheet3 = stock_detail_bs.copy()
        sheet3 = sheet3.merge(net_vol_calc, on=['no_cust', 'no_share'], how='left')
        
        def apply_formula_logic(row):
            total_net = row['vol_net']
            status = row['bors']
            if total_net < 0: return total_net if status == 'S' else 0
            elif total_net > 0: return total_net if status == 'B' else 0
            else: return 0

        sheet3['Volume_Formula'] = sheet3.apply(apply_formula_logic, axis=1)
        sheet3_final = sheet3[['no_cust', 'no_share', 'bors', 'tot_vol', 'Volume_Formula', 'amt_pay']]

        # --- SHEET 4: NETTING PER EMITEN (Murni Per Saham) ---
        # Mengelompokkan berdasarkan client dan saham saja (B/S digabung)
        net_emiten = df.groupby(['no_cust', 'no_share']).agg({
            'vol_net': 'sum',
            'amt_pay_net': 'sum'
        }).reset_index()
        
        # Rename kolom sesuai dengan screenshot yang kamu kirim
        net_emiten.rename(columns={
            'vol_net': 'Net_Volume_Stock', 
            'amt_pay_net': 'Net_Amount_IDR'
        }, inplace=True)

        st.success("✅ Data Berhasil Diproses!")

        # 5. Tombol Download di Atas
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            stock_detail_bs.to_excel(writer, index=False, sheet_name='Detail_BS_per_Saham')
            client_final_net.to_excel(writer, index=False, sheet_name='Total_per_Client')
            sheet3_final.to_excel(writer, index=False, sheet_name='Replika_Formula_Volume')
            net_emiten.to_excel(writer, index=False, sheet_name='Netting_per_Saham')
        
        st.download_button(
            label="📥 Download Netting Complete Version.xlsx",
            data=output.getvalue(),
            file_name="MNCN_Netting_Complete.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.divider()
        
        # 6. Preview Sheet Baru (Netting per Saham)
        st.subheader("Preview: Netting per Saham (Buy - Sell)")
        st.write("💡 *Gunakan fitur search untuk memfilter ID Client atau Kode Saham.*")
        
        # Menampilkan hasil netting per emiten dengan format ribuan di preview
        net_emiten_display = net_emiten.copy()
        for col in ['Net_Volume_Stock', 'Net_Amount_IDR']:
            net_emiten_display[col] = net_emiten_display[col].map('{:,.2f}'.format)
            
        st.dataframe(net_emiten_display, use_container_width=True, height=500)

    except Exception as e:
        # Menangani error jika kolom yang diperlukan tidak ada (seperti pada gambar error 'bors')
        st.error(f"Terjadi kesalahan: Pastikan kolom CSV sudah sesuai. Error: {e}")
