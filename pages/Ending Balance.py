import streamlit as st
import pandas as pd
import io

def process_soa(uploaded_file):
    try:
        # 1. Baca data, paksa ClAcNo jadi string agar 00 tidak hilang
        df = pd.read_csv(uploaded_file, dtype={'ClAcNo': str})
        
        # 2. Pembersihan Angka
        for col in ['DbAmount', 'CrAmount']:
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str).str.replace(',', '').str.replace('"', '')
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # 3. Logika Netting per Client (ClAcNo)
        summary = df.groupby('ClAcNo').agg({
            'DbAmount': 'sum',
            'CrAmount': 'sum'
        }).reset_index()

        summary['Ending_Balance'] = summary['CrAmount'] - summary['DbAmount']
        
        # Merapikan tampilan angka di preview
        summary_display = summary.copy()
        # Menggunakan format ribuan dengan koma dan 2 desimal
        for col in ['DbAmount', 'CrAmount', 'Ending_Balance']:
            summary_display[col] = summary_display[col].apply(lambda x: f"{x:,.2f}")

        return summary, summary_display

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses: {e}")
        return None, None

# --- UI Streamlit ---
st.set_page_config(page_title="MNCN - SOA Generator", layout="wide") # Layout lebar agar enak dilihat

st.title("📑 MNCN - SOA Generator")
st.info("Aplikasi ini akan merapikan data SOA dan menghitung Ending Balance per Client.")

uploaded_file = st.file_uploader("Upload file SOA kamu di sini", type=['csv'])

if uploaded_file:
    raw_res, display_res = process_soa(uploaded_file)
    
    if raw_res is not None:
        st.success("✅ Data berhasil diproses!")
        
        # --- Bagian Download (Pindah ke Atas) ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            raw_res.to_excel(writer, index=False, sheet_name='Ending_Balance')
        
        st.download_button(
            label="📥 Download File Ending_Balance.xlsx",
            data=output.getvalue(),
            file_name="Ending_Balance_MNCN.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.divider() # Garis pembatas agar rapi

        # --- Bagian Preview Lengkap dengan Pencarian ---
        st.subheader("Preview Hasil Ending Balance")
        st.write("💡 *Gunakan ikon kaca pembesar di pojok kanan tabel untuk mencari ID Client.*")
        
        # Menampilkan seluruh data nasabah tanpa pagination (gulir saja)
        # height=600 memberikan area scroll yang luas agar semua data bisa diakses
        st.dataframe(
            display_res, 
            use_container_width=True, 
            height=600 
        )
