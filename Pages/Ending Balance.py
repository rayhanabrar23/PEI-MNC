import streamlit as st
import pandas as pd
import io

def process_soa(uploaded_file):
    try:
        # 1. Baca data, paksa ClAcNo jadi string agar 00 tidak hilang
        df = pd.read_csv(uploaded_file, dtype={'ClAcNo': str})
        
        # 2. Pembersihan Angka (Menghapus koma ribuan agar bisa jadi float)
        # Kita cek jika kolom berupa string, kita hapus komanya.
        for col in ['DbAmount', 'CrAmount']:
            if df[col].dtype == 'object':
                df[col] = df[col].str.replace(',', '')
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # 3. Logika Netting per Client (ClAcNo)
        # Mengelompokkan berdasarkan nomor client
        summary = df.groupby('ClAcNo').agg({
            'DbAmount': 'sum',
            'CrAmount': 'sum'
        }).reset_index()

        # Rumus: Balance = Total Credit - Total Debit
        # Sesuai permintaan: Db itu negatif (pengurang), Cr itu positif (penambah)
        summary['Ending_Balance'] = summary['CrAmount'] - summary['DbAmount']
        
        # Merapikan tampilan angka di preview agar ada pemisah ribuan
        summary_display = summary.copy()
        for col in ['DbAmount', 'CrAmount', 'Ending_Balance']:
            summary_display[col] = summary_display[col].map('{:,.2f}'.format)

        return summary, summary_display

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses: {e}")
        return None, None

# --- UI Streamlit ---
st.title("📊 MNCN - SOA Generator")
st.info("Aplikasi ini akan merapikan data SOA dan menghitung Ending Balance per Client.")

uploaded_file = st.file_uploader("Upload file SOA kamu di sini", type=['csv'])

if uploaded_file:
    raw_res, display_res = process_soa(uploaded_file)
    
    if raw_res is not None:
        st.success("✅ Data berhasil diproses!")
        
        # Tampilkan Preview
        st.subheader("Preview Hasil Ending_Balance")
        st.dataframe(display_res, use_container_width=True)

        # Download Button Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            raw_res.to_excel(writer, index=False, sheet_name='Ending_Balance')
        
        st.download_button(
            label="📥 Download File Ending_Balance.xlsx",
            data=output.getvalue(),
            file_name="Ending_Balance.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
