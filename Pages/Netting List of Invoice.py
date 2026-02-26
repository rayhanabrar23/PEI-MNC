import streamlit as st
import pandas as pd
import io

# Konfigurasi Halaman
st.set_page_config(page_title="MNCN Project Dashboard", layout="wide")

# --- FUNGSI HELPER ---
def clean_numeric(df, columns):
    for col in columns:
        if col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].str.replace(',', '').str.replace('"', '')
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    return df

# --- HALAMAN 1: SOA GENERATOR ---
def show_soa_page():
    st.title("📂 SOA Statement of Account")
    uploaded_file = st.file_uploader("Upload SOA CSV", type=['csv'], key="soa")
    
    if uploaded_file:
        df = pd.read_csv(uploaded_file, dtype={'ClAcNo': str})
        df = clean_numeric(df, ['DbAmount', 'CrAmount'])
        
        # Netting
        summary = df.groupby('ClAcNo').agg({'DbAmount':'sum', 'CrAmount':'sum'}).reset_index()
        summary['Ending_Balance'] = summary['CrAmount'] - summary['DbAmount']
        
        st.success("Proses Selesai!")
        st.subheader("Preview Ending Balance")
        st.dataframe(summary.style.format({'DbAmount': '{:,.2f}', 'CrAmount': '{:,.2f}', 'Ending_Balance': '{:,.2f}'}))

# --- HALAMAN 2: LIST OF INVOICE NETTING ---
def show_invoice_page():
    st.title("📑 List of Invoice Netting")
    uploaded_file = st.file_uploader("Upload Invoice CSV", type=['csv'], key="inv")
    
    if uploaded_file:
        # 1. Load & Clean (No_inv as string)
        df = pd.read_csv(uploaded_file, dtype={'no_inv': str, 'no_share': str})
        df = clean_numeric(df, ['amt_pay', 'tot_vol'])
        
        # 2. Detail Buy & Sell (Raw Data Preview)
        st.subheader("1. Detail Transaksi (Buy/Sell)")
        st.dataframe(df[['no_inv', 'bors', 'no_share', 'tot_vol', 'amt_pay']], use_container_width=True)

        # 3. Netting per Emiten
        # Aturan: Sell (+), Buy (-)
        df['net_amount'] = df.apply(lambda x: x['amt_pay'] if x['bors'] == 'S' else -x['amt_pay'], axis=1)
        
        emiten_netting = df.groupby(['no_inv', 'no_share']).agg({
            'tot_vol': 'sum',
            'net_amount': 'sum'
        }).reset_index()
        
        st.subheader("2. Netting per Emiten")
        st.dataframe(emiten_netting.style.format({'net_amount': '{:,.2f}'}), use_container_width=True)

        # 4. Final Netting per Client
        client_final = emiten_netting.groupby('no_inv').agg({'net_amount': 'sum'}).reset_index()
        client_final['Status'] = client_final['net_amount'].apply(lambda x: 'Positif (Kredit)' if x >= 0 else 'Negatif (Debit)')
        
        st.subheader("3. Final Netting per Client")
        st.table(client_final.style.format({'net_amount': '{:,.2f}'}))

        # Download Result
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            emiten_netting.to_excel(writer, index=False, sheet_name='Netting_Emiten')
            client_final.to_excel(writer, index=False, sheet_name='Final_Client')
        
        st.download_button("📥 Download Ending_Balance_Invoice.xlsx", output.getvalue(), "Ending_Balance_Invoice.xlsx")

# --- NAVIGASI SIDEBAR ---
page = st.sidebar.radio("Pilih Halaman:", ["SOA Generator", "List of Invoice Netting"])

if page == "SOA Generator":
    show_soa_page()
else:
    show_invoice_page()
