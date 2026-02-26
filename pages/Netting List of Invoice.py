import streamlit as st
import pandas as pd
import io

st.title("📑 List of Invoice Netting")

uploaded_file = st.file_uploader("Upload Invoice CSV", type=['csv'])

if uploaded_file:
    try:
        # 1. Load Data
        df = pd.read_csv(uploaded_file, dtype={'no_cust': str, 'no_share': str})
        
        # 2. Cleaning Angka
        for col in ['amt_pay', 'tot_vol']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace(',', '').str.replace('"', '')
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # 3. Logika Netting
        # Sell (+) Buy (-)
        df['net_amount'] = df.apply(lambda x: x['amt_pay'] if x['bors'] == 'S' else -x['amt_pay'], axis=1)
        
        # Grouping Emiten & Final Client
        emiten_netting = df.groupby(['no_cust', 'no_share']).agg({'tot_vol': 'sum', 'net_amount': 'sum'}).reset_index()
        client_final = emiten_netting.groupby('no_cust').agg({'net_amount': 'sum'}).reset_index()
        client_final['Status'] = client_final['net_amount'].apply(lambda x: 'Positif (Kredit)' if x >= 0 else 'Negatif (Debit)')

        st.success("✅ Data berhasil diproses!")

        # 4. Export ke Excel (Tanpa Preview)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            emiten_netting.to_excel(writer, index=False, sheet_name='Netting_Emiten')
            client_final.to_excel(writer, index=False, sheet_name='Final_Client')
        
        st.download_button(
            label="📥 Download Hasil Netting (Excel)",
            data=output.getvalue(),
            file_name="Netting_Invoice_MNCN.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")
