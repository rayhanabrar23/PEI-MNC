import streamlit as st
import pandas as pd
import io

st.title("📑 TRX PEI Details - With Loan Logic")

uploaded_file = st.file_uploader("Upload Invoice CSV", type=['csv'], key="pei_loan")

if uploaded_file:
    try:
        # 1. Load Data
        df = pd.read_csv(uploaded_file, dtype={'no_cust': str, 'no_share': str, 'bors': str})
        
        # 2. Cleaning
        for col in ['amt_pay', 'tot_vol']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace(',', '').str.replace('"', '')
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # 3. Hitung Volume_Formula (Netting Logic)
        df['vol_net'] = df.apply(lambda x: x['tot_vol'] if x['bors'] == 'B' else -x['tot_vol'], axis=1)
        net_vol_calc = df.groupby(['no_cust', 'no_share']).agg({'vol_net': 'sum'}).reset_index()
        
        detail_bs = df.groupby(['no_cust', 'no_share', 'bors']).agg({
            'tot_vol': 'sum',
            'amt_pay': 'sum'
        }).reset_index()
        
        merged_data = detail_bs.merge(net_vol_calc, on=['no_cust', 'no_share'], how='left')
        
        def apply_volume_formula(row):
            total_net = row['vol_net']
            status = row['bors']
            if total_net < 0: return total_net if status == 'S' else 0
            elif total_net > 0: return total_net if status == 'B' else 0
            else: return 0

        merged_data['Volume_Formula'] = merged_data.apply(apply_volume_formula, axis=1)

        # 4. REPLIKA FORMULA EXCEL BARU (Loan/Repay Logic)
        # Catatan: Karena kita tidak punya table VLOOKUP eksternal, 
        # saya asumsikan semua share diproses. Jika kamu punya list khusus Margin, 
        # kita bisa tambahkan filternya nanti.
        
        def apply_loan_logic(row):
            p4 = row['Volume_Formula'] # Volume Net
            s4 = row['amt_pay']        # Value Amount Pay
            
            if p4 == 0:
                return ""
            
            # Logika jika Net Buy (P4 > 0)
            if p4 > 0:
                if s4 < 0: return "" # Safety check
                return "LOAN PEI" if s4 > p4 else "LOAN PARTIAL"
            
            # Logika jika Net Sell (P4 < 0)
            else:
                if s4 == 0: return ""
                return "REPAY PEI" if s4 < p4 else "ALL STOCK REPAY"

        merged_data['Loan_Status'] = merged_data.apply(apply_loan_logic, axis=1)

        # 5. Output Final
        trx_pei_details = merged_data[['no_cust', 'no_share', 'Volume_Formula', 'amt_pay', 'Loan_Status']].copy()
        trx_pei_details.columns = ['Client Number', 'Kode Efek', 'Volume', 'Value', 'Status Loan/Repay']

        st.success("✅ TRX PEI Details dengan Logika Loan Berhasil Dibuat!")
        st.dataframe(trx_pei_details.head(20))

        # 6. Download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            trx_pei_details.to_excel(writer, index=False, sheet_name='TRX_PEI_Details')
        
        st.download_button(
            label="📥 Download TRX PEI Loan Details.xlsx",
            data=output.getvalue(),
            file_name="TRX_PEI_Loan_Details.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")
