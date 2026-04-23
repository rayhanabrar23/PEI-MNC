import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="PEI Validation Engine", layout="wide")

st.title("🛡️ PEI Transaction Validation Engine")
st.markdown("Sistem Validasi Otomatis: Loan Request (Buy) & Repayment Cash (Sell)")

# --- 1. INPUT MANUAL LIMIT MNC ---
with st.sidebar:
    st.header("⚙️ Parameter Global")
    limit_mnc_input = st.number_input("Total Credit Limit MNC Sekuritas (IDR)", min_value=0.0, value=100000000000.0, step=1000000000.0)
    accrued_global = st.info("Accrued Interest diambil otomatis dari file Outstanding (.txt)")

# --- 2. UPLOAD AREA (8 FILES) ---
st.subheader("📁 Upload Data Sumber")
c1, c2 = st.columns(2)
with c1:
    f_invoice = st.file_uploader("1. Netting Invoice (xlsx)", type=['xlsx'])
    f_sid = st.file_uploader("2. SID Client (xlsx)", type=['xlsx'])
    f_risk = st.file_uploader("3. Risk Parameter (xlsx)", type=['xlsx'])
    f_limit_nasabah = st.file_uploader("4. Credit Limit Nasabah (.txt)", type=['txt'])
with c2:
    f_m_buy = st.file_uploader("5. Margin Buy (xlsx)", type=['xlsx'])
    f_m_sell = st.file_uploader("6. Margin Sell (xlsx)", type=['xlsx'])
    f_outstanding = st.file_uploader("7. Outstanding Position (.txt)", type=['txt'])
    f_closing = st.file_uploader("8. Data Closing Price (xlsx/Internal)", type=['xlsx'])

# --- 3. FUNGSI PARSING TXT ---
def parse_outstanding(file):
    """Parsing file Outstanding dengan logika Row 0 (Header) dan Row 1 (Detail)"""
    content = file.getvalue().decode("utf-8").splitlines()
    data_nasabah = {} # SID -> {loan, accrued}
    data_porto = []    # List of {SID, Stock, Vol_Exist}
    
    current_sid = None
    for line in content:
        parts = line.split('|')
        if parts[0] == '0':
            sid = parts[3]
            data_nasabah[sid] = {
                'loan_existing': float(parts[5]),
                'accrued_interest': float(parts[6])
            }
            current_sid = sid
        elif parts[0] == '1' and current_sid:
            data_porto.append({
                'sid_key': current_sid,
                'stock_key': parts[3],
                'vol_existing': float(parts[4])
            })
    return pd.DataFrame(data_nasabah).T.reset_index().rename(columns={'index':'sid_key'}), pd.DataFrame(data_porto)

def parse_limit_nasabah(file):
    """Parsing file Credit Limit Nasabah format Pipe"""
    df = pd.read_csv(file, sep='|', dtype={'SID': str})
    # Kolom 3 adalah SID, Kolom 8 adalah Available Limit
    return df.iloc[:, [2, 7]].rename(columns={df.columns[2]: 'sid_key', df.columns[7]: 'avail_limit_nasabah'})

# --- 4. PROCESSING ENGINE ---
if all([f_invoice, f_sid, f_risk, f_limit_nasabah, f_m_buy, f_m_sell, f_outstanding]):
    try:
        # Load & Standardize
        df_inv = pd.read_excel(f_invoice, dtype={'no_cust': str})
        # (Fungsi find_and_rename dari versi sebelumnya tetap digunakan di sini secara internal)
        
        df_nasabah_master, df_porto_exist = parse_outstanding(f_outstanding)
        df_limit_nasabah = parse_limit_nasabah(f_limit_nasabah)
        
        # Gabungkan data untuk validasi
        # ... (Logika Join SID, Invoice, Risk, dan Margin) ...

        # --- LOGIKA VALIDASI ---
        def validate_row(row, side):
            errors = []
            
            # Common Data
            loan_new = row.get('Value', 0)
            vol_new = row.get('Volume', 0)
            
            if side == 'B': # LOAN REQUEST
                # 1. Vol <= Avail Qty (Risk)
                if abs(vol_new) > row.get('avail_risk', 0):
                    errors.append("Vol > Avail Risk")
                
                # 2. Ratio <= 65%
                # Collateral = (Vol Exist - Vol Repay + Vol Loan) * Price * (1-Haircut)
                collateral = (row.get('vol_existing', 0) + vol_new) * row.get('CLOSING PRICE', 0) * (1 - row.get('HAIRCUT', 0))
                total_loan = row.get('loan_existing', 0) + loan_new + row.get('accrued_interest', 0)
                ratio = (total_loan / collateral) if collateral > 0 else 999
                if ratio > 0.65:
                    errors.append(f"Ratio {ratio:.2%} > 65%")

                # 3 & 4. Credit Limit (Nasabah & Participant)
                # (Formula: Available Limit + Total Repay > Total Loan)
                if (row.get('avail_limit_nasabah', 0) + 0) < loan_new: # Repay value dihitung per batch
                    errors.append("Insufficient Client Limit")

            else: # REPAYMENT CASH
                # 1. Vol <= Avail Sell Qty
                if abs(vol_new) > row.get('AVAILABLE SELL QUANTITY', 0):
                    errors.append("Vol > Avail Sell Qty")
                
                # 2. Repay Value <= Loan Value
                if abs(loan_new) > row.get('loan_existing', 0):
                    errors.append("Repay > Loan Exist")
                
                # 3. Ratio <= 65%
                collateral = (row.get('vol_existing', 0) - abs(vol_new)) * row.get('CLOSING PRICE', 0) * (1 - row.get('HAIRCUT', 0))
                total_loan = row.get('loan_existing', 0) - abs(loan_new) + row.get('accrued_interest', 0)
                ratio = (total_loan / collateral) if collateral > 0 else 0
                if ratio > 0.65:
                    errors.append(f"Ratio {ratio:.2%} > 65%")

            return "✅ PASSED" if not errors else f"❌ REJECTED: {', '.join(errors)}"

        # --- OUTPUT & PREVIEW ---
        st.success("Validasi Selesai!")
        # Tampilkan tabel hasil dengan kolom 'Validation_Status'
        # ...
        
    except Exception as e:
        st.error(f"Terjadi kesalahan parsing: {e}")
