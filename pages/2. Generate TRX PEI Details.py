import streamlit as st
import pandas as pd
import io

# ─────────────────────────────────────────────
# 1. KONFIGURASI HALAMAN
# ─────────────────────────────────────────────
st.set_page_config(page_title="TRX PEI Details", layout="wide", page_icon="📑")

st.title("📑 TRX PEI Details Generator")
st.info(
    "Sistem ini memproses transaksi nasabah PEI: "
    "**Buy (Loan)** · **Sell (Repayment)** · "
    "**Repayment Proceed** · **Loan Request**"
)

# ─────────────────────────────────────────────
# 2. FUNGSI UTILITAS
# ─────────────────────────────────────────────

def find_and_rename(df):
    mapping = {
        'stock_key':  ['no_share', 'no_shares', 'Stock Code', 'Stockcode', 'Stock', 'SYMBOL', 'StockCode'],
        'sid_key':    ['SID', 'SID_No', 'Client_SID'],
        'cid_key':    ['no_cust', 'CID', 'Client_ID', 'Account_No'],
        'avail_risk': ['Available Quantity', 'availablequantity', 'Available Qty', 'AvailableQuantity'],
        'name_key':   ['Name', 'Client_Name', 'Nama'],
        'haircut_key':['Haircut', 'haircut', 'HC'],
        'margin_flag':['Margin', 'margin', 'MARGIN', 'flag_margin', 'MarginFlag'],
    }
    rename_dict = {}
    for official, aliases in mapping.items():
        for col in df.columns:
            if str(col).strip() in aliases:
                rename_dict[col] = official
                break
    return df.rename(columns=rename_dict)


def clean_num(df, extra_keys=None):
    import re
    num_keys = ['amt', 'vol', 'qty', 'val', 'price', 'avail',
                'haircut', 'collateral', 'quantity', 'margin']
    if extra_keys:
        num_keys += extra_keys

    def parse_number(s):
        s = str(s).strip().replace('"', '').replace('%', '')
        if s in ('', 'nan', 'None', '-'):
            return 0.0
        s = s.replace(',', '')
        if re.fullmatch(r'\d{1,3}(\.\d{3})+', s):
            s = s.replace('.', '')
        return pd.to_numeric(s, errors='coerce') or 0.0

    for c in df.columns:
        if any(k in str(c).lower() for k in num_keys):
            df[c] = df[c].apply(parse_number)

    return df


def fmt_rp(val):
    try:
        return f"Rp {int(val):,}".replace(',', '.')
    except Exception:
        return str(val)


def fmt_vol(val):
    try:
        return f"{int(val):,}".replace(',', '.')
    except Exception:
        return str(val)


def load_price_file(uploaded_file):
    """
    Load file harga Excel. Kolom yang diharapkan:
    STK_CODE (B) dan STK_CLOS (G) — atau deteksi otomatis.
    """
    df = pd.read_excel(uploaded_file, dtype=str)
    df.columns = df.columns.str.strip()

    # Cari kolom kode saham
    code_col = None
    for c in df.columns:
        if str(c).upper() in ['STK_CODE', 'STOCK_CODE', 'CODE', 'KODE']:
            code_col = c
            break
    if code_col is None:
        # fallback: kolom kedua (index 1)
        code_col = df.columns[1]

    # Cari kolom closing price
    price_col = None
    for c in df.columns:
        if str(c).upper() in ['STK_CLOS', 'CLOSE', 'CLOSING', 'CLOSING_PRICE', 'HARGA']:
            price_col = c
            break
    if price_col is None:
        # fallback: kolom G (index 6)
        price_col = df.columns[6] if len(df.columns) > 6 else df.columns[-1]

    price_map = {}
    for _, row in df.iterrows():
        code  = str(row[code_col]).strip().upper()
        price = pd.to_numeric(str(row[price_col]).replace(',', '').replace('.', ''), errors='coerce')
        # Handle format angka Indonesia: titik = ribuan
        if pd.isna(price):
            price = 0.0
        price_map[code] = float(price)

    return price_map


# ─────────────────────────────────────────────
# 3. AREA UPLOAD FILE
# ─────────────────────────────────────────────
st.subheader("📂 Upload File")
col_u1, col_u2, col_u3, col_u4 = st.columns(4)

with col_u1:
    file_invoice    = st.file_uploader("1. Netting Invoice (xlsx)",  type=['xlsx'])
    file_sid_client = st.file_uploader("2. SID Client (xlsx)",       type=['xlsx'])

with col_u2:
    file_risk  = st.file_uploader("3. Risk Parameter (.txt)",  type=['txt'])
    file_m_buy = st.file_uploader("4. Margin Buy (.txt)",      type=['txt'])

with col_u3:
    file_m_sell = st.file_uploader("5. Margin Sell (.txt)",    type=['txt'])
    file_ep     = st.file_uploader("6. Outstanding Position (.txt)", type=['txt'])

with col_u4:
    file_price = st.file_uploader("7. File Harga Saham (xlsx)", type=['xlsx'],
                                   help="File Excel dengan kolom STK_CODE dan STK_CLOS (Closing Price)")

# ─────────────────────────────────────────────
# 4. PROSES DATA
# ─────────────────────────────────────────────
required_files = [file_invoice, file_sid_client, file_risk, file_m_buy, file_m_sell, file_price]

if all(required_files):
    try:
        with st.spinner("⚙️ Memproses seluruh data..."):

            # ── 4a. LOAD & STANDARISASI ──────────────────────
            df_inv  = find_and_rename(pd.read_excel(file_invoice,    dtype=str))
            df_sid  = find_and_rename(pd.read_excel(file_sid_client, dtype=str))

            df_inv['stock_key'] = df_inv['stock_key'].astype(str).str.strip().str.upper()
            df_inv['cid_key']   = df_inv['cid_key'].astype(str).str.strip()
            df_sid['cid_key']   = df_sid['cid_key'].astype(str).str.strip()
            df_sid['sid_key']   = df_sid['sid_key'].astype(str).str.strip()

            # Risk Parameter — ambil margin flag dan avail_risk
            df_risk = find_and_rename(pd.read_csv(file_risk, sep='|', dtype=str))
            df_risk.columns = df_risk.columns.str.strip()
            df_risk = clean_num(df_risk)
            df_risk['stock_key'] = df_risk['stock_key'].astype(str).str.strip().str.upper()

            # ── STEP 1: Filter saham yang dicover PEI (flag margin) ──
            # Saham yang tidak ada di Risk Parameter atau margin_flag != 'Y'/'1'/'true'
            # dianggap tidak dicover PEI → take out dari proses
            margin_covered_stocks = set()
            if 'margin_flag' in df_risk.columns:
                margin_ok = df_risk[
                    df_risk['margin_flag'].astype(str).str.strip().str.upper().isin(['Y', '1', 'TRUE', 'YES'])
                ]
                margin_covered_stocks = set(margin_ok['stock_key'].unique())
                st.info(f"ℹ️ {len(margin_covered_stocks)} saham terdeteksi dicover margin PEI dari Risk Parameter.")
            else:
                # Jika kolom margin_flag tidak ada, semua saham di Risk Parameter dianggap covered
                margin_covered_stocks = set(df_risk['stock_key'].unique())
                st.warning("⚠️ Kolom flag margin tidak ditemukan di Risk Parameter — semua saham di Risk Parameter dianggap dicover PEI.")

            # Margin Buy
            df_mbuy = pd.read_csv(file_m_buy, sep='|', dtype=str).rename(columns={
                'SID':        'sid_key',
                'STOCK CODE': 'stock_key',
            })
            df_mbuy.columns = df_mbuy.columns.str.strip()
            df_mbuy = clean_num(df_mbuy)
            df_mbuy['stock_key'] = df_mbuy['stock_key'].astype(str).str.strip().str.upper()
            df_mbuy['sid_key']   = df_mbuy['sid_key'].astype(str).str.strip()

            # Margin Sell
            df_msell = pd.read_csv(file_m_sell, sep='|', dtype=str).rename(columns={
                'SID':        'sid_key',
                'STOCK CODE': 'stock_key',
            })
            df_msell.columns = df_msell.columns.str.strip()
            df_msell = clean_num(df_msell)
            df_msell['stock_key'] = df_msell['stock_key'].astype(str).str.strip().str.upper()
            df_msell['sid_key']   = df_msell['sid_key'].astype(str).str.strip()

            # File Harga
            price_map = load_price_file(file_price)

            # ── 4b. VOLUME FORMULA ────────────────────────────
            if 'Volume_Formula' not in df_inv.columns:
                df_inv['_sign'] = df_inv['bors'].map({'B': 1, 'S': -1}).fillna(0)
                df_inv['tot_vol'] = pd.to_numeric(
                    df_inv['tot_vol'].astype(str).str.replace(',', ''), errors='coerce'
                ).fillna(0)
                df_inv['_signed_vol'] = df_inv['tot_vol'] * df_inv['_sign']
                net_map = df_inv.groupby(['cid_key', 'stock_key'])['_signed_vol'].sum()
                df_inv['vol_net_total'] = df_inv.set_index(
                    ['cid_key', 'stock_key']).index.map(net_map)

                def calc_vol_formula(row):
                    total = row['vol_net_total']
                    vol   = row['tot_vol']
                    if total < 0:
                        return -vol if row['bors'] == 'S' else 0
                    elif total > 0:
                        return vol if row['bors'] == 'B' else 0
                    return 0

                df_inv['Volume_Formula'] = df_inv.apply(calc_vol_formula, axis=1)
            else:
                df_inv['Volume_Formula'] = pd.to_numeric(
                    df_inv['Volume_Formula'].astype(str).str.replace(',', ''),
                    errors='coerce'
                ).fillna(0)

            if 'amt_pay' in df_inv.columns:
                df_inv['amt_pay'] = pd.to_numeric(
                    df_inv['amt_pay'].astype(str).str.replace(',', ''),
                    errors='coerce'
                ).fillna(0)

            # ── 4c. FILTER NASABAH PEI ────────────────────────
            pei_cids = set(df_sid['cid_key'].astype(str).str.strip())

            df_sid_unique = df_sid.drop_duplicates(subset='cid_key', keep='first')
            sid_lookup = df_sid_unique.set_index('cid_key')[['sid_key', 'name_key']].to_dict('index')

            risk_sub = df_risk[['stock_key', 'avail_risk']].drop_duplicates('stock_key').copy()
            if 'haircut_key' in df_risk.columns:
                risk_sub = df_risk[['stock_key', 'avail_risk', 'haircut_key']].drop_duplicates('stock_key').copy()

            risk_lookup = risk_sub.set_index('stock_key').to_dict('index')

            # ── 4d. PORTOFOLIO KOLATERAL (dari OP file) ───────
            # Sekaligus ambil lot per (cid, stock) untuk hitung value = lot × price
            porto_sheets     = {}
            df_porto_all     = pd.DataFrame()
            porto_coll_lookup = {}   # (cid, stock) → volume
            op_lot_lookup     = {}   # (cid, stock) → lot dari OP (untuk repayment value)

            if file_ep is not None:
                content = file_ep.read().decode('utf-8')
                current_sid = None
                porto_rows  = []

                for line in content.strip().splitlines():
                    parts = line.strip().split('|')
                    if not parts:
                        continue
                    if parts[0] == '0':
                        current_sid = parts[3].strip() if len(parts) > 3 else None
                    elif parts[0] == '1' and current_sid:
                        if len(parts) < 5:
                            continue
                        stock = parts[3].strip().upper() if len(parts) > 3 else ''
                        vol   = pd.to_numeric(parts[4], errors='coerce') if len(parts) > 4 else 0
                        vol   = vol if pd.notna(vol) else 0
                        if stock and vol > 0:
                            sid_match = df_sid[df_sid['sid_key'].astype(str).str.strip() == current_sid]
                            if not sid_match.empty:
                                cid = str(sid_match['cid_key'].values[0]).strip()
                                porto_coll_lookup[(cid, stock)] = vol
                                op_lot_lookup[(cid, stock)]     = vol  # lot dari OP
                                porto_rows.append({
                                    'CID':       cid,
                                    'SID':       current_sid,
                                    'stock_key': stock,
                                    'coll_vol':  vol,
                                })

                if porto_rows:
                    df_porto_all = pd.DataFrame(porto_rows)
                    df_porto_all = df_porto_all.merge(
                        df_sid[['cid_key', 'name_key']].rename(
                            columns={'cid_key': 'CID', 'name_key': 'Name'}
                        ),
                        on='CID', how='left'
                    )
                    porto_sheets = dict(tuple(df_porto_all.groupby('CID')))

            st.session_state['porto_coll_lookup'] = porto_coll_lookup
            st.session_state['sid_cid_map'] = df_sid.set_index('sid_key')['cid_key'].to_dict()

            # ── 4e. SHEET BUY ─────────────────────────────────
            # STEP 1 filter: hanya saham yang dicover margin PEI
            df_buy_inv = df_inv[
                (df_inv['bors'] == 'B') &
                (df_inv['cid_key'].isin(pei_cids)) &
                (df_inv['Volume_Formula'].abs() > 0) &
                (df_inv['stock_key'].isin(margin_covered_stocks))   # ← STEP 1
            ].copy()

            if df_buy_inv.empty:
                buy_out = pd.DataFrame(columns=[
                    'SID','STOCK CODE','MARGIN BUY QUANTITY','LOAN QUANTITY',
                    'AVAILABLE QUANTITY','CLOSING PRICE','AVAILABLE MARKET VALUE',
                    'HAIRCUT','AVAILABLE COLLATERAL VALUE','B/S','CID','Name',
                    'Stock','Volume','Value (Lot×Price)','PEI (Risk/Porto)','NETT'
                ])
            else:
                if 'sid_key' not in df_buy_inv.columns:
                    buy_merged = df_buy_inv.merge(
                        df_sid[['cid_key', 'sid_key', 'name_key']], on='cid_key', how='left'
                    ).merge(df_mbuy, on=['sid_key', 'stock_key'], how='left', suffixes=('', '_m'))
                else:
                    buy_merged = df_buy_inv.merge(
                        df_mbuy, on=['sid_key', 'stock_key'], how='left', suffixes=('', '_m')
                    )

                buy_merged = buy_merged.merge(risk_sub, on='stock_key', how='left')

                # STEP 4: Value = Lot × Price (closing price dari file harga)
                def calc_buy_value(row):
                    cid   = str(row.get('cid_key', '')).strip()
                    stock = str(row.get('stock_key', '')).strip().upper()
                    lot   = abs(pd.to_numeric(row.get('Volume_Formula', 0), errors='coerce') or 0)
                    price = price_map.get(stock, 0)
                    return lot * price

                buy_out = pd.DataFrame()
                buy_out['SID']        = buy_merged.get('sid_key', '')
                buy_out['STOCK CODE'] = buy_merged['stock_key']
                for col in ['MARGIN BUY QUANTITY', 'LOAN QUANTITY', 'AVAILABLE QUANTITY',
                            'CLOSING PRICE', 'AVAILABLE MARKET VALUE', 'HAIRCUT',
                            'AVAILABLE COLLATERAL VALUE']:
                    buy_out[col] = buy_merged[col] if col in buy_merged.columns else 0
                buy_out['B/S']              = 'B'
                buy_out['CID']              = buy_merged['cid_key']
                buy_out['Name']             = buy_merged.get('name_key', '')
                buy_out['Stock']            = buy_merged['stock_key']
                buy_out['Volume']           = buy_merged['Volume_Formula'].abs()
                buy_out['Value (Lot×Price)'] = buy_merged.apply(calc_buy_value, axis=1)
                buy_out['PEI (Risk/Porto)'] = buy_merged.get('avail_risk', 0)

                def nett_buy(row):
                    p = pd.to_numeric(row['PEI (Risk/Porto)'], errors='coerce')
                    s = pd.to_numeric(row['Volume'], errors='coerce')
                    if pd.isna(p) or p == 0: return 'NON MARGIN'
                    if p > 0:
                        if pd.isna(s) or s < 0: return ''
                        return 'LOAN PEI' if s > p else 'LOAN PARTIAL'
                    return ''

                buy_out['NETT'] = buy_out.apply(nett_buy, axis=1)

            # ── 4f. SHEET SELL ────────────────────────────────
            # STEP 1 filter: hanya saham yang dicover margin PEI
            df_sell_inv = df_inv[
                (df_inv['bors'] == 'S') &
                (df_inv['cid_key'].isin(pei_cids)) &
                (df_inv['Volume_Formula'].abs() > 0) &
                (df_inv['stock_key'].isin(margin_covered_stocks))   # ← STEP 1
            ].copy()

            if df_sell_inv.empty:
                sell_out = pd.DataFrame(columns=[
                    'SID','STOCK CODE','REGULAR SELL QUANTITY','REPAYMENT QUANTITY',
                    'AVAILABLE SELL QUANTITY','CLOSING PRICE','AVAILABLE SELL VALUE',
                    'B/S','CID','Name','Stock','Volume',
                    'Value (Lot×Price)','PEI (Risk/Porto)','Ada di Kolateral','NETT'
                ])
            else:
                if 'sid_key' not in df_sell_inv.columns:
                    sell_merged = df_sell_inv.merge(
                        df_sid[['cid_key', 'sid_key', 'name_key']], on='cid_key', how='left'
                    ).merge(df_msell, on=['sid_key', 'stock_key'], how='left', suffixes=('', '_m'))
                else:
                    sell_merged = df_sell_inv.merge(
                        df_msell, on=['sid_key', 'stock_key'], how='left', suffixes=('', '_m')
                    )

                sell_merged = sell_merged.merge(risk_sub, on='stock_key', how='left')

                # STEP 3: Value = Lot (dari OP) × Closing Price (dari file harga)
                def calc_sell_value(row):
                    cid   = str(row.get('cid_key', '')).strip()
                    stock = str(row.get('stock_key', '')).strip().upper()
                    # Lot dari OP file; fallback ke Volume_Formula jika tidak ada di OP
                    lot   = op_lot_lookup.get((cid, stock),
                            abs(pd.to_numeric(row.get('Volume_Formula', 0), errors='coerce') or 0))
                    price = price_map.get(stock, 0)
                    return lot * price

                # STEP 3: Cek apakah saham ada di kolateral client (OP file)
                def cek_kolateral(row):
                    cid   = str(row.get('cid_key', '')).strip()
                    stock = str(row.get('stock_key', '')).strip().upper()
                    return 'YA' if (cid, stock) in porto_coll_lookup else 'TIDAK'

                def get_pei_sell(row):
                    cid   = str(row.get('cid_key', '')).strip()
                    stock = str(row.get('stock_key', '')).strip().upper()
                    vol   = porto_coll_lookup.get((cid, stock), 0)
                    return -vol

                sell_out = pd.DataFrame()
                sell_out['SID']        = sell_merged.get('sid_key', '')
                sell_out['STOCK CODE'] = sell_merged['stock_key']
                for col in ['REGULAR SELL QUANTITY', 'REPAYMENT QUANTITY',
                            'AVAILABLE SELL QUANTITY', 'CLOSING PRICE', 'AVAILABLE SELL VALUE']:
                    sell_out[col] = sell_merged[col] if col in sell_merged.columns else 0
                sell_out['B/S']              = 'S'
                sell_out['CID']              = sell_merged['cid_key']
                sell_out['Name']             = sell_merged.get('name_key', '')
                sell_out['Stock']            = sell_merged['stock_key']
                sell_out['Volume']           = sell_merged['Volume_Formula'].abs()
                sell_out['Value (Lot×Price)'] = sell_merged.apply(calc_sell_value, axis=1)
                sell_out['Ada di Kolateral'] = sell_merged.apply(cek_kolateral, axis=1)
                sell_out['PEI (Risk/Porto)'] = sell_merged.apply(get_pei_sell, axis=1)

                def nett_sell(row):
                    # STEP 3: Jika saham tidak ada di kolateral → baris ini di-exclude (tandai EXCLUDED)
                    if row['Ada di Kolateral'] == 'TIDAK':
                        return 'EXCLUDED (tidak ada di kolateral)'
                    p = pd.to_numeric(row['PEI (Risk/Porto)'], errors='coerce')
                    s = pd.to_numeric(row['Volume'], errors='coerce')
                    if pd.isna(p) or p == 0: return 'NON MARGIN'
                    if p < 0:
                        if pd.isna(s) or s == 0: return ''
                        return 'REPAY PEI' if s < abs(p) else 'ALL STOCK REPAY'
                    return ''

                sell_out['NETT'] = sell_out.apply(nett_sell, axis=1)

            # ── 4g. LOAN REQUEST ──────────────────────────────
            loan_req_rows = []
            if not df_porto_all.empty and 'coll_vol' in df_porto_all.columns:
                df_porto_risk = df_porto_all.merge(risk_sub, on='stock_key', how='left')

                for _, row in df_porto_risk.iterrows():
                    vol = pd.to_numeric(row.get('coll_vol', 0), errors='coerce') or 0
                    if vol > 0:
                        stock  = str(row.get('stock_key', '')).strip().upper()
                        price  = price_map.get(stock, 0)
                        hc_pct = pd.to_numeric(row.get('haircut_key', 5), errors='coerce') or 5
                        cv     = vol * price * (1 - hc_pct / 100)
                        loan_req_rows.append({
                            'CID':              row.get('CID', ''),
                            'Name':             row.get('Name', ''),
                            'Stock':            stock,
                            'HC (%)':           hc_pct,
                            'Collateral Vol':   vol,
                            'Price (STK_CLOS)': price,
                            'Collateral Value': cv,
                            'Loan Eligible':    cv,
                            'Status':           'STOCK DEPOSIT',
                        })

            df_loan_req = pd.DataFrame(loan_req_rows)

            # ── 4h. SUMMARY ───────────────────────────────────
            # Sell aktif = hanya yang ADA di kolateral (exclude baris EXCLUDED)
            sell_aktif = sell_out[sell_out['NETT'] != 'EXCLUDED (tidak ada di kolateral)'] if not sell_out.empty else sell_out

            summary_buy = buy_out.groupby(['CID', 'Name']).agg(
                Total_Volume=('Volume', 'sum'),
                Total_Value=('Value (Lot×Price)', 'sum'),
            ).reset_index() if not buy_out.empty else pd.DataFrame()

            summary_sell = sell_aktif.groupby(['CID', 'Name']).agg(
                Total_Volume=('Volume', 'sum'),
                Total_Value=('Value (Lot×Price)', 'sum'),
            ).reset_index() if not sell_aktif.empty else pd.DataFrame()

            # Hitung jumlah baris excluded
            n_excluded = len(sell_out[sell_out['NETT'] == 'EXCLUDED (tidak ada di kolateral)']) if not sell_out.empty else 0

        # ─────────────────────────────────────────────────────────────
        # 5. TAMPILKAN HASIL
        # ─────────────────────────────────────────────────────────────
        st.success("✅ Data Berhasil Diproses!")

        col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)
        col_m1.metric("Nasabah PEI Terdaftar",  len(pei_cids))
        col_m2.metric("Saham Dicover Margin",    len(margin_covered_stocks))
        col_m3.metric("Baris BUY (Loan)",        len(buy_out))
        col_m4.metric("Baris SELL (Repay)",       len(sell_aktif))
        col_m5.metric("Sell Excluded (tdk ada di kolateral)", n_excluded,
                       delta=f"-{n_excluded}" if n_excluded else None, delta_color="inverse")

        # ─────────────────────────────────────────────────────────────
        # 6. DOWNLOAD EXCEL
        # ─────────────────────────────────────────────────────────────
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as wr:
            buy_out.to_excel(wr,       index=False, sheet_name='Buy (Loan)')
            sell_out.to_excel(wr,      index=False, sheet_name='Sell (Repayment)')
            sell_aktif.to_excel(wr,    index=False, sheet_name='Sell Aktif (tanpa excluded)')
            if not df_porto_all.empty:
                for cid_sheet, df_p in porto_sheets.items():
                    safe_name = f"Porto_{cid_sheet}"[:31]
                    df_p.to_excel(wr, index=False, sheet_name=safe_name)
                df_porto_all.to_excel(wr, index=False, sheet_name='Portofolio_All')
            if not df_loan_req.empty:
                df_loan_req.to_excel(wr, index=False, sheet_name='Loan Request')
            if not summary_buy.empty:
                summary_buy.to_excel(wr,  index=False, sheet_name='Summary Buy')
            if not summary_sell.empty:
                summary_sell.to_excel(wr, index=False, sheet_name='Summary Sell')

        st.download_button(
            "📥 Download Hasil_TRX_PEI.xlsx",
            out.getvalue(),
            "Hasil_TRX_PEI.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.divider()

        # ─────────────────────────────────────────────────────────────
        # 7. TABS PREVIEW
        # ─────────────────────────────────────────────────────────────
        tabs = ["📊 BUY (Loan)", "📊 SELL (Repayment)"]
        if not df_porto_all.empty:
            tabs += ["🏦 Portofolio Kolateral"]
        if not df_loan_req.empty:
            tabs += ["💰 Loan Request"]
        tabs += ["📋 Summary"]

        tab_objs = st.tabs(tabs)
        idx = 0

        with tab_objs[idx]:
            idx += 1
            st.caption(f"Total {len(buy_out)} baris · hanya nasabah PEI · hanya saham dicover margin")
            if not buy_out.empty:
                def color_nett(val):
                    if val == 'LOAN PEI':     return 'background-color:#d4edda;color:#155724'
                    if val == 'LOAN PARTIAL': return 'background-color:#fff3cd;color:#856404'
                    return ''
                st.dataframe(buy_out.style.map(color_nett, subset=['NETT']), use_container_width=True)
            else:
                st.info("Tidak ada transaksi BUY nasabah PEI pada hari ini.")

        with tab_objs[idx]:
            idx += 1
            st.caption(
                f"Total {len(sell_out)} baris · "
                f"{len(sell_aktif)} aktif · "
                f"{n_excluded} excluded (saham tidak ada di kolateral)"
            )
            if not sell_out.empty:
                def color_nett_sell(val):
                    if val == 'REPAY PEI':       return 'background-color:#cce5ff;color:#004085'
                    if val == 'ALL STOCK REPAY': return 'background-color:#d4edda;color:#155724'
                    if 'EXCLUDED' in str(val):   return 'background-color:#f8d7da;color:#721c24'
                    return ''
                st.dataframe(sell_out.style.map(color_nett_sell, subset=['NETT']), use_container_width=True)
            else:
                st.info("Tidak ada transaksi SELL nasabah PEI pada hari ini.")

        if not df_porto_all.empty:
            with tab_objs[idx]:
                idx += 1
                st.caption("Portofolio saham yang dijadikan kolateral PEI, dikelompokkan per nasabah")
                cid_options  = ['Semua'] + list(porto_sheets.keys())
                selected_cid = st.selectbox("Filter Nasabah:", cid_options, key='porto_filter')
                show_df = df_porto_all if selected_cid == 'Semua' else porto_sheets[selected_cid]
                disp = show_df.copy()
                for c in ['coll_vol']:
                    if c in disp.columns:
                        disp[c] = pd.to_numeric(disp[c], errors='coerce').fillna(0)
                st.dataframe(disp, use_container_width=True)

        if not df_loan_req.empty:
            with tab_objs[idx]:
                idx += 1
                st.caption("Saham yang didepositkan sebagai kolateral dan nilai loan yang dapat diajukan")
                disp_loan = df_loan_req.copy()
                for c in ['Collateral Vol', 'Price (STK_CLOS)', 'Collateral Value', 'Loan Eligible']:
                    if c in disp_loan.columns:
                        disp_loan[c] = pd.to_numeric(disp_loan[c], errors='coerce').fillna(0)
                st.dataframe(disp_loan, use_container_width=True)
                total_loan = disp_loan['Loan Eligible'].sum()
                st.metric("Total Loan Eligible (semua nasabah)", fmt_rp(total_loan))

        with tab_objs[idx]:
            st.subheader("📋 Summary BUY (Loan PEI)")
            if not summary_buy.empty:
                st.dataframe(summary_buy, use_container_width=True)
            else:
                st.info("Tidak ada transaksi BUY.")
            st.subheader("📋 Summary SELL Aktif (Repayment PEI)")
            if not summary_sell.empty:
                st.dataframe(summary_sell, use_container_width=True)
            else:
                st.info("Tidak ada transaksi SELL aktif.")

    except Exception as e:
        st.error(f"❌ Gagal memproses data: {e}")
        st.exception(e)

else:
    st.warning(
        "⬆️ Upload 7 file wajib (Netting Invoice, SID Client, Risk Parameter, "
        "Margin Buy, Margin Sell, Outstanding Position, File Harga Saham) untuk memulai."
    )

    with st.expander("📖 Panduan Struktur File"):
        st.markdown("""
        | # | File | Sumber File |
        |---|------|-------------|
        | 1 | **Netting Invoice** | Hasil generate dari List of Invoice |
        | 2 | **SID Client** | Template excel berisi kolom SID, CID, Name |
        | 3 | **Risk Parameter** | File .txt dari menu Report I-Fast Web |
        | 4 | **Margin Buy** | File .txt dari menu Daily Transaction I-Fast Web |
        | 5 | **Margin Sell** | File .txt dari menu Daily Transaction I-Fast Web |
        | 6 | **Outstanding Position** | File .txt dari I-Fast Web (untuk lot & kolateral) |
        | 7 | **File Harga Saham** | File Excel dengan kolom STK_CODE dan STK_CLOS |

        > **Repayment Value** = Lot (dari OP) × STK_CLOS (dari File Harga)  
        > **Loan Value** = Volume (dari Invoice) × STK_CLOS (dari File Harga)  
        > Saham yang tidak dicover margin PEI (dari Risk Parameter) otomatis di-take out.
        """)
