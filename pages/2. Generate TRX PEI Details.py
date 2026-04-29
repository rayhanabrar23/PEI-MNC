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
        'stock_key':  ['no_share', 'no_shares', 'Stock Code', 'Stockcode', 'Stock', 'SYMBOL'],
        'sid_key':    ['SID', 'SID_No', 'Client_SID'],
        'cid_key':    ['no_cust', 'CID', 'Client_ID', 'Account_No'],
        'avail_risk': ['Available Quantity', 'availablequantity', 'Available Qty'],
        'name_key':   ['Name', 'Client_Name', 'Nama'],
        'haircut_key':['Haircut', 'haircut', 'HC'],
        'price_key':  ['close_prc', 'Price', 'PRICE'],
        # CLOSING PRICE dan STOCK CODE (kapital semua) tidak di sini
        # karena itu nama resmi di Margin Buy/Sell yang tidak pakai find_and_rename
    }
    rename_dict = {}
    for official, aliases in mapping.items():
        for col in df.columns:
            if str(col).strip() in aliases:
                rename_dict[col] = official
                break
    return df.rename(columns=rename_dict)


def clean_num(df, extra_keys=None):
    """Konversi kolom numerik: hapus koma, tanda kutip, dll."""
    num_keys = ['amt', 'vol', 'qty', 'val', 'price', 'avail',
                'haircut', 'collateral', 'quantity', 'margin']
    if extra_keys:
        num_keys += extra_keys
    for c in df.columns:
        if any(k in str(c).lower() for k in num_keys):
            df[c] = pd.to_numeric(
                df[c].astype(str).str.replace(',', '').str.replace('"', ''),
                errors='coerce'
            ).fillna(0)
    return df


def fmt_rp(val):
    """Format angka sebagai Rupiah."""
    try:
        return f"Rp {int(val):,}".replace(',', '.')
    except Exception:
        return str(val)


def fmt_vol(val):
    """Format volume dengan pemisah ribuan."""
    try:
        return f"{int(val):,}".replace(',', '.')
    except Exception:
        return str(val)


# ─────────────────────────────────────────────
# 3. AREA UPLOAD FILE
# ─────────────────────────────────────────────
st.subheader("📂 Upload File")
col_u1, col_u2, col_u3 = st.columns(3)

with col_u1:
    file_invoice    = st.file_uploader("1. Netting Invoice (xlsx)",   type=['xlsx'])
    file_sid_client = st.file_uploader("2. SID Client (xlsx)",        type=['xlsx'])

with col_u2:
    file_risk       = st.file_uploader("3. Risk Parameter (xlsx)",    type=['xlsx'])
    file_m_buy      = st.file_uploader("4. Margin Buy (xlsx)",        type=['xlsx'])

with col_u3:
    file_m_sell     = st.file_uploader("5. Margin Sell (xlsx)",       type=['xlsx'])
    file_ep         = st.file_uploader("6. File OP", type=['txt'])

# ─────────────────────────────────────────────
# 4. PROSES DATA (hanya jika 5 file wajib ada)
# ─────────────────────────────────────────────
required_files = [file_invoice, file_sid_client, file_risk, file_m_buy, file_m_sell]

if all(required_files):
    try:
        with st.spinner("⚙️ Memproses seluruh data..."):

            # ── 4a. LOAD & STANDARISASI ──────────────────────
            df_inv   = find_and_rename(pd.read_excel(file_invoice,    dtype=str))
            df_sid   = find_and_rename(pd.read_excel(file_sid_client, dtype=str))
            df_risk  = find_and_rename(pd.read_excel(file_risk,       dtype=str))

            # Margin Buy & Sell: rename manual supaya CLOSING PRICE, HAIRCUT, dll tetap nama aslinya
            df_mbuy  = pd.read_excel(file_m_buy,  dtype=str).rename(columns={
            'SID':        'sid_key',
            'STOCK CODE': 'stock_key',
            })
            df_msell = pd.read_excel(file_m_sell, dtype=str).rename(columns={
            'SID':        'sid_key',
            'STOCK CODE': 'stock_key',
            })

            # ── 4b. VOLUME FORMULA (netting B vs S per nasabah-saham) ──
            # Volume_Formula sudah dihitung di Netting Invoice → pakai langsung
            if 'Volume_Formula' not in df_inv.columns:
                # fallback: hitung ulang
                df_inv['_sign'] = df_inv['bors'].map({'B': 1, 'S': -1}).fillna(0)
                df_inv['_signed_vol'] = df_inv['tot_vol'] * df_inv['_sign']
                net_map = df_inv.groupby(['cid_key', 'stock_key'])['_signed_vol'].sum()
                df_inv['vol_net_total'] = df_inv.set_index(
                    ['cid_key', 'stock_key']).index.map(net_map)

                def calc_vol_formula(row):
                    total = row['vol_net_total']
                    if total < 0:
                        return row['tot_vol'] * -1 if row['bors'] == 'S' else 0
                    elif total > 0:
                        return row['tot_vol'] if row['bors'] == 'B' else 0
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

            # ── 4c. FILTER HANYA NASABAH PEI (ada di SID Client) ────
            pei_cids = set(df_sid['cid_key'].astype(str).str.strip())

            # Drop duplikat CID sebelum dijadikan lookup
            df_sid_unique = df_sid.drop_duplicates(subset='cid_key', keep='first')
            sid_lookup = df_sid_unique.set_index('cid_key')[['sid_key', 'name_key']].to_dict('index')
            # Risk parameter lookup: stock → (avail_qty, haircut)
            risk_sub = df_risk[['stock_key', 'avail_risk']].drop_duplicates('stock_key').copy()
            if 'haircut_key' in df_risk.columns:
                risk_sub = df_risk[['stock_key', 'avail_risk', 'haircut_key']].drop_duplicates('stock_key').copy()

            risk_lookup = risk_sub.set_index('stock_key').to_dict('index')

            # ── 4d. SHEET BUY (Margin Buy – Loan PEI) ────────────────
            df_buy_inv = df_inv[
                (df_inv['bors'] == 'B') &
                (df_inv['cid_key'].isin(pei_cids)) &
                (df_inv['Volume_Formula'].abs() > 0)
            ].copy()

            buy_merged = df_buy_inv.merge(
                df_mbuy, on=['sid_key', 'stock_key'], how='left', suffixes=('', '_m')
            ) if 'sid_key' in df_buy_inv.columns else df_buy_inv.merge(
                df_sid[['cid_key', 'sid_key', 'name_key']], on='cid_key', how='left'
            ).merge(df_mbuy, on=['sid_key', 'stock_key'], how='left', suffixes=('', '_m'))

            # Jika sid_key belum ada di invoice, enrich dari df_sid
            if 'sid_key' not in df_buy_inv.columns:
                buy_merged = df_buy_inv.merge(
                    df_sid[['cid_key', 'sid_key', 'name_key']], on='cid_key', how='left'
                ).merge(df_mbuy, on=['sid_key', 'stock_key'], how='left', suffixes=('', '_m'))
            else:
                buy_merged = df_buy_inv.merge(
                    df_mbuy, on=['sid_key', 'stock_key'], how='left', suffixes=('', '_m')
                )
            # Sebelum buy_merged.merge(risk_sub...)
            buy_merged['stock_key'] = buy_merged['stock_key'].astype(str).str.strip().str.upper()
            risk_sub['stock_key']   = risk_sub['stock_key'].astype(str).str.strip().str.upper()

            buy_merged = buy_merged.merge(risk_sub, on='stock_key', how='left')

            # Kolom-kolom output BUY
            buy_out = pd.DataFrame()
            buy_out['SID']           = buy_merged.get('sid_key', '')
            buy_out['STOCK CODE']    = buy_merged['stock_key']
            for col in ['MARGIN BUY QUANTITY', 'LOAN QUANTITY', 'AVAILABLE QUANTITY',
                        'CLOSING PRICE', 'AVAILABLE MARKET VALUE', 'HAIRCUT',
                        'AVAILABLE COLLATERAL VALUE']:
                buy_out[col] = buy_merged[col] if col in buy_merged.columns else 0
            buy_out['B/S']           = 'B'
            buy_out['CID']           = buy_merged['cid_key']
            buy_out['Name']          = buy_merged.get('name_key', '')
            buy_out['Stock']         = buy_merged['stock_key']
            buy_out['Volume']        = buy_merged['Volume_Formula'].abs()
            buy_out['Value']         = buy_merged.apply(
                lambda r: r.get('AVAILABLE MARKET VALUE', 0) if r['Volume_Formula'] != 0 else 0,
                axis=1
            )
            buy_out['PEI (Risk/Porto)'] = buy_merged.get('avail_risk', 0)

            def nett_buy(row):
                k = row.get('CLOSING PRICE', '')
                m = row.get('AVAILABLE MARKET VALUE', '')
                p = pd.to_numeric(row['PEI (Risk/Porto)'], errors='coerce')
                s = pd.to_numeric(row['Volume'], errors='coerce')

                if k == '' or m == '': return ''
                if pd.isna(p) or p == 0: return 'NON MARGIN'
                if p > 0:
                    if pd.isna(s) or s < 0: return ''
                    return 'LOAN PEI' if s > p else 'LOAN PARTIAL'
                return ''

            buy_out['NETT'] = buy_out.apply(nett_buy, axis=1)

            # ── 4f. SHEET PORTOFOLIO KOLATERAL ───────────────────────
            porto_sheets = {}
            df_porto_all = pd.DataFrame()
            porto_coll_lookup = {}

            if file_ep is not None:
                content = file_ep.read().decode('utf-8')
                current_sid = None
                porto_rows = []

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
                                porto_rows.append({
                                    'CID':       cid,
                                    'SID':       current_sid,
                                    'stock_key': stock,
                                    'coll_vol':  vol,
                                })

                if porto_rows:
                    df_porto_all = pd.DataFrame(porto_rows)
                    # Enrich nama dari df_sid
                    df_porto_all = df_porto_all.merge(
                        df_sid[['cid_key', 'name_key']].rename(columns={'cid_key': 'CID', 'name_key': 'Name'}),
                        on='CID', how='left'
                    )
                    porto_sheets = dict(tuple(df_porto_all.groupby('CID')))


             # ── 4e. SHEET SELL (Margin Sell – Repayment PEI) ─────────
            df_sell_inv = df_inv[
                (df_inv['bors'] == 'S') &
                (df_inv['cid_key'].isin(pei_cids)) &
                (df_inv['Volume_Formula'].abs() > 0)
            ].copy()

            if 'sid_key' not in df_sell_inv.columns:
                sell_merged = df_sell_inv.merge(
                    df_sid[['cid_key', 'sid_key', 'name_key']], on='cid_key', how='left'
                ).merge(df_msell, on=['sid_key', 'stock_key'], how='left', suffixes=('', '_m'))
            else:
                sell_merged = df_sell_inv.merge(
                    df_msell, on=['sid_key', 'stock_key'], how='left', suffixes=('', '_m')
                )
            # Tambahkan ini SEBELUM baris sell_merged.merge(risk_sub...)
            sell_merged['stock_key'] = sell_merged['stock_key'].astype(str).str.strip().str.upper()
            risk_sub['stock_key']    = risk_sub['stock_key'].astype(str).str.strip().str.upper()

            sell_merged = sell_merged.merge(risk_sub, on='stock_key', how='left')
            
            sell_out = pd.DataFrame()
            sell_out['SID']          = sell_merged.get('sid_key', '')
            sell_out['STOCK CODE']   = sell_merged['stock_key']
            for col in ['REGULAR SELL QUANTITY', 'REPAYMENT QUANTITY',
                        'AVAILABLE SELL QUANTITY', 'CLOSING PRICE', 'AVAILABLE SELL VALUE']:
                sell_out[col] = sell_merged[col] if col in sell_merged.columns else 0
            sell_out['B/S']          = 'S'
            sell_out['CID']          = sell_merged['cid_key']
            sell_out['Name']         = sell_merged.get('name_key', '')
            sell_out['Stock']        = sell_merged['stock_key']
            sell_out['Volume']       = sell_merged['Volume_Formula'].abs()
            sell_out['Value']        = sell_merged.apply(
                lambda r: r.get('AVAILABLE SELL VALUE', r.get('amt_pay', 0))
                          if r['Volume_Formula'] != 0 else 0,
                axis=1
            )
            # BARU - dari COLLATERAL (VOL) Portofolio Client
            def get_pei_sell(row):
                cid   = str(row.get('cid_key', '')).strip()
                stock = str(row.get('stock_key', '')).strip().upper()
                vol   = porto_coll_lookup.get((cid, stock), 0)
                return -vol  # negatif karena Sell = repayment

            sell_out['PEI (Risk/Porto)'] = sell_merged.apply(get_pei_sell, axis=1)

            # Fix nett_sell sesuai logika Excel
            def nett_sell(row):
                k = row.get('CLOSING PRICE', '')
                m = row.get('AVAILABLE SELL VALUE', '')
                p = pd.to_numeric(row['PEI (Risk/Porto)'], errors='coerce')
                s = pd.to_numeric(row['Volume'], errors='coerce')

                if k == '' or m == '': return ''
                if pd.isna(p) or p == 0: return 'NON MARGIN'
                if p < 0:
                    if pd.isna(s) or s == 0: return ''
                    return 'REPAY PEI' if s < abs(p) else 'ALL STOCK REPAY'
                return ''
    
            sell_out['NETT'] = sell_out.apply(nett_sell, axis=1)
            
            # ── 4g. SHEET LOAN REQUEST ────────────────────────────────
            # Loan Request = nasabah yang Stock Deposit di Portofolio
            # Hitung Loan Eligible = Collateral Value × (1 - additional_margin)
            # Disini kita ambil dari data Porto yang coll_vol > 0
            loan_req_rows = []

            if not df_porto_all.empty and 'coll_vol' in df_porto_all.columns:
                for _, row in df_porto_all.iterrows():
                    vol = pd.to_numeric(row.get('coll_vol', 0), errors='coerce') or 0
                    if vol > 0:
                        price  = pd.to_numeric(row.get('price_key',  0), errors='coerce') or 0
                        hc_pct = pd.to_numeric(row.get('haircut_key', 5), errors='coerce') or 5
                        cv     = vol * price * (1 - hc_pct / 100)
                        loan_req_rows.append({
                            'CID':              row.get('CID', ''),
                            'Name':             row.get('Name', ''),
                            'Stock':            row.get('stock_key', ''),
                            'HC (%)':           hc_pct,
                            'Collateral Vol':   vol,
                            'Price':            price,
                            'Collateral Value': cv,
                            'Loan Eligible':    cv,   # Bisa dikurangi LTV ratio kalau ada
                            'Status':           'STOCK DEPOSIT',
                        })

            df_loan_req = pd.DataFrame(loan_req_rows)

            # ── 4h. RINGKASAN (COMP-like summary) ────────────────────
            # Dari semua BUY outstanding
            summary_buy = buy_out.groupby(['CID', 'Name']).agg(
                Total_Volume=('Volume', 'sum'),
                Total_Value=('Value', 'sum'),
                Loan_Eligible=('AVAILABLE COLLATERAL VALUE', 'sum'),
            ).reset_index()
            summary_buy['Status'] = 'LOAN PEI'

            # Dari semua SELL repayment
            summary_sell = sell_out.groupby(['CID', 'Name']).agg(
                Total_Volume=('Volume', 'sum'),
                Total_Value=('Value', 'sum'),
            ).reset_index()
            summary_sell['Status'] = 'REPAY PEI'

        # ─────────────────────────────────────────────────────────────
        # 5. TAMPILKAN HASIL
        # ─────────────────────────────────────────────────────────────
        st.success("✅ Data Berhasil Diproses!")

        # Metric Cards
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        col_m1.metric("Nasabah PEI Terdaftar", len(pei_cids))
        col_m2.metric("Baris BUY (Loan)",  len(buy_out))
        col_m3.metric("Baris SELL (Repay)", len(sell_out))
        col_m4.metric("Portofolio Kolateral",
                       len(df_porto_all) if not df_porto_all.empty else "—")

        # ─────────────────────────────────────────────────────────────
        # 6. DOWNLOAD EXCEL
        # ─────────────────────────────────────────────────────────────
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as wr:
            buy_out.to_excel(wr,  index=False, sheet_name='Buy (Loan)')
            sell_out.to_excel(wr, index=False, sheet_name='Sell (Repayment)')
            if not df_porto_all.empty:
                # Satu sheet per nasabah di Portofolio
                for cid_sheet, df_p in porto_sheets.items():
                    safe_name = f"Porto_{cid_sheet}"[:31]
                    df_p.to_excel(wr, index=False, sheet_name=safe_name)
                df_porto_all.to_excel(wr, index=False, sheet_name='Portofolio_All')
            if not df_loan_req.empty:
                df_loan_req.to_excel(wr, index=False, sheet_name='Loan Request')
            summary_buy.to_excel(wr,  index=False, sheet_name='Summary Buy')
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
            st.caption(f"Total {len(buy_out)} baris · hanya nasabah PEI terdaftar")
            if not buy_out.empty:
                # Highlight NETT column
                def color_nett(val):
                    if val == 'LOAN PEI':       return 'background-color:#d4edda;color:#155724'
                    if val == 'LOAN PARTIAL':   return 'background-color:#fff3cd;color:#856404'
                    return ''
                st.dataframe(
                    buy_out.style.map(color_nett, subset=['NETT']),
                    use_container_width=True
                )
            else:
                st.info("Tidak ada transaksi BUY nasabah PEI pada hari ini.")

        with tab_objs[idx]:
            idx += 1
            st.caption(f"Total {len(sell_out)} baris · hanya nasabah PEI terdaftar")
            if not sell_out.empty:
                def color_nett_sell(val):
                    if val == 'REPAY PEI':       return 'background-color:#cce5ff;color:#004085'
                    if val == 'ALL STOCK REPAY': return 'background-color:#d4edda;color:#155724'
                    return ''
                st.dataframe(
                    sell_out.style.map(color_nett_sell, subset=['NETT']),
                    use_container_width=True
                )
            else:
                st.info("Tidak ada transaksi SELL nasabah PEI pada hari ini.")

        if not df_porto_all.empty:
            with tab_objs[idx]:
                idx += 1
                st.caption("Portofolio saham yang dijadikan kolateral PEI, dikelompokkan per nasabah")
                # Filter per nasabah
                cid_options = ['Semua'] + list(porto_sheets.keys())
                selected_cid = st.selectbox("Filter Nasabah:", cid_options, key='porto_filter')
                if selected_cid == 'Semua':
                    show_df = df_porto_all
                else:
                    show_df = porto_sheets[selected_cid]

                # Format kolom numerik untuk display
                disp = show_df.copy()
                for c in ['coll_vol', 'price_key', 'coll_value']:
                    if c in disp.columns:
                        disp[c] = pd.to_numeric(disp[c], errors='coerce').fillna(0)

                st.dataframe(disp, use_container_width=True)

                # Mini summary per nasabah
                st.subheader("📌 Ringkasan Kolateral per Nasabah")
                if 'coll_value' in df_porto_all.columns:
                    summary_porto = df_porto_all.groupby(['CID', 'Name']).agg(
                        Jumlah_Saham=('stock_key', 'count'),
                        Total_Coll_Value=('coll_value', 'sum')
                    ).reset_index()
                    summary_porto['Total_Coll_Value_Rp'] = summary_porto['Total_Coll_Value'].apply(fmt_rp)
                    st.dataframe(summary_porto[['CID', 'Name', 'Jumlah_Saham', 'Total_Coll_Value_Rp']],
                                 use_container_width=True)

        if not df_loan_req.empty:
            with tab_objs[idx]:
                idx += 1
                st.caption("Saham yang didepositkan sebagai kolateral dan nilai loan yang dapat diajukan")
                disp_loan = df_loan_req.copy()
                for c in ['Collateral Vol', 'Price', 'Collateral Value', 'Loan Eligible']:
                    if c in disp_loan.columns:
                        disp_loan[c] = pd.to_numeric(disp_loan[c], errors='coerce').fillna(0)

                st.dataframe(disp_loan, use_container_width=True)

                # Total loan request
                total_loan = disp_loan['Loan Eligible'].sum()
                st.metric("Total Loan Eligible (semua nasabah)", fmt_rp(total_loan))

        with tab_objs[idx]:
            st.subheader("📋 Summary BUY (Loan PEI)")
            st.dataframe(summary_buy, use_container_width=True)
            st.subheader("📋 Summary SELL (Repayment PEI)")
            st.dataframe(summary_sell, use_container_width=True)

    except Exception as e:
        st.error(f"❌ Gagal memproses data: {e}")
        st.exception(e)

else:
    st.warning("⬆️ Upload 6 file wajib (Netting Invoice, SID Client, Risk Parameter, Margin Buy, Margin Sell, OP) untuk memulai.")

    # Panduan struktur file
    with st.expander("📖 Panduan Struktur File"):
        st.markdown("""
        | # | File | Sumber File |
        |---|------|-------------|
        | 1 | **Netting Invoice** | `Hasil generate Netting List of Invoice dari data mentah List of Invoice` |
        | 2 | **SID Client** | `Template excel yang berisi kolom SID, CID, Name` |
        | 3 | **Risk Parameter** | `File .txt yang diunduh pada menu Report melalui I-Fast Web` |
        | 4 | **Margin Buy** | `File .txt yang diunduh pada menu Daily Transaction melalui I-Fast Web` |
        | 5 | **Margin Sell** | `File .txt yang diunduh pada menu Daily Transaction melalui I-Fast Web` |
        | 6 | **Outstanding Position** | file .txt yang diunduh melalui I-Fast Web' |
        
        > **Outstanding Position (OP)** dipakai untuk menentukan nilai portofolio client.
        """)
