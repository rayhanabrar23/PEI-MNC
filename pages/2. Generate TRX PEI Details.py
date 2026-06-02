import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="TRX PEI Details", layout="wide", page_icon="📑")
st.title("📑 TRX PEI Details Generator")
st.info("Sistem ini memproses transaksi nasabah PEI: **Repayment (RP)** → **Loan Request (LR)**")

# ─────────────────────────────────────────────
# FUNGSI UTILITAS
# ─────────────────────────────────────────────

def find_and_rename(df):
    mapping = {
        'stock_key':  ['no_share','no_shares','Stock Code','Stockcode','Stock','SYMBOL','StockCode'],
        'sid_key':    ['SID','SID_No','Client_SID'],
        'cid_key':    ['no_cust','CID','Client_ID','Account_No'],
        'avail_risk': ['Available Quantity','availablequantity','Available Qty','AvailableQuantity'],
        'name_key':   ['Name','Client_Name','Nama'],
        'haircut_key':['Haircut','haircut','HC'],
        'margin_flag':['Margin','margin','MARGIN','flag_margin','MarginFlag'],
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
    num_keys = ['amt','vol','qty','val','price','avail','haircut','collateral','quantity','margin']
    if extra_keys:
        num_keys += extra_keys
    def parse_number(s):
        s = str(s).strip().replace('"','').replace('%','')
        if s in ('','nan','None','-'): return 0.0
        s = s.replace(',','')
        if re.fullmatch(r'\d{1,3}(\.\d{3})+', s): s = s.replace('.','')
        return pd.to_numeric(s, errors='coerce') or 0.0
    for c in df.columns:
        if any(k in str(c).lower() for k in num_keys):
            df[c] = df[c].apply(parse_number)
    return df

def fmt_rp(val):
    try: return f"Rp {float(val):,.0f}".replace(',','.')
    except: return str(val)

def load_price_file(uploaded_file):
    df = pd.read_excel(uploaded_file, sheet_name=0, header=0)
    price_map = {}
    for _, row in df.iterrows():
        code  = str(row['no_share']).strip().upper()
        price = pd.to_numeric(str(row['kurs_now']).replace(',',''), errors='coerce')
        if pd.notna(price) and code and code != 'NAN':
            price_map[code] = float(price)
    return price_map

def calc_collateral_value(stocks_dict, price_map, risk_params):
    """Hitung total collateral value dari dict {stock: lot}"""
    total = 0.0
    detail = {}
    for stock, lot in stocks_dict.items():
        price = price_map.get(stock, 0)
        hc    = risk_params.get(stock, 0.05)
        cv    = lot * price * (1 - hc)
        total += cv
        detail[stock] = {'lot': lot, 'price': price, 'hc': hc, 'cv': cv}
    return total, detail

# ─────────────────────────────────────────────
# UPLOAD FILE
# ─────────────────────────────────────────────
st.subheader("📂 Upload File")
col_u1, col_u2, col_u3, col_u4 = st.columns(4)
with col_u1:
    file_invoice    = st.file_uploader("1. Netting Invoice (xlsx)", type=['xlsx'])
    file_sid_client = st.file_uploader("2. SID Client (xlsx)",      type=['xlsx'])
with col_u2:
    file_risk  = st.file_uploader("3. Risk Parameter (.txt)", type=['txt'])
    file_m_buy = st.file_uploader("4. Margin Buy (.txt)",     type=['txt'])
with col_u3:
    file_m_sell = st.file_uploader("5. Margin Sell (.txt)",         type=['txt'])
    file_ep     = st.file_uploader("6. Outstanding Position (.txt)", type=['txt'])
with col_u4:
    file_price = st.file_uploader("7. Closing Price (xlsx)", type=['xlsx'])

required_files = [file_invoice, file_sid_client, file_risk, file_m_buy, file_m_sell, file_price, file_ep]

if all(required_files):
    try:
        with st.spinner("⚙️ Memproses data..."):

            # ── LOAD DATA ────────────────────────────────────
            df_inv = find_and_rename(pd.read_excel(file_invoice,    dtype=str))
            df_sid = find_and_rename(pd.read_excel(file_sid_client, dtype=str))
            df_inv['stock_key'] = df_inv['stock_key'].astype(str).str.strip().str.upper()
            df_inv['cid_key']   = df_inv['cid_key'].astype(str).str.strip()
            df_sid['cid_key']   = df_sid['cid_key'].astype(str).str.strip()
            df_sid['sid_key']   = df_sid['sid_key'].astype(str).str.strip()

            # Risk Parameter
            df_risk = find_and_rename(pd.read_csv(file_risk, sep='|', dtype=str))
            df_risk.columns = df_risk.columns.str.strip()
            df_risk = clean_num(df_risk)
            df_risk['stock_key'] = df_risk['stock_key'].astype(str).str.strip().str.upper()

            # Haircut lookup: {stock: hc_decimal}
            risk_params = {}
            if 'haircut_key' in df_risk.columns:
                for _, r in df_risk.iterrows():
                    hc_raw = pd.to_numeric(r.get('haircut_key', 5), errors='coerce') or 5
                    hc = hc_raw / 100 if hc_raw > 1 else hc_raw
                    risk_params[str(r['stock_key']).strip().upper()] = hc

            # Margin covered stocks
            if 'margin_flag' in df_risk.columns:
                margin_ok = df_risk[df_risk['margin_flag'].astype(str).str.strip().str.upper().isin(['Y','1','TRUE','YES'])]
                margin_covered_stocks = set(margin_ok['stock_key'].unique())
            else:
                margin_covered_stocks = set(df_risk['stock_key'].unique())

            risk_sub = df_risk[['stock_key','avail_risk']].drop_duplicates('stock_key').copy()
            if 'haircut_key' in df_risk.columns:
                risk_sub = df_risk[['stock_key','avail_risk','haircut_key']].drop_duplicates('stock_key').copy()

            # Margin Buy — nilai & lot transaksi BELI kemarin per SID
            df_mbuy = pd.read_csv(file_m_buy, sep='|', dtype=str).rename(columns={
                'SID': 'sid_key',
                'STOCK CODE': 'stock_key',
                'MARGIN BUY QUANTITY': 'qty',       # <--- Sesuaikan dengan nama di file
                'AVAILABLE MARKET VALUE': 'value'   # <--- Sesuaikan dengan nama di file
            })
            df_mbuy.columns = df_mbuy.columns.str.strip()
            df_mbuy = clean_num(df_mbuy)
            df_mbuy['stock_key'] = df_mbuy['stock_key'].astype(str).str.strip().str.upper()
            df_mbuy['sid_key']   = df_mbuy['sid_key'].astype(str).str.strip()

            # Margin Sell — nilai & lot transaksi JUAL kemarin per SID
            df_msell = pd.read_csv(file_m_sell, sep='|', dtype=str).rename(columns={
                'SID': 'sid_key',
                'STOCK CODE': 'stock_key',
                'REGULAR SELL QUANTITY': 'qty',
                'AVAILABLE SELL VALUE': 'value'
            })
            df_msell.columns = df_msell.columns.str.strip()
            df_msell = clean_num(df_msell)
            df_msell['stock_key'] = df_msell['stock_key'].astype(str).str.strip().str.upper()
            df_msell['sid_key']   = df_msell['sid_key'].astype(str).str.strip()

            # Closing Price
            price_map = load_price_file(file_price)

            # ── OUTSTANDING POSITION ─────────────────────────
            # op_data[sid] = {stock: lot} — kolateral saham saat ini
            # op_loan[sid] = {loan_existing, accrued_interest, available_limit, name}
            op_stocks  = {}   # {sid: {stock: lot}}
            op_loan    = {}   # {sid: {loan_existing, accrued_interest, available_limit, name}}
            sid_to_cid = df_sid.set_index('sid_key')['cid_key'].to_dict()
            cid_to_sid = df_sid.set_index('cid_key')['sid_key'].to_dict()
            cid_to_name= df_sid.drop_duplicates('cid_key').set_index('cid_key')['name_key'].to_dict() if 'name_key' in df_sid.columns else {}

            content = file_ep.read().decode('utf-8')
            current_sid = None
            for line in content.strip().splitlines():
                parts = line.strip().split('|')
                if not parts: continue
                if parts[0] == '0':
                    current_sid = parts[3].strip() if len(parts) > 3 else None
                    if current_sid:
                        try: loan_ex = float(parts[5]) if len(parts) > 5 else 0.0
                        except: loan_ex = 0.0
                        try: accrued = float(parts[6]) if len(parts) > 6 else 0.0
                        except: accrued = 0.0
                        try: avail_lim = float(parts[7]) if len(parts) > 7 else 0.0
                        except: avail_lim = 0.0
                        name = parts[4].strip() if len(parts) > 4 else current_sid
                        op_loan[current_sid] = {
                            'loan_existing': loan_ex,
                            'accrued_interest': accrued,
                            'available_limit': avail_lim,
                            'name': name,
                        }
                        op_stocks[current_sid] = {}
                elif parts[0] == '1' and current_sid:
                    if len(parts) < 5: continue
                    stock = parts[3].strip().upper()
                    try: vol = float(parts[4])
                    except: vol = 0.0
                    if stock and vol > 0:
                        op_stocks[current_sid][stock] = op_stocks[current_sid].get(stock, 0) + vol

            # ── VOLUME FORMULA ────────────────────────────────
            if 'Volume_Formula' not in df_inv.columns:
                df_inv['_sign'] = df_inv['bors'].map({'B':1,'S':-1}).fillna(0)
                df_inv['tot_vol'] = pd.to_numeric(df_inv['tot_vol'].astype(str).str.replace(',',''), errors='coerce').fillna(0)
                df_inv['_signed_vol'] = df_inv['tot_vol'] * df_inv['_sign']
                net_map = df_inv.groupby(['cid_key','stock_key'])['_signed_vol'].sum()
                df_inv['vol_net_total'] = df_inv.set_index(['cid_key','stock_key']).index.map(net_map)
                def calc_vol_formula(row):
                    total = row['vol_net_total']
                    vol   = row['tot_vol']
                    if total < 0: return -vol if row['bors'] == 'S' else 0
                    elif total > 0: return vol if row['bors'] == 'B' else 0
                    return 0
                df_inv['Volume_Formula'] = df_inv.apply(calc_vol_formula, axis=1)
            else:
                df_inv['Volume_Formula'] = pd.to_numeric(df_inv['Volume_Formula'].astype(str).str.replace(',',''), errors='coerce').fillna(0)

            pei_cids = set(df_sid['cid_key'].astype(str).str.strip())

            # ── SELL REGULAR VALUE per SID (dari file Margin Sell) ──
            # Ini nilai transaksi jual kemarin = RP Maks per nasabah
            # {sid: {stock: {'lot_sell': x, 'value': y}}}
            sell_reg_by_sid = {}
            for _, row in df_msell.iterrows():
                sid   = str(row.get('sid_key','')).strip()
                stock = str(row.get('stock_key','')).strip().upper()
                # coba ambil qty dan price dari kolom yang ada
                qty   = pd.to_numeric(row.get('qty', row.get('QUANTITY', row.get('volume', 0))), errors='coerce') or 0
                price_s = price_map.get(stock, 0)
                val   = qty * price_s
                if sid not in sell_reg_by_sid:
                    sell_reg_by_sid[sid] = {}
                if stock not in sell_reg_by_sid[sid]:
                    sell_reg_by_sid[sid][stock] = {'lot_sell': 0, 'value': 0}
                sell_reg_by_sid[sid][stock]['lot_sell'] += qty
                sell_reg_by_sid[sid][stock]['value']    += val

            # ── MARGIN BUY VALUE per SID ─────────────────────
            # {sid: {stock: {'lot_buy': x, 'value': y}}}
            buy_reg_by_sid = {}
            for _, row in df_mbuy.iterrows():
                sid   = str(row.get('sid_key','')).strip()
                stock = str(row.get('stock_key','')).strip().upper()
                qty   = pd.to_numeric(row.get('qty', row.get('QUANTITY', row.get('volume', 0))), errors='coerce') or 0
                price_b = price_map.get(stock, 0)
                val   = qty * price_b
                if sid not in buy_reg_by_sid:
                    buy_reg_by_sid[sid] = {}
                if stock not in buy_reg_by_sid[sid]:
                    buy_reg_by_sid[sid][stock] = {'lot_buy': 0, 'value': 0}
                buy_reg_by_sid[sid][stock]['lot_buy'] += qty
                buy_reg_by_sid[sid][stock]['value']   += val

            # ── HITUNG PER NASABAH ────────────────────────────
            # Untuk setiap SID PEI, hitung:
            # 1. RP Min, RP Maks, collateral after RP, loan after RP, rasio RP
            # 2. Ceiling LR, collateral LR (setelah RP + saham beli baru), rasio LR, max LR
            results = {}
            pei_sids = set()
            for _, row in df_sid.iterrows():
                cid = str(row['cid_key']).strip()
                sid = str(row['sid_key']).strip()
                if cid in pei_cids:
                    pei_sids.add(sid)

            for sid in pei_sids:
                loan_info = op_loan.get(sid, {'loan_existing':0,'accrued_interest':0,'available_limit':0,'name':sid})
                stocks_op = op_stocks.get(sid, {})
                loan_ex   = loan_info['loan_existing']
                accrued   = loan_info['accrued_interest']
                avail_lim = loan_info['available_limit']
                name      = loan_info['name']

                sell_stocks = sell_reg_by_sid.get(sid, {})
                buy_stocks  = buy_reg_by_sid.get(sid, {})

                # ── RP CALCULATION ──
                rp_detail = []
                for stock, sdata in sell_stocks.items():
                    lot_sell = sdata['lot_sell']
                    lot_op   = stocks_op.get(stock, 0)
                    lot_keluar = min(lot_sell, lot_op)   # saham yang bisa keluar dari kolateral
                    price_s  = price_map.get(stock, 0)
                    rp_min   = lot_keluar * price_s      # nilai minimum RP (sesuai lot di OP)
                    rp_maks  = sdata['value']            # nilai maksimum RP (nilai transaksi jual kemarin)
                    ada_di_op = lot_op > 0
                    rp_detail.append({
                        'stock':      stock,
                        'lot_sell':   lot_sell,
                        'lot_op':     lot_op,
                        'lot_keluar': lot_keluar,
                        'price':      price_s,
                        'rp_min':     rp_min,
                        'rp_maks':    rp_maks,
                        'ada_di_op':  ada_di_op,
                    })

                total_rp_maks = sum(d['rp_maks'] for d in rp_detail if d['ada_di_op'])
                total_rp_min  = sum(d['rp_min']  for d in rp_detail if d['ada_di_op'])

                # Collateral setelah RP (saham yang keluar dikurangi dari OP)
                stocks_after_rp = dict(stocks_op)
                for d in rp_detail:
                    if d['ada_di_op'] and d['lot_keluar'] > 0:
                        stocks_after_rp[d['stock']] = stocks_after_rp.get(d['stock'], 0) - d['lot_keluar']
                        if stocks_after_rp[d['stock']] <= 0:
                            del stocks_after_rp[d['stock']]

                coll_before_rp, _ = calc_collateral_value(stocks_op,        price_map, risk_params)
                coll_after_rp,  _ = calc_collateral_value(stocks_after_rp,  price_map, risk_params)

                loan_after_rp = max(loan_ex - total_rp_maks, 0)
                rasio_rp = (loan_after_rp + accrued) / coll_after_rp if coll_after_rp > 0 else None

                # ── LR CALCULATION ──
                # Collateral LR = collateral after RP + saham beli baru
                stocks_after_lr = dict(stocks_after_rp)
                for stock, bdata in buy_stocks.items():
                    stocks_after_lr[stock] = stocks_after_lr.get(stock, 0) + bdata['lot_buy']

                coll_after_lr, _ = calc_collateral_value(stocks_after_lr, price_map, risk_params)

                total_buy_val = sum(b['value'] for b in buy_stocks.values())
                # Ceiling LR = min(nilai beli kemarin, avail_limit + rp_maks)
                avail_efektif = avail_lim + total_rp_maks
                ceiling_lr    = min(total_buy_val, avail_efektif)

                # Rasio LR = (loan_after_rp + accrued + lr_diajukan) / coll_after_lr
                numerator_lr = loan_after_rp + accrued + ceiling_lr
                rasio_lr     = numerator_lr / coll_after_lr if coll_after_lr > 0 else None

                # Max LR agar rasio = 63%
                max_lr_63 = max(coll_after_lr * 0.63 - (loan_after_rp + accrued), 0) if coll_after_lr > 0 else 0
                max_lr_65 = max(coll_after_lr * 0.65 - (loan_after_rp + accrued), 0) if coll_after_lr > 0 else 0
                max_lr_final = min(ceiling_lr, max_lr_65)

                results[sid] = {
                    'name':           name,
                    'cid':            sid_to_cid.get(sid, sid),
                    'loan_existing':  loan_ex,
                    'accrued':        accrued,
                    'avail_limit':    avail_lim,
                    # RP
                    'rp_detail':      rp_detail,
                    'total_rp_maks':  total_rp_maks,
                    'total_rp_min':   total_rp_min,
                    'stocks_op':      stocks_op,
                    'stocks_after_rp':stocks_after_rp,
                    'coll_before_rp': coll_before_rp,
                    'coll_after_rp':  coll_after_rp,
                    'loan_after_rp':  loan_after_rp,
                    'rasio_rp':       rasio_rp,
                    # LR
                    'buy_stocks':     buy_stocks,
                    'stocks_after_lr':stocks_after_lr,
                    'coll_after_lr':  coll_after_lr,
                    'total_buy_val':  total_buy_val,
                    'avail_efektif':  avail_efektif,
                    'ceiling_lr':     ceiling_lr,
                    'rasio_lr':       rasio_lr,
                    'max_lr_63':      max_lr_63,
                    'max_lr_65':      max_lr_65,
                    'max_lr_final':   max_lr_final,
                }

        st.session_state['pei_results'] = results
        st.session_state['price_map']   = price_map
        st.session_state['risk_params'] = risk_params
        st.success("✅ Data Berhasil Diproses!")

    except Exception as e:
        st.error(f"❌ Gagal memproses data: {e}")
        st.exception(e)

# ─────────────────────────────────────────────
# TAMPILKAN HASIL
# ─────────────────────────────────────────────
if st.session_state.get('pei_results'):
    results    = st.session_state['pei_results']
    price_map  = st.session_state['price_map']
    risk_params= st.session_state['risk_params']

    n_has_rp = sum(1 for v in results.values() if v['total_rp_maks'] > 0)
    n_has_lr = sum(1 for v in results.values() if v['total_buy_val'] > 0)
    n_rp_lolos = sum(1 for v in results.values() if v['rasio_rp'] is not None and v['rasio_rp'] < 0.65)
    n_lr_lolos = sum(1 for v in results.values() if v['rasio_lr'] is not None and v['rasio_lr'] < 0.65)

    m1,m2,m3,m4,m5 = st.columns(5)
    m1.metric("Total Nasabah PEI", len(results))
    m2.metric("Nasabah Ada RP", n_has_rp)
    m3.metric("Nasabah Ada LR", n_has_lr)
    m4.metric("RP Lolos Rasio", n_rp_lolos)
    m5.metric("LR Lolos Rasio", n_lr_lolos)

    tab_rp, tab_lr, tab_simulator, tab_summary, tab_export = st.tabs([
        "📤 LANGKAH 1 — Repayment (RP)",
        "📥 LANGKAH 2 — Loan Request (LR)",
        "🎛️ Simulator RP → LR",
        "📋 Summary",
        "📥 Export",
    ])

    # ── TAB RP ────────────────────────────────────────────────
    with tab_rp:
        st.info("💡 Input RP **terlebih dahulu** pagi ini. Setelah RP diinput, available limit nasabah naik dan kolateral berkurang.")
        rp_rows = []
        for sid, d in results.items():
            if d['total_rp_maks'] <= 0: continue
            for rd in d['rp_detail']:
                if not rd['ada_di_op']: continue
                rasio_val = f"{d['rasio_rp']*100:.2f}%" if d['rasio_rp'] is not None else "N/A"
                status_rasio = "✅ LOLOS" if d['rasio_rp'] is not None and d['rasio_rp'] < 0.65 else "❌ GAGAL"
                rp_rows.append({
                    'SID':              sid,
                    'Nama':             d['name'],
                    'Saham':            rd['stock'],
                    'Lot Transaksi Jual':int(rd['lot_sell']),
                    'Lot di OP':        int(rd['lot_op']),
                    'Lot Keluar Kolateral': int(rd['lot_keluar']),
                    'Harga':            rd['price'],
                    'RP Min (Lot OP × Harga)':   rd['rp_min'],
                    'RP Maks (Nilai Jual Kemarin)': rd['rp_maks'],
                    'Loan Before RP':   d['loan_existing'],
                    'Loan After RP':    d['loan_after_rp'],
                    'Coll After RP':    d['coll_after_rp'],
                    'Rasio After RP':   rasio_val,
                    'Status Rasio':     status_rasio,
                })

        if rp_rows:
            df_rp = pd.DataFrame(rp_rows)
            def color_rp(val):
                if val == '✅ LOLOS': return 'background-color:#d4edda;color:#155724'
                if val == '❌ GAGAL': return 'background-color:#f8d7da;color:#721c24'
                return ''
            st.dataframe(df_rp.style.map(color_rp, subset=['Status Rasio']), use_container_width=True)

            st.subheader("Rekap RP per Nasabah")
            rekap_rp = []
            for sid, d in results.items():
                if d['total_rp_maks'] <= 0: continue
                rekap_rp.append({
                    'SID':          sid,
                    'Nama':         d['name'],
                    'Loan Existing':fmt_rp(d['loan_existing']),
                    'Coll Before RP':fmt_rp(d['coll_before_rp']),
                    'RP Min':       fmt_rp(d['total_rp_min']),
                    'RP Maks':      fmt_rp(d['total_rp_maks']),
                    'Loan After RP':fmt_rp(d['loan_after_rp']),
                    'Coll After RP':fmt_rp(d['coll_after_rp']),
                    'Rasio After RP': f"{d['rasio_rp']*100:.2f}%" if d['rasio_rp'] is not None else "N/A",
                    'Status':       "✅ LOLOS" if d['rasio_rp'] is not None and d['rasio_rp'] < 0.65 else ("⚠️ Coll=0" if d['coll_after_rp']==0 else "❌ GAGAL"),
                })
            st.dataframe(pd.DataFrame(rekap_rp), use_container_width=True, hide_index=True)
        else:
            st.info("Tidak ada nasabah PEI dengan transaksi jual kemarin.")

    # ── TAB LR ────────────────────────────────────────────────
    with tab_lr:
        st.info("💡 Input LR **setelah RP diinput**. Ceiling LR = min(Nilai Beli Kemarin, Avail Limit + RP). Collateral sudah ditambah saham beli baru.")
        lr_rows = []
        for sid, d in results.items():
            if d['total_buy_val'] <= 0: continue
            rasio_val = f"{d['rasio_lr']*100:.2f}%" if d['rasio_lr'] is not None else "N/A"
            status_rasio = "✅ LOLOS" if d['rasio_lr'] is not None and d['rasio_lr'] < 0.65 else "❌ GAGAL (perlu dipotong)"
            for stock, bdata in d['buy_stocks'].items():
                lr_rows.append({
                    'SID':              sid,
                    'Nama':             d['name'],
                    'Saham Beli':       stock,
                    'Lot Beli Kemarin': int(bdata['lot_buy']),
                    'Harga':            price_map.get(stock, 0),
                    'Nilai Beli':       bdata['value'],
                    'RP Lolos (pagi)':  d['total_rp_maks'],
                    'Avail Limit':      d['avail_limit'],
                    'Avail Efektif':    d['avail_efektif'],
                    'Ceiling LR':       d['ceiling_lr'],
                    'Loan After RP':    d['loan_after_rp'],
                    'Coll After LR':    d['coll_after_lr'],
                    'Numerator LR':     d['loan_after_rp'] + d['accrued'] + d['ceiling_lr'],
                    'Rasio LR':         rasio_val,
                    'Max LR (63%)':     d['max_lr_63'],
                    'Max LR Final':     d['max_lr_final'],
                    'Status Rasio':     status_rasio,
                })

        if lr_rows:
            df_lr = pd.DataFrame(lr_rows)
            def color_lr(val):
                if val == '✅ LOLOS': return 'background-color:#d4edda;color:#155724'
                if '❌' in str(val):  return 'background-color:#f8d7da;color:#721c24'
                return ''
            st.dataframe(df_lr.style.map(color_lr, subset=['Status Rasio']), use_container_width=True)

            st.subheader("Rekap LR per Nasabah")
            rekap_lr = []
            for sid, d in results.items():
                if d['total_buy_val'] <= 0: continue
                rekap_lr.append({
                    'SID':           sid,
                    'Nama':          d['name'],
                    'Ceiling LR':    fmt_rp(d['ceiling_lr']),
                    'Loan After RP': fmt_rp(d['loan_after_rp']),
                    'Coll After LR': fmt_rp(d['coll_after_lr']),
                    'Rasio LR':      f"{d['rasio_lr']*100:.2f}%" if d['rasio_lr'] is not None else "N/A",
                    'Max LR (63%)':  fmt_rp(d['max_lr_63']),
                    'Max LR Final':  fmt_rp(d['max_lr_final']),
                    'Status':        "✅ LOLOS" if d['rasio_lr'] is not None and d['rasio_lr'] < 0.65 else "❌ Perlu dipotong",
                })
            st.dataframe(pd.DataFrame(rekap_lr), use_container_width=True, hide_index=True)
        else:
            st.info("Tidak ada nasabah PEI dengan transaksi beli kemarin.")

    # ── TAB SIMULATOR ─────────────────────────────────────────
    with tab_simulator:
        st.subheader("🎛️ Simulator — Ubah Nilai RP dan Lihat Dampak ke LR")
        st.info("Ubah nilai RP yang akan diinput. Sistem akan menghitung ulang rasio RP, loan after RP, collateral, dan max LR secara otomatis.")

        sid_options = [sid for sid, d in results.items() if d['total_rp_maks'] > 0 or d['total_buy_val'] > 0]
        if not sid_options:
            st.warning("Tidak ada nasabah dengan transaksi RP atau LR.")
        else:
            selected_sid = st.selectbox("Pilih Nasabah (SID):", sid_options,
                format_func=lambda s: f"{s} — {results[s]['name']}")
            d = results[selected_sid]

            st.divider()
            col_info1, col_info2, col_info3 = st.columns(3)
            col_info1.metric("Loan Outstanding", fmt_rp(d['loan_existing']))
            col_info2.metric("Collateral Awal (OP)", fmt_rp(d['coll_before_rp']))
            col_info3.metric("Current Ratio", f"{(d['loan_existing']/d['coll_before_rp']*100):.2f}%" if d['coll_before_rp'] > 0 else "N/A")

            st.subheader("Step 1 — Atur Nilai RP per Saham")
            st.caption(f"Range RP: Min = Lot OP × Harga | Maks = Nilai Transaksi Jual Kemarin")

            rp_inputs = {}
            for rd in d['rp_detail']:
                if not rd['ada_di_op']: continue
                col_s1, col_s2, col_s3, col_s4 = st.columns([2,1,1,2])
                with col_s1:
                    st.markdown(f"**{rd['stock']}**  \nLot Jual: {int(rd['lot_sell']):,} | Lot OP: {int(rd['lot_op']):,}")
                with col_s2:
                    st.caption("RP Min")
                    st.write(fmt_rp(rd['rp_min']))
                with col_s3:
                    st.caption("RP Maks")
                    st.write(fmt_rp(rd['rp_maks']))
                with col_s4:
                    rp_val = st.number_input(
                        f"RP Value untuk {rd['stock']}",
                        min_value=float(rd['rp_min']),
                        max_value=float(rd['rp_maks']),
                        value=float(rd['rp_maks']),
                        step=1_000_000.0,
                        format="%.0f",
                        key=f"rp_input_{selected_sid}_{rd['stock']}",
                        label_visibility="collapsed"
                    )
                rp_inputs[rd['stock']] = {
                    'rp_value':   rp_val,
                    'lot_keluar': rd['lot_keluar'],
                    'stock':      rd['stock'],
                }

            total_rp_sim = sum(v['rp_value'] for v in rp_inputs.values())

            # Hitung ulang setelah simulator
            stocks_after_rp_sim = dict(d['stocks_op'])
            for stock, v in rp_inputs.items():
                if v['lot_keluar'] > 0:
                    stocks_after_rp_sim[stock] = stocks_after_rp_sim.get(stock, 0) - v['lot_keluar']
                    if stocks_after_rp_sim[stock] <= 0:
                        del stocks_after_rp_sim[stock]

            coll_after_rp_sim, coll_detail_rp = calc_collateral_value(stocks_after_rp_sim, price_map, risk_params)
            loan_after_rp_sim = max(d['loan_existing'] - total_rp_sim, 0)
            rasio_rp_sim = (loan_after_rp_sim + d['accrued']) / coll_after_rp_sim if coll_after_rp_sim > 0 else None

            st.divider()
            st.subheader("Hasil Setelah RP")
            res_col1, res_col2, res_col3, res_col4 = st.columns(4)
            res_col1.metric("Total RP Diinput", fmt_rp(total_rp_sim))
            res_col2.metric("Loan After RP", fmt_rp(loan_after_rp_sim))
            res_col3.metric("Collateral After RP", fmt_rp(coll_after_rp_sim))
            rasio_rp_str = f"{rasio_rp_sim*100:.2f}%" if rasio_rp_sim is not None else "N/A"
            res_col4.metric("Rasio After RP", rasio_rp_str,
                delta="✅ LOLOS" if rasio_rp_sim is not None and rasio_rp_sim < 0.65 else "❌ GAGAL",
                delta_color="normal" if rasio_rp_sim is not None and rasio_rp_sim < 0.65 else "inverse")

            st.divider()
            st.subheader("Step 2 — Dampak ke Loan Request (LR)")

            # Collateral LR = after RP + saham beli baru
            stocks_after_lr_sim = dict(stocks_after_rp_sim)
            for stock, bdata in d['buy_stocks'].items():
                stocks_after_lr_sim[stock] = stocks_after_lr_sim.get(stock, 0) + bdata['lot_buy']
            coll_after_lr_sim, _ = calc_collateral_value(stocks_after_lr_sim, price_map, risk_params)

            avail_efektif_sim = d['avail_limit'] + total_rp_sim
            ceiling_lr_sim    = min(d['total_buy_val'], avail_efektif_sim)
            max_lr_63_sim     = max(coll_after_lr_sim * 0.63 - (loan_after_rp_sim + d['accrued']), 0) if coll_after_lr_sim > 0 else 0
            max_lr_65_sim     = max(coll_after_lr_sim * 0.65 - (loan_after_rp_sim + d['accrued']), 0) if coll_after_lr_sim > 0 else 0
            max_lr_final_sim  = min(ceiling_lr_sim, max_lr_65_sim)
            numerator_lr_sim  = loan_after_rp_sim + d['accrued'] + ceiling_lr_sim
            rasio_lr_sim      = numerator_lr_sim / coll_after_lr_sim if coll_after_lr_sim > 0 else None

            lr_col1, lr_col2, lr_col3 = st.columns(3)
            lr_col1.metric("Avail Limit Efektif", fmt_rp(avail_efektif_sim),
                help="Avail Limit + RP yang diinput")
            lr_col2.metric("Ceiling LR", fmt_rp(ceiling_lr_sim),
                help="min(Nilai Beli Kemarin, Avail Limit Efektif)")
            lr_col3.metric("Collateral After LR", fmt_rp(coll_after_lr_sim),
                help="Collateral After RP + Saham Beli Baru")

            lr_col4, lr_col5, lr_col6 = st.columns(3)
            lr_col4.metric("Numerator LR", fmt_rp(numerator_lr_sim),
                help="Loan After RP + Accrued + Ceiling LR")
            rasio_lr_str = f"{rasio_lr_sim*100:.2f}%" if rasio_lr_sim is not None else "N/A"
            lr_col5.metric("Rasio LR", rasio_lr_str,
                delta="✅ LOLOS" if rasio_lr_sim is not None and rasio_lr_sim < 0.65 else "❌ Perlu dipotong",
                delta_color="normal" if rasio_lr_sim is not None and rasio_lr_sim < 0.65 else "inverse")
            lr_col6.metric("Max LR Final (aman 63%)", fmt_rp(max_lr_final_sim))

            # Detail collateral
            with st.expander("📊 Detail Collateral After LR"):
                coll_rows = []
                for stock, lot in stocks_after_lr_sim.items():
                    price_c = price_map.get(stock, 0)
                    hc_c    = risk_params.get(stock, 0.05)
                    cv_c    = lot * price_c * (1 - hc_c)
                    source  = "OP (sisa)" if stock in stocks_after_rp_sim else "Beli Baru"
                    coll_rows.append({'Saham': stock, 'Lot': int(lot), 'Harga': price_c,
                                      'HC': f"{hc_c*100:.0f}%", 'Coll Value': cv_c, 'Sumber': source})
                st.dataframe(pd.DataFrame(coll_rows), use_container_width=True, hide_index=True)

    # ── TAB SUMMARY ───────────────────────────────────────────
    with tab_summary:
        st.subheader("📋 Summary Semua Nasabah PEI")
        summary_rows = []
        for sid, d in results.items():
            rasio_rp_str = f"{d['rasio_rp']*100:.2f}%" if d['rasio_rp'] is not None else "-"
            rasio_lr_str = f"{d['rasio_lr']*100:.2f}%" if d['rasio_lr'] is not None else "-"
            summary_rows.append({
                'SID':            sid,
                'Nama':           d['name'],
                'Loan Existing':  d['loan_existing'],
                'Coll Awal':      d['coll_before_rp'],
                'Current Ratio':  f"{d['loan_existing']/d['coll_before_rp']*100:.2f}%" if d['coll_before_rp'] > 0 else "-",
                'RP Min':         d['total_rp_min'],
                'RP Maks':        d['total_rp_maks'],
                'Loan After RP':  d['loan_after_rp'],
                'Coll After RP':  d['coll_after_rp'],
                'Rasio After RP': rasio_rp_str,
                'Status RP':      "✅" if d['rasio_rp'] is not None and d['rasio_rp'] < 0.65 else ("-" if d['total_rp_maks']==0 else "❌"),
                'Ceiling LR':     d['ceiling_lr'],
                'Coll After LR':  d['coll_after_lr'],
                'Max LR Final':   d['max_lr_final'],
                'Rasio LR':       rasio_lr_str,
                'Status LR':      "✅" if d['rasio_lr'] is not None and d['rasio_lr'] < 0.65 else ("-" if d['total_buy_val']==0 else "❌"),
            })
        st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

    # ── TAB EXPORT ────────────────────────────────────────────
    with tab_export:
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as wr:
            # Sheet RP
            rp_exp = []
            for sid, d in results.items():
                for rd in d['rp_detail']:
                    if not rd['ada_di_op']: continue
                    rp_exp.append({
                        'SID': sid, 'Nama': d['name'], 'Saham': rd['stock'],
                        'Lot Jual': int(rd['lot_sell']), 'Lot OP': int(rd['lot_op']),
                        'Lot Keluar Kolateral': int(rd['lot_keluar']),
                        'Harga': rd['price'],
                        'RP Min': rd['rp_min'], 'RP Maks': rd['rp_maks'],
                        'Loan Before RP': d['loan_existing'], 'Loan After RP': d['loan_after_rp'],
                        'Coll After RP': d['coll_after_rp'],
                        'Rasio After RP': f"{d['rasio_rp']*100:.2f}%" if d['rasio_rp'] is not None else "N/A",
                    })
            pd.DataFrame(rp_exp).to_excel(wr, sheet_name='Repayment (RP)', index=False)

            # Sheet LR
            lr_exp = []
            for sid, d in results.items():
                if d['total_buy_val'] <= 0: continue
                lr_exp.append({
                    'SID': sid, 'Nama': d['name'],
                    'RP Maks': d['total_rp_maks'], 'Avail Limit': d['avail_limit'],
                    'Avail Efektif': d['avail_efektif'], 'Ceiling LR': d['ceiling_lr'],
                    'Loan After RP': d['loan_after_rp'], 'Coll After LR': d['coll_after_lr'],
                    'Numerator LR': d['loan_after_rp'] + d['accrued'] + d['ceiling_lr'],
                    'Rasio LR': f"{d['rasio_lr']*100:.2f}%" if d['rasio_lr'] is not None else "N/A",
                    'Max LR (63%)': d['max_lr_63'], 'Max LR Final': d['max_lr_final'],
                })
            pd.DataFrame(lr_exp).to_excel(wr, sheet_name='Loan Request (LR)', index=False)

            # Sheet Summary
            pd.DataFrame(summary_rows if 'summary_rows' in dir() else []).to_excel(wr, sheet_name='Summary', index=False)

        st.download_button(
            "📥 Download Hasil_TRX_PEI.xlsx",
            out.getvalue(), "Hasil_TRX_PEI.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.warning("⬆️ Upload semua file wajib untuk memulai.")
    with st.expander("📖 Panduan & Alur Proses"):
        st.markdown("""
        ### Alur Proses Pagi Hari

        **LANGKAH 1 — Input RP terlebih dahulu:**
        - RP Min  = Lot di OP × Closing Price  
        - RP Maks = Nilai transaksi jual kemarin (dari Margin Sell)  
        - Saham keluar kolateral = min(lot jual, lot di OP)  
        - Rasio RP = (Loan After RP + Accrued) / Collateral After RP < 65%

        **LANGKAH 2 — Input LR setelah RP selesai:**
        - Collateral LR = Collateral After RP + Saham Beli Baru (dari Margin Buy)  
        - Ceiling LR    = min(Nilai Beli Kemarin, Avail Limit + RP)  
        - Rasio LR      = (Loan After RP + Accrued + LR Diajukan) / Collateral LR < 65%  
        - Jika > 65% → dipotong agar rasio = 63%

        | # | File | Keterangan |
        |---|------|------------|
        | 1 | Netting Invoice | Hasil generate dari List of Invoice |
        | 2 | SID Client | Kolom SID, CID, Name |
        | 3 | Risk Parameter | .txt dari I-Fast Web |
        | 4 | Margin Buy | Transaksi beli kemarin |
        | 5 | Margin Sell | Transaksi jual kemarin |
        | 6 | Outstanding Position | Kolateral & loan existing |
        | 7 | Closing Price | Kolom no_share & kurs_now |
        """)
