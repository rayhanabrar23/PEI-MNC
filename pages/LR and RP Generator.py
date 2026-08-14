import streamlit as st
import pandas as pd
import io
from datetime import datetime
import copy

st.set_page_config(page_title="PEI Tools", layout="wide", page_icon="📊")
st.title("📊 PEI Tools — TRX PEI & Validasi MNC")

# ─────────────────────────────────────────────
# SESSION STATE INIT
# ─────────────────────────────────────────────
for key in [
    'shared_closing_prices', 'shared_risk_params', 'shared_risk_avail_qty',
    'shared_netting', 'shared_net_buy', 'shared_net_sell', 'shared_cid_to_sid',
    'shared_op_data', 'shared_cl_data', 'shared_cl_value_date',
    # TRX PEI
    'pei_results', 'pei_results_original',
    # Validasi MNC
    'df_sell_edited', 'sid_results', 'global_result', 'df_buy',
    'df_buy_adjusted',
]:
    if key not in st.session_state:
        st.session_state[key] = None

if 'clamped_warnings' not in st.session_state:
    st.session_state['clamped_warnings'] = []
if 'sid_results_original' not in st.session_state:
    st.session_state['sid_results_original'] = {}

RATIO_THRESHOLD    = 0.65   # batas rasio maksimum (dipakai utk solve RP Min sbg referensi)
AUTO_ADJUST_TARGET = 0.63   # batas rasio final utk LR (dipakai sbg cap utama, bukan opsional lagi)
CREDIT_LIMIT_PARTISIPAN = 160_000_000_000.0

# ─────────────────────────────────────────────
# HELPERS UMUM
# ─────────────────────────────────────────────
def fmt_rp(val):
    try: return f"Rp {float(val):,.0f}".replace(',', '.')
    except: return str(val)

def fmt_pct(val):
    try: return f"{float(val)*100:.2f}%"
    except: return str(val)

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

def calc_collateral(stocks_dict, closing_prices, risk_params):
    total = 0.0
    detail = []
    for stock, qty in stocks_dict.items():
        cp   = closing_prices.get(stock, 0.0)
        hc   = risk_params.get(stock, 0.05)
        coll = qty * cp * (1 - hc)
        total += coll
        detail.append({"stock": stock, "qty": qty, "cp": cp, "hc": hc, "collateral": coll})
    return total, detail

def solve_rp_min(loan_ex, accrued, coll_after_rp):
    """
    RP MINIMUM — nilai RP terkecil yang cukup membawa rasio (loan+accrued)/kolateral
    turun ke tepat 65% (RATIO_THRESHOLD). Ini hanya REFERENSI batas bawah, BUKAN nilai
    yang dipakai untuk rekomendasi di file output (lihat solve_rp_max).
    RP_min = (loan O/S + accrued) − (kolateral sisa setelah jual × 65%), di-clamp ke [0, loan_ex].
    """
    if coll_after_rp > 0:
        rp = (loan_ex + accrued) - coll_after_rp * RATIO_THRESHOLD
    else:
        rp = loan_ex + accrued
    return min(max(rp, 0.0), loan_ex)

def solve_rp_max(loan_ex, accrued):
    """
    RP MAXIMUM — dipakai sebagai REKOMENDASI RESMI di file output RP.
    RP hanya dihitung kalau nasabah punya transaksi sell saham marginable (dicek di
    pemanggil lewat has_rp/rp_detail). Begitu terpicu, kewajibannya adalah melunasi
    seluruh Loan O/S + Accrued Interest — tidak dibatasi rasio 65% seperti RP Min.
    """
    return loan_ex + accrued

def solve_lr(coll_after_lr, loan_after_rp, accrued, avail_lim, total_rp):
    """
    LR final = MIN( LR@63% , Available Limit efektif )
      LR@63%          = (kolateral setelah beli × 63%) − (loan setelah RP + accrued)
      Available Limit efektif = avail_limit (dari Credit Limit file) + total RP yang sudah masuk
    Mengganti ceiling_lr lama (min(nilai_beli*1.1, avail_efektif)) dan cap 65% lama.
    """
    lr_63 = max(coll_after_lr * AUTO_ADJUST_TARGET - (loan_after_rp + accrued), 0) if coll_after_lr > 0 else 0
    avail_efektif = avail_lim + total_rp
    lr_final = min(lr_63, avail_efektif)
    binding = "Available Limit" if avail_efektif < lr_63 else "Rasio 63%"
    return lr_63, avail_efektif, lr_final, binding

# ─────────────────────────────────────────────
# PARSERS FILE SHARED
# ─────────────────────────────────────────────
def load_closing_price(uploaded_file) -> dict:
    df = pd.read_excel(uploaded_file, sheet_name=0, header=0)
    result = {}
    for _, row in df.iterrows():
        code  = str(row['no_share']).strip().upper()
        price = pd.to_numeric(str(row['kurs_now']).replace(',',''), errors='coerce')
        if pd.notna(price) and code and code != 'NAN':
            result[code] = float(price)
    return result

def load_risk_parameter(uploaded_file) -> tuple:
    result_hc  = {}
    result_avq = {}
    content = uploaded_file.read().decode("utf-8", errors="replace")
    uploaded_file.seek(0)
    for line in content.strip().splitlines():
        line = line.strip()
        if not line or line.startswith("StockCode"): continue
        parts = line.split("|")
        if len(parts) < 4: continue
        code = parts[0].strip().upper()
        try: hc  = float(parts[2]) / 100.0
        except: hc  = 0.0
        try: avq = float(parts[3])
        except: avq = 0.0
        result_hc[code]  = hc
        result_avq[code] = avq
    return result_hc, result_avq

def get_loan_status(avail_qty, volume_buy):
    if volume_buy <= 0:
        return ""
    if avail_qty < 0:
        return ""
    if avail_qty > volume_buy:
        return "LOAN PEI"
    else:
        return "LOAN PARTIAL"

def parse_op_file(content: str) -> dict:
    result = {}
    for line in content.strip().splitlines():
        line  = line.strip()
        if not line: continue
        parts = line.split("|")
        if parts[0] == "0":
            if len(parts) < 7: continue
            sid = parts[3].strip()
            try: loan_ex = float(parts[5])
            except: loan_ex = 0.0
            try: accrued = float(parts[6])
            except: accrued = 0.0
            try: avail = float(parts[7]) if len(parts) > 7 else 0.0
            except: avail = 0.0
            result[sid] = {
                "loan_existing":    loan_ex,
                "accrued_interest": accrued,
                "available_limit":  avail,
                "name":             parts[4].strip() if len(parts) > 4 else sid,
                "stocks":           {},
            }
        elif parts[0] == "1":
            if len(parts) < 5: continue
            sid   = parts[2].strip()
            stock = parts[3].strip().upper()
            try: vol = float(parts[4])
            except: vol = 0.0
            if sid in result and stock and vol > 0:
                result[sid]["stocks"][stock] = result[sid]["stocks"].get(stock, 0) + vol
    return result

def parse_credit_limit_file(content: str):
    """
    Parse file Credit Limit (.txt, pipe-delimited).
    Kolom: Value Date|Broker Code|SID|Client Name|Client Credit Limit|Used Limit|Available Limit
    Ini SATU-SATUNYA sumber Available Limit yang valid — field 'available_limit' di file OP
    selalu 0 (bukan bug data, memang formatnya begitu), jadi jangan dipakai.
    """
    result = {}
    value_date = None
    for i, line in enumerate(content.strip().splitlines()):
        line  = line.strip()
        if not line: continue
        parts = line.split("|")
        if i == 0 and parts[0].strip().lower() == "value date": continue
        if len(parts) < 7: continue
        sid = parts[2].strip()
        try: avail_limit = float(parts[6].replace(",",""))
        except: avail_limit = 0.0
        if value_date is None: value_date = parts[0].strip()
        result[sid] = {"available_limit": avail_limit, "name": parts[3].strip(), "value_date": parts[0].strip()}
    return result, value_date

def parse_netting_invoice(uploaded_file):
    """
    Parse file Netting Invoice (list of invoice, format .xls/.xlsx/.csv).
    Kolom yang dipakai (setelah find_and_rename): sid_key (dari SID), cid_key (dari no_cust),
    stock_key (dari no_share), bors (B/S), tot_vol, amt_done/amt_pay (nilai transaksi).
    Difilter board=RG & lorf=D (transaksi EP reguler, bukan subreg) sebagai pengaman.

    Netting dilakukan sederhana per SID + saham:
        net_lot   = total_buy_lot   - total_sell_lot
        net_value = total_buy_value - total_sell_value

    Return:
        netting     : {sid: {stock: {buy_lot, sell_lot, buy_value, sell_value, net_lot, net_value}}}
        cid_to_sid  : {cid: sid} — mapping langsung dari file invoice
    """
    fname = getattr(uploaded_file, 'name', '') or ''
    if fname.lower().endswith('.csv'):
        df = pd.read_csv(uploaded_file, dtype=str)
    else:
        df = pd.read_excel(uploaded_file, dtype=str)
    df = find_and_rename(df)
    for req in ('stock_key', 'sid_key', 'cid_key', 'bors', 'tot_vol'):
        if req not in df.columns:
            raise ValueError(f"Kolom '{req}' tidak ditemukan di file Netting Invoice.")

    df['stock_key'] = df['stock_key'].astype(str).str.strip().str.upper()
    df['sid_key']   = df['sid_key'].astype(str).str.strip()
    df['cid_key']   = df['cid_key'].astype(str).str.strip()
    df['bors']      = df['bors'].astype(str).str.strip().str.upper()

    # Pengaman filter EP reguler (subreg exclusion) — kalau kolomnya ada
    if 'board' in df.columns:
        df = df[df['board'].astype(str).str.strip().str.upper() == 'RG']
    if 'lorf' in df.columns:
        df = df[df['lorf'].astype(str).str.strip().str.upper() == 'D']

    df['tot_vol'] = pd.to_numeric(df['tot_vol'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

    val_col = 'amt_done' if 'amt_done' in df.columns else ('amt_pay' if 'amt_pay' in df.columns else None)
    if val_col is None:
        raise ValueError("Kolom nilai transaksi (amt_done/amt_pay) tidak ditemukan di file Netting Invoice.")
    df[val_col] = pd.to_numeric(df[val_col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

    netting = {}
    cid_to_sid = {}
    for _, row in df.iterrows():
        sid, cid, stock = row['sid_key'], row['cid_key'], row['stock_key']
        if not sid or not stock:
            continue
        vol, val, side = row['tot_vol'], row[val_col], row['bors']
        if cid:
            cid_to_sid[cid] = sid
        d = netting.setdefault(sid, {}).setdefault(
            stock, {'buy_lot': 0.0, 'sell_lot': 0.0, 'buy_value': 0.0, 'sell_value': 0.0})
        if side == 'B':
            d['buy_lot']   += vol
            d['buy_value'] += val
        elif side == 'S':
            d['sell_lot']   += vol
            d['sell_value'] += val

    for sid, stocks in netting.items():
        for stock, d in stocks.items():
            d['net_lot']   = d['buy_lot']   - d['sell_lot']
            d['net_value'] = d['buy_value'] - d['sell_value']

    return netting, cid_to_sid

def split_netting(netting: dict):
    """
    Pecah hasil netting jadi net_buy & net_sell per SID, dengan bentuk {sid: {stock: {lot, value}}}
      net_lot > 0 -> net Beli -> basis LR
      net_lot < 0 -> net Jual -> basis RP (lot & value disimpan positif)
      net_lot == 0 -> tidak ada posisi net, dilewati
    """
    net_buy, net_sell = {}, {}
    for sid, stocks in netting.items():
        for stock, d in stocks.items():
            if d['net_lot'] > 0:
                net_buy.setdefault(sid, {})[stock] = {'lot': d['net_lot'], 'value': d['net_value']}
            elif d['net_lot'] < 0:
                net_sell.setdefault(sid, {})[stock] = {'lot': -d['net_lot'], 'value': -d['net_value']}
    return net_buy, net_sell

# ─────────────────────────────────────────────
# UPLOAD FILE SHARED (selalu tampil di atas)
# ─────────────────────────────────────────────
st.subheader("📂 Upload File Bersama")
st.caption("File berikut digunakan oleh kedua modul. Upload sekali, dipakai di kedua tab.")

col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    file_cp     = st.file_uploader("1. Closing Price (.xlsx)", type=['xlsx'], key='shared_cp')
with col2:
    file_rp     = st.file_uploader("2. Risk Parameter (.txt)", type=['txt'], key='shared_rp')
with col3:
    file_op     = st.file_uploader("3. Outstanding Position (.txt)", type=['txt'], key='shared_op')
with col4:
    file_netinv = st.file_uploader("4. Netting Invoice / List of Invoice (.csv/.xls/.xlsx)",
                                    type=['csv', 'xls', 'xlsx'], key='shared_netinv')
with col5:
    file_cl     = st.file_uploader("5. Credit Limit (.txt)", type=['txt'], key='shared_cl')

shared_ready = all([file_cp, file_rp, file_op, file_netinv, file_cl])

if shared_ready:
    file_sig = f"{file_cp.name}_{file_rp.name}_{file_op.name}_{file_netinv.name}_{file_cl.name}"
    if st.session_state.get('_shared_file_sig') != file_sig:
        with st.spinner("⚙️ Memuat file bersama..."):
            st.session_state['shared_closing_prices'] = load_closing_price(file_cp)
            risk_hc, risk_avq = load_risk_parameter(file_rp)
            st.session_state['shared_risk_params']    = risk_hc
            st.session_state['shared_risk_avail_qty'] = risk_avq
            op_content = file_op.read().decode("utf-8", errors="replace")
            file_op.seek(0)
            st.session_state['shared_op_data']        = parse_op_file(op_content)
            netting, cid_to_sid = parse_netting_invoice(file_netinv)
            net_buy, net_sell   = split_netting(netting)
            st.session_state['shared_netting']    = netting
            st.session_state['shared_net_buy']    = net_buy
            st.session_state['shared_net_sell']   = net_sell
            st.session_state['shared_cid_to_sid'] = cid_to_sid
            cl_content = file_cl.read().decode("utf-8", errors="replace")
            file_cl.seek(0)
            cl_data, cl_vdate = parse_credit_limit_file(cl_content)
            st.session_state['shared_cl_data']       = cl_data
            st.session_state['shared_cl_value_date'] = cl_vdate
            st.session_state['_shared_file_sig']  = file_sig
        st.success(f"✅ File bersama dimuat — {len(st.session_state['shared_op_data'])} SID di OP, "
                   f"{len(st.session_state['shared_closing_prices'])} harga saham, "
                   f"{len(st.session_state['shared_netting'])} SID di Netting Invoice, "
                   f"{len(st.session_state['shared_cl_data'])} SID di Credit Limit.")
    cl_vdate = st.session_state.get('shared_cl_value_date')
    if cl_vdate:
        today_check = datetime.today().strftime("%Y/%m/%d")
        if cl_vdate != today_check:
            st.warning(f"⚠️ Value Date Credit Limit: **{cl_vdate}** (hari ini: {today_check})")
else:
    st.info("⬆️ Upload kelima file bersama di atas untuk mengaktifkan kedua modul.")

st.divider()

# ─────────────────────────────────────────────
# DUA TAB UTAMA
# ─────────────────────────────────────────────
tab_pei, tab_mnc = st.tabs(["📑 TRX PEI — Generator", "✅ Validasi MNC"])


# ══════════════════════════════════════════════════════════════════
# TAB 1 — TRX PEI
# ══════════════════════════════════════════════════════════════════
with tab_pei:
    st.info("Sistem ini memproses transaksi nasabah PEI: **Repayment (RP)** → **Loan Request (LR)**")

    st.subheader("📂 File Tambahan (TRX PEI)")
    st.caption("Netting Invoice & Credit Limit sekarang dipakai bersama (upload di bagian atas). Di sini cukup SID Client untuk menentukan nasabah PEI.")
    file_sid_client = st.file_uploader("A. SID Client (.xlsx)", type=['xlsx'], key='pei_sid')

    pei_ready = shared_ready and file_sid_client

    if pei_ready:
        if st.button("▶ Proses TRX PEI", type="primary", use_container_width=True, key='btn_pei'):
            try:
                with st.spinner("⚙️ Memproses TRX PEI..."):
                    closing_prices  = st.session_state['shared_closing_prices']
                    risk_params     = st.session_state['shared_risk_params']
                    op_data         = st.session_state['shared_op_data']
                    net_buy_by_sid  = st.session_state['shared_net_buy']
                    net_sell_by_sid = st.session_state['shared_net_sell']
                    cl_data         = st.session_state['shared_cl_data']

                    df_sid = find_and_rename(pd.read_excel(file_sid_client, dtype=str))
                    df_sid['cid_key']   = df_sid['cid_key'].astype(str).str.strip()
                    df_sid['sid_key']   = df_sid['sid_key'].astype(str).str.strip()

                    risk_params_hc = risk_params  # {stock: hc_decimal}

                    margin_covered_stocks = set()
                    for sid_val, d in op_data.items():
                        margin_covered_stocks.update(d.get('stocks', {}).keys())

                    sid_to_cid = df_sid.set_index('sid_key')['cid_key'].to_dict()
                    cid_to_sid = df_sid.set_index('cid_key')['sid_key'].to_dict()
                    cid_to_name = df_sid.drop_duplicates('cid_key').set_index('cid_key')['name_key'].to_dict() if 'name_key' in df_sid.columns else {}

                    op_stocks = {sid: d['stocks'] for sid, d in op_data.items()}
                    op_loan   = {sid: {
                        'loan_existing': d['loan_existing'],
                        'accrued_interest': d['accrued_interest'],
                        'name': d['name'],
                    } for sid, d in op_data.items()}

                    pei_cids = set(df_sid['cid_key'].astype(str).str.strip())

                    # RP & LR murni dari hasil netting Netting Invoice (net beli - net jual)
                    sell_reg_by_sid = net_sell_by_sid
                    buy_reg_by_sid  = net_buy_by_sid

                    results = {}
                    pei_sids = set()
                    for _, row in df_sid.iterrows():
                        cid = str(row['cid_key']).strip()
                        sid = str(row['sid_key']).strip()
                        if cid in pei_cids:
                            pei_sids.add(sid)

                    for sid in pei_sids:
                        loan_info = op_loan.get(sid, {'loan_existing':0,'accrued_interest':0,'name':sid})
                        stocks_op = op_stocks.get(sid, {})
                        loan_ex   = loan_info['loan_existing']
                        accrued   = loan_info['accrued_interest']
                        # Available Limit HARUS dari file Credit Limit — field di OP selalu 0.
                        avail_lim = cl_data.get(sid, {}).get('available_limit', 0.0)
                        name      = loan_info['name']

                        sell_stocks = sell_reg_by_sid.get(sid, {})
                        buy_stocks  = buy_reg_by_sid.get(sid, {})

                        # ── RP: lot keluar tetap dihitung dari saham yg dijual (utk info & kolateral) ──
                        rp_detail = []
                        for stock, sdata in sell_stocks.items():
                            lot_sell   = sdata['lot']
                            lot_op     = stocks_op.get(stock, 0)
                            lot_keluar = min(lot_sell, lot_op)
                            price_s    = closing_prices.get(stock, 0)
                            rp_min     = lot_keluar * price_s     # referensi: nilai jual (bukan basis RP)
                            ada_di_op  = lot_op > 0
                            rp_detail.append({
                                'stock': stock, 'lot_sell': lot_sell, 'lot_op': lot_op,
                                'lot_keluar': lot_keluar, 'price': price_s,
                                'rp_min': rp_min, 'ada_di_op': ada_di_op,
                            })

                        total_rp_min  = sum(d['rp_min']  for d in rp_detail if d['ada_di_op'])

                        stocks_after_rp = dict(stocks_op)
                        for d in rp_detail:
                            if d['ada_di_op'] and d['lot_keluar'] > 0:
                                stocks_after_rp[d['stock']] = stocks_after_rp.get(d['stock'], 0) - d['lot_keluar']
                                if stocks_after_rp[d['stock']] <= 0:
                                    del stocks_after_rp[d['stock']]

                        coll_before_rp, _ = calc_collateral(stocks_op,       closing_prices, risk_params_hc)
                        coll_after_rp,  _ = calc_collateral(stocks_after_rp, closing_prices, risk_params_hc)

                        # ── RP: kalau ada transaksi sell (rp_detail tidak kosong) & loan_ex>0 → kewajiban
                        #    RP = RP Max (lunasi penuh Loan O/S + Accrued). RP Min (solve-to-65%) tetap
                        #    dihitung sbg referensi batas bawah saja, tidak dipakai untuk rekomendasi.
                        has_rp_trx    = bool(rp_detail) and loan_ex > 0
                        total_rp_min_ratio = solve_rp_min(loan_ex, accrued, coll_after_rp)
                        total_rp_maks = solve_rp_max(loan_ex, accrued) if has_rp_trx else 0.0

                        loan_after_rp = max(loan_ex - total_rp_maks, 0)
                        rasio_rp = (loan_after_rp + accrued) / coll_after_rp if coll_after_rp > 0 else None

                        risk_avq = st.session_state['shared_risk_avail_qty']

                        stocks_after_lr   = dict(stocks_after_rp)
                        buy_status_detail = []
                        for stock, bdata in buy_stocks.items():
                            stocks_after_lr[stock] = stocks_after_lr.get(stock, 0) + bdata['lot']
                            avq    = risk_avq.get(stock, 0.0)
                            status = get_loan_status(avq, bdata['lot'])
                            buy_status_detail.append({
                                'stock': stock, 'lot_beli': bdata['lot'],
                                'available_qty': avq, 'status': status,
                            })

                        coll_after_lr, _ = calc_collateral(stocks_after_lr, closing_prices, risk_params_hc)
                        total_buy_val = sum(b['value'] for b in buy_stocks.values())

                        # ── LR final = MIN(LR@63%, Available Limit efektif) ──
                        max_lr_63, avail_efektif, max_lr_final, lr_binding = solve_lr(
                            coll_after_lr, loan_after_rp, accrued, avail_lim, total_rp_maks)
                        numerator_lr = loan_after_rp + accrued + min(total_buy_val, max_lr_final)
                        rasio_lr     = numerator_lr / coll_after_lr if coll_after_lr > 0 else None

                        results[sid] = {
                            'name': name, 'cid': sid_to_cid.get(sid, sid),
                            'loan_existing': loan_ex, 'accrued': accrued, 'avail_limit': avail_lim,
                            'rp_detail': rp_detail, 'total_rp_maks': total_rp_maks, 'total_rp_min': total_rp_min,
                            'total_rp_min_ratio': total_rp_min_ratio,
                            'stocks_op': stocks_op, 'stocks_after_rp': stocks_after_rp, 'stocks_after_lr': stocks_after_lr,
                            'buy_stocks': buy_stocks,
                            'coll_before_rp': coll_before_rp, 'coll_after_rp': coll_after_rp, 'coll_after_lr': coll_after_lr,
                            'loan_after_rp': loan_after_rp, 'rasio_rp': rasio_rp,
                            'total_buy_val': total_buy_val, 'avail_efektif': avail_efektif,
                            'ceiling_lr': avail_efektif, 'rasio_lr': rasio_lr,
                            'max_lr_63': max_lr_63, 'max_lr_65': max_lr_63, 'max_lr_final': max_lr_final,
                            'lr_binding': lr_binding,
                            'buy_status_detail': buy_status_detail,
                        }

                st.session_state['pei_results'] = results
                st.session_state['pei_results_original'] = copy.deepcopy(results)
                st.success("✅ TRX PEI berhasil diproses!")
            except Exception as e:
                st.error(f"❌ Gagal: {e}")
                st.exception(e)
    else:
        st.info("⬆️ Lengkapi file bersama (termasuk Netting Invoice & Credit Limit) + SID Client untuk memproses TRX PEI.")

    # ── TAMPILKAN HASIL TRX PEI ───────────────────────────────
    if st.session_state.get('pei_results'):
        results        = st.session_state['pei_results']
        closing_prices = st.session_state['shared_closing_prices']
        risk_params_hc = st.session_state['shared_risk_params']

        n_has_rp   = sum(1 for v in results.values() if v['total_rp_maks'] > 0)
        n_has_lr   = sum(1 for v in results.values() if v['total_buy_val'] > 0)
        n_rp_lolos = sum(1 for v in results.values() if v['rasio_rp'] is not None and v['rasio_rp'] < 0.65)
        n_lr_lolos = sum(1 for v in results.values() if v['rasio_lr'] is not None and v['rasio_lr'] < 0.65)
        n_sell_trx = sum(len(v['rp_detail'])  for v in results.values())
        n_buy_trx  = sum(len(v['buy_stocks']) for v in results.values())

        m1,m2,m3,m4,m5 = st.columns(5)
        m1.metric(
            "Total Nasabah PEI", len(results),
            help="Jumlah SID yang tercatat sebagai nasabah PEI.\n\n"
                 "**Sumber:** file *SID Client* (CID dicocokkan → SID)."
        )
        m2.metric(
            "Nasabah Ada RP", n_has_rp,
            help=f"Nasabah PEI dengan kewajiban Repayment (RP Max > 0), "
                 f"dari total **{n_sell_trx}** baris transaksi Sell (saham × SID).\n\n"
                 "**Sumber:** transaksi *sell* di file *Netting Invoice*, "
                 "dicocokkan dengan saham di file *Outstanding Position*."
        )
        m3.metric(
            "Nasabah Ada LR", n_has_lr,
            help=f"Nasabah PEI dengan transaksi Beli margin (Nilai Beli > 0), "
                 f"dari total **{n_buy_trx}** baris transaksi Buy (saham × SID).\n\n"
                 "**Sumber:** transaksi *buy* di file *Netting Invoice*."
        )
        m4.metric(
            "RP Lolos Rasio", n_rp_lolos,
            help="Nasabah dengan Rasio After RP < 65%.\n\n"
                 "**Dihitung:** (Loan setelah RP + Accrued) ÷ Collateral setelah RP."
        )
        m5.metric(
            "LR Lolos Rasio", n_lr_lolos,
            help="Nasabah dengan Rasio After LR < 65%.\n\n"
                 "**Dihitung:** (Loan setelah RP + Accrued + Nilai Beli dicairkan) ÷ Collateral setelah LR."
        )

        sub_rp, sub_lr, sub_sim, sub_sum, sub_exp = st.tabs([
            "📤 Langkah 1 — Repayment (RP)",
            "📥 Langkah 2 — Loan Request (LR)",
            "🎛️ Simulator RP → LR",
            "📋 Summary",
            "📥 Export",
        ])

        with sub_rp:
            st.info("💡 Input RP **terlebih dahulu**. RP dipicu oleh transaksi **sell** saham marginable — "
                     "begitu terpicu, kewajibannya adalah RP Max (lunasi penuh Loan O/S + Accrued). "
                     "RP Min (solve-to-65%) ditampilkan sebagai referensi batas bawah saja.")
            rp_rows = []
            for sid, d in results.items():
                if d['total_rp_maks'] <= 0: continue
                for rd in d['rp_detail']:
                    if not rd['ada_di_op']: continue
                    rasio_val    = f"{d['rasio_rp']*100:.2f}%" if d['rasio_rp'] is not None else "N/A"
                    status_rasio = "✅ LOLOS" if d['rasio_rp'] is not None and d['rasio_rp'] < 0.65 else "❌ GAGAL"
                    rp_rows.append({
                        'SID': sid, 'Nama': d['name'], 'Saham': rd['stock'],
                        'Lot Jual': int(rd['lot_sell']), 'Lot OP': int(rd['lot_op']),
                        'Lot Keluar': int(rd['lot_keluar']), 'Harga': rd['price'],
                        'Nilai Jual (ref)': rd['rp_min'],
                        'RP Min (ref, target 65%)': d.get('total_rp_min_ratio', 0),
                        'Total RP Max (rekomendasi)': d['total_rp_maks'],
                        'Loan After RP': d['loan_after_rp'], 'Coll After RP': d['coll_after_rp'],
                        'Rasio After RP': rasio_val, 'Status': status_rasio,
                    })
            if rp_rows:
                df_rp = pd.DataFrame(rp_rows)
                def color_rp(val):
                    if val == '✅ LOLOS': return 'background-color:#d4edda;color:#155724'
                    if val == '❌ GAGAL': return 'background-color:#f8d7da;color:#721c24'
                    return ''
                st.dataframe(df_rp.style.map(color_rp, subset=['Status']), use_container_width=True)
            else:
                st.info("Tidak ada transaksi jual kemarin, sehingga tidak ada kewajiban RP.")

        with sub_lr:
            st.info("💡 Input LR **setelah RP selesai**. LR Final = MIN(rasio kolateral 63%, Available Limit efektif).")
            lr_rows = []
            for sid, d in results.items():
                if d['max_lr_final'] <= 0 and d['total_buy_val'] <= 0: continue
                rasio_val    = f"{d['rasio_lr']*100:.2f}%" if d['rasio_lr'] is not None else "N/A"
                status_rasio = "✅ LOLOS" if d['rasio_lr'] is not None and d['rasio_lr'] < 0.65 else "❌ GAGAL"
                if d['buy_stocks']:
                    for stock, bdata in d['buy_stocks'].items():
                        status_info = next((s for s in d.get('buy_status_detail', []) if s['stock'] == stock), {})
                        lr_rows.append({
                            'SID': sid, 'Nama': d['name'], 'Saham Beli': stock,
                            'Lot Beli': int(bdata['lot']), 'Nilai Beli': bdata['value'],
                            'Available Qty': status_info.get('available_qty', 0),
                            'Status Margin': status_info.get('status', ''),
                            'Avail Efektif': d['avail_efektif'], 'LR @63%': d['max_lr_63'],
                            'Loan After RP': d['loan_after_rp'], 'Coll After LR': d['coll_after_lr'],
                            'Rasio LR': rasio_val, 'Max LR Final': d['max_lr_final'],
                            'Binding': d.get('lr_binding',''), 'Status': status_rasio,
                        })
                else:
                    lr_rows.append({
                        'SID': sid, 'Nama': d['name'], 'Saham Beli': '-',
                        'Lot Beli': 0, 'Nilai Beli': 0,
                        'Available Qty': 0, 'Status Margin': '',
                        'Avail Efektif': d['avail_efektif'], 'LR @63%': d['max_lr_63'],
                        'Loan After RP': d['loan_after_rp'], 'Coll After LR': d['coll_after_lr'],
                        'Rasio LR': "-", 'Max LR Final': d['max_lr_final'],
                        'Binding': d.get('lr_binding',''), 'Status': "Kapasitas tersedia (belum ada transaksi beli)",
                    })
            if lr_rows:
                df_lr = pd.DataFrame(lr_rows)
                def color_lr(val):
                    if val == '✅ LOLOS': return 'background-color:#d4edda;color:#155724'
                    if '❌' in str(val):  return 'background-color:#f8d7da;color:#721c24'
                    return ''
                st.dataframe(df_lr.style.map(color_lr, subset=['Status']), use_container_width=True)
            else:
                st.info("Tidak ada kapasitas LR untuk nasabah manapun.")

        with sub_sim:
            st.subheader("🎛️ Simulator — Ubah Nilai RP, Lihat Dampak ke LR")
            sid_options = [s for s,d in results.items() if d['total_rp_maks'] > 0 or d['total_buy_val'] > 0 or d['max_lr_final'] > 0]
            if not sid_options:
                st.warning("Tidak ada nasabah dengan transaksi.")
            else:
                sel = st.selectbox("Pilih Nasabah:", sid_options,
                    format_func=lambda s: f"{s} — {results[s]['name']}", key='pei_sim_sel')
                d = results[sel]
                if d.get('is_simulated'):
                    st.success("✏️ Menggunakan nilai simulasi yang sudah disimpan.")
                c1,c2,c3 = st.columns(3)
                c1.metric("Loan Outstanding", fmt_rp(d['loan_existing']))
                c2.metric("Collateral Awal",  fmt_rp(d['coll_before_rp']))
                c3.metric("Current Ratio", f"{d['loan_existing']/d['coll_before_rp']*100:.2f}%" if d['coll_before_rp'] > 0 else "N/A")

                rp_ref_rows = [rd for rd in d['rp_detail'] if rd['ada_di_op']]
                if rp_ref_rows:
                    st.caption("Ringkasan saham RP (referensi — lot keluar tetap dihitung penuh)")
                    df_rp_ref = pd.DataFrame([{
                        'Saham': rd['stock'], 'Lot Jual': int(rd['lot_sell']),
                        'Lot OP': int(rd['lot_op']), 'Lot Keluar': int(rd['lot_keluar']),
                        'Harga': rd['price'], 'Nilai Jual (ref)': rd['rp_min'],
                    } for rd in rp_ref_rows])
                    st.dataframe(df_rp_ref, use_container_width=True, hide_index=True)

                num_key = f"pei_rp_total_{sel}"
                if num_key not in st.session_state:
                    st.session_state[num_key] = float(d['total_rp_maks'])

                st.caption(f"RP Max (rekomendasi, default): {fmt_rp(d['total_rp_maks'])} "
                           f"| RP Min (ref, target 65%): {fmt_rp(d.get('total_rp_min_ratio', 0))}. "
                           f"Ubah nilai di bawah untuk mensimulasikan RP berbeda.")
                total_rp_sim = st.number_input(
                    "Total RP Adjust", step=1_000_000.0, format="%.0f", key=num_key,
                )

                stocks_ar_sim = dict(d['stocks_op'])
                for rd in rp_ref_rows:
                    if rd['lot_keluar'] > 0:
                        stocks_ar_sim[rd['stock']] = stocks_ar_sim.get(rd['stock'], 0) - rd['lot_keluar']
                        if stocks_ar_sim.get(rd['stock'], 0) <= 0:
                            stocks_ar_sim.pop(rd['stock'], None)
                coll_ar_sim, _ = calc_collateral(stocks_ar_sim, closing_prices, risk_params_hc)
                loan_ar_sim    = max(d['loan_existing'] - total_rp_sim, 0)
                rasio_rp_sim   = (loan_ar_sim + d['accrued']) / coll_ar_sim if coll_ar_sim > 0 else None

                st.divider()
                r1,r2,r3,r4 = st.columns(4)
                r1.metric("Total RP", fmt_rp(total_rp_sim))
                r2.metric("Loan After RP", fmt_rp(loan_ar_sim))
                r3.metric("Coll After RP", fmt_rp(coll_ar_sim))
                rp_ok = rasio_rp_sim is not None and rasio_rp_sim < 0.65
                r4.metric("Rasio RP", f"{rasio_rp_sim*100:.2f}%" if rasio_rp_sim else "N/A",
                    delta="✅ LOLOS" if rp_ok else "❌ GAGAL",
                    delta_color="normal" if rp_ok else "inverse")

                stocks_al_sim = dict(stocks_ar_sim)
                for stock, bdata in d['buy_stocks'].items():
                    stocks_al_sim[stock] = stocks_al_sim.get(stock,0) + bdata['lot']
                coll_al_sim, _ = calc_collateral(stocks_al_sim, closing_prices, risk_params_hc)

                max63_sim, avail_eff_sim, max_final_sim, binding_sim = solve_lr(
                    coll_al_sim, loan_ar_sim, d['accrued'], d['avail_limit'], total_rp_sim)
                num_lr_sim   = loan_ar_sim + d['accrued'] + min(d['total_buy_val'], max_final_sim)
                rasio_lr_sim = num_lr_sim / coll_al_sim if coll_al_sim > 0 else None

                st.subheader("Dampak ke LR")
                l1,l2,l3 = st.columns(3)
                l1.metric("Avail Efektif",  fmt_rp(avail_eff_sim))
                l2.metric("LR @63%",        fmt_rp(max63_sim))
                l3.metric("Coll After LR",  fmt_rp(coll_al_sim))
                l4,l5,l6 = st.columns(3)
                l4.metric("Numerator LR",   fmt_rp(num_lr_sim))
                lr_ok = rasio_lr_sim is not None and rasio_lr_sim < 0.65
                l5.metric("Rasio LR", f"{rasio_lr_sim*100:.2f}%" if rasio_lr_sim else "N/A",
                    delta="✅ LOLOS" if lr_ok else "❌ Perlu dipotong",
                    delta_color="normal" if lr_ok else "inverse")
                l6.metric("Max LR Final", fmt_rp(max_final_sim), delta=f"Binding: {binding_sim}")

                st.divider()
                sv1, sv2, sv3 = st.columns(3)
                if sv1.button("💾 Simpan Simulasi", use_container_width=True, type="primary", key=f'pei_save_{sel}'):
                    updated = copy.deepcopy(st.session_state['pei_results'][sel])
                    updated['is_simulated']    = True
                    updated['loan_after_rp']   = loan_ar_sim
                    updated['coll_after_rp']   = coll_ar_sim
                    updated['coll_after_lr']   = coll_al_sim
                    updated['rasio_rp']        = rasio_rp_sim
                    updated['rasio_lr']        = rasio_lr_sim
                    updated['avail_efektif']   = avail_eff_sim
                    updated['ceiling_lr']      = avail_eff_sim
                    updated['max_lr_63']       = max63_sim
                    updated['max_lr_65']       = max63_sim
                    updated['max_lr_final']    = max_final_sim
                    updated['lr_binding']      = binding_sim
                    updated['total_rp_maks']   = total_rp_sim
                    updated['stocks_after_rp'] = stocks_ar_sim
                    updated['stocks_after_lr'] = stocks_al_sim
                    st.session_state['pei_results'][sel] = updated
                    st.success(f"✅ Simulasi {sel} disimpan — cek tab Summary/Export untuk hasil terbaru.")
                    st.rerun()
                if sv2.button("↩️ Reset Nasabah Ini", use_container_width=True, key=f'pei_reset1_{sel}'):
                    orig = st.session_state.get('pei_results_original', {})
                    if sel in orig:
                        st.session_state['pei_results'][sel] = copy.deepcopy(orig[sel])
                        st.session_state.pop(num_key, None)
                        st.rerun()
                if sv3.button("🔄 Reset Semua", use_container_width=True, key='pei_reset_all'):
                    orig = st.session_state.get('pei_results_original', {})
                    if orig:
                        st.session_state['pei_results'] = copy.deepcopy(orig)
                        st.rerun()

        with sub_sum:
            st.subheader("📋 Summary Semua Nasabah PEI")
            summary_rows = []
            for sid, d in results.items():
                summary_rows.append({
                    'SID': sid, 'Nama': d['name'],
                    'Loan Existing': d['loan_existing'], 'Coll Awal': d['coll_before_rp'],
                    'Current Ratio': f"{d['loan_existing']/d['coll_before_rp']*100:.2f}%" if d['coll_before_rp'] > 0 else "-",
                    'RP Min (ref)': d.get('total_rp_min_ratio', 0),
                    'RP Max (rekomendasi)': d['total_rp_maks'],
                    'Loan After RP': d['loan_after_rp'], 'Coll After RP': d['coll_after_rp'],
                    'Rasio RP': f"{d['rasio_rp']*100:.2f}%" if d['rasio_rp'] is not None else "-",
                    'Status RP': "✅" if d['rasio_rp'] is not None and d['rasio_rp'] < 0.65 else ("-" if d['total_rp_maks']==0 else "❌"),
                    'Avail Efektif': d['avail_efektif'], 'Coll After LR': d['coll_after_lr'],
                    'Max LR Final': d['max_lr_final'], 'Binding': d.get('lr_binding',''),
                    'Rasio LR': f"{d['rasio_lr']*100:.2f}%" if d['rasio_lr'] is not None else "-",
                    'Status LR': "✅" if d['rasio_lr'] is not None and d['rasio_lr'] < 0.65 else ("-" if d['total_buy_val']==0 else "❌"),
                })
            st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

        with sub_exp:
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as wr:
                rp_exp = []
                for sid, d in results.items():
                    for rd in d['rp_detail']:
                        if not rd['ada_di_op']: continue
                        rp_exp.append({
                            'SID': sid, 'Nama': d['name'], 'Saham': rd['stock'],
                            'Lot Jual': int(rd['lot_sell']), 'Lot OP': int(rd['lot_op']),
                            'Lot Keluar': int(rd['lot_keluar']), 'Harga': rd['price'],
                            'Nilai Jual (ref)': rd['rp_min'],
                            'RP Min (ref)': d.get('total_rp_min_ratio', 0),
                            'Total RP Max (rekomendasi)': d['total_rp_maks'],
                            'Loan After RP': d['loan_after_rp'], 'Coll After RP': d['coll_after_rp'],
                            'Rasio After RP': f"{d['rasio_rp']*100:.2f}%" if d['rasio_rp'] is not None else "N/A",
                        })
                pd.DataFrame(rp_exp).to_excel(wr, sheet_name='Repayment (RP)', index=False)
                lr_exp = []
                for sid, d in results.items():
                    if d['max_lr_final'] <= 0: continue
                    lr_exp.append({
                        'SID': sid, 'Nama': d['name'],
                        'RP Max (rekomendasi)': d['total_rp_maks'], 'Avail Limit': d['avail_limit'],
                        'Avail Efektif': d['avail_efektif'], 'LR @63%': d['max_lr_63'],
                        'Loan After RP': d['loan_after_rp'], 'Coll After LR': d['coll_after_lr'],
                        'Rasio LR': f"{d['rasio_lr']*100:.2f}%" if d['rasio_lr'] is not None else "N/A",
                        'Max LR Final': d['max_lr_final'], 'Binding': d.get('lr_binding',''),
                    })
                pd.DataFrame(lr_exp).to_excel(wr, sheet_name='Loan Request (LR)', index=False)
                pd.DataFrame(summary_rows).to_excel(wr, sheet_name='Summary', index=False)
            st.download_button("📥 Download Hasil_TRX_PEI.xlsx", out.getvalue(),
                "Hasil_TRX_PEI.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ══════════════════════════════════════════════════════════════════
# TAB 2 — VALIDASI MNC
# ══════════════════════════════════════════════════════════════════
with tab_mnc:
    st.info("Sistem Validasi Repayment & Loan Request nasabah PEI: **RP dulu → LR setelahnya**")

    st.subheader("📂 File Tambahan (Validasi MNC)")
    st.caption("Credit Limit sekarang dipakai bersama (upload di bagian atas). Di sini cukup Hasil MNC.")
    hasil_file = st.file_uploader("A. Hasil MNC (.xlsx)", type=["xlsx"], key='mnc_hasil')

    mnc_ready = shared_ready and hasil_file

    # ── HELPERS VALIDASI MNC ──────────────────────────────────
    def load_hasil_mnc(uploaded_file):
        xls = pd.ExcelFile(uploaded_file)
        if "Repayment (RP)" in xls.sheet_names:
            df_sell = pd.read_excel(xls, sheet_name="Repayment (RP)", header=0)
            if 'NETT' in df_sell.columns:
                df_sell = df_sell[~df_sell['NETT'].astype(str).str.contains('EXCLUDED', na=False)].copy()
        elif "Sell Aktif (tanpa excluded)" in xls.sheet_names:
            df_sell = pd.read_excel(xls, sheet_name="Sell Aktif (tanpa excluded)", header=0)
        elif "Sell Aktif" in xls.sheet_names:
            df_sell = pd.read_excel(xls, sheet_name="Sell Aktif", header=0)
        elif "Sell (Repayment)" in xls.sheet_names:
            df_sell = pd.read_excel(xls, sheet_name="Sell (Repayment)", header=0)
        else:
            st.error(f"❌ Sheet RP tidak ditemukan. Sheet tersedia: {xls.sheet_names}"); st.stop()

        if "Loan Request (LR)" in xls.sheet_names:
            df_buy = pd.read_excel(xls, sheet_name="Loan Request (LR)", header=0)
        elif "Buy (Loan)" in xls.sheet_names:
            df_buy = pd.read_excel(xls, sheet_name="Buy (Loan)", header=0)
        else:
            st.error(f"❌ Sheet LR tidak ditemukan. Sheet tersedia: {xls.sheet_names}"); st.stop()

        df_buy_raw = df_buy.copy()
        return df_sell, df_buy, df_buy_raw

    def validate_sid_mnc(sid, op_data, cl_data, sell_regular, margin_buy,
                          closing_prices, risk_params, df_sell, df_buy, risk_avq):
        op  = op_data.get(sid, {"loan_existing":0,"accrued_interest":0,"name":sid,"stocks":{}})
        cl  = cl_data.get(sid, {"available_limit":0,"name":sid})
        loan_ex   = op["loan_existing"]
        accrued   = op["accrued_interest"]
        avail_lim = cl["available_limit"]
        name      = op.get("name") or cl.get("name") or sid
        stocks_op = op.get("stocks", {})

        sell_stocks = sell_regular.get(sid, {})
        buy_stocks  = margin_buy.get(sid, {})

        checks = []
        def add(label, passed, detail=""):
            checks.append({"label": label, "passed": passed, "detail": detail})

        rp_detail       = []
        stocks_after_rp = dict(stocks_op)
        total_rp_min    = 0.0
        has_rp          = bool(sell_stocks)

        if not has_rp or loan_ex <= 0:
            msg = "⏭ Dilewati — Loan Existing = 0" if loan_ex <= 0 else "Tidak ada transaksi jual"
            add("RP-1. Pengecekan Saham Jual", True, msg)
            add("RP-2. Lot Sell ≤ Lot di OP",  True, msg)
            add("RP-3. Rasio After RP < 65%",  True, msg)
            total_rp_maks = 0.0
            total_rp_min_ratio = 0.0
        else:
            rp1_detail = [s for s in sell_stocks if stocks_op.get(s, 0) == 0]
            if rp1_detail:
                add("RP-1. Pengecekan Saham Jual", True, f"ℹ️ Saham reguler (bukan OP): {', '.join(rp1_detail)}")
            else:
                add("RP-1. Pengecekan Saham Jual", True, "✅ Semua saham jual ada di OP")

            rp2_detail = []
            for stock, sdata in sell_stocks.items():
                lot_sell   = sdata['lot']
                lot_op     = stocks_op.get(stock, 0)
                lot_keluar = min(lot_sell, lot_op)
                price      = closing_prices.get(stock, 0)
                rp_min     = lot_keluar * price
                ada        = lot_op > 0
                if lot_sell > lot_op and ada:
                    rp2_detail.append(f"{stock}: {lot_sell:,.0f} → {lot_op:,.0f}")
                if ada:
                    total_rp_min += rp_min
                    rp_detail.append({'stock': stock, 'lot_sell': lot_sell, 'lot_op': lot_op,
                                      'lot_keluar': lot_keluar, 'price': price, 'rp_min': rp_min})
                    stocks_after_rp[stock] = stocks_after_rp.get(stock, 0) - lot_keluar
                    if stocks_after_rp[stock] <= 0: del stocks_after_rp[stock]

            coll_after_rp_tmp, _ = calc_collateral(stocks_after_rp, closing_prices, risk_params)
            # RP dipicu oleh transaksi sell → kewajiban = RP Max (lunasi penuh loan+accrued).
            # RP Min (solve-to-65%) tetap dihitung & ditampilkan sbg referensi batas bawah saja.
            total_rp_min_ratio = solve_rp_min(loan_ex, accrued, coll_after_rp_tmp)
            total_rp_maks = solve_rp_max(loan_ex, accrued)

            add("RP-2. Lot Sell ≤ Lot di OP", True,
                ("⚠️ Auto-adjusted: " + "; ".join(rp2_detail)) if rp2_detail else
                f"Total RP Max (rekomendasi): {fmt_rp(total_rp_maks)} | RP Min (ref): {fmt_rp(total_rp_min_ratio)}")

            coll_after_rp, _ = calc_collateral(stocks_after_rp, closing_prices, risk_params)
            loan_after_rp    = max(loan_ex - total_rp_maks, 0)
            numerator_rp     = loan_after_rp + accrued
            rasio_rp = numerator_rp / coll_after_rp if coll_after_rp > 0 else None
            if rasio_rp is not None:
                add("RP-3. Rasio After RP < 65%", rasio_rp < RATIO_THRESHOLD,
                    f"Rasio: {fmt_pct(rasio_rp)} | Numerator: {fmt_rp(numerator_rp)} | Coll: {fmt_rp(coll_after_rp)}")
            elif numerator_rp <= 0:
                add("RP-3. Rasio After RP < 65%", True, "Numerator ≤ 0")
            else:
                add("RP-3. Rasio After RP < 65%", False, "Collateral = 0")

        has_lr = bool(buy_stocks)
        coll_after_rp2, _ = calc_collateral(stocks_after_rp, closing_prices, risk_params)
        loan_after_rp2    = max(loan_ex - total_rp_maks, 0)

        stocks_after_lr   = dict(stocks_after_rp)
        buy_status_detail = []
        for stock, bdata in buy_stocks.items():
            stocks_after_lr[stock] = stocks_after_lr.get(stock, 0) + bdata['lot']
            avq    = risk_avq.get(stock, 0.0)
            status = get_loan_status(avq, bdata['lot'])
            buy_status_detail.append({
                'stock': stock, 'lot_beli': bdata['lot'],
                'available_qty': avq, 'status': status,
            })
        coll_after_lr, _ = calc_collateral(stocks_after_lr, closing_prices, risk_params)

        total_buy_val  = sum(b['value'] for b in buy_stocks.values())

        # ── LR final = MIN(LR@63%, Available Limit efektif) ──
        max_lr_63, avail_efektif, max_lr_final, lr_binding = solve_lr(
            coll_after_lr, loan_after_rp2, accrued, avail_lim, total_rp_maks)

        if not has_lr:
            add("LR-1. Volume Buy ≤ Available Qty", True, "Tidak ada Loan Request")
            add("LR-2. Kapasitas LR (63% / Avail Limit)", True, "Tidak ada Loan Request")
            add("LR-3. Nilai Beli ≤ Max LR Final",  True, "Tidak ada Loan Request")
        else:
            lr1_pass   = True
            lr1_detail = []
            if df_buy is not None and not df_buy.empty and df_buy.shape[1] > 0:
                try:
                    buy_rows = df_buy[df_buy.iloc[:, 0].astype(str) == sid]
                except IndexError:
                    buy_rows = pd.DataFrame()
            else:
                buy_rows = pd.DataFrame()

            if buy_rows.empty and df_buy is not None and (df_buy.empty or df_buy.shape[1] == 0):
                lr1_detail.append("⚠️ Sheet Loan Request (LR) di Hasil MNC kosong/tidak sesuai format — LR-1 dilewati")
            for _, row in buy_rows.iterrows():
                vol = pd.to_numeric(row.iloc[13] if len(row) > 13 else 0, errors='coerce') or 0
                avq = pd.to_numeric(row.iloc[4]  if len(row) > 4  else 0, errors='coerce') or 0
                stk = str(row.iloc[1]) if len(row) > 1 else ''
                if avq == 0:
                    lr1_pass = False; lr1_detail.append(f"{stk}: DIBATALKAN (Available=0)")
                elif vol > avq:
                    lr1_pass = False; lr1_detail.append(f"{stk}: Vol {vol:,.0f} > Avail {avq:,.0f}")
            add("LR-1. Volume Buy ≤ Available Qty", lr1_pass,
                "; ".join(lr1_detail) if lr1_detail else f"Total Nilai Beli: {fmt_rp(total_buy_val)}")
            add("LR-2. Kapasitas LR (63% / Avail Limit)", True,
                f"LR @63%: {fmt_rp(max_lr_63)} | Avail Efektif: {fmt_rp(avail_efektif)} | "
                f"Max LR Final: {fmt_rp(max_lr_final)} (binding: {lr_binding})")

            # LR-3: apakah nilai beli aktual muat dalam kapasitas LR final yg disetujui
            lr3_pass   = total_buy_val <= max_lr_final
            numerator_lr = loan_after_rp2 + accrued + min(total_buy_val, max_lr_final)
            rasio_lr = numerator_lr / coll_after_lr if coll_after_lr > 0 else None
            detail_str = (f"Nilai Beli: {fmt_rp(total_buy_val)} vs Max LR Final: {fmt_rp(max_lr_final)} | "
                          f"Rasio jika full: {fmt_pct(rasio_lr) if rasio_lr is not None else 'N/A'}")
            if not lr3_pass:
                detail_str += f" || ⚠️ Potong Loan Request ke: {fmt_rp(max_lr_final)}"
            add("LR-3. Nilai Beli ≤ Max LR Final", lr3_pass, detail_str)

        return {
            "name": name, "checks": checks, "has_rp": has_rp, "has_lr": has_lr,
            "rp_skipped": loan_ex <= 0,
            "loan_existing": loan_ex, "accrued": accrued, "avail_limit": avail_lim,
            "stocks_op": stocks_op, "stocks_after_rp": stocks_after_rp, "stocks_after_lr": stocks_after_lr,
            "rp_detail": rp_detail, "total_rp_maks": total_rp_maks, "total_rp_min": total_rp_min,
            "total_rp_min_ratio": total_rp_min_ratio,
            "loan_after_rp": loan_after_rp2,
            "coll_before_rp": calc_collateral(stocks_op, closing_prices, risk_params)[0],
            "coll_after_rp": coll_after_rp2, "coll_after_lr": coll_after_lr,
            "ceiling_lr": avail_efektif, "max_lr_63": max_lr_63, "max_lr_65": max_lr_63,
            "max_lr_final": max_lr_final, "lr_binding": lr_binding,
            "avail_efektif": avail_efektif, "total_buy_val": total_buy_val,
            "buy_status_detail": buy_status_detail,
        }

    def lolos_rp(data):
        if data.get("rp_skipped"): return False
        if not data.get("has_rp"): return False
        return all(c["passed"] for c in data["checks"] if c["label"].startswith("RP-"))

    def lolos_lr(data):
        if not data.get("has_lr"): return False
        return all(c["passed"] for c in data["checks"] if c["label"].startswith("LR-"))

    # ── TOMBOL VALIDASI ───────────────────────────────────────
    if mnc_ready:
        if st.button("▶ Jalankan Validasi MNC", type="primary", use_container_width=True, key='btn_mnc'):
            with st.spinner("⚙️ Memproses Validasi MNC..."):
                df_sell, df_buy, df_buy_raw = load_hasil_mnc(hasil_file)
                closing_prices = st.session_state['shared_closing_prices']
                risk_params    = st.session_state['shared_risk_params']
                op_data        = st.session_state['shared_op_data']
                margin_buy     = st.session_state['shared_net_buy']    # net beli hasil Netting Invoice
                sell_regular   = st.session_state['shared_net_sell']   # net jual hasil Netting Invoice
                cl_data        = st.session_state['shared_cl_data']    # Available Limit — dari file Credit Limit bersama

                risk_avq = st.session_state['shared_risk_avail_qty']

                all_sids = sorted(set(list(op_data.keys()) + list(cl_data.keys()) +
                                      list(margin_buy.keys()) + list(sell_regular.keys())))

                sid_results = {}
                for sid in all_sids:
                    sid_results[sid] = validate_sid_mnc(
                        sid, op_data, cl_data, sell_regular, margin_buy,
                        closing_prices, risk_params, df_sell, df_buy, risk_avq)

                total_lr_all = sum(d['max_lr_final'] for d in sid_results.values() if lolos_lr(d))
                total_rp_all = sum(d['total_rp_maks'] for d in sid_results.values() if lolos_rp(d))
                global_result = {
                    "passed": (CREDIT_LIMIT_PARTISIPAN + total_rp_all) > total_lr_all,
                    "detail": (f"CL Partisipan: {fmt_rp(CREDIT_LIMIT_PARTISIPAN)} + "
                               f"RP: {fmt_rp(total_rp_all)} = "
                               f"{fmt_rp(CREDIT_LIMIT_PARTISIPAN+total_rp_all)} | LR: {fmt_rp(total_lr_all)}"),
                    "total_rp": total_rp_all, "total_lr": total_lr_all,
                }

            st.session_state.update({
                'sid_results': sid_results, 'global_result': global_result,
                'df_sell_edited': df_sell.copy(), 'df_buy': df_buy.copy(),
                'df_buy_adjusted': df_buy_raw.copy(),
                'sid_results_original': copy.deepcopy(sid_results),
            })
            st.success("✅ Validasi Selesai!")
    else:
        st.info("⬆️ Lengkapi file bersama (termasuk Credit Limit) + Hasil MNC untuk menjalankan validasi.")

    # ── TAMPILKAN HASIL VALIDASI MNC ─────────────────────────
    if st.session_state.get('sid_results'):
        sid_results   = st.session_state['sid_results']
        global_result = st.session_state['global_result']
        closing_prices= st.session_state['shared_closing_prices']
        risk_params   = st.session_state['shared_risk_params']
        margin_buy    = st.session_state['shared_net_buy']
        sell_regular  = st.session_state['shared_net_sell']
        op_data       = st.session_state['shared_op_data']
        cl_data       = st.session_state.get('shared_cl_data', {})

        n_rp     = sum(1 for v in sid_results.values() if v.get('has_rp'))
        n_lr     = sum(1 for v in sid_results.values() if v.get('has_lr'))
        n_rp_ok  = sum(1 for v in sid_results.values() if lolos_rp(v))
        n_lr_ok  = sum(1 for v in sid_results.values() if lolos_lr(v))
        n_rp_fail= sum(1 for v in sid_results.values() if v.get('has_rp') and not v.get('rp_skipped') and not lolos_rp(v))
        n_lr_fail= sum(1 for v in sid_results.values() if v.get('has_lr') and not lolos_lr(v))

        m1,m2,m3,m4,m5,m6,m7 = st.columns(7)
        m1.metric("Total Nasabah",  len(sid_results))
        m2.metric("Ada RP",         n_rp)
        m3.metric("Ada LR",         n_lr)
        m4.metric("RP Lolos",       n_rp_ok)
        m5.metric("LR Lolos",       n_lr_ok)
        m6.metric("RP/LR Gagal",    f"{n_rp_fail}/{n_lr_fail}",
                  delta=f"-{n_rp_fail+n_lr_fail}" if (n_rp_fail+n_lr_fail) else None, delta_color="inverse")
        m7.metric("CL Partisipan",  "✅ LOLOS" if global_result["passed"] else "❌ GAGAL")

        st.divider()

        tab_rp, tab_lr, tab_sim, tab_global, tab_gagal, tab_adj, tab_exp = st.tabs([
            "📤 Langkah 1 — RP", "📥 Langkah 2 — LR",
            "🎛️ Simulator", "🌐 Limit Participant",
            "❌ Nasabah Gagal", "⚡ Auto-Adjust LR", "📥 Export",
        ])

        with tab_rp:
            st.info("💡 Input RP **terlebih dahulu**. RP dipicu oleh transaksi sell — kewajiban = RP Max (lunasi penuh).")
            for sid, data in sid_results.items():
                if not data.get('has_rp') or data.get('rp_skipped') or data['total_rp_maks'] <= 0: continue
                ok   = lolos_rp(data)
                icon = "✅" if ok else "❌"
                if data.get('is_simulated'): icon += " ✏️"
                with st.expander(f"{icon} {sid} — {data['name']}  |  Max LR Final: {fmt_rp(data['max_lr_final'])}", expanded=not ok):
                    c1,c2,c3,c4 = st.columns(4)
                    c1.metric("Loan After RP",   fmt_rp(data['loan_after_rp']))
                    c2.metric("Avail Efektif",   fmt_rp(data['avail_efektif']))
                    c3.metric("LR @63%",         fmt_rp(data['max_lr_63']))
                    c4.metric("Coll After LR",   fmt_rp(data['coll_after_lr']))
                    c5,c6,_,_ = st.columns(4)
                    c5.metric("RP Min (ref, 65%)", fmt_rp(data.get('total_rp_min_ratio', 0)))
                    c6.metric("RP Max (rekomendasi)", fmt_rp(data['total_rp_maks']))
                    if data.get('buy_status_detail'):
                        df_buy_status = pd.DataFrame(data['buy_status_detail']).rename(columns={
                            'stock': 'Saham', 'lot_beli': 'Lot Beli',
                            'available_qty': 'Available Qty', 'status': 'Status Margin'
                        })
                        st.dataframe(df_buy_status, use_container_width=True, hide_index=True)
                    if data['rp_detail']:
                        df_rd = pd.DataFrame([{
                            'Saham': r['stock'], 'Lot Jual': int(r['lot_sell']),
                            'Lot OP': int(r['lot_op']), 'Lot Keluar': int(r['lot_keluar']),
                            'Harga': r['price'], 'Nilai Jual (ref)': r['rp_min']
                        } for r in data['rp_detail']])
                        st.dataframe(df_rd, use_container_width=True, hide_index=True)
                    for c in data['checks']:
                        if not c['label'].startswith('RP-'): continue
                        if c['passed']: st.success(f"✅ **{c['label']}** {c['detail']}")
                        else:           st.error(  f"❌ **{c['label']}** {c['detail']}")

        with tab_lr:
            st.info("💡 Input LR **setelah RP selesai**. LR Final = MIN(LR@63%, Available Limit efektif).")
            for sid, data in sid_results.items():
                if not data.get('has_lr'): continue
                ok   = lolos_lr(data)
                icon = "✅" if ok else "❌"
                if data.get('is_simulated'): icon += " ✏️"
                with st.expander(f"{icon} {sid} — {data['name']}  |  Max LR Final: {fmt_rp(data['max_lr_final'])}", expanded=not ok):
                    c1,c2,c3,c4 = st.columns(4)
                    c1.metric("Loan After RP",   fmt_rp(data['loan_after_rp']))
                    c2.metric("Avail Efektif",   fmt_rp(data['avail_efektif']))
                    c3.metric("LR @63%",         fmt_rp(data['max_lr_63']))
                    c4.metric("Coll After LR",   fmt_rp(data['coll_after_lr']))
                    c5,c6,_,_ = st.columns(4)
                    c5.metric("Max LR Final",    fmt_rp(data['max_lr_final']))
                    c6.metric("Binding",         data.get('lr_binding',''))
                    if not ok:
                        st.warning(f"💡 Potong LR ke: **{fmt_rp(data['max_lr_final'])}**")
                    for c in data['checks']:
                        if not c['label'].startswith('LR-'): continue
                        if c['passed']: st.success(f"✅ **{c['label']}** {c['detail']}")
                        else:           st.error(  f"❌ **{c['label']}** {c['detail']}")

        with tab_sim:
            st.subheader("🎛️ Simulator — Ubah Nilai RP, Lihat Dampak ke LR")
            sid_options = [s for s,d in sid_results.items() if d.get('has_rp') or d.get('has_lr')]
            if not sid_options:
                st.warning("Tidak ada nasabah dengan transaksi.")
            else:
                sel_sid = st.selectbox("Pilih Nasabah:", sid_options,
                    format_func=lambda s: f"{s} — {sid_results[s]['name']}", key='mnc_sim_sel')
                d = st.session_state['sid_results'][sel_sid]
                if d.get('is_simulated'):
                    st.success("✏️ Menggunakan nilai simulasi yang sudah disimpan.")
                c1,c2,c3 = st.columns(3)
                c1.metric("Loan Outstanding", fmt_rp(d['loan_existing']))
                c2.metric("Collateral Awal",  fmt_rp(d['coll_before_rp']))
                c3.metric("Current Ratio", f"{d['loan_existing']/d['coll_before_rp']*100:.2f}%" if d['coll_before_rp'] > 0 else "N/A")

                original_d  = (st.session_state.get('sid_results_original') or {}).get(sel_sid, d)
                saved_total = st.session_state.get(f'mnc_sim_saved_total_{sel_sid}')

                if d['rp_skipped'] or not d.get('has_rp'):
                    st.info("Loan Existing = 0 → RP tidak diperlukan." if d['rp_skipped'] else "Tidak ada transaksi jual kemarin.")
                    total_rp_sim  = 0.0
                    stocks_ar_sim = dict(original_d.get('stocks_op', d['stocks_op']))
                else:
                    rp_ref_rows = original_d.get('rp_detail', d['rp_detail'])
                    st.caption("Ringkasan saham RP (referensi — lot keluar tetap dihitung penuh)")
                    df_rp_ref = pd.DataFrame([{
                        'Saham': rd['stock'], 'Lot Jual': int(rd['lot_sell']),
                        'Lot OP': int(rd['lot_op']), 'Lot Keluar': int(rd['lot_keluar']),
                        'Harga': rd['price'], 'Nilai Jual (ref)': rd['rp_min'],
                    } for rd in rp_ref_rows])
                    st.dataframe(df_rp_ref, use_container_width=True, hide_index=True)

                    num_key = f"mnc_rp_total_{sel_sid}"
                    if num_key not in st.session_state:
                        st.session_state[num_key] = saved_total if saved_total is not None else float(d['total_rp_maks'])

                    st.caption(f"RP Max (rekomendasi, default): {fmt_rp(d['total_rp_maks'])} "
                               f"| RP Min (ref, 65%): {fmt_rp(d.get('total_rp_min_ratio', 0))}")
                    total_rp_sim = st.number_input(
                        "Total RP Adjust", step=1_000_000.0, format="%.0f", key=num_key,
                    )

                    stocks_ar_sim = dict(original_d.get('stocks_op', d['stocks_op']))
                    for rd in rp_ref_rows:
                        if rd['lot_keluar'] > 0:
                            stocks_ar_sim[rd['stock']] = stocks_ar_sim.get(rd['stock'], 0) - rd['lot_keluar']
                            if stocks_ar_sim.get(rd['stock'], 0) <= 0:
                                stocks_ar_sim.pop(rd['stock'], None)

                coll_ar_sim, _ = calc_collateral(stocks_ar_sim, closing_prices, risk_params)
                loan_ar_sim    = max(original_d.get('loan_existing', d['loan_existing']) - total_rp_sim, 0)
                rasio_rp_sim   = (loan_ar_sim + d['accrued']) / coll_ar_sim if coll_ar_sim > 0 else None

                st.divider()
                r1,r2,r3,r4 = st.columns(4)
                r1.metric("Total RP", fmt_rp(total_rp_sim))
                r2.metric("Loan After RP", fmt_rp(loan_ar_sim))
                r3.metric("Coll After RP", fmt_rp(coll_ar_sim))
                rp_ok = rasio_rp_sim is not None and rasio_rp_sim < RATIO_THRESHOLD
                r4.metric("Rasio RP", f"{rasio_rp_sim*100:.2f}%" if rasio_rp_sim else "N/A",
                    delta="✅ LOLOS" if rp_ok else "❌ GAGAL",
                    delta_color="normal" if rp_ok else "inverse")

                mb_sid = margin_buy.get(sel_sid, {})
                stocks_al_sim = dict(stocks_ar_sim)
                for stock, bdata in mb_sid.items():
                    stocks_al_sim[stock] = stocks_al_sim.get(stock,0) + bdata['lot']
                coll_al_sim, _ = calc_collateral(stocks_al_sim, closing_prices, risk_params)
                total_beli_sim = sum(b['value'] for b in mb_sid.values())

                max63_sim, avail_eff_sim, max_final_sim, binding_sim = solve_lr(
                    coll_al_sim, loan_ar_sim, d['accrued'], d['avail_limit'], total_rp_sim)
                num_lr_sim   = loan_ar_sim + d['accrued'] + min(total_beli_sim, max_final_sim)
                rasio_lr_sim = num_lr_sim / coll_al_sim if coll_al_sim > 0 else None

                st.subheader("Dampak ke LR")
                l1,l2,l3 = st.columns(3)
                l1.metric("Avail Efektif",  fmt_rp(avail_eff_sim))
                l2.metric("LR @63%",        fmt_rp(max63_sim))
                l3.metric("Coll After LR",  fmt_rp(coll_al_sim))
                l4,l5,l6 = st.columns(3)
                l4.metric("Numerator LR",   fmt_rp(num_lr_sim))
                lr_ok = rasio_lr_sim is not None and rasio_lr_sim < RATIO_THRESHOLD
                l5.metric("Rasio LR", f"{rasio_lr_sim*100:.2f}%" if rasio_lr_sim else "N/A",
                    delta="✅ LOLOS" if lr_ok else "❌ Perlu dipotong",
                    delta_color="normal" if lr_ok else "inverse")
                l6.metric("Max LR Final", fmt_rp(max_final_sim), delta=f"Binding: {binding_sim}")

                st.divider()
                b1,b2,b3 = st.columns(3)
                if b1.button("💾 Simpan Simulasi", use_container_width=True, type="primary", key=f'mnc_save_{sel_sid}'):
                    updated = copy.deepcopy(st.session_state['sid_results'][sel_sid])
                    updated['is_simulated']    = True
                    updated['loan_after_rp']   = loan_ar_sim
                    updated['coll_after_rp']   = coll_ar_sim
                    updated['coll_after_lr']   = coll_al_sim
                    updated['ceiling_lr']      = avail_eff_sim
                    updated['avail_efektif']   = avail_eff_sim
                    updated['max_lr_63']       = max63_sim
                    updated['max_lr_65']       = max63_sim
                    updated['max_lr_final']    = max_final_sim
                    updated['lr_binding']      = binding_sim
                    updated['total_rp_maks']   = total_rp_sim
                    updated['stocks_after_rp'] = stocks_ar_sim
                    updated['stocks_after_lr'] = stocks_al_sim
                    new_checks = []
                    for c in updated['checks']:
                        if c['label'] == 'RP-3. Rasio After RP < 65%':
                            new_checks.append({'label': c['label'],
                                'passed': rasio_rp_sim is not None and rasio_rp_sim < RATIO_THRESHOLD,
                                'detail': f"✏️ Simulasi | Rasio: {fmt_pct(rasio_rp_sim)} | Loan: {fmt_rp(loan_ar_sim)} | Coll: {fmt_rp(coll_ar_sim)}"})
                        elif c['label'] == 'LR-3. Nilai Beli ≤ Max LR Final':
                            new_checks.append({'label': c['label'],
                                'passed': total_beli_sim <= max_final_sim,
                                'detail': f"✏️ Simulasi | Nilai Beli: {fmt_rp(total_beli_sim)} vs Max LR Final: {fmt_rp(max_final_sim)}"})
                        else:
                            new_checks.append(c)
                    updated['checks'] = new_checks
                    st.session_state['sid_results'][sel_sid] = updated
                    st.session_state[f'mnc_sim_saved_total_{sel_sid}'] = total_rp_sim
                    st.rerun()
                if b2.button("↩️ Reset Nasabah Ini", use_container_width=True, key=f'mnc_reset1_{sel_sid}'):
                    orig = st.session_state.get('sid_results_original', {})
                    if sel_sid in orig:
                        st.session_state['sid_results'][sel_sid] = copy.deepcopy(orig[sel_sid])
                        st.session_state.pop(f"mnc_rp_total_{sel_sid}", None)
                        st.session_state.pop(f"mnc_sim_saved_total_{sel_sid}", None)
                        st.rerun()
                if b3.button("🔄 Reset Semua", use_container_width=True, key='mnc_reset_all'):
                    orig = st.session_state.get('sid_results_original', {})
                    if orig:
                        st.session_state['sid_results'] = copy.deepcopy(orig); st.rerun()

        with tab_global:
            st.subheader("Validasi Limit Participant")
            if global_result["passed"]:
                st.success(f"✅ LOLOS — {global_result['detail']}")
            else:
                st.error(f"❌ GAGAL — {global_result['detail']}")

        with tab_gagal:
            gc1, gc2 = st.columns(2)
            with gc1:
                gagal_rp = [(s,d) for s,d in sid_results.items()
                            if d.get('has_rp') and not d.get('rp_skipped') and not lolos_rp(d)]
                st.markdown(f"#### 🔴 Gagal RP — {len(gagal_rp)} nasabah")
                if not gagal_rp: st.success("Semua lolos RP.")
                for sid, data in gagal_rp:
                    with st.expander(f"❌ {sid} — {data['name']}"):
                        for c in data['checks']:
                            if c['label'].startswith('RP-') and not c['passed']:
                                st.error(f"**{c['label']}** — {c['detail']}")
            with gc2:
                gagal_lr = [(s,d) for s,d in sid_results.items()
                            if d.get('has_lr') and not lolos_lr(d)]
                st.markdown(f"#### 🔴 Gagal LR — {len(gagal_lr)} nasabah")
                if not gagal_lr: st.success("Semua lolos LR.")
                for sid, data in gagal_lr:
                    with st.expander(f"❌ {sid} — {data['name']}"):
                        if data['max_lr_final'] > 0:
                            st.warning(f"💡 Max LR Final (dapat dicairkan): {fmt_rp(data['max_lr_final'])}")
                        for c in data['checks']:
                            if c['label'].startswith('LR-') and not c['passed']:
                                st.error(f"**{c['label']}** — {c['detail']}")

        with tab_adj:
            st.subheader("⚡ Auto-Adjust LR — Potong ke Max LR Final")
            gagal_lr3 = [(s,d) for s,d in sid_results.items()
                         if d.get('has_lr') and any(
                             c['label'] == 'LR-3. Nilai Beli ≤ Max LR Final' and not c['passed']
                             for c in d['checks'])]
            if not gagal_lr3:
                st.success("✅ Tidak ada nasabah yang perlu di-adjust.")
            else:
                prev_rows = [{'SID': sid, 'Nama': data['name'],
                              'Nilai Beli': fmt_rp(data['total_buy_val']),
                              'Max LR Final': fmt_rp(data['max_lr_final']),
                              'Binding': data.get('lr_binding',''),
                              'Coll After LR': fmt_rp(data['coll_after_lr']),
                              'Status': f"✂️ Dipotong ke {fmt_rp(data['max_lr_final'])}" if data['max_lr_final'] > 0 else "❌ Dikeluarkan"}
                             for sid, data in gagal_lr3]
                st.dataframe(pd.DataFrame(prev_rows), use_container_width=True, hide_index=True)

                if st.button("⚡ Terapkan Auto-Adjust", type="primary", use_container_width=True):
                    new_results = dict(sid_results)
                    for sid, data in gagal_lr3:
                        updated = dict(data)
                        new_checks = []
                        for c in data['checks']:
                            if c['label'] == 'LR-3. Nilai Beli ≤ Max LR Final':
                                numerator  = data['loan_after_rp'] + data['accrued'] + data['max_lr_final']
                                rasio_baru = numerator / data['coll_after_lr'] if data['coll_after_lr'] > 0 else None
                                new_checks.append({'label': c['label'], 'passed': True,
                                    'detail': f"✂️ Di-adjust ke {fmt_rp(data['max_lr_final'])} | Rasio: {fmt_pct(rasio_baru)}"})
                            else:
                                new_checks.append(c)
                        updated['checks'] = new_checks
                        new_results[sid]  = updated
                    total_lr_new = sum(d['max_lr_final'] for d in new_results.values() if lolos_lr(d))
                    total_rp_new = sum(d['total_rp_maks'] for d in new_results.values() if lolos_rp(d))
                    st.session_state['sid_results']   = new_results
                    st.session_state['global_result'] = {
                        "passed": (CREDIT_LIMIT_PARTISIPAN + total_rp_new) > total_lr_new,
                        "detail": (f"CL Partisipan: {fmt_rp(CREDIT_LIMIT_PARTISIPAN)} + "
                                   f"RP: {fmt_rp(total_rp_new)} = "
                                   f"{fmt_rp(CREDIT_LIMIT_PARTISIPAN+total_rp_new)} | LR: {fmt_rp(total_lr_new)}"),
                        "total_rp": total_rp_new, "total_lr": total_lr_new,
                    }
                    st.success("✅ Auto-Adjust diterapkan!"); st.rerun()

        with tab_exp:
            st.subheader("📋 Ringkasan Hasil Validasi")
            sum_rows = []
            for sid, d in sid_results.items():
                sum_rows.append({
                    'SID': sid, 'Nama': d['name'],
                    'Loan Existing': d['loan_existing'], 'Accrued': d['accrued'],
                    'Avail Limit': d['avail_limit'], 'Coll Awal': d['coll_before_rp'],
                    'Current Ratio': f"{d['loan_existing']/d['coll_before_rp']*100:.2f}%" if d['coll_before_rp'] > 0 else "-",
                    'RP Min (ref)': d.get('total_rp_min_ratio', 0),
                    'RP Max (rekomendasi)': d['total_rp_maks'],
                    'Loan After RP': d['loan_after_rp'], 'Coll After RP': d['coll_after_rp'],
                    'Status RP': "✅" if lolos_rp(d) else ("⏭" if d['rp_skipped'] else ("-" if not d['has_rp'] else "❌")),
                    'Avail Efektif': d['avail_efektif'], 'Coll After LR': d['coll_after_lr'],
                    'Max LR Final': d['max_lr_final'], 'Binding': d.get('lr_binding',''),
                    'Status LR': "✅" if lolos_lr(d) else ("-" if not d['has_lr'] else "❌"),
                })
            df_sum = pd.DataFrame(sum_rows)
            st.dataframe(df_sum, use_container_width=True, hide_index=True)

            buf_sum = io.BytesIO()
            df_sum.to_excel(buf_sum, index=False); buf_sum.seek(0)
            st.download_button("⬇️ Download Ringkasan (.xlsx)", data=buf_sum,
                file_name="hasil_validasi.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            st.divider()
            st.subheader("📤 Export ke Sistem")
            today = datetime.today().strftime("%Y%m%d")

            def gen_rp_excel():
                s1, s2 = [], []
                for sid, data in sid_results.items():
                    if not lolos_rp(data): continue
                    if data['total_rp_maks'] <= 0: continue  # RP tidak diperlukan -> tidak ada saham yang keluar
                    s1.append({"Participant Code":"EP","SID Client":sid,"Repayment Value":data['total_rp_maks']})
                    for rd in data['rp_detail']:
                        if rd['lot_keluar'] > 0:
                            s2.append({"SID Client":sid,"Stock Code":rd['stock'],"Quantity":int(rd['lot_keluar'])})
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as w:
                    pd.DataFrame(s1).to_excel(w, sheet_name="Repayment Direct", index=False)
                    pd.DataFrame(s2).to_excel(w, sheet_name="Stock",  index=False)
                buf.seek(0); return buf

            def gen_lr_excel():
                s1, s2 = [], []
                for sid, data in sid_results.items():
                    if not lolos_lr(data): continue
                    if data['max_lr_final'] <= 0: continue  # LR tidak cair -> tidak ada saham yang masuk
                    s1.append({"Participant Code":"EP","SID Client":sid,"Loan Value":data['max_lr_final']})
                    stocks_lr = data.get('stocks_after_lr', {})
                    stocks_rp = data.get('stocks_after_rp', {})
                    for stock, lot_lr in stocks_lr.items():
                        lot_beli = lot_lr - stocks_rp.get(stock, 0)
                        if lot_beli > 0:
                            s2.append({"SID Client":sid,"Stock Code":stock,"Quantity":int(lot_beli)})
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as w:
                    pd.DataFrame(s1).to_excel(w, sheet_name="Loan",      index=False)
                    pd.DataFrame(s2).to_excel(w, sheet_name="Stock", index=False)
                buf.seek(0); return buf

            e1, e2 = st.columns(2)
            with e1:
                st.download_button("⬇️ Repayment Proceed (.xlsx)", data=gen_rp_excel(),
                    file_name=f"Repayment Proceed {today}.xlsx", use_container_width=True,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with e2:
                st.download_button("⬇️ Loan Request (.xlsx)", data=gen_lr_excel(),
                    file_name=f"Loan Request {today}.xlsx", use_container_width=True,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            st.divider()
            st.subheader("📅 Rekap Belum Settled — Untuk Besok")
            def gen_rekap(mode):
                rows = []
                for sid, d in sid_results.items():
                    if mode == 'rp' and lolos_rp(d) and d['total_rp_maks'] > 0:
                        rows.append({"SID":sid,"Name":d["name"],"Repayment Value":d["total_rp_maks"]})
                    elif mode == 'lr' and lolos_lr(d) and d['max_lr_final'] > 0:
                        rows.append({"SID":sid,"Name":d["name"],"Loan Value":d["max_lr_final"]})
                buf = io.BytesIO()
                pd.DataFrame(rows).to_excel(buf, index=False); buf.seek(0)
                return buf

            r1, r2 = st.columns(2)
            with r1:
                st.download_button("⬇️ Rekap RP Belum Settled", data=gen_rekap('rp'),
                    file_name=f"RP Belum Settled {today}.xlsx", use_container_width=True, type="primary",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with r2:
                st.download_button("⬇️ Rekap LR Belum Settled", data=gen_rekap('lr'),
                    file_name=f"LR Belum Settled {today}.xlsx", use_container_width=True, type="primary",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
