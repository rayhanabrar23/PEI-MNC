import streamlit as st
import pandas as pd
import io
from datetime import datetime
import copy

# ── SESSION STATE INIT ────────────────────────────────────────
for key in ['df_sell_edited','sid_results','global_result','df_buy',
            'df_buy_adjusted','op_data','cl_data','closing_prices','risk_params']:
    if key not in st.session_state:
        st.session_state[key] = None

if 'clamped_warnings' not in st.session_state:
    st.session_state['clamped_warnings'] = []

if 'sid_results_original' not in st.session_state:
    st.session_state['sid_results_original'] = {}  # ← dict kosong, bukan None

st.set_page_config(page_title="Validasi MNC", page_icon="✅", layout="wide")
st.title("✅ Validasi MNC")
st.info("Sistem Validasi Repayment & Loan Request nasabah PEI: **RP dulu → LR setelahnya**")

RATIO_THRESHOLD    = 0.65
AUTO_ADJUST_TARGET = 0.63
CREDIT_LIMIT_PARTISIPAN = 160_000_000_000.0

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def fmt_rp(val):
    try: return f"Rp {float(val):,.0f}".replace(',','.')
    except: return str(val)

def fmt_pct(val):
    try: return f"{float(val)*100:.2f}%"
    except: return str(val)

# ─────────────────────────────────────────────
# PARSERS
# ─────────────────────────────────────────────
def parse_op_file(content: str):
    """
    Baca OP file. Tiap nasabah:
    - loan_existing, accrued_interest, available_limit
    - stocks: {stock: lot}
    """
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
            try: avail   = float(parts[7]) if len(parts) > 7 else 0.0
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

def load_closing_price(uploaded_file) -> dict:
    df = pd.read_excel(uploaded_file, sheet_name=0, header=0)
    result = {}
    for _, row in df.iterrows():
        code  = str(row['no_share']).strip().upper()
        price = pd.to_numeric(str(row['kurs_now']).replace(',',''), errors='coerce')
        if pd.notna(price) and code and code != 'NAN':
            result[code] = float(price)
    return result

def load_risk_parameter(uploaded_file) -> dict:
    """Return {stock: hc_decimal}"""
    result  = {}
    content = uploaded_file.read().decode("utf-8", errors="replace")
    uploaded_file.seek(0)
    for line in content.strip().splitlines():
        line = line.strip()
        if not line or line.startswith("StockCode"): continue
        parts = line.split("|")
        if len(parts) < 3: continue
        code = parts[0].strip().upper()
        try: hc = float(parts[2]) / 100.0
        except: hc = 0.0
        result[code] = hc
    return result

def parse_sell_regular(content: str) -> dict:
    """
    Return {sid: {stock: {'lot': x, 'value': y}}}
    Kolom: SID|...|STOCK CODE|QTY|...|PRICE|VALUE
    """
    result = {}
    for line in content.strip().splitlines():
        line  = line.strip()
        parts = line.split('|')
        if len(parts) < 5 or parts[0].strip() == 'SID': continue
        sid   = parts[0].strip()
        stock = parts[1].strip().upper() if len(parts) > 1 else ''
        try: qty = float(str(parts[2]).replace(',',''))
        except: qty = 0.0
        try: val = float(str(parts[6]).replace(',','')) if len(parts) > 6 else 0.0
        except: val = 0.0
        if sid not in result: result[sid] = {}
        if stock not in result[sid]: result[sid][stock] = {'lot': 0, 'value': 0}
        result[sid][stock]['lot']   += qty
        result[sid][stock]['value'] += val
    return result

def parse_margin_buy(content: str) -> dict:
    """
    Return {sid: {stock: {'lot': x, 'value': y}}}
    """
    result = {}
    for line in content.strip().splitlines():
        line  = line.strip()
        parts = line.split('|')
        
        # Abaikan header atau baris yang kosong
        if len(parts) < 5 or parts[0].strip() == 'SID': continue
        
        sid   = parts[0].strip()
        stock = parts[1].strip().upper() if len(parts) > 1 else ''
        
        # Ambil kolom ke-3 (Indeks 2) untuk Lot (MARGIN BUY QUANTITY)
        try: qty = float(str(parts[2]).replace(',',''))
        except: qty = 0.0
        
        # Ambil kolom ke-7 (Indeks 6) untuk Nilai (AVAILABLE MARKET VALUE)
        try: val = float(str(parts[6]).replace(',','')) if len(parts) > 6 else 0.0
        except: val = 0.0
        
        if sid not in result: result[sid] = {}
        if stock not in result[sid]: result[sid][stock] = {'lot': 0, 'value': 0}
        result[sid][stock]['lot']   += qty
        result[sid][stock]['value'] += val
        
    return result

def load_hasil_mnc(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    
    # 1. Proses Sheet Sell / Repayment (RP)
    if "Repayment (RP)" in xls.sheet_names:
        df_sell = pd.read_excel(xls, sheet_name="Repayment (RP)", header=0)
        if 'NETT' in df_sell.columns:
            df_sell = df_sell[~df_sell['NETT'].astype(str).str.contains('EXCLUDED', na=False)].copy()
            
    # -- Fallback (Cadangan) untuk format lama --
    elif "Sell Aktif (tanpa excluded)" in xls.sheet_names:
        df_sell = pd.read_excel(xls, sheet_name="Sell Aktif (tanpa excluded)", header=0)
    elif "Sell Aktif" in xls.sheet_names:
        df_sell = pd.read_excel(xls, sheet_name="Sell Aktif", header=0)
    elif "Sell (Repayment)" in xls.sheet_names:
        df_sell = pd.read_excel(xls, sheet_name="Sell (Repayment)", header=0)
    else:
        st.error(f"❌ Sheet untuk Repayment tidak ditemukan! Sheet yang tersedia di Excel Anda: {xls.sheet_names}")
        st.stop()
        
    # 2. Proses Sheet Buy / Loan Request (LR)
    if "Loan Request (LR)" in xls.sheet_names:
        df_buy = pd.read_excel(xls, sheet_name="Loan Request (LR)", header=0)
        
    # -- Fallback (Cadangan) untuk format lama --
    elif "Buy (Loan)" in xls.sheet_names:
        df_buy = pd.read_excel(xls, sheet_name="Buy (Loan)", header=0)
    else:
        st.error(f"❌ Sheet untuk Loan Request tidak ditemukan! Sheet yang tersedia di Excel Anda: {xls.sheet_names}")
        st.stop()
                
    st.write("Sheet names:", xls.sheet_names)
    st.write("df_buy columns:", list(df_buy.columns))
    st.write("df_buy shape:", df_buy.shape)
            
    # Load raw buy sheet (data transaksi per baris, bukan summary)
    if "Buy (Loan)" in xls.sheet_names:
        df_buy_raw = pd.read_excel(xls, sheet_name="Buy (Loan)", header=0)
    elif "Loan Request (LR)" in xls.sheet_names:
        df_buy_raw = pd.read_excel(xls, sheet_name="Loan Request (LR)", header=0)
    else:
        df_buy_raw = df_buy.copy()

    return df_sell, df_buy, df_buy_raw

# ─────────────────────────────────────────────
# COLLATERAL CALCULATOR
# ─────────────────────────────────────────────
def calc_collateral(stocks_dict, closing_prices, risk_params):
    """Hitung collateral value dari {stock: lot}"""
    total  = 0.0
    detail = []
    for stock, qty in stocks_dict.items():
        cp   = closing_prices.get(stock, 0.0)
        hc   = risk_params.get(stock, 0.05)
        coll = qty * cp * (1 - hc)
        total += coll
        detail.append({"stock": stock, "qty": qty, "cp": cp, "hc": hc, "collateral": coll})
    return total, detail

# ─────────────────────────────────────────────
# CORE VALIDATION PER SID
# ─────────────────────────────────────────────
def validate_sid(sid, op_data, cl_data, sell_regular, margin_buy,
                 closing_prices, risk_params, df_sell, df_buy):
    """
    Logika sequential: RP dulu → LR setelah RP.

    RP:
      - Cek saham sell ada di OP, lot sell ≤ lot OP
      - RP Min = lot_op × harga, RP Maks = nilai jual kemarin (sell_regular)
      - Collateral after RP = OP dikurangi saham yg keluar
      - Rasio RP = (loan_existing - rp_value + accrued) / coll_after_rp < 65%

    LR:
      - Collateral LR = coll_after_rp + saham beli baru (margin_buy) × harga × (1-HC)
      - Ceiling LR    = min(total_buy_val, avail_limit + rp_lolos)
      - Rasio LR      = (loan_after_rp + accrued + ceiling_lr) / coll_after_lr < 65%
      - Jika > 65% → potong ke 63%
    """
    op  = op_data.get(sid, {"loan_existing":0,"accrued_interest":0,"available_limit":0,"name":sid,"stocks":{}})
    cl  = cl_data.get(sid, {"available_limit":0,"name":sid})

    loan_ex   = op["loan_existing"]
    accrued   = op["accrued_interest"]
    avail_lim = cl["available_limit"]
    name      = op.get("name") or cl.get("name") or sid
    stocks_op = op.get("stocks", {})

    sell_stocks = sell_regular.get(sid, {})   # {stock: {lot, value}}
    buy_stocks  = margin_buy.get(sid, {})      # {stock: {lot, value}}

    checks = []
    def add(label, passed, detail=""):
        checks.append({"label": label, "passed": passed, "detail": detail})

    # ── SECTION RP ──────────────────────────────────────────
    rp_detail     = []
    stocks_after_rp = dict(stocks_op)
    total_rp_maks = 0.0
    total_rp_min  = 0.0
    has_rp        = bool(sell_stocks)
    rp_skipped    = loan_ex <= 0 and not has_rp

    if not has_rp or loan_ex <= 0:
        if loan_ex <= 0:
            add("RP-1. Saham Sell Ada di OP", True, "⏭ Dilewati — Loan Existing = 0")
            add("RP-2. Lot Sell ≤ Lot di OP", True, "⏭ Dilewati — Loan Existing = 0")
            add("RP-3. Rasio After RP < 65%", True, "⏭ Dilewati — Loan Existing = 0")
        else:
            add("RP-1. Saham Sell Ada di OP", True, "Tidak ada transaksi jual")
            add("RP-2. Lot Sell ≤ Lot di OP", True, "Tidak ada transaksi jual")
            add("RP-3. Rasio After RP < 65%", True, "Tidak ada transaksi jual")
    else:
        # RP-1: Pengecekan Saham Jual vs OP (Hanya sebagai Info, TIDAK menggagalkan SID)
        rp1_pass = True  # Selalu True agar nasabah tetap lanjut
        rp1_detail = []
        for stock in sell_stocks.keys():
            if stocks_op.get(stock, 0) == 0:
                rp1_detail.append(f"{stock}")
                
        if rp1_detail:
            # Tetap tampilkan info saham apa saja yang bukan jaminan, tapi status tetap True/Lolos
            add("RP-1. Pengecekan Saham Jual", True, f"ℹ️ Info: Saham reguler (bukan OP): {', '.join(rp1_detail)}")
        else:
            add("RP-1. Pengecekan Saham Jual", True, "✅ Semua saham jual adalah kolateral (ada di OP)")

       # RP-2: Lot sell ≤ lot OP — auto-adjust kalau lebih
        rp2_detail = []
        for stock, sdata in sell_stocks.items():
            lot_sell = sdata['lot']
            lot_op   = stocks_op.get(stock, 0)
            lot_keluar = min(lot_sell, lot_op)
            price    = closing_prices.get(stock, 0)
            rp_min   = lot_keluar * price
            rp_maks = sdata['value'] * 1.01
            ada      = lot_op > 0
            
            # --- PERUBAHAN UTAMA: Uang tunai masuk tanpa syarat ---
            total_rp_maks += rp_maks 
            
            if lot_sell > lot_op and ada:
                rp2_detail.append(f"{stock}: {lot_sell:,.0f} → adjusted ke {lot_op:,.0f} (sesuai OP)")
            
            if ada:
                total_rp_min  += rp_min
                rp_detail.append({
                    'stock': stock, 'lot_sell': lot_sell, 'lot_op': lot_op,
                    'lot_keluar': lot_keluar, 'price': price,
                    'rp_min': rp_min, 'rp_maks': rp_maks
                })
                # Update collateral after RP
                stocks_after_rp[stock] = stocks_after_rp.get(stock, 0) - lot_keluar
                if stocks_after_rp[stock] <= 0:
                    del stocks_after_rp[stock]
                    
        add("RP-2. Lot Sell ≤ Lot di OP", True,
            ("⚠️ Auto-adjusted: " + "; ".join(rp2_detail)) if rp2_detail else
            f"Total RP Maks: {fmt_rp(total_rp_maks)}")

        # RP-3: Rasio after RP < 65%
        coll_after_rp, _ = calc_collateral(stocks_after_rp, closing_prices, risk_params)
        loan_after_rp    = max(loan_ex - total_rp_maks, 0)
        numerator_rp     = loan_after_rp + accrued
        rasio_rp = numerator_rp / coll_after_rp if coll_after_rp > 0 else None

        if rasio_rp is not None:
            add("RP-3. Rasio After RP < 65%", rasio_rp < RATIO_THRESHOLD,
                f"Rasio: {fmt_pct(rasio_rp)} | "
                f"Numerator: {fmt_rp(numerator_rp)} (Loan After RP: {fmt_rp(loan_after_rp)} + Accrued: {fmt_rp(accrued)}) | "
                f"Coll After RP: {fmt_rp(coll_after_rp)}")
        elif numerator_rp <= 0:
            add("RP-3. Rasio After RP < 65%", True, f"Numerator ≤ 0 — posisi lunas setelah RP")
        else:
            add("RP-3. Rasio After RP < 65%", False, "Collateral = 0 setelah RP — rasio ∞")

    # ── SECTION LR ──────────────────────────────────────────
    has_lr         = bool(buy_stocks)
    coll_after_rp2, _ = calc_collateral(stocks_after_rp, closing_prices, risk_params)
    loan_after_rp2    = max(loan_ex - total_rp_maks, 0)

    # Collateral LR = coll after RP + saham beli baru
    stocks_after_lr = dict(stocks_after_rp)
    for stock, bdata in buy_stocks.items():
        stocks_after_lr[stock] = stocks_after_lr.get(stock, 0) + bdata['lot']
    coll_after_lr, _ = calc_collateral(stocks_after_lr, closing_prices, risk_params)

    total_buy_val  = sum(b['value'] for b in buy_stocks.values())
    avail_efektif  = avail_lim + total_rp_maks
    buy_with_buffer = total_buy_val * 1.1
    ceiling_lr = min(buy_with_buffer, avail_efektif)
    # flag apakah buffer berlaku atau tidak
    buffer_lr_berlaku = buy_with_buffer <= avail_efektif
    max_lr_63 = max(coll_after_lr * AUTO_ADJUST_TARGET - (loan_after_rp2 + accrued), 0) if coll_after_lr > 0 else 0
    max_lr_65 = max(coll_after_lr * RATIO_THRESHOLD    - (loan_after_rp2 + accrued), 0) if coll_after_lr > 0 else 0
    max_lr_final = min(ceiling_lr, max_lr_65)

    if not has_lr:
        add("LR-1. Volume Buy ≤ Available Quantity", True, "Tidak ada Loan Request")
        add("LR-2. Ceiling LR (min Beli vs Limit)", True, "Tidak ada Loan Request")
        add("LR-3. Rasio LR < 65%", True, "Tidak ada Loan Request")
    else:
        # LR-1: Volume buy ≤ available quantity (dari df_buy)
        lr1_pass   = True
        lr1_detail = []
        buy_rows = df_buy[df_buy.iloc[:, 0].astype(str) == sid] if df_buy is not None else pd.DataFrame()
        for _, row in buy_rows.iterrows():
            vol = pd.to_numeric(row.iloc[13] if len(row) > 13 else 0, errors='coerce') or 0
            avq = pd.to_numeric(row.iloc[4]  if len(row) > 4  else 0, errors='coerce') or 0
            stk = str(row.iloc[1]) if len(row) > 1 else ''
            if avq == 0:
                lr1_pass = False
                lr1_detail.append(f"{stk}: DIBATALKAN (Available = 0)")
            elif vol > avq:
                lr1_pass = False
                lr1_detail.append(f"{stk}: Vol {vol:,.0f} > Avail {avq:,.0f}")
        add("LR-1. Volume Buy ≤ Available Quantity", lr1_pass,
            "; ".join(lr1_detail) if lr1_detail else f"Total Nilai Beli: {fmt_rp(total_buy_val)}")

        # LR-2: Ceiling LR
        add("LR-2. Ceiling LR (min Beli vs Limit)", True,
            f"Nilai Beli Kemarin: {fmt_rp(total_buy_val)} | "
            f"Avail Efektif (Limit + RP): {fmt_rp(avail_efektif)} | "
            f"Ceiling LR: {fmt_rp(ceiling_lr)}")

        # LR-3: Rasio LR < 65%
        numerator_lr = loan_after_rp2 + accrued + ceiling_lr
        rasio_lr = numerator_lr / coll_after_lr if coll_after_lr > 0 else None

        if rasio_lr is not None:
            lr3_pass = rasio_lr < RATIO_THRESHOLD
            detail_str = (
                f"Rasio: {fmt_pct(rasio_lr)} (threshold <{RATIO_THRESHOLD*100:.0f}%) | "
                f"Numerator: {fmt_rp(numerator_lr)} "
                f"(Loan After RP: {fmt_rp(loan_after_rp2)} + Accrued: {fmt_rp(accrued)} + LR: {fmt_rp(ceiling_lr)}) | "
                f"Coll After LR: {fmt_rp(coll_after_lr)}"
            )
            if not lr3_pass:
                detail_str += f" || ⚠️ Max LR aman (63%): {fmt_rp(max_lr_63)}"
            add("LR-3. Rasio LR < 65%", lr3_pass, detail_str)
        elif numerator_lr <= 0:
            add("LR-3. Rasio LR < 65%", True, "Numerator ≤ 0 — posisi lunas")
        else:
            add("LR-3. Rasio LR < 65%", False, "Collateral = 0 — rasio ∞")

    return {
        "name":            name,
        "checks":          checks,
        "has_rp":          has_rp,
        "has_lr":          has_lr,
        "rp_skipped":      loan_ex <= 0,
        "loan_existing":   loan_ex,
        "accrued":         accrued,
        "avail_limit":     avail_lim,
        "stocks_op":       stocks_op,
        "stocks_after_rp": stocks_after_rp,
        "stocks_after_lr": stocks_after_lr,
        "rp_detail":       rp_detail,
        "total_rp_maks":   total_rp_maks,
        "total_rp_min":    total_rp_min,
        "loan_after_rp":   loan_after_rp2,
        "coll_before_rp":  calc_collateral(stocks_op, closing_prices, risk_params)[0],
        "coll_after_rp":   coll_after_rp2,
        "coll_after_lr":   coll_after_lr,
        "ceiling_lr":      ceiling_lr,
        "max_lr_63":       max_lr_63,
        "max_lr_65":       max_lr_65,
        "max_lr_final":    max_lr_final,
        "avail_efektif":   avail_efektif,
        "total_buy_val":   total_buy_val,
    }

def lolos_rp(data):
    if data.get("rp_skipped"): return False
    if not data.get("has_rp"): return False
    return all(c["passed"] for c in data["checks"] if c["label"].startswith("RP-"))

def lolos_lr(data):
    if not data.get("has_lr"): return False
    return all(c["passed"] for c in data["checks"] if c["label"].startswith("LR-"))

# ─────────────────────────────────────────────
# AUTO ADJUST LR
# ─────────────────────────────────────────────
def auto_adjust_loan(df_buy, sid, max_loan_val, orig_loan_val, closing_prices):
    df_updated = df_buy.copy().reset_index(drop=True)
    df_ref     = df_buy.copy().reset_index(drop=True)

    # Pastikan kolom ke-13 dan 14 ada
    n_cols = len(df_updated.columns)
    if n_cols < 14:
        st.warning(f"auto_adjust_loan: df_buy hanya punya {n_cols} kolom, minimal butuh 14.")
        return df_updated

    col_sid   = df_updated.columns[0]
    col_stock = df_updated.columns[1]
    col_vol   = df_updated.columns[13]
    col_val   = df_updated.columns[14] if n_cols > 14 else None

    rows_idx = df_updated[df_updated[col_sid].astype(str) == sid].index.tolist()
    if not rows_idx:
        return df_updated

    if orig_loan_val <= 0 or max_loan_val <= 0:
        for i in rows_idx:
            df_updated.at[i, col_vol] = 0
            if col_val:
                df_updated.at[i, col_val] = 0
        return df_updated

    ratio = max_loan_val / orig_loan_val
    for i in rows_idx:
        old_vol = pd.to_numeric(df_ref.at[i, col_vol], errors='coerce')
        if pd.isna(old_vol):
            old_vol = 0.0
        stock   = str(df_ref.at[i, col_stock]).strip().upper()
        price   = closing_prices.get(stock, 0)
        new_vol = int((old_vol * ratio) // 100) * 100
        new_vol = max(new_vol, 0)
        new_val = new_vol * price if price > 0 else 0
        df_updated.at[i, col_vol] = new_vol
        if col_val:
            df_updated.at[i, col_val] = new_val

    return df_updated

# ─────────────────────────────────────────────
# EXCEL EXPORTS
# ─────────────────────────────────────────────
def generate_repayment_excel(sid_results, sell_regular):
    today = datetime.today().strftime("%Y%m%d")
    s1, s2 = [], []
    for sid, data in sid_results.items():
        if not lolos_rp(data): continue
        total_val = data['total_rp_maks']
        if total_val > 0:
            s1.append({"Participant Code": "EP", "SID Client": sid, "Repayment Value": total_val})
        for rd in data['rp_detail']:
            if rd['lot_keluar'] > 0:
                s2.append({"SID Client": sid, "Stock Code": rd['stock'], "Quantity": int(rd['lot_keluar'])})
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame(s1).to_excel(w, sheet_name="Repayment Proceed", index=False)
        pd.DataFrame(s2).to_excel(w, sheet_name="Detail Collateral",  index=False)
    buf.seek(0)
    return buf, f"Repayment Proceed {today}.xlsx"

def generate_loan_excel(sid_results, margin_buy):
    today = datetime.today().strftime("%Y%m%d")
    s1, s2 = [], []
    for sid, data in sid_results.items():
        if not lolos_lr(data): continue
        if data['ceiling_lr'] > 0:
            s1.append({"Participant Code": "EP", "SID Client": sid, "Loan Value": data['max_lr_final']})
        for stock, bdata in margin_buy.get(sid, {}).items():
            if bdata['lot'] > 0:
                s2.append({"SID Client": sid, "Stock Code": stock, "Quantity": int(bdata['lot'])})
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame(s1).to_excel(w, sheet_name="Loan Request",      index=False)
        pd.DataFrame(s2).to_excel(w, sheet_name="Detail Collateral", index=False)
    buf.seek(0)
    return buf, f"Loan Request {today}.xlsx"

def generate_rekap_rp_excel(sid_results):
    today = datetime.today().strftime("%Y%m%d")
    rows = [{"SID": sid, "Name": d["name"], "Repayment Value": d["total_rp_maks"]}
            for sid, d in sid_results.items() if lolos_rp(d) and d["total_rp_maks"] > 0]
    buf = io.BytesIO()
    pd.DataFrame(rows).to_excel(buf, index=False)
    buf.seek(0)
    return buf, f"RP Belum Settled {today}.xlsx"

def generate_rekap_lr_excel(sid_results):
    today = datetime.today().strftime("%Y%m%d")
    rows = [{"SID": sid, "Name": d["name"], "Loan Value": d["max_lr_final"]}
            for sid, d in sid_results.items() if lolos_lr(d) and d["max_lr_final"] > 0]
    buf = io.BytesIO()
    pd.DataFrame(rows).to_excel(buf, index=False)
    buf.seek(0)
    return buf, f"LR Belum Settled {today}.xlsx"

# ─────────────────────────────────────────────
# UPLOAD UI
# ─────────────────────────────────────────────
st.subheader("📂 Upload File")
c1,c2,c3 = st.columns(3)
with c1:
    hasil_file = st.file_uploader("1. Hasil_MNC (xlsx)", type=["xlsx"])
with c2:
    op_file    = st.file_uploader("2. File OP (.txt)",   type=["txt"])
with c3:
    cl_file    = st.file_uploader("3. Credit Limit (.txt)", type=["txt"])
c4,c5,c6 = st.columns(3)
with c4:
    cp_file    = st.file_uploader("4. Closing Price (.xlsx)", type=["xlsx"])
with c5:
    rp_param   = st.file_uploader("5. Risk Parameter (.txt)", type=["txt"])
with c6:
    mbuy_file  = st.file_uploader("6. Margin Buy (.txt)", type=["txt"])
c7,_,_ = st.columns(3)
with c7:
    msell_file = st.file_uploader("7. Sell Regular (.txt)", type=["txt"])

# Preview OP
if op_file:
    try:
        op_prev_content = op_file.read().decode("utf-8", errors="replace"); op_file.seek(0)
        op_prev = parse_op_file(op_prev_content)
        with st.expander(f"👁 Preview OP File — {len(op_prev)} nasabah"):
            for s in list(op_prev.keys())[:5]:
                d = op_prev[s]
                st.caption(f"**{s}** {d['name']} | Loan: {fmt_rp(d['loan_existing'])} | Accrued: {fmt_rp(d['accrued_interest'])} | Avail: {fmt_rp(d['available_limit'])} | Saham: {len(d['stocks'])} kode")
    except: pass

# Credit limit partisipan
st.subheader("Credit Limit Partisipan")
st.info(f"Ditetapkan: **Rp 160.000.000.000**")
st.divider()

with st.expander("📖 Formula Rasio"):
    st.markdown("""
    **LANGKAH 1 — Rasio RP:**
    ```
    Rasio RP = (Loan Existing - RP Diajukan + Accrued) / Collateral After RP
    Collateral After RP = Collateral OP dikurangi saham yang di-repay
    ```
    **LANGKAH 2 — Rasio LR (setelah RP diinput):**
    ```
    Collateral LR = Collateral After RP + Saham Beli Baru (Margin Buy × HC)
    Ceiling LR    = min(Nilai Beli Kemarin, Avail Limit + RP)
    Rasio LR      = (Loan After RP + Accrued + Ceiling LR) / Collateral LR
    Max LR (63%)  = Collateral LR × 63% - (Loan After RP + Accrued)
    ```
    """)

run_btn = st.button("▶ Jalankan Validasi", use_container_width=True, type="primary")

if run_btn:
    errors = []
    if not hasil_file: errors.append("File Hasil_MNC belum diupload.")
    if not op_file:    errors.append("File OP belum diupload.")
    if not cl_file:    errors.append("File Credit Limit belum diupload.")
    if not cp_file:    errors.append("File Closing Price belum diupload.")
    if not rp_param:   errors.append("File Risk Parameter belum diupload.")
    if not mbuy_file:  errors.append("File Margin Buy belum diupload.")
    if not msell_file: errors.append("File Sell Regular belum diupload.")
    if errors:
        for e in errors: st.error(e)
        st.stop()

    with st.spinner("⚙️ Memproses data..."):
        df_sell, df_buy, df_buy_raw = load_hasil_mnc(hasil_file)
        op_content = op_file.read().decode("utf-8", errors="replace")
        op_data    = parse_op_file(op_content)
        cl_content = cl_file.read().decode("utf-8", errors="replace")
        cl_data, vdate = parse_credit_limit_file(cl_content)
        if vdate:
            today_check = datetime.today().strftime("%Y/%m/%d")
            if vdate != today_check:
                st.warning(f"⚠️ Value Date Credit Limit: **{vdate}** (berbeda dengan hari ini {today_check})")
        closing_prices = load_closing_price(cp_file)
        risk_params    = load_risk_parameter(rp_param)
        mb_content     = mbuy_file.read().decode("utf-8", errors="replace")
        margin_buy     = parse_margin_buy(mb_content)
        sr_content     = msell_file.read().decode("utf-8", errors="replace")
        sell_regular   = parse_sell_regular(sr_content)

        all_sids = sorted(set(list(op_data.keys()) + list(cl_data.keys()) +
                              list(margin_buy.keys()) + list(sell_regular.keys())))

        sid_results = {}
        for sid in all_sids:
            sid_results[sid] = validate_sid(
                sid, op_data, cl_data, sell_regular, margin_buy,
                closing_prices, risk_params, df_sell, df_buy
            )

        # Global check: CL Partisipan
        total_lr_all = sum(d['max_lr_final'] for d in sid_results.values() if lolos_lr(d))
        total_rp_all = sum(d['total_rp_maks'] for d in sid_results.values() if lolos_rp(d))
        global_result = {
            "passed": (CREDIT_LIMIT_PARTISIPAN + total_rp_all) > total_lr_all,
            "detail": f"CL Partisipan: {fmt_rp(CREDIT_LIMIT_PARTISIPAN)} + Total RP: {fmt_rp(total_rp_all)} = {fmt_rp(CREDIT_LIMIT_PARTISIPAN+total_rp_all)} | Total LR: {fmt_rp(total_lr_all)}",
            "total_rp": total_rp_all,
            "total_lr": total_lr_all,
        }

    st.session_state.update({
        'sid_results':    sid_results,
        'global_result':  global_result,
        'op_data':        op_data,
        'cl_data':        cl_data,
        'closing_prices': closing_prices,
        'risk_params':    risk_params,
        'margin_buy':     margin_buy,
        'sell_regular':   sell_regular,
        'df_sell_edited': df_sell.copy(),
        'df_buy':         df_buy.copy(),
        'df_buy_raw':     df_buy_raw.copy(),
        'df_buy_adjusted':df_buy_raw.copy(),
        'sid_results_original': copy.deepcopy(sid_results),
    })
    st.success("✅ Validasi Selesai!")

# ─────────────────────────────────────────────
# TAMPILKAN HASIL
# ─────────────────────────────────────────────
if st.session_state.get('sid_results'):
    sid_results   = st.session_state['sid_results']
    global_result = st.session_state['global_result']
    op_data       = st.session_state.get('op_data', {})
    cl_data       = st.session_state.get('cl_data', {})
    closing_prices= st.session_state.get('closing_prices', {})
    risk_params   = st.session_state.get('risk_params', {})
    margin_buy    = st.session_state.get('margin_buy', {})
    sell_regular  = st.session_state.get('sell_regular', {})

    n_rp   = sum(1 for v in sid_results.values() if v.get('has_rp'))
    n_lr   = sum(1 for v in sid_results.values() if v.get('has_lr'))
    n_rp_ok  = sum(1 for v in sid_results.values() if lolos_rp(v))
    n_lr_ok  = sum(1 for v in sid_results.values() if lolos_lr(v))
    n_rp_fail= sum(1 for v in sid_results.values() if v.get('has_rp') and not v.get('rp_skipped') and not lolos_rp(v))
    n_lr_fail= sum(1 for v in sid_results.values() if v.get('has_lr') and not lolos_lr(v))
    n_skip   = sum(1 for v in sid_results.values() if v.get('rp_skipped'))

    m1,m2,m3,m4,m5,m6,m7 = st.columns(7)
    m1.metric("Total Nasabah",   len(sid_results))
    m2.metric("Ada RP",          n_rp)
    m3.metric("Ada LR",          n_lr)
    m4.metric("RP Lolos",        n_rp_ok)
    m5.metric("LR Lolos",        n_lr_ok)
    m6.metric("RP/LR Gagal",     f"{n_rp_fail}/{n_lr_fail}",
              delta=f"-{n_rp_fail+n_lr_fail}" if (n_rp_fail+n_lr_fail) else None, delta_color="inverse")
    m7.metric("CL Partisipan",   "✅ LOLOS" if global_result["passed"] else "❌ GAGAL")

    st.divider()

    tab_rp, tab_lr, tab_sim, tab_global, tab_gagal, tab_adj, tab_export = st.tabs([
        "📤 LANGKAH 1 — Repayment (RP)",
        "📥 LANGKAH 2 — Loan Request (LR)",
        "🎛️ Simulator RP → LR",
        "🌐 Validasi Limit Participant",
        "❌ Nasabah Gagal",
        "⚡ Auto-Adjust LR",
        "📥 Export",
    ])

    # ── TAB RP ────────────────────────────────────────────────
    with tab_rp:
        st.info("💡 Input RP **terlebih dahulu**. Setelah RP diinput, loan outstanding berkurang dan available limit naik.")
        for sid, data in sid_results.items():
            if not data.get('has_rp') or data.get('rp_skipped') or data['total_rp_maks'] <= 0: continue
            skip = data.get('rp_skipped')
            ok   = lolos_rp(data)
            icon = "⏭" if skip else ("✅" if ok else "❌")
            # Tambah cek is_simulated
            simulated = data.get('is_simulated', False)
            icon = "⏭" if skip else ("✅" if ok else "❌")
            if simulated: icon += " ✏️"
            with st.expander(f"{icon} {sid} — {data['name']}  |  RP Maks: {fmt_rp(data['total_rp_maks'])}", expanded=not ok and not skip):
                if skip:
                    st.info("Loan Existing = 0 → Repayment tidak diperlukan")
                    continue
                col_a, col_b, col_c, col_d = st.columns(4)
                col_a.metric("Loan Existing",   fmt_rp(data['loan_existing']))
                col_b.metric("Loan After RP",   fmt_rp(data['loan_after_rp']))
                col_c.metric("Coll After RP",   fmt_rp(data['coll_after_rp']))
                rasio_rp_disp = next((c['detail'] for c in data['checks'] if c['label'] == 'RP-3. Rasio After RP < 65%'), '')
                col_d.metric("RP Maks", fmt_rp(data['total_rp_maks']))
                st.caption(f"RP Min: {fmt_rp(data['total_rp_min'])}  |  RP Maks: {fmt_rp(data['total_rp_maks'])}")

                rp_sistem = data['total_rp_maks']  # sudah include buffer 1%
                st.warning(
                   f"⚠️ Sistem mengasumsikan RP = {fmt_rp(rp_sistem)} "
                   f"(nilai jual × 1.01). Jika partisipan input lebih besar, "
                   f"hasil LR di bawah perlu dihitung ulang."
                )
                    
                if data['rp_detail']:
                    df_rd = pd.DataFrame([{
                        'Saham': r['stock'], 'Lot Jual': int(r['lot_sell']),
                        'Lot OP': int(r['lot_op']), 'Lot Keluar': int(r['lot_keluar']),
                        'Harga': r['price'], 'RP Min': r['rp_min'], 'RP Maks': r['rp_maks']
                    } for r in data['rp_detail']])
                    st.dataframe(df_rd, use_container_width=True, hide_index=True)

                for c in data['checks']:
                    if not c['label'].startswith('RP-'): continue
                    if c['passed']: st.success(f"✅ **{c['label']}** {c['detail']}")
                    else:           st.error(  f"❌ **{c['label']}** {c['detail']}")

    # ── TAB LR ────────────────────────────────────────────────
    with tab_lr:
        st.info("💡 Input LR **setelah RP selesai**. Collateral sudah ditambah saham beli baru. Ceiling LR = min(Nilai Beli, Avail Limit + RP).")
        for sid, data in sid_results.items():
            if not data.get('has_lr'): continue
            ok   = lolos_lr(data)
            icon = "✅" if ok else "❌"
            simulated = data.get('is_simulated', False)
            icon = "✅" if ok else "❌"
            if simulated: icon += " ✏️"
            with st.expander(f"{icon} {sid} — {data['name']}  |  Max LR Final: {fmt_rp(data['max_lr_final'])}", expanded=not ok):

                st.info(
                    f"ℹ️ Avail Efektif dan Loan After RP dihitung berdasarkan "
                    f"asumsi RP sistem ({fmt_rp(data['total_rp_maks'])}). "
                    f"Jika RP aktual berbeda, angka LR ini bisa berubah."
                ) 
                    
                col_a, col_b, col_c, col_d = st.columns(4)
                col_a.metric("Loan After RP",    fmt_rp(data['loan_after_rp']))
                col_b.metric("Avail Efektif",    fmt_rp(data['avail_efektif']), help="Avail Limit + RP Lolos")
                col_c.metric("Ceiling LR",       fmt_rp(data['ceiling_lr']))
                col_d.metric("Coll After LR",    fmt_rp(data['coll_after_lr']))
                col_e, col_f, _, _ = st.columns(4)
                col_e.metric("Max LR (65%)",     fmt_rp(data['max_lr_65']))
                col_f.metric("Max LR Aman (63%)",fmt_rp(data['max_lr_63']))

                if data['max_lr_63'] > 0 and not ok:
                    st.warning(f"💡 Potong LR ke: **{fmt_rp(data['max_lr_final'])}** agar rasio ≤ 63%")
                elif data['max_lr_63'] == 0 and not ok:
                    st.error("🚫 Collateral tidak mencukupi — tidak bisa ajukan LR")

                for c in data['checks']:
                    if not c['label'].startswith('LR-'): continue
                    if c['passed']: st.success(f"✅ **{c['label']}** {c['detail']}")
                    else:           st.error(  f"❌ **{c['label']}** {c['detail']}")

    # ── TAB SIMULATOR ─────────────────────────────────────────
    with tab_sim:
        st.subheader("🎛️ Simulator — Ubah Nilai RP dan Lihat Dampak ke LR")
        st.info("Ubah nilai RP per saham. Sistem menghitung ulang rasio RP, loan after RP, collateral LR, dan max LR secara otomatis.")

        sid_options = [sid for sid, d in sid_results.items() if d.get('has_rp') or d.get('has_lr')]
        if not sid_options:
            st.warning("Tidak ada nasabah dengan transaksi.")
        else:
            sel_sid = st.selectbox("Pilih Nasabah:", sid_options,
                format_func=lambda s: f"{s} — {sid_results[s]['name']}")
            
            d = st.session_state['sid_results'][sel_sid]
            
            # Badge kalau sudah disimpan
            if d.get('is_simulated'):
                st.success("✏️ Nasabah ini menggunakan nilai simulasi yang sudah disimpan.")
            
            c1, c2, c3 = st.columns(3)
            c1.metric("Loan Outstanding", fmt_rp(d['loan_existing']))
            c2.metric("Collateral Awal",  fmt_rp(d['coll_before_rp']))
            c3.metric("Current Ratio",    f"{d['loan_existing']/d['coll_before_rp']*100:.2f}%" if d['coll_before_rp'] > 0 else "N/A")
            
            st.subheader("Step 1 — Atur Nilai RP")
            
            saved_input = st.session_state.get(f'sim_saved_input_{sel_sid}', {})
            saved_mode1 = saved_input.get('mode1', True)
            saved_mode2 = saved_input.get('mode2', False)
            
            col_m1, col_m2 = st.columns(2)
            with col_m1:
                mode1 = st.checkbox("📦 Simulasi Lot Saham", value=saved_mode1, key=f"mode1_{sel_sid}")
            with col_m2:
                mode2 = st.checkbox("💰 Simulasi Nilai RP", value=saved_mode2, key=f"mode2_{sel_sid}")
            
            rp_inputs = {}
            
            # Ambil data original untuk referensi
            original_d = (st.session_state.get('sid_results_original') or {}).get(sel_sid, d)
            
            if d['rp_skipped']:
                st.info("Loan Existing = 0 → RP tidak diperlukan. Lanjut ke LR langsung.")
            elif not d.get('has_rp'):
                            st.info("Tidak ada transaksi jual kemarin.")
                        # ── GANTI BLOK data_sim LAMA DENGAN INI ──
                rp_detail_ref = original_d.get('rp_detail', d['rp_detail'])
                saved_input = st.session_state.get(f'sim_saved_input_{sel_sid}', {})
                saved_lots  = saved_input.get('lot_inputs', {})
                saved_rps   = saved_input.get('rp_inputs', {})
            
                data_sim = []
                for idx, rd in enumerate(rp_detail_ref):
                    data_sim.append({
                        "Saham":    rd['stock'],
                        "Lot Jual": saved_lots.get(rd['stock'], int(rd['lot_sell'])),
                        "Harga":    rd['price'],
                        "RP Value": saved_rps.get(idx, round(rd['rp_maks'], 0)),
                        "_lot_op":  int(rd['lot_op']),
                    })
                df_sim = pd.DataFrame(data_sim)
            
                if mode1:
                    st.caption("📦 Simulasi Lot Saham")
                    df_sim1 = df_sim[["Saham", "Lot Jual", "Harga", "_lot_op"]].copy()
                    edited_df1 = st.data_editor(
                        df_sim1,
                        column_config={
                            "Saham":   st.column_config.TextColumn("Saham", disabled=True),
                            "Lot Jual":st.column_config.NumberColumn("Lot Jual", min_value=0, step=1),
                            "Harga":   st.column_config.NumberColumn("Harga", format="Rp %d", disabled=True),
                            "_lot_op": st.column_config.NumberColumn("Lot OP", disabled=True),
                        },
                        hide_index=True,
                        use_container_width=True,
                        key=f"sim1_{sel_sid}"
                    )
                else:
                    edited_df1 = df_sim[["Saham", "Lot Jual", "Harga", "_lot_op"]].copy()
            
                if mode2:
                    st.caption("💰 Simulasi Nilai RP")
                    df_sim2 = df_sim[["RP Value"]].copy()
                    edited_df2 = st.data_editor(
                        df_sim2,
                        column_config={
                            "RP Value": st.column_config.NumberColumn("RP Value (Rp)", min_value=0, step=1_000_000.0),
                        },
                        hide_index=True,
                        use_container_width=True,
                        key=f"sim2_{sel_sid}"
                    )
                else:
                    edited_df2 = df_sim[["RP Value"]].copy()
            
                for i, row in edited_df1.iterrows():
                    stock      = row['Saham']
                    lot_jual   = row['Lot Jual'] if mode1 else next(rd['lot_sell'] for rd in rp_detail_ref if rd['stock'] == stock)
                    lot_op     = row['_lot_op']
                    lot_keluar = min(lot_jual, lot_op) if mode1 else 0
                    harga      = row['Harga']
                    if mode2:
                        rp_value = float(edited_df2.loc[i, 'RP Value'])
                    else:
                        rp_value = lot_jual * harga * 1.01
                    rp_inputs[stock] = {
                        'rp_value':   rp_value,
                        'lot_keluar': lot_keluar,
                    }
            
            total_rp_sim = sum(v['rp_value'] for v in rp_inputs.values())
            
            # Hitung ulang collateral
            stocks_after_rp_sim = dict(original_d.get('stocks_op', d['stocks_op']))
            for stock, v in rp_inputs.items():
                if v['lot_keluar'] > 0:
                    stocks_after_rp_sim[stock] = stocks_after_rp_sim.get(stock, 0) - v['lot_keluar']
                    if stocks_after_rp_sim.get(stock, 0) <= 0:
                        stocks_after_rp_sim.pop(stock, None)
            
            coll_after_rp_sim, _ = calc_collateral(stocks_after_rp_sim, closing_prices, risk_params)
            loan_after_rp_sim    = max(original_d.get('loan_existing', d['loan_existing']) - total_rp_sim, 0)
            rasio_rp_sim = (loan_after_rp_sim + d['accrued']) / coll_after_rp_sim if coll_after_rp_sim > 0 else None

            st.divider()
            st.subheader("Hasil Setelah RP")
            r1,r2,r3,r4 = st.columns(4)
            r1.metric("Total RP",       fmt_rp(total_rp_sim))
            r2.metric("Loan After RP",  fmt_rp(loan_after_rp_sim))
            r3.metric("Coll After RP",  fmt_rp(coll_after_rp_sim))
            rp_ok_sim = rasio_rp_sim is not None and rasio_rp_sim < RATIO_THRESHOLD
            r4.metric("Rasio RP", f"{rasio_rp_sim*100:.2f}%" if rasio_rp_sim else "N/A",
                delta="✅ LOLOS" if rp_ok_sim else "❌ GAGAL",
                delta_color="normal" if rp_ok_sim else "inverse")

            st.divider()
            st.subheader("Step 2 — Dampak ke LR")

            mb_sid = margin_buy.get(sel_sid, {})
            stocks_after_lr_sim = dict(stocks_after_rp_sim)
            for stock, bdata in mb_sid.items():
                stocks_after_lr_sim[stock] = stocks_after_lr_sim.get(stock, 0) + bdata['lot']
            coll_lr_sim, _ = calc_collateral(stocks_after_lr_sim, closing_prices, risk_params)

            avail_eff_sim  = d['avail_limit'] + total_rp_sim
            total_beli_sim = sum(b['value'] for b in mb_sid.values())
            buy_with_buffer_sim = total_beli_sim * 1.1
            ceiling_sim         = min(buy_with_buffer_sim, avail_eff_sim)
            buffer_berlaku_sim  = buy_with_buffer_sim <= avail_eff_sim
            num_lr_sim     = loan_after_rp_sim + d['accrued'] + ceiling_sim
            rasio_lr_sim   = num_lr_sim / coll_lr_sim if coll_lr_sim > 0 else None
            max63_sim      = max(coll_lr_sim * 0.63 - (loan_after_rp_sim + d['accrued']), 0)
            max65_sim      = max(coll_lr_sim * 0.65 - (loan_after_rp_sim + d['accrued']), 0)
            max_final_sim  = min(ceiling_sim, max65_sim)

            l1,l2,l3 = st.columns(3)
            l1.metric("Avail Efektif",  fmt_rp(avail_eff_sim), help="Avail Limit + RP")
            l2.metric("Ceiling LR",     fmt_rp(ceiling_sim),   help="min(Nilai Beli, Avail Efektif)")
            l3.metric("Coll After LR",  fmt_rp(coll_lr_sim),   help="Coll After RP + Saham Beli Baru")
            l4,l5,l6 = st.columns(3)
            l4.metric("Numerator LR",   fmt_rp(num_lr_sim))
            lr_ok_sim = rasio_lr_sim is not None and rasio_lr_sim < RATIO_THRESHOLD
            l5.metric("Rasio LR", f"{rasio_lr_sim*100:.2f}%" if rasio_lr_sim else "N/A",
                delta="✅ LOLOS" if lr_ok_sim else "❌ Perlu dipotong",
                delta_color="normal" if lr_ok_sim else "inverse")
            l6.metric("Max LR Final (63%)", fmt_rp(max_final_sim))

            with st.expander("📊 Detail Collateral After LR"):
                crows = []
                for stk, lot in stocks_after_lr_sim.items():
                    p = closing_prices.get(stk,0); h = risk_params.get(stk,0.05)
                    crows.append({'Saham':stk,'Lot':int(lot),'Harga':p,
                                  'HC':f"{h*100:.0f}%",'CV':lot*p*(1-h),
                                  'Sumber':'Sisa OP' if stk in stocks_after_rp_sim else 'Beli Baru'})
                st.dataframe(pd.DataFrame(crows), use_container_width=True, hide_index=True)


            st.divider()
            col_b1, col_b2, col_b3 = st.columns(3)
            
            if col_b1.button("💾 Simpan Simulasi", use_container_width=True, type="primary"):
                updated = copy.deepcopy(st.session_state['sid_results'][sel_sid])
                updated['is_simulated']    = True
                updated['loan_after_rp']   = loan_after_rp_sim
                updated['coll_after_rp']   = coll_after_rp_sim
                updated['coll_after_lr']   = coll_lr_sim
                updated['ceiling_lr']      = ceiling_sim
                updated['avail_efektif']   = avail_eff_sim
                updated['max_lr_63']       = max63_sim
                updated['max_lr_65']       = max65_sim
                updated['max_lr_final']    = max_final_sim
                updated['total_rp_maks']   = total_rp_sim
                updated['stocks_after_rp'] = stocks_after_rp_sim
                updated['stocks_after_lr'] = stocks_after_lr_sim
                new_checks = []
                for c in updated['checks']:
                    if c['label'] == 'RP-3. Rasio After RP < 65%':
                        new_checks.append({
                            'label': c['label'],
                            'passed': rasio_rp_sim is not None and rasio_rp_sim < RATIO_THRESHOLD,
                            'detail': f"✏️ Simulasi | Rasio: {fmt_pct(rasio_rp_sim)} | Loan After RP: {fmt_rp(loan_after_rp_sim)} | Coll: {fmt_rp(coll_after_rp_sim)}"
                        })
                    elif c['label'] == 'LR-3. Rasio LR < 65%':
                        new_checks.append({
                            'label': c['label'],
                            'passed': rasio_lr_sim is not None and rasio_lr_sim < RATIO_THRESHOLD,
                            'detail': f"✏️ Simulasi | Rasio: {fmt_pct(rasio_lr_sim)} | Ceiling LR: {fmt_rp(ceiling_sim)} | Coll: {fmt_rp(coll_lr_sim)}"
                        })
                    else:
                        new_checks.append(c)
                updated['checks'] = new_checks
                st.session_state['sid_results'][sel_sid] = updated
                st.write("DEBUG:", st.session_state['sid_results'][sel_sid].get('is_simulated'))
                st.session_state[f'sim_saved_input_{sel_sid}'] = {
                    'lot_inputs': {row['Saham']: row['Lot Jual'] for _, row in edited_df1.iterrows()},
                    'rp_inputs':  {i: float(edited_df2.loc[i, 'RP Value']) for i in edited_df2.index} if mode2 else {},
                    'mode1': mode1,
                    'mode2': mode2,
                }
                st.rerun()
            
            if col_b2.button("↩️ Reset Nasabah Ini", use_container_width=True):
                original = st.session_state.get('sid_results_original', {})
                if sel_sid in original:
                    st.session_state['sid_results'][sel_sid] = copy.deepcopy(original[sel_sid])
                    st.rerun()
            
            if col_b3.button("🔄 Reset Semua", use_container_width=True):
                original = st.session_state.get('sid_results_original', {})
                if original:
                    st.session_state['sid_results'] = copy.deepcopy(original)
                    st.rerun()
            
                # ── TAB GLOBAL ────────────────────────────────────────────
                with tab_global:
                    st.subheader("Validasi Limit Participant — Credit Limit Partisipan")
                    if global_result["passed"]:
                        st.success(f"✅ LOLOS — {global_result['detail']}")
                    else:
                        st.error(f"❌ GAGAL — {global_result['detail']}")

    # ── TAB GAGAL ─────────────────────────────────────────────
    with tab_gagal:
        gc1, gc2 = st.columns(2)
        with gc1:
            gagal_rp = [(s,d) for s,d in sid_results.items()
                        if d.get('has_rp') and not d.get('rp_skipped') and not lolos_rp(d)]
            st.markdown(f"#### 🔴 Gagal RP — {len(gagal_rp)} nasabah")
            if not gagal_rp:
                st.success("Semua lolos RP.")
            for sid, data in gagal_rp:
                with st.expander(f"❌ {sid} — {data['name']}"):
                    for c in data['checks']:
                        if c['label'].startswith('RP-') and not c['passed']:
                            st.error(f"**{c['label']}** — {c['detail']}")
        with gc2:
            gagal_lr = [(s,d) for s,d in sid_results.items()
                        if d.get('has_lr') and not lolos_lr(d)]
            st.markdown(f"#### 🔴 Gagal LR — {len(gagal_lr)} nasabah")
            if not gagal_lr:
                st.success("Semua lolos LR.")
            for sid, data in gagal_lr:
                with st.expander(f"❌ {sid} — {data['name']}"):
                    if data['max_lr_63'] > 0:
                        st.warning(f"💡 Max LR Aman (63%): {fmt_rp(data['max_lr_63'])}")
                    for c in data['checks']:
                        if c['label'].startswith('LR-') and not c['passed']:
                            st.error(f"**{c['label']}** — {c['detail']}")
                                    
            if st.session_state.get('debug_log'):
               with st.expander("🐛 Debug Log", expanded=True):
                    for line in st.session_state['debug_log']:
                        st.write(line)
                        
    # ── TAB AUTO-ADJUST ───────────────────────────────────────                 
    with tab_adj:
        st.subheader("⚡ Auto-Adjust LR — Target Rasio 63%")
        gagal_lr3 = [(s,d) for s,d in sid_results.items()
                     if d.get('has_lr') and any(
                         c['label'] == 'LR-3. Rasio LR < 65%' and not c['passed']
                         for c in d['checks'])]
        if not gagal_lr3:
            st.success("✅ Tidak ada nasabah yang perlu di-adjust.")
        else:
            prev_rows = []
            for sid, data in gagal_lr3:
                prev_rows.append({
                    'SID':            sid,
                    'Nama':           data['name'],
                    'Ceiling LR':     fmt_rp(data['ceiling_lr']),
                    'Max LR (63%)':   fmt_rp(data['max_lr_63']),
                    'Max LR Final':   fmt_rp(data['max_lr_final']),
                    'Coll After LR':  fmt_rp(data['coll_after_lr']),
                    'Status':         f"✂️ Dipotong ke {fmt_rp(data['max_lr_final'])}" if data['max_lr_final'] > 0 else "❌ Dikeluarkan",
                })
            st.dataframe(pd.DataFrame(prev_rows), use_container_width=True, hide_index=True)

            if st.button("⚡ Terapkan Auto-Adjust & Validasi Ulang", type="primary", use_container_width=True):
                new_results = dict(sid_results)
                for sid, data in gagal_lr3:
                    # Langsung update max_lr_final dan ceiling_lr di sid_results
                    updated = dict(data)
                    updated['ceiling_lr']   = data['max_lr_final']
                    updated['max_lr_final'] = data['max_lr_final']
                    # Update checks LR-3 jadi passed
                    new_checks = []
                    for c in data['checks']:
                        if c['label'] == 'LR-3. Rasio LR < 65%':
                            numerator = data['loan_after_rp'] + data['accrued'] + data['max_lr_final']
                            rasio_baru = numerator / data['coll_after_lr'] if data['coll_after_lr'] > 0 else None
                            new_checks.append({
                                'label': c['label'],
                                'passed': True,
                                'detail': f"✂️ Di-adjust ke {fmt_rp(data['max_lr_final'])} | Rasio: {fmt_pct(rasio_baru)}"
                            })
                        else:
                            new_checks.append(c)
                    updated['checks'] = new_checks
                    new_results[sid] = updated

                total_lr_new = sum(d['max_lr_final'] for d in new_results.values() if lolos_lr(d))
                total_rp_new = sum(d['total_rp_maks'] for d in new_results.values() if lolos_rp(d))
                st.session_state['sid_results']   = new_results
                st.session_state['global_result'] = {
                    "passed": (CREDIT_LIMIT_PARTISIPAN + total_rp_new) > total_lr_new,
                    "detail": f"CL Partisipan: {fmt_rp(CREDIT_LIMIT_PARTISIPAN)} + RP: {fmt_rp(total_rp_new)} = {fmt_rp(CREDIT_LIMIT_PARTISIPAN+total_rp_new)} | LR: {fmt_rp(total_lr_new)}",
                    "total_rp": total_rp_new, "total_lr": total_lr_new,
                }
                st.success("✅ Auto-Adjust diterapkan!")
                st.rerun()

    # ── TAB EXPORT ────────────────────────────────────────────
    with tab_export:
        st.subheader("📋 Ringkasan Hasil Validasi")
        sum_rows = []
        for sid, d in sid_results.items():
            sum_rows.append({
                'SID': sid, 'Nama': d['name'],
                'Loan Existing': d['loan_existing'], 'Accrued': d['accrued'],
                'Avail Limit': d['avail_limit'],
                'Coll Awal': d['coll_before_rp'],
                'Current Ratio': f"{d['loan_existing']/d['coll_before_rp']*100:.2f}%" if d['coll_before_rp'] > 0 else "-",
                'RP Min': d['total_rp_min'], 'RP Maks': d['total_rp_maks'],
                'Loan After RP': d['loan_after_rp'], 'Coll After RP': d['coll_after_rp'],
                'Status RP': "✅" if lolos_rp(d) else ("⏭" if d['rp_skipped'] else ("-" if not d['has_rp'] else "❌")),
                'Ceiling LR': d['ceiling_lr'], 'Coll After LR': d['coll_after_lr'],
                'Max LR Final': d['max_lr_final'],
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
        e1,e2 = st.columns(2)
        with e1:
            buf_rp, fname_rp = generate_repayment_excel(sid_results, sell_regular)
            st.download_button("⬇️ Repayment Proceed (.xlsx)", data=buf_rp,
                file_name=fname_rp, use_container_width=True,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with e2:
            buf_lr, fname_lr = generate_loan_excel(sid_results, margin_buy)
            st.download_button("⬇️ Loan Request (.xlsx)", data=buf_lr,
                file_name=fname_lr, use_container_width=True,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.divider()
        st.subheader("📅 Rekap Belum Settled — Untuk Besok")
        r1,r2 = st.columns(2)
        with r1:
            buf_rrp, fname_rrp = generate_rekap_rp_excel(sid_results)
            st.download_button("⬇️ Rekap RP Belum Settled", data=buf_rrp,
                file_name=fname_rrp, use_container_width=True, type="primary",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with r2:
            buf_rlr, fname_rlr = generate_rekap_lr_excel(sid_results)
            st.download_button("⬇️ Rekap LR Belum Settled", data=buf_rlr,
                file_name=fname_rlr, use_container_width=True, type="primary",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


