import streamlit as st
import pandas as pd
import io
from datetime import datetime

if 'df_sell_edited' not in st.session_state:
    st.session_state['df_sell_edited'] = None
if 'sid_results' not in st.session_state:
    st.session_state['sid_results'] = None
if 'global_result' not in st.session_state:
    st.session_state['global_result'] = None
if 'clamped_warnings' not in st.session_state:
    st.session_state['clamped_warnings'] = []
if 'df_buy' not in st.session_state:
    st.session_state['df_buy'] = None
if 'closing_prices' not in st.session_state:
    st.session_state['closing_prices'] = {}
if 'risk_params' not in st.session_state:
    st.session_state['risk_params'] = {}
if 'lr_data' not in st.session_state:
    st.session_state['lr_data'] = {}
if 'rp_data' not in st.session_state:
    st.session_state['rp_data'] = {}
if 'df_buy_adjusted' not in st.session_state:
    st.session_state['df_buy_adjusted'] = None

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Validasi Input LR & RP",
    page_icon="✅",
    layout="wide",
)

st.title("✅ Validasi MNC")
st.info(
    "Sistem Validasi Repayment & Loan Request nasabah PEI: "
    "**Repayment Proceed** · **Loan Request** · **Credit Limit**"
)

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

def fmt_rp(val):
    try:
        return f"Rp {val:,.2f}"
    except Exception:
        return str(val)

def fmt_pct(val):
    try:
        return f"{val*100:.2f}%"
    except Exception:
        return str(val)

def parse_hc(raw):
    try:
        s = str(raw).replace("%", "").strip()
        v = float(s)
        return v / 100 if v > 1 else v
    except Exception:
        return 0.0

# ─────────────────────────────────────────────
# PARSERS
# ─────────────────────────────────────────────

def parse_op_file(content: str):
    result = {}
    lines  = content.strip().splitlines()
    for line in lines:
        line  = line.strip()
        if not line:
            continue
        parts = line.split("|")
        if parts[0] == "0":
            if len(parts) < 7:
                continue
            sid = parts[3].strip()
            try:
                loan_existing    = float(parts[5])
            except Exception:
                loan_existing    = 0.0
            try:
                accrued_interest = float(parts[6])
            except Exception:
                accrued_interest = 0.0
            result[sid] = {
                "loan_existing":    loan_existing,
                "accrued_interest": accrued_interest,
                "volume_existing":  0.0,
                "name":             parts[4].strip() if len(parts) > 4 else sid,
                "stocks":           {},
            }
        elif parts[0] == "1":
            if len(parts) < 5:
                continue
            sid   = parts[2].strip()
            stock = parts[3].strip()
            try:
                vol = float(parts[4])
            except Exception:
                vol = 0.0
            if sid in result:
                result[sid]["volume_existing"]     += vol
                result[sid]["stocks"][stock]         = result[sid]["stocks"].get(stock, 0.0) + vol
    return result


def parse_credit_limit_file(content: str):
    result     = {}
    value_date = None
    lines      = content.strip().splitlines()

    for i, line in enumerate(lines):
        line  = line.strip()
        if not line:
            continue
        parts = line.split("|")

        if i == 0 and parts[0].strip().lower() == "value date":
            continue

        if len(parts) < 7:
            continue

        sid = parts[2].strip()
        try:
            available_limit = float(parts[6].replace(",", ""))
        except Exception:
            available_limit = 0.0

        if value_date is None:
            value_date = parts[0].strip()

        result[sid] = {
            "available_limit": available_limit,
            "name":            parts[3].strip(),
            "value_date":      parts[0].strip(),
        }

    return result, value_date


def parse_lr_excel(uploaded_file) -> dict:
    """
    Baca file LR belum settled dari Excel.
    Kolom: SID (A), Name (B), Loan Value (C)
    Return: ({sid: loan_value}, {sid: name})
    """
    df = pd.read_excel(uploaded_file, header=0, dtype=str)
    df.columns = df.columns.str.strip()
    sid_col  = df.columns[0]
    name_col = df.columns[1]
    val_col  = df.columns[2]
    result   = {}
    name_map = {}
    for _, row in df.iterrows():
        sid = str(row[sid_col]).strip()
        if not sid or sid == 'nan':
            continue
        val_str = str(row[val_col]).replace('.', '').replace(',', '').strip()
        try:
            val = float(val_str)
        except Exception:
            val = 0.0
        name = str(row[name_col]).strip()
        result[sid]   = result.get(sid, 0.0) + val
        name_map[sid] = name
    return result, name_map


def parse_rp_excel(uploaded_file) -> dict:
    """
    Baca file RP belum settled dari Excel.
    Kolom: SID (A), Name (B), Repayment Value (C)
    Return: ({sid: repayment_value}, {sid: name})
    """
    df = pd.read_excel(uploaded_file, header=0, dtype=str)
    df.columns = df.columns.str.strip()
    sid_col  = df.columns[0]
    name_col = df.columns[1]
    val_col  = df.columns[2]
    result   = {}
    name_map = {}
    for _, row in df.iterrows():
        sid = str(row[sid_col]).strip()
        if not sid or sid == 'nan':
            continue
        val_str = str(row[val_col]).replace('.', '').replace(',', '').strip()
        try:
            val = float(val_str)
        except Exception:
            val = 0.0
        name = str(row[name_col]).strip()
        result[sid]   = result.get(sid, 0.0) + val
        name_map[sid] = name
    return result, name_map


def load_closing_price(uploaded_file) -> dict:
    df = pd.read_excel(uploaded_file, sheet_name=0, header=0)
    result = {}
    for _, row in df.iterrows():
        code  = str(row['no_share']).strip().upper()
        price = pd.to_numeric(str(row['kurs_now']).replace(',', ''), errors='coerce')
        if pd.notna(price) and code and code != 'NAN':
            result[code] = float(price)
    return result

def load_risk_parameter(uploaded_file) -> dict:
    result = {}
    content = uploaded_file.read().decode("utf-8", errors="replace")
    uploaded_file.seek(0)
    for line in content.strip().splitlines():
        line = line.strip()
        if not line or line.startswith("StockCode"):
            continue
        parts = line.split("|")
        if len(parts) < 3:
            continue
        code = parts[0].strip()
        try:
            hc = float(parts[2]) / 100.0
        except Exception:
            hc = 0.0
        result[code] = hc
    return result


def load_hasil_mnc(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    if "Sell Aktif (tanpa excluded)" in xls.sheet_names:
        df_sell = pd.read_excel(xls, sheet_name="Sell Aktif (tanpa excluded)", header=0)
    else:
        df_sell = pd.read_excel(xls, sheet_name="Sell (Repayment)", header=0)
        if 'NETT' in df_sell.columns:
            df_sell = df_sell[~df_sell['NETT'].astype(str).str.contains('EXCLUDED', na=False)].copy()
    df_buy = pd.read_excel(xls, sheet_name="Buy (Loan)", header=0)
    return df_sell, df_buy

# ─────────────────────────────────────────────
# COLUMN ACCESSORS (0-based)
# ─────────────────────────────────────────────

def col(df, idx):
    return df.iloc[:, idx]

SELL_SID   = 0
SELL_STOCK = 1
SELL_AVQ   = 4
SELL_CP    = 5
SELL_VOL   = 11
SELL_VAL   = 12

BUY_SID    = 0
BUY_STOCK  = 1
BUY_AVQ    = 4
BUY_CP     = 5
BUY_HC     = 7
BUY_VOL    = 13
BUY_VAL    = 14

RATIO_THRESHOLD    = 0.65
AUTO_ADJUST_TARGET = 0.63

# ─────────────────────────────────────────────
# COLLATERAL CALCULATOR
# ─────────────────────────────────────────────

def calc_collateral_from_op(op_stocks, closing_prices, risk_params):
    total  = 0.0
    detail = []
    for stock, qty in op_stocks.items():
        cp   = closing_prices.get(stock, 0.0)
        hc   = risk_params.get(stock, 0.0)
        coll = qty * cp * (1 - hc)
        total += coll
        detail.append({"stock": stock, "qty": qty, "cp": cp, "hc": hc, "collateral": coll})
    return total, detail


def calc_ratio_baru(loan_existing, accrued_interest, lr_value, rp_value, collateral):
    numerator = loan_existing + accrued_interest + lr_value - rp_value
    if collateral > 0:
        return numerator / collateral, numerator
    return None, numerator


def calc_max_loan_baru(loan_existing, accrued_interest, lr_existing, rp_existing, collateral, available_limit, total_sell_val, threshold=0.63):
    current_numerator = loan_existing + accrued_interest + lr_existing - rp_existing
    max_from_ratio    = collateral * threshold - current_numerator
    effective_cl      = available_limit + total_sell_val
    return max(min(max_from_ratio, effective_cl), 0)
    


# ─────────────────────────────────────────────
# AUTO-ADJUST LOAN (proporsional per saham)
# ─────────────────────────────────────────────

def auto_adjust_loan(df_buy, sid, max_loan_value, closing_prices):
    df_updated = df_buy.copy()
    rows_idx   = df_buy[col(df_buy, BUY_SID).astype(str) == sid].index

    if len(rows_idx) == 0:
        return df_updated

    total_val_now = sum(
        pd.to_numeric(df_buy.at[i, df_buy.columns[BUY_VAL]], errors='coerce') or 0
        for i in rows_idx
    )

    if total_val_now <= 0 or max_loan_value <= 0:
        for i in rows_idx:
            df_updated.at[i, df_updated.columns[BUY_VOL]] = 0
            df_updated.at[i, df_updated.columns[BUY_VAL]] = 0
        return df_updated

    ratio = max_loan_value / total_val_now

    for i in rows_idx:
        old_vol = pd.to_numeric(df_buy.at[i, df_buy.columns[BUY_VOL]], errors='coerce') or 0
        stock   = str(df_buy.at[i, df_buy.columns[BUY_STOCK]]).strip().upper()
        price   = closing_prices.get(stock, 0)
        new_vol = int((old_vol * ratio) // 100) * 100
        if new_vol < 0:
            new_vol = 0
        new_val = new_vol * price if price > 0 else 0
        df_updated.at[i, df_updated.columns[BUY_VOL]] = new_vol
        df_updated.at[i, df_updated.columns[BUY_VAL]] = new_val

    return df_updated


# ─────────────────────────────────────────────
# VALIDATION LOGIC
# ─────────────────────────────────────────────

def run_validations(df_sell, df_buy, op_data, cl_data, credit_limit_partisipan,
                    closing_prices, risk_params, lr_data, rp_data):
    results = {}

    all_sids = sorted(set(
        col(df_sell, SELL_SID).dropna().astype(str).unique().tolist() +
        col(df_buy,  BUY_SID ).dropna().astype(str).unique().tolist()
    ))

    def agg_sell(sid):
        rows      = df_sell[col(df_sell, SELL_SID).astype(str) == sid]
        total_vol = pd.to_numeric(col(rows, SELL_VOL), errors="coerce").abs().sum()
        total_val = pd.to_numeric(col(rows, SELL_VAL), errors="coerce").sum()
        return {"total_volume": total_vol, "total_value": total_val, "rows": rows}

    def agg_buy(sid):
        rows      = df_buy[col(df_buy, BUY_SID).astype(str) == sid]
        total_vol = pd.to_numeric(col(rows, BUY_VOL), errors="coerce").sum()
        total_val = pd.to_numeric(col(rows, BUY_VAL), errors="coerce").sum()
        return {"total_volume": total_vol, "total_value": total_val, "rows": rows}

    global_total_sell_value = pd.to_numeric(col(df_sell, SELL_VAL), errors="coerce").sum()
    global_total_buy_value  = pd.to_numeric(col(df_buy,  BUY_VAL),  errors="coerce").sum()

    for sid in all_sids:
        sell = agg_sell(sid)
        buy  = agg_buy(sid)
        op   = op_data.get(sid, {
            "loan_existing": 0, "accrued_interest": 0,
            "volume_existing": 0, "name": sid, "stocks": {}
        })
        cl   = cl_data.get(sid, {"available_limit": 0, "name": sid})

        loan_existing    = op["loan_existing"]
        accrued_interest = op["accrued_interest"]
        available_limit  = cl["available_limit"]
        name             = op.get("name") or cl.get("name") or sid
        op_stocks        = op.get("stocks", {})

        lr_value = lr_data.get(sid, 0.0)
        rp_value = rp_data.get(sid, 0.0)

        total_sell_vol = sell["total_volume"]
        total_sell_val = sell["total_value"]
        total_buy_vol  = buy["total_volume"]
        total_buy_val  = buy["total_value"]

        has_repayment    = total_sell_vol > 0
        has_loan_request = total_buy_vol  > 0

        collateral_existing, _ = calc_collateral_from_op(op_stocks, closing_prices, risk_params)
        repayment_skipped_no_loan = has_repayment and loan_existing <= 0

        max_loan_63 = calc_max_loan_baru(
            loan_existing, accrued_interest, lr_value,
            rp_value + total_sell_val, collateral_existing,
            available_limit, total_sell_val,
            threshold=AUTO_ADJUST_TARGET
        )

        sid_results = {
            "name":                      name,
            "checks":                    [],
            "has_repayment":             has_repayment,
            "has_loan_request":          has_loan_request,
            "repayment_skipped_no_loan": repayment_skipped_no_loan,
            "loan_existing":             loan_existing,
            "max_loan_recommendation":   0.0,
            "max_loan_63":               max_loan_63,
            "collateral_existing":       collateral_existing,
            "lr_value":                  lr_value,
            "rp_value":                  rp_value,
            "total_buy_val":             total_buy_val,
            "total_sell_val":            total_sell_val,
            "accrued_interest":          accrued_interest,
        }

        def add(label, passed, detail=""):
            sid_results["checks"].append({"label": label, "passed": passed, "detail": detail})

        # ── 1a ──────────────────────────────────────────────────
        if repayment_skipped_no_loan:
            add("1a. Volume Sell ≤ Available Sell Quantity", True,
                "⏭ Dilewati — Existing Loan = 0, repayment tidak diperlukan")
        else:
            rep_1a_pass   = True
            rep_1a_detail = []
            for _, row in sell["rows"].iterrows():
                vol = abs(pd.to_numeric(row.iloc[SELL_VOL], errors="coerce") or 0)
                avq = pd.to_numeric(row.iloc[SELL_AVQ], errors="coerce") or 0
                stk = str(row.iloc[SELL_STOCK])
                if vol > avq:
                    rep_1a_pass = False
                    rep_1a_detail.append(f"{stk}: Vol {vol:,.0f} > Avail {avq:,.0f}")
            add("1a. Volume Sell ≤ Available Sell Quantity", rep_1a_pass,
                "; ".join(rep_1a_detail) if rep_1a_detail else f"Total Volume Sell: {total_sell_vol:,.0f}")

        # ── 1a-OP ────────────────────────────────────────────────
        if repayment_skipped_no_loan:
            add("1a-OP. Saham Sell Terverifikasi di OP File", True,
                "⏭ Dilewati — Existing Loan = 0, repayment tidak diperlukan")
        elif not has_repayment:
            add("1a-OP. Saham Sell Terverifikasi di OP File", True, "Tidak ada Repayment")
        elif not op_stocks:
            add("1a-OP. Saham Sell Terverifikasi di OP File", False,
                f"SID {sid} tidak ditemukan di OP file")
        else:
            op_1a_pass      = True
            op_1a_detail    = []
            op_1a_ok_detail = []
            op_1a_adjusted  = []
            for _, row in sell["rows"].iterrows():
                stk = str(row.iloc[SELL_STOCK]).strip()
                vol = abs(pd.to_numeric(row.iloc[SELL_VOL], errors="coerce") or 0)
                if stk not in op_stocks:
                    op_1a_pass = False
                    op_1a_detail.append(f"{stk}: tidak ada di OP file")
                elif vol > op_stocks[stk]:
                    avail = op_stocks[stk]
                    op_1a_adjusted.append(f"{stk}: {vol:,.0f} → {avail:,.0f} (dipotong sesuai OP)")
                    op_1a_ok_detail.append(f"{stk}: {avail:,.0f} ✓ (adjusted)")
                    df_sell.loc[
                        (col(df_sell, SELL_SID).astype(str) == sid) &
                        (col(df_sell, SELL_STOCK).astype(str).str.strip() == stk),
                        df_sell.columns[SELL_VOL]
                    ] = avail
                else:
                    op_1a_ok_detail.append(f"{stk}: {vol:,.0f} ✓")
            if op_1a_adjusted:
                detail_msg = "⚠️ Auto-adjusted: " + "; ".join(op_1a_adjusted) + (
                    " || OK: " + "; ".join(op_1a_ok_detail) if op_1a_ok_detail else "")
            else:
                detail_msg = ("; ".join(op_1a_ok_detail) if op_1a_pass
                    else "; ".join(op_1a_detail) + (
                        " || OK: " + "; ".join(op_1a_ok_detail) if op_1a_ok_detail else ""))
            add("1a-OP. Saham Sell Terverifikasi di OP File", op_1a_pass, detail_msg)

        # ── 1b ──────────────────────────────────────────────────
        if repayment_skipped_no_loan:
            add("1b. Total Repayment Value ≤ Total Loan Value", True,
                "⏭ Dilewati — Existing Loan = 0, repayment tidak diperlukan")
        elif not has_loan_request:
            add("1b. Total Repayment Value ≤ Total Loan Value", True,
                "Pure repayment tanpa Loan Request — check 1b dilewati")
        else:
            rep_1b_pass = total_sell_val <= total_buy_val
            add("1b. Total Repayment Value ≤ Total Loan Value", rep_1b_pass,
                f"Total Sell Value: {fmt_rp(total_sell_val)} | Total Buy Value: {fmt_rp(total_buy_val)}")

        # ── 1c ──────────────────────────────────────────────────
        if repayment_skipped_no_loan:
            add("1c. Rasio Repayment < 65%", True,
                "⏭ Dilewati — Existing Loan = 0, repayment tidak diperlukan")
        elif not has_repayment:
            add("1c. Rasio Repayment < 65%", True, "Tidak ada Repayment — check 1c dilewati")
        else:
            rp_after_repayment = rp_value + total_sell_val
            ratio_rep, num_rep = calc_ratio_baru(
                loan_existing, accrued_interest, lr_value, rp_after_repayment, collateral_existing)
            if ratio_rep is not None:
                add("1c. Rasio Repayment < 65%", ratio_rep < RATIO_THRESHOLD,
                    f"Rasio: {fmt_pct(ratio_rep)} | Numerator: {fmt_rp(num_rep)} "
                    f"(Loan: {fmt_rp(loan_existing)} + Accrued: {fmt_rp(accrued_interest)} "
                    f"+ LR: {fmt_rp(lr_value)} - RP: {fmt_rp(rp_after_repayment)}) | "
                    f"Collateral: {fmt_rp(collateral_existing)}")
            elif num_rep <= 0:
                add("1c. Rasio Repayment < 65%", True, f"Numerator ≤ 0 ({fmt_rp(num_rep)}) — posisi lunas")
            else:
                add("1c. Rasio Repayment < 65%", False,
                    "Collateral = 0 sementara Loan masih ada — Rasio tidak terhitung (∞)")

        # ── 2a ──────────────────────────────────────────────────
        loan_2a_pass   = True
        loan_2a_detail = []
        for _, row in buy["rows"].iterrows():
            vol = pd.to_numeric(row.iloc[BUY_VOL], errors="coerce") or 0
            avq = pd.to_numeric(row.iloc[BUY_AVQ], errors="coerce") or 0
            stk = str(row.iloc[BUY_STOCK])
            if avq == 0:
                loan_2a_pass = False
                loan_2a_detail.append(f"{stk}: DIBATALKAN (Stock Available = 0)")
            elif vol > avq:
                loan_2a_pass = False
                loan_2a_detail.append(f"{stk}: Vol {vol:,.0f} > Avail {avq:,.0f}")
        add("2a. Volume Buy ≤ Available Quantity", loan_2a_pass,
            "; ".join(loan_2a_detail) if loan_2a_detail else f"Total Volume Buy: {total_buy_vol:,.0f}")

        # ── 2b ──────────────────────────────────────────────────
        if not has_loan_request:
            add("2b. Rasio Loan Request < 65%", True, "Tidak ada Loan Request — check 2b dilewati")
        else:
            lr_after_buy  = lr_value + total_buy_val
            rp_after_sell = rp_value + total_sell_val
            ratio_loan, num_loan = calc_ratio_baru(
                loan_existing, accrued_interest, lr_after_buy, rp_after_sell, collateral_existing)
            max_loan_rec = calc_max_loan_baru(
                loan_existing, accrued_interest, lr_value, rp_value + total_sell_val,
                collateral_existing, available_limit, total_sell_val,
                threshold=RATIO_THRESHOLD)
            sid_results["max_loan_recommendation"] = max_loan_rec

            if ratio_loan is not None:
                loan_2b_pass = ratio_loan < RATIO_THRESHOLD
                detail_str = (
                    f"Rasio: {fmt_pct(ratio_loan)} (threshold <{RATIO_THRESHOLD*100:.0f}%) | "
                    f"Numerator: {fmt_rp(num_loan)} "
                    f"(Loan: {fmt_rp(loan_existing)} + Accrued: {fmt_rp(accrued_interest)} "
                    f"+ LR baru: {fmt_rp(lr_after_buy)} - RP: {fmt_rp(rp_after_sell)}) | "
                    f"Collateral: {fmt_rp(collateral_existing)}"
                )
                if not loan_2b_pass:
                    detail_str += (
                        f" || ⚠️ Loan diajukan ({fmt_rp(total_buy_val)}) melebihi kapasitas. "
                        f"Max loan aman (63%): {fmt_rp(max_loan_63)}"
                    )
                add("2b. Rasio Loan Request < 65%", loan_2b_pass, detail_str)
            elif num_loan <= 0:
                add("2b. Rasio Loan Request < 65%", True,
                    f"Numerator ≤ 0 ({fmt_rp(num_loan)}) — posisi lunas")
            else:
                add("2b. Rasio Loan Request < 65%", False,
                    "Collateral = 0 sementara ada Loan — Rasio tidak terhitung (∞)")

        # ── 3 ───────────────────────────────────────────────────
        if not has_loan_request:
            add("3. Credit Limit Nasabah", True, "Tidak ada Loan Request — check 3 dilewati")
        else:
            effective_limit = available_limit + total_sell_val
            cl_nasabah_pass = effective_limit > total_buy_val
            add("3. Credit Limit Nasabah", cl_nasabah_pass,
                f"Avail Limit: {fmt_rp(available_limit)} + Sell: {fmt_rp(total_sell_val)} = "
                f"{fmt_rp(effective_limit)} | Loan Diajukan: {fmt_rp(total_buy_val)}")

        results[sid] = sid_results

    # ── 4 ───────────────────────────────────────────────────────
    cl_partisipan_pass = (credit_limit_partisipan + global_total_sell_value) > global_total_buy_value
    global_result = {
        "passed": cl_partisipan_pass,
        "detail": (
            f"CL Partisipan: {fmt_rp(credit_limit_partisipan)} + "
            f"Total Sell: {fmt_rp(global_total_sell_value)} = "
            f"{fmt_rp(credit_limit_partisipan + global_total_sell_value)} | "
            f"Total Loan Diajukan: {fmt_rp(global_total_buy_value)}"
        ),
        "total_sell": global_total_sell_value,
        "total_buy":  global_total_buy_value,
    }

    return results, global_result

# ─────────────────────────────────────────────
# LOLOS CHECKER
# ─────────────────────────────────────────────

def lolos_repayment(sid_data):
    if sid_data.get("repayment_skipped_no_loan"):
        return False
    if not sid_data.get("has_repayment", False):
        return False
    repayment_labels = {"1a.", "1a-OP.", "1b.", "1c."}
    return all(
        c["passed"] for c in sid_data["checks"]
        if any(c["label"].startswith(p) for p in repayment_labels)
    )

def lolos_loan(sid_data):
    if not sid_data.get("has_loan_request", False):
        return False
    loan_labels = {"1a.", "1a-OP.", "1c.", "2a.", "2b.", "3."}
    return all(
        c["passed"] for c in sid_data["checks"]
        if any(c["label"].startswith(p) for p in loan_labels)
    )

# ─────────────────────────────────────────────
# EXCEL EXPORT
# ─────────────────────────────────────────────

def generate_repayment_excel(df_sell, sid_results):
    today_str   = datetime.today().strftime("%Y%m%d")
    passed_sids = [sid for sid, data in sid_results.items() if lolos_repayment(data)]
    sheet1_rows = []
    for sid in passed_sids:
        rows      = df_sell[col(df_sell, SELL_SID).astype(str) == sid]
        total_val = pd.to_numeric(col(rows, SELL_VAL), errors="coerce").abs().sum()
        total_vol = pd.to_numeric(col(rows, SELL_VOL), errors="coerce").abs().sum()
        if total_vol > 0:
            sheet1_rows.append({"Participant Code": "EP", "SID Client": sid,
                                 "Repayment Value": total_val})
    sheet2_rows = []
    for sid in passed_sids:
        rows = df_sell[col(df_sell, SELL_SID).astype(str) == sid]
        for _, row in rows.iterrows():
            qty = abs(pd.to_numeric(row.iloc[SELL_VOL], errors="coerce") or 0)
            if qty > 0:
                sheet2_rows.append({"SID Client": sid,
                                    "Stock Code": str(row.iloc[SELL_STOCK]),
                                    "Quantity":   qty})
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(sheet1_rows).to_excel(writer, sheet_name="Repayment Proceed", index=False)
        pd.DataFrame(sheet2_rows).to_excel(writer, sheet_name="Detail Collateral",  index=False)
    buf.seek(0)
    return buf, f"Repayment Proceed {today_str}.xlsx"


def generate_loan_excel(df_buy, sid_results):
    today_str   = datetime.today().strftime("%Y%m%d")
    passed_sids = [sid for sid, data in sid_results.items() if lolos_loan(data)]
    sheet1_rows = []
    for sid in passed_sids:
        rows      = df_buy[col(df_buy, BUY_SID).astype(str) == sid]
        total_val = pd.to_numeric(col(rows, BUY_VAL), errors="coerce").sum()
        total_vol = pd.to_numeric(col(rows, BUY_VOL), errors="coerce").sum()
        if total_vol > 0:
            sheet1_rows.append({"Participant Code": "EP", "SID Client": sid,
                                 "Loan Value": total_val})
    sheet2_rows = []
    for sid in passed_sids:
        rows = df_buy[col(df_buy, BUY_SID).astype(str) == sid]
        for _, row in rows.iterrows():
            qty = pd.to_numeric(row.iloc[BUY_VOL], errors="coerce") or 0
            if qty > 0:
                sheet2_rows.append({"SID Client": sid,
                                    "Stock Code": str(row.iloc[BUY_STOCK]),
                                    "Quantity":   qty})
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(sheet1_rows).to_excel(writer, sheet_name="Loan Request",      index=False)
        pd.DataFrame(sheet2_rows).to_excel(writer, sheet_name="Detail Collateral", index=False)
    buf.seek(0)
    return buf, f"Loan Request {today_str}.xlsx"


def generate_lr_rekap_excel(df_buy, sid_results):
    """Rekap LR lolos hari ini → input LR belum settled besok"""
    today_str   = datetime.today().strftime("%Y%m%d")
    passed_sids = [sid for sid, data in sid_results.items() if lolos_loan(data)]
    rows = []
    for sid in passed_sids:
        buy_rows  = df_buy[col(df_buy, BUY_SID).astype(str) == sid]
        total_val = pd.to_numeric(col(buy_rows, BUY_VAL), errors="coerce").sum()
        total_vol = pd.to_numeric(col(buy_rows, BUY_VOL), errors="coerce").sum()
        if total_vol > 0:
            rows.append({
                "SID":         sid,
                "Name":        sid_results[sid].get("name", sid),
                "Loan Value":  total_val,
            })
    buf = io.BytesIO()
    pd.DataFrame(rows).to_excel(buf, index=False)
    buf.seek(0)
    return buf, f"LR Belum Settled {today_str}.xlsx"


def generate_rp_rekap_excel(df_sell, sid_results):
    """Rekap RP lolos hari ini → input RP belum settled besok"""
    today_str   = datetime.today().strftime("%Y%m%d")
    passed_sids = [sid for sid, data in sid_results.items() if lolos_repayment(data)]
    rows = []
    for sid in passed_sids:
        sell_rows = df_sell[col(df_sell, SELL_SID).astype(str) == sid]
        total_val = pd.to_numeric(col(sell_rows, SELL_VAL), errors="coerce").abs().sum()
        total_vol = pd.to_numeric(col(sell_rows, SELL_VOL), errors="coerce").abs().sum()
        if total_vol > 0:
            rows.append({
                "SID":              sid,
                "Name":             sid_results[sid].get("name", sid),
                "Repayment Value":  total_val,
            })
    buf = io.BytesIO()
    pd.DataFrame(rows).to_excel(buf, index=False)
    buf.seek(0)
    return buf, f"RP Belum Settled {today_str}.xlsx"


def generate_revisi_repayment_excel(df_sell, sid_results, op_data):
    today_str   = datetime.today().strftime("%Y%m%d")
    failed_sids = [
        sid for sid, data in sid_results.items()
        if data.get("has_repayment") and not data.get("repayment_skipped_no_loan")
        and not lolos_repayment(data)
    ]
    alasan_rows = []
    for sid in failed_sids:
        data = sid_results[sid]
        op   = op_data.get(sid, {})
        for check in data["checks"]:
            if check["label"].startswith(("1a.", "1a-OP.", "1b.", "1c.")) and not check["passed"]:
                alasan_rows.append({
                    "SID": sid, "Nama": data["name"],
                    "Check Gagal": check["label"], "Detail Alasan": check["detail"],
                    "Loan Existing": op.get("loan_existing", "-"),
                    "Accrued Interest": op.get("accrued_interest", "-"),
                    "Volume Existing": op.get("volume_existing", "-"),
                    "Saham OP": ", ".join(op.get("stocks", {}).keys()) or "-",
                    "Collateral": data.get("collateral_existing", "-"),
                    "LR Belum Settled": data.get("lr_value", 0),
                    "RP Belum Settled": data.get("rp_value", 0),
                })
    df_failed_sell = df_sell[col(df_sell, SELL_SID).astype(str).isin([str(s) for s in failed_sids])].copy()
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(alasan_rows).to_excel(writer, sheet_name="Alasan Gagal",    index=False)
        df_failed_sell.to_excel(writer,            sheet_name="Sell (Repayment)", index=False)
    buf.seek(0)
    return buf, f"Revisi Repayment {today_str}.xlsx", len(failed_sids)


def generate_revisi_loan_excel(df_buy, sid_results, op_data, cl_data):
    today_str   = datetime.today().strftime("%Y%m%d")
    failed_sids = [
        sid for sid, data in sid_results.items()
        if data.get("has_loan_request") and not lolos_loan(data)
    ]
    alasan_rows = []
    for sid in failed_sids:
        data     = sid_results[sid]
        op       = op_data.get(sid, {})
        cl       = cl_data.get(sid, {})
        max_loan = data.get("max_loan_recommendation", 0)
        max_63   = data.get("max_loan_63", 0)
        for check in data["checks"]:
            if check["label"].startswith(("1a.", "1a-OP.", "1c.", "2a.", "2b.", "3.")) and not check["passed"]:
                alasan_rows.append({
                    "SID": sid, "Nama": data["name"],
                    "Check Gagal": check["label"], "Detail Alasan": check["detail"],
                    "Loan Existing": op.get("loan_existing", "-"),
                    "Accrued Interest": op.get("accrued_interest", "-"),
                    "Available Limit": cl.get("available_limit", "-"),
                    "Collateral": data.get("collateral_existing", "-"),
                    "LR Belum Settled": data.get("lr_value", 0),
                    "RP Belum Settled": data.get("rp_value", 0),
                    "Max Loan Rekomendasi (65%)": max_loan if max_loan > 0 else "Tidak bisa",
                    "Max Loan Aman (63%)": max_63 if max_63 > 0 else "Tidak bisa",
                })
    df_failed_buy = df_buy[col(df_buy, BUY_SID).astype(str).isin([str(s) for s in failed_sids])].copy()
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(alasan_rows).to_excel(writer, sheet_name="Alasan Gagal", index=False)
        df_failed_buy.to_excel(writer,             sheet_name="Buy (Loan)",   index=False)
    buf.seek(0)
    return buf, f"Revisi Loan Request {today_str}.xlsx", len(failed_sids)

# ─────────────────────────────────────────────
# CREDIT LIMIT PARTISIPAN
# ─────────────────────────────────────────────

CREDIT_LIMIT_PARTISIPAN = 148_000_000_000.0

# ─────────────────────────────────────────────
# UI — UPLOAD
# ─────────────────────────────────────────────

st.subheader("📂 Upload File")
col1, col2, col3 = st.columns(3)

with col1:
    hasil_file = st.file_uploader("1. Hasil_MNC (xlsx)", type=["xlsx", "xls"], key="hasil")
with col2:
    op_file = st.file_uploader("2. File OP (.txt)", type=["txt"], key="op")
with col3:
    cl_file = st.file_uploader("3. Credit Limit (.txt)", type=["txt"], key="cl")

col4, col5, col6 = st.columns(3)
with col4:
    cp_file = st.file_uploader("4. Closing Price (.xlsx)", type=["xlsx", "xls"], key="cp")
with col5:
    rp_file = st.file_uploader("5. File RiskParameter (.txt)", type=["txt"], key="rp_param")
with col6:
    lr_file = st.file_uploader("6. LR Belum Settled (.xlsx)", type=["xlsx", "xls"], key="lr",
                                help="Format: SID | Name | Loan Value")

col7, col8, col9 = st.columns(3)
with col7:
    rp_txn_file = st.file_uploader("7. RP Belum Settled (.xlsx)", type=["xlsx", "xls"], key="rp_txn",
                                    help="Format: SID | Name | Repayment Value")

# Preview setelah upload
if op_file or cl_file:
    prev_col1, prev_col2 = st.columns(2)

    if op_file:
        with prev_col1:
            st.subheader("👁 Preview OP File")
            try:
                op_content_prev = op_file.read().decode("utf-8", errors="replace")
                op_file.seek(0)
                op_prev     = parse_op_file(op_content_prev)
                sample_sids = list(op_prev.keys())[:3]
                for s in sample_sids:
                    d = op_prev[s]
                    with st.container(border=True):
                        st.markdown(f"**SID: {s}** — {d['name']}")
                        if d['loan_existing'] <= 0:
                            st.warning("⚠️ Loan Existing = 0 → Repayment akan dilewati otomatis")
                        st.caption(
                            f"Loan: {fmt_rp(d['loan_existing'])} | Accrued: {fmt_rp(d['accrued_interest'])} | "
                            f"Vol Existing: {d['volume_existing']:,.0f} lot | Saham: {len(d['stocks'])} kode"
                        )
                st.caption(f"Total {len(op_prev)} nasabah terbaca dari OP file.")
            except Exception as ex:
                st.error(f"Gagal preview OP: {ex}")

    if cl_file:
        with prev_col2:
            st.subheader("👁 Preview Credit Limit File")
            try:
                cl_content_prev = cl_file.read().decode("utf-8", errors="replace")
                cl_file.seek(0)
                cl_prev, vdate  = parse_credit_limit_file(cl_content_prev)
                today_str_check = datetime.today().strftime("%Y/%m/%d")
                if vdate and vdate != today_str_check:
                    st.warning(f"⚠️ Value Date file: **{vdate}** — berbeda dengan hari ini ({today_str_check}).")
                sample_sids = list(cl_prev.keys())[:3]
                for s in sample_sids:
                    d = cl_prev[s]
                    with st.container(border=True):
                        st.markdown(f"**SID: {s}** — {d['name']}")
                        st.caption(f"Available Limit: {fmt_rp(d['available_limit'])}")
                st.caption(f"Total {len(cl_prev)} nasabah terbaca dari Credit Limit file.")
            except Exception as ex:
                st.error(f"Gagal preview Credit Limit: {ex}")

    st.divider()

st.subheader("Credit Limit Partisipan")
st.info(f"Credit Limit Partisipan ditetapkan: **{fmt_rp(CREDIT_LIMIT_PARTISIPAN)}**")
st.divider()

with st.expander("📖 Panduan Validasi — Formula Rasio"):
    st.markdown("""
    ### Formula Rasio

    **Collateral Value (CV)**  
    `CV = Σ (Qty Emiten Existing × Closing Price × (1 - Haircut%))`

    **Rasio Repayment (1c)**  
    `Rasio = (Loan Existing + Accrued + LR belum settled − (RP belum settled + Repayment diajukan)) / CV`

    **Rasio Loan Request (2b)**  
    `Rasio = (Loan Existing + Accrued + (LR belum settled + Loan diajukan) − (RP belum settled + Repayment diajukan)) / CV`

    **Auto-Adjust Target: 63%** (buffer 2% dari threshold 65%)

    ---

    | # | Cek | Keterangan |
    |---|-----|------------|
    | — | **Guard: Existing Loan = 0** | Repayment check dilewati |
    | 1a | Volume Sell ≤ Available Sell Qty | Per baris Sell sheet |
    | 1a-OP | Saham Sell ada & Vol ≤ Outstanding di OP | Cross-check ke OP file |
    | 1b | Total Repayment Value ≤ Total Loan Value | Hanya jika ada Loan Request |
    | 1c | Rasio Repayment < 65% | |
    | 2a | Volume Buy ≤ Available Quantity | Avail=0 → dibatalkan |
    | 2b | Rasio Loan Request < 65% | Auto-adjust ke 63% tersedia |
    | 3  | Credit Limit Nasabah | (Avail Limit + Sell Value) > Loan Diajukan |
    | 4  | Credit Limit Partisipan | Global check |
    """)

with st.expander("📖 Panduan Struktur File"):
    st.markdown("""
    | # | File | Format |
    |---|------|--------|
    | 1 | Hasil_MNC | Sheet Sell & Buy |
    | 2 | OP File | .txt pipe-delimited |
    | 3 | Credit Limit | .txt pipe-delimited |
    | 4 | Closing Price | .xlsx kolom no_share & kurs_now |
    | 5 | RiskParameter | .txt pipe-delimited |
    | 6 | LR Belum Settled | .xlsx: SID \| Name \| Loan Value |
    | 7 | RP Belum Settled | .xlsx: SID \| Name \| Repayment Value |

    > File 6 & 7 bisa digunakan dari rekap yang digenerate sistem hari sebelumnya.
    """)

run_btn = st.button("▶ Jalankan Validasi", use_container_width=True, type="primary")

# ─────────────────────────────────────────────
# RUN VALIDASI
# ─────────────────────────────────────────────

if run_btn:
    errors = []
    if not hasil_file: errors.append("File Hasil_MNC belum diupload.")
    if not op_file:    errors.append("File OP belum diupload.")
    if not cl_file:    errors.append("File Credit Limit belum diupload.")
    if not cp_file:    errors.append("File Closing Price belum diupload.")
    if not rp_file:    errors.append("File RiskParameter belum diupload.")

    if errors:
        for e in errors:
            st.error(e)
        st.stop()

    with st.spinner("⚙️ Memproses seluruh data..."):
        try:
            df_sell, df_buy = load_hasil_mnc(hasil_file)
        except Exception as ex:
            st.error(f"Gagal membaca Hasil_MNC: {ex}"); st.stop()

        try:
            op_content = op_file.read().decode("utf-8", errors="replace")
            op_data    = parse_op_file(op_content)
        except Exception as ex:
            st.error(f"Gagal membaca OP file: {ex}"); st.stop()

        try:
            cl_content = cl_file.read().decode("utf-8", errors="replace")
            cl_data, _ = parse_credit_limit_file(cl_content)
        except Exception as ex:
            st.error(f"Gagal membaca Credit Limit file: {ex}"); st.stop()

        try:
            closing_prices = load_closing_price(cp_file)
        except Exception as ex:
            st.error(f"Gagal membaca Closing Price: {ex}"); st.stop()

        try:
            risk_params = load_risk_parameter(rp_file)
        except Exception as ex:
            st.error(f"Gagal membaca RiskParameter: {ex}"); st.stop()

        # LR belum settled — opsional
        lr_data  = {}
        lr_names = {}
        if lr_file is not None:
            try:
                lr_data, lr_names = parse_lr_excel(lr_file)
                st.info(f"ℹ️ LR belum settled: {len(lr_data)} nasabah, "
                        f"total {fmt_rp(sum(lr_data.values()))}")
            except Exception as ex:
                st.error(f"Gagal membaca file LR Excel: {ex}"); st.stop()
        else:
            st.warning("⚠️ File LR Belum Settled tidak diupload — LR dianggap 0.")

        # RP belum settled — opsional
        rp_data  = {}
        rp_names = {}
        if rp_txn_file is not None:
            try:
                rp_data, rp_names = parse_rp_excel(rp_txn_file)
                st.info(f"ℹ️ RP belum settled: {len(rp_data)} nasabah, "
                        f"total {fmt_rp(sum(rp_data.values()))}")
            except Exception as ex:
                st.error(f"Gagal membaca file RP Excel: {ex}"); st.stop()
        else:
            st.warning("⚠️ File RP Belum Settled tidak diupload — RP dianggap 0.")

        sid_results, global_result = run_validations(
            df_sell, df_buy, op_data, cl_data, CREDIT_LIMIT_PARTISIPAN,
            closing_prices, risk_params, lr_data, rp_data
        )

    # Simpan ke session state
    st.session_state['df_sell_edited']   = df_sell.copy()
    st.session_state['df_buy']           = df_buy.copy()
    st.session_state['df_buy_adjusted']  = df_buy.copy()
    st.session_state['sid_results']      = sid_results
    st.session_state['global_result']    = global_result
    st.session_state['op_data']          = op_data
    st.session_state['cl_data']          = cl_data
    st.session_state['closing_prices']   = closing_prices
    st.session_state['risk_params']      = risk_params
    st.session_state['lr_data']          = lr_data
    st.session_state['lr_names']         = lr_names
    st.session_state['rp_data']          = rp_data
    st.session_state['rp_names']         = rp_names
    st.session_state['clamped_warnings'] = []

    st.success("✅ Validasi Selesai!")

# ─────────────────────────────────────────────
# TAMPILKAN HASIL
# ─────────────────────────────────────────────

if st.session_state.get('sid_results') is not None:

    sid_results    = st.session_state['sid_results']
    global_result  = st.session_state['global_result']
    op_data        = st.session_state.get('op_data', {})
    cl_data        = st.session_state.get('cl_data', {})
    df_buy         = st.session_state.get('df_buy')
    closing_prices = st.session_state.get('closing_prices', {})
    risk_params    = st.session_state.get('risk_params', {})
    lr_data        = st.session_state.get('lr_data', {})
    rp_data        = st.session_state.get('rp_data', {})

    total_sids         = len(sid_results)
    total_pass_rep     = sum(1 for v in sid_results.values() if lolos_repayment(v))
    total_pass_loan    = sum(1 for v in sid_results.values() if lolos_loan(v))
    total_fail_rep     = sum(1 for v in sid_results.values()
                             if v.get("has_repayment") and not v.get("repayment_skipped_no_loan")
                             and not lolos_repayment(v))
    total_fail_loan    = sum(1 for v in sid_results.values()
                             if v.get("has_loan_request") and not lolos_loan(v))
    total_skip_no_loan = sum(1 for v in sid_results.values() if v.get("repayment_skipped_no_loan"))
    global_pass        = global_result["passed"]

    m1, m2, m3, m4, m5, m6, m7 = st.columns(7)
    m1.metric("Total Nasabah",   total_sids)
    m2.metric("Lolos Repayment", total_pass_rep)
    m3.metric("Lolos Loan",      total_pass_loan)
    m4.metric("Gagal Repayment", total_fail_rep,
              delta=f"-{total_fail_rep}" if total_fail_rep else None, delta_color="inverse")
    m5.metric("Gagal Loan",      total_fail_loan,
              delta=f"-{total_fail_loan}" if total_fail_loan else None, delta_color="inverse")
    m6.metric("Skip (Loan=0)",   total_skip_no_loan)
    m7.metric("CL Partisipan",   "✅ LOLOS" if global_pass else "❌ GAGAL")

    st.divider()

    tab_global, tab_per_sid, tab_gagal, tab_autoadjust, tab_export = st.tabs([
        "🌐 Validasi 4 (Global)",
        "👤 Validasi Per Nasabah",
        "❌ Nasabah Gagal",
        "⚡ Auto-Adjust Rasio",
        "📥 Export Hasil",
    ])

    with tab_global:
        st.subheader("Validasi 4 — Credit Limit Partisipan (Global)")
        if global_result["passed"]:
            st.success(f"✅ LOLOS — {global_result['detail']}")
        else:
            st.error(f"❌ GAGAL — {global_result['detail']}")

    with tab_per_sid:
        st.caption("🟢 R = Lolos Repayment  |  🟢 L = Lolos Loan  |  🔴 = Gagal  |  ⏭ = Skip (Loan=0)")
        for sid, data in sid_results.items():
            r_pass  = lolos_repayment(data)
            l_pass  = lolos_loan(data)
            skipped = data.get("repayment_skipped_no_loan")
            r_icon  = "⏭ R (Loan=0)" if skipped else ("✅ R" if r_pass else "❌ R")
            l_icon  = "✅ L" if l_pass else "❌ L"
            expanded = not (r_pass and l_pass) and not skipped

            with st.expander(f"{r_icon} | {l_icon} | {sid} — {data['name']}", expanded=expanded):
                if skipped:
                    st.info("ℹ️ Existing Loan = 0 → Repayment tidak diproses.")
                coll = data.get("collateral_existing", 0)
                lr_v = data.get("lr_value", 0)
                rp_v = data.get("rp_value", 0)
                m63  = data.get("max_loan_63", 0)
                st.caption(
                    f"📊 Collateral: {fmt_rp(coll)} | LR belum settled: {fmt_rp(lr_v)} | "
                    f"RP belum settled: {fmt_rp(rp_v)} | Max Loan Aman 63%: {fmt_rp(m63)}"
                )
                max_loan = data.get("max_loan_recommendation", 0)
                if max_loan > 0 and not l_pass:
                    st.warning(f"💡 Max Loan (65%): **{fmt_rp(max_loan)}** | Max Loan Aman (63%): **{fmt_rp(m63)}**")
                elif max_loan == 0 and data.get("has_loan_request") and not l_pass:
                    st.error("🚫 Collateral tidak mencukupi — tidak bisa ajukan loan.")
                for check in data["checks"]:
                    if check["passed"]:
                        st.success(f"✅ **{check['label']}**  {check['detail']}")
                    else:
                        st.error(f"❌ **{check['label']}**  {check['detail']}")

    with tab_gagal:
        st.subheader("Nasabah yang Tidak Lolos Validasi")

        gagal_rep  = [(sid, d) for sid, d in sid_results.items()
                      if d.get("has_repayment") and not d.get("repayment_skipped_no_loan")
                      and not lolos_repayment(d)]
        gagal_loan = [(sid, d) for sid, d in sid_results.items()
                      if d.get("has_loan_request") and not lolos_loan(d)]
        skip_loan0 = [(sid, d) for sid, d in sid_results.items()
                      if d.get("repayment_skipped_no_loan")]

        gcol1, gcol2 = st.columns(2)

        with gcol1:
            st.markdown(f"#### 🔴 Gagal Repayment — {len(gagal_rep)} nasabah")
            if skip_loan0:
                with st.expander(f"⏭ Dilewati (Loan=0) — {len(skip_loan0)} nasabah"):
                    for sid, data in skip_loan0:
                        st.info(f"**{sid}** — {data['name']} | Loan Existing = 0")
            if not gagal_rep:
                st.success("Semua nasabah lolos Repayment.")
            else:
                for sid, data in gagal_rep:
                    failed_checks = [c for c in data["checks"]
                                     if c["label"].startswith(("1a.", "1a-OP.", "1b.", "1c."))
                                     and not c["passed"]]
                    with st.expander(f"❌ {sid} — {data['name']}"):
                        for c in failed_checks:
                            st.error(f"**{c['label']}** — {c['detail']}")

        with gcol2:
            st.markdown(f"#### 🔴 Gagal Loan Request — {len(gagal_loan)} nasabah")
            if not gagal_loan:
                st.success("Semua nasabah lolos Loan Request.")
            else:
                for sid, data in gagal_loan:
                    max_63 = data.get("max_loan_63", 0)
                    failed_checks = [c for c in data["checks"]
                                     if c["label"].startswith(("1a.", "1a-OP.", "1c.", "2a.", "2b.", "3."))
                                     and not c["passed"]]
                    with st.expander(f"❌ {sid} — {data['name']}"):
                        if max_63 > 0:
                            st.warning(f"💡 Max Loan Aman (63%): **{fmt_rp(max_63)}**")
                        else:
                            st.error("🚫 Tidak bisa ajukan loan — collateral tidak cukup.")
                        for c in failed_checks:
                            st.error(f"**{c['label']}** — {c['detail']}")

        # ── EDITOR 1b ─────────────────────────────────────────────
        st.divider()
        st.markdown("#### ✏️ Revisi Volume Sell — Nasabah Gagal 1b")

        if st.session_state.get('clamped_warnings'):
            st.warning("⚠️ Volume dipotong ke AVQ: " + " | ".join(st.session_state['clamped_warnings']))
            st.session_state['clamped_warnings'] = []

        _sid_results = st.session_state.get('sid_results') or sid_results
        gagal_1b = [
            (sid, d) for sid, d in _sid_results.items()
            if any(c["label"].startswith("1b.") and not c["passed"] for c in d["checks"])
        ]

        if not gagal_1b:
            st.success("Tidak ada nasabah yang gagal validasi 1b.")
        else:
            df_sell_edit = st.session_state.get('df_sell_edited')
            if df_sell_edit is not None:
                edit_rows = []
                for sid, data in gagal_1b:
                    rows = df_sell_edit[col(df_sell_edit, SELL_SID).astype(str) == sid]
                    for idx, row in rows.iterrows():
                        price = pd.to_numeric(row.iloc[SELL_CP], errors='coerce') or 0
                        vol   = abs(pd.to_numeric(row.iloc[SELL_VOL], errors='coerce') or 0)
                        avq   = pd.to_numeric(row.iloc[SELL_AVQ], errors='coerce') or 0
                        edit_rows.append({
                            '_df_index': idx, '_price': price, '_avq': avq,
                            'SID': sid, 'Nama': data['name'],
                            'Stock': str(row.iloc[SELL_STOCK]),
                            'AVQ (Maks)': int(avq),
                            'Volume Sell (editable)': int(vol),
                            'Closing Price': price,
                            'Value': vol * price,
                        })

                df_editor_input = pd.DataFrame(edit_rows)
                st.caption("⚠️ Ubah kolom **Volume Sell (editable)**. "
                           "Volume otomatis dibatasi maksimal **AVQ (Maks)** saat tombol Terapkan ditekan.")

                edited = st.data_editor(
                    df_editor_input.drop(columns=['_df_index', '_price', '_avq']),
                    column_config={
                        "SID":        st.column_config.TextColumn(disabled=True),
                        "Nama":       st.column_config.TextColumn(disabled=True),
                        "Stock":      st.column_config.TextColumn(disabled=True),
                        "AVQ (Maks)": st.column_config.NumberColumn(disabled=True),
                        "Volume Sell (editable)": st.column_config.NumberColumn(min_value=0, step=100),
                        "Closing Price": st.column_config.NumberColumn(disabled=True),
                        "Value":      st.column_config.NumberColumn(disabled=True),
                    },
                    use_container_width=True, key="editor_1b", hide_index=True,
                )

                if st.button("✅ Terapkan Revisi & Validasi Ulang", type="primary"):
                    df_sell_updated  = st.session_state['df_sell_edited'].copy()
                    clamped_warnings = []
                    for i, edit_row in edited.iterrows():
                        orig_idx = df_editor_input.iloc[i]['_df_index']
                        avq      = df_editor_input.iloc[i]['_avq']
                        price    = df_editor_input.iloc[i]['_price']
                        new_vol  = abs(int(edit_row['Volume Sell (editable)'] or 0))
                        if new_vol > int(avq):
                            clamped_warnings.append(
                                f"{df_editor_input.iloc[i]['Stock']} ({df_editor_input.iloc[i]['SID']}): "
                                f"{new_vol:,} → {int(avq):,}")
                            new_vol = int(avq)
                        new_val = new_vol * price
                        df_sell_updated.at[orig_idx, df_sell_updated.columns[SELL_VOL]] = new_vol
                        df_sell_updated.at[orig_idx, df_sell_updated.columns[SELL_VAL]] = new_val

                    st.session_state['df_sell_edited']   = df_sell_updated
                    st.session_state['clamped_warnings'] = clamped_warnings

                    new_sid_results, new_global_result = run_validations(
                        df_sell_updated,
                        st.session_state['df_buy'],
                        st.session_state['op_data'],
                        st.session_state['cl_data'],
                        CREDIT_LIMIT_PARTISIPAN,
                        st.session_state['closing_prices'],
                        st.session_state['risk_params'],
                        st.session_state['lr_data'],
                        st.session_state['rp_data'],
                    )
                    st.session_state['sid_results']   = new_sid_results
                    st.session_state['global_result'] = new_global_result
                    st.rerun()

    # ── TAB AUTO-ADJUST ────────────────────────────────────────
    with tab_autoadjust:
        st.subheader("⚡ Auto-Adjust Loan Request — Target Rasio 63%")
        st.info(
            "Sistem memotong volume buy **proporsional per saham** agar rasio ≤ 63%. "
            "Nasabah yang collateral-nya tidak mencukupi dikeluarkan dari loan request."
        )

        gagal_2b = [
            (sid, d) for sid, d in sid_results.items()
            if d.get("has_loan_request") and any(
                c["label"].startswith(("2b.", "3.")) and not c["passed"]
                for c in d["checks"]
            )
        ]

        if not gagal_2b:
            st.success("✅ Tidak ada nasabah yang perlu di-adjust rasionya.")
        else:
            preview_rows = []
            for sid, data in gagal_2b:
                max_63     = data.get("max_loan_63", 0)
                total_buy  = data.get("total_buy_val", 0)
                collateral = data.get("collateral_existing", 0)
                loan_exist = data.get("loan_existing", 0)
                accrued    = data.get("accrued_interest", 0)
                lr_v       = data.get("lr_value", 0)
                rp_v       = data.get("rp_value", 0)
                sell_val   = data.get("total_sell_val", 0)

                effective_cl   = cl_data.get(sid, {}).get('available_limit', 0) + sell_val
                max_loan_final = min(max_63, effective_cl) if max_63 > 0 else 0

                if max_loan_final <= 0:
                    status    = "❌ Dikeluarkan (collateral/CL tidak cukup)"
                    adj_ratio, _ = calc_ratio_baru(
                        loan_exist, accrued, lr_v, rp_v + sell_val, collateral)
                else:
                    status    = f"✂️ Dipotong ke {fmt_rp(max_loan_final)}"
                    adj_ratio, _ = calc_ratio_baru(
                        loan_exist, accrued, lr_v + max_loan_final, rp_v + sell_val, collateral)

                preview_rows.append({
                    'SID':                  sid,
                    'Nama':                 data['name'],
                    'Loan Diajukan':        fmt_rp(total_buy),
                    'Max Loan (63%)':       fmt_rp(max_63) if max_63 > 0 else "Rp 0",
                    'Effective CL':         fmt_rp(effective_cl),
                    'Max Loan Final':       fmt_rp(max_loan_final) if max_loan_final > 0 else "Rp 0",
                    'Rasio Setelah':        fmt_pct(adj_ratio) if adj_ratio else "N/A",
                    'Status':               status,
                })

            st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)
            n_dikeluarkan = sum(1 for r in preview_rows if "Dikeluarkan" in r['Status'])
            n_dipotong    = sum(1 for r in preview_rows if "Dipotong" in r['Status'])
            st.caption(f"**{n_dipotong}** nasabah dipotong · **{n_dikeluarkan}** nasabah dikeluarkan")

            if st.button("⚡ Terapkan Auto-Adjust & Validasi Ulang", type="primary",
                         use_container_width=True):
                df_buy_updated = st.session_state['df_buy_adjusted'].copy()
                for sid, data in gagal_2b:
                    max_63       = data.get("max_loan_63", 0)
                    sell_val     = data.get("total_sell_val", 0)
                    effective_cl = cl_data.get(sid, {}).get('available_limit', 0) + sell_val
                    max_loan_final = min(max_63, effective_cl) if max_63 > 0 else 0
                    df_buy_updated = auto_adjust_loan(
                        df_buy_updated, sid, max_loan_final, st.session_state['closing_prices'])

                st.session_state['df_buy_adjusted'] = df_buy_updated
                st.session_state['df_buy']          = df_buy_updated

                new_sid_results, new_global_result = run_validations(
                    st.session_state['df_sell_edited'],
                    df_buy_updated,
                    st.session_state['op_data'],
                    st.session_state['cl_data'],
                    CREDIT_LIMIT_PARTISIPAN,
                    st.session_state['closing_prices'],
                    st.session_state['risk_params'],
                    st.session_state['lr_data'],
                    st.session_state['rp_data'],
                )
                st.session_state['sid_results']   = new_sid_results
                st.session_state['global_result'] = new_global_result
                st.success("✅ Auto-Adjust diterapkan!")
                st.rerun()

    # ── TAB EXPORT ─────────────────────────────────────────────
    with tab_export:
        st.subheader("📋 Ringkasan Hasil Validasi")
        summary_rows = []
        for sid, data in sid_results.items():
            row = {
                "SID": sid, "Nama": data["name"],
                "Loan Existing": data.get("loan_existing", 0),
                "Collateral Existing": data.get("collateral_existing", 0),
                "LR Belum Settled": data.get("lr_value", 0),
                "RP Belum Settled": data.get("rp_value", 0),
                "Max Loan Aman (63%)": data.get("max_loan_63", 0),
                "Skip Repayment (Loan=0)": "YA" if data.get("repayment_skipped_no_loan") else "TIDAK",
                "Max Loan Rekomendasi": data.get("max_loan_recommendation", 0),
            }
            for check in data["checks"]:
                row[check["label"]] = "LOLOS" if check["passed"] else "GAGAL"
            row["Status Repayment"] = (
                "SKIP (Loan=0)" if data.get("repayment_skipped_no_loan")
                else ("LOLOS" if lolos_repayment(data) else "GAGAL")
            )
            row["Status Loan"] = "LOLOS" if lolos_loan(data) else "GAGAL"
            summary_rows.append(row)

        df_summary = pd.DataFrame(summary_rows)
        st.dataframe(df_summary, use_container_width=True)

        buf_sum = io.BytesIO()
        df_summary.to_excel(buf_sum, index=False)
        buf_sum.seek(0)
        st.download_button("⬇️ Download Hasil Validasi (.xlsx)", data=buf_sum,
                           file_name="hasil_validasi.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.divider()
        st.subheader("📤 Export ke Dashboard Kantor")

        dl_col1, dl_col2 = st.columns(2)
        with dl_col1:
            df_sell_final = st.session_state.get('df_sell_edited')
            if df_sell_final is not None:
                rep_buf, rep_fname = generate_repayment_excel(df_sell_final, sid_results)
                st.download_button("⬇️ Repayment Proceed (.xlsx)", data=rep_buf,
                                   file_name=rep_fname,
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)
        with dl_col2:
            df_buy_final = st.session_state.get('df_buy')
            if df_buy_final is not None:
                loan_buf, loan_fname = generate_loan_excel(df_buy_final, sid_results)
                st.download_button("⬇️ Loan Request (.xlsx)", data=loan_buf,
                                   file_name=loan_fname,
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)

        st.divider()
        st.subheader("📅 Rekap Belum Settled — Untuk Besok")
        st.caption("Upload file ini besok sebagai input LR/RP Belum Settled.")

        rek_col1, rek_col2 = st.columns(2)
        with rek_col1:
            df_buy_final = st.session_state.get('df_buy')
            if df_buy_final is not None:
                lr_rek_buf, lr_rek_fname = generate_lr_rekap_excel(df_buy_final, sid_results)
                st.download_button("⬇️ Rekap LR Belum Settled (.xlsx)", data=lr_rek_buf,
                                   file_name=lr_rek_fname,
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True, type="primary")
        with rek_col2:
            df_sell_final = st.session_state.get('df_sell_edited')
            if df_sell_final is not None:
                rp_rek_buf, rp_rek_fname = generate_rp_rekap_excel(df_sell_final, sid_results)
                st.download_button("⬇️ Rekap RP Belum Settled (.xlsx)", data=rp_rek_buf,
                                   file_name=rp_rek_fname,
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True, type="primary")

        st.divider()
        st.subheader("📝 Export File Revisi (Nasabah Gagal)")

        rv_col1, rv_col2 = st.columns(2)
        with rv_col1:
            df_sell_final = st.session_state.get('df_sell_edited')
            if df_sell_final is not None:
                rev_rep_buf, rev_rep_fname, n_gagal_rep = generate_revisi_repayment_excel(
                    df_sell_final, sid_results, op_data)
                if n_gagal_rep == 0:
                    st.info("✅ Tidak ada nasabah yang gagal Repayment.")
                else:
                    st.download_button(f"⬇️ Revisi Repayment — {n_gagal_rep} nasabah gagal",
                                       data=rev_rep_buf, file_name=rev_rep_fname,
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       use_container_width=True, type="secondary")

        with rv_col2:
            df_buy_final = st.session_state.get('df_buy')
            if df_buy_final is not None:
                rev_loan_buf, rev_loan_fname, n_gagal_loan = generate_revisi_loan_excel(
                    df_buy_final, sid_results, op_data, cl_data)
                if n_gagal_loan == 0:
                    st.info("✅ Tidak ada nasabah yang gagal Loan Request.")
                else:
                    st.download_button(f"⬇️ Revisi Loan Request — {n_gagal_loan} nasabah gagal",
                                       data=rev_loan_buf, file_name=rev_loan_fname,
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       use_container_width=True, type="secondary")
