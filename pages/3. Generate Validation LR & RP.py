import streamlit as st
import pandas as pd
import io
from datetime import datetime

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
    lines = content.strip().splitlines()
    for line in lines:
        line = line.strip()
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
                "stocks": {},
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
                result[sid]["volume_existing"] += vol
                result[sid]["stocks"][stock]    = result[sid]["stocks"].get(stock, 0.0) + vol
    return result


def parse_credit_limit_file(content: str):
    result     = {}
    value_date = None
    lines      = content.strip().splitlines()

    for i, line in enumerate(lines):
        line = line.strip()
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


def load_hasil_mnc(uploaded_file):
    xls     = pd.ExcelFile(uploaded_file)
    df_sell = pd.read_excel(xls, sheet_name="Sell (Repayment)", header=0)
    df_buy  = pd.read_excel(xls, sheet_name="Buy (Loan)",       header=0)
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

# ─────────────────────────────────────────────
# COLLATERAL CALCULATOR
# ─────────────────────────────────────────────

def calc_collateral_from_rows(rows, qty_col_idx, cp_col_idx, hc_col_idx):
    total  = 0.0
    detail = []
    for _, row in rows.iterrows():
        qty  = abs(pd.to_numeric(row.iloc[qty_col_idx], errors="coerce") or 0)
        cp   = pd.to_numeric(row.iloc[cp_col_idx], errors="coerce") or 0
        hc   = parse_hc(row.iloc[hc_col_idx])
        coll = qty * cp * (1 - hc)
        total += coll
        detail.append({
            "stock":      str(row.iloc[BUY_STOCK]),
            "qty":        qty,
            "cp":         cp,
            "hc":         hc,
            "collateral": coll,
        })
    return total, detail

# ─────────────────────────────────────────────
# VALIDATION LOGIC
# ─────────────────────────────────────────────

RATIO_THRESHOLD = 0.65

def run_validations(df_sell, df_buy, op_data, cl_data, credit_limit_partisipan):
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
        volume_existing  = op["volume_existing"]
        available_limit  = cl["available_limit"]
        name             = op.get("name") or cl.get("name") or sid

        total_sell_vol = sell["total_volume"]
        total_sell_val = sell["total_value"]
        total_buy_vol  = buy["total_volume"]
        total_buy_val  = buy["total_value"]

        has_repayment    = total_sell_vol > 0
        has_loan_request = total_buy_vol  > 0

        sid_results = {
            "name":             name,
            "checks":           [],
            "has_repayment":    has_repayment,
            "has_loan_request": has_loan_request,
        }

        def add(label, passed, detail=""):
            sid_results["checks"].append({
                "label":  label,
                "passed": passed,
                "detail": detail,
            })

        buy_collateral_total, buy_coll_detail = calc_collateral_from_rows(
            buy["rows"], BUY_AVQ, BUY_CP, BUY_HC
        )

        # ── 1a. Volume Sell ≤ Available Sell Quantity (dari Hasil_MNC) ──
        rep_1a_pass   = True
        rep_1a_detail = []
        for _, row in sell["rows"].iterrows():
            vol = abs(pd.to_numeric(row.iloc[SELL_VOL], errors="coerce") or 0)
            avq = pd.to_numeric(row.iloc[SELL_AVQ], errors="coerce") or 0
            stk = str(row.iloc[SELL_STOCK])
            if vol > avq:
                rep_1a_pass = False
                rep_1a_detail.append(f"{stk}: Vol {vol:,.0f} > Avail {avq:,.0f}")
        add(
            "1a. Volume Sell ≤ Available Sell Quantity",
            rep_1a_pass,
            "; ".join(rep_1a_detail) if rep_1a_detail else f"Total Volume Sell: {total_sell_vol:,.0f}"
        )

        # ── 1a-OP. Saham Sell ada & Volume ≤ Outstanding di OP File ──────
        if not has_repayment:
            add(
                "1a-OP. Saham Sell Terverifikasi di OP File",
                True,
                "Tidak ada Repayment — check 1a-OP dilewati"
            )
        elif not op.get("stocks"):
            # Nasabah tidak ditemukan di OP file sama sekali
            add(
                "1a-OP. Saham Sell Terverifikasi di OP File",
                False,
                f"SID {sid} tidak ditemukan di OP file — tidak ada posisi saham outstanding"
            )
        else:
            op_stocks        = op["stocks"]   # {"BBCA": 500.0, ...}
            op_1a_pass       = True
            op_1a_detail     = []
            op_1a_ok_detail  = []

            for _, row in sell["rows"].iterrows():
                stk = str(row.iloc[SELL_STOCK]).strip()
                vol = abs(pd.to_numeric(row.iloc[SELL_VOL], errors="coerce") or 0)

                if stk not in op_stocks:
                    op_1a_pass = False
                    op_1a_detail.append(f"{stk}: tidak ada di OP file")
                elif vol > op_stocks[stk]:
                    op_1a_pass = False
                    op_1a_detail.append(
                        f"{stk}: Vol Sell {vol:,.0f} > Outstanding OP {op_stocks[stk]:,.0f}"
                    )
                else:
                    op_1a_ok_detail.append(
                        f"{stk}: Vol Sell {vol:,.0f} ✓ (OP: {op_stocks[stk]:,.0f})"
                    )

            if op_1a_pass:
                detail_msg = "; ".join(op_1a_ok_detail) if op_1a_ok_detail else "Semua saham terverifikasi di OP file"
            else:
                detail_msg = "; ".join(op_1a_detail)
                if op_1a_ok_detail:
                    detail_msg += " || OK: " + "; ".join(op_1a_ok_detail)

            add(
                "1a-OP. Saham Sell Terverifikasi di OP File",
                op_1a_pass,
                detail_msg
            )

        # ── 1b. Total Repayment Value ≤ Total Loan Value ─────────────────
        if not has_loan_request:
            rep_1b_pass   = True
            rep_1b_detail = "Pure repayment tanpa Loan Request — check 1b dilewati"
        else:
            rep_1b_pass   = total_sell_val <= total_buy_val
            rep_1b_detail = (
                f"Total Sell Value: {fmt_rp(total_sell_val)} | "
                f"Total Buy Value (Loan): {fmt_rp(total_buy_val)}"
            )
        add("1b. Total Repayment Value ≤ Total Loan Value", rep_1b_pass, rep_1b_detail)

        # ── 1c. Rasio Repayment < 65% ─────────────────────────────────────
        if not has_repayment:
            rep_1c_pass   = True
            rep_1c_detail = "Tidak ada Repayment — check 1c dilewati"
        else:
            loan_num_repayment = loan_existing - total_sell_val + accrued_interest
            if buy_collateral_total > 0:
                collateral_repayment = buy_collateral_total
            else:
                sell_cp_vals = pd.to_numeric(col(sell["rows"], SELL_CP), errors="coerce").dropna()
                avg_sell_cp  = sell_cp_vals.mean() if len(sell_cp_vals) > 0 else 0
                sisa_vol     = max(volume_existing - total_sell_vol, 0)
                collateral_repayment = sisa_vol * avg_sell_cp

            if collateral_repayment > 0:
                ratio_rep     = loan_num_repayment / collateral_repayment
                rep_1c_pass   = ratio_rep < RATIO_THRESHOLD
                rep_1c_detail = (
                    f"Rasio: {fmt_pct(ratio_rep)} (threshold <{RATIO_THRESHOLD*100:.0f}%) | "
                    f"Numerator: {fmt_rp(loan_num_repayment)} | "
                    f"Collateral: {fmt_rp(collateral_repayment)}"
                )
            elif loan_num_repayment <= 0:
                rep_1c_pass   = True
                rep_1c_detail = f"Loan Numerator ≤ 0 ({fmt_rp(loan_num_repayment)}) — posisi lunas"
            else:
                rep_1c_pass   = False
                rep_1c_detail = "Collateral = 0 sementara Loan masih ada — Rasio tidak terhitung (∞)"
        add("1c. Rasio Repayment < 65%", rep_1c_pass, rep_1c_detail)

        # ── 2a. Volume Buy ≤ Available Quantity ───────────────────────────
        loan_2a_pass   = True
        loan_2a_detail = []
        for _, row in buy["rows"].iterrows():
            vol = pd.to_numeric(row.iloc[BUY_VOL], errors="coerce") or 0
            avq = pd.to_numeric(row.iloc[BUY_AVQ], errors="coerce") or 0
            stk = str(row.iloc[BUY_STOCK])
            if vol > avq:
                loan_2a_pass = False
                loan_2a_detail.append(f"{stk}: Vol {vol:,.0f} > Avail {avq:,.0f}")
        add(
            "2a. Volume Buy ≤ Available Quantity",
            loan_2a_pass,
            "; ".join(loan_2a_detail) if loan_2a_detail else f"Total Volume Buy: {total_buy_vol:,.0f}"
        )

        # ── 2b. Rasio Loan Request < 65% ──────────────────────────────────
        if not has_loan_request:
            loan_2b_pass   = True
            loan_2b_detail = "Tidak ada Loan Request — check 2b dilewati"
        else:
            loan_num_2b     = loan_existing - total_sell_val + accrued_interest + total_buy_val
            collateral_loan = buy_collateral_total
            if collateral_loan > 0:
                ratio_loan     = loan_num_2b / collateral_loan
                loan_2b_pass   = ratio_loan < RATIO_THRESHOLD
                loan_2b_detail = (
                    f"Rasio: {fmt_pct(ratio_loan)} (threshold <{RATIO_THRESHOLD*100:.0f}%) | "
                    f"Numerator: {fmt_rp(loan_num_2b)} | "
                    f"Collateral: {fmt_rp(collateral_loan)}"
                )
                if len(buy_coll_detail) > 1:
                    breakdown = " | ".join(
                        f"{d['stock']}: {d['qty']:,.0f}×{fmt_rp(d['cp'])}×(1-{d['hc']*100:.0f}%)={fmt_rp(d['collateral'])}"
                        for d in buy_coll_detail
                    )
                    loan_2b_detail += f" || Breakdown: {breakdown}"
            elif loan_num_2b <= 0:
                loan_2b_pass   = True
                loan_2b_detail = f"Loan Numerator ≤ 0 ({fmt_rp(loan_num_2b)}) — posisi lunas"
            else:
                loan_2b_pass   = False
                loan_2b_detail = "Collateral = 0 sementara ada Loan — Rasio tidak terhitung (∞)"
        add("2b. Rasio Loan Request < 65%", loan_2b_pass, loan_2b_detail)

        # ── 3. Credit Limit Nasabah ────────────────────────────────────────
        if not has_loan_request:
            cl_nasabah_pass   = True
            cl_nasabah_detail = "Tidak ada Loan Request — check 3 dilewati"
        else:
            effective_limit   = available_limit + total_sell_val
            cl_nasabah_pass   = effective_limit > total_buy_val
            cl_nasabah_detail = (
                f"Avail Limit: {fmt_rp(available_limit)} + "
                f"Sell: {fmt_rp(total_sell_val)} = "
                f"{fmt_rp(effective_limit)} | "
                f"Loan Diajukan: {fmt_rp(total_buy_val)}"
            )
        add("3. Credit Limit Nasabah", cl_nasabah_pass, cl_nasabah_detail)

        results[sid] = sid_results

    # ── 4. Global Credit Limit Partisipan ────────────────────────────────
    cl_partisipan_pass = (credit_limit_partisipan + global_total_sell_value) > global_total_buy_value
    global_result = {
        "passed":     cl_partisipan_pass,
        "detail":     (
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
    if not sid_data.get("has_repayment", False):
        return False
    repayment_labels = {"1a.", "1a-OP.", "1b.", "1c."}
    return all(
        c["passed"]
        for c in sid_data["checks"]
        if any(c["label"].startswith(p) for p in repayment_labels)
    )

def lolos_loan(sid_data):
    if not sid_data.get("has_loan_request", False):
        return False
    loan_labels = {"1a.", "1a-OP.", "1c.", "2a.", "2b.", "3."}
    return all(
        c["passed"]
        for c in sid_data["checks"]
        if any(c["label"].startswith(p) for p in loan_labels)
    )

# ─────────────────────────────────────────────
# EXCEL EXPORT — LOLOS
# ─────────────────────────────────────────────

def generate_repayment_excel(df_sell, sid_results):
    today_str = datetime.today().strftime("%Y%m%d")
    passed_sids = [sid for sid, data in sid_results.items() if lolos_repayment(data)]

    sheet1_rows = []
    for sid in passed_sids:
        rows      = df_sell[col(df_sell, SELL_SID).astype(str) == sid]
        total_vol = pd.to_numeric(col(rows, SELL_VOL), errors="coerce").abs().sum()
        total_val = pd.to_numeric(col(rows, SELL_VAL), errors="coerce").abs().sum()
        if total_vol > 0 and total_val > 0:
            sheet1_rows.append({"Participant Code": "EP", "SID Client": sid, "Repayment Value": total_val})

    sheet2_rows = []
    for sid in passed_sids:
        rows = df_sell[col(df_sell, SELL_SID).astype(str) == sid]
        for _, row in rows.iterrows():
            qty = abs(pd.to_numeric(row.iloc[SELL_VOL], errors="coerce") or 0)
            if qty > 0:
                sheet2_rows.append({"SID Client": sid, "Stock Code": str(row.iloc[SELL_STOCK]), "Quantity": qty})

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(sheet1_rows).to_excel(writer, sheet_name="Repayment Proceed", index=False)
        pd.DataFrame(sheet2_rows).to_excel(writer, sheet_name="Detail Collateral",  index=False)
    buf.seek(0)
    return buf, f"Repayment Proceed {today_str}.xlsx"


def generate_loan_excel(df_buy, sid_results):
    today_str = datetime.today().strftime("%Y%m%d")
    passed_sids = [sid for sid, data in sid_results.items() if lolos_loan(data)]

    sheet1_rows = []
    for sid in passed_sids:
        rows      = df_buy[col(df_buy, BUY_SID).astype(str) == sid]
        total_vol = pd.to_numeric(col(rows, BUY_VOL), errors="coerce").sum()
        total_val = pd.to_numeric(col(rows, BUY_VAL), errors="coerce").sum()
        if total_vol > 0 and total_val > 0:
            sheet1_rows.append({"Participant Code": "EP", "SID Client": sid, "Loan Value": total_val})

    sheet2_rows = []
    for sid in passed_sids:
        rows = df_buy[col(df_buy, BUY_SID).astype(str) == sid]
        for _, row in rows.iterrows():
            qty = pd.to_numeric(row.iloc[BUY_VOL], errors="coerce") or 0
            if qty > 0:
                sheet2_rows.append({"SID Client": sid, "Stock Code": str(row.iloc[BUY_STOCK]), "Quantity": qty})

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(sheet1_rows).to_excel(writer, sheet_name="Loan Request",      index=False)
        pd.DataFrame(sheet2_rows).to_excel(writer, sheet_name="Detail Collateral", index=False)
    buf.seek(0)
    return buf, f"Loan Request {today_str}.xlsx"

# ─────────────────────────────────────────────
# EXCEL EXPORT — REVISI (NASABAH GAGAL)
# ─────────────────────────────────────────────

def generate_revisi_repayment_excel(df_sell, sid_results, op_data):
    today_str   = datetime.today().strftime("%Y%m%d")
    failed_sids = [
        sid for sid, data in sid_results.items()
        if data.get("has_repayment") and not lolos_repayment(data)
    ]

    alasan_rows = []
    for sid in failed_sids:
        data = sid_results[sid]
        op   = op_data.get(sid, {})
        for check in data["checks"]:
            if check["label"].startswith(("1a.", "1a-OP.", "1b.", "1c.")) and not check["passed"]:
                alasan_rows.append({
                    "SID"             : sid,
                    "Nama"            : data["name"],
                    "Check Gagal"     : check["label"],
                    "Detail Alasan"   : check["detail"],
                    "Loan Existing"   : op.get("loan_existing", "-"),
                    "Accrued Interest": op.get("accrued_interest", "-"),
                    "Volume Existing" : op.get("volume_existing", "-"),
                    "Saham OP"        : ", ".join(op.get("stocks", {}).keys()) or "-",
                })

    df_failed_sell = df_sell[
        col(df_sell, SELL_SID).astype(str).isin([str(s) for s in failed_sids])
    ].copy()

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(alasan_rows).to_excel(writer, sheet_name="Alasan Gagal",    index=False)
        df_failed_sell.to_excel(           writer, sheet_name="Sell (Repayment)", index=False)
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
        data = sid_results[sid]
        op   = op_data.get(sid, {})
        cl   = cl_data.get(sid, {})
        for check in data["checks"]:
            if check["label"].startswith(("1a.", "1a-OP.", "1c.", "2a.", "2b.", "3.")) and not check["passed"]:
                alasan_rows.append({
                    "SID"             : sid,
                    "Nama"            : data["name"],
                    "Check Gagal"     : check["label"],
                    "Detail Alasan"   : check["detail"],
                    "Loan Existing"   : op.get("loan_existing", "-"),
                    "Accrued Interest": op.get("accrued_interest", "-"),
                    "Available Limit" : cl.get("available_limit", "-"),
                })

    df_failed_buy = df_buy[
        col(df_buy, BUY_SID).astype(str).isin([str(s) for s in failed_sids])
    ].copy()

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(alasan_rows).to_excel(writer, sheet_name="Alasan Gagal", index=False)
        df_failed_buy.to_excel(            writer, sheet_name="Buy (Loan)",   index=False)
    buf.seek(0)
    return buf, f"Revisi Loan Request {today_str}.xlsx", len(failed_sids)

# ─────────────────────────────────────────────
# CREDIT LIMIT PARTISIPAN — HARDCODED
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

# ── Preview setelah upload ────────────────────────────────────────────
if op_file or cl_file:
    prev_col1, prev_col2 = st.columns(2)

    if op_file:
        with prev_col1:
            st.subheader("👁 Preview OP File")
            try:
                op_content_prev = op_file.read().decode("utf-8", errors="replace")
                op_file.seek(0)
                op_prev = parse_op_file(op_content_prev)
                sample_sids = list(op_prev.keys())[:3]
                for s in sample_sids:
                    d = op_prev[s]
                    with st.container(border=True):
                        st.markdown(f"**SID: {s}** — {d['name']}")
                        st.caption(
                            f"Loan: {fmt_rp(d['loan_existing'])} | "
                            f"Accrued: {fmt_rp(d['accrued_interest'])} | "
                            f"Vol Existing: {d['volume_existing']:,.0f} lot | "
                            f"Saham: {len(d['stocks'])} kode"
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
                cl_prev, vdate = parse_credit_limit_file(cl_content_prev)

                today_str_check = datetime.today().strftime("%Y/%m/%d")
                if vdate and vdate != today_str_check:
                    st.warning(
                        f"⚠️ Value Date file: **{vdate}** — "
                        f"berbeda dengan hari ini ({today_str_check}). "
                        f"Pastikan file sudah sesuai tanggal transaksi."
                    )

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

# ── Credit Limit Partisipan — Info ───────────────────────────────────
st.subheader("Credit Limit Partisipan")
st.info(f"Credit Limit Partisipan ditetapkan: **{fmt_rp(CREDIT_LIMIT_PARTISIPAN)}**")

st.divider()

# ── Panduan ───────────────────────────────────────────────────────────
with st.expander("📖 Panduan Validasi"):
    st.markdown("""
    | # | Cek | Keterangan |
    |---|-----|------------|
    | 1a | Volume Sell ≤ Available Sell Qty | Per baris di Sell sheet (dari Hasil_MNC) |
    | 1a-OP | Saham Sell ada & Vol ≤ Outstanding di OP File | Cross-check ke OP file — saham harus ada & vol tidak melebihi posisi outstanding |
    | 1b | Total Repayment Value ≤ Total Loan Value | Hanya jika ada Loan Request |
    | 1c | Rasio Repayment < 65% | (Loan Existing − Repayment + Accrued) / Collateral |
    | 2a | Volume Buy ≤ Available Quantity | Per baris di Buy sheet |
    | 2b | Rasio Loan Request < 65% | (Loan Existing − Repayment + Accrued + Loan Baru) / Collateral |
    | 3  | Credit Limit Nasabah | (Avail Limit + Sell Value) > Loan Diajukan |
    | 4  | Credit Limit Partisipan | Global: (CL Partisipan + Total Sell) > Total Loan |

    > **Lolos Repayment:** 1a + **1a-OP** + 1b + 1c  
    > **Lolos Loan Request:** 1a + **1a-OP** + 1c + 2a + 2b + 3
    """)

with st.expander("📖 Panduan Struktur File"):
    st.markdown("""
    | # | File | Sumber File |
    |---|-----|------------|
    | 1 | Hasil_MNC | Didapat dari Generate TRX PEI DETAILS |
    | 2 | File OP | Unduh .txt dari I-Fast Web |
    | 3 | Credit Limit Nasabah | Unduh .txt dari I-Fast Web |
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

    if errors:
        for e in errors:
            st.error(e)
        st.stop()

    with st.spinner("⚙️ Memproses seluruh data..."):
        try:
            df_sell, df_buy = load_hasil_mnc(hasil_file)
        except Exception as ex:
            st.error(f"Gagal membaca Hasil_MNC: {ex}")
            st.stop()

        try:
            op_content = op_file.read().decode("utf-8", errors="replace")
            op_data    = parse_op_file(op_content)
        except Exception as ex:
            st.error(f"Gagal membaca OP file: {ex}")
            st.stop()

        try:
            cl_content  = cl_file.read().decode("utf-8", errors="replace")
            cl_data, _  = parse_credit_limit_file(cl_content)
        except Exception as ex:
            st.error(f"Gagal membaca Credit Limit file: {ex}")
            st.stop()

        sid_results, global_result = run_validations(
            df_sell, df_buy, op_data, cl_data, CREDIT_LIMIT_PARTISIPAN
        )

    st.success("✅ Validasi Selesai!")

    # ── Metric Cards ─────────────────────────────────────────────────
    total_sids      = len(sid_results)
    total_pass_rep  = sum(1 for v in sid_results.values() if lolos_repayment(v))
    total_pass_loan = sum(1 for v in sid_results.values() if lolos_loan(v))
    total_fail_rep  = sum(1 for v in sid_results.values() if v.get("has_repayment")    and not lolos_repayment(v))
    total_fail_loan = sum(1 for v in sid_results.values() if v.get("has_loan_request") and not lolos_loan(v))
    global_pass     = global_result["passed"]

    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("Total Nasabah",      total_sids)
    m2.metric("Lolos Repayment",    total_pass_rep)
    m3.metric("Lolos Loan",         total_pass_loan)
    m4.metric("Gagal Repayment",    total_fail_rep,  delta=f"-{total_fail_rep}"  if total_fail_rep  else None, delta_color="inverse")
    m5.metric("Gagal Loan",         total_fail_loan, delta=f"-{total_fail_loan}" if total_fail_loan else None, delta_color="inverse")
    m6.metric("CL Partisipan",      "✅ LOLOS" if global_pass else "❌ GAGAL")

    st.divider()

    # ── Tabs ─────────────────────────────────────────────────────────
    tab_global, tab_per_sid, tab_gagal, tab_export = st.tabs([
        "🌐 Validasi 4 (Global)",
        "👤 Validasi Per Nasabah",
        "❌ Nasabah Gagal",
        "📥 Export Hasil",
    ])

    # ── Tab 1: Global CL Partisipan ───────────────────────────────────
    with tab_global:
        st.subheader("Validasi 4 — Credit Limit Partisipan (Global)")
        if global_result["passed"]:
            st.success(f"✅ Credit Limit Partisipan LOLOS — {global_result['detail']}")
        else:
            st.error(f"❌ Credit Limit Partisipan GAGAL — {global_result['detail']}")

    # ── Tab 2: Per-SID ────────────────────────────────────────────────
    with tab_per_sid:
        st.caption("🟢 R = Lolos Repayment  |  🟢 L = Lolos Loan  |  🔴 R = Gagal Repayment  |  🔴 L = Gagal Loan")
        for sid, data in sid_results.items():
            checks   = data["checks"]
            r_pass   = lolos_repayment(data)
            l_pass   = lolos_loan(data)
            r_icon   = "✅ R" if r_pass else "❌ R"
            l_icon   = "✅ L" if l_pass else "❌ L"
            expanded = not (r_pass and l_pass)

            with st.expander(f"{r_icon} | {l_icon} | {sid} — {data['name']}", expanded=expanded):
                for check in checks:
                    if check["passed"]:
                        st.success(f"✅ **{check['label']}**  {check['detail']}")
                    else:
                        st.error(f"❌ **{check['label']}**  {check['detail']}")

    # ── Tab 3: Nasabah Gagal ──────────────────────────────────────────
    with tab_gagal:
        st.subheader("Nasabah yang Tidak Lolos Validasi")
        st.caption(
            "Daftar ini menunjukkan nasabah yang perlu direvisi sebelum dapat di-upload ke sistem. "
            "Download file revisi di tab **Export Hasil**."
        )

        gagal_rep  = [(sid, d) for sid, d in sid_results.items()
                      if d.get("has_repayment")    and not lolos_repayment(d)]
        gagal_loan = [(sid, d) for sid, d in sid_results.items()
                      if d.get("has_loan_request") and not lolos_loan(d)]

        gcol1, gcol2 = st.columns(2)

        with gcol1:
            st.markdown(f"#### 🔴 Gagal Repayment — {len(gagal_rep)} nasabah")
            if not gagal_rep:
                st.success("Semua nasabah lolos Repayment.")
            else:
                for sid, data in gagal_rep:
                    failed_checks = [
                        c for c in data["checks"]
                        if c["label"].startswith(("1a.", "1a-OP.", "1b.", "1c.")) and not c["passed"]
                    ]
                    with st.expander(f"❌ {sid} — {data['name']}"):
                        for c in failed_checks:
                            st.error(f"**{c['label']}** — {c['detail']}")

        with gcol2:
            st.markdown(f"#### 🔴 Gagal Loan Request — {len(gagal_loan)} nasabah")
            if not gagal_loan:
                st.success("Semua nasabah lolos Loan Request.")
            else:
                for sid, data in gagal_loan:
                    failed_checks = [
                        c for c in data["checks"]
                        if c["label"].startswith(("1a.", "1a-OP.", "1c.", "2a.", "2b.", "3.")) and not c["passed"]
                    ]
                    with st.expander(f"❌ {sid} — {data['name']}"):
                        for c in failed_checks:
                            st.error(f"**{c['label']}** — {c['detail']}")

    # ── Tab 4: Export ─────────────────────────────────────────────────
    with tab_export:

        # — Summary table —
        st.subheader("📋 Ringkasan Hasil Validasi")
        summary_rows = []
        for sid, data in sid_results.items():
            row = {"SID": sid, "Nama": data["name"]}
            for check in data["checks"]:
                row[check["label"]] = "LOLOS" if check["passed"] else "GAGAL"
            row["Status Repayment"] = "LOLOS" if lolos_repayment(data) else "GAGAL"
            row["Status Loan"]      = "LOLOS" if lolos_loan(data)      else "GAGAL"
            summary_rows.append(row)

        df_summary = pd.DataFrame(summary_rows)
        st.dataframe(df_summary, use_container_width=True)

        buf_sum = io.BytesIO()
        df_summary.to_excel(buf_sum, index=False)
        buf_sum.seek(0)
        st.download_button(
            "⬇️ Download Hasil Validasi (.xlsx)",
            data=buf_sum,
            file_name="hasil_validasi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.divider()

        # — Export lolos ke Dashboard Kantor —
        st.subheader("📤 Export ke Dashboard Kantor")
        st.caption("Repayment Proceed: lolos 1a + 1a-OP + 1b + 1c  |  Loan Request: lolos 1a + 1a-OP + 1c + 2a + 2b + 3")

        dl_col1, dl_col2 = st.columns(2)
        with dl_col1:
            rep_buf, rep_fname = generate_repayment_excel(df_sell, sid_results)
            st.download_button(
                label="⬇️ Repayment Proceed (.xlsx)",
                data=rep_buf,
                file_name=rep_fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with dl_col2:
            loan_buf, loan_fname = generate_loan_excel(df_buy, sid_results)
            st.download_button(
                label="⬇️ Loan Request (.xlsx)",
                data=loan_buf,
                file_name=loan_fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        st.divider()

        # — Export revisi nasabah gagal —
        st.subheader("📝 Export File Revisi (Nasabah Gagal)")
        st.caption(
            "File berisi **Sheet 1 — Alasan Gagal** (referensi) dan "
            "**Sheet 2 — Data Transaksi** nasabah yang tidak lolos. "
            "Perbaiki nilai di Sheet 2, lalu salin kembali ke Hasil_MNC dan upload ulang."
        )

        rv_col1, rv_col2 = st.columns(2)

        with rv_col1:
            rev_rep_buf, rev_rep_fname, n_gagal_rep = generate_revisi_repayment_excel(
                df_sell, sid_results, op_data
            )
            if n_gagal_rep == 0:
                st.info("✅ Tidak ada nasabah yang gagal Repayment.")
            else:
                st.download_button(
                    label=f"⬇️ Revisi Repayment — {n_gagal_rep} nasabah gagal",
                    data=rev_rep_buf,
                    file_name=rev_rep_fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="secondary",
                )

        with rv_col2:
            rev_loan_buf, rev_loan_fname, n_gagal_loan = generate_revisi_loan_excel(
                df_buy, sid_results, op_data, cl_data
            )
            if n_gagal_loan == 0:
                st.info("✅ Tidak ada nasabah yang gagal Loan Request.")
            else:
                st.download_button(
                    label=f"⬇️ Revisi Loan Request — {n_gagal_loan} nasabah gagal",
                    data=rev_loan_buf,
                    file_name=rev_loan_fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="secondary",
                )
