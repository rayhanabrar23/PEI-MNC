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
                "stocks":           {},   # stock → lot
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


def load_hasil_mnc(uploaded_file):
    xls     = pd.ExcelFile(uploaded_file)

    # Coba baca sheet "Sell Aktif" dulu, fallback ke "Sell (Repayment)"
    if "Sell Aktif (tanpa excluded)" in xls.sheet_names:
        df_sell = pd.read_excel(xls, sheet_name="Sell Aktif (tanpa excluded)", header=0)
    else:
        df_sell = pd.read_excel(xls, sheet_name="Sell (Repayment)", header=0)
        # Buang baris yang excluded
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
SELL_VAL   = 12   # Value (Lot×Price)

BUY_SID    = 0
BUY_STOCK  = 1
BUY_AVQ    = 4
BUY_CP     = 5
BUY_HC     = 7
BUY_VOL    = 13
BUY_VAL    = 14   # Value (Lot×Price)

RATIO_THRESHOLD = 0.65

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


def calc_max_loan(loan_existing, sell_value, accrued_interest, collateral, threshold=0.65):
    """
    Hitung nilai loan maksimum yang masih di bawah threshold (65%).
    Formula: max_loan = collateral × threshold − (loan_existing − sell_value + accrued_interest)
    Jika hasil negatif → max_loan = 0 (tidak bisa ajukan loan sama sekali).
    """
    numerator_before_loan = loan_existing - sell_value + accrued_interest
    max_loan = collateral * threshold - numerator_before_loan
    return max(max_loan, 0)

# ─────────────────────────────────────────────
# VALIDATION LOGIC
# ─────────────────────────────────────────────

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

        # ── STEP 3: Guard — skip repayment jika Existing Loan = 0 ──
        # Jika loan_existing = 0, repayment tidak diperlukan → lewati semua check repayment
        repayment_skipped_no_loan = has_repayment and loan_existing <= 0

        sid_results = {
            "name":                      name,
            "checks":                    [],
            "has_repayment":             has_repayment,
            "has_loan_request":          has_loan_request,
            "repayment_skipped_no_loan": repayment_skipped_no_loan,
            "loan_existing":             loan_existing,
            "max_loan_recommendation":   0.0,
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

        # ── 1a. Volume Sell ≤ Available Sell Quantity ──────────────
        if repayment_skipped_no_loan:
            add(
                "1a. Volume Sell ≤ Available Sell Quantity",
                True,
                f"⏭ Dilewati — Existing Loan = 0, repayment tidak diperlukan"
            )
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
            add(
                "1a. Volume Sell ≤ Available Sell Quantity",
                rep_1a_pass,
                "; ".join(rep_1a_detail) if rep_1a_detail else f"Total Volume Sell: {total_sell_vol:,.0f}"
            )

        # ── 1a-OP. Saham Sell terverifikasi di OP File ────────────
        if repayment_skipped_no_loan:
            add(
                "1a-OP. Saham Sell Terverifikasi di OP File",
                True,
                f"⏭ Dilewati — Existing Loan = 0, repayment tidak diperlukan"
            )
        elif not has_repayment:
            add("1a-OP. Saham Sell Terverifikasi di OP File", True, "Tidak ada Repayment")
        elif not op.get("stocks"):
            add(
                "1a-OP. Saham Sell Terverifikasi di OP File",
                False,
                f"SID {sid} tidak ditemukan di OP file"
            )
        else:
            op_stocks       = op["stocks"]
            op_1a_pass      = True
            op_1a_detail    = []
            op_1a_ok_detail = []

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
                    op_1a_ok_detail.append(f"{stk}: {vol:,.0f} ✓")

            detail_msg = (
                "; ".join(op_1a_ok_detail) if op_1a_pass
                else "; ".join(op_1a_detail) + (
                    " || OK: " + "; ".join(op_1a_ok_detail) if op_1a_ok_detail else ""
                )
            )
            add("1a-OP. Saham Sell Terverifikasi di OP File", op_1a_pass, detail_msg)

        # ── 1b. Total Repayment Value ≤ Total Loan Value ──────────
        if repayment_skipped_no_loan:
            add(
                "1b. Total Repayment Value ≤ Total Loan Value",
                True,
                "⏭ Dilewati — Existing Loan = 0, repayment tidak diperlukan"
            )
        elif not has_loan_request:
            add("1b. Total Repayment Value ≤ Total Loan Value", True,
                "Pure repayment tanpa Loan Request — check 1b dilewati")
        else:
            rep_1b_pass   = total_sell_val <= total_buy_val
            add("1b. Total Repayment Value ≤ Total Loan Value", rep_1b_pass,
                f"Total Sell Value: {fmt_rp(total_sell_val)} | Total Buy Value: {fmt_rp(total_buy_val)}")

        # ── 1c. Rasio Repayment < 65% ──────────────────────────────
        if repayment_skipped_no_loan:
            add(
                "1c. Rasio Repayment < 65%",
                True,
                "⏭ Dilewati — Existing Loan = 0, repayment tidak diperlukan"
            )
        elif not has_repayment:
            add("1c. Rasio Repayment < 65%", True, "Tidak ada Repayment — check 1c dilewati")
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
                ratio_rep = loan_num_repayment / collateral_repayment
                add("1c. Rasio Repayment < 65%", ratio_rep < RATIO_THRESHOLD,
                    f"Rasio: {fmt_pct(ratio_rep)} | Numerator: {fmt_rp(loan_num_repayment)} | "
                    f"Collateral: {fmt_rp(collateral_repayment)}")
            elif loan_num_repayment <= 0:
                add("1c. Rasio Repayment < 65%", True,
                    f"Loan Numerator ≤ 0 ({fmt_rp(loan_num_repayment)}) — posisi lunas")
            else:
                add("1c. Rasio Repayment < 65%", False,
                    "Collateral = 0 sementara Loan masih ada — Rasio tidak terhitung (∞)")

        # ── 2a. Volume Buy ≤ Available Quantity ───────────────────
        loan_2a_pass   = True
        loan_2a_detail = []
        for _, row in buy["rows"].iterrows():
            vol = pd.to_numeric(row.iloc[BUY_VOL], errors="coerce") or 0
            avq = pd.to_numeric(row.iloc[BUY_AVQ], errors="coerce") or 0
            stk = str(row.iloc[BUY_STOCK])
            # STEP 4: Jika Stock Available = 0 → loan saham ini dibatalkan
            if avq == 0:
                loan_2a_pass = False
                loan_2a_detail.append(f"{stk}: DIBATALKAN (Stock Available = 0)")
            elif vol > avq:
                loan_2a_pass = False
                loan_2a_detail.append(f"{stk}: Vol {vol:,.0f} > Avail {avq:,.0f}")
        add(
            "2a. Volume Buy ≤ Available Quantity",
            loan_2a_pass,
            "; ".join(loan_2a_detail) if loan_2a_detail else f"Total Volume Buy: {total_buy_vol:,.0f}"
        )

        # ── 2b. Rasio Loan Request < 65% + Rekomendasi Max Loan ───
        if not has_loan_request:
            add("2b. Rasio Loan Request < 65%", True,
                "Tidak ada Loan Request — check 2b dilewati")
        else:
            loan_num_2b     = loan_existing - total_sell_val + accrued_interest + total_buy_val
            collateral_loan = buy_collateral_total

            # STEP 4: Hitung max loan rekomendasi
            max_loan_rec = calc_max_loan(
                loan_existing, total_sell_val, accrued_interest, collateral_loan
            )
            sid_results["max_loan_recommendation"] = max_loan_rec

            if collateral_loan > 0:
                ratio_loan   = loan_num_2b / collateral_loan
                loan_2b_pass = ratio_loan < RATIO_THRESHOLD
                detail_str   = (
                    f"Rasio: {fmt_pct(ratio_loan)} (threshold <{RATIO_THRESHOLD*100:.0f}%) | "
                    f"Numerator: {fmt_rp(loan_num_2b)} | Collateral: {fmt_rp(collateral_loan)}"
                )
                if not loan_2b_pass:
                    detail_str += (
                        f" || ⚠️ Loan diajukan ({fmt_rp(total_buy_val)}) melebihi kapasitas. "
                        f"Max loan yang aman: {fmt_rp(max_loan_rec)}"
                    )
                add("2b. Rasio Loan Request < 65%", loan_2b_pass, detail_str)
            elif loan_num_2b <= 0:
                add("2b. Rasio Loan Request < 65%", True,
                    f"Loan Numerator ≤ 0 ({fmt_rp(loan_num_2b)}) — posisi lunas")
            else:
                add("2b. Rasio Loan Request < 65%", False,
                    "Collateral = 0 sementara ada Loan — Rasio tidak terhitung (∞)")

        # ── 3. Credit Limit Nasabah ────────────────────────────────
        if not has_loan_request:
            add("3. Credit Limit Nasabah", True, "Tidak ada Loan Request — check 3 dilewati")
        else:
            effective_limit   = available_limit + total_sell_val
            cl_nasabah_pass   = effective_limit > total_buy_val
            add("3. Credit Limit Nasabah", cl_nasabah_pass,
                f"Avail Limit: {fmt_rp(available_limit)} + Sell: {fmt_rp(total_sell_val)} = "
                f"{fmt_rp(effective_limit)} | Loan Diajukan: {fmt_rp(total_buy_val)}")

        results[sid] = sid_results

    # ── 4. Global Credit Limit Partisipan ────────────────────────
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
    # Jika repayment di-skip karena loan = 0, tidak ada yang perlu diproses
    if sid_data.get("repayment_skipped_no_loan"):
        return False   # tidak masuk output repayment
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

# ─────────────────────────────────────────────
# EXCEL EXPORT — REVISI
# ─────────────────────────────────────────────

def generate_revisi_repayment_excel(df_sell, sid_results, op_data):
    today_str   = datetime.today().strftime("%Y%m%d")
    # Nasabah gagal = punya repayment, bukan skip loan=0, dan tidak lolos
    failed_sids = [
        sid for sid, data in sid_results.items()
        if data.get("has_repayment")
        and not data.get("repayment_skipped_no_loan")
        and not lolos_repayment(data)
    ]

    alasan_rows = []
    for sid in failed_sids:
        data = sid_results[sid]
        op   = op_data.get(sid, {})
        for check in data["checks"]:
            if check["label"].startswith(("1a.", "1a-OP.", "1b.", "1c.")) and not check["passed"]:
                alasan_rows.append({
                    "SID":              sid,
                    "Nama":             data["name"],
                    "Check Gagal":      check["label"],
                    "Detail Alasan":    check["detail"],
                    "Loan Existing":    op.get("loan_existing", "-"),
                    "Accrued Interest": op.get("accrued_interest", "-"),
                    "Volume Existing":  op.get("volume_existing", "-"),
                    "Saham OP":         ", ".join(op.get("stocks", {}).keys()) or "-",
                })

    df_failed_sell = df_sell[
        col(df_sell, SELL_SID).astype(str).isin([str(s) for s in failed_sids])
    ].copy()

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
        for check in data["checks"]:
            if check["label"].startswith(("1a.", "1a-OP.", "1c.", "2a.", "2b.", "3.")) and not check["passed"]:
                alasan_rows.append({
                    "SID":                    sid,
                    "Nama":                   data["name"],
                    "Check Gagal":            check["label"],
                    "Detail Alasan":          check["detail"],
                    "Loan Existing":          op.get("loan_existing", "-"),
                    "Accrued Interest":       op.get("accrued_interest", "-"),
                    "Available Limit":        cl.get("available_limit", "-"),
                    "Max Loan Rekomendasi":   max_loan if max_loan > 0 else "Tidak bisa ajukan loan",
                })

    df_failed_buy = df_buy[
        col(df_buy, BUY_SID).astype(str).isin([str(s) for s in failed_sids])
    ].copy()

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
                        # Highlight jika loan = 0
                        loan_info = fmt_rp(d['loan_existing'])
                        if d['loan_existing'] <= 0:
                            st.warning(f"⚠️ Loan Existing = 0 → Repayment akan dilewati otomatis")
                        st.caption(
                            f"Loan: {loan_info} | Accrued: {fmt_rp(d['accrued_interest'])} | "
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
                    st.warning(
                        f"⚠️ Value Date file: **{vdate}** — berbeda dengan hari ini ({today_str_check})."
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

st.subheader("Credit Limit Partisipan")
st.info(f"Credit Limit Partisipan ditetapkan: **{fmt_rp(CREDIT_LIMIT_PARTISIPAN)}**")
st.divider()

with st.expander("📖 Panduan Validasi"):
    st.markdown("""
    | # | Cek | Keterangan |
    |---|-----|------------|
    | — | **Guard: Existing Loan = 0** | Jika loan existing = 0 → **semua check repayment dilewati**, nasabah tidak masuk output Repayment Proceed |
    | 1a | Volume Sell ≤ Available Sell Qty | Per baris di Sell sheet |
    | 1a-OP | Saham Sell ada & Vol ≤ Outstanding di OP | Cross-check ke OP file |
    | 1b | Total Repayment Value ≤ Total Loan Value | Hanya jika ada Loan Request |
    | 1c | Rasio Repayment < 65% | (Loan Existing − Repayment + Accrued) / Collateral |
    | 2a | Volume Buy ≤ Available Quantity | Jika Avail = 0 → saham **dibatalkan** |
    | 2b | Rasio Loan Request < 65% | Jika gagal → tampilkan **max loan rekomendasi** |
    | 3  | Credit Limit Nasabah | (Avail Limit + Sell Value) > Loan Diajukan |
    | 4  | Credit Limit Partisipan | Global: (CL Partisipan + Total Sell) > Total Loan |

    > **Lolos Repayment:** 1a + 1a-OP + 1b + 1c (hanya jika Loan Existing > 0)  
    > **Lolos Loan Request:** 1a + 1a-OP + 1c + 2a + 2b + 3
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
            cl_content = cl_file.read().decode("utf-8", errors="replace")
            cl_data, _ = parse_credit_limit_file(cl_content)
        except Exception as ex:
            st.error(f"Gagal membaca Credit Limit file: {ex}")
            st.stop()

        sid_results, global_result = run_validations(
            df_sell, df_buy, op_data, cl_data, CREDIT_LIMIT_PARTISIPAN
        )

    st.success("✅ Validasi Selesai!")

    # Metric cards
    total_sids           = len(sid_results)
    total_pass_rep       = sum(1 for v in sid_results.values() if lolos_repayment(v))
    total_pass_loan      = sum(1 for v in sid_results.values() if lolos_loan(v))
    total_fail_rep       = sum(1 for v in sid_results.values()
                               if v.get("has_repayment") and not v.get("repayment_skipped_no_loan")
                               and not lolos_repayment(v))
    total_fail_loan      = sum(1 for v in sid_results.values()
                               if v.get("has_loan_request") and not lolos_loan(v))
    total_skip_no_loan   = sum(1 for v in sid_results.values() if v.get("repayment_skipped_no_loan"))
    global_pass          = global_result["passed"]

    m1, m2, m3, m4, m5, m6, m7 = st.columns(7)
    m1.metric("Total Nasabah",        total_sids)
    m2.metric("Lolos Repayment",      total_pass_rep)
    m3.metric("Lolos Loan",           total_pass_loan)
    m4.metric("Gagal Repayment",      total_fail_rep,
              delta=f"-{total_fail_rep}"  if total_fail_rep  else None, delta_color="inverse")
    m5.metric("Gagal Loan",           total_fail_loan,
              delta=f"-{total_fail_loan}" if total_fail_loan else None, delta_color="inverse")
    m6.metric("Skip (Loan=0)",        total_skip_no_loan)
    m7.metric("CL Partisipan",        "✅ LOLOS" if global_pass else "❌ GAGAL")

    st.divider()

    tab_global, tab_per_sid, tab_gagal, tab_export = st.tabs([
        "🌐 Validasi 4 (Global)",
        "👤 Validasi Per Nasabah",
        "❌ Nasabah Gagal",
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
                    st.info(f"ℹ️ Existing Loan = 0 → Repayment tidak diproses untuk nasabah ini.")
                # Tampilkan max loan recommendation jika ada
                max_loan = data.get("max_loan_recommendation", 0)
                if max_loan > 0 and not l_pass:
                    st.warning(f"💡 Max Loan Rekomendasi (rasio <65%): **{fmt_rp(max_loan)}**")
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
                    max_loan = data.get("max_loan_recommendation", 0)
                    failed_checks = [c for c in data["checks"]
                                     if c["label"].startswith(("1a.", "1a-OP.", "1c.", "2a.", "2b.", "3."))
                                     and not c["passed"]]
                    with st.expander(f"❌ {sid} — {data['name']}"):
                        if max_loan > 0:
                            st.warning(f"💡 Max Loan Rekomendasi: **{fmt_rp(max_loan)}**")
                        for c in failed_checks:
                            st.error(f"**{c['label']}** — {c['detail']}")

    with tab_export:
        st.subheader("📋 Ringkasan Hasil Validasi")
        summary_rows = []
        for sid, data in sid_results.items():
            row = {"SID": sid, "Nama": data["name"],
                   "Loan Existing": data.get("loan_existing", 0),
                   "Skip Repayment (Loan=0)": "YA" if data.get("repayment_skipped_no_loan") else "TIDAK",
                   "Max Loan Rekomendasi": data.get("max_loan_recommendation", 0)}
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
        st.download_button(
            "⬇️ Download Hasil Validasi (.xlsx)",
            data=buf_sum,
            file_name="hasil_validasi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.divider()
        st.subheader("📤 Export ke Dashboard Kantor")

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
        st.subheader("📝 Export File Revisi (Nasabah Gagal)")

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
