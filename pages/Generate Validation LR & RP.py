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

# ─────────────────────────────────────────────
# CUSTOM CSS — Light Theme
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
}
.stApp {
    background-color: #f5f7fa;
    color: #1a1f2e;
}
h1, h2, h3 {
    font-family: 'IBM Plex Mono', monospace;
    color: #1a5fe8;
}
.stButton > button {
    background: #1a5fe8;
    color: white;
    border: none;
    border-radius: 6px;
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 600;
    padding: 0.5rem 1.5rem;
    transition: background 0.2s;
}
.stButton > button:hover {
    background: #3a7af0;
}
.validation-box {
    border-radius: 8px;
    padding: 1rem 1.25rem;
    margin: 0.5rem 0;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.85rem;
}
.pass  { background:#e8f8ed; border-left:4px solid #22a84a; color:#166332; }
.fail  { background:#fdecea; border-left:4px solid #e53935; color:#b71c1c; }
.warn  { background:#fff8e1; border-left:4px solid #f9a825; color:#7a5700; }
.section-header {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.75rem;
    letter-spacing: 0.15em;
    text-transform: uppercase;
    color: #5a6278;
    margin: 1.5rem 0 0.5rem 0;
    border-bottom: 1px solid #d0d7e3;
    padding-bottom: 0.3rem;
}
.metric-card {
    background: #ffffff;
    border: 1px solid #d0d7e3;
    border-radius: 8px;
    padding: 0.75rem 1rem;
    text-align: center;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}
.metric-label {
    font-size: 0.7rem;
    color: #5a6278;
    font-family: 'IBM Plex Mono', monospace;
    text-transform: uppercase;
    letter-spacing: 0.1em;
}
.metric-value {
    font-size: 1.4rem;
    font-weight: 600;
    font-family: 'IBM Plex Mono', monospace;
}
.preview-box {
    background: #ffffff;
    border: 1px solid #d0d7e3;
    border-radius: 8px;
    padding: 0.75rem 1rem;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.78rem;
    color: #1a1f2e;
    margin-top: 0.5rem;
}
.preview-label {
    font-size: 0.68rem;
    color: #5a6278;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    margin-bottom: 0.4rem;
}
.date-warn {
    background: #fff8e1;
    border: 1px solid #f9a825;
    border-radius: 6px;
    padding: 0.5rem 1rem;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.78rem;
    color: #7a5700;
    margin-top: 0.4rem;
}
.stDataFrame { background: #ffffff; }
[data-testid="stExpander"] {
    background: #ffffff;
    border: 1px solid #d0d7e3;
    border-radius: 8px;
}
</style>
""", unsafe_allow_html=True)

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

def pass_box(msg):
    st.markdown(f'<div class="validation-box pass">✅ {msg}</div>', unsafe_allow_html=True)

def fail_box(msg):
    st.markdown(f'<div class="validation-box fail">❌ {msg}</div>', unsafe_allow_html=True)

def warn_box(msg):
    st.markdown(f'<div class="validation-box warn">⚠️ {msg}</div>', unsafe_allow_html=True)

def section(title):
    st.markdown(f'<div class="section-header">{title}</div>', unsafe_allow_html=True)

def parse_hc(raw):
    """
    Konversi nilai Haircut ke desimal (0–1).
    Menerima: '18%', '0.18', '18', 18.0, dll.
    """
    try:
        s = str(raw).replace("%", "").strip()
        v = float(s)
        # Jika nilainya > 1, asumsikan dalam bentuk persen (mis. 18 → 0.18)
        return v / 100 if v > 1 else v
    except Exception:
        return 0.0

# ─────────────────────────────────────────────
# PARSERS
# ─────────────────────────────────────────────

def parse_op_file(content: str):
    """
    Parse OP txt file.
    Format baris '0': header nasabah
        0|YYYYMMDD|BrokerCode|SID|Name|LoanExisting|AccruedInterest|...|HC%|...|...
    Format baris '1': detail saham per nasabah
        1|BrokerCode|SID|StockCode|Volume

    Returns dict keyed by SID:
        loan_existing, accrued_interest, volume_existing (sum lot seluruh saham), name
    """
    result = {}
    lines = content.strip().splitlines()
    for line in lines:
        line = line.strip()
        if not line:
            continue
        parts = line.split("|")
        if parts[0] == "0":
            # Minimal 7 kolom: 0|date|broker|SID|name|loan|accrued
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
                # Simpan detail saham untuk collateral akurat (diisi dari baris '1')
                "stocks": {},
            }
        elif parts[0] == "1":
            # Format: 1|BrokerCode|SID|StockCode|Volume
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
    """
    Parse CreditLimit txt file.
    Format: Value Date|Broker Code|SID|Client Name|Client Credit Limit|Used Limit|Available Limit

    Returns:
        dict keyed by SID: {available_limit, name, value_date}
        str value_date (dari baris pertama data)
    """
    result     = {}
    value_date = None
    lines      = content.strip().splitlines()

    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue
        parts = line.split("|")

        # Skip header baris pertama
        if i == 0 and parts[0].strip().lower() == "value date":
            continue

        if len(parts) < 7:
            continue

        sid = parts[2].strip()
        try:
            available_limit = float(parts[6].replace(",", ""))
        except Exception:
            available_limit = 0.0

        # Ambil value_date dari baris data pertama
        if value_date is None:
            value_date = parts[0].strip()

        result[sid] = {
            "available_limit": available_limit,
            "name":            parts[3].strip(),
            "value_date":      parts[0].strip(),
        }

    return result, value_date


def load_hasil_mnc(uploaded_file):
    """Load Hasil_MNC excel. Returns (df_sell, df_buy)."""
    xls    = pd.ExcelFile(uploaded_file)
    df_sell = pd.read_excel(xls, sheet_name="Sell (Repayment)", header=0)
    df_buy  = pd.read_excel(xls, sheet_name="Buy (Loan)",  header=0)
    return df_sell, df_buy

# ─────────────────────────────────────────────
# COLUMN ACCESSORS (by index, 0-based)
# ─────────────────────────────────────────────

def col(df, idx):
    return df.iloc[:, idx]

# Sell sheet columns (0-based)
SELL_SID   = 0
SELL_STOCK = 1
SELL_AVQ   = 4   # Available Sell Quantity
SELL_CP    = 5   # Closing Price
SELL_VOL   = 11  # Volume
SELL_VAL   = 12  # Value

# Buy sheet columns (0-based)
BUY_SID    = 0
BUY_STOCK  = 1
BUY_AVQ    = 4   # Available Quantity
BUY_CP     = 5   # Closing Price
BUY_HC     = 7   # Haircut
BUY_VOL    = 13  # Volume
BUY_VAL    = 14  # Value

# ─────────────────────────────────────────────
# COLLATERAL CALCULATOR (FIX UTAMA)
# ─────────────────────────────────────────────

def calc_collateral_from_rows(rows, qty_col_idx, cp_col_idx, hc_col_idx):
    """
    Hitung total collateral value per saham secara akurat.
    Collateral per baris = Qty × ClosingPrice × (1 - Haircut)

    Berbeda dari versi lama yang memakai rata-rata HC × net_vol,
    fungsi ini menghitung per baris sehingga akurat meski tiap saham
    punya harga dan HC yang berbeda.

    Returns:
        total_collateral (float)
        detail (list of dict, untuk debug/display)
    """
    total  = 0.0
    detail = []
    for _, row in rows.iterrows():
        qty = abs(pd.to_numeric(row.iloc[qty_col_idx], errors="coerce") or 0)
        cp  = pd.to_numeric(row.iloc[cp_col_idx], errors="coerce") or 0
        hc  = parse_hc(row.iloc[hc_col_idx])
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

        # ── Collateral dari Buy rows (FIX: per saham, bukan rata-rata) ──
        # Digunakan untuk check 1c (setelah repayment) dan 2b (setelah loan)
        buy_collateral_total, buy_coll_detail = calc_collateral_from_rows(
            buy["rows"], BUY_AVQ, BUY_CP, BUY_HC
        )

        # ── 1. Validasi Repayment Proceed ─────────────────────────────

        # 1a: Per baris Volume Sell ≤ Available Sell Qty
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
            "; ".join(rep_1a_detail) if rep_1a_detail
            else f"Total Volume Sell: {total_sell_vol:,.0f}"
        )

        # 1b: Total Value Sell ≤ Total Value Buy (Loan Value)
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

        # 1c: Rasio setelah Repayment < 65%
        # Numerator   = Loan Existing - Repayment Value + Accrued Interest
        # Denominator = Collateral dari sisa saham setelah repayment
        #
        # FIX: collateral dihitung dari buy_rows per saham (bukan avg × net_vol).
        # Sisa collateral setelah repayment = buy_collateral_total
        # (karena repayment melunasi loan lama, buy_collateral_total adalah
        # collateral baru yang akan menggantikan posisi lama).
        #
        # Jika tidak ada buy rows, gunakan approximasi:
        # collateral_repayment ≈ (volume_existing - total_sell_vol) × avg_price_hc dari sell rows
        if not has_repayment:
            rep_1c_pass   = True
            rep_1c_detail = "Tidak ada Repayment — check 1c dilewati"
        else:
            loan_num_repayment = loan_existing - total_sell_val + accrued_interest

            if buy_collateral_total > 0:
                # Ada buy rows → pakai collateral baru yang akurat
                collateral_repayment = buy_collateral_total
            else:
                # Tidak ada buy rows (pure repayment) → estimasi dari sisa volume × harga sell
                # Ambil closing price rata-rata dari sell rows sebagai proxy
                sell_cp_vals = pd.to_numeric(col(sell["rows"], SELL_CP), errors="coerce").dropna()
                avg_sell_cp  = sell_cp_vals.mean() if len(sell_cp_vals) > 0 else 0
                sisa_vol     = max(volume_existing - total_sell_vol, 0)
                # Tanpa HC dari sell file, gunakan 0% sebagai konservatif (worst case)
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
                # Loan sudah lunas atau negatif → lolos
                rep_1c_pass   = True
                rep_1c_detail = f"Loan Numerator ≤ 0 ({fmt_rp(loan_num_repayment)}) — posisi lunas"
            else:
                rep_1c_pass   = False
                rep_1c_detail = "Collateral = 0 sementara Loan masih ada — Rasio tidak terhitung (∞)"

        add("1c. Rasio Repayment < 65%", rep_1c_pass, rep_1c_detail)

        # ── 2. Validasi Loan Request ──────────────────────────────────

        # 2a: Per baris Volume Buy ≤ Available Qty
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
            "; ".join(loan_2a_detail) if loan_2a_detail
            else f"Total Volume Buy: {total_buy_vol:,.0f}"
        )

        # 2b: Rasio setelah Loan Request < 65%
        # Numerator   = Loan Existing - Repayment Value + Accrued Interest + Loan Baru
        # Denominator = Collateral dari seluruh posisi baru (buy rows)
        #
        # FIX: collateral dihitung per saham dari buy_rows, bukan avg × net_vol.
        # buy_collateral_total sudah dihitung di atas.
        if not has_loan_request:
            loan_2b_pass   = True
            loan_2b_detail = "Tidak ada Loan Request — check 2b dilewati"
        else:
            loan_num_2b      = loan_existing - total_sell_val + accrued_interest + total_buy_val
            collateral_loan  = buy_collateral_total  # per-saham, akurat

            if collateral_loan > 0:
                ratio_loan     = loan_num_2b / collateral_loan
                loan_2b_pass   = ratio_loan < RATIO_THRESHOLD
                loan_2b_detail = (
                    f"Rasio: {fmt_pct(ratio_loan)} (threshold <{RATIO_THRESHOLD*100:.0f}%) | "
                    f"Numerator: {fmt_rp(loan_num_2b)} | "
                    f"Collateral: {fmt_rp(collateral_loan)}"
                )
                # Tambahkan breakdown collateral per saham jika lebih dari 1 saham
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

        # ── 3. Validasi Credit Limit Nasabah ─────────────────────────
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

    # ── 4. Validasi Credit Limit Partisipan (global) ─────────────────
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
    repayment_labels = {"1a.", "1b.", "1c."}
    return all(
        c["passed"]
        for c in sid_data["checks"]
        if any(c["label"].startswith(p) for p in repayment_labels)
    )

def lolos_loan(sid_data):
    if not sid_data.get("has_loan_request", False):
        return False
    loan_labels = {"1a.", "1c.", "2a.", "2b.", "3."}
    return all(
        c["passed"]
        for c in sid_data["checks"]
        if any(c["label"].startswith(p) for p in loan_labels)
    )


# ─────────────────────────────────────────────
# EXCEL EXPORT GENERATORS
# ─────────────────────────────────────────────

def generate_repayment_excel(df_sell, sid_results):
    today_str = datetime.today().strftime("%Y%m%d")

    passed_sids = [
        sid for sid, data in sid_results.items()
        if lolos_repayment(data)
    ]

    sheet1_rows = []
    for sid in passed_sids:
        rows      = df_sell[col(df_sell, SELL_SID).astype(str) == sid]
        total_vol = pd.to_numeric(col(rows, SELL_VOL), errors="coerce").abs().sum()
        if total_vol > 0:
            total_val = pd.to_numeric(col(rows, SELL_VAL), errors="coerce").abs().sum()
            sheet1_rows.append({
                "Participant Code": "EP",
                "SID Client":       sid,
                "Repayment Value":  total_val,
            })

    sheet2_rows = []
    for sid in passed_sids:
        rows = df_sell[col(df_sell, SELL_SID).astype(str) == sid]
        for _, row in rows.iterrows():
            qty = abs(pd.to_numeric(row.iloc[SELL_VOL], errors="coerce") or 0)
            if qty > 0:
                sheet2_rows.append({
                    "SID Client": sid,
                    "Stock Code": str(row.iloc[SELL_STOCK]),
                    "Quantity":   qty,
                })

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(sheet1_rows).to_excel(writer, sheet_name="Repayment Proceed", index=False)
        pd.DataFrame(sheet2_rows).to_excel(writer, sheet_name="Detail Collateral",  index=False)
    buf.seek(0)
    return buf, f"Repayment Proceed {today_str}.xlsx"


def generate_loan_excel(df_buy, sid_results):
    today_str = datetime.today().strftime("%Y%m%d")

    passed_sids = [
        sid for sid, data in sid_results.items()
        if lolos_loan(data)
    ]

    sheet1_rows = []
    for sid in passed_sids:
        rows      = df_buy[col(df_buy, BUY_SID).astype(str) == sid]
        total_vol = pd.to_numeric(col(rows, BUY_VOL), errors="coerce").sum()
        total_val = pd.to_numeric(col(rows, BUY_VAL), errors="coerce").sum()
        if total_vol > 0:
            sheet1_rows.append({
                "Participant Code": "EP",
                "SID Client":       sid,
                "Loan Value":       total_val,
            })

    sheet2_rows = []
    for sid in passed_sids:
        rows = df_buy[col(df_buy, BUY_SID).astype(str) == sid]
        for _, row in rows.iterrows():
            qty = pd.to_numeric(row.iloc[BUY_VOL], errors="coerce") or 0
            if qty > 0:
                sheet2_rows.append({
                    "SID Client": sid,
                    "Stock Code": str(row.iloc[BUY_STOCK]),
                    "Quantity":   qty,
                })

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(sheet1_rows).to_excel(writer, sheet_name="Loan Request",      index=False)
        pd.DataFrame(sheet2_rows).to_excel(writer, sheet_name="Detail Collateral", index=False)
    buf.seek(0)
    return buf, f"Loan Request {today_str}.xlsx"


# ─────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────

st.markdown("# 🔍 Validasi MNC")
st.markdown(
    '<p style="color:#5a6278;font-family:\'IBM Plex Mono\',monospace;font-size:0.85rem;">'
    'Sistem Validasi Repayment &amp; Loan Request</p>',
    unsafe_allow_html=True,
)

st.divider()

# ── Upload section ────────────────────────────────────────────────────
col1, col2, col3 = st.columns(3)

with col1:
    section("📂 File Hasil_MNC")
    hasil_file = st.file_uploader("Upload Hasil_MNC.xlsx", type=["xlsx", "xls"], key="hasil")

with col2:
    section("📄 File OP (.txt)")
    op_file = st.file_uploader("Upload OP_YYYYMMDD_EP.txt", type=["txt"], key="op")

with col3:
    section("📄 File Credit Limit (.txt)")
    cl_file = st.file_uploader("Upload CreditLimit_YYYYMMDD_EP.txt", type=["txt"], key="cl")

st.divider()

# ── Preview file txt setelah upload ──────────────────────────────────
if op_file or cl_file:
    prev_col1, prev_col2 = st.columns(2)

    if op_file:
        with prev_col1:
            section("👁 Preview OP File")
            try:
                op_content_prev = op_file.read().decode("utf-8", errors="replace")
                op_file.seek(0)  # reset agar bisa dibaca lagi saat validasi
                op_prev = parse_op_file(op_content_prev)
                # Tampilkan 3 SID pertama sebagai sampel
                sample_sids = list(op_prev.keys())[:3]
                for s in sample_sids:
                    d = op_prev[s]
                    st.markdown(
                        f'<div class="preview-box">'
                        f'<div class="preview-label">SID: {s}</div>'
                        f'<b>{d["name"]}</b><br>'
                        f'Loan: {fmt_rp(d["loan_existing"])} | '
                        f'Accrued: {fmt_rp(d["accrued_interest"])} | '
                        f'Vol Existing: {d["volume_existing"]:,.0f} lot | '
                        f'Saham: {len(d["stocks"])} kode'
                        f'</div>',
                        unsafe_allow_html=True,
                    )
                st.caption(f"Total {len(op_prev)} nasabah terbaca dari OP file.")
            except Exception as ex:
                st.error(f"Gagal preview OP: {ex}")

    if cl_file:
        with prev_col2:
            section("👁 Preview Credit Limit File")
            try:
                cl_content_prev = cl_file.read().decode("utf-8", errors="replace")
                cl_file.seek(0)  # reset
                cl_prev, vdate = parse_credit_limit_file(cl_content_prev)

                # Cek value date vs hari ini
                today_str_check = datetime.today().strftime("%Y/%m/%d")
                if vdate and vdate != today_str_check:
                    st.markdown(
                        f'<div class="date-warn">⚠️ Value Date file: <b>{vdate}</b> — '
                        f'berbeda dengan hari ini ({today_str_check}). '
                        f'Pastikan file sudah sesuai tanggal transaksi.</div>',
                        unsafe_allow_html=True,
                    )

                sample_sids = list(cl_prev.keys())[:3]
                for s in sample_sids:
                    d = cl_prev[s]
                    st.markdown(
                        f'<div class="preview-box">'
                        f'<div class="preview-label">SID: {s}</div>'
                        f'<b>{d["name"]}</b><br>'
                        f'Available Limit: {fmt_rp(d["available_limit"])}'
                        f'</div>',
                        unsafe_allow_html=True,
                    )
                st.caption(f"Total {len(cl_prev)} nasabah terbaca dari Credit Limit file.")
            except Exception as ex:
                st.error(f"Gagal preview Credit Limit: {ex}")

    st.divider()

section("💼 Credit Limit Partisipan")
credit_limit_partisipan = st.number_input(
    "Masukkan nilai Credit Limit Partisipan (Rp)",
    min_value=0.0,
    value=0.0,
    step=1_000_000.0,
    format="%.2f",
    help="Input manual Credit Limit Partisipan untuk validasi global (check 4)"
)

st.divider()

run_btn = st.button("▶ Jalankan Validasi", use_container_width=True)

# ─────────────────────────────────────────────
# RUN
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

    with st.spinner("Memproses file..."):
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
            df_sell, df_buy, op_data, cl_data, credit_limit_partisipan
        )

    # ── Summary metrics ──────────────────────────────────────────────
    total_sids      = len(sid_results)
    total_pass_rep  = sum(1 for v in sid_results.values() if lolos_repayment(v))
    total_pass_loan = sum(1 for v in sid_results.values() if lolos_loan(v))
    total_pass_all  = sum(
        1 for v in sid_results.values()
        if all(c["passed"] for c in v["checks"])
    )
    global_pass = global_result["passed"]

    st.markdown("### 📊 Ringkasan Hasil Validasi")
    m1, m2, m3, m4, m5 = st.columns(5)

    with m1:
        st.markdown(f'''<div class="metric-card">
            <div class="metric-label">Total Nasabah</div>
            <div class="metric-value" style="color:#1a5fe8">{total_sids}</div>
        </div>''', unsafe_allow_html=True)
    with m2:
        st.markdown(f'''<div class="metric-card">
            <div class="metric-label">Lolos Repayment</div>
            <div class="metric-value" style="color:#22a84a">{total_pass_rep}</div>
        </div>''', unsafe_allow_html=True)
    with m3:
        st.markdown(f'''<div class="metric-card">
            <div class="metric-label">Lolos Loan</div>
            <div class="metric-value" style="color:#22a84a">{total_pass_loan}</div>
        </div>''', unsafe_allow_html=True)
    with m4:
        st.markdown(f'''<div class="metric-card">
            <div class="metric-label">Lolos Semua</div>
            <div class="metric-value" style="color:#f9a825">{total_pass_all}</div>
        </div>''', unsafe_allow_html=True)
    with m5:
        cl_color = "#22a84a" if global_pass else "#e53935"
        cl_label = "LOLOS" if global_pass else "GAGAL"
        st.markdown(f'''<div class="metric-card">
            <div class="metric-label">CL Partisipan</div>
            <div class="metric-value" style="color:{cl_color}">{cl_label}</div>
        </div>''', unsafe_allow_html=True)

    st.divider()

    # ── Validasi 4: Credit Limit Partisipan (global) ─────────────────
    section("VALIDASI 4 — Credit Limit Partisipan (Global)")
    if global_result["passed"]:
        pass_box(f"Credit Limit Partisipan LOLOS — {global_result['detail']}")
    else:
        fail_box(f"Credit Limit Partisipan GAGAL — {global_result['detail']}")

    st.divider()

    # ── Per-SID results ───────────────────────────────────────────────
    section("VALIDASI PER NASABAH")
    st.markdown(
        '<p style="color:#5a6278;font-family:\'IBM Plex Mono\',monospace;font-size:0.75rem;">'
        '🟢 R = Lolos Repayment &nbsp;|&nbsp; 🟢 L = Lolos Loan &nbsp;|&nbsp; '
        '🔴 R = Gagal Repayment &nbsp;|&nbsp; 🔴 L = Gagal Loan</p>',
        unsafe_allow_html=True,
    )

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
                    pass_box(f"{check['label']} &nbsp;|&nbsp; {check['detail']}")
                else:
                    fail_box(f"{check['label']} &nbsp;|&nbsp; {check['detail']}")

    st.divider()

    # ── Export summary table ──────────────────────────────────────────
    section("📥 Export Hasil Validasi")

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

    buf = io.BytesIO()
    df_summary.to_excel(buf, index=False)
    buf.seek(0)
    st.download_button(
        "⬇️ Download Hasil Validasi (.xlsx)",
        data=buf,
        file_name="hasil_validasi.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.divider()

    # ── Export ke Dashboard Kantor ────────────────────────────────────
    section("📤 Export ke Dashboard Kantor")
    st.markdown(
        '<p style="color:#5a6278;font-family:\'IBM Plex Mono\',monospace;font-size:0.8rem;">'
        'Repayment Proceed: lolos 1a + 1b + 1c. &nbsp;|&nbsp; '
        'Loan Request: lolos 1a + 1c + 2a + 2b + 3.</p>',
        unsafe_allow_html=True,
    )

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
