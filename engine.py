"""
engine.py
Core logic for IDX Securities Financing Portfolio Tracker.

Pipeline per hari:
  1. Baca RiskParameter (haircut per saham) - file .xls
  2. Baca Closing Price (harga per saham) - file .xls
  3. Baca list_invoice (transaksi harian, mentah dari bursa) - file .xls
  4. Baca template hasil hari sebelumnya (opsional, kosong kalau hari pertama)
  5. Gabungkan transaksi baru ke histori tiap client (append, bukan replace)
  6. Auto-assign Tranche/LN (FIFO: Sell melunasi tranche tertua yang masih outstanding;
     Buy selalu membuka tranche baru). Bisa dioverride manual lewat UI sebelum final export.
  7. Hitung ulang FUNDING / OUTSTANDING / INTEREST untuk SELURUH histori tiap client
     (bukan cuma baris baru), supaya konsisten kalau ada revisi di tengah.
  8. Bangun recap per sheet: PELUNASAN FUNDING per tranche + ringkasan collateral saham (posisi saat ini).
  9. Tulis workbook Excel, 1 sheet per client.

Catatan asumsi penting (didiskusikan & dikonfirmasi dengan user):
  - Baris "PORTOFOLIO" (initial holding) di-skip; sumber transaksi HANYA dari list_invoice.
  - Bunga flat 9.5% / tahun, basis 360 hari.
  - Tranche/LN adalah keputusan bisnis yang aslinya diisi manual di file sumber (bukan formula
    Excel) -> di sini di-auto-assign dengan aturan FIFO, tapi tetap bisa dikoreksi manual di UI.
  - Nilai FUNDING per transaksi = nilai transaksi (amt_done) dari list_invoice, bertanda
    (+) untuk Buy dan (-) untuk Sell.

Catatan format file mentah (dikonfirmasi dari sample riil 08/07/26):
  - RiskParameter, Closing Price, dan List Invoice SEMUA berupa file Excel biner lama
    (.xls, "Composite Document File V2"), BUKAN txt pipe-delimited / csv. Ketiganya dibaca
    dengan pandas.read_excel (butuh package `xlrd` terpasang untuk format .xls lama).
  - RiskParameter kolom: StockCode, StockName, Haircut, AvailableQuantity.
  - Closing Price kolom kunci: no_share, kurs_now (ada banyak kolom lain yang diabaikan).
  - List Invoice kolom kunci: dt_inv, no_inv, no_cust, name, bors, no_share, tot_vol, rate,
    amt_done, dt_due (ada banyak kolom lain yang diabaikan, mis. SID, KSEI01, npwp, dll).
  - dt_inv & dt_due berbentuk string "DD/MM/YYYY HH:MM:SS" -> parse dengan dayfirst=True.
"""

from __future__ import annotations
import io
import re
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

RATE = 0.095          # bunga tahunan flat 9.5% (dikonfirmasi user)
DAY_BASIS = 360        # basis hari

TEMPLATE_COLUMNS = [
    "CLIENT_ID", "NAME", "B_S", "TRX_DATE", "DUE_DATE", "ACTIVITY", "STOCK",
    "HC", "VOL", "PRICE", "COLLATERAL_IDR_HC", "AMOUNT_TRX",
    "TRANCHE", "FUNDING", "OUTSTANDING", "INTEREST", "RATIO", "INV_NO",
]

DISPLAY_HEADERS = {
    "CLIENT_ID": "CLIENT ID",
    "NAME": "NAME",
    "B_S": "B/S",
    "TRX_DATE": "TRX DATE",
    "DUE_DATE": "DUE DATE",
    "ACTIVITY": "ACTIVITY",
    "STOCK": "STOCK",
    "HC": "HC",
    "VOL": "COLLATERAL (VOL)",
    "PRICE": "PRICE",
    "COLLATERAL_IDR_HC": "COLLATERAL (IDR-HC)",
    "AMOUNT_TRX": "AMOUNT TRX",
    "TRANCHE": "LN (TRANCHE)",
    "FUNDING": "FUNDING",
    "OUTSTANDING": "OUTSTANDING",
    "INTEREST": "INTEREST",
    "RATIO": "RATIO",
    "INV_NO": "INV NO",
}


# --------------------------------------------------------------------------
# 1. PARSERS
# --------------------------------------------------------------------------

def parse_risk_parameter(file) -> dict:
    """RiskParameter: file Excel (.xls) dengan kolom
    StockCode | StockName | Haircut | AvailableQuantity."""
    df = pd.read_excel(file, dtype=str)
    df.columns = [c.strip() for c in df.columns]
    required = {"StockCode", "Haircut"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Kolom wajib hilang di RiskParameter: {sorted(missing)}")
    df["StockCode"] = df["StockCode"].astype(str).str.strip()
    df["Haircut"] = pd.to_numeric(df["Haircut"], errors="coerce").fillna(0)
    return dict(zip(df["StockCode"], df["Haircut"]))


def parse_closing_price(file) -> dict:
    """Closing price: file Excel (.xls). Kolom kunci: no_share (kode saham),
    kurs_now (harga). Kolom lain (descr, isin, dll) diabaikan."""
    df = pd.read_excel(file)
    df.columns = [str(c).strip() for c in df.columns]
    if "no_share" not in df.columns or "kurs_now" not in df.columns:
        raise ValueError(
            "File closing price harus punya kolom 'no_share' dan 'kurs_now'."
        )
    df["no_share"] = df["no_share"].astype(str).str.strip()
    df["kurs_now"] = pd.to_numeric(df["kurs_now"], errors="coerce")
    return dict(zip(df["no_share"], df["kurs_now"]))


def parse_list_invoice(file, hc_map: dict, price_map: dict) -> pd.DataFrame:
    """list_invoice: file Excel (.xls) -> baris transaksi ternormalisasi (kolom internal).
    Kolom sumber yang dipakai: dt_inv, no_inv, no_cust, name, bors, no_share, tot_vol,
    rate, amt_done, dt_due. Kolom lain (SID, KSEI01, npwp, board, dll) diabaikan."""
    df = pd.read_excel(file, dtype=str)
    df.columns = [c.strip() for c in df.columns]

    required = {"dt_inv", "no_cust", "name", "bors", "no_share", "tot_vol",
                "rate", "amt_done", "dt_due", "no_inv"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Kolom wajib hilang di list_invoice: {sorted(missing)}")

    out = pd.DataFrame()
    out["CLIENT_ID"] = df["no_cust"].astype(str).str.strip()
    out["NAME"] = df["name"].astype(str).str.strip()
    out["B_S"] = df["bors"].astype(str).str.strip().str.upper()
    # dt_inv / dt_due datang sebagai teks "DD/MM/YYYY HH:MM:SS" dari file .xls sumber
    out["TRX_DATE"] = pd.to_datetime(df["dt_inv"], dayfirst=True, errors="coerce")
    out["DUE_DATE"] = pd.to_datetime(df["dt_due"], dayfirst=True, errors="coerce")
    out["ACTIVITY"] = "TRX"
    out["STOCK"] = df["no_share"].astype(str).str.strip()

    sign = out["B_S"].map({"B": 1, "S": -1}).fillna(1)

    out["HC"] = out["STOCK"].map(hc_map).fillna(0.0)
    out["VOL"] = pd.to_numeric(df["tot_vol"], errors="coerce").fillna(0.0) * sign

    price_from_file = out["STOCK"].map(price_map)
    price_fallback = pd.to_numeric(df["rate"], errors="coerce")
    out["PRICE"] = price_from_file.fillna(price_fallback)

    out["COLLATERAL_IDR_HC"] = out["VOL"] * out["PRICE"] * (100 - out["HC"]) / 100
    amt = pd.to_numeric(df["amt_done"], errors="coerce").fillna(0.0)
    out["AMOUNT_TRX"] = amt * sign

    out["TRANCHE"] = ""
    out["FUNDING"] = out["AMOUNT_TRX"]
    out["OUTSTANDING"] = 0.0
    out["INTEREST"] = 0.0
    out["RATIO"] = ""
    out["INV_NO"] = df["no_inv"].astype(str).str.strip()

    return out[TEMPLATE_COLUMNS]


def parse_previous_template(file) -> dict[str, pd.DataFrame]:
    """Baca workbook hasil hari sebelumnya (format yang dihasilkan app ini sendiri).
    Return dict: {client_id: DataFrame histori transaksi (tanpa recap)}.
    Kalau file kosong/None, return {}.
    """
    if file is None:
        return {}

    wb = load_workbook(file, data_only=True)
    result: dict[str, pd.DataFrame] = {}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue

        header_row_idx = None
        for i, r in enumerate(rows):
            cells = [str(c).strip() if c is not None else "" for c in r]
            if "CLIENT ID" in cells:
                header_row_idx = i
                break
        if header_row_idx is None:
            continue  # bukan sheet transaksi

        header = [str(h).strip() if h is not None else "" for h in rows[header_row_idx]]
        col_idx = {h: i for i, h in enumerate(header)}

        data_rows = []
        for r in rows[header_row_idx + 1:]:
            ci = col_idx.get("CLIENT ID", 0)
            client_id = r[ci] if len(r) > ci else None
            if client_id in (None, ""):
                break  # ketemu baris kosong -> akhir tabel transaksi, sisanya recap
            data_rows.append(r)

        if not data_rows:
            continue

        rec = []
        for r in data_rows:
            def g(col):
                i = col_idx.get(col)
                return r[i] if i is not None and i < len(r) else None

            rec.append({
                "CLIENT_ID": g("CLIENT ID"),
                "NAME": g("NAME"),
                "B_S": g("B/S"),
                "TRX_DATE": g("TRX DATE"),
                "DUE_DATE": g("DUE DATE"),
                "ACTIVITY": g("ACTIVITY"),
                "STOCK": g("STOCK"),
                "HC": g("HC"),
                "VOL": g("COLLATERAL (VOL)"),
                "PRICE": g("PRICE"),
                "COLLATERAL_IDR_HC": g("COLLATERAL (IDR-HC)"),
                "AMOUNT_TRX": g("AMOUNT TRX"),
                "TRANCHE": g("LN (TRANCHE)"),
                "FUNDING": g("FUNDING"),
                "OUTSTANDING": g("OUTSTANDING"),
                "INTEREST": g("INTEREST"),
                "RATIO": g("RATIO"),
                "INV_NO": g("INV NO"),
            })
        cdf = pd.DataFrame(rec)
        cdf["TRX_DATE"] = pd.to_datetime(cdf["TRX_DATE"], errors="coerce")
        cdf["DUE_DATE"] = pd.to_datetime(cdf["DUE_DATE"], errors="coerce")
        for c in ["HC", "VOL", "PRICE", "COLLATERAL_IDR_HC", "AMOUNT_TRX",
                  "FUNDING", "OUTSTANDING", "INTEREST"]:
            cdf[c] = pd.to_numeric(cdf[c], errors="coerce").fillna(0.0)
        cdf["TRANCHE"] = cdf["TRANCHE"].fillna("").astype(str)
        cdf["RATIO"] = cdf["RATIO"].fillna("").astype(str)
        cdf["INV_NO"] = cdf["INV_NO"].fillna("").astype(str)
        result[str(client_id_from_sheet(sheet_name, cdf))] = cdf[TEMPLATE_COLUMNS]

    return result


def client_id_from_sheet(sheet_name: str, cdf: pd.DataFrame) -> str:
    if len(cdf) and cdf["CLIENT_ID"].iloc[0]:
        return str(cdf["CLIENT_ID"].iloc[0])
    return sheet_name


# --------------------------------------------------------------------------
# 2. MERGE (histori lama + transaksi baru), dedup by INV_NO
# --------------------------------------------------------------------------

def merge_client_history(old_df: pd.DataFrame | None, new_df: pd.DataFrame) -> pd.DataFrame:
    if old_df is None or old_df.empty:
        combined = new_df.copy()
    else:
        existing_inv = set(old_df["INV_NO"].astype(str))
        add_df = new_df[~new_df["INV_NO"].astype(str).isin(existing_inv)]
        combined = pd.concat([old_df, add_df], ignore_index=True)
    combined = combined.sort_values(["TRX_DATE", "DUE_DATE", "INV_NO"]).reset_index(drop=True)
    return combined


# --------------------------------------------------------------------------
# 3. TRANCHE AUTO-ASSIGN (FIFO) + FUNDING/OUTSTANDING/INTEREST
# --------------------------------------------------------------------------

def _next_letter(used: list[str]) -> str:
    n = len(used)
    if n < 26:
        return chr(ord("A") + n)
    # setelah Z -> AA, AB, ...
    first = (n // 26) - 1
    second = n % 26
    return chr(ord("A") + first) + chr(ord("A") + second)


def assign_tranches(df: pd.DataFrame) -> pd.DataFrame:
    """Auto-assign kolom TRANCHE untuk baris yang belum punya tranche (baru).
    Baris yang sudah punya TRANCHE (dari histori sebelumnya / edit manual user) dibiarkan.
    Aturan: Buy -> buka tranche baru. Sell -> FIFO ke tranche TERTUA yang outstanding-nya > 0.
    """
    df = df.sort_values(["TRX_DATE", "DUE_DATE", "INV_NO"]).reset_index(drop=True)
    used_letters: list[str] = []
    balance: dict[str, float] = {}

    # pass awal: kumpulkan tranche yang sudah ada (dari histori / manual override)
    for t in df["TRANCHE"]:
        if isinstance(t, str) and t.strip() and t.strip() not in used_letters:
            used_letters.append(t.strip())
            balance[t.strip()] = 0.0

    for idx, row in df.iterrows():
        tranche = str(row["TRANCHE"]).strip() if row["TRANCHE"] else ""
        if not tranche:
            if str(row["B_S"]).upper() == "B":
                tranche = _next_letter(used_letters)
                used_letters.append(tranche)
                balance[tranche] = 0.0
            else:
                tranche = None
                for L in used_letters:
                    if balance.get(L, 0.0) > 0.01:
                        tranche = L
                        break
                if tranche is None:
                    if used_letters:
                        tranche = used_letters[-1]
                    else:
                        tranche = _next_letter(used_letters)
                        used_letters.append(tranche)
                        balance[tranche] = 0.0
            df.at[idx, "TRANCHE"] = tranche
        balance[tranche] = balance.get(tranche, 0.0) + row["AMOUNT_TRX"]

    return df


def compute_funding_outstanding(df: pd.DataFrame) -> pd.DataFrame:
    """FUNDING = AMOUNT_TRX apa adanya. OUTSTANDING = saldo kumulatif per tranche
    (urut kronologis), sesuai urutan TRX_DATE."""
    df = df.sort_values(["TRX_DATE", "DUE_DATE", "INV_NO"]).reset_index(drop=True)
    df["FUNDING"] = df["AMOUNT_TRX"]
    df["OUTSTANDING"] = df.groupby("TRANCHE")["FUNDING"].cumsum()
    return df


def compute_interest(df: pd.DataFrame, as_of_date) -> pd.DataFrame:
    """Bunga day-weighted-balance per tranche, basis DUE_DATE.
    Interest baris ke-i = (hari sampai baris berikutnya di tranche yg sama, atau
    as_of_date kalau baris terakhir) x OUTSTANDING baris ke-i x rate/360.
    """
    df = df.copy()
    df["INTEREST"] = 0.0
    as_of_date = pd.Timestamp(as_of_date)

    for tranche, sub in df.groupby("TRANCHE"):
        sub = sub.sort_values(["DUE_DATE", "INV_NO"])
        idxs = sub.index.tolist()
        dates = sub["DUE_DATE"].tolist()
        outstandings = sub["OUTSTANDING"].tolist()
        for k, ix in enumerate(idxs):
            start = dates[k]
            end = dates[k + 1] if k + 1 < len(idxs) else as_of_date
            if pd.isna(start):
                continue
            if pd.isna(end) or end < start:
                end = max(start, as_of_date)
            days = (end - start).days
            if days < 0:
                days = 0
            bal = outstandings[k]
            df.at[ix, "INTEREST"] = days * bal * RATE / DAY_BASIS
    return df


def compute_ratio_flag(df: pd.DataFrame) -> pd.DataFrame:
    """Tandai SHORT kalau posisi bersih saham di client tsb negatif saat baris Sell terjadi."""
    df = df.copy()
    net_by_stock = df.groupby("STOCK")["VOL"].sum()

    def flag(row):
        if str(row["B_S"]).upper() == "S" and net_by_stock.get(row["STOCK"], 0) < 0:
            return "SHORT"
        return ""

    df["RATIO"] = df.apply(flag, axis=1)
    return df


def process_client(df: pd.DataFrame, as_of_date) -> pd.DataFrame:
    df = assign_tranches(df)
    df = compute_funding_outstanding(df)
    df = compute_interest(df, as_of_date)
    df = compute_ratio_flag(df)
    return df.sort_values(["TRX_DATE", "DUE_DATE", "INV_NO"]).reset_index(drop=True)


# --------------------------------------------------------------------------
# 4. RECAP (per tranche + ringkasan collateral saham posisi saat ini)
# --------------------------------------------------------------------------

def build_recap(df: pd.DataFrame, hc_map: dict, price_map: dict):
    tranche_summary = (
        df.groupby("TRANCHE")
        .agg(TOTAL_FUNDING=("FUNDING", "sum"),
             TOTAL_INTEREST=("INTEREST", "sum"),
             LAST_DUE=("DUE_DATE", "max"))
        .reset_index()
        .sort_values("TRANCHE")
    )

    stock_pos = df.groupby("STOCK")["VOL"].sum().reset_index()
    stock_pos = stock_pos[stock_pos["VOL"].round(0) != 0]
    stock_pos["HC"] = stock_pos["STOCK"].map(hc_map).fillna(0.0)
    stock_pos["PRICE"] = stock_pos["STOCK"].map(price_map).fillna(0.0)
    stock_pos["COLLATERAL_IDR_HC"] = (
        stock_pos["VOL"] * stock_pos["PRICE"] * (100 - stock_pos["HC"]) / 100
    )
    stock_pos = stock_pos.sort_values("STOCK")
    portfolio_total = stock_pos["COLLATERAL_IDR_HC"].sum()
    total_outstanding = df.groupby("TRANCHE")["OUTSTANDING"].last().sum()
    total_interest = df["INTEREST"].sum()

    return tranche_summary, stock_pos, portfolio_total, total_outstanding, total_interest


# --------------------------------------------------------------------------
# 5. WRITE EXCEL
# --------------------------------------------------------------------------

THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
BOLD = Font(bold=True)


def _write_df(ws, start_row, df, headers, number_cols=(), date_cols=()):
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=start_row, column=j, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal="center")
    r = start_row + 1
    for _, row in df.iterrows():
        for j, col in enumerate(df.columns, start=1):
            v = row[col]
            if col in date_cols and pd.notna(v):
                v = pd.Timestamp(v).to_pydatetime()
            cell = ws.cell(row=r, column=j, value=(None if pd.isna(v) else v))
            cell.border = BORDER
            if col in number_cols:
                cell.number_format = "#,##0"
            if col in date_cols:
                cell.number_format = "dd-mmm-yy"
        r += 1
    return r  # baris berikutnya yang kosong


def write_workbook(client_results: dict[str, dict], as_of_date) -> bytes:
    """client_results: {client_id: {"df":..., "name":..., "tranche_summary":...,
    "stock_pos":..., "portfolio_total":..., "total_outstanding":..., "total_interest":...}}
    """
    wb = Workbook()
    wb.remove(wb.active)

    for client_id, res in client_results.items():
        df = res["df"]
        name = res["name"]
        sheet_name = re.sub(r"[\\/*?:\[\]]", "_", str(client_id))[:31]
        ws = wb.create_sheet(sheet_name)

        ws.cell(row=1, column=1, value=str(client_id)).font = BOLD
        ws.cell(row=1, column=4, value=str(name)).font = BOLD
        ws.cell(row=2, column=1, value=f"As of: {pd.Timestamp(as_of_date).strftime('%d-%b-%Y')}")

        disp = df.copy()
        disp = disp[TEMPLATE_COLUMNS]
        headers = [DISPLAY_HEADERS[c] for c in TEMPLATE_COLUMNS]
        next_row = _write_df(
            ws, start_row=4, df=disp, headers=headers,
            number_cols=[DISPLAY_HEADERS[c] for c in
                         ["HC", "VOL", "PRICE", "COLLATERAL_IDR_HC", "AMOUNT_TRX",
                          "FUNDING", "OUTSTANDING", "INTEREST"]],
            date_cols=[DISPLAY_HEADERS[c] for c in ["TRX_DATE", "DUE_DATE"]],
        )

        # ---- Ringkasan total ----
        r = next_row + 1
        ws.cell(row=r, column=1, value="TOTAL OUTSTANDING").font = BOLD
        ws.cell(row=r, column=2, value=res["total_outstanding"]).number_format = "#,##0"
        r += 1
        ws.cell(row=r, column=1, value="TOTAL INTEREST").font = BOLD
        ws.cell(row=r, column=2, value=res["total_interest"]).number_format = "#,##0"
        r += 2

        # ---- Recap: PELUNASAN FUNDING per tranche ----
        ws.cell(row=r, column=1, value="PELUNASAN FUNDING (per tranche)").font = BOLD
        r += 1
        ts = res["tranche_summary"].rename(columns={
            "TRANCHE": "Tranche", "TOTAL_FUNDING": "Total Funding",
            "TOTAL_INTEREST": "Total Interest", "LAST_DUE": "Last Due Date",
        })
        r = _write_df(
            ws, start_row=r, df=ts,
            headers=list(ts.columns),
            number_cols=["Total Funding", "Total Interest"],
            date_cols=["Last Due Date"],
        )
        r += 1

        # ---- Recap: Stock collateral summary ----
        ws.cell(row=r, column=1, value="COLLATERAL SAHAM (posisi saat ini)").font = BOLD
        r += 1
        sp = res["stock_pos"].rename(columns={
            "STOCK": "Stock", "HC": "HC", "VOL": "Collateral (Vol)",
            "PRICE": "Price", "COLLATERAL_IDR_HC": "Collateral (IDR-HC)",
        })
        r = _write_df(
            ws, start_row=r, df=sp,
            headers=list(sp.columns),
            number_cols=["Collateral (Vol)", "Price", "Collateral (IDR-HC)"],
        )
        ws.cell(row=r, column=1, value="PORTOFOLIO").font = BOLD
        ws.cell(row=r, column=5, value=res["portfolio_total"]).number_format = "#,##0"
        ws.cell(row=r, column=5).font = BOLD

        # lebar kolom
        for col_cells in ws.columns:
            length = max((len(str(c.value)) for c in col_cells if c.value is not None), default=8)
            col_letter = get_column_letter(col_cells[0].column)
            ws.column_dimensions[col_letter].width = min(max(length + 2, 10), 32)

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()
