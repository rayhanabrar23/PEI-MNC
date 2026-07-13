"""
engine.py
Core logic for IDX Securities Financing Portfolio Tracker.

Pipeline per hari:
  1. Baca RiskParameter (haircut per saham)
  2. Baca Closing Price (harga per saham)
  3. Baca list_invoice (transaksi harian, mentah dari bursa)
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
# 0. UNIVERSAL FILE READER (dukung .csv, .txt pipe-delimited, .xls, .xlsx)
# --------------------------------------------------------------------------

def _detect_ext(file) -> str:
    name = getattr(file, "name", None)
    if name is None:
        name = str(file)
    return name.lower().rsplit(".", 1)[-1] if "." in name else ""


def _read_table(file, prefer_pipe: bool = False) -> pd.DataFrame:
    """Baca file tabel apapun formatnya, deteksi otomatis dari ekstensi:
    - .xls  -> pd.read_excel(engine='xlrd')
    - .xlsx -> pd.read_excel(engine='openpyxl')
    - .txt  -> pipe-delimited (StockCode|StockName|...)
    - .csv  -> comma-delimited
    Kalau file adalah objek upload Streamlit tanpa ekstensi terbaca, coba tebak
    dari isinya (xls/xlsx = binary, sisanya text).
    """
    ext = _detect_ext(file)
    if hasattr(file, "seek"):
        file.seek(0)

    if ext == "xls":
        return pd.read_excel(file, dtype=str, engine="xlrd")
    if ext == "xlsx":
        return pd.read_excel(file, dtype=str, engine="openpyxl")
    if ext == "txt" or prefer_pipe:
        return pd.read_csv(file, sep="|", dtype=str)
    if ext == "csv":
        return pd.read_csv(file, dtype=str)

    # fallback: coba deteksi dari byte pertama (biner Excel vs teks)
    head = file.read(8) if hasattr(file, "read") else b""
    if hasattr(file, "seek"):
        file.seek(0)
    if head[:2] == b"PK":  # xlsx = zip
        return pd.read_excel(file, dtype=str, engine="openpyxl")
    if head[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":  # xls = OLE2
        return pd.read_excel(file, dtype=str, engine="xlrd")
    if prefer_pipe:
        return pd.read_csv(file, sep="|", dtype=str)
    return pd.read_csv(file, dtype=str)


# --------------------------------------------------------------------------
# 1. PARSERS
# --------------------------------------------------------------------------

def parse_risk_parameter(file) -> dict:
    """RiskParameter: StockCode, StockName, Haircut, AvailableQuantity.
    Terima .txt (pipe-delimited), .csv, .xls, atau .xlsx."""
    df = _read_table(file, prefer_pipe=True)
    df.columns = [c.strip() for c in df.columns]
    df["StockCode"] = df["StockCode"].astype(str).str.strip()
    df["Haircut"] = pd.to_numeric(df["Haircut"], errors="coerce").fillna(0)
    return dict(zip(df["StockCode"], df["Haircut"]))


def parse_closing_price(file) -> dict:
    """Closing price. Kolom kunci: no_share (kode saham), kurs_now (harga).
    Terima .xls atau .xlsx."""
    df = _read_table(file)
    df.columns = [str(c).strip() for c in df.columns]
    if "no_share" not in df.columns or "kurs_now" not in df.columns:
        raise ValueError(
            "File closing price harus punya kolom 'no_share' dan 'kurs_now'."
        )
    df["no_share"] = df["no_share"].astype(str).str.strip()
    df["kurs_now"] = pd.to_numeric(df["kurs_now"], errors="coerce")
    return dict(zip(df["no_share"], df["kurs_now"]))


def parse_list_invoice(file, hc_map: dict, price_map: dict) -> pd.DataFrame:
    """list_invoice -> baris transaksi ternormalisasi (kolom internal).
    Terima .csv, .xls, atau .xlsx."""
    df = _read_table(file)
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
# 2. NETTING (per client + saham + tanggal -> 1 baris hasil akhir bersih)
# --------------------------------------------------------------------------

def net_transactions(df: pd.DataFrame) -> pd.DataFrame:
    """Netting harian: semua transaksi Buy & Sell untuk kombinasi
    (client, saham, tanggal) yang sama digabung jadi SATU baris posisi akhir.
    - Kalau hasil akhirnya net Buy -> 1 baris B dengan volume net-nya.
    - Kalau hasil akhirnya net Sell -> 1 baris S dengan volume net-nya.
    - Kalau net = 0 (saling menutup persis) -> baris tidak muncul sama sekali.
    Jadi tidak akan ada 1 saham tercatat 2x di hari yang sama untuk client yang sama.
    """
    if df.empty:
        return df

    out_rows = []
    for (cid, stock, trx_date), g in df.groupby(["CLIENT_ID", "STOCK", "TRX_DATE"], dropna=False):
        net_vol = g["VOL"].sum()
        if round(net_vol) == 0:
            continue  # net habis, tidak dicatat
        net_amt = g["AMOUNT_TRX"].sum()
        b_s = "B" if net_vol > 0 else "S"

        due_dates = g["DUE_DATE"].dropna()
        due_date = due_dates.mode().iloc[0] if not due_dates.empty else pd.NaT

        hc = g["HC"].iloc[0]
        price = g["PRICE"].iloc[0]
        collateral = net_vol * price * (100 - hc) / 100
        inv_no = ",".join(sorted(set(g["INV_NO"].astype(str))))

        out_rows.append({
            "CLIENT_ID": cid,
            "NAME": g["NAME"].iloc[0],
            "B_S": b_s,
            "TRX_DATE": trx_date,
            "DUE_DATE": due_date,
            "ACTIVITY": "TRX",
            "STOCK": stock,
            "HC": hc,
            "VOL": net_vol,
            "PRICE": price,
            "COLLATERAL_IDR_HC": collateral,
            "AMOUNT_TRX": net_amt,
            "TRANCHE": "",
            "FUNDING": net_amt,
            "OUTSTANDING": 0.0,
            "INTEREST": 0.0,
            "RATIO": "",
            "INV_NO": inv_no,
        })

    if not out_rows:
        return pd.DataFrame(columns=TEMPLATE_COLUMNS)

    result = pd.DataFrame(out_rows)
    return result.sort_values(["TRX_DATE", "CLIENT_ID", "STOCK"]).reset_index(drop=True)[TEMPLATE_COLUMNS]


# --------------------------------------------------------------------------
# 3. MERGE (histori lama + transaksi baru), dedup by INV_NO
# --------------------------------------------------------------------------

def _inv_no_set(inv_no) -> set:
    """Satu baris (setelah netting) bisa mewakili beberapa no invoice
    yang digabung jadi string 'inv1,inv2,inv3'. Pecah jadi set untuk dedup."""
    if not inv_no or (isinstance(inv_no, float) and pd.isna(inv_no)):
        return set()
    return set(str(inv_no).split(","))


def merge_client_history(old_df: pd.DataFrame | None, new_df: pd.DataFrame) -> pd.DataFrame:
    if old_df is None or old_df.empty:
        combined = new_df.copy()
    else:
        existing_invs: set = set()
        for s in old_df["INV_NO"].astype(str):
            existing_invs |= _inv_no_set(s)

        def is_new(inv_no):
            return len(_inv_no_set(inv_no) & existing_invs) == 0

        add_df = new_df[new_df["INV_NO"].apply(is_new)]
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
    """number_cols/date_cols merujuk ke NAMA KOLOM ASLI df (bukan display header)."""
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
            number_cols=["HC", "VOL", "PRICE", "COLLATERAL_IDR_HC", "AMOUNT_TRX",
                         "FUNDING", "OUTSTANDING", "INTEREST"],
            date_cols=["TRX_DATE", "DUE_DATE"],
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
