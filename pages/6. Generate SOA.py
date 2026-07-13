"""
IDX Securities Financing - Portfolio Tracker (Streamlit)

Alur pakai harian:
  1. Upload RiskParameter (txt), Closing Price (xlsx), list_invoice (csv) hari ini.
  2. (Opsional) Upload file hasil (template) hari sebelumnya -> histori otomatis disambung.
  3. Cek / edit kolom Tranche (LN) kalau perlu (auto-assign FIFO by default).
  4. Klik "Generate & Download" -> download file Excel baru (1 sheet per client).
     Simpan file ini, upload lagi besok sebagai "template" bareng raw data baru.
"""

import streamlit as st
import pandas as pd
from datetime import date

import engine

st.set_page_config(page_title="IDX Porto Tracker", layout="wide")
st.title("📊 IDX Securities Financing — Portfolio Tracker")
st.caption(
    "Upload raw data harian + template hasil kemarin (opsional) → app hitung ulang "
    "Funding / Outstanding / Interest per client, lalu keluarkan Excel baru untuk di-download."
)

with st.sidebar:
    st.header("1. Upload File Hari Ini")
    risk_param_file = st.file_uploader("RiskParameter (.txt / .csv / .xls / .xlsx)",
                                        type=["txt", "csv", "xls", "xlsx"])
    price_file = st.file_uploader("Closing Price (.xls / .xlsx)", type=["xls", "xlsx"])
    invoice_file = st.file_uploader("List Invoice (.csv / .xls / .xlsx)",
                                     type=["csv", "xls", "xlsx"])

    st.header("2. Template Hasil Kemarin (opsional)")
    st.caption("Kosongkan kalau ini hari pertama / client baru.")
    template_file = st.file_uploader("Template sebelumnya (.xlsx)", type=["xlsx"])

    st.header("3. Tanggal Acuan (as of)")
    as_of = st.date_input("Tanggal untuk hitung bunga berjalan", value=date.today())

process_btn = st.button("🔄 Proses Data", type="primary", use_container_width=True,
                         disabled=not (risk_param_file and price_file and invoice_file))

if "processed" not in st.session_state:
    st.session_state.processed = None

if process_btn:
    try:
        hc_map = engine.parse_risk_parameter(risk_param_file)
        price_map = engine.parse_closing_price(price_file)
        raw_tx = engine.parse_list_invoice(invoice_file, hc_map, price_map)
        new_tx = engine.net_transactions(raw_tx)
        old_history = engine.parse_previous_template(template_file) if template_file else {}

        all_client_ids = sorted(set(new_tx["CLIENT_ID"]) | set(old_history.keys()))

        merged_by_client = {}
        for cid in all_client_ids:
            old_df = old_history.get(cid)
            new_df = new_tx[new_tx["CLIENT_ID"] == cid]
            merged = engine.merge_client_history(old_df, new_df)
            if merged.empty:
                continue
            merged_by_client[cid] = merged

        st.session_state.processed = {
            "hc_map": hc_map,
            "price_map": price_map,
            "merged_by_client": merged_by_client,
            "as_of": as_of,
        }
        st.success(
            f"Berhasil diproses: {len(merged_by_client)} client. "
            f"Netting: {len(raw_tx)} baris raw invoice → {len(new_tx)} baris net posisi harian."
        )
    except Exception as e:
        st.error(f"Gagal memproses file: {e}")
        st.session_state.processed = None

if st.session_state.processed:
    data = st.session_state.processed
    merged_by_client = data["merged_by_client"]
    hc_map = data["hc_map"]
    price_map = data["price_map"]
    as_of = data["as_of"]

    st.divider()
    st.subheader("Review & Edit Tranche (opsional)")
    st.caption(
        "Kolom **LN (TRANCHE)** di-auto-assign FIFO (Sell melunasi tranche tertua yang "
        "masih outstanding, Buy selalu buka tranche baru). Kalau ada kasus yang beda dari "
        "aturan itu, edit langsung di tabel client terkait sebelum generate final."
    )

    client_ids = list(merged_by_client.keys())
    selected_client = st.selectbox("Pilih client untuk review", client_ids)

   raw_df = merged_by_client[selected_client]
    preview_df = engine.assign_tranches(raw_df.copy())
    
    # KUNCI PERBAIKAN: Pastikan kolom MATURITY ada di DataFrame agar tidak KeyError
    if "MATURITY" not in preview_df.columns:
        preview_df["MATURITY"] = ""  # Isi kosong dulu karena nanti diisi rumus otomatis oleh Excel
        
    # Tambahkan MATURITY ke dalam list display
    display_cols = ["TRX_DATE", "DUE_DATE", "MATURITY", "B_S", "STOCK", "VOL", "PRICE",
                    "AMOUNT_TRX", "TRANCHE", "INV_NO"]
    
    # Memastikan tidak ada kolom lain di display_cols yang typo/hilang
    # (Hanya mengambil kolom yang benar-benar ada di preview_df)
    available_cols = [c for c in display_cols if c in preview_df.columns]
    
    edited = st.data_editor(
        preview_df[available_cols],  # Menggunakan kolom yang sudah divalidasi aman
        column_config={
            "TRANCHE": st.column_config.TextColumn("LN (Tranche)", help="Edit manual kalo perlu"),
            "TRX_DATE": st.column_config.DateColumn("TRX DATE", format="YYYY-MM-DD"),
            "DUE_DATE": st.column_config.DateColumn("DUE DATE", format="YYYY-MM-DD"),
            "MATURITY": st.column_config.TextColumn("MATURITY (Auto Excel)", disabled=True),
        },
        disabled=[c for c in available_cols if c != "TRANCHE"],
        use_container_width=True,
        hide_index=True,
        key=f"editor_{selected_client}",
    )

    # simpan balik override tranche ke merged_by_client
    if st.button("💾 Simpan perubahan tranche untuk client ini"):
        raw_df = raw_df.set_index("INV_NO")
        edited_idx = edited.set_index("INV_NO")
        raw_df.loc[edited_idx.index, "TRANCHE"] = edited_idx["TRANCHE"]
        merged_by_client[selected_client] = raw_df.reset_index()[engine.TEMPLATE_COLUMNS]
        st.success("Tersimpan. Lanjut ke Generate kalau sudah selesai review semua client.")

    st.divider()
    st.subheader("Generate Hasil Akhir")

    if st.button("✅ Generate & Download Excel", type="primary"):
        client_results = {}
        for cid, df in merged_by_client.items():
            processed = engine.process_client(df, as_of)
            tranche_summary, stock_pos, portfolio_total, total_outstanding, total_interest = (
                engine.build_recap(processed, hc_map, price_map)
            )
            name = processed["NAME"].dropna().iloc[0] if processed["NAME"].notna().any() else ""
            client_results[cid] = {
                "df": processed,
                "name": name,
                "tranche_summary": tranche_summary,
                "stock_pos": stock_pos,
                "portfolio_total": portfolio_total,
                "total_outstanding": total_outstanding,
                "total_interest": total_interest,
            }

        excel_bytes = engine.write_workbook(client_results, as_of)
        fname = f"DATA_PORTO_{pd.Timestamp(as_of).strftime('%d%m%y')}.xlsx"
        st.download_button(
            "⬇️ Download Excel Hasil Hari Ini",
            data=excel_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

        st.divider()
        st.subheader("Ringkasan Portofolio")
        summary_rows = []
        for cid, res in client_results.items():
            summary_rows.append({
                "Client ID": cid,
                "Nama": res["name"],
                "Total Outstanding": res["total_outstanding"],
                "Total Interest": res["total_interest"],
                "Total Portfolio (IDR-HC)": res["portfolio_total"],
            })
        st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

st.divider()
with st.expander("ℹ️ Catatan & Asumsi"):
    st.markdown("""
- Sumber transaksi **hanya** dari `list_invoice` (baris PORTOFOLIO/initial holding tidak dipakai).
- Bunga dihitung flat **9.5% / tahun**, basis 360 hari, day-weighted-balance per tranche
  (berdasarkan DUE DATE).
- **Tranche (LN)** di-auto-assign pakai aturan FIFO (Sell melunasi tranche tertua yang masih
  outstanding). Ini adalah pendekatan otomatis — kalau ada kasus khusus, edit manual di tabel
  sebelum generate.
- Setiap kali generate, seluruh histori client (lama + baru) dihitung ulang dari awal supaya
  konsisten (bukan cuma baris baru).
- App ini **tidak menyimpan data di server** — file hasil harus kamu download & upload lagi
  besok sebagai "template" bareng raw data baru.
- Dedup transaksi otomatis berdasarkan nomor invoice (`no_inv`), jadi aman kalau file yang sama
  ter-upload dua kali.
""")
