import streamlit as st
import streamlit.components.v1 as components

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto+Condensed:wght@300;400;600;700&display=swap');

    /* Sidebar styling */
    [data-testid="stSidebar"] {
        background-color: #ffffff !important;
        border-right: 1px solid #e8e0db !important;
    }
    [data-testid="stSidebar"] * {
        font-family: 'Roboto Condensed', sans-serif !important;
    }
    [data-testid="stSidebar"] .stMarkdown p {
        font-size: 14px !important;
        color: #3a2020 !important;
    }
    /* Active nav item warna merah */
    [data-testid="stSidebarNavLink"][aria-current="page"] {
        background-color: #fff0f2 !important;
        color: #D90020 !important;
        border-left: 3px solid #D90020 !important;
        font-weight: 700 !important;
    }

    /* Hapus padding atas konten utama */
    .block-container {
        padding-top: 1rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
        max-width: 100% !important;
    }

    /* Sembunyikan header bawaan Streamlit */
    header[data-testid="stHeader"] {
        background: transparent !important;
        height: 0 !important;
    }
    #MainMenu, footer { visibility: hidden; }
    </style>
""", unsafe_allow_html=True)

components.html("""
<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8"/>
<link href="https://fonts.googleapis.com/css2?family=Roboto+Condensed:wght@300;400;600;700&display=swap" rel="stylesheet"/>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  :root {
    --red:        #D90020;
    --red-dark:   #b0001a;
    --red-light:  #fff0f2;
    --red-border: #f5c0c8;
    --bg:         #f8f5f3;
    --white:      #ffffff;
    --text:       #1a0a0f;
    --muted:      #5a4040;
    --subtle:     #9b8a8f;
    --border:     #e8e0db;
  }

  body {
    font-family: 'Roboto Condensed', sans-serif;
    background: transparent;
    color: var(--text);
    padding: 32px 8px 32px 4px;
  }

  /* ── Eyebrow ── */
  .eyebrow {
    display: flex;
    align-items: center;
    gap: 8px;
    font-size: 11px;
    font-weight: 700;
    color: var(--red);
    text-transform: uppercase;
    letter-spacing: 0.14em;
    margin-bottom: 14px;
  }
  .eyebrow::before {
    content: '';
    width: 22px;
    height: 2px;
    background: var(--red);
    display: inline-block;
    flex-shrink: 0;
  }

  /* ── Hero title ── */
  .hero-title {
    font-size: 56px;
    font-weight: 700;
    color: var(--text);
    line-height: 1.0;
    text-transform: uppercase;
    letter-spacing: -0.01em;
  }
  .hero-title .accent { color: var(--red); }

  .hero-divider {
    width: 44px;
    height: 3px;
    background: var(--red);
    border-radius: 2px;
    margin: 18px 0 16px;
  }

  .hero-sub {
    font-size: 16px;
    font-weight: 300;
    color: var(--muted);
    line-height: 1.65;
    max-width: 480px;
    margin-bottom: 36px;
  }

  /* ── Section label ── */
  .section-label {
    font-size: 10px;
    font-weight: 700;
    color: var(--subtle);
    text-transform: uppercase;
    letter-spacing: 0.12em;
    margin-bottom: 12px;
  }

  /* ── Menu cards ── */
  .menu-grid {
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 12px;
    margin-bottom: 28px;
    max-width: 740px;
  }
  .menu-card {
    background: var(--white);
    border: 0.5px solid var(--border);
    border-left: 3px solid var(--red);
    border-radius: 10px;
    padding: 18px 20px;
    transition: transform 0.15s ease, box-shadow 0.15s ease;
    cursor: default;
  }
  .menu-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(217, 0, 32, 0.07);
  }
  .menu-num {
    font-size: 10px;
    font-weight: 700;
    color: var(--red);
    letter-spacing: 0.1em;
    opacity: 0.6;
    margin-bottom: 7px;
  }
  .menu-title {
    font-size: 14px;
    font-weight: 700;
    color: var(--text);
    text-transform: uppercase;
    letter-spacing: 0.03em;
    margin-bottom: 5px;
  }
  .menu-desc {
    font-size: 12.5px;
    font-weight: 300;
    color: #7a5a5a;
    line-height: 1.5;
  }

  /* ── Info banner ── */
  .info-banner {
    max-width: 740px;
    background: var(--red-light);
    border: 0.5px solid var(--red-border);
    border-left: 3px solid var(--red);
    border-radius: 0 8px 8px 0;
    padding: 13px 16px;
    font-size: 13px;
    color: #3a1010;
    line-height: 1.6;
  }
  .info-banner strong { font-weight: 700; }

  /* ── Animations ── */
  @keyframes fadeUp {
    from { opacity: 0; transform: translateY(14px); }
    to   { opacity: 1; transform: translateY(0); }
  }
  .eyebrow     { animation: fadeUp 0.35s ease both 0.05s; }
  .hero-title  { animation: fadeUp 0.35s ease both 0.12s; }
  .hero-divider{ animation: fadeUp 0.35s ease both 0.17s; }
  .hero-sub    { animation: fadeUp 0.35s ease both 0.22s; }
  .section-label{ animation: fadeUp 0.35s ease both 0.27s; }
  .menu-grid   { animation: fadeUp 0.35s ease both 0.31s; }
  .info-banner { animation: fadeUp 0.35s ease both 0.37s; }
</style>
</head>
<body>

  <div class="eyebrow">PEI I-Fast Converter</div>

  <div class="hero-title">
    Platform Konversi Data<br>
    <span class="accent">MNC Sekuritas</span>
  </div>
  <div class="hero-divider"></div>

  <div class="hero-sub">
    Kelola dan konversi data investasi Anda dengan cepat, aman, dan akurat.
    Pilih salah satu fitur di bawah atau gunakan menu di sidebar kiri.
  </div>

  <div class="section-label">Fitur Tersedia</div>

  <div class="menu-grid">
    <div class="menu-card">
      <div class="menu-num">01</div>
      <div class="menu-title">Netting List of Invoice</div>
      <div class="menu-desc">Generate daftar netting invoice dari data transaksi PEI.</div>
    </div>
    <div class="menu-card">
      <div class="menu-num">02</div>
      <div class="menu-title">TRX PEI Details</div>
      <div class="menu-desc">Tampilkan dan ekspor detail transaksi PEI secara terstruktur.</div>
    </div>
    <div class="menu-card">
      <div class="menu-num">03</div>
      <div class="menu-title">Validation LR &amp; RP</div>
      <div class="menu-desc">Validasi data Laporan Rekening dan Rekening Perantara.</div>
    </div>
    <div class="menu-card">
      <div class="menu-num">04</div>
      <div class="menu-title">Ending Balance (SOA)</div>
      <div class="menu-desc">Generate laporan saldo akhir Statement of Account.</div>
    </div>
  </div>

  <div class="info-banner">
    <strong>Cara menggunakan:</strong>
    Pilih menu di sidebar kiri untuk mulai mengolah data Anda.
  </div>

</body>
</html>
""", height=620, scrolling=False)
