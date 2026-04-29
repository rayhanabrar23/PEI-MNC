import streamlit as st
import streamlit.components.v1 as components

# Sembunyikan padding default Streamlit agar komponen HTML full-width
st.markdown("""
    <style>
    [data-testid="stSidebar"] {
        background-color: #ffffff !important;
        border-right: 0.5px solid #e8e0db !important;
    }
    [data-testid="stSidebar"] * {
        font-family: 'Roboto Condensed', sans-serif !important;
    }
    .block-container {
        padding: 0 !important;
        max-width: 100% !important;
    }
    #MainMenu, footer, header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

components.html("""
<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<link href="https://fonts.googleapis.com/css2?family=Roboto+Condensed:wght@300;400;500;600;700&display=swap" rel="stylesheet"/>
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
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
  }

  /* ── Topbar ── */
  .topbar {
    background: var(--white);
    border-bottom: 0.5px solid var(--border);
    padding: 0 40px;
    height: 52px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    position: sticky;
    top: 0;
    z-index: 10;
  }
  .topbar-left {
    display: flex;
    align-items: center;
    gap: 10px;
  }
  .logo-box {
    width: 28px;
    height: 28px;
    background: var(--red);
    border-radius: 5px;
    display: flex;
    align-items: center;
    justify-content: center;
  }
  .logo-box svg { width: 16px; height: 16px; fill: #fff; }
  .brand-name {
    font-size: 13px;
    font-weight: 700;
    color: var(--text);
    letter-spacing: 0.02em;
    text-transform: uppercase;
  }
  .brand-sub {
    font-size: 11px;
    font-weight: 400;
    color: var(--subtle);
    margin-left: 2px;
  }
  .topbar-right {
    display: flex;
    align-items: center;
    gap: 12px;
  }
  .user-chip {
    display: flex;
    align-items: center;
    gap: 8px;
    background: var(--red-light);
    border: 0.5px solid var(--red-border);
    border-radius: 20px;
    padding: 4px 12px 4px 4px;
  }
  .user-avatar {
    width: 24px;
    height: 24px;
    background: var(--red);
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 10px;
    font-weight: 700;
    color: #fff;
  }
  .user-name { font-size: 12px; font-weight: 600; color: var(--red); }

  /* ── Main content ── */
  .content {
    padding: 48px 56px 48px;
    max-width: 900px;
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
    letter-spacing: 0.12em;
    margin-bottom: 12px;
  }
  .eyebrow::before {
    content: '';
    display: inline-block;
    width: 22px;
    height: 2px;
    background: var(--red);
    flex-shrink: 0;
  }

  /* ── Hero title ── */
  .hero-title {
    font-size: 54px;
    font-weight: 700;
    color: var(--text);
    line-height: 1.05;
    text-transform: uppercase;
    letter-spacing: -0.01em;
    margin-bottom: 6px;
  }
  .hero-title .accent { color: var(--red); }

  .hero-divider {
    width: 48px;
    height: 3px;
    background: var(--red);
    border-radius: 2px;
    margin: 16px 0 18px;
  }

  .hero-sub {
    font-size: 16px;
    font-weight: 300;
    color: var(--muted);
    line-height: 1.65;
    max-width: 500px;
    margin-bottom: 36px;
  }

  /* ── Menu cards ── */
  .section-label {
    font-size: 11px;
    font-weight: 700;
    color: var(--subtle);
    text-transform: uppercase;
    letter-spacing: 0.1em;
    margin-bottom: 14px;
  }
  .menu-grid {
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 14px;
    margin-bottom: 32px;
  }
  .menu-card {
    background: var(--white);
    border: 0.5px solid var(--border);
    border-left: 3px solid var(--red);
    border-radius: 10px;
    padding: 20px 22px;
    cursor: default;
    transition: transform 0.15s, box-shadow 0.15s;
  }
  .menu-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 16px rgba(217, 0, 32, 0.08);
  }
  .menu-card-num {
    font-size: 10px;
    font-weight: 700;
    color: var(--red);
    letter-spacing: 0.1em;
    text-transform: uppercase;
    margin-bottom: 7px;
    opacity: 0.75;
  }
  .menu-card-title {
    font-size: 14px;
    font-weight: 700;
    color: var(--text);
    text-transform: uppercase;
    letter-spacing: 0.03em;
    margin-bottom: 6px;
  }
  .menu-card-desc {
    font-size: 13px;
    font-weight: 300;
    color: #7a5a5a;
    line-height: 1.5;
  }

  /* ── Info banner ── */
  .info-banner {
    background: var(--red-light);
    border: 0.5px solid var(--red-border);
    border-left: 3px solid var(--red);
    border-radius: 0 8px 8px 0;
    padding: 14px 18px;
    font-size: 13px;
    color: #3a1010;
    line-height: 1.6;
  }
  .info-banner strong { font-weight: 700; }

  /* ── Fade-in animation ── */
  @keyframes fadeUp {
    from { opacity: 0; transform: translateY(12px); }
    to   { opacity: 1; transform: translateY(0); }
  }
  .eyebrow    { animation: fadeUp 0.4s ease both; animation-delay: 0.05s; }
  .hero-title { animation: fadeUp 0.4s ease both; animation-delay: 0.12s; }
  .hero-divider { animation: fadeUp 0.4s ease both; animation-delay: 0.18s; }
  .hero-sub   { animation: fadeUp 0.4s ease both; animation-delay: 0.22s; }
  .section-label { animation: fadeUp 0.4s ease both; animation-delay: 0.28s; }
  .menu-grid  { animation: fadeUp 0.4s ease both; animation-delay: 0.32s; }
  .info-banner { animation: fadeUp 0.4s ease both; animation-delay: 0.38s; }
</style>
</head>
<body>

<!-- Topbar -->
<div class="topbar">
  <div class="topbar-left">
    <div class="logo-box">
      <svg viewBox="0 0 24 24"><path d="M3 3h7v7H3zm11 0h7v7h-7zM3 14h7v7H3zm14 3h-3v-3h3v3zm0 4h-3v-3h3v3zm4-4h-3v-3h3v3zm0 4h-3v-3h3v3z"/></svg>
    </div>
    <span class="brand-name">I-Fast Converter</span>
    <span class="brand-sub">· MNC Sekuritas</span>
  </div>
  <div class="topbar-right">
    <div class="user-chip">
      <div class="user-avatar">MN</div>
      <span class="user-name">MNC Sekuritas</span>
    </div>
  </div>
</div>

<!-- Main Content -->
<div class="content">

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
      <div class="menu-card-num">01</div>
      <div class="menu-card-title">Netting List of Invoice</div>
      <div class="menu-card-desc">Generate daftar netting invoice dari data transaksi PEI.</div>
    </div>
    <div class="menu-card">
      <div class="menu-card-num">02</div>
      <div class="menu-card-title">TRX PEI Details</div>
      <div class="menu-card-desc">Tampilkan dan ekspor detail transaksi PEI secara terstruktur.</div>
    </div>
    <div class="menu-card">
      <div class="menu-card-num">03</div>
      <div class="menu-card-title">Validation LR &amp; RP</div>
      <div class="menu-card-desc">Validasi data Laporan Rekening dan Rekening Perantara.</div>
    </div>
    <div class="menu-card">
      <div class="menu-card-num">04</div>
      <div class="menu-card-title">Ending Balance (SOA)</div>
      <div class="menu-card-desc">Generate laporan saldo akhir Statement of Account.</div>
    </div>
  </div>

  <div class="info-banner">
    <strong>Cara menggunakan:</strong>
    Pilih menu di sidebar kiri untuk mulai mengolah data Anda.
  </div>

</div>
</body>
</html>
""", height=700, scrolling=False)
