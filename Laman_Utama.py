import streamlit as st
import streamlit.components.v1 as components

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto+Condensed:wght@300;400;600;700&display=swap');

    /* ── Sembunyikan HANYA toolbar atas (keyboard_double_arrow, Share, dll) ── */
    [data-testid="stToolbar"]    { display: none !important; }
    [data-testid="stDecoration"] { display: none !important; }
    #MainMenu                    { display: none !important; }
    footer                       { display: none !important; }

    /* ── Sidebar tetap normal ── */
    [data-testid="stSidebar"] {
        background-color: #ffffff !important;
        border-right: 1px solid #e8e0db !important;
    }
    [data-testid="stSidebar"] * {
        font-family: 'Roboto Condensed', sans-serif !important;
    }
    [data-testid="stSidebarNavLink"][aria-current="page"] {
        background-color: #fff0f2 !important;
        color: #D90020 !important;
        border-left: 3px solid #D90020 !important;
        font-weight: 700 !important;
    }

    /* ── Kurangi padding atas konten agar tidak ada gap besar ── */
    .block-container {
        padding-top: 1rem !important;
        padding-left: 0 !important;
        padding-right: 0 !important;
        max-width: 100% !important;
    }
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
    padding: 0;
    margin: 0;
  }

  /* ── Mini topbar inline ── */
  .mini-topbar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 13px 48px;
    background: var(--white);
    border-bottom: 0.5px solid var(--border);
  }
  .mini-left {
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
    flex-shrink: 0;
  }
  .logo-box svg { width: 15px; height: 15px; fill: #fff; }
  .brand-name {
    font-size: 13px;
    font-weight: 700;
    color: var(--text);
    text-transform: uppercase;
    letter-spacing: 0.04em;
  }
  .brand-divider {
    width: 1px;
    height: 14px;
    background: var(--border);
    margin: 0 2px;
  }
  .brand-sub {
    font-size: 12px;
    font-weight: 400;
    color: var(--subtle);
  }
  .user-chip {
    display: flex;
    align-items: center;
    gap: 7px;
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
    flex-shrink: 0;
  }
  .user-name {
    font-size: 12px;
    font-weight: 600;
    color: var(--red);
  }

  /* ── Main content ── */
  .content { padding: 36px 48px 40px; }

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

  .hero-title {
    font-size: 54px;
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
    margin: 16px 0 14px;
  }

  .hero-sub {
    font-size: 15px;
    font-weight: 300;
    color: var(--muted);
    line-height: 1.65;
    max-width: 520px;
    margin-bottom: 32px;
  }

  .section-label {
    font-size: 10px;
    font-weight: 700;
    color: var(--subtle);
    text-transform: uppercase;
    letter-spacing: 0.12em;
    margin-bottom: 12px;
  }

  .menu-grid {
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 12px;
    margin-bottom: 24px;
  }
  .menu-card {
    background: var(--white);
    border: 0.5px solid var(--border);
    border-left: 3px solid var(--red);
    border-radius: 10px;
    padding: 18px 22px;
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
    margin-bottom: 6px;
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

  .info-banner {
    background: var(--red-light);
    border: 0.5px solid var(--red-border);
    border-left: 3px solid var(--red);
    border-radius: 0 8px 8px 0;
    padding: 12px 16px;
    font-size: 13px;
    color: #3a1010;
    line-height: 1.6;
  }
  .info-banner strong { font-weight: 700; }

  @keyframes fadeUp {
    from { opacity: 0; transform: translateY(12px); }
    to   { opacity: 1; transform: translateY(0); }
  }
  .mini-topbar   { animation: fadeUp 0.3s ease both 0.00s; }
  .eyebrow       { animation: fadeUp 0.3s ease both 0.08s; }
  .hero-title    { animation: fadeUp 0.3s ease both 0.13s; }
  .hero-divider  { animation: fadeUp 0.3s ease both 0.17s; }
  .hero-sub      { animation: fadeUp 0.3s ease both 0.21s; }
  .section-label { animation: fadeUp 0.3s ease both 0.25s; }
  .menu-grid     { animation: fadeUp 0.3s ease both 0.29s; }
  .info-banner   { animation: fadeUp 0.3s ease both 0.34s; }
</style>
</head>
<body>

<div class="mini-topbar">
  <div class="mini-left">
    <div class="logo-box">
      <svg viewBox="0 0 24 24"><path d="M3 3h7v7H3zm11 0h7v7h-7zM3 14h7v7H3zm14 3h-3v-3h3v3zm0 4h-3v-3h3v3zm4-4h-3v-3h3v3zm0 4h-3v-3h3v3z"/></svg>
    </div>
    <span class="brand-name">I-Fast Converter</span>
    <div class="brand-divider"></div>
    <span class="brand-sub">MNC Sekuritas</span>
  </div>
  <div class="user-chip">
    <div class="user-avatar">MN</div>
    <span class="user-name">MNC Sekuritas</span>
  </div>
</div>

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
</div>

</body>
</html>
""", height=660, scrolling=False)
