import streamlit as st

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto+Condensed:wght@300;400;600;700&display=swap');

    /* Sembunyikan toolbar Streamlit (keyboard_double_arrow, Share, dll) */
    [data-testid="stToolbar"]    { display: none !important; }
    [data-testid="stDecoration"] { display: none !important; }
    #MainMenu                    { display: none !important; }
    footer                       { display: none !important; }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background-color: #ffffff !important;
        border-right: 1px solid #ece4e4 !important;
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

    /* Konten utama */
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 2rem !important;
        max-width: 900px !important;
    }

    /* Seluruh teks pakai Roboto Condensed */
    .stMarkdown, .stMarkdown p, .stMarkdown div {
        font-family: 'Roboto Condensed', sans-serif !important;
    }

    /* ── Topbar strip ── */
    .topbar {
        display: flex;
        align-items: center;
        justify-content: space-between;
        background: #ffffff;
        border: 0.5px solid #ece4e4;
        border-radius: 10px;
        padding: 12px 20px;
        margin-bottom: 32px;
    }
    .topbar-left {
        display: flex;
        align-items: center;
        gap: 10px;
        font-family: 'Roboto Condensed', sans-serif;
    }
    .topbar-logo {
        width: 30px;
        height: 30px;
        background: #D90020;
        border-radius: 6px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        color: white;
        font-size: 14px;
        font-weight: 700;
        letter-spacing: -1px;
        flex-shrink: 0;
    }
    .topbar-brand {
        font-size: 13px;
        font-weight: 700;
        color: #1a0a0f;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        font-family: 'Roboto Condensed', sans-serif;
    }
    .topbar-sep {
        width: 1px;
        height: 14px;
        background: #ddd;
        display: inline-block;
        margin: 0 2px;
    }
    .topbar-sub {
        font-size: 12px;
        font-weight: 400;
        color: #9b8a8f;
        font-family: 'Roboto Condensed', sans-serif;
    }
    .topbar-chip {
        display: inline-flex;
        align-items: center;
        gap: 7px;
        background: #fff0f2;
        border: 0.5px solid #f5c0c8;
        border-radius: 20px;
        padding: 4px 12px 4px 4px;
        font-family: 'Roboto Condensed', sans-serif;
    }
    .topbar-avatar {
        width: 22px;
        height: 22px;
        background: #D90020;
        border-radius: 50%;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-size: 9px;
        font-weight: 700;
        color: white;
    }
    .topbar-username {
        font-size: 12px;
        font-weight: 600;
        color: #D90020;
        font-family: 'Roboto Condensed', sans-serif;
    }

    /* ── Eyebrow ── */
    .eyebrow {
        display: flex;
        align-items: center;
        gap: 8px;
        font-family: 'Roboto Condensed', sans-serif;
        font-size: 11px;
        font-weight: 700;
        color: #D90020;
        text-transform: uppercase;
        letter-spacing: 0.14em;
        margin-bottom: 12px;
    }
    .eyebrow-line {
        display: inline-block;
        width: 22px;
        height: 2px;
        background: #D90020;
        vertical-align: middle;
        margin-right: 2px;
    }

    /* ── Hero title ── */
    .main-title {
        font-family: 'Roboto Condensed', sans-serif !important;
        font-size: 54px;
        font-weight: 700;
        color: #1a0a0f;
        line-height: 1.0;
        text-transform: uppercase;
        letter-spacing: -0.01em;
        margin-bottom: 0;
    }
    .sub-title {
        font-family: 'Roboto Condensed', sans-serif !important;
        font-size: 54px;
        font-weight: 700;
        color: #D90020;
        line-height: 1.0;
        text-transform: uppercase;
        letter-spacing: -0.01em;
        margin-bottom: 0;
    }
    .hero-divider {
        width: 44px;
        height: 3px;
        background: #D90020;
        border-radius: 2px;
        margin: 16px 0 14px;
    }
    .hero-desc {
        font-family: 'Roboto Condensed', sans-serif !important;
        font-size: 15px;
        font-weight: 300;
        color: #5a4040;
        line-height: 1.65;
        max-width: 500px;
        margin-bottom: 32px;
    }

    /* ── Section label ── */
    .section-label {
        font-family: 'Roboto Condensed', sans-serif;
        font-size: 10px;
        font-weight: 700;
        color: #9b8a8f;
        text-transform: uppercase;
        letter-spacing: 0.12em;
        margin-bottom: 12px;
    }

    /* ── Menu cards grid ── */
    .menu-grid {
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 12px;
        margin-bottom: 24px;
    }
    .menu-card {
        background: #ffffff;
        border: 0.5px solid #ece4e4;
        border-left: 3px solid #D90020;
        border-radius: 10px;
        padding: 18px 20px;
        font-family: 'Roboto Condensed', sans-serif;
    }
    .menu-num {
        font-size: 10px;
        font-weight: 700;
        color: #D90020;
        letter-spacing: 0.1em;
        opacity: 0.55;
        margin-bottom: 6px;
    }
    .menu-title {
        font-size: 13px;
        font-weight: 700;
        color: #1a0a0f;
        text-transform: uppercase;
        letter-spacing: 0.03em;
        margin-bottom: 5px;
    }
    .menu-desc {
        font-size: 12px;
        font-weight: 300;
        color: #7a5a5a;
        line-height: 1.5;
    }

    /* ── Info banner ── */
    .info-banner {
        background: #fff0f2;
        border: 0.5px solid #f5c0c8;
        border-left: 3px solid #D90020;
        border-radius: 0 8px 8px 0;
        padding: 12px 16px;
        font-family: 'Roboto Condensed', sans-serif;
        font-size: 13px;
        color: #3a1010;
        line-height: 1.6;
        margin-top: 4px;
    }
    .info-banner strong {
        font-weight: 700;
    }
    </style>
""", unsafe_allow_html=True)

# ── Topbar
st.markdown("""
<div class="topbar">
    <div class="topbar-left">
        <span class="topbar-logo">IF</span>
        <span class="topbar-brand">I-Fast Converter</span>
        <span class="topbar-sep"></span>
        <span class="topbar-sub">MNC Sekuritas</span>
    </div>
    <div class="topbar-chip">
        <span class="topbar-avatar">MN</span>
        <span class="topbar-username">MNC Sekuritas</span>
    </div>
</div>
""", unsafe_allow_html=True)

# ── Eyebrow
st.markdown("""
<div class="eyebrow">
    <span class="eyebrow-line"></span>PEI I-Fast Converter
</div>
""", unsafe_allow_html=True)

# ── Hero Title
st.markdown('<div class="main-title">Platform Konversi Data</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">MNC Sekuritas</div>', unsafe_allow_html=True)
st.markdown('<div class="hero-divider"></div>', unsafe_allow_html=True)

# ── Deskripsi
st.markdown("""
<div class="hero-desc">
    Kelola dan konversi data investasi Anda dengan cepat, aman, dan akurat.
    Pilih salah satu fitur di bawah atau gunakan menu di sidebar kiri.
</div>
""", unsafe_allow_html=True)

# ── Menu Cards
st.markdown('<div class="section-label">Fitur Tersedia</div>', unsafe_allow_html=True)

st.markdown("""
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
""", unsafe_allow_html=True)

# ── Info Banner
st.markdown("""
<div class="info-banner">
    <strong>Cara menggunakan:</strong>
    Pilih menu di sidebar kiri untuk mulai mengolah data Anda.
</div>
""", unsafe_allow_html=True)
