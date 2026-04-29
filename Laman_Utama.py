import streamlit as st

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto+Condensed:wght@300;400;500;600;700&display=swap');

    :root {
        --red:        #D90020;
        --red-dark:   #b0001a;
        --red-light:  #ffeaed;
        --red-border: #f5c0c8;
    }

    html, body, [class*="css"], .stApp,
    .stMarkdown, .stButton, label,
    p, h1, h2, h3, h4, h5, h6, span, div {
        font-family: 'Roboto Condensed', sans-serif !important;
    }

    .stApp {
        background-color: #f8f5f3;
    }

    [data-testid="stSidebar"] {
        background-color: #ffffff !important;
        border-right: 0.5px solid #e8e0db;
    }

    .block-container {
        padding-top: 3rem;
        padding-bottom: 2rem;
        max-width: 820px;
    }

    .eyebrow {
        font-family: 'Roboto Condensed', sans-serif !important;
        font-size: 12px;
        font-weight: 600;
        color: var(--red);
        text-transform: uppercase;
        letter-spacing: 0.12em;
        display: flex;
        align-items: center;
        gap: 8px;
        margin-bottom: 10px;
    }
    .eyebrow::before {
        content: '';
        display: inline-block;
        width: 22px;
        height: 2px;
        background: var(--red);
        flex-shrink: 0;
    }

    .hero-title {
        font-family: 'Roboto Condensed', sans-serif !important;
        font-size: 52px;
        font-weight: 700;
        color: #1a0a0f;
        line-height: 1.1;
        margin-bottom: 6px;
        letter-spacing: -0.01em;
        text-transform: uppercase;
    }
    .hero-title span {
        color: var(--red);
    }

    .hero-divider {
        width: 48px;
        height: 3px;
        background: var(--red);
        margin: 14px 0 20px;
        border-radius: 2px;
    }

    .hero-sub {
        font-family: 'Roboto Condensed', sans-serif !important;
        font-size: 17px;
        font-weight: 300;
        color: #5a4040;
        line-height: 1.6;
        max-width: 520px;
        margin-bottom: 4px;
    }

    .menu-grid {
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 14px;
        margin: 28px 0 32px;
    }
    .menu-card {
        background: #ffffff;
        border: 0.5px solid #e8e0db;
        border-radius: 10px;
        padding: 20px 22px;
        border-left: 3px solid var(--red);
    }
    .menu-card-num {
        font-size: 11px;
        font-weight: 700;
        color: var(--red);
        letter-spacing: 0.08em;
        margin-bottom: 8px;
        text-transform: uppercase;
    }
    .menu-card-title {
        font-family: 'Roboto Condensed', sans-serif !important;
        font-size: 15px;
        font-weight: 700;
        color: #1a0a0f;
        text-transform: uppercase;
        letter-spacing: 0.03em;
        margin-bottom: 5px;
    }
    .menu-card-desc {
        font-family: 'Roboto Condensed', sans-serif !important;
        font-size: 13px;
        font-weight: 300;
        color: #7a5a5a;
        line-height: 1.5;
    }

    .info-banner {
        background: var(--red-light);
        border: 0.5px solid var(--red-border);
        border-left: 3px solid var(--red);
        border-radius: 0 8px 8px 0;
        padding: 14px 18px;
        font-family: 'Roboto Condensed', sans-serif !important;
        font-size: 14px;
        color: #3a1010;
        line-height: 1.55;
    }
    .info-banner strong {
        font-weight: 700;
    }
    </style>
""", unsafe_allow_html=True)

# ── Eyebrow
st.markdown('<div class="eyebrow">PEI I-Fast Converter</div>', unsafe_allow_html=True)

# ── Hero Title
st.markdown("""
    <div class="hero-title">
        Platform Konversi Data<br>
        <span>MNC Sekuritas</span>
    </div>
    <div class="hero-divider"></div>
""", unsafe_allow_html=True)

# ── Subtitle
st.markdown("""
    <div class="hero-sub">
        Kelola dan konversi data investasi Anda dengan cepat, aman, dan akurat.
        Pilih salah satu menu di sidebar kiri untuk memulai.
    </div>
""", unsafe_allow_html=True)

# ── Menu shortcut cards (visual overview, navigasi via sidebar)
st.markdown("""
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
""", unsafe_allow_html=True)

# ── Info Banner
st.markdown("""
    <div class="info-banner">
        <strong>Cara menggunakan:</strong>
        Pilih menu di sidebar kiri untuk mulai mengolah data Anda.
    </div>
""", unsafe_allow_html=True)
