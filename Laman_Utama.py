import streamlit as st

st.set_page_config(
    page_title="PEI I-Fast Converter | MNC Sekuritas",
    page_icon="📊",
    layout="wide",
)

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto+Condensed:ital,wght@0,300;0,400;0,500;0,600;0,700;1,300;1,400&display=swap');

    /* ── CSS Variables ── */
    :root {
        --red:        #D90020;
        --red-dark:   #b0001a;
        --red-light:  #ffeaed;
        --red-border: #f5c0c8;
    }

    /* ── Global font ── */
    html, body, [class*="css"], .stApp,
    .stMarkdown, .stButton, .stSelectbox,
    .stTextInput, .stFileUploader, label,
    p, h1, h2, h3, h4, h5, h6, span, div {
        font-family: 'Roboto Condensed', sans-serif !important;
    }

    /* ── App background ── */
    .stApp {
        background-color: #f8f5f3;
    }

    /* ── Sidebar ── */
    [data-testid="stSidebar"] {
        background-color: #ffffff !important;
        border-right: 0.5px solid #e8e0db;
    }
    [data-testid="stSidebar"] .stMarkdown p,
    [data-testid="stSidebar"] span,
    [data-testid="stSidebar"] div {
        font-family: 'Roboto Condensed', sans-serif !important;
        font-size: 14px;
        color: #3a2020;
    }

    /* ── Main content block ── */
    .block-container {
        padding-top: 2.5rem;
        padding-bottom: 2rem;
        max-width: 860px;
    }

    /* ── Eyebrow label ── */
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

    /* ── Hero title ── */
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

    /* ── Subtitle ── */
    .hero-sub {
        font-family: 'Roboto Condensed', sans-serif !important;
        font-size: 17px;
        font-weight: 300;
        color: #5a4040;
        line-height: 1.6;
        max-width: 520px;
        margin-bottom: 28px;
    }

    /* ── Divider line ── */
    .hero-divider {
        width: 48px;
        height: 3px;
        background: var(--red);
        margin: 14px 0 20px;
        border-radius: 2px;
    }

    /* ── Stat cards ── */
    .stat-row {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 14px;
        margin: 28px 0;
    }
    .stat-card {
        background: #ffffff;
        border: 0.5px solid #e8e0db;
        border-radius: 10px;
        padding: 18px 22px;
    }
    .stat-card.accent {
        background: var(--red);
        border-color: var(--red);
    }
    .stat-label {
        font-family: 'Roboto Condensed', sans-serif !important;
        font-size: 11px;
        font-weight: 400;
        color: #9b8a8f;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        margin-bottom: 6px;
    }
    .stat-card.accent .stat-label {
        color: rgba(255, 255, 255, 0.65);
    }
    .stat-value {
        font-family: 'Roboto Condensed', sans-serif !important;
        font-size: 28px;
        font-weight: 700;
        color: #1a0a0f;
        line-height: 1;
        letter-spacing: -0.01em;
    }
    .stat-card.accent .stat-value {
        color: #ffffff;
    }
    .stat-change {
        font-family: 'Roboto Condensed', sans-serif !important;
        font-size: 12px;
        color: #6b8a6e;
        margin-top: 5px;
    }
    .stat-card.accent .stat-change {
        color: rgba(255, 255, 255, 0.5);
    }

    /* ── Info banner ── */
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

    /* ── Primary button ── */
    .stButton > button {
        background-color: var(--red) !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 6px !important;
        font-family: 'Roboto Condensed', sans-serif !important;
        font-size: 14px !important;
        font-weight: 600 !important;
        letter-spacing: 0.04em !important;
        text-transform: uppercase !important;
        padding: 0.5rem 1.5rem !important;
        transition: background 0.15s ease !important;
    }
    .stButton > button:hover {
        background-color: var(--red-dark) !important;
    }
    .stButton > button:active {
        background-color: var(--red-dark) !important;
        transform: scale(0.98);
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
        Pilih menu di sebelah kiri untuk memulai.
    </div>
""", unsafe_allow_html=True)

# ── CTA Button
st.button("Mulai Konversi →")

# ── Stat Cards
st.markdown("""
    <div class="stat-row">
        <div class="stat-card accent">
            <div class="stat-label">File Diproses</div>
            <div class="stat-value">1,284</div>
            <div class="stat-change">Total keseluruhan</div>
        </div>
        <div class="stat-card">
            <div class="stat-label">Konversi Hari Ini</div>
            <div class="stat-value">38</div>
            <div class="stat-change">▲ 12% dari kemarin</div>
        </div>
        <div class="stat-card">
            <div class="stat-label">Terakhir Diproses</div>
            <div class="stat-value">14:32</div>
            <div class="stat-change">WIB — Hari ini</div>
        </div>
    </div>
""", unsafe_allow_html=True)

# ── Info Banner
st.markdown("""
    <div class="info-banner">
        <strong>Cara menggunakan:</strong> Pilih menu
        <em>Konversi Data</em> atau <em>Upload File</em>
        di sidebar kiri untuk mulai mengolah data Anda.
    </div>
""", unsafe_allow_html=True)
