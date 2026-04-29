import streamlit as st

# ================================
# 🎨 CSS Styling
# ================================
st.markdown("""
    <style>
    @import url('[fonts.googleapis.com](https://fonts.googleapis.com/css2?family=Roboto+Condensed:wght@400;700&display=swap)');
    
    :root {
        --primary-color: #004aad;
        --secondary-color: #f5f7fb;
        --text-color: #222;
        --muted-text: #666;
        --card-bg: #ffffff;
        --border-color: #e5e7eb;
    }

    html, body, [class*="css"] {
        background-color: var(--secondary-color);
        font-family: 'Roboto Condensed', sans-serif;
        color: var(--text-color);
    }

    .main-title {
        font-size: 50px;
        font-weight: 700;
        color: var(--primary-color);
        line-height: 1.2;
        margin-bottom: 5px;
        letter-spacing: -0.5px;
    }

    .sub-title {
        font-size: 26px;
        font-weight: 400;
        color: var(--muted-text);
        margin-bottom: 30px;
    }

    /* Container Card */
    .content-card {
        background-color: var(--card-bg);
        padding: 25px 35px;
        border-radius: 12px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.05);
        border: 1px solid var(--border-color);
        margin-top: 15px;
    }

    /* Paragraph & Text */
    .stMarkdown p, .content-text {
        font-size: 18px;
        line-height: 1.6;
        color: var(--text-color);
        font-family: 'Roboto Condensed', sans-serif;
    }

    /* Divider */
    hr {
        border: none;
        border-top: 1px solid var(--border-color);
        margin: 25px 0;
    }

    /* Tombol Streamlit */
    div.stButton > button {
        background-color: var(--primary-color);
        color: white;
        border: none;
        padding: 0.6rem 1.5rem;
        border-radius: 6px;
        font-size: 17px;
        font-weight: 500;
        transition: all 0.3s ease;
        font-family: 'Roboto Condensed', sans-serif;
    }
    div.stButton > button:hover {
        background-color: #003b8a;
        transform: translateY(-2px);
        box-shadow: 0 3px 6px rgba(0,0,0,0.1);
    }
    </style>
""", unsafe_allow_html=True)

# ================================
# 🖥️ Render Layout
# ================================
st.markdown('<div class="main-title">Selamat Datang di PEI I-Fast Converter</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">MNC Sekuritas</div>', unsafe_allow_html=True)

st.markdown('<div class="content-card">', unsafe_allow_html=True)
st.markdown(
    '<p class="content-text">Platform ini membantu Anda mengolah dan mengonversi data dengan cepat dan akurat. '
    'Gunakan menu di sebelah kiri untuk mengakses berbagai fitur konversi dan alat analisis.</p>',
    unsafe_allow_html=True
)
st.markdown('</div>', unsafe_allow_html=True)

st.markdown("<hr>", unsafe_allow_html=True)
st.write("📂 Silakan pilih menu di sidebar untuk memulai proses konversi data.")
