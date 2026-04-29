import streamlit as st

# CSS untuk Font & Merapikan Tampilan
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto+Condensed:wght@400;700&display=swap');

    /* Terapkan font ke seluruh aplikasi */
    html, body, [class*="st-"], .stMarkdown {
        font-family: 'Roboto Condensed', sans-serif !important;
    }

    /* Ukuran Judul Utama */
    .main-title {
        font-family: 'Roboto Condensed', sans-serif;
        font-weight: 700;
        font-size: 50px; /* Ukuran bisa Anda tambah di sini */
        margin-bottom: -10px;
        line-height: 1.2;
    }

    /* Ukuran Sub-Judul */
    .sub-title {
        font-family: 'Roboto Condensed', sans-serif;
        font-weight: 400;
        font-size: 32px;
        color: #666;
        margin-bottom: 20px;
    }
    
    /* Sembunyikan artifact teks aneh jika muncul di area header */
    .css-1544893 { display: none; }
    </style>
    """,
    unsafe_allow_html=True
)

# Render Judul
st.markdown('<div class="main-title">Selamat Datang di PEI I-Fast Converter</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">MNC Sekuritas</div>', unsafe_allow_html=True)

st.write("Silakan pilih menu di sebelah kiri untuk mulai mengolah data.")
