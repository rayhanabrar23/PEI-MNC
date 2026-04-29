import streamlit as st

# CSS yang lebih aman (tidak merusak ikon sidebar)
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto+Condensed:wght@400;700&display=swap');

    /* Font hanya untuk class kustom kita, bukan elemen sistem Streamlit */
    .custom-font {
        font-family: 'Roboto Condensed', sans-serif !important;
    }

    .main-title {
        font-size: 52px; /* Ukuran lebih besar */
        font-weight: 700;
        line-height: 1.1;
        margin-bottom: 5px;
    }

    .sub-title {
        font-size: 32px; /* Ukuran lebih besar */
        font-weight: 400;
        color: #555;
        margin-bottom: 30px;
    }
    
    /* Memperbaiki font untuk st.write agar seragam tapi aman */
    .stMarkdown p {
        font-family: 'Roboto Condensed', sans-serif;
        font-size: 18px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Render Judul menggunakan div kustom
st.markdown('<div class="custom-font main-title">Selamat Datang di PEI I-Fast Converter</div>', unsafe_allow_html=True)
st.markdown('<div class="custom-font sub-title">MNC Sekuritas</div>', unsafe_allow_html=True)

st.write("Silakan pilih menu di sebelah kiri untuk mulai mengolah data.")
