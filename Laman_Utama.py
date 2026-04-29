import streamlit as st

# Menyisipkan font Roboto Condensed dari Google Fonts
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto+Condensed:ital,wght@0,100..900;1,100..900&display=swap');

    /* Mengatur agar font diterapkan pada seluruh aplikasi atau elemen tertentu */
    html, body, [class*="st-"] {
        font-family: 'Roboto Condensed', sans-serif;
    }
    
    .title-text {
        font-family: 'Roboto Condensed', sans-serif;
        font-weight: 700;
        font-size: 42px;
        margin-bottom: 0px;
    }
    
    .subtitle-text {
        font-family: 'Roboto Condensed', sans-serif;
        font-weight: 400;
        font-size: 24px;
        color: #555;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Tampilan Judul
st.markdown('<p class="title-text">Selamat Datang di PEI I-Fast Converter</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitle-text">MNC Sekuritas</p>', unsafe_allow_html=True)

st.write("Silakan pilih menu di sebelah kiri untuk mulai mengolah data.")
