import streamlit as st

st.set_page_config(
    page_title="Oracle Fusion Excel Runner",
    page_icon="🧩",
    layout="wide"
)

st.title("🧩 Oracle Fusion Excel Runner")
st.caption("MVP untuk mencari reference ID, GET existing Inventory Organizations, membuat template Excel, generate JSON payload, dan upload data ke Oracle Fusion lewat REST API.")

st.markdown(
    """
    ### Alur kerja

    1. Buka **Reference Finder** untuk mengambil ID Business Unit, Legal Entity, Inventory Organization, dan Location dari instance.
    2. GET existing Inventory Organization kalau mau membuat template dari contoh real.
    3. Atau buka **Template & JSON Builder** untuk membuat template manual dari schema lokal.
    4. Isi template Excel sesuai data yang mau dinaikkan.
    5. Buka **Upload Runner** untuk validasi Excel, preview payload, lalu test upload ke Oracle Fusion.

    ### Fokus MVP

    Endpoint awal yang disiapkan:

    ```http
    POST /fscmRestApi/resources/11.13.18.05/inventoryOrganizations
    ```

    ### Catatan keamanan

    - Gunakan **Mock/Dry Run mode** dulu sebelum upload live.
    - Tes di instance **DEV/TEST** sebelum masuk PROD.
    - Jangan simpan username dan password asli di repository.
    """
)

st.info("Mulai dari menu di sidebar: Reference Finder, Template & JSON Builder, atau Upload Runner.")
