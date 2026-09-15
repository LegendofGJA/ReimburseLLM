"""
app.py — Aplikasi Scan & Ekstraksi Reimburse Struk (Streamlit).

Modul terpisah:
  - llm_core.py    : koneksi provider + ping sungguhan + call vision (retry/timeout)
  - extract_core.py: prompt + aturan ekstraksi + dedup Flazz + skip top up
  - excel_core.py  : isi template Excel FORM_REIMBURSE
  - pdf_core.py    : gabung gambar asli -> PDF tanpa kompresi
  - gps_core.py    : baca lokasi GPS dari EXIF foto (dipertahankan)

Jalankan:
    pip install -r requirements.txt
    streamlit run app.py
"""

import os
from datetime import date

import streamlit as st

from excel_core import TEMPLATE_PATH, fill_excel_template
from extract_core import (
    EXTRACTION_PROMPT,
    FLAZZ_PROMPT,
    build_rows,
    dedupe_flazz_against_receipts,
    normalise_flazz_rows,
)
from gps_core import gps_location_name, read_gps
from llm_core import PROVIDERS, check_all_providers, fetch_models, ping_model
from llm_core import call_vision
from pdf_core import merge_images_to_pdf

st.set_page_config(page_title="Scan Reimburse Struk", page_icon="🧾", layout="wide")

st.markdown(
    """
    <h2>🧾 Aplikasi Scan & Ekstraksi Reimburse Struk</h2>
    <p>Upload foto struk (bensin, parkir, Teazzi) dan screenshot Flazz/e-money.
    Sistem mengekstrak data, mengurutkan kronologis, mengisi template Excel, dan
    menggabungkan bukti asli ke satu PDF tanpa kompresi.</p>
    """,
    unsafe_allow_html=True,
)

if not PROVIDERS:
    st.error(
        "Belum ada API key yang dikonfigurasi. Buat file `.streamlit/secrets.toml` "
        "(lihat `.streamlit/secrets.toml.example`) atau set environment variable "
        "KAGIRO_API_KEY / BANDEL_API_KEY, lalu jalankan ulang aplikasi."
    )
    st.stop()


# ─────────────────────────────────────────────────────────────────────────
# Sidebar: data pemohon
# ─────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.subheader("Data Pemohon")

    def _default(key: str, fallback: str = "") -> str:
        try:
            if key in st.secrets:
                return str(st.secrets[key])
        except Exception:
            pass
        return os.environ.get(key, fallback)

    name = st.text_input("Name", value=_default("DEFAULT_NAME"))
    department = st.text_input("Department", value=_default("DEFAULT_DEPARTMENT"))
    purpose = st.text_input("Purpose", value=_default("DEFAULT_PURPOSE", "Reimburse"))
    bank_acc = st.text_input("Bank Acc.", value=_default("DEFAULT_BANK_ACC"))
    st.caption("Periode (kolom E6 & E7) diisi otomatis: awal & akhir bulan dari tanggal struk ter-upload.")


# ─────────────────────────────────────────────────────────────────────────
# Session state
# ─────────────────────────────────────────────────────────────────────────
DEFAULTS = {
    "extracted_items": [],
    "image_bytes_list": [],
    "flazz_rows": [],
    "flazz_kept": [],
    "flazz_skipped": 0,
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ─────────────────────────────────────────────────────────────────────────
# Provider & model
# ─────────────────────────────────────────────────────────────────────────
st.subheader("Pemilihan Provider OCR")

col_mode, col_prov = st.columns([1, 1])
with col_mode:
    mode = st.radio(
        "Mode pemilihan provider",
        ["Auto (cek semua, pakai yang aktif)", "Pilih manual"],
        horizontal=True,
    )

available_models = []
provider_name = None
status_text = ""

if mode.startswith("Auto"):
    # Auto: ping model sungguhan per provider, pilih provider pertama yang punya
    # minimal satu model merespons beneran.
    scan = check_all_providers()
    chosen = None
    chosen_models = []
    detail_lines = []
    for pname, info in scan.items():
        models = info["models"]
        results = info["results"]
        ok_models = [m for m in models if results.get(m, (False,))[0]]
        line = f"- **{pname}**: "
        if ok_models:
            line += "✅ " + ", ".join(ok_models)
        else:
            line += "❌ tidak ada model merespons"
        detail_lines.append(line)
        if chosen is None and ok_models:
            chosen = pname
            chosen_models = ok_models

    if chosen:
        provider_name = chosen
        available_models = chosen_models
        status_text = f"Provider aktif: **{chosen}** — ✅ model merespons beneran"
    else:
        provider_name = list(PROVIDERS.keys())[0]
        available_models = fetch_models(provider_name)
        status_text = "⚠️ Tidak ada provider yang merespons beneran. Daftar model default ditampilkan (belum terverifikasi)."

    st.markdown(status_text)
    with st.expander("Status cek semua provider (ping sungguhan)"):
        for line in detail_lines:
            st.markdown(line)
else:
    provider_name = st.selectbox("Pilih Provider OCR", list(PROVIDERS.keys()))
    available_models = fetch_models(provider_name)
    st.markdown(f"Provider: **{provider_name}**")

selected_model = st.selectbox(
    "Pilih Model Vision",
    available_models if available_models else ["Model tidak tersedia"],
)

if provider_name and selected_model != "Model tidak tersedia":
    with st.expander("🩺 Cek API hidup (ping sungguhan ke model terpilih)"):
        if st.button("🔄 Test ping model sekarang"):
            with st.spinner(f"Mengirim ping ke {provider_name} / {selected_model}..."):
                ok, pesan, _ms = ping_model(provider_name, selected_model)
            if ok:
                st.success(pesan)
            else:
                st.error(pesan)


# ─────────────────────────────────────────────────────────────────────────
# Uploader
# ─────────────────────────────────────────────────────────────────────────
st.subheader("Unggah Bukti")
col_receipt, col_flazz = st.columns(2)

with col_receipt:
    uploaded_files = st.file_uploader(
        "Unggah foto struk (Bensin, Parkir, Teazzi) untuk satu bulan",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key="receipt_uploader",
    )

with col_flazz:
    flazz_files = st.file_uploader(
        "Unggah screenshot Flazz / e-money (opsional)",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key="flazz_uploader",
    )

st.info(
    """
**Aturan ekstraksi & format output:**
1. **Sorting kronologis:** hasil Excel diurutkan dari tanggal transaksi struk paling awal ke akhir (bukan urutan upload).
2. **Bensin:** Pertalite → description cukup nama BBM. Pertamax/jenis lain → description berisi jenis BBM + jumlah liter, nominal tetap total akhir struk.
3. **Drink (Teazzi):** description berisi jenis minuman & nama outlet, nominal total akhir.
4. **Parkir:** description berisi nama tempat (atau `-` jika tidak ada), nominal total.
5. **Dedup Flazz:** baris "Parking" di screenshot Flazz yang tanggal & nominalnya sama persis dengan struk parkir fisik di-skip (tidak dobel). "Top Up" selalu di-skip.
6. **Periode E6/E7:** otomatis diisi awal & akhir bulan berdasarkan tanggal struk.
7. **PDF gabungan:** seluruh foto struk **dan** screenshot Flazz digabung ke satu PDF **tanpa kompresi/downsizing**.
    """
)


# ─────────────────────────────────────────────────────────────────────────
# Proses
# ─────────────────────────────────────────────────────────────────────────
if st.button("🚀 Mulai Proses OCR, Sorting, Generate Excel & PDF", type="primary"):
    if not uploaded_files and not flazz_files:
        st.warning("Silakan unggah minimal satu foto struk atau screenshot Flazz.")
    elif not available_models or selected_model == "Model tidak tersedia":
        st.error("Provider tidak aktif atau model tidak ditemukan. Gagal memproses.")
    else:
        extracted_items = []
        image_bytes_list = []
        failures = []

        total_files = len(uploaded_files or [])
        if total_files:
            progress_bar = st.progress(0)
            with st.spinner("Memproses foto struk dengan AI Vision..."):
                for i, file in enumerate(uploaded_files):
                    img_bytes = file.getvalue()
                    image_bytes_list.append(img_bytes)
                    try:
                        parsed = call_vision(provider_name, selected_model, img_bytes, EXTRACTION_PROMPT)
                        if isinstance(parsed, dict) and (parsed.get("type") or "").strip().lower() == "parkir":
                            loc = (parsed.get("location_name") or "").strip()
                            if not loc or loc == "-":
                                gps_name = gps_location_name(img_bytes)
                                if gps_name:
                                    parsed["location_name"] = gps_name
                        extracted_items.append(parsed)
                    except Exception as e:
                        failures.append((file.name, str(e)))
                    progress_bar.progress((i + 1) / total_files)

        # Screenshot Flazz
        flazz_items = []
        if flazz_files:
            with st.spinner("Memproses screenshot Flazz/e-money..."):
                for file in flazz_files:
                    img_bytes = file.getvalue()
                    image_bytes_list.append(img_bytes)  # ikut masuk PDF gabungan
                    try:
                        parsed = call_vision(provider_name, selected_model, img_bytes, FLAZZ_PROMPT)
                        if isinstance(parsed, list):
                            flazz_items.extend(parsed)
                        elif isinstance(parsed, dict):
                            flazz_items.append(parsed)
                    except Exception as e:
                        failures.append((file.name, str(e)))

        if failures:
            with st.expander(f"⚠️ {len(failures)} file gagal diproses"):
                for fname, err in failures:
                    st.write(f"- {fname}: {err}")

        # Dedup Flazz vs struk parkir + skip top up
        receipt_df = build_rows(extracted_items)
        flazz_norm = normalise_flazz_rows(flazz_items)
        flazz_kept, flazz_skipped, matched = dedupe_flazz_against_receipts(
            flazz_norm, receipt_df.to_dict("records")
        )

        # Baris Flazz parkir yang tetap -> jadi baris struk "parkir" (description '-')
        flazz_as_receipts = []
        for fr in flazz_kept:
            flazz_as_receipts.append(
                {
                    "date": fr["date"],
                    "type": "parkir",
                    "nominal": fr["nominal"],
                    "location_name": "-",
                    "description_override": "-",
                }
            )

        final_items = extracted_items + flazz_as_receipts

        st.session_state.extracted_items = final_items
        st.session_state.image_bytes_list = image_bytes_list
        st.session_state.flazz_kept = flazz_kept
        st.session_state.flazz_skipped = flazz_skipped

        msg = f"Ekstraksi selesai: {len(extracted_items)} struk fisik, {len(flazz_kept)} transaksi Flazz ditambahkan"
        if flazz_skipped:
            msg += f", {flazz_skipped} duplikat di-skip"
        st.success(msg + ".")


# ─────────────────────────────────────────────────────────────────────────
# Hasil & download
# ─────────────────────────────────────────────────────────────────────────
if st.session_state.extracted_items:
    df = build_rows(st.session_state.extracted_items)
    display_df = df[["date", "category", "description", "nominal"]].copy()
    display_df["date"] = display_df["date"].apply(lambda d: d.strftime("%d %b %Y") if d else "-")
    st.dataframe(display_df, use_container_width=True)
    st.markdown(f"**Total: Rp {df['nominal'].sum():,.0f}**".replace(",", "."))

    # GPS dari foto struk (dipertahankan)
    gps_list = []
    for f in uploaded_files or []:
        g = read_gps(f.getvalue())
        if g:
            gps_list.append((f.name, g))
    if gps_list:
        with st.expander("📍 Lokasi GPS dari EXIF foto"):
            for fname, g in gps_list:
                st.write(
                    f"- **{fname}**: {g['lat']:.6f}, {g['lon']:.6f}"
                    + (f" (diambil {g['timestamp']})" if g.get("timestamp") else "")
                )

    excel_bytes = fill_excel_template(
        df,
        TEMPLATE_PATH,
        {"name": name, "department": department, "purpose": purpose, "bank_acc": bank_acc},
    )
    pdf_bytes = merge_images_to_pdf(st.session_state.image_bytes_list)

    dcol1, dcol2 = st.columns(2)
    with dcol1:
        st.download_button(
            "📥 Download Excel Reimburse",
            data=excel_bytes,
            file_name=f"FORM_REIMBURSE_{date.today().strftime('%Y%m')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with dcol2:
        st.download_button(
            "📥 Download PDF Bukti Gabungan (tanpa kompresi)",
            data=pdf_bytes,
            file_name=f"Bukti_Gabungan_{date.today().strftime('%Y%m')}.pdf",
            mime="application/pdf",
        )
