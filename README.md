<div align="center">

# 🧾 ReimburseLLM

**Scan & ekstraksi struk jadi Excel reimburse — otomatis, kronologis, tanpa kompresi.**

Upload sebulan foto struk (bensin, parkir, Teazzi) dan screenshot Flazz/e-money.
Model vision membaca tiap gambar, mengisi template Excel, dan menggabungkan semua
bukti asli ke satu PDF.

[![Streamlit](https://img.shields.io/badge/Streamlit-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)](https://streamlit.io)
[![Python](https://img.shields.io/badge/Python-3.11%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org)
[![License](https://img.shields.io/badge/License-MIT-22c55e?style=for-the-badge)](#-lisensi)

</div>

---

## ✨ Fitur

| | Fitur | Keterangan |
|---|---|---|
| 🔍 | **OCR multi-upload** | Unggah banyak foto struk sekaligus untuk satu bulan penuh. |
| 🧠 | **Multi-provider vision** | Kagiro, Bandel & 9router (gateway custom) — pilih provider + model manual. |
| 🚗 | **Aturan per kategori** | Bensin, Parkir, dan Drink (Teazzi) dipetakan sesuai aturan baku. |
| 💳 | **Dedup Flazz** | Baris parkir di screenshot Flazz yang sama dengan struk fisik tidak dobel. |
| ⏱️ | **Urut kronologis** | Diurutkan berdasarkan tanggal + jam (timestamp) struk/transaksi. |
| 📅 | **Periode otomatis** | Sel **E6** = transaksi terawal, **E7** = transaksi terakhir. |
| 📍 | **Fallback GPS** | Nama parkir tanpa teks di struk diisi dari EXIF GPS + reverse geocode. |
| 📄 | **PDF gabungan** | Semua bukti asli digabung jadi satu PDF **tanpa kompresi / downsizing**. |
| 🩺 | **Cek API hidup** | Ping sungguhan ke model terpilih — hanya saat tombol ditekan. |

---

## 🚀 Mulai Cepat

### 1. Clone & install

```bash
git clone https://github.com/<username>/<repo>.git
cd <repo>
pip install -r requirements.txt
```

### 2. Siapkan secrets

```bash
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
```

Isi `.streamlit/secrets.toml` dengan API key asli (satu blok per provider; provider tanpa key otomatis hilang dari dropdown):

```toml
KAGIRO_API_KEY = "mk-...."
BANDEL_API_KEY = "sk-qwen-...."

# 9router (gateway custom) — nama provider di dropdown: "9router"
ROUTER9_API_KEY = "sk-...."
ROUTER9_BASE_URL = "https://rrakv37.abc-tunnel.us/v1"
```

> 🔒 File `secrets.toml` **tidak pernah** ikut ter-commit — sudah dikecualikan oleh `.gitignore`.

### 3. Jalankan

```bash
streamlit run app.py
```

---

## ☁️ Deploy ke Streamlit Community Cloud

1. Push repo ini ke GitHub (tanpa `.streamlit/secrets.toml`).
2. Buat app baru di [share.streamlit.io](https://share.streamlit.io) → arahkan ke `app.py`.
3. Buka **App → Settings → Secrets**, tempel isi yang sama seperti `secrets.toml`.
4. Deploy. 🎉

---

## 📖 Cara Pakai

1. **Isi Data Pemohon** di sidebar (Name, Department, Purpose, Bank Acc.).
2. **Pilih Provider OCR** — Kagiro, Bandel, atau 9router, lalu pilih **Model Vision** (dropdown hanya menampilkan model yang bisa membaca gambar).
   Daftar model diambil sekali per provider (tanpa ping) dan di-cache.
3. *(Opsional)* **🩺 Cek API hidup** untuk memastikan model terpilih merespons.
4. **Unggah Bukti** — foto struk fisik, dan screenshot Flazz/e-money (opsional).
5. Klik **Ekstraksi** → unduh **Excel** dan **PDF gabungan**.

---

## 📐 Aturan Ekstraksi

- **Bensin** — `Pertalite` → description cukup nama BBM. `Pertamax`/jenis lain →
  description berisi jenis BBM + jumlah liter; nominal tetap total akhir struk.
- **Drink (Teazzi)** — description berisi jenis minuman & nama outlet; nominal total akhir.
- **Parkir** — description berisi nama tempat, atau `-` jika tidak tertera (isi manual nanti).
- **Dedup Flazz** — baris `Parking` di screenshot Flazz yang tanggal & nominalnya sama
  persis dengan struk parkir fisik **di-skip**. Baris `Top Up` **selalu** di-skip.
- **Sorting** — hasil di Excel diurutkan dari tanggal transaksi paling awal ke akhir.
- **Periode** — sel **E6** (awal) & **E7** (akhir) diisi otomatis dari timestamp transaksi.
- **PDF** — seluruh foto struk dan screenshot Flazz digabung **tanpa kompresi**.

---

## 🗂️ Struktur Proyek

```
├── 🏠 app.py                      → UI Streamlit + orkestrasi proses
├── 🧠 llm_core.py                 → Provider, ping model sungguhan, call vision (retry/timeout)
├── 📝 extract_core.py             → Prompt + aturan ekstraksi + dedup Flazz + skip Top Up
├── 📊 excel_core.py               → Isi template Excel (header, periode E6/E7, total)
├── 📄 pdf_core.py                 → Gabung gambar asli → PDF tanpa kompresi
├── 📍 gps_core.py                 → Baca lokasi GPS dari EXIF foto struk
├── 📗 FORM_REIMBURSE_template.xlsx → Template Excel yang diisi otomatis
├── 📦 requirements.txt            → Dependency Python
│
├── 📁 .streamlit/                 → secrets.toml.example
├── 📁 .devcontainer/              → devcontainer.json (Python 3.11 + auto-run Streamlit)
└── 📁 .github/workflows/          → keep-alive.yml (ping opsional)
```

---

## 🛠️ Teknologi

- **[Streamlit](https://streamlit.io)** — UI web interaktif
- **[Pillow](https://python-pillow.org)** — normalisasi & resize gambar
- **[img2pdf](https://pypi.org/project/img2pdf/)** — gabung gambar jadi PDF tanpa kompresi
- **[openpyxl](https://openpyxl.readthedocs.io)** — tulis template Excel
- **[pandas](https://pandas.pydata.org)** — agregasi & sorting data
- **[requests](https://requests.readthedocs.io)** — panggilan provider vision

---

## 🔐 Keamanan

> ⚠️ **Jangan pernah** commit `secrets.toml` atau API key ke repo.
>
> Kalau key pernah terekspos di kode lama/history, segera **revoke & rotate**
> di dashboard provider dan buat key baru.
>
> Isi API key lewat **App → Settings → Secrets** di Streamlit Cloud, bukan lewat file di repo.

---

## 📄 Lisensi

Dirilis di bawah lisensi **MIT** — lihat [LICENSE](LICENSE) untuk detail.

<div align="center">

**Dibuat dengan ❤️ untuk mempercepat urusan reimburse.**

</div>
