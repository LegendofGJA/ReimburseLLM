# 📦 folder-bersih — Isi Repo GitHub (siap upload)

Folder ini berisi **seluruh file yang dipakai aplikasi** dan **aman untuk repo publik**.
Upload isi folder ini (bukan folder induknya) ke root repo GitHub, atau pakai
`LLMREIM15092026.zip` di folder induk yang isinya sudah identik.

## Isi

```
folder-bersih/
├── app.py                        → UI Streamlit + orkestrasi proses
├── llm_core.py                   → Provider, ping model sungguhan, call vision
├── extract_core.py               → Prompt + aturan ekstraksi + dedup Flazz
├── excel_core.py                 → Isi template Excel (header, periode E6/E7, total)
├── pdf_core.py                   → Gabung gambar asli → PDF tanpa kompresi
├── gps_core.py                   → Baca lokasi GPS dari EXIF foto
├── requirements.txt              → Dependency Python
├── README.md                     → Dokumentasi repo
├── .gitignore                    → Kecualikan secrets.toml & cache
├── FORM_REIMBURSE_template.xlsx  → Template Excel kosong (tanpa data pribadi)
├── .streamlit/
│   └── secrets.toml.example      → Contoh secrets (placeholder, BUKAN key asli)
├── .devcontainer/
│   └── devcontainer.json         → Dev container (Python 3.11 + auto-run)
└── .github/workflows/
    └── keep-alive.yml            → Ping opsional (set repo variable APP_URL)
```

## Sudah diverifikasi bersih

- ✅ Tidak ada API key asli (`secrets.toml` tidak disertakan).
- ✅ Tidak ada data pribadi (nama / no. rekening) di kode maupun template Excel.
- ✅ Tidak ada foto contoh / sampel struk.
- ✅ Semua modul Python lolos `py_compile`.

## Cara upload

1. Ekstrak `LLMREIM15092026.zip`, **atau** langsung salin isi folder ini.
2. Masukkan ke root repo GitHub (`git add . && git commit && git push`).
3. Di Streamlit Cloud: **App → Settings → Secrets** → tempel key asli.

> ⚠️ Jangan lupa `git add -f` tidak diperlukan; biarkan `.gitignore` bekerja.
