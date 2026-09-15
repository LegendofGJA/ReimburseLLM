"""
extract_core.py — Aturan ekstraksi data struk & Flazz (bebas dari UI).

Berisi prompt + logika murni (testable tanpa Streamlit):
  - klasifikasi kategori (bensin / drink / parkir),
  - penyusunan description per kategori,
  - ekstraksi screenshot Flazz (JSON array),
  - dedup Flazz vs struk parkir berdasar (tanggal, nominal),
  - skip "Top Up".
"""

import re

from dateutil import parser as dateparser

# ─────────────────────────────────────────────────────────────────────────
# Prompt
# ─────────────────────────────────────────────────────────────────────────

EXTRACTION_PROMPT = """Kamu adalah sistem OCR untuk struk belanja Indonesia.
Analisis gambar struk ini dan kembalikan HANYA JSON murni (tanpa markdown, tanpa teks tambahan) dengan struktur persis berikut:

{
  "date": "YYYY-MM-DD",
  "type": "bensin" | "parkir" | "drink",
  "nominal": <angka total akhir struk, tanpa titik/koma/Rp>,
  "fuel_type": "<nama jenis BBM apa adanya di struk, contoh: Pertalite / Pertamax / Pertamax Turbo, atau null jika bukan bensin>",
  "liters": <jumlah liter sebagai angka, atau null jika tidak ada / bukan bensin>,
  "drink_name": "<jenis/menu minuman yang dibeli, atau null jika bukan drink>",
  "outlet_name": "<nama toko/outlet, contoh Teazzi, atau null jika bukan drink>",
  "location_name": "<nama tempat/lokasi parkir, atau null jika bukan parkir. Jika parkir tapi nama tempat tidak tertera, isi dengan '-'>"
}

Aturan klasifikasi "type":
- Struk pom bensin / SPBU -> "bensin"
- Struk parkir -> "parkir"
- Struk Teazzi atau minuman lain -> "drink"

Ambil "nominal" sebagai TOTAL AKHIR yang benar-benar dibayar pada struk.
Tanggal WAJIB diambil dari tanggal transaksi yang tertera di struk, bukan diasumsikan.
Kembalikan JSON murni saja, tidak ada teks lain sebelum atau sesudahnya."""

FLAZZ_PROMPT = """Kamu adalah sistem OCR untuk screenshot riwayat transaksi e-money (Flazz/BCA).
Gambar ini berisi DAFTAR transaksi (mungkin banyak baris). Ekstrak SEMUA baris transaksi.

Kembalikan HANYA JSON array murni (tanpa markdown), setiap elemen:

{
  "date": "YYYY-MM-DD",
  "label": "<teks baris transaksi apa adanya, misal: Parking, Top Up, Merchant name>",
  "nominal": <angka nominal transaksi, tanpa titik/koma/Rp>,
  "is_topup": true|false
}

Aturan:
- "is_topup" true hanya jika baris itu top up / isi saldo / topup; sisanya false.
- Baris "Parking" biasanya adalah pembayaran parkir.
- Ambil tanggal sesuai baris transaksi (format mungkin "06 Jul 2026").
- Kembalikan JSON array murni saja, tidak ada teks lain sebelum/sesudahnya."""


# ─────────────────────────────────────────────────────────────────────────
# Description per kategori (logika dipastikan di Python, bukan di model)
# ─────────────────────────────────────────────────────────────────────────


def build_description(item: dict) -> str:
    itype = (item.get("type") or "").strip().lower()

    if itype == "bensin":
        fuel = (item.get("fuel_type") or "").strip()
        liters = item.get("liters")
        if not fuel:
            return "-"
        if "pertalite" in fuel.lower():
            return fuel
        if liters:
            try:
                liters_fmt = f"{float(liters):g}"
            except (TypeError, ValueError):
                liters_fmt = str(liters)
            return f"{fuel} {liters_fmt} Liter"
        return fuel

    if itype == "drink":
        drink = (item.get("drink_name") or "").strip()
        outlet = (item.get("outlet_name") or "").strip()
        parts = [p for p in [drink, outlet] if p]
        return " - ".join(parts) if parts else "-"

    if itype == "parkir":
        loc = (item.get("location_name") or "").strip()
        return loc if loc else "-"

    return "-"


# ─────────────────────────────────────────────────────────────────────────
# Parsing
# ─────────────────────────────────────────────────────────────────────────


def parse_date_safe(raw):
    if not raw:
        return None
    if isinstance(raw, (int, float)):
        return None
    text = str(raw).strip()
    if not text:
        return None
    try:
        return dateparser.parse(text, dayfirst=False, yearfirst=True).date()
    except Exception:
        try:
            return dateparser.parse(text, dayfirst=True).date()
        except Exception:
            return None


def parse_nominal(value):
    """Nominal -> float. Abaikan titik/koma/Rp, pertahankan tanda minus."""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = re.sub(r"[^\d.\-]", "", str(value))
    if cleaned in ("", "-", ".", "-."):
        return 0.0
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def _to_date(d):
    return d if d else __import__("datetime").date.max


# ─────────────────────────────────────────────────────────────────────────
# Struk fisik -> rows
# ─────────────────────────────────────────────────────────────────────────


def build_rows(extracted_items: list) -> list:
    """Ubah daftar hasil struk menjadi list dict (date, category, description,
    nominal) dan urutkan kronologis (struk tanpa tanggal di akhir)."""
    import pandas as pd

    rows = []
    for item in extracted_items:
        rows.append(
            {
                "date": parse_date_safe(item.get("date")),
                "category": (item.get("type") or "-").strip().lower(),
                "description": build_description(item),
                "nominal": parse_nominal(item.get("nominal")),
                "location_name": (item.get("location_name") or "").strip(),
            }
        )
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["_sort_key"] = df["date"].apply(_to_date)
    df = df.sort_values("_sort_key").drop(columns=["_sort_key"]).reset_index(drop=True)
    return df


# ─────────────────────────────────────────────────────────────────────────
# Flazz: normalisasi + dedup + skip topup
# ─────────────────────────────────────────────────────────────────────────


def _is_topup(row: dict) -> bool:
    if row.get("is_topup") is True:
        return True
    label = (row.get("label") or "").strip().lower()
    return "top up" in label or "topup" in label or "top-up" in label or "isi saldo" in label


def normalise_flazz_rows(flazz_items: list) -> list:
    """Buang baris Top Up, lalu kembalikan baris parkir (label mengandung
    'park') sebagai list dict {date, nominal} + raw untuk description."""
    rows = []
    for it in flazz_items:
        if _is_topup(it):
            continue
        label = (it.get("label") or "").strip()
        is_parking = "park" in label.lower()
        rows.append(
            {
                "date": parse_date_safe(it.get("date")),
                "nominal": parse_nominal(it.get("nominal")),
                "label": label,
                "is_parking": is_parking,
            }
        )
    return rows


def dedupe_flazz_against_receipts(flazz_rows: list, receipt_rows: list) -> tuple:
    """Cocokkan baris Flazz 'parkir' ke struk parkir fisik berdasar
    (tanggal, nominal). Baris yang match persis di-SKIP (sudah terwakili struk).

    receipt_rows: list dict hasil build_rows (memiliki date, category, nominal).

    Return (kept: list[flazz_rows], skipped: int, matched_pairs: list[tuple]).
    """
    receipt_pairs = set()
    for r in receipt_rows:
        if (r.get("category") or "") == "parkir" and r.get("date") and r.get("nominal"):
            receipt_pairs.add((r["date"], round(float(r["nominal"]), 2)))

    kept = []
    skipped = 0
    matched = []
    for fr in flazz_rows:
        if not fr["is_parking"]:
            kept.append(fr)
            continue
        key = (fr["date"], round(float(fr["nominal"]), 2))
        if fr["date"] and key in receipt_pairs:
            skipped += 1
            matched.append((fr["date"], fr["nominal"]))
            continue
        kept.append(fr)
    return kept, skipped, matched
