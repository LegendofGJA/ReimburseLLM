"""
pdf_core.py — Gabungkan gambar asli menjadi PDF TANPA kompresi/downsizing.

Pakai img2pdf yang menyisipkan byte gambar apa adanya (lossless embedding).
Gambar asli tidak di-resize/di-re-encode di sini.
"""

import io

import img2pdf

def merge_images_to_pdf(image_bytes_list):
    """Return bytes PDF dari list bytes gambar (JPEG/PNG), tanpa kompresi."""
    if not image_bytes_list:
        return b""
    # Memasang rotation=img2pdf.Rotation.ifvalid agar EXIF orientation tidak valid/0 diabaikan
    return img2pdf.convert(image_bytes_list, rotation=img2pdf.Rotation.ifvalid)
