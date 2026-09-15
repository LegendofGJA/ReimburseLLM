"""
pdf_core.py — Gabungkan gambar asli menjadi PDF TANPA kompresi/downsizing.

Pakai img2pdf yang menyisipkan byte gambar apa adanya (lossless embedding).
Gambar asli tidak di-resize/di-re-encode di sini.

Perhatian pada metadata EXIF: sebagian foto HP menyimpan nilai Orientation
yang tidak standar (mis. 0), sehingga img2pdf melempar `ExifOrientationError`.
Kita pasang `rotation=img2pdf.Rotation.ifvalid` agar nilai orientasi yang tidak
valid diabaikan dan gambar tetap digabung tanpa memutar aslinya.
"""

import io

import img2pdf


def _to_io(image_bytes):
    return io.BytesIO(image_bytes)


def merge_images_to_pdf(image_bytes_list: list) -> bytes:
    """Return bytes PDF dari list bytes gambar (JPEG/PNG), tanpa kompresi."""
    if not image_bytes_list:
        return b""
    streams = [_to_io(b) for b in image_bytes_list]
    kwargs = {"rotation": img2pdf.Rotation.ifvalid}
    try:
        return img2pdf.convert(streams, **kwargs)
    except img2pdf.ExifOrientationError:
        # Fallback terakhir: paksa orientasi default tanpa membaca EXIF rotation.
        return img2pdf.convert(streams)
