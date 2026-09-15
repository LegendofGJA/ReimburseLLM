"""
llm_core.py — Koneksi provider AI vision (Kagiro / Bandel).

Menyediakan:
  - `PROVIDERS`       : konfigurasi provider dari st.secrets (bukan hardcoded).
  - `fetch_models`    : ambil daftar model dari endpoint /models (fallback default).
  - `ping_model`      : tes API "hidup" dengan KIRIM completion kecil BENERAN,
                        bukan cuma HTTP 200 pada /models — sehingga model yang
                        terdaftar tapi tidak merespons akan terdeteksi.
  - `prepare_image`   : resize + re-encode JPEG agar payload aman (menghindari
                        HTTP 413 Payload Too Large dari server).
  - `call_vision`     : panggil chat/completions dengan timeout panjang + retry
                        + backoff (menghindari read timeout).
"""

import base64
import io
import json
import os
import re
import time

import requests
import streamlit as st
from PIL import Image

# ─────────────────────────────────────────────────────────────────────────
# Konfigurasi provider (secrets, tidak pernah hardcoded)
# ─────────────────────────────────────────────────────────────────────────


def _get_secret(key: str, default: str = "") -> str:
    try:
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return os.environ.get(key, default)


PROVIDERS = {
    "Kagiro": {
        "base_url": _get_secret("KAGIRO_BASE_URL", "https://api.kagiro.net/v1"),
        "api_key": _get_secret("KAGIRO_API_KEY"),
    },
    "Bandel": {
        "base_url": _get_secret("BANDEL_BASE_URL", "https://bandelbanget.xyz/v1"),
        "api_key": _get_secret("BANDEL_API_KEY"),
    },
}
# Provider tanpa API key di-skip (supaya tidak error saat salah satu belum di-set).
PROVIDERS = {k: v for k, v in PROVIDERS.items() if v["api_key"]}

# Fallback kalau endpoint /models tidak tersedia / kosong.
FALLBACK_MODELS = [
    "qwen-vl-max",
    "qwen2.5-vl-72b-instruct",
    "qwen-vl-max-latest",
    "gpt-4o",
    "gpt-4o-mini",
]


def _auth_headers(cfg: dict) -> dict:
    return {"Authorization": f"Bearer {cfg['api_key']}"}


# ─────────────────────────────────────────────────────────────────────────
# Daftar model
# ─────────────────────────────────────────────────────────────────────────


def fetch_models(provider_name: str) -> list:
    """Ambil daftar model id dari /models. Gagal -> fallback default."""
    cfg = PROVIDERS[provider_name]
    url = f"{cfg['base_url']}/models"
    try:
        r = requests.get(url, headers=_auth_headers(cfg), timeout=12)
        if r.status_code == 200:
            data = r.json()
            models = data.get("data", data)
            ids = [
                (m.get("id") if isinstance(m, dict) else m)
                for m in (models if isinstance(models, list) else [])
            ]
            ids = [i for i in ids if isinstance(i, str)]
            if ids:
                return ids
    except Exception:
        pass
    return list(FALLBACK_MODELS)


# ─────────────────────────────────────────────────────────────────────────
# Ping API "hidup" — kirim completion asli, bukan cuma HTTP 200
# ─────────────────────────────────────────────────────────────────────────


def ping_model(provider_name: str, model: str) -> tuple:
    """Kirim completion kecil beneran ke model dan laporkan hasilnya.

    Return (ok: bool, pesan: str, ms: int|None).
    """
    cfg = PROVIDERS[provider_name]
    url = f"{cfg['base_url']}/chat/completions"
    payload = {
        "model": model,
        "messages": [{"role": "user", "content": "ping"}],
        "max_tokens": 5,
        "temperature": 0,
    }
    headers = {**_auth_headers(cfg), "Content-Type": "application/json"}
    start = time.time()
    try:
        r = requests.post(url, headers=headers, json=payload, timeout=30)
        ms = int((time.time() - start) * 1000)
        if r.status_code == 200:
            data = r.json()
            content = ""
            try:
                content = data["choices"][0]["message"]["content"]
            except (KeyError, IndexError, TypeError):
                content = ""
            if content is None or str(content).strip() == "":
                return False, f"⚠️ HTTP 200 tapi balasan kosong ({ms} ms)", ms
            return True, f"✅ Model merespons beneran ({ms} ms)", ms
        else:
            return False, f"❌ Status {r.status_code}: {r.text[:120]}", ms
    except Exception as e:
        return False, f"❌ Gagal terhubung ({e})", None


def check_all_providers() -> dict:
    """Scan model per provider: setiap model di-ping sungguhan.

    Return {provider_name: {"models": [id...], "results": {model: (ok,pesan)}}}
    """
    out = {}
    for name in PROVIDERS:
        models = fetch_models(name)
        results = {}
        for m in models:
            ok, pesan, _ms = ping_model(name, m)
            results[m] = (ok, pesan)
        out[name] = {"models": models, "results": results}
    return out


# ─────────────────────────────────────────────────────────────────────────
# Siapkan gambar (hindari 413 + byte kecil)
# ─────────────────────────────────────────────────────────────────────────


def prepare_image(file_bytes: bytes, max_side: int = 1600, quality: int = 85) -> bytes:
    """Resize gambar ke max_side dan re-encode jadi JPEG sehingga payload
    base64-nya tetap kecil (teks struk tetap terbaca). Gambar asli TIDAK
    diubah (dipakai untuk PDF gabungan)."""
    try:
        img = Image.open(io.BytesIO(file_bytes))
        img = img.convert("RGB")
    except Exception:
        return file_bytes
    w, h = img.size
    longest = max(w, h)
    if longest > max_side:
        ratio = max_side / float(longest)
        img = img.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=quality, optimize=True)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────
# Panggil vision API
# ─────────────────────────────────────────────────────────────────────────


def call_vision(
    provider_name: str,
    model: str,
    image_bytes: bytes,
    prompt: str,
    timeout: int = 120,
    retries: int = 2,
) -> dict:
    """Kirim satu gambar + prompt ke model vision. Return objek JSON (dict/list).

    - image_bytes di-resize dulu (prepare_image) untuk hindari 413.
    - timeout panjang (default 120 dtk) + retry dengan backoff utk read timeout.
    """
    cfg = PROVIDERS[provider_name]
    url = f"{cfg['base_url']}/chat/completions"
    headers = {**_auth_headers(cfg), "Content-Type": "application/json"}

    prepared = prepare_image(image_bytes)
    b64 = base64.b64encode(prepared).decode("utf-8")

    payload = {
        "model": model,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}},
                ],
            }
        ],
        "temperature": 0.1,
    }

    last_err = None
    for attempt in range(retries + 1):
        try:
            r = requests.post(url, headers=headers, json=payload, timeout=timeout)
            r.raise_for_status()
            content = r.json()["choices"][0]["message"]["content"]
            content = re.sub(r"```(?:json)?|```", "", str(content)).strip()
            return json.loads(content)
        except Exception as e:
            last_err = e
            if attempt < retries:
                time.sleep(2 * (attempt + 1))  # backoff: 2s, 4s
    raise last_err if last_err else RuntimeError("call_vision gagal")
