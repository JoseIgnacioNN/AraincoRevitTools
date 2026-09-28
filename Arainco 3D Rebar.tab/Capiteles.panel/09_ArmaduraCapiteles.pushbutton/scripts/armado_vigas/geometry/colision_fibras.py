# -*- coding: utf-8 -*-
# === BIZARDS_OBFUSCATED_MODULE ===
# Modulo de produccion ofuscado (no es codigo fuente legible).
# Generado por prod_builder — no editar.
# Decoder portable: CPython 3 + IronPython/pyRevit (str/bytes indexing).
from __future__ import print_function
import base64 as _b64
import zlib as _zlib


def _biz_ord(x):
    # int (Py3 bytes) o char (Py2/IronPython str)
    return x if isinstance(x, int) else ord(x)


def _biz_xor_decode(payload, key):
    klen = len(key)
    n = len(payload)
    out = bytearray(n)
    for i in range(n):
        out[i] = _biz_ord(payload[i]) ^ _biz_ord(key[i % klen])
    return out


def _biz_to_unicode_source(raw):
    # CPython3: bytes. CPython2: str-bytes. IronPython: str==unicode y zlib
    # mapea cada byte a un codepoint (p. ej. C3 A1 se ve como mojibake).
    if not isinstance(raw, type(u"")):
        return raw.decode("utf-8")
    try:
        return raw.encode("latin-1").decode("utf-8")
    except Exception:
        return raw


_K = _b64.b64decode("Qml6YXJkcy5Ub29sLlByb2QuT2JmdXNjYXRpb24udjE=")
_P = _b64.b64decode(
"""
OrP3Mr/qoB5E0ZRFaLWgKbnAPSIk5IXT41eL79YPEv54J8oFCU6Mevir7uvOVUodgj+5Lx5nVS3a
r/3B+DJk/rrypY65b8tGei88C+2AMKaZ66SzRWP45hk3p0kV0U2x4VmiPWrgfRRX0tLhtrC9mHvO
IpBiT6kPw1EhChVHNT41iepcmiipJqW5V18OyHLCw4Aa4UcaAddJ+bY94MbbONvNg/vRlec6eHzq
Ys+FZxfbqFYIS9ea2vK9c/xYo1BmPolwSLXGAM+WIqwPK4gQKRzpZsSylcsWEghKE+cq4XH1QKWe
pqFiInKzkOZ4A7ei7oF5F1PX/IcgvtcfQw1TaUSsdvyvlbkGuG53UVh0BRmkmpZtG8JKLxh8q8Ld
NJXytnvf0rNEFDEnDuguZtm8HvlFtzZw8nM+aNaZdzenL01UKw7H9QMDZVWTG8ilt6O5sTRELUGu
Vvr5Q8RWHFyfHipQ81JRWj6Qms34oH9l0CaAWGkofVn0Ph74rPOyHgV2LHpAqPhUQUvYlpVRq4N5
UYjz5NS4+Zh83kMcVwbVZEXp96VAGU7p7+UDAEYFACtWbLg2zz/g6dTFYCKMpudxlBg5IlPbC/mb
lYA+Jfx5OlsOxZ8gq6JHNPZbrjCoD54jNaHIPXQG/C3ZDuiFfqnxDk2re29fwMHPQtjoCu3V499P
t5gr271NTNmQYF0DT1ijzbGcneJx7fGgSztXSallyk4ZLqbDlXsXBaWzUC1LBxwIYt/3nn6Fi8bD
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
try:
    _SRC = _zlib.decompress(_P)
except Exception:
    try:
        _SRC = _zlib.decompress(bytes(_P))
    except Exception:
        _SRC = _zlib.decompress("".join(chr(b) for b in _P))
_SRC = _biz_to_unicode_source(_SRC)
exec(compile(_SRC, 'colision_fibras.py', "exec"), globals())
