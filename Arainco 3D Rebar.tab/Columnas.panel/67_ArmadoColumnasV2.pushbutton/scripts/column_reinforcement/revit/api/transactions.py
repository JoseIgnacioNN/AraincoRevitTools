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
OrMX87PqoB5E6YAVpAGE/SY14cp08hZUczwB2gQ/ofXzTr/a/DkKCPFTeOPu20MB/0Mdg+3Wgly+
i+byVXbLasNnvJuvfnXAAnA8SeWjjrHMEyxm5bUiGjMuso//TvU2iQzsHHlgTdo5+/WtOO0CUwUS
52qOFWzS0ltoLYXCOEDKXzoSbfGGFPjOfqy41+SPJ3yhNzKvjZ6Coym57ZoUde3EWOnxA6gNfKLs
HmZ1jv8wY7w6WyI87LSnqrgEdYCXU8ver8AoSGtQsKD1d2IdaYsfWcMbNL7oFHLbK5QClqGEqlnC
1YB9B1Objkwgo9PUZcuLqjrR0L5STLVGfeTlVEBzlsHcycRyg2rOh7gw8DHqttvBsPtdlQhROEZ0
1uKlsZs39joKWXkH1vDnsQdPQYWC1pqWxD1Eo0wE2W3MLpWtLDPNULdc5Wvkk4FDovfTew8DSJaL
qVdgD+iRQUFH//R1pwYvgUX3Qvw1kFOd6fmYfcwBx+rqcIj8EWcTQY/XYkt1sC3erGCkxz4ygmof
Lf91hh9Cq2QmxiAB5atyRgss6wn+5xrO2wD5X5cpEndu1Q==
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
exec(compile(_SRC, 'transactions.py', "exec"), globals())
