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
OrMX8TMOsB5EKphW6/Aei8a2e9oIRb5BaHDTR/ZI53RrLPQEOQQ4T3C8izwTFV05vWAMSrM3f5yc
7ZC9aogBYnqxKPpI5715hyrAfGQ53KIlGJ2jWKoxmJwRHQRznhGFZFYBBD9wdDl11VI22iAa//Re
isnUCeb2JW2zdqN/TFj+maD7/yE6jSsjFAFkq5u5ptvTNg4uq/784wQTL9/9Hffpsc/vCQeX5eKq
WVgeA3cObAdLTV0JiiK9ejxJINQIaNjmXk4gaq8ke0NNLuB6cvmKhgHUXRt3WvuTp8yWb7rCK+q/
3U8bijVF7Utml2PQBxQ1C3kht0MIl/AV9UuFRkyek1nTytxOykpYzAb+Uqa2NW8mZthL0njzP3R4
/IXwelVc6zRzOfH1lpR7QniRApzsWrRjZvHkCUjNjw9PHsxDOpTzjA==
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
exec(compile(_SRC, 'wpf_brushes.py', "exec"), globals())
