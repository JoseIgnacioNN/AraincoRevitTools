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
OrP3NLMKqB5Y0pRHJr8QAM35OktWwbtCIs9Lr7QrK7sfY3LLX8MxQh3/i8pvCxLIGe7/Xb8YpJ2o
yg37ggoAIy6iW+87jpq/2wsxrPCbTycB6IqWbl/ojNwH/09i7Za1nZj2CiqkN30HPypZ0q6IbO7J
0DH4JwtulfpyBvwOzlU1tbigjCWMO2jtNbjB30K9NqJYlDB//HuFLv8vJnCve3Er+vY4DnhkJw6+
RnybQ1tEvZ6GjJZUeyFhZ1qLDON28up0Q3R9uK8ELEWtH11kNhJBrMFmhbAQERupYjJmt1vhkhLn
hW93IemmhamY/3+I+cJxB6y3dk643b89ATNHZ3WIT8ogxvZEtsR7ObJ6p2xK/hebWg4rFADt2WxK
/Cb3aGQnWQ59DK6jzJzlutVLqX+i7SdA9LHn+61gyOuyUFCB1v/oVw9kjbHG/ARAjHoEyN85V9oV
RN/V+m4Gr7+6QhpDEg+X+OfiTFUK4YXb0zIjvvaor+n4b+LhXbNdQGbbwnu/R8UX4iv/jIFh12p3
ONIgiDS2czs+lQx8tiZWnC46njTtvCFI3N5V6PZ4vQfjyRRCuOQC8fWDQ6bBPeD+WmeF9jKVqQsI
l7F5N8yjb9f7c+itHSXEv6tlBlsC3oGjagO2FzFQGYvRMUToAZlCpEp3gVZ1XPghXouqnoZEYvef
V2I3ZF+94kZw/TCjxvqE2nBM9QD/8ddLzM/+QcWh5PJXK+R9nzz5tjVT0Fda4TOOvJUKsZHeqtSy
g1wEMFtEN+GLnibO7df9yXZPAAPgpZZoNXe9TnoEl+hDKZNI5KFzmQtZqixaazV3nnQ+zW9zhqRy
GY3KiOSmbPK8xJ72vzpcgQcR8I57XQmlOTp35z0oab4QfW/fRiUh4WcPId21opzJHso4M8MpeRsa
i2umC80FXpuwLlRWh52lCh3nGO2drc2O8rYRZAQ/L+forNrxtqLNdeT/5zAZiuzDhxRv7mMwx4bD
04IyBfrf/nrswQcSq5sJw5vuHGVBtexRyLkSQw==
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
exec(compile(_SRC, 'dividir_rebar_punto.py', "exec"), globals())
