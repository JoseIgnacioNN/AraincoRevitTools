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
OrP39c8KqB5E7xhS62Uvy0a/1M8W5FIrUxABf0rDWmKQsBtA40j8XMYhEvYsTUiKddJpH/RrVjrT
GGzCTm6iwDzWQZNf+M8V1bN7JoG2IpSB58hCdzI0JNY6cMjQZS4dbAM9py+GSe9z4PX4/2O48taw
KeljMTt9wCRnHKEIUDeWBT+u5jyaMULkK+QBdHlsshZ3FTMtDcn8VRHpBN0k5WdGxbR6D8ZIBxEk
ZZckZBKcxILpMMz3YeLIIu4Lp85mXk4p9EXC1xpQTKGrBYR62mJ1nCEwDhxXWShAneANabLVPNxX
zVz5T8swCmRGXdtTXKsccz+cVwzPTMJ2JgyC9f3QVvCqxPfiLLnj/64Z+spXO+uL1IPGoFVxPB73
MNpimNouhuwdev/dYgI74hf6m5rMoZiQhg4ZGAluCB5NQjtjYybXc7d9XZ4bzBQA0G1hHvf8hJyK
j5tYT4FkDQb6HxjNqlXaR+6q8OzwOUdiK60KKMI+SzJBM7XQvGlhoQH+gSYCl/hUHvFHoTanJPOf
0KgGM0EA0DlFdLPUeBLkp0v6CI9KleiIuat6lZCf2ChoarfzfHGnQX9nFGaPg/oS8TiaIS9+sEK+
tmM2Z1UoUWk5bYgUxKrFekZJt2Nrio3oT5fsoDSO0Wn9Rwe0Ak7xhk6Z1rrnDaYfGWJvMqVtyroo
DrYz2f6D8kcoUWb/6lLz4RhzpUp0TdQAyf6hzhWOD+aagqGvCq4GWbQ0DXKYJPVcgoN4PaKMODUb
xVvP3afKIQGX2lOEcenTxkIZttYBgrUcXYyboXaUcPikNkrlfGhRkXy5yz21uRNQTzheg1WhFJ32
Zfv/Jb9tvo07Yj23rnUOlhzMU16B9lJXwb3swRfA3XWXZY8gflmRZwvN7T80NYqvL/rbyOdErQ==
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
exec(compile(_SRC, 'bootstrap_paths.py', "exec"), globals())
