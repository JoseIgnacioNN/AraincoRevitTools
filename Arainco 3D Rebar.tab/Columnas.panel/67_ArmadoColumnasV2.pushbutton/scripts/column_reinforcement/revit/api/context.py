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
OrPXNbnKoB5EsohHaDXhzVA1eYos4FvXKalJa+tDEq2L1Py1RBr3uIuNfJQfljz1QlMOR9Q975Jd
ulBxoR3uDN1oLcTi63zOCVoUdQKZFl0djeuh+yaDYnqV6WY1nSRwyzlJgT7Xw0ruefMd36vlo00I
BvB0XmPV0bz0P9x+0PFYcmweaOVM/KT1UUhuXzcD6Gn/QIIZ9xZtpGuqqKLWsZdWObAmTqMbx+EK
nEzJMwY8y0cQYJ0LXezRKi0/bz5+Qonn86oyeqEViOLomfzEd+ngo+3dKCvd12+1PXMZ/o1a5EHv
qKVZUNy2vjEdO71kMGgiA4foWbIfv8rPl9xDue8/vKAVzs7CTerQ44Mqn6ACgX6LNfiBFYHxfP3a
vWqS+LBSggaIDb+bAiFPcCcThxxyyHLoisxDCWiltC3CU4oSRwKsCTWAIEpHHNadaoaDW6z0wSUz
3dSO/34jhRrICksuxtxacVT0hU4wPKxWTcLzpLRPLEKfjbeHhTkyw/P/Oppkpu45po6Pm1GykvaB
AwPlJeW6cpQNhdCDpDkEDS1Q4j2CG7p3nQd58fhbZEFLGNkuer9N/k8JG4GysAfi7Jozxj7F51qK
XjV46b6SZyoOn+FQJDAtZohDZr0KgyTsuNEQsaok77sYVFu9FDiY/07BfLPvgZjSqoWph1XwHAaM
pV3kMev8A3neOD3jfchIMdgDobZch10DPRACaUqJ+6rpzu7Kui0tHpYOH/ojINa8Wc8WlAQGuXJw
R3IOpedgvFrmspZryONf5A==
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
exec(compile(_SRC, 'context.py', "exec"), globals())
