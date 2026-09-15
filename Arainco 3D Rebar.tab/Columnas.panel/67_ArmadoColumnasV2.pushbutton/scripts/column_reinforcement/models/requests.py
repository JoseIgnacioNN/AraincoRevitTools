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
OrMHMz8rsB5Y0pg7drXAzhlynsZ4pWNxdjS6VzlwlGcDACiDdFBjkd7wr3PtjajOEa3Tfp1ItKxO
Y1nS8qyE6/obe9BbwiRDcia8mEitss9t5npjU7tqzrB8BfMKRQSr3ltyu3Y2fpicetL2KP5tgAJs
APAiLBflxMPYHUUc2MEDOF+cJJBhp8/+sFAlJ84nJKA/9Lhg1e9g4ZCRHJWxNFKezVsKSdKiBJRv
es/5CchBRWVCqfhALv/7H/aYHq9TMNE/x642+kTP/EThjTxvTZn4hoOgJ73W9FrcKilSaHCX7w7K
ETQVMREP+dL7FcigEvy1OwBSUCNnb4aChopGD2T348zV+lqzc3Cpch5fi24JAQiiO1UovCiPtjbg
2KLpP3PPZpUnVjG90uIHIlg3A6BeWVWB9Le654xmNBlAdHbn4UPaN69Hrx4NQwnHuTaVCYMI/7sO
qrwBh3TC+fArNsjSHGOYxyVdLmOFMj8AbWaMzXybZ/umb4Q=
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
exec(compile(_SRC, 'requests.py', "exec"), globals())
