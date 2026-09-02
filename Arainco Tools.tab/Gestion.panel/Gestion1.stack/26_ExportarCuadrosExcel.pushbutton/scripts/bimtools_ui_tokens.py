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
OrOnNz/qqBhAsYR4IaLUWl4whF11QjN19TDm6rScV/hzu/el4w30Bnlly1jnNOcsgCxyIjAoB7zY
PSc9hOlofDnrYyFCqJzNR7CUlFCE8cSaO+sAqceQOEn/qrrCbHriduQynTNXdtkY+us/K+ni38Mp
1XRkCw+eZ8Le5erOfJ9DILjTJmh/Lf/tHm1pyiMmXH7acCibS+b9PGuSkGlcFHtY4HQ2JSX8L82V
EcMQdGhFYWpF9HyrDod3qnguXLhjYCioI3j2ZcLK89tJK1IpIGXTgzY32GZL93X4SP+fWRUn5gUE
6KiKzaP7eihV5Vb/VlNrvhfn1mc17ToS/vqIMJVBqy8zItktIOffD/cODj2UfbwliV9tnK2CLZE8
MxIVC+r9pHHDnGILC7sgvM91712Kavff3S3Bvg826vc/XiRDo9vBVWYxi/HD+55gadU6XKOMDUc8
nbvI3H3iUcvxaT/N5m6A41evK5Hp7DBp25lomvZwlC/lmS/apnk5C22ttADotUUDh86T2MCWr9iX
jUznrpNdHNL0mmyvE4Nphy3v1Z9mCmPWN4bMz1Y7nPRm2fE+DsVlGLF7Admn5bhDCanpm7fMhCXZ
UQSYr1Hh94gjO3mOvDJbued5s69Vwqsf4KITxBtvzRnLHiQIzTWZlO0bs7EL8CY5fyyjGzSQhprw
BhnVq7Di2NIAw80bY1kD59J5nLjlK1TClXp1jQsIvmCFzFXEiIKG8gDsmVhDXV8WX5/y8EDp7bzu
nZOjGBx85LMSt2ewZVyrvdqQkeSZj0bqh/i6OW7QpM4LNwaiUN/xNzpeJfmIOz2+W+n1HrealEJx
+41pows6vHOTG0S5LeVvLbo54q5iiiMJGxvqhAyAFpddQWfBswmqvGHLlCe6Ff0z56Ngb0FWAOru
ePjbgmhzI/zfBZpeBK6hcg1jr6BdxureNwlqoYS7dlPlY94dWfyHPKY2Trc+oI1e4iME0t+OwYov
V0KEodYChYciK6mqVPojKE5k/xSUgm1SQeqtw9MNCJLCUP6vmdyMO3UhTDzCY+CTfwSJ0IlZW6E1
tLS8mIZZMIUcGP49BCKW8hVVpNsnF5dpHvQsRVNLZYYUbuoAhmqK1Wwl
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
exec(compile(_SRC, 'bimtools_ui_tokens.py', "exec"), globals())
