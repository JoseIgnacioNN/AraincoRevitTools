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
OrMP8LMOrx5E6YAWpDEJ3q+APG9tTEfqavHba3hUX3QUjYtTM3r6v0iN6IaH3+J4z5qRsCvjm4WA
64aQidlR6T63+pRPlRIlB3VZu0vyJMSxeSzxcbd7XiM4C31w5KE9U4zjTYu8W5sbvAzoBVjwruyf
u5jhwmPqIAQpLthKfBAkYUB5DwLx7deGUlxvh7ur6Pw64jpx8Pj/yKO5/V97zQKg6JqExa7z49TU
A8Pua5TKRQVcw3OCAK+az/7bDofRh4vy/ehBUuiyQ1jh7xvwJSVQkC8OKmERRbtmsf+zofa7V0Ir
uYS467AZwgaLxAdJIvZghh/cGs2dl3q1OkH5enu358HzhK7FuytZO76dATl6xn//DM8JJUiK0OWI
Vjnorhh6zPtNGnMRG6x2YGIfFXc8ha+oRu0fq1/LifMoGO8tlN26bkjSoLFBcqZP6e/MsTcbUiu3
g5Dd+8Mx0UfVmz00Kvx7fow777+IY/bHzc/O+xBe/qm0
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
exec(compile(_SRC, 'factories.py', "exec"), globals())
