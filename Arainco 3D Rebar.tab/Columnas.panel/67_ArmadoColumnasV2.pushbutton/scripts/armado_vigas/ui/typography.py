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
OrMX8LMKoG5E6YASpA+bSiVgJ0Ey4HlHMxchStl64z6ew/MWmMgXwcyyGnhALXox5CVdZYipgEUc
Zq9LLHuzVDWc1YeJg+wXlpuqeLtImYUX2r4cHENbqglAoywrI1frDAPFmgwLWOwONcYbEU1hcXuQ
7kDR1PPMawS4JK0vfe7QeBYNPwXVunBWLZrBh8in+oQnVXRFvcc8vewDPaq+bIVXP/XICguMETi5
oR5vN9hZMb861NikkqyI/LWzvm8rhO8HtcR+a0BEuLNeqFcdzEQeczWEVt4L9xCMAh/Zl+I3EA2S
awsCXGhUsQflrICMlo6BLTMgmF+7W1QoI0guAvT0Gt2J+ZFz3uNVZlTNTlGYtmunhFcM6RzJxkAZ
kRTuiraINFGusWUgi9vbX5ZcyayZNfYI5+f0BOWW6B6mJsSM+c2LrQ5MqPbHljaJSCLPCX27rHxH
b0WoITdB8K64x+LlLD2brHx/aGKkTsH9HEw9kPAfAChfzZIT+VBNiNsS
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
exec(compile(_SRC, 'typography.py', "exec"), globals())
