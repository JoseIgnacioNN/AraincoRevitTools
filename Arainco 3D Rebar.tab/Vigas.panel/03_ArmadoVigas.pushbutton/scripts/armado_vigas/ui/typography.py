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
    out = [_biz_ord(payload[i]) ^ _biz_ord(key[i % klen]) for i in range(len(payload))]
    try:
        return bytes(bytearray(out))
    except Exception:
        return "".join(chr(v) for v in out)


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
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'typography.py', "exec"), globals())
