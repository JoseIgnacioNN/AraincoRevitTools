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
OrMPMM8K8B5EsRbmfeoyZP2FB/ecdyq5zP8ct2lcE12mUljkRY4A2pEFZe1KsFTw23GWLTZIJMPi
xddSrw47xHi0HuaxhW8i7a+/PsYuDUs+s1nUO5r7KHEWSMO9N+ywYOtbmdLv5k3GCTbKM2KxSNEz
NpusNo4uO/46aQc8FTBANVLrp4zzElixdVjngQjOjL8oJa4OOD8hrGE33Wqio2ogTRSTXQScY7vv
PvgLI0UDXAJrxReyS0TsGIZe4/IsVWGkwQNlVnGijf4J6MXiQ4KyuFzgS6bcz93KTQHQVd4QS7Nx
jwer5ZkaBOCCNDoA3w2AjZOQPj0NDnSVhaAAzMd09V4UUtH2
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
exec(compile(_SRC, 'segments.py', "exec"), globals())
