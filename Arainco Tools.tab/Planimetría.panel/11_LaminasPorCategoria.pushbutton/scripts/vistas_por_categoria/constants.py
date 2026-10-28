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
OrPXN7/qqBhE0YRFKJrTGYVboi/QPH6HoeHAwnTnJBSbsN5Jesk3QV2nzcex9Bx/RCGedY7QxP2T
fJJJF7x4GDtEG/Lu3uQzx7AfVPGcDrtHFrSUaifq6oMCaCLTEw2XFGWpVV89J1sGu9PuYU+M7enu
EHeKEHo1M4cyNrDT7fEqThSqVnEEMV+nF/lfCGzlAAl0Gka+fU9zIszKtvml3RLl3RnMBR9l7Wyp
2mON5vNVPnRe2/MOSutuPuTaOMQBTqJf8BFfZZ9IEGqjj2EVHCJ5PirJvdYoLJfc0ExCP1zwHxnc
DSEkGnDBBkCzmABts3cevun8kNHpLhgUopcX9WMwPyVd1KunCcXSJbfYUPrvUIWLG78LjCaUk1+0
yT4/LI7IK9iVcQKVHGZrJhRX2isPkZIa9906YzlI/oss5EPVDTf48kzFkkDnVSz7vMdWUQJdsU1K
B2uidDTNwEFsKsXXHrCo0u5FMlbbt/qGAlwUjDktf7ZeBASnuZ4tkwvNc+0+ipJgStBkYrgLr9s2
XjhOYb82YnIsdXnfp4V8IGk7EVlLA0cBv3F265AIefN4xAddG+hZGqXzhoeAc2lhEV9mFhxRNp8x
3VIOaAp0bEzFadmJ9sWGBt6SZlIizVdF2uqHrBVVcAQ4Mr/hzj+fC0PcDB6+hsUfEwM6HXZ7O2TJ
00MR0l2s3/YZXZQDeUFzfrArWJ1pBzR6b1IJRcynb9PLwA0giJwK02+27L8+4d4d7JG1inQnFjNr
AJSF++T1BS8AGSlSNrj8Hr7CE8/cggnMdozN+7uXViWaoR3nnXWmVP0PMtqfPyVS+PaYsCaP7sHb
Y8X7RJ0m5r8926iUJyDXuQ/a3RklsAQwpi3OUR5YCuUn3a05YH9vGKf0U/hEhOD7oMbXBMcFP7Hf
YU+MDovQ2yeUq4MrdRcSqbrtKLwXcw54k4msN9QQ4vUEir+q50IN3KuklXo5aLLY1kN/rTYHAU3z
kokkjSgDunzpin2jaRO3d82LQdRnitIMuSBY4MEvbMTxloCdfLvljZL8Xep6xGAdxad/LlftnB8Z
pCcD2zi0X+71/5EDbSMBdkLwxR9wNfAv6cjUh7ADwszrRR/mTtcIrRy9z3PoTJOr6uY/AljNo5p+
hVLqq0mnS/+egRqSF50LJU+5BV+NIpwtCdugtKCActnYleFdck06zyhJx7ZoBnIJB+3mJmiz3qGn
eGGhszxnVCQQmMgyUQx11XuaMuomclEH/jjERC8eTZLt2J7tv5O0dczYaBWNcnt+PKvWb9nIcvim
strbNJJdd5F099BySfmg1GHKQ++pddvgW4+QrVc/d7f66hQblVxRK3rYV3r17IHtwwkKY+5HaL/V
ZEWL6PZvXzk5Y6KaaWPq084bMP4nNW6dkKfYa7SafBX09BmFqcQBasCYLmwV+bms2ucG2ORyubDk
k1BdhUKk3rMIefWNS9zovYd2e4rPV8jAO5LXdkgqp+WGFa0s/d3ZyurfSHergDG8nNtgnGo/mFqx
8LutkNA4MY5QIr1Cn6vuxFA6KIo5a2q/fUWu0B77TTTUlVv2EQ49P5TkcHQ2ztyNWgmDM3w2rBOz
CMJCWpvz3hlRZtiNcRQadAqKLwLEarqRhwt6z8fWJRHrg3Vnsh8PfojRkLOgPgBNI8mqvKbiKiPB
cbv2jzDRKFlw2ZATrYqGjzLZFcgxpdK0SM2I9Whdk2kMB94j0Vmjr3LXXhxt
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
exec(compile(_SRC, 'constants.py', "exec"), globals())
