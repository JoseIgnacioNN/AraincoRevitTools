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
OrPfN68KkWhA0Zg/MsFpvEEqB8pE185TgTLBiSlu0mUXDHwaEiPCSYie9wL9/qbEklCadb6uOd6L
MYZYAbxR2NtYap/kRKY84gihW0mevZjpF8yRybtSc0CXmmkKRABP7w7WI9O0YRIB60T2bBhTe1XN
A+IrV7VKaD1jS+x2cj4ba3uAefZl5eyl3dbg7a0b4lD4aQnUKcg1Xv9x2NQBjPO2l5UC0sUXt7OM
oJuFAGmqRibNyDO1Fdho3LIBsk1eLoUk3gWPxBUauIBj6HPu3c+bXTbZWzLO48tTkcBlBmhpM5QJ
TohH1SrIWB7pKSKarlQ5VhYvbhf0mxKnDmzOm64y7ICKis4nvWfNidhk1hlknMM8Azw3DiiM9bVp
bF3w1yDTP14+BPOyv54QmWsC08qVFgeNealBBEd//+4Ti/Kj9OpQ6Y/EV3DJt0J1wRvr+X3dxnc8
Wz9h+Xp9qzroTrl6GPe0YPGmeEdbV/PhYn5S+yWXHdjSWBeJss69WDuA8asXzCxMchB7+KDHVPHK
bLnGF9wm8pvdNuavTeumO1SMxHJWZzbBvhRLl/CHXJvHp3N/BkbFtHyAjZ7g5rpu/Fv5NmgG7fhr
wIr2SmwSWBcsQDsgDul6minQd2ciGBClaRh5aCOg/HCYFWMy9+CjgtHTIX3lA52lHGfYscmeG/Lo
cmmbr9V/suWMFKq2v8ACU3uUucBxyL0WEhvAq8Pbb+jTnRFpjEpjg7xUk0xRLBQMOjOjtzzllYxs
4lnVtXMNxqJZtaLlWgg2UYdaFcHQ/Xcwf9CA7Z0xIWZrGRBOtip+KTh74zRnNewptD5O2KSrtAMI
OE5/5yDhjyLOp5wWQuBSw+POZYjP9jmTkAFCaPs0Cg0tkLaoeV2LmH5coMoBFmY+7Wg30Sp9rKZA
j2Jj8y7OFkpqeNiGrgs0M4wjL4fQ/p0dlWoOKRGDVm/RC9+NVgSCaI95S6lqO6oohGH/xWtmB/z/
yRy22sF78qnD+8m1Sk3dpBHx5jnKjHP3Yrqg68gWBxhlX6tK6frD+cfWRfcyQVOmFeefmHH1BWxc
YJ2hQd0rwKBduHeefPlYgVKGn65paPPdbC+R45w0oHh44Id4HaZv+uEvgv/LWHOjJjAv23qULy6p
oqUwCNFJD+kP32uvuxyIeF1P3k7m+eAhddqsFnx++DCpu7H08OYVdwcFHMRy3Yf9lK7fDUMkx901
3emRh5qCmmw6hMR8HEEnW/+v5DxvYvEp3MiRMDwE8tYLZF7dJ93MaLVrF9k4QSHJNOqrTexaNgJA
Tw8trPfN6l1zP6fAMf6kTh6dn7nBznuxca2ruPaX0Zyb59u9JQhL73uQFJhkb4y9Y5WSaaLgN6w=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'constants.py', "exec"), globals())
