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
OrPfN78Kpx5E0YhFaLHgr1JaZEYGsG05taopxzrM9mGmJMbq47N9mAcZLiQSVF2vZ16tc5xTT2qC
66SgSPX0wefHR1KLgujhZxRg4sBPM04vCBlRzriBi+hJ4I1WEFV7U0FTNTjKZ/tGxYJJkzZ0pT4/
LkrcBCrz46mVfbig3gTzPur9qPn3ReJNVaEeBmZC8XIM/EkeIEj7k6m5+3pwUnEIijVm4jo/ET4m
+74Hec0WBIfgtUc6BHCUWwWFGQCIKrwsWEpFp3z+2mnOD6dJ12NsMUZVzQsXXRlVeisudJ8WJKd4
r3DHHCSBkUyqicyKYc/VVKarYC6SNr+tSrQcvExoZEcL5QQ0x7ljHUXwcFflqyxTh7EjJo3FxWWa
DKrCIDR/rZPVqJ9xw8dN6Pgvm236PoEdm2rqRmo9tQw7hlFMV7NhqRYqXvZaUxt7AtXOXoEzVxNX
ZHwJFTOXoQugRSEJN2UrNfbKSaZlWEWazJwgT69HNhgdHSG9HUrr467mm28JNZcoQ72z92czYPA9
+TnYc6mlTEFzFw+eKXTsh0fWkEc6VC7Jb4+M0p9E0fVAwY40hdTXhaDAtEQf5BLesf5qYupRwKo5
cye679Rc1MIviVG0wEycCpgaVOhlgZ2ZpbGz0R3yBoTBmzNwLjZK8mPyOf160uDiNDUtLleZ7/9f
IgfyxpOsDvNJwLt1PZpVH6268/kBheH3+IX92swiIjpEQ0qvmj/7rhJbZJqbQdSgaBvJ5J/EFn3w
EDV6dkTWWtY5Q4Tmfp3jpLILOdidHdFm0SxBRQ4ZU/X6/P0/VMvbdxR4yTGSlU9CPLqEe29ebKQj
+LWoLFjFhGpFcT0TZrvytPK3a5RIDz7Q1whsbkKeQe2Rc1e5dXGnh9Y2esWKx7tfqfFcj9erRqlB
skzQ3VspVdCaxoI32Lr6YShpa67zpj70aMrXymeH4SnWIo6ahbmmtV7GvPb5yiMeRghIGX7pkwUC
s5oWSdbJQRIsAzpkyXRF2Be5WHHFMPAwPYeoAkRyTaNgxW5ePHzitpGuk1Ab2BTKmvU7UtsT6pm7
aGebOqq7Uc9eXsu1B5dAtuoybyUk7FpXKpOarx6HfPibb26EnRjKt1Fe4GRMugkwFQM5u3cHlYUg
JKadOSjNlJtYNUYyoy6VA446wxGHiLHI3WlsSkc51JWs9OYDkw0pJ81n+6FQjrn4NCk22NfbHHVW
UJYSV9eXxZOeZpBgYCt+zA==
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'colocar_progress.py', "exec"), globals())
