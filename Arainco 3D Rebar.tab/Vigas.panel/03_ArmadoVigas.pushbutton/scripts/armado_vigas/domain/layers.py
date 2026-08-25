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
OrPPOa8KqBhA0bg/sulFnYKWxGFq9kJsEacdZ/6IK6JAYvAhE0CpzX7+6UQrHflvH6byoghCnx25
mvtYt72aLdD7Ujg8OB8bxLdRThHOnZOVMnDg+7oEQzf9+aRHurNTpQLoZfGlE7H59+k45/VaJRjg
Hcwn4tikI1oJyhyWZHp0r/8gwqwmKHigN84Mwpgexhxphl4jIXQtYnMr6HV8OUGprfKLzxvRkYSM
v8apk57MBdWdBjCIkPaH+oeH78XIqcopPqY9Eh7SLp9k2KFa/tQqWm+DZd2Cjzu5c3yQZqMJyg7w
wSiyh9Cf5YteIRPRI3i46e2UazLvG5atQz5lzmuzXDrDeBI7dGsgvQxV9V93FCzSrKDR92jpIDRC
IvQtj6Z0Pze77LARoGFjoM7sr2jszl5PIvBybZDIM/ZZig/rbekrwDM62FIs6A9HqjqpUzvTmyTt
MDefMMylxGC6hNIY1MUKq1cefHv+3CKWUnD3W9sF8ASDvrD/TpzN5mhKMB/Gvu7e3WREawc1ZolH
oMeZszjhg9BP0pW9lyl6OWBG2mONKFXOXAr71ugh5fIxsmoclVifT0Drp6MTg0n/UduyIByePZIu
/uDyeQfrlkZaAUjZ0coKbA72NxI1lAjJsA9KP1oQDM9NDnuYBK+SjREYFA8PpggTKrce7RL5CgSD
IbtpWGMIm2EB86vNxBfEjLyqM8o/KgPLORR0C8Ir6YwZsuV+v2ogVliJ1BORn5cW/t6nUHSv5LoQ
1a4CxLcXfhRXglqBODgLKL2NcCIkKky44IUyQcuHAWdgLPMgtHMF8fFAEpJAZi0zx3DFmv1Tg7ai
KP2/NHPlgq8Kj86qXGUIDz/2PIXiRduLZEG159mZlaU+xBqQRk6Ia19qj5dpC7fe9xfv+VqDDO7E
+Kg4fGjHopQ3Vcsse9BDY9Qj8GcHnyeyAL9pWHvS9ZBoKtTLmVk1ZpXcm3qN/nf41OIPVRrzLgs3
/lepbsK4EcLGorpe9S/FBuasbxPE4NlvGzkGpesD7cem2C4WLaGIvNfhPhfFjXXWTh9fkOLFtTnT
nSQqBbGaqT3wEzm6MrPfP623zlSRgZDjGIjksNk6BcC1n6iDTJGHvoB48DqNTY7lBNAOPWqNCzqn
pa9iLaKsoVCLbxXHF+/MPu7Sx+bzq5gHDoeV9mCaTQFaslriQSsaWctDpYRw6qs+R/qVVqw1iUZa
EdFWxL4TVTvMi5CWec5M19CMOAih01qnuXGk3IhHeGHiW+70LsUIFrh3lIvHnbwmlKCeUuZx2qmY
PZxVcsZeCiRJLxEs3avCIFZ/bQE7GnhRP1KO+GX2Us55YUkwSa+sooXHF/bD8aW0bRiQsPSZDDlA
qlQO4tU+GSMmG8UZTy2ddFYM/YFS0Jvv9epZqTpkHt5OpzX090E03bnCYAqPZ1IOPbb4KdnYIZwB
9sRl4AvvNPaRRm2+0SIIDCh6IkWau1DrOybGVgEIHod0dUYeTGF3u3Q5L63qyRYQ9YVFKQ2QV6xv
4lvS5RdIye5h0bo9xeKmyWK8EoVNwctxg7qzDm4aAy9V4D9Mm7Kpuf4bGFgHz7wqqsysGyPaz1ph
N/SdygzTnw47WMX1JEUy+LaLoSAtSst15B3sr/LCPj8zf851pzYmt2UnKF3EGVpo0KXp7khJuds1
AEJXDM5IIylqhuB8Bdqb6N8XmkOk14fzaYCC2+m8j85yzWV6BTDNfi5rdrtBAzUrbEvvosAJZWEL
gi7j4IjEQ+i6Zx/74cjM7AzebaKdZtoKzK6bklLvaI89CRKP6vC7Blfb7GiY9ITemkreP79M2tLW
r9aS3QDJMJ9FeUiAncauZfIW+p2TbSEEOZlZlBJa3im3GSCVO4Ws+Q3P4IpFsifdz2eQa60+Brg5
/UeNFRk9ECdOGtO73Y3F6bI3etRX2xVyxXMxCouVuLmNgAWFgBKIFLsknkF3XtXo
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'layers.py', "exec"), globals())
