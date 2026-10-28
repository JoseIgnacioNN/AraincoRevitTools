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
OrP3NL8KkGhY0ZRFJqUQGHkrojzscUrbEAL4TQOtqzX/fSoJ+xJVN5nHnI8KbYdOniw0jgc6iGsf
51cmp78aetl98m6kBPXS0pOX/d1dbhMfLurP8ixuIm2/6nAKSeJUmzUXXxvfwTAq1BNrO9E010HJ
/zQCd7HcHolkkfiacMMjmrM1moS03cvAw7kQN42+DMcjb20gdLpL4JiuqxxeWsgjWyLxzmjFVsy3
64x2nR4h857Wu8GEJujSTbIDoTD0vhUMkJkfS6/MonNGfYpR11LkjugKOLIOX1flAMOFo1uxrDYA
3xizJxNcxQQnlZ/szSQICXOzIgKyOOcprpGWr5jshjOurdVWPPmq7xxFtb08gWuaC6HeIqxVTb+/
4CscYyj5kxII83H4zlLU0kz3h6j5Rz8mvOF9Kh5r88shKYu1BqqaI2s3/u+LfQmwZMdLkya4T5D8
axR4P+yLFacqfY/cn2tZxzlLVdbLG3qWsnpBI0R5/rxPOq4ddCbBJxhf7DXe+mmUOUL24Ie0m9mo
TKgKJRahGYZt1JultSRMSIkce50DTMs4EZN42lzwAaAGBpSU6r/AUpuoXmqANXtKMCm/a+Y4Udx2
gPOBIgtyYNq0ypYoAutHYKmgxF6h1M7up1khWfYMxXjYPLTmmtUDzOHlZt77OsrGPdRAP0bdP8/5
oAB5rYODv1RiBZkJStrM11BFKY2bw/E0111V4Mt/e2eoGMhWouGLAwxPq/cTbeC+HAilndKl9y1u
FBqL9Wofeo5cgEDv5oQEJ5+o31ekZKPv7WQJQEXRpjIehzWBG1XBIl/7ob2W14G26vzoQkLZOh6M
8F7WUruBpOrZcNPHrbG8u2uG7wtREMU/NgypuYa2X6GuNVebE8wiEff2F5qCMEXt6rP/BxjUAzf8
oHnk5fJ4iLiuwsU+6f3NTiO52m+Ywy25sNJW9GoH80Hni26hGJdcESloX7Mh+4jraX394XmC7rrv
+8XgKezSkmE0q7w4RJOWuRs92Q/bRLhzc/edAbJMrGYeTgFOi0vhcsif9BoIqQ6hIBNWyEGCGZ3h
1dOQdUiEKsbk+86qNtk6HwCXJJAOP5fRziEeSyaC+ZHYskXqj+QzW20ikgjy8Mg5Xcil73ndiE5P
vzaXJSE330AqGJVVhF22rXxr84/DM4W9MPN3Q15SiQ6dIJqmwKGZ2fvciYfiwNSmPWezjM3Rhga2
a3HXhUiF9Wb/cKbEoEmR42mCPc+jHCip0JzFe8dv1Sv+TQ==
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
exec(compile(_SRC, 'join_geometry_instruction_dialog.py', "exec"), globals())
