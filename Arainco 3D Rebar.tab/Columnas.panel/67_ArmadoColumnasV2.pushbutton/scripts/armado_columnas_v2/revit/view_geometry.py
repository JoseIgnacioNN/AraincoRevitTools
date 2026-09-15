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
OrPnNp0KqBhEkMHLDoxdPY+GubBrMMMfCe8Hd96AM+FMJHgi00BJTzP87Shc5vVSqLAu7bZiqH4o
Pf2DWud0CdSHBBa1h47bRx2bmKZEwIuGyMsK00kJ6vUcW/YShpOnV2yFm91EVg04qrkgBoniKYIb
8fEMMPYnv/57rHVrfxxWXu6lclKnp3pWNEkHuPnNAQOYRq3jKYr7EXCOhWSgu9Z1/thmNZcnkcXh
TtbHzEWvXfkGU2Z26EaTpR+bxLUe0a3RG9fxX3enKAVVx4mZOYfJxKlOvw2Z+QkV8NMN0VXUd7U/
zghaQcQbUydVmkkoHb/x7sTdFyjTIPD9txQom6ZNWJElV0sdbNz9OA/FTkrJwg/GHuW/+MCkkw7o
dNd9SccYonhLsblQMWcD5j8gXyMlG0O3pyfejO1CbkAVXRHel7a13Sivho9XJlYieHS360qB7ICq
hsHmAG5bOllwEG8wxnM8fag6+D2LwdPSOnBajbysFoMU9bLyXl5QVGkAPZmBjGWyb+bsfE0HUJN2
FiLgtq/LV4JOcQI7R9QMcyTEX3ORTDF3kvl6QBZcSiuqA+In1FT5LzyCFcfAJ+0NGtKPblL5dvUp
3MR7cnzQMgasbJSgZjiSKtgHjP4Ph5vYDbJi4vCwSVqSBDmM+6WjvrE+uMtcdbEIgGwDZBdBqS9T
e5XpSfGEKiDNgozjS8u4N9GxNWHtS6qo/T5egxegFXKFDSGlcWNbsUcQl7L5GJZhEoBIr/C51+iZ
qvu2GpwCihxOadBy1/xEXR0QaHg2U7J9LXDuTBczpC2yjYLjSbbmYGhPf4YYg9B4TshcZm13CSfO
057ehUV3yQmgHVnHKQS3iYFMHjWXveW18gOt7v9HYd+StHbEQIExYdNmrs0gslt+P3S9exZkkLuf
RGSPY4na22ktrHb3Iaplihub0TEKCncheDCLhZI3UdbFNAtBU3duCaefVa2SytyJl7P/TPexdYgU
UlZ1rIMJseu1C5XqJb8FXFEv72j1RvRXhvf+LFcqIkVHRm7sd3xCMYgGkpUFmuJrisHnTap5R4g2
a+iO+iRrKz+4bKvZTsdLv14U48DAu7gnOcjr/VvpZzMZcUlJh6GD43GiW+YwVEI0As48YSJxXbAj
l1SJMwXv5b1a+6iQFJ9e9fApJNhl770WKrKeas4CLcrLj4oXas8Ra8WD39sC72s2KW2/ps63pPTG
SBbq+QxohY1JaVWuz0IKtS8SDZ650sy1aIDMDwZaF1cXHJOlVGrCQznXgBJg/tnOynLH8wzBH5kQ
YzzXTzJuZE3F3Ws8J0iNHCzYWNktjcsYKrlfV6eCFlHV3phOTNSsbXG9YBA5ev5j/ofW6LwPX7Kg
unIWKMguWKjWAZku/OmPZ1h6rQorKbzlsCH5wTcJma5HyMfhmsmdGDS+bwXpPuFtn9u+ItLSaJum
vQC+8LpjCY39B6NkpF+yAOmU+wsCykjEwoXy/5oDlHRzsKElwm9+yVSeaItrJZ+uJ4JLmumUpECs
E/JCYqfAcYxfYb5qupQ6CrpQqw8NIiS1mvfTFncVTCWoK5F1g2vrBJ2UWypEvVb89ba1Zc+svYHn
nDoVG3eNFqOQtBl15oCp/8dQEVnuTELGFRsdFuyBfubjhhR8gD6J76Dqb/Ry2ZmmSHpaaENEBDmx
bYXp1H6IWVdrM7Re/uGLY0nvV22rDELJFqyB1Fep8o1T+OyiDj+vLI7tz/B7zorD2xZi1yNYXkl7
D9sPFIoWfg+UzsbeCMJZQlrLOGgDUxXVtZZgqzH1KJ1cR5H/ibihCz+oaRys11FSqxUie2VS7zB5
rHXkCcLfrdIjiSU2xNcImSs4yt+AHJWXAIU9tPOQT+qfd//t1T0=
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
exec(compile(_SRC, 'view_geometry.py', "exec"), globals())
