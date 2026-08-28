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
OrPnMjkKqB5Espx4fSdRatX9oehmLGcOdy1dZ/LAX3vgBXhz9YQjY4YEmmSTCPXdAKXFuWZQf7P9
ixKcmjCgJinBylcEE35TPLnwzmv8JMTJCCOZrImNBsCiiZTOJurDSFeBL6dsrFTba3Y58OgyP03l
7sQEbvTUa4Aav/hMCI+jEzhVbF/jnL0sgNm6MEjoWq68pGDDqUl5wCvjCIqfLv/5apODr4XAtZTM
kFTn93TcaeTndsAdA300v6kJ18MxfNJsKqemT0nk9Endr/HTv0UH0YI5O/wbdxn4WSCri7c5MK+I
Hh7YvLbyXlb9YsRTW3ierfMDco1YiL8UmkGj7qs5R7SrxfP5mPoi7TIAyfGSS1BgvEyDB5/l170a
L8cUW04eR1XTWJc1heOhy+KAxjo9PPbL+kd762XHoz1+leXsfUljHOP4GSjlGuZGXgHZMHgBdq8Z
ZZ1sYObSNn957+JeRXRmodsb4DjMdoDacGcvgU7hAZQZqDbZbwJ4K11bQjlu4kAHToTN4ismvTKq
nH4dmptZIUBS7mzhlNQLg8Bu9WEG3gjCfPnzyQSi90KMG+ibUF84P7SXA5f8c7Jzw/YOQyZ06Pan
4FsNjECkHtOK3sjsgok8lAJtCiaVvguJxsCowafXsRLd/cAbGYPElwUf030drJSARPNOsHX7lD7M
ZfCb
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
# type(u"") = unicode en Py2/IronPython, str en Py3. No usar isinstance(..., str):
# en IronPython zlib devuelve str (bytes UTF-8) y compile() lo leeria como cp1252.
if not isinstance(_SRC, type(u"")):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'armado_muros_portable_path.py', "exec"), globals())
