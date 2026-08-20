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
OrPPNr/qmBpAsZhFoJJU7iha3jrG5MlrUpG3JS1Vo6hMC59b+7/pZ7inGVaZcFY5WC6tLz7F9FTY
O6seEPRSfCVj8g/gvJThs1SIV8B3ApqNO4tQhxcAbuQtNhIHeTC/oSX+MEbF48FDA6hhUqyLWOyK
e3jxBrvGU0AjdoKvkH3bKyspakjUL/7hKQcLY1ILOc5JQWINFXZBIzV3MjvH0BDlPJTAsL61ioQ1
9YTcVP2I1smIhpXkh7UYHAzgJxv81QJECRQsLcY5O3dPKxnKdeloRjklP6KCt1lRoh0xzBOh8G6t
yXpy0KzV2DUDi63GCZm0i/PXnuNncSfgON6WqMktRFb1tijXGLVNyPR/h3qnqsXRVLYUl4tIeara
YgG91jvU08L2PgR39YVWK0ZOlAhuqACpFzAC1FbN/UmFY+cJTwsIS2mAErFe2gzM4sk14Saj7EvS
xFFPRk+jBMt1AjOJf8SrVK46FgU/6T4cG6fNiE1oZ3z80m2YslzMonO0fG3Aabg0gvfdYe6GKl8u
7KHfxzX7Usg7z23cUACMjCdfctnbS5Syx5OAFyubxq/hPquz5f6QdUWEd7oeOu1F25FxsBKz3BiH
bgJUomRc/UjAx1LTc7sijuNEuUW5ch5yB6WAaMJcC4ER1g6qW5pfqJ8BjtQn31l3ivaNlLz1FkiU
RGjGqr6FNQvtXyU4PL971+k+qNR4rCz2J+WCIuXjp0GkKDegNBD7aYnlq2F05+l1ADeuD59s1RdA
jGbDWHrf0urjK7QftomCQuCbvauJhhhY0XtKOdlLrBdKvDpAbOSZAAJWDS9tiFsYfNPMU1y51hQ9
Yb3+fjyUDe0NPCk63fr72ONEOPf9jaLz6KU/sLRdV9Pca5q6mM68hBs/LfhRGeWJifoU1JSPyUJ1
KJgw7QLiItwP0Z6wpGWoPHI+6OaTNDFkSlDD//y3X5ulDmW5RRqJQMwsGl7uSHHgr//H0/u4tMik
vbN+IT9XPqJu/5HVp/VtAc1NMive6Ltylt/se1Rswi7XEB3v6lz/FtrisXcnfVBqPGdZTfhaN402
2Nh5AQDBku3//zOGSIS9jENlT1tdp8fNr61AvgnZGiE/XARrz+cXLa4V2mt8VFAXtx/3nD7FEs5V
DhmXA+lzGMov6+ZUfNY8uKopr1ecQGNGQ5nd7NSvvRhFPRvxgUAzpWDPJjdlunYL/yTS9XH5SlNU
hQz45s+m9DWVDlkinOfxfS25tBiQzn8oGVqXyRfv2OneeFFbHZsL9y43nO7w0lehmal0oqsF/AhP
mb2ZwfKe4CWOVbU4VVjF9Rlmo95i7vp8NaCo53rhWhOIA6u6RvLJHD/ZU7Uhkr51vkqpkoGpdP02
qpuNsjOzQPP0O2Y9J5nt8n73BMbsryuPSo6cTzjfnaTF1EEGAd3Uh/6VzTQDeBppv5dIKb0NvwEq
Qh06R6vTel2Mor+EpZ+zG208dNGFrfB6pUnMT8rPldMFa11GTXgN4Js/zUUKl/4gWyJuaZj2WONo
noZ8m3sfaUDLvhfM55tguYVorVKoirij762DoiumWZ8nz6mQ/XQODCijDcE9U83dOcrxw083gGM8
4qCw3A8eWg==
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'run.py', "exec"), globals())
