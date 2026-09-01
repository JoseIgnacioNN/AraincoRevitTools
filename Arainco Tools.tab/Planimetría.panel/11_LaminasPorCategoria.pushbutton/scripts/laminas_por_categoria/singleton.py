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
OrPPMjnqoG5EsphFZDYRzKBFe3BtRuoiQlQtC2pMoxuuGxIABT+uO2bcjHxy0HEfZZj2q/rxgQST
pVF1745hh/nCJNwlq6ZhNJUbZpd7/jL579Wpj8aFJyQlHYFCVSFwAwkWAcbdMvaq3UkrlquEcjpB
MLagF/OlGLz4BovSgBbkay/Lhsp1g59awFFHXdT5SKiJea0HGRBXukNacsnjZgh3Ezdmi9sKG1nI
VHQoeAt99ZMp6UAe/55mTEz3Mfocnk/GoTJ2nPfD2O7FkDe89QLIJVaQRett2Y/NEdU6MCvIPaKU
2kHfqUpiydXFXJNyVHVjloqY6I8L2jYyEno8JBZrgTr/tpt5lQbVXa+VllddfW8p5Ig02H1ItmQi
4aFvmIsbluAc9aKlXire7cogj0Lukzu8Ih/0RkW2rUt490fD5MsjlLYY2n+NeZE3U0CUvGNviCnW
dytRlGcmNM5e/GuIsTCT991spte/T61Wf7lVpvEsyJ1iojaV1GaIfTPg5q5dptIbUN7wp6692jAC
q7mtElgRy7WJYFRyiUwwA9OYxo19Bvcg7SS+0ogIrAWYKz1Rx5A2nQpE9ySykTbKzhEgVxPa8iSq
KpvlkFPP1bQNHkuXffpQ63k=
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
exec(compile(_SRC, 'singleton.py', "exec"), globals())
