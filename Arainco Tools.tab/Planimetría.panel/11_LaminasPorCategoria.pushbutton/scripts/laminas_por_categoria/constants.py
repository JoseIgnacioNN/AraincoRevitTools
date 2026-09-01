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
OrOXN78KkBhE0YRFKIo5PVxs1XTtV/h9Ce7fdiWKPPicD1XjELXaOOYLh0vz8CrPS/4M49i17ej/
grG0lGOMFV8jNEOluu/0DxJeUfFnsZadZrVn+ggCZapWIJy1ZR7bl7BNpxC9xWPf08WPkoZnk5Hd
IIIISt+wXdoJPPqe+lo/Mu66TQZLd/PYCtBi520cLBorE62H+PEX6F3hxEvvXyhOw649eX9CDcTd
NG1E3/XeBzp2J8+1w1QYo1vB/+weqd2yTFkwNyrQ6TPmBrTTpIR1Jw40sxrdTVpW+P9DZ+7NL6hC
LiBAWIlVYtU5Rjc5XcAuHGi0jo2Vn10g5DFHiiKjInl7hgbVBUYtStYY8CRYAMcPH5GIyzpzuP/q
WvyN7aZX/lqGNvpxbnAfZHjUsmHmrUg70DgPM+WyVLSszk22rSOlK/6fJjnGvry7JGRqWqjUXEO2
8VNDT0JrOh9QXQmlx0Ah9TxohQRRZmNqLTZafc5Wf/tKJRp42b4LexYy3t7vuLkech0pME44xScH
7nfiH2+N7pF2vbEwzH92GA5kgBdgqOhiB27o0ubOjurte5rhJ9XLosYCJXvd53rqdZBeqWpCqPor
rFYGAFEyfa8Y71YCe1dXUwD2QI20ifU20EIh81n1l+TUa1KgDX2UfqCS9vAbGUIO0LF93muxmsiT
Z9ln0bbuVTZgCTIlY30OCJ3VsRbbextNsDg+vElwLHEB4/XK/A8hSt55ch/UNal/q8Rj740drRtA
jNhay1U8V4iVtjrfaYf9fmokBgGJupwzjMXdQzm/anoxrsJLcwImyFLGCR0FK9vTWSNhMudXz1DJ
gUi2wqrx+Y4j07+LTzRZT1GXaBQbcgG/XmDdwse1P+N5O6AWb0nY3MEwN9tVn6j/KwXGyf+khhj4
rlNliHMBcJ9vUsm5eRloVTjoBeKmHAuLKsJm/TfNXW7EdBGNkLxT0ccP8xOMmICnuv0/eLdoSDWu
CXJyuRQzUxYMhYIzXY5ZZiKI45p80RQ1YhIhNGFUmjqkwxSyDrmxysitKEGkvWOkAGxxE6WarDFT
BrXxFHiQ5ZdLCDH18kkrBFuDuSD/6NVPdRCTn4RoNulvncgbZ0b3Z6hxpksyBl0hFkZaPT/tokei
Gv2wJ4CAb4c+S/RbjNVS5ugKqbbe51qJHQPH4+s4sD41JgxuhBmmCwiq6BPMpXFpYUoC8jEWKLSG
GNVkPG6K88tianjljb8NQkITdFFIAXiIpOVVLOle9j1V+apldkQiNqmxmM9vFxSOuBVCvftr+nyp
cxlPVW+q2s/6OAsD6/aEmgGXENEUOI1r4uGPUxnNTUOQCqtSV4L6c0K7Psw8CbDdWZfir0y5NWQQ
RnAOd1cLkwl8fK8zTxtuMveAeE5F1ghvJv7u5HV9e+l9nvp3xWR+6pBkIEClbld+wQ2e5BsHU5NH
Vsg+t4l4U2jrVxVSoSKS0pleYTcl3oIyXfib0CKexGSZ+xoOq5TAj1ZNUbHy
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
