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
OrPfNi8WqBhEEYhFHrp5tj6yi5sVqZrgAr3Hyf+8x6y2SOnAUgbWTlhtde66vZub0SmnXxz/rcdG
nWuPwWYvZwmm4kPCm4jDep9ydVBvVt4GcuEsL+te5EYSrHRiKKil5FgfS3cpNb4FQaGfWlP2brZJ
G9l5B7pGLEyhI38RIBMNTa8gmh2UugVNkZJlUkf7aAtHY+L/96VBXWguiN3HsUYwXQ1MUzesDr4L
IXdaU3ZvHNUclJMAxKfmL+NJf8hXdupOAzvq/w4s7Sl8j2Q7Bf/N0yjA/XXRMR0jJI+Pg/wZz1BA
I9C/nWjqR04SCdlbR8XbwegQXYDtX4AWq1xaY42qBYBrBySe0HsLYNCweFlYFbRaf3H2o9JIOmOa
Qx5GgLjBM9xpLs+5PlzwafT4rnWXr2C7dyVP/F/lN/KHArayGQb+87APD7R1xoQkGgcM/d+jg+4p
tLxB1YyUM+BzMrZNMX48WfHh9R2gjjbkDGB/PlG+ocU1NpfxqB18NWLVZVspc/i0LxG8Y/0+/jW8
vrungzx6ViddrVpP3J9aSorLL+kIhhpB1XHVUZNXbE/kShECJE69o5hignsxQ3GlBoGsM50EnIH3
p4xOF9+ip5+RtdzVEmRAO18P59ZGVxFWkViIb6wp1RkeZhIy7Ti2fODZYsZlWolf//lXMyt4+ZQH
ET67NEj5B+o2iuWpemIAaqfTJ///Y0xfBH9XmT8tB8iNgriDj2xgz39gU33BsROaDZYizEQT6zBe
WLCy7qOhIFEaZBslZDTIdfZweyI019Jod1+APHLSSukMidXkHFLx2bSIfIgdUYpVbRaqZJ9kTuyS
kND4B4LqZDPJiHLyzGAZUcBkMONYiCSY25zR4/9OxcAk6XUvomXOe2exBQOq7ZQH8XkupqjtODBr
D6jPAiiX63XzdDZuX6Vp0maBV0LeWuDhreAF/pnVzJMiiIq9bd0kCok+jvkLKdOzQs8o5nlRCwvd
5MEo+7CXcVCcZe2jb0C1yVy8g8gY5gzDXHQUIcrEpyxHk6zXX8LvCdPQjmCVJ1QRLld1WTwRrkv5
qU3100mQ/8BIij+nzYH/xarBdydhDZuXRNCyVOOyVm5U4jn89i31uAOW37WyCesl9IwIrexXqPDZ
JX+NjuBR847tXKW3L6DHAEkz1s8ulYawfZCdlJX8w8mDDZoTZiUZPBoSN9v10oW7movrWaJDNsRU
rTDDXUhnfiXwGhmVRh+uPDVNT6mP5u1Tef6q/s5BaOTI45Ix5ylDjPCpStlJlfCDsmOD458Mr4pt
c31bqbynjqwAbx5GcYjkuosgEEv/U40d6KGK+YIWO5l5KaCiFuxpzGOUVu/vG//2DBFmGOGzQ48/
43zErAT3Vjc/x5x1pMomAktlaY3wrr8RE8CXmONHX4EPw9bJrha4A76U0P1Q3f5hpHHQ+uYXNZvL
wYgIYKOZlcchS1MEoir3BPOHaEXkEUXRQl8YLvbx41Eb9LAqBKiLfJQa+13ru8XjQkivIVYQ669/
aCbAiEcm0nIQdug6XtEEkhg3FEDdEfWfrWRwVMQdZ4H8Pu4yEuVrvtQ9g1FMiu4Q69Ny+dXcFiDR
EXWZEQZfOVuAIsdUcB4P4cCunJzKKJhO/GAoZRF1ton1tO50Ey/ZJm0S94xDDlpOw2fzjjdWTi/H
oJCznKmyUi44/7W6dMYsomQYGisPO+m4kTtyhpwb/4N9pY15rZ8pBW0qjBTYKuUKEfYqP2V1a3TH
UIn7eGV/F0mzrqQ411wXH1yi7dsNfLMOA9ABBs+7WqLj6LXqd/5+MzjYAbgN3Pg6+FiG8Q/llt+U
FReN6sVcKDbvwLqL6f8agbK87TN076gvja5TFE6x4JLlwe4gNaOLpXRrATVq9uMcMoDfuQSWQ8uP
G5FBzDda8eNPt9BA6fN7
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
exec(compile(_SRC, 'confinement_dim_link_schema.py', "exec"), globals())
