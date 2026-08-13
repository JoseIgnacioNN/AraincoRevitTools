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
OrPPMr8Kpx5E0ZRFaLHgz7myHkdo1KjBa8XFtyzrvwOMDcspngyfWo90YK39IJciHDxp691KGTV0
fBW3wkmQqq3yVfIVVAwQU4IlDxk3ouoO0O3t7lE6rfx5jQVGP/rEAzAXdY6YEqMCn4S3WagaV4yP
OmZ5CdSr2eg4LTrKbKHwgzgmWkHrqB9t2whCsNq6pRZkJmqe6eeB69VHpdkuEOO84kdZJvwY/NGZ
/nrTKSg7vSPAlLOrsQyJQor/3VXKuLd9DWz8+JgJ1yl1Vci+TIqPJ97I7GIA8q8IUad3r/S73WHt
el8nnw6mqteZlac92pe3RuJdwuEhgzue+pmqeqB0zQABt59iVHAh4lGqkY+BcBIcuGS8d4OZXOnr
JFwM5KEMOSFVr1YMVeSXoK1TC4MDE8i8ll3SVOKgMxCDHrLjLgUwLOEzh88uXSq9XOJehCAGGIZR
vkqAnyCO3Y81x/a5WeLQYpgs5NEuQsUxszj6+YAeyETq7pmYcZsBKuglW3S3bsTirhPsiiQro/pt
bu6WWIvCZqS3jE2ZIQ01llM6mR2RrblMYyw7VUIXsfR4f6m4c9uKICpcRh+8WwTszQaXBK+hh8fk
L5505hfUeA==
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'singleton.py', "exec"), globals())
