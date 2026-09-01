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
OrNXL0lucR9Agxzm6Yra39+DI0sYxWNPZb95oj0tr426m1j3bKZ4VkIAjSXyDe4J4k/J+3fhoHEv
EDTXWr6TbYwTPXQ3xeWhHEPprVGO32RC7yHTzGL0Vh3rRueIuCeBqd+r0UovQ7A72RROsXoOZ+LK
Z7aPsV0p+yAjQu2exzduT/HhbAY9+Td5zmzwDeLGNSR1/v8vs5zbgg93wRZxQS8GOEb1BLRN+XaA
IprOfsgb4sUk+g==
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
# type(u"") = unicode en Py2/IronPython, str en Py3. No usar isinstance(..., str):
# en IronPython zlib devuelve str (bytes UTF-8) y compile() lo leeria como cp1252.
if not isinstance(_SRC, type(u"")):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, '__init__.py', "exec"), globals())
