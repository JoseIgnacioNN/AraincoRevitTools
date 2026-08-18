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
OrPPMjnqoG5EsphFZDYRzKBFe3Bt5r/DcFYtC2pMoxuuGxIABT+mw52j4S9eXVyvEkCNK04KVXyA
fOjPeV/MmxzULa6XQ0NRbuizeeg+Xc9py7M7cSQid0zoZINEBn69t2gXD2J6zXXHTJOU3EWsdm1U
J36ld3AOfEsVrDquPKSuA6rbHie4imH3NcYsZB4D9IaC4Rcj2lL9I8X7koovYP7M6PkeKgKaMp0L
b1jlJeNRw1Md6Of6qZayfwGJcdYWi2aW+gl721AwofUKxzpGmVy3wdqh15LZcDJz9LLYVVuLgwZ5
owTplYri36tLu8cx2lfSP8mB2KhuPsDtJCYKnhOPxnTuwkhHcGGopatSS/140Dulp4bw3RWl2yN7
6K16SCJ2ee02V1X4RmoCEvzuyifpoVLvgmXN0YbdIIKiYwPVt1oOWxCbYbxea+LzXNBOfgsP4GcC
P7FgUwChZ/E6O92cZ/Otb04fPLsPMS6BpaRxE22avL59+e5hLNxvseqzvOS60lAjNP330n4boT6b
pRWjmk3mGa/oGabgdAAhakXVOmmBwBu7qaSwslNhKCBVSBecNHDuiZRQDpbI/BSS6W/GchiAd4rV
baDgv46nmEho4ck9
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'singleton.py', "exec"), globals())
