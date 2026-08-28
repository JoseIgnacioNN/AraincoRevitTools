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
OrP3Nb8qrx5E0ZRFiKTgO/QD+Ojsdl5/WfR5vx34P1dVPvOMCPh2VL68O4aPd2wOu4oAsBfnPlXa
NQR7R+AOT/yQnNKtpqKVd6cizdqidgXbFrVySpeTg4GyQIy6WX/Cu86xygkHOWSkpsxKS6T+HcOC
6Lsae2LxJlz9rLJd8rhDXH2wQUEl1NOWU33MW7ut5BtBrDpg5TTUR22gPoMpu/iNn9ZoBdZ4SKUA
RQ/HvM1rcfIct1/kMwHEHn+2uP4KhxxINItkXSOR554Uq0rTuQRFNaoMmv5l5j8ssR/LNcHwo66b
ZTNCsqfEY2mDcefowF5jQE3ol65E1/PuSmw5zdmiaOhxLruiM6BXZQJZwx/CrHs/sssF+8fUQwG9
rQN7Od+QvUE7MVoyLPR53XM2FZPXuZCAi11e3sjsPNYYbLRSbFS78A00XZ6DKZ2nUAzXjFeAV5GK
lKSurbcz9qi02eDMbvtzET8JVRun3wIdVCS8aJ+6sxcXYCR5xdkWVspKXBe40qHgGsnKXwlBBdxE
kFaNv+/+XiZRMgW1CK+JapzFAgzeAqrLUbEwzaMx29oihL7WtYUzKW1mdYV/Fh3TPRrEBap7JW4s
fSmnMvbDRY3ENadaKIGoVQKkiqK+bsy6LKHNaUM2ZBPB5e3r47JDXjnJeXE9US1XfqKrjwV3suoW
BLNNx85a4Wcei89+/VmtTJouf9Vwi7/ngWfzr1m+jI2nmWI93z6NYAMB6V9zaQSGnn1JCiaMDDIJ
2b3q4DHeTcj3jjakfqBbPLSGjtZBgrUcal6d3HiNQmzy3Ce4Y4NlmXfRB9f/AdsxTzjegFWhFJ32
IVxg3tpwvs1FaGuG1o9AHV6rNbiFgIHyumfuwReA3XWJHZk8LoHsTv/jIFEkfuQtRlOacRoaf2E=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
# type(u"") = unicode en Py2/IronPython, str en Py3. No usar isinstance(..., str):
# en IronPython zlib devuelve str (bytes UTF-8) y compile() lo leeria como cp1252.
if not isinstance(_SRC, type(u"")):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'bootstrap_paths.py', "exec"), globals())
