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
OrPXNSMLoB5EEaDDFpY5Nd6iCDZOp2uwdfC5a4JH7Ct112wO+0urYA0RuxLrrmoIQ0dU2Nc843Sm
q/TQ1dUK1tUZQl/2+Nf8/0a7AleotEy1744jkmwANHp173nbW3QVgBEGwOflZnzdobOzp6FqjSuK
uzvtQWQtdEJQLIlv4pZH/9L8b0EiujGGZnkrw1i9ixQESLkSams5F+d5Wq8E+WN+cFQVwVLlwguP
erqPZIOuViUq+UKdZIjX+hnjFsObu+MRhKzmNob5nygjhob8oNjMsHfcDJnTrcbKlFUDI78OMWh5
xvFbCnHOiwRH+PaMvr8LGfVxtYeCczkKK+i3F6tglnRvIe/rRolWzechnAfqg7g+rwzSz+r5kVr0
aLauGGsVISucMTQPoxBqdyB/+DMv7rin2awT+dn2uIp0P55ImhaQ/F/rks5Bn+/+YhJAAvVF8DLa
NsKFcnHdGt9lYBfBH63VI7Kg41Y5Ysy/P+rD6N/7psK10c+pj7dXjWCEKeTZTBtB38M/CI3RtEmp
wDddsihhkOET+PdNIa3suwV+mUCLVzbebi7q3rJnlTleb+Qjwieh3mymyT04de2VQF2NYpFWGXDo
PluflhW5UZaasQQcjvcEzcIZAfCZW27X3+/M0NYYcTOEBoRr5v/b25qDp+4C1KUpk1Zj/yz8sv5y
fzongKu8NI0WMg+fFhqISb7T+Xxi8WQWm0R26MiHX5oCekfbAE83cKrl3WVvAHPYWTEIbIjTjcmb
kqUlDLr7uJ/wVrVuIgDw961fsJMTZvdR6ZJjwl9YDMq+4AMLC8OzgeMcj+Nz7aPEANPPwfcuFZ2U
LCnJ+bIO0SmdgQ==
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
# type(u"") = unicode en Py2/IronPython, str en Py3. No usar isinstance(..., str):
# en IronPython zlib devuelve str (bytes UTF-8) y compile() lo leeria como cp1252.
if not isinstance(_SRC, type(u"")):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'bimtools_joined_geometry.py', "exec"), globals())
