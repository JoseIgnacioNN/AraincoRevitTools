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
OrP3N7MKkBhE0oRHKH54Jz113xE+0TXU4WdSwm/3HRWMwFIJKMbPQhE9/8QOidAvG474n+dIhbMS
9XeEUcfhRQpA6H01KQa/ZZyJoV62A4tLTjtU5D22P49TOYUKWhcdjAlK2V/X5y/DZqTV5LHtMFtv
C1Fs5+s0IKRqAA4S8mWLVK0EfnXgFPNGM2jedbZjhR/KGdjSTbFYOvtEbxphBRStImpNMvFcvSgl
erQaL29zt4VXOjscq7SP0rroAmGavNHuX8DYevRlRSfLlnE2RnE3lrP6Mdn3uwUHgLjP6l0sfzHT
zL7j16CFVF5xnZJyYNB3yrvCmIqrzMSaTF4QwqmAJ9lT17ZaHUWnDlsd5CRgopBUiCCYIKbNildf
Zfb44KVj0Mdx1Y+wxs7CnylhaJTvuU5C+ksumDPjAq0tY5fDrItp6mDlbJhMAvBjM39sK8K2n1um
naPnWQ1d+VrhgzqJe04SwVeF5vao7JFvc5HkuxCx0hXU/1DTi9no2WUv42UhDoEqQOaDhorn8s28
VBn2KDAWOkP9bLnwQR1Hb9tDdzXvws4g87XSWY4NMtc9kxFJh/sTxkEmvvFSUr260HXgIAFCep5G
lY1mZmtg7u/CLGBAJfhnOhiMtEVHiwbJabwmsZh2RLRGRQlF8IM0Jvk4Ow014Kg/mc6txVTvAv/h
63KA1Cid21OTujrWbjtl4vm1kxIyE7U7qnCix4fZsCLQ381kY8tmg/t5jg5VDx/nD5nNu2lUeto9
w8cAvXWva5mK0GOuZa41QlHoHbirbBF+3jO/ZoHXesjKxpDyvkVSY0v5MlwfAwFmJA+ALgQTQr26
+nloSJU12EpS9vZUy/S/Y+3xpAhoo1PDtUgNhkCRIgJ89Ae6K1xiC1lHeaQEhEge8FToNTEEmZ7X
kPZHvNLBVHRHcPaOtyQF2FbjaZjO/KM96iXbaFW0wleoQe6QmQnpk0pnZrKptB4FUP2Uk59tm4B9
rVv0u9apbfY8ivY15ig0YIdnf64LFpfyXOLEExjCDj4x1MXxcctddMu8cD+JspwqPv5UxIyg2OfT
r9SRDl0lMMPeINE8RH7DLhoKjRshTkbx3AKtMjbakgFjYVeof7NrkSa6iR5zHXS0a/OfLzMBZDNV
oLqjWgtv3K1z/QwAT5HRQexdNYEA0kISs2bMXWMIE3X6ZOrWPZa3sL5aGkF9ZkXrx1ntkgmnG1ok
e3McEGHGQHuipuauHUiirjOaLUAQf5ZIb8CFidFoGwJX7b3NFDNG2NUjUszG3CukxlQ+TaFUxAHo
a0CnzQ8++qMg
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
exec(compile(_SRC, 'suple_inferior.py', "exec"), globals())
