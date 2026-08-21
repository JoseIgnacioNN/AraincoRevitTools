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
OrP3Nq8K2GhE0YhFeJ5/O6DbQJzPICy77wkny5Zpx5JRwNCxcG/fLHgyj/CKBddToFfOvJLLmUCF
OZtL83NGN4UgZzhD+orIz7a2tbOxcDKaXsYOgKW7eMBYHEDpbNQFWORnGlI18ZCGS6xtZ8C9LYK/
xzJhTpuWmE9j9pO+s7mHXx/wm/wmbkdne6nLfDa0wwpqA5V+/n+XWShb4jc54MAsk3YwWUxjINuc
xdNA40tdS57+Z+HAw1XpW4F+ICDVc/R8cbYt2AslR7H3Qgky8n0lV/2UWYIWRevib18hx/KOxq6i
O1fQVc+bRgowl8OQ4F5tHU//kXN27f+PXlOEmbjmQKMziJBf87i8KiTSTXDpP2dyhV0YP5X56F7a
NCwCC1D3NN99lyH0BQd+YQ1m/yzbepZVn5R3Gu7FPvZlX9ZiyyvLiNvHaMkwGrjAt37xMwR1QcXX
w88zFeWVeyuwkUqY2bugNTsNQc24SX5+NLwWOmgLoNaQh8iG2Nbu+juJJcBiimL+fX+8j38gOzUF
VD2Ago8cxEBcLhq/j8X1IKVJnYCca1obtSV4tDx8qR0KnufTDlAHznTMNdK6cBjyXXEwtw54TR3Z
b0JG0XA9nBkRws5bSfhwkpuCAxH1rU1c4bbhVHj9ypKPosBmozNq0jqH5loO5h20Ud5nGemMxnRY
Iy3qPYKUxYjlYzHt1cIv0JJhHpqhwOXXWQxdcam8mdIGQmxD4wGdOX3KJGMYMgP/kLN5x2riqA3S
hNIjSMXI6GHIsgPe93ot2Uuofxofyz72vio07x6EKgfXsyJA4MKev89NxWd4Ot6CsoqIQJGutFw+
TFhNmXBTxG/WGlApOR+yk3s8ZS6BQrjEMBAg7sFvJRkux+w9zV99h1AKcxO8ZuJ5s2pFxlr4zCEz
/VtHZgcyy0fGfSrkS15ITS1Vyznzz/7q09DOgf8DedUael8KoCuvrYQJlKRasnhmus21bF7gsHTC
S6fMdr8FSBbevczk1HArJtZLKK7X+RoVvLKj5YtmdwtmTc5krENJ4Ik+lk3Vx42KgCath50lpF1d
cGI8HTQTyvqa/WdHizVR31PLqldbElHJtHLRh0EBuGl6w4QhP1pclljPrSptd3m32EoqPTIK6TZX
KynJEbWD8H8WmheKWblGlFITKNKNc++Llg/7PPH2JfbW+oaHErVUqRerccbG8edqErFRI7epBjQ7
zzZklKf5ZHvDpTjjcx9rvMXsafkeAqEsMX4axSPjcf0guXmEaKmnKqEmff5iqGAIRKA0Tbv/oBw1
H3Y6+fmnP5ZleTnmiKz938FFqEeeWK6XH3fYuzMmbexurMV0/CVEw73yxaUZAQlUaz2q4EkSrrp5
RXFhwkpMnrHLavZyc2N0DlaNnbdSsjL1aIWey1ngkkmrhcXoz9KXj5V6fk2/nDIvXH/SyaSE3sFb
iquYLa9rHmQjTtDRiBM9OpT6AQtOJCFnUELo/egANsRnIUi7sfMjPYMNcSrlGcpP44z667vS0+Iv
3UALaJ4ExjyQWq6FTe+LeJF1kpmMzjYe9jSbys4sWLf5H8S6q4n5KPXqX6uJRlLg8bdrCW0MRSr/
IhBA9rQVL0dncNDN9xUJ6E9PuqFbiWFgS3h1rtS3eB/XRfiX84s+yiJIEBIrrItBOvab/jQ6/3CD
c8ywsjLFCCeUvCROHckQ5Hr+k2t5+0Ye
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'theme.py', "exec"), globals())
