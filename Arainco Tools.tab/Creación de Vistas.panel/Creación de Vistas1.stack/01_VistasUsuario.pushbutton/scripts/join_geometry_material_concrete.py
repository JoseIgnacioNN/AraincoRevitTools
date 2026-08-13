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
OrOnOc8WkBhEgZZ4JoPicf1nEwdaljFEqv0Pg91p4KLQ7TMp9fx3NlQbwxe1HJg7HSrwOnCJjHSM
nMdndWe9zJ0GDWQO7z1hGv6gb+QmlUc19Zx9U7BDLrLMTeWWEAvMxcj6wXpPUQZafv4mCu44axjR
iyciAumgOwSPZP6ea/Ia11LgLZt9lkKb03HPQSNPiNRzxvQHlV2yaf/xiY5xuMNEnKOC2KwWMQGh
0CmOPk6UZMpfaSwIE/W9Tc1vOTKTnXr410Yw8Z7OTFqb7kVv6NU7cyCVMYLjqVMOIvCHBB99SWDg
K+FNCvQMqno+j/zk2U/36Z+TrT6/W49w+ZmY/D7aJs6UszP2+HB596mljnU5DoQilsEX/rqP5BQx
9CnBcuDi/T8Ezjxs7X1gna74rBRqnDeuzOsF6FNFr2SkRgIuqkpjwEkYxYL9aLPlQipLQL8y/Ayi
arb+b2iG+TpncUbVXy0DQDp0dXV/7Vu5PSGifzwQtLH+tRpTQXiaHXDtFielhRhSiLZs6uvqatrR
du5Hx2N0Rod6piwAZT+QdRv3r0/wQ2pwUNoNQzLRhJqN56ZJKkYwLUN5/2/7KXbU3lKKq8JSEepy
If5j1mPD8scgMXZBP9OHYB9S685FMjyjUGijtKkaExQRs8vsRi9Q4zahw2L3r3IE4ceM4YiRNjaQ
8nDwZfu3JHm8YKxitQnjwbJVM82gnVPBaUXnRKzOOSIu7jrM06ksGrrsOU2Ov7iXNfQ2Y8D7FkBs
PgZfPhAwyPTl0I5h9Gf5WXKLpbcr7VOffK1iVT6z+OK0Dm1ttpj30lMDrtb+Bl/H4LEit+42q/yh
OM+gMJQzl5QftBWfLs/NcgVe4hoPKK5nSk48eTKdarnqpNc6r2vMwS/AFzW4O7qpsRSpnGDeJbsx
Xa7mq3HAjzRVIlrLp/RdN+wZdtsypzAEaLwsJcSMAXFNXr6/VnVuyQZzEOpqyEqiZzkGzrnLIvau
K0dd6Yy8m8P2XSkWbyPXWHh57l5lzKLDLN1iN6Svt9kNrHIAUVRH2NwxdfCqIQzvdyi5KBlCiN1u
OV4Fa4nBpQGp4eCchWTd4n+c18C/Xn10DiC13XxANR+TfiVNQZRVLrgs/ond+J68I35UM7r2Tw8I
+lu6nTtDxN1r3SJn5gZbyRK7IvaABkHWha2tdTHhlcytuFichk6RCw8LjKpl3VsJ0DbKK1Ku2/f9
KvOJNlV6a3595WRJphSWa1ndDtWI7x5jHtX1WNCEYzxXUz/dEYxuCNHSDuuSzXvilxTe0MnxxrQg
wkuK+6Wfgl41YoyMq96FbxQtWC7NiWjj26dBzzdvINhjYezin/xOHS7iaKDbJ54eVRo0s1vYFHj6
dOmdPYOXavc2Ryj3XF8Pxb6VpUVjpcFBZCHjUfbO9XwiTsjOEvhKoxBjiNbBcN+EIB+QF9Ywhkqx
x2c1BMO+otbdcygsVzKaCpPulPc7G6prQb5W+lrXULOBIbb5vHEfDMBD2holewiVLi6TNh7nUqc3
EWl4faSUqkxwlELpYaZRb276dIEWOweN5pocY+5Fx7hTNBnSalQkgu/+zcYsRrR7CVH4UjzakYt7
KvYcRYNhxLgPwqFBpitZa/B3HfOT62ue1jqfZ9SMIze/ZdaDL4BcB+ylzKPEsEC7aDIc2WrNJWGK
95SBVT3TMBLrEfkp66llyODxx5YoTFFPorZAjcD94pvSCb9uyNmO1AaYbMeR9g8W5b/oLbX39J1s
gP6GlsCAQoK//EGKmQHO4QVRYyAMps8ZtcUp5ikw0HMpDdvX
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'join_geometry_material_concrete.py', "exec"), globals())
