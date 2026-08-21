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
OrPHN78qqG5E0ZRFaLbg20XscGbyX+TBQYFrUr5tE2iUryTCdx9ug7k5Z01ftKxRIYvxJp1TTyoG
oycD6fZSNagiDNzBl47iLlwU3NE+bZNATrMQagxljX6/XvXAxaXl5fVy99Z0IyK1ToxISMQVZiiP
KJ6fcE368kpVsZGqDxNwZfFBPJiYUmqrkvY42lZpRkxdRrhCBSqzFfmr6t8XdSzBMuC3DijmEPUW
hrNnUa2XvdMF/MN0sH8P1ZAf+/E4kl1pDeRk0SVeEBCfrqlmXKQjEx/mOeN6YJOzqqZMOshgYdTm
6o7na+x/S/7GKrKx6LxXp1pJ8wBvlLFyBAmRVuX2tN39O869p0AIe7asJGwY4n5aDFXsOFxWWnWH
xujP22NsMnbuhy+8nAVaqFvjPJ2gb1QoaGREeNeslwCOpIiH5o8nm7+g76/q6uqVBRgM4ZSyXpN/
uL9vvwW8X8Nr0XNYgfkLvkBePJrwSBB1I7WswlNuprGH7VB0ooLG6Beq5musyxKJfk/5fbkadfvG
c62JVNGPEgCAe1pCbgOS7FAXYNDk1C/QowhdHhSW3TmXRTChijT0qkuzuk9GNOHvxJbCbfEsGuX7
jBkwJ29LAUrPg50lu+DFSGDXZEQDufJ3bcQOYsFWufSFx1XD2xYrNLKWPmlwksLf15ibpIbO3d+E
X8AnC/+qcyESCto1a89UF4W6WnhC/YDXQ+t7HNrsBkDDQhupfxJvLpKKXFyv5XcnUFeJzAfUCjs0
KtfKtyFFz7fI8l7+359eircjWyQQRSL+Y3V6I2nSQNKHnlBOXYbw64UxV16xKF8X87uv+8wZtm/9
okGqQiuvYc/hA8b/u5R8gBP5F5MUMYoz2R/E2a8cO308PLAxK/alWfYmugfypjsMO4XE/Xs+y1Oh
K6VW0jYqwoukylEAU7vNI8kozpWUx8exkoPRY+g1+Edgvd+D2stRtsELC+IvhaHG/wooIrYNwZd9
AxFz6Q==
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'confinement.py', "exec"), globals())
