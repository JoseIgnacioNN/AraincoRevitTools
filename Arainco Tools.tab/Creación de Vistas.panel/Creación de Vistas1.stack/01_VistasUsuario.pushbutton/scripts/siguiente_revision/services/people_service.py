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
OrPnN78KqGhE0ZxFiLr3QwCwRGNczcI5yYQmD0+JVhGUjbK/9oDBlR8qwX2O904vtTmHYYKhNJP6
kCCVg9wxlWa++2ixcARZ6vby9UE2DMv3sQFKM+xfDAYgspqn4RgU2iVcoBkyJuB/uXT1HC2JhHtk
ytxzATzmySw4qTqfQJtWb8WsklcNoQNza4vilEtm+fCG73G914W1XkSojNlGkHY/rZZObz3GCHVE
sK8CZ1j+IxogbV9nLWk5NVW+Oi27gI/+mvDJJb927oDhOvkOk9MKbOMnApHd5P4LrzGmfNmQPcPZ
jZKKQSeoA/+jmBkBcO/Zbw1exrCSjLwdL6gFfxFgGVIarLrcQqWR7OoW4s94UcA9bCz5u9paMQQ7
fuSdKTIRL2D2NovgyiWn/DLHI0mezcI784xaVKn+HDDZLahi2YAfZHTtlMCBC6z1xFkk84yv57A/
/BUIOhk8Ba/nhveNqBOQGBGOZ5nheotl7yLFuRHMMBplyIUYEFKb8dkmlcLRqqExgcWmE/4wsA5e
KY/nurcfBwIW10xWUtWUSO0YN4TXyjhdgPv1wt/Lfx0R2+ylkLOHhx/Q7m8+fqIi8S4MmCuq34uT
xgpX3MnYX/9qHBExNF8u6z2K4GTEEjI94a9jZwqK4LryFV3lTKg2OQXKeIe8mLzseqih1Ckd9GZc
Vm9dA8QEZz8PwSi45Nv55BjuZ+PlHtPXwb3K5VXJw55zS2nGrAt9J9EOZhXAbrP8WJb5PuBHimqY
waYwylrhayrc4/BqiarSrdzyexoVg3b0u+ChPEaebyLkmmLj7FB33I6xqUQdFsu+7xmxeCFKVlkH
Jv5yIW7Uwhhmz7LBiUUuSmpEi2r1ihaKIRowBqLk4NaXxmE+FO8BRy5r9P7Fs0KGX/jtd9sZjbPI
vx8dsJ3rslkEkpIw9iPSnyyTU2gQEeZQtbdZFJzR7mWpAZ3bXntSGcIsuP+iZQ4xbPUhLLUV7pfc
ySGf/tdzr2E50EicyZSlbx+9R0PMHdL92Qgi6oAJPq8GNKUkst/dgzZ5kCarNdttJZMfDD9QG6SI
SIdt0MEAhnynf+ICqjSN3CZB/iA4ZR7dxPmbGOapOWm7hPDFA0nKM8mETBZ8It5D7BtxBe19gPYH
GWdkLR57gTPYK8ow5ZA8elpb51z8RziGUKN1ACOifInIdtR9R2hbjR8rZkUCBysLru50L7STbFq/
n11Qus8YQvJdazBgYq3iCsKBwtlPw2FpNv4KTWpWYtTwpgrnYrrJGRoEllE98MA6rTXcFWVd1+BE
tc+zFJWBM3AZ1kLlR6s+hk1bc45qt0alHanqfO7AgeDOcAECpx8Fjuj0aJWVDeCjIE1eeN9lYhMf
ql7o5V59hQENHqOW0Yop+Y6yoiNYVCECenvr7dwur4NDkQ+LA6vJ4hKQ/B7wX3NyCKiSQM5caBYL
B0NA0itVnxwWrKk/KWoXBQBe+a+by9PtMNG9/THMi3UntTfMcu/TuDK8Zz416AfIxieGKIudowdd
QkvBgj9Yp5WAC+w3pJVd4FwhvfDNh4D55kHk23ILEYGn14TD2ttDiiPg+7JU5EE8ifFrIPird73h
ThwJfkoBFsIskZWv6hkQjhU652tUFgbGkdZbv6eEOxFVjW7V77BG
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'people_service.py', "exec"), globals())
