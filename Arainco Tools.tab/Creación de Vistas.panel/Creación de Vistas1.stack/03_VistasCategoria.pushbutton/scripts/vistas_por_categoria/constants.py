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
OrPX9z8WqBhAqJhQK9arp3l32v7ePEbwZDjoRXVQ45VWrudv9EVSKWdLzp6N+ZieKCCeaRKmLjRY
9gb9/dp3kS4bJgYnC48iqO/om1FPhlwQ1IJLrDG9yIpRg7Wzx12IIPHcCPvOGv3+CmKW3Lk924He
GOGQ1zU1u8YCPrHT5ZMLoNWKRocEsX/73EpJOAfHtny7Gk6fbYVIGdv1gfUvT+onHBHNBZ9EBU0l
G3OLp+M26gTa2zCI+TZZxTk35TNUWv/xArUSX0Jt7uAyjZ7uQEzf3OrKhPYoLZdDn19zZyottT6M
3e08DLj1qaanfVUbebGPHORo5KWlPp1aCajndZrmyzvS7eCWEPKpGeT7PkPjDyiVR7bI/KB6kgfJ
C1k/iwi1dSPtgVOZv8uFtMbFsuUPHdL0IN7wwnsA8u3oWvwHakh0Ou0TQrnrlbQNhADnfwaQEz0C
aqvwYkFDvCQ7HYxjCwNM1NI9a+7p0DfBBN+VeXdb7EOTDrAhEEW7ezE2JRPvDZ5dnZUK4oNvQQkN
HaH023Xlds79Luvh9iSN5Hx4AgMw8sx+BfHuSdFAvIVbDyvHwdIwekiWcQHhNqs2nj9EdBH9/DGH
Fh6GKXMZAgKofpGPbevyaUhufGjT5/jENSLFHYj/rdJ1xZu63iDu/+wIi2k76btTqV7vuLE2fk94
WunT6a80/xADOOzjIhZeOm6cCmfSjtJc0CnylKSudnS+KE/ttovmYdLvpwn173XTfBKwAtznNXes
J7sLedIEVa78/ZZo18hBuw/BGs6371eGDeXa0MLmINghOeW6F0N89tj2lTfHttfeV/+4DCin0Z22
niKPCHtipk5Ly4OvRLjX+ke6ik3JqHW9qXSuGWiVsoOyUFg3bqKQXIUZwSnnAocgBdA3F9BokGo+
RBE4uSXKL3jix8m38uB1fAbQTzyVj26UZPjxOq5m4T3xplZZK4i4luTP8sXEQcmyTuhVw84TIvKg
ZuGWMMToIjYU+pMsqLsR3IgysLNuuDQJE8jIR9wANyhHpSSI4ojj3yM4U+llBIAXZx7PBWcIiy+t
zAFJ7oYWZiYlkuDv+DHCkkZMToOeahh8Ts0AHeSPatVJXAQoxMMjzBJHWY9nF8+4qIDNSSYwPHah
aCEZYiiFD8ixjgAmOPHIs71KrmQG90rkeaREEaSusx1/beEy879ucZZuSfnqekpWuYSDIXyjehGe
+OggVkpYVlDlm+Qd31Y3kpzezmH47M4Phbq3gxcNm+NNZQqNY+VVq5Fk6psyRDprkIrf5AdFO2Ih
Ove6WS/nw8lr/CbY3WvvK4CZPu//IUZyLTG4F6OiqOoqHcDv/ww3tDNgODA2WOZ5JizK1Np7SLW5
Mvwbpd/Ob3f19bBxCFqK8Cid3wR2bX9chGm35PuY3MdH7E3XjImbegBeIMPQDTtTSWwlz3lfjVyO
he1al0nLeT9SCZg851isYtOf586loWmcmnGH7MDg
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'constants.py', "exec"), globals())
