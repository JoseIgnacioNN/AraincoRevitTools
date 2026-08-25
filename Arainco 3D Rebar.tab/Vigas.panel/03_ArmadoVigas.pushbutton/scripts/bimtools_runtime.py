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
OrPPNj8KqGhAsZx4gZr05QHk//HEBdNhK+5Tbgz5Pr0s4/h74yu5Dpd2GQrGCC8asS37Ra7H+E3r
aihYp2eM79L7MirAj3H70r+fkVCjz3BarIyl4eXqP4uNQDsQWI4k3S3HP12edRMy9jOSg3E13/HY
/SJOYhySOw5xLZOiKWrhRm2fz7UDIoNhMVhRHnQfGFqzhpVazLCCvxd9K0EGH/BsvtRQ2/ctPMSY
I8+C6NYhydFpXfVYSNVn8PgosS9i10nOrTcKb05wUjU9aWfr88XRgYQORMx3N2B91xu9fwnraNdU
FvQg7s/9Iv8boGS+NWao5NJBau9I6M4s4FrG+OJytVuHkxEEwZtNe+GwmAN3+m0Q/30T9qNaVj52
uDj2LUD0GPeYUQoncusPScRzekQPP3qWHRUuYB1TxBJLR3wDf9/QbC5Ym/M8cvym8pS3SXJbG91N
/uMfNR5FZUH6Q2tDVvHmAM0cNlabmLTQU3SNLMrvllOwsaK/mc+99Ca/8XS4EBzWDQOUUOR4jNbR
R+52EdhSXnALRh0q+9a/0PXsBf/lc8RxZDCW6MlFoAZyangPIdq3r9doBbsgdjlAgozan/z+oXds
5uU8hrepk4ejEDNf5OZT31q3/+YBqwSJCyvojkjrkJMtn9q9S+YL9Lhu80Ah0+X8/bjro+jGDSPT
2ocoDlXxzdmLOaCZXtmxQo5Re2OSHrgPi4gXugqW+g3wJkhDxLYsngjTkOVkSwlgLdorTyEdy3/L
xvIwfl9jW2dNRr1P2tydC+Jqdddjngwl1YRrtsDk7BRFc7O5IfZDqGefYVuLfOLdYDn/WOBoBioT
euuOlllw/9tLIYrhvYk7sy32Di6TyT7brlWbBXwzk1kg23TfnGYIs2h5e0d0mYIN+KKVyCAaE0uT
XaSeScJGSaOHVZezjvyGqMVF+gR5xnJ5H1jM8wJOYRSH+9ztozUA5P7417cuWP8lCmZ+PQDzqoo0
okNXp27W1q0iOUw9R2vsloACHFryMxo6EWXDaKLC3lchKgaAHn+G7Th2IMynqO/0Acr/jsJ3Z7u+
xGecQLhVqYKrOJZi+IIPIFvz70A8DbvFmZrw8mMOGYfGNM/+FPz0gzjDWSJV1Dh6/sSSNlEWwUW+
JrcMOGabfVPgwMg7Z/DPrlpul4gzUfMpYJdrW/CA+PLBf9xdsYbQ2P/EVWuoEl/hqBcY2LezrZyB
Yp+Sy3ihbzmZRkx3+SiR+2NIRkSZas48Xzy+DPvJ2Nck955KMq1pjKiQdJVZbxjdXDp5JRzq18Ea
hJ69d8l8NQdMknUxZp2Nehqs6LoIhKXuokas/jZB2p699UXA2IK9LnlHZTUvZTORinqbTz0lzRvB
XC2niBhP1FKMbLFO8kYM3e2VXGKMTrsGODA59xpna2rxOz5xlRUsVPY8ZXguvT7qR7Maxvtx2ZeJ
a6AFqIbEGVT1YMoS3/+a3bqzmL8xHkUwP9LjzGenrrr9rHZyYk93GxAPGQoSVE1RAR8xpsCzaa1U
K3cSUjqKQO8cndR6vZbvpYUpS3A1rbZPVhd3JZKFbbaLxd0MKE0j1ed9XuCQJDG6XgrY28ufNwMU
Fukjggi2VUUGLNaSZLbWU2gNh1BPwb4Oc9ddql8HBLhGvm8iQ099wwpcymOcTwdh8KQpR/J3tlqS
COq7ByVjDO+JT112W0Yjmz+esLYH9MG44iqkhPC8Eq/+XXPRIiuTY/a0WNZIE8IhVw==
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'bimtools_runtime.py', "exec"), globals())
