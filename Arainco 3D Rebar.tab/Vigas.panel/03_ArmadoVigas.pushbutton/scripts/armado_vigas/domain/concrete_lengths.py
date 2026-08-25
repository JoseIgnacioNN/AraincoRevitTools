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
OrOvN68KqBhA0bg/MuhF/YUGAbVyGMNvKS9GR46INqIcfPmj+vAxumtpJndj066BZVj6uVi2XWWN
Bhf8U2d5GIvj5gkD0mVrUXd5RbBvDLCd5kjSU6nxPzLao1L2lRTay/bzZ/GlJ4XG1FYJOHhxCOw7
K6QGbuv/CGgjofr9Tq9TLrrTAJ3kLP5CQBRgIsyTJ28zmhYCiJNmiFS1I3yCKAfsq3P69ffEOK4k
Dn36V45O4/rQ6F0U4sILio+XxdXWLU9d8p6Uvu/fJpkO3KOztI3nR/YgHrJc03gkHa0yvXsBQfHD
+tZxPeX5BNnYYdzIk6H7AK8rd5sBSu6XI9WE/5NAVRy8XGOQSHRFVCbJMKdwTEL6f9tbO+/hCh47
f6goHSx8D0MjdBAGeqdTtpqux7rKrPBfl+8Kk2ioE76ChnjwqFJZPvaT/83kuZnjfUcDpXP1qMGv
vICohZikuWmnjBHDd4467Q7mjaOuSNIxqnTB8cEF628hSPJGQtunBlEjAfGi3pKAe5Qtwi500B9I
g0BsbQPTVnEBSOfv9KXxl5MJuc67/tPlrncIbU1QpoC67fiNoFOT1ROt1Ukt/j/qr7iB3cYbQw8p
WmWUPCYdLjykqLSBe0InwZGXV4Peg9HM6SN2snm5uYRvgDLD4eQWakICpIn53/tdl+kYhpKYE7ff
eOn8GZlNPRj6gRD7nv/vYLaeRDUvASalfgaTvlwFu5cEGHChdefnnmOE91gaUOWyTgNxnrwcioJC
nAS19NqIQcNj0EQ+NI3Drrbg3sq0S3QX8/y74qDiotzcScaeonDNQM3v0ZopiqExqjKyRkJzLlzu
rxb9WPX+RGTMBbhVuO9/YIc4ZxnUNpYmHITh2f6Z4v2iFPC0Yx+bilbbJr2KHasBgOd0S6SoAmpK
w84Ht4axCrlGPJRgw8PBLxT0w5kgzh5XEe61pYLFcMUgCpq3V6yVnBDftNSI5E1wDETlKtoc/Si/
57PTq0kKP0yaAnXhPTp5LNeBEBzj7xJ3P4Q3eeapkM1MI0mdmElT+bZEOMQjwtzXfQodQ1jpQlGs
6+wzUxTZqt3a08ucGt5YtbEkRR8JaTfEwDX5/QPhP+EXkm2rT0ilOJYyM1/j03RFTicxVexuEd+d
0h1J32iMEKGyt9/GXlb/AtV/n9SJ1v2whn8+el1QD/1RuRmGCz1/2ot6g9v/zYAuoPuPgFN/OlwY
ys7QPRmGRynWdzVGMhCqeoWsbylLLHiRE0FBK7Ao/+gA0zOWSHKiBOPFicw=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'concrete_lengths.py', "exec"), globals())
