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
OrP3Nb8KqGhE0ZxFKLw3JQE4Ikzjdmy3qX11y8e/rtnnaUc7/IDGuW+jkaRYHw/n2m7BX5KhNLP6
gCBVJXxkKPnpA/Fk38CdixOwW91dbtJcLIr30J4Q6wABnQy054QNHYGXdrEPhd6/lnTEJlywWl2p
+kR9ZY2egjuZuAoi5j6ZRswjdVvu0yiuAxAPp8EDgZbJx8g2qUnUgagqexqtg5BB0yClacgODZfB
wOcrLjwbGqRUn8MnC3TtF1Sl5LPovpRVDpNgERAjx8KOTejGtaW098SGwwxALdqwuhSTi/aMCdkc
ss3T5341LhVTN9m9tWjIK7KtR7CCeI1kgI8tle2sp6fmE76E6LzZfgJdCfKkwR8VZ8AzC/ISkoQU
KxtbkndWZoqk5SCdvIhEajfNobQtJUGPKyJwgcCU/CkhgSq2V2FFO72IzA6lU3XEGGx/ipWTRU/N
3FtI4JegP+pCk+9B1swk/nJrhnQ9XabyFg+tb9dOvJ12DkPtZK67+/O6Nhx/P9X8mw1u1rUozUGU
ZlGUGHCTc/6Mok5M3nYT6TtVWGWKQMiCZ3sjmfjazuqCU0wxC2+rwtWRwhqqbaYRLEs90ARXwuVc
o0EmSv8bMTHGvII5qXV5FnpYmw6wwX2pbXQ41clMun0nKukN/Mi2imhwYQS0J12zbVzBDTTSEMID
52VriiDzONmRckH6gcPwpaSSbzAcpgviOtlT7vx/HWj0SkvDOv6vyDtgr1KEu3P5wvAEzfFHmYu9
/mrSueyqdMbKOZtuYCfKaxB8kVh8XlFWebrCSHrZIItj42rD5rxclGbIsACSejwUxc7KtSbXWPz/
z3T9Qu8SBW6Csr+mUSXQtAb/BzawX1eaBkHGc5QGrg+VhzFp3KpEUYTz/+mOrEd1YnyaSWef7skB
TeTNsoKaoMT88CQXRP6ns/dccLLcI2C6Up67ra3OxooHDJDrUX8EDaJiqNyPAQmXICh9G7YAz3CD
zz32PW5w6R5k5yPWJbvV7aOtppJODVvWsG6V0biaGKZNnpJOu1tW9oWlkOshQAj60yASdZOK55GL
iniYUpCmYZbKjaSem/LpcD+GC6XAi5uXgQ9WwWZrcSiAo+luGxmypwO4o7rgS9azWZ/JNvNX2tmJ
WJpfajZRpzCi32cX/xUpj9FUSvMJkQ==
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'join_geometry_instruction_dialog.py', "exec"), globals())
