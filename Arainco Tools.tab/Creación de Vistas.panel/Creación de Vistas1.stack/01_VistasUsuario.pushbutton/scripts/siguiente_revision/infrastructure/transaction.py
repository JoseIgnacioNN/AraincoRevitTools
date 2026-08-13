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
OrPHNDkKqB5EsoR4LTcReOr4yuNmACF0jNYA/8pGATNv5ChHZs3jej5wyGvm+jWfoRTrXMBRpCvh
1jM3bMmB79mIgmU8u+9s/l1BgzgAq5OQg3TrQOY6F8rtXUPbcj/J5DAjw/T2GiiDl6tBAv9MqjVa
6jfPATldbV4g9srr+7TAIAwmjOtmLA8E81dlngeuBCXVgDcuwTc3YdIw+InacUjlFjygwTAtTbnK
ik4oI2X6vdCuctUDmpAqUGC6dB9DBGD4ILgulWBZhUgiRxyexiCXGWdCPC26Kxf7b4t9rHPr0M4V
WFm4pr4AeH3+Azshxk7L+XI+ePTlG4xlm4IGN4SDXNiTE2ZZcH0oo37q6kd/5etXC8yojKiRm5xE
pNRlcgoQF/OoIJvqqLuixHp81x9xQ1bDt+4t2qvbBmk7dZzGMcmVU3QlL1VP+nChCiQF/uW1t86w
0Z0VuyYBEaKHVWj0e1ISU/jgh1JIzEZKKrzDNE8ZygO8Xs6mGBugOdwZ3i7yK/IF8UWGItmEppx5
b6mP1U5TlkE2aow7rM7mv3DiQaosJIiVmNwAvyKxyNofsiejXVS5GlvZkF8myCgzD0U2MjOpq12n
twPwubez9qOwU7V/+s57UZnjuzKt1ThHAUdUDCq5/qA6pWiJ0jxfJtwigiGKgPu8/gNq2vnIJrwS
0HSWTuRJNRGqWPpHa690UPza8/vIdBpyfC9SIsLA9bjiLg/8snI3yoC+Hc11mtjPB7njb/buwtrg
KPXYtD0S11TqCyJLjebYyOm1unY3CVoC+8wCkx+ATUkQwYek6+JXYm5beJqs77zzhbc6UBr5sOTw
yrGIq/Idiz72jc6sBFN0brDt2libD1kZt2NORWb55q43f2Ru7w9MQvGJhWNJ44sZb3wRyLLT3wZn
c+cJdFOLl1/BqVmrSq4Fe0aoiZgzksWRxd+Sr2N1cnWNU49VKuW4y8XmpTBZUrpDldE/1NdeuHKL
Z/sG/nM=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'transaction.py', "exec"), globals())
