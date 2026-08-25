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
OrPnNr8KrxhE0YA/PrwTyjI402Zud2Ra2YLPGv4WN+cr40DnsnHKKEfmeyTk4BGXvimnx2mnNKv6
gCBVStuxbCte8m+Jp4qQO+mLHBaqc3KYVs8QL8zLXSPBMHBSnVUCzpjoSDxujauwk4TTLAErHaUL
tY3JyqKohO8ae0eYELvL9Wk1uJgAc/WwQTKRqmvSW23a8NK7TtK3cBQXlQw4cDFfIUorRtFkFgoZ
r4HZPC78Xt6ENB0w57D8/vo6KOst/VQkB9XvZOm7MDJIRuHOFxPYhPkKjIWbhZiVuLi+6x2MHiJ2
DTg/lXincj+zb1QbyXb5AMK5zwb7HX1GC+miwaMRAFfpNcVoCqunQAmy9Ifo3JMmEFkyPzECcnCj
PkrqXHRSFzQ9aNpIstSz2JHujUdKootTRuKeelfAP+NRSYl9FbqIKCbTCKZh/fJJvlh5VPMPoRsH
i5b+0+4XbgHuMbbT6qX1iEpYgD5hv+3xCmI8AMmqfMLgf4OwYMbm8HQGgSdDuZ0JnZz6SYe6zETM
JJNi40US2JbYFTMqgsxSalEFtA3/cUiWikj8Aaszs6O8KeTX5RAJ96HVLbX1PS4agERPNV2uQl0F
/HiNFqyLKHIZ6da/tkgSAVPWp4lasfsBNe1OVZ3ounjNYoZTi3CI2sDPrmhquFJA1OURTEAHoG0u
W+Gw9BK5eaFI5maQAh7rGoSHuOpcC+BynaLoDXgm4wlvv8pyunh8PPmvTM+a7KOvBbcBXobMWIO3
qdzAhfRwnEV+OGeopg195K7+QFKQ/phbz8tCg7Xy1WD6I3/89BKhi3uanOR0rizj0lhqyc99ZQQO
dcwTghtNgHKtzxXii/ezK2zOFQEP8yW77Dh8FkFvGfBCuc+yRQAFIe8nQrTckmHivQCZ/fOUalfa
BtHHMuwpuGsRawFUBlKDfcit1PZLE6g93ZJiNHo+WN/PTas+hjCHf3anG2zxTqwJ+9kXSMm01aql
mSEnSUzB/b5W7ySloM0xaviyxrEjybMAsYsKY62NaIM0+2FcvhJ9/2BQVXE/knV3eFDBDDDj9jQn
fVjYRbfVlDarGL/86QNDROaezDchEKYAFsO2wwE+r5sfYA9WnUxES1yBNoDBTlmAcL7UKzxSPWu0
eF0cFQ35vEpwafL3bQT8WkzM9OTigdZLMK0V5j3n+k2EAPXhXsSjWFgTMJsA+0upb2VHbieAvVe7
qodvuAK4AxO3iPBehIApxP45SqRel+5Cwq5GQTre9ZmQtp0VK968Qd3Iuun7x3GHCIg45mRyL8RV
rJqgLYBAONgcJpVQFHO8V7pGpmPfgOeOvdYyGBIPIkCq8kylyt+dRZ6U/jNzXr92zFO1+uU8Eyg7
wv2bzT7xagmcZmOUpF2/cIkkvS8gOfbhB7Yn9ZIi3xqEvR277W5/R8ZaSX+VmAnoEqydfqrGP/OA
qAPMz50pNYGDm2ERsY+OBjddBqJi9XinEeQb37VLkdn+VKQ1WzHDw4ecwX7ctkkunMQxYJCCPsCC
HeJayktSAVupmFAvn1YA+cXhotTZjEJCrTWJeaF/xx+Ms2RlBsyaWWeZebUQYnTOtFTKFgHDRWTi
I/yIbbQwnu6SgkcaqHvhmJWu98yGS5t7hXeJy0t4/bEIHuh6KIhJwsGTPUzq6aHBvPujcfZoVKIL
E5XpLk1pyoifZ6ZF1K01pL7w30yIRgHDZDb8M+AqKeg0eSulRhI26BihQXoP3SgL+gWed1jTklo=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'canvas_bridge.py', "exec"), globals())
