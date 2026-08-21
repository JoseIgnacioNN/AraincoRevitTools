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
OrPfNj/qkGpAsYh4YTYRLL40a3TO0F8vFhRLYqcGnwUAc9MzxE86P/OlxUaJa1JLMjJPE3Vjn4JH
B3sx+8Ou/w0mHKm0ktjZq9LwpkkggZZpCn2HBAB45jiqFQJHE7nCxKtS+yCshH3ijGmy0aZVqzY0
woZ6B9pv7EdvFM4XJZcSfShngTX7jOmxPaUDQMBDlK5eYBoedkM2KHHL5OoGGDjU8E0HmdXZ7n3J
BmUDOxkYazoa1j4BCbfOP6jqmJws72bgeOBzCTNDqSctfP4I0jrk2fI3QcuVajqa02Fz4PAaY2y9
hGQMUi0bN/7DvybYu9NDfg2u0tBt08h6UlKUE9cmDzEISexmz+r7i2AOQyPMlocRWLPTqWIlJaQC
6f7SsPz/9N5Wav0YOEbczokVYl45NPR+f/hmQFFWNkKeKIdQIIkw5oNqYfxqfKXQMmNO1BV5fbu+
17tTFX/kUB7ltQOBrierzmN810XrDGQTPFpxfkcqH/cWA77qeQpCBsIQki2SPbcDtBVQxYxQXl1k
qGKBZS3wnRsljO8qezGNF1IWXn0cguqO2dYpw8mZcyRNm7L/vqalTlN8+XGpxmELQTOo2srb6MXD
M6bpjSc5/C0Il5zqO5R1CYTuYgLGZq5AQEV0oPXYLx+3FIGkJa793Am+hWijz1Oavy1SNPB7/RVb
38yctPNDZIdGNF7rPErxT/Wnjw5LgTl8Y3wpyGIZczV1tCbHwnSEY5jzClyFyb/74MIy1fnK+jR/
kyIMV8QQ2132EbN3p/s7yekB/dNa3i15yhmxeKDtB0jMytZvh3fijXNfkvylA+MJrLVCZGv4pfqn
376CFtimvE8apur4Gmsj/sIbUGnxGSdbkRDq7f+j/ziz6smO5b37zzoAPwU6N3dNIQXuz4OOoat8
qcHSNIL5kFawg5qAx5RwgecOnvlszHi9b7wcevj3kLl/YyRBsOQtyP+m7aRgeG+Ph2qH7SqGLovQ
B1tUwlY54EAqbcLGRIhTtPI5KhRSXOh/+xiR4Qg82BQpPNXRWPnUvyC4yHZKI3qoJ5tvAiisddlf
Q/QB7dFL+jBVRoe2ZuNJDZHqgRS+wA/gA9XX36/kvh/n+RnhJjzwzziBW6HQWfMTKAbybuJPj/OJ
7YKcYDT9jWoy/wKNQqRCEuH8a8BAySYSkJYo4kGa5XlcWUuswbOQB6Stw2HQJ5WNziMVien/ANhm
CuFpKtfQNWcu8ueAPsXuge6O76k/e8VU0RVBY7ECEzdzaZDLjeK2MvS6SgghM/1DsUCf4+z4XKia
Q8YfRz8Q8ILKvgxi3+KVo5wLJbdItzWhm4q4gdEOc55Oz3G335twwzon5HET6lMsYbcOr2vrwT2W
aNadoEjfn8Qw0T53lvl+MWJkMZGDsXO1uDq7j6oA1wid/x5diA2Fkw4uSP9VULwzf94dNU+rhQVR
En3TPsG3JE5+rdmdWTAz8quiPrvm9aj+CPHf4c1SpyTWTnXcB7QjQxm7gDqTyMD7+h1Xt8gviuGq
6h8w1dnzgOWvOfWS5nTUM2yYgBko5H+O80PENzU3ba0YF4OJnJsWE2cQlLgquOXE3eeTrriukYVw
WtzWlZtgOCGK8BMriW6stnF/
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'run.py', "exec"), globals())
