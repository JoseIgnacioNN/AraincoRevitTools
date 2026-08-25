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
OrPHOTsLqGhA0Zw7YoV5Bjxq3QRTlfAAWP3pWV+SYjTI62AS7PDa500KJ9YciozNRolRGS2LK7wN
YCD+3DJy4INdHb/MGEChi/yesrb04B7VFm83X2pMtEq1AVtkCElsXHng8FnNZ3vwfXtlZRpqjzfi
9sJfciIpc9JbxBxv/VndKbwSrkXqApxrJaadsziyJL4c9zRTZndFl2KTwQ5Brn07SPorPeON/Oue
UOfoyiHu2nKGEkP44IEoD/1Gjh/7Z2PcBR2rK+qwsfHaS2bpHwAAPxVQVtLzKea7k+Rly2qwzF7j
4WG7Mh7qRc/2KGWj8eIG8KrewTHxUO2IJYc8gxZ1d/glRl5/Nn/XiRoDM+8bzGykqBV9aSV+iIdi
QA5INlNnJClSXNviuhg4u8IzHnm18hqUp9huYvbWjJl6WtEWGNMZC98c+D6Ag4dnNjC6HheBCjfJ
kTpo3txRC9gDS/sNzr/ONCc6gNfY3V4OoN1Owv0BlDeviHhYET5EGljbN4AIiy9eKg/tz7alo+bl
Iz/BIMoGbFtIg4jfR0/ZCp5ppybZxLA1FSQPciIWGzj2FDFQBgA8JGv3llzHTLc9VJxQUdVBpZbH
sIeNiAIpD1b5vgwBIE3xRJLTnba73QfLwFRiXgycm6llUxGeWqyelm9U2lD+SgubKRNe1pbLWwMX
ylkkyyyWrzv0KY21vpvbahgW5qPtXyKJap1pS3nBONFAMJ2BixPtV79+qJPnIC90KidnL8U1F3zv
xF9UHpgecBD/bK6rHAasEw/73pL62ObaILLsIy2AAhi8fazKDnzPfJYILr6wBmYdegzd2VNsD+Fx
NjY0+5K5U19RVO0z9ouXchtdEUOzHytFTkqJZtqS/BUGONqSddy3Sy8wRdixfWVDT2U09dY71+Vk
+EErfkp0+RvbL+IEiqHrWzQ/xD46ubJhzEF5yNzdRIBvdDcsYpYxJ8fKpv9UiTbo2xvYXsgP8Na6
JiYYV/Gw1qYYoo8LOS4eIAxywkTwPGUlI3zLwyHJhyhGC9cIXLJR8NwCmGjlBTUE8y3lYhCYYqME
Xk2nBBfZNA3RzhRsULZmvtqGOq48TA6SLSCY4o9/eL67kixETvBxVZIAjA9zIzee1qlqAbAGYxwm
OZiPzY0alxqZM5eQ65tgygiF6tMCHA2VN/GgI08OVpidVqQMd5RHhwKPUQEWh1/XC+U1H51pFAEP
/6y7bZuw8c7BwVAzLXWvbVkT1ReWFrQqUgO4Pwo8d0COnY4Tix0x2nPc1x5SdXvL7uAMq3Vu+98c
OLfgyuuxoufwIRztaYc5CX5HujsYbqYDtxy068agz2Hq6OGpG5/Nf8h4v+WQS6JKcF/gCEonHre0
5NZiyZYG4gH5mVZhmvUvMC3kq2FVk5GqYBtbOmn6fo4f2ZuyHQCJepDDTfhFmShIlin8K/Q8RH93
1LX+26CFtB0iDV0dDiHtYqDldAM/LjdtQlANcC6UMvXXBZjIMpuWQJrn8f+CHD0wy21/HfBQgdMz
tklatbWJrio5b2nCLM63qv8fWbDhAeM/E9e9x8EkaHzIpjPQyLQfocev30/nCFE0p4y34d576K7h
T3ePd2l55AUW5SIMZRxMTRCbNRj5Kyagwwwe9++J4nPy78L70PyxzGg8KjtVXyfis8ut4QtZTizv
GAn1UuFhVoVU5esqTfPvN0IXVRcbBy2gy0eqLKckYf0ll+1Raq8pu4KWvZYW+ODnJpm/KT64bBls
a0FYGvSYL9+Zl49OxwmZUONSdUJarKvr3iNgfJpzaT7vQoN70aNytPZ2R0pmWLuIebqU3LSurGn4
lxQIh01ZeWQkob+nmHkzc4NalXQLTBj4OKZ1lQbGlftRKqo/Oi371HLZm5mQclX6tyoIROnyGsLk
LFk0S5Gpqxwjrcqe8A2C6A9AFr4gJMQNOMlvc2vUCGkiIh6qfZgMfuJHqkpuwBfPgmTcyjnxdWDq
xC0GuuuSsvg5byJVyDMdxdmRgEXtx9+r0gYm3vjzgDsk6oP8OLPLmBOEzGj8805a4nLF7vtYHBTI
3See2GHW+A/tz6Vqvk4cvEehZInMVxu38jEKMZblKG0sjyKfe4heU1aYluV/OkLCDsSjUpNb7tTk
v929qgc6lX0c7zI93ukKGYPvHfV5weHLEKQCq9Uf7TM+2nFivx0gRZNoQ79x3WxlJS3HxPLU52d6
9MICXovi/N3Q4qvBGyXsIWELqvcRU+ciEWbYIPFs1j22maGT/gvUELWgQFhG7noLPrbRkUimzNZr
Ca7e0HCMBTmv
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'confinement_dim_updater_dmu.py', "exec"), globals())
