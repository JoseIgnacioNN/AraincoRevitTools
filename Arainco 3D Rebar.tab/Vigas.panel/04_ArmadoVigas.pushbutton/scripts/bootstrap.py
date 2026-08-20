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
OrPnNr/qqBhE0YRFJqdQ2jpBlGIjxJ51KaUsQAM80nxsqPjHttW+gBDFSwnpqzxSkT+9AHsn153c
Ly1oNcyKQaTyRZN5FobLQUZ7e1wvMXXflbW8y1HYtosTzMPftKOVFLPI3XaWkRWuPTjghacQoKgj
nu8DEz6JJ5vob398YnNTZzLxJDDDb2Em3b+EcuaBfvUnwvL9uUpFwX+9jIRanRJbM6PAVoLjyCL9
vPEPjZ2KmfRdjurcivoDBhDTDcEVD2FEyk5+y3TjZNr8mgR3RM0H+8rL/Pft8qKlcdTLniZvsCNU
REqcbwEP2gw6gGJ7zyoHyU4Jx3lLEY8T6wTzcGWBq4BAw9C9yHrhzU7KI04TjsqwFTNUvJPsz5lS
uvVUJ2Bxp9MjUXmSZCpeMgmr68k0MgbSsztD7r7kCit9gOdxfrInxR9p8uCBQru+Y4/zsQ44YNKz
H8OHtB/5cLZPafSGOnbnMdka4LoZPMTR5E+TrU8uPisDV6f8NfussBEN5RM9WKFrpzD07XYVOgl/
KAWTiysu2z3XNzLRW+UiU9fAKiz1xnCBD9QPZTjzMyuKpAfcpyccxeT3BaXtOEXi+tePGlGFaNBV
e18h8T+2+DnpeEGTemtCNHBGWg7mR+IUC6sb5ncCfVc3QunuJc0UWPr1pRGdl21uOSl7xAnN9uli
8h501vPNdzqzEOZLecJvDLgqPHMJX8PXkHydsEfbxZBvAdG37dT/7D6qEYAtARRa+VM4Dd0yEf5k
wJRnyP4s02Ohb9x01tMvszJ0Saby2LvvMBj/LkFWt3+lMlMzprQULzK2t9gnYSMgM7GW3gDr5YSv
kB3ZkMquLMhJd2d9WAKyjopHDWZqXA0MBxV5p8Gbi34miF6LOoacv1vVDCx5JvvBXvWx0QUf6Ytn
vqhjXw2ZPmxuNlT3VC0dT4f610AxwF+eVL2BESR2qYzgrxO4PCYOsp6UTQPft1eMSCo+Dkrvn4Hu
E1yZ+/vrDDZetIISnZY+EVekIoUOsWzYdeRTJahUwrg2Yq0VRztBdgsG9vuBAX+QD/BaHD+LMUDt
YdL4KngUs7A/wHWSibcvAnfpP14pA/aeDEpApz+S9IKUbY/v23kltSsePjMgFnXmbs8NqzqQv1Vx
eTqeqa22ZCwSBf8/v6lko1abM2wXxs39f9JsjELdzfZTCkTDU6x8zfIa1QuSI1hBnImRt8/RA2W1
idhI7keswrFTdkwBm1l1P975f/6M8V0mrEhluiuHtXVpdSfWr4ZSkzM7qHDy2gpy7fayiaoM48ra
G/7ocm4YbLJngd3yYxk/QNKgnj5/lTV0lIuTWaDeqV5QWpfan1HxQ+UOTBHFa+35eN8dR5FuryVT
5/qvfu8ceInvy89tTO7a3x+SkicAhjhCeg7btTK/NNjdZ6T/OJzl8tXWy+4cjmz0Ty1lllPxgqWu
eJReSksNZrAcGfniGOd6c730iEb0Av4TljSzLY3ONfyOT+16T5Ud4tqfJVw4CY/Q/Kati2ZLBixP
85c+avxqbsCFDHjQm/zZVh1gFRIAa15bJdOkXutVE2xpbU1bwI/nhFww1aih5Qbhwb9/wU9Xnxar
Sy4oAyiG4vBMLY8cl7tIw7x770XdT0bzVP42YsQmB7CpzotrRPN9V0ieMYYi4WCIAwBsWMf3iMoq
ms406aqTOKawMoIomn7J1fbJMft8M/+dWRWdrMoYkPFndR9Xw4zO3U+r+2UJvCnb+dIDplc5A4QP
2/hGYf4THHzs3uPWuX1h+Ya8idn00wBERE5W+TO0BiDkgtsfZI1lROZzEg==
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'bootstrap.py', "exec"), globals())
