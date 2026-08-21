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
OrPfN7MKqGhE0pxHKLo34wFoBfKyBNhhy1nKbQKJxqxgBBrNpIAh9DTWxP5lEwB/xInthcZKk2yJ
3e5tU3gJ022Atk9XNcH6QsdBeEk3FBiC4zS+qRWMql+wT/4UWn6UD551qWBQaaP7FvGfLETFTYNk
W/JucC8GfWoXlXtM2JwOEajjMCgL4DdWkDHtXGBW2+HLr2vPED38PqRcR5pDIKRl+FVFfUO4wtmj
aU1wweatA8m/rDdpfmXhSIwYcg/CforzAMUWt6V8Xnb7FAIOSZZMXh9FDbVX6o3teCLEwoQsQLCo
nCj2uy0xIfqXMuEiLgUXBwECNejj3uVC8KH7tF7kV6rAREvDCZpAFom8mtu7lS9FwYjEj414078H
mdXHKJyP+oFPBVeBkLnDEPsk0Itx/7VBUuSeShsWPmXEKWcjDvkSOJyr8C/kmTuBOnmo5/BWWYjq
h739KjijpQ5+I9MrK+80xsP22Er9oR/OTVNoLjBmbEj8SV29NouWwWa3gTEyhO6FWrNv6xKEItdf
p4KRJvt/nTdN14tXVerh6YRDdGUfjGBdwi2bXeipD0wuY2G8XjyuP8WBxqNtoJ/LUTCWhARlvo12
oKRqoqbWX1IVAFWRLBjoV0jlsl07qC1uXncCSnLF2A/8yqCUmukXZEh8pYTwy8Qj5pWHsVw0Qneh
Pk6c9p3ke++gl2DFtXgYg2pd/ko6bYH2Rj+XcECUIG/z/ZSDTaVgyRzpeP2BbLdphiArCwgE3ouT
cexwoweiHsrr0b/i7mMQbAQ31kM4nzEz3piQlnkmHS2pRzoPKg7f0H+CwuKF5LtQTDADv+oPWnOP
Mv8s32euY1wCotlBlMpkpuuurusFPL3jzAS1cD0dbXX1TETeV21dhv83QKEz3OEWH51RQzklFvwn
gTpcluuLz3thSyL7fQ2KG2jPg3uhPBJZ/gTTTQGha0UqIVf6i2MO+E//+oEW4SdVj4p+pvjQrbU1
kPV6vusQkd+SN+E1QXaUfdz1VvkE9DEsOV6fyvjha+7mfDFpDCdl/xxiGj6gc/lcrIfSAEFv9bra
aRMyHFmje1dgfd+sgo/S/rrVMrtBneCFd1VFxOfXSdwgTE8I1TvcicdFRHc97+szFqt+OZav6jyP
EdhhT/s2jZHNPCOgV24KFTVaxOYdhkI+Sqpnw8EnetDHdyhEBoSMxjQZwiKTZEU3+IJh3SV0y725
nQkrMG/qI7JB/W+1ZHr+3vpDv+vJBmml3erRFw6xpRyEWZYoO//Wsmw+ftka5/l+qn141iw3dAFu
agcPGdgvAGTjz5m5rzK30XrUNMTO0tpBIl6vqRI9MYOAdcQ+6c5eK/mhfviSB1yg+nytTK5+qZE0
BfaDoIqvXIDH6PwAOsuzIlStF3+JMNoD7WQxvi+qVGhMniEvpzzGhSjzGzp38pC/iX9edqH/Xy8y
gHaQtHRgm47T7v13S03VMgQ6KDk8xp8FHCHV3j88ap3Kd9SGwHgVrd8s+bAGxutUg2c/rN4Ifq7v
XNut1uo/ocGoFNxke2JfkavUhPb8DCumybzGrITxTcrN3vu+2Hq/qY3R0shZTQII/A==
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'laterales.py', "exec"), globals())
