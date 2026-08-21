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
OrOfNq/ukBhA0Zg/Moxp1I8fVkrDcbGh91lF2xWfW4xLKCvN4ehkd17dcwo2Voh/Ga9xIL/uzo3d
PfbLKbwJ03rw5Ypo4rf8hRiR0t15gWI4UzRn42PWuazV2LTFMap7Vq6L7iiO8TmxjHhpLDGx4oUb
Bwy3BfpWwfbr1EuvZInuUV7eGp7dz0mTlKux4CwcN1+eeqh7lu222IkUMIGhWAyJK9VH3nHo4uIC
YBxWF9qz6M1KaRAuCbRuEWdzLgSQaTfQHXrFTv163y1OGfzYtzI1Gb72mGpz4vO7YXaCpuN5hbpX
P012xXXaY7eC4MV4tSIdF/60KVTHbIsdZdz2/S+zwVwgDCJKmbJTblRWwvx2W4hxCZNtdeOGqEf5
jTmfJGf4wLZu5oPnrhNDV9EPYXDSSRuTqU2A06hAe28SLpZTt2zZxZStUQRyLYABWnp5lH9k7kjO
D70DaD0IcWt0hoYzPTGxni0IFnjiH7MR/9qPdFWgE/EA/OCmLnPIFGW4R0OUsS5OYSbx6SDCjb3E
jpThkD5zQkmuQvuaZuea3K+1fCfOYaGxqjCv5q20bk9Ejg/zbrbtB94P04ob9RjeyoTL4Vfbsd2a
WeLKt+QXK/s/xl3n1INmYmvAJNY1NPQCfDPiBMI49pFykfrm99Eav4qeVJbagtPX5GkzAcNnLNVO
i0S5he4cCt5X7POW/EMSrnWQgHDD9qjW8j5LVRMpr8yR2edpGnRQGvOZ/MYYH/5GwP89X1ouwpsn
PaxlGguAp2An7r3BvIgDzEy3wKG5FjFmtYDBKKHe98wu1g5GNBs88iYGIk6oG3o4UJui96LAt/pQ
WFmvtDO2teRHk4Y42KJfAIlKEEnbetGYaxWx0awDmr8KOmgjuZTFVhrn5zWSn1qxJ90wvQmL1D3P
BOTA6c7DKJ97Qsxa/g9ege1uSk//9kchJ4ECYMMlqgHME29Pt37tIla/gCa172O6Cjak0stgYs7q
azkjpFytjd9/K2UZvBoleWJXdXdxrNIy19ORH6rBtORAAi/nvyDOVp1EGMg6qJuGWhu1+sGg/5cf
T99d64HMIbNDCT79DRu5w+2ks1U9djZjU7FglZGFSLZRH7RI/xGFeYKfue0r+SosKr844gj2NCuq
PpK74632gAiR+J81F87rlkJhsobKTOqRyBGUo4gbszK0EmKH8CuRRoR/RIbDiRZtmHuIDZgYZmzF
7log3sz7EBIcATyQaQt74rgPm6Qt3Sq12TcjuTMHroavUq1b7jW9lBogSx5RdDAXka83j9D1fM+g
y05ppqO99X755/CYlLNnkntLkM4CmOG+px8imIRXLy9WN6KRC9d8qOWYtF/aHIAGpgKNS2eKWxKu
Q70tW4dcG1eOG3V7hsKQ1AVtWYD17eZ/llBCP3sB
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'colocar_laterales.py', "exec"), globals())
