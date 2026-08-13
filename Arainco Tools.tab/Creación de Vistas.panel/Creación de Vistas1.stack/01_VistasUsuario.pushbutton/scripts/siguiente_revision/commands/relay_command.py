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
OrPvNDkKqB5EsoR4LTcR2NKUYj4CpGkKgDHqTwO9M2BKXUMr+jTXZAYm04cPKvvDHyqQxkGc45k4
/YTLakL/1Y7yTQ70sGLiCqLvdw+lp3NOsYC84OAXV3tRjse2VBe7wAB7Hg2ZdXfAZWS4sCHLlsdv
b069iOE6Zz8Cm7f0KSpPxpuTngfpCjMHZIU/oMxbfgN/rEPIS+k47JkooOhzUsAa/ul4SlfwKVlt
uRqxdoooR2Mj/kfLUXSWR/dQlzFvkxv2Hb8yIWebg3Qux3b2wNkbRT98RjkRCn6h3dLc4/RuYmNw
RvAlRBvE9RekwFaiVnl5wzmPLesSuLjNiBZobAgLBoZMVPSNHWyWCcNsZAbXu2EXP4xl3/RcIGfi
F3Hq+OehvLtYcjTVChdCMRkvJ0kl6caVmeOPRQydgYVuIG/gL8osdPZKPs9oPeSkkNK3WFZcSEZA
EuQM7tVs6V4jAGnx9OG2f/ESIYY/b90G3EUrUIvEFdiagwyglwh7pY9YU9TvPSOc9fFjCImf1GV+
6E5hgi+g7tehWWqjeyMc1QIxWmObA+7/hKeyiDu1Hfsk1TxD0B+YbrX6yZgLvLtWj0ql0zY6HOxL
RQclMTIrXKSROs0FWUpNRgjuSxmymIC4GlYDBUcgUeW1rW4HUJ2Geu6iUSDHAJeYnDOkkWZIxWid
Qa6BGhe1H3iYll0P4oxDSUT3RL9zhyoz82HLwgLnRS0UgSlkbwClq36jt1tvCw6QFAd92wC2GO/R
7mTIBEqAlLFbEzujwUmOPAPUq/5cv2v2hh4RhNfgsX0qE+VaqVm6A1G97ToQmmYp4htv3wl1Kzx0
62lP7piCAoBS8fVo1DU7A0DhxbG4HejURjcZH17FnTcI2rlWBBUj7OAqliOHbBb7ygLRkgMB+Tbr
Ryvz6fU1lJRzgEznugeVE31j8P3d4WsnEb8bNtuRr4qeSO9Pmy7BUgeX
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'relay_command.py', "exec"), globals())
