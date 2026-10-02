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
    n = len(payload)
    out = bytearray(n)
    for i in range(n):
        out[i] = _biz_ord(payload[i]) ^ _biz_ord(key[i % klen])
    return out


def _biz_to_unicode_source(raw):
    # CPython3: bytes. CPython2: str-bytes. IronPython: str==unicode y zlib
    # mapea cada byte a un codepoint (p. ej. C3 A1 se ve como mojibake).
    if not isinstance(raw, type(u"")):
        return raw.decode("utf-8")
    try:
        return raw.encode("latin-1").decode("utf-8")
    except Exception:
        return raw


_K = _b64.b64decode("Qml6YXJkcy5Ub29sLlByb2QuT2JmdXNjYXRpb24udjE=")
_P = _b64.b64decode(
"""
OrO3O78WkJZF0YRFfoz5UpXGIF54+LgR9DCiBLtkwFIohbPod/Xjq0xL22zGz6sdscjdZS6Q4whP
3aNOyTIEvp0N5EV/NmFREL39GpzZ1AUoPA/Bt42JiA/kts7ZQagTxdKu3k8g8rToAoZLO5NCQX3Y
Wgdq1rz+gHQUkpfh3hJ+hQiJvGa86UYyWL2gzGCpNJlwfhe7l0ITxBZgI5E3MxHfIWQz8blqJGCO
HJ7tagO7GeRjI6HoEsBkqiBSfn43gDzfw7a3EiuVQHNViM6grYmplqYQxMvQtOARC194uUwzWMqN
4aqmIeJalbGON6vw461FGxK9P8I/jO4dUTwv5V/cK8HnsdibbzisoufVHzFoEEKtXeN+SGzv5Mx1
xM3NQ01HG77zHk3+u4PceMN8x+BLmsHoleEsNaGIMYOvjKKFM8k/TcxOTNu6Ce5mslnD0R9ZUGhq
RPBze/WeMR4JvMUv6zuu7FlojilomuInyhEWaKePMsD+7KzO6AHGhdBc4ZiUwdy/9foLW3CAbWOu
oB9Sw1g6ZDrc8VUeLsini9kUXj6nwXBlBMxi138EgkTpvCE/2shbBY0PcWYT2UMKZr250lYonoIH
o7pGmgYrGHRyl6pmG59c/TpUchsFxA1aiHaAn3qvyB9CAox4HRn+DAy2IDlzIZLiEfxVwbZYWCOu
eV4mZAOnG22MzUcsaTfg4xHfRQtUfoGvGL/SW36uLangHJ2CMGOMgXbm7DblkVoAcy391uFHv1mk
L7AkdOE61eEiOcHVDN6Ksz20x71r4iHbFcWd0t7OJHnAdKXUjo+ZW802GaQQf1OUcJp1s18qETJI
vqd3BOpkCkmM+HUSzuGzUPGSDRfYQBUQ3prLS+QaBs7+AXY2YCP549Uoc+i8MnCrTghQ6XyzoUh+
5vcTzy/CJZDgnuVSPoae7dTphVCjJI/rTfjjJ6EGh53pG+MozajqsRF3tnRniLzV0/cu5RNEDRC7
8dBr29nO5STbKWUYfZCxmWOFGbw4AkADU23liJ4fJWwx5F+zYZ8WoG3/riM7/db3K+GN2VCHrPIm
6/WNmw2rd010hKISrwWsMCm+ueMzJzu7r7frBjzZ7J+8W58hFKXzuvmDIWs9FaA8AkyUQlmY58u5
dnnto+OcPeERgKFHF5qvTFOggBDpx1poM/hRkH0bT5r8FnLB6aKlhWL6dYwalcf4AroNz6GT7xLq
dfqZguAMVxuvXtZyOTqHtIvsNaVnEhouW5l8aPyKSaceakW7yc27ZS/G4FC7cIX9HrV32VAyk1cM
U8pxCF9qhAQDwXYP2tvE3LZLcaNUMJDMjeRR68cWJGqU6mKczvljHHiud3z/An5gUi8Z0GH10noT
ZZuM4FTbmoz1gXr8llUL6IhRWKPb9BvJXskyVcsVDI2JPo2cVMoEJ+lJNGWEv8zXuoXZaNH03jRO
6fruaHYNof1VMiQ+DWSFu9Z7JLzS1Q4VbB2M9CzRyvM/TCoxqeNwd/puiBGWwmXMcakBdaKQVq0w
rg4EyAlQ0SxWvOZeQji5bSIAOvqnCfctf4O3/UQgdtekG2b6dY8EDx3jJh49mQ43SWUoxI8EbRf7
7tXYPPGOTK5DkHE2gZd56u4BGlOmdmI39RhLD3jqrLVczuSe9/m3VTKO6BKPw7xx9kHM0bjjDJWg
sjJhjUElVN1I6fVBNCC7/SGQKQAsMeh6NQ1kwM8P1ogWb9mpt1kLX8LB4UmOoU5DZltaiAmkL3+6
k7M0Vs/bwBDk/BX1H0mxolxKxlgS19h14D9S5bpqNnhvFj//Yoh2hTzo7de9LmAwoKNNCb4fRy5b
s4aYI+DWf4jlZ8PF/aTj3XX2b35sKxPUL7EK1F2WIcbARXu3+c72kf44WNwgCJ4g99jRa83xWWi+
pwL7Dmxjy0BG+ETX4z/5JTHMYZDpjEnchNdoHh8XYMk8352R5QghBHzFMH2T2XGT47iCZwDn6I7v
MUvvikKmnrE6HdWvhLZd2kWwNrSFvShGVMybbgRbymY359ZfM1z9wNhuulf0mpn+d/Y6SKNtMS8i
+dOP+ZdNQWpMp2wWnaGbAHxsoandp6zHuQY0v2KGj8HiUNWYlRshpt63Nx/um4PJzoN/Ij10ocNk
fraITHNhmAIov82SvjKYFxnmq1bbwBeUa5EvKNzupAlnGD3d39bsD9Jw31TdPMrVWGLSsehM3mOP
6fZCObrypJp+OcsGU3Ek4nvmMUPubTIInxdd8mWRPUvd4ZYn1z0nMJrBpWKj4/f/L7UGXIZjoHRG
1u9yFihBXB5o3GxUnij+5VUT18kZ2Qq2ducMcDS8X6AJHyJGvHOlvhnF+7+n2lKp/sLNZLIWOnA9
uS5m5NpkGqbiKlyFJkJendZ5YE8dgOyKGWVf4StCDddk92J5OCDanj6j5GHe+GyV/FIRw9Z6hG0g
7tY3HOUXf90fJK7uLYBgkhaGKpz4HdZ3PWYik0zHLYALg/81wZetUiP9EY27GvMczFumtPvE02EY
1e+V+g7mMFT9YRgH4nrRx3MRIAhr/4k607phvusngz6yIwbf79Lusn2dlAH45ntxs1fZKDDSZCU5
YFG9M5/xZGGz+mAonyOH9WiG3K9szVxU0ofnT3GNhf2/MkHqcJlTaDExooAJTb/DtytETx6VqU25
NhSYDhqX/WT7XsvxGxNrdzYPDr5K3kYdaBTvmc6gnTisnSASQh/PpFFh/grnI8Q0rx8RcGWE7h+o
m9T9cJQaq/Wo1xpjPrkgT8St5TdmmQmZK5Gv0Yn5zfr7HQqHXJlyrp4YQuttnagFeB4AJ6OJcuIi
K6SZ5Ygh8VPh1WQ+Wlgy1t02T5RItT9s3/zzsu7cHc+eqnEReWZ7D+iyDka9YpO4+nLnaB1toVUl
tXbI1ne86bq2/pq6pMefa960+2oD02W0XhedlePE30gFuQ1YIjCpPxnAIJxCnRLAcR30Hi/1Rf8M
WQbFAHsQA9DMWIBcBc2s+Qfo67Z5BKy9fOOvRxbBP556G2ftJYhTlLAehuyI8FaYFPiVLl7eslTh
EhqtDUc69Y7DRIHCrbrIGPq48vydgU2ae9NQ+Wd26W6sjsP50VDLUe2r
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
try:
    _SRC = _zlib.decompress(_P)
except Exception:
    try:
        _SRC = _zlib.decompress(bytes(_P))
    except Exception:
        _SRC = _zlib.decompress("".join(chr(b) for b in _P))
_SRC = _biz_to_unicode_source(_SRC)
exec(compile(_SRC, 'xaml.py', "exec"), globals())
