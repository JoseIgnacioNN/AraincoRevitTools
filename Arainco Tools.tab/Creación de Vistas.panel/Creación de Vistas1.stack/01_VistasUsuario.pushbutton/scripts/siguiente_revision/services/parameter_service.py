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
OrPPNr8KkBhE0YRFKLo3RXiPiEAje7pFJHd9pkwgwLbgeRIIntTiZ8FQAR7FE4ZqkvaM/JDDJNwo
681cyko9gLHeP+V21ltcbJwSs+IAra+lmUgyjy/ZpkAMXlCWsDGjdiTplrDsVhRgPFtHd3GOlLPx
b8hs2HXoahtdZwarL07mMDw62nVpVMQlN9SpGo7hKgT7J/Tp2inVol5thiHbAoCIOGssVShR5RAX
wOtxvtPJ5u3zLVGoy2FvkU+f4Z884Erx7LVh/AnBx2okS80iOtA8eTkZJ+M+KqOJJs2Fra2gy0Gu
Vz/PWSDB78VA5QODLUvrkSBR5jVVL8S/NpeInQIR3faZ6z28F3uElqXS27wiAr+Y2LBRg/K45DEs
d7gM26nIZcsAwDLUHVvGBHfHJXZJoW111bhIKDwK7Rk7wD9sYadSEdBqk9FMTxC89EE5UQjunYTN
3cwR+m8663cMtGmA9L7twwpatvezwf+OZI4hXzWwnQeRBoK6vFSode3KS5UyrhbBTUUk5qXR/R3x
eiEYVqTxUTruEgRlCldNWvpMBM443896dvJL4OI9RtXXlvbj6hnXDSDtaVi8/U9pgYM5yPjryFvY
UVWO8OcmbiBLiZiH8CnoLYE2QqOtdE8+icx3AtIpB15sLwGboYEf0Wg8OmgYiF7vDfxz0wSKA+gk
+dWmYnhzfci4lgRjIUuAR+eHs2FaDNpAD/TWHVXXbLXy06nEF5CBGlMIjWiqqHgoiuCznKPEOpEc
a67VuCWLFk3GTAWFDTwGoMQtEFN7Z/BwgW4o8zraxBr5/X782tbYW+TsLGH0woKsgyp1BxDjCizU
YPX29e/4nVFqRUcRe37bJhjoDzcfJviXOvRdsBZRy5S/ENwcWmZ35lBpJkFPJGqGV8Gt5LjjDuiq
dV3EhCxG0MPWYKXl5k3QgAi6PB4Pq+dBiCSvwhE+U3PncGhYj202WHrL28kzY5l+iHOGkhZEwBNX
LD//42br6ylKT9kWj80qj1iSQPvVM9REQ4uJEq7vjevsuP4mEsoaGXEGcm1BkQqHC3u2PtYYV8LT
2mI+Jbe2tEdYuzBFt9nbN0Rtj1GhxrUhGiZycCU2LyPlUHKrMlxklvscG2F9D3Nzuk1i4gWaUxH1
WUKtn/L/6X/ZbwepR5JfP7xkP/O3BUSHyg5YZALuFqS6vyJY73Wg+HeqJnmxfVBqOfWkt9VSD8Bf
1IGVhfotKcN4G938Y82ruL6lEs6YSHsOLDzfXvR8FTMY1BL4tUBBktWgnu2YURUKlxesgdkmqCP8
X7darB58Ufl+rJYP5M0CCbvfx8Onp/H0fwoM2Huacq/kLABOiJ4ABPrmJsprR0D19Y799DR5Vm6M
jhztuVxivR19hsCDr1h9+JTfax90wCD5hU6GPtUCEYCKr8/IpkGaBVJ47jtPem4U4u8pdSrzLkrI
RXfkM12JET0GUBjgjqOYDDLcikC8feK1ALxjtFWvv5tB0f3Kfs2N0zRtfJlQPdAd0Nohs4AAxBiM
lgn1FefNvSg7Zi3T8ByNTiGhNdn3Y03asxQQUW5rPBkDwBud6S8Wy49Ei5iRFjbvD0LRTnTkKNa4
boQnGGWmxq+qSxZ/vkFzdKXpdcaVikrZvJW6soEzkEopQJIeqTaU7cbuXsRxoBDY1buQ
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'parameter_service.py', "exec"), globals())
