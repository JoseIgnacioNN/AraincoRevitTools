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
OrPXe6kKqOahMjBdTtH0g1byK+MIJkRJQa/jDKoZPvvvtPh540f6H52+oWAD6/PumaaxotiBHvyb
pJ6JIdg9S4yliks4kh3FVF5W8F3vvZdTyqaN+Of7pCjKVU/rdEpMw0J5pgO67byjOMCk4W+O7Ogq
ApDuauNq6ucQnuckdvhXfkNuEoMTnkRTTA2WeCZOPDbdRj0G0qs7UNDtuxIkOI1uedq+AYqaNZs1
3GulVyWhTk+dY9ZL/dFxekPhtp1tS08jz1oGCgs8qVkwcmusbF+h9wLb7q1EGLhenf7kiN5RH+nJ
o7Ga8NsAlFNhv9YNfVJ9WFfa1l0nn5xeoW7hD1u2evrdDW2JGFvYarjpYAumJYqAgBvs+0wn+f2b
r3CWu9MqRPBTn+CTkHFONswhHjmCBNJpQThE7bKMiT+PDv2mHNqEfaRKwpKewg4M99upi72ctB2w
ge+rn4rryFIfJHEmSKQLTrAMXXrQCld1g46IDHN9O8lu/piQvCYWVg/NU3dUThSFh9p7DJKyb+Er
QKY2922z/xbPowcTUdWSSqrErtmo6RRJ1SxZ8eoxnr0V5MvFxw98REzzQGdK0ota9X+IV4mwJOZw
S8HXpLx1vaz/amU45M9wYV6VmIaGwoQV3vS0r7TaKox0B/EssmHBr1qf5zIy8VYl0uhbSSwERCtV
6/zRW2Du6Zg/3jeuBnDv9otj6Ka4vtEBZTZtHutu+tk5gFAuV4uGxtfaxDzGwR8pH+H2hliBnaFH
62TalKDkBNIx2ZRWsg4/5FHroTF26WNgP1GWFFNQ/1NJyM91tAezMFo5HgnG/kRdjga9l7qgRGr8
1PI59Q7Ee6k6QGC49/Gi8ZG/nFBSenP0r+x693LxmMsqYOlsofKFrBdctEJGtQxGZkt8DLbRyHjf
53WeWcY5Advzxce5LRYRXljXhJ5eXOStISIrYAYFmyMGdLV0bud4BUKQ4Vt3abFYbYu2O2JKnXGe
ZnGmeNdPizKBbN1xkwaCEz0z9SlErBO2BXYWAKn0FDjl9jy1YnxGp1B3bedfRgp4qxDtjfgjxOHP
tSxDBW6XW/7XziVW/OgAFcPJVYNW1iIH7a5k96TaJsDsG8Ob9DIEJ9BqaxENMs7Qj+OWoIjRt+zF
P8jKWi/zUI8YJZkhY0zJuPLcTTr+XCJ8ZUocPeSacRf5tXCGRz8CicNRbFc3jkWwRtS1d/rNfON2
S9ex3M5shPLRWOn1RMDu1zTJaS+KBQdQ8nJsxvMbutO121qvP48UQ5QGrwG7T7ySKes1jnc53I/O
YXIQXJhZxyZljulaa0n8dr8SM6f7Qsh5qg1onUy8RcJDOAbGbI3shZC1I/LFYonJlJ6huTP0ofH+
M18G+orhy5afo9clXtstmLebsFwqIw48Osn3Vo9EEW7a9XK4QY2etx3dMsb7tNjlLlkZ7hZOssR4
JO9OloASLPFGvHdsuqm0gGeArpYXApEqAr4H015RHNU4+bnFwGZVYe78RVC6MttBm8q1E85w1ouP
p3BX3MAHwZ0dtGEFN1nYodEknVy15ur1hOtmmpI3pA10OOuH5PxxbrYH0FU3MFLt6p+INLJ/Gzk3
MPxpXGG9lS/ZJSCG+8mORTSx6xxGhg4YzGr6yjbdchTzyydg53tBjrr2T8k80Syw9/4v56BIrkbf
NBTb7eLyIuygR5fL/0zMeN7v7xLO+Hn9wEItwp5pwHYYRrDnZi2rgwLRA9BEktu/+Cl0Sc3OyDTo
DjnIZwrpmHq6cyLADo5/VVAo1QVGf2AEDsurA8XUqWv1Lijamzphp2ycw/3Xa1Lji+eJjQDH7is8
sjBfLzMQl+E2CQJ1vHwkzXCJ4RuO6bMVYXTygl+mnCm9aT0DSCXUQNEtOEMgHqRvf2NCFIlQaOFY
/Qy+IBn+bnmKa7kE1kOkMPIc2p0NuzqS1VxHZekuLlng1/YSvJ/EIPHNiM31lEr5U4HMnWnJ7RI0
i1r47UQrzd6pDN2T0kGMRw0ZpvDe+XQRE0OJv6z85gY+6ojvGhGwlFF4ed1NG7RnfbMsysn7Vhpk
vhO3wW2qyfQ0x4q8DoXb4M1GBcyo3r90F4R9JzQanpqgZsib2AwlDBC6J6sf5V/ExVFZR6poziHa
HIOJdlaqe7tkbe/QoQBSXU6rn71HPxe0+MPtwOXZ5q78hwZsirwmBnuUTM/i+9rymaUmnHSuXkNY
KF1gJWEVjSfA2NYaWFaORM91BcHK4ZvGQLERaB/NAVG9xkE8C+8gygro+Qf7SI3gMi21NYw8dkyA
Crgt6uX+5WklvFdvcXv05Aa2C2/sdL5Ry3wHggvQw1Caq+BdjtMt9xp0nX/zU7yNSOqdLy7N0wKY
UCehehiv0YEvRJtTZrJwPvoxwdsV42zWUCbehdI2j3EUNcUkQdjYKMVCkfQwd86G+SMtQj5rI6T7
DFYGYleors6IsNt09lOk32COMmteJyZwu4EJ+OzU8U7qlsb7kanBcyiiPXQvu2IU3OgXVJOz7sbL
829lhSvhGXV5OpEt9Cf4hLw+rcI4iv+6MBmxhbEPCxdBFU8KI4WZC99xmfrAga6v++VeaUUKUzUI
EvlJj7xTiyTpbikZpIoJ3L5fOOLEd26zGvzM57j2wh7jxcOPmOESICGlRhjPRx0mGd93JQYbwALn
Cv0KMQMQEbiZCeuKbg7kWY7uPxQHhnHliYUiM8Wb5ehahpPn9h1I7dJ7cZlaK7VPS9pjHkgkCbXU
rrqgAD9lAk1G47LvxQUW2Jb8hnOgYuTb4bH0gbxt5DvNaFl6tzRTXUxGIVx4orNbpgAi1uoXBqMa
zLDVivE9WX1E/k7QqHyAw7lD4G4yfoiKCZ7wLVoIjLhP4JoTvb7WvMjJNdRUEbFI+meuaR/6Hufr
OnmDmebEw9asEPsvVcC15MGNXoYtOwFUZWiA3fyZDLF2+lwbPt7fLXXf6yQNo2YqQPQgMVgxv5aj
PIFBFOUyU7bZHXrDOdUz2tpUr1stJ9V186JUlz6xpnQ4MK4pOc+wmKPJL7k/X/k/lascu0qSNjh0
bNw0WZsL7jEcg64ftg/wgGxOqUR9Sq7YrmEwGtt0Ey48+hiF7URiXWh2D7BslZMrUR4Kw8yas/Bk
6KDFkwGwLe86xP6FGqgCSO5SBs5gW1Gixb0ClmX9c4UAALSn/OETXOJL1ggPfxTKIJXf9jR2Fgsd
CZQ0nSvVjvC6zlNyUkpLRhTSH4w7vPrq+9DgcKm6+6xfAXG1Y1YkBuWNHbyuH0AOK7yXGdvOy/cS
ahC6iIPUwgtrgamjC8vyZDHDB3B1loPtmwYEXgqnNRWhJlarWMk+t5lNLQPTSCMx5yi2EwJ31YHE
mnPktMVL0gtZXPDA9B3Ga+G2cvyy9gWZ1ddAStdmItqnOndR4soB1Ot8hcC5XiC6E3AGZUpoprjF
TSDpAD7YESOxGZDIC5Q9yTkwP9CsufqPP7kDmK+hK9xf0IzvXK9P8J4L66LxL8L92kkayBhzdSCW
zUSSziXZzdNu93pdShNUBCm+XzwU7rqp7ZOhrT65ATzA4Kdmuk0hixOH8/yGML3perauGeetCNLB
+D400+BBgsoYibTnFgXCPiQekpY2vBZLtBbm1H7UELTzaWcJeROCkaQKY518QIUHyyuXtbzGvcM1
Fpm1Vi0shD6FS5ZJ8bGporBTtvRs+vL5IwUBFde157gPHLdMEQ/jVY7MZ4Wljeoib2UkUkPsSN+x
KEI+fuUONUcQhrmb08Ss1wzWWbJR6Vr2zb0ftdnS/8Pw14siLtvRmsH14DLZ3G6UsfA/zgfi1LFe
jaUfTqP1AuDgV5O2Kyf9E7+2HgJUwC5Xws748/QNekGgJ7jmx1JJrBY73G2bItohwrz9kyJebz/S
4XlYadntwf7Qb548sWyn04/HE83iDZxUq97Eb/+rA69/08n2qoV3YpXud10OP6cl9/X3WjKShZbI
l5JWu2ZAN+wdtkADWvKZV8WZ0wOcZBv5WuB3zrcfub7pSmiK/MyZkfaChldoIopu8pQ5NraWpONW
RttoMFNCf17Q7Irq+IOAV4LZ2LF/0f4ADqFFwlbMJ2LgKTE+FAHXz8KJFcwmVwNjJ0CsToObH6Uo
qy449pv7GHcKDZN2VMKDEUAJJOXe+9lpzbK9a0l6iI4=
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
exec(compile(_SRC, 'contorno_hormigon_vista_ui.py', "exec"), globals())
