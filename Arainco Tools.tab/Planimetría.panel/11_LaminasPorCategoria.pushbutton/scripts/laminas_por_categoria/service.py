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
OrPPeqkKUJmhMjCthl/jdR1MILBEBZBj09gN6hZLtUN0Yczqc1HRSvoorKVelmVDcuYcXcaYfIfE
NHed+ql0SO6WHt7Wg+OyTR1B/pLLL0/VqdihFcrinXKtkpxwJ7KzBcI589evzgEky7yBWGtB+bje
z2kUtB0bK6qXmwE8Ur1ElZvW+khrJ0sRMZ+mGmDk4Cp7B03kT58ryScIkk2wXiqh2PXjl4jaHoea
OANvZaYsAam28TSxgLUvtSjps0sD/IpRJxD+c0rCSe1QiXEtr+0uzyeCSWLMOeT7Jto6v6vkUaYC
BTjVAl5XghcuMjYw1boILvTGqsYevbr/c6lUidTtp2ReKnY3l5e/87PPfAgr5yMWV2s8tiflTUNq
jkGz/m8Qv57Z7WDzAGdJscG3rMsF0WgvUShenWr4PiYLKcMCCPz61QZOYUw2nTus3LzJ882mhoL5
723+Oc8s8hFZedCh80R1kLqJMKbtR9IJ87i1IjDQmZGA1F6JYhGSu4re/Q3o6bpUwDmLxEosH1rY
3Cjq1ZqYRzjQ5Y1XjgU+P3ug7T1I/iJynP68Xq9vyAZVlBEg2j4pw0Yq9BLDuXtmtqX6FL5yeBWH
DF8WgYdJnYtcqZSJdo0E/5BE9CsouIEyOXmSxDJKONq8uE+V3i4E0Pqln+6mR0sbUHqwdbpKUvNS
TVGtiH2PaAODnVJIKpJdxb/r6Oqvs0/EHU+anl24XyU+9sTxKj3ryMA8TWy5QRLw4Vhbe5D77CoJ
ibmuMDo0TABKJwsZq8ZT1pUWr9gWJcd69c7/C4SPGQiI8GFhit5lTu0uRxaU7d+OXBMqKkWWHhxV
7EASIBiDjZkbHEz7/9znvfaMEItCOCfxMLDej4y3nZ7VIA6equqVHVlyc7RfHUQubTr3iG1h8ipw
QswH22m+ZmoB6VdBEBKgfv8H3jeqb+/z7rK+e7mG4nk+sA4wCylaNnXdqlSnFLgqZt5d4NDGyKV+
oHoGxm8SM+5AiOn0vBtpNFiD91PDdhCvULqpDhIY7xyk75A0JD8RMKGmmFHZvv0ClAEzHdGBaZIX
jDDSAkM/erZtpAbPrzZqszaj4ZNtbtSlaSPL0JzSjt2apLfE1j0Nfu9dWFJ4XT7QSY9KAP/nE0vb
wXAvxneBVVYOXJ6AtA/+V/ZShbOnEXvZW/07nxjWYN0BJXlyWrScMcLo2z9uxIij8viDv+8Da8Y4
ERnn0xSINUbyLdEmqd76m15YVMd+XY/M8VRMw1dMwWoIeOmDIawJrXOt2Z34CGd1I1rvAioiQ6Iu
NL92PvkOQ04XxqKu3gt1shyCloHAZO1E2oJi4jQ0digjtjFQai19mAwTIJ7+q8IEe3npTLz3AGmK
sfziWRStB/ypXzvu19lhaij0fdkBJOzsn6Mpp+woOnUMcGIKRVS8TORlLjgTC+WfNnHfj6stMHjr
wnEwLbbtMZB2Q9AkI/6xg+tJlxXIrtO0SB6BHi8CUFN9xwzq4SooGI5b8eNlMQNEfZ9n5lzEDBWX
nHEPWZGduAs3AkCDyHgPPhjc9WYTLLA5TGahIeGWgYD0GHkf7EVHg2jWj1ONozipo/Xuuiz1LIPh
nGhSfy5mmwvwpU3/cjvYdLNT82klBGus2tAi68nOCFt55FT/qydJElYy9Ciyt6aGfOyx7zz45EhX
lIJD5TSsjbRCu3qGBjPnB/LPFQXCZrjWGsVw2A2j5IzKH2XAwW1Y+BfivnTrQ2S3C/fQ0rQhxio+
+9u8UiPXsocZyz+KgUB6JKwJWcjD07ab+GVjsYFYKO5YJJkZ78z1NGmcR8zrxkOCXDNk/MY//Bkh
CwOqaaCAwGcHyruaHlgPnW8UqN//P7axIHB/2F7+sZhsYkPurTu1YX6GTbudx64LOA8GYyG5jc5y
g6+42KdOARaEZCV3xXC1sOxNJBnGxD4VNmVNSnILDNcMqm4itbH1d/O6pIYY8wpMqTpHAsQUzg+P
2LxNpqgFxajlqySwoID7NcQrgzHAQ5YoYvxXtHj4gzmeR8ApFaTYoZ8R0tShB5HEwLwsl9ye/obV
gWIXtWoaKgUSaW2emjk7q28HGFDKjtXKYrQ6Hie6P7dGLznB1XI3Mmv2jAiXUaxzPaNqI8s5ZHAs
y16ToPENu0Mwp1aza5P97uv6tDPAdBznXITnzsyeKoDrweDfAApuFVbIuCZQhiUCjezv+kxS4xcc
Q8qWRM0XW8eqo9ZpXpQ5Kkz9WiYKah5+C386MkfvqUTCmeM0OuvmGKpln4S9WFTc4CZjpTekD+Uu
24XTdOxot4RVaGZXNbX2aKKgAF8CdfixG///O/Lehnfp89v0Zyf0g926nCUhuJ0rU9hjIC9M6zfE
Y040r5gVZePWMgr+jwGu/H1JIdSJ+W6ZRrMiMFT5ohFR03NglVBVM5lYik7ZdJTxRu68iNnJefNi
a3048OhvPHBp2T2/1f0QIlZCG9Lics2v9z8XuD1lOsbgHSrbI60ld+B/S3n5sA294GS9k7iy7e1J
ZGBML4CeDFIyTLlzqrABVDIXcfNRjwkL0a9RlINsnEVz9pJjEKOeSmlFl/SfdOvUiPPx/os44YYO
pfYhszMCKrluey8/HmZ4avFdwiHt2OaBMdv2T8DHw7i7q/GROlh5uWN4X5uZ9IOiPMWm4n7eThVH
BN86qeEaYE6xRCugWEIjFgaZnSfWRQwt1VCDFCFrsJ7ZXlQbMGevedlRbJuzqyrQ+Zral0aSQlto
fb354Rfp+EZq/LPLRS9NN8cGA2JVnhILaPgPwrJa3ibCeLaR+woQrkBiZLFU/AxYMqZN5939A/p7
/DFZ+sCVi3jEwXWzu5uAZIdziJRb+eTniP2IXC7U4FQ+7rBlUx/tzn5H81BXMTdDrDpeC6WYtBx7
9Lu8//M8RHjPt6BA9Ven95z4BAdTM4XG09OKZMMPXBQ5jAn4PQdPRDHsU3M7hPzLeZZ5l/W/kJqX
hWYAys+xrP2xrXmb/OMT341f7ZG6vKOlqOaEZAmDMZzhLnpYhGHkECblge/IT+/VOiLdaey/6JKr
GqVyIromRF9YG7uPCuTi8ea+GJsQmbnJOJa3v5nnsNODILFA8CRP4mbsVYKoNJ5dzyu9HIIodIpv
WbGcf+64joZ+8c6NEd14gVCjJUqC0xHzCHw+0tGwC6zvuUtiT+TEtgR65hcFa9NBh/2mvex+1gcB
lFHJc7pgo9c2rbMHwm7UPxk8pUZdMWoEkQ0GjLbOYDBKTOwO/YiY2LslWHIRYFHdeSx4moYergpw
Q2EfnZVr1u5lqEhttyHiLhJ08sgDvldHjZXeWrDtIk2xqlwb0uRIiL31crmiLTxsMi02odmd/k5H
2a63ozM4qnOaguSenAl+fEbDnkTBWCxc5g7UKyIfuc58KOsJPveJz1Gqv9nHXwUInfjwcXguYwLO
us5yYAdV2eNrJeoOzalMp7tmx0eyQYMDDOKD1HROlYKNb/tceVWxnGrxvJ2gk9fpXysPOdqvHOcJ
OjuK+5aX8AUVUKUIz1bcMae1VlFWs03C5ugip+WV6Mtq3UPpxAuaj7mdPEJejGTJotmXeo7W/SGr
XUvv/qVACyH+bGvce0QqqO4GLKWN6oYlY2pt/Iwy1WQEcJJipq1BM0KHg8CBj0x4/iKHLRTz2uQ7
f4zHSm3UeX9n5EA4HDTuYYHpKK4nKZfIUTTe5Au1loHxMkSR1C2uK5qSsOV4UtTuDf0MD3R+3UVi
nCHdJTtInzbUCjGXRb+IWwGkueDHlRR3/9qacpNt01d5+DeU+iI4EAVvLo3FC6r3N5qnGa4+JxN7
/Xknp+O8dQ9oNn/DxWG+qMyHvfk54xxwIgjXWB897m4PRDmvkOLHYVgYQWuQ+3IHQvJbsrdm0zEN
fpZwd/Ckd1ZtYf4OEofhpBreP7n1Yhu7/tzMe7Z+abx50LGj95vreZ/iA1qdHWeL70vSJPnMPjlC
OD/cq3Fvc9RAcLgzeUdi+rHDUe7FmQhIUzuT6dMOZXz8nPEKh69nw2RiiGFmJkVLTv2aXA4ca1VP
+J1uj69YzWmr2ym8hG5ixsaw5U5s6DeP+vSEYtE6lTFCOtlTXl1H+GYX10KfZye5VVNJc5qYkfnz
juNoITr0BU6JBgGF8hRVW6v4xEAgtNX67lT17789hIpgAr/xEQKRy/MLwPwDVu9WXBHRtxPgYJhh
rAYG3j3GWi8vuxgFMWqMoRVp0lc5TxDZfPZ2zw9YbE/fRnaYhERmaX3e2lbfW1WntWL7Bvk77m12
Da6xnupMfPe8QQBqbUWBDFk6VjtRcIV/CYwBkWYFC2zb05OtzlKzyIpPFVw3CNfhcIkPmmQHgjtU
HaSPTtzMfQsyoF3UjPKItABiu/sCHkhTW2R4qWF2tX+mbqWRqKG9rY+jWVKpkdZVDrWmVZqCIfrI
P/u7Sv2YI8PLwwwhIqj+Dt+OGBzofPzwBAmuIiQC7MHNCjDNxA5SVHV8G46y+veksqcr5cpMLwyi
pIKfmikQxTxbUT1lfE0P4ltn8qJ+1TIUhM9AtSw2/OyoHpKu9mt2MeIp6+MM1uxix15DPPkP6N7o
kLwQ1JbLqI5DmEClbSBFE0agGlirbAzhwJt8IvLywkO8OX+h7qx3HlqV2lCExIZdimyO2UXZ59Tn
b2QPwqH5DTDYwIY4SfTqR4AodR/aNcDuV8Y3XcVwAraCzNbeH7lGI5t3lE3XagAKKFJRjZY2/FMh
2s3Nc3QI7F4ev0188fuQrKVh7hMAKI/ue1Xf957VYruZ/SELhtcw9RomI1X2C+Rwo7DcgCjTUxEG
TCALzMwu1ZgAGqQpLpW0G/tP/gemIJBKE0UmJv9bpMrWdfRfCevx6Uc9lyLnYQjIPkOtb49j4L0I
z9WHiE1KaPL3MYPTMgS6FNVZn/oGIvoxDjBDS/tHNI6dpymWOJll/YvlfLRuod9GwSub/TffZ4OO
5QjMNMaoy9XTSKmFEI0hLV8vFw==
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
exec(compile(_SRC, 'service.py', "exec"), globals())
