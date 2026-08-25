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
OrOvOL8Kr+ZF0aA/GjZdc0nrBxNo/gJ/IeIMKO4ZQB0h2sE/8adSVbAxHNYJjb/u4KvM3iXrU6R9
BpsCwVTlDQvD0PoWFzugEtQu7MZwg8jOlDHcnNbNA6YzrkM5q+bmMT+jtIMzi0pm8iFLOYvew9Zy
kx/wmZvrvgihRQ3rQAp4C38GjqIZkAJpwJ9kP0Gs1nKWzub+5oWzBv9e2xqA310QL2OvIhqYAQra
ihzJPuG9Rfet5+ledYa0pnhl/jXzn8wDy1g9Xzgn1xatw9VeMU8PIhZUdumLqxKzobbDuPoRE3PW
wWD3sfePIkmDMvTMOIV8AVVPkoup3it2Xi2j02du900u0bmspPmnOCjeyD3BLtp8KW3gR3L4vE/+
uRpop0wdJ0K2//O2RspcbJnJoBcoVUNfchS9nPjGufW89Wug07ArTv3gkPRO+KL5DKQR/p14Hk1Z
W0B0DjSKlvoQhHxgyAQwfgEvjGDeNAftdNqk9iqrNyhL31O8G2YpRrVCuQ4bXqZKRET97i/em5jH
VWIZ4ndie40NjorgrkOzbXASXbWqvvzKFc9rW07viN06pT01GKPaIi4SXeRHlkzggWmosO3v2c1K
jWmyyh3AdR6LhFDByHduYkca3Tcdg3YbFXBm9r9tM6rgd6yUTRUzbO8B0LP5Q+Pm83wj6+5pX+0y
yXt1Wla61SMdpNQ7L9vXFwfJbTqHaNvfNJW+u30TDrqiRUh7SjgGbUlEwfi83m3wXrxTx01WrP9G
2Cltd3LHc58oKWy3ppafmdITm5GtqfVStL+RhBeFbsqPBaODYGgYCle8+A87EyvRNe456Txirntp
z/tG6h+gKkMWQ0LzgR5iBsnXl/mt1B6+PwKKegECwuBFl9Y0D9wKZTfv+q5xsklCdC2ed9WjsJEC
wG6RXq0KspiUib7IlfsOQPGJISZtu50Oq3cFtMjvlHKzrJzEgEp9+3IJtjFP/n3cuMXjuyBfXUq3
7sbhzF9j8cAOk+CCmJ3weKdAyA4PRtQcvRmRyJhjcIP/70F41M9TzuPk2kdJE3RgkumIIZAX3tUE
DdqLBlPm6l1v1GQqm0tIS5ny5W6M3Qm0rjsBYuJibLoVrIBbe1CxEKIsZkzvWaXi/zlTJ2u8Dbgk
PwIRoImTqlXQYv6laXGRaQs2/rNuLvlgLnJvgVwgZ9qBPAgVVjx95L+seSpT2epWLQmuNg+/Lez6
Ag77QgMg2a9bY0yw/o9FxzkfsO49G0jGpDm9FFKmbdX2rpEJTnuliVtCuuFliUdoc2KPrcHWt/RP
wovAhQh9C+UkBMcslftDl8dn4jnQm9zENhLMBjuxjowuvl0VGTwZgPFAgpwScKYehlnZ6EvqhWOA
4jDygy7ur0rKuVHDFC9jq14OLcsWpO3mTpbgUAkukEQ5212jmya0L+YZFWmlU9eyCfNue/kyKz8D
He+6Ddo0XWHOXax8c5c7yqn+3vh1j4RIdPk7v4MtTK4KPhEHEvBXXKxI5DWYCtVhCuEek2ui6iQf
aBv3wox7hYtPN5/75/JUsgjeP4+1eXPGX4unIzChCRO2SL8tdAnc277ADgEtX3lK++BkxrndEecD
r2sL20bRNiqUd7JVJ/2FCPhhn122lCmc8jF7DZ+Q19klL8PxIATBMkNHwoZs+S7hTrpSXsLQfVSL
ciCOrPyK6lRVCKM4zvYlLmqqCQ4GvMCpp7AB3uo0Of+YQwQgFmXr5JHZpUxFXuE8nHqNzHrQI9od
Ijpbs1dpapojIOjVzF0+YnxjewFJ9TuYhkmBsbL0OsG6zl9bVmlS+1t8TVUbw3g58WKvOg2aoQT4
SyMJ8c5xqhpjmshzaPawbKuORiCvLtgXJpdLwQVCaE2dNov4ePQl2+vwzpApt72UQAEb13XCbWxp
4qMapOzNbWEtsuQN5kVOr07qt3MXVY54DpgIK08tYw0OJTrhicZIqfeUffeN16Tt5IXP4YM/Q/RL
pWAu6yqXtiAtCnYrX2OuHWQM2a9zHyjIRnvsXT15PCCjDL73N8vEjMwlsl5B/6QLhaK14LteSu37
xmVvZ8ogRisCY0ezqCSukeR0fxixbE5t5ZuCMVJSccI/sjUxcNOxKnISU9kTqMAx9EQi1UxgJoK4
nEEG+3DEbpYjxIlPHt+SDaQbUbJnUsxZOq/xOC1n1QIgVpMSCHe81EJPGxA6UyC2PyIMxdEmFCmK
z9uQgVNG16ACO2x/6j65KYkIi+mKxKcuqngLxByecqzEEt8aS2JFJmT88vNQD+lunzDKa6BXDoTb
2mwSNJuBG6Y8EXCPvajADDdqT3oqdTr6vtmCNHsl+rMn2DOZkh4gFXDefAOezKEgX/Uh+YqVs0HP
M9hnhaL31qkPCm1HbVsCM8KZ1oiMhyPWAJ5GBawsZc14+hiUTDHgVU+BRkA1B3pMD1UbjRqUl3mV
PVCiGDXHOyj3xKlSEBQ8A7gF/j+7pE4lpYGzdyz/adDStS2K8dR+vO48BiTPUEaTGlrYaCOqyNxa
vjbVCfMZ9BxE5xiFj3ta6ymsvKVt0NJNH2NxPm+w5Eiu1cSJmCcDvpATjimY2oghAvCiz9YopVAQ
y89PtLO7Hb+rS1KqTVlfOpafiR1No+rygp8NT1jRmyVgIFuLzJGAkIgvejFH8AD8vUNz8DKJSWAn
v67C3kdG3z5UHx0rnJpOkLm5RqLIZrlOrX6u84m45UXfC+Xl5eKfJrk3vuMc7bOMDfnH6FJSEXvi
kwf3RqIbfdhWeNWfxqW7CHy5u2PQCG4Nbu+HyFpI09WXIac2brQKczPeKQkL3l+JiCl/LwUUIpVH
ouEuFS+ZL0wcJGNJsptRoEkGdqiGuhI84n6BkMSBIN5AafdFsQH7DTbLd2YFZhuYgkEzEYKPrb+N
bEDWZXwTNd9kzmVt9x7MzhcQC3PbBQHMSXObuPzZGTwr6xmPViQ0AutDutThn7Mlo6WJWg9GSNpT
7d+xPZILptee7TumMzXaOYO3WyZjVmDZvP4S7/+glntIPDcn1L9CFXRWVGn8MgnGSgQ02qeenAfJ
VwDWKPaHtiy81wq9WEnBKxvnjDIUHMm3SYFHipQXsMTGEtFXYip1cEQuHB3maMsl+QPi7OFi+euh
01ds89NiZhqsO6XcxqDZshdRB6bpfaqACvUXZxA7qspocRgG+C1MMfS4uhbqPfybLm9HIq0=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'evaluacion_curva_puntos_obstaculos.py', "exec"), globals())
