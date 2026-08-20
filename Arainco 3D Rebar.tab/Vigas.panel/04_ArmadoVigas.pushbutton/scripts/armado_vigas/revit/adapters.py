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
OrPXeakKqBigMjAthn+EiczAdXYj9B8TfBNa/1pr5ESB4zjBV76VDMVtJomsct2skpTYHK5L0hzT
W7COK5xS4vUHBClW+oirs81IhMb7UfuUuv2nTSHZmrA/FH3qT057M8gKq3f1m7yIPeqS161d+5ay
fahFNXvh2JRQ5Y1X33sDi2KwzoQxj8CxPSUXW3AY9EqNQwXkcKiZ4kff4q5hl24jQvJ4V1YFrp6H
gcZ6TPb4oDFY3WvTqrSBjJubawrENF3k34vFceq3pvRB/plduCUuzcE37y3xWGXtSso5bHoEBt0A
rQ1HBEEg5tCnXFZVHXMmk9zR6vnTRT8Lr+u1Md+L5gd9N9BCFmmH4AkI8/aHfJPwRv08TCaTM+yf
xmwsYaMuzrFGyZbu4cD4BkOCeKpgSlIC/g1z7lunmcg0hZ2O8H7L7SMUlZwRqTWIAT2gSgBhFYbr
Ff9qNGft+uJn+AonWkfDEiECnDOZPrDEiAPk2iSuIylSlj5xYRHSerbOv26JU5VmImsm/9jqoVCE
kecLnV30dK54rzwUzartfJ09HfT3KgbWahsccsxeMusit56QJMtMJWkmS3mK3iMHZdGf+Vej3ucD
uFop63DCcTo9Bw+MdSdbV1CLScQwKx19GHfYWzTsOersXX/od1o1vjCQjHsrWhXc8DDc1i6uVG1W
7d9c4D2irBQ5XmFf0e20yJo+zWSjqHYI2WMs9cydumQBIMTttS+JMYxD9oiB5q5XPy8sL4+bm3lB
xrfUjfcY1nQNXPMzhOgujdetB+uRPgQHspzlzvGwmfrPBXeJNovOjtAQB9D4KHy+Sq/HgwChzmaN
P9EQQS1RPMLF/KELfK/9QyH3BMiHCVzovd1RjhY193VBXvqmICCWZaAmmMmsNcmoYdeUZNy8PxJB
rXQFqijeXQBRuks1XLne4NBsubadL1WwIZAfNlmURMaLz1iyKsvDTaX6sKSu61dRFSJV+bu18ioy
Ub5bw0gEcuHX7VDZNslPaRP3TqGetTNdsc0D3Omge2gxLgmkyFIMWGohSImzbbQWtf1vMwAPOwYV
FzksUUIg9PEWgEyrW+HDjYR2guaTh4mrip4dHqMNKYY7dh9qcAlMkK0J719Baami6d0cpe6MdzA3
zG272Ndn1Ts2Vc8+dKPrKzSZs4QkbA/0Db3dtNrvZTobSQ4SMFZsJUxxv+d5BleBhs3/ubavcaUq
MtEk1m6o0uOLj9kFnR9YyUHmH+Ar+JQOVwVh1R3eoWCqAHJv1BG6MmsBtz/udY8ytn5ih+KAQG+E
SjJA+1lnH2w4TkSsRYPj1TDUU4G19+A7Ogm3GH3mcqffAJL2Ptv9sQZmrpZzrdnsY+0sTo/dicgw
wzE56bqThmJFBLoplwEyOSiH4hNVCzZ261iwg4hNpYN2QTVv0VT1xgCLV4ilSWv55an7kw3k3zcq
dMu0rxC7BhhCRZCdb2nqcwc8O+4WOKu2rmL5+PdR6AtyK+VkxFBaDQOLUFwrwyBcQp+8sYlkuXgn
fyyjOiekvkFBPIzILQ/ZPnWaw45lLvzJ5R5h+m7tyawPgCwF8bC13zNNsa+xaYIgIzLF6GdMb4Ke
y5xUZzOctHjW/EIstz7RrHvw0vJmaYl/AWkN4G5njeqxsE6tMBQ+yy6ZEie/ATf+G+Kb1O651EgL
GltTdKmTiwPPyniAmzst3Mo9BMo5RP1nMBNm+r8LXk5o/Ybdh3qid0LtAvdDRVE9Ict8HDChVrS7
yf/XIIcMyGY5mcg8XfSuAqo+mchvsw62wRM01sT3vZ55fzLxh4PaJc9XJxzUkOnJTrXqjy6NQ0Sf
IsMKwPgaz1AGo2dDPv6Lq4w39OryIS38naiWnW/vu+EpfNCBOX6k4X0WhOACwkTDAeTsYtj+3Wxs
S3ajPycV/J761ZO4pDrC9cWwrDihVv8JvLx4Em4shbqe1plw2l2eQKwzv9gDTccU0ewCTYbrF4Ja
cdtvZlSXDokTF0S+ESB+3uLUFR7j7QLl7q9F4jJ7rHYnf36XCTIVuA5BY7ld0q5T7rlQd/nvmW6W
x/Ji4rPnmAJLEFVrY8xeGeSJ+xOPyjNXoOKGe4cRtzyChHkNkThvAcTFS/DaGjQzitQo1uMPckRj
tTdC+gOS5jUtHtWIXxcTiQlgb8xn/DTMgMQeisyCvWFHAg4EJwqjitDVwnChwKVuvottLhlWHFS5
8BQ6HKdU2frzlaWgtWhdCYVyd/Zaf/1EIn3sW+B6Fo7wLXZZ9W1v6gMpWpVAvYEmo4uLS9X03dcY
zimfBcgrkOnUWf8R4SSrbfkxjI/ueP53SQUrjEfPnrker1g/N8y1kxA7cTWcI+6Pj9i8Rtul92HV
bJaUlGPe8lreq6FSPNRZQWHv6lljjVos4jlqH1EPRRIjtFhRZS6PXM6mlB06zogjRE2pOoRjF7Ar
gsesUYl+JmJkrRqM6pwPQrRN5ACy5u0diAN+vgQRqINvFbOXEMrISS3GFi43ZJdj3rkesFHzHoMF
eNjC3Z+hZE02gHihDGLSX/gf5aY=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'adapters.py', "exec"), globals())
