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
OrPXND8LVxtE0pjDHrGgYi1F+uhEwSFz4GnI1ioG/oxivYi5p+eyQ6OVTICeslUehgmjfcW/BZQV
E9umZM+9fZmMncuF+rix5+RAQrbgIPaZBVkhGYRmrjUTDgi7M6qagrf6krh+BQ2td6/Ze6Vobezu
1oSK62VcnRfxa/jtt4HiCWJMP6UcDvYnbmvoWfJctFZTzN21dGZHrONcr6nrg9zfpp3CIf/OdAw0
b7CFkih0pKRjMfXOH5CGmyRIg7fBz2/gvTjiAnnDeFR5tGYXMJHgQc6Nj2vB4toKRXPkMJYOqSNC
cMYkeyToOHntQf8teo7NFWSzsw2YdHKIwnXbgymP+crYLhvAKcHWXHpY2IJqhUEvuDPqs46yYVGU
hmVa0GFuch3Db2AI5jMSSFX7SMSDvuN9UTA5rf19l/1PK3roWGRui3Hv5Vgca3CX928X2NtGpeXO
JOEpC3Ludz6WJEBWqmdN6dcelxpVCmTm78QvUWuWLsBBSRUphC6Zyfq+IYj6p76WRwb8Um8b5nh4
PXNQJL6wYzEMgsRquNpmziupz+OWZeFFx1NsUOJiBrt2Gh11WeR+9Fq+nKgYObmeqUzb8VtXN4Gl
6N6twGdmL5sYT3ZG4ciSeZg3v0MkPdgfmSALe3y0CEjFvrgjiQResDXrKZ1W4vF1KjKGkFRFDdpE
d5CY8oJb6eVJu0nuTu/24S9qtWMbrWam9UeJU/kSFE4ek58019s467weNOBuSCFTcwvMuTN2aD5Y
olB/kyJxOApsDJCKIEtpbhRoXst5EuTVc5Rkn7bAbSZGTnQl8/mR87iOXjo4OVtI+5XiMtSb+NIs
Bhd/Yvs9wYtDG4PtKbJHgr6pwAfVrrTGFeOAmZd7IXGiVnBWlwYx8c6EOiliNz0tkg1sASrLDpD3
MmNiBe8eBXx/7QYx39fHqu30HxlIZLAVBkXdp7XaObRzBN1rK6Mgm+D8S+9bqeqOrfEw0M8f+lA8
LYjC+fGhV0jl9b9dumA/DttGuanjjct7xeLXvR5qzUHNOIE7cZXtVZyk3G0dcm6+4e5egQcKpiP4
PwGn9NqQYRLYlUaxEWPPFpmGFwW+OysaXl1GPDIZDP1Iq5e65P1Sq2xfs8aWlDGI5FD6egK+Gtdy
ULdZO14h9W5BqI4lh3uxIPYuMCBRYBmedcNFKi08g7+78YRwsbQagei3evRdN05am6zH7v368ujo
SuyqiDkWmr3A
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'revit_version.py', "exec"), globals())
