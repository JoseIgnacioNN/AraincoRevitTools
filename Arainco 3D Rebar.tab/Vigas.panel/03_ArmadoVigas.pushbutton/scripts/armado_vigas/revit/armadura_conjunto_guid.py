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
OrPfOTkLkBhE0YQ7nr43AXlPNElk0TP5HKCpDTUqr8fyMbLVRhLKZcdBJsnnrunxH8L+1btKpUIk
nkQ2IKO2v4rwm61FK5lFJp6iaec0vKuavFsEEYTOPuhi4iqaxJvLe6B7FucGSgulCRzBIGRdeylB
65o2fVAwKsCwQ6e/XHvCqnsRJ/c/L5memEQ7IT79uJ1yty9hznYpjGjdTE0pz1Of+q4NuzsAZ8Vn
mdxvjeTDMB/EZyd7EQGOSea9KXKIx0QvSnDijn/UzhV5y6ixVwUoQbhsQfkJL/spcT+i0xEbKH+L
crhZ1LHAbHSrH30jf2VLeJlHttUNv0geqGlNFE1NNxb3TWeKfQ5o2OvA0Q2mzuWqauZR6Z0Grzp6
zyFD347N0XFhUa0sHmSb4gyN3ToCBBux8OJaifvgtCawiSmG11LwKYNY3MYe7JCXDzAHXEoJwKMe
4qBwhCL3VZT5E+F3VULwe1QVX0G2ewJfAAzTwSf6xukOwUt5SchDoKLKBuMvpLFguPYNc7W87SH2
EXp7RUAzo7DZxSsvn/7vZxAwVgL7xR3FE7Dl0/0s8gFDDtwDC13i3dVLBC1325U8Lt0KF8cKPenk
IpyaCvTZzcPUeM9jSh9eyMEdBaH7MFkspdk52/Xxdfg6VQqON15vXtZBkB7/Uj5YM9nUIsuNNYAJ
XDHu0morUzbQ2QYJs+c60hibcWNee/WpvXl0Hcc9+7fRNdYglgtkuvfdwYm1PuDzlNuS8vTt0T3T
MUSNHJqMu9tK+XJ+/+kv6KnHA1ZSnr2vgI24opFyT8+0Kq5QJn6BLXu0EG92UQwtnFBYCh78EbDj
5+dI2zxIW6n0WEMo1Lvv10lEPBnGSsJCpxPz145VAOEqbfaplNmjndT+iDsIow1cj0PyHuv99Wq2
zYbY6uUijzHqeFECrjo37gMzDfkdCsWEZ/t9V2zlCjmpHP6Nvm4ZjI0uop8Z/4c7rUnR4K3QgCX4
alHuNW/tX4O3PhMgzftvBl5cGux/AhPI9pwMtnTOgzN4GW9MZlluhIqukoTincKYn1mnoVzYGiHr
pkCQZi2FK+sCF7jFMQLjoyGK5Ai5tn4Qwuzq8SzD79UUdMisNTPpHhZuyy4IVFMMQIMV2IxxBJ7h
9loYjG8zIITpsFIuMxVlo+dXFw1yYrWwU/ix3XFl8KmU0EQulNSj5ft/p79VvAElulmN8a2x7taG
7fDJEZX+U+8NQoQeG51XZQnF6FgcNrHfJo9pFRX/AXVKN8b/20i1gp0SqNv00pn3u6V5g0/hMhqI
3e1Og3pPTalgw62csHRo7+G2QWpNZU9FGBY5n0Q9UIqyH5hJjDTZJ9qUjU5OF/cfVyoZsxNoEOwA
NVQEy4qVozrwGpMb5nSZGJKKCpKFCm6F9QHUEqctDI417u4pflq6m5ko+kZVp9N5q9kA8/qiSCXO
2xZOapYuuRuT6K7WFMpDmxRBFZZsvy61MUWcIG7BM6ddWZbBF6SjWh61SxCVXjKYiCQW7wpeyx2O
/1wBWO73J3A9/nDemTvNyFqC9gskJ/s3Vt+1J7QreO+jSx/IyVorgZn0H36EtiMBY6a4lpAli7N4
iAVkPKfDwCSm3xla5dAYfcBHWN85ux8EnVrFUHHBh2VxtfV/sQ3Jkgkm+9vEXN3ZD0qwReq/Qxz5
Y3YkZYGo9VnBdVP3iZoUTO2+T7QFOoNLmI0h7/3J1OXgGcw6EIrH/VLFif+nzs30KKKN2sJKLHbx
XI/oTkwtsEcnGl5Xc6tUBcYBLohDiYUWN0BNODme8RifBSjlsUqffUUecnRpdBYFiOYJ96CNIGb/
nnQ7JyJy
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'armadura_conjunto_guid.py', "exec"), globals())
