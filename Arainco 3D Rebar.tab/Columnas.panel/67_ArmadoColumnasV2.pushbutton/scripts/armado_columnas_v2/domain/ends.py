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
OrO3OL8WaBlG0ZxFfpwXqX4fW1Uo0CfCysB78iO7WppesP002wTjLcYnCWryqdNKwVySqF/SzoXZ
kxRr0WfoVflodmeSFDvksbvBYthc1NS1UaecfInxf15PT1x3pRpACMMIH45p3dqukpGIkJwLRW88
PPCuGVKpa9wxHSSAL/iKTOvl1U8qyHh6U0cLWSo6rtqsjcPtcnSuvNAjiqKT0ELNjxuBONWqhhY6
Ohf35mcE4ARw7cBb5614aKQ9OD0L0SDvB97EmP/0j2p++lxl/PNJIE0JnQQMr1pPfqe3PlluMTAw
uzocXCEuwKNAnrr3MNNaCizmzTn9yNM5BY7CbzqcgAOAZ18rYSQrYHP4zDD60zoQ73FMlGmIvDuO
iMjHAlbIfvJtpAmOTp+NuY5li/jLQ9/R4QpJd2uW/P8RiMhZQCPXchsCywMpPyk8RFBMlwTEuyp1
LluJlSebktvN+FkTCReHuZfE0tEVaZTc7GDiDjVgZIcmTWsKIxc/pOrNs4IVHvWLXy+jXd2MPtkN
+N22ftGpViBYYG0QUUpHPDO8Sv3sLNoNVSEeGZ/Tu1mVp+PUq6uMrIukkIuC13Qnanx2uPeCLCzw
lLSmDGbuVhB9UHqaYGPM3MPj+kSTK5/d5TdVhwH28st4fuDw3HzUxuhX/wQvsFAzC7rDy3L5aA4X
x1Y19p6Zl5hwm87TwDMNsTzHrvdHb3cBkENnaaeK8bFJLn2OGrV5cfbHkw4laFb6/hVS7k0ziw+1
8rXlF8pzEmrn5Y/Muuiu/oeri8RNxuMrhAQtrHx4wE3uDwgYhQdHClwmSMcG+GGg1hW3ososRBby
hUrsC8MZj8Fmmx+XeINE5W+c5uRjDKJO4WZxEAES0sZdjmZ0JzItn/qA3u+9MzzXScP+0bp/gtUW
0eFHswNqDT35r1wWwFyh+qJWvV1hsU8aOebxd/jr49w40Wv+djaV1mv4eOkvYy1ruaJVP5IE8UM4
5NFllXhUZXTeByHq8DcC5HXhMRwyqe61h7aO1KMS9mgfOJcMSxMb+jcnYYgV3SE3EIQW4cVH7JCk
uzj7ZMqsDse4AMvsu1YaGODfTWFNRXY7jYiLSmlS7QChdwws1aSQDdZHTsfh/4hBK4uopLjI7R8U
MPDiO0lr44+lpz0LRk4/5+X2gQiiNFjWpyId1gDNIA7uItHuYFMFTvQ2UeoKIStYKmhhHUw1Y7pm
qGaN0P65YKqrVCQsnKx6zKAPJgZbkmf0QURlQ62ID59R2sgFIPeiaEqaAWvwELffo1Bju2H55DPa
CTwER7PcBwzzMeaV1T+YlazB5eOkNkfsUFptyyYc6ejnVjlkKOO+PzCLds5yW0s7hxSBsh5wk8VC
3PnSd8Z55cr8A+k0Sle9/7wAdtCaJjdy0GuyPxq6SJVnuYs0kUZQX667OdOMeqAbfblP7jARbY5U
6CKmtW9p2u5kvez1NWrwJmLLDgRAbgSoEUpzPvH5mGnXFw988Rtau6l+JWvMejsG6QGv2ntEnHdu
uNQHUeEiGUt6eBtI75jiglPxw5Up7BMvzUaew4B1M7+YqMuiJRZCkWXn50TfHOgQecYspC1jSCx6
FalEmwXM7qTuia1QmXElQQZKcPcz86VbjItIYIB1YEzJQP5pXk3zkL3FhgviK3A9Vjep1CkEZcph
qVMvCNREeBNWBpHJlFm7/GIxhCIrZEV6C5xWXP4c8XivQ74TgD4aYM5+3WsaDOECBMJsBiD4DlK1
ehsAgkGK4Ybxy4PWApVduVnfEciN+mWSUV9vgzmqcERxUse5xiKI8C3o4a2v9LMe2i4wEx1l89HJ
ZW1B149fXi6V6/ntviovfi+jq2BAFnamAxZirQJf7lpPwPwIUXTa+NzzdLqDxHZ3+qok3XaSEiYx
pKNTnmkTtUK68GY/ylLMuhL78tqIQ9D9ydm9x9a84WdOO4GDGToAmFWqTdaD1uWZJgoT3MtcVuWN
v1w8P8RZhxq75auMagRgHIHvPW9EZoSm0cOWZw50+hzRetJTe7FFHQR8h331YJOKlJ+TVXaa0tAe
qGz46jGdXwsUX1LAT9AOhTcQuiotJbJw8mWwkaOw7dWKEcmAtPoG3g63Bhxl9VtFF6ruBdOGTqkG
VWRRdVzpYSSE8AmAFsUJYHnlcmsd1lsqPDzSaDAi48OMWAdY90Ys1lUTex64hJ/Cv0lNTQCq4ALW
zmLP3iRXnpkfKLtdXgvN8QzjCKD/pvDjaC+UElW8c3htevqVujWjonTBddoiJr7M9bPKybu2uXrZ
vfxq0WCPszDFPvKYjodx5CjhWcxGzHNfioxQAzORQhkVy5bcgKmND52RhMCZnygkGeiFwWIPAc0M
oqAjsv71dYGXoCljGWPrK/+jpzyYgthfslpNG/DCcQzs0BNnp7FnbNgCtmjuqj8O//pOzwrBMNP/
DAAHLhG+xUw3hdUIaRFAaDjzlg74XY6uXj2HDrjxLfab1u6+BDDi7EGUnEP0/CAoaqoNHBAoa3lj
PDmCZL6jkP07HpvSmZLBgyGpPk5PtvbcerggMyq3HiTJHkG8gqjt8Yh5xNZ2eJoeh+4yXwFrbcQb
D0I8hTZygUt4nmwVIJ1ZQT0+E/nhjR+DXIrM8RKmPgSngAagM5z3FDngqxTXd2BvsoAdcenQwAs6
QMWBa5ET6SV+g5SwsqYLFFacK0Iq3PsuddtQTE9gr+RQRpn49mB4aVHZKFvXREftO/CHxPO+0E9S
Rm8etVos6jnuvtvVWGmY+9jVACv8C5Vehp7eqb45Jls32OyLYKU9hFSrN8eqSNyQbKQFsoAZHjel
jYrvRweKDBGhb2WkLIQg7Ar2XUZKxAgPJF4n1nZk8dPIzjYTAn7oG8qe8wsVM7d7ELTHfzK4rWsG
+LbRobB1OqyXjUnYccOlsMEEennleGVmXeIxW8B5sIFbuAx74/o3JhXNuuvmkZH3Ij1fWa2ztfwS
ZDAb8rUZobWuxZfedP5dQt/+mns3LXAfKkhAT1bjfwR0BPlp9SPNPxKYhCxoZOF/KgOBNjg+GAbf
kz+kaQNIhuVsexKvY3NZdFaOB8zoWPrXLNdBmvESRalve56RCLeBlRJYlJfHkkcnkEGYGSrBEEDP
YuvIYb7a/j/E/3zBooZ0jP+N6mEx31B9FsckpLgWkeNu5LDUip+iNkTtIWOKbC6D6pif3cN3vfRd
IR46JQQTqTGfJoT3vvTJfSAnKQ8A6P1V/Abc7Fak2sYN6wdcJqd6149gltrF3gvd+RjQ
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
exec(compile(_SRC, 'ends.py', "exec"), globals())
