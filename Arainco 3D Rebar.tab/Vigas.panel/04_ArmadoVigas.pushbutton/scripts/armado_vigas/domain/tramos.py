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
OrPXeKkKr5ahEtHuewAy3aOt/oFSu8ewQ4EP3mj4P4A0uHhiMYI7upniQW+/okhaaNpsdL56miC+
zxa9QcdPFWf4wK0K5pUP3acc092cWR0QXgvARJOKvuMgW1ZU7eW4pRCtRL9uMkwSudtAB+m85ns7
KqsFf3tKrm0kpk+/EyDz6xbr4NQfbYfjal/uzj791uZHT1VvZIJiy8oNeDlv1LtFUBNTDV0rtXbP
JVypjICFsZCddist7UJ9oKTc9EIJWTNwd9x9vmyVTabnUmJ7wFqyED0aj2udjJbyhIHsnX6d7p32
Fi2Vna30pojU14X4++y4hYFqTT0+8T4QkyWqyVFrBUudsbpnqPVeA3QvodgDxcvUudo8gLabt9qx
5wys3vXMvhWjFH8muRHz6qGbVeXB76mel5DUV5cNkpjWmnzJobS/zC8faPMAILqR48wMHmTclF3v
7P5H5RyQ4LCzltiYUZNMgSe+LznddAIajqLonlAF4jNf9IEvOAJm5T87REUwR3pc9vEB6mcKgIBS
H6tU3KtCE3pygXiMaY+FybLH+94FwE4AW58NEl8Lz6v4AWT8V2KlCSK3x1JJE4GElyZ3fufuiaPe
c3Gklx6aZtFF0lweKTz7pEEhVmV6gtAd/QXiZuZJCrzgz70rNGYxfu5VFWB7L/UG9EPNWUtV7qPq
Q93l+1TRC1kW8dy8R5edP7RFBAK3A6Qeu2InbuehDcd181AB5CHqK5I44qwC4Dtm5p3wMkzV8y/R
U9AIPTPSynVmPSd/tsq075fbmZPeAAutSqvevXmTOwA2k/CK4uFW3d4Bh9Aj6lw/ePe6rvGNCeyY
Ej/mUh0zkHl8rjFlPB43yd5CdFDC+6QlH0HE+S+LugYoGlC5zN70depPdq23ZW23yE9nls0XcRRt
151gor8yESs8P7w7tAbyvdYIpSTqbqeTo3UzuEJ4EghsWVNO7qjMbEkeMyN+8epUSeiea7rvjvA0
absxwgvfVLdw9uVMVxHjAFN3HvxMyCZysLHdMy5q4xdwa9zOAN7M/4kU37amCIaeM/ILDI95WNi0
Kah6JKQ0TukYaoW+GMeMWH2NyTQWKU4TA5PQvNzpZypBrmhou0UihNmiwTtd66dFFg/dtEV17Hly
aRhVHpQLbsmnLPZMszaeXWL5u+MCGHz9Euc+dyLL6zz9UMfeT7NMQRRcEfAmPqSWpItTWFg2SAJe
IF+aArqhcZMBhnu2KUtAIMCeN+EV/zmcBI0jfYh8fdzITYRvASRgqJBX0zdzR4MhqquyeNIf/9UV
KXtgPypDUDMpFsRB9VVXX/e2s9ZuGvotIO/H8scIq89qWDY0AHLUsiK1dF9VtJAYObbnatnsGdT3
mT82PAmfrOH2xqngrFIpPUBIHKl2sZTGaHbmDi3YP/H7KISBdhhpMwVXM9BrsVwusioVJK9DVW0P
NIEdxK9NcNtGzq8Uenj5uQh1nzLav/WBOX4BBo2JH9WsHcfhdEdg6QS/qFCEIaNwFOIW2GP4YqS7
PLpJX+cXAYMyqIP9r000q3eU2lz5eRifaLfaqVLe5LSazwkBH5T0gnip6QOfCRQRxnvVeRMPKqZn
iiTrOsTh/HoWSYreJPsiM67AFXQDrSYzY8cj8OpFBC2Hamvnubden9VmYGGh/DXmyW0MgMDy0yhw
BegZ7mhWF1VAYFWSr8oLuRubp1ABrk/VHgs6XylMqwvz9CJOt/GvCWu5b2p0zr0OniwXk+/D+JXY
2Q3Ubo8KZYOJKff4V4y3fjzrfvu0JO0qYP+c0SXNIPaJTKqrtHcefcIiCbNoUee02/mrt4ggJ+0V
DKitWTB0aOogvPQkY5WMPYOU9x6x2B3GPU+gVhvY3VCF3c9AmqWhmC6D3jpV8t4+Gqo22uBLbGCF
5/R6wn39iXwVx+5fm550Z7IW9E7i/M1ySq6CttyJos916aDCm8RJpA/MsN/QiEziKJWv065/8PVY
lN3O6U053U8xnnXhnMVmXCg4ILaoiTHA0vP2W3HDnkyQlRlRLeTwQEXNVQ+rAnWcIcOi5cbPQTei
MwlctQmsCU+rfeoURrh4kZe1iWrGGU4XEVRXgFfiLnxSbsZiQ39KAtybPis9C4UTDIzf/VgJ2XQq
AVa06BiF/Efk+dZBmHhs5AQ4OXkb4b2paiUsX917ILInKgqCViyQlmAAOy9OXoo6LqL0MZpbU+sM
lv71lE7OYIWKQLRIq7fIIOZ68poTEtnEDvdPspgksgQy3gi2P7bf6aTTPs61pJsoJNcQcYmQLATo
z6oFlBaPG1Rq8FXSO8Ae/xgCKX5JRR6XieuNASewg1Af0n2K8jaTwwh27VdI0ANmg9h6GA8E/9iG
02wKGglP5NSpiWHg+2vyOyTZgrke3ytN4489b9a3tAFR2Xfv1ZFIH9CxfJ+61nga5yBS9Sdyd69d
hS3D52tfQOskq/SA+9SrDKB245VSXLzbtVi1AlqQyOUzlFTTmfgJM5lsnSvVt3qiRq8euU51biZ4
rQHP2WLXK4dkVY3vdgiJ0r+OiALLKHtg3aqK4NagoAofsUXHiyF+CHWqZ1b4kTQO0T8i02LmtLMB
Mu+yQ1WMTRCvMZQGuMHILTe3NaNEVj8+qK4Npy3lLFlcwu66BLO/xQK3j6LHoZ3YoVWy2ax1pCjM
vSiIjQai9M0oEaAWWLp+RtacGh0+7jsx0JCAxEQ1DdHvtS4gudWJSc9W2xsJtWI8+b/cPVYMcZM8
KHRLI7FM4Y6ojLlH/Usq/Ir4/8+K8rFfqWeAjr2VI06TftnbPQA8zaF3vKYxmxZyVf8th6h2jr4I
NkO95FyXoUeFHZ+PElCKtG6pSXII1tfrnZfe/IUr4p+VqTnXkjK5JKWpKPurx2tSbdiUBLnv/7Su
jvL7orQwgAR5POgNjCQSGuH49xEsbU2FqgcO8GwkcbBva597zXI/7A==
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'tramos.py', "exec"), globals())
