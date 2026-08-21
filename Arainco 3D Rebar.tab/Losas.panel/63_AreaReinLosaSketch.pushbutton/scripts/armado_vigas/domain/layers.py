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
OrPPNi8KqGhEEYhFiJp5urZEYmEjVCdmdbNAY0ZgbAoC5LpBpIYhc8wzZTAc2GF3jA896HrU/For
O7u2McAkpefJcmDpsJ/i41zcn5kfkoefDRc3f5aYetU5F5vlijKSk4HuDfrk0cw/udvUYFpYL9C1
UHupFCC2kfrEpmo31RntwmgpL/bfZHE6I96EdP4gMjtYb4f4czXDPGoeatZKbP4nNT7M3ctkC/Mw
SFOCkCApFBLePCnVrK/qRrFCgQpjzSob7dGMerK2OUNPN3oIYJjlcNL0Q7+z3wYqBdcK9sPkz6o1
K0H+PYS7y6j5RfXjcmYmikIvXc2uKs4QMnPR/o2m7K5cBX3f3DDZ939SURGqXbqkuSY4+zu+4Mry
orA3BQKTQluzKRahLA3ijY/lQPsA2ctwzs/dWZnN2TrUd4eTylKrGmp2l/MaFSvBFd5QHXXBBawW
v81DfY8cVvgSyJbdEwp8e25varWF9Vd3z1rN24MYQim4e0j5wZJyzkYxtNq9MOI/Kc4uVFWS3Ikp
E+woum8SiIuwohz7OHa9+9bGTWA1DEuAmGFo5ujIvWeGiLPt3z0SRmY7fOPESeBtPqStLWRbU5oD
/bw6udx2lInco8AolEt9EG8+r4xeyfaQIuSYePiHEI4H4sJeR1jFpCWrlgeOW08SriMALrlKA64w
6gEq++kjAnxAC3Z9HjOIYmjwvm4+/Kets+nWIo7WmuTDS7KNxkbryMDrxmtCMM3ZG8ZX3ESR8QmD
In4lwe9xAi4CLMW+fWXtM/3TCDuol10TSUSG15HXKlkEQkT8kfbe3kbfoHY1xDmLC0y2f3PTDjr1
bODVNZkmZ3W0eNjWSIyR54fbTobRmfxfs+Rs9wKpjRPWUzjm9kvh6yeN3Da8lLIdDSoAwLW7nzzY
xYxmHdDlI2fMIwSE4y6vNY97dvfjTwYskQ2h2xUDbKvNvlaBwHnIyNmWMwQKzxb0GfbFkayTfR6F
MiT7z+4MY6CGQyVvrRWRfsPeyOa0UrH1DqQdO2WSzJH7MG2xhnICht7vbGh7wz0LndKT7wOyP/Ou
65BoFwg3+SK0WYNtk1QcaWH5Ib1t4FC2cylfdZyy+0hNUEMfHlQv3foeeh+f/y2GzRAbLGU08ubp
ggHhOKuOF7i+Z7+Ox/DXwgWAc2KAmX7PzvsSYcka1zhFvjYH4rTJYFaNkMBEd2XKqUh2fnkq0/6K
aJVF5V5WPobUbMG/cO+Om+xRgUFb+kJMlHXATpRwf+2i+laU3tlIlGAtcIdskDVm8kfc2ZseLwoc
sEWvjxuF1rE5fXy4aEHV2WqB48EeZE/HgKzZSGmVP0gvF37wdR5ABAwSdbNYWu+3Ui8H/+ucTnqw
UvNe3choJlC9WI5COyY1hjHBZXjOzKl91UwrZc2JEYYtFMdFGKwtC/1mXVXPOnc3KIycd8Iq0EV5
FgG/4QO+Oo1NnapRL8PLpKf8XSP4zuXmuujCIVIxPBN2lYGkja0AsYFOt8g/LTJHW8l6yxU8hCeP
T0ypNivwf5HglbjTovQvQVlcas0+cNzhjtOEURKfCl20Z64LzsKX348DMUq2s0/CflYLltKk9I+a
Zxf0FH0DADxqkHoc63ewfG8bkU3ncGAt23ciYIJP7dXNCjdRFXIfVaA/+LwmpIDz0EuIXRcUOTrU
/s8YOXzFHZUUOVuI0PnUAcutoHljXQXUPkv+HDnyHau9hP2kuyyvrH0LcEYU
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'layers.py', "exec"), globals())
