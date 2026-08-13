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
OrPPNq8KqBhA0bg/sjlFPYRelhRVbZtjtoI/GJet4yVpeToAFl4SCcV1Jmkf2nDUvi1DGecondnu
5F6G/cImkzwEXQnzvfM7SXHx08GslwRcEnfmItvL23w3CLnpES67pZLEVpHuVmh+QQ6pDjc4JCUb
M5iFGrvbr1zFZEgi9e5Jb2veXxyKmDCYM5/p33DR5bnE8eOjdkD/A8BGMnHJLB00Z3y3tQ/XMl5U
N+w8fOS3xy1Uoeo6FxQhQk6B3gdMD9UDa7uHtDvrQgIkfPqL9UMaZndGcAZ61f0AJ0cyauFZ7JDK
kedwShyOGhhYunfmcNkDNrG7RIUvfh01LnB+ErdhCGOeeTItC/w8kHAP9QeWbGsL1zTeV9gPlXiX
u9fxRiWIHVRXud9uOnvoKpUgqfAbDtPzEDqL63YsACKvalmljPj1Q5GlZDJ/qIcsHypzZ+E99m3M
kaFBIU3pJc8Pv24ch/aTfYAS91kMY94DeuHEH8BfQllEEytHXzgkhKgY/mkX9ZWhoOegulOSS65k
RuGh5z4u7BmRx3i31jq5zJFKWkOgOU0uZ5Gbe/FILg+ahsx37JaHxLx85UqgnXgaPPPkceod/Xrs
HgIM8dPE/Cl/Qxc0Xax3rQrMG9UnCyuQqo/qLXnBQ7uItKCg4vjTLo9ql76AHv/dCPAhOtBMjDg2
JV5NAkNQHhHx3bxc9D67gD7xGGeqxEW6+wtn4q4lGBoqTflTeI7uEymccGVVnxpWlPlmcFERHONk
Zw7jyVABeGWbJppigVfZ4flObhHbVnjvxUVABYnIcUyZiIstpVeuhOi4MS9ACeQeiqefi/U3gD6k
G7+3RImkHZ3r/QOm/rvyKczvxy/yh/vAxE8hpMkXyzH/zi36aKYS4rl7lieH0cXCwHge/iZfNwxZ
D3fk3snMsEsGqGs6iUqqhLTwTpakJw2dXvtymQk3qEeAFl+E/vXdCr3MWdPemQCJytxU08JET8a6
4dT2rfdZ49FwGFXzrJDGChGoSVdpslg9CUSQQEwatsT5gsPMmyO3IPPc2qOEm0xjpD++kvAA/V2n
Ay35nkC1lZdWGK0me5OaaOxQT4J78mzWmIjqXNPzKd8agLbvf7Vs6QVGjeAjecHiqVFZ3LPDoJ+V
a1pL2D/8WeS64k9sH11Ljv71AHQ9GYv1l6v2MqO0DnXiJF7M9LY4qzmoXlkCSUsRAujfdQZk9qIj
ZymxQxIa/HKfeK7CwAAWa2aRVGbuebrbUZebi3GAb7t0wtSJB+YGnhuo+BdOA6MRlGVA8LL8GkXI
OOo0yWQv+s8qiOaCQTaQukmpvRnICAbUDFnQcc9I0WhYbrAPh5dP1hvzt/VeZYyS59Y2e1RDwbuA
lJNKOu0zk0gI39ehdryWlzN8KMoLLrKA70nR26Wf+9kFvR5BD8Hs5aSanr2LI65S4DFqzPH+X0HB
+yWvY8jGKnls/x9WjSeDdsvYaCY5jI5RYOtuEAE5l6wfck3ArBS86Cgew8diDjgNtysEJA8w8juI
Lq+OuDRckM0IlVHuxJF3Y6UVJ/FlCQ1wnZUhkVq9NkvbNMw8tV+E+U6VpCT8dmjtYKvwg4nf09Nd
sViT0tcpyse7dneZiqq/pXQLdy7U7SM18Z+eAYJZwfP1o4/egvQyTGbpQWOp+NXjaa3dMhOQ7Glv
RoyMqWPaH2sYtRmFv8kZFo8y5gAnyZbD5xdmS+lfGy495v9RRk6BwapMSee2MrDh724o2/kfsAlr
otUqI7em103HVhZpMAGZg7+wccafXCfcYcr5uypoubzGTQoC4UdmiK5VBA2a76t90mCc+g==
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'people.py', "exec"), globals())
