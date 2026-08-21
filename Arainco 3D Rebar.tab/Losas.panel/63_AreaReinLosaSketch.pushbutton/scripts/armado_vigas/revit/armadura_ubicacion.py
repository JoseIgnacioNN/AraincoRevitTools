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
OrO3Ob8KqGhE0ZxFOLpTC0cdPizjdGx3NjJtt+QPsd9jYewqAMFdZLp91MhEk0L/NzHvDfY+HF52
vFgR/RI/VtVWIDPWyFdGRa7ZqNuevaEY3YLsGDV0XH/UIavigfhh00Z1H6BPvUZgHUHkUqjivTgk
ygDsbSRagGVK5xw6ZBjLK+vFVuKEYhxb/kWnVSjccX/xgXrVU1Mq8U2XREpnuNqJOxH2MOJ/nS9N
ZGDPdm6e2kCKOEM4IrHiVE53/xRwoncEmwqGCG9kuWBDtNe9gJtsxcx4SwF3oBIdF348M1IyUzvI
8tvRgyWlvnVnSUUFne2I9izUPPuhBQh6UAQOM6PF8+GMk+lvVC+giYHElAc6im73htl99tLgSyp2
Af2N0n0ZpSBsOqBmphdihWy4jpSgfDv6Og3dOziZuqX3v9UmnwDvXUL4521VKalGayjMWq+WnhLP
694/bPC8sAPhZlW/7Eak+uDnkcugyCDixwJe6SjxlazPFA5n1/4d1xwLP+qdNpVL1CUoUvSSTL3F
rGA/PQDcMWtBKzgpHkp890I4+nG2Xh40INZ4NE7d2OotF+8m28wq4SENTUFx94sq7yyXfTMugNaK
q+XHB7xE5AfrljFLlsZn+sm/dT3bXEtq57s0Cc++Ymos42WijgbaxFhxziKAD44l18Jc1LiYJ9Ik
9p+Uwo2GsOK4UqoVYdJYdeM9VlMac9S/LP7mJQbjaavlGA9A7j6/NDx8grSBrUQLaNqMVT/EKaSg
ccaEpjZ47mAIukeyUbc/0n7iz0j/+AgNg2esbJa+Q2HhxZLPdWvRgZsOQ2L2xdZJBGaaQCFHCJ0B
nvDWSbmBiv6OQ61SW2uWNumt7p+MEFr6gm4RHsoEjl/eSbh4PfIwITO6opdjx0oq7Bki3vwGZoHO
PVg2OHzs5Eq8NEP2duT7W62KULnPXugj0Xwcx2lRSsxW0+rG7gpEiyayM+kUhiUAbiti7MALgrDA
SD+GO8q7dDCc214LOAKaRuwoEm3xIdntx4NygzGho7B0qpl+UsTxarA1iav6TViMHu2/l2KeqR9N
rI7m0BFdrWarNQ2/kqx6ul58vYoflbdfVQbAU4154EqIW0Sxno+2EC/areBQtzOKBo1z27XEP6ZP
CXOhY0vpbjSG9gVGg81yS1oOCo6yE+9F+SLQHaI/B0RGVpqIdx9mK2RzTyP5VFDW/s8249l/4ypy
zssax9yskKwDC68c1IykgbX3b8pBuCwCbiEa8RoXTVG4d1OH852My+qJaJ47u8eIzBW9JPP/MExU
OubzWJawNKS/IVUtJ1ca8kigDSMDTTGVzX8XKCOlo4u904lzNRpIPgM2HWNk4vXmqvfHw8Mi1dJM
uu/A45RfdpfliSscY0vFhHP+tLNm6AiTD19fRPCJ3P3qMNrJoWJM9eTH5RBi2wm6TKC7fRXosGUT
hJXQzQBYaBU6PTI3TqFhjQZmuoxxSt/Seec7M7ihIgXxVRtxpN1pYCn0dImg9GYRDTKjBgG7kCQ/
PfIoBJk8trO+zqWrRXaLIPWHqFOEKluAbEVvUeA3jEG5c+TpnPL5bf9QcP8UKcMzKLI5Exytuh41
b7oCqaBrU6XYpxQvR4UVkHn19fA3KQNFUUUYGeWBgimhl1FiHkBfe60XQqE+CKwBtNtLi0cWUO/T
cBIxVGTBcOZFOV60Z7LsbEAIz8Gu4A2WdN4u+KztjTuF50KUuSGdSY1D41uf16nY6g/59HI6Lf2x
Es/cSaENqyTgeiwbnoCWFL1++Trxdk3bTZUXFOsp57T3QyjDsa0yVt6Plm1nBlF4UnY8lQbRspu9
s5XxyBT+culKnFZzDm3YB9jM1pz1ndueMYXnTlAeuy6rNXHvZ05yPmD4cmk34xp0fFWxCzx/Y0wG
j/AJtjPZKzfxtq3Wy5IXDy+s5uPwgfbhewOaVM6/7QLRWBauuw3VyoKqGY/Btxe9a4Ta+j4=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'armadura_ubicacion.py', "exec"), globals())
