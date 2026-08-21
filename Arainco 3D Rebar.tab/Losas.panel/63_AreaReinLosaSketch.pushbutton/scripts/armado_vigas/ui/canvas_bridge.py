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
OrPnNrMKrxhE0oA7PrwTyjI402Zud2Ra2brTGo4W8OYrY0LnMu8vrd0rZU5LiEcTiSm7sGmcVIcu
jkxy/IxS3wjSOxUqHxTwL1xzYoeJbb+6oTRK8eH4uGKggbqwBRWoI+XtzZTAk1HG8xCReXBYJEqW
C52bgEQZ2ggW6cy022mPLbRoWAGNi/2/APvBHg0pe2AfNy2vl+25XOtR8DGxGCqnHMoo8vv6cc5d
183nNspEYvcqFVznRSNfraQk+GwsX/dlT8FT+M2XYMj7JwpeQEtgMRv5WLxs7jqnEZugmlkz37/a
iocBFDAucTTzlREtURuEDH5cQIT4EoP1+pwVG7rZRlTpqVt1dUjNCbhDAygwaDbMRRY4c3bgtKhp
HXL+JC1x6983RTTpKeCaUaMSIZUvfZyS45qrHfn4wz+tCe45tQWjmy4xjG4dbv6w921Fch7iw4PL
q5ybu6ISUq5hfUITLymQNgUyN9r7kVovWrFkuMRD1LQN8iLoUU17f23RlAXv4o+6lwDPBe0CEgaT
1BBhiwgkogvdKgNdukDlkTA0XQZ2S6QhYJ5eBgOZ3BVAswcsc9oI96H1KrWEWS6uUU819NnNLS94
JDDntyf+hTWeuHzZwA63pwUy9KAEkGxQLCATiaz+4JcxucrJdiFj88jwlXzJBssgKnLeKl8SFJvr
Y3yELWwBfWQUw8kCVuvCZ8zeVNb12j0Ruviye706IJDrtNKL5AxVlmwvUOgOGD2Ps11jTDTkFqKW
mDOGGdJy9UEP1KgW27In3f8/CS0yZG63hKDIWpv5r6xfYmAnxrk30Uvs8q1OeUODYbINMzdm2JBf
LZBmLX6pyuldv7rGLVZUNbH5kl5LSs6PMUoHRBOxWtvaYQd451euaMs4n9vDbEWNua/eD/vDJcd0
X9X2u4xiRbRJ0g7cP30cYqRQjf2VYFvfOez0bCgt54Lk2PR2tYfDisj5G/twgXC6IhnDAdIfLDA+
TV2rzTlRR9b7BmLTT5v4vWvDoGycnPwgegaUf8Nn6u6jh8uBGoOue7RzSW00AVBwc34ZVEpOrpI+
raemYMBQr6wTqQ0iW99FAYOa1Mlkm0agWqeJD3dlKprYTDTgPQT351bGzHtF0CIasv0/HWebU2A3
is1Bbgijs29sfZ0vrvl+5bYxAU+gHGjXKunN7l1b0YfB/Oww6/khPIJwlLIPIBxc7HBnTrmuMTWg
ZUbWtdqnGlT1H6mUl91biGGnlbze+gZkueHg7YzoYQs1lRpMYwPs0CvX7JJtj3jnp9vo6fRwxUku
BVSSoKnuLf+fvq5QvqsJehsj9Nx86Ky6MOPJ7pekQK8ZSzbvxRgVkkSxil2bGR28zQSEG7aoX1im
DDxpvi6wGk8Z0TPOxHoDKxJxlzcscRUT3QE8IG2Y4FHe19X+WEArlC8z8nV6SRyz92Naa7KUQCc7
OiL+Hr+P9AKRiogawHspfFugZ8pjrthi6MVlgFcMWGWuHjkbIyx3MsRjeA1Plo2N2EgFMw/5XXMt
tomgRF37TraNsKbrrcu1fE0Ghz2WZiFTZX5Q2NT21/xgsljHGQTGeSFukeaLs+9S7XamHjsP3NFv
c3eLvMBHoE2mvDC4ufqRHOwl4D7g+TnsAqg4t76FIyuVPew0w/bmfaWmqvQcn1DROrHj
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'canvas_bridge.py', "exec"), globals())
