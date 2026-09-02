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
OrPXNi8KaBlEEbA/PsvlsibwdSd4f3BKgbwfD9pfdPYKdnYi0PLAJzVBxv6VIvZPJCGeCZJhziou
poigkERTiiucWsy6/fzdp6JjmzdpiIHsysUTn4T3tzQFncwzp5QJzKcuqFcz9y9F+xNugzm7LjK6
0x04Xi4yLIbzXcMiWNVi21Lb9Dh65gUL1pDusZMSmicXTwbMN3cgPuyw5AqGCFWcbEUhAF6G+Nd/
uIo0OW/li3QsFiaiSzbRK8eZdSVw2/Sr68jLaJQm9wTGBwkDBskJG3LdgikQJwoUGP8XaZMAwuhN
ZCEu+055ZK5lO0jOT40LP6e3jOzDPgYgTKzpazBlfSUaR98FJiokVPkKm65KVwYuw9Zbw4Q2ep3i
ZoMqf0htTvz74MGdFRRhNd661/2zfRRrcBvzWFDYtFZTASjleuG/pWcoAFbXgDEV/Vnbo/IobFJO
CLYZTM56EnoMCmF+JjV6pb0rwPfR1TxmNyWrZnAn+3aUTO8LmY+Q7nDA16umM0tHnAU5bNE77wmq
B9HXVashZPDhTfTaquikbCOyj1dZORnf9cC34Gpi8FYzMhOOjuWgP/4xvL7Wlc88R+hq5pwq82Et
UtVoFTbm9RvhiiJ+bHkqezs0i1P1D4E6GOKQBG+jIJCl7Vct7N3x9VgNL+5lZ/2QRtFw3fyB5V7P
YRAuc1lKINniTnhX8JIttDvH5s8JNS/Brz1zPmes+0hLo3hmvPPCPcJnLkz+IqeuSEKlSSxM9SBy
6kY7G2fhX4Zn0THR596+tudroZcSQJAKs1a96w0xZg6dE95HdbGpVuuCe0Maq14iu975AEegWheR
1DsGDzOWMx0r7U97kSQPwYn+sILx65B+U2ieR1pCxHy+D0DvCanF0ajYMf7EXVD/v0vN7VXQJYlF
WDkM6Qn9bEUvKrQQVnGE86Irq7HL9pCAjQiaNRjJiTjygAk/IZ+IuFG1C+3JFMWqC/6PJhQmgwz1
XNDi9pcriVCaxpZWuC5DJwNDuL5HsbtOGKuN9Tlvr3gOr2MUtAhYfTl0/VZkP+vi04G0/mCGZDy7
XTw6mMoRAU9hNO5NxztNF5AWgSESTNmx76ZGwmEISri+eb564ut5KcA0eDX9GSUP5aQIzs/IC0iu
vSQ5GzvQ9bVy+jkMVpUdQKVFCbD0WNG5sXJdQ+kvj1CY0D6D8oL4fA0lCkphr910SNbyJcpf10o6
gPZC9h4Q2EV7NMf/bZzwFszlLN0vuctYgtR+tC0rr6d+yrSYCHKpJtJ93x+tVxrrFH114R9dKqN1
Ba8Et66Xlw+yLnbXHeHuDODEW6gW1dp3NxPyXWf79MdhyguNpbIvnx2RiBcEAesPOFM5Z2Lr3c1q
i86zDhiflX0kgk+rIL8Wq4qhKD2xDRqTtgosE/Z/ENW5vXT3qJtg4qlAXfXNxd1OEoNJba1jPFEm
2+6X5EmRqQ1gXjhf/6qKEd45ZNfh89kdMBUXLG0hWP5cEibFvRGWjefd8QyUJjKtU3FiztLqxZcD
BdeogPJSjEm+6j9VvDEZzpdM8imTXBSbWdMmh//DVfAWPD2UysEPgUTV/cRDNJGyRcyKLijszef4
Uc/ZyV4WyjgW6g+a+kWFb82avvjr/QqHtMhNtonCY8VNeOosCFg2fm14FiDIGpjEbHtZefsMAJlt
yUxNZOpM25ROsiNMcUYw38oZZh3FnNXgOt5gk50i+97wRCW0zjLCuGn63WfTIjFZzNuQrGctWVzt
CCqhIOKz9fINW8tCyMDNWR7L5vc2hgGiSaxuo3Ru97sfuQYboTxXjyB9rtmEJTOtJ9d09FErw3oV
v7s8QsWJe5tUpkuroC0g10V99nI2A1wxX6PxUYbjrG6YRgOR6Ykishzuyrh9Uj9yJVQiR8LfYNQJ
6Ub1Cy30+6weL/52Buvfrl/Zpe/1xTeOhri00Ryp4sdfa51nHz8xZA==
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
exec(compile(_SRC, 'core.py', "exec"), globals())
