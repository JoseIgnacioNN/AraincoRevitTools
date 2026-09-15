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
OrPnMjkKqB5EsoR4LSVR2NY5EndgySNzwkT5yS7FACTCCVarVEgBKsbl/DEhgWGZQDMd8Ms6oWeE
8o2Tl/CVCtrW5aJ8F9rZH8cyXdG0Xz546Zbh0h0yAq8/CI6hgamhW6wJxYLWJbQPl5Z9PCtiCm15
1E09saGtd+BwKutvc71X3NHsEmjXhIRXx+D63t4VgwqlStyrF1Q+S9qCCFXFt1e9CwfKsEhyarzm
xmjpWu8V5SVPDxHiW5CcD6RmRr8zvmTe5Cmw0R+uCmuhjvkEcyjgcy8biAGiSGoezRxevaQacvkH
4Abut9LArCoB6/ZgvpPXnYlv0DRVREHLfPJG3xsjEKjZTnaHnTdoHH2SPFTF88UQNLVQkIyc25m4
cpAp9aFnL0hDgxM+POtuvxZOJ6NgGqMoJ5QbA2EzhvVaygdD48ICzD4HYHM5rQSwRv8BIGPgmc43
gMO2tMTX2wH6TW38iRsWWTt2ehbml0A4Cc/zBi5Puf6lFniIcAhcts0enx2aiOFUA1rQ25S890Gm
gY0qa3RGL4qAsW3JNVskTgokSrufTfYqw7wAt4/My53QnJ+mkr0HCN9/tR/glz3H6L3KR1mIBlTf
2m00NEgtg+vcjy+4cBSII36WJ7CHtjgQtnUDeByV7oYYaRvpymf7hEuikOaB
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
exec(compile(_SRC, 'instruction_dialog.py', "exec"), globals())
