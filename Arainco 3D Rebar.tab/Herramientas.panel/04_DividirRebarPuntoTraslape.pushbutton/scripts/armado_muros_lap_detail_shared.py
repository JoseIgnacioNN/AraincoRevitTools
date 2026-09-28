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
OrOXNr8KkBhE0YRFKLpTPYUmZrglZFNkDNZpaYJCxkBN4k5xavU5l/Un7c6YZmSULBO6c5CwSfeF
nK7d8wKsajf5VhOUh78Zi/+dY16qkUYApD2bEnvmri4LHEzwwY1mjntUgSBxKWdyQj9qgoA2e2is
Awso1YuV9pBsbsOiU77A6FznMgh5bEFDRM9kqIv4VGx/ra34WUMsh6M56ItfwYAQ+1cuPm+7zYol
FnBr7maHVzfpBD9M0/1tT1x41mnpJQeZust7xa4lCnoJyyMmoCB5oNTemllNjSg3faWw4uiOiWC+
Y+JVrEH3osk80hc5fbJuZ3HuyC6Q3O2Iy6WOB/8jLJAO5fzFrNiqRpPe5Zyh6btEOFMDYvTrtcK8
Beq8FZp/rSWtQWPC2J4Ot6jhgc+h+CupzJITCW5mONdtvfFB1XRMRH8qFvDean7LAeQsTrtmMC9/
ZfeADYQHo4iPYP3RokU1QdMvfOzfJJzLE4ToR0Qrm7Vp7h3E3/wOGER9Sesvyo672J3nl6Pzj2Vu
KrGKY99oa4+L8pZb17m+FWlY0T6NoH7drWR9hSNIoCsFpUNfYKk0TFdKbO8lLfVbMHv+bg4ZIiBv
ICUEFYk1L5nw/c/IylB1IpFxEWZZhoof8H+rbMgwHTqoLlmimiL4PSR+jP0zoWPx8lbcq1KfSIKF
ra+TTCqcK8RntO642Chdv+BCuhFBuqGQACEO89POzaoSUu+f4V/3dacr+D/xRt/LAHYAEMmGUft0
hIM0JEDLNxUNvIoSxcEvAttiRfal4Yyfen1SqUkfaQNcvl0W7nqNx2H+Wl9Z99FGFwWtESuZiVL6
gGHjopidJglC7NDgiSm7lMizDPyjitQSxED75SzHwvPt79X+nJ8VwmvHQktJ7mi9ZidtPyxHF9vE
oxsGtPR2Y+YVkGnBgyacl9bzq6RrAMPNKjv5kwnAQrpW9C35qocYKspVwrVlzIdjkqibKt6pV7O3
dAWCuYgEVrTGRfN8G7dawFuAyzireI8yWT/p2S+dfcuM9GRtc5ZyAJcjWYDAKowdTajO9h9WmoKg
KiaTB0HmqRna8bYj+8Bf6MrvFS24lFRxVfFX3zebVNVj1CuMlDnst1XlY12TBDYpjHix0fDQ2SKq
nNCyt8HITuS5KJ7JXpX7pxem6GhMQPw8fYL1kALfIztnH4ZBV5dcgMQmj5zKYMPnG/72qBMi0spn
3e4waUEqDZkbZpD8hWje/sCxuWaEaP1Sa+pUtEF9rqTz5q/Kj83ZbPx1ePwS++Lxe+VIX+YBXKO9
UzAERiJ40hsurUiFyA==
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
exec(compile(_SRC, 'armado_muros_lap_detail_shared.py', "exec"), globals())
