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
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'armado_muros_lap_detail_shared.py', "exec"), globals())
