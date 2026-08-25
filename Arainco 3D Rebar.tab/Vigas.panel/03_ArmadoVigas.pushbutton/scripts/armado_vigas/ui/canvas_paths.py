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
OrOvNLMKpx5E0ohHaLHgL/+Ej2MjYdo39NwJNDr0ZXMbTosXr0LFn8k6pFNfsmGVB6LcDgSaBWpm
Wo3fN1UQr6XUYwbspY3dKUcrpdYdeOI9wvDTXTRr5fb5SC9EhjNAlSGCEUefdwv5z08Qa2Q2PUxt
veikftOU5wakcRHfTvCKSRrUx+ajb8SiQSIQDg2hqqEoxRZxUjJUXSzZDBkyVl8/rHzjSD5zNR8/
vit8akAZXWnrPzPBsqrIWCmeghLbqjUKe9Vlk01b2YCvNraB2Mfx9OFIK2wvfUn22J8BG8252qP0
YvE9vrau3uFFRDtLhZpVgoAGu6rM9OPvTKFJTuJT9HNfstg26ltWP0aCGnL3rGqzTlqW5oEyeGpG
Aif00MT+JnTXX3cdY7rE84cdgnp1z80ksqwULvz9+/n/FoPNufN0Q91mojtH1mtPn1mzZ5kpk+al
jlCLFVUAz2p7n+4gVOWjrNsFw+nRf/QimC1VR/tfSG1T/yzQ0YbZv5OpTuOa2dd6PgIWyV9j6gtB
KXtgJXgL227vWYIpScry5EKJgxzThcBGWu9Ds43NZ61By485LfyoCQ0ajNXUe5NZmXYZoN/qhNDu
xgdyINgx0dvU27t/oMsyAGf4iqqPGUILPnTdniyYXaerIZWxNCjmFBIV4M9pU/rBnd1ivq/WuIfy
1IUw3QbJuy7BR9Ul5ofZIUHM2UCk5xG9pr0rsw/AxZD/BIXl3sS0t9L6d/PQrSeiNXUOdJxDgiOh
hAc2QeZtNcpYl0xz1xOBiVztQA1gPfI/3nymb+yKZN/B2pRu/KsEGEZGAjIFbqFMwnNNZopAAnd4
HBVY2R1vG2WPChNUHZaTsVhLVJDuwXaCTytnEwhZA3ycxqdtpCfFVtxVP2C9p8W4d1o3l5s+B9rO
jnypJXBmN0J5BHK09hPfPZ6T19tjd7KWOyo7NVPQbsDK3gE=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'canvas_paths.py', "exec"), globals())
