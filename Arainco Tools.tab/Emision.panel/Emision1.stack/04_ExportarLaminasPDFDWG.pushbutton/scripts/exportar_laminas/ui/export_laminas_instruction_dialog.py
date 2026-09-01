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
OrPPNDkKqB5Espx4LSdRWuqcYmAelNdqcdRg5UprJxHCJTE1MDyRJnD6YLxrX6NkgC4Sl27LOiyz
pUbNQvUHYRZi8/tdg77NpHKD0/HXegzNQEDD1ryN5DO2L8y/2igrdvkbrSEFUzgknffPmAyxyIFd
aXdRhZSkLkQeTSNvboo+P0tI/aq9gUU8fFxYeTaCZp2WPgDDZ+QzbcLkpmQiUG4Q8NbBfQu/AAtZ
I5faSiNRuWUgvvbbc1szKXx7NIuxYPPmxkNm9EU6eXBHUQCupR8qzEvbUVdfLE0VkhI04kRy/xKt
Ei8Nvg3WJJw2gBGBlA+D5leJJddnIAPXKBCYLTJ4PNUF9X1yAloMih2WEP4kRubmFvuC8IYYAyOD
fiAZqXYsgG9xP1Qrx87M1BrRqeEilffF4TpIzFf/4HiYpptVB/V9HLD+NMHXuJpC7RvUmzynW+lj
J8AG0SSqdphiU7c67d6YrceeTGTcZACVz9ulufzUdSgcuuozTmTAJqEHowJam9lBLMUWO12dNtmu
fc/lQJtqtUqGyGXqNfGqfyS8xgCmvJ5rnVi6Pg0DepFFjDpby4vVwlJUyjq4Fd1rRDInfYsec1FD
tuR8i1bFGQHG8NEZ3Z8a4Q5AO9EB+ijRKEZSFacRbAjHbn0JmufFG4nH4wUmqK+2gkMf4LSqZppH
fbjrtTZDA4cdsqS16P+B3NQ0M7hI21u1UWRJeYbsbceNA7+OU1eMVSkXS3w3Tv2v8eX8oZUuTh5p
Nb8OqlDWA3KlbjQpWFWoBIXz/ViGFSkj+KK2Hg1k6I9ljx+xYJukUk4J0w/NRXZXd1bF8dPsA5lT
+LN9JIzfxlJCkEA6S5EpCnil2zHQ3QBIyQjGQ9H5epqPKLxhRwch3sg5c+k831/K6VQny5/cQiO1
5sLfY73kwX98y8W6HwvPghAHku1a3+xh3bgNwnvs1Ibj
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
# type(u"") = unicode en Py2/IronPython, str en Py3. No usar isinstance(..., str):
# en IronPython zlib devuelve str (bytes UTF-8) y compile() lo leeria como cp1252.
if not isinstance(_SRC, type(u"")):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'export_laminas_instruction_dialog.py', "exec"), globals())
