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
OrOXOL/qqBhE0ZRFpr03tVzAWdtyLovTrNLsQAe+q5XFScoJ+7pjZcfhJ1Vb0nIS7JAK7VuD5d0o
681cegslI6f853Ub5YueP54IChZjraND1GQQmYFNKjSyjKJXB7rDvN5CU1gqPWDKHKJ7cqD+CUBN
y4RN72ZaRDj7YrmoHMOi+z6NB6VEemQGD2XXoIkcqvDfBTRYBzmFWJKfSYAJ+3S1vnJiXmkDJ9pN
e6sBa7zTPXRR34ntk7QujOmS7vNQ8obrAVn0in5drj3P3GEzroHxSPVbN1U0R1s7bBQFByZ7yjV4
5S/miQKPdwB3gK6S/+fkz1AZRnnsRt9Lcnq5Atd/whPMJSb3CEVW6ws7M2LvTn1ceWSLW27Vcv3K
mlx4gib4NXBwCnNkTXq9ZEKqQWk5zThB1XPdpeHtBJx2P2dIrDWvkmShENa6Cy+WvIfl0Um9sZTU
0zkm0XyeIDre6TZ8mkYnBP4CDFIHnkMjd4H54rvXXQwAUUn2lP3cm1e+rlbUBQGvaQL4aNnwp9we
87/9vVqVqXy3/T9yblLM7tF1pouXuTflhE9psduPHOAekZvFSKIWeAD8aOmeY8jAUVj6SSgiHuFi
EjmDX2MxGvCjZTxE1Vl1Kh6Lk3Rqdl7zIva/pGoHncmVq6Vh6c0S53UckdEyc23aZS1bCr+IOM36
9F+Ft68iCS53VgdxH4L1P9XsNCE5ZuUt35z17+XQwY0ty3wyOQCwlfecJrPq27/8kh5qF5ZdV9yn
HN3/b24F6uFACAoT0+7s19zBKF1qdvzTtneuAVEEm8VbgZXbAPzgqYAKl3Kb+DhTf3L7Ue/4TzkN
eOdyuCoZCJjdbXwqIAARasMzuynbYu8jOM9LFJsFFCG7cAEDCkOkX/Yv7c5QuxpQa2TjB3BM/Jwp
T4jRUzNXGD42MiQsIbXZ7INq10KZB714fdaFfw4fYxg/fu+/SKUvCbVJ87jl2r4jCD9l1NbqVq3K
9nqbRuWh6ps6VET3noOf+lhux6GqdbQ1lqPdjK7tH6/lXk1GYoPMX+dErppAPiBd0ytjRkvbZTqG
X4/hce+ajAE4UytXUC+zy/HIHXi4Qc3Z36oXtuAE8VwE9I58tK5W8p3mix6oNRXJz6vuVYxuYz3n
mS7uuyCo+J6J38zKcPDqjgoIT/neRYeyOYo23AwDFSJQ3tqq02iwUuY7qOiFGYzHejyXKVgJ0uDk
Qx2//Dhe6tKO1GbMs8tSBRfTbF3XIOILOV5GQ5OLb9pb8xRD1eZ4xmUuZbIdaRSMUq8NRhhf9Frl
Nb7wkTUF1M7sIqOUwjnkCLuqbVh9lJHsXex2oAauqvLpALg2OiCWaVJJEkdsdjW71FVCCKzfY6ku
NR9IzoHAkEcsqedpS+yY81K8a8FtaBhfvHTnYZmcZoCDaqIUlShXLAEHE9IPu4Kpk3/LKCvltyf2
C2tlHOLGEimTvxqtI9rseMvH0TKdqQZq4kf9281Awep12B6X3vZ5Fg6fv9vBagJ6nVQNbxUcYcOI
RrGQvuMyV5cYVR0raWuJ2u7NHgQzsDi5MorYt6oZoaewGZQltEXfscNnlak6T+ondRTao+bXIlDF
7qxJdAJs3J4oRxRDRf8sh9f2LxfypLZOa1HNaCzJ5unTbLGMg/yugcqI2CldzK7D/sZ99RNlw2W8
Wo7jXAh929r5QyQHDcZ5keSoGm0aQDP1vQnvuXkCD3xiQ+hSOPNusN9KH2yg8o9nvqLDMKw+o+dJ
6lb6zMQHDhpn/+xKJe4FoouLu3x8yfSyVqwdkUoQwvKIlRYSEK+TjP0HE1hTE5XUSLjVdsY0c0dB
wlD76iLD86dwUK/QyWU7vInO/YiyBBbK0Ot5NvzCKE3z4eOltlvZOiwBVL2pG+EaqO0N7f1gLzx4
XlPdT23NrMzMN0bcj3RjXcbq+4QYJ+G6G4pfcZqyIRpYu/+lrKP5t2QHYL5vDkRFWvbubxs9QsAR
h9aAPDwgWWYbRoODgPLwOlyM8ZZ5S86IzCJzFMN7MW/PF4RbvpmU8GhqTH6ofoVZ/EJZD/R7Km9V
wMUgjNtTd8CGxTtZGwtTpyN8H0hzB/NOsN8+jW1CUojlZGr4PFmHuKpfIRAHLpm7hG2Cpbwrpglk
aRMrz/Lhdd161yFsAypsITUmX/BL2L9uCiaCUpaFVAXU2WG2KmGA2wvxwjYi/F6cGPMDzLWaOzpu
yPhuR0HRuPuF9buXTKkamLSNrKfKuRCHZY38Pk32jHdC9Z5VlCgJr4sLRV8meL1dGIX0HzZKReUn
Rs8oGgburN11RwubXgLHWEnmdBarfLxs8Tc84Gj9PfBYdZJNPI4Nx4iccs20tjTRIuzwGHVxsK2G
P8EUz6JoB0khOcECoreYnuN3I9zW/sdx72qVeSsiszg=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'bimtools_rebar_3d_visibility.py', "exec"), globals())
