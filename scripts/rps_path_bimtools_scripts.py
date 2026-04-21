# -*- coding: utf-8 -*-
"""
RevitPythonShell / IronPython 3.4: ejecutar con Open… en cada consola NUEVA.

IronPython no implementa importlib.util.module_from_spec como CPython.
Aquí se usa imp.load_source (si existe) o carga con compile/exec.

Tras ejecutar: variable global 'm' y sys.modules['rebar_shape_tag_updater_dmu'].
"""

from __future__ import print_function

import os
import sys
import types

m = None

_SCRIPTS = r"c:\Users\jinun\CustomRevitExtensions\BIMTools.extension\scripts"
_MOD_KEY = "rebar_shape_tag_updater_dmu"
_MOD_FILE = os.path.join(_SCRIPTS, _MOD_KEY + ".py")


def _load_via_exec(path, mod_key):
    """Carga un .py en un módulo nuevo (compatible IronPython 3.4)."""
    mod = types.ModuleType(mod_key)
    mod.__file__ = path
    mod.__name__ = mod_key
    mod.__package__ = None
    with open(path, "rb") as f:
        blob = f.read()
    try:
        src = blob.decode("utf-8")
    except Exception:
        src = blob.decode("utf-8", "replace")
    code = compile(src, path, "exec")
    try:
        mod.__dict__["__builtins__"] = __builtins__
    except NameError:
        import builtins as _builtins

        mod.__dict__["__builtins__"] = _builtins
    sys.modules[mod_key] = mod
    exec(code, mod.__dict__)
    return mod


if _SCRIPTS not in sys.path:
    sys.path.insert(0, _SCRIPTS)

print("sys.path[0]: {!r}".format(sys.path[0] if sys.path else ""))

if not os.path.isfile(_MOD_FILE):
    print(
        "ERROR: no está el archivo:\n  {!r}\n"
        "Corrige _SCRIPTS si la extensión está en otra ruta.".format(_MOD_FILE)
    )
else:
    try:
        del sys.modules[_MOD_KEY]
    except Exception:
        pass

    _loaded = False
    try:
        import imp

        if getattr(imp, "load_source", None) is not None:
            m = imp.load_source(_MOD_KEY, _MOD_FILE)
            sys.modules[_MOD_KEY] = m
            _loaded = True
            print("OK: imp.load_source -> m")
    except Exception as _ex:
        print("(imp.load_source falló: {!r})".format(_ex))

    if not _loaded:
        try:
            m = _load_via_exec(_MOD_FILE, _MOD_KEY)
            _loaded = True
            print("OK: carga por compile/exec -> m")
        except Exception as _ex2:
            print("ERROR exec: {!r}".format(_ex2))
            m = None

    if not _loaded:
        try:
            import rebar_shape_tag_updater_dmu as m

            sys.modules[_MOD_KEY] = m
            _loaded = True
            print("OK: import estándar -> m")
        except Exception as _ex3:
            print("ERROR import: {!r}".format(_ex3))

    if m is not None:
        print("Prueba: m.print_diagnostic_rebar_shape_vs_tag_types(doc)")
        print("  doc = __revit__.ActiveUIDocument.Document")
