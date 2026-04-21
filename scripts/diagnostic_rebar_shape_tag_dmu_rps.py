# -*- coding: utf-8 -*-
"""
RevitPythonShell / consola IronPython: diagnóstico del DMU de etiquetas vs RebarShape.

Ejecutar este archivo desde RPS (Open...) o pegar su contenido. Añade scripts/
al path para poder importar rebar_shape_tag_updater_dmu.
"""

from __future__ import print_function

import os
import sys

try:
    _scripts_dir = os.path.dirname(os.path.abspath(__file__))
except NameError:
    _scripts_dir = os.path.join(
        os.environ.get(u"USERPROFILE", u""),
        u"CustomRevitExtensions",
        u"BIMTools.extension",
        u"scripts",
    )
if _scripts_dir not in sys.path:
    sys.path.insert(0, _scripts_dir)

import rebar_shape_tag_updater_dmu as _dmu


def run_diagnostic(document):
    _dmu.print_diagnostic_rebar_shape_vs_tag_types(document)


def _resolve_doc():
    try:
        return doc  # noqa: F821 — RPS suele inyectar doc
    except NameError:
        pass
    try:
        return __revit__.ActiveUIDocument.Document  # noqa: F821
    except (NameError, AttributeError):
        pass
    return None


_d = _resolve_doc()
if _d is None:
    print(
        u"No se encontró documento. En RPS usa: "
        u"exec(open(r'...\\diagnostic_rebar_shape_tag_dmu_rps.py').read()) "
        u"o asigna doc = uidoc.Document antes de importar."
    )
else:
    run_diagnostic(_d)
