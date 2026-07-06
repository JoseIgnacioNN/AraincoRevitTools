# -*- coding: utf-8 -*-
"""
Canvas WPF con la geometría en planta de una losa (Floor).

Shim de desarrollo: delega en el paquete portable del pushbutton
``15_FloorGeometryCanvas.pushbutton/scripts/run.py``.
"""

from __future__ import print_function

import os
import sys


def _find_portable_run():
    here = os.path.dirname(os.path.abspath(__file__))
    ext_root = os.path.dirname(os.path.dirname(here))
    run_path = os.path.join(
        ext_root,
        u"BIMTools.tab",
        u"Armadura.panel",
        u"15_FloorGeometryCanvas.pushbutton",
        u"scripts",
        u"run.py",
    )
    if os.path.isfile(run_path):
        return os.path.dirname(run_path)
    return None


def main(revit):
    scripts_dir = _find_portable_run()
    if scripts_dir is None:
        from Autodesk.Revit.UI import TaskDialog

        TaskDialog.Show(
            u"Arainco: Geometría losa (canvas)",
            u"No se encontró el paquete portable en "
            u"15_FloorGeometryCanvas.pushbutton/scripts/run.py",
        )
        return
    if scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)
    import run as _run_mod
    _run_mod.run(revit)
