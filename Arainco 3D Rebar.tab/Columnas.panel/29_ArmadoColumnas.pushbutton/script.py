# -*- coding: utf-8 -*-
"""
Punto de entrada pyRevit para "Armado Columnas".

Responsabilidades de este archivo:
1. Inyección dinámica de sys.path para que los sub-módulos del .pushbutton
   sean importables sin dependencia de rutas absolutas.
2. Inicialización del contexto Revit (doc, uidoc, uiapp).
3. Coordinación de las capas: selection → geometry → ui → creators.
"""
from __future__ import print_function

import os
import sys


# ---------------------------------------------------------------------------
# Inyección de path: DEBE ocurrir antes de cualquier import local
# ---------------------------------------------------------------------------

def _inject_pushbutton_root_to_syspath():
    """
    Agrega la raíz del .pushbutton a sys.path para que
    ``import core.geometry``, ``from ui.controllers import …``, etc. funcionen
    sin depender de la estructura de la extension.
    """
    try:
        _this_dir = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        _this_dir = os.getcwd()
    nd = os.path.normpath(_this_dir)
    if nd not in sys.path:
        sys.path.insert(0, nd)


_inject_pushbutton_root_to_syspath()


# ---------------------------------------------------------------------------
# Contexto pyRevit / Revit
# ---------------------------------------------------------------------------

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import FilteredElementCollector, Transaction, TransactionGroup
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import TaskDialog

try:
    # Contexto pyRevit
    from pyrevit import revit, script
    doc   = revit.doc
    uidoc = revit.uidoc
    uiapp = revit.uiapp
except Exception:
    # Contexto RPS / fallback manual
    try:
        doc   = __revit__.ActiveUIDocument.Document    # noqa
        uidoc = __revit__.ActiveUIDocument             # noqa
        uiapp = __revit__                              # noqa
    except Exception:
        raise RuntimeError(
            u"No se pudo obtener el contexto de Revit. "
            u"Ejecuta el script desde pyRevit o RPS."
        )


# ---------------------------------------------------------------------------
# Imports locales (después de la inyección de path)
# ---------------------------------------------------------------------------

from core.revit_compat  import create_version_adapter
from core.selection     import pick_structural_columns_optional, build_column_elements_ordered
from core.geometry      import (
    _element_id_iv,
    _canonical_section_mm_key,
    get_column_dimensions,
    build_troceo_scheme_rows,
)
from core.jobs          import (
    generate_bar_points,
    fuse_vertical_world_intervals_from_jobs,
)
from ui.controllers     import show_column_layout_wizard_singleton, ColumnLayoutWizardOutcome
from creators.longitudinal_creator import run_longitudinal_layout
from creators.stirrup_creator      import run_stirrup_layout


# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

_DEFAULT_LONG_BAR_DIAM_MM = 12.0
_COVER_FT = 0.15   # ~46 mm de recubrimiento por defecto (unidades internas Revit = pies)


# ---------------------------------------------------------------------------
# Función principal
# ---------------------------------------------------------------------------

def main():
    # 1. Selección interactiva de columnas
    try:
        refs = pick_structural_columns_optional(uidoc)
    except OperationCanceledException:
        return

    if not refs:
        TaskDialog.Show(
            u"Arainco: Armado Columnas",
            u"No se seleccionaron columnas estructurales.",
        )
        return

    try:
        columns_ordered = build_column_elements_ordered(doc, refs)
    except Exception as ex:
        TaskDialog.Show(u"Arainco: Armado Columnas", u"{}".format(ex))
        return

    # 2. Dimensiones de todas las columnas (cache)
    dims_cache = {}
    for col in columns_ordered:
        iv = _element_id_iv(col)
        if iv < 0:
            continue
        try:
            dims = get_column_dimensions(col)
            dims_cache[iv] = dims
        except Exception:
            pass

    if not dims_cache:
        TaskDialog.Show(
            u"Arainco: Armado Columnas",
            u"No se pudo obtener la geometría de ninguna columna seleccionada.",
        )
        return

    # 3. Metadatos de sección (claves únicas + título)
    section_keys_seen = {}
    for iv, dims in dims_cache.items():
        width, depth = dims[0], dims[1]
        sk = _canonical_section_mm_key(width, depth)
        if sk not in section_keys_seen:
            s_mm, L_mm = sk
            section_keys_seen[sk] = u"{}×{} mm".format(s_mm, L_mm)

    section_meta = [(sk, title) for sk, title in sorted(section_keys_seen.items())]

    # 4. Filas del esquema de troceo (para el Paso 2 del wizard)
    troceo_rows = build_troceo_scheme_rows(columns_ordered)

    # 5. Wizard WPF
    wiz = show_column_layout_wizard_singleton(
        section_meta=section_meta,
        troceo_rows=troceo_rows,
        uiapp=uiapp,
        uidoc=uidoc,
        doc=doc,
        default_bar_diam_mm=_DEFAULT_LONG_BAR_DIAM_MM,
    )

    if wiz is None or wiz.cancelled or wiz.already_running:
        return

    if not wiz.section_grid_config:
        TaskDialog.Show(
            u"Arainco: Armado Columnas",
            u"No se configuró ninguna sección. No se creó armadura.",
        )
        return

    # 6. Determinar diámetro global del longitudinal
    global_diam_mm = float(getattr(wiz, "global_long_bar_diam_mm", _DEFAULT_LONG_BAR_DIAM_MM))

    # 7. Generar rejilla de puntos y fusionar tramos verticales
    jobs = []
    for col in columns_ordered:
        iv = _element_id_iv(col)
        if iv < 0 or iv not in dims_cache:
            continue
        width, depth, height, center, grid_vs, grid_vl = dims_cache[iv]
        sk = _canonical_section_mm_key(width, depth)
        cfg = wiz.section_grid_config.get(sk)
        if cfg is None:
            continue
        ba = int(cfg.get("bars_a", 4))
        bb = int(cfg.get("bars_b", 4))
        inc_in = bool(cfg.get("include_inner_outline", False))
        side_short = min(width, depth)
        side_long  = max(width, depth)
        short_on_x = width <= depth

        pts = generate_bar_points(
            center, side_short, side_long, short_on_x,
            ba, bb, _COVER_FT, inc_in, grid_vs, grid_vl,
        )
        jobs.append(dict(
            height=height, nominal_n=len(pts),
            raw_pts=pts, width=width, depth=depth,
            short_on_x=short_on_x, elem=col,
            section_key_mm=sk, bars_a=ba, bars_b=bb,
            include_inner_outline=inc_in,
        ))

    if not jobs:
        TaskDialog.Show(
            u"Arainco: Armado Columnas",
            u"No se generaron puntos de rejilla para ninguna columna.",
        )
        return

    try:
        tol_ft = float(doc.Application.ShortCurveTolerance)
    except Exception:
        tol_ft = 1.0 / 304.8

    fused_world = fuse_vertical_world_intervals_from_jobs(jobs, tol_ft)

    if not fused_world:
        TaskDialog.Show(
            u"Arainco: Armado Columnas",
            u"La fusión de tramos no produjo barras. Revisa la selección.",
        )
        return

    # 8. Crear armadura longitudinal
    run_longitudinal_layout(
        doc=doc,
        fused_world=fused_world,
        wiz=wiz,
        dims_cache=dims_cache,
        columns_ordered=columns_ordered,
    )

    # 9. Crear estribos (si el wizard los configuró)
    if wiz.has_stirrups:
        run_stirrup_layout(
            doc=doc,
            columns_ordered=columns_ordered,
            dims_cache=dims_cache,
            wiz=wiz,
        )


# ---------------------------------------------------------------------------
# Ejecutar
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    main()
else:
    # pyRevit ejecuta el módulo directamente sin __name__ == '__main__'
    main()
