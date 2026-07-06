# -*- coding: utf-8 -*-
"""
Cota de anchura para zapatas de muro (WallFoundation).

- Solo aplica a Wall Foundation (no fundaciones aisladas ni losas de cimentación).
- Vista: planta (ViewPlan).
- Selección múltiple al iniciar (PickObjects) o zapatas ya preseleccionadas.
- Por cada zapata, crea una cota lineal alineada entre las dos caras laterales (anchura).
- La cota se coloca a la mitad del largo, usando el punto medio de la LocationCurve
  del muro host.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    Line,
    ReferenceArray,
    Transaction,
    ViewPlan,
    ViewSheet,
    WallFoundation,
    XYZ,
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

try:
    from Autodesk.Revit.Exceptions import OperationCanceledException
except Exception:
    OperationCanceledException = Exception

from cota_wall_foundation_planta import (
    _mm_a_pies,
    _perpendicular_planta,
    _refs_ancho_por_caras,
    _tangente_wall_foundation,
    _z_plano_vista,
)
from geometria_wall_foundation_cortes_muro import (
    host_wall_from_wall_foundation,
    punto_centro_location_curve_muro,
)

_TITULO = u"Arainco: Cota anchura zapata de muro"
_OFFSET_ANCHO_MM = 300.0
_MARGEN_LINEA_MM = 150.0


def _punto_sobre_width_dir(anchor, width_dir, w_abs, z):
    """Punto en planta con proyección ``w_abs`` sobre ``width_dir`` (Z fija)."""
    w_anchor = float(anchor.DotProduct(width_dir))
    pt = anchor.Add(width_dir.Multiply(float(w_abs) - w_anchor))
    return XYZ(float(pt.X), float(pt.Y), float(z))


def _crear_cota(doc, view, ref_a, ref_b, pt1, pt2):
    try:
        ra = ReferenceArray()
        ra.Append(ref_a)
        ra.Append(ref_b)
        linea = Line.CreateBound(pt1, pt2)
        return doc.Create.NewDimension(view, linea, ra)
    except Exception:
        return None


def _cotar_anchura_wall_foundation(doc, view, wf):
    wall = host_wall_from_wall_foundation(wf)
    if wall is None:
        return 0, [u"No se encontró el muro host de la zapata de muro."]

    pm = punto_centro_location_curve_muro(wall)
    if pm is None:
        return 0, [
            u"No se pudo obtener el punto medio de la LocationCurve del muro host."
        ]

    tangent = _tangente_wall_foundation(wf, doc)
    if tangent is None:
        return 0, [u"No se pudo determinar el eje del muro / zapata."]

    ref_lo, ref_hi, ancho_aprox, p_lo, p_hi = _refs_ancho_por_caras(
        wf, view, tangent
    )
    if ref_lo is None or ref_hi is None or p_lo is None or p_hi is None:
        return 0, [
            u"No se encontraron caras laterales válidas para cotar el ancho de la zapata."
        ]

    width_dir = _perpendicular_planta(tangent)
    z = _z_plano_vista(view)
    centro = XYZ(float(pm.X), float(pm.Y), z)

    margen = _mm_a_pies(_MARGEN_LINEA_MM)
    off_ancho = _mm_a_pies(_OFFSET_ANCHO_MM)

    # Punto ancla: mitad del largo (centro del muro) + desplazamiento lateral
    # para que la cota quede fuera de la huella de la zapata.
    anchor = centro.Add(width_dir.Multiply(-off_ancho))

    # La línea de cota debe ser PARALELA a width_dir (⟂ al eje del muro), no al
    # tangent. Si va paralela al muro, Revit no puede medir entre las caras laterales.
    p1 = _punto_sobre_width_dir(anchor, width_dir, float(p_lo) - margen, z)
    p2 = _punto_sobre_width_dir(anchor, width_dir, float(p_hi) + margen, z)

    dim = _crear_cota(doc, view, ref_lo, ref_hi, p1, p2)
    if dim is not None:
        return 1, []

    return 0, [
        u"No se pudo crear la cota de anchura (ancho aprox. {:.0f} mm). "
        u"Compruebe que la vista muestre las caras laterales de la zapata.".format(
            float(ancho_aprox or 0.0) * 304.8
        )
    ]


def _wall_foundation_filter():
    class _FiltroWallFoundation(ISelectionFilter):
        def AllowElement(self, elem):
            try:
                return elem is not None and isinstance(elem, WallFoundation)
            except Exception:
                return False

        def AllowReference(self, ref, pt):
            return False

    return _FiltroWallFoundation()


def _pick_wall_foundations(uidoc):
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            _wall_foundation_filter(),
            u"Seleccione zapatas de muro (Wall Foundation). "
            u"Finalizar en la cinta o Esc para cancelar.",
        )
    except OperationCanceledException:
        return []
    except Exception:
        return []
    if not refs:
        return []

    doc = uidoc.Document
    wfs = []
    vistos = set()
    for ref in refs:
        try:
            el = doc.GetElement(ref.ElementId)
        except Exception:
            continue
        if el is None or not isinstance(el, WallFoundation):
            continue
        try:
            eid = int(el.Id.IntegerValue)
        except Exception:
            eid = None
        if eid is not None and eid in vistos:
            continue
        if eid is not None:
            vistos.add(eid)
        wfs.append(el)
    return wfs


def _obtener_wall_foundations(uidoc, doc):
    """
    Usa todas las WF preseleccionadas o, si no hay ninguna, abre el picker múltiple.
    """
    wfs = []
    vistos = set()
    try:
        for eid in uidoc.Selection.GetElementIds():
            el = doc.GetElement(eid)
            if el is None or not isinstance(el, WallFoundation):
                continue
            try:
                key = int(el.Id.IntegerValue)
            except Exception:
                key = None
            if key is not None and key in vistos:
                continue
            if key is not None:
                vistos.add(key)
            wfs.append(el)
    except Exception:
        pass

    if wfs:
        return wfs

    return _pick_wall_foundations(uidoc)


def ejecutar(uidoc):
    doc = uidoc.Document
    view = doc.ActiveView

    if isinstance(view, ViewSheet):
        TaskDialog.Show(_TITULO, u"Abra una vista de modelo, no una lámina.")
        return

    if not isinstance(view, ViewPlan):
        TaskDialog.Show(
            _TITULO,
            u"La vista activa debe ser una planta (ViewPlan).\n"
            u"Abra una planta donde se vea la zapata de muro y vuelva a ejecutar.",
        )
        return

    wfs = _obtener_wall_foundations(uidoc, doc)
    if not wfs:
        return

    creadas = 0
    errores = []

    with Transaction(doc, u"Arainco: Cota anchura zapata de muro") as t:
        t.Start()
        try:
            for wf in wfs:
                try:
                    wf_id = wf.Id.IntegerValue
                except Exception:
                    wf_id = u"?"
                n, errs = _cotar_anchura_wall_foundation(doc, view, wf)
                creadas += n
                for err in errs:
                    errores.append(u"Id {0}: {1}".format(wf_id, err))
            t.Commit()
        except Exception as ex:
            try:
                t.RollBack()
            except Exception:
                pass
            TaskDialog.Show(_TITULO, u"Error inesperado al crear la cota:\n{}".format(ex))
            return

    if creadas == 0:
        msg = u"No se pudo crear ninguna cota de anchura."
        if errores:
            msg += u"\n\nDetalles:\n" + u"\n".join(errores)
        TaskDialog.Show(_TITULO, msg)
        return

    if errores:
        msg = u"\n".join(errores)
        TaskDialog.Show(_TITULO, msg)


def _main():
    try:
        _uidoc = __revit__.ActiveUIDocument  # noqa: F821
    except NameError:
        TaskDialog.Show(
            _TITULO,
            u"Este script debe ejecutarse desde pyRevit con la variable __revit__ disponible.",
        )
        return
    if _uidoc is None:
        TaskDialog.Show(_TITULO, u"No hay documento activo.")
        return
    ejecutar(_uidoc)


if __name__ == "__main__":
    _main()
