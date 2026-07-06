# -*- coding: utf-8 -*-
"""
Cota de anchura para muros (Wall).

- Vista: planta (ViewPlan).
- Selección múltiple al iniciar (PickObjects) o muros ya preseleccionados.
- Por cada muro, crea una cota lineal alineada entre las dos caras laterales (espesor).
- La cota se coloca a la mitad del largo, usando el punto medio de la LocationCurve.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    Line,
    LocationCurve,
    ReferenceArray,
    Transaction,
    ViewPlan,
    ViewSheet,
    Wall,
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
    _tangente_desde_bbox_plano,
    _tangente_desde_caras_verticales,
    _tangente_planta_desde_curva,
    _z_plano_vista,
)
from geometria_wall_foundation_cortes_muro import (
    location_curve_muro_host,
    punto_centro_location_curve_muro,
)

_TITULO = u"Arainco: Cota anchura muro"
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


def _tangente_muro(wall, view):
    crv = location_curve_muro_host(wall)
    if crv is not None:
        t = _tangente_planta_desde_curva(crv)
        if t is not None:
            return t
    try:
        loc = wall.Location
        crv = getattr(loc, "Curve", None)
        if crv is None and isinstance(loc, LocationCurve):
            try:
                crv = loc.Curve
            except Exception:
                crv = None
        if crv is not None:
            t = _tangente_planta_desde_curva(crv)
            if t is not None:
                return t
    except Exception:
        pass
    t = _tangente_desde_caras_verticales(wall, view)
    if t is not None:
        return t
    bb = None
    try:
        bb = wall.get_BoundingBox(view)
    except Exception:
        pass
    if bb is None:
        try:
            bb = wall.get_BoundingBox(None)
        except Exception:
            bb = None
    if bb is not None:
        t = _tangente_desde_bbox_plano(bb)
        if t is not None:
            return t
    return None


def _cotar_anchura_muro(doc, view, wall):
    pm = punto_centro_location_curve_muro(wall)
    if pm is None:
        return 0, [u"No se pudo obtener el punto medio de la LocationCurve del muro."]

    tangent = _tangente_muro(wall, view)
    if tangent is None:
        return 0, [u"No se pudo determinar el eje del muro."]

    ref_lo, ref_hi, ancho_aprox, p_lo, p_hi = _refs_ancho_por_caras(
        wall, view, tangent
    )
    if ref_lo is None or ref_hi is None or p_lo is None or p_hi is None:
        return 0, [
            u"No se encontraron caras laterales válidas para cotar el espesor del muro."
        ]

    width_dir = _perpendicular_planta(tangent)
    z = _z_plano_vista(view)
    centro = XYZ(float(pm.X), float(pm.Y), z)

    margen = _mm_a_pies(_MARGEN_LINEA_MM)
    off_ancho = _mm_a_pies(_OFFSET_ANCHO_MM)

    anchor = centro.Add(width_dir.Multiply(-off_ancho))

    p1 = _punto_sobre_width_dir(anchor, width_dir, float(p_lo) - margen, z)
    p2 = _punto_sobre_width_dir(anchor, width_dir, float(p_hi) + margen, z)

    dim = _crear_cota(doc, view, ref_lo, ref_hi, p1, p2)
    if dim is not None:
        return 1, []

    return 0, [
        u"No se pudo crear la cota de anchura (espesor aprox. {:.0f} mm). "
        u"Compruebe que la vista muestre las caras laterales del muro.".format(
            float(ancho_aprox or 0.0) * 304.8
        )
    ]


def _wall_filter():
    class _FiltroWall(ISelectionFilter):
        def AllowElement(self, elem):
            try:
                return elem is not None and isinstance(elem, Wall)
            except Exception:
                return False

        def AllowReference(self, ref, pt):
            return False

    return _FiltroWall()


def _pick_walls(uidoc):
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            _wall_filter(),
            u"Seleccione muros (Wall). "
            u"Finalizar en la cinta o Esc para cancelar.",
        )
    except OperationCanceledException:
        return []
    except Exception:
        return []
    if not refs:
        return []

    doc = uidoc.Document
    walls = []
    vistos = set()
    for ref in refs:
        try:
            el = doc.GetElement(ref.ElementId)
        except Exception:
            continue
        if el is None or not isinstance(el, Wall):
            continue
        try:
            eid = int(el.Id.IntegerValue)
        except Exception:
            eid = None
        if eid is not None and eid in vistos:
            continue
        if eid is not None:
            vistos.add(eid)
        walls.append(el)
    return walls


def _obtener_muros(uidoc, doc):
    """Usa muros preseleccionados o, si no hay ninguno, abre el picker múltiple."""
    walls = []
    vistos = set()
    try:
        for eid in uidoc.Selection.GetElementIds():
            el = doc.GetElement(eid)
            if el is None or not isinstance(el, Wall):
                continue
            try:
                key = int(el.Id.IntegerValue)
            except Exception:
                key = None
            if key is not None and key in vistos:
                continue
            if key is not None:
                vistos.add(key)
            walls.append(el)
    except Exception:
        pass

    if walls:
        return walls

    return _pick_walls(uidoc)


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
            u"Abra una planta donde se vea el muro y vuelva a ejecutar.",
        )
        return

    walls = _obtener_muros(uidoc, doc)
    if not walls:
        return

    creadas = 0
    errores = []

    with Transaction(doc, u"Arainco: Cota anchura muro") as t:
        t.Start()
        try:
            for wall in walls:
                try:
                    wall_id = wall.Id.IntegerValue
                except Exception:
                    wall_id = u"?"
                n, errs = _cotar_anchura_muro(doc, view, wall)
                creadas += n
                for err in errs:
                    errores.append(u"Id {0}: {1}".format(wall_id, err))
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
