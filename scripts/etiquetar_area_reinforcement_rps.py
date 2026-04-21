# -*- coding: utf-8 -*-
"""
Etiqueta uno o varios Area Reinforcement seleccionados en la vista activa.

Ejecutable en RevitPythonShell (RPS) o pyRevit — Revit 2024+ (API IndependentTag.Create).

Requisitos:
- Vista activa válida para etiquetar: no plantilla, no perspectiva; en 3D la vista debe estar bloqueada.
- Debe existir en el proyecto una familia de etiqueta cargada para la categoría Area Reinforcement
  (modo "Por categoría" / TM_ADDBY_CATEGORY).

Punto de inserción: centroide volumétrico del elemento host (muro/losa/Part, etc.)
obtenido con GetHostId + geometría (Solid.ComputeCentroid ponderado por volumen).
Si no hay host o no hay sólidos válidos, se usa el centro del bounding box del host
y en último caso el del propio Area Reinforcement.

Uso RPS: selecciona el/los Area Reinforcement y ejecuta el script.

Si antes ejecutaste seleccionar_tipo_etiqueta_por_familia_rps.py en la misma consola,
se usará __main__.TAG_SYMBOL_ID a menos que definas TAG_SYMBOL_ID en este archivo.
"""

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    ElementId,
    GeometryInstance,
    IndependentTag,
    Options,
    Reference,
    Solid,
    TagMode,
    TagOrientation,
    Transaction,
    View3D,
    XYZ,
)
from Autodesk.Revit.DB.Structure import AreaReinforcement
from Autodesk.Revit.UI import TaskDialog

# True = etiqueta con línea de llamada; False = solo cabeza en el punto indicado
ADD_LEADER = False

# Orientación de la cabeza de etiqueta
TAG_ORIENTATION = TagOrientation.Horizontal

# Opcional: ElementId de FamilySymbol de etiqueta concreta. None = usar tipo por categoría del proyecto.
TAG_SYMBOL_ID = None


def _effective_tag_symbol_id():
    if TAG_SYMBOL_ID is not None and TAG_SYMBOL_ID != ElementId.InvalidElementId:
        return TAG_SYMBOL_ID
    try:
        import __main__ as _main

        sid = getattr(_main, "TAG_SYMBOL_ID", None)
        if sid is not None and sid != ElementId.InvalidElementId:
            return sid
    except Exception:
        pass
    return None


def _get_doc_uidoc():
    try:
        return doc, uidoc
    except NameError:
        u = __revit__.ActiveUIDocument
        return u.Document, u


def _view_ok_for_tag(view):
    if view is None:
        return False, u"Vista nula."
    if view.IsTemplate:
        return False, u"La vista activa es una plantilla de vista; abre una vista de modelo."
    # No usar ViewType.Perspective: en IronPython 3.4 `from ... import ViewType` puede
    # resolverse al builtin `type` y fallar con MissingMemberException.
    if str(view.ViewType) == "Perspective":
        return False, u"No se pueden crear etiquetas en vista en perspectiva."
    if isinstance(view, View3D) and not view.IsLocked:
        return False, u"En vista 3D la cámara debe estar bloqueada para etiquetar."
    return True, None


def _bbox_center_xyz(element, view):
    """Centro del bounding box del elemento (vista o modelo)."""
    bb = element.get_BoundingBox(view)
    if bb is None:
        bb = element.get_BoundingBox(None)
    if bb is None or bb.Min is None or bb.Max is None:
        return None
    mn, mx = bb.Min, bb.Max
    return XYZ((mn.X + mx.X) * 0.5, (mn.Y + mx.Y) * 0.5, (mn.Z + mx.Z) * 0.5)


def _solid_centroid(solid):
    try:
        return solid.ComputeCentroid()
    except Exception:
        return None


def _volume_weighted_centroid(element):
    """
    Centroide de la geometría del elemento: promedio de ComputeCentroid de cada
    Solid ponderado por volumen (mismo criterio que masa uniforme).
    """
    if element is None:
        return None
    opts = Options()
    opts.ComputeReferences = False
    try:
        geom_elem = element.get_Geometry(opts)
    except Exception:
        return None
    if geom_elem is None:
        return None
    vol_sum = 0.0
    sx = sy = sz = 0.0
    for obj in geom_elem:
        if obj is None:
            continue
        if isinstance(obj, Solid) and obj.Volume > 1e-12:
            c = _solid_centroid(obj)
            if c is None:
                continue
            v = obj.Volume
            sx += c.X * v
            sy += c.Y * v
            sz += c.Z * v
            vol_sum += v
        elif isinstance(obj, GeometryInstance):
            try:
                inst_geom = obj.GetInstanceGeometry()
                if inst_geom is None:
                    continue
                for g in inst_geom:
                    if isinstance(g, Solid) and g.Volume > 1e-12:
                        c = _solid_centroid(g)
                        if c is None:
                            continue
                        v = g.Volume
                        sx += c.X * v
                        sy += c.Y * v
                        sz += c.Z * v
                        vol_sum += v
            except Exception:
                pass
    if vol_sum < 1e-12:
        return None
    return XYZ(sx / vol_sum, sy / vol_sum, sz / vol_sum)


def _area_reinforcement_host_id(area_rein):
    try:
        hid = area_rein.GetHostId()
        if hid is not None and hid != ElementId.InvalidElementId:
            return hid
    except Exception:
        pass
    return None


def _insertion_point(document, area_rein, view):
    """
    Punto para IndependentTag: centroide del host; reservas bbox host y bbox refuerzo.
    """
    hid = _area_reinforcement_host_id(area_rein)
    if hid is not None:
        host = document.GetElement(hid)
        if host is not None:
            c = _volume_weighted_centroid(host)
            if c is not None:
                return c
            c = _bbox_center_xyz(host, None)
            if c is not None:
                return c
    return _bbox_center_xyz(area_rein, view)


def _create_tag(document, view, area_rein):
    ref = Reference(area_rein)
    pnt = _insertion_point(document, area_rein, view)
    if pnt is None:
        raise Exception(
            u"No se pudo calcular un punto para la etiqueta (host sin geometría ni bbox). "
            u"Comprueba que el refuerzo y su host sean válidos."
        )

    tag_sym_id = _effective_tag_symbol_id()
    if tag_sym_id is not None:
        return IndependentTag.Create(
            document,
            tag_sym_id,
            view.Id,
            ref,
            ADD_LEADER,
            TAG_ORIENTATION,
            pnt,
        )

    return IndependentTag.Create(
        document,
        view.Id,
        ref,
        ADD_LEADER,
        TagMode.TM_ADDBY_CATEGORY,
        TAG_ORIENTATION,
        pnt,
    )


def main():
    document, uidoc_ = _get_doc_uidoc()
    view = uidoc_.ActiveView

    ok, msg = _view_ok_for_tag(view)
    if not ok:
        print(u"Error: {}".format(msg))
        try:
            TaskDialog.Show(u"Etiquetar Area Reinforcement", msg)
        except Exception:
            pass
        return

    ids = list(uidoc_.Selection.GetElementIds())
    if not ids:
        print(u"Error: No hay nada seleccionado. Selecciona al menos un Area Reinforcement.")
        return

    targets = []
    for eid in ids:
        el = document.GetElement(eid)
        if el is not None and isinstance(el, AreaReinforcement):
            targets.append(el)

    if not targets:
        print(
            u"Error: Ningún elemento seleccionado es Area Reinforcement. "
            u"Selecciona refuerzos de área y vuelve a ejecutar."
        )
        return

    trans = Transaction(document, u"Etiquetar Area Reinforcement")
    trans.Start()
    try:
        creadas = []
        for ar in targets:
            tag = _create_tag(document, view, ar)
            if tag is not None:
                creadas.append(tag.Id.IntegerValue)
        trans.Commit()
        resumen = u"Etiquetas creadas: {} (IDs: {}).".format(
            len(creadas),
            u", ".join(str(i) for i in creadas),
        )
        print(resumen)
        try:
            TaskDialog.Show(u"Etiquetar Area Reinforcement", resumen)
        except Exception:
            pass
    except Exception as ex:
        trans.RollBack()
        err = u"No se pudo crear la etiqueta: {}".format(str(ex))
        print(err)
        try:
            TaskDialog.Show(u"Etiquetar Area Reinforcement", err)
        except Exception:
            pass


main()
