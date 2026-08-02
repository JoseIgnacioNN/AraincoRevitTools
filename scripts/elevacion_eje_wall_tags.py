# -*- coding: utf-8 -*-
"""
Etiquetas de espesor de muro en Elevación Eje.

Por cada ViewSection creada:
  - Muros visibles con Material for Model Behavior = Concrete
  - Paralelos al plano de la vista (|Orientation · ViewDirection| ≈ 1)
  - IndependentTag: EST_A_WALL TAG_ELEVACION_MHA / Espesor Muro
  - Cabeza en el centro de la cara vista del bbox (hacia ViewDirection)
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import Wall, XYZ

from armado_muros_wall_tags_rebase import (
    _crear_wall_tag,
    _default_snap_espesor,
    resolve_wall_espesor_tag_symbol,
)
from elevacion_eje_collect import (
    _eid_int,
    _vector_unitario,
    muro_paralelo_a_vista,
    recoger_concrete_en_vista,
)


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _esquinas_bbox(bb):
    mn, mx = bb.Min, bb.Max
    return (
        XYZ(mn.X, mn.Y, mn.Z),
        XYZ(mx.X, mn.Y, mn.Z),
        XYZ(mn.X, mx.Y, mn.Z),
        XYZ(mx.X, mx.Y, mn.Z),
        XYZ(mn.X, mn.Y, mx.Z),
        XYZ(mx.X, mn.Y, mx.Z),
        XYZ(mn.X, mx.Y, mx.Z),
        XYZ(mx.X, mx.Y, mx.Z),
    )


def _punto_etiqueta_muro(wall, view):
    """
    Centro de la cara del bbox orientada hacia la vista (máx. ViewDirection).

    Fallback: centro del bbox / punto medio de la LocationCurve.
    """
    if wall is None:
        return None
    bb = None
    try:
        bb = wall.get_BoundingBox(view)
    except Exception:
        bb = None
    if bb is None:
        try:
            bb = wall.get_BoundingBox(None)
        except Exception:
            bb = None

    vd = None
    try:
        if view is not None:
            vd = _vector_unitario(view.ViewDirection)
    except Exception:
        vd = None

    if bb is not None and vd is not None:
        try:
            corners = _esquinas_bbox(bb)
            dmax = max(float(p.DotProduct(vd)) for p in corners)
            face = [
                p for p in corners
                if abs(float(p.DotProduct(vd)) - dmax) <= 1e-6
            ]
            if face:
                n = float(len(face))
                return XYZ(
                    sum(p.X for p in face) / n,
                    sum(p.Y for p in face) / n,
                    sum(p.Z for p in face) / n,
                )
        except Exception:
            pass

    if bb is not None:
        try:
            return XYZ(
                (bb.Min.X + bb.Max.X) * 0.5,
                (bb.Min.Y + bb.Max.Y) * 0.5,
                (bb.Min.Z + bb.Max.Z) * 0.5,
            )
        except Exception:
            pass
    try:
        loc = wall.Location
        curve = loc.Curve if loc is not None else None
        if curve is not None and curve.IsBound:
            return curve.Evaluate(0.5, True)
    except Exception:
        pass
    return None


def recoger_muros_concrete_paralelos(document, view, material_cache=None):
    """Muros Concrete visibles en ``view`` y paralelos a su plano."""
    packed = recoger_concrete_en_vista(document, view, material_cache)
    return list(packed.get(u"muros") or [])


def etiquetar_muros_concrete_paralelos(
    document,
    view,
    symbol=None,
    muros=None,
    tagged_hosts=None,
):
    """
    Crea etiquetas Espesor Muro en ``view``.

    Debe llamarse dentro de una Transaction abierta (crop ya activo).

    Args:
        muros: lista precargada (opcional)
        tagged_hosts: set de host Id int ya etiquetados en la vista

    Returns:
        dict ``n_ok``, ``n_skip``, ``n_fail``, ``error_symbol``
    """
    result = {
        u"n_ok": 0,
        u"n_skip": 0,
        u"n_fail": 0,
        u"error_symbol": None,
    }
    if document is None or view is None:
        result[u"n_fail"] = 1
        return result

    if symbol is None:
        symbol, err = resolve_wall_espesor_tag_symbol(document)
        if symbol is None:
            result[u"error_symbol"] = err or u"Tipo de etiqueta no encontrado."
            result[u"n_fail"] = 1
            return result
    try:
        type_id = symbol.Id
    except Exception:
        type_id = None
    if type_id is None:
        result[u"error_symbol"] = u"Id de tipo de etiqueta inválido."
        result[u"n_fail"] = 1
        return result

    snap = _default_snap_espesor(type_id)
    if muros is None:
        muros = recoger_muros_concrete_paralelos(document, view)
    if not muros:
        return result

    already = tagged_hosts if tagged_hosts is not None else set()

    for wall in muros:
        if wall is None:
            continue
        try:
            if not isinstance(wall, Wall):
                continue
        except Exception:
            continue
        wid = _eid_int(getattr(wall, u"Id", None))
        if wid is not None and wid in already:
            result[u"n_skip"] = int(result[u"n_skip"]) + 1
            continue
        if not muro_paralelo_a_vista(wall, view):
            continue
        head = _punto_etiqueta_muro(wall, view)
        if head is None:
            result[u"n_fail"] = int(result[u"n_fail"]) + 1
            continue
        tag, err = _crear_wall_tag(document, view, wall, snap, head)
        if tag is None:
            result[u"n_fail"] = int(result[u"n_fail"]) + 1
        else:
            result[u"n_ok"] = int(result[u"n_ok"]) + 1
            if wid is not None:
                already.add(wid)
    return result
