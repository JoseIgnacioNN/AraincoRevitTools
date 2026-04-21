# -*- coding: utf-8 -*-
"""
Evaluación **exhaustiva** de extremos de curvas acotadas frente a obstáculos por volumen.

- Por cada ``Curve`` con ``IsBound``: ``GetEndPoint(0)`` (Start) y ``GetEndPoint(1)`` (End).
- Cada punto se contrasta con **todos** los elementos obstáculo y **todos** sus ``Solid``;
  no se acorta el bucle si Start ya colisionó: End se evalúa completo igualmente.
- Contención en volumen: Revit **no** documenta ``Solid.ContainsPoint`` de forma estable;
  se usa ``IntersectWithCurve`` sobre un micro-tramo a través del punto
  (:class:`SolidCurveIntersectionMode.CurveSegmentsInside`). Si la API aportara
  ``ContainsPoint``, se intenta primero.

Salida: por curva, conjuntos de IDs de elementos donde cada extremo cae dentro del volumen
y etiqueta ``Start`` / ``End`` / ``Both`` / ``None``.
"""

from Autodesk.Revit.DB import (
    Line,
    SolidCurveIntersectionMode,
    SolidCurveIntersectionOptions,
    XYZ,
)

from geometria_colision_vigas import obtener_solidos_elemento

# Longitud semientorno del micro-segmento (pies) ~0,33 mm; solo para prueba punto–sólido.
_MICRO_HALF_SEG_FT = 1.0 / 3048.0


def _eid_int(element_id):
    if element_id is None:
        return None
    try:
        return int(element_id.Value)
    except Exception:
        try:
            return int(element_id.IntegerValue)
        except Exception:
            return None


def _punto_en_volumen_solido(solid, pt):
    """
    True si ``pt`` está en el interior del volumen (aprox.). Preferencia: ``ContainsPoint``
    si existe en esta versión de API; si no, micro-segmento + ``IntersectWithCurve``.
    """
    if solid is None or pt is None:
        return False
    try:
        if float(solid.Volume) < 1e-12:
            return False
    except Exception:
        return False
    try:
        fn = getattr(solid, "ContainsPoint", None)
        if callable(fn):
            try:
                return bool(fn(pt))
            except TypeError:
                try:
                    return bool(fn(pt, 0.0))
                except Exception:
                    pass
            except Exception:
                pass
    except Exception:
        pass
    try:
        d = float(_MICRO_HALF_SEG_FT)
        p_a = XYZ(pt.X - d, pt.Y, pt.Z)
        p_b = XYZ(pt.X + d, pt.Y, pt.Z)
        micro = Line.CreateBound(p_a, p_b)
        scio = SolidCurveIntersectionOptions()
        try:
            scio.ResultType = SolidCurveIntersectionMode.CurveSegmentsInside
        except Exception:
            pass
        sci = solid.IntersectWithCurve(micro, scio)
        if sci is None:
            return False
        return int(sci.SegmentCount) > 0
    except Exception:
        return False


def curvas_acotadas_desde_ids_elementos(document, element_ids):
    """
    Obtiene ``Curve`` acotadas desde ``ModelCurve`` / ``DetailCurve`` u otro elemento
    con ``GeometryCurve`` acotada.
    """
    out = []
    if document is None or not element_ids:
        return out
    for eid in element_ids:
        try:
            el = document.GetElement(eid)
        except Exception:
            el = None
        if el is None:
            continue
        crv = getattr(el, "GeometryCurve", None)
        if crv is None:
            continue
        try:
            if crv.IsBound:
                out.append(crv)
        except Exception:
            continue
    return out


def evaluar_extremos_curvas_vs_obstaculos(
    document,
    curvas,
    obstacle_element_ids,
    geometry_options=None,
    incluir_detalle_por_solido=False,
):
    """
    Evalúa Start y End de cada curva contra **todos** los obstáculos y **todos** sus sólidos.

    Args:
        document: ``Document``
        curvas: iterable de ``Curve`` con ``IsBound`` True
        obstacle_element_ids: iterable de ``ElementId`` (elementos obstáculo)
        geometry_options: opcional, ``Options`` para ``get_Geometry`` de obstáculos
        incluir_detalle_por_solido: si True, lista ``detail`` por cada pareja punto–sólido

    Returns:
        ``list`` de dicts:
            - ``curve_index`` (int)
            - ``start_colliding_element_ids`` (list int, únicos)
            - ``end_colliding_element_ids`` (list int, únicos)
            - ``endpoints_colliding``: ``u\"Start\"``, ``u\"End\"``, ``u\"Both\"``, ``u\"None\"``
        Opcional por entrada: ``detail`` list de ``{endpoint, element_id, solid_index}``
    """
    resultados = []
    if document is None:
        return resultados
    obs_ids = list(obstacle_element_ids or [])
    curvas_list = [c for c in (curvas or []) if c is not None]
    for ci, curva in enumerate(curvas_list):
        try:
            if not curva.IsBound:
                continue
            p_start = curva.GetEndPoint(0)
            p_end = curva.GetEndPoint(1)
        except Exception:
            continue

        start_ids = set()
        end_ids = set()
        detail = [] if incluir_detalle_por_solido else None

        for oid in obs_ids:
            try:
                obs_el = document.GetElement(oid)
            except Exception:
                obs_el = None
            if obs_el is None:
                continue
            solids = obtener_solidos_elemento(obs_el, geometry_options)
            for si, solid in enumerate(solids or []):
                hit_start = False
                hit_end = False
                try:
                    hit_start = _punto_en_volumen_solido(solid, p_start)
                except Exception:
                    hit_start = False
                try:
                    hit_end = _punto_en_volumen_solido(solid, p_end)
                except Exception:
                    hit_end = False

                if incluir_detalle_por_solido:
                    if hit_start:
                        detail.append(
                            {
                                u"endpoint": u"Start",
                                u"element_id": _eid_int(oid),
                                u"solid_index": si,
                            }
                        )
                    if hit_end:
                        detail.append(
                            {
                                u"endpoint": u"End",
                                u"element_id": _eid_int(oid),
                                u"solid_index": si,
                            }
                        )

                if hit_start:
                    ik = _eid_int(oid)
                    if ik is not None:
                        start_ids.add(ik)
                if hit_end:
                    ik = _eid_int(oid)
                    if ik is not None:
                        end_ids.add(ik)

        if start_ids and end_ids:
            label = u"Both"
        elif start_ids:
            label = u"Start"
        elif end_ids:
            label = u"End"
        else:
            label = u"None"

        entry = {
            u"curve_index": ci,
            u"start_colliding_element_ids": sorted(start_ids),
            u"end_colliding_element_ids": sorted(end_ids),
            u"endpoints_colliding": label,
        }
        if detail is not None:
            entry[u"detail"] = detail
        resultados.append(entry)

    return resultados


def resumen_colision_curvas(resultados):
    """
    A partir de la lista devuelta por :func:`evaluar_extremos_curvas_vs_obstaculos`,
    devuelve índices de curva con colisión solo Start, solo End, o Both.
    """
    solo_start = []
    solo_end = []
    ambos = []
    ninguno = []
    for r in resultados or []:
        try:
            idx = int(r.get(u"curve_index", -1))
            lbl = r.get(u"endpoints_colliding")
        except Exception:
            continue
        if lbl == u"Both":
            ambos.append(idx)
        elif lbl == u"Start":
            solo_start.append(idx)
        elif lbl == u"End":
            solo_end.append(idx)
        else:
            ninguno.append(idx)
    return {
        u"curve_indices_start_only": solo_start,
        u"curve_indices_end_only": solo_end,
        u"curve_indices_both_endpoints": ambos,
        u"curve_indices_none": ninguno,
    }
