# -*- coding: utf-8 -*-
"""
Quiebre en L en el leader de una etiqueta (IndependentTag) seleccionada.

Geometría tipo **segunda imagen** (L de 90° en el plano de la vista):
- Tramo 1: vertical (solo ``View.UpDirection``) desde el anclaje del leader.
- Tramo 2: horizontal (solo ``View.RightDirection``) hacia el texto.
- El encuentro con el texto debe caer en el **borde que mira al ancla** y a **altura media**
  del bloque (≈ línea central de un texto de 3 líneas), no en ``TagHeadPosition`` crudo
  (suele ser esquina de familia).

Con ``LEADER_TARGET_USE_TAG_BBOX = True`` (por defecto), el punto objetivo se calcula con el
``BoundingBox`` de la etiqueta en la vista: borde izquierdo/derecho según hacia dónde esté
el texto respecto al ancla, y altura media en pantalla — como un leader “Free” que termina
al inicio de la zona del texto central.

Mejoras respecto a la v1:
- Vectores proyectados al plano de la vista (⟂ ``ViewDirection``), igual que en herramientas
  del propio proyecto (p. ej. criterio ortogonal en ``enfierrado_shaft_hashtag``).
- Tramo vertical mínimo y “estante” horizontal mínimo (evita leaders casi degenerados).
- Si ``ADJUST_TAG_HEAD_TO_FIT_L`` es True, recalcula ``TagHeadPosition`` para que el tramo
  horizontal tenga longitud coherente (Revit a veces deja la cabecera desplazada si solo se
  asigna el codo).
- Opcional: segunda pasada tras ``Regenerate`` para estabilizar.

Ejecutable en RevitPythonShell (RPS): selecciona uno o varios IndependentTag con leader y ejecuta.

Requisitos:
- Revit 2022+ recomendado (API ``GetTaggedReferences`` / ``SetLeaderElbow`` por Reference).
- La etiqueta debe tener línea de llamada visible (``HasLeader``).
"""

from __future__ import division

import clr
import math

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import IndependentTag, Reference, Transaction, TransactionStatus, XYZ
from Autodesk.Revit.UI import TaskDialog


# Tramos mínimos en el plano de la vista (evita codos invisibles o rechazados por Revit)
MIN_VERTICAL_LEG_MM = 3.0
MIN_HORIZONTAL_SHELF_MM = 3.0

# Si True, ajusta TagHeadPosition al final del estante horizontal (recomendado; corrige saltos de Revit)
ADJUST_TAG_HEAD_TO_FIT_L = True

# Cuando anclaje y cabecera están casi a la misma “altura” en pantalla (t ≈ 0), sentido del
# tramo vertical: -1 suele coincidir con “hacia abajo” en alzados/planos (revisa en tu vista).
DEFAULT_VERTICAL_SIGN_WHEN_LEVEL = -1.0

# Segunda iteración tras Regenerate (mejora estabilidad del API)
LEADER_L_SECOND_PASS = True

# Objetivo del L según bbox del tag en vista (recomendado para coincidir con la 2.ª imagen)
LEADER_TARGET_USE_TAG_BBOX = True


def _mm_to_ft(mm):
    return float(mm) / 304.8


def _xyz_dot(a, b):
    return (
        float(a.X) * float(b.X)
        + float(a.Y) * float(b.Y)
        + float(a.Z) * float(b.Z)
    )


def _xyz_sub(a, b):
    return XYZ(float(a.X) - float(b.X), float(a.Y) - float(b.Y), float(a.Z) - float(b.Z))


def _xyz_add(a, b):
    return XYZ(float(a.X) + float(b.X), float(a.Y) + float(b.Y), float(a.Z) + float(b.Z))


def _xyz_scale(v, s):
    return XYZ(float(v.X) * s, float(v.Y) * s, float(v.Z) * s)


def _unit(v):
    if v is None:
        return None
    try:
        ln = float(v.GetLength())
    except Exception:
        return None
    if ln < 1e-12:
        return None
    try:
        return v.Normalize()
    except Exception:
        return XYZ(float(v.X) / ln, float(v.Y) / ln, float(v.Z) / ln)


def _view_axes(view):
    """Direcciones unitarias: normal (hacia el observador), derecha, arriba en pantalla."""
    if view is None:
        return None, None, None
    n = _unit(view.ViewDirection)
    r = _unit(view.RightDirection)
    u = _unit(view.UpDirection)
    return n, r, u


def _project_onto_plane_through_origin(v, plane_normal):
    """Componente de ``v`` en el plano perpendicular a ``plane_normal`` (unitario)."""
    if v is None:
        return None
    if plane_normal is None:
        return v
    nv = _xyz_dot(v, plane_normal)
    return _xyz_sub(v, _xyz_scale(plane_normal, nv))


def _bbox_corners(bb):
    """Lista de las 8 esquinas de un ``BoundingBox``."""
    if bb is None or bb.Min is None or bb.Max is None:
        return []
    mn, mx = bb.Min, bb.Max
    return [
        XYZ(mn.X, mn.Y, mn.Z),
        XYZ(mx.X, mn.Y, mn.Z),
        XYZ(mn.X, mx.Y, mn.Z),
        XYZ(mx.X, mx.Y, mn.Z),
        XYZ(mn.X, mn.Y, mx.Z),
        XYZ(mx.X, mn.Y, mx.Z),
        XYZ(mn.X, mx.Y, mx.Z),
        XYZ(mx.X, mx.Y, mx.Z),
    ]


def _leader_target_second_image(
    tag, view, leader_end, head_fallback, use_bbox
):
    """
    Punto objetivo para un L como en la 2.ª imagen: borde del bloque de texto que mira al
    ancla + altura media (≈ línea central). Si ``use_bbox`` es False o no hay bbox, usa
    ``head_fallback``.
    """
    if not use_bbox or tag is None or view is None:
        return head_fallback
    if leader_end is None or head_fallback is None:
        return head_fallback
    bb = None
    try:
        bb = tag.get_BoundingBox(view)
    except Exception:
        bb = None
    if bb is None:
        try:
            bb = tag.get_BoundingBox(None)
        except Exception:
            bb = None
    corners = _bbox_corners(bb)
    if not corners:
        return head_fallback
    n, r, u = _view_axes(view)
    if r is None or u is None:
        return head_fallback
    r_dots = []
    u_dots = []
    for c in corners:
        dv = _xyz_sub(c, leader_end)
        r_dots.append(_xyz_dot(dv, r))
        u_dots.append(_xyz_dot(dv, u))
    try:
        to_tag = _xyz_dot(_xyz_sub(head_fallback, leader_end), r)
    except Exception:
        to_tag = 0.0
    if to_tag > 1e-9:
        r_edge = min(r_dots)
    elif to_tag < -1e-9:
        r_edge = max(r_dots)
    else:
        r_edge = min(r_dots)
    u_mid = 0.5 * (min(u_dots) + max(u_dots))
    v = _xyz_add(_xyz_scale(r, float(r_edge)), _xyz_scale(u, float(u_mid)))
    if n is not None:
        try:
            gamma = _xyz_dot(_xyz_sub(head_fallback, leader_end), n)
            v = _xyz_add(v, _xyz_scale(n, float(gamma)))
        except Exception:
            pass
    return _xyz_add(leader_end, v)


def _get_doc_uidoc():
    try:
        return doc, uidoc
    except NameError:
        u = __revit__.ActiveUIDocument
        return u.Document, u


def _refs_count(refs):
    if refs is None:
        return 0
    try:
        return int(refs.Count)
    except Exception:
        try:
            return len(refs)
        except Exception:
            return 0


def _tagged_reference_for_leader(tag, document):
    """Una Reference válida para GetLeaderEnd / SetLeaderElbow."""
    try:
        refs = tag.GetTaggedReferences()
    except Exception:
        refs = None
    n = _refs_count(refs)
    if n > 0:
        try:
            return refs[0]
        except Exception:
            try:
                return refs.Item[0]
            except Exception:
                pass
    try:
        ids = tag.GetTaggedLocalElementIds()
    except Exception:
        ids = None
    if ids is not None:
        try:
            cnt = int(ids.Count)
        except Exception:
            try:
                cnt = len(ids)
            except Exception:
                cnt = 0
        if cnt > 0:
            try:
                eid = ids[0]
            except Exception:
                eid = None
            if eid is not None:
                try:
                    el = document.GetElement(eid)
                except Exception:
                    el = None
                if el is not None:
                    try:
                        return Reference(el)
                    except Exception:
                        pass
    return None


def _leader_end_or_fallback(tag, ref_tagged, document):
    try:
        return tag.GetLeaderEnd(ref_tagged)
    except Exception:
        pass
    try:
        eid = ref_tagged.ElementId
        el = document.GetElement(eid)
    except Exception:
        el = None
    if el is None:
        return None
    try:
        loc = el.Location
    except Exception:
        loc = None
    if loc is not None:
        try:
            p = loc.Point
            if p is not None:
                return p
        except Exception:
            pass
        try:
            crv = loc.Curve
            if crv is not None:
                p = crv.Evaluate(0.5, True)
                if p is not None:
                    return p
        except Exception:
            pass
    try:
        bb = el.get_BoundingBox(None)
        if bb is not None and bb.Min is not None and bb.Max is not None:
            return XYZ(
                (float(bb.Min.X) + float(bb.Max.X)) * 0.5,
                (float(bb.Min.Y) + float(bb.Max.Y)) * 0.5,
                (float(bb.Min.Z) + float(bb.Max.Z)) * 0.5,
            )
    except Exception:
        pass
    return None


def compute_l_elbow_and_head(
    view,
    leader_end,
    geometry_target,
    min_vertical_ft,
    min_horizontal_ft,
    default_t_sign,
    depth_reference=None,
):
    """
    Codo en L en el plano de la vista: ``elbow = end + t·û``, cabecera coherente
    ``head_new = elbow + s·r̂ + n̂·(ref−elbow)·n̂``.

    ``geometry_target`` define el L (p. ej. borde medio del bbox del tag).
    ``depth_reference`` (opcional) conserva la profundidad respecto a la vista al ajustar
    ``TagHeadPosition``; por defecto es ``geometry_target``.
    """
    if view is None or leader_end is None or geometry_target is None:
        return None, None
    n, r, u = _view_axes(view)
    if r is None or u is None:
        return None, None
    ref_depth = geometry_target if depth_reference is None else depth_reference
    d = _xyz_sub(geometry_target, leader_end)
    d_in = _project_onto_plane_through_origin(d, n)
    if d_in is None:
        return None, None
    t_raw = _xyz_dot(d_in, u)
    s_raw = _xyz_dot(d_in, r)
    t_use = float(t_raw)
    if abs(t_use) < float(min_vertical_ft):
        if abs(t_raw) > 1e-9:
            t_use = math.copysign(float(min_vertical_ft), t_raw)
        else:
            t_use = math.copysign(float(min_vertical_ft), float(default_t_sign))
    s_use = float(s_raw)
    if abs(s_use) < float(min_horizontal_ft):
        if abs(s_raw) > 1e-9:
            s_use = math.copysign(float(min_horizontal_ft), s_raw)
        else:
            s_use = float(min_horizontal_ft)
    elbow = _xyz_add(leader_end, _xyz_scale(u, t_use))
    head_new = _xyz_add(elbow, _xyz_scale(r, s_use))
    if n is not None:
        try:
            gamma = _xyz_dot(_xyz_sub(ref_depth, elbow), n)
            head_new = _xyz_add(head_new, _xyz_scale(n, gamma))
        except Exception:
            pass
    return elbow, head_new


def _apply_quiebre_once(
    tag,
    document,
    view,
    ref_tagged,
    min_v_ft,
    min_h_ft,
    default_t_sign,
    adjust_head,
    use_bbox_target,
):
    """Una pasada: lee puntos, asigna codo y opcionalmente cabecera."""
    try:
        head = tag.TagHeadPosition
    except Exception:
        head = None
    if head is None:
        return False, u"No se pudo leer TagHeadPosition."
    end = _leader_end_or_fallback(tag, ref_tagged, document)
    if end is None:
        return False, u"No se pudo obtener LeaderEnd ni un punto de respaldo en el elemento."
    target = _leader_target_second_image(tag, view, end, head, use_bbox_target)
    elbow, head_new = compute_l_elbow_and_head(
        view,
        end,
        target,
        min_v_ft,
        min_h_ft,
        default_t_sign,
        depth_reference=head,
    )
    if elbow is None:
        return False, u"No se pudo calcular el codo (ejes de vista inválidos)."
    try:
        if elbow.DistanceTo(end) < 1e-4:
            return False, u"El codo coincide con el anclaje; aumenta MIN_VERTICAL_LEG_MM."
    except Exception:
        pass
    try:
        tag.SetLeaderElbow(ref_tagged, elbow)
    except Exception as ex:
        return False, u"SetLeaderElbow falló: {0}".format(ex)
    if adjust_head and head_new is not None:
        for _ in (0, 1):
            try:
                tag.TagHeadPosition = head_new
            except Exception as ex:
                return False, u"TagHeadPosition falló: {0}".format(ex)
            try:
                document.Regenerate()
            except Exception:
                pass
    else:
        try:
            document.Regenerate()
        except Exception:
            pass
    return True, u""


def apply_leader_quiebre_l(
    tag,
    document,
    view,
    min_vertical_mm=None,
    min_horizontal_mm=None,
    adjust_head=None,
    default_vertical_sign=None,
    use_tag_bbox_target=None,
):
    """
    Aplica el quiebre en L al primer leader del tag.

    Args:
        min_vertical_mm: tramo mínimo en dirección Up (pantalla).
        min_horizontal_mm: estante mínimo en dirección Right.
        adjust_head: si True, ajusta ``TagHeadPosition`` al cierre del L (recomendado).
        default_vertical_sign: si ancla y cabecera están alineados en horizontal, sentido del
            tramo vertical (+1 / -1 según ``View.UpDirection``).
        use_tag_bbox_target: si True, el L apunta al borde medio del bbox del tag en vista
            (estilo 2.ª imagen). Si None, usa ``LEADER_TARGET_USE_TAG_BBOX``.

    Returns:
        (ok: bool, message: unicode)
    """
    if tag is None or not isinstance(tag, IndependentTag):
        return False, u"No es un IndependentTag válido."
    try:
        if not bool(tag.HasLeader):
            return False, u"La etiqueta no tiene línea de llamada (HasLeader = False)."
    except Exception:
        return False, u"No se pudo leer HasLeader."
    ref_tagged = _tagged_reference_for_leader(tag, document)
    if ref_tagged is None:
        return False, u"No se pudo obtener una Reference del elemento etiquetado."
    min_v = _mm_to_ft(
        float(min_vertical_mm if min_vertical_mm is not None else MIN_VERTICAL_LEG_MM)
    )
    min_h = _mm_to_ft(
        float(
            min_horizontal_mm
            if min_horizontal_mm is not None
            else MIN_HORIZONTAL_SHELF_MM
        )
    )
    adj = ADJUST_TAG_HEAD_TO_FIT_L if adjust_head is None else bool(adjust_head)
    dvs = (
        DEFAULT_VERTICAL_SIGN_WHEN_LEVEL
        if default_vertical_sign is None
        else float(default_vertical_sign)
    )
    use_bbox = (
        LEADER_TARGET_USE_TAG_BBOX
        if use_tag_bbox_target is None
        else bool(use_tag_bbox_target)
    )
    try:
        document.Regenerate()
    except Exception:
        pass
    ok, msg = _apply_quiebre_once(
        tag, document, view, ref_tagged, min_v, min_h, dvs, adj, use_bbox
    )
    if not ok:
        return ok, msg
    if LEADER_L_SECOND_PASS:
        try:
            document.Regenerate()
        except Exception:
            pass
        ok2, msg2 = _apply_quiebre_once(
            tag, document, view, ref_tagged, min_v, min_h, dvs, adj, use_bbox
        )
        if not ok2:
            return ok2, msg2
    return True, u"Quiebre aplicado."


def main():
    document, uidoc = _get_doc_uidoc()
    view = uidoc.ActiveView
    sel = uidoc.Selection.GetElementIds()
    if sel is None or sel.Count < 1:
        TaskDialog.Show(u"Leader L", u"Selecciona una o más etiquetas (IndependentTag).")
        return
    tags = []
    for eid in sel:
        try:
            el = document.GetElement(eid)
        except Exception:
            el = None
        if el is not None and isinstance(el, IndependentTag):
            tags.append(el)
    if not tags:
        TaskDialog.Show(u"Leader L", u"Ninguna selección es un IndependentTag.")
        return
    ok_n = 0
    errs = []
    t = Transaction(document, u"Quiebre leader etiqueta L")
    t.Start()
    try:
        for tg in tags:
            ok, msg = apply_leader_quiebre_l(tg, document, view)
            if ok:
                ok_n += 1
            else:
                try:
                    tid = int(tg.Id.IntegerValue)
                except Exception:
                    tid = -1
                errs.append(u"Id {0}: {1}".format(tid, msg))
        t.Commit()
    except Exception as ex:
        try:
            if t.GetStatus() == TransactionStatus.Started:
                t.RollBack()
        except Exception:
            pass
        TaskDialog.Show(u"Leader L", u"Error: {0}".format(ex))
        return
    lines = [u"Procesadas: {0} / {1}.".format(ok_n, len(tags))]
    if errs:
        lines.append(u"")
        lines.extend(errs[:8])
        if len(errs) > 8:
            lines.append(u"...")
    TaskDialog.Show(u"Leader L", u"\n".join(lines))


if __name__ == "__main__":
    main()
