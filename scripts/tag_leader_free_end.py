# -*- coding: utf-8 -*-
"""
Arainco: Leader Free End en etiquetas (IndependentTag).

1. Usa etiquetas preseleccionadas o pide seleccionar una o varias.
2. Activa la línea de llamada (HasLeader).
3. Aplica extremo libre (LeaderEndCondition.Free).
4. Alinea el leader en horizontal con el tag (misma altura en vista
   que la 1.ª línea de texto; extremo libre hacia el elemento).

Revit 2024+ | pyRevit (importable via run).
"""

from __future__ import division, print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    IndependentTag,
    LeaderEndCondition,
    Reference,
    Transaction,
    TransactionStatus,
    XYZ,
)
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

_TOOL_TITLE = u"Arainco: Tag Leader Free End"
_TXN_NAME = u"Arainco: Tag Leader Free End"
_PROMPT = (
    u"Selecciona una o varias etiquetas (tags). "
    u"Finish para continuar / Esc cancela."
)

# Longitud mínima del leader horizontal si el extremo queda casi sobre el texto (mm modelo).
_MIN_LEADER_HORIZ_MM = 15.0


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _eid_int(eid):
    try:
        return int(eid.IntegerValue)
    except Exception:
        try:
            return int(eid.Value)
        except Exception:
            return -1


def _mm_to_ft(mm):
    return float(mm) / 304.8


def _mostrar_aviso(uiapp, instruction, content=u"", ok_text=u"Entendido"):
    try:
        from bimtools_instruction_dialog import show_message_dialog
        from revit_wpf_window_position import revit_main_hwnd

        hwnd = revit_main_hwnd(uiapp) if uiapp is not None else None
        show_message_dialog(
            _TOOL_TITLE,
            instruction=_as_unicode(instruction),
            content=_as_unicode(content),
            ok_text=_as_unicode(ok_text),
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass
    try:
        body = _as_unicode(instruction)
        extra = _as_unicode(content).strip()
        if extra:
            body = body + u"\n\n" + extra
        TaskDialog.Show(_TOOL_TITLE, body)
    except Exception:
        pass


class _IndependentTagFilter(ISelectionFilter):
    def AllowElement(self, element):
        return element is not None and isinstance(element, IndependentTag)

    def AllowReference(self, reference, point):
        return False


def _collect_tags_from_ids(doc, eids):
    tags = []
    seen = set()
    if eids is None:
        return tags
    for eid in eids:
        try:
            el = doc.GetElement(eid)
        except Exception:
            el = None
        if el is None or not isinstance(el, IndependentTag):
            continue
        iid = _eid_int(el.Id)
        if iid in seen:
            continue
        seen.add(iid)
        tags.append(el)
    return tags


def _get_preselected_tags(uidoc, doc):
    try:
        eids = uidoc.Selection.GetElementIds()
    except Exception:
        return []
    return _collect_tags_from_ids(doc, eids)


def _pick_tags(uidoc):
    refs = list(
        uidoc.Selection.PickObjects(
            ObjectType.Element, _IndependentTagFilter(), _PROMPT
        )
    )
    tags = []
    seen = set()
    doc = uidoc.Document
    for pref in refs:
        try:
            el = doc.GetElement(pref.ElementId)
        except Exception:
            el = None
        if el is None or not isinstance(el, IndependentTag):
            continue
        iid = _eid_int(el.Id)
        if iid in seen:
            continue
        seen.add(iid)
        tags.append(el)
    return tags


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
    if view is None:
        return None, None, None
    return (
        _unit(view.ViewDirection),
        _unit(view.RightDirection),
        _unit(view.UpDirection),
    )


def _project_onto_plane(v, plane_normal):
    if v is None:
        return None
    if plane_normal is None:
        return v
    return _xyz_sub(v, _xyz_scale(plane_normal, _xyz_dot(v, plane_normal)))


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


def _tagged_references(tag, document):
    """Lista de Reference del tag (multihost o respaldo)."""
    out = []
    try:
        refs = tag.GetTaggedReferences()
    except Exception:
        refs = None
    n = _refs_count(refs)
    if n > 0:
        for i in range(n):
            try:
                out.append(refs[i])
            except Exception:
                try:
                    out.append(refs.Item[i])
                except Exception:
                    pass
        if out:
            return out
    try:
        r = tag.GetTaggedReference()
        if r is not None:
            return [r]
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
        for i in range(cnt):
            try:
                eid = ids[i]
            except Exception:
                eid = None
            if eid is None:
                continue
            try:
                el = document.GetElement(eid)
            except Exception:
                el = None
            if el is None:
                continue
            try:
                out.append(Reference(el))
            except Exception:
                pass
    return out


def _bbox_corners(bb):
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


def _tag_bbox(tag, view):
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
    return bb


def _attach_point_first_line(tag, view, leader_end, head_fallback):
    """
    Punto de encuentro visual del leader con el tag: borde del bbox que mira
    al extremo libre, a la altura media de la 1.ª línea de texto (≈ tercio superior).
    """
    if head_fallback is None:
        return None
    if tag is None or view is None or leader_end is None:
        return head_fallback
    n, r, u = _view_axes(view)
    if r is None or u is None:
        return head_fallback
    corners = _bbox_corners(_tag_bbox(tag, view))
    if not corners:
        return head_fallback

    origin = head_fallback
    r_dots = [_xyz_dot(_xyz_sub(c, origin), r) for c in corners]
    u_dots = [_xyz_dot(_xyz_sub(c, origin), u) for c in corners]
    u_min, u_max = min(u_dots), max(u_dots)
    u_span = float(u_max - u_min)
    # Centro de la primera línea (arriba) asumiendo ~3 líneas de altura similar.
    if u_span > 1e-9:
        u_first = float(u_max) - (u_span / 6.0)
    else:
        u_first = 0.0

    to_end = _xyz_dot(_xyz_sub(leader_end, origin), r)
    # Borde del texto hacia el extremo libre.
    if to_end < -1e-9:
        r_edge = min(r_dots)
    elif to_end > 1e-9:
        r_edge = max(r_dots)
    else:
        r_edge = min(r_dots)

    attach = _xyz_add(origin, _xyz_add(_xyz_scale(r, float(r_edge)), _xyz_scale(u, float(u_first))))
    if n is not None:
        try:
            gamma = _xyz_dot(_xyz_sub(head_fallback, attach), n)
            attach = _xyz_add(attach, _xyz_scale(n, float(gamma)))
        except Exception:
            pass
    return attach


def _leader_end_or_fallback(tag, ref_tagged, document, view):
    try:
        return tag.GetLeaderEnd(ref_tagged)
    except Exception:
        pass
    try:
        el = document.GetElement(ref_tagged.ElementId)
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
                return crv.Evaluate(0.5, True)
        except Exception:
            pass
    try:
        bb = el.get_BoundingBox(view)
        if bb is None:
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


def _horizontal_free_end(view, attach, current_end):
    """
    Extremo libre a la misma altura (Up) que ``attach``, conservando el
    desplazamiento horizontal (Right) hacia el extremo actual → leader recto horizontal.
    """
    if view is None or attach is None or current_end is None:
        return None
    n, r, u = _view_axes(view)
    if r is None or u is None:
        return None
    d = _project_onto_plane(_xyz_sub(current_end, attach), n)
    if d is None:
        return None
    s = float(_xyz_dot(d, r))
    min_h = _mm_to_ft(_MIN_LEADER_HORIZ_MM)
    if abs(s) < min_h:
        if abs(s) > 1e-9:
            s = (1.0 if s > 0.0 else -1.0) * min_h
        else:
            # Por defecto hacia la izquierda del texto (como en la imagen de referencia).
            s = -min_h
    new_end = _xyz_add(attach, _xyz_scale(r, s))
    if n is not None:
        try:
            gamma = _xyz_dot(_xyz_sub(current_end, attach), n)
            new_end = _xyz_add(new_end, _xyz_scale(n, float(gamma)))
        except Exception:
            pass
    return new_end


def _straighten_elbow_horizontal(tag, ref_tagged, attach, free_end):
    """Si hay codo, lo coloca en el segmento horizontal para que se vea recto."""
    try:
        if not bool(tag.HasLeaderElbow(ref_tagged)):
            return
    except Exception:
        # Sin codo o API distinta: no forzar.
        return
    if attach is None or free_end is None:
        return
    mid = XYZ(
        (float(attach.X) + float(free_end.X)) * 0.5,
        (float(attach.Y) + float(free_end.Y)) * 0.5,
        (float(attach.Z) + float(free_end.Z)) * 0.5,
    )
    try:
        tag.SetLeaderElbow(ref_tagged, mid)
    except Exception:
        pass


def _align_leaders_horizontal(tag, document, view):
    """Alinea cada leader del tag en horizontal respecto al texto. (ok, err)."""
    try:
        head = tag.TagHeadPosition
    except Exception:
        head = None
    if head is None:
        return False, u"No se pudo leer TagHeadPosition"

    refs = _tagged_references(tag, document)
    if not refs:
        return False, u"Sin referencia etiquetada"

    aligned = 0
    last_err = None
    for ref in refs:
        end = _leader_end_or_fallback(tag, ref, document, view)
        if end is None:
            last_err = u"sin LeaderEnd"
            continue
        attach = _attach_point_first_line(tag, view, end, head)
        if attach is None:
            attach = head
        new_end = _horizontal_free_end(view, attach, end)
        if new_end is None:
            last_err = u"no se pudo calcular extremo horizontal"
            continue
        try:
            tag.SetLeaderEnd(ref, new_end)
        except Exception as ex:
            last_err = u"SetLeaderEnd: {0}".format(_as_unicode(ex))
            continue
        _straighten_elbow_horizontal(tag, ref, attach, new_end)
        aligned += 1

    if aligned < 1:
        return False, last_err or u"no se alineó ningún leader"
    return True, None


def _apply_leader_free_end(tag, document, view):
    """Activa leader, Free End y alineación horizontal. Devuelve (ok, mensaje_error)."""
    if tag is None:
        return False, u"etiqueta nula"
    try:
        tag.HasLeader = True
    except Exception as ex:
        return False, u"No se pudo activar leader: {0}".format(_as_unicode(ex))

    try:
        if not bool(tag.HasLeader):
            return False, u"la etiqueta no admite línea de llamada"
    except Exception:
        pass

    try:
        tag.LeaderEndCondition = LeaderEndCondition.Free
    except Exception as ex:
        return False, u"No se pudo aplicar Free End: {0}".format(_as_unicode(ex))

    try:
        if tag.LeaderEndCondition != LeaderEndCondition.Free:
            return False, u"LeaderEndCondition no quedó en Free"
    except Exception:
        pass

    try:
        document.Regenerate()
    except Exception:
        pass

    ok_h, err_h = _align_leaders_horizontal(tag, document, view)
    if not ok_h:
        return False, u"Free End ok; alineación horizontal falló: {0}".format(
            err_h or u"?"
        )

    try:
        document.Regenerate()
    except Exception:
        pass

    # Segunda pasada: el bbox del tag ya refleja el leader y estabiliza la 1.ª línea.
    _align_leaders_horizontal(tag, document, view)
    return True, None


def run(uiapp):
    """Entrada pyRevit: uiapp = __revit__."""
    if uiapp is None:
        _mostrar_aviso(None, u"No hay aplicación Revit activa.")
        return

    try:
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document if uidoc is not None else None
    except Exception as ex:
        _mostrar_aviso(uiapp, u"No hay documento activo.", _as_unicode(ex))
        return

    if uidoc is None or doc is None:
        _mostrar_aviso(uiapp, u"No hay documento activo.")
        return

    if doc.IsReadOnly:
        _mostrar_aviso(uiapp, u"El documento está en solo lectura.")
        return

    view = uidoc.ActiveView
    if view is None:
        _mostrar_aviso(uiapp, u"No hay vista activa.")
        return

    tags = _get_preselected_tags(uidoc, doc)
    if not tags:
        try:
            tags = _pick_tags(uidoc)
        except OperationCanceledException:
            return
        except Exception as ex:
            _mostrar_aviso(uiapp, u"Error en la selección.", _as_unicode(ex))
            return

    if not tags:
        _mostrar_aviso(
            uiapp,
            u"No se seleccionó ninguna etiqueta válida.",
            u"Seleccione uno o varios IndependentTag (etiquetas de elemento).",
        )
        return

    ok_n = 0
    err_rows = []
    t = Transaction(doc, _TXN_NAME)
    t.Start()
    try:
        for tag in tags:
            tid = _eid_int(tag.Id)
            ok, err = _apply_leader_free_end(tag, doc, view)
            if ok:
                ok_n += 1
            else:
                err_rows.append(u"Id {0}: {1}".format(tid, err or u"error"))
        t.Commit()
    except Exception as ex:
        try:
            if t.GetStatus() == TransactionStatus.Started:
                t.RollBack()
        except Exception:
            pass
        _mostrar_aviso(uiapp, u"Error al aplicar Leader Free End.", _as_unicode(ex))
        return

    content_lines = [
        u"Etiquetas procesadas: {0} / {1}.".format(ok_n, len(tags)),
        u"Leader activado, Free End y alineación horizontal aplicados.",
    ]
    if err_rows:
        content_lines.append(u"")
        content_lines.append(u"No aplicadas:")
        content_lines.extend(err_rows[:10])
        if len(err_rows) > 10:
            content_lines.append(u"… (+{0})".format(len(err_rows) - 10))

    _mostrar_aviso(
        uiapp,
        u"Tag Leader Free End finalizado.",
        u"\n".join(content_lines),
    )
