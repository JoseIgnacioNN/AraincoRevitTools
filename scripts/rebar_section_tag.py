# -*- coding: utf-8 -*-
"""
Arainco: Rebar Section Tag

Etiqueta uno o varios Rebar / RebarInSystem en la vista activa.
Solo admite barras cuyo eje NO es paralelo al plano de la vista (corte/sección).
Por cada set: etiqueta multihost, Merge Leaders y leader en L
(vertical +Up + shoulder horizontal), sin ajustar por bbox.
Familia/tipo: EST_A_STRUCTURAL REBAR TAG_DETAIL / Cantidad - Diametro.

Revit 2025+ | pyRevit / RPS (importable via run).
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    Family,
    FamilySymbol,
    FilteredElementCollector,
    IndependentTag,
    LeaderEndCondition,
    Options,
    Reference,
    TagMode,
    TagOrientation,
    Transaction,
    View3D,
    ViewDetailLevel,
    XYZ,
)
from Autodesk.Revit.DB.Structure import Rebar, RebarInSystem
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from System.Collections.Generic import List

try:
    from Autodesk.Revit.DB.Structure import MultiplanarOption
except Exception:
    MultiplanarOption = None

# ---------------------------------------------------------------------------
# CONFIGURACION
# ---------------------------------------------------------------------------
ADD_LEADER = True
TAG_ORIENTATION = TagOrientation.Horizontal
TAG_SYMBOL_ID = None
# Familia/tipo de etiqueta (selector de tipo en detalle).
PREFERRED_TAG_FAMILY = u"EST_A_STRUCTURAL REBAR TAG_DETAIL"
PREFERRED_TAG_TYPE = u"Cantidad - Diametro"
# Tramo horizontal (shoulder) en el plano de vista, mm de modelo.
# Hueco mínimo elbow → inicio del texto de la etiqueta.
TAG_OFFSET_MM = 1.0
# Tramo vertical del leader hacia arriba (+Up), mm de modelo.
SHOULDER_VERTICAL_MM = 220.0
# +1 = leaders suben (como la imagen); -1 = bajan.
SHOULDER_VERTICAL_SIGN = 1.0
# "Right" / "Left" / "Auto" (Auto = hacia fuera del centroide de la selección).
# Right = View.RightDirection (derecha de pantalla en la vista).
TAG_SIDE = u"Right"
INVERT_TAG_SIDE = False
APPLY_SHOULDER_L = True
# Tras colocar, desplaza TagHeadPosition para que el texto nazca al final del shoulder.
ALIGN_TAG_TO_SHOULDER_END = True
# Separación mínima entre cabezas de etiquetas distintas (plano de vista), mm.
MIN_HEAD_SEPARATION_MM = 90.0
# |dir_barra · ViewDirection| debe superar este valor para admitir la barra.
PARALLEL_VIEW_COS_MAX = 0.17

_TOOL_TITLE = u"Arainco: Rebar Section Tag"
_TXN_NAME = u"Arainco: Rebar Section Tag"
PROMPT_REBAR = (
    u"Selecciona Rebar no paralelos al plano de vista. "
    u"Finish para continuar / Esc cancela."
)


def _as_unicode(text):
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _mostrar_aviso(uiapp, instruction, content=u"", ok_text=u"Entendido"):
    try:
        from bimtools_instruction_dialog import show_message_dialog
        from revit_wpf_window_position import revit_main_hwnd

        hwnd = revit_main_hwnd(uiapp) if uiapp is not None else None
        show_message_dialog(
            _TOOL_TITLE,
            instruction=_as_unicode(instruction),
            content=_as_unicode(content) if content else None,
            ok_text=ok_text,
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass
    msg = _as_unicode(instruction)
    if content:
        msg = u"{0}\n\n{1}".format(msg, _as_unicode(content))
    try:
        TaskDialog.Show(_TOOL_TITLE, msg)
    except Exception:
        pass


def _eid_int(eid):
    if eid is None:
        return -1
    try:
        return int(eid.IntegerValue)
    except Exception:
        try:
            return int(eid.Value)
        except Exception:
            return -1


def _view_ok(view):
    if view is None:
        return False, u"Vista nula."
    try:
        if view.IsTemplate:
            return False, u"Vista plantilla: abre una vista de modelo."
    except Exception:
        pass
    try:
        if str(view.ViewType) == "Perspective":
            return False, u"No se puede etiquetar en perspectiva."
    except Exception:
        pass
    try:
        if isinstance(view, View3D) and not view.IsLocked:
            return False, u"Vista 3D: bloquea la camara para etiquetar."
    except Exception:
        pass
    return True, None


def _safe_normalize(v):
    if v is None:
        return None
    try:
        if float(v.GetLength()) < 1e-9:
            return None
        return v.Normalize()
    except Exception:
        return None


def _centerline_curves(rb, bar_index=0):
    curves = []
    if MultiplanarOption is not None:
        for mpo_name in ("IncludeAllMultiplanarCurves", "IncludeOnlyPlanarCurves"):
            mpo = getattr(MultiplanarOption, mpo_name, None)
            if mpo is None:
                continue
            try:
                curves = list(
                    rb.GetCenterlineCurves(False, False, False, mpo, bar_index)
                )
                if curves:
                    return curves
            except Exception:
                curves = []
    try:
        curves = list(rb.GetCenterlineCurves(False, False, False))
    except Exception:
        curves = []
    return curves


def _rebar_axis_direction(rb):
    """Direccion longitudinal unitaria de la barra (posicion 0)."""
    curves = _centerline_curves(rb, 0)
    if not curves:
        return None
    try:
        longest = max(curves, key=lambda c: float(c.Length))
        return _safe_normalize(longest.GetEndPoint(1) - longest.GetEndPoint(0))
    except Exception:
        return None


def _rebar_not_parallel_to_view_plane(rb, view):
    """
    True si el eje de la barra NO es paralelo al plano de la vista.

    Plano de vista ⟂ ViewDirection. Paralelo al plano <=> |dir · ViewDirection| ≈ 0.
    """
    if rb is None or view is None:
        return False
    bar_dir = _rebar_axis_direction(rb)
    vd = _safe_normalize(getattr(view, "ViewDirection", None))
    if bar_dir is None or vd is None:
        return False
    try:
        return abs(float(bar_dir.DotProduct(vd))) > float(PARALLEL_VIEW_COS_MAX)
    except Exception:
        return False


class _RebarSectionFilter(ISelectionFilter):
    """Solo Structural Rebar / RebarInSystem no paralelos al plano de vista."""

    def __init__(self, view):
        self._view = view

    def AllowElement(self, elem):
        try:
            if not isinstance(elem, (Rebar, RebarInSystem)):
                return False
            return _rebar_not_parallel_to_view_plane(elem, self._view)
        except Exception:
            return False

    def AllowReference(self, reference, position):
        return True


def _tag_symbol_id(document=None):
    if TAG_SYMBOL_ID is not None and TAG_SYMBOL_ID != ElementId.InvalidElementId:
        return TAG_SYMBOL_ID
    try:
        import __main__ as m

        sid = getattr(m, "TAG_SYMBOL_ID", None)
        if sid is not None and sid != ElementId.InvalidElementId:
            return sid
    except Exception:
        pass
    if document is not None and PREFERRED_TAG_FAMILY:
        sid = _find_family_tag_symbol_id(
            document, PREFERRED_TAG_FAMILY, PREFERRED_TAG_TYPE
        )
        if sid is not None:
            return sid
    return None


def _norm_name(s):
    if s is None:
        return u""
    try:
        t = unicode(s)
    except Exception:
        t = str(s)
    return t.strip().lower()


def _symbol_type_name(sym):
    try:
        nm = getattr(sym, "Name", None)
        if nm:
            return _as_unicode(nm)
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import BuiltInParameter, StorageType

        for bip_name in (u"SYMBOL_NAME_PARAM", u"ALL_MODEL_TYPE_NAME"):
            bip = getattr(BuiltInParameter, bip_name, None)
            if bip is None:
                continue
            p = sym.get_Parameter(bip)
            if p is None or not p.HasValue:
                continue
            if p.StorageType != StorageType.String:
                continue
            raw = p.AsString()
            if raw:
                return _as_unicode(raw)
    except Exception:
        pass
    return u""


def _find_family_tag_symbol_id(document, family_name, type_name=None):
    """FamilySymbol OST_RebarTags de la familia (y tipo, si se indica)."""
    if document is None or not family_name:
        return None
    wanted_fam = _norm_name(family_name)
    wanted_typ = _norm_name(type_name) if type_name else u""
    try:
        for fam in FilteredElementCollector(document).OfClass(Family):
            try:
                if _norm_name(fam.Name) != wanted_fam:
                    continue
            except Exception:
                continue
            for sid in fam.GetFamilySymbolIds():
                sym = document.GetElement(sid)
                if sym is None or not isinstance(sym, FamilySymbol):
                    continue
                try:
                    cat = sym.Category
                    if cat is None:
                        continue
                    if int(cat.Id.IntegerValue) != int(BuiltInCategory.OST_RebarTags):
                        continue
                except Exception:
                    continue
                if not wanted_typ:
                    return sym.Id
                if _norm_name(_symbol_type_name(sym)) == wanted_typ:
                    return sym.Id
    except Exception:
        pass
    return None


def _ref_key(document, r):
    try:
        return r.ConvertToStableRepresentation(document)
    except Exception:
        try:
            return str(_eid_int(r.ElementId))
        except Exception:
            return id(r)


def _layout_bar_references(document, rb):
    refs = []
    seen = set()

    def add(r):
        if r is None:
            return
        key = _ref_key(document, r)
        if key in seen:
            return
        seen.add(key)
        refs.append(r)

    npos = 0
    try:
        npos = int(rb.NumberOfBarPositions)
    except Exception:
        npos = 0

    if npos > 0 and hasattr(rb, "GetReferenceToBarPosition"):
        for i in range(npos):
            try:
                if hasattr(rb, "DoesBarExistAtPosition"):
                    if not bool(rb.DoesBarExistAtPosition(i)):
                        continue
            except Exception:
                pass
            try:
                add(rb.GetReferenceToBarPosition(i))
            except Exception:
                continue

    if len(refs) < 2:
        try:
            for sub in rb.GetSubelements() or []:
                try:
                    add(sub.GetReference())
                except Exception:
                    pass
        except Exception:
            pass

    if not refs:
        try:
            add(Reference(rb))
        except Exception:
            pass
    return refs


def _refs_for_rebar_fallback(document, view, rb):
    refs = list(_layout_bar_references(document, rb))
    seen = set(_ref_key(document, r) for r in refs)

    def add(r):
        if r is None:
            return
        key = _ref_key(document, r)
        if key in seen:
            return
        seen.add(key)
        refs.append(r)

    try:
        add(Reference(rb))
    except Exception:
        pass
    try:
        opts = Options()
        opts.ComputeReferences = True
        try:
            opts.View = view
            opts.DetailLevel = ViewDetailLevel.Fine
        except Exception:
            pass
        geom = rb.get_Geometry(opts)
        if geom:
            for go in geom:
                try:
                    if go.Reference:
                        add(go.Reference)
                except Exception:
                    pass
    except Exception:
        pass
    return refs


def _create_independent_tag(document, view, ref, point, sid):
    if sid is not None:
        return IndependentTag.Create(
            document, sid, view.Id, ref, ADD_LEADER, TAG_ORIENTATION, point
        )
    return IndependentTag.Create(
        document,
        view.Id,
        ref,
        ADD_LEADER,
        TagMode.TM_ADDBY_CATEGORY,
        TAG_ORIENTATION,
        point,
    )


def _apply_merge_leaders(tag):
    if tag is None:
        return False, u"tag nulo"
    try:
        tag.HasLeader = True
    except Exception:
        pass
    try:
        if not bool(getattr(tag, "HasLeader", True)):
            return False, u"la etiqueta no tiene leader"
    except Exception:
        pass
    try:
        tag.LeaderEndCondition = LeaderEndCondition.Attached
    except Exception:
        pass
    try:
        tag.MergeElbows = True
        return True, None
    except Exception as ex:
        return False, _as_unicode(ex)


def _view_axes(view):
    n = _safe_normalize(getattr(view, "ViewDirection", None))
    r = _safe_normalize(getattr(view, "RightDirection", None))
    u = _safe_normalize(getattr(view, "UpDirection", None))
    return n, r, u


def _project_onto_view_plane(vec, view_dir):
    if vec is None or view_dir is None:
        return vec
    try:
        dn = float(vec.DotProduct(view_dir))
        return XYZ(
            float(vec.X) - float(view_dir.X) * dn,
            float(vec.Y) - float(view_dir.Y) * dn,
            float(vec.Z) - float(view_dir.Z) * dn,
        )
    except Exception:
        return vec


def _average_xyz(points):
    pts = [p for p in (points or []) if p is not None]
    if not pts:
        return None
    sx = sy = sz = 0.0
    for p in pts:
        sx += float(p.X)
        sy += float(p.Y)
        sz += float(p.Z)
    n = float(len(pts))
    return XYZ(sx / n, sy / n, sz / n)


def _tagged_refs(tag):
    try:
        refs = list(tag.GetTaggedReferences() or [])
        if refs:
            return refs
    except Exception:
        pass
    try:
        r = tag.GetTaggedReference()
        if r is not None:
            return [r]
    except Exception:
        pass
    return []


def _leader_ends(tag, document):
    out = []
    for ref in _tagged_refs(tag):
        end = None
        try:
            end = tag.GetLeaderEnd(ref)
        except Exception:
            end = None
        if end is None:
            try:
                el = document.GetElement(ref.ElementId)
                end = _centerline_midpoint(el) if el is not None else None
            except Exception:
                end = None
        if end is not None:
            out.append((ref, end))
    return out


def _horizontal_side_sign(side_sign):
    """Signo efectivo en View.RightDirection (+ = Right API, con flip opcional)."""
    side = 1.0 if float(side_sign) >= 0.0 else -1.0
    if INVERT_TAG_SIDE:
        side = -side
    return side


def _tag_bbox_right_span(tag, view, right_dir):
    """
    (min, max) de la proyección del bbox del tag sobre RightDirection.
    None si no hay bbox usable.
    """
    if tag is None or view is None or right_dir is None:
        return None
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
    if bb is None or bb.Min is None or bb.Max is None:
        return None
    corners = (
        XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z),
        XYZ(bb.Min.X, bb.Min.Y, bb.Max.Z),
        XYZ(bb.Min.X, bb.Max.Y, bb.Min.Z),
        XYZ(bb.Min.X, bb.Max.Y, bb.Max.Z),
        XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z),
        XYZ(bb.Max.X, bb.Min.Y, bb.Max.Z),
        XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z),
        XYZ(bb.Max.X, bb.Max.Y, bb.Max.Z),
    )
    try:
        projs = [float(c.DotProduct(right_dir)) for c in corners]
        return min(projs), max(projs)
    except Exception:
        return None


def _align_head_to_shoulder_end(document, view, tag, elbow, right_dir, side, shoulder_ft):
    """
    Mueve TagHeadPosition para que el borde interior del texto quede al final
    del shoulder (elbow + side·shoulder), y el texto se extienda hacia afuera.
    """
    if not ALIGN_TAG_TO_SHOULDER_END:
        return None
    if tag is None or elbow is None or right_dir is None:
        return None
    try:
        head = tag.TagHeadPosition
    except Exception:
        head = None
    if head is None:
        return None
    try:
        document.Regenerate()
    except Exception:
        pass
    span = _tag_bbox_right_span(tag, view, right_dir)
    if span is None:
        return None
    bb_min_r, bb_max_r = span
    elbow_r = float(elbow.DotProduct(right_dir))
    # Borde del texto hacia el codo (inicio visual de la etiqueta).
    if float(side) >= 0.0:
        inner_r = bb_min_r
        target_inner = elbow_r + abs(float(shoulder_ft))
    else:
        inner_r = bb_max_r
        target_inner = elbow_r - abs(float(shoulder_ft))
    shift = target_inner - inner_r
    if abs(shift) < 1e-9:
        return head
    head_new = _xyz_add_scaled(head, right_dir, shift)
    try:
        tag.TagHeadPosition = head_new
        document.Regenerate()
    except Exception:
        return None
    return head_new


def _apply_leader_shoulder_l(document, view, tag, head_target=None, side_sign=1.0):
    """
    Shoulder en L:
    MergeElbows + vertical (+Up) + estante horizontal; texto nace al final del shoulder.
    """
    if not APPLY_SHOULDER_L:
        return False, u"desactivado"
    if tag is None or view is None or document is None:
        return False, u"tag/vista nulos"

    n, r, u = _view_axes(view)
    if r is None or u is None:
        return False, u"ejes de vista invalidos"

    merged_ok, merge_err = _apply_merge_leaders(tag)
    if not merged_ok:
        return False, merge_err or u"MergeElbows fallo"

    side = _horizontal_side_sign(side_sign)
    min_v = _mm_to_ft(SHOULDER_VERTICAL_MM)
    min_h = _mm_to_ft(TAG_OFFSET_MM)

    if head_target is not None:
        try:
            tag.TagHeadPosition = head_target
        except Exception:
            pass
    try:
        document.Regenerate()
    except Exception:
        pass

    ends = _leader_ends(tag, document)
    if not ends:
        return False, u"sin LeaderEnd"

    avg_end = _average_xyz([e for _, e in ends])
    if avg_end is None:
        return False, u"sin ancla media"

    try:
        head_now = tag.TagHeadPosition
    except Exception:
        head_now = head_target
    if head_now is None:
        head_now = _xyz_add_scaled(
            _xyz_add_scaled(avg_end, u, min_v * float(SHOULDER_VERTICAL_SIGN)),
            r,
            min_h * side,
        )

    d = _project_onto_view_plane(head_now - avg_end, n)
    t_raw = float(d.DotProduct(u)) if d is not None else 0.0

    v_sign = 1.0 if float(SHOULDER_VERTICAL_SIGN) >= 0.0 else -1.0
    if abs(t_raw) < min_v:
        t_use = v_sign * min_v
    else:
        t_use = v_sign * max(abs(t_raw), min_v)

    # Shoulder horizontal fijo hacia el lado pedido.
    s_use = side * min_h

    elbow = _xyz_add_scaled(avg_end, u, t_use)
    head_final = _xyz_add_scaled(elbow, r, s_use)
    if n is not None:
        try:
            gamma = float((avg_end - elbow).DotProduct(n))
            head_final = _xyz_add_scaled(head_final, n, gamma)
        except Exception:
            pass

    try:
        if elbow.DistanceTo(avg_end) < 1e-4:
            return False, u"codo coincide con ancla"
    except Exception:
        pass

    for ref, _end in ends:
        try:
            tag.SetLeaderElbow(ref, elbow)
        except Exception:
            continue
    try:
        document.Regenerate()
    except Exception:
        pass

    for _ in (0, 1):
        try:
            tag.TagHeadPosition = head_final
        except Exception as ex:
            return False, u"TagHeadPosition: {0}".format(_as_unicode(ex))
        try:
            document.Regenerate()
        except Exception:
            pass

    # Alinear borde interior del texto al término del shoulder.
    aligned = _align_head_to_shoulder_end(
        document, view, tag, elbow, r, side, min_h
    )
    if aligned is not None:
        head_final = aligned
        # Segunda pasada por si el bbox cambió al mover la cabeza.
        aligned2 = _align_head_to_shoulder_end(
            document, view, tag, elbow, r, side, min_h
        )
        if aligned2 is not None:
            head_final = aligned2

    # Reafirmar codo común tras mover la cabeza (Merge Leaders).
    for ref, _end in ends:
        try:
            tag.SetLeaderElbow(ref, elbow)
        except Exception:
            continue
    try:
        tag.MergeElbows = True
        tag.TagHeadPosition = head_final
        document.Regenerate()
    except Exception:
        pass
    return True, None


def _add_multihost_refs(document, tag, extra_refs):
    if tag is None or not extra_refs:
        return 0, None
    add_fn = getattr(tag, "AddReferences", None)
    if add_fn is None:
        return 0, u"Esta version de Revit no expone AddReferences."
    refs_add = List[Reference]()
    for ref in extra_refs:
        refs_add.Add(ref)
    if refs_add.Count < 1:
        return 0, None
    try:
        document.Regenerate()
    except Exception:
        pass
    try:
        add_fn(refs_add)
        return int(refs_add.Count), None
    except Exception as ex:
        return 0, (
            u"No se pudieron agregar hosts del layout (AddReferences):\n{0}\n\n"
            u"Usa un tipo de etiqueta MULTI HOST o define TAG_SYMBOL_ID."
        ).format(_as_unicode(ex))


def _create_tag(document, view, rb, head_point, side_sign=1.0):
    layout_refs = _layout_bar_references(document, rb)
    if not layout_refs:
        layout_refs = _refs_for_rebar_fallback(document, view, rb)
    if not layout_refs:
        raise Exception(u"Sin referencia valida del rebar.")

    sid = _tag_symbol_id(document)
    if sid is not None:
        try:
            sym = document.GetElement(sid)
            if sym is not None and hasattr(sym, "IsActive") and not sym.IsActive:
                sym.Activate()
                document.Regenerate()
        except Exception:
            pass

    tag = None
    primary_idx = 0
    last = None
    layout_keys = set(_ref_key(document, x) for x in layout_refs)
    candidates = list(layout_refs) + [
        r
        for r in _refs_for_rebar_fallback(document, view, rb)
        if _ref_key(document, r) not in layout_keys
    ]
    for i, ref in enumerate(candidates):
        try:
            tag = _create_independent_tag(document, view, ref, head_point, sid)
            if tag is not None:
                primary_idx = i
                break
        except Exception as ex:
            last = ex
            tag = None
    if tag is None:
        raise last or Exception(u"IndependentTag.Create fallo.")

    primary_key = _ref_key(document, candidates[primary_idx])
    extras = []
    for ref in layout_refs:
        if _ref_key(document, ref) == primary_key:
            continue
        extras.append(ref)

    n_extra, warn = _add_multihost_refs(document, tag, extras)
    n_hosts = 1 + int(n_extra or 0)

    try:
        tag.TagHeadPosition = head_point
        document.Regenerate()
    except Exception:
        pass

    merged_ok, merge_err = _apply_merge_leaders(tag)
    if not merged_ok:
        note = u"Merge Leaders no aplicado: {0}".format(merge_err or u"?")
        warn = (warn + u"\n" + note) if warn else note

    shoulder_ok, shoulder_msg = _apply_leader_shoulder_l(
        document, view, tag, head_target=head_point, side_sign=side_sign
    )
    if not shoulder_ok:
        note = u"Shoulder L no aplicado: {0}".format(shoulder_msg or u"?")
        warn = (warn + u"\n" + note) if warn else note

    if extras and n_extra == 0:
        note = u"El layout tiene {0} barras pero no se agregaron hosts extra.".format(
            len(layout_refs)
        )
        warn = (warn + u"\n" + note) if warn else note

    return tag, n_hosts, warn, bool(merged_ok), bool(shoulder_ok)


def _pick_rebars(uidoc, doc, view):
    refs = list(
        uidoc.Selection.PickObjects(
            ObjectType.Element, _RebarSectionFilter(view), PROMPT_REBAR
        )
    )
    out = []
    seen = set()
    for pref in refs:
        el = doc.GetElement(pref.ElementId)
        if not isinstance(el, (Rebar, RebarInSystem)):
            continue
        if not _rebar_not_parallel_to_view_plane(el, view):
            continue
        iid = _eid_int(el.Id)
        if iid in seen:
            continue
        seen.add(iid)
        out.append(el)
    return out


def _mm_to_ft(mm):
    return float(mm) / 304.8


def _xyz_add_scaled(origin, direction, scale):
    try:
        return origin + direction.Multiply(scale)
    except Exception:
        return XYZ(
            float(origin.X) + float(direction.X) * scale,
            float(origin.Y) + float(direction.Y) * scale,
            float(origin.Z) + float(direction.Z) * scale,
        )


def _centerline_midpoint(rb):
    curves = _centerline_curves(rb, 0)
    if not curves:
        return None
    try:
        longest = max(curves, key=lambda c: float(c.Length))
        return longest.Evaluate(0.5, True)
    except Exception:
        return None


def _bbox_center(rb, view):
    try:
        bb = rb.get_BoundingBox(view)
        if bb is None:
            bb = rb.get_BoundingBox(None)
        if bb is not None:
            return (bb.Min + bb.Max) * 0.5
    except Exception:
        pass
    return None


def _layout_bar_midpoints(rb):
    """Midpoints de posiciones existentes del layout (para anclar el L)."""
    pts = []
    try:
        npos = int(rb.NumberOfBarPositions)
    except Exception:
        npos = 0
    if npos <= 0:
        mid = _centerline_midpoint(rb)
        return [mid] if mid is not None else []
    for i in range(npos):
        try:
            if hasattr(rb, "DoesBarExistAtPosition") and not bool(
                rb.DoesBarExistAtPosition(i)
            ):
                continue
        except Exception:
            pass
        curves = _centerline_curves(rb, i)
        if not curves:
            continue
        try:
            longest = max(curves, key=lambda c: float(c.Length))
            pts.append(longest.Evaluate(0.5, True))
        except Exception:
            continue
    if not pts:
        mid = _centerline_midpoint(rb)
        if mid is not None:
            pts.append(mid)
    return pts


def _rebar_anchor_point_in_view(rb, view):
    """Ancla del set: punto medio del layout (fallback: bbox / top)."""
    pts = _layout_bar_midpoints(rb)
    avg = _average_xyz(pts)
    if avg is not None:
        return avg
    bb = _bbox_center(rb, view)
    if bb is not None:
        return bb
    return _rebar_top_point_in_view(rb, view)


def _rebar_top_point_in_view(rb, view):
    """Punto del layout mas alto en pantalla (max proyeccion en Up)."""
    up = _safe_normalize(getattr(view, "UpDirection", None))
    pts = _layout_bar_midpoints(rb)
    if not pts:
        return _bbox_center(rb, view)
    if up is None:
        return pts[0]
    try:
        return max(pts, key=lambda p: float(p.DotProduct(up)))
    except Exception:
        return pts[0]


def _side_signs_for_rebars(rebars, view):
    """Signo horizontal por set: +1 = RightDirection, -1 = Left."""
    mode = _norm_name(TAG_SIDE)
    n = len(rebars or [])
    if n < 1:
        return []
    if mode == u"left":
        return [-1.0] * n
    if mode != u"auto":
        return [1.0] * n

    right = _safe_normalize(getattr(view, "RightDirection", None))
    if right is None:
        return [1.0] * n
    anchors = []
    for rb in rebars:
        a = _rebar_anchor_point_in_view(rb, view)
        anchors.append(a)
    vals = []
    for a in anchors:
        if a is None:
            vals.append(0.0)
        else:
            vals.append(float(a.DotProduct(right)))
    mean_r = sum(vals) / float(len(vals))
    signs = []
    for v in vals:
        signs.append(1.0 if v >= mean_r else -1.0)
    # Si todos quedan iguales (casi colineales en Right), fuerza derecha.
    if len(set(signs)) == 1 and abs(max(vals) - min(vals)) < _mm_to_ft(5.0):
        return [1.0] * n
    return signs


def _auto_tag_head_for_rebar(rb, view, side_sign=1.0):
    """
    Cabeza objetivo por set:
    barra superior → +Up (vertical) → ±Right (shoulder).
    """
    top = _rebar_top_point_in_view(rb, view)
    if top is None:
        return None
    up = _safe_normalize(getattr(view, "UpDirection", None))
    right = _safe_normalize(getattr(view, "RightDirection", None))
    side = _horizontal_side_sign(side_sign)
    head = top
    if up is not None:
        head = _xyz_add_scaled(
            head, up, _mm_to_ft(SHOULDER_VERTICAL_MM) * float(SHOULDER_VERTICAL_SIGN)
        )
    if right is not None:
        head = _xyz_add_scaled(head, right, _mm_to_ft(TAG_OFFSET_MM) * side)
    return head


def _separate_overlapping_heads(placements, view):
    """
    placements: [(rb, head, side_sign), ...]
    Empuja cabezas demasiado cercanas a lo largo de Up (y un poco Right).
    """
    if not placements or len(placements) < 2:
        return placements
    n_ax, r, u = _view_axes(view)
    if u is None:
        return placements
    min_sep = _mm_to_ft(MIN_HEAD_SEPARATION_MM)
    items = [[rb, head, side] for rb, head, side in placements]

    def sort_key(item):
        _rb, head, _s = item
        up_v = float(head.DotProduct(u)) if head is not None else 0.0
        rt_v = float(head.DotProduct(r)) if (head is not None and r is not None) else 0.0
        return (up_v, rt_v)

    for _pass in range(max(1, len(items))):
        items.sort(key=sort_key)
        moved = False
        for i in range(1, len(items)):
            h_prev = items[i - 1][1]
            h_cur = items[i][1]
            if h_prev is None or h_cur is None:
                continue
            delta = _project_onto_view_plane(h_cur - h_prev, n_ax)
            try:
                dist = float(delta.GetLength()) if delta is not None else 0.0
            except Exception:
                dist = 0.0
            if dist >= min_sep:
                continue
            need = min_sep - dist + _mm_to_ft(5.0)
            nudge = _xyz_add_scaled(XYZ(0.0, 0.0, 0.0), u, need)
            d_up = 0.0
            try:
                d_up = abs(float(delta.DotProduct(u))) if delta is not None else 0.0
            except Exception:
                d_up = 0.0
            if r is not None and d_up < min_sep * 0.25:
                side = 1.0 if float(items[i][2]) >= 0.0 else -1.0
                nudge = _xyz_add_scaled(nudge, r, need * 0.35 * side)
            items[i][1] = h_cur + nudge
            moved = True
        if not moved:
            break

    return [(h[0], h[1], h[2]) for h in items]


def run(uiapp):
    """Entrada pyRevit / RPS: uiapp = __revit__."""
    if uiapp is None:
        _mostrar_aviso(None, u"No hay aplicacion Revit activa.")
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

    view = uidoc.ActiveView
    ok, msg = _view_ok(view)
    if not ok:
        _mostrar_aviso(uiapp, msg)
        return

    try:
        rebars = _pick_rebars(uidoc, doc, view)
    except OperationCanceledException:
        return
    except Exception as ex:
        _mostrar_aviso(uiapp, u"Error en seleccion.", _as_unicode(ex))
        return

    if not rebars:
        _mostrar_aviso(
            uiapp,
            u"No se selecciono ningun Rebar valido.",
            u"Solo se admiten Structural Rebar cuyo eje no sea paralelo "
            u"al plano de la vista activa.",
        )
        return

    side_signs = _side_signs_for_rebars(rebars, view)
    placements = []
    skip_rows = []
    for rb, side in zip(rebars, side_signs):
        head_pt = _auto_tag_head_for_rebar(rb, view, side_sign=side)
        if head_pt is None:
            skip_rows.append(
                u"Rebar {0}: sin ancla/geometria para posicion automatica.".format(
                    _eid_int(rb.Id)
                )
            )
            continue
        placements.append((rb, head_pt, side))

    placements = _separate_overlapping_heads(placements, view)

    if not placements:
        _mostrar_aviso(
            uiapp,
            u"No se pudo calcular posicion automatica para ningun rebar.",
            u"\n".join(skip_rows[:8]),
        )
        return

    t = Transaction(doc, _TXN_NAME)
    t.Start()
    ok_rows = []
    err_rows = []
    try:
        for rb, head_pt, side in placements:
            rid = _eid_int(rb.Id)
            try:
                tag, n_hosts, warn, merged, shoulder = _create_tag(
                    doc, view, rb, head_pt, side_sign=side
                )
                side_txt = u"Right" if float(side) >= 0.0 else u"Left"
                ok_rows.append(
                    u"Rebar {0} -> Tag {1} | Hosts={2} | Merge={3} | "
                    u"Shoulder={4} | Side={5}{6}".format(
                        rid,
                        _eid_int(tag.Id),
                        n_hosts,
                        u"Si" if merged else u"No",
                        u"Si" if shoulder else u"No",
                        side_txt,
                        (u" | " + warn) if warn else u"",
                    )
                )
            except Exception as ex:
                err_rows.append(u"Rebar {0}: {1}".format(rid, _as_unicode(ex)))
        if not ok_rows and err_rows:
            raise Exception(u"\n".join(err_rows))
        t.Commit()
    except Exception as ex:
        try:
            if t.HasStarted():
                t.RollBack()
        except Exception:
            pass
        _mostrar_aviso(uiapp, u"No se pudo crear las etiquetas.", _as_unicode(ex))
        return

    side_label = _as_unicode(TAG_SIDE)
    content_lines = [
        u"Etiquetas creadas: {0} / {1}.".format(len(ok_rows), len(placements)),
        u"Leader L: vertical {0:g} mm (+Up) + shoulder {1:g} mm (lado {2}).".format(
            float(SHOULDER_VERTICAL_MM), float(TAG_OFFSET_MM), side_label
        ),
        u"Etiqueta: {0} / {1}.".format(PREFERRED_TAG_FAMILY, PREFERRED_TAG_TYPE),
        u"",
    ]
    content_lines.extend(ok_rows[:12])
    if len(ok_rows) > 12:
        content_lines.append(u"...")
    if skip_rows:
        content_lines.append(u"")
        content_lines.append(u"Omitidos:")
        content_lines.extend(skip_rows[:8])
    if err_rows:
        content_lines.append(u"")
        content_lines.append(u"Errores:")
        content_lines.extend(err_rows[:8])
    _mostrar_aviso(uiapp, u"Rebar Section Tag finalizado.", u"\n".join(content_lines))
