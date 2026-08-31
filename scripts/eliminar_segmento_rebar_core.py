# -*- coding: utf-8 -*-
"""
Eliminar segmento rebar — pick, alzado y CreateFromCurves.

Revit 2024+ | pyRevit | IronPython
"""

from __future__ import print_function

import math

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    Transaction,
    UnitTypeId,
    UnitUtils,
)
from Autodesk.Revit.DB.Structure import Rebar, RebarHookOrientation, RebarStyle
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

from dividir_rebar_punto_core import (
    _attach_rebar_outside_host_swallower,
    _build_view_projection_frame,
    _cantidad_posiciones,
    _centerline_curves,
    _clone_curve,
    _collect_concrete_context_polylines_uv,
    _copy_bar_included,
    _copy_layout,
    _create_from_curves,
    _curve_endpoints,
    _curve_length,
    _element_id_int,
    _exception_text,
    _hook_type,
    _is_free_form,
    _is_in_group,
    _layout_rule_name,
    _rebar_normal,
    apply_rebar_presentation,
    apply_unobscured_to_rebars,
    capture_rebar_presentation,
    copy_armadura_instance_parameters,
    internal_to_mm,
    resolve_active_model_view,
)
from dividir_rebar_punto_geom import project_xyz_mm_to_uv
from dividir_rebar_punto_shapes import (
    find_rebar_shape_by_name,
    get_rebar_shape_name,
    normalize_shape_key,
    rebar_shape_display_name,
    set_rebar_shape,
)
from dividir_rebar_punto_tags import capture_rebar_tag_infos, tag_divided_rebars

_TRANSACTION_NAME = u"Arainco: Eliminar segmento rebar"
_LETTERS = u"ABCDEFGHIJKLMNOPQRSTUVWXYZ"
_MIN_SEG_FT = 1e-6
_DIR_MATCH = 0.82


class _RebarSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, Rebar)

    def AllowReference(self, reference, point):
        return False


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _is_cancelled(ex):
    try:
        import Autodesk.Revit.Exceptions as _RvtEx

        if isinstance(ex, _RvtEx.OperationCanceledException):
            return True
    except Exception:
        pass
    msg = _exception_text(ex).lower()
    return u"cancel" in msg or u"abort" in msg or u"operationcanceled" in msg


def segment_letter(idx):
    try:
        i = int(idx)
    except Exception:
        i = 0
    if 0 <= i < len(_LETTERS):
        return _LETTERS[i]
    return u"{0}".format(i + 1)


def is_end_segment(idx, n):
    try:
        i = int(idx)
        n = int(n)
    except Exception:
        return False
    if n < 2:
        return False
    return i == 0 or i == n - 1


def rebars_from_selection(uidoc):
    """Rebars de la selección actual (sin pick)."""
    if uidoc is None:
        return []
    doc = uidoc.Document
    out = []
    seen = set()
    try:
        ids = uidoc.Selection.GetElementIds()
    except Exception:
        ids = None
    if ids is None:
        return []
    for eid in ids:
        try:
            el = doc.GetElement(eid)
        except Exception:
            el = None
        if not isinstance(el, Rebar):
            continue
        iid = _element_id_int(el.Id)
        if iid in seen:
            continue
        seen.add(iid)
        out.append(el)
    return out


def pick_rebars(uidoc):
    """
    Pick múltiple de Rebar.

    Returns:
        (lista_rebars, error_o_None). Lista vacía + None si canceló.
    """
    if uidoc is None:
        return [], u"No hay documento activo."
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            _RebarSelectionFilter(),
            u"Seleccione una o más armaduras estructurales (Rebar). Pulse Finalizar.",
        )
    except Exception as ex:
        if _is_cancelled(ex):
            return [], None
        return [], _exception_text(ex)
    if refs is None:
        return [], None
    doc = uidoc.Document
    out = []
    seen = set()
    for ref in refs:
        try:
            el = doc.GetElement(ref.ElementId)
        except Exception:
            el = None
        if not isinstance(el, Rebar):
            continue
        iid = _element_id_int(el.Id)
        if iid in seen:
            continue
        seen.add(iid)
        out.append(el)
    return out, None


def _diameter_mm(rebar):
    for getter in (
        lambda: rebar.get_Parameter(BuiltInParameter.REBAR_BAR_DIAMETER),
        lambda: rebar.LookupParameter(u"Bar Diameter"),
    ):
        try:
            p = getter()
            if p is not None and p.HasValue:
                return float(
                    UnitUtils.ConvertFromInternalUnits(p.AsDouble(), UnitTypeId.Millimeters)
                )
        except Exception:
            pass
    try:
        bt = rebar.Document.GetElement(rebar.GetTypeId())
        p = bt.get_Parameter(BuiltInParameter.REBAR_BAR_DIAMETER)
        if p is not None and p.HasValue:
            return float(
                UnitUtils.ConvertFromInternalUnits(p.AsDouble(), UnitTypeId.Millimeters)
            )
    except Exception:
        pass
    return None


def _xyz_mm(pt):
    if pt is None:
        return None
    try:
        return (
            internal_to_mm(pt.X),
            internal_to_mm(pt.Y),
            internal_to_mm(pt.Z),
        )
    except Exception:
        return None


def _curve_xyz_samples(curve, n_arc=10):
    pts = []
    try:
        tess = curve.Tessellate()
        if tess is not None and tess.Count >= 2:
            for i in range(tess.Count):
                mm = _xyz_mm(tess[i])
                if mm is not None:
                    pts.append(mm)
            if len(pts) >= 2:
                return pts
    except Exception:
        pass
    p0, p1 = _curve_endpoints(curve)
    a = _xyz_mm(p0)
    b = _xyz_mm(p1)
    if a is None or b is None:
        return []
    try:
        if hasattr(curve, u"Evaluate") and abs(float(curve.Length)) > _MIN_SEG_FT:
            samples = []
            steps = max(2, int(n_arc))
            for i in range(steps):
                t = float(i) / float(steps - 1)
                try:
                    mm = _xyz_mm(curve.Evaluate(t, True))
                except Exception:
                    mm = None
                if mm is not None:
                    samples.append(mm)
            if len(samples) >= 2:
                return samples
    except Exception:
        pass
    return [a, b]


def _unit2(dx, dy):
    L = math.sqrt(dx * dx + dy * dy)
    if L < 1e-9:
        return 0.0, 0.0
    return dx / L, dy / L


def _project_pts_uv(pts_mm, origin_mm, u_axis, v_axis):
    uv = []
    for p in pts_mm or []:
        uv.append(list(project_xyz_mm_to_uv(p, origin_mm, u_axis, v_axis)))
    return uv


def _fallback_frame_from_bars(rebars):
    """Marco UV si la vista no da proyección (planta / 3D)."""
    pts = []
    for rb in rebars or []:
        curves = _centerline_curves(rb, 0, True, True) or _centerline_curves(
            rb, 0, True, False
        )
        for c in curves:
            for ep in _curve_endpoints(c):
                mm = _xyz_mm(ep)
                if mm is not None:
                    pts.append(mm)
    if not pts:
        return None
    origin = pts[0]
    n = None
    try:
        n = _rebar_normal(rebars[0])
    except Exception:
        n = None
    # U horizontal en XY, V ≈ Z (alzado “hacia +Z”).
    u = (1.0, 0.0, 0.0)
    v = (0.0, 0.0, 1.0)
    if n is not None:
        try:
            nx, ny, nz = float(n.X), float(n.Y), float(n.Z)
            # Eje derecho ⟂ normal, preferir horizontal.
            ux, uy, uz = -ny, nx, 0.0
            Lu = math.sqrt(ux * ux + uy * uy + uz * uz)
            if Lu > 1e-9:
                u = (ux / Lu, uy / Lu, uz / Lu)
            vx = ny * u[2] - nz * u[1]
            vy = nz * u[0] - nx * u[2]
            vz = nx * u[1] - ny * u[0]
            Lv = math.sqrt(vx * vx + vy * vy + vz * vz)
            if Lv > 1e-9:
                v = (vx / Lv, vy / Lv, vz / Lv)
            if v[2] < 0:
                v = (-v[0], -v[1], -v[2])
                u = (-u[0], -u[1], -u[2])
        except Exception:
            pass
    return origin, u, v


def _bar_payload(rebar, origin_mm, u_axis, v_axis):
    curves = _centerline_curves(rebar, 0, True, True)
    if not curves:
        curves = _centerline_curves(rebar, 0, True, False)
    segs = []
    for i, crv in enumerate(curves or []):
        samples = _curve_xyz_samples(crv)
        uv = _project_pts_uv(samples, origin_mm, u_axis, v_axis)
        if len(uv) < 2:
            continue
        ln_mm = internal_to_mm(_curve_length(crv))
        du, dv = _unit2(uv[-1][0] - uv[0][0], uv[-1][1] - uv[0][1])
        segs.append(
            {
                u"idx": i,
                u"len_mm": float(ln_mm),
                u"uv": uv,
                u"dir": (du, dv),
            }
        )
    d_mm = _diameter_mm(rebar)
    return {
        u"id": rebar.Id,
        u"id_int": _element_id_int(rebar.Id),
        u"n_segments": len(segs),
        u"n_positions": _cantidad_posiciones(rebar),
        u"layout": _layout_rule_name(rebar),
        u"diameter_mm": d_mm,
        u"segments": segs,
        u"skip": u"",
    }


def _eligibility_skip(rebar):
    if not isinstance(rebar, Rebar):
        return u"No es Rebar."
    if _is_in_group(rebar):
        return u"Pertenece a un Group."
    if _is_free_form(rebar):
        return u"Free-form no soportada."
    return u""


def build_elevation_session(doc, rebars, view):
    """
    Proyecta las barras al plano de alzado de la vista (o marco de respaldo).

    Returns:
        (ok, err, session)
    """
    if doc is None or not rebars:
        return False, u"No hay barras para dibujar.", None

    usable = []
    skipped = []
    for rb in rebars:
        reason = _eligibility_skip(rb)
        if reason:
            skipped.append(( _element_id_int(rb.Id), reason))
            continue
        usable.append(rb)
    if not usable:
        return False, u"Ninguna barra es editable (group / free-form).", None

    origin_xyz = None
    curves0 = _centerline_curves(usable[0], 0, True, True) or _centerline_curves(
        usable[0], 0, True, False
    )
    if curves0:
        origin_xyz = _curve_endpoints(curves0[0])[0]

    frame = None
    if view is not None and origin_xyz is not None:
        frame = _build_view_projection_frame(view, origin_xyz)
    if frame is None:
        frame = _fallback_frame_from_bars(usable)
    if frame is None:
        return False, u"No se pudo construir el plano de alzado.", None

    origin_mm, u_axis, v_axis = frame
    context_polylines_uv = []
    context_fill_rects_uv = []
    context_n_elems = 0
    if view is not None:
        try:
            (
                context_polylines_uv,
                context_n_elems,
                context_fill_rects_uv,
            ) = _collect_concrete_context_polylines_uv(
                doc,
                view,
                origin_mm,
                u_axis,
                v_axis,
                skip_element_id=usable[0].Id,
            )
        except Exception:
            context_polylines_uv = []
            context_fill_rects_uv = []
            context_n_elems = 0

    bars = []
    for rb in usable:
        payload = _bar_payload(rb, origin_mm, u_axis, v_axis)
        if payload[u"n_segments"] < 1:
            skipped.append((_element_id_int(rb.Id), u"Sin centerline."))
            continue
        bars.append(payload)

    if not bars:
        return False, u"No se pudo leer la centerline de las barras.", None

    session = {
        u"bars": bars,
        u"skipped": skipped,
        u"origin_mm": origin_mm,
        u"u_axis": u_axis,
        u"v_axis": v_axis,
        u"flip_v": True,
        u"context_polylines_uv": context_polylines_uv,
        u"context_fill_rects_uv": context_fill_rects_uv,
        u"context_n_elems": int(context_n_elems),
        u"view_id": view.Id if view is not None else None,
    }
    return True, None, session


def match_segments(session, bar_index, seg_idx):
    """
    Segmentos equivalentes en el resto de barras (mismo índice o mismo extremo + dirección).

    Returns:
        lista de (bar_index, seg_idx)
    """
    bars = list((session or {}).get(u"bars") or [])
    if bar_index < 0 or bar_index >= len(bars):
        return []
    src = bars[bar_index]
    segs = list(src.get(u"segments") or [])
    if seg_idx < 0 or seg_idx >= len(segs):
        return []
    src_dir = segs[seg_idx].get(u"dir") or (0.0, 0.0)
    src_n = int(src.get(u"n_segments") or 0)
    src_end = is_end_segment(seg_idx, src_n)
    src_is_first = int(seg_idx) == 0
    hits = [(int(bar_index), int(seg_idx))]

    for bi, bar in enumerate(bars):
        if bi == bar_index:
            continue
        other = list(bar.get(u"segments") or [])
        n = len(other)
        if n < 1:
            continue
        if n == src_n:
            hits.append((bi, int(seg_idx)))
            continue
        if not src_end or n < 2:
            continue
        candidates = [0, n - 1]
        best = None
        best_score = -1.0
        for ci in candidates:
            d = other[ci].get(u"dir") or (0.0, 0.0)
            score = abs(src_dir[0] * d[0] + src_dir[1] * d[1])
            same_end = (src_is_first and ci == 0) or ((not src_is_first) and ci == n - 1)
            if same_end:
                score += 0.15
            if score > best_score:
                best_score = score
                best = ci
        if best is not None and best_score >= _DIR_MATCH:
            hits.append((bi, int(best)))
    return hits


def _shape_segment_count(shape):
    """Número de segmentos de la definición, o None."""
    if shape is None:
        return None
    try:
        from Autodesk.Revit.DB.Structure import RebarShapeDefinitionBySegments

        defn = shape.GetRebarShapeDefinition()
        if defn is not None and isinstance(defn, RebarShapeDefinitionBySegments):
            return int(defn.NumberOfSegments)
    except Exception:
        pass
    return None


def resolve_resulting_rebar_shape(doc, n_remaining, original_shape_name=None):
    """
    RebarShape del resultado: «01»/«02»/«03» según tramos, o homónimo
    al original decrementado, o el primer shape con ese NumberOfSegments.
    """
    try:
        n = int(n_remaining)
    except Exception:
        n = 0
    if doc is None or n < 1:
        return None, u""
    key = u"{0:02d}".format(n)
    sh = find_rebar_shape_by_name(doc, key)
    if sh is not None:
        return sh, rebar_shape_display_name(sh) or key

    orig_key = normalize_shape_key(original_shape_name)
    try:
        orig_n = int(orig_key)
        if orig_n >= 2:
            alt = u"{0:02d}".format(orig_n - 1)
            sh = find_rebar_shape_by_name(doc, alt)
            if sh is not None:
                return sh, rebar_shape_display_name(sh) or alt
    except Exception:
        pass

    best = None
    best_name = u""
    try:
        from Autodesk.Revit.DB import FilteredElementCollector
        from Autodesk.Revit.DB.Structure import RebarShape

        for cand in FilteredElementCollector(doc).OfClass(RebarShape):
            if _shape_segment_count(cand) != n:
                continue
            nm = rebar_shape_display_name(cand)
            if best is None:
                best = cand
                best_name = nm
            nk = normalize_shape_key(nm)
            if nk == key or _as_unicode(nm).strip() == key:
                return cand, nm or key
    except Exception:
        pass
    return best, best_name or key


def _remaining_curves(curves, seg_idx):
    out = []
    for i, c in enumerate(curves or []):
        if i == int(seg_idx):
            continue
        cc = _clone_curve(c)
        if cc is not None:
            out.append(cc)
    return out


def _remove_one_rebar_segment(doc, rebar, seg_idx, view):
    """
    Recrea la barra sin el tramo ``seg_idx`` (solo extremos).

    Debe ejecutarse dentro de una Transaction abierta.

    Returns:
        (ok, msg, new_rebar_or_None, tag_infos)
    """
    reason = _eligibility_skip(rebar)
    if reason:
        return False, reason, None, []
    curves = _centerline_curves(rebar, 0, True, True)
    if not curves:
        curves = _centerline_curves(rebar, 0, True, False)
    n = len(curves or [])
    if n < 2:
        return False, u"La barra tiene un solo tramo; no se puede eliminar.", None, []
    if not is_end_segment(seg_idx, n):
        return (
            False,
            u"El tramo {0} es intermedio: el sketch quedaría discontinuo.".format(
                segment_letter(seg_idx)
            ),
            None,
            [],
        )
    remaining = _remaining_curves(curves, seg_idx)
    if len(remaining) < 1:
        return False, u"No quedarían tramos válidos.", None, []

    try:
        host = doc.GetElement(rebar.GetHostId())
    except Exception:
        host = None
    if host is None:
        return False, u"Sin host válido.", None, []
    bar_type = doc.GetElement(rebar.GetTypeId())
    try:
        style = rebar.Style
    except Exception:
        style = RebarStyle.Standard
    norm = _rebar_normal(rebar)
    hook_start = _hook_type(doc, rebar.GetHookTypeId(0))
    hook_end = _hook_type(doc, rebar.GetHookTypeId(1))
    try:
        so0 = rebar.GetHookOrientation(0)
        so1 = rebar.GetHookOrientation(1)
    except Exception:
        so0 = RebarHookOrientation.Right
        so1 = RebarHookOrientation.Left
    if int(seg_idx) == 0:
        hook_start = None
        so0 = RebarHookOrientation.Right
    else:
        hook_end = None
        so1 = RebarHookOrientation.Left

    n_pos = _cantidad_posiciones(rebar)
    snaps = capture_rebar_presentation(doc, rebar, view)
    orig_shape_name = u""
    try:
        orig_shape_name = get_rebar_shape_name(doc, rebar)
    except Exception:
        orig_shape_name = u""
    tag_infos = []
    try:
        tag_infos = capture_rebar_tag_infos(doc, rebar.Id) or []
    except Exception:
        tag_infos = []

    shape_obj, shape_name = resolve_resulting_rebar_shape(
        doc, len(remaining), orig_shape_name
    )
    try:
        rb_new = _create_from_curves(
            doc,
            remaining,
            host,
            norm,
            bar_type,
            style,
            hook_start,
            hook_end,
            so0,
            so1,
            rebar_shape=shape_obj,
        )
    except Exception as ex:
        return False, _exception_text(ex), None, []
    if rb_new is None:
        return False, u"CreateFromCurves no creó la barra.", None, []

    if shape_obj is not None:
        try:
            set_rebar_shape(doc, rb_new, shape_obj)
        except Exception:
            pass
        try:
            doc.Regenerate()
        except Exception:
            pass

    rule_src = _layout_rule_name(rebar)
    if n_pos > 1 or rule_src == u"MaximumSpacing":
        ok_lay, err_lay = _copy_layout(rebar, rb_new)
        if ok_lay:
            try:
                doc.Regenerate()
            except Exception:
                pass
            _copy_bar_included(rebar, rb_new, n_pos)
        elif err_lay:
            pass
    copy_armadura_instance_parameters(rebar, rb_new)
    try:
        apply_rebar_presentation(doc, [rb_new], snaps)
    except Exception:
        pass

    old_id = rebar.Id
    try:
        doc.Delete(old_id)
    except Exception as ex:
        return False, u"No se pudo borrar la barra original: {0}".format(
            _exception_text(ex)
        ), None, []
    note = u""
    if shape_name:
        note = u"Shape {0}".format(shape_name)
    return True, note, rb_new, tag_infos


def apply_remove_segments(doc, targets, view):
    """
    ``targets``: lista de (ElementId, seg_idx).

    Returns:
        (ok, mensaje, n_ok, errores, ids_nuevos)
    """
    if doc is None or not targets:
        return False, u"Nada que aplicar.", 0, [], []

    # Una transacción = un deshacer para todo el lote.
    t = Transaction(doc, _TRANSACTION_NAME)
    try:
        _attach_rebar_outside_host_swallower(t)
    except Exception:
        pass
    t.Start()
    n_ok = 0
    errors = []
    new_ids = []
    pending_tags = []
    n_tags = 0
    try:
        for eid, seg_idx in targets:
            try:
                rb = doc.GetElement(eid)
            except Exception:
                rb = None
            if not isinstance(rb, Rebar):
                errors.append(u"Id {0}: ya no existe.".format(_element_id_int(eid)))
                continue
            ok, msg, rb_new, tag_infos = _remove_one_rebar_segment(
                doc, rb, int(seg_idx), view
            )
            if ok and rb_new is not None:
                n_ok += 1
                try:
                    new_ids.append(rb_new.Id)
                except Exception:
                    pass
                if tag_infos:
                    pending_tags.append((rb_new, tag_infos))
            else:
                errors.append(
                    u"Id {0}: {1}".format(_element_id_int(eid), msg or u"falló")
                )
        if n_ok < 1:
            t.RollBack()
            return False, u"No se modificó ninguna barra.", 0, errors, []
        try:
            doc.Regenerate()
        except Exception:
            pass
        for rb_new, tag_infos in pending_tags:
            try:
                n_tags += int(tag_divided_rebars(doc, tag_infos, [rb_new]) or 0)
            except Exception as ex_tag:
                errors.append(
                    u"Etiqueta Id {0}: {1}".format(
                        _element_id_int(getattr(rb_new, u"Id", None)),
                        _exception_text(ex_tag),
                    )
                )
        if view is not None and new_ids:
            try:
                apply_unobscured_to_rebars(doc, new_ids, view)
            except Exception:
                pass
        t.Commit()
    except Exception as ex:
        try:
            t.RollBack()
        except Exception:
            pass
        return False, _exception_text(ex), 0, errors, []

    msg = u"Se eliminó el tramo en {0} barra(s).".format(n_ok)
    if n_tags:
        msg = msg + u" Se recrearon {0} etiqueta(s) con tipo según el shape.".format(
            n_tags
        )
    if errors:
        msg = msg + u" {0} no se pudieron actualizar.".format(len(errors))
    return True, msg, n_ok, errors, new_ids


def collect_all_uv(session):
    pts = []
    for bar in (session or {}).get(u"bars") or []:
        for seg in bar.get(u"segments") or []:
            for uv in seg.get(u"uv") or []:
                pts.append(uv)
    for pl in (session or {}).get(u"context_polylines_uv") or []:
        for uv in pl or []:
            pts.append(uv)
    for rect in (session or {}).get(u"context_fill_rects_uv") or []:
        try:
            u0, u1, v0, v1 = [float(x) for x in rect]
            pts.extend([[u0, v0], [u1, v1]])
        except Exception:
            pass
    return pts
