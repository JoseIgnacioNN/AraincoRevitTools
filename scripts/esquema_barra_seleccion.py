# -*- coding: utf-8 -*-
"""
Esquema dibujado de una barra seleccionada en la vista activa.

- Geometría fiel: cada segmento de la centerline (con ganchos y radios de doblez).
- Dibujo con DetailCurves a la derecha de la barra (mismo alineado vertical).
- Etiquetas por tramo y totales con ceiling a 10 mm (criterio despiece Armado Columnas).
- Texto 2.5 mm.

Revit 2024–2026 · IronPython 2.7 / pyRevit.
"""

from __future__ import print_function

import math

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    Arc,
    BuiltInCategory,
    BuiltInParameter,
    Category,
    ElementId,
    FilteredElementCollector,
    GraphicsStyle,
    GraphicsStyleType,
    HorizontalTextAlignment,
    Line,
    StorageType,
    TextNote,
    TextNoteOptions,
    TextNoteType,
    Transaction,
    UnitTypeId,
    UnitUtils,
    VerticalTextAlignment,
    View3D,
    ViewSchedule,
    ViewSheet,
    XYZ,
)
from Autodesk.Revit.DB.Structure import MultiplanarOption, Rebar, RebarBarType
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from System.Collections.Generic import List

try:
    from Autodesk.Revit.Exceptions import OperationCanceledException
except Exception:
    OperationCanceledException = Exception

from column_reinforcement.linea_fierro import (
    _despiece_diameter_text,
    _despiece_parcial_mm_ceiling_10,
)

_TOOL_TITLE = u"Arainco: Esquema barra"
OFFSET_RIGHT_MM = 2000.0
# Offset de etiqueta de tramo respecto al punto medio (hacia −Right / +Up).
SEGMENT_LABEL_OFFSET_MM = 50.0
SEGMENT_LABEL_ALONG_UP_MM = 40.0
# No etiquetar tramos menores (radios de doblez muy cortos); sí se dibujan.
MIN_SEGMENT_LABEL_MM = 30.0
TEXT_NOTE_TYPE_NAMES = (
    u"2.5mm Arial_Arrow Filled 15 Degree",
    u"2.5mm Arial",
)
TEXT_HEIGHT_MM = 2.5
# Estilo de línea del croquis (mismo criterio que despiece Armado Columnas).
LINE_STYLE_NAME = u"<Wide Lines>"
_WIDE_LINES_NAMES = (u"<Wide Lines>", u"Wide Lines")


def _u(text):
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _norm_upper(s):
    try:
        return (s or u"").strip().upper()
    except Exception:
        return u""


def _mostrar_aviso(uiapp, instruction, content=u"", ok_text=u"Entendido"):
    hwnd = None
    try:
        if uiapp is not None:
            from revit_wpf_window_position import revit_main_hwnd

            hwnd = revit_main_hwnd(uiapp)
    except Exception:
        pass
    try:
        from bimtools_instruction_dialog import show_message_dialog

        show_message_dialog(
            _TOOL_TITLE,
            instruction,
            content=content,
            ok_text=ok_text,
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass
    try:
        body = instruction
        if content:
            body = instruction + u"\n\n" + content
        TaskDialog.Show(_TOOL_TITLE, body)
    except Exception:
        pass


class _FiltroRebar(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, Rebar)

    def AllowReference(self, reference, position):
        return False


def _mm_to_ft(mm):
    return float(mm) / 304.8


def _ft_to_mm(val_ft):
    try:
        return float(
            UnitUtils.ConvertFromInternalUnits(float(val_ft), UnitTypeId.Millimeters)
        )
    except Exception:
        try:
            return float(val_ft) * 304.8
        except Exception:
            return 0.0


def _dot(a, b):
    return float(a.X) * float(b.X) + float(a.Y) * float(b.Y) + float(a.Z) * float(b.Z)


def _view_local_uv(view, pt):
    o = view.Origin
    d = XYZ(
        float(pt.X) - float(o.X),
        float(pt.Y) - float(o.Y),
        float(pt.Z) - float(o.Z),
    )
    return _dot(d, view.RightDirection), _dot(d, view.UpDirection)


def _map_uv_to_view(view, u, v):
    right = view.RightDirection
    up = view.UpDirection
    o = view.Origin
    return XYZ(
        float(o.X) + float(right.X) * float(u) + float(up.X) * float(v),
        float(o.Y) + float(right.Y) * float(u) + float(up.Y) * float(v),
        float(o.Z) + float(right.Z) * float(u) + float(up.Z) * float(v),
    )


def _view_allows_detail_curves(view):
    if view is None:
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass
    try:
        if isinstance(view, (View3D, ViewSchedule, ViewSheet)):
            return False
    except Exception:
        pass
    return True


def _rebar_from_selection_or_pick(uidoc, uiapp):
    doc = uidoc.Document
    ids = list(uidoc.Selection.GetElementIds() or [])
    if len(ids) == 1:
        el = doc.GetElement(ids[0])
        if isinstance(el, Rebar):
            return el
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            _FiltroRebar(),
            u"Selecciona la barra (Rebar) para dibujar su esquema a la derecha.",
        )
    except OperationCanceledException:
        return None
    except Exception:
        return None
    if ref is None:
        return None
    el = doc.GetElement(ref.ElementId)
    if not isinstance(el, Rebar):
        _mostrar_aviso(uiapp, u"El elemento seleccionado no es un Rebar.")
        return None
    return el


def _curves_list(raw):
    if raw is None:
        return []
    out = []
    try:
        n = raw.Count
        for i in range(n):
            c = raw[i]
            if c is not None:
                out.append(c)
    except Exception:
        try:
            for c in raw:
                if c is not None:
                    out.append(c)
        except Exception:
            pass
    return out


def _centerline_segments(rebar, bar_index=0):
    """
    Cada curva de la centerline real (ganchos + radios de doblez), en orden.
    Preferir GetTransformedCenterlineCurves.
    """
    bi = int(bar_index)
    # adjustForSelfIntersections, suppressHooks=False, suppressBendRadius=False
    args_list = (
        (False, False, False, MultiplanarOption.IncludeAllMultiplanarCurves, bi),
        (False, False, False, MultiplanarOption.IncludeOnlyPlanarCurves, bi),
        (False, False, True, MultiplanarOption.IncludeAllMultiplanarCurves, bi),
    )
    for args in args_list:
        try:
            raw = rebar.GetTransformedCenterlineCurves(*args)
            curves = _curves_list(raw)
            if curves:
                return curves
        except Exception:
            pass
    for args in args_list:
        try:
            raw = rebar.GetCenterlineCurves(*args)
            curves = _curves_list(raw)
            if curves:
                return curves
        except Exception:
            pass
    return []


def _curve_length_ft(curve):
    try:
        return abs(float(curve.Length))
    except Exception:
        try:
            return float(curve.GetEndPoint(0).DistanceTo(curve.GetEndPoint(1)))
        except Exception:
            return 0.0


def _bar_type_diameter_mm(doc, rebar):
    try:
        tid = rebar.GetTypeId()
        if tid is None or tid == ElementId.InvalidElementId:
            return None
        bt = doc.GetElement(tid)
        if not isinstance(bt, RebarBarType):
            return None
        try:
            return _ft_to_mm(bt.BarDiameter)
        except Exception:
            pass
        try:
            p = bt.get_Parameter(BuiltInParameter.REBAR_BAR_DIAMETER)
            if p is not None and p.StorageType == StorageType.Double:
                return _ft_to_mm(p.AsDouble())
        except Exception:
            pass
    except Exception:
        pass
    return None


def _bbox_uv_from_curves(view, curves):
    us = []
    vs = []
    for c in curves or []:
        try:
            pts = list(c.Tessellate() or [])
        except Exception:
            pts = []
        if len(pts) < 2:
            try:
                pts = [c.GetEndPoint(0), c.GetEndPoint(1)]
            except Exception:
                continue
        for p in pts:
            if p is None:
                continue
            u, v = _view_local_uv(view, p)
            us.append(u)
            vs.append(v)
    if not us:
        return None
    return min(us), max(us), min(vs), max(vs)


def _segment_mid_uv(view, curve):
    try:
        pm = curve.Evaluate(0.5, True)
        return _view_local_uv(view, pm)
    except Exception:
        try:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            u0, v0 = _view_local_uv(view, p0)
            u1, v1 = _view_local_uv(view, p1)
            return 0.5 * (u0 + u1), 0.5 * (v0 + v1)
        except Exception:
            return None


def _text_type_display_name(text_type):
    if text_type is None:
        return u""
    for bip in (
        BuiltInParameter.SYMBOL_NAME_PARAM,
        BuiltInParameter.ALL_MODEL_TYPE_NAME,
    ):
        try:
            p = text_type.get_Parameter(bip)
            if p is not None and p.HasValue:
                s = (p.AsString() or u"").strip()
                if s:
                    return s
        except Exception:
            pass
    try:
        return (text_type.Name or u"").strip()
    except Exception:
        return u""


def _text_type_size_mm(text_note_type):
    if text_note_type is None:
        return None
    try:
        p = text_note_type.get_Parameter(BuiltInParameter.TEXT_SIZE)
        if p is not None and p.StorageType == StorageType.Double and p.HasValue:
            return _ft_to_mm(p.AsDouble())
    except Exception:
        pass
    return None


def _apply_text_type_height_mm(text_note_type, size_mm):
    if text_note_type is None:
        return False
    size_ft = _mm_to_ft(size_mm)
    try:
        p = text_note_type.get_Parameter(BuiltInParameter.TEXT_SIZE)
        if p is not None and (not p.IsReadOnly) and p.StorageType == StorageType.Double:
            p.Set(size_ft)
            return True
    except Exception:
        pass
    return False


def _name_looks_like_25mm(name):
    nu = _norm_upper(name)
    if not nu:
        return False
    try:
        nu = nu.replace(u"\u00A0", u" ").replace(u",", u".")
    except Exception:
        pass
    if u"2.5" in nu or u"2.50" in nu:
        return True
    compact = nu.replace(u" ", u"")
    return u"2.5MM" in compact or u"2.5M" in compact


def _collect_text_note_types(document):
    by_id = {}

    def _register(tnt):
        if tnt is None:
            return
        try:
            eid = int(tnt.Id.IntegerValue)
        except Exception:
            return
        if eid not in by_id:
            by_id[eid] = tnt

    try:
        col = FilteredElementCollector(document).OfClass(TextNoteType)
        try:
            col = col.WhereElementIsElementType()
        except Exception:
            pass
        for tnt in col:
            _register(tnt)
    except Exception:
        pass
    if not by_id:
        try:
            for tnt in FilteredElementCollector(document).OfClass(
                clr.GetClrType(TextNoteType)
            ):
                _register(tnt)
        except Exception:
            pass
    return list(by_id.values())


def _ensure_text_note_type(document):
    all_types = _collect_text_note_types(document)
    if not all_types:
        return None

    def _names(tnt):
        out = []
        n = _text_type_display_name(tnt)
        if n:
            out.append(n)
        try:
            raw = (tnt.Name or u"").strip()
            if raw and raw not in out:
                out.append(raw)
        except Exception:
            pass
        return out

    for wanted in TEXT_NOTE_TYPE_NAMES:
        w = (wanted or u"").strip()
        for tnt in all_types:
            for nm in _names(tnt):
                if nm == w:
                    return tnt

    for tnt in all_types:
        for nm in _names(tnt):
            if _name_looks_like_25mm(nm) and u"ARIAL" in _norm_upper(nm):
                return tnt

    for tnt in all_types:
        for nm in _names(tnt):
            if _name_looks_like_25mm(nm):
                return tnt

    for tnt in all_types:
        sz = _text_type_size_mm(tnt)
        if sz is not None and abs(float(sz) - float(TEXT_HEIGHT_MM)) < 0.15:
            return tnt

    base = all_types[0]
    for candidate in all_types:
        for nm in _names(candidate):
            if u"ARIAL" in _norm_upper(nm):
                base = candidate
                break
    try:
        new_name = TEXT_NOTE_TYPE_NAMES[-1]
        for tnt in all_types:
            for nm in _names(tnt):
                if nm == new_name:
                    return tnt
        new_id = base.Duplicate(new_name)
        if new_id is not None and new_id != ElementId.InvalidElementId:
            nt = document.GetElement(new_id)
            if nt is not None:
                _apply_text_type_height_mm(nt, TEXT_HEIGHT_MM)
                return nt
    except Exception:
        pass
    return all_types[0]


def _draw_detail_line(document, view, p0, p1, line_style_id=None):
    try:
        if p0.DistanceTo(p1) < 1e-9:
            return None
        line = Line.CreateBound(p0, p1)
        if line is None:
            return None
        dc = document.Create.NewDetailCurve(view, line)
        _apply_detail_line_style(document, dc, line_style_id)
        return dc
    except Exception:
        return None


def _draw_curve_segment_shifted(document, view, curve, delta_u, line_style_id=None):
    """
    Dibuja un segmento de la centerline proyectado al plano de la vista,
    desplazado en +Right (delta_u). Devuelve cantidad de DetailCurves creadas.
    """
    if curve is None:
        return 0

    def world_shifted(pt):
        u, v = _view_local_uv(view, pt)
        return _map_uv_to_view(view, float(u) + float(delta_u), v)

    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
    except Exception:
        p0 = p1 = None

    if isinstance(curve, Arc) and p0 is not None and p1 is not None:
        try:
            pm = curve.Evaluate(0.5, True)
            arc = Arc.Create(world_shifted(p0), world_shifted(p1), world_shifted(pm))
            if arc is not None:
                dc = document.Create.NewDetailCurve(view, arc)
                _apply_detail_line_style(document, dc, line_style_id)
                return 1
        except Exception:
            pass

    if isinstance(curve, Line) and p0 is not None and p1 is not None:
        if _draw_detail_line(
            document, view, world_shifted(p0), world_shifted(p1), line_style_id
        ):
            return 1

    pts = []
    try:
        pts = list(curve.Tessellate() or [])
    except Exception:
        pts = []
    if len(pts) < 2 and p0 is not None and p1 is not None:
        pts = [p0, p1]
    drawn = 0
    for i in range(len(pts) - 1):
        if _draw_detail_line(
            document,
            view,
            world_shifted(pts[i]),
            world_shifted(pts[i + 1]),
            line_style_id,
        ):
            drawn += 1
    return drawn


def _get_lines_category(document):
    try:
        return Category.GetCategory(document, BuiltInCategory.OST_Lines)
    except Exception:
        return None


def _graphics_style_id_for_subcategory(sub):
    if sub is None:
        return None
    try:
        gs = sub.GetGraphicsStyle(GraphicsStyleType.Projection)
        if gs is not None:
            return gs.Id
    except Exception:
        pass
    return None


def _line_style_display_name(document, style_element_id):
    if document is None or style_element_id is None:
        return u""
    el = document.GetElement(style_element_id)
    if el is None:
        return u""
    try:
        nm = el.Name
        if nm:
            return (nm or u"").strip()
    except Exception:
        pass
    try:
        cg = getattr(el, u"GraphicsStyleCategory", None)
        if cg is not None:
            return (cg.Name or u"").strip()
    except Exception:
        pass
    return u""


def _resolve_wide_lines_style_id(document):
    """GraphicsStyle.Id de «Wide Lines» / «<Wide Lines>»."""
    targets = set()
    for name in _WIDE_LINES_NAMES:
        targets.add(_norm_upper(name))
        bare = name.strip(u"<>").strip()
        if bare:
            targets.add(_norm_upper(bare))

    fallback_id = None
    lines_cat = _get_lines_category(document)
    if lines_cat is not None:
        try:
            for sub in lines_cat.SubCategories:
                try:
                    nm_u = _norm_upper(sub.Name)
                except Exception:
                    continue
                style_id = _graphics_style_id_for_subcategory(sub)
                if style_id is None:
                    continue
                if nm_u in targets:
                    return style_id
                for t in list(targets):
                    bare = t.strip(u"<>").strip()
                    if bare and (nm_u == bare or bare in nm_u or u"WIDE" in nm_u):
                        fallback_id = style_id
                        break
        except Exception:
            pass

    if fallback_id is not None:
        return fallback_id

    try:
        p_iv = lines_cat.Id.IntegerValue if lines_cat is not None else None
    except Exception:
        p_iv = None
    for gs in FilteredElementCollector(document).OfClass(GraphicsStyle):
        try:
            if gs.GraphicsStyleType != GraphicsStyleType.Projection:
                continue
        except Exception:
            pass
        try:
            cg = getattr(gs, u"GraphicsStyleCategory", None)
            if cg is None:
                cg = getattr(gs, u"Category", None)
            if cg is None:
                continue
            if p_iv is not None:
                pc = getattr(cg, u"Parent", None)
                if pc is None or pc.Id.IntegerValue != p_iv:
                    continue
            nm_u = _norm_upper(cg.Name)
            if nm_u in targets or u"WIDE" in nm_u:
                return gs.Id
        except Exception:
            pass
    return None


def _pick_applicable_line_style_id(document, detail_curve, preferred_id):
    target = _norm_upper(LINE_STYLE_NAME)
    bare = target.strip(u"<>").strip()
    applicable = []
    try:
        applicable = list(detail_curve.GetLineStyleIds())
    except Exception:
        pass
    if not applicable:
        return preferred_id
    if preferred_id is not None:
        try:
            piv = preferred_id.IntegerValue
            for aid in applicable:
                if aid.IntegerValue == piv:
                    return aid
        except Exception:
            pass
    for aid in applicable:
        nm_u = _norm_upper(_line_style_display_name(document, aid))
        if nm_u == target or (bare and (nm_u == bare or bare in nm_u)):
            return aid
        if u"WIDE" in nm_u:
            return aid
    return preferred_id


def _apply_detail_line_style(document, detail_curve, style_id):
    if detail_curve is None:
        return
    style_id = _pick_applicable_line_style_id(document, detail_curve, style_id)
    if style_id is None:
        return
    try:
        if style_id == ElementId.InvalidElementId:
            return
    except Exception:
        return
    try:
        detail_curve.LineStyleId = style_id
    except Exception:
        pass
    try:
        gs = document.GetElement(style_id)
        if gs is not None:
            detail_curve.LineStyle = gs
    except Exception:
        pass


def _create_label(document, view, u, v, text, text_note_type, rotation_rad=0.0):
    if not text or text_note_type is None:
        return None
    origin = _map_uv_to_view(view, u, v)
    tn = None
    try:
        opts = TextNoteOptions(text_note_type.Id)
        try:
            opts.HorizontalAlignment = HorizontalTextAlignment.Center
        except Exception:
            pass
        try:
            opts.VerticalAlignment = VerticalTextAlignment.Middle
        except Exception:
            pass
        try:
            opts.Rotation = float(rotation_rad)
        except Exception:
            pass
        tn = TextNote.Create(document, view.Id, origin, text, opts)
    except Exception:
        pass
    if tn is None:
        try:
            opts = TextNoteOptions()
            opts.TypeId = text_note_type.Id
            tn = TextNote.Create(document, view.Id, origin, text, opts)
        except Exception:
            pass
    return tn


def _principal_segment(segments):
    """Segmento de mayor longitud (tramo principal de la barra)."""
    best = None
    best_len = -1.0
    for c in segments or []:
        L = _curve_length_ft(c)
        if L > best_len:
            best_len = L
            best = c
    return best


def _build_caption(diam_mm, partials_ceiling):
    """ø12 L=total (p1+p2+…) con ceiling 10 mm por segmento (criterio despiece)."""
    dia = _despiece_diameter_text(diam_mm)
    if not partials_ceiling:
        return u"{0} L=0".format(dia)
    total = sum(int(p) for p in partials_ceiling)
    if len(partials_ceiling) == 1:
        return u"{0} L={1}".format(dia, partials_ceiling[0])
    parts = u"+".join([u"{0}".format(p) for p in partials_ceiling])
    return u"{0} L={1} ({2})".format(dia, total, parts)


def _segment_label_rotation_rad(view, curve):
    """Rota la etiqueta según la dirección del tramo en el plano de la vista."""
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        u0, v0 = _view_local_uv(view, p0)
        u1, v1 = _view_local_uv(view, p1)
        ang = math.atan2(float(v1) - float(v0), float(u1) - float(u0))
        # Mantener texto legible (no boca abajo).
        if ang > math.pi * 0.5:
            ang -= math.pi
        elif ang < -math.pi * 0.5:
            ang += math.pi
        return ang
    except Exception:
        return 0.0


def _caption_anchor_uv(view, principal, u_max, v_min, v_max, delta_u):
    """
    Ancla de la etiqueta ø/L junto al segmento principal del croquis
    (desplazado a la derecha del trazo).
    """
    mid = _segment_mid_uv(view, principal) if principal is not None else None
    if mid is None:
        return (
            float(u_max) + float(delta_u) + _mm_to_ft(SEGMENT_LABEL_OFFSET_MM),
            0.5 * (float(v_min) + float(v_max)),
        )
    mu, mv = mid
    # Offset perpendicular al tramo principal en el plano de la vista (+90°).
    try:
        p0 = principal.GetEndPoint(0)
        p1 = principal.GetEndPoint(1)
        u0, v0 = _view_local_uv(view, p0)
        u1, v1 = _view_local_uv(view, p1)
        du = float(u1) - float(u0)
        dv = float(v1) - float(v0)
        L = math.sqrt(du * du + dv * dv)
        if L < 1e-12:
            nu, nv = 1.0, 0.0
        else:
            # Normal hacia el lado “exterior” preferente (+Right del croquis).
            nu, nv = -dv / L, du / L
            if nu < 0.0:
                nu, nv = -nu, -nv
        off = _mm_to_ft(SEGMENT_LABEL_OFFSET_MM)
        return (
            float(mu) + float(delta_u) + nu * off,
            float(mv) + nv * off,
        )
    except Exception:
        return (
            float(mu) + float(delta_u) + _mm_to_ft(SEGMENT_LABEL_OFFSET_MM),
            float(mv),
        )


def run(revit_app):
    """Entrada pyRevit: ``run(__revit__)``."""
    try:
        uiapp = revit_app
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        view = uidoc.ActiveView
    except Exception:
        _mostrar_aviso(None, u"No hay documento activo.")
        return

    if not _view_allows_detail_curves(view):
        _mostrar_aviso(
            uiapp,
            u"La vista activa no admite líneas de detalle.",
            content=u"Usa una planta, sección, alzado o vista de dibujo.",
        )
        return

    rebar = _rebar_from_selection_or_pick(uidoc, uiapp)
    if rebar is None:
        return

    segments = _centerline_segments(rebar, 0)
    if not segments:
        _mostrar_aviso(
            uiapp,
            u"No se pudo obtener la geometría (segmentos) de la barra.",
        )
        return

    bb = _bbox_uv_from_curves(view, segments)
    if bb is None:
        _mostrar_aviso(uiapp, u"La proyección de la barra en esta vista está vacía.")
        return

    u_min, u_max, v_min, v_max = bb
    width = max(0.0, float(u_max) - float(u_min))
    # Borde izquierdo del croquis = borde derecho de la barra + OFFSET.
    delta_u = width + _mm_to_ft(OFFSET_RIGHT_MM)
    if abs(delta_u) < 1e-12:
        delta_u = _mm_to_ft(OFFSET_RIGHT_MM)

    diam_mm = _bar_type_diameter_mm(doc, rebar)
    if diam_mm is None:
        diam_mm = 12.0

    # Largos por segmento (mm crudos) → ceiling 10 mm para etiquetas.
    seg_mm_raw = []
    for c in segments:
        seg_mm_raw.append(_ft_to_mm(_curve_length_ft(c)))

    partials_ceiling = []
    for mm in seg_mm_raw:
        if mm < 1e-6:
            continue
        partials_ceiling.append(_despiece_parcial_mm_ceiling_10(mm))

    caption = _build_caption(diam_mm, partials_ceiling)

    t = Transaction(doc, u"Arainco: Esquema barra seleccionada")
    t.Start()
    try:
        line_style_id = _resolve_wide_lines_style_id(doc)
        drawn = 0
        for c in segments:
            drawn += int(
                _draw_curve_segment_shifted(
                    doc, view, c, delta_u, line_style_id=line_style_id
                )
                or 0
            )

        if drawn < 1:
            t.RollBack()
            _mostrar_aviso(
                uiapp,
                u"No se pudo dibujar ningún segmento en la vista activa.",
            )
            return

        tnt = _ensure_text_note_type(doc)
        if tnt is None:
            t.RollBack()
            _mostrar_aviso(
                uiapp,
                u"No hay ningún TextNoteType en el proyecto.",
                content=u"Carga una plantilla con tipos de texto (p. ej. 2.5 mm).",
            )
            return

        # Etiqueta por cada segmento significativo.
        for c, mm in zip(segments, seg_mm_raw):
            if mm < float(MIN_SEGMENT_LABEL_MM):
                continue
            mid = _segment_mid_uv(view, c)
            if mid is None:
                continue
            mu, mv = mid
            ceil_mm = _despiece_parcial_mm_ceiling_10(mm)
            rot = _segment_label_rotation_rad(view, c)
            # Ligeramente hacia −Right y +Up respecto al trazo.
            _create_label(
                doc,
                view,
                float(mu) + float(delta_u) - _mm_to_ft(SEGMENT_LABEL_OFFSET_MM),
                float(mv) + _mm_to_ft(SEGMENT_LABEL_ALONG_UP_MM),
                u"{0}".format(ceil_mm),
                tnt,
                rotation_rad=rot,
            )

        # Resumen ø / L orientado según el segmento principal.
        principal = _principal_segment(segments)
        cap_rot = (
            _segment_label_rotation_rad(view, principal)
            if principal is not None
            else 0.0
        )
        label_u, label_v = _caption_anchor_uv(
            view, principal, u_max, v_min, v_max, delta_u
        )
        _create_label(
            doc, view, label_u, label_v, caption, tnt, rotation_rad=cap_rot
        )

        t.Commit()
    except Exception as ex:
        if t.HasStarted():
            t.RollBack()
        _mostrar_aviso(uiapp, u"Error al generar el esquema.", content=_u(ex))
        return

    try:
        ids = List[ElementId]()
        ids.Add(rebar.Id)
        uidoc.Selection.SetElementIds(ids)
    except Exception:
        pass


def _run_rps():
    try:
        run(__revit__)
    except NameError:
        print(u"Ejecuta desde pyRevit (falta __revit__).")
