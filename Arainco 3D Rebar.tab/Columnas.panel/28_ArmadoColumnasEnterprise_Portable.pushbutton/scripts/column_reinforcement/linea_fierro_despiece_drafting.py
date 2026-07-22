# -*- coding: utf-8 -*-
"""
Vista de dibujo (Drafting View) con despiece por línea de fierro (vista **1:50**; geometría en unidades reales del modelo).

Geometría de barras en unidades internas Revit (pies): ``span_seg``, ``pata_hook_ft_seg``.
Posición vertical del despiece: cada pieza a su largo completo; la siguiente inicia
``L_anterior − traslape`` más abajo (solape visible como en plano de taller).
Dentro de cada línea: tramos alternan eje **izquierdo / derecho**; **50 mm** entre ejes.
Traslape vertical sin cambio (``L − traslape``). Entre líneas A, B, C…: **800 mm**, izq. → der.
Cota de arranque: una fila horizontal; ``y_base = y_ref + (z_start − z_min)`` en escala 1:1 mm.

Además, en vista activa de **sección o alzado**, el mismo despiece se dibuja a
``ACTIVE_VIEW_DESPIECE_OFFSET_MM`` (2000 mm) a la derecha del conjunto de pilares armados.
Cada ejecución **añade** un nuevo esquema sin borrar los anteriores.
Antes del croquis se crea un **lienzo** (masking region) que oculta el modelo detrás.
"""

from __future__ import print_function

import math

import clr
clr.AddReference("RevitAPI")

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    Arc,
    BuiltInCategory,
    BuiltInParameter,
    Category,
    CurveElement,
    CurveLoop,
    DetailCurve,
    ElementId,
    FilledRegion,
    FilteredElementCollector,
    GraphicsStyle,
    GraphicsStyleType,
    HorizontalTextAlignment,
    Line,
    LinePatternElement,
    OverrideGraphicSettings,
    Plane,
    Sketch,
    SketchPlane,
    TextNote,
    TextNoteOptions,
    TextNoteType,
    Transaction,
    VerticalTextAlignment,
    View,
    ViewDrafting,
    ViewFamily,
    ViewFamilyType,
    ViewSection,
    ViewType,
    XYZ,
)

try:
    from Autodesk.Revit.DB import LeaderAtachement
except Exception:
    try:
        from Autodesk.Revit.DB import LeaderAttachment as LeaderAtachement
    except Exception:
        LeaderAtachement = None

from column_reinforcement.linea_fierro import (
    etiqueta_despiece_mm,
    etiqueta_empalme_mm,
    fingerprint_seg_linea_fierro,
    linea_fierro_key_from_seg_jobs,
    linea_fierro_label_map_from_keys,
    linea_fierro_sort_index_from_letter,
    pata_flags_from_fingerprint,
)


def _stamp_lienzo_conjunto_guid(element, conjunto_guid=None):
    """Asigna ``Armadura_Conjunto_GUID`` a curvas, textos y regiones del lienzo."""
    if element is None:
        return False
    try:
        from conjunto_guid import stamp_armadura_conjunto_guid
        return stamp_armadura_conjunto_guid(element, conjunto_guid=conjunto_guid)
    except Exception:
        return False


DRAFTING_VIEW_NAME = u"Arainco Despiece Columnas"
DEPIECE_LINE_STYLE_NAME = u"<Wide Lines>"
LINEA_FIERRO_MARKER_LINE_STYLE_NAME = u"<Thin Lines>"
INVISIBLE_LINE_STYLE_NAME = u"<Invisible lines>"
TEXT_NOTE_TYPE_NAME = u"2.5mm Arial_Arrow Filled 15 Degree"
# Texto paralelo al tramo vertical (90° respecto al eje X local de la vista).
TEXT_NOTE_ROTATION_RAD = math.pi * 0.5
DRAFTING_VIEW_SCALE = 50

MARGIN_MM = 80.0
TRAMO_HORIZONTAL_GAP_MM = 50.0
LINEA_FIERRO_GAP_MM = 800.0
LABEL_GAP_MM = 8.0
TITLE_GAP_MM = 10.0
TEXT_OFFSET_MM = 6.0
# Marcador de línea de fierro: círculo R=100 mm, letra, directriz 100 mm hacia la 1.ª barra.
LINEA_FIERRO_MARKER_CIRCLE_RADIUS_MM = 100.0
LINEA_FIERRO_MARKER_LEADER_MM = 100.0
# Separación entre el extremo superior de la directriz y la base de la primera barra.
LINEA_FIERRO_MARKER_GAP_BELOW_FIRST_BAR_MM = 150.0
LINEA_FIERRO_MARKER_CELL_PAD_MM = 8.0
LINEA_FIERRO_MARKER_OFFSET_LEFT_MM = 25.0
LINEA_FIERRO_MARKER_OFFSET_UP_MM = 50.0
# Separación etiqueta ↔ eje de la barra: siempre a la izquierda del trazo (eje local −X).
SEGMENT_LABEL_OFFSET_MM = 50.0
# Etiqueta de empalme: a la izquierda del par de tramos en la zona de traslape.
LAP_LABEL_OFFSET_MM = 50.0
CELL_SIDE_MARGIN_MM = 40.0
# Separación del despiece respecto al borde derecho (vista local +X) del conjunto de pilares.
ACTIVE_VIEW_DESPIECE_OFFSET_MM = 2000.0
# Margen alrededor del croquis para el lienzo (masking region) en vista activa.
MASK_CANVAS_PAD_MM = 120.0


def _z_line_start_mm(z_ft):
    """Cota de arranque del hilo en mm (modelo)."""
    try:
        return round(float(z_ft) * 304.8, 3)
    except Exception:
        return 0.0


def _dot_xyz(a, b):
    return float(a.X) * float(b.X) + float(a.Y) * float(b.Y) + float(a.Z) * float(b.Z)


def _view_local_xy(view, pt):
    """Coordenadas en el plano de la vista (Right, Up) respecto a ``view.Origin``."""
    o = view.Origin
    d = XYZ(float(pt.X) - float(o.X), float(pt.Y) - float(o.Y), float(pt.Z) - float(o.Z))
    return _dot_xyz(d, view.RightDirection), _dot_xyz(d, view.UpDirection)


def _bbox_corners_world(bb):
    if bb is None:
        return []
    try:
        tr = bb.Transform
        mn, mx = bb.Min, bb.Max
        out = []
        for x in (float(mn.X), float(mx.X)):
            for y in (float(mn.Y), float(mx.Y)):
                for z in (float(mn.Z), float(mx.Z)):
                    out.append(tr.OfPoint(XYZ(x, y, z)))
        return out
    except Exception:
        try:
            return [bb.Min, bb.Max]
        except Exception:
            return []


def _is_section_or_elevation_view(view):
    if view is None:
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass
    try:
        if isinstance(view, ViewDrafting):
            return False
    except Exception:
        pass
    try:
        if isinstance(view, ViewSection):
            return True
    except Exception:
        pass
    try:
        vt = view.ViewType
        return vt == ViewType.Section or vt == ViewType.Elevation
    except Exception:
        return False


def _columns_max_view_local_right(view, columns):
    max_r = None
    for col in columns or []:
        bb = None
        try:
            bb = col.get_BoundingBox(view)
        except Exception:
            pass
        if bb is None:
            try:
                bb = col.get_BoundingBox(None)
            except Exception:
                bb = None
        for pt in _bbox_corners_world(bb):
            x_loc, _ = _view_local_xy(view, pt)
            if max_r is None or x_loc > max_r:
                max_r = x_loc
    return float(max_r) if max_r is not None else 0.0


def _columns_reference_xyz_at_z(columns, z_ft):
    xs = []
    ys = []
    for col in columns or []:
        bb = None
        try:
            bb = col.get_BoundingBox(None)
        except Exception:
            pass
        if bb is None:
            continue
        try:
            xs.append(0.5 * (float(bb.Min.X) + float(bb.Max.X)))
            ys.append(0.5 * (float(bb.Min.Y) + float(bb.Max.Y)))
        except Exception:
            pass
    if not xs:
        return XYZ(0.0, 0.0, float(z_ft))
    return XYZ(sum(xs) / len(xs), sum(ys) / len(ys), float(z_ft))


def _view_max_local_right(doc, view):
    """Extremo derecho (+X local) del contenido ya dibujado en la vista."""
    max_r = [None]
    view_id = view.Id

    def _consider_pt(pt):
        if pt is None:
            return
        try:
            x_loc, _ = _view_local_xy(view, pt)
            if max_r[0] is None or x_loc > max_r[0]:
                max_r[0] = x_loc
        except Exception:
            pass

    for ce in FilteredElementCollector(doc).OfClass(CurveElement):
        try:
            if ce is None or ce.OwnerViewId != view_id:
                continue
            if not isinstance(ce, DetailCurve):
                continue
            curve = ce.GeometryCurve
            if curve is None:
                continue
            try:
                for pt in curve.Tessellate():
                    _consider_pt(pt)
            except Exception:
                _consider_pt(curve.GetEndPoint(0))
                _consider_pt(curve.GetEndPoint(1))
        except Exception:
            pass

    try:
        for el in FilteredElementCollector(doc, view_id):
            if isinstance(el, TextNote):
                bb = el.get_BoundingBox(view)
                for pt in _bbox_corners_world(bb):
                    _consider_pt(pt)
    except Exception:
        pass

    return max_r[0]


def _resolve_invisible_line_style(document):
    """Estilo de proyección invisible (inglés / español)."""
    preferred = []
    fallback = []
    for gs in FilteredElementCollector(document).OfClass(GraphicsStyle):
        try:
            if gs.GraphicsStyleType != GraphicsStyleType.Projection:
                continue
            name = _norm_upper(gs.Name)
            low = (gs.Name or u"").lower()
            if u"INVISIBLE" in name or u"INVISIBLES" in name:
                if u"<" in low:
                    preferred.append(gs)
                else:
                    fallback.append(gs)
        except Exception:
            pass
    if preferred:
        return preferred[0]
    if fallback:
        return fallback[0]
    return None


def _ensure_view_sketch_plane(document, view):
    try:
        if view.SketchPlane is not None:
            return True
    except Exception:
        pass
    try:
        plane = Plane.CreateByNormalAndOrigin(view.ViewDirection, view.Origin)
        sketch_plane = SketchPlane.Create(document, plane)
        view.SketchPlane = sketch_plane
        return True
    except Exception:
        return False


def _local_rect_to_curve_loop(view, x_min_ft, y_min_ft, x_max_ft, y_max_ft):
    right = view.RightDirection
    up = view.UpDirection
    o = view.Origin

    def _pt(xl, yl):
        return XYZ(
            o.X + right.X * xl + up.X * yl,
            o.Y + right.Y * xl + up.Y * yl,
            o.Z + right.Z * xl + up.Z * yl,
        )

    p1 = _pt(x_min_ft, y_min_ft)
    p2 = _pt(x_max_ft, y_min_ft)
    p3 = _pt(x_max_ft, y_max_ft)
    p4 = _pt(x_min_ft, y_max_ft)
    loop = CurveLoop()
    loop.Append(Line.CreateBound(p1, p2))
    loop.Append(Line.CreateBound(p2, p3))
    loop.Append(Line.CreateBound(p3, p4))
    loop.Append(Line.CreateBound(p4, p1))
    return loop


def _pick_invisible_line_style_id_for_filled_region(document, filled_region, invisible_style):
    """Elige un estilo invisible válido para ``FilledRegion.SetLineStyleId``."""
    if filled_region is None or invisible_style is None:
        return None
    preferred_id = invisible_style.Id
    valid = []
    try:
        valid = list(filled_region.GetValidLineStyleIds())
    except Exception:
        pass
    target_names = (u"INVISIBLE", u"INVISIBLES")

    def _matches_invisible(style_id):
        nm = _norm_upper(_line_style_display_name(document, style_id))
        return any(t in nm for t in target_names)

    for sid in valid:
        if _matches_invisible(sid):
            return sid
    if preferred_id is not None:
        try:
            if FilledRegion.IsValidLineStyleId(document, preferred_id):
                return preferred_id
        except Exception:
            pass
        try:
            piv = preferred_id.IntegerValue
            for sid in valid:
                try:
                    if sid.IntegerValue == piv:
                        return sid
                except Exception:
                    pass
        except Exception:
            pass
    return valid[0] if valid else None


def _resolve_invisible_line_pattern_id(document):
    for lp in FilteredElementCollector(document).OfClass(LinePatternElement):
        try:
            low = (lp.Name or u"").lower()
            if u"invisible" in low or u"invisibles" in low:
                return lp.Id
        except Exception:
            pass
    return None


def _apply_filled_region_view_overrides(document, view, filled_region):
    """Refuerzo gráfico en la vista: patrón de línea invisible en proyección/corte."""
    if view is None or filled_region is None:
        return
    pat_id = _resolve_invisible_line_pattern_id(document)
    if pat_id is None:
        return
    try:
        ogs = OverrideGraphicSettings()
        ogs.SetProjectionLinePatternId(pat_id)
        ogs.SetCutLinePatternId(pat_id)
        view.SetElementOverrides(filled_region.Id, ogs)
    except Exception:
        pass


def _apply_invisible_boundary_to_filled_region(document, view, filled_region):
    """
    Contorno invisible del lienzo: ``FilledRegion.SetLineStyleId`` (API oficial)
    + curvas del sketch + overrides de vista.
    """
    invisible_style = _resolve_invisible_line_style(document)
    if filled_region is None:
        return False
    applied = False
    style_id = _pick_invisible_line_style_id_for_filled_region(
        document, filled_region, invisible_style
    )
    if style_id is not None:
        try:
            filled_region.SetLineStyleId(style_id)
            applied = True
        except Exception:
            pass
    if invisible_style is not None:
        _apply_invisible_style_to_filled_region_sketch(document, filled_region)
    _apply_filled_region_view_overrides(document, view, filled_region)
    try:
        document.Regenerate()
    except Exception:
        pass
    return applied


def _apply_invisible_style_to_filled_region_sketch(document, filled_region):
    invisible_style = _resolve_invisible_line_style(document)
    if invisible_style is None or filled_region is None:
        return
    try:
        document.Regenerate()
    except Exception:
        pass
    sketch = None
    try:
        for dep_id in filled_region.GetDependentElements(None):
            elem = document.GetElement(dep_id)
            if isinstance(elem, Sketch):
                sketch = elem
                break
    except Exception:
        sketch = None
    if sketch is None:
        return
    try:
        for line_id in sketch.GetAllElements():
            line_elem = document.GetElement(line_id)
            if isinstance(line_elem, CurveElement):
                _apply_invisible_line_style_to_curve_element(
                    document, line_elem, invisible_style
                )
    except Exception:
        pass
    try:
        document.Regenerate()
    except Exception:
        pass


def _compute_despiece_row_bounds_ft(
    row_items,
    x_origin_ft,
    y_sheet_origin_ft,
    line_gap_ft=None,
):
    """Envolvente local (+X derecha, +Y arriba) de una fila A→B→C…"""
    if not row_items:
        return None
    if line_gap_ft is None:
        line_gap_ft = _mm_to_ft(LINEA_FIERRO_GAP_MM)

    z_min_mm = min(_z_line_start_mm(it[7]) for it in row_items)
    x_cursor = float(x_origin_ft)
    y_sheet_origin = float(y_sheet_origin_ft)
    y_min = None
    y_max = None
    x_min = x_cursor
    x_max = x_cursor

    marker_left_ft = _mm_to_ft(
        LINEA_FIERRO_MARKER_OFFSET_LEFT_MM + LINEA_FIERRO_MARKER_CIRCLE_RADIUS_MM
    )
    label_left_ft = _mm_to_ft(SEGMENT_LABEL_OFFSET_MM + CELL_SIDE_MARGIN_MM)

    for item in row_items:
        cw = float(item[2])
        z_start_ft = float(item[7])
        bot_m = float(item[8])
        top_m = float(item[9])
        bar_h_ft = float(item[10])

        elev_off_ft = _mm_to_ft(
            max(0.0, _z_line_start_mm(z_start_ft) - z_min_mm)
        )
        y_bar_base = y_sheet_origin + bot_m + elev_off_ft
        cy, _, _, r_ft, _ = _linea_fierro_marker_geometry_ft(y_bar_base)
        cy += _mm_to_ft(LINEA_FIERRO_MARKER_OFFSET_UP_MM)
        y_cell_bottom = cy - r_ft
        y_cell_top = y_bar_base + bar_h_ft + top_m

        if y_min is None or y_cell_bottom < y_min:
            y_min = y_cell_bottom
        if y_max is None or y_cell_top > y_max:
            y_max = y_cell_top

        x_cell_min = x_cursor - max(marker_left_ft, label_left_ft)
        x_cell_max = x_cursor + cw
        if x_cell_min < x_min:
            x_min = x_cell_min
        if x_cell_max > x_max:
            x_max = x_cell_max

        x_cursor += cw + line_gap_ft

    pad_ft = _mm_to_ft(MASK_CANVAS_PAD_MM)
    return (x_min - pad_ft, y_min - pad_ft, x_max + pad_ft, y_max + pad_ft)


def _create_despiece_masking_canvas(document, view, x_min_ft, y_min_ft, x_max_ft, y_max_ft):
    """Lienzo blanco (masking region) que oculta el modelo detrás del despiece."""
    if x_max_ft <= x_min_ft or y_max_ft <= y_min_ft:
        return None
    try:
        if view.ViewType == ViewType.ThreeD:
            return None
    except Exception:
        pass
    if not _ensure_view_sketch_plane(document, view):
        return None
    try:
        loop = _local_rect_to_curve_loop(view, x_min_ft, y_min_ft, x_max_ft, y_max_ft)
        loop_list = List[CurveLoop]()
        loop_list.Add(loop)
        region = FilledRegion.CreateMaskingRegion(document, view.Id, loop_list)
        _apply_invisible_boundary_to_filled_region(document, view, region)
        _stamp_lienzo_conjunto_guid(region)
        return region
    except Exception:
        return None


def build_seg_pata_flags_by_line_idx(line_plans, line_rb_accum):
    """
    Alinea patas por ``seg_i`` (no por ``troceo_ui_i``), coherente con ``seg_jobs`` del plan.
    """
    out = {}
    for lp in line_plans or []:
        line_idx = lp.get("line_idx")
        items = line_rb_accum.get(line_idx) or []
        by_seg_i = {}
        for item in items:
            try:
                _troceo_ui, seg_i, _rb, fp = item
                by_seg_i[int(seg_i)] = pata_flags_from_fingerprint(fp)
            except Exception:
                pass
        segs = sorted(lp.get("seg_jobs") or [], key=lambda s: int(s["seg_i"]))
        if not segs:
            continue
        flags_list = []
        for sj in segs:
            si = int(sj["seg_i"])
            if si in by_seg_i:
                flags_list.append(by_seg_i[si])
            else:
                flags_list.append(
                    (
                        bool(sj.get("want_bot_pata")),
                        bool(sj.get("want_top_pata")),
                    )
                )
        out[line_idx] = flags_list
    return out


def _segment_pata_flags(seg_index, n_segments, did_b, did_t, seg_job):
    """
    Patas en el croquis: inferior solo en tramo 0, superior solo en el último tramo.
    Usa flags del Rebar creado; respaldo ``want_*`` del ``seg_job``.
    """
    n = max(1, int(n_segments))
    i = int(seg_index)
    try:
        bot = bool(did_b[i]) if did_b is not None and i < len(did_b) else False
        top = bool(did_t[i]) if did_t is not None and i < len(did_t) else False
    except Exception:
        bot = False
        top = False
    if not bot and not top and seg_job is not None:
        bot = bool(seg_job.get("want_bot_pata"))
        top = bool(seg_job.get("want_top_pata"))
    if n > 1:
        bot = bot and i == 0
        top = top and i == n - 1
    return bot, top


def _mm_to_ft(mm):
    return float(mm) / 304.8


def _traslape_mm_for_diameter(diameter_mm):
    try:
        import bimtools_rebar_hook_lengths as bhl
        fn = getattr(bhl, "traslape_mm_from_nominal_diameter_mm", None)
        if fn is not None:
            v = fn(float(diameter_mm))
            if v is not None and float(v) > 1e-6:
                return float(v)
    except Exception:
        pass
    return 400.0


def _set_view_scale(view, scale=DRAFTING_VIEW_SCALE):
    if view is None:
        return
    try:
        view.Scale = int(scale)
    except Exception:
        try:
            view.Scale = scale
        except Exception:
            pass


def _first_drafting_view_family_type(document):
    for vft in FilteredElementCollector(document).OfClass(ViewFamilyType):
        try:
            if vft and vft.ViewFamily == ViewFamily.Drafting:
                return vft
        except Exception:
            pass
    return None


def _find_drafting_view_by_name(document, view_name):
    for v in FilteredElementCollector(document).OfClass(ViewDrafting):
        try:
            if v and v.Name == view_name:
                return v
        except Exception:
            pass
    return None


def _create_or_get_drafting_view(document, vft, view_name):
    v = _find_drafting_view_by_name(document, view_name)
    if v:
        _set_view_scale(v)
        return v
    for v in FilteredElementCollector(document).OfClass(View):
        try:
            if v and v.Name == view_name and not isinstance(v, ViewDrafting):
                raise RuntimeError(
                    u'Ya existe una vista "{}" que no es de dibujo.'.format(view_name)
                )
        except RuntimeError:
            raise
        except Exception:
            pass
    v = ViewDrafting.Create(document, vft.Id)
    v.Name = view_name
    _set_view_scale(v)
    return v


def _clear_detail_curves_in_view(document, drafting_view):
    view_id = drafting_view.Id
    to_delete = []
    for ce in FilteredElementCollector(document).OfClass(CurveElement):
        try:
            if ce is None or ce.OwnerViewId != view_id:
                continue
            if not isinstance(ce, DetailCurve):
                continue
            to_delete.append(ce.Id)
        except Exception:
            pass
    for did in to_delete:
        try:
            document.Delete(did)
        except Exception:
            pass


def _clear_text_notes_in_view(document, view):
    view_id = view.Id
    for el in FilteredElementCollector(document, view_id):
        try:
            if isinstance(el, TextNote):
                document.Delete(el.Id)
        except Exception:
            pass


def _clear_drafting_view_content(document, drafting_view):
    _clear_detail_curves_in_view(document, drafting_view)
    _clear_text_notes_in_view(document, drafting_view)


def _map_local_to_view_plane(view, x_local_ft, y_local_ft):
    right = view.RightDirection
    up = view.UpDirection
    o = view.Origin
    return XYZ(
        o.X + right.X * x_local_ft + up.X * y_local_ft,
        o.Y + right.Y * x_local_ft + up.Y * y_local_ft,
        o.Z + right.Z * x_local_ft + up.Z * y_local_ft,
    )


def _norm_upper(s):
    try:
        return (s or u"").strip().upper()
    except Exception:
        return u""


def _get_lines_category(document):
    try:
        lc = Category.GetCategory(document, BuiltInCategory.OST_Lines)
        if lc is not None:
            return lc
    except Exception:
        pass
    try:
        return document.Settings.Categories.get_Item(BuiltInCategory.OST_Lines)
    except Exception:
        return None


def _iter_lines_subcategories(lines_category):
    if lines_category is None:
        return
    try:
        subcats = lines_category.SubCategories
    except Exception:
        return
    if subcats is None:
        return
    try:
        for sub in subcats:
            if sub is not None:
                yield sub
        return
    except Exception:
        pass
    try:
        it = subcats.ForwardIterator()
        while it.MoveNext():
            try:
                entry = it.Current
                sub = entry
                try:
                    if not hasattr(sub, "Name"):
                        sub = entry.Value
                except Exception:
                    pass
                if sub is not None:
                    yield sub
            except Exception:
                pass
    except Exception:
        pass


def _graphics_style_id_for_subcategory(subcategory):
    if subcategory is None:
        return None
    try:
        gs = subcategory.GetGraphicsStyle(GraphicsStyleType.Projection)
        if gs is not None:
            return gs.Id
    except Exception:
        pass
    try:
        return subcategory.Id
    except Exception:
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
        cg = getattr(el, "GraphicsStyleCategory", None)
        if cg is not None:
            return (cg.Name or u"").strip()
    except Exception:
        pass
    return u""


def _resolve_lines_style_id(document, style_name, fallback_keywords=None):
    """``GraphicsStyle.Id`` de una subcategoría Líneas (p. ej. ``<Thin Lines>``)."""
    target = _norm_upper(style_name)
    if not target:
        return None
    lines_cat = _get_lines_category(document)
    fallback_id = None
    keywords = tuple(fallback_keywords or ())
    for sub in _iter_lines_subcategories(lines_cat):
        try:
            nm = (sub.Name or u"").strip()
        except Exception:
            continue
        nm_u = _norm_upper(nm)
        style_id = _graphics_style_id_for_subcategory(sub)
        if style_id is None:
            continue
        if nm_u == target:
            return style_id
        bare = target.strip("<>").strip()
        if bare and (nm_u == bare or bare in nm_u):
            fallback_id = style_id
        elif fallback_id is None and keywords and all(kw in nm_u for kw in keywords):
            fallback_id = style_id
    if fallback_id is not None:
        return fallback_id
    if lines_cat is None:
        return None
    try:
        p_iv = lines_cat.Id.IntegerValue
    except Exception:
        p_iv = None
    for gs in FilteredElementCollector(document).OfClass(GraphicsStyle):
        try:
            cg = getattr(gs, "GraphicsStyleCategory", None)
            if cg is None:
                cg = getattr(gs, "Category", None)
            if cg is None:
                continue
            if p_iv is not None:
                pc = getattr(cg, "Parent", None)
                if pc is None or pc.Id.IntegerValue != p_iv:
                    continue
            nm_u = _norm_upper(cg.Name)
            if nm_u == target:
                return gs.Id
        except Exception:
            pass
    return None


def _resolve_despiece_line_style_id(document):
    return _resolve_lines_style_id(
        document, DEPIECE_LINE_STYLE_NAME, (u"WIDE", u"LINE")
    )


def _resolve_marker_line_style_id(document):
    return _resolve_lines_style_id(
        document, LINEA_FIERRO_MARKER_LINE_STYLE_NAME, (u"THIN", u"LINE")
    )


def _pick_applicable_line_style_id(document, detail_curve, preferred_id, style_name):
    """Elige un ``LineStyleId`` válido para esta ``DetailCurve``."""
    target = _norm_upper(style_name)
    bare = target.strip("<>").strip()
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
    return preferred_id


def _apply_detail_line_style(document, detail_curve, style_id, style_name=None):
    if detail_curve is None:
        return
    if style_name is None:
        style_name = DEPIECE_LINE_STYLE_NAME
    style_id = _pick_applicable_line_style_id(
        document, detail_curve, style_id, style_name
    )
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


def _apply_invisible_line_style_to_curve_element(
    document, curve_elem, invisible_style
):
    """``LineStyleId`` + validación aplicable (mismo patrón que detail curves)."""
    if curve_elem is None or invisible_style is None:
        return
    style_name = invisible_style.Name or INVISIBLE_LINE_STYLE_NAME
    style_id = _pick_applicable_line_style_id(
        document, curve_elem, invisible_style.Id, style_name
    )
    if style_id is None:
        return
    try:
        if style_id == ElementId.InvalidElementId:
            return
    except Exception:
        return
    try:
        curve_elem.LineStyleId = style_id
    except Exception:
        pass
    try:
        gs = document.GetElement(style_id)
        if gs is not None:
            curve_elem.LineStyle = gs
        else:
            curve_elem.LineStyle = invisible_style
    except Exception:
        pass


def _linea_fierro_marker_geometry_ft(y_base_ft):
    """
    Posición del marcador bajo la primera barra (eje Y local hacia arriba).

    Orden de abajo a arriba: círculo (R) → directriz 100 mm → separación 150 mm → ``y_base``.
    """
    r_ft = _mm_to_ft(LINEA_FIERRO_MARKER_CIRCLE_RADIUS_MM)
    leader_ft = _mm_to_ft(LINEA_FIERRO_MARKER_LEADER_MM)
    gap_ft = _mm_to_ft(LINEA_FIERRO_MARKER_GAP_BELOW_FIRST_BAR_MM)
    y_leader_top = float(y_base_ft) - gap_ft
    y_circle_top = y_leader_top - leader_ft
    cy = y_circle_top - r_ft
    return cy, y_circle_top, y_leader_top, r_ft, leader_ft


def _linea_fierro_marker_reserved_height_ft():
    """Margen inferior de celda: gap 150 + directriz + diámetro círculo + pad."""
    r = _mm_to_ft(LINEA_FIERRO_MARKER_CIRCLE_RADIUS_MM)
    leader = _mm_to_ft(LINEA_FIERRO_MARKER_LEADER_MM)
    gap = _mm_to_ft(LINEA_FIERRO_MARKER_GAP_BELOW_FIRST_BAR_MM)
    pad = _mm_to_ft(LINEA_FIERRO_MARKER_CELL_PAD_MM)
    return gap + leader + 2.0 * r + pad


def _draw_detail_arc(document, view, arc, line_style_id=None, line_style_name=None):
    if arc is None:
        return None
    try:
        dc = document.Create.NewDetailCurve(view, arc)
    except Exception:
        return None
    _apply_detail_line_style(document, dc, line_style_id, line_style_name)
    _stamp_lienzo_conjunto_guid(dc)
    return dc


def _draw_circle_detail(
    document, view, cx_ft, cy_ft, radius_ft, line_style_id=None, line_style_name=None
):
    """Circunferencia en el plano de la vista de dibujo."""
    origin = _map_local_to_view_plane(view, cx_ft, cy_ft)
    r = float(radius_ft)
    if r < 1e-9:
        return
    right = view.RightDirection
    up = view.UpDirection
    try:
        arc = Arc.Create(origin, r, 0.0, 2.0 * math.pi, right, up)
        if (
            _draw_detail_arc(
                document, view, arc, line_style_id, line_style_name
            )
            is not None
        ):
            return
    except Exception:
        pass
    try:
        a1 = Arc.Create(origin, r, 0.0, math.pi, right, up)
        a2 = Arc.Create(origin, r, math.pi, 2.0 * math.pi, right, up)
        _draw_detail_arc(document, view, a1, line_style_id, line_style_name)
        _draw_detail_arc(document, view, a2, line_style_id, line_style_name)
    except Exception:
        pass


def _draw_linea_fierro_marker(
    document,
    view,
    cx_ft,
    y_base_ft,
    letter,
    text_note_type,
    marker_line_style_id=None,
    marker_line_style_name=None,
):
    """
    Identificador de línea de fierro: letra en círculo (R=100 mm), directriz 100 mm
    desde el cuadrante superior; el extremo de la directriz queda 150 mm bajo la 1.ª barra.
    """
    cx = float(cx_ft) - _mm_to_ft(LINEA_FIERRO_MARKER_OFFSET_LEFT_MM)
    cy, y_circle_top, y_leader_top, r_ft, _leader_ft = _linea_fierro_marker_geometry_ft(
        y_base_ft
    )
    dy = _mm_to_ft(LINEA_FIERRO_MARKER_OFFSET_UP_MM)
    cy += dy
    y_circle_top += dy
    y_leader_top += dy

    _draw_circle_detail(
        document,
        view,
        cx,
        cy,
        r_ft,
        marker_line_style_id,
        marker_line_style_name,
    )

    p0 = _map_local_to_view_plane(view, cx, y_circle_top)
    p1 = _map_local_to_view_plane(view, cx, y_leader_top)
    _draw_detail_line(
        document,
        view,
        p0,
        p1,
        marker_line_style_id,
        marker_line_style_name,
    )

    lbl = (letter or u"?").strip()
    if lbl:
        _create_text_note_on_plane(
            document,
            view,
            cx,
            cy,
            lbl,
            text_note_type,
            rotation_rad=0.0,
            h_align=HorizontalTextAlignment.Center,
            v_align=VerticalTextAlignment.Middle,
            keep_rotated_readable=True,
            segment_label=False,
        )


def _draw_detail_line(
    document, view, p1, p2, line_style_id=None, line_style_name=None
):
    line = Line.CreateBound(p1, p2)
    if line is None:
        return None
    try:
        dc = document.Create.NewDetailCurve(view, line)
    except Exception:
        return None
    _apply_detail_line_style(document, dc, line_style_id, line_style_name)
    _stamp_lienzo_conjunto_guid(dc)
    return dc


def _active_revit_document(fallback_document=None):
    """Documento activo pyRevit; si no hay, el pasado al despiece."""
    if fallback_document is not None:
        return fallback_document
    try:
        return __revit__.ActiveUIDocument.Document  # noqa: F821
    except Exception:
        return None


def _text_note_type_display_name(text_type):
    """
    Nombre del tipo de texto (mismo criterio que el script de inspección en consola):
    ``BuiltInParameter.SYMBOL_NAME_PARAM``.
    """
    if text_type is None:
        return u""
    try:
        p = text_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p and p.HasValue:
            s = (p.AsString() or u"").strip()
            if s:
                return s
    except Exception:
        pass
    for bip in (BuiltInParameter.ALL_MODEL_TYPE_NAME,):
        try:
            p = text_type.get_Parameter(bip)
            if p and p.HasValue:
                s = (p.AsString() or u"").strip()
                if s:
                    return s
        except Exception:
            pass
    try:
        s = (text_type.Name or u"").strip()
        if s:
            return s
    except Exception:
        pass
    return u""


def _collect_text_note_type_elements(document):
    """
    ``FilteredElementCollector(doc).OfClass(TextNoteType)`` — tipos de anotación de texto
    del proyecto (independientes de la vista activa).
    """
    doc = _active_revit_document(document)
    if doc is None:
        return []
    by_id = {}

    def _register(text_type):
        if text_type is None:
            return
        try:
            eid = int(text_type.Id.IntegerValue)
        except Exception:
            return
        if eid not in by_id:
            by_id[eid] = text_type

    try:
        for text_type in FilteredElementCollector(doc).OfClass(TextNoteType):
            _register(text_type)
    except Exception:
        pass
    if not by_id:
        try:
            for text_type in FilteredElementCollector(doc).OfClass(
                clr.GetClrType(TextNoteType)
            ):
                _register(text_type)
        except Exception:
            pass
    return list(by_id.values())


def get_project_text_note_types(document):
    """
    Todos los ``TextNoteType`` del proyecto.

    Returns:
        list: ``[(nombre, TextNoteType), ...]`` ordenado por nombre.
    """
    out = []
    seen_names = set()
    for tnt in _collect_text_note_type_elements(document):
        nm = _text_note_type_display_name(tnt)
        if not nm:
            try:
                nm = u"(id {0})".format(int(tnt.Id.IntegerValue))
            except Exception:
                nm = u"(sin nombre)"
        key = (nm, int(tnt.Id.IntegerValue))
        if key in seen_names:
            continue
        seen_names.add(key)
        out.append((nm, tnt))
    try:
        out.sort(key=lambda pair: pair[0])
    except Exception:
        pass
    return out


def find_text_note_type_by_name(document, type_name, text_note_types=None):
    """
    Busca un ``TextNoteType`` por nombre **exacto** (sin aproximaciones).

    Args:
        document: Documento Revit.
        type_name: Nombre del tipo, p. ej. ``2.5mm Arial_Arrow Filled 15 Degree``.
        text_note_types: Lista opcional de ``get_project_text_note_types``; si no se
            pasa, se recorre el proyecto.

    Returns:
        ``TextNoteType`` o ``None``.
    """
    target = (type_name or u"").strip()
    if not target:
        return None
    pairs = text_note_types
    if pairs is None:
        pairs = get_project_text_note_types(document)
    for nm, tnt in pairs:
        if nm == target:
            return tnt
    return None


def _duplicate_text_note_type(document, base_type, new_name):
    if document is None or base_type is None:
        return None
    try:
        new_id = base_type.Duplicate((new_name or u"").strip())
        if new_id is None or new_id == ElementId.InvalidElementId:
            return None
        return document.GetElement(new_id)
    except Exception:
        return None


def _pick_base_text_note_type_for_duplicate(all_types):
    """Prefiere un tipo existente parecido a 2.5 mm Arial para duplicar."""
    if not all_types:
        return None
    target_u = _norm_upper(TEXT_NOTE_TYPE_NAME)
    for nm, tnt in all_types:
        nu = _norm_upper(nm)
        if u"2.5" in nu and u"ARIAL" in nu:
            return tnt
    for nm, tnt in all_types:
        nu = _norm_upper(nm)
        if u"ARIAL" in nu:
            return tnt
    return all_types[0][1]


def _ensure_text_note_type(document):
    """
    Tipo de texto obligatorio ``TEXT_NOTE_TYPE_NAME``.
    Si no existe pero hay otros tipos, duplica uno con ese nombre (dentro de Transaction).
    """
    all_types = get_project_text_note_types(document)
    tnt = find_text_note_type_by_name(
        document, TEXT_NOTE_TYPE_NAME, text_note_types=all_types
    )
    if tnt is not None:
        return tnt
    if all_types:
        base = _pick_base_text_note_type_for_duplicate(all_types)
        tnt = _duplicate_text_note_type(document, base, TEXT_NOTE_TYPE_NAME)
        if tnt is not None:
            return tnt
        for _nm, candidate in all_types:
            tnt = _duplicate_text_note_type(document, candidate, TEXT_NOTE_TYPE_NAME)
            if tnt is not None:
                return tnt
        names_block = u"\n".join([nm for nm, _ in all_types])
    else:
        names_block = (
            u"(no se encontraron tipos de texto; revise Anotar > Tipos de texto en el proyecto)"
        )
    raise RuntimeError(
        u'El despiece requiere el tipo de texto "{0}". No se encontró en el proyecto.\n\n'
        u"Tipos de texto disponibles:\n{1}".format(TEXT_NOTE_TYPE_NAME, names_block)
    )


def _apply_text_note_options(
    opts,
    rotation_rad,
    h_align,
    v_align,
    keep_rotated_readable=False,
):
    try:
        opts.Rotation = float(rotation_rad)
    except Exception:
        pass
    try:
        opts.HorizontalAlignment = h_align
    except Exception:
        pass
    try:
        opts.VerticalAlignment = v_align
    except Exception:
        pass
    try:
        opts.KeepRotatedTextReadable = bool(keep_rotated_readable)
    except Exception:
        pass


def _configure_segment_label_text_note(text_note):
    """
    Etiquetas de tramo: Horizontal Center, Vertical Bottom;
    enganches Left Top, Right Bottom (``LeaderAtachement`` API).
    """
    if text_note is None:
        return
    try:
        text_note.KeepRotatedTextReadable = False
    except Exception:
        pass
    try:
        text_note.HorizontalAlignment = HorizontalTextAlignment.Center
    except Exception:
        pass
    try:
        text_note.VerticalAlignment = VerticalTextAlignment.Bottom
    except Exception:
        pass
    if LeaderAtachement is not None:
        try:
            text_note.LeaderLeftAttachment = LeaderAtachement.TopLine
        except Exception:
            pass
        try:
            text_note.LeaderRightAttachment = LeaderAtachement.BottomLine
        except Exception:
            pass


def _create_text_note_on_plane(
    document,
    view,
    x_ft,
    y_ft,
    txt,
    text_note_type,
    rotation_rad=0.0,
    h_align=HorizontalTextAlignment.Center,
    v_align=VerticalTextAlignment.Top,
    keep_rotated_readable=False,
    segment_label=False,
):
    if not txt or text_note_type is None:
        return None
    origin = _map_local_to_view_plane(view, x_ft, y_ft)
    tn = None
    try:
        opts = TextNoteOptions(text_note_type.Id)
        _apply_text_note_options(
            opts, rotation_rad, h_align, v_align, keep_rotated_readable
        )
        tn = TextNote.Create(document, view.Id, origin, txt, opts)
    except Exception:
        pass
    if tn is None:
        try:
            opts = TextNoteOptions()
            opts.TypeId = text_note_type.Id
            _apply_text_note_options(
                opts, rotation_rad, h_align, v_align, keep_rotated_readable
            )
            tn = TextNote.Create(document, view.Id, origin, txt, opts)
        except Exception:
            pass
    if tn is not None:
        _stamp_lienzo_conjunto_guid(tn)
        if segment_label:
            _configure_segment_label_text_note(tn)
    return tn


def _segment_label_x_ft(x_bar_ft):
    """Etiqueta siempre a la izquierda del eje de la barra (``−X`` local), a 50 mm."""
    return float(x_bar_ft) - _mm_to_ft(SEGMENT_LABEL_OFFSET_MM)


def _lap_mm_for_seg(sj, diam_resolver):
    try:
        d_mm = float(diam_resolver(sj))
    except Exception:
        d_mm = 12.0
    return _traslape_mm_for_diameter(d_mm)


def _create_lap_zone_label(
    document,
    view,
    x_ft,
    y_center_ft,
    lap_mm,
    text_note_type,
):
    """Etiqueta vertical en la zona de empalme, p. ej. ``(860)``."""
    txt = etiqueta_empalme_mm(lap_mm)
    _create_text_note_on_plane(
        document,
        view,
        x_ft,
        y_center_ft,
        txt,
        text_note_type,
        rotation_rad=TEXT_NOTE_ROTATION_RAD,
        h_align=HorizontalTextAlignment.Center,
        v_align=VerticalTextAlignment.Bottom,
        keep_rotated_readable=False,
        segment_label=True,
    )


def _create_segment_label_vertical(
    document,
    view,
    x_bar_ft,
    y_mid_ft,
    txt,
    text_note_type,
):
    """
    Etiqueta por tramo: vertical (90°), centrada en el medio del tramo,
    a la izquierda de la barra dibujada (``SEGMENT_LABEL_OFFSET_MM``).
    """
    x_lbl = _segment_label_x_ft(x_bar_ft)
    _create_text_note_on_plane(
        document,
        view,
        x_lbl,
        y_mid_ft,
        txt,
        text_note_type,
        rotation_rad=TEXT_NOTE_ROTATION_RAD,
        h_align=HorizontalTextAlignment.Center,
        v_align=VerticalTextAlignment.Bottom,
        keep_rotated_readable=False,
        segment_label=True,
    )


def _seg_height_ft(sj, fp=None):
    """Alto del tramo vertical (pies). Prioriza ``lz`` de la huella modelada."""
    if fp is not None:
        try:
            if len(fp) > 1:
                lz_mm = float(fp[1])
                if lz_mm > 1e-9:
                    return _mm_to_ft(lz_mm)
        except Exception:
            pass
    try:
        return max(1e-9, abs(float(sj.get("span_seg", 0.0))))
    except Exception:
        return 1e-9


def _hook_length_ft(sj, want_hook, fp=None, want_bot=False, want_top=False):
    if not want_hook:
        return 0.0
    if fp is not None:
        try:
            kind = int(fp[0])
            if kind == 1 and want_bot:
                return _mm_to_ft(float(fp[2]))
            if kind == 2 and want_top:
                return _mm_to_ft(float(fp[2]))
            if kind == 3:
                if want_bot:
                    return _mm_to_ft(float(fp[2]))
                if want_top:
                    return _mm_to_ft(float(fp[3]))
        except Exception:
            pass
    try:
        return max(0.0, float(sj.get("pata_hook_ft_seg", 0.0)))
    except Exception:
        return 0.0


def _lap_offset_ft(sj, diam_resolver):
    """Desfase horizontal entre tramos = traslape tabular real (escala 1:1)."""
    try:
        d_mm = float(diam_resolver(sj))
    except Exception:
        d_mm = 12.0
    return _mm_to_ft(_traslape_mm_for_diameter(d_mm))


def _draw_segment_schematic(
    document,
    view,
    x_center_ft,
    y_bottom_ft,
    height_ft,
    hook_ft,
    want_bot,
    want_top,
    line_style_id=None,
):
    x = float(x_center_ft)
    y0 = float(y_bottom_ft)
    y1 = y0 + float(height_ft)
    p_bot = _map_local_to_view_plane(view, x, y0)
    p_top = _map_local_to_view_plane(view, x, y1)
    _draw_detail_line(document, view, p_bot, p_top, line_style_id)
    h_ft = float(hook_ft)
    if h_ft > 1e-9:
        if want_bot:
            p_h0 = _map_local_to_view_plane(view, x, y0)
            p_h1 = _map_local_to_view_plane(view, x + h_ft, y0)
            _draw_detail_line(document, view, p_h0, p_h1, line_style_id)
        if want_top:
            p_h0 = _map_local_to_view_plane(view, x, y1)
            p_h1 = _map_local_to_view_plane(view, x + h_ft, y1)
            _draw_detail_line(document, view, p_h0, p_h1, line_style_id)


def _representative_plan_for_key(line_plans, key, seg_pata_flags):
    for lp in line_plans or []:
        segs = sorted(lp.get("seg_jobs") or [], key=lambda s: int(s["seg_i"]))
        if not segs:
            continue
        flags = seg_pata_flags.get(lp.get("line_idx"))
        if flags is None:
            flags = [
                (bool(s.get("want_bot_pata")), bool(s.get("want_top_pata")))
                for s in segs
            ]
        did_b = [f[0] for f in flags]
        did_t = [f[1] for f in flags]
        k = linea_fierro_key_from_seg_jobs(segs, did_b, did_t)
        if k == key:
            return lp, segs, did_b, did_t
    return None, [], [], []


def _schematic_layout_with_laps_ft(segs, diam_resolver, segment_fingerprints=None):
    """
    Despiece con traslape (convención de taller):

    - Pieza 0: base en y=0, largo L0.
    - Pieza i>0: base en y = y_prev + L_{i-1} − traslape_{i-1}.
    - Cada pieza se dibuja a su largo modelado (``lz`` de huella o ``span_seg``).

    Devuelve ``[(y_bottom_rel_ft, height_ft), ...]`` por ``seg_i`` ascendente.
    """
    layouts = []
    y_rel = 0.0
    for i, sj in enumerate(segs):
        fp_i = None
        if segment_fingerprints is not None and i < len(segment_fingerprints):
            fp_i = segment_fingerprints[i]
        h_ft = _seg_height_ft(sj, fp_i)
        layouts.append((y_rel, h_ft))
        if i >= len(segs) - 1:
            break
        lap_ft = _lap_offset_ft(sj, diam_resolver)
        y_rel += max(0.0, h_ft - lap_ft)
    return layouts


def _schematic_envelope_height_ft(layouts):
    if not layouts:
        return 0.0
    y0, h = layouts[-1]
    return max(1e-9, float(y0) + float(h))


def _axis_separation_ft(segs, diam_resolver=None):
    """Distancia fija entre eje izquierdo y derecho de tramos alternados (misma línea)."""
    if not segs or len(segs) < 2:
        return 0.0
    return _mm_to_ft(TRAMO_HORIZONTAL_GAP_MM)


def _segment_x_center_alternating(cx_cell, seg_index, axis_sep_ft):
    """
    Alterna eje como en despiece de taller: tramo 0 (base) izquierda, 1 derecha, 2 izq., …
    """
    half = float(axis_sep_ft) * 0.5
    if int(seg_index) % 2 == 0:
        return float(cx_cell) - half
    return float(cx_cell) + half


def _cell_bar_extent_ft(segs, did_b, did_t, diam_resolver, segment_fingerprints=None):
    """Caja del croquis: ancho (2 ejes + pata) y alto (envolvente con traslapos)."""
    n = len(segs)
    if n < 1:
        return 0.0, 0.0
    layouts = _schematic_layout_with_laps_ft(
        segs, diam_resolver, segment_fingerprints
    )
    total_h = _schematic_envelope_height_ft(layouts)
    max_hook = 0.0
    for i, sj in enumerate(segs):
        db = did_b[i] if i < len(did_b) else False
        dt = did_t[i] if i < len(did_t) else False
        fp_i = None
        if segment_fingerprints is not None and i < len(segment_fingerprints):
            fp_i = segment_fingerprints[i]
        max_hook = max(
            max_hook,
            _hook_length_ft(sj, db or dt, fp_i, want_bot=db, want_top=dt),
        )
    axis_sep = _axis_separation_ft(segs, diam_resolver) if n > 1 else 0.0
    width = axis_sep + max_hook + _mm_to_ft(
        CELL_SIDE_MARGIN_MM + SEGMENT_LABEL_OFFSET_MM
    )
    return width, total_h


def _cell_layout_margins_ft(n_labels):
    """Márgenes fijos alrededor del croquis (texto no escala la barra)."""
    return (
        _mm_to_ft(TITLE_GAP_MM + TEXT_OFFSET_MM + LABEL_GAP_MM * max(1, n_labels)),
        _linea_fierro_marker_reserved_height_ft(),
    )


def _draw_linea_fierro_cell(
    document,
    view,
    text_note_type,
    cell_origin_x_ft,
    cell_origin_y_ft,
    letter,
    segs,
    did_b,
    did_t,
    diam_resolver,
    line_style_id=None,
    marker_line_style_id=None,
    marker_line_style_name=None,
    segment_fingerprints=None,
    y_bar_base_ft=None,
):
    n = len(segs)
    if n < 1:
        return 0.0, 0.0

    layouts = _schematic_layout_with_laps_ft(
        segs, diam_resolver, segment_fingerprints
    )
    bar_w_ft, bar_h_ft = _cell_bar_extent_ft(
        segs, did_b, did_t, diam_resolver, segment_fingerprints
    )
    top_margin_ft, bot_margin_ft = _cell_layout_margins_ft(n)

    cx_cell = float(cell_origin_x_ft) + bar_w_ft * 0.5
    if y_bar_base_ft is not None:
        y_base_ft = float(y_bar_base_ft)
    else:
        y_base_ft = float(cell_origin_y_ft) + bot_margin_ft

    axis_sep_ft = _axis_separation_ft(segs, diam_resolver)

    for i, sj in enumerate(segs):
        y_rel, h_ft = layouts[i]
        y_bottom = y_base_ft + y_rel
        db, dt = _segment_pata_flags(i, n, did_b, did_t, sj)
        fp_i = None
        if segment_fingerprints is not None and i < len(segment_fingerprints):
            fp_i = segment_fingerprints[i]
        hook_ft = _hook_length_ft(sj, db or dt, fp_i, want_bot=db, want_top=dt)
        x_seg = _segment_x_center_alternating(cx_cell, i, axis_sep_ft)
        _draw_segment_schematic(
            document,
            view,
            x_seg,
            y_bottom,
            h_ft,
            hook_ft,
            db,
            dt,
            line_style_id,
        )
        if fp_i is not None:
            fp = fp_i
        else:
            fp = fingerprint_seg_linea_fierro(
                sj["span_seg"],
                db,
                dt,
                sj.get("pata_hook_ft_seg", 0.0),
            )
        lbl = etiqueta_despiece_mm(diam_resolver(sj), fp)
        y_mid_ft = y_bottom + h_ft * 0.5
        _create_segment_label_vertical(
            document,
            view,
            x_seg,
            y_mid_ft,
            lbl,
            text_note_type,
        )

    lap_off_ft = _mm_to_ft(LAP_LABEL_OFFSET_MM)
    for i in range(n - 1):
        sj = segs[i]
        lap_ft = _lap_offset_ft(sj, diam_resolver)
        if lap_ft < 1e-9:
            continue
        y_rel_i, h_i = layouts[i]
        y_overlap_center = y_base_ft + float(y_rel_i) + float(h_i) - lap_ft * 0.5
        x_i = _segment_x_center_alternating(cx_cell, i, axis_sep_ft)
        x_next = _segment_x_center_alternating(cx_cell, i + 1, axis_sep_ft)
        x_lap = min(x_i, x_next) - lap_off_ft
        lap_mm = _lap_mm_for_seg(sj, diam_resolver)
        _create_lap_zone_label(
            document,
            view,
            x_lap,
            y_overlap_center,
            lap_mm,
            text_note_type,
        )

    _draw_linea_fierro_marker(
        document,
        view,
        cx_cell,
        y_base_ft,
        letter,
        text_note_type,
        marker_line_style_id,
        marker_line_style_name,
    )

    cell_w = bar_w_ft + _mm_to_ft(CELL_SIDE_MARGIN_MM)
    cell_h = bar_h_ft + top_margin_ft + bot_margin_ft
    return cell_w, cell_h


def _prepare_despiece_row_items(model_groups, label_map, rebar_nominal_diameter_mm_fn):
    def _diam(sj):
        bt = sj.get("layout_bar_type_seg")
        if bt is None:
            return 12.0
        try:
            v = rebar_nominal_diameter_mm_fn(bt)
            return float(v) if v is not None else 12.0
        except Exception:
            return 12.0

    row_items = []
    for grp in model_groups:
        k = grp["key"]
        segs = grp.get("seg_jobs") or []
        pata_flags = grp.get("pata_flags") or []
        if not segs:
            continue
        did_b = [f[0] for f in pata_flags]
        did_t = [f[1] for f in pata_flags]
        bw, bh = _cell_bar_extent_ft(
            segs, did_b, did_t, _diam, grp.get("fingerprints")
        )
        top_m, bot_m = _cell_layout_margins_ft(len(segs))
        cw = bw + _mm_to_ft(CELL_SIDE_MARGIN_MM)
        letter = label_map.get(k, u"?")
        z_start_ft = float(
            grp.get("z_line_start_ft")
            if grp.get("z_line_start_ft") is not None
            else (grp.get("line_plan") or {}).get("z_line_start", 0.0)
        )
        row_items.append(
            (
                k,
                letter,
                cw,
                segs,
                did_b,
                did_t,
                grp.get("fingerprints"),
                z_start_ft,
                bot_m,
                top_m,
                bh,
            )
        )

    row_items.sort(
        key=lambda it: linea_fierro_sort_index_from_letter(it[1]),
    )
    return row_items


def _draw_despiece_row_items(
    doc,
    view,
    row_items,
    tnt,
    line_style_id,
    marker_line_style_id,
    marker_line_style_name,
    rebar_nominal_diameter_mm_fn,
    x_origin_ft,
    y_sheet_origin_ft,
    line_gap_ft=None,
):
    if not row_items:
        return

    if line_gap_ft is None:
        line_gap_ft = _mm_to_ft(LINEA_FIERRO_GAP_MM)

    def _diam(sj):
        bt = sj.get("layout_bar_type_seg")
        if bt is None:
            return 12.0
        try:
            v = rebar_nominal_diameter_mm_fn(bt)
            return float(v) if v is not None else 12.0
        except Exception:
            return 12.0

    z_min_mm = min(_z_line_start_mm(it[7]) for it in row_items)
    x_cursor = float(x_origin_ft)
    y_sheet_origin = float(y_sheet_origin_ft)

    for (
        k,
        letter,
        cw,
        segs,
        did_b,
        did_t,
        seg_fps,
        z_start_ft,
        bot_m,
        top_m,
        bar_h_ft,
    ) in row_items:
        elev_off_ft = _mm_to_ft(
            max(0.0, _z_line_start_mm(z_start_ft) - z_min_mm)
        )
        y_bar_base = y_sheet_origin + bot_m + elev_off_ft
        _draw_linea_fierro_cell(
            doc,
            view,
            tnt,
            x_cursor,
            y_sheet_origin,
            letter,
            segs,
            did_b,
            did_t,
            _diam,
            line_style_id,
            marker_line_style_id,
            marker_line_style_name,
            segment_fingerprints=seg_fps,
            y_bar_base_ft=y_bar_base,
        )
        x_cursor += cw + line_gap_ft


def generate_despiece_drafting_view(
    doc,
    model_groups,
    label_map,
    rebar_nominal_diameter_mm_fn,
    manage_transaction=True,
):
    """
    Vista de despiece alineada al modelo: ``model_groups`` y ``label_map`` deben
    provenir de ``collect_linea_fierro_model_groups`` (misma regla que
    ``Armadura_Ubicacion``).
    """
    if not model_groups or not label_map:
        return None

    vft = _first_drafting_view_family_type(doc)
    if vft is None:
        return None

    t = None
    if manage_transaction:
        t = Transaction(doc, u"Arainco: Despiece líneas fierro columnas")
        t.Start()
    try:
        view = _create_or_get_drafting_view(doc, vft, DRAFTING_VIEW_NAME)
        tnt = _ensure_text_note_type(doc)
        line_style_id = _resolve_despiece_line_style_id(doc)
        marker_line_style_id = _resolve_marker_line_style_id(doc)
        marker_line_style_name = LINEA_FIERRO_MARKER_LINE_STYLE_NAME

        row_items = _prepare_despiece_row_items(
            model_groups, label_map, rebar_nominal_diameter_mm_fn
        )

        if not row_items:
            if manage_transaction and t is not None:
                t.Commit()
            return view

        margin_ft = _mm_to_ft(MARGIN_MM)
        line_gap_ft = _mm_to_ft(LINEA_FIERRO_GAP_MM)
        existing_max = _view_max_local_right(doc, view)
        if existing_max is None:
            x_origin = margin_ft
        else:
            x_origin = existing_max + line_gap_ft
        _draw_despiece_row_items(
            doc,
            view,
            row_items,
            tnt,
            line_style_id,
            marker_line_style_id,
            marker_line_style_name,
            rebar_nominal_diameter_mm_fn,
            x_origin,
            margin_ft,
        )

        if manage_transaction and t is not None:
            t.Commit()
        return view
    except Exception:
        if (
            manage_transaction
            and t is not None
            and t.HasStarted()
            and not t.HasEnded()
        ):
            t.RollBack()
        raise


def _refresh_active_view(uidoc):
    if uidoc is None:
        return
    try:
        uidoc.RefreshActiveView()
    except Exception:
        pass


def generate_despiece_in_active_section_view(
    doc,
    view,
    columns,
    model_groups,
    label_map,
    rebar_nominal_diameter_mm_fn,
    manage_transaction=True,
    uidoc=None,
):
    """
    Dibujo del despiece en la vista activa (sección o alzado), a
    ``ACTIVE_VIEW_DESPIECE_OFFSET_MM`` a la derecha del conjunto de pilares.
    Antes del croquis crea un lienzo (masking region) que oculta el modelo detrás.
    No elimina despieces generados en ejecuciones anteriores.
    """
    if not model_groups or not label_map:
        return None
    if not _is_section_or_elevation_view(view):
        return None

    t = None
    if manage_transaction:
        t = Transaction(doc, u"Arainco: Despiece líneas fierro en vista")
        t.Start()
    try:
        tnt = _ensure_text_note_type(doc)
        line_style_id = _resolve_despiece_line_style_id(doc)
        marker_line_style_id = _resolve_marker_line_style_id(doc)
        marker_line_style_name = LINEA_FIERRO_MARKER_LINE_STYLE_NAME

        row_items = _prepare_despiece_row_items(
            model_groups, label_map, rebar_nominal_diameter_mm_fn
        )
        if not row_items:
            if manage_transaction and t is not None:
                t.Commit()
            return view

        z_min_mm = min(_z_line_start_mm(it[7]) for it in row_items)
        z_min_ft = z_min_mm / 304.8
        ref_pt = _columns_reference_xyz_at_z(columns, z_min_ft)
        _, y_local_ref = _view_local_xy(view, ref_pt)
        max_bot_m = max(it[8] for it in row_items)
        y_sheet_origin = y_local_ref - max_bot_m
        x_origin = _columns_max_view_local_right(view, columns) + _mm_to_ft(
            ACTIVE_VIEW_DESPIECE_OFFSET_MM
        )

        bounds = _compute_despiece_row_bounds_ft(
            row_items,
            x_origin,
            y_sheet_origin,
        )
        if bounds is not None:
            _create_despiece_masking_canvas(doc, view, *bounds)

        _draw_despiece_row_items(
            doc,
            view,
            row_items,
            tnt,
            line_style_id,
            marker_line_style_id,
            marker_line_style_name,
            rebar_nominal_diameter_mm_fn,
            x_origin,
            y_sheet_origin,
        )

        if manage_transaction and t is not None:
            t.Commit()
        if manage_transaction:
            _refresh_active_view(uidoc)
        return view
    except Exception:
        if (
            manage_transaction
            and t is not None
            and t.HasStarted()
            and not t.HasEnded()
        ):
            t.RollBack()
        raise
