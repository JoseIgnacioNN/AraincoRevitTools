# -*- coding: utf-8 -*-
"""
RevitPythonShell (RPS): RebarShape del proyecto.

- DRAW_DETAIL_VIEWS = False: solo selecciona todos los RebarShape en el modelo.
- DRAW_DETAIL_VIEWS = True: dibuja todas las formas en una sola vista de dibujo (tabla en
  filas y columnas: GRID_COLS, tamaños de celda en mm). Franja inferior opcional «BARRA TIPO nn»,
  trazo + ganchos, y TextNote con letras de segmento (A,B,… / Gi,Gf). Opcionalmente se usan
  curvas de get_Geometry solo si la forma no define ganchos por API (si define ganchos,
  GetCurvesForBrowser + trazos sintéticos). Escala de vista 1:1 (DRAFTING_VIEW_SCALE = 1).
  TextNote: tipo TEXT_NOTE_TYPE_NAME (p. ej. 2.5mm Arial); si no existe, se duplica el primer
  tipo del proyecto y se intenta ajustar tamaño/fuente. Si la vista ya existe (TABLE_VIEW_NAME),
  borra líneas y notas de texto y regenera. Al final deja los RebarShape seleccionados.

  Nota API / catálogo: la base recomendada es GetCurvesForBrowser + definición por segmentos
  y ganchos (GetDefaultHookAngle/Orientation); conviene inspeccionar cada RebarShape con
  RevitLookup o RevitDBExplorer antes de afinar heurísticas. The Building Coder trata
  «Rebar Shape» y ubicación de ganchos. Otra vía distinta es rellenar el parámetro de tipo
  Image con bitmaps exportados; este script dibuja en vista de dibujo, no asigna imágenes.

  Nota: se usa ViewDrafting (no ViewDetail), porque ViewDetail no existe en todas las
  versiones de la API / IronPython; el resultado son líneas de detalle igual que en una leyenda.

Ejecutar: File > Run script (no pegar línea a línea).
"""

import clr
import math

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from System.Collections.Generic import List
from Autodesk.Revit.DB import (
    Arc,
    Curve,
    CurveElement,
    DetailCurve,
    ElementId,
    FilteredElementCollector,
    GeometryInstance,
    HorizontalTextAlignment,
    Line,
    Mesh,
    Options,
    PolyLine,
    Solid,
    StorageType,
    TextNote,
    TextNoteOptions,
    TextNoteType,
    Transaction,
    VerticalTextAlignment,
    View,
    ViewDetailLevel,
    ViewDrafting,
    ViewFamily,
    ViewFamilyType,
    XYZ,
)
from Autodesk.Revit.DB.Structure import RebarHookOrientation, RebarShape
from Autodesk.Revit.UI import TaskDialog

# --- Comportamiento ---
DRAW_DETAIL_VIEWS = True

# Tamaño aproximado del gancho respecto al span del bbox de la directriz (GetCurvesForBrowser).
HOOK_LENGTH_RATIO = 0.28
# Escala de la vista de dibujo (1 = 1:1).
DRAFTING_VIEW_SCALE = 1
# Una sola vista de dibujo con este nombre (se reutiliza al volver a ejecutar).
TABLE_VIEW_NAME = u"RS_TABLA_FORMAS"
# Grilla tipo tabla.
GRID_COLS = 4
CELL_W_MM = 70.0
CELL_H_MM = 55.0
MARGIN_MM = 3.0
# Factor respecto al máximo que cabe en la celda (0.85 ≈ 15 % más pequeño, más aire para título y etiquetas).
CELL_GEOMETRY_SCALE_FACTOR = 0.85
# Intentar curvas desde element.get_Geometry(Options) (suele incluir ganchos); si no convence, se usa GetCurvesForBrowser.
USE_REBAR_SHAPE_ELEMENT_GEOMETRY = True
# Contorno de cada celda (líneas de detalle).
DRAW_CELL_BORDERS = True
# Letra por tramo de la directriz (A,B,… o nombre de parámetro de la definición).
DRAW_SEGMENT_LABELS = True
# Desplazamiento base del texto respecto al punto medio del tramo (mm).
SEGMENT_LABEL_OFFSET_MM = 6.0
# Extra para etiquetas de gancho (mm), suele apiñarse más.
HOOK_LABEL_EXTRA_OFFSET_MM = 2.0
# Posición a lo largo del gancho (0=anclaje, 1=punta). Gi/Gf distintos evitan solapar dos TextNote en un solo trazo.
HOOK_LABEL_T_PARAM = 0.72
HOOK_LABEL_T_PARAM_GI = 0.40
HOOK_LABEL_T_PARAM_GF = 0.82
# Si Gi y Gf quedan muy cerca (estribo), separación lateral en el plano de la forma (mm).
HOOK_PAIR_SPLIT_MM = 5.0
HOOK_PAIR_PROXIMITY_MM = 10.0
# Evitar solapes entre letras en la misma celda.
LABEL_MIN_SEPARATION_MM = 3.5
LABEL_NUDGE_STEP_MM = 2.5
# Mantener TextNotes de tramo fuera de la franja del título y del borde de celda (coord. locales).
LABEL_CLEARANCE_ABOVE_CAPTION_MM = 5.0
LABEL_CLEARANCE_TOP_CELL_MM = 2.5
LABEL_CLEARANCE_SIDE_CELL_MM = 2.0
# Vértice cerrado (Gi+Gf en el mismo punto): dos trazos paralelos al tramo dominante (p. ej. sobre la horizontal).
CLOSED_HOOK_MERGE_ATTACH_FRAC_OF_SPAN = 0.035
CLOSED_HOOK_MERGE_ATTACH_MIN_MM = 2.0
CLOSED_HOOK_PARALLEL_SEP_FRAC = 0.012
CLOSED_HOOK_PARALLEL_SEP_MIN_MM = 0.7
# Si GetDefaultHookAngle es 0, trazo corto en el eje de la barra (Gi/Gf en catálogos tipo 09/13/16).
DRAW_SCHEMATIC_AXIAL_HOOK_WHEN_ANGLE_ZERO = True
SCHEMATIC_AXIAL_HOOK_LEN_FRAC = 0.11
# Tramos casi horizontales / verticales: refuerzo del offset (mm).
HORIZONTAL_SEGMENT_EXTRA_OFFSET_MM = 1.5
VERTICAL_SEGMENT_EXTRA_OFFSET_MM = 1.0
# Nombre del tipo de texto anotación (debe existir o se duplica desde el primero del proyecto).
TEXT_NOTE_TYPE_NAME = u"2.5mm Arial"
# Texto en el segmento recto del gancho (índices API 0 = inicio polilínea, 1 = final).
HOOK_SEGMENT_LABEL_0 = u"Gi"
HOOK_SEGMENT_LABEL_1 = u"Gf"
# --- Franja de título por celda ---
DRAW_CELL_CAPTION = True
CELL_CAPTION_RATIO = 0.18
CAPTION_GAP_MM = 2.5
# Dos líneas horizontales en la franja; el título queda centrado entre ellas (fracciones de h_cap).
CAPTION_LOWER_LINE_FRAC = 0.12
CAPTION_UPPER_LINE_FRAC = 0.88
# Posición vertical del título en la franja: 0.5 = mitad; >0.5 sube el texto (entre las dos líneas, menos pegado al borde inferior).
CAPTION_TEXT_CENTER_FRAC = 0.56
# Ajuste fino vertical del título (mm, + = hacia arriba en coord. locales de la tabla).
CAPTION_TEXT_VERTICAL_BIAS_MM = 0.4
CAPTION_TEXT_TEMPLATE = u"BARRA TIPO {:02d}"


def _get_doc():
    try:
        return doc
    except NameError:
        return __revit__.ActiveUIDocument.Document


def _get_uidoc():
    try:
        return __revit__.ActiveUIDocument
    except Exception:
        return None


def _collect_rebar_shapes(document):
    shapes = []
    for rs in FilteredElementCollector(document).OfClass(RebarShape):
        if rs is None:
            continue
        try:
            eid = rs.Id
            if eid is None or eid == ElementId.InvalidElementId:
                continue
            if not hasattr(rs, "GetCurvesForBrowser"):
                continue
            shapes.append(rs)
        except Exception:
            pass
    try:
        shapes.sort(key=lambda s: (s.Name or u"", int(s.Id.IntegerValue)))
    except Exception:
        pass
    return shapes


def _first_drafting_view_family_type(document):
    for vft in FilteredElementCollector(document).OfClass(ViewFamilyType):
        try:
            if vft and vft.ViewFamily == ViewFamily.Drafting:
                return vft
        except Exception:
            pass
    return None


def _mm_to_ft(mm):
    return float(mm) / 304.8


def _find_drafting_view_by_name(document, view_name):
    for v in FilteredElementCollector(document).OfClass(ViewDrafting):
        try:
            if v and v.Name == view_name:
                return v
        except Exception:
            pass
    return None


def _create_or_get_table_view(document, vft, view_name):
    v = _find_drafting_view_by_name(document, view_name)
    if v:
        return v
    for v in FilteredElementCollector(document).OfClass(View):
        try:
            same = v and v.Name == view_name
        except Exception:
            continue
        if same and not isinstance(v, ViewDrafting):
            raise Exception(
                u'Ya existe una vista llamada "{}" que no es de dibujo.'.format(view_name)
            )
    v = ViewDrafting.Create(document, vft.Id)
    v.Name = view_name
    return v


def _clear_detail_curves_in_view(document, drafting_view):
    """
    OfClass(DetailCurve) falla en algunas versiones/API: el colector exige CurveElement
    y luego se filtra por DetailCurve.
    """
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
    ok = False
    try:
        for el in FilteredElementCollector(document, view_id):
            try:
                if isinstance(el, TextNote):
                    document.Delete(el.Id)
                    ok = True
            except Exception:
                pass
    except Exception:
        pass
    if ok:
        return
    try:
        for tn in FilteredElementCollector(document).OfClass(TextNote):
            try:
                if tn and tn.OwnerViewId == view_id:
                    document.Delete(tn.Id)
            except Exception:
                pass
    except Exception:
        pass


def _clear_table_view_content(document, drafting_view):
    _clear_detail_curves_in_view(document, drafting_view)
    _clear_text_notes_in_view(document, drafting_view)


def _first_text_note_type(document):
    for tnt in FilteredElementCollector(document).OfClass(TextNoteType):
        if tnt:
            return tnt
    return None


def _norm_upper(s):
    try:
        return (s or u"").strip().upper()
    except Exception:
        return u""


def _find_text_note_type_named(document, exact_name):
    en = (exact_name or u"").strip()
    if not en:
        return None
    for tnt in FilteredElementCollector(document).OfClass(TextNoteType):
        try:
            if tnt and (tnt.Name or u"").strip() == en:
                return tnt
        except Exception:
            pass
    return None


def _find_text_note_type_25_arial_loose(document):
    for tnt in FilteredElementCollector(document).OfClass(TextNoteType):
        try:
            nu = _norm_upper(tnt.Name)
            if u"2.5" in nu and u"ARIAL" in nu:
                return tnt
        except Exception:
            pass
    return None


def _apply_text_type_size_and_font(text_note_type, size_mm, font_name):
    if text_note_type is None:
        return
    size_ft = float(size_mm) / 304.8
    try:
        for p in text_note_type.Parameters:
            if p is None or p.IsReadOnly:
                continue
            try:
                dn = _norm_upper(p.Definition.Name)
            except Exception:
                continue
            try:
                if p.StorageType == StorageType.Double:
                    if any(k in dn for k in (u"SIZE", u"TAMAÑO", u"HEIGHT", u"ALTO", u"TEXT")):
                        p.Set(size_ft)
                elif p.StorageType == StorageType.String:
                    if any(k in dn for k in (u"FONT", u"FUENTE", u"TYPEFACE", u"TIPO DE LETRA")):
                        p.Set(str(font_name))
            except Exception:
                pass
    except Exception:
        pass


def _ensure_text_note_type_table(document):
    """
    Usa un tipo llamado TEXT_NOTE_TYPE_NAME o que contenga 2.5 + Arial; si no existe,
    duplica el primer TextNoteType del proyecto, lo renombra y ajusta tamaño/fuente.
    Debe ejecutarse dentro de una Transaction.
    """
    t = _find_text_note_type_named(document, TEXT_NOTE_TYPE_NAME)
    if t:
        return t
    t = _find_text_note_type_25_arial_loose(document)
    if t:
        return t
    base = _first_text_note_type(document)
    if base is None:
        return None
    try:
        new_id = base.Duplicate(TEXT_NOTE_TYPE_NAME)
        if new_id is None or new_id == ElementId.InvalidElementId:
            return base
        nt = document.GetElement(new_id)
        if nt is None:
            return base
        _apply_text_type_size_and_font(nt, 2.5, u"Arial")
        return nt
    except Exception:
        return base


def _outward_perp(px, py, mx, my, bb_cx, bb_cy):
    vx = float(mx) - float(bb_cx)
    vy = float(my) - float(bb_cy)
    if px * vx + py * vy < 0:
        return -px, -py
    return px, py


def _label_xy_with_nudge(
    placed,
    cell_cx,
    cell_cy,
    bb_cx,
    bb_cy,
    scale,
    mx,
    my,
    px0,
    py0,
    base_off_ft,
    y_min_local=None,
    y_max_local=None,
    x_min_local=None,
    x_max_local=None,
):
    px, py = _outward_perp(px0, py0, mx, my, bb_cx, bb_cy)
    step = _mm_to_ft(LABEL_NUDGE_STEP_MM)
    min_sep = _mm_to_ft(LABEL_MIN_SEPARATION_MM)
    vx = vy = None
    for k in range(10):
        off = base_off_ft + k * step
        vx = cell_cx + (float(mx) - float(bb_cx)) * float(scale) + px * off
        vy = cell_cy + (float(my) - float(bb_cy)) * float(scale) + py * off
        if y_min_local is not None:
            vy = max(vy, float(y_min_local))
        if y_max_local is not None:
            vy = min(vy, float(y_max_local))
        if x_min_local is not None:
            vx = max(vx, float(x_min_local))
        if x_max_local is not None:
            vx = min(vx, float(x_max_local))
        clash = False
        for ox, oy in placed:
            if math.hypot(vx - ox, vy - oy) < min_sep:
                clash = True
                break
        if not clash:
            placed.append((vx, vy))
            return vx, vy
    if vx is not None:
        if y_min_local is not None:
            vy = max(vy, float(y_min_local))
        if y_max_local is not None:
            vy = min(vy, float(y_max_local))
        if x_min_local is not None:
            vx = max(vx, float(x_min_local))
        if x_max_local is not None:
            vx = min(vx, float(x_max_local))
        placed.append((vx, vy))
        return vx, vy
    return None, None


def _apply_detail_line_style(detail_curve, style_id):
    if detail_curve is None or style_id is None:
        return
    try:
        if style_id == ElementId.InvalidElementId:
            return
        detail_curve.LineStyleId = style_id
    except Exception:
        pass


def _draw_detail_line(document, view, p1, p2, line_style_id=None):
    line = Line.CreateBound(p1, p2)
    if line is None:
        return None
    dc = document.Create.NewDetailCurve(view, line)
    _apply_detail_line_style(dc, line_style_id)
    return dc


def _create_text_note_at_local_xy(
    document,
    view,
    lx,
    ly,
    txt,
    text_note_type,
    horizontal_center=True,
    vertical_middle=False,
):
    if not txt or text_note_type is None:
        return
    origin = _map_local_to_view_plane(view, lx, ly)
    try:
        opts = TextNoteOptions(text_note_type.Id)
        if horizontal_center:
            try:
                opts.HorizontalAlignment = HorizontalTextAlignment.Center
            except Exception:
                pass
        if vertical_middle:
            try:
                opts.VerticalAlignment = VerticalTextAlignment.Middle
            except Exception:
                pass
        TextNote.Create(document, view.Id, origin, txt, opts)
    except Exception:
        try:
            opts = TextNoteOptions()
            opts.TypeId = text_note_type.Id
            if horizontal_center:
                try:
                    opts.HorizontalAlignment = HorizontalTextAlignment.Center
                except Exception:
                    pass
            if vertical_middle:
                try:
                    opts.VerticalAlignment = VerticalTextAlignment.Middle
                except Exception:
                    pass
            TextNote.Create(document, view.Id, origin, txt, opts)
        except Exception:
            pass


def _bbox_from_curves_local(curves):
    min_x = 1e99
    min_y = 1e99
    max_x = -1e99
    max_y = -1e99
    any_pts = False
    for c in curves:
        if c is None:
            continue
        try:
            p0 = c.GetEndParameter(0)
            p1 = c.GetEndParameter(1)
            dom = p1 - p0
            if abs(dom) < 1e-12:
                dom = 0.0
            for sp in (p0, p0 + 0.25 * dom, p0 + 0.5 * dom, p0 + 0.75 * dom, p1):
                pt = c.Evaluate(sp, False)
                if pt is None:
                    continue
                any_pts = True
                min_x = min(min_x, float(pt.X))
                min_y = min(min_y, float(pt.Y))
                max_x = max(max_x, float(pt.X))
                max_y = max(max_y, float(pt.Y))
        except Exception:
            try:
                pt_a = c.GetEndPoint(0)
                pt_b = c.GetEndPoint(1)
                any_pts = True
                min_x = min(min_x, float(pt_a.X), float(pt_b.X))
                min_y = min(min_y, float(pt_a.Y), float(pt_b.Y))
                max_x = max(max_x, float(pt_a.X), float(pt_b.X))
                max_y = max(max_y, float(pt_a.Y), float(pt_b.Y))
            except Exception:
                pass
    if not any_pts:
        return None
    return (min_x, max_x, min_y, max_y)


def _map_local_to_view_plane(view, x_local, y_local):
    right = view.RightDirection
    up = view.UpDirection
    o = view.Origin
    return XYZ(
        o.X + right.X * x_local + up.X * y_local,
        o.Y + right.Y * x_local + up.Y * y_local,
        o.Z + right.Z * x_local + up.Z * y_local,
    )


def _draw_browser_curve(document, view, curve, center_x, center_y, bb_cx, bb_cy, scale):
    if curve is None:
        return

    def map_pt(pt):
        x = center_x + (float(pt.X) - bb_cx) * scale
        y = center_y + (float(pt.Y) - bb_cy) * scale
        return _map_local_to_view_plane(view, x, y)

    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
    except Exception:
        p0 = p1 = None

    if isinstance(curve, Arc) and p0 is not None and p1 is not None:
        try:
            pm = curve.Evaluate(0.5, True)
            arc = Arc.Create(map_pt(p0), map_pt(p1), map_pt(pm))
            if arc:
                document.Create.NewDetailCurve(view, arc)
            return
        except Exception:
            pass

    if isinstance(curve, Line) and p0 is not None and p1 is not None:
        try:
            line = Line.CreateBound(map_pt(p0), map_pt(p1))
            if line:
                document.Create.NewDetailCurve(view, line)
            return
        except Exception:
            pass

    pts = []
    try:
        pts = list(curve.Tessellate() or [])
    except Exception:
        pts = []
    if len(pts) < 2 and p0 is not None and p1 is not None:
        pts = [p0, p1]
    for i in range(len(pts) - 1):
        try:
            line = Line.CreateBound(map_pt(pts[i]), map_pt(pts[i + 1]))
            if line:
                document.Create.NewDetailCurve(view, line)
        except Exception:
            pass


def _curves_from_browser(shape):
    """Directriz del navegador de formas (sin ganchos)."""
    try:
        raw = shape.GetCurvesForBrowser()
    except Exception:
        return []
    if raw is None:
        return []
    out = []
    try:
        for c in raw:
            if c:
                out.append(c)
    except Exception:
        pass
    return out


def _curve_length_safe(c):
    try:
        return abs(float(c.Length))
    except Exception:
        try:
            return float(c.GetEndPoint(0).DistanceTo(c.GetEndPoint(1)))
        except Exception:
            return 0.0


def _total_curves_length(curves):
    s = 0.0
    for c in curves or []:
        s += _curve_length_safe(c)
    return s


def _walk_geometry_element_split(ge, acc_wire, acc_solid):
    if ge is None:
        return
    try:
        for go in ge:
            _walk_geometry_object_split(go, acc_wire, acc_solid)
    except Exception:
        pass


def _walk_geometry_object_split(go, acc_wire, acc_solid):
    if go is None:
        return
    try:
        if isinstance(go, Curve):
            acc_wire.append(go)
            return
    except Exception:
        pass
    try:
        if isinstance(go, PolyLine):
            pts = list(go.GetCoordinates())
            for i in range(len(pts) - 1):
                try:
                    ln = Line.CreateBound(pts[i], pts[i + 1])
                    if ln:
                        acc_wire.append(ln)
                except Exception:
                    pass
            return
    except Exception:
        pass
    try:
        if isinstance(go, Solid):
            for edge in go.Edges:
                try:
                    ec = edge.AsCurve()
                    if ec:
                        acc_solid.append(ec)
                except Exception:
                    pass
            return
    except Exception:
        pass
    try:
        if isinstance(go, GeometryInstance):
            sub = go.GetInstanceGeometry(go.Transform)
            if sub is not None:
                _walk_geometry_element_split(sub, acc_wire, acc_solid)
            return
    except Exception:
        pass
    try:
        if isinstance(go, Mesh):
            return
    except Exception:
        pass


def _curves_from_element_geometry(shape):
    """
    Curvas desde la representación geométrica del elemento (p. ej. símbolo con ganchos).
    Prefiere Curves / PolyLine; si no hay, aristas de Solid (p. ej. contorno).
    """
    try:
        opt = Options()
        opt.ComputeReferences = False
        opt.DetailLevel = ViewDetailLevel.Fine
        try:
            opt.IncludeNonVisibleObjects = True
        except Exception:
            pass
    except Exception:
        return []
    ge = None
    try:
        ge = shape.GetGeometry(opt)
    except Exception:
        try:
            ge = shape.get_Geometry(opt)
        except Exception:
            try:
                ge = shape.Geometry[opt]
            except Exception:
                return []
    if ge is None:
        return []
    acc_wire = []
    acc_solid = []
    _walk_geometry_element_split(ge, acc_wire, acc_solid)
    return acc_wire if acc_wire else acc_solid


def _filter_short_geometry_curves(curves):
    if not curves:
        return []
    bb = _bbox_from_curves_local(curves)
    if bb is None:
        return list(curves)
    x_min, x_max, y_min, y_max = bb
    span = max(x_max - x_min, y_max - y_min, 1e-9)
    min_len = max(1e-5, span * 0.0015)
    out = [c for c in curves if _curve_length_safe(c) >= min_len]
    return out if out else list(curves)


def _prefer_element_geometry_over_browser(geo, browser):
    """Evita usar mallas sólidas ruidosas; pide longitud total razonable frente al browser."""
    if not geo:
        return False
    ng = len(geo)
    if ng > 500:
        return False
    nb = len(browser) if browser else 0
    if ng > max(120, nb * 20):
        return False
    Lg = _total_curves_length(geo)
    if Lg < 1e-6:
        return False
    if not browser:
        return True
    Lb = _total_curves_length(browser)
    if Lb < 1e-6:
        return True
    if Lg < Lb * 0.65:
        return False
    return True


def _rebar_shape_has_default_hooks(shape):
    """True si la familia define gancho en inicio o fin (GetDefaultHookAngle > 0)."""
    try:
        a0 = int(shape.GetDefaultHookAngle(0))
    except Exception:
        a0 = 0
    try:
        a1 = int(shape.GetDefaultHookAngle(1))
    except Exception:
        a1 = 0
    return a0 > 0 or a1 > 0


def _fallback_segment_letter(i):
    if i < 26:
        return chr(65 + i)
    return str(i + 1)


def _letter_from_dimension_name(name):
    if not name:
        return None
    s = name.strip()
    if len(s) == 1 and s.isalpha():
        return s.upper()
    parts = s.replace(u",", u" ").split()
    for p in reversed(parts):
        if len(p) == 1 and p.isalpha():
            return p.upper()
    return None


def _parameter_name_for_element_id(shape, eid):
    try:
        for p in shape.Parameters:
            try:
                if p is None or p.Id != eid:
                    continue
                d = p.Definition
                if d is None:
                    continue
                return d.Name
            except Exception:
                continue
    except Exception:
        pass
    return None


def _letters_for_browser_segments(shape, browser_curves):
    n = len(browser_curves)
    letters = [_fallback_segment_letter(i) for i in range(n)]
    if n == 0:
        return letters
    try:
        from Autodesk.Revit.DB.Structure import RebarShapeDefinitionBySegments

        rsd = shape.GetRebarShapeDefinition()
        if not isinstance(rsd, RebarShapeDefinitionBySegments):
            return letters
        nseg = int(rsd.NumberOfSegments)
        for i in range(min(nseg, n)):
            try:
                seg = rsd.GetSegment(i)
                cons = seg.GetConstraints()
                if cons is None:
                    continue
                for c in cons:
                    try:
                        pid = c.GetParamId()
                        if pid is None or pid == ElementId.InvalidElementId:
                            continue
                        pname = _parameter_name_for_element_id(shape, pid)
                        ch = _letter_from_dimension_name(pname)
                        if ch:
                            letters[i] = ch
                            break
                    except Exception:
                        pass
            except Exception:
                pass
    except Exception:
        pass
    return letters


def _curve_midpoint_local_xy(curve):
    try:
        p = curve.Evaluate(0.5, True)
        return float(p.X), float(p.Y)
    except Exception:
        try:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            return 0.5 * (float(p0.X) + float(p1.X)), 0.5 * (float(p0.Y) + float(p1.Y))
        except Exception:
            return None


def _perp_offset_unit_xy(curve):
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        dx = float(p1.X - p0.X)
        dy = float(p1.Y - p0.Y)
        L = math.hypot(dx, dy)
        if L < 1e-12:
            return 0.0, 1.0
        return -dy / L, dx / L
    except Exception:
        return 0.0, 1.0


def _hook_label_point_local_xy(curve, t_param=None):
    """Punto sobre el tramo del gancho; t_param 0..1 desde anclaje hacia la punta."""
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        if t_param is None:
            t = float(HOOK_LABEL_T_PARAM)
        else:
            t = float(t_param)
        return (
            float(p0.X) + t * (float(p1.X) - float(p0.X)),
            float(p0.Y) + t * (float(p1.Y) - float(p0.Y)),
        )
    except Exception:
        return None


def _arc_label_point_and_perp(curve, bb_cx, bb_cy):
    """Punto medio del arco y dirección radial hacia el exterior respecto al centro de la forma."""
    try:
        if not isinstance(curve, Arc):
            return None
        mid = curve.Evaluate(0.5, True)
        mx, my = float(mid.X), float(mid.Y)
        ctr = curve.Center
        cx, cy = float(ctr.X), float(ctr.Y)
        rx, ry = mx - cx, my - cy
        lr = math.hypot(rx, ry)
        if lr < 1e-12:
            return None
        ox, oy = rx / lr, ry / lr
        vx, vy = mx - float(bb_cx), my - float(bb_cy)
        if ox * vx + oy * vy < 0:
            ox, oy = -ox, -oy
        return mx, my, ox, oy
    except Exception:
        return None


def _segment_anchor_and_perp(curve, bb_cx, bb_cy):
    arc = _arc_label_point_and_perp(curve, bb_cx, bb_cy)
    if arc is not None:
        return arc
    mid = _curve_midpoint_local_xy(curve)
    if mid is None:
        return None
    mx, my = mid
    px0, py0 = _perp_offset_unit_xy(curve)
    return mx, my, px0, py0


def _line_orientation_extra_offset_mm(curve):
    """Refuerzo de separación en tramos casi horizontales o verticales."""
    if not isinstance(curve, Line):
        return 0.0
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        dx = abs(float(p1.X - p0.X))
        dy = abs(float(p1.Y - p0.Y))
        if dx + dy < 1e-12:
            return 0.0
        if dy <= 0.06 * dx:
            return float(HORIZONTAL_SEGMENT_EXTRA_OFFSET_MM)
        if dx <= 0.06 * dy:
            return float(VERTICAL_SEGMENT_EXTRA_OFFSET_MM)
    except Exception:
        pass
    return 0.0


def _place_segment_letter(
    document,
    view,
    placed,
    cell_cx,
    cell_cy,
    bb_cx,
    bb_cy,
    scale,
    mx,
    my,
    px0,
    py0,
    base_off_ft,
    txt,
    text_note_type,
    y_min_local=None,
    y_max_local=None,
    x_min_local=None,
    x_max_local=None,
):
    if not txt or text_note_type is None:
        return
    vx, vy = _label_xy_with_nudge(
        placed,
        cell_cx,
        cell_cy,
        bb_cx,
        bb_cy,
        scale,
        mx,
        my,
        px0,
        py0,
        base_off_ft,
        y_min_local,
        y_max_local,
        x_min_local,
        x_max_local,
    )
    if vx is None:
        return
    _create_text_note_at_local_xy(document, view, vx, vy, txt, text_note_type, True)


def _draw_segment_labels(
    document,
    view,
    shape,
    bb_cx,
    bb_cy,
    scale,
    cell_cx,
    cell_cy,
    text_note_type,
    hook_labeled,
    segment_curves_for_letters,
    gx_min,
    gx_max,
    gy_draw_min,
    gy_draw_max,
):
    if not DRAW_SEGMENT_LABELS or text_note_type is None:
        return
    if not segment_curves_for_letters:
        return
    browser_ref = _curves_from_browser(shape)
    if not browser_ref:
        return
    letters = _letters_for_browser_segments(shape, browser_ref)
    off_hook = _mm_to_ft(SEGMENT_LABEL_OFFSET_MM + HOOK_LABEL_EXTRA_OFFSET_MM)
    placed = []

    pad_y_lo = _mm_to_ft(float(LABEL_CLEARANCE_ABOVE_CAPTION_MM))
    pad_y_hi = _mm_to_ft(float(LABEL_CLEARANCE_TOP_CELL_MM))
    pad_x = _mm_to_ft(float(LABEL_CLEARANCE_SIDE_CELL_MM))
    y_lo = float(gy_draw_min) + pad_y_lo
    y_hi = float(gy_draw_max) - pad_y_hi
    x_lo = float(gx_min) + pad_x
    x_hi = float(gx_max) - pad_x
    if y_lo > y_hi:
        y_lo, y_hi = float(gy_draw_min), float(gy_draw_max)
    if x_lo > x_hi:
        x_lo, x_hi = float(gx_min), float(gx_max)

    for i, curve in enumerate(segment_curves_for_letters):
        if i >= len(letters):
            break
        ap = _segment_anchor_and_perp(curve, bb_cx, bb_cy)
        if ap is None:
            continue
        mx, my, px0, py0 = ap
        extra_mm = _line_orientation_extra_offset_mm(curve)
        off_use = _mm_to_ft(SEGMENT_LABEL_OFFSET_MM + extra_mm)
        _place_segment_letter(
            document,
            view,
            placed,
            cell_cx,
            cell_cy,
            bb_cx,
            bb_cy,
            scale,
            mx,
            my,
            px0,
            py0,
            off_use,
            letters[i],
            text_note_type,
            y_lo,
            y_hi,
            x_lo,
            x_hi,
        )

    if hook_labeled:
        hook_items = []
        for curve, htxt in hook_labeled:
            if curve is None or not htxt:
                continue
            try:
                if htxt == HOOK_SEGMENT_LABEL_1:
                    t_use = float(HOOK_LABEL_T_PARAM_GF)
                elif htxt == HOOK_SEGMENT_LABEL_0:
                    t_use = float(HOOK_LABEL_T_PARAM_GI)
                else:
                    t_use = float(HOOK_LABEL_T_PARAM)
            except Exception:
                t_use = float(HOOK_LABEL_T_PARAM)
            hp = _hook_label_point_local_xy(curve, t_use)
            if hp is None:
                continue
            mx, my = hp
            hook_items.append((curve, htxt, mx, my))

        if len(hook_items) == 2:
            c0, t0, m0x, m0y = hook_items[0]
            c1, t1, m1x, m1y = hook_items[1]
            d_cell_mm = math.hypot(m1x - m0x, m1y - m0y) * float(scale) * 304.8
            if d_cell_mm < float(HOOK_PAIR_PROXIMITY_MM):
                ddx, ddy = m1x - m0x, m1y - m0y
                L12 = math.hypot(ddx, ddy)
                if L12 < 1e-12:
                    px, py = 1.0, 0.0
                else:
                    px, py = -ddy / L12, ddx / L12
                half = _mm_to_ft(0.5 * float(HOOK_PAIR_SPLIT_MM)) / max(
                    float(scale), 1e-12
                )
                hook_items[0] = (c0, t0, m0x - px * half, m0y - py * half)
                hook_items[1] = (c1, t1, m1x + px * half, m1y + py * half)

        for curve, htxt, mx, my in hook_items:
            px0, py0 = _perp_offset_unit_xy(curve)
            _place_segment_letter(
                document,
                view,
                placed,
                cell_cx,
                cell_cy,
                bb_cx,
                bb_cy,
                scale,
                mx,
                my,
                px0,
                py0,
                off_hook,
                htxt,
                text_note_type,
                y_lo,
                y_hi,
                x_lo,
                x_hi,
            )


def _v_sub(a, b):
    return XYZ(a.X - b.X, a.Y - b.Y, a.Z - b.Z)


def _v_add(a, b):
    return XYZ(a.X + b.X, a.Y + b.Y, a.Z + b.Z)


def _v_scale(v, s):
    return XYZ(v.X * s, v.Y * s, v.Z * s)


def _v_dot(a, b):
    return a.X * b.X + a.Y * b.Y + a.Z * b.Z


def _v_len(a):
    return math.sqrt(a.X * a.X + a.Y * a.Y + a.Z * a.Z)


def _v_unit(v):
    L = _v_len(v)
    if L < 1e-12:
        return None
    return XYZ(v.X / L, v.Y / L, v.Z / L)


def _orient_is_right(orient):
    try:
        return int(orient) == int(RebarHookOrientation.Right)
    except Exception:
        return False


def _vector_unit_xy(v):
    """Proyección unitaria al plano XY (misma base que el dibujo de la tabla)."""
    if v is None:
        return None
    x, y = float(v.X), float(v.Y)
    L = math.hypot(x, y)
    if L < 1e-12:
        return None
    return XYZ(x / L, y / L, 0.0)


def _curve_tangent_into_bar_at_start(curve):
    """Tangente unitaria en el parámetro inicial, hacia el interior de la directriz."""
    if isinstance(curve, Line):
        try:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            u = _v_unit(_v_sub(p1, p0))
            if u is not None:
                return u
        except Exception:
            pass
    try:
        u0 = curve.GetEndParameter(0)
        deriv = curve.ComputeDerivatives(u0, True)
        if deriv is not None:
            u = _v_unit(deriv.BasisX)
            if u is not None:
                return u
    except Exception:
        pass
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        return _v_unit(_v_sub(p1, p0))
    except Exception:
        return None


def _curve_tangent_into_bar_at_end(curve):
    """
    En el extremo final: tangente que entra en el acero hacia el interior de la barra
    (opuesta a la derivada en el parámetro final). Evita invertir el gancho Gf respecto a Gi.
    """
    if isinstance(curve, Line):
        try:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            return _v_unit(_v_sub(p0, p1))
        except Exception:
            pass
    try:
        u1 = curve.GetEndParameter(1)
        deriv = curve.ComputeDerivatives(u1, True)
        if deriv is not None:
            bx = deriv.BasisX
            u = _v_unit(XYZ(-bx.X, -bx.Y, -bx.Z))
            if u is not None:
                return u
    except Exception:
        pass
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        return _v_unit(_v_sub(p0, p1))
    except Exception:
        return None


def _curve_tangent_into_bar_at_start_projected_xy(curve):
    """Tangente «hacia dentro» en el inicio, proyectada al XY de la forma (coherente con el trazo)."""
    if isinstance(curve, Line):
        try:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            dx = float(p1.X - p0.X)
            dy = float(p1.Y - p0.Y)
            Lxy = math.hypot(dx, dy)
            if Lxy > 1e-12:
                return XYZ(dx / Lxy, dy / Lxy, 0.0)
        except Exception:
            pass
        return None
    return _vector_unit_xy(_curve_tangent_into_bar_at_start(curve))


def _curve_tangent_into_bar_at_end_projected_xy(curve):
    """Tangente «hacia dentro» en el extremo final, proyectada al XY de la forma."""
    if isinstance(curve, Line):
        try:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            dx = float(p0.X - p1.X)
            dy = float(p0.Y - p1.Y)
            Lxy = math.hypot(dx, dy)
            if Lxy > 1e-12:
                return XYZ(dx / Lxy, dy / Lxy, 0.0)
        except Exception:
            pass
        return None
    return _vector_unit_xy(_curve_tangent_into_bar_at_end(curve))


def _hook_curves_at_attach(attach, t_into_bar_xy, angle_deg, orient, hook_len):
    """
    Tramo recto del gancho en el plano XY de la forma (igual que map_pt al dibujar).
    Así el gancho comparte extremo con la directriz y no «flota» por componente Z.
    t_into_bar_xy: unitario con Z=0 (dirección de la barra hacia el interior en planta).
    """
    out = []
    t = _v_unit(t_into_bar_xy)
    if t is None:
        return out
    ad = int(angle_deg)
    if ad <= 0 or ad > 180:
        return out
    if hook_len < 1e-9:
        return out

    tx, ty = float(t.X), float(t.Y)
    h0 = XYZ(-tx, -ty, 0.0)
    sx, sy = -ty, tx
    ls = math.hypot(sx, sy)
    if ls < 1e-12:
        return out
    sx, sy = sx / ls, sy / ls
    if _orient_is_right(orient):
        sx, sy = -sx, -sy
    side = XYZ(sx, sy, 0.0)

    theta = math.radians(float(ad))
    ux = h0.X * math.cos(theta) + side.X * math.sin(theta)
    uy = h0.Y * math.cos(theta) + side.Y * math.sin(theta)
    u = _v_unit(XYZ(ux, uy, 0.0))
    if u is None:
        return out

    tip = _v_add(attach, _v_scale(u, hook_len))
    ln = Line.CreateBound(attach, tip)
    if ln:
        out.append(ln)
    return out


def _synthetic_hook_curves_labeled(shape, curves, span):
    """
    Lista de (curva Line del gancho, etiqueta). Misma geometría que antes; las etiquetas
    son HOOK_SEGMENT_LABEL_0 / _1 (extremos API 0 y 1).
    """
    pairs = []
    if not curves:
        return pairs

    hook_len = max(float(span) * HOOK_LENGTH_RATIO, 1e-4)
    schematic_len = max(
        float(span) * float(SCHEMATIC_AXIAL_HOOK_LEN_FRAC),
        _mm_to_ft(1.2),
    )

    try:
        a0 = int(shape.GetDefaultHookAngle(0))
    except Exception:
        a0 = 0
    if a0 > 0:
        try:
            o0 = shape.GetDefaultHookOrientation(0)
        except Exception:
            o0 = RebarHookOrientation.Left
        try:
            c0 = curves[0]
            p_a = c0.GetEndPoint(0)
            t_raw = _curve_tangent_into_bar_at_start_projected_xy(c0)
            if t_raw is not None:
                for ln in _hook_curves_at_attach(p_a, t_raw, a0, o0, hook_len):
                    pairs.append((ln, HOOK_SEGMENT_LABEL_0))
        except Exception:
            pass
    elif DRAW_SCHEMATIC_AXIAL_HOOK_WHEN_ANGLE_ZERO:
        try:
            c0 = curves[0]
            p_a = c0.GetEndPoint(0)
            t_in = _curve_tangent_into_bar_at_start_projected_xy(c0)
            if t_in is not None:
                ux = -float(t_in.X)
                uy = -float(t_in.Y)
                ul = math.hypot(ux, uy)
                if ul > 1e-12:
                    ux, uy = ux / ul, uy / ul
                    tip = XYZ(
                        float(p_a.X) + ux * float(schematic_len),
                        float(p_a.Y) + uy * float(schematic_len),
                        float(p_a.Z),
                    )
                    ln_s = Line.CreateBound(p_a, tip)
                    if ln_s:
                        pairs.append((ln_s, HOOK_SEGMENT_LABEL_0))
        except Exception:
            pass

    try:
        a1 = int(shape.GetDefaultHookAngle(1))
    except Exception:
        a1 = 0
    if a1 > 0:
        try:
            o1 = shape.GetDefaultHookOrientation(1)
        except Exception:
            o1 = RebarHookOrientation.Left
        try:
            cL = curves[-1]
            p_a = cL.GetEndPoint(1)
            t_raw = _curve_tangent_into_bar_at_end_projected_xy(cL)
            if t_raw is not None:
                for ln in _hook_curves_at_attach(p_a, t_raw, a1, o1, hook_len):
                    pairs.append((ln, HOOK_SEGMENT_LABEL_1))
        except Exception:
            pass
    elif DRAW_SCHEMATIC_AXIAL_HOOK_WHEN_ANGLE_ZERO:
        try:
            cL = curves[-1]
            p_a = cL.GetEndPoint(1)
            t_in = _curve_tangent_into_bar_at_end_projected_xy(cL)
            if t_in is not None:
                ux = -float(t_in.X)
                uy = -float(t_in.Y)
                ul = math.hypot(ux, uy)
                if ul > 1e-12:
                    ux, uy = ux / ul, uy / ul
                    tip = XYZ(
                        float(p_a.X) + ux * float(schematic_len),
                        float(p_a.Y) + uy * float(schematic_len),
                        float(p_a.Z),
                    )
                    ln_s = Line.CreateBound(p_a, tip)
                    if ln_s:
                        pairs.append((ln_s, HOOK_SEGMENT_LABEL_1))
        except Exception:
            pass

    return _merge_coincident_hook_pairs(pairs, span, hook_len, curves)


def _dominant_segment_direction_at_vertex_xy(curves, attach, span):
    """
    En el vértice de cierre, dirección XY unitaria del tramo incidente más largo
    (suele coincidir con el lado largo del estribo, p. ej. horizontal en catálogos).
    """
    if not curves:
        return None
    tol = max(float(span) * 0.025, _mm_to_ft(4.0))
    ax, ay = float(attach.X), float(attach.Y)
    best_l = -1.0
    best_px, best_py = 1.0, 0.0
    for c in curves:
        try:
            p0 = c.GetEndPoint(0)
            p1 = c.GetEndPoint(1)
        except Exception:
            continue
        d0 = math.hypot(float(p0.X) - ax, float(p0.Y) - ay)
        d1 = math.hypot(float(p1.X) - ax, float(p1.Y) - ay)
        if d0 < tol and d1 >= tol:
            vx = float(p1.X - p0.X)
            vy = float(p1.Y - p0.Y)
        elif d1 < tol and d0 >= tol:
            vx = float(p0.X - p1.X)
            vy = float(p0.Y - p1.Y)
        else:
            continue
        lg = math.hypot(vx, vy)
        if lg < 1e-12:
            continue
        if lg > best_l:
            best_l = lg
            best_px, best_py = vx / lg, vy / lg
    if best_l < 0:
        return None
    return XYZ(best_px, best_py, 0.0)


def _merge_coincident_hook_pairs(pairs, span, hook_len, curves):
    """
    Estribo cerrado: Gi y Gf en el mismo vértice. En catálogos suelen verse dos trazos
    paralelos al lado dominante (no una bisectriz en «V» ni un solo segmento).
    Si no hay tramo dominante claro, se usa la dirección media de los dos ganchos.
    """
    if len(pairs) != 2:
        return pairs
    ln0, lab0 = pairs[0]
    ln1, lab1 = pairs[1]
    if ln0 is None or ln1 is None:
        return pairs
    try:
        p0a = ln0.GetEndPoint(0)
        p1a = ln1.GetEndPoint(0)
        dxy = math.hypot(float(p0a.X - p1a.X), float(p0a.Y - p1a.Y))
    except Exception:
        return pairs
    thresh = max(
        float(span) * float(CLOSED_HOOK_MERGE_ATTACH_FRAC_OF_SPAN),
        _mm_to_ft(float(CLOSED_HOOK_MERGE_ATTACH_MIN_MM)),
    )
    if dxy > thresh:
        return pairs

    try:
        q0 = ln0.GetEndPoint(1)
        q1 = ln1.GetEndPoint(1)
        v0x = float(q0.X - p0a.X)
        v0y = float(q0.Y - p0a.Y)
        v1x = float(q1.X - p1a.X)
        v1y = float(q1.Y - p1a.Y)
        l0 = math.hypot(v0x, v0y)
        l1 = math.hypot(v1x, v1y)
        if l0 > 1e-12 and l1 > 1e-12:
            v0x, v0y = v0x / l0, v0y / l0
            v1x, v1y = v1x / l1, v1y / l1
            if v0x * v1x + v0y * v1y < -0.75:
                return pairs
    except Exception:
        pass

    d_dom = _dominant_segment_direction_at_vertex_xy(curves, p0a, span)
    if d_dom is not None:
        try:
            px = float(d_dom.X)
            py = float(d_dom.Y)
            pl = math.hypot(px, py)
            if pl > 1e-12:
                px, py = px / pl, py / pl
            az = float(p0a.Z)
            ax, ay = float(p0a.X), float(p0a.Y)
            off = max(
                float(span) * float(CLOSED_HOOK_PARALLEL_SEP_FRAC),
                _mm_to_ft(float(CLOSED_HOOK_PARALLEL_SEP_MIN_MM)),
            )
            perp_x = -py
            perp_y = px
            half = 0.5 * float(off)
            x1 = ax + perp_x * (-half)
            y1 = ay + perp_y * (-half)
            x2 = x1 + px * float(hook_len)
            y2 = y1 + py * float(hook_len)
            x3 = ax + perp_x * half
            y3 = ay + perp_y * half
            x4 = x3 + px * float(hook_len)
            y4 = y3 + py * float(hook_len)
            ln_a = Line.CreateBound(XYZ(x1, y1, az), XYZ(x2, y2, az))
            ln_b = Line.CreateBound(XYZ(x3, y3, az), XYZ(x4, y4, az))
            if ln_a and ln_b:
                if lab0 == HOOK_SEGMENT_LABEL_1 and lab1 == HOOK_SEGMENT_LABEL_0:
                    return [(ln_b, lab1), (ln_a, lab0)]
                return [(ln_a, lab0), (ln_b, lab1)]
        except Exception:
            pass

    try:
        q0 = ln0.GetEndPoint(1)
        q1 = ln1.GetEndPoint(1)
        u0x = float(q0.X - p0a.X)
        u0y = float(q0.Y - p0a.Y)
        u1x = float(q1.X - p1a.X)
        u1y = float(q1.Y - p1a.Y)
        l0 = math.hypot(u0x, u0y)
        l1 = math.hypot(u1x, u1y)
        if l0 < 1e-12 or l1 < 1e-12:
            return pairs
        u0x, u0y = u0x / l0, u0y / l0
        u1x, u1y = u1x / l1, u1y / l1
        mx = u0x + u1x
        my = u0y + u1y
        lm = math.hypot(mx, my)
        if lm < 1e-12:
            return pairs
        mx, my = mx / lm, my / lm
        attach = XYZ(float(p0a.X), float(p0a.Y), float(p0a.Z))
        tip = XYZ(
            float(p0a.X) + mx * float(hook_len),
            float(p0a.Y) + my * float(hook_len),
            float(p0a.Z),
        )
        ln_m = Line.CreateBound(attach, tip)
        if ln_m is None:
            return pairs
        if lab0 == HOOK_SEGMENT_LABEL_1 and lab1 == HOOK_SEGMENT_LABEL_0:
            return [(ln_m, lab1), (ln_m, lab0)]
        return [(ln_m, lab0), (ln_m, lab1)]
    except Exception:
        return pairs


def _unique_hook_curves_for_draw(hook_labeled):
    """Evita dibujar dos veces la misma Line cuando Gi y Gf comparten geometría."""
    out = []
    seen = set()
    for crv, _txt in hook_labeled or []:
        if crv is None:
            continue
        oid = id(crv)
        if oid in seen:
            continue
        seen.add(oid)
        out.append(crv)
    return out


def _prepare_shape_curves(shape):
    """
    Curvas para dibujar + bbox + ganchos sintéticos si la forma tiene ganchos API o no se usa
    get_Geometry. Devuelve (all_curves, bb_cx, bb_cy, bbox_w, bbox_h, hook_labeled, segment_curves_for_labels).
    """
    browser_curves = _curves_from_browser(shape)
    if not browser_curves:
        return None

    chosen = list(browser_curves)
    hook_labeled = []

    if USE_REBAR_SHAPE_ELEMENT_GEOMETRY:
        geo_raw = _curves_from_element_geometry(shape)
        geo_filt = _filter_short_geometry_curves(geo_raw) if geo_raw else []
        # La Geometry del símbolo suele ser solo el contorno; los ganchos no salen como trazos
        # separados → Gi/Gf quedarían sin línea. Con ganchos API usamos browser + sintéticos.
        use_geo = (
            bool(geo_filt)
            and _prefer_element_geometry_over_browser(geo_filt, browser_curves)
            and not _rebar_shape_has_default_hooks(shape)
        )
        if use_geo:
            chosen = list(geo_filt)
        else:
            bb0 = _bbox_from_curves_local(browser_curves)
            if bb0 is None:
                return None
            x_min, x_max, y_min, y_max = bb0
            span0 = max(x_max - x_min, y_max - y_min, 1e-9)
            hook_labeled = _synthetic_hook_curves_labeled(
                shape, browser_curves, span0
            )
    else:
        bb0 = _bbox_from_curves_local(browser_curves)
        if bb0 is None:
            return None
        x_min, x_max, y_min, y_max = bb0
        span0 = max(x_max - x_min, y_max - y_min, 1e-9)
        hook_labeled = _synthetic_hook_curves_labeled(
            shape, browser_curves, span0
        )

    hook_curves = _unique_hook_curves_for_draw(hook_labeled)
    all_curves = list(chosen) + hook_curves
    bb = _bbox_from_curves_local(all_curves)
    if bb is None:
        return None
    x_min, x_max, y_min, y_max = bb
    bw = max(x_max - x_min, 1e-9)
    bh = max(y_max - y_min, 1e-9)
    bb_cx = 0.5 * (x_min + x_max)
    bb_cy = 0.5 * (y_min + y_max)

    # Letras A,B,C… según la definición / browser; el dibujo puede ser get_Geometry (orden distinto).
    segment_curves_for_labels = list(browser_curves)

    return (
        all_curves,
        bb_cx,
        bb_cy,
        bw,
        bh,
        hook_labeled,
        segment_curves_for_labels,
    )


def _draw_shapes_table_grid(document, view, shapes):
    """
    Dibuja cada forma en la zona superior de la celda (debajo de la franja de título si
    DRAW_CELL_CAPTION). Escala uniforme al rectángulo interior; letras de segmento con TextNote.

    Orden: la primera forma de la lista (p. ej. por nombre) va en la fila superior
    (izquierda→derecha), luego la siguiente fila hacia abajo. En coordenadas locales
    de la vista de dibujo Y suele crecer hacia arriba, por eso se invierte el índice de fila.
    """
    cols = int(GRID_COLS)
    if cols < 1:
        cols = 1
    n = len(shapes)
    if n == 0:
        return 0, [], None

    need_text_type = DRAW_SEGMENT_LABELS or DRAW_CELL_CAPTION
    text_note_type = _ensure_text_note_type_table(document) if need_text_type else None
    warn_labels = None
    if need_text_type and text_note_type is None:
        warn_labels = (
            u"No hay ningún TextNoteType en el proyecto; no se pueden crear anotaciones. "
            u"Carga un tipo base o crea uno manualmente."
        )

    rows = int(math.ceil(float(n) / float(cols)))
    cell_w_ft = _mm_to_ft(CELL_W_MM)
    cell_h_ft = _mm_to_ft(CELL_H_MM)
    margin_ft = _mm_to_ft(MARGIN_MM)

    total_w = cols * cell_w_ft
    total_h = rows * cell_h_ft
    grid_min_x = -0.5 * total_w
    grid_min_y = -0.5 * total_h

    drawn = 0
    skipped = []

    for idx, shape in enumerate(shapes):
        r_logical = idx // cols
        c = idx % cols
        # r_logical=0 → fila superior; sin invertir, grid_min_y+r*cell quedaría abajo en pantalla.
        r = rows - 1 - r_logical
        x_min_cell = grid_min_x + c * cell_w_ft
        x_max_cell = x_min_cell + cell_w_ft
        y_min_cell = grid_min_y + r * cell_h_ft
        y_max_cell = y_min_cell + cell_h_ft
        cell_cx = 0.5 * (x_min_cell + x_max_cell)
        cell_cy = 0.5 * (y_min_cell + y_max_cell)

        gx_min = x_min_cell + margin_ft
        gx_max = x_max_cell - margin_ft
        if DRAW_CELL_CAPTION and float(CELL_CAPTION_RATIO) > 1e-6:
            h_cap = cell_h_ft * float(CELL_CAPTION_RATIO)
            y_sep = y_min_cell + h_cap
            y_cap_lo = y_min_cell + float(CAPTION_LOWER_LINE_FRAC) * h_cap
            y_cap_hi = y_min_cell + float(CAPTION_UPPER_LINE_FRAC) * h_cap
            cap_gap = _mm_to_ft(CAPTION_GAP_MM)
            gy_draw_min = y_sep + cap_gap
            gy_draw_max = y_max_cell - margin_ft
            s_sep1 = _map_local_to_view_plane(view, x_min_cell, y_sep)
            s_sep2 = _map_local_to_view_plane(view, x_max_cell, y_sep)
            _draw_detail_line(document, view, s_sep1, s_sep2)
            slo1 = _map_local_to_view_plane(view, x_min_cell, y_cap_lo)
            slo2 = _map_local_to_view_plane(view, x_max_cell, y_cap_lo)
            _draw_detail_line(document, view, slo1, slo2)
            shi1 = _map_local_to_view_plane(view, x_min_cell, y_cap_hi)
            shi2 = _map_local_to_view_plane(view, x_max_cell, y_cap_hi)
            _draw_detail_line(document, view, shi1, shi2)
            if text_note_type:
                cap_y = (
                    y_cap_lo
                    + float(CAPTION_TEXT_CENTER_FRAC) * (y_cap_hi - y_cap_lo)
                    + _mm_to_ft(float(CAPTION_TEXT_VERTICAL_BIAS_MM))
                )
                cap_txt = CAPTION_TEXT_TEMPLATE.format(idx + 1)
                try:
                    cap_txt = cap_txt.upper()
                except Exception:
                    pass
                _create_text_note_at_local_xy(
                    document,
                    view,
                    cell_cx,
                    cap_y,
                    cap_txt,
                    text_note_type,
                    horizontal_center=True,
                    vertical_middle=True,
                )
        else:
            gy_draw_min = y_min_cell + margin_ft
            gy_draw_max = y_max_cell - margin_ft

        inner_w = gx_max - gx_min
        inner_h = gy_draw_max - gy_draw_min
        if inner_w <= 0 or inner_h <= 0:
            try:
                skipped.append(u"{} — celda sin espacio (revisa MARGIN_MM)".format(shape.Name))
            except Exception:
                skipped.append(u"celda sin espacio")
            continue

        prep = _prepare_shape_curves(shape)
        if prep is None:
            try:
                skipped.append(u"{} — sin curvas / bbox".format(shape.Name))
            except Exception:
                skipped.append(u"sin curvas / bbox")
            continue

        all_curves, bb_cx, bb_cy, bw, bh, hook_labeled, seg_for_lbl = prep
        scale = (
            min(inner_w / bw, inner_h / bh) * float(CELL_GEOMETRY_SCALE_FACTOR)
        )

        if DRAW_CELL_BORDERS:
            q1 = _map_local_to_view_plane(view, x_min_cell, y_min_cell)
            q2 = _map_local_to_view_plane(view, x_max_cell, y_min_cell)
            q3 = _map_local_to_view_plane(view, x_max_cell, y_max_cell)
            q4 = _map_local_to_view_plane(view, x_min_cell, y_max_cell)
            _draw_detail_line(document, view, q1, q2)
            _draw_detail_line(document, view, q2, q3)
            _draw_detail_line(document, view, q3, q4)
            _draw_detail_line(document, view, q4, q1)

        for curve in all_curves:
            _draw_browser_curve(document, view, curve, cell_cx, cell_cy, bb_cx, bb_cy, scale)
        _draw_segment_labels(
            document,
            view,
            shape,
            bb_cx,
            bb_cy,
            scale,
            cell_cx,
            cell_cy,
            text_note_type,
            hook_labeled,
            seg_for_lbl,
            gx_min,
            gx_max,
            gy_draw_min,
            gy_draw_max,
        )
        drawn += 1

    return drawn, skipped, warn_labels


def main():
    uidoc = _get_uidoc()
    if uidoc is None:
        print(u"No hay ActiveUIDocument (__revit__).")
        return

    document = _get_doc()
    shapes = _collect_rebar_shapes(document)
    if not shapes:
        msg = u"No se encontró ningún RebarShape usable en este proyecto."
        print(msg)
        try:
            TaskDialog.Show(u"Rebar shapes", msg)
        except Exception:
            pass
        return

    vft = None
    if DRAW_DETAIL_VIEWS:
        vft = _first_drafting_view_family_type(document)
        if vft is None:
            msg = (
                u"No hay ViewFamilyType de familia Drafting en el proyecto. "
                u"Carga una plantilla con vista de dibujo / leyenda o crea una manualmente."
            )
            print(msg)
            try:
                TaskDialog.Show(u"Rebar shapes", msg)
            except Exception:
                pass
            return

    skipped_draw = []

    if DRAW_DETAIL_VIEWS:
        t = Transaction(document, u"RebarShape: tabla en vista de dibujo")
        t.Start()
        try:
            table_view = _create_or_get_table_view(document, vft, TABLE_VIEW_NAME)
            try:
                table_view.Scale = int(DRAFTING_VIEW_SCALE)
            except Exception:
                pass
            _clear_table_view_content(document, table_view)
            n_drawn, skipped_draw, warn_labels = _draw_shapes_table_grid(
                document, table_view, shapes
            )
            t.Commit()
        except Exception as ex:
            if t.HasStarted():
                t.RollBack()
            msg = u"Error al crear/actualizar la vista tabla:\n\n{}".format(str(ex))
            print(msg)
            try:
                TaskDialog.Show(u"Rebar shapes", msg)
            except Exception:
                pass
            return

        cols = max(1, int(GRID_COLS))
        n = len(shapes)
        rows = int(math.ceil(float(n) / float(cols))) if n else 0
        print(
            u"Vista tabla: {!r} — {} forma(s), grilla {} filas × {} columnas ({} dibujadas).".format(
                TABLE_VIEW_NAME, n, rows, cols, n_drawn
            )
        )
        if skipped_draw:
            print(u"Omitidas al dibujar ({}): ".format(len(skipped_draw)))
            for s in skipped_draw:
                print(u"  - {}".format(s))
        if warn_labels:
            print(warn_labels)

    ids = List[ElementId]()
    nombres = []
    for rs in shapes:
        try:
            ids.Add(rs.Id)
            nombres.append(rs.Name or u"")
        except Exception:
            pass

    uidoc.Selection.SetElementIds(ids)
    print(u"\nSeleccionados {} RebarShape(s).".format(ids.Count))
    for i, nm in enumerate(sorted(set(n for n in nombres if n)), 1):
        print(u"  {}. {}".format(i, nm))


main()
