# -*- coding: utf-8 -*-
"""
Bending Detail — esquema agrupado de armaduras de la misma forma.

Selecciona una barra matriz (define la forma), opcionalmente más barras de la
misma forma, y dibuja un croquis representativo con DetailCurves + TextNotes
en un único punto de inserción.

Revit 2024–2026 · IronPython 2.7 / pyRevit.
"""

from __future__ import print_function

import math

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    HorizontalTextAlignment,
    Line,
    Plane,
    SketchPlane,
    StorageType,
    TextNote,
    TextNoteOptions,
    TextNoteType,
    Transaction,
    Transform,
    VerticalTextAlignment,
    View3D,
    ViewSchedule,
    ViewSheet,
    XYZ,
)
from Autodesk.Revit.DB.Structure import MultiplanarOption, Rebar, RebarLayoutRule
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from System.Collections.Generic import List

try:
    from Autodesk.Revit.Exceptions import OperationCanceledException
except Exception:
    OperationCanceledException = Exception

_TOOL_TITLE = u"Arainco: Bending Detail"
_TXN_SKETCH = u"Arainco: Bending Detail SketchPlane"
_TXN_MAIN = u"Arainco: Bending Detail"

MM_PER_FOOT = 304.8
TARGET_TEXT_SIZE_MM = 2.5
TEXT_SIZE_TOLERANCE_FT = 0.0005
HOOK_ALLOWANCE_MM = 100
HOOK_LABEL_TEXT = u"100"
LENGTH_ROUND_STEP_MM = 10
DEFAULT_BAR_DIAM_MM = 10.0
ROTATION_ANGLE_EPS = 0.001
SEGMENT_OFFSET_MM = 2.5
LABEL_OFFSET_MM = 4.5
_ZERO_LEN_EPS = 1e-9

# Parámetros de forma típicos (letras de segmento / ganchos).
_SHAPE_PARAM_NAMES = frozenset(
    [u"A", u"B", u"C", u"D", u"E", u"F", u"G", u"H", u"J", u"K", u"O", u"R"]
)


def _as_unicode(text):
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _mostrar_aviso(uiapp, instruction, content=u"", ok_text=u"Entendido"):
    """Aviso WPF estándar; respaldo a TaskDialog."""
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


# ---------------------------------------------------------
# Filtros de selección
# ---------------------------------------------------------
class RebarSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, Rebar)

    def AllowReference(self, reference, position):
        return False


class StrictShapeRebarFilter(ISelectionFilter):
    def __init__(self, target_shape_id):
        self.target_shape_id = target_shape_id

    def AllowElement(self, elem):
        if not isinstance(elem, Rebar):
            return False
        try:
            return elem.GetShapeId() == self.target_shape_id
        except Exception:
            return False

    def AllowReference(self, reference, position):
        return False


# ---------------------------------------------------------
# Auxiliares
# ---------------------------------------------------------
def _view_supports_tool(view):
    if view is None:
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass
    if isinstance(view, (ViewSheet, ViewSchedule, View3D)):
        return False
    return True


def ensure_sketch_plane(doc, view):
    """Crea SketchPlane en la vista si falta. Retorna True si OK."""
    if view.SketchPlane is not None:
        return True
    t_sp = Transaction(doc, _TXN_SKETCH)
    t_sp.Start()
    try:
        plane = Plane.CreateByNormalAndOrigin(view.ViewDirection, view.Origin)
        view.SketchPlane = SketchPlane.Create(doc, plane)
        t_sp.Commit()
        return True
    except Exception:
        if t_sp.HasStarted():
            t_sp.RollBack()
        return False


def get_existing_25mm_text_type(doc):
    """Prefiere TextNoteType de 2.5 mm; si no, el más cercano por tamaño."""
    target_size_feet = TARGET_TEXT_SIZE_MM / MM_PER_FOOT
    collector = list(FilteredElementCollector(doc).OfClass(TextNoteType).ToElements())
    if not collector:
        return None

    best = None
    best_delta = None
    for t_type in collector:
        size_param = t_type.get_Parameter(BuiltInParameter.TEXT_SIZE)
        if not size_param:
            continue
        try:
            size_val = size_param.AsDouble()
        except Exception:
            continue
        delta = abs(size_val - target_size_feet)
        if delta < TEXT_SIZE_TOLERANCE_FT:
            return t_type
        if best_delta is None or delta < best_delta:
            best_delta = delta
            best = t_type
    return best if best is not None else collector[0]


def round_up_to_nearest_10(value):
    return int(math.ceil(float(value) / float(LENGTH_ROUND_STEP_MM)) * LENGTH_ROUND_STEP_MM)


def _collect_shape_nominals_mm(rebar):
    """Longitudes nominales (mm) desde parámetros de forma A–K / cortos."""
    valid_nominals = set()
    for p in rebar.Parameters:
        try:
            if p.StorageType != StorageType.Double:
                continue
            name = (p.Definition.Name or u"").strip().upper()
            # Nombres cortos (≤2) o letras de segmento conocidas.
            if len(name) > 2 and name not in _SHAPE_PARAM_NAMES:
                continue
            if len(name) <= 2 or name in _SHAPE_PARAM_NAMES:
                val_mm = int(round(p.AsDouble() * MM_PER_FOOT))
                if val_mm > 0:
                    valid_nominals.add(val_mm)
        except Exception:
            continue
    return list(valid_nominals)


def _bar_diameter_mm(rebar):
    try:
        diam_param = rebar.get_Parameter(BuiltInParameter.REBAR_BAR_DIAMETER)
        if diam_param is not None:
            return diam_param.AsDouble() * MM_PER_FOOT
    except (AttributeError, TypeError, ValueError):
        pass
    return DEFAULT_BAR_DIAM_MM


def get_segment_nominal_lengths_fuzzy(rebar, segment_curves):
    """Asocia cada segmento al nominal de forma más cercano (+ diámetro); ceiling 10 mm."""
    lengths_mm = []
    valid_nominals = _collect_shape_nominals_mm(rebar)

    if not valid_nominals:
        for curve in segment_curves:
            val_mm = int(round(curve.Length * MM_PER_FOOT))
            lengths_mm.append(round_up_to_nearest_10(val_mm))
        return lengths_mm

    bar_diam_mm = _bar_diameter_mm(rebar)
    for curve in segment_curves:
        curve_len_mm = curve.Length * MM_PER_FOOT
        adjusted_len = curve_len_mm + bar_diam_mm
        closest_nominal = min(valid_nominals, key=lambda x: abs(x - adjusted_len))
        lengths_mm.append(round_up_to_nearest_10(closest_nominal))
    return lengths_mm


def _hook_count(rebar):
    count = 0
    try:
        if rebar.GetHookTypeId(0) != ElementId.InvalidElementId:
            count += 1
        if rebar.GetHookTypeId(1) != ElementId.InvalidElementId:
            count += 1
    except Exception:
        pass
    return count


def _spacing_suffix(rebar):
    """Devuelve u'@NNN' o cadena vacía (solo si LayoutRule != Single)."""
    try:
        if rebar.LayoutRule != RebarLayoutRule.Single:
            spacing_mm = int(round(rebar.MaxSpacing * MM_PER_FOOT))
            return u"@{}".format(spacing_mm)
        return u""
    except (AttributeError, TypeError, ValueError):
        pass
    try:
        spacing_param = rebar.get_Parameter(BuiltInParameter.REBAR_ELEM_BAR_SPACING)
        if spacing_param is not None:
            spacing_mm = int(round(spacing_param.AsDouble() * MM_PER_FOOT))
            return u"@{}".format(spacing_mm)
    except (AttributeError, TypeError, ValueError):
        pass
    return u""


def _rebar_shape_display_name(doc, rebar):
    """Nombre visible del RebarShape (``.Name`` suele estar vacío; usar parámetros de tipo)."""
    shape_id = None
    try:
        shape_id = rebar.GetShapeId()
    except Exception:
        shape_id = None
    if shape_id is None or shape_id == ElementId.InvalidElementId:
        try:
            shape_id = rebar.RebarShapeId
        except Exception:
            shape_id = None
    if shape_id is None or shape_id == ElementId.InvalidElementId:
        return u""
    try:
        sh = doc.GetElement(shape_id)
    except Exception:
        return u""
    if sh is None:
        return u""
    try:
        n = _as_unicode(getattr(sh, "Name", None) or u"").strip()
        if n:
            return n
    except Exception:
        pass
    for bip_name in (u"SYMBOL_NAME_PARAM", u"ALL_MODEL_TYPE_NAME"):
        try:
            bip = getattr(BuiltInParameter, bip_name, None)
            if bip is None:
                continue
            p = sh.get_Parameter(bip)
            if p is None or not p.HasValue:
                continue
            if p.StorageType == StorageType.String:
                s = _as_unicode(p.AsString() or u"").strip()
                if s:
                    return s
        except Exception:
            continue
    return u""


def _shape_code_digits(shape_name):
    """Normaliza el nombre visible a dígitos (p. ej. '10', '08')."""
    if not shape_name:
        return u""
    s = _as_unicode(shape_name).strip()
    if s in (u"10", u"08"):
        return s
    digits = u"".join(ch for ch in s if ch in u"0123456789")
    if digits in (u"10", u"08"):
        return digits
    # Códigos con cero a la izquierda: "8" → no; solo "08"/"10" exactos tras normalizar.
    if len(digits) == 1 and digits == u"8":
        return u"08"
    return digits


def _shape_prefix(doc, rebar):
    """E. = estribo (shape 10), T. = traba (shape 08)."""
    code = _shape_code_digits(_rebar_shape_display_name(doc, rebar))
    if code == u"10":
        return u"E."
    if code == u"08":
        return u"T."
    return u""


def _rebar_normal(rebar, view):
    try:
        return rebar.GetShapeDrivenAccessor().Normal
    except Exception:
        return view.ViewDirection


def _build_final_transform(rebar_normal, view, base_point, insert_point):
    """Origen → alinear normal a ViewDirection → punto de inserción."""
    translation_to_origin = Transform.CreateTranslation(-base_point)
    view_dir = view.ViewDirection
    cross_prod = rebar_normal.CrossProduct(view_dir)
    angle = rebar_normal.AngleTo(view_dir)

    rotation = Transform.Identity
    if not cross_prod.IsAlmostEqualTo(XYZ.Zero) and angle > ROTATION_ANGLE_EPS:
        rotation = Transform.CreateRotation(cross_prod, angle)
    elif cross_prod.IsAlmostEqualTo(XYZ.Zero) and angle > (math.pi - ROTATION_ANGLE_EPS):
        # Normal casi anti-paralela: CrossProduct ~0; voltear con eje de la vista.
        aux_axis = view.RightDirection
        if aux_axis.IsAlmostEqualTo(XYZ.Zero):
            aux_axis = view.UpDirection
        if not aux_axis.IsAlmostEqualTo(XYZ.Zero):
            rotation = Transform.CreateRotation(aux_axis, math.pi)

    translation_to_insert = Transform.CreateTranslation(insert_point)
    return translation_to_insert.Multiply(rotation.Multiply(translation_to_origin))


def _bottom_center_in_view(all_points, view):
    """Centro inferior del croquis proyectado en ejes Right/Up de la vista."""
    right = view.RightDirection
    up = view.UpDirection
    view_dir = view.ViewDirection

    min_r = min(p.DotProduct(right) for p in all_points)
    max_r = max(p.DotProduct(right) for p in all_points)
    min_u = min(p.DotProduct(up) for p in all_points)
    # Profundidad media para no desalinear en secciones.
    avg_n = sum(p.DotProduct(view_dir) for p in all_points) / float(len(all_points))

    center_r = (min_r + max_r) / 2.0
    # Inferior: misma R, proyección Up mínima.
    return (right * center_r) + (up * min_u) + (view_dir * avg_n)


def _centroid_xyz(points):
    n = float(len(points))
    sx = sum(p.X for p in points) / n
    sy = sum(p.Y for p in points) / n
    sz = sum(p.Z for p in points) / n
    return XYZ(sx, sy, sz)


def _curve_mid_and_tangent(curve):
    """Punto medio y tangente unitaria de una curva transformada."""
    mid = curve.Evaluate(0.5, True)
    try:
        if isinstance(curve, Line):
            direction = curve.Direction
        else:
            derivs = curve.ComputeDerivatives(0.5, True)
            direction = derivs.BasisX
        if direction.GetLength() < _ZERO_LEN_EPS:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            direction = p1 - p0
        if direction.GetLength() < _ZERO_LEN_EPS:
            return mid, None
        return mid, direction.Normalize()
    except Exception:
        try:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            direction = p1 - p0
            if direction.GetLength() < _ZERO_LEN_EPS:
                return mid, None
            return mid, direction.Normalize()
        except Exception:
            return mid, None


def _outward_label_position(mid_point, tangent, view, centroid, offset_feet):
    """Desplaza el texto perpendicular al tramo, hacia fuera del centroide del croquis."""
    if mid_point is None:
        return None
    if tangent is None or centroid is None:
        return mid_point
    offset_dir = view.ViewDirection.CrossProduct(tangent)
    if offset_dir.GetLength() < _ZERO_LEN_EPS:
        return mid_point
    offset_dir = offset_dir.Normalize()
    # Elegir el lado que se aleja del centroide (exterior del esquema).
    to_mid = mid_point - centroid
    if offset_dir.DotProduct(to_mid) < 0.0:
        offset_dir = XYZ(-offset_dir.X, -offset_dir.Y, -offset_dir.Z)
    return mid_point + (offset_dir * offset_feet)


def _hook_curves_for_labels(rebar):
    """
    Curvas de gancho (sin radio de doblez) en extremos 0 y/o 1.
    Retorna lista de curvas a etiquetar con 100 mm.
    """
    hooks = []
    try:
        with_hooks = list(
            rebar.GetCenterlineCurves(
                False, False, True, MultiplanarOption.IncludeAllMultiplanarCurves, 0
            )
            or []
        )
    except Exception:
        with_hooks = []
    if not with_hooks:
        return hooks

    has0 = False
    has1 = False
    try:
        has0 = rebar.GetHookTypeId(0) != ElementId.InvalidElementId
    except Exception:
        pass
    try:
        has1 = rebar.GetHookTypeId(1) != ElementId.InvalidElementId
    except Exception:
        pass

    if has0:
        hooks.append(with_hooks[0])
    if has1 and (not has0 or len(with_hooks) > 1):
        hooks.append(with_hooks[-1])
    return hooks


# ---------------------------------------------------------
# Flujo principal
# ---------------------------------------------------------
def run(uiapp):
    """Entrada desde pushbutton ligero."""
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        _mostrar_aviso(uiapp, u"No hay documento activo.")
        return
    doc = uidoc.Document
    view = doc.ActiveView

    if not _view_supports_tool(view):
        _mostrar_aviso(
            uiapp,
            u"La vista actual no admite esquema con DetailCurves / PickPoint.",
            u"Use planta, sección, alzado o vista de detalle (no 3D, lámina ni tabla).",
        )
        return

    if not ensure_sketch_plane(doc, view):
        _mostrar_aviso(
            uiapp,
            u"No se pudo configurar el plano de trabajo (SketchPlane) en la vista.",
        )
        return

    # 1. Armadura matriz
    try:
        first_ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            RebarSelectionFilter(),
            u"1/2: Selecciona la PRIMERA armadura (esto definirá la forma permitida)",
        )
        first_rebar = doc.GetElement(first_ref)
        target_shape_id = first_rebar.GetShapeId()
    except OperationCanceledException:
        return

    # 2. Selección múltiple restringida
    try:
        strict_filter = StrictShapeRebarFilter(target_shape_id)
        other_refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            strict_filter,
            u"2/2: Selecciona más armaduras (Finalizar en la barra de opciones para continuar)",
        )
    except OperationCanceledException:
        other_refs = []

    # 3. IDs únicos
    unique_rebar_ids = set()
    unique_rebar_ids.add(first_rebar.Id.IntegerValue)
    for ref in other_refs:
        unique_rebar_ids.add(ref.ElementId.IntegerValue)
    total_rebars = len(unique_rebar_ids)

    # 4. Punto de inserción
    try:
        if total_rebars > 1:
            prompt_msg = (
                u"Selecciona el punto de inserción para el esquema agrupado "
                u"({} elementos)".format(total_rebars)
            )
        else:
            prompt_msg = u"Selecciona el punto de inserción para el esquema"
        insert_point = uidoc.Selection.PickPoint(prompt_msg)
    except OperationCanceledException:
        return
    except Exception as ex:
        _mostrar_aviso(
            uiapp,
            u"La vista actual no admite selección de puntos.",
            _as_unicode(ex),
        )
        return

    rebar = first_rebar
    drawn_curves = rebar.GetCenterlineCurves(
        False, False, False, MultiplanarOption.IncludeAllMultiplanarCurves, 0
    )
    segment_curves = rebar.GetCenterlineCurves(
        False, True, True, MultiplanarOption.IncludeAllMultiplanarCurves, 0
    )

    if not drawn_curves or not segment_curves:
        _mostrar_aviso(
            uiapp,
            u"No se pudieron extraer las curvas del rebar representativo.",
        )
        return

    nominal_lengths_mm = get_segment_nominal_lengths_fuzzy(rebar, segment_curves)
    exact_length_mm = sum(nominal_lengths_mm) + (_hook_count(rebar) * HOOK_ALLOWANCE_MM)
    total_length_calc = round_up_to_nearest_10(exact_length_mm)

    base_point = drawn_curves[0].GetEndPoint(0)
    rebar_normal = _rebar_normal(rebar, view)
    final_transform = _build_final_transform(
        rebar_normal, view, base_point, insert_point
    )

    prefix = _shape_prefix(doc, rebar)
    diam_mm = int(round(_bar_diameter_mm(rebar)))
    spacing_str = _spacing_suffix(rebar)
    count_prefix = u"{}".format(total_rebars) if total_rebars > 1 else u""
    # Formato: 2T.ø10@100  (cantidad + E.|T. + ø + diam + @espaciado)
    main_label = u"{}{}\u00F8{}{}\nL={}".format(
        count_prefix, prefix, diam_mm, spacing_str, total_length_calc
    )

    text_type = get_existing_25mm_text_type(doc)
    if text_type is None:
        _mostrar_aviso(
            uiapp,
            u"No hay tipos de texto (TextNoteType) en el documento.",
            u"Se dibujarán solo las curvas del esquema.",
        )

    hook_curves = _hook_curves_for_labels(rebar)

    t = Transaction(doc, _TXN_MAIN)
    t.Start()
    try:
        created_elements = List[ElementId]()
        all_points = []

        for curve in drawn_curves:
            transformed_curve = curve.CreateTransformed(final_transform)
            detail_curve = doc.Create.NewDetailCurve(view, transformed_curve)
            created_elements.Add(detail_curve.Id)
            for pt in transformed_curve.Tessellate():
                all_points.append(pt)

        if text_type is not None and all_points:
            centroid = _centroid_xyz(all_points)
            segment_text_options = TextNoteOptions()
            segment_text_options.TypeId = text_type.Id
            segment_text_options.HorizontalAlignment = HorizontalTextAlignment.Center
            segment_text_options.VerticalAlignment = VerticalTextAlignment.Middle

            view_scale = view.Scale
            offset_feet = (SEGMENT_OFFSET_MM * view_scale) / MM_PER_FOOT

            for i, curve in enumerate(segment_curves):
                if i < len(nominal_lengths_mm):
                    length_str = str(nominal_lengths_mm[i])
                else:
                    raw_len = int(round(curve.Length * MM_PER_FOOT))
                    length_str = str(round_up_to_nearest_10(raw_len))

                transformed_seg = curve.CreateTransformed(final_transform)
                mid_point, tangent = _curve_mid_and_tangent(transformed_seg)
                text_pos = _outward_label_position(
                    mid_point, tangent, view, centroid, offset_feet
                )
                if text_pos is None:
                    text_pos = mid_point

                seg_note = TextNote.Create(
                    doc, view.Id, text_pos, length_str, segment_text_options
                )
                created_elements.Add(seg_note.Id)

            # Etiquetas parciales de ganchos (siempre 100 mm).
            for hook_curve in hook_curves:
                try:
                    transformed_hook = hook_curve.CreateTransformed(final_transform)
                except Exception:
                    continue
                mid_point, tangent = _curve_mid_and_tangent(transformed_hook)
                text_pos = _outward_label_position(
                    mid_point, tangent, view, centroid, offset_feet
                )
                if text_pos is None:
                    text_pos = mid_point
                hook_note = TextNote.Create(
                    doc, view.Id, text_pos, HOOK_LABEL_TEXT, segment_text_options
                )
                created_elements.Add(hook_note.Id)

            bottom_center_pt = _bottom_center_in_view(all_points, view)
            main_text_options = TextNoteOptions()
            main_text_options.TypeId = text_type.Id
            main_text_options.HorizontalAlignment = HorizontalTextAlignment.Center
            main_text_options.VerticalAlignment = VerticalTextAlignment.Top

            label_offset_feet = (LABEL_OFFSET_MM * view_scale) / MM_PER_FOOT
            final_label_pos = bottom_center_pt - (view.UpDirection * label_offset_feet)

            main_note = TextNote.Create(
                doc, view.Id, final_label_pos, main_label, main_text_options
            )
            created_elements.Add(main_note.Id)

        if created_elements.Count > 0:
            doc.Create.NewGroup(created_elements)

        t.Commit()
    except Exception as ex:
        if t.HasStarted():
            t.RollBack()
        _mostrar_aviso(uiapp, u"El esquema falló.", _as_unicode(ex))
