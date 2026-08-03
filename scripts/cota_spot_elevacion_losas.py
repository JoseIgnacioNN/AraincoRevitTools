# -*- coding: utf-8 -*-
"""
Cota alineada + Spot Elevations en caras superiores de losas (sección/alzado).

Revit 2024+ | pyRevit (IronPython).

Flujo:
1. Validar vista Sección o Alzado (no plantilla).
2. Selección múltiple de Floors.
3. PickPoint para ubicar la cota.
4. Crear cota alineada entre caras superiores horizontales + Spot Elevation por losa.
5. Ocultar lo creado en el resto del árbol de vistas dependientes (solo visible en la activa).
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    FilteredElementCollector,
    HostObjectUtils,
    Line,
    Plane,
    PlanarFace,
    ReferenceArray,
    SketchPlane,
    SpotDimensionType,
    Transaction,
    TransactionGroup,
    UnitTypeId,
    UnitUtils,
    ViewType,
    XYZ,
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

try:
    from Autodesk.Revit.Exceptions import OperationCanceledException
except Exception:
    OperationCanceledException = Exception

_DIALOG_TITLE = u"Arainco: Cota y Spot Losas"
_TXN_GROUP = u"Arainco: Cota y Spot Elevations losas"
_TXN_SKETCH = u"Arainco: SketchPlane para cota losas"
_TXN_CREATE = u"Arainco: Generar cota y Spot Elevations losas"

# Tipo de Spot Elevation por categoría (nombre exacto en el proyecto).
CATEGORY_SPOT_MAPPING = {
    int(BuiltInCategory.OST_Floors): u"Survey Point_Nivel Tope de Losa",
}

# Offsets en mm → pies internos vía UnitUtils.
_OFFSET_LEADER_MM = 450.0
_OFFSET_END_MM = 150.0
_DIM_LINE_LENGTH_MM = 3000.0
_HORIZONTAL_DOT_MIN = 0.9999


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _mm_to_internal(mm):
    return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)


def mostrar_aviso(uiapp, instruction, content=u""):
    """Aviso WPF estándar; respaldo a TaskDialog."""
    try:
        from bimtools_instruction_dialog import show_message_dialog
        from revit_wpf_window_position import revit_main_hwnd

        hwnd = revit_main_hwnd(uiapp) if uiapp is not None else None
        show_message_dialog(
            _DIALOG_TITLE,
            instruction=_as_unicode(instruction),
            content=_as_unicode(content) if content else None,
            ok_text=u"Entendido",
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
        TaskDialog.Show(_DIALOG_TITLE, msg)
    except Exception:
        pass


class FloorSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            if elem is None or elem.Category is None:
                return False
            return elem.Category.Id.IntegerValue == int(BuiltInCategory.OST_Floors)
        except Exception:
            return False

    def AllowReference(self, ref, pos):
        return False


def _is_invalid_id(eid):
    try:
        return eid is None or eid == ElementId.InvalidElementId
    except Exception:
        return True


def _get_spot_type_by_name(doc, type_name):
    collector = FilteredElementCollector(doc).OfClass(SpotDimensionType)
    for spot_type in collector:
        try:
            if spot_type.Name == type_name:
                return spot_type
        except Exception:
            continue
    return None


def _ensure_sketch_plane(doc, active_view):
    """Crea SketchPlane en la vista si falta. Retorna True si se creó o ya existía."""
    if active_view.SketchPlane is not None:
        return True
    t_sp = Transaction(doc, _TXN_SKETCH)
    t_sp.Start()
    try:
        plane = Plane.CreateByNormalAndOrigin(
            active_view.ViewDirection, active_view.Origin
        )
        sketch_plane = SketchPlane.Create(doc, plane)
        active_view.SketchPlane = sketch_plane
        t_sp.Commit()
        return True
    except Exception:
        if t_sp.HasStarted():
            t_sp.RollBack()
        return False


def _project_point_on_face(planar_face, pt, view_up):
    """Proyecta el clic sobre la cara planar (intersección con Up de la vista)."""
    normal = planar_face.FaceNormal
    face_origin = planar_face.Origin
    denominator = normal.DotProduct(view_up)
    if abs(denominator) > 1e-6:
        t_param = -(normal.DotProduct(pt - face_origin)) / denominator
        return pt + view_up * t_param
    try:
        result = planar_face.Project(pt)
        if result is not None:
            return result.XYZPoint
    except Exception:
        pass
    return face_origin


def _collect_floor_spot_data(doc, sel_refs, pt, view_up):
    """
    Filtra losas horizontales y prepara referencias + puntos de spot.
    Retorna (ref_array, spots_data, discarded_slope, discarded_other).
    """
    ref_array = ReferenceArray()
    spots_data = []
    discarded_slope = 0
    discarded_other = 0

    for ref in sel_refs:
        elem = doc.GetElement(ref.ElementId)
        if elem is None:
            discarded_other += 1
            continue
        try:
            top_faces = HostObjectUtils.GetTopFaces(elem)
        except Exception:
            discarded_other += 1
            continue
        try:
            n_faces = int(top_faces.Count)
        except Exception:
            try:
                n_faces = len(top_faces)
            except Exception:
                n_faces = 0
        if n_faces < 1:
            discarded_other += 1
            continue

        top_face_ref = top_faces[0]
        try:
            planar_face = elem.GetGeometryObjectFromReference(top_face_ref)
        except Exception:
            planar_face = None
        if planar_face is None or not isinstance(planar_face, PlanarFace):
            discarded_other += 1
            continue

        normal = planar_face.FaceNormal
        if abs(normal.Z) < _HORIZONTAL_DOT_MIN:
            discarded_slope += 1
            continue

        intersection_pt = _project_point_on_face(planar_face, pt, view_up)
        ref_array.Append(top_face_ref)
        cat_id = elem.Category.Id.IntegerValue
        spots_data.append((top_face_ref, intersection_pt, cat_id, elem.Id))

    return ref_array, spots_data, discarded_slope, discarded_other


def _views_to_hide_created_elements(doc, active_view):
    """
    Vistas del mismo árbol de dependientes excepto la activa.
    - En dependiente: padre + hermanas.
    - En primaria: todas las hijas dependientes.
    """
    views = []
    try:
        primary_view_id = active_view.GetPrimaryViewId()
    except Exception:
        return views

    if not _is_invalid_id(primary_view_id):
        primary_view = doc.GetElement(primary_view_id)
        if primary_view is None:
            return views
        views.append(primary_view)
        try:
            dep_ids = primary_view.GetDependentViewIds()
        except Exception:
            dep_ids = []
        for dep_id in dep_ids:
            if dep_id == active_view.Id:
                continue
            v = doc.GetElement(dep_id)
            if v is not None:
                views.append(v)
        return views

    try:
        dep_ids = active_view.GetDependentViewIds()
    except Exception:
        dep_ids = []
    for dep_id in dep_ids:
        v = doc.GetElement(dep_id)
        if v is not None:
            views.append(v)
    return views


def _hide_in_views(views, element_ids, hide_warnings):
    for v in views:
        if v is None:
            continue
        try:
            v.HideElements(element_ids)
        except Exception as ex:
            try:
                name = v.Name
            except Exception:
                name = u"(sin nombre)"
            hide_warnings.append(
                u"No se pudo ocultar en «{0}»: {1}".format(name, _as_unicode(ex))
            )


def create_smart_dimensions_and_spots(uiapp):
    uidoc = uiapp.ActiveUIDocument if uiapp is not None else None
    if uidoc is None:
        mostrar_aviso(uiapp, u"No hay documento activo.")
        return

    doc = uidoc.Document
    active_view = doc.ActiveView
    if active_view is None:
        mostrar_aviso(uiapp, u"No hay vista activa.")
        return

    try:
        if active_view.IsTemplate:
            mostrar_aviso(
                uiapp,
                u"Vista incorrecta",
                u"No se puede ejecutar sobre una plantilla de vista.",
            )
            return
    except Exception:
        pass

    if active_view.ViewType not in (ViewType.Section, ViewType.Elevation):
        mostrar_aviso(
            uiapp,
            u"Vista incorrecta",
            u"Ejecute la herramienta en una vista de Sección o Alzado.",
        )
        return

    try:
        sel_refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            FloorSelectionFilter(),
            u"Seleccione losas (Finish para confirmar, Esc para cancelar).",
        )
    except OperationCanceledException:
        return
    except Exception as ex:
        mostrar_aviso(uiapp, u"No se pudo completar la selección.", _as_unicode(ex))
        return

    if sel_refs is None or len(sel_refs) < 2:
        mostrar_aviso(
            uiapp,
            u"Selección insuficiente",
            u"Seleccione al menos dos losas para crear la cota alineada.",
        )
        return

    tg = TransactionGroup(doc, _TXN_GROUP)
    tg.Start()
    try:
        if not _ensure_sketch_plane(doc, active_view):
            tg.RollBack()
            mostrar_aviso(
                uiapp,
                u"No se pudo configurar el plano de trabajo (SketchPlane) en la vista.",
            )
            return

        try:
            pt = uidoc.Selection.PickPoint(
                u"Haga clic para definir la posición de la cota."
            )
        except OperationCanceledException:
            tg.RollBack()
            return
        except Exception as ex:
            tg.RollBack()
            mostrar_aviso(uiapp, u"No se pudo obtener el punto de cota.", _as_unicode(ex))
            return

        view_up = active_view.UpDirection
        view_right = active_view.RightDirection
        ref_array, spots_data, discarded_slope, discarded_other = _collect_floor_spot_data(
            doc, sel_refs, pt, view_up
        )

        if ref_array.Size < 2:
            tg.RollBack()
            parts = [
                u"No hay suficientes losas horizontales para crear la cota "
                u"(se requieren al menos 2 caras superiores planas horizontales)."
            ]
            if discarded_slope:
                parts.append(
                    u"Descartadas por pendiente: {}.".format(discarded_slope)
                )
            if discarded_other:
                parts.append(
                    u"Descartadas por geometría inválida: {}.".format(discarded_other)
                )
            mostrar_aviso(uiapp, u"No se pudo crear la cota", u"\n".join(parts))
            return

        spot_types_cache = {}
        missing_types = set()
        created_ids = List[ElementId]()
        spot_ok = 0
        spot_fail = 0
        hide_warnings = []

        offset_leader = _mm_to_internal(_OFFSET_LEADER_MM)
        offset_end = _mm_to_internal(_OFFSET_END_MM)
        dim_len = _mm_to_internal(_DIM_LINE_LENGTH_MM)

        t = Transaction(doc, _TXN_CREATE)
        t.Start()
        try:
            dim_line = Line.CreateBound(pt, pt + view_up * dim_len)
            new_dim = doc.Create.NewDimension(active_view, dim_line, ref_array)
            if new_dim is None:
                raise Exception(u"NewDimension devolvió None.")
            created_ids.Add(new_dim.Id)

            for face_ref, intersect_pt, cat_id, _elem_id in spots_data:
                bend = intersect_pt - view_right * offset_leader
                end = bend - view_right * offset_end
                try:
                    new_spot = doc.Create.NewSpotElevation(
                        active_view,
                        face_ref,
                        intersect_pt,
                        bend,
                        end,
                        intersect_pt,
                        True,
                    )
                    if new_spot is None:
                        spot_fail += 1
                        continue
                    created_ids.Add(new_spot.Id)
                    spot_ok += 1

                    target_type_name = CATEGORY_SPOT_MAPPING.get(cat_id)
                    if not target_type_name:
                        continue
                    if target_type_name not in spot_types_cache:
                        spot_types_cache[target_type_name] = _get_spot_type_by_name(
                            doc, target_type_name
                        )
                    target_type = spot_types_cache[target_type_name]
                    if target_type is None:
                        missing_types.add(target_type_name)
                        continue
                    try:
                        new_spot.ChangeTypeId(target_type.Id)
                    except Exception:
                        missing_types.add(target_type_name)
                except Exception:
                    spot_fail += 1

            if created_ids.Count > 0:
                views_hide = _views_to_hide_created_elements(doc, active_view)
                _hide_in_views(views_hide, created_ids, hide_warnings)

            t.Commit()
        except Exception as ex:
            if t.HasStarted():
                t.RollBack()
            tg.RollBack()
            mostrar_aviso(
                uiapp,
                u"Error al crear cota o Spot Elevations.",
                _as_unicode(ex),
            )
            return

        tg.Assimilate()

        summary = [
            u"Cota alineada: 1.",
            u"Spot Elevations creados: {}.".format(spot_ok),
        ]
        if spot_fail:
            summary.append(u"Spot Elevations fallidos: {}.".format(spot_fail))
        if discarded_slope:
            summary.append(u"Losas descartadas por pendiente: {}.".format(discarded_slope))
        if discarded_other:
            summary.append(
                u"Losas descartadas por geometría inválida: {}.".format(discarded_other)
            )
        if missing_types:
            summary.append(
                u"Tipos de Spot no encontrados en el proyecto (se usó el tipo por defecto):\n- "
                + u"\n- ".join(sorted(missing_types))
            )
        if hide_warnings:
            summary.append(u"Avisos de ocultamiento:\n" + u"\n".join(hide_warnings))
        else:
            summary.append(
                u"Elementos ocultos en el resto del árbol de vistas dependientes "
                u"(visibles solo en la vista activa)."
            )

        mostrar_aviso(uiapp, u"Operación completada", u"\n".join(summary))

    except Exception as ex:
        try:
            if tg.HasStarted():
                tg.RollBack()
        except Exception:
            pass
        mostrar_aviso(uiapp, u"Error inesperado", _as_unicode(ex))


def run(uiapp):
    create_smart_dimensions_and_spots(uiapp)
