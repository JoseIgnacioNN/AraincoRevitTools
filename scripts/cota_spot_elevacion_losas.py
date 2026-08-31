# -*- coding: utf-8 -*-
"""
Cota alineada + Spot Elevations en caras superiores de losas (sección/alzado).

Revit 2024+ | pyRevit (IronPython).

Flujo:
1. Validar vista Sección o Alzado (no plantilla).
2. Selección múltiple de Floors.
3. PickPoint para ubicar la cota.
4. Crear cota alineada entre caras superiores horizontales + Spot Elevation por losa
   (origen sobre la cara; líder hacia la posición de la cota).
5. Ocultar lo creado en el resto del árbol de vistas dependientes (solo visible en la activa).
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    DimensionStyleType,
    ElementId,
    FilteredElementCollector,
    GeometryInstance,
    HostObjectUtils,
    Line,
    Options,
    Plane,
    PlanarFace,
    ReferenceArray,
    SketchPlane,
    Solid,
    SpotDimensionType,
    Transaction,
    TransactionGroup,
    UnitTypeId,
    UnitUtils,
    UV,
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
_SPOT_FLOOR = u"Survey Point_Nivel Tope de Losa"
_SPOT_RADIER = u"Survey Point_Nivel Tope de Radier"
_SPOT_FOUNDATION = u"Survey Point_Nivel Sello de Fundacion"
_SPOT_WALL = u"Survey Point_Nivel Tope de Concreto"
_SPOT_FRAMING_CONCRETE = u"Survey Point_Nivel Tope de Concreto"
_SPOT_FRAMING_STEEL = u"Survey Point_Nivel Tope de Acero"

CATEGORY_SPOT_MAPPING = {
    int(BuiltInCategory.OST_Floors): _SPOT_FLOOR,
    int(BuiltInCategory.OST_StructuralFoundation): _SPOT_FOUNDATION,
    int(BuiltInCategory.OST_Walls): _SPOT_WALL,
    int(BuiltInCategory.OST_StructuralFraming): _SPOT_FRAMING_CONCRETE,
}

_CAT_FLOORS = int(BuiltInCategory.OST_Floors)
_CAT_FOUNDATIONS = int(BuiltInCategory.OST_StructuralFoundation)
_CAT_WALLS = int(BuiltInCategory.OST_Walls)
_CAT_FRAMING = int(BuiltInCategory.OST_StructuralFraming)
_ALLOWED_ELEV_CATS = (_CAT_FLOORS, _CAT_FOUNDATIONS, _CAT_WALLS, _CAT_FRAMING)

# Offsets en mm de modelo, calibrados a escala de vista 1:50.
# En otras escalas se escalan: mm = mm_at_50 * (Scale / 50) para mantener
# la misma separación aparente en papel.
_REF_VIEW_SCALE = 50
_OFFSET_LEADER_MM_AT_50 = 450.0
_OFFSET_END_MM_AT_50 = 150.0
_LEADER_SHOULDER_MM_AT_50 = 300.0
# Separación Spot respecto a la línea de cota (hacia fuera del modelo) @ 1:50.
_SPOT_OFFSET_PAST_DIM_MM_AT_50 = 450.0
# Segunda cota de fundación: separación hacia el modelo respecto a la principal @ 1:50.
_SECOND_DIM_OFFSET_MM_AT_50 = 350.0
_DIM_LINE_LENGTH_MM = 3000.0
_HORIZONTAL_DOT_MIN = 0.9999
_MIN_LEADER_LEN_MM = 25.0


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _mm_to_internal(mm):
    return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)


def _view_scale_int(view):
    """Denominador de escala de vista (50 → 1:50). Fallback a la escala de calibración."""
    try:
        s = int(view.Scale)
        if s > 0:
            return s
    except Exception:
        pass
    return _REF_VIEW_SCALE


def _mm_scaled_to_view(mm_at_ref_scale, view):
    """Convierte un offset calibrado a 1:50 a mm de modelo según Scale de la vista."""
    scale = _view_scale_int(view)
    return float(mm_at_ref_scale) * (float(scale) / float(_REF_VIEW_SCALE))


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
        return _is_elevation_host(elem)

    def AllowReference(self, ref, pos):
        return False


def _is_invalid_id(eid):
    try:
        return eid is None or eid == ElementId.InvalidElementId
    except Exception:
        return True


def _canon_type_name(name):
    s = _as_unicode(name).strip().lower()
    for src, dst in (
        (u"á", u"a"),
        (u"é", u"e"),
        (u"í", u"i"),
        (u"ó", u"o"),
        (u"ú", u"u"),
        (u"ü", u"u"),
        (u"ñ", u"n"),
    ):
        s = s.replace(src, dst)
    return u" ".join(s.split())


def _spot_type_display_name(elem):
    if elem is None:
        return u""
    try:
        n = elem.Name
        if n:
            s = _as_unicode(n).strip()
            if s:
                return s
    except Exception:
        pass
    for bip in (
        BuiltInParameter.ALL_MODEL_TYPE_NAME,
        BuiltInParameter.SYMBOL_NAME_PARAM,
        BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM,
    ):
        try:
            p = elem.get_Parameter(bip)
            if p is None or not p.HasValue:
                continue
            s = _as_unicode(p.AsString() or p.AsValueString() or u"").strip()
            if s:
                return s
        except Exception:
            continue
    return u""


def _spot_type_key(name):
    """Clave comparable: quita prefijo 'survey point_' y acentos."""
    n = _canon_type_name(name)
    marker = u"survey point_"
    if marker in n:
        n = n.split(marker, 1)[1].strip()
    return n


def _build_spot_type_registry(doc):
    rows = []
    for spot_type in FilteredElementCollector(doc).OfClass(SpotDimensionType):
        try:
            if spot_type.StyleType != DimensionStyleType.SpotElevation:
                continue
        except Exception:
            pass
        name = _spot_type_display_name(spot_type)
        if not name:
            continue
        rows.append((name, _canon_type_name(name), _spot_type_key(name), spot_type))
    return rows


def _get_spot_type_by_name(doc, type_name, registry=None, cache=None):
    wanted = _as_unicode(type_name)
    if not wanted:
        return None
    if cache is not None and wanted in cache:
        return cache[wanted]
    if registry is None:
        registry = _build_spot_type_registry(doc)
    wanted_c = _canon_type_name(wanted)
    wanted_k = _spot_type_key(wanted)
    hit = None
    for name, norm, key, spot_type in registry:
        if name == wanted or norm == wanted_c or key == wanted_k:
            hit = spot_type
            break
    if hit is None and wanted_k:
        for name, norm, key, spot_type in registry:
            if wanted_k in key or key in wanted_k:
                hit = spot_type
                break
    if cache is not None:
        cache[wanted] = hit
    return hit


def _apply_spot_type(spot, target_type):
    if spot is None or target_type is None:
        return False
    try:
        spot.ChangeTypeId(target_type.Id)
    except Exception:
        pass
    try:
        spot.DimensionType = target_type
    except Exception:
        pass
    try:
        return _eid_key(spot.GetTypeId()) == _eid_key(target_type.Id)
    except Exception:
        return False


def _category_id_int(elem):
    try:
        if elem is None or elem.Category is None:
            return None
        return int(elem.Category.Id.IntegerValue)
    except Exception:
        return None


def _is_elevation_host(elem):
    return _category_id_int(elem) in _ALLOWED_ELEV_CATS


def _is_floor(elem):
    return _is_elevation_host(elem)


def collect_floors_from_ids(doc, eids):
    floors = []
    seen = set()
    if doc is None or eids is None:
        return floors
    for eid in eids:
        try:
            el = doc.GetElement(eid)
        except Exception:
            el = None
        if not _is_elevation_host(el):
            continue
        try:
            key = int(el.Id.IntegerValue)
        except Exception:
            try:
                key = int(el.Id.Value)
            except Exception:
                continue
        if key in seen:
            continue
        seen.add(key)
        floors.append(el)
    return floors


def preselected_floors(uidoc):
    if uidoc is None:
        return []
    try:
        eids = uidoc.Selection.GetElementIds()
    except Exception:
        return []
    return collect_floors_from_ids(uidoc.Document, eids)


def is_section_or_elevation_view(view):
    if view is None:
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass
    try:
        return view.ViewType in (ViewType.Section, ViewType.Elevation)
    except Exception:
        return False


def _ensure_sketch_plane(doc, active_view, txn_name=None):
    """Crea SketchPlane en la vista si falta. Retorna True si se creó o ya existía."""
    if active_view.SketchPlane is not None:
        return True
    t_sp = Transaction(doc, txn_name or _TXN_SKETCH)
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
        try:
            if t_sp.HasStarted():
                t_sp.RollBack()
        except Exception:
            pass
        return False


def _annotation_on_face_plane(planar_face, pt, view_up):
    """
    Punto de anotación (línea de cota) en el plano infinito de la cara.
    Puede quedar fuera del polígono de la losa; sirve para el extremo del líder.
    """
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


def _xyz_avg(points):
    if not points:
        return None
    n = float(len(points))
    try:
        return XYZ(
            sum(float(p.X) for p in points) / n,
            sum(float(p.Y) for p in points) / n,
            sum(float(p.Z) for p in points) / n,
        )
    except Exception:
        return None


def _point_accepted_on_face(planar_face, candidate):
    """Solo acepta puntos que Face.Project confirma sobre la cara."""
    if planar_face is None or candidate is None:
        return None
    try:
        ir = planar_face.Project(candidate)
        if ir is None:
            return None
        return ir.XYZPoint
    except Exception:
        return None


def _face_interior_from_uv_grid(planar_face):
    """Muestrea la caja UV; el centro bbox falla en losas en L / con huecos."""
    try:
        bb = planar_face.GetBoundingBox()
    except Exception:
        return None
    try:
        u0 = float(bb.Min.U)
        u1 = float(bb.Max.U)
        v0 = float(bb.Min.V)
        v1 = float(bb.Max.V)
    except Exception:
        return None
    if abs(u1 - u0) < 1e-12 or abs(v1 - v0) < 1e-12:
        return None

    fractions = (
        (0.5, 0.5),
        (0.33, 0.33),
        (0.33, 0.67),
        (0.67, 0.33),
        (0.67, 0.67),
        (0.25, 0.5),
        (0.75, 0.5),
        (0.5, 0.25),
        (0.5, 0.75),
        (0.2, 0.2),
        (0.2, 0.8),
        (0.8, 0.2),
        (0.8, 0.8),
        (0.15, 0.5),
        (0.85, 0.5),
        (0.5, 0.15),
        (0.5, 0.85),
    )
    for fu, fv in fractions:
        try:
            uv = UV(u0 + (u1 - u0) * fu, v0 + (v1 - v0) * fv)
            pt = planar_face.Evaluate(uv)
        except Exception:
            continue
        accepted = _point_accepted_on_face(planar_face, pt)
        if accepted is not None:
            return accepted

    steps = 8
    for i in range(1, steps):
        for j in range(1, steps):
            try:
                uv = UV(
                    u0 + (u1 - u0) * (float(i) / steps),
                    v0 + (v1 - v0) * (float(j) / steps),
                )
                pt = planar_face.Evaluate(uv)
            except Exception:
                continue
            accepted = _point_accepted_on_face(planar_face, pt)
            if accepted is not None:
                return accepted
    return None


def _face_interior_from_triangulation(planar_face):
    """Centroides de triángulos de la malla de la cara."""
    try:
        mesh = planar_face.Triangulate()
    except Exception:
        return None
    if mesh is None:
        return None
    try:
        ntri = int(mesh.NumTriangles)
    except Exception:
        return None
    for i in range(ntri):
        try:
            tri = mesh.get_Triangle(i)
            pts = [tri.get_Vertex(0), tri.get_Vertex(1), tri.get_Vertex(2)]
        except Exception:
            continue
        c = _xyz_avg(pts)
        accepted = _point_accepted_on_face(planar_face, c)
        if accepted is not None:
            return accepted
        for vt in pts:
            mid = _xyz_avg([c, vt]) if c is not None else vt
            accepted = _point_accepted_on_face(planar_face, mid)
            if accepted is not None:
                return accepted
    return None


def _face_interior_from_edges(planar_face):
    """Puntos medios de aristas, ligeramente hacia el origen de la cara."""
    try:
        edge_loops = planar_face.EdgeLoops
    except Exception:
        return None
    if edge_loops is None:
        return None
    try:
        face_origin = planar_face.Origin
    except Exception:
        face_origin = None
    try:
        for loop in edge_loops:
            for edge in loop:
                try:
                    curve = edge.AsCurve()
                    mid = curve.Evaluate(0.5, True)
                except Exception:
                    continue
                accepted = _point_accepted_on_face(planar_face, mid)
                if accepted is not None:
                    return accepted
                if face_origin is not None and mid is not None:
                    try:
                        toward = face_origin - mid
                        ln = toward.GetLength()
                        if ln > 1e-9:
                            nudged = mid + toward * (min(0.05, ln * 0.1) / ln)
                            accepted = _point_accepted_on_face(planar_face, nudged)
                            if accepted is not None:
                                return accepted
                    except Exception:
                        pass
    except Exception:
        pass
    return None


def _face_interior_point(planar_face):
    """
    Punto interior que Face.Project acepta.
    El centro del BoundingBox UV suele caer fuera en losas en L o con huecos.
    """
    if planar_face is None or not isinstance(planar_face, PlanarFace):
        return None
    for finder in (
        _face_interior_from_uv_grid,
        _face_interior_from_triangulation,
        _face_interior_from_edges,
    ):
        try:
            pt = finder(planar_face)
        except Exception:
            pt = None
        if pt is not None:
            return pt
    try:
        return _point_accepted_on_face(planar_face, planar_face.Origin)
    except Exception:
        return None


def _origin_on_face(planar_face):
    """Origen del Spot: interior real de la cara (validado con Face.Project)."""
    return _face_interior_point(planar_face)


def _eid_key(eid):
    if eid is None:
        return None
    try:
        return int(eid.IntegerValue)
    except Exception:
        try:
            return int(eid.Value)
        except Exception:
            return None


def _pick_horizontal_face(elem, face_refs, prefer_lowest=False, prefer_highest=False):
    """
    Elige una cara horizontal con origen interior válido.
    ``prefer_lowest``: en fundaciones, la cara más baja (sello).
    """
    try:
        n_faces = int(face_refs.Count)
    except Exception:
        try:
            n_faces = len(face_refs)
        except Exception:
            n_faces = 0

    saw_sloped = False
    candidates = []
    for i in range(n_faces):
        try:
            face_ref = face_refs[i]
        except Exception:
            continue
        try:
            planar_face = elem.GetGeometryObjectFromReference(face_ref)
        except Exception:
            planar_face = None
        if planar_face is None or not isinstance(planar_face, PlanarFace):
            continue
        try:
            nz = abs(float(planar_face.FaceNormal.Z))
        except Exception:
            continue
        if nz < _HORIZONTAL_DOT_MIN:
            saw_sloped = True
            continue
        origin_pt = _origin_on_face(planar_face)
        if origin_pt is None:
            continue
        try:
            z = float(origin_pt.Z)
        except Exception:
            z = 0.0
        candidates.append((z, face_ref, planar_face, origin_pt))

    if candidates:
        if prefer_lowest:
            candidates.sort(key=lambda item: item[0])
        elif prefer_highest:
            candidates.sort(key=lambda item: item[0], reverse=True)
        _z, face_ref, planar_face, origin_pt = candidates[0]
        return face_ref, planar_face, origin_pt, u"ok"

    if saw_sloped:
        return None, None, None, u"slope"
    return None, None, None, u"other"


def _pick_horizontal_top_face(elem, top_faces):
    return _pick_horizontal_face(elem, top_faces, prefer_lowest=False)


def _annotation_past_dimension(origin, annotation_pt, view_right, view):
    """
    Desplaza el extremo del Spot más allá de la línea de cota, alejándolo del modelo.
    El offset es proporcional a la escala de la vista (calibrado a 1:50).
    """
    if annotation_pt is None:
        return None
    offset_mm = _mm_scaled_to_view(_SPOT_OFFSET_PAST_DIM_MM_AT_50, view)
    offset = _mm_to_internal(offset_mm)
    if origin is None:
        return annotation_pt - view_right * offset
    try:
        along = view_right.DotProduct(annotation_pt - origin)
    except Exception:
        along = 0.0
    if abs(along) > 1e-9:
        sign = 1.0 if along > 0 else -1.0
        return annotation_pt + view_right * (sign * offset)
    return annotation_pt - view_right * offset


def _leader_bend_end(origin, annotation_pt, view_right, view):
    """
    Líder largo: origen (interior losa) → bend → end (más allá de la cota).
    Shoulder y offsets de respaldo también siguen la escala de la vista.
    """
    shoulder = _mm_to_internal(
        _mm_scaled_to_view(_LEADER_SHOULDER_MM_AT_50, view)
    )
    min_len = _mm_to_internal(_MIN_LEADER_LEN_MM)
    offset_leader = _mm_to_internal(
        _mm_scaled_to_view(_OFFSET_LEADER_MM_AT_50, view)
    )
    offset_end = _mm_to_internal(_mm_scaled_to_view(_OFFSET_END_MM_AT_50, view))

    if origin is None:
        return None, None

    end = _annotation_past_dimension(origin, annotation_pt, view_right, view)

    if end is None:
        bend = origin - view_right * offset_leader
        end = bend - view_right * offset_end
        return bend, end

    try:
        dist = (end - origin).GetLength()
    except Exception:
        dist = 0.0

    if dist < min_len:
        bend = origin - view_right * offset_leader
        end = bend - view_right * offset_end
        return bend, end

    # Codo cerca del texto, hacia la losa (brazo largo origin→bend).
    along = view_right.DotProduct(origin - end)
    if abs(along) > min_len:
        sign = 1.0 if along > 0 else -1.0
        shoulder_use = min(shoulder, abs(along) * 0.35)
        if shoulder_use < min_len:
            shoulder_use = min(shoulder, abs(along) * 0.5)
        bend = end + view_right * (sign * shoulder_use)
    else:
        bend = XYZ(
            (origin.X + end.X) * 0.5,
            (origin.Y + end.Y) * 0.5,
            (origin.Z + end.Z) * 0.5,
        )
    return bend, end


def _try_new_spot_elevation(doc, view, face_ref, origin, bend, end, ref_pt):
    """Crea Spot con líder hacia la cota. Retorna el Spot o None."""
    candidates = [
        (origin, bend, end, ref_pt, True),
        (origin, bend, end, origin, True),
    ]
    if origin is not None and end is not None:
        mid = XYZ(
            (origin.X + end.X) * 0.5,
            (origin.Y + end.Y) * 0.5,
            (origin.Z + end.Z) * 0.5,
        )
        candidates.append((origin, mid, end, origin, True))

    for o, b, e, rp, hl in candidates:
        if o is None or b is None or e is None:
            continue
        try:
            spot = doc.Create.NewSpotElevation(view, face_ref, o, b, e, rp, hl)
            if spot is not None:
                return spot
        except Exception:
            continue
    return None


def _append_planar_face_ref(out, face, side=None):
    if face is None or not isinstance(face, PlanarFace):
        return
    try:
        nz = float(face.FaceNormal.Z)
    except Exception:
        return
    if abs(nz) < _HORIZONTAL_DOT_MIN:
        return
    if side == u"bottom" and nz >= 0:
        return
    if side == u"top" and nz <= 0:
        return
    try:
        href = face.Reference
    except Exception:
        href = None
    if href is not None:
        out.append(href)


def _walk_geometry_for_face_refs(geom, out, side=None):
    if geom is None:
        return
    for obj in geom:
        if obj is None:
            continue
        if isinstance(obj, Solid):
            try:
                faces = obj.Faces
            except Exception:
                faces = None
            if faces is None:
                continue
            for face in faces:
                _append_planar_face_ref(out, face, side=side)
            continue
        if isinstance(obj, GeometryInstance):
            try:
                inst_geom = obj.GetInstanceGeometry()
            except Exception:
                inst_geom = None
            _walk_geometry_for_face_refs(inst_geom, out, side=side)


def _is_foundation(elem):
    return _category_id_int(elem) == _CAT_FOUNDATIONS


def _is_framing(elem):
    return _category_id_int(elem) == _CAT_FRAMING


def _framing_material_text(elem):
    parts = []
    try:
        from Autodesk.Revit.DB.Structure import StructuralMaterialType

        sm = elem.StructuralMaterialType
        if sm == StructuralMaterialType.Steel:
            return u"steel"
        if sm in (
            StructuralMaterialType.Concrete,
            getattr(StructuralMaterialType, u"PrecastConcrete", StructuralMaterialType.Concrete),
        ):
            return u"concrete"
    except Exception:
        pass
    try:
        p = elem.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)
        if p is not None and p.HasValue:
            try:
                parts.append(_as_unicode(p.AsValueString() or u""))
            except Exception:
                pass
            try:
                parts.append(_as_unicode(p.AsString() or u""))
            except Exception:
                pass
    except Exception:
        pass
    blob = _canon_type_name(u" ".join(parts))
    if u"acero" in blob or u"steel" in blob:
        return u"steel"
    if u"hormigon" in blob or u"concreto" in blob or u"concrete" in blob:
        return u"concrete"
    return u""


def _floor_type_name(elem):
    if elem is None:
        return u""
    try:
        ft = elem.FloorType
        if ft is not None:
            n = _spot_type_display_name(ft)
            if n:
                return n
    except Exception:
        pass
    try:
        tid = elem.GetTypeId()
        if tid is None or tid == ElementId.InvalidElementId:
            return u""
        et = elem.Document.GetElement(tid)
        return _spot_type_display_name(et)
    except Exception:
        return u""


def _floor_is_radier(elem):
    return u"radier" in _canon_type_name(_floor_type_name(elem))


def _spot_type_name_for_elem(elem, cat_id=None):
    """
    Tipo de Spot Elevation según categoría:

    - Wall → Survey Point_Nivel Tope de Concreto
    - Structural Foundation → Survey Point_Nivel Sello de Fundacion
    - Structural Framing (hormigón) → Survey Point_Nivel Tope de Concreto
    - Structural Framing (acero) → Survey Point_Nivel Tope de Acero
    - Floor (tipo con Radier) → Survey Point_Nivel Tope de Radier
    - Floor (resto) → Survey Point_Nivel Tope de Losa
    """
    cid = cat_id
    if cid is None:
        cid = _category_id_int(elem)
    try:
        cid = int(cid)
    except Exception:
        cid = None
    if cid == _CAT_WALLS:
        return _SPOT_WALL
    if cid == _CAT_FOUNDATIONS:
        return _SPOT_FOUNDATION
    if cid == _CAT_FLOORS:
        if _floor_is_radier(elem):
            return _SPOT_RADIER
        return _SPOT_FLOOR
    if cid == _CAT_FRAMING:
        if _framing_material_text(elem) == u"steel":
            return _SPOT_FRAMING_STEEL
        return _SPOT_FRAMING_CONCRETE
    return CATEGORY_SPOT_MAPPING.get(cid)


def _host_face_refs(elem, want_bottom=None):
    """
    Losas: cara superior. Fundaciones: cara inferior (sello) por defecto.
    ``want_bottom`` fuerza inferior o superior.
    Si HostObjectUtils no aplica, recorre sólidos de la instancia.
    """
    if want_bottom is None:
        want_bottom = _is_foundation(elem)
    host_faces = None
    try:
        if want_bottom:
            host_faces = HostObjectUtils.GetBottomFaces(elem)
        else:
            host_faces = HostObjectUtils.GetTopFaces(elem)
    except Exception:
        host_faces = None
    if host_faces:
        try:
            n_faces = int(host_faces.Count)
        except Exception:
            try:
                n_faces = len(host_faces)
            except Exception:
                n_faces = 0
        if n_faces > 0:
            return host_faces, want_bottom

    out = []
    try:
        opts = Options()
        opts.ComputeReferences = True
        opts.IncludeNonVisibleObjects = False
        geom = elem.get_Geometry(opts)
    except Exception:
        geom = None
    side = u"bottom" if want_bottom else u"top"
    _walk_geometry_for_face_refs(geom, out, side=side)
    return out, want_bottom


def _foundation_top_face(elem):
    """Cara superior horizontal de la fundación (la más alta si hay varias)."""
    host_faces, _wb = _host_face_refs(elem, want_bottom=False)
    return _pick_horizontal_face(
        elem, host_faces, prefer_lowest=False, prefer_highest=True
    )


def _offset_dim_point_toward_model(pt, toward_pt, view_right, view):
    """Desplaza el origen de la línea de cota hacia el modelo (350 mm @ 1:50)."""
    offset = _mm_to_internal(
        _mm_scaled_to_view(_SECOND_DIM_OFFSET_MM_AT_50, view)
    )
    try:
        along = float(view_right.DotProduct(toward_pt - pt))
    except Exception:
        along = 0.0
    sign = 1.0 if along >= 0.0 else -1.0
    return pt + view_right * (sign * offset)


def _nearest_main_face(spots_data, origin_pt, skip_elem_id):
    """Referencia de la cota principal más cercana en Z a la cara superior."""
    skip = _eid_key(skip_elem_id)
    try:
        z0 = float(origin_pt.Z)
    except Exception:
        return None
    best = None
    best_dz = 1e18
    for face_ref, other_origin, _ann, _cat, elem_id in spots_data:
        if _eid_key(elem_id) == skip:
            continue
        if other_origin is None:
            continue
        try:
            dz = abs(float(other_origin.Z) - z0)
        except Exception:
            continue
        if dz < best_dz:
            best_dz = dz
            best = (face_ref, other_origin)
    return best


def _collect_floor_spot_data(doc, sel_refs, pt, view_up):
    """
    Filtra losas / fundaciones horizontales y prepara referencias + spots.
    Cada item: (face_ref, origin_on_face, annotation_pt, cat_id, elem_id).
    ``foundation_chains``: datos extra para la segunda cota de fundación.
    """
    ref_array = ReferenceArray()
    spots_data = []
    foundation_chains = []
    discarded_slope = 0
    discarded_other = 0

    for ref in sel_refs:
        elem = doc.GetElement(ref.ElementId)
        if elem is None or not _is_elevation_host(elem):
            discarded_other += 1
            continue
        host_faces, prefer_lowest = _host_face_refs(elem)
        try:
            n_faces = int(host_faces.Count)
        except Exception:
            try:
                n_faces = len(host_faces)
            except Exception:
                n_faces = 0
        if n_faces < 1:
            discarded_other += 1
            continue

        face_ref, planar_face, origin_pt, status = _pick_horizontal_face(
            elem,
            host_faces,
            prefer_lowest=prefer_lowest,
            prefer_highest=not prefer_lowest,
        )
        if status == u"slope":
            discarded_slope += 1
            continue
        if face_ref is None or planar_face is None or origin_pt is None:
            discarded_other += 1
            continue

        annotation_pt = _annotation_on_face_plane(planar_face, pt, view_up)
        ref_array.Append(face_ref)
        cat_id = elem.Category.Id.IntegerValue
        spots_data.append((face_ref, origin_pt, annotation_pt, cat_id, elem.Id))

        if _is_foundation(elem):
            top_ref, top_face, top_origin, top_status = _foundation_top_face(elem)
            if top_status == u"ok" and top_ref is not None and top_origin is not None:
                foundation_chains.append(
                    {
                        u"elem_id": elem.Id,
                        u"bottom_ref": face_ref,
                        u"bottom_origin": origin_pt,
                        u"top_ref": top_ref,
                        u"top_origin": top_origin,
                    }
                )

    return ref_array, spots_data, foundation_chains, discarded_slope, discarded_other


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


def run_floor_cota_and_spots(
    uidoc,
    aviso_fn,
    use_preselection=True,
    txn_group=None,
    txn_sketch=None,
    txn_create=None,
    show_success_dialog=True,
):
    """
    Flujo de losas: selección + PickPoint + cota alineada + Spot Elevations.

    ``aviso_fn(instruction, content=u"")``.
    Returns: (ok, status_text)
    """
    if uidoc is None:
        aviso_fn(u"No hay documento activo.")
        return False, u"Sin documento activo."

    doc = uidoc.Document
    active_view = None
    try:
        active_view = getattr(uidoc, "ActiveGraphicalView", None)
    except Exception:
        active_view = None
    if active_view is None:
        try:
            active_view = uidoc.ActiveView
        except Exception:
            active_view = None
    if active_view is None:
        aviso_fn(u"No hay vista activa.")
        return False, u"Sin vista activa."

    try:
        if active_view.IsTemplate:
            aviso_fn(
                u"Vista incorrecta",
                u"No se puede ejecutar sobre una plantilla de vista.",
            )
            return False, u"Vista plantilla."
    except Exception:
        pass

    if not is_section_or_elevation_view(active_view):
        aviso_fn(
            u"Vista incorrecta",
            u"Este escenario opera en Sección o Alzado.",
        )
        return False, u"La vista no es Sección ni Alzado."

    sel_refs = None
    if use_preselection:
        floors = preselected_floors(uidoc)
        if len(floors) >= 2:

            class _ElemRef(object):
                def __init__(self, eid):
                    self.ElementId = eid

            sel_refs = [_ElemRef(fl.Id) for fl in floors]

    if sel_refs is None or len(sel_refs) < 2:
        try:
            sel_refs = uidoc.Selection.PickObjects(
                ObjectType.Element,
                FloorSelectionFilter(),
                u"Seleccione losas, fundaciones, muros o Structural Framing "
                u"(Finish para confirmar, Esc para cancelar).",
            )
        except OperationCanceledException:
            return False, u"Selección cancelada."
        except Exception as ex:
            aviso_fn(u"No se pudo completar la selección.", _as_unicode(ex))
            return False, u"Error al seleccionar elementos."

    if sel_refs is None or len(sel_refs) < 2:
        aviso_fn(
            u"Selección insuficiente",
            u"Seleccione al menos dos losas, fundaciones, muros o Structural Framing "
            u"para crear la cota alineada.",
        )
        return False, u"Hacen falta al menos dos elementos."

    name_group = txn_group or _TXN_GROUP
    name_sketch = txn_sketch or _TXN_SKETCH
    name_create = txn_create or _TXN_CREATE

    tg = TransactionGroup(doc, name_group)
    tg.Start()
    try:
        if not _ensure_sketch_plane(doc, active_view, txn_name=name_sketch):
            tg.RollBack()
            aviso_fn(
                u"No se pudo configurar el plano de trabajo (SketchPlane) en la vista.",
            )
            return False, u"Sin plano de trabajo."

        try:
            pt = uidoc.Selection.PickPoint(
                u"Haga clic para definir la posición de la cota."
            )
        except OperationCanceledException:
            tg.RollBack()
            return False, u"Punto de cota cancelado."
        except Exception as ex:
            tg.RollBack()
            aviso_fn(u"No se pudo obtener el punto de cota.", _as_unicode(ex))
            return False, u"Error al indicar el punto."

        view_up = active_view.UpDirection
        view_right = active_view.RightDirection
        (
            ref_array,
            spots_data,
            foundation_chains,
            discarded_slope,
            discarded_other,
        ) = _collect_floor_spot_data(doc, sel_refs, pt, view_up)

        if ref_array.Size < 2:
            tg.RollBack()
            parts = [
                u"No hay suficientes caras superiores horizontales para crear "
                u"la cota (se requieren al menos 2 losas o fundaciones planas)."
            ]
            if discarded_slope:
                parts.append(
                    u"Descartadas por pendiente: {}.".format(discarded_slope)
                )
            if discarded_other:
                parts.append(
                    u"Descartadas por geometría inválida: {}.".format(discarded_other)
                )
            aviso_fn(u"No se pudo crear la cota", u"\n".join(parts))
            return False, u"Caras horizontales insuficientes."

        spot_types_cache = {}
        spot_registry = _build_spot_type_registry(doc)
        missing_types = set()
        created_ids = List[ElementId]()
        spot_ok = 0
        spot_fail = 0
        second_ok = 0
        hide_warnings = []

        dim_len = _mm_to_internal(_DIM_LINE_LENGTH_MM)

        t = Transaction(doc, name_create)
        t.Start()
        try:
            dim_line = Line.CreateBound(pt, pt + view_up * dim_len)
            new_dim = doc.Create.NewDimension(active_view, dim_line, ref_array)
            if new_dim is None:
                raise Exception(u"NewDimension devolvió None.")
            created_ids.Add(new_dim.Id)

            for chain in foundation_chains or []:
                top_ref = chain.get(u"top_ref")
                bot_ref = chain.get(u"bottom_ref")
                top_origin = chain.get(u"top_origin")
                if top_ref is None or bot_ref is None or top_origin is None:
                    continue
                nearest = _nearest_main_face(
                    spots_data, top_origin, chain.get(u"elem_id")
                )
                ra2 = ReferenceArray()
                if nearest is not None:
                    ra2.Append(nearest[0])
                ra2.Append(top_ref)
                ra2.Append(bot_ref)
                if ra2.Size < 2:
                    continue
                pt2 = _offset_dim_point_toward_model(
                    pt, top_origin, view_right, active_view
                )
                try:
                    dim_line2 = Line.CreateBound(pt2, pt2 + view_up * dim_len)
                    dim2 = doc.Create.NewDimension(active_view, dim_line2, ra2)
                except Exception:
                    dim2 = None
                if dim2 is None:
                    continue
                created_ids.Add(dim2.Id)
                second_ok += 1

            for face_ref, origin_pt, annotation_pt, cat_id, _elem_id in spots_data:
                host_elem = None
                try:
                    host_elem = doc.GetElement(_elem_id)
                except Exception:
                    host_elem = None
                bend, end = _leader_bend_end(
                    origin_pt, annotation_pt, view_right, active_view
                )
                new_spot = _try_new_spot_elevation(
                    doc,
                    active_view,
                    face_ref,
                    origin_pt,
                    bend,
                    end,
                    origin_pt,
                )
                if new_spot is None:
                    spot_fail += 1
                    continue
                created_ids.Add(new_spot.Id)
                spot_ok += 1

                target_type_name = _spot_type_name_for_elem(host_elem, cat_id)
                if not target_type_name:
                    missing_types.add(u"(sin regla de tipo para la categoría)")
                    continue
                target_type = _get_spot_type_by_name(
                    doc, target_type_name, registry=spot_registry, cache=spot_types_cache
                )
                if target_type is None:
                    missing_types.add(target_type_name)
                    continue
                if not _apply_spot_type(new_spot, target_type):
                    missing_types.add(target_type_name)

            if created_ids.Count > 0:
                views_hide = _views_to_hide_created_elements(doc, active_view)
                _hide_in_views(views_hide, created_ids, hide_warnings)

            t.Commit()
        except Exception as ex:
            try:
                if t.HasStarted():
                    t.RollBack()
            except Exception:
                pass
            tg.RollBack()
            aviso_fn(
                u"Error al crear cota o Spot Elevations.",
                _as_unicode(ex),
            )
            return False, u"Error al crear cota o spots."

        tg.Assimilate()

        summary = [
            u"Cota alineada principal: 1.",
            u"Cotas de espesor de fundación: {}.".format(second_ok),
            u"Spot Elevations creados: {}.".format(spot_ok),
        ]
        if spot_fail:
            summary.append(u"Spot Elevations fallidos: {}.".format(spot_fail))
        if discarded_slope:
            summary.append(
                u"Elementos descartados por pendiente: {}.".format(discarded_slope)
            )
        if discarded_other:
            summary.append(
                u"Elementos descartados por geometría inválida: {}.".format(
                    discarded_other
                )
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

        status = u"Cota principal + {0} cota(s) de fundación + {1} spot(s).".format(
            second_ok, spot_ok
        )
        if missing_types:
            status = status + u" Tipos no aplicados: {0}.".format(
                u", ".join(sorted(missing_types))
            )
        if show_success_dialog:
            aviso_fn(u"Operación completada", u"\n".join(summary))
        return True, status

    except Exception as ex:
        try:
            if tg.HasStarted():
                tg.RollBack()
        except Exception:
            pass
        aviso_fn(u"Error inesperado", _as_unicode(ex))
        return False, u"Error inesperado."


def create_smart_dimensions_and_spots(uiapp):
    uidoc = uiapp.ActiveUIDocument if uiapp is not None else None

    def _aviso(instruction, content=u""):
        mostrar_aviso(uiapp, instruction, content)

    run_floor_cota_and_spots(
        uidoc,
        _aviso,
        use_preselection=True,
        show_success_dialog=True,
    )


def run(uiapp):
    create_smart_dimensions_and_spots(uiapp)
