# -*- coding: utf-8 -*-
"""
Arainco: Spots by Dimension — Spot Elevations alineados desde cota lineal.

Selecciona una cota lineal en la vista activa y crea Spot Elevations en cada
referencia de la cota (excepto ejes, planos de referencia y líneas de modelo),
alineados
con la cota y con leader paramétrico según la escala de la vista.

Revit 2024–2026 | pyRevit / IronPython.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    Curve,
    DetailCurve,
    Dimension,
    DimensionStyleType,
    Edge,
    ElementId,
    Face,
    FilteredElementCollector,
    Grid,
    HostObjectUtils,
    PlanarFace,
    Point,
    ReferencePlane,
    SpotDimensionType,
    Transaction,
    ViewType,
    XYZ,
)
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

_TOOL_TITLE = u"Arainco: Spots by Dimension"
_TXN_NAME = u"Arainco: Spots by Dimension"
_PICK_PROMPT = u"Selecciona una cota lineal"

_MM_TO_FT = 1.0 / 304.8
_GAP_FROM_DIM_MM = 3.0
_SHOULDER_LEN_MM = 6.0
_TEXT_LEN_MM = 1.0
_LATERAL_MIN_LEN_FT = 0.001

_SPOT_WALL = u"Survey Point_Nivel Tope de Concreto"
_SPOT_FRAMING_CONCRETE = u"Survey Point_Nivel Tope de Concreto"
_SPOT_FRAMING_STEEL = u"Survey Point_Nivel Tope de Acero"
_SPOT_FOUNDATION = u"Survey Point_Nivel Sello de Fundacion"


def _ref_stable_repr(doc, ref):
    try:
        return _as_unicode(ref.ConvertToStableRepresentation(doc))
    except Exception:
        return u"(sin stable rep)"


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


class _DimensionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, Dimension)

    def AllowReference(self, ref, pos):
        return False


def _view_accepts_spot_dimensions(view):
    try:
        vt = view.ViewType
    except Exception:
        return True, u""
    if vt in (
        ViewType.DrawingSheet,
        ViewType.Schedule,
        ViewType.Legend,
        ViewType.Rendering,
    ):
        return False, (
            u"La vista activa no puede mostrar Spot Elevation (hoja, leyenda, etc.). "
            u"Abre una planta, alzado, sección o 3D y vuelve a ejecutar la herramienta."
        )
    return True, u""


def _iter_dimension_references(dim):
    refs = []
    try:
        ra = dim.References
        if ra is None:
            return refs
        it = ra.GetEnumerator()
        while it.MoveNext():
            refs.append(it.Current)
    except Exception:
        pass
    return refs


def _dimension_curve(dim):
    dim_curve = dim.Curve
    if dim_curve is None:
        try:
            segs = dim.Segments
            if segs is not None and segs.Size > 0:
                dim_curve = segs.Item[0].Curve
        except Exception:
            pass
    return dim_curve


def _dimension_axis_points(dim_curve):
    if dim_curve.IsBound:
        dim_origin_pt = dim_curve.GetEndPoint(0)
        dim_end_pt = dim_curve.GetEndPoint(1)
        dim_mid_pt = dim_curve.Evaluate(0.5, True)
    else:
        dim_origin_pt = dim_curve.Origin
        dim_end_pt = None
        dim_mid_pt = dim_curve.Origin
    dim_dir = dim_curve.Direction
    return dim_origin_pt, dim_mid_pt, dim_dir, dim_end_pt


def _point_on_dimension_near(pt, dim_origin, dim_dir, dim_end=None):
    """Punto en la cota (segmento o infinita) más cercano a ``pt``."""
    if pt is None or dim_origin is None or dim_dir is None:
        return dim_origin
    t = (pt - dim_origin).DotProduct(dim_dir)
    if dim_end is not None:
        t_end = (dim_end - dim_origin).DotProduct(dim_dir)
        t_min = min(0.0, t_end)
        t_max = max(0.0, t_end)
        t = max(t_min, min(t_max, t))
    return dim_origin + dim_dir * t


def _distance_xyz(a, b):
    try:
        return a.DistanceTo(b)
    except Exception:
        return None


def _closest_point_on_curve_to_target(curve, target):
    if curve is None or target is None:
        return None
    try:
        proj = curve.Project(target)
        if proj is not None:
            return proj.XYZPoint
    except Exception:
        pass
    if not curve.IsBound:
        return None
    best = None
    best_d = None
    for param in (0.0, 0.5, 1.0):
        try:
            candidate = curve.Evaluate(param, True)
            d = _distance_xyz(candidate, target)
            if d is not None and (best_d is None or d < best_d):
                best_d = d
                best = candidate
        except Exception:
            continue
    return best


def _closest_point_on_face_to_target(face, target):
    if face is None or target is None:
        return None
    snapped = _project_point_on_face(face, target)
    if snapped is not None:
        return snapped
    try:
        from Autodesk.Revit.DB import UV

        bbox = face.GetBoundingBox()
        candidates = []
        for u in (bbox.Min.U, bbox.Max.U):
            for v in (bbox.Min.V, bbox.Max.V):
                try:
                    candidates.append(face.Evaluate(UV(u, v)))
                except Exception:
                    pass
        uv_mid = UV(
            (bbox.Min.U + bbox.Max.U) * 0.5,
            (bbox.Min.V + bbox.Max.V) * 0.5,
        )
        try:
            candidates.append(face.Evaluate(uv_mid))
        except Exception:
            pass
        best = None
        best_d = None
        for candidate in candidates:
            d = _distance_xyz(candidate, target)
            if d is not None and (best_d is None or d < best_d):
                best_d = d
                best = candidate
        if best is not None:
            return _project_point_on_face(face, best)
    except Exception:
        pass
    return None


def _project_point_on_face(face, target):
    """Punto exactamente sobre la cara (UV de Project → Evaluate)."""
    if face is None or target is None:
        return None
    try:
        res = face.Project(target)
        if res is None:
            return None
        try:
            return face.Evaluate(res.UVPoint)
        except Exception:
            return res.XYZPoint
    except Exception:
        return None


def _is_horizontal_top_face(face, dot_min=0.985):
    try:
        if isinstance(face, PlanarFace):
            return abs(float(face.FaceNormal.Z)) >= dot_min
    except Exception:
        pass
    return False


def _face_interior_point(planar_face):
    """Punto interior válido sobre PlanarFace (muestreo UV)."""
    if planar_face is None or not isinstance(planar_face, PlanarFace):
        return None
    try:
        from Autodesk.Revit.DB import UV

        bb = planar_face.GetBoundingBox()
        u0 = float(bb.Min.U)
        u1 = float(bb.Max.U)
        v0 = float(bb.Min.V)
        v1 = float(bb.Max.V)
        if abs(u1 - u0) < 1e-12 or abs(v1 - v0) < 1e-12:
            return None
        for fu, fv in (
            (0.5, 0.5),
            (0.33, 0.33),
            (0.67, 0.67),
            (0.25, 0.75),
            (0.75, 0.25),
        ):
            try:
                uv = UV(u0 + (u1 - u0) * fu, v0 + (v1 - v0) * fv)
                pt = planar_face.Evaluate(uv)
            except Exception:
                continue
            snapped = _project_point_on_face(planar_face, pt)
            if snapped is not None:
                return snapped
    except Exception:
        pass
    try:
        return _project_point_on_face(planar_face, planar_face.Origin)
    except Exception:
        return None


def _origin_on_face_reference(elem, face_ref, near_pt):
    """
    Origen del Spot sobre la misma referencia de cara que recibirá NewSpotElevation.
    """
    face = None
    try:
        face = elem.GetGeometryObjectFromReference(face_ref)
    except Exception:
        return None

    origin = _project_point_on_face(face, near_pt)
    if origin is not None:
        return origin

    if isinstance(face, PlanarFace):
        interior = _face_interior_point(face)
        if interior is not None:
            origin = _project_point_on_face(face, near_pt)
            if origin is None:
                origin = interior
            return origin

    return None


def _rough_geometry_point(geom_obj, dim_origin_pt, dim_mid_pt):
    """Estimación inicial sobre la geometría referenciada."""
    for probe in (dim_mid_pt, dim_origin_pt):
        if probe is None:
            continue
        if hasattr(geom_obj, "Project"):
            try:
                res = geom_obj.Project(probe)
                if res is not None:
                    return res.XYZPoint
            except Exception:
                pass

    if isinstance(geom_obj, Edge):
        try:
            return _closest_point_on_curve_to_target(geom_obj.AsCurve(), dim_mid_pt)
        except Exception:
            pass

    if isinstance(geom_obj, Curve):
        return _closest_point_on_curve_to_target(geom_obj, dim_mid_pt)

    if isinstance(geom_obj, Face):
        return _closest_point_on_face_to_target(geom_obj, dim_mid_pt)

    if isinstance(geom_obj, Point):
        return geom_obj.Coord

    return None


def _resolve_origin(geom_obj, dim_origin_pt, dim_dir, dim_mid_pt, dim_end_pt):
    """
    Punto de anclaje del Spot sobre la geometría referenciada, lo más cercano posible
    a la cota seleccionada (proyección sobre el segmento de la cota).
    """
    rough = _rough_geometry_point(geom_obj, dim_origin_pt, dim_mid_pt)
    if rough is None:
        return None

    pt_on_dim = _point_on_dimension_near(rough, dim_origin_pt, dim_dir, dim_end_pt)

    if hasattr(geom_obj, "Project"):
        try:
            res = geom_obj.Project(pt_on_dim)
            if res is not None:
                return res.XYZPoint
        except Exception:
            pass

    if isinstance(geom_obj, Edge):
        try:
            pt = _closest_point_on_curve_to_target(geom_obj.AsCurve(), pt_on_dim)
            if pt is not None:
                return pt
        except Exception:
            pass

    if isinstance(geom_obj, Curve):
        pt = _closest_point_on_curve_to_target(geom_obj, pt_on_dim)
        if pt is not None:
            return pt

    if isinstance(geom_obj, Face):
        pt = _closest_point_on_face_to_target(geom_obj, pt_on_dim)
        if pt is not None:
            return pt

    if isinstance(geom_obj, Point):
        return geom_obj.Coord

    return rough


def _category_id_int(elem):
    try:
        if elem is None or elem.Category is None:
            return None
        return elem.Category.Id.IntegerValue
    except Exception:
        return None


def _is_detail_line(elem):
    return isinstance(elem, DetailCurve)


def _should_skip_reference(elem):
    if isinstance(elem, Grid) or isinstance(elem, ReferencePlane):
        return True
    cat = _category_id_int(elem)
    if cat == int(BuiltInCategory.OST_Lines) and not _is_detail_line(elem):
        return True
    return False


def _type_display_name(elem):
    if elem is None:
        return u""
    try:
        n = elem.Name
        if n is not None:
            s = _as_unicode(n).strip()
            if s:
                return s
    except Exception:
        pass
    for bip in (
        BuiltInParameter.ALL_MODEL_TYPE_NAME,
        BuiltInParameter.SYMBOL_NAME_PARAM,
    ):
        try:
            p = elem.get_Parameter(bip)
            if p is not None and p.HasValue:
                s = p.AsString()
                if s and s.strip():
                    return _as_unicode(s).strip()
        except Exception:
            continue
    try:
        pi = elem.GetType().GetProperty("Name")
        if pi is not None:
            v = pi.GetValue(elem, None)
            if v is not None:
                s = _as_unicode(v).strip()
                if s:
                    return s
    except Exception:
        pass
    return u""


def _element_type_name(doc, elem):
    try:
        from Autodesk.Revit.DB import Floor

        if isinstance(elem, Floor):
            try:
                ft = elem.FloorType
                if ft is not None:
                    n = _type_display_name(ft)
                    if n:
                        return n
            except Exception:
                pass
    except Exception:
        pass
    try:
        tid = elem.GetTypeId()
        if tid is None or tid == ElementId.InvalidElementId:
            return u""
        et = doc.GetElement(tid)
        if et is None:
            return u""
        return _type_display_name(et)
    except Exception:
        return u""


def _texto_indica_acero(text):
    if not text:
        return False
    t = _as_unicode(text).lower()
    for token in (u"acero", u"steel", u"metálic", u"metalic"):
        if token in t:
            return True
    return False


def _texto_indica_hormigon(text):
    if not text:
        return False
    t = _as_unicode(text).lower()
    for token in (u"hormig", u"concreto", u"concrete"):
        if token in t:
            return True
    return False


def _structural_material_type(elem):
    try:
        from Autodesk.Revit.DB.Structure import StructuralMaterialType

        return elem.StructuralMaterialType
    except Exception:
        return None


def _is_steel(elem):
    try:
        from Autodesk.Revit.DB.Structure import StructuralMaterialType

        sm = _structural_material_type(elem)
        if sm == StructuralMaterialType.Steel:
            return True
    except Exception:
        pass
    try:
        p = elem.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)
        if p is not None and p.HasValue:
            for attr in (u"AsValueString", u"AsString"):
                try:
                    val = getattr(p, attr)()
                    if _texto_indica_acero(val):
                        return True
                except Exception:
                    pass
    except Exception:
        pass
    return False


def _is_concrete(elem):
    try:
        from Autodesk.Revit.DB.Structure import StructuralMaterialType

        sm = _structural_material_type(elem)
        if sm == StructuralMaterialType.Concrete:
            return True
    except Exception:
        pass
    try:
        p = elem.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)
        if p is not None and p.HasValue:
            for attr in (u"AsValueString", u"AsString"):
                try:
                    val = getattr(p, attr)()
                    if _texto_indica_hormigon(val):
                        return True
                except Exception:
                    pass
    except Exception:
        pass
    return False


def _element_type_id_int(elem):
    try:
        tid = elem.GetTypeId()
        if tid is None:
            return None
        return tid.IntegerValue
    except Exception:
        return None


def _normalize_type_key(name):
    if not name:
        return u""
    t = _as_unicode(name).strip().lower()
    for src, dst in (
        (u"á", u"a"),
        (u"é", u"e"),
        (u"í", u"i"),
        (u"ó", u"o"),
        (u"ú", u"u"),
        (u"ñ", u"n"),
    ):
        t = t.replace(src, dst)
    return t


def _distinctive_type_key(name):
    n = _normalize_type_key(name)
    marker = u"survey point_"
    if marker in n:
        return n.split(marker, 1)[1]
    return n


def _build_spot_type_registry(doc):
    registry = []
    for spot_type in FilteredElementCollector(doc).OfClass(SpotDimensionType):
        try:
            style_ok = True
            try:
                style_ok = spot_type.StyleType == DimensionStyleType.SpotElevation
            except Exception:
                style_ok = True
            if not style_ok:
                continue
        except Exception:
            pass
        name = _type_display_name(spot_type)
        if not name:
            continue
        norm = _normalize_type_key(name)
        registry.append((name, norm, _distinctive_type_key(name), spot_type))
    return registry


def _get_spot_type_by_name(type_name, spot_types_cache, spot_registry):
    if not type_name:
        return None

    if type_name in spot_types_cache:
        return spot_types_cache[type_name]

    norm_target = _normalize_type_key(type_name)
    dist_target = _distinctive_type_key(type_name)

    exact = None
    normalized_map = {}
    distinctive_map = {}
    for name, norm, dist, spot_type in spot_registry:
        if name == type_name:
            exact = spot_type
            break
        normalized_map[norm] = spot_type
        distinctive_map[dist] = spot_type

    if exact is not None:
        spot_types_cache[type_name] = exact
        return exact

    hit = normalized_map.get(norm_target)
    if hit is not None:
        spot_types_cache[type_name] = hit
        return hit

    hit = distinctive_map.get(dist_target)
    if hit is not None:
        spot_types_cache[type_name] = hit
        return hit

    for name, norm, dist, spot_type in spot_registry:
        if dist_target and (dist_target in dist or dist in dist_target):
            spot_types_cache[type_name] = spot_type
            return spot_type

    spot_types_cache[type_name] = None
    return None


def _get_spot_type_by_keyword(keyword, spot_types_cache, spot_registry):
    """Spot Elevation cuyo nombre contiene la palabra clave (p. ej. radier, losa)."""
    if not keyword:
        return None

    cache_key = u"__keyword__:{0}".format(_normalize_type_key(keyword))
    if cache_key in spot_types_cache:
        return spot_types_cache[cache_key]

    kw = _normalize_type_key(keyword)
    candidates = []
    for name, norm, dist, spot_type in spot_registry:
        if kw in norm or kw in dist:
            candidates.append((name, norm, dist, spot_type))

    if not candidates:
        spot_types_cache[cache_key] = None
        return None

    for name, norm, dist, spot_type in candidates:
        if u"tope" in dist:
            spot_types_cache[cache_key] = spot_type
            return spot_type

    hit = candidates[0][3]
    spot_types_cache[cache_key] = hit
    return hit


def _floor_spot_keyword(floor_type_name):
    """Palabra clave según el nombre del tipo de Floor (solo tipo, no instancia)."""
    if not floor_type_name:
        return None
    tn = _normalize_type_key(floor_type_name)
    if u"radier" in tn:
        return u"radier"
    if u"losa" in tn:
        return u"losa"
    return None


def _resolve_spot_type_name(doc, elem):
    cat = _category_id_int(elem)
    if cat is None:
        return None

    if cat == int(BuiltInCategory.OST_Walls):
        return _SPOT_WALL

    if cat == int(BuiltInCategory.OST_StructuralFraming):
        if _is_steel(elem):
            return _SPOT_FRAMING_STEEL
        if _is_concrete(elem):
            return _SPOT_FRAMING_CONCRETE
        return None

    if cat == int(BuiltInCategory.OST_StructuralFoundation):
        return _SPOT_FOUNDATION

    return None


def _apply_spot_type(doc, spot, elem, spot_types_cache, spot_registry, missing_types):
    cat = _category_id_int(elem)
    if cat == int(BuiltInCategory.OST_Floors):
        floor_type_name = _element_type_name(doc, elem)
        keyword = _floor_spot_keyword(floor_type_name)
        if not keyword:
            return False

        target_type = _get_spot_type_by_keyword(
            keyword,
            spot_types_cache,
            spot_registry,
        )
        missing_label = u"Spot Elevation con «{0}» en el nombre".format(
            keyword.capitalize()
        )
        if target_type is None:
            missing_types.add(missing_label)
            return False

        try:
            spot.ChangeTypeId(target_type.Id)
        except Exception:
            missing_types.add(missing_label)
            return False

        if _element_type_id_int(spot) != target_type.Id.IntegerValue:
            missing_types.add(missing_label)
            return False

        return True

    target_type_name = _resolve_spot_type_name(doc, elem)
    if not target_type_name:
        return False

    target_type = _get_spot_type_by_name(
        target_type_name,
        spot_types_cache,
        spot_registry,
    )
    if target_type is None:
        missing_types.add(target_type_name)
        return False

    try:
        spot.ChangeTypeId(target_type.Id)
    except Exception:
        missing_types.add(target_type_name)
        return False

    if _element_type_id_int(spot) != target_type.Id.IntegerValue:
        missing_types.add(target_type_name)
        return False

    return True


def _element_guess_near_dimension(elem, dim_ref, dim_origin_pt, dim_mid_pt):
    """Punto aproximado del elemento para ubicarlo a lo largo de la cota."""
    geom_obj = None
    try:
        geom_obj = elem.GetGeometryObjectFromReference(dim_ref)
    except Exception:
        pass
    if geom_obj is not None:
        rough = _rough_geometry_point(geom_obj, dim_origin_pt, dim_mid_pt)
        if rough is not None:
            return rough

    try:
        bb = elem.get_BoundingBox(None)
        if bb is not None:
            return XYZ(
                (bb.Min.X + bb.Max.X) * 0.5,
                (bb.Min.Y + bb.Max.Y) * 0.5,
                bb.Max.Z,
            )
    except Exception:
        pass

    return dim_mid_pt


def _host_top_face_spot_candidates(
    elem,
    dim_ref,
    dim_origin_pt,
    dim_dir,
    dim_mid_pt,
    dim_end_pt,
):
    """
    Host (Floor, fundación, etc.): caras superiores vía GetTopFaces.
    Devuelve candidatos (referencia, origen) ordenados por cercanía a la cota.
    """
    guess = _element_guess_near_dimension(elem, dim_ref, dim_origin_pt, dim_mid_pt)
    pt_on_dim = _point_on_dimension_near(guess, dim_origin_pt, dim_dir, dim_end_pt)

    ranked = []
    seen_refs = set()

    def _add_candidate(face_ref, near_pt):
        try:
            doc = elem.Document
            key = _ref_stable_repr(doc, face_ref)
        except Exception:
            key = unicode(face_ref)
        if key in seen_refs:
            return
        origin = _origin_on_face_reference(elem, face_ref, near_pt)
        if origin is None:
            return
        d = _distance_xyz(origin, pt_on_dim)
        seen_refs.add(key)
        ranked.append((d if d is not None else 1e9, face_ref, origin))

    geom_obj = None
    try:
        geom_obj = elem.GetGeometryObjectFromReference(dim_ref)
    except Exception:
        pass

    if isinstance(geom_obj, Face) and _is_horizontal_top_face(geom_obj):
        _add_candidate(dim_ref, pt_on_dim)

    top_faces = None
    try:
        top_faces = HostObjectUtils.GetTopFaces(elem)
    except Exception:
        pass

    if top_faces is not None:
        try:
            n_faces = int(top_faces.Count)
        except Exception:
            try:
                n_faces = len(top_faces)
            except Exception:
                n_faces = 0

        for i in range(n_faces):
            try:
                face_ref = top_faces[i]
            except Exception:
                continue
            _add_candidate(face_ref, pt_on_dim)

    ranked.sort(key=lambda item: item[0])
    return [(face_ref, origin) for _, face_ref, origin in ranked]


def _uses_host_top_faces(elem):
    cat = _category_id_int(elem)
    if cat in (
        int(BuiltInCategory.OST_Floors),
        int(BuiltInCategory.OST_StructuralFoundation),
    ):
        return True
    return False


def _spot_reference_candidates(
    elem,
    dim_ref,
    dim_origin_pt,
    dim_dir,
    dim_mid_pt,
    dim_end_pt,
):
    if _uses_host_top_faces(elem):
        return _host_top_face_spot_candidates(
            elem,
            dim_ref,
            dim_origin_pt,
            dim_dir,
            dim_mid_pt,
            dim_end_pt,
        )

    geom_obj = None
    try:
        geom_obj = elem.GetGeometryObjectFromReference(dim_ref)
    except Exception:
        pass
    if geom_obj is None:
        return []

    origin = _resolve_origin(
        geom_obj,
        dim_origin_pt,
        dim_dir,
        dim_mid_pt,
        dim_end_pt,
    )
    if origin is None:
        return []
    return [(dim_ref, origin)]


def _compute_spot_points(origin, dim_origin_pt, dim_dir, view_normal, scale):
    vec_to_origin = origin - dim_origin_pt
    dist_along_dim = vec_to_origin.DotProduct(dim_dir)
    pt_on_dim = dim_origin_pt + (dim_dir * dist_along_dim)

    depth_diff = (pt_on_dim - origin).DotProduct(view_normal)
    pt_on_dim_aligned = pt_on_dim - (view_normal * depth_diff)

    vec_outward = pt_on_dim_aligned - origin
    if vec_outward.GetLength() > _LATERAL_MIN_LEN_FT:
        lateral_dir = vec_outward.Normalize()
    else:
        lateral_dir = dim_dir.CrossProduct(view_normal)
        if lateral_dir.GetLength() > _LATERAL_MIN_LEN_FT:
            lateral_dir = lateral_dir.Normalize()
        else:
            lateral_dir = dim_dir

    gap_from_dim = _GAP_FROM_DIM_MM * scale * _MM_TO_FT
    shoulder_len = _SHOULDER_LEN_MM * scale * _MM_TO_FT
    text_len = _TEXT_LEN_MM * scale * _MM_TO_FT

    bend = pt_on_dim_aligned + (lateral_dir * gap_from_dim)
    end = bend + (lateral_dir * shoulder_len)
    text = end + (lateral_dir * text_len)
    return bend, end, text


def _try_new_spot_elevation(doc, view, face_ref, origin, bend, end, ref_pt):
    if origin is None:
        return None

    candidates = [
        (origin, bend, end, ref_pt, True),
        (origin, bend, end, origin, True),
        (origin, origin, end, origin, True),
        (origin, bend, end, bend, True),
    ]
    if end is not None:
        candidates.append((origin, end, end, end, True))
        candidates.append((origin, origin, origin, origin, False))
    if origin is not None and end is not None:
        mid = XYZ(
            (origin.X + end.X) * 0.5,
            (origin.Y + end.Y) * 0.5,
            (origin.Z + end.Z) * 0.5,
        )
        candidates.append((origin, mid, end, origin, True))

    seen = set()
    for o, b, e, rp, has_leader in candidates:
        if o is None or b is None or e is None or rp is None:
            continue
        key = (
            round(o.X, 4),
            round(o.Y, 4),
            round(o.Z, 4),
            round(b.X, 4),
            round(b.Y, 4),
            round(b.Z, 4),
            round(e.X, 4),
            round(e.Y, 4),
            round(e.Z, 4),
            has_leader,
        )
        if key in seen:
            continue
        seen.add(key)
        try:
            spot = doc.Create.NewSpotElevation(view, face_ref, o, b, e, rp, has_leader)
            if spot is not None:
                return spot
        except Exception:
            continue

    return None


def place_aligned_spot_elevations(doc, uidoc, uiapp=None):
    view = doc.ActiveView
    ok_view, view_msg = _view_accepts_spot_dimensions(view)
    if not ok_view:
        _mostrar_aviso(uiapp, view_msg)
        return

    try:
        sel_ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            _DimensionFilter(),
            _PICK_PROMPT,
        )
        dim = doc.GetElement(sel_ref.ElementId)
    except OperationCanceledException:
        return
    except Exception as ex:
        _mostrar_aviso(
            uiapp,
            u"No se pudo leer la cota seleccionada.",
            _as_unicode(ex),
        )
        return

    refs = _iter_dimension_references(dim)
    if not refs:
        _mostrar_aviso(
            uiapp,
            u"La cota seleccionada no tiene referencias utilizables.",
        )
        return

    dim_curve = _dimension_curve(dim)
    if dim_curve is None:
        _mostrar_aviso(
            uiapp,
            u"No se pudo obtener la curva de la cota seleccionada.",
        )
        return

    try:
        dim_origin_pt, dim_mid_pt, dim_dir, dim_end_pt = _dimension_axis_points(dim_curve)
    except Exception:
        _mostrar_aviso(
            uiapp,
            u"La curva de la cota no admite dirección para alinear los Spot Elevations.",
        )
        return

    view_normal = view.ViewDirection
    scale = view.Scale

    skipped = 0
    skipped_geom = 0
    failed = 0
    spots_creados = 0
    spot_types_cache = {}
    spot_registry = _build_spot_type_registry(doc)
    missing_types = set()

    t = Transaction(doc, _TXN_NAME)
    t.Start()
    try:
        for ref in refs:
            elem = doc.GetElement(ref.ElementId)
            if elem is None or _should_skip_reference(elem):
                skipped += 1
                continue

            is_detail_line = _is_detail_line(elem)

            candidates = _spot_reference_candidates(
                elem,
                ref,
                dim_origin_pt,
                dim_dir,
                dim_mid_pt,
                dim_end_pt,
            )
            if not candidates:
                skipped_geom += 1
                continue

            spot = None
            for spot_ref, origin in candidates:
                bend, end, text = _compute_spot_points(
                    origin,
                    dim_origin_pt,
                    dim_dir,
                    view_normal,
                    scale,
                )
                spot = _try_new_spot_elevation(
                    doc,
                    view,
                    spot_ref,
                    origin,
                    bend,
                    end,
                    text,
                )
                if spot is not None:
                    break

            if spot is not None:
                spots_creados += 1
                if not is_detail_line:
                    _apply_spot_type(
                        doc,
                        spot,
                        elem,
                        spot_types_cache,
                        spot_registry,
                        missing_types,
                    )
            else:
                failed += 1

        if spots_creados > 0:
            t.Commit()
        else:
            t.RollBack()
    except Exception as ex:
        t.RollBack()
        _mostrar_aviso(
            uiapp,
            u"Error al crear Spot Elevations.",
            _as_unicode(ex),
        )
        return

    if spots_creados == 0:
        detail = []
        if skipped:
            detail.append(u"Referencias omitidas: {}.".format(skipped))
        if skipped_geom:
            detail.append(u"Referencias sin geometría válida: {}.".format(skipped_geom))
        if failed:
            detail.append(u"Referencias no creadas por la API: {}.".format(failed))
        content = u"\n".join(detail)
        _mostrar_aviso(
            uiapp,
            u"No se creó ningún Spot Elevation.",
            content,
        )
        return

    summary = [u"Spot Elevations creados: {}.".format(spots_creados)]
    if failed:
        summary.append(u"No creados (API): {}.".format(failed))
    if skipped:
        summary.append(u"Omitidos: {}.".format(skipped))
    if skipped_geom:
        summary.append(u"Sin geometría/candidatos: {}.".format(skipped_geom))
    if missing_types:
        summary.append(
            u"Tipos no encontrados en el proyecto: {}.".format(
                u", ".join(sorted(_as_unicode(n) for n in missing_types))
            )
        )

    content = u"\n".join(summary[1:])
    _mostrar_aviso(uiapp, summary[0], content)


def run(revit):
    uiapp = revit
    try:
        uidoc = uiapp.ActiveUIDocument
    except Exception:
        uidoc = revit.ActiveUIDocument
    if uidoc is None:
        _mostrar_aviso(None, u"No hay documento activo.")
        return
    doc = uidoc.Document
    if doc is None:
        _mostrar_aviso(uiapp, u"No hay documento activo.")
        return
    place_aligned_spot_elevations(doc, uidoc, uiapp=uiapp)
