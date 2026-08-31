# -*- coding: utf-8 -*-
"""
Muros invertidos → Filled Region en Ceiling Plan (ViewDirection +Z).

Revit 2024+ | IronPython (pyRevit)

1. Muros en el recorte XY de la vista (Z libre) más los visibles.
2. Lote corte / sin corte según colisión con el Cut Plane.
3. Sección longitudinal XY (corte: Z_cut; sin corte: media altura del sólido).
4. Empareja huellas coincidentes (solape de área).
5. Región = sin corte − unión(cortes coincidentes).
"""

from __future__ import print_function

import math

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    Arc,
    BooleanOperationsType,
    BooleanOperationsUtils,
    BoundingBoxIntersectsFilter,
    BuiltInParameter,
    Element,
    ElementId,
    FilledRegion,
    FilledRegionType,
    FilteredElementCollector,
    GeometryCreationUtilities,
    GeometryInstance,
    Line,
    LocationCurve,
    Options,
    Outline,
    Plane,
    PlanarFace,
    PlanViewPlane,
    PlanViewRange,
    Solid,
    SolidUtils,
    Transaction,
    ViewDetailLevel,
    ViewFamily,
    ViewPlan,
    Wall,
    XYZ,
)
from bimtools_clr_collections import (
    iterate_net_collection,
    list_curve_from_iterable,
    list_curve_loop_from_iterable,
)

TXN_NAME = u"Arainco: Muros invertidos"
_SOLID_VOL_TOL = 1e-12
_MIN_LINE_FT = 1.0 / 304.8
_CUT_FACE_DISTS = (0.04, 0.08, 0.16)
_OVERLAP_RATIO = 0.10
_EXTRUDE_H_FT = 1.0
_Z_TOL_FT = 2.0 / 304.8


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _elem_name(el):
    if el is None:
        return u""
    try:
        n = el.Name
        if n:
            return _as_unicode(n)
    except Exception:
        pass
    try:
        return _as_unicode(Element.Name.__get__(el))
    except Exception:
        return u""


def is_ceiling_plan_looking_up(view):
    """Ceiling Plan y ViewDirection hacia +Z."""
    if view is None or not isinstance(view, ViewPlan):
        return False
    try:
        if float(view.ViewDirection.Z) <= 0.1:
            return False
    except Exception:
        return False
    try:
        if str(view.ViewType) == u"CeilingPlan":
            return True
    except Exception:
        pass
    try:
        vft = view.Document.GetElement(view.GetTypeId())
        if vft is not None and vft.ViewFamily == ViewFamily.CeilingPlan:
            return True
    except Exception:
        pass
    return False


def cut_plane_z(view):
    """Elevación del Cut Plane de una ViewPlan, o None."""
    if view is None or not isinstance(view, ViewPlan):
        return None
    doc = view.Document
    try:
        vr = view.GetViewRange()
        lid = vr.GetLevelId(PlanViewPlane.CutPlane)
        offset = float(vr.GetOffset(PlanViewPlane.CutPlane))
    except Exception:
        return None
    level = None
    try:
        unlim = PlanViewRange.Unlimited
    except Exception:
        unlim = None
    if (
        lid is not None
        and lid != ElementId.InvalidElementId
        and (unlim is None or lid != unlim)
        and doc is not None
    ):
        try:
            level = doc.GetElement(lid)
        except Exception:
            level = None
    if level is None:
        try:
            level = view.GenLevel
        except Exception:
            level = None
    if level is None:
        return None
    try:
        return float(level.Elevation) + offset
    except Exception:
        return None


def collect_walls_in_view(doc, view):
    if doc is None or view is None:
        return []
    try:
        col = (
            FilteredElementCollector(doc, view.Id)
            .OfClass(Wall)
            .WhereElementIsNotElementType()
        )
        return [w for w in col if isinstance(w, Wall)]
    except Exception:
        return []


def _view_crop_xy_outline(view):
    """Outline modelo: recorte XY de la vista y Z libre (incluye invertidos fuera del View Range)."""
    if view is None:
        return None
    try:
        cb = view.CropBox
    except Exception:
        cb = None
    if cb is None:
        return None
    try:
        tf = cb.Transform
        mn = cb.Min
        mx = cb.Max
    except Exception:
        return None
    corners = []
    for x in (float(mn.X), float(mx.X)):
        for y in (float(mn.Y), float(mx.Y)):
            for z in (float(mn.Z), float(mx.Z)):
                try:
                    corners.append(tf.OfPoint(XYZ(x, y, z)))
                except Exception:
                    continue
    if len(corners) < 2:
        return None
    xs = [float(p.X) for p in corners]
    ys = [float(p.Y) for p in corners]
    pad = 1.0
    minx, maxx = min(xs) - pad, max(xs) + pad
    miny, maxy = min(ys) - pad, max(ys) + pad
    if maxx <= minx or maxy <= miny:
        return None
    try:
        return Outline(
            XYZ(minx, miny, -10000.0),
            XYZ(maxx, maxy, 10000.0),
        )
    except Exception:
        return None


def _wall_id_int(wall):
    try:
        return int(wall.Id.IntegerValue)
    except Exception:
        try:
            return int(wall.Id.Value)
        except Exception:
            return None


def collect_walls_for_plan(doc, view):
    """
    Muros del recorte XY (Z ilimitada) más los visibles en la vista.

    El collector por ``view.Id`` omite invertidos enteros por encima del
    Top Clip / View Depth de la planta de techo.
    """
    by_id = {}
    for w in collect_walls_in_view(doc, view):
        wid = _wall_id_int(w)
        if wid is not None:
            by_id[wid] = w
    outline = _view_crop_xy_outline(view)
    if outline is not None and doc is not None:
        try:
            filt = BoundingBoxIntersectsFilter(outline)
            col = (
                FilteredElementCollector(doc)
                .OfClass(Wall)
                .WherePasses(filt)
                .WhereElementIsNotElementType()
            )
            for w in col:
                if not isinstance(w, Wall):
                    continue
                wid = _wall_id_int(w)
                if wid is not None:
                    by_id[wid] = w
        except Exception:
            pass
    if not by_id and doc is not None:
        try:
            col = (
                FilteredElementCollector(doc)
                .OfClass(Wall)
                .WhereElementIsNotElementType()
            )
            for w in col:
                if isinstance(w, Wall):
                    wid = _wall_id_int(w)
                    if wid is not None:
                        by_id[wid] = w
        except Exception:
            pass
    return list(by_id.values())


def list_filled_region_types(doc):
    if doc is None:
        return []
    out = []
    try:
        col = FilteredElementCollector(doc).OfClass(FilledRegionType)
        for t in col:
            if t is None:
                continue
            out.append(t)
    except Exception:
        return []
    out.sort(key=lambda t: _elem_name(t).lower())
    return out


def wall_solids(wall):
    if wall is None:
        return []
    out = []
    for detail in (ViewDetailLevel.Medium, ViewDetailLevel.Fine):
        opts = Options()
        opts.ComputeReferences = False
        try:
            opts.DetailLevel = detail
        except Exception:
            pass
        try:
            opts.IncludeNonVisibleObjects = True
        except Exception:
            pass
        try:
            geom = wall.get_Geometry(opts)
        except Exception:
            geom = None
        if geom is None:
            continue
        for obj in iterate_net_collection(geom):
            if obj is None:
                continue
            if isinstance(obj, Solid):
                try:
                    if float(obj.Volume) > _SOLID_VOL_TOL:
                        out.append(obj)
                except Exception:
                    pass
            elif isinstance(obj, GeometryInstance):
                inst = None
                try:
                    inst = obj.GetInstanceGeometry()
                except Exception:
                    inst = None
                if inst is None:
                    continue
                for g in iterate_net_collection(inst):
                    if isinstance(g, Solid):
                        try:
                            if float(g.Volume) > _SOLID_VOL_TOL:
                                out.append(g)
                        except Exception:
                            pass
        if out:
            return out
    return out


def _solids_vertex_z_range(solids):
    zmin = None
    zmax = None
    for s in solids or []:
        try:
            edges = s.Edges
        except Exception:
            edges = None
        if edges is None:
            continue
        for edge in iterate_net_collection(edges):
            try:
                crv = edge.AsCurve()
                pts = (crv.GetEndPoint(0), crv.GetEndPoint(1))
            except Exception:
                continue
            for p in pts:
                try:
                    z = float(p.Z)
                except Exception:
                    continue
                if zmin is None or z < zmin:
                    zmin = z
                if zmax is None or z > zmax:
                    zmax = z
    return zmin, zmax


def wall_model_z_range(wall, solids=None):
    """
    Elevación modelo del muro. Prioriza bbox del elemento (coords de modelo),
    no ``Solid.GetBoundingBox`` (puede ir en coords locales).
    """
    if wall is not None:
        try:
            bb = wall.get_BoundingBox(None)
        except Exception:
            bb = None
        if bb is not None:
            try:
                return float(bb.Min.Z), float(bb.Max.Z)
            except Exception:
                pass
    vz0, vz1 = _solids_vertex_z_range(solids)
    if vz0 is not None:
        return vz0, vz1
    if wall is None:
        return None, None
    try:
        base_id = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)
        base_off = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
        height = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)
        doc = wall.Document
        lid = base_id.AsElementId() if base_id is not None else None
        level = doc.GetElement(lid) if lid is not None else None
        z0 = float(level.Elevation) + float(base_off.AsDouble())
        z1 = z0 + float(height.AsDouble())
        return z0, z1
    except Exception:
        return None, None


def wall_location_loop(wall, z=0.0):
    """Huella rectangular (o anillo) desde LocationCurve × Width, en Z dada."""
    if wall is None:
        return None
    try:
        loc = wall.Location
        curve = loc.Curve if isinstance(loc, LocationCurve) else None
    except Exception:
        curve = None
    if curve is None:
        return None
    try:
        width = float(wall.Width)
    except Exception:
        width = 0.0
    if width < _MIN_LINE_FT:
        return None
    half = 0.5 * width
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
    except Exception:
        return None
    if isinstance(curve, Line):
        dx = float(p1.X) - float(p0.X)
        dy = float(p1.Y) - float(p0.Y)
        ln = math.hypot(dx, dy)
        if ln < _MIN_LINE_FT:
            return None
        nx, ny = -dy / ln * half, dx / ln * half
        a = XYZ(float(p0.X) + nx, float(p0.Y) + ny, z)
        b = XYZ(float(p1.X) + nx, float(p1.Y) + ny, z)
        c = XYZ(float(p1.X) - nx, float(p1.Y) - ny, z)
        d = XYZ(float(p0.X) - nx, float(p0.Y) - ny, z)
        try:
            lines = [
                Line.CreateBound(a, b),
                Line.CreateBound(b, c),
                Line.CreateBound(c, d),
                Line.CreateBound(d, a),
            ]
        except Exception:
            return None
        return CurveLoop_Create(list_curve_from_iterable(lines))
    tes = []
    try:
        tes = list(curve.Tessellate())
    except Exception:
        tes = []
    if len(tes) < 2:
        return None
    left = []
    right = []
    npts = len(tes)
    for i in range(npts):
        if i == 0:
            t0, t1 = tes[0], tes[1]
        elif i == npts - 1:
            t0, t1 = tes[i - 1], tes[i]
        else:
            t0, t1 = tes[i - 1], tes[i + 1]
        dx = float(t1.X) - float(t0.X)
        dy = float(t1.Y) - float(t0.Y)
        ln = math.hypot(dx, dy)
        if ln < 1e-12:
            continue
        nx, ny = -dy / ln * half, dx / ln * half
        p = tes[i]
        left.append(XYZ(float(p.X) + nx, float(p.Y) + ny, z))
        right.append(XYZ(float(p.X) - nx, float(p.Y) - ny, z))
    ring = left + list(reversed(right))
    lines = []
    for i in range(len(ring)):
        a = ring[i]
        b = ring[(i + 1) % len(ring)]
        if a.DistanceTo(b) < _MIN_LINE_FT:
            continue
        try:
            lines.append(Line.CreateBound(a, b))
        except Exception:
            continue
    if len(lines) < 3:
        return None
    return CurveLoop_Create(list_curve_from_iterable(lines))


def section_loops_for_wall(wall, solids, z):
    loops = section_loops_at_z(solids, z)
    if loops:
        return loops
    loc = wall_location_loop(wall, 0.0)
    if loc is not None:
        return [loc]
    return []


def _dist_point_to_plane(pt, plane):
    try:
        n = plane.Normal
        o = plane.Origin
        return abs(
            (float(pt.X) - float(o.X)) * float(n.X)
            + (float(pt.Y) - float(o.Y)) * float(n.Y)
            + (float(pt.Z) - float(o.Z)) * float(n.Z)
        )
    except Exception:
        return 1.0e9


def _largest_face_on_plane(solid, plane, tol):
    if solid is None or plane is None:
        return None
    try:
        pn = plane.Normal
        plen = math.sqrt(float(pn.X) ** 2 + float(pn.Y) ** 2 + float(pn.Z) ** 2)
        if plen < 1e-12:
            return None
        nx, ny, nz = float(pn.X) / plen, float(pn.Y) / plen, float(pn.Z) / plen
    except Exception:
        return None
    best = None
    best_a = -1.0
    try:
        faces = solid.Faces
    except Exception:
        return None
    for face in iterate_net_collection(faces):
        if not isinstance(face, PlanarFace):
            continue
        try:
            fn = face.FaceNormal
            fl = math.sqrt(float(fn.X) ** 2 + float(fn.Y) ** 2 + float(fn.Z) ** 2)
            if fl < 1e-12:
                continue
            fx, fy, fz = float(fn.X) / fl, float(fn.Y) / fl, float(fn.Z) / fl
            if abs(abs(fx * nx + fy * ny + fz * nz) - 1.0) > 0.02:
                continue
        except Exception:
            continue
        try:
            origin = face.Origin
        except Exception:
            origin = None
        if origin is None:
            continue
        if _dist_point_to_plane(origin, plane) > tol:
            continue
        try:
            a = float(face.Area)
        except Exception:
            a = 0.0
        if a > best_a:
            best_a = a
            best = face
    return best


def _cut_face_from_solid(solid, plane):
    if solid is None or plane is None:
        return None
    try:
        pn = plane.Normal
        origin = plane.Origin
    except Exception:
        return None
    for td in _CUT_FACE_DISTS:
        for flip in (False, True):
            try:
                if flip:
                    nn = XYZ(-float(pn.X), -float(pn.Y), -float(pn.Z))
                else:
                    nn = pn
                cut_pl = Plane.CreateByNormalAndOrigin(nn, origin)
            except Exception:
                continue
            try:
                s_cut = BooleanOperationsUtils.CutWithHalfSpace(solid, cut_pl)
            except Exception:
                s_cut = None
            if s_cut is None:
                continue
            try:
                if float(s_cut.Volume) <= _SOLID_VOL_TOL:
                    continue
            except Exception:
                pass
            face = _largest_face_on_plane(s_cut, cut_pl, td)
            if face is not None:
                return face
    return None


def _loops_from_face(face):
    if face is None:
        return []
    try:
        raw = face.GetEdgesAsCurveLoops()
    except Exception:
        return []
    return [lp for lp in iterate_net_collection(raw) if lp is not None]


def _section_pts_from_edges(solid, plane):
    pts = []
    if solid is None or plane is None:
        return pts
    try:
        edges = solid.Edges
    except Exception:
        return pts
    try:
        n = plane.Normal
        o = plane.Origin
    except Exception:
        return pts
    for edge in iterate_net_collection(edges):
        try:
            crv = edge.AsCurve()
        except Exception:
            crv = None
        if crv is None:
            continue
        try:
            p0 = crv.GetEndPoint(0)
            p1 = crv.GetEndPoint(1)
        except Exception:
            continue
        d0 = (
            (float(p0.X) - float(o.X)) * float(n.X)
            + (float(p0.Y) - float(o.Y)) * float(n.Y)
            + (float(p0.Z) - float(o.Z)) * float(n.Z)
        )
        d1 = (
            (float(p1.X) - float(o.X)) * float(n.X)
            + (float(p1.Y) - float(o.Y)) * float(n.Y)
            + (float(p1.Z) - float(o.Z)) * float(n.Z)
        )
        if d0 * d1 > 1e-10:
            continue
        den = d1 - d0
        if abs(den) < 1e-12:
            continue
        t = -d0 / den
        if t < -1e-6 or t > 1.0 + 1e-6:
            continue
        pts.append(
            XYZ(
                float(p0.X) + t * (float(p1.X) - float(p0.X)),
                float(p0.Y) + t * (float(p1.Y) - float(p0.Y)),
                float(p0.Z) + t * (float(p1.Z) - float(p0.Z)),
            )
        )
    return pts


def _loop_from_pts_xy(pts, z):
    if not pts or len(pts) < 3:
        return None
    uniq = []
    for p in pts:
        px, py = float(p.X), float(p.Y)
        skip = False
        for q in uniq:
            if abs(px - q[0]) < _MIN_LINE_FT and abs(py - q[1]) < _MIN_LINE_FT:
                skip = True
                break
        if not skip:
            uniq.append((px, py))
    if len(uniq) < 3:
        return None
    cx = sum(p[0] for p in uniq) / float(len(uniq))
    cy = sum(p[1] for p in uniq) / float(len(uniq))
    uniq.sort(key=lambda p: math.atan2(p[1] - cy, p[0] - cx))
    lines = []
    n = len(uniq)
    for i in range(n):
        a = uniq[i]
        b = uniq[(i + 1) % n]
        if math.hypot(b[0] - a[0], b[1] - a[1]) < _MIN_LINE_FT:
            continue
        try:
            lines.append(
                Line.CreateBound(
                    XYZ(a[0], a[1], z),
                    XYZ(b[0], b[1], z),
                )
            )
        except Exception:
            continue
    if len(lines) < 3:
        return None
    cl = list_curve_from_iterable(lines)
    try:
        return CurveLoop_Create(cl)
    except Exception:
        return None


def CurveLoop_Create(curves_list):
    from Autodesk.Revit.DB import CurveLoop

    try:
        return CurveLoop.Create(curves_list)
    except Exception:
        return None


def flatten_loop(loop, z=0.0):
    """Proyecta un CurveLoop a cota Z constante (plano horizontal)."""
    from Autodesk.Revit.DB import CurveLoop

    if loop is None:
        return None
    curves = []
    for c in iterate_net_collection(loop):
        if c is None:
            continue
        try:
            p0 = c.GetEndPoint(0)
            p1 = c.GetEndPoint(1)
        except Exception:
            continue
        a = XYZ(float(p0.X), float(p0.Y), z)
        b = XYZ(float(p1.X), float(p1.Y), z)
        if a.DistanceTo(b) < _MIN_LINE_FT:
            continue
        if isinstance(c, Line):
            try:
                curves.append(Line.CreateBound(a, b))
            except Exception:
                continue
        elif isinstance(c, Arc):
            try:
                mid = c.Evaluate(0.5, True)
                m = XYZ(float(mid.X), float(mid.Y), z)
                curves.append(Arc.Create(a, b, m))
            except Exception:
                try:
                    curves.append(Line.CreateBound(a, b))
                except Exception:
                    continue
        else:
            tes = []
            try:
                tes = list(c.Tessellate())
            except Exception:
                tes = []
            for i in range(len(tes) - 1):
                ta = tes[i]
                tb = tes[i + 1]
                pa = XYZ(float(ta.X), float(ta.Y), z)
                pb = XYZ(float(tb.X), float(tb.Y), z)
                if pa.DistanceTo(pb) < _MIN_LINE_FT:
                    continue
                try:
                    curves.append(Line.CreateBound(pa, pb))
                except Exception:
                    continue
    if len(curves) < 3:
        return None
    cl = list_curve_from_iterable(curves)
    try:
        return CurveLoop.Create(cl)
    except Exception:
        return None


def section_loops_at_z(solids, z):
    """CurveLoops horizontales del sólido ∩ plano Z=z."""
    if not solids:
        return []
    plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ(0.0, 0.0, float(z)))
    all_loops = []
    for solid in solids:
        face = _cut_face_from_solid(solid, plane)
        loops = _loops_from_face(face) if face is not None else []
        if not loops:
            pts = _section_pts_from_edges(solid, plane)
            lp = _loop_from_pts_xy(pts, float(z))
            if lp is not None:
                loops = [lp]
        for lp in loops:
            flat = flatten_loop(lp, 0.0)
            if flat is not None:
                all_loops.append(flat)
    return all_loops


def loops_to_solid(loops, height=_EXTRUDE_H_FT):
    if not loops:
        return None
    net = list_curve_loop_from_iterable(loops)
    if net is None or net.Count < 1:
        return None
    try:
        sol = GeometryCreationUtilities.CreateExtrusionGeometry(
            net, XYZ.BasisZ, float(height)
        )
    except Exception:
        return None
    if sol is None:
        return None
    try:
        if float(sol.Volume) <= _SOLID_VOL_TOL:
            return None
    except Exception:
        return None
    return sol


def loops_area(loops):
    sol = loops_to_solid(loops, 1.0)
    if sol is None:
        return 0.0
    try:
        return float(sol.Volume)
    except Exception:
        return 0.0


def boolean_solids(a, b, op):
    if a is None or b is None:
        return None
    try:
        return BooleanOperationsUtils.ExecuteBooleanOperation(a, b, op)
    except Exception:
        return None


def overlap_area(loops_a, loops_b):
    sa = loops_to_solid(loops_a, 1.0)
    sb = loops_to_solid(loops_b, 1.0)
    inter = boolean_solids(sa, sb, BooleanOperationsType.Intersect)
    if inter is None:
        return 0.0
    try:
        return float(inter.Volume)
    except Exception:
        return 0.0


def sections_coincident(loops_a, loops_b):
    aa = loops_area(loops_a)
    ab = loops_area(loops_b)
    if aa <= _SOLID_VOL_TOL or ab <= _SOLID_VOL_TOL:
        return False
    ov = overlap_area(loops_a, loops_b)
    return ov > _OVERLAP_RATIO * min(aa, ab)


def union_loops_as_solid(list_of_loops):
    acc = None
    for loops in list_of_loops or []:
        s = loops_to_solid(loops, _EXTRUDE_H_FT)
        if s is None:
            continue
        if acc is None:
            acc = s
        else:
            merged = boolean_solids(acc, s, BooleanOperationsType.Union)
            if merged is not None:
                acc = merged
    return acc


def _z_normal_faces(solid):
    faces = []
    if solid is None:
        return faces
    try:
        raw = solid.Faces
    except Exception:
        return faces
    for face in iterate_net_collection(raw):
        if not isinstance(face, PlanarFace):
            continue
        try:
            n = face.FaceNormal
            if abs(float(n.Z)) < 0.95:
                continue
            faces.append(face)
        except Exception:
            continue
    faces.sort(key=lambda f: float(getattr(f, "Area", 0.0)), reverse=True)
    return faces


def solid_to_region_loop_groups(solid):
    """Lista de grupos de CurveLoop (exterior + huecos) listos para FilledRegion."""
    if solid is None:
        return []
    parts = [solid]
    try:
        split = SolidUtils.SplitVolumes(solid)
        extra = [s for s in iterate_net_collection(split) if s is not None]
        if extra:
            parts = extra
    except Exception:
        parts = [solid]
    groups = []
    for part in parts:
        try:
            if float(part.Volume) <= _SOLID_VOL_TOL:
                continue
        except Exception:
            continue
        faces = _z_normal_faces(part)
        if not faces:
            continue
        loops = _loops_from_face(faces[0])
        flat = []
        for lp in loops:
            f = flatten_loop(lp, 0.0)
            if f is not None:
                flat.append(f)
        if flat:
            groups.append(flat)
    return groups


def difference_loop_groups(host_loops, subtract_loops_list):
    """Grupos de loops = host − unión(subtract). Si no hay subtract, el host entero."""
    if not host_loops:
        return []
    if not subtract_loops_list:
        return [host_loops]
    host_s = loops_to_solid(host_loops, _EXTRUDE_H_FT)
    if host_s is None:
        return []
    cut_s = union_loops_as_solid(subtract_loops_list)
    if cut_s is None:
        return [host_loops]
    diff = boolean_solids(host_s, cut_s, BooleanOperationsType.Difference)
    if diff is None:
        return [host_loops]
    try:
        if float(diff.Volume) <= _SOLID_VOL_TOL:
            return []
    except Exception:
        return []
    return solid_to_region_loop_groups(diff)


def classify_walls(doc, view):
    """
    Retorna (cut_items, nocut_items, skipped, zcut).
    Cada item: {wall, loops, id}.

    Colisión = el bbox **de modelo** del muro cruza Z_cut. No se usa el bbox
    del sólido (coords locales) ni solo la visibilidad de la vista.
    """
    zcut = cut_plane_z(view)
    cut_items = []
    nocut_items = []
    skipped = 0
    if zcut is None:
        return cut_items, nocut_items, skipped, None
    for wall in collect_walls_for_plan(doc, view):
        solids = wall_solids(wall)
        zmin, zmax = wall_model_z_range(wall, solids)
        if zmin is None or zmax is None:
            skipped += 1
            continue
        crosses = (zmin < zcut - _Z_TOL_FT) and (zmax > zcut + _Z_TOL_FT)
        if crosses:
            loops = section_loops_for_wall(wall, solids, zcut)
            if loops:
                cut_items.append(
                    {u"wall": wall, u"loops": loops, u"id": wall.Id}
                )
                continue
        mid = 0.5 * (zmin + zmax)
        loops = section_loops_for_wall(wall, solids, mid)
        if not loops:
            skipped += 1
            continue
        nocut_items.append({u"wall": wall, u"loops": loops, u"id": wall.Id})
    return cut_items, nocut_items, skipped, zcut


def build_region_groups(cut_items, nocut_items):
    """Lista de grupos de CurveLoop a crear (uno por pieza de desacople)."""
    groups = []
    for host in nocut_items or []:
        mates = []
        hloops = host.get(u"loops") or []
        for cut in cut_items or []:
            cloops = cut.get(u"loops") or []
            if sections_coincident(hloops, cloops):
                mates.append(cloops)
        for g in difference_loop_groups(hloops, mates):
            if g:
                groups.append(g)
    return groups


def create_filled_regions(doc, view, type_id, loop_groups):
    created = 0
    failed = 0
    ids = []
    if doc is None or view is None or type_id is None:
        return created, failed, ids
    for loops in loop_groups or []:
        net = list_curve_loop_from_iterable(loops)
        if net is None or net.Count < 1:
            failed += 1
            continue
        try:
            region = FilledRegion.Create(doc, type_id, view.Id, net)
        except Exception:
            region = None
        if region is None:
            failed += 1
            continue
        created += 1
        try:
            ids.append(region.Id)
        except Exception:
            pass
    return created, failed, ids


def generate(doc, view, type_id):
    """
    Crea las Filled Region. Retorna dict con conteos.
    Debe llamarse dentro o envuelve transacción.
    """
    cut_items, nocut_items, skipped, zcut = classify_walls(doc, view)
    groups = build_region_groups(cut_items, nocut_items)
    t = Transaction(doc, TXN_NAME)
    t.Start()
    try:
        created, failed, ids = create_filled_regions(doc, view, type_id, groups)
        t.Commit()
    except Exception:
        try:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
        except Exception:
            pass
        raise
    return {
        u"cut": len(cut_items),
        u"nocut": len(nocut_items),
        u"skipped": skipped,
        u"groups": len(groups),
        u"created": created,
        u"failed": failed,
        u"ids": ids,
        u"zcut": zcut,
    }
