# -*- coding: utf-8 -*-
"""
Arainco: Magic Dimensions — cotas automáticas por escenario.

Escenario actual: ejes (Grid) en vistas Building Section.
La línea de cota se coloca por PickPoint en la vista activa.

Revit 2024+ | IronPython (pyRevit).
"""

from __future__ import division, print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    DatumEnds,
    DatumExtentType,
    Grid,
    Line,
    Plane,
    Reference,
    ReferenceArray,
    SketchPlane,
    Transaction,
    TransactionStatus,
    View,
    XYZ,
)
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectSnapTypes, ObjectType

_TOOL_TITLE = u"Arainco: Magic Dimensions"
_TXN_NAME = u"Arainco: Magic Dimensions"
_TXN_SKETCH = u"Arainco: Magic Dimensions (plano de trabajo)"

SCENARIO_GRIDS_BUILDING_SECTION = u"grids_building_section"
SCENARIO_FLOORS_SECTION_ELEVATION = u"floors_section_elevation"

SCENARIOS = (
    (
        SCENARIO_GRIDS_BUILDING_SECTION,
        u"Ejes (Grid) en Building Section",
        u"Seleccione dos o más ejes visibles en la sección. "
        u"Luego indique un punto: ahí se coloca la línea de cota "
        u"y se acorta el 2D de los ejes (400 mm a 1:50) para acercar la cabeza a la cota.",
    ),
    (
        SCENARIO_FLOORS_SECTION_ELEVATION,
        u"Cota de elevación — losas, fundaciones, muros y vigas",
        u"Seleccione dos o más losas, fundaciones, muros o Structural Framing "
        u"en Sección o Alzado. Luego indique un punto: cota principal "
        u"(losas, muros y vigas por cara superior, fundaciones por sello), "
        u"segunda cota de espesor de fundación a 350 mm (1:50) y Spot Elevation.",
    ),
)

_PROMPT_GRIDS = (
    u"Seleccione dos o más ejes (Grid). Finish para continuar / Esc cancela."
)
_PROMPT_POINT = u"Indique el punto donde colocar la línea de cota."

_SKIP_SECTION_DOT = 0.85
_MIN_LINE_FT = 0.05
# Hueco entre la cota y la cabeza 2D del eje, calibrado a 1:50.
_REF_VIEW_SCALE = 50
_GRID_HEAD_GAP_MM_AT_50 = 400.0


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _eid_int(eid):
    if eid is None:
        return -1
    try:
        return int(eid.IntegerValue)
    except Exception:
        pass
    try:
        return int(eid.Value)
    except Exception:
        return -1


def _xyz_dot(a, b):
    return (
        float(a.X) * float(b.X)
        + float(a.Y) * float(b.Y)
        + float(a.Z) * float(b.Z)
    )


def _xyz_sub(a, b):
    return XYZ(float(a.X) - float(b.X), float(a.Y) - float(b.Y), float(a.Z) - float(b.Z))


def _xyz_add(a, b):
    return XYZ(float(a.X) + float(b.X), float(a.Y) + float(b.Y), float(a.Z) + float(b.Z))


def _xyz_scale(v, s):
    return XYZ(float(v.X) * s, float(v.Y) * s, float(v.Z) * s)


def _unit(v):
    if v is None:
        return None
    try:
        ln = float(v.GetLength())
    except Exception:
        return None
    if ln < 1e-12:
        return None
    try:
        return v.Normalize()
    except Exception:
        return XYZ(float(v.X) / ln, float(v.Y) / ln, float(v.Z) / ln)


def resolve_active_view(uidoc):
    if uidoc is None:
        return None
    view = None
    try:
        view = getattr(uidoc, "ActiveGraphicalView", None)
    except Exception:
        view = None
    if view is None:
        try:
            view = uidoc.ActiveView
        except Exception:
            view = None
    if view is None:
        return None
    try:
        if not isinstance(view, View):
            view = uidoc.Document.GetElement(view.Id)
    except Exception:
        pass
    return view


def is_building_section_view(view):
    try:
        from filtro_armadura_eje import es_vista_building_section

        return bool(es_vista_building_section(view))
    except Exception:
        pass
    if view is None:
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import ViewType

        return view.ViewType == ViewType.Section
    except Exception:
        return False


def view_display_name(view):
    if view is None:
        return u"(sin vista)"
    try:
        name = _as_unicode(view.Name).strip()
    except Exception:
        name = u""
    return name or u"(sin nombre)"


class _GridSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return _is_grid(element)

    def AllowReference(self, reference, point):
        return False


def _is_grid(element):
    if element is None:
        return False
    if isinstance(element, Grid):
        return True
    try:
        cat = element.Category
        if cat is None:
            return False
        return _eid_int(cat.Id) == int(BuiltInCategory.OST_Grids)
    except Exception:
        return False


def collect_grids_from_ids(doc, eids):
    grids = []
    seen = set()
    if doc is None or eids is None:
        return grids
    for eid in eids:
        try:
            el = doc.GetElement(eid)
        except Exception:
            el = None
        if not _is_grid(el):
            continue
        iid = _eid_int(el.Id)
        if iid in seen:
            continue
        seen.add(iid)
        grids.append(el)
    return grids


def preselected_grids(uidoc):
    if uidoc is None:
        return []
    try:
        eids = uidoc.Selection.GetElementIds()
    except Exception:
        return []
    return collect_grids_from_ids(uidoc.Document, eids)


def pick_grids(uidoc):
    refs = list(
        uidoc.Selection.PickObjects(
            ObjectType.Element, _GridSelectionFilter(), _PROMPT_GRIDS
        )
    )
    eids = []
    for pref in refs:
        try:
            eids.append(pref.ElementId)
        except Exception:
            continue
    return collect_grids_from_ids(uidoc.Document, eids)


def _ensure_view_sketch_plane(doc, view):
    if doc is None or view is None:
        return False
    try:
        if view.SketchPlane is not None:
            return True
    except Exception:
        pass
    t = Transaction(doc, _TXN_SKETCH)
    t.Start()
    try:
        plane = Plane.CreateByNormalAndOrigin(view.ViewDirection, view.Origin)
        view.SketchPlane = SketchPlane.Create(doc, plane)
        t.Commit()
        return True
    except Exception:
        try:
            if t.GetStatus() == TransactionStatus.Started:
                t.RollBack()
        except Exception:
            pass
        return False


def pick_dimension_point(uidoc, view):
    doc = uidoc.Document
    _ensure_view_sketch_plane(doc, view)
    snaps = (
        ObjectSnapTypes.Endpoints
        | ObjectSnapTypes.Intersections
        | ObjectSnapTypes.Nearest
        | ObjectSnapTypes.Perpendicular
    )
    try:
        return uidoc.Selection.PickPoint(snaps, _PROMPT_POINT)
    except Exception:
        return uidoc.Selection.PickPoint(_PROMPT_POINT)


def _project_onto_view_plane(pt, view):
    origin = view.Origin
    normal = _unit(view.ViewDirection)
    if pt is None or origin is None or normal is None:
        return pt
    delta = _xyz_sub(pt, origin)
    dist = _xyz_dot(delta, normal)
    return _xyz_sub(pt, _xyz_scale(normal, dist))


def _grid_curves(grid, view):
    for ext in (DatumExtentType.ViewSpecific, DatumExtentType.Model):
        try:
            curves = grid.GetCurvesInView(ext, view)
        except Exception:
            curves = None
        if curves:
            out = []
            for c in curves:
                if c is not None:
                    out.append(c)
            if out:
                return out
    try:
        c = grid.Curve
        if c is not None:
            return [c]
    except Exception:
        pass
    return []


def _curve_direction(curve):
    try:
        return _unit(curve.Direction)
    except Exception:
        pass
    try:
        return _unit(_xyz_sub(curve.GetEndPoint(1), curve.GetEndPoint(0)))
    except Exception:
        return None


def _curve_midpoint(curve):
    try:
        return curve.Evaluate(0.5, True)
    except Exception:
        pass
    try:
        a = curve.GetEndPoint(0)
        b = curve.GetEndPoint(1)
        return XYZ(
            0.5 * (float(a.X) + float(b.X)),
            0.5 * (float(a.Y) + float(b.Y)),
            0.5 * (float(a.Z) + float(b.Z)),
        )
    except Exception:
        return None


def _grid_plane_normal(grid, view):
    curves = _grid_curves(grid, view)
    for curve in curves:
        direction = _curve_direction(curve)
        if direction is None:
            continue
        normal = _unit(direction.CrossProduct(XYZ.BasisZ))
        if normal is not None:
            return normal
        normal = _unit(direction.CrossProduct(XYZ(0, 0, 1)))
        if normal is not None:
            return normal
    return None


def _grid_appears_in_section(grid, view):
    """True si el plano del eje corta la sección (aparece como línea)."""
    n_grid = _grid_plane_normal(grid, view)
    vd = _unit(view.ViewDirection)
    if n_grid is None or vd is None:
        return bool(_grid_curves(grid, view))
    return abs(_xyz_dot(vd, n_grid)) < _SKIP_SECTION_DOT


def _grid_anchor_on_view(grid, view):
    for curve in _grid_curves(grid, view):
        mid = _curve_midpoint(curve)
        if mid is not None:
            return _project_onto_view_plane(mid, view)
    return None


def _grid_dir_key(grid, view):
    n_grid = _grid_plane_normal(grid, view)
    if n_grid is None:
        return None
    if float(n_grid.X) < 0 or (abs(float(n_grid.X)) < 1e-9 and float(n_grid.Y) < 0):
        n_grid = _xyz_scale(n_grid, -1.0)
    return (
        round(float(n_grid.X), 3),
        round(float(n_grid.Y), 3),
        round(float(n_grid.Z), 3),
    )


def _largest_parallel_group(grids, view):
    groups = {}
    leftovers = []
    for grid in grids:
        key = _grid_dir_key(grid, view)
        if key is None:
            leftovers.append(grid)
            continue
        groups.setdefault(key, []).append(grid)
    if not groups:
        return leftovers
    best = []
    for members in groups.values():
        if len(members) > len(best):
            best = members
    return best


def _grid_reference(grid):
    try:
        return Reference(grid)
    except Exception:
        return None


def _view_scale_int(view):
    try:
        s = int(view.Scale)
        if s > 0:
            return s
    except Exception:
        pass
    return _REF_VIEW_SCALE


def _grid_head_gap_ft(view):
    mm = _GRID_HEAD_GAP_MM_AT_50 * (
        float(_view_scale_int(view)) / float(_REF_VIEW_SCALE)
    )
    return float(mm) / 304.8


def _xyz_length(v):
    try:
        return float(v.GetLength())
    except Exception:
        return (
            float(v.X) * float(v.X)
            + float(v.Y) * float(v.Y)
            + float(v.Z) * float(v.Z)
        ) ** 0.5


def _bubble_visible(grid, view, end):
    try:
        return bool(grid.IsBubbleVisibleInView(end, view))
    except Exception:
        return False


def _ensure_grid_2d_extents(grid, view):
    for end in (DatumEnds.End0, DatumEnds.End1):
        try:
            if grid.GetDatumExtentTypeInView(end, view) != DatumExtentType.ViewSpecific:
                grid.SetDatumExtentType(end, view, DatumExtentType.ViewSpecific)
        except Exception:
            pass


def _shorten_grid_2d_to_dimension(grid, view, pick_pt):
    """
    Recorta el largo 2D del eje en la vista: la cabeza (burbuja) queda
    cerca de la línea de cota, con un hueco proporcional a la escala.
    """
    if grid is None or view is None or pick_pt is None:
        return False
    _ensure_grid_2d_extents(grid, view)
    curves = []
    try:
        raw = grid.GetCurvesInView(DatumExtentType.ViewSpecific, view)
        if raw:
            curves = [c for c in raw if c is not None]
    except Exception:
        curves = []
    if not curves:
        curves = _grid_curves(grid, view)
    if not curves:
        return False
    curve = curves[0]
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
    except Exception:
        return False
    origin = _project_onto_view_plane(pick_pt, view)
    axis = _unit(_xyz_sub(p1, p0))
    if origin is None or axis is None:
        return False
    t = _xyz_dot(_xyz_sub(origin, p0), axis)
    p_dim = _xyz_add(p0, _xyz_scale(axis, t))
    gap = max(_MIN_LINE_FT, _grid_head_gap_ft(view))

    move = []
    if _bubble_visible(grid, view, DatumEnds.End0):
        move.append(0)
    if _bubble_visible(grid, view, DatumEnds.End1):
        move.append(1)
    if not move:
        d0 = _xyz_length(_xyz_sub(p0, p_dim))
        d1 = _xyz_length(_xyz_sub(p1, p_dim))
        move = [0] if d0 >= d1 else [1]

    new0, new1 = p0, p1
    for i in move:
        end_pt = p0 if i == 0 else p1
        vec = _xyz_sub(end_pt, p_dim)
        ln = _xyz_length(vec)
        if ln <= gap + 1e-6:
            continue
        direction = _unit(vec)
        if direction is None:
            continue
        new_end = _xyz_add(p_dim, _xyz_scale(direction, gap))
        if i == 0:
            new0 = new_end
        else:
            new1 = new_end

    if _xyz_length(_xyz_sub(new1, new0)) < _MIN_LINE_FT:
        return False
    try:
        new_line = Line.CreateBound(new0, new1)
        grid.SetCurveInView(DatumExtentType.ViewSpecific, view, new_line)
        return True
    except Exception:
        return False


def _shorten_grids_2d_to_dimension(grids, view, pick_pt):
    n = 0
    for grid in grids or []:
        if _shorten_grid_2d_to_dimension(grid, view, pick_pt):
            n += 1
    return n


def create_grid_dimension_in_building_section(doc, view, grids, pick_pt):
    """
    Crea una cota lineal entre ejes en una Building Section.

    Returns:
        (dimension_or_None, message)
    """
    if doc is None or view is None:
        return None, u"No hay documento o vista activa."
    if pick_pt is None:
        return None, u"No se indicó el punto de la cota."

    usable = []
    seen = set()
    for grid in grids or []:
        if not _is_grid(grid):
            continue
        iid = _eid_int(grid.Id)
        if iid in seen:
            continue
        seen.add(iid)
        if not _grid_appears_in_section(grid, view):
            continue
        usable.append(grid)

    usable = _largest_parallel_group(usable, view)
    if len(usable) < 2:
        return None, (
            u"Se necesitan al menos dos ejes que aparezcan como líneas "
            u"en esta Building Section (ejes que cortan el plano de la vista)."
        )

    refs = ReferenceArray()
    anchors = []
    for grid in usable:
        ref = _grid_reference(grid)
        if ref is None:
            continue
        anchor = _grid_anchor_on_view(grid, view)
        refs.Append(ref)
        if anchor is not None:
            anchors.append(anchor)

    if refs.Size < 2:
        return None, u"No se pudieron obtener referencias de los ejes seleccionados."

    origin = _project_onto_view_plane(pick_pt, view)
    right = _unit(view.RightDirection)
    if origin is None or right is None:
        return None, u"No se pudo proyectar el punto sobre el plano de la vista."

    ts = [0.0]
    for pt in anchors:
        ts.append(_xyz_dot(_xyz_sub(pt, origin), right))
    t_min = min(ts)
    t_max = max(ts)
    if abs(t_max - t_min) < _MIN_LINE_FT:
        t_min -= 1.0
        t_max += 1.0
    else:
        pad = max(0.5, 0.05 * abs(t_max - t_min))
        t_min -= pad
        t_max += pad

    p1 = _xyz_add(origin, _xyz_scale(right, t_min))
    p2 = _xyz_add(origin, _xyz_scale(right, t_max))
    try:
        dim_line = Line.CreateBound(p1, p2)
    except Exception as ex:
        return None, u"No se pudo construir la línea de cota. {}".format(_as_unicode(ex))

    t = Transaction(doc, _TXN_NAME)
    t.Start()
    try:
        dim = doc.Create.NewDimension(view, dim_line, refs)
        if dim is None:
            raise Exception(u"NewDimension no devolvió una cota.")
        n_short = _shorten_grids_2d_to_dimension(usable, view, pick_pt)
        t.Commit()
        msg = u"Cota creada entre {0} ejes.".format(int(refs.Size))
        if n_short:
            msg = msg + u" Se acortó el 2D de {0} eje(s).".format(n_short)
        return dim, msg
    except Exception as ex:
        try:
            if t.GetStatus() == TransactionStatus.Started:
                t.RollBack()
        except Exception:
            pass
        return None, u"No se pudo crear la cota. {}".format(_as_unicode(ex))


def select_created_dimension(uidoc, dim):
    if uidoc is None or dim is None:
        return
    try:
        from System.Collections.Generic import List
        from Autodesk.Revit.DB import ElementId

        ids = List[ElementId]()
        ids.Add(dim.Id)
        uidoc.Selection.SetElementIds(ids)
    except Exception:
        pass


def run_grids_building_section(uidoc, aviso_fn, use_preselection=True):
    """
    Flujo de selección + punto + cota para el escenario de ejes.

    ``aviso_fn(instruction, content=u"")`` muestra avisos al usuario.
    Returns: (ok, status_text)
    """
    view = resolve_active_view(uidoc)
    if view is None:
        aviso_fn(u"No hay una vista gráfica activa.")
        return False, u"Sin vista activa."
    if not is_building_section_view(view):
        aviso_fn(
            u"Este escenario solo opera en vistas Building Section.",
            content=u"Vista activa: {0}.".format(view_display_name(view)),
        )
        return False, u"La vista activa no es Building Section."

    grids = []
    if use_preselection:
        grids = preselected_grids(uidoc)
    if len(grids) < 2:
        try:
            grids = pick_grids(uidoc)
        except OperationCanceledException:
            return False, u"Selección de ejes cancelada."
        except Exception as ex:
            aviso_fn(u"No se pudieron seleccionar los ejes.", content=_as_unicode(ex))
            return False, u"Error al seleccionar ejes."

    if len(grids) < 2:
        aviso_fn(u"Seleccione al menos dos ejes (Grid).")
        return False, u"Hacen falta al menos dos ejes."

    try:
        pick_pt = pick_dimension_point(uidoc, view)
    except OperationCanceledException:
        return False, u"Punto de cota cancelado."
    except Exception as ex:
        aviso_fn(
            u"No se pudo indicar el punto de la cota.",
            content=_as_unicode(ex),
        )
        return False, u"Error al indicar el punto."

    dim, msg = create_grid_dimension_in_building_section(
        uidoc.Document, view, grids, pick_pt
    )
    if dim is None:
        aviso_fn(msg)
        return False, msg
    select_created_dimension(uidoc, dim)
    return True, msg


def run_floors_section_elevation(uidoc, aviso_fn, use_preselection=True):
    """Cota alineada + Spot Elevations en losas (misma lógica que Cota Spot Losas)."""
    try:
        from cota_spot_elevacion_losas import run_floor_cota_and_spots
    except Exception as ex:
        aviso_fn(
            u"No se pudo cargar la lógica de cotas de losas.",
            content=_as_unicode(ex),
        )
        return False, u"Módulo de losas no disponible."

    return run_floor_cota_and_spots(
        uidoc,
        aviso_fn,
        use_preselection=use_preselection,
        txn_group=_TXN_NAME,
        txn_sketch=_TXN_SKETCH,
        txn_create=_TXN_NAME,
        show_success_dialog=False,
    )


def run_scenario(scenario_key, uidoc, aviso_fn, use_preselection=True):
    if scenario_key == SCENARIO_FLOORS_SECTION_ELEVATION:
        return run_floors_section_elevation(
            uidoc, aviso_fn, use_preselection=use_preselection
        )
    return run_grids_building_section(
        uidoc, aviso_fn, use_preselection=use_preselection
    )
