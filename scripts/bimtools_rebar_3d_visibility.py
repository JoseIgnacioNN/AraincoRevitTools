# -*- coding: utf-8 -*-
u"""Visibilidad de armadura en vistas 3D (equivalente a «View Unobscured» en Revit)."""

from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementCategoryFilter,
    FilteredElementCollector,
    View,
    View3D,
)
from Autodesk.Revit.DB.Structure import (
    AreaReinforcement,
    PathReinforcement,
    Rebar,
    RebarInSystem,
)


def iter_document_view3d_non_template(doc):
    u"""Vistas ``View3D`` del documento que no son plantilla."""
    if doc is None:
        return
    for v in FilteredElementCollector(doc).OfClass(View3D):
        if v is None:
            continue
        try:
            if v.IsTemplate:
                continue
        except Exception:
            pass
        yield v


def collect_reinforcement_in_view(doc, view):
    u"""
    Recoge ``Rebar``, ``RebarInSystem``, ``AreaReinforcement`` y
    ``PathReinforcement`` visibles en ``view``.

    Usa ``FilteredElementCollector(doc, view.Id)`` — solo esa vista, no el documento.
    """
    if doc is None or view is None or not isinstance(view, View):
        return []
    try:
        if view.IsTemplate:
            return []
    except Exception:
        pass
    out = []
    seen = set()
    view_id = view.Id
    for cls in (Rebar, RebarInSystem, AreaReinforcement, PathReinforcement):
        try:
            elems = (
                FilteredElementCollector(doc, view_id)
                .OfClass(cls)
                .WhereElementIsNotElementType()
                .ToElements()
            )
        except Exception:
            elems = []
        for el in elems or []:
            if el is None:
                continue
            try:
                eid = el.Id
            except Exception:
                continue
            try:
                key = int(eid.IntegerValue)
            except AttributeError:
                try:
                    key = int(eid.Value)
                except Exception:
                    key = None
            if key is None or key in seen:
                continue
            seen.add(key)
            out.append(el)
    return out


def _resolve_reinforcement_element(doc, ref):
    if ref is None:
        return None
    if isinstance(ref, (AreaReinforcement, PathReinforcement, Rebar, RebarInSystem)):
        return ref
    try:
        ref = doc.GetElement(ref)
    except Exception:
        ref = None
    if isinstance(ref, (AreaReinforcement, PathReinforcement, Rebar, RebarInSystem)):
        return ref
    return None


def _set_solid_in_view(ref, view, solid):
    try:
        ref.SetSolidInView(view, solid)
    except Exception:
        try:
            fn = getattr(ref, "SetSolidInView", None)
            if fn is not None:
                fn(view, solid)
        except Exception:
            pass


def is_reinforcement_unobscured_in_view(ref, view):
    u"""``True``/``False`` si la API responde; ``None`` si no aplica o falla."""
    if ref is None or view is None:
        return None
    try:
        return bool(ref.IsUnobscuredInView(view))
    except Exception:
        return None


def summarize_reinforcement_unobscured_in_view(doc, refuerzos, view):
    u"""
    Cuenta elementos con View Unobscured activo, inactivo o sin dato en ``view``.
    """
    if not refuerzos or doc is None or view is None or not isinstance(view, View):
        return {"total": 0, "unobscured": 0, "obscured": 0, "unknown": 0}
    n_unobscured = 0
    n_obscured = 0
    n_unknown = 0
    for ref in refuerzos:
        ref = _resolve_reinforcement_element(doc, ref)
        if ref is None:
            n_unknown += 1
            continue
        state = is_reinforcement_unobscured_in_view(ref, view)
        if state is True:
            n_unobscured += 1
        elif state is False:
            n_obscured += 1
        else:
            n_unknown += 1
    return {
        "total": len(refuerzos),
        "unobscured": n_unobscured,
        "obscured": n_obscured,
        "unknown": n_unknown,
    }


def _element_id_key(eid):
    if eid is None:
        return None
    try:
        return int(eid.IntegerValue)
    except AttributeError:
        try:
            return int(eid.Value)
        except Exception:
            return None
    except Exception:
        return None


def _iter_system_reinforcement_bar_children(doc, system_rein):
    u"""
    ``RebarInSystem`` / ``Rebar`` hijos de ``AreaReinforcement`` o ``PathReinforcement``.

    Tras ``Regenerate``: ``GetRebarInSystemIds`` y, si hace falta,
    ``GetDependentElements`` (p. ej. cuando HostStructuralRebar genera ``Rebar``).
    """
    if doc is None or system_rein is None:
        return
    seen = set()
    try:
        sys_ids = system_rein.GetRebarInSystemIds()
    except Exception:
        sys_ids = None
    if sys_ids is not None:
        try:
            n = int(sys_ids.Count)
        except Exception:
            n = 0
        for i in range(n):
            try:
                rid = sys_ids[i]
                key = _element_id_key(rid)
                if key is None or key in seen:
                    continue
                child = doc.GetElement(rid)
                if isinstance(child, (RebarInSystem, Rebar)):
                    seen.add(key)
                    yield child
            except Exception:
                continue
    for cat, cls in (
        (BuiltInCategory.OST_RebarInSystem, RebarInSystem),
        (BuiltInCategory.OST_Rebar, Rebar),
    ):
        try:
            dep = system_rein.GetDependentElements(ElementCategoryFilter(cat))
        except Exception:
            dep = None
        if dep is None:
            continue
        try:
            nd = int(dep.Count)
        except Exception:
            nd = 0
        for i in range(nd):
            try:
                rid = dep[i]
                key = _element_id_key(rid)
                if key is None or key in seen:
                    continue
                child = doc.GetElement(rid)
                if isinstance(child, cls):
                    seen.add(key)
                    yield child
            except Exception:
                continue


def _iter_area_reinforcement_bar_children(doc, area_rein):
    u"""Compat: mismos hijos que ``_iter_system_reinforcement_bar_children``."""
    for child in _iter_system_reinforcement_bar_children(doc, area_rein):
        yield child


def _apply_visibility_to_element(ref, doc, view, unobscured, solid_in_view):
    applied = False
    try:
        ref.SetUnobscuredInView(view, unobscured)
        applied = True
    except Exception:
        pass
    _set_solid_in_view(ref, view, solid_in_view)
    if isinstance(ref, (AreaReinforcement, PathReinforcement)):
        for child in _iter_system_reinforcement_bar_children(doc, ref):
            try:
                child.SetUnobscuredInView(view, unobscured)
            except Exception:
                pass
            _set_solid_in_view(child, view, solid_in_view)
    return applied


def apply_reinforcement_unobscured_in_view(doc, refuerzos, view, unobscured=True, solid_in_view=None):
    u"""
    ``SetUnobscuredInView`` + ``SetSolidInView`` para ``AreaReinforcement``,
    ``PathReinforcement``, ``Rebar`` o ``RebarInSystem`` **solo** en la vista
    indicada (nunca en otras).

    ``unobscured``: ``True`` para activar View Unobscured; ``False`` para quitarlo.
    ``solid_in_view``: si es ``None``, sigue el mismo valor que ``unobscured``.
    """
    if not refuerzos or doc is None or view is None:
        return 0
    if not isinstance(view, View):
        return 0
    try:
        if view.IsTemplate:
            return 0
    except Exception:
        pass
    try:
        view = doc.GetElement(view.Id)
    except Exception:
        return 0
    if not isinstance(view, View):
        return 0
    if solid_in_view is None:
        solid_in_view = unobscured
    n_ok = 0
    for ref in refuerzos:
        ref = _resolve_reinforcement_element(doc, ref)
        if ref is None:
            continue
        if _apply_visibility_to_element(ref, doc, view, unobscured, solid_in_view):
            n_ok += 1
    return n_ok


def _resolve_rebar_element(doc, ref):
    if ref is None:
        return None
    if isinstance(ref, Rebar):
        return ref
    try:
        ref = doc.GetElement(ref)
    except Exception:
        ref = None
    if isinstance(ref, Rebar):
        return ref
    return None


def _apply_rebar_visibility_in_view(doc, rebars, view, unobscured, solid_in_view=None):
    u"""``SetUnobscuredInView`` + ``SetSolidInView`` en ``Rebar`` para ``view``."""
    if not rebars or doc is None or view is None:
        return 0
    if not isinstance(view, View):
        return 0
    try:
        if view.IsTemplate:
            return 0
    except Exception:
        pass
    try:
        view = doc.GetElement(view.Id)
    except Exception:
        return 0
    if not isinstance(view, View):
        return 0
    if solid_in_view is None:
        solid_in_view = unobscured
    n_ok = 0
    for ref in rebars:
        rb = _resolve_rebar_element(doc, ref)
        if rb is None:
            continue
        applied = False
        try:
            rb.SetUnobscuredInView(view, unobscured)
            applied = True
        except Exception:
            pass
        _set_solid_in_view(rb, view, solid_in_view)
        if applied:
            n_ok += 1
    return n_ok


def apply_rebar_unobscured_in_view(doc, rebars, view):
    u"""
    Para cada ``Rebar`` en ``rebars``, activa «View Unobscured» y sólido en vista
    **solo** en la vista indicada.
    """
    return _apply_rebar_visibility_in_view(doc, rebars, view, True, True)


def apply_rebar_unobscured_off_in_view(doc, rebars, view):
    u"""
    Para cada ``Rebar`` en ``rebars``, desactiva «View Unobscured» y sólido en vista
    **solo** en la vista indicada.

    Usado por Armado Muros para barras de **malla** (no deben quedar Unobscured).
    """
    return _apply_rebar_visibility_in_view(doc, rebars, view, False, False)


def ensure_rebar_obscured_in_view(doc, rebars, view):
    u"""Desactiva «View Unobscured» (+ sólido) **solo** en la vista indicada."""
    return apply_rebar_unobscured_off_in_view(doc, rebars, view)


def apply_rebar_unobscured_in_3d_views(doc, rebars):
    u"""
    Para cada ``Rebar`` en ``rebars``, activa visible sin oscurecer y sólido
    en todas las vistas 3D no plantilla del documento.
    """
    if not rebars or doc is None:
        return
    views = list(iter_document_view3d_non_template(doc))
    if not views:
        return
    for rb in rebars:
        if rb is None:
            continue
        if not isinstance(rb, Rebar):
            continue
        for v in views:
            try:
                rb.SetUnobscuredInView(v, True)
            except Exception:
                pass
            try:
                rb.SetSolidInView(v, True)
            except Exception:
                try:
                    fn = getattr(rb, "SetSolidInView", None)
                    if fn is not None:
                        fn(v, True)
                except Exception:
                    pass
