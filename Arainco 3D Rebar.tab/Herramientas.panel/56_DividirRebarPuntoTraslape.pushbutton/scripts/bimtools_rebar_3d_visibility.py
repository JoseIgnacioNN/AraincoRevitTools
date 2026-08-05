# -*- coding: utf-8 -*-
u"""Visibilidad de armadura en vistas 3D (equivalente a «View Unobscured» en Revit)."""

from Autodesk.Revit.DB import FilteredElementCollector, View, View3D
from Autodesk.Revit.DB.Structure import AreaReinforcement, Rebar, RebarInSystem


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
    Recoge ``Rebar``, ``RebarInSystem`` y ``AreaReinforcement`` visibles en ``view``.

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
    for cls in (Rebar, RebarInSystem, AreaReinforcement):
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
    if isinstance(ref, (AreaReinforcement, Rebar, RebarInSystem)):
        return ref
    try:
        ref = doc.GetElement(ref)
    except Exception:
        ref = None
    if isinstance(ref, (AreaReinforcement, Rebar, RebarInSystem)):
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


def _apply_visibility_to_element(ref, doc, view, unobscured, solid_in_view):
    applied = False
    try:
        ref.SetUnobscuredInView(view, unobscured)
        applied = True
    except Exception:
        pass
    _set_solid_in_view(ref, view, solid_in_view)
    if isinstance(ref, AreaReinforcement):
        try:
            ids = ref.GetRebarInSystemIds()
        except Exception:
            ids = None
        if ids:
            for rid in ids:
                child = doc.GetElement(rid)
                if child is None:
                    continue
                try:
                    child.SetUnobscuredInView(view, unobscured)
                except Exception:
                    pass
                _set_solid_in_view(child, view, solid_in_view)
    return applied


def apply_reinforcement_unobscured_in_view(doc, refuerzos, view, unobscured=True, solid_in_view=None):
    u"""
    ``SetUnobscuredInView`` + ``SetSolidInView`` para ``AreaReinforcement``,
    ``Rebar`` o ``RebarInSystem`` **solo** en la vista indicada (nunca en otras).

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


def apply_rebar_unobscured_in_view(doc, rebars, view):
    u"""
    Para cada ``Rebar`` en ``rebars``, activa «View Unobscured» y sólido en vista
    **solo** en la vista indicada.
    """
    if not rebars or doc is None or view is None:
        return
    if not isinstance(view, View):
        return
    try:
        if view.IsTemplate:
            return
    except Exception:
        pass
    try:
        view = doc.GetElement(view.Id)
    except Exception:
        return
    if not isinstance(view, View):
        return
    for rb in rebars:
        if rb is None:
            continue
        if not isinstance(rb, Rebar):
            try:
                rb = doc.GetElement(rb)
            except Exception:
                rb = None
            if not isinstance(rb, Rebar):
                continue
        try:
            rb.SetUnobscuredInView(view, True)
        except Exception:
            pass
        try:
            rb.SetSolidInView(view, True)
        except Exception:
            try:
                fn = getattr(rb, "SetSolidInView", None)
                if fn is not None:
                    fn(view, True)
            except Exception:
                pass


def ensure_rebar_obscured_in_view(doc, rebars, view):
    u"""Desactiva «View Unobscured» **solo** en la vista indicada."""
    if not rebars or doc is None or view is None:
        return
    if not isinstance(view, View):
        return
    try:
        if view.IsTemplate:
            return
    except Exception:
        pass
    try:
        view = doc.GetElement(view.Id)
    except Exception:
        return
    if not isinstance(view, View):
        return
    for rb in rebars:
        if rb is None:
            continue
        if not isinstance(rb, Rebar):
            try:
                rb = doc.GetElement(rb)
            except Exception:
                rb = None
            if not isinstance(rb, Rebar):
                continue
        try:
            rb.SetUnobscuredInView(view, False)
        except Exception:
            pass


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
