# -*- coding: utf-8 -*-
u"""Visibilidad de armadura en vistas 3D (equivalente a «View Unobscured» en Revit)."""

from Autodesk.Revit.DB import FilteredElementCollector, View3D
from Autodesk.Revit.DB.Structure import Rebar


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
