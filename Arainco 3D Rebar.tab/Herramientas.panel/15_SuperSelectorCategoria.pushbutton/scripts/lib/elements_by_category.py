# -*- coding: utf-8 -*-
"""Elementos de categorías de modelado presentes en una vista."""

from __future__ import print_function

from Autodesk.Revit.DB import CategoryType, ElementId, FilteredElementCollector


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _model_category_label(element):
    try:
        cat = element.Category
        if cat is None:
            return None
        if cat.CategoryType != CategoryType.Model:
            return None
        name = _as_unicode(cat.Name).strip()
        return name or None
    except Exception:
        return None


def collect_model_elements_in_view(doc, view):
    """Instancias de categoría Model visibles en la vista activa."""
    out = []
    if doc is None or view is None:
        return out
    try:
        elems = (
            FilteredElementCollector(doc, view.Id)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        return out
    for el in elems or []:
        if _model_category_label(el) is not None:
            out.append(el)
    return out


def group_elements_by_model_category(elements):
    """
    Agrupa instancias por nombre de categoría de modelado.

    Returns:
        list[dict]: ``{u"label", u"element_ids", u"count"}`` ordenado por label.
    """
    groups = {}
    for el in elements or []:
        label = _model_category_label(el)
        if label is None:
            continue
        bucket = groups.get(label)
        if bucket is None:
            bucket = []
            groups[label] = bucket
        try:
            bucket.append(el.Id)
        except Exception:
            continue

    return [
        {
            u"label": label,
            u"element_ids": groups[label],
            u"count": len(groups[label]),
        }
        for label in sorted(groups.keys())
    ]


def summarize_view_elements(elements, groups):
    total = len(elements or [])
    return {
        u"total": total,
        u"categories": len(groups or []),
    }


def select_elements_in_model(uidoc, element_ids):
    """Selecciona en el modelo los ids indicados."""
    if uidoc is None:
        return 0
    from System.Collections.Generic import List

    sel = List[ElementId]()
    for eid in element_ids or []:
        if eid is None:
            continue
        if isinstance(eid, ElementId):
            sel.Add(eid)
        else:
            try:
                sel.Add(ElementId(int(eid)))
            except Exception:
                pass
    try:
        uidoc.Selection.SetElementIds(sel)
        return int(sel.Count)
    except Exception:
        return 0
