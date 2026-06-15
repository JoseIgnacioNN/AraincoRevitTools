# -*- coding: utf-8 -*-
"""
Conversión segura Python → IList / iteración .NET (Revit 2025, CPython 3).

Evita listas Python crudas en la API Revit y fallos al iterar ``GeometryElement``.
"""

from __future__ import print_function

from System.Collections.Generic import List

from Autodesk.Revit.DB import Curve, CurveLoop, GeometryCreationUtilities, Line, XYZ


def iterate_net_collection(collection):
    """Lista Python desde cualquier IEnumerable / GeometryElement."""
    if collection is None:
        return []
    out = []
    try:
        for item in collection:
            out.append(item)
    except Exception:
        pass
    if out:
        return out
    try:
        n = int(collection.Count)
    except Exception:
        n = 0
    for i in range(n):
        item = None
        try:
            item = collection[i]
        except Exception:
            try:
                item = collection.get_Item(i)
            except Exception:
                item = None
        if item is not None:
            out.append(item)
    return out


def list_curve_from_iterable(curves):
    cl = List[Curve]()
    for c in curves or []:
        if c is not None:
            cl.Add(c)
    return cl


def list_curve_loop_from_iterable(loops):
    out = List[CurveLoop]()
    for lp in loops or []:
        if lp is not None:
            out.Add(lp)
    return out


def safe_solid_volume(solid):
    """Volumen de sólido sin propagar errores LINQ/pythonnet."""
    if solid is None:
        return None
    try:
        v = float(solid.Volume)
        if v != v:
            return None
        return v
    except Exception:
        return None


def curve_loop_create_from_lines(lines):
    """``CurveLoop`` desde líneas; ``None`` si la API rechaza el lazo."""
    cl = list_curve_from_iterable(lines)
    if cl.Count < 1:
        return None
    try:
        return CurveLoop.Create(cl)
    except Exception:
        return None


def create_vertical_square_prism_solid(px, py, z_start_ft, half_side_ft, height_ft):
    """
    Prisma cuadrado extruido en +Z (sólido de prueba colisión embed).
    Retorna ``Solid`` o ``None``.
    """
    hw = abs(float(half_side_ft))
    hgt = abs(float(height_ft))
    if hgt <= 1e-12:
        return None
    hs = XYZ(float(px), float(py), float(z_start_ft))
    p1 = XYZ(hs.X - hw, hs.Y - hw, hs.Z)
    p2 = XYZ(hs.X + hw, hs.Y - hw, hs.Z)
    p3 = XYZ(hs.X + hw, hs.Y + hw, hs.Z)
    p4 = XYZ(hs.X - hw, hs.Y + hw, hs.Z)
    loop = curve_loop_create_from_lines([
        Line.CreateBound(p1, p2),
        Line.CreateBound(p2, p3),
        Line.CreateBound(p3, p4),
        Line.CreateBound(p4, p1),
    ])
    if loop is None:
        return None
    loops = list_curve_loop_from_iterable([loop])
    if loops.Count < 1:
        return None
    try:
        sol = GeometryCreationUtilities.CreateExtrusionGeometry(
            loops, XYZ.BasisZ, hgt,
        )
    except Exception:
        return None
    if sol is None:
        return None
    try:
        if float(sol.Volume) < 1e-15:
            return None
    except Exception:
        return None
    return sol


def net_collection_count(collection):
    """Número de elementos en IList / IEnumerable."""
    if collection is None:
        return 0
    try:
        if bool(collection.IsEmpty):
            return 0
    except Exception:
        pass
    try:
        return int(collection.Count)
    except Exception:
        pass
    return len(iterate_net_collection(collection))


def as_python_list(collection):
    """Copia a ``list`` Python; IList .NET vacío no es falsy en CPython 3."""
    if collection is None:
        return []
    if isinstance(collection, list):
        return collection
    try:
        if isinstance(collection, tuple):
            return list(collection)
    except Exception:
        pass
    return list(iterate_net_collection(collection))


def list_get_or_last(collection, index, default=None):
    """Índice seguro; si ``index`` excede la lista, retorna el último elemento."""
    lst = as_python_list(collection)
    if not lst:
        return default
    try:
        i = int(index)
    except Exception:
        i = 0
    if 0 <= i < len(lst):
        return lst[i]
    return lst[-1]


def curves_from_curve_loop(cl):
    """Curvas de un ``CurveLoop`` (iteración segura)."""
    if cl is None:
        return []
    out = []
    try:
        n = int(cl.Count)
    except Exception:
        n = 0
    if n > 0:
        for j in range(n):
            try:
                c = cl.get_Item(j)
            except Exception:
                try:
                    c = cl[j]
                except Exception:
                    c = None
            if c is not None and c.IsBound and float(c.Length) > 1e-8:
                out.append(c)
        if out:
            return out
    for c in iterate_net_collection(cl):
        try:
            if c is not None and c.IsBound and float(c.Length) > 1e-8:
                out.append(c)
        except Exception:
            continue
    return out
