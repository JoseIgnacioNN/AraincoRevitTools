# -*- coding: utf-8 -*-
"""
Contorno de hormigón para Elevación Eje — detail lines agrupadas.

Revit 2024+ | pyRevit | IronPython 2.7 / 3.4

Flujo (por sección recién creada):
  1. Elementos Concrete visibles (opcionalmente precargados).
  2. Prefiltro espacial (bbox ∩ plano del eje).
  3. Unión booleana de sus sólidos.
  4. Plano de corte = plano vertical del Grid (no el de la ViewSection).
  5. ``CutWithHalfSpace`` → cara → perímetro (``GetEdgesAsCurveLoops``).
  6. ``NewDetailCurve`` en la ViewSection (estilo Medium Lines);
     grupo con el nombre de la vista.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    Category,
    ElementId,
    FilteredElementCollector,
    GraphicsStyle,
    GraphicsStyleType,
    XYZ,
)
from System.Collections.Generic import List

from contorno_hormigon_eje import (
    _as_unicode,
    _buscar_cara_corte,
    _nombre_grupo_unico,
    _proyectar_curva_a_plano,
    _vector_unitario,
    curveloops_perimetro,
    plano_corte_desde_eje,
    recoger_hormigon_en_vista,
    recopilar_nombres_grupos,
)
from contorno_hormigon_vista import unir_solidos_hormigon
from elevacion_eje_collect import plano_desde_vista

_MEDIUM_LINES_NAMES = (
    u"Medium Lines",
    u"<Medium Lines>",
    u"Líneas medias",
    u"<Líneas medias>",
)
# Holgura al prefiltrar por plano (ft): ~300 mm.
_TOL_PREFILTRO_PLANO_FT = 300.0 / 304.8
# Respaldo si no se puede leer ShortCurveTolerance (~0.8 mm).
_MIN_MODEL_CURVE_LEN_FT = 1.0 / 304.8


def _min_curve_len_ft(document):
    try:
        return max(
            float(document.Application.ShortCurveTolerance),
            _MIN_MODEL_CURVE_LEN_FT,
        )
    except Exception:
        return _MIN_MODEL_CURVE_LEN_FT


def _curva_longitud_ok(curve, min_len_ft):
    if curve is None:
        return False
    try:
        if not curve.IsBound:
            return False
    except Exception:
        return False
    try:
        return float(curve.Length) >= float(min_len_ft)
    except Exception:
        try:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            return p0.DistanceTo(p1) >= float(min_len_ft)
        except Exception:
            return False


def _nombre_vista(view):
    if view is None:
        return u""
    try:
        return _as_unicode(view.Name).strip()
    except Exception:
        return u""


def _norm_upper(text):
    return _as_unicode(text).strip().upper()


def _lines_category(document):
    if document is None:
        return None
    try:
        return Category.GetCategory(document, BuiltInCategory.OST_Lines)
    except Exception:
        pass
    try:
        return document.Settings.Categories.get_Item(BuiltInCategory.OST_Lines)
    except Exception:
        return None


def resolver_medium_lines_style_id(document):
    """``GraphicsStyle.Id`` del estilo de línea «Medium Lines» (subcategoría Líneas)."""
    if document is None:
        return None
    targets = set()
    for name in _MEDIUM_LINES_NAMES:
        targets.add(_norm_upper(name))
        bare = name.strip(u"<>").strip()
        if bare:
            targets.add(_norm_upper(bare))

    lines_cat = _lines_category(document)
    fallback_id = None
    if lines_cat is not None:
        try:
            for sub in lines_cat.SubCategories:
                try:
                    nm_u = _norm_upper(sub.Name)
                except Exception:
                    continue
                style_id = None
                try:
                    gs = sub.GetGraphicsStyle(GraphicsStyleType.Projection)
                    if gs is not None:
                        style_id = gs.Id
                except Exception:
                    style_id = None
                if style_id is None:
                    continue
                if nm_u in targets:
                    return style_id
                for t in targets:
                    bare = t.strip(u"<>").strip()
                    if bare and (nm_u == bare or bare in nm_u):
                        fallback_id = style_id
                        break
        except Exception:
            pass
    if fallback_id is not None:
        return fallback_id

    parent_iv = None
    if lines_cat is not None:
        try:
            parent_iv = lines_cat.Id.IntegerValue
        except Exception:
            parent_iv = None
    try:
        for gs in FilteredElementCollector(document).OfClass(GraphicsStyle):
            try:
                cg = getattr(gs, u"GraphicsStyleCategory", None)
                if cg is None:
                    cg = getattr(gs, u"Category", None)
                if cg is None:
                    continue
                if parent_iv is not None:
                    pc = getattr(cg, u"Parent", None)
                    if pc is None or pc.Id.IntegerValue != parent_iv:
                        continue
                nm_u = _norm_upper(cg.Name)
                if nm_u in targets:
                    return gs.Id
                for t in targets:
                    bare = t.strip(u"<>").strip()
                    if bare and (nm_u == bare or bare in nm_u):
                        return gs.Id
            except Exception:
                continue
    except Exception:
        pass
    return None


def _aplicar_line_style(curve_element, style_id, style_verificado=False):
    """
    Aplica ``style_id`` a la model line.

    Si ``style_verificado`` es True, asigna directo (ya validado una vez).
    """
    if curve_element is None or style_id is None:
        return False
    try:
        if style_id == ElementId.InvalidElementId:
            return False
    except Exception:
        pass
    if style_verificado:
        try:
            curve_element.LineStyleId = style_id
            return True
        except Exception:
            pass
    try:
        applicable = list(curve_element.GetLineStyleIds())
    except Exception:
        applicable = []
    chosen = style_id
    if applicable:
        try:
            piv = style_id.IntegerValue
            for aid in applicable:
                if aid is not None and aid.IntegerValue == piv:
                    chosen = aid
                    break
            else:
                for aid in applicable:
                    try:
                        gs = curve_element.Document.GetElement(aid)
                        nm = u""
                        if gs is not None:
                            cg = getattr(gs, u"GraphicsStyleCategory", None)
                            if cg is not None:
                                nm = _norm_upper(cg.Name)
                            if not nm:
                                nm = _norm_upper(getattr(gs, u"Name", u""))
                        for name in _MEDIUM_LINES_NAMES:
                            t = _norm_upper(name)
                            bare = t.strip(u"<>").strip()
                            if nm == t or (bare and (nm == bare or bare in nm)):
                                chosen = aid
                                break
                        else:
                            continue
                        break
                    except Exception:
                        continue
        except Exception:
            chosen = style_id
    try:
        curve_element.LineStyleId = chosen
        return True
    except Exception:
        pass
    try:
        gs = curve_element.Document.GetElement(chosen)
        if gs is not None:
            curve_element.LineStyle = gs
            return True
    except Exception:
        pass
    return False


def _esquinas_bbox(bb):
    mn, mx = bb.Min, bb.Max
    return (
        XYZ(mn.X, mn.Y, mn.Z),
        XYZ(mx.X, mn.Y, mn.Z),
        XYZ(mn.X, mx.Y, mn.Z),
        XYZ(mx.X, mx.Y, mn.Z),
        XYZ(mn.X, mn.Y, mx.Z),
        XYZ(mx.X, mn.Y, mx.Z),
        XYZ(mn.X, mx.Y, mx.Z),
        XYZ(mx.X, mx.Y, mx.Z),
    )


def _bbox_cruza_plano(bb, plane, tol_ft):
    """True si el AABB intersecta la banda del plano (±tol)."""
    if bb is None or plane is None:
        return True
    n = _vector_unitario(plane.Normal)
    if n is None:
        return True
    try:
        o = plane.Origin
        dmin = None
        dmax = None
        for p in _esquinas_bbox(bb):
            d = float(p.Subtract(o).DotProduct(n))
            if dmin is None or d < dmin:
                dmin = d
            if dmax is None or d > dmax:
                dmax = d
        if dmin is None:
            return True
        return (dmin - tol_ft) <= 0.0 and (dmax + tol_ft) >= 0.0
    except Exception:
        return True


def filtrar_elementos_cerca_plano(elementos, plane, tol_ft=_TOL_PREFILTRO_PLANO_FT):
    """Conserva elementos cuyo bbox cruza (o roza) el plano de corte."""
    if plane is None:
        return list(elementos or [])
    out = []
    for el in elementos or []:
        if el is None:
            continue
        try:
            bb = el.get_BoundingBox(None)
        except Exception:
            bb = None
        if bb is None or _bbox_cruza_plano(bb, plane, tol_ft):
            out.append(el)
    return out


def crear_detail_lines_y_grupo(
    document,
    view,
    curve_loops,
    group_name_base,
    style_id=None,
    nombres_grupos=None,
):
    """
    Crea DetailCurves en ``view`` (estilo Medium Lines) y las agrupa.

    Returns:
        dict con ``line_count``, ``group_name``, ``loops``, ``style_id``.
    """
    if document is None or view is None:
        raise ValueError(u"Documento o vista no válidos para detail lines.")

    plane_vista = plano_desde_vista(view)
    if plane_vista is None:
        raise ValueError(u"No se pudo obtener el plano de la vista.")

    if style_id is None:
        style_id = resolver_medium_lines_style_id(document)
    min_len = _min_curve_len_ft(document)
    style_ok = False
    line_ids = []
    creadas = 0
    for cl in curve_loops or []:
        if cl is None:
            continue
        for c in cl:
            if c is None:
                continue
            try:
                if not c.IsBound:
                    continue
            except Exception:
                continue
            curva = _proyectar_curva_a_plano(c, plane_vista)
            if curva is None:
                curva = c
            if not _curva_longitud_ok(curva, min_len):
                continue
            try:
                dc = document.Create.NewDetailCurve(view, curva)
            except Exception:
                curva2 = _proyectar_curva_a_plano(c, plane_vista)
                if curva2 is None or not _curva_longitud_ok(curva2, min_len):
                    continue
                try:
                    dc = document.Create.NewDetailCurve(view, curva2)
                except Exception:
                    continue
            if dc is not None:
                if _aplicar_line_style(dc, style_id, style_verificado=style_ok):
                    style_ok = True
                line_ids.append(dc.Id)
                creadas += 1

    group_name = _nombre_grupo_unico(
        document, group_name_base, nombres_existentes=nombres_grupos,
    )
    if line_ids:
        ids = List[ElementId]()
        for eid in line_ids:
            ids.Add(eid)
        grp = document.Create.NewGroup(ids)
        gt = document.GetElement(grp.GroupType.Id)
        gt.Name = group_name

    return {
        u"line_count": creadas,
        u"group_name": group_name if line_ids else None,
        u"loops": len(curve_loops or []),
        u"style_id": style_id,
    }


def _nombre_eje_grid(grid):
    if grid is None:
        return u"Eje"
    try:
        nombre = _as_unicode(grid.Name).strip()
    except Exception:
        nombre = u""
    if nombre:
        return nombre
    try:
        return u"Id {0}".format(grid.Id.IntegerValue)
    except Exception:
        return u"Eje"


def generar_contorno_model_lines(
    document,
    view_section,
    grid,
    group_name=None,
    style_id=None,
    nombres_grupos=None,
    regenerate=False,
    elementos=None,
):
    """
    Genera el contorno de hormigón de ``view_section``.

    Debe llamarse dentro de una Transaction abierta.
    El crop de la vista debe estar activo (lo gestiona el orquestador).

    Args:
        document: Document
        view_section: ViewSection creada
        grid: Grid asociado (nombre / respaldo de plano)
        group_name: nombre base del grupo (por defecto ``view.Name``)
        style_id: GraphicsStyle Medium Lines precacheado
        nombres_grupos: set mutable de nombres de GroupType
        regenerate: si True, regenera el documento antes de colectar
        elementos: lista precargada de hosts Concrete (evita re-colectar)

    Returns:
        (True, dict_resultado) o (False, mensaje_error)
    """
    if document is None or view_section is None:
        return False, u"Vista de sección no válida."

    if regenerate:
        try:
            document.Regenerate()
        except Exception:
            pass

    if elementos is None:
        elementos = recoger_hormigon_en_vista(document, view_section)
    if not elementos:
        return False, (
            u"No hay elementos con Material for Model Behavior = Concrete "
            u"visibles en la sección."
        )

    # Corte geométrico: plano del eje (Grid). La vista solo proyecta detail lines.
    plane, nombre_eje_pl = plano_corte_desde_eje(grid, view_section)
    if plane is None:
        return False, _as_unicode(nombre_eje_pl) or (
            u"No se pudo obtener el plano de corte desde el eje."
        )
    nombre_eje = _as_unicode(nombre_eje_pl).strip() or _nombre_eje_grid(grid)

    elementos = filtrar_elementos_cerca_plano(elementos, plane)
    if not elementos:
        return False, (
            u"Ningún elemento de hormigón intersecta el plano del eje «{0}»."
        ).format(_as_unicode(nombre_eje))

    solido = unir_solidos_hormigon(elementos, view_section)
    if solido is None:
        return False, u"No se pudo unir la geometría sólida del hormigón."

    preferred = None
    try:
        preferred = view_section.ViewDirection
    except Exception:
        preferred = None

    origin = plane.Origin
    cara = _buscar_cara_corte(solido, plane, origin, preferred_normal=preferred)
    if cara is None:
        return False, (
            u"El plano del eje «{0}» no produce una sección válida "
            u"sobre el sólido unificado."
        ).format(_as_unicode(nombre_eje))

    loops = curveloops_perimetro(cara)
    if not loops:
        return False, u"No se obtuvieron curvas de perímetro en la sección."

    base_name = _as_unicode(group_name).strip() if group_name else u""
    if not base_name:
        base_name = _nombre_vista(view_section)
    if not base_name:
        base_name = u"CONTORNO {0}".format(_as_unicode(nombre_eje))

    try:
        resultado = crear_detail_lines_y_grupo(
            document,
            view_section,
            loops,
            base_name,
            style_id=style_id,
            nombres_grupos=nombres_grupos,
        )
    except Exception as ex:
        return False, _as_unicode(ex)

    if not resultado.get(u"line_count"):
        return False, u"No se pudo crear ninguna detail line del contorno."

    resultado[u"concrete_count"] = len(elementos)
    resultado[u"eje"] = _as_unicode(nombre_eje)
    return True, resultado
