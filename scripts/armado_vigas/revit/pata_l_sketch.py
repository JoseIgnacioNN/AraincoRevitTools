# -*- coding: utf-8 -*-
"""
Pata L en vigas longitudinales — polilínea (``CreateFromCurves`` sin ``RebarHookType``).

Misma pauta que Armado muros: tramo recto + segmento 90° hacia el interior del host.
El largo del tramo de pata en la polilínea usa **eje = tabla − Ø/2** (``pata_eje_curve_loop_mm_desde_tabla_mm``).
"""

from __future__ import division

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import Curve, ElementId, Line, XYZ
from Autodesk.Revit.DB.Structure import MultiplanarOption, Rebar, RebarBarType, RebarHookOrientation, RebarStyle

try:
    from arearein_verticales_empotramiento_rps import (
        _copy_layout_rebar_shape_driven,
        _create_from_curves_no_hooks,
        _hook_orient_for_create,
        _rebar_normal,
        _tangent_at_end_of_curve,
        _tangent_start_curve,
    )
except Exception:
    _copy_layout_rebar_shape_driven = None
    _create_from_curves_no_hooks = None
    _hook_orient_for_create = None
    _rebar_normal = None
    _tangent_at_end_of_curve = None
    _tangent_start_curve = None

try:
    from rebar_extender_l_ganchos_135_rps import _try_create_l_from_rebar_shape_2seg
except Exception:
    _try_create_l_from_rebar_shape_2seg = None

try:
    from rebar_fundacion_cara_inferior import (
        buscar_rebar_shape_por_nombre,
        REBAR_SHAPE_NOMBRE_DEFECTO,
    )
except Exception:
    buscar_rebar_shape_por_nombre = None
    REBAR_SHAPE_NOMBRE_DEFECTO = u"03"

try:
    from geometria_empotramiento_extremos import DIAM_NOMINAL_RESPALDO_MM
except Exception:
    DIAM_NOMINAL_RESPALDO_MM = 16.0

_MM_TO_FT = 1.0 / 304.8


def _mm_to_ft(mm):
    return float(mm) * _MM_TO_FT


def _largo_pata_l_tabla_mm(meta, diam_mm):
    """Largo de pata según tabla BIMTools (mm) — usado en metadatos de extremo."""
    if meta:
        try:
            h = meta.get(u"hook_mm")
            if h is not None and float(h) > 0.1:
                return float(h)
        except Exception:
            pass
    try:
        from geometria_empotramiento_extremos import _hook_mm_desde_diametro

        return float(_hook_mm_desde_diametro(diam_mm))
    except Exception:
        pass
    try:
        d = float(diam_mm or DIAM_NOMINAL_RESPALDO_MM)
    except Exception:
        d = DIAM_NOMINAL_RESPALDO_MM
    return max(6.0 * d, 150.0)


def _largo_pata_l_eje_sketch_mm(meta, diam_mm):
    """
    Longitud de eje para polilínea ``CreateFromCurves`` (tabla − Ø/2).

    Revit modela la pata desde el eje de la barra; sin compensación la geometría
    queda ~tabla + Ø/2 (misma pauta que Armado muros / columnas).
    """
    tabla_mm = _largo_pata_l_tabla_mm(meta, diam_mm)
    try:
        d = float(diam_mm or DIAM_NOMINAL_RESPALDO_MM)
    except Exception:
        d = DIAM_NOMINAL_RESPALDO_MM
    if d <= 1e-9:
        d = DIAM_NOMINAL_RESPALDO_MM
    try:
        from bimtools_rebar_hook_lengths import pata_eje_curve_loop_mm_desde_tabla_mm

        leje = pata_eje_curve_loop_mm_desde_tabla_mm(tabla_mm, d)
        if leje is not None and float(leje) > 0.1:
            return float(leje)
    except Exception:
        pass
    try:
        d_int = float(int(round(d)))
    except Exception:
        d_int = 0.0
    if d_int > 1e-6:
        return max(40.0, float(tabla_mm) - 0.5 * d_int)
    return float(tabla_mm)


def _dir_pata_l_interior(tangent, n_face):
    """Dirección unitaria 90° al eje de la barra, hacia el interior del alma (−n exterior)."""
    if tangent is None or n_face is None:
        return None
    try:
        interior = n_face.Normalize().Negate()
        t = tangent.Normalize()
        along = t.Multiply(float(interior.DotProduct(t)))
        d = interior - along
    except Exception:
        return None
    if d.GetLength() < 1e-10:
        return None
    try:
        return d.Normalize()
    except Exception:
        return None


def _dir_pata_l_lateral_alma(tangent, w_lay, w_face_sign, n_plane):
    """
    Pata L en caras del alma: hacia el interior (±ancho). Si el eje es casi // ancho,
    dobla en canto (``n_plane``) hacia el núcleo de la sección.
    """
    n_ext = None
    try:
        n_ext = w_lay.Multiply(float(w_face_sign)).Normalize()
    except Exception:
        n_ext = w_lay
    d = _dir_pata_l_interior(tangent, n_ext)
    if d is not None:
        return d
    if tangent is None or w_lay is None:
        return None
    try:
        interior = w_lay.Normalize().Multiply(-float(w_face_sign))
    except Exception:
        return None
    try:
        t = tangent.Normalize()
        along = t.Multiply(float(interior.DotProduct(t)))
        d2 = interior - along
        if d2.GetLength() >= 1e-10:
            return d2.Normalize()
    except Exception:
        pass
    if n_plane is None:
        return None
    try:
        t = tangent.Normalize()
        np = n_plane.Normalize()
        along = t.Multiply(float(np.DotProduct(t)))
        d3 = np - along
        if d3.GetLength() < 1e-10:
            d3 = t.CrossProduct(w_lay)
        if d3.GetLength() < 1e-10:
            return None
        d3 = d3.Normalize()
        if d3.DotProduct(interior) < 0.0:
            d3 = d3.Negate()
        return d3
    except Exception:
        return None


def _leg_dir_pata(tangent, n_face, leg_dir_fn):
    if leg_dir_fn is not None:
        try:
            d = leg_dir_fn(tangent)
            if d is not None:
                return d
        except Exception:
            pass
    return _dir_pata_l_interior(tangent, n_face)


def _centerline_chain(rebar, pos_idx=0):
    if rebar is None:
        return None
    mpo = MultiplanarOption.IncludeAllMultiplanarCurves
    try:
        crvs = rebar.GetCenterlineCurves(False, False, False, mpo, int(pos_idx))
    except Exception:
        return None
    if crvs is None or int(crvs.Count) < 1:
        return None
    return [crvs[i] for i in range(crvs.Count)]


def _cadena_con_patas_l(chain, n_face, largo_inicio_mm, largo_fin_mm, leg_dir_fn=None):
    """
    Inserta segmentos de pata L al inicio y/o fin de ``chain`` (lista de ``Curve``).
    """
    if not chain:
        return None, u"Sin curvas de eje."
    out = list(chain)
    le_i = _mm_to_ft(largo_inicio_mm) if largo_inicio_mm and float(largo_inicio_mm) > 0.1 else 0.0
    le_f = _mm_to_ft(largo_fin_mm) if largo_fin_mm and float(largo_fin_mm) > 0.1 else 0.0

    if le_i > 1e-9:
        c0 = out[0]
        t0 = _tangent_start_curve(c0) if _tangent_start_curve else None
        leg_dir = _leg_dir_pata(t0, n_face, leg_dir_fn)
        if leg_dir is None:
            return None, u"Pata L inicio: dirección interior nula."
        p0 = c0.GetEndPoint(0)
        p_leg = p0 + leg_dir.Multiply(le_i)
        try:
            leg = Line.CreateBound(p_leg, p0)
        except Exception:
            return None, u"Pata L inicio: curva inválida."
        out = [leg] + out

    if le_f > 1e-9:
        c_last = out[-1]
        t1 = _tangent_at_end_of_curve(c_last) if _tangent_at_end_of_curve else None
        leg_dir = _leg_dir_pata(t1, n_face, leg_dir_fn)
        if leg_dir is None:
            return None, u"Pata L fin: dirección interior nula."
        p_end = c_last.GetEndPoint(1)
        p_tip = p_end + leg_dir.Multiply(le_f)
        try:
            leg = Line.CreateBound(p_end, p_tip)
        except Exception:
            return None, u"Pata L fin: curva inválida."
        out = out + [leg]

    if le_i <= 1e-9 and le_f <= 1e-9:
        return None, u"Sin patas L solicitadas."
    return out, None


def _create_from_curves_no_hooks_try(
    document, curves_list, host, norm, bar_type, style, o0, o1
):
    """``CreateFromCurves`` sin ganchos — varias combinaciones useExisting/createNew."""
    if document is None or not curves_list or host is None or bar_type is None:
        return None
    import clr
    import System
    from Autodesk.Revit.DB.Structure import Rebar

    ct = clr.GetClrType(Line).BaseType
    n = len(curves_list)
    arr = System.Array.CreateInstance(ct, n)
    for i in range(n):
        arr[i] = curves_list[i]
    inv = ElementId.InvalidElementId
    if style is None:
        style = RebarStyle.Standard
    for use_ex in (True, False):
        for create_new in (True, False):
            try:
                r = Rebar.CreateFromCurves(
                    document,
                    style,
                    bar_type,
                    inv,
                    inv,
                    host,
                    norm,
                    arr,
                    o0,
                    o1,
                    use_ex,
                    create_new,
                )
                if r is not None:
                    return r
            except Exception:
                continue
    if _create_from_curves_no_hooks is not None:
        try:
            return _create_from_curves_no_hooks(
                document, curves_list, host, norm, bar_type, style, o0, o1
            )
        except Exception:
            pass
    return None


def _orient_pairs_create():
    return (
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
        (RebarHookOrientation.Right, RebarHookOrientation.Right),
        (RebarHookOrientation.Left, RebarHookOrientation.Left),
    )


def _norms_unicas(norm_candidates, rebar_src):
    out = []
    seen = set()
    for nv in norm_candidates or []:
        if nv is None:
            continue
        for cand in (nv,):
            try:
                n = cand.Normalize()
            except Exception:
                n = cand
            key = None
            try:
                key = (
                    round(float(n.X), 6),
                    round(float(n.Y), 6),
                    round(float(n.Z), 6),
                )
            except Exception:
                key = unicode(n)
            if key in seen:
                continue
            seen.add(key)
            out.append(n)
            try:
                nn = n.Negate()
                key2 = (
                    round(float(nn.X), 6),
                    round(float(nn.Y), 6),
                    round(float(nn.Z), 6),
                )
                if key2 not in seen:
                    seen.add(key2)
                    out.append(nn)
            except Exception:
                pass
    if not out and _rebar_normal is not None and rebar_src is not None:
        try:
            n0 = _rebar_normal(rebar_src)
            if n0 is not None:
                out.append(n0)
        except Exception:
            pass
    if not out:
        out = [XYZ.BasisZ]
    return out


def _apply_rebar_style_if_writable(rebar, style):
    if rebar is None or style is None:
        return
    try:
        rebar.Style = style
    except Exception:
        pass


def _try_create_from_rebar_shape_named(
    document, curves_list, host, norm, bar_type, style, o0, o1, shape_name
):
    """
    Crea con un ``RebarShape`` del proyecto por nombre exacto (p. ej. «03»).
    Usado cuando la barra lleva pata L en ambos extremos.
    """
    if (
        document is None
        or not curves_list
        or host is None
        or bar_type is None
        or buscar_rebar_shape_por_nombre is None
    ):
        return None
    target = (shape_name or u"").strip()
    if not target:
        return None
    shape = buscar_rebar_shape_por_nombre(document, target)
    if shape is None:
        return None
    import System
    from System.Collections.Generic import List

    try:
        cl = List[Curve]()
        for c in curves_list:
            cl.Add(c)
    except Exception:
        return None
    orient_tries = [
        (o0, o1),
        (RebarHookOrientation.Right, RebarHookOrientation.Right),
        (RebarHookOrientation.Left, RebarHookOrientation.Left),
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
    ]
    seen_or = set()
    pairs = []
    for a in orient_tries:
        try:
            k = (int(a[0]), int(a[1]))
        except Exception:
            k = (str(a[0]), str(a[1]))
        if k not in seen_or:
            seen_or.add(k)
            pairs.append(a)
    inv = ElementId.InvalidElementId
    for so, eo in pairs:
        try:
            r = Rebar.CreateFromCurvesAndShape(
                document,
                shape,
                bar_type,
                None,
                None,
                host,
                norm,
                cl,
                so,
                eo,
                0.0,
                0.0,
                inv,
                inv,
            )
            if r is not None:
                _apply_rebar_style_if_writable(r, style)
                return r
        except Exception:
            pass
        try:
            r = Rebar.CreateFromCurvesAndShape(
                document, shape, bar_type, None, None, host, norm, cl, so, eo
            )
            if r is not None:
                _apply_rebar_style_if_writable(r, style)
                return r
        except Exception:
            pass
    return None


def _crear_rebar_desde_cadena(
    document, host, rebar_src, new_chain, norm_candidates=None, shape_name_priority=None
):
    if (
        document is None
        or host is None
        or rebar_src is None
        or not new_chain
        or _create_from_curves_no_hooks is None
    ):
        return None
    bar_type = document.GetElement(rebar_src.GetTypeId())
    if not isinstance(bar_type, RebarBarType):
        return None
    try:
        style = rebar_src.Style
    except Exception:
        style = None
    o_def = (
        (
            _hook_orient_for_create(rebar_src, 0)
            if _hook_orient_for_create
            else RebarHookOrientation.Left
        ),
        (
            _hook_orient_for_create(rebar_src, 1)
            if _hook_orient_for_create
            else RebarHookOrientation.Left
        ),
    )
    orient_pairs = [o_def]
    for pair in _orient_pairs_create():
        if pair not in orient_pairs:
            orient_pairs.append(pair)

    norms = _norms_unicas(norm_candidates, rebar_src)
    for nvec in norms:
        for o0, o1 in orient_pairs:
            new_rb = None
            if shape_name_priority:
                new_rb = _try_create_from_rebar_shape_named(
                    document,
                    new_chain,
                    host,
                    nvec,
                    bar_type,
                    style,
                    o0,
                    o1,
                    shape_name_priority,
                )
            if new_rb is not None:
                return new_rb
            if _try_create_l_from_rebar_shape_2seg is not None:
                try:
                    new_rb = _try_create_l_from_rebar_shape_2seg(
                        document, new_chain, host, nvec, bar_type, style, o0, o1
                    )
                except Exception:
                    new_rb = None
            if new_rb is not None:
                return new_rb
            try:
                new_rb = _create_from_curves_no_hooks_try(
                    document, new_chain, host, nvec, bar_type, style, o0, o1
                )
            except Exception:
                new_rb = None
            if new_rb is not None:
                return new_rb
    return None


def aplicar_patas_l_polilinea(
    document,
    rebar,
    host,
    n_face,
    gi,
    gf,
    meta_inicio=None,
    meta_fin=None,
    diam_mm=None,
    leg_dir_fn=None,
    layout_fallback=None,
    norm_candidatos=None,
):
    """
    Sustituye ``rebar`` por una polilínea con patas L en extremos con gancho (sin ``RebarHookType``).

    Debe llamarse **dentro** de la transacción del caller (no abre transacción propia).

    Returns:
        ``(rebar_resultado, mensaje_error_o_None)``
    """
    if rebar is None or document is None or host is None:
        return rebar, u"Rebar/host inválidos."
    if not gi and not gf:
        return rebar, None
    if n_face is None:
        return rebar, u"Sin normal de cara para orientar pata L."
    if _copy_layout_rebar_shape_driven is None or _create_from_curves_no_hooks is None:
        return rebar, u"Módulos de boceto pata L no disponibles."

    chain = _centerline_chain(rebar, 0)
    if not chain:
        return rebar, u"Sin centerline para pata L."

    largo_i = _largo_pata_l_eje_sketch_mm(meta_inicio, diam_mm) if gi else 0.0
    largo_f = _largo_pata_l_eje_sketch_mm(meta_fin, diam_mm) if gf else 0.0
    new_chain, err_chain = _cadena_con_patas_l(
        chain, n_face, largo_i, largo_f, leg_dir_fn=leg_dir_fn
    )
    if new_chain is None:
        return rebar, err_chain or u"Cadena L inválida."

    shape_prio = None
    if gi and gf:
        shape_prio = REBAR_SHAPE_NOMBRE_DEFECTO

    orig_id = rebar.Id
    new_rb = _crear_rebar_desde_cadena(
        document,
        host,
        rebar,
        new_chain,
        norm_candidates=norm_candidatos,
        shape_name_priority=shape_prio,
    )
    if new_rb is None:
        return rebar, u"CreateFromCurves (polilínea L) devolvió None (n_curvas={0}).".format(
            len(new_chain)
        )

    ok_lay, err_lay = _copy_layout_rebar_shape_driven(rebar, new_rb)
    if not ok_lay and layout_fallback is not None:
        try:
            ok_lay = bool(layout_fallback(new_rb))
            err_lay = u"" if ok_lay else u"Fallback de layout falló."
        except Exception as ex:
            ok_lay = False
            err_lay = unicode(ex)
    if not ok_lay:
        try:
            document.Delete(new_rb.Id)
        except Exception:
            pass
        return rebar, u"Layout pata L: {0}".format(err_lay or u"?")

    try:
        document.Delete(orig_id)
    except Exception as ex:
        try:
            document.Delete(new_rb.Id)
        except Exception:
            pass
        return rebar, u"No se pudo reemplazar rebar tras pata L: {0!s}".format(ex)

    return new_rb, None
