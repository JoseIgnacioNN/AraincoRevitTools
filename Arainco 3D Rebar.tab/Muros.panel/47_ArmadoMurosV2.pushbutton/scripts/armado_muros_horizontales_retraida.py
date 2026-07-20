# -*- coding: utf-8 -*-
"""
Post-proceso barras horizontales ext/int (Armado muros lineales / mallas).

1. Estiramiento negativo en **ambos** extremos del eje (interior y exterior):
   - **Extremo inicial** (startpoint): ``25 mm + Ø/2``.
   - **Extremo final** (endpoint): ``25 mm`` fijos.
2. ``RebarShape`` **«06»** (polilínea en L; ganchos 135° estilo stirrup/tie en la forma):
   - **Interior** → pata en startpoint.
   - **Exterior** → pata en endpoint.
   Largo pata = espesor muro − 50 mm. Gancho: **«Stirrup/Tie - 135 deg.»** (o equivalente ~135°).
"""

from __future__ import print_function

import math
import os
import sys

import clr

clr.AddReference("RevitAPI")

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    Line,
    Transaction,
    UnitUtils,
    UnitTypeId,
    Wall,
)
from Autodesk.Revit.DB.Structure import (
    MultiplanarOption,
    Rebar,
    RebarBarType,
    RebarHookOrientation,
    RebarHookType,
    RebarShape,
    RebarStyle,
)

HORIZONTAL_RETRACT_BASE_MM = 25.0
HORIZONTAL_REBAR_SHAPE_NOMBRE = u"06"
HORIZONTAL_HOOK_TYPE_NOMBRE = u"Stirrup/Tie - 135 deg."
PLAN_PREFER_X_FOR_HORIZONTAL = True
HORIZONTAL_MAX_ABS_TZ = 0.02


def _ensure_pushbutton_path():
    try:
        import bootstrap_paths
        return bootstrap_paths.pin_local_scripts_first()
    except Exception:
        d = os.path.dirname(os.path.abspath(__file__))
        if d and d not in sys.path:
            sys.path.insert(0, d)
        return d


_ensure_pushbutton_path()

try:
    from armado_muros_nodo_shared import ajustar_inclusion_extremos_rebar_set
except Exception:
    ajustar_inclusion_extremos_rebar_set = None

try:
    from arearein_verticales_empotramiento_rps import (
        _copy_layout_rebar_shape_driven,
        _create_from_curves_no_hooks,
        _hook_orient_for_create,
        _nominal_diameter_mm_from_rebar,
        _rebar_es_vertical_por_criterio,
        _rebar_normal,
        _tangent_at_end_of_curve,
        _tangent_start_curve,
    )
except Exception:
    _copy_layout_rebar_shape_driven = None
    _create_from_curves_no_hooks = None
    _hook_orient_for_create = None
    _nominal_diameter_mm_from_rebar = None
    _rebar_es_vertical_por_criterio = None
    _rebar_normal = None
    _tangent_at_end_of_curve = None
    _tangent_start_curve = None

try:
    from arearein_exterior_h_l135_rps import (
        _rebar_horizontal_en_plano,
        _rebar_solo_cara_exterior,
    )
except Exception:
    _rebar_horizontal_en_plano = None
    _rebar_solo_cara_exterior = None

try:
    from arearein_interior_h_l135_rps import _rebar_solo_cara_interior
except Exception:
    _rebar_solo_cara_interior = None

try:
    import rebar_extender_l_ganchos_135_rps as l135
except Exception:
    l135 = None

try:
    from rebar_fundacion_cara_inferior import (
        buscar_rebar_hook_type_por_nombre,
        buscar_rebar_shape_por_nombre,
    )
except Exception:
    buscar_rebar_hook_type_por_nombre = None
    buscar_rebar_shape_por_nombre = None


def _mm_to_internal(mm):
    return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)


def _element_id_int(eid):
    if eid is None:
        return None
    try:
        return int(eid.Value)
    except Exception:
        try:
            return int(eid.IntegerValue)
        except Exception:
            return None


def _retract_mm_extremo_inicial(d_mm):
    """Acortamiento en el extremo inicial del trazado: 25 mm + mitad del Ø nominal."""
    return float(HORIZONTAL_RETRACT_BASE_MM) + float(d_mm) / 2.0


def _retract_mm_extremo_final():
    """Acortamiento en el extremo final del trazado: solo 25 mm."""
    return float(HORIZONTAL_RETRACT_BASE_MM)


def _shape_label(sh):
    if sh is None:
        return u""
    try:
        t = (sh.Name or u"").strip()
        if t:
            return t
    except Exception:
        pass
    try:
        p = sh.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p is not None and p.HasValue:
            return (p.AsString() or u"").strip()
    except Exception:
        pass
    return u""


def _shape_key_matches_06(label):
    if not label:
        return False
    s = label.replace(u"\u00A0", u" ").strip()
    if s == HORIZONTAL_REBAR_SHAPE_NOMBRE:
        return True
    digits = u"".join(ch for ch in s if ch in u"0123456789")
    return digits == HORIZONTAL_REBAR_SHAPE_NOMBRE


def _resolver_rebar_shape_06(doc):
    if doc is None:
        return None
    cache = getattr(_resolver_rebar_shape_06, u"_cache", None)
    if cache is None:
        _resolver_rebar_shape_06._cache = {}
        cache = _resolver_rebar_shape_06._cache
    try:
        dkey = int(doc.GetHashCode())
    except Exception:
        dkey = id(doc)
    if dkey in cache:
        return cache[dkey]
    found = None
    if buscar_rebar_shape_por_nombre is not None:
        try:
            sh = buscar_rebar_shape_por_nombre(doc, HORIZONTAL_REBAR_SHAPE_NOMBRE)
            if sh is not None:
                found = sh
        except Exception:
            pass
    if found is None:
        try:
            for sh in FilteredElementCollector(doc).OfClass(RebarShape):
                if _shape_key_matches_06(_shape_label(sh)):
                    found = sh
                    break
        except Exception:
            pass
    cache[dkey] = found
    return found


def _hook_name_variants(nombre):
    n0 = (nombre or u"").strip()
    if not n0:
        return []
    out = [n0]
    if n0.endswith(u"."):
        out.append(n0[:-1].strip())
    else:
        out.append(n0 + u".")
    seen = set()
    uniq = []
    for n in out:
        if n and n not in seen:
            seen.add(n)
            uniq.append(n)
    return uniq


def _resolver_hook_stirrup_tie_135(doc):
    if doc is None:
        return None
    cache = getattr(_resolver_hook_stirrup_tie_135, u"_cache", None)
    if cache is None:
        _resolver_hook_stirrup_tie_135._cache = {}
        cache = _resolver_hook_stirrup_tie_135._cache
    try:
        dkey = int(doc.GetHashCode())
    except Exception:
        dkey = id(doc)
    if dkey in cache:
        return cache[dkey]
    found = None
    if buscar_rebar_hook_type_por_nombre is not None:
        for nombre in _hook_name_variants(HORIZONTAL_HOOK_TYPE_NOMBRE):
            try:
                ht = buscar_rebar_hook_type_por_nombre(doc, nombre)
                if ht is not None:
                    found = ht
                    break
            except Exception:
                pass
    if found is None:
        target_deg = 135.0
        tol_deg = 2.0
        stirrup_cands = []
        try:
            for ht in FilteredElementCollector(doc).OfClass(RebarHookType):
                name = u""
                try:
                    name = (ht.Name or u"").lower()
                except Exception:
                    pass
                try:
                    ang = math.degrees(float(ht.HookAngle))
                except Exception:
                    ang = None
                if ang is None or abs(ang - target_deg) > tol_deg:
                    continue
                if u"stirrup" in name or u"tie" in name:
                    stirrup_cands.append(ht)
        except Exception:
            pass
        if stirrup_cands:
            found = stirrup_cands[0]
    if found is None and l135 is not None:
        try:
            hid, _err = l135._resolve_rebar_hook_135_id(doc, 100.0)
            if hid is not None and hid != ElementId.InvalidElementId:
                ht = doc.GetElement(hid)
                if isinstance(ht, RebarHookType):
                    found = ht
        except Exception:
            pass
    cache[dkey] = found
    return found


def _curves_to_list(curves_list):
    cl = List[object]()
    for c in curves_list:
        cl.Add(c)
    return cl


def _crear_rebar_shape_06_desde_cadena(
    doc, curves_list, host, norm, bar_type, style, o0, o1,
):
    """Crea ``Rebar`` con forma «06» y gancho stirrup/tie 135° (explícito o definido por forma)."""
    if doc is None or not curves_list or host is None or bar_type is None:
        return None, u"Parámetros incompletos."
    shape = _resolver_rebar_shape_06(doc)
    if shape is None:
        return None, u"No se encontró RebarShape «{0}» en el proyecto.".format(
            HORIZONTAL_REBAR_SHAPE_NOMBRE,
        )
    hook = _resolver_hook_stirrup_tie_135(doc)
    try:
        cl = _curves_to_list(curves_list)
    except Exception as ex:
        return None, u"Curvas: {0!s}".format(ex)
    if cl is None or int(cl.Count) < 1:
        return None, u"Lista de curvas vacía."

    norms = [norm]
    try:
        norms.append(norm.Negate())
    except Exception:
        pass
    orient_pairs = [
        (o0, o1),
        (RebarHookOrientation.Right, RebarHookOrientation.Right),
        (RebarHookOrientation.Left, RebarHookOrientation.Left),
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
    ]
    seen_or = set()
    pairs = []
    for a in orient_pairs:
        try:
            k = (int(a[0]), int(a[1]))
        except Exception:
            k = (str(a[0]), str(a[1]))
        if k not in seen_or:
            seen_or.add(k)
            pairs.append(a)

    hook_trials = []
    if hook is not None:
        hook_trials.append((hook, hook))
    hook_trials.append((None, None))
    invalid = ElementId.InvalidElementId
    last_err = None

    for nvec in norms:
        if nvec is None:
            continue
        for h0, h1 in hook_trials:
            for so, eo in pairs:
                try:
                    if h0 is None and h1 is None:
                        r = Rebar.CreateFromCurvesAndShape(
                            doc,
                            shape,
                            bar_type,
                            None,
                            None,
                            host,
                            nvec,
                            cl,
                            so,
                            eo,
                            0.0,
                            0.0,
                            invalid,
                            invalid,
                        )
                    else:
                        r = Rebar.CreateFromCurvesAndShape(
                            doc,
                            shape,
                            bar_type,
                            h0,
                            h1,
                            host,
                            nvec,
                            cl,
                            so,
                            eo,
                            0.0,
                            0.0,
                            invalid,
                            invalid,
                        )
                    if r is not None:
                        try:
                            if style is not None:
                                r.Style = style
                        except Exception:
                            pass
                        return r, None
                except Exception as ex:
                    last_err = ex
                try:
                    if h0 is None and h1 is None:
                        r = Rebar.CreateFromCurvesAndShape(
                            doc, shape, bar_type, None, None, host, nvec, cl, so, eo,
                        )
                    else:
                        r = Rebar.CreateFromCurvesAndShape(
                            doc, shape, bar_type, h0, h1, host, nvec, cl, so, eo,
                        )
                    if r is not None:
                        try:
                            if style is not None:
                                r.Style = style
                        except Exception:
                            pass
                        return r, None
                except Exception as ex:
                    last_err = ex

    msg = u"CreateFromCurvesAndShape (forma «{0}») falló.".format(
        HORIZONTAL_REBAR_SHAPE_NOMBRE,
    )
    if last_err is not None:
        msg += u" {0!s}".format(last_err)
    return None, msg


def _rebar_es_horizontal_cara_ext_o_int(rebar, host_wall):
    if rebar is None or host_wall is None or not isinstance(host_wall, Wall):
        return False, None
    try:
        import armado_muros_lineales as _lin_malla
        if _lin_malla._rebar_es_vertical_en_muro(rebar, host_wall):
            return False, None
    except Exception:
        if _rebar_es_vertical_por_criterio is not None:
            try:
                if _rebar_es_vertical_por_criterio(rebar, host_wall, 0):
                    return False, None
            except Exception:
                pass
    if _rebar_horizontal_en_plano is None:
        return False, None
    try:
        if not _rebar_horizontal_en_plano(
            rebar,
            PLAN_PREFER_X_FOR_HORIZONTAL,
            HORIZONTAL_MAX_ABS_TZ,
            host_wall,
        ):
            return False, None
    except Exception:
        return False, None
    if _rebar_solo_cara_exterior is not None and _rebar_solo_cara_exterior(rebar, host_wall):
        return True, u"exterior"
    if _rebar_solo_cara_interior is not None and _rebar_solo_cara_interior(rebar, host_wall):
        return True, u"interior"
    return False, None


def _excluir_solo_ultima_barra_horizontal(doc, rebar):
    if doc is None or rebar is None or ajustar_inclusion_extremos_rebar_set is None:
        return rebar
    try:
        ajustar_inclusion_extremos_rebar_set(rebar, doc, True, False)
    except Exception:
        pass
    try:
        rb = doc.GetElement(rebar.Id)
        if rb is not None and isinstance(rb, Rebar):
            return rb
    except Exception:
        pass
    return rebar


def _copy_layout_y_excluir_ultima_barra(doc, src, dst):
    if _copy_layout_rebar_shape_driven is None:
        return False, u"Copy layout no disponible.", dst
    ok_lay, err_lay = _copy_layout_rebar_shape_driven(src, dst)
    if not ok_lay:
        return False, err_lay or u"?", dst
    dst = _excluir_solo_ultima_barra_horizontal(doc, dst)
    return True, u"", dst


def _acortar_rebar_eje_inicio_fin_mm(doc, rebar, mm_inicio, mm_fin, pos_idx=0):
    """Acorta el eje en inicio y/o fin (mm positivos hacia el interior del trazado)."""
    if doc is None or rebar is None:
        return False, u"Doc o rebar no válido.", None
    la = max(0.0, float(mm_inicio))
    lb = max(0.0, float(mm_fin))
    if la < 0.1 and lb < 0.1:
        return True, u"Retiro 0 mm, sin cambio.", rebar
    if (
        _tangent_start_curve is None
        or _tangent_at_end_of_curve is None
        or _create_from_curves_no_hooks is None
        or _copy_layout_rebar_shape_driven is None
    ):
        return False, u"Módulo de acortamiento no disponible.", None

    mpo = MultiplanarOption.IncludeAllMultiplanarCurves
    try:
        crvs = rebar.GetCenterlineCurves(False, False, False, mpo, int(pos_idx))
    except Exception as ex:
        return False, u"GetCenterlineCurves: {0!s}".format(ex), None
    if crvs is None or int(crvs.Count) < 1:
        return False, u"Sin curvas de eje (pos. {0}).".format(pos_idx), None

    chain = [crvs[i] for i in range(crvs.Count)]
    c_first = chain[0]
    c_last = chain[-1]
    t0 = _tangent_start_curve(c_first)
    t1 = _tangent_at_end_of_curve(c_last)
    if t0 is None or t1 is None:
        return False, u"Tangente nula (geometría de eje).", None

    p0s = c_first.GetEndPoint(0)
    p0e = c_first.GetEndPoint(1)
    p1s = c_last.GetEndPoint(0)
    p1e = c_last.GetEndPoint(1)
    a = _mm_to_internal(la)
    b = _mm_to_internal(lb)
    new_p0 = p0s + t0.Multiply(a) if la >= 0.1 else p0s
    new_p1e = p1e - t1.Multiply(b) if lb >= 0.1 else p1e

    if int(crvs.Count) == 1:
        try:
            if new_p0.DistanceTo(new_p1e) < 1e-6:
                return False, u"Retiro deja barra nula.", None
        except Exception:
            pass
        c_new = Line.CreateBound(new_p0, new_p1e)
        new_chain = [c_new]
    else:
        c_first_new = Line.CreateBound(new_p0, p0e)
        c_last_new = Line.CreateBound(p1s, new_p1e)
        new_chain = [c_first_new] + chain[1:-1] + [c_last_new]
        for i in range(len(new_chain) - 1):
            e_prev = new_chain[i].GetEndPoint(1)
            s_next = new_chain[i + 1].GetEndPoint(0)
            gap = (e_prev - s_next).GetLength()
            if gap > 0.01:
                return (
                    False,
                    u"Polilínea no consecutiva (gap {0:,.4f} ft).".format(gap),
                    None,
                )

    host = doc.GetElement(rebar.GetHostId())
    if host is None:
        return False, u"Host inválido.", None
    bar_type = doc.GetElement(rebar.GetTypeId())
    if not isinstance(bar_type, RebarBarType):
        return False, u"RebarBarType no resuelto.", None
    try:
        style = rebar.Style
    except Exception:
        style = RebarStyle.Standard
    norm = _rebar_normal(rebar)
    o0 = _hook_orient_for_create(rebar, 0)
    o1 = _hook_orient_for_create(rebar, 1)

    from armado_muros_txn import TxnScope

    scope = TxnScope(
        doc, u"Arainco: Armado muros lineales — retraer horizontal ext/int",
    )
    try:
        new_rb = _create_from_curves_no_hooks(
            doc, new_chain, host, norm, bar_type, style, o0, o1,
        )
        if new_rb is None:
            scope.rollback()
            return False, u"CreateFromCurves devolvió None.", None
        ok_lay, err_lay, new_rb = _copy_layout_y_excluir_ultima_barra(doc, rebar, new_rb)
        if not ok_lay:
            scope.rollback()
            return False, u"Layout: {0}".format(err_lay or u"?"), None
        try:
            doc.Delete(rebar.Id)
        except Exception as ex2:
            scope.rollback()
            return False, u"Delete rebar: {0!s}".format(ex2), None
        try:
            from armado_muros_rebar_params import stamp_malla_horizontal_rebar
            stamp_malla_horizontal_rebar(new_rb)
        except Exception:
            pass
        scope.commit()
    except Exception as ex:
        scope.rollback()
        return False, u"{0!s}".format(ex), None
    return (
        True,
        u"Retiro inicio {0} mm, fin {1} mm; nuevo id {2}.".format(
            int(round(la)), int(round(lb)), _element_id_int(getattr(new_rb, "Id", None)),
        ),
        new_rb,
    )


def _largo_pata_l_mm(doc, host):
    if l135 is None or doc is None or host is None:
        return None
    try:
        return float(l135.largo_pata_mm_desde_espesor_host(doc, host))
    except Exception:
        return None


def _agregar_pata_l_extremo_sketch(
    doc,
    rebar,
    host,
    largo_p_mm,
    pata_en_final,
    invertir,
    txn_name,
    pos_idx=0,
):
    """Polilínea en L con ``RebarShape`` «06» y ganchos stirrup/tie 135°."""
    if l135 is None or doc is None or rebar is None or host is None:
        return False, u"Módulo geometría L no disponible.", None
    if _copy_layout_rebar_shape_driven is None:
        return False, u"Copy layout no disponible.", None
    if largo_p_mm is None or float(largo_p_mm) < 0.1:
        return False, u"Largo pata L inválido.", None

    mpo = MultiplanarOption.IncludeAllMultiplanarCurves
    try:
        crvs = rebar.GetCenterlineCurves(False, False, False, mpo, int(pos_idx))
    except Exception as ex:
        return False, u"GetCenterlineCurves: {0!s}".format(ex), None
    if crvs is None or int(crvs.Count) < 1:
        return False, u"Sin curvas de eje.", None

    chain = [crvs[i] for i in range(crvs.Count)]
    le = _mm_to_internal(float(largo_p_mm))
    norm = l135._rebar_normal(rebar)
    if bool(getattr(l135, u"INVERTIR_NORMAL_REBAR", False)):
        norm = norm.Negate()

    if not pata_en_final:
        c0 = chain[0]
        t_vec = l135._tangent_start_first_curve(c0)
        if t_vec is None:
            return False, u"Tangente nula en inicio.", None
        b_vec = l135._perp_in_plane(norm, t_vec)
        if b_vec is None:
            return False, u"Perpendicular in-plane nula.", None
        if invertir:
            b_vec = b_vec.Negate()
        p0 = c0.GetEndPoint(0)
        p_leg = p0 - b_vec.Multiply(le)
        leg = Line.CreateBound(p_leg, p0)
        new_chain = [leg] + chain
    else:
        c_last = chain[-1]
        t_vec = l135._tangent_start_first_curve(c_last)
        if t_vec is None:
            return False, u"Tangente nula en extremo.", None
        b_vec = l135._perp_in_plane(norm, t_vec)
        if b_vec is None:
            return False, u"Perpendicular in-plane nula.", None
        if invertir:
            b_vec = b_vec.Negate()
        p_end = c_last.GetEndPoint(1)
        p_tip = p_end - b_vec.Multiply(le)
        leg = Line.CreateBound(p_end, p_tip)
        new_chain = chain + [leg]

    bar_type = doc.GetElement(rebar.GetTypeId())
    if not isinstance(bar_type, RebarBarType):
        return False, u"RebarBarType no resuelto.", None
    try:
        style = rebar.Style
    except Exception:
        style = RebarStyle.Standard
    o0 = _hook_orient_for_create(rebar, 0)
    o1 = _hook_orient_for_create(rebar, 1)
    orig_id = rebar.Id

    from armado_muros_txn import TxnScope

    scope = TxnScope(doc, txn_name)
    try:
        new_rb, err_shape = _crear_rebar_shape_06_desde_cadena(
            doc, new_chain, host, norm, bar_type, style, o0, o1,
        )
        if new_rb is None:
            scope.rollback()
            return False, err_shape or u"Forma 06 no creada.", None
        ok_lay, err_lay, new_rb = _copy_layout_y_excluir_ultima_barra(doc, rebar, new_rb)
        if not ok_lay:
            scope.rollback()
            return False, u"Layout: {0}".format(err_lay or u"?"), None
        try:
            doc.Delete(orig_id)
        except Exception as ex2:
            scope.rollback()
            return False, u"Delete rebar: {0!s}".format(ex2), None
        try:
            from armado_muros_rebar_params import stamp_malla_horizontal_rebar
            stamp_malla_horizontal_rebar(new_rb)
        except Exception:
            pass
        scope.commit()
    except Exception as ex:
        scope.rollback()
        return False, u"{0!s}".format(ex), None

    return (
        True,
        u"Forma «{0}» (pata {1} mm); nuevo id {2}.".format(
            HORIZONTAL_REBAR_SHAPE_NOMBRE,
            int(round(float(largo_p_mm))),
            _element_id_int(getattr(new_rb, "Id", None)),
        ),
        new_rb,
    )


def _aplicar_pata_l_horizontal(doc, rebar, host, cara, res):
    largo_p = _largo_pata_l_mm(doc, host)
    if largo_p is None or float(largo_p) < 0.1:
        res[u"n_fail"] += 1
        rid = _element_id_int(getattr(rebar, "Id", None))
        res[u"messages"].append(
            u"Rebar {0} (pata L horizontal {1}): largo inválido.".format(rid, cara or u"?"),
        )
        return rebar

    if cara == u"interior":
        pata_en_final = False
        invertir = not bool(getattr(l135, u"INVERTIR_DIRECCION_PATA", False))
        txn = u"Arainco: Armado muros lineales — horizontal forma 06 interior"
    elif cara == u"exterior":
        pata_en_final = True
        invertir = bool(getattr(l135, u"INVERTIR_DIRECCION_PATA", False))
        txn = u"Arainco: Armado muros lineales — horizontal forma 06 exterior"
    else:
        return rebar

    ok_l, msg_l, rb_l = _agregar_pata_l_extremo_sketch(
        doc,
        rebar,
        host,
        float(largo_p),
        pata_en_final,
        invertir,
        txn,
        0,
    )
    if ok_l:
        res[u"n_horiz_pata_l"] += 1
        if cara == u"exterior":
            res[u"n_horiz_pata_l_ext"] += 1
        elif cara == u"interior":
            res[u"n_horiz_pata_l_int"] += 1
        return rb_l if rb_l is not None else rebar

    res[u"n_fail"] += 1
    rid = _element_id_int(getattr(rebar, "Id", None))
    res[u"messages"].append(
        u"Rebar {0} (pata L horizontal {1}): {2}".format(rid, cara or u"?", msg_l or u"error"),
    )
    return rebar


def _registrar_rebar_malla_horizontal(res, wid, rebar):
    """IDs de malla horizontal (identificados antes de forma 06)."""
    if res is None or rebar is None:
        return
    try:
        wkey = int(wid)
    except Exception:
        return
    rid = _element_id_int(getattr(rebar, u"Id", None))
    if rid is None:
        return
    mp = res.setdefault(u"rebars_malla_horizontal_por_muro_id", {})
    lst = mp.setdefault(wkey, [])
    if int(rid) not in lst:
        lst.append(int(rid))


def _quitar_rebar_del_registro_horizontal(res, wid, rebar_id):
    if res is None or rebar_id is None:
        return
    try:
        wkey = int(wid)
        rid = int(rebar_id)
    except Exception:
        return
    mp = res.get(u"rebars_malla_horizontal_por_muro_id") or {}
    lst = mp.get(wkey)
    if not lst:
        return
    mp[wkey] = [x for x in lst if int(x) != rid]


def _procesar_rebar_horizontal_retraida(doc, rebar, host, res, wid=None):
    ok_cara, cara = _rebar_es_horizontal_cara_ext_o_int(rebar, host)
    if not ok_cara:
        res[u"n_skip"] += 1
        return rebar

    old_rid = _element_id_int(getattr(rebar, u"Id", None))
    d_mm = _nominal_diameter_mm_from_rebar(rebar, doc)
    if d_mm is None or float(d_mm) <= 0.0:
        res[u"n_skip"] += 1
        return rebar

    mm_inicio = _retract_mm_extremo_inicial(d_mm)
    mm_fin = _retract_mm_extremo_final()
    ok, msg, rb_out = _acortar_rebar_eje_inicio_fin_mm(
        doc, rebar, mm_inicio, mm_fin, 0,
    )
    if not ok:
        res[u"n_fail"] += 1
        rid = _element_id_int(getattr(rebar, "Id", None))
        res[u"messages"].append(
            u"Rebar {0} (horizontal {1}): {2}".format(rid, cara or u"?", msg or u"error"),
        )
        return rebar

    res[u"n_horiz_retract"] += 1
    if cara == u"exterior":
        res[u"n_horiz_retract_ext"] += 1
    elif cara == u"interior":
        res[u"n_horiz_retract_int"] += 1
    rb_work = rb_out if rb_out is not None else rebar
    rb_final = _aplicar_pata_l_horizontal(doc, rb_work, host, cara, res)
    if wid is not None:
        new_rid = _element_id_int(getattr(rb_final, u"Id", None))
        if old_rid is not None and new_rid is not None and old_rid != new_rid:
            _quitar_rebar_del_registro_horizontal(res, wid, old_rid)
        _registrar_rebar_malla_horizontal(res, wid, rb_final)
    return rb_final


def aplicar_retraida_horizontales_ext_int(doc, walls, rebars_por_muro_id):
    """
    Post-proceso horizontal ext/int: retraída en ambos extremos + ``RebarShape`` «06».

    Inicio: ``25 mm + Ø/2``; fin: ``25 mm``. Forma 06 (L + ganchos stirrup/tie 135°):
    interior en startpoint, exterior en endpoint.
    """
    res = {
        u"n_horiz_retract": 0,
        u"n_horiz_retract_ext": 0,
        u"n_horiz_retract_int": 0,
        u"n_horiz_pata_l": 0,
        u"n_horiz_pata_l_ext": 0,
        u"n_horiz_pata_l_int": 0,
        u"n_skip": 0,
        u"n_fail": 0,
        u"messages": [],
    }
    if doc is None or not walls or not rebars_por_muro_id:
        return res
    if (
        _nominal_diameter_mm_from_rebar is None
        or _rebar_horizontal_en_plano is None
        or l135 is None
        or (_rebar_solo_cara_exterior is None and _rebar_solo_cara_interior is None)
    ):
        res[u"messages"].append(
            u"Post-proceso horizontal ext/int: módulos locales no disponibles.",
        )
        return res

    wall_by_id = {}
    for w in walls:
        try:
            _wi = _element_id_int(w.Id)
            if _wi is not None:
                wall_by_id[int(_wi)] = w
        except Exception:
            pass

    # Sin Transaction envolvente: cada retraída/pata L abre Transaction propia (como V1).
    for wid, eid_list in rebars_por_muro_id.items():
        host = wall_by_id.get(int(wid))
        if host is None:
            continue
        for idx, eid in enumerate(eid_list or []):
            rebar = doc.GetElement(eid)
            if rebar is None or not isinstance(rebar, Rebar):
                res[u"n_skip"] += 1
                continue
            rebar_work = _procesar_rebar_horizontal_retraida(
                doc, rebar, host, res, wid=wid,
            )
            try:
                eid_list[idx] = rebar_work.Id
            except Exception:
                pass

    return res
