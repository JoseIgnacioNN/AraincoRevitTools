# -*- coding: utf-8 -*-
"""
1) Seleccionar un **Area Reinforcement**.
2) Quitar el sistema de refuerzo por área (``RebarInSystem`` → ``Rebar``) con
   ``AreaReinforcement.RemoveAreaReinforcementSystem``.
3) Entre las barras creadas, tratar la **cara exterior o la interior** (muro/losa: mismos
   criterios que ``arearein_exterior_h_l135_rps`` y ``arearein_interior_h_l135_rps``) **y** que
   sean **verticales** (muro: eje Z dominante; losa: no «horizontal en planta» con el criterio
   X/Y de BIMTools).
4) A cada barra que cumple ambos criterios añade una **pata en 90°** al **fin** del boceto.
   En **muro** el sentido in-plane se ajusta hacia el **núcleo** (no hacia la cara
   exterior) vía ``_B_pata_90_hacia_nucleo_muro``. En **losa** u host no muro, en
   **cara exterior** aplica ``INVERTIR_DIRECCION_PATA``; en **cara interior** el
   opuesto, como en la pata L (``arearein_interior_h_l135_rps``).

Revit 2024+ | RPS (IronPython 3.4). Resolución de ``sys.path`` alineada con
``arearein_exterior_h_l135_rps`` / ``arearein_interior_h_l135_rps`` / ``arearein_verticales_empotramiento_rps``.
"""

from __future__ import print_function

import os
import sys
import clr

import System

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    Floor,
    Line,
    Transaction,
    UnitUtils,
    UnitTypeId,
    Wall,
    XYZ,
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB.Structure import (
    AreaReinforcement,
    MultiplanarOption,
    Rebar,
    RebarBarType,
    RebarHookOrientation,
    RebarStyle,
)

# --- sys.path (RPS a veces sin __file__) ---
RPS_SCRIPTS_DIR_OVERRIDE = u""


def _script_containing_dir():
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except NameError:
        return None


def _paths_for_bimtools_scripts():
    out = []
    d0 = _script_containing_dir()
    if d0:
        out.append(d0)
    ovr = (RPS_SCRIPTS_DIR_OVERRIDE or u"").strip()
    if ovr and os.path.isdir(ovr):
        out.append(ovr)
    try:
        cwd = os.getcwd()
        if cwd and os.path.isdir(cwd):
            out.append(cwd)
    except System.Exception:
        pass
    try:
        home = os.path.expanduser(u"~")
        guess = os.path.join(
            home, u"CustomRevitExtensions", u"BIMTools.extension", u"scripts"
        )
        if os.path.isdir(guess):
            out.append(guess)
    except System.Exception:
        pass
    seen = set()
    uniq = []
    for p in out:
        ap = os.path.normpath(p) if p else u""
        if ap and ap not in seen and os.path.isdir(ap):
            seen.add(ap)
            uniq.append(ap)
    return uniq


for _d in _paths_for_bimtools_scripts():
    if _d not in sys.path:
        sys.path.insert(0, _d)

# Largo pata: espesor host − resta (misma función que rebar L + 135°)
try:
    from rebar_extender_l_ganchos_135_rps import largo_pata_mm_desde_espesor_host
except System.Exception as ex:
    largo_pata_mm_desde_espesor_host = None
    _LARGO_IMPORT_ERR = u"{0!s}".format(ex)
else:
    _LARGO_IMPORT_ERR = u""

# Parámetros (mm). Si :func:`largo_pata_mm_desde_espesor_host` no importa, se usa LARGO_PATA_MM
LARGO_PATA_MM = 200.0
PATA_RESTA_ESPESOR_HOST_MM = 50.0
PATA_LARGO_MIN_MM = 10.0

# Sentido in-plane de la pata (cara **exterior**). En cara **interior** se usa siempre
# el opuesto, para emparejar con la pata hacia el otro lado del elemento.
INVERTIR_DIRECCION_PATA = False
# Índice de barra (rebar) si varias en la misma posición; casi siempre 0
INDICE_POSICION = 0

# Criterio vertical (muro / losa) — igual que arearein_verticales_empotramiento_rps
MURO_VERT_MIN_ABS_TZ = 0.45
HORIZONTAL_MAX_ABS_TZ = 0.02
PLAN_PREFER_X_FOR_HORIZONTAL = True
# Cara exterior / interior — alineado con arearein_exterior / arearein_interior
EXTERIOR_IS_TOP_LAYER = True
MEDIA_CAPA_TOL = 0.0
WALL_EXTERIOR_MIN_ALONG_NORMAL_FT = 1e-5
WALL_INTERIOR_MAX_ALONG_NORMAL_FT = -1e-5


def _z_center(elem):
    bb = elem.get_BoundingBox(None) if elem is not None else None
    if bb is None:
        return None
    return 0.5 * (float(bb.Min.Z) + float(bb.Max.Z))


def _bb_center_xyz(el):
    bb = el.get_BoundingBox(None) if el is not None else None
    if bb is None:
        return None
    return XYZ(
        0.5 * (float(bb.Min.X) + float(bb.Max.X)),
        0.5 * (float(bb.Min.Y) + float(bb.Max.Y)),
        0.5 * (float(bb.Min.Z) + float(bb.Max.Z)),
    )


def _rebar_exterior_cara_sup_losa(rebar, host, exterior_is_top, tol):
    if rebar is None or host is None:
        return False
    zr = _z_center(rebar)
    zh = _z_center(host)
    if zr is None or zh is None:
        return False
    if exterior_is_top:
        return float(zr) >= float(zh) - float(tol)
    return float(zr) < float(zh) + float(tol)


def _rebar_muro_solo_cara_exterior(rebar, wall, min_along_normal):
    if rebar is None or wall is None:
        return False
    try:
        ori = wall.Orientation
    except System.Exception:
        return False
    if ori is None or ori.GetLength() < 1e-12:
        return False
    ori = ori.Normalize()
    pr = _bb_center_xyz(rebar)
    pw = _bb_center_xyz(wall)
    if pr is None or pw is None:
        return False
    d = pr - pw
    return float(d.DotProduct(ori)) > float(min_along_normal)


def _rebar_solo_cara_exterior(rebar, host):
    if rebar is None or host is None:
        return False
    if isinstance(host, Wall):
        return _rebar_muro_solo_cara_exterior(
            rebar, host, WALL_EXTERIOR_MIN_ALONG_NORMAL_FT
        )
    if isinstance(host, Floor):
        return _rebar_exterior_cara_sup_losa(
            rebar, host, EXTERIOR_IS_TOP_LAYER, MEDIA_CAPA_TOL
        )
    return _rebar_exterior_cara_sup_losa(
        rebar, host, EXTERIOR_IS_TOP_LAYER, MEDIA_CAPA_TOL
    )


def _rebar_interior_cara_inf_losa(rebar, host, exterior_is_top, tol):
    if rebar is None or host is None:
        return False
    zr = _z_center(rebar)
    zh = _z_center(host)
    if zr is None or zh is None:
        return False
    if exterior_is_top:
        return float(zr) < float(zh) + float(tol)
    return float(zr) >= float(zh) - float(tol)


def _rebar_muro_solo_cara_interior(rebar, wall, max_along_normal_neg):
    if rebar is None or wall is None:
        return False
    try:
        ori = wall.Orientation
    except System.Exception:
        return False
    if ori is None or ori.GetLength() < 1e-12:
        return False
    ori = ori.Normalize()
    pr = _bb_center_xyz(rebar)
    pw = _bb_center_xyz(wall)
    if pr is None or pw is None:
        return False
    d = pr - pw
    return float(d.DotProduct(ori)) < float(max_along_normal_neg)


def _rebar_solo_cara_interior(rebar, host):
    if rebar is None or host is None:
        return False
    if isinstance(host, Wall):
        return _rebar_muro_solo_cara_interior(
            rebar, host, WALL_INTERIOR_MAX_ALONG_NORMAL_FT
        )
    if isinstance(host, Floor):
        return _rebar_interior_cara_inf_losa(
            rebar, host, EXTERIOR_IS_TOP_LAYER, MEDIA_CAPA_TOL
        )
    return _rebar_interior_cara_inf_losa(
        rebar, host, EXTERIOR_IS_TOP_LAYER, MEDIA_CAPA_TOL
    )


def _face_cara_ext_o_int(rebar, host):
    """
    u'ext' = solo cara exterior, u'int' = solo interior, None = ninguna o ambigua
    (no se modifica la barra).
    """
    ex = _rebar_solo_cara_exterior(rebar, host)
    inn = _rebar_solo_cara_interior(rebar, host)
    if ex and not inn:
        return u"ext"
    if inn and not ex:
        return u"int"
    return None


def _obtener_espesor_host_mm(document, host):
    if host is None or document is None:
        return None
    try:
        if isinstance(host, Wall):
            p = host.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)
            if p is not None and p.HasValue:
                return float(
                    UnitUtils.ConvertFromInternalUnits(
                        p.AsDouble(), UnitTypeId.Millimeters
                    )
                )
        if isinstance(host, Floor):
            p = host.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)
            if p is not None and p.HasValue:
                return float(
                    UnitUtils.ConvertFromInternalUnits(
                        p.AsDouble(), UnitTypeId.Millimeters
                    )
                )
    except System.Exception:
        pass
    return None


def _largo_pata_mm_resuelto(doc, host):
    if largo_pata_mm_desde_espesor_host is not None:
        return largo_pata_mm_desde_espesor_host(
            doc,
            host,
            resta_mm=PATA_RESTA_ESPESOR_HOST_MM,
            fallback_mm=LARGO_PATA_MM,
            min_largo_mm=PATA_LARGO_MIN_MM,
        )
    th = _obtener_espesor_host_mm(doc, host)
    if th is None:
        return max(PATA_LARGO_MIN_MM, LARGO_PATA_MM)
    return max(PATA_LARGO_MIN_MM, float(th) - float(PATA_RESTA_ESPESOR_HOST_MM))


def _mm_to_internal(mm):
    return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)


def _eid(ei):
    try:
        return int(ei.Value)
    except System.Exception:
        try:
            return int(ei.IntegerValue)
        except System.Exception:
            return 0


def _tangent_start_first_segment(rebar, pos_idx=0):
    mpo = MultiplanarOption.IncludeAllMultiplanarCurves
    try:
        crvs = rebar.GetCenterlineCurves(
            False, False, False, mpo, int(pos_idx)
        )
    except System.Exception:
        return None, None, None
    if crvs is None or crvs.Count < 1:
        return None, None, None
    c0 = crvs[0]
    try:
        p0 = c0.GetEndPoint(0)
        p1 = c0.GetEndPoint(1)
        v = p1 - p0
        if v.GetLength() < 1e-12:
            return None, None, None
        t = v.Normalize()
    except System.Exception:
        return None, None, None
    txy = XYZ(t.X, t.Y, 0.0)
    return t, txy, crvs


def _rebar_horizontal_en_plano(rebar, prefer_x, max_abs_tz, pos_idx=0):
    t, txy, _ = _tangent_start_first_segment(rebar, pos_idx)
    if t is None or txy is None:
        return False
    if abs(float(t.Z)) > float(max_abs_tz):
        return False
    if txy.GetLength() < 1e-9:
        return True
    txy = txy.Normalize()
    ex, ey = abs(float(txy.X)), abs(float(txy.Y))
    if prefer_x:
        return ex >= ey
    return ey > ex


def _rebar_es_vertical_por_criterio(rebar, host, pos_idx=0):
    t, txy, _ = _tangent_start_first_segment(rebar, pos_idx)
    if t is None:
        return False
    if host is not None and isinstance(host, Wall):
        return abs(float(t.Z)) >= float(MURO_VERT_MIN_ABS_TZ)
    if host is not None and isinstance(host, Floor):
        return not _rebar_horizontal_en_plano(
            rebar, PLAN_PREFER_X_FOR_HORIZONTAL, HORIZONTAL_MAX_ABS_TZ, pos_idx
        )
    return abs(float(t.Z)) >= float(MURO_VERT_MIN_ABS_TZ)


def _rebar_normal(rebar):
    try:
        acc = rebar.GetShapeDrivenAccessor()
        if acc is not None:
            n = acc.Normal
            if n is not None and n.GetLength() > 1e-12:
                return n.Normalize()
    except System.Exception:
        pass
    return XYZ.BasisZ


def _tangent_start_curve(crv):
    if crv is None:
        return None
    p0 = crv.GetEndPoint(0)
    p1 = crv.GetEndPoint(1)
    v = p1 - p0
    if v.GetLength() < 1e-12:
        return None
    return v.Normalize()


def _perp_in_plane(normal, tangent):
    c = normal.CrossProduct(tangent)
    if c.GetLength() < 1e-10:
        c = tangent.CrossProduct(normal)
    if c.GetLength() < 1e-10:
        return None
    return c.Normalize()


def _B_pata_90_hacia_nucleo_muro(B, p_end, host):
    u"""
    Ajusta el signo de ``B`` (``normal × tangente``) para que el tramo corto
    ``p_end → p_end - B*L`` (sentido **-B**) vaya hacia el **núcleo** del muro
    (cara a cara, según el espesor y ``Wall.Orientation``), no hacia afuera.
    Se usa el componente del vector al centro de BBox **sobre** la orientación del
    muro para no confundir con muros alargados en planta.
    """
    if B is None or not isinstance(host, Wall) or p_end is None:
        return B
    v_dir = None
    try:
        wbb = host.get_BoundingBox(None)
    except System.Exception:
        wbb = None
    if wbb is not None:
        try:
            c = XYZ(
                0.5 * (wbb.Min.X + wbb.Max.X),
                0.5 * (wbb.Min.Y + wbb.Max.Y),
                0.5 * (wbb.Min.Z + wbb.Max.Z),
            )
        except System.Exception:
            c = None
        if c is not None:
            v_full = c - p_end
            try:
                o = host.Orientation
                if o is not None and o.GetLength() > 1e-12:
                    o = o.Normalize()
                    tcomp = v_full.Dot(o)
                    v_thin = o.Multiply(float(tcomp))
                    if v_thin is not None and v_thin.GetLength() > 1e-8:
                        v_dir = v_thin.Normalize()
            except System.Exception:
                v_dir = None
            if v_dir is None and v_full is not None and v_full.GetLength() > 1e-8:
                try:
                    v_dir = v_full.Normalize()
                except System.Exception:
                    v_dir = None
    if v_dir is None:
        return B
    try:
        if float(v_dir.Dot(B)) > 0.0:
            return B.Negate()
    except System.Exception:
        return B
    return B


def _layout_rule_name(rebar, acc):
    if acc is None:
        return u""
    try:
        r = rebar.LayoutRule
        if r is not None:
            return r.ToString()
    except System.Exception:
        pass
    try:
        r = acc.GetLayoutRule()
        if r is not None:
            return r.ToString()
    except System.Exception:
        pass
    return u""


def _spacing_internal(rebar):
    try:
        return float(rebar.MaxSpacing)
    except System.Exception:
        return 0.0


def _array_length_internal(acc):
    if acc is None:
        return 0.0
    try:
        return float(acc.ArrayLength)
    except System.Exception:
        try:
            return float(acc.GetArrayLength())
        except System.Exception:
            return 0.0


def _copy_layout_rebar_shape_driven(src, dst):
    a0 = src.GetShapeDrivenAccessor()
    a1 = dst.GetShapeDrivenAccessor()
    if a0 is None or a1 is None:
        return False, u"ShapeDrivenAccessor nulo (layout no copiable)."

    rule_name = _layout_rule_name(src, a0)
    sp = _spacing_internal(src)
    alen = _array_length_internal(a0)
    b_side = bool(a0.BarsOnNormalSide)
    inc0 = bool(src.IncludeFirstBar)
    inc1 = bool(src.IncludeLastBar)
    nbars = int(src.Quantity)

    try:
        if rule_name == u"Single":
            a1.SetLayoutAsSingle()
        elif rule_name == u"MaximumSpacing":
            a1.SetLayoutAsMaximumSpacing(sp, alen, b_side, inc0, inc1)
        elif rule_name in (u"Number", u"FixedNumber"):
            a1.SetLayoutAsFixedNumber(nbars, alen, b_side, inc0, inc1)
        elif rule_name == u"NumberWithSpacing":
            a1.SetLayoutAsNumberWithSpacing(nbars, sp, alen, b_side, inc0, inc1)
        elif rule_name == u"MinimumClearSpacing":
            a1.SetLayoutAsMinimumClearSpacing(sp, alen, b_side, inc0, inc1)
        else:
            if rule_name:
                try:
                    a1.SetLayoutAsFixedNumber(nbars, alen, b_side, inc0, inc1)
                except System.Exception:
                    a1.SetLayoutAsMaximumSpacing(sp, alen, b_side, inc0, inc1)
            else:
                a1.SetLayoutAsMaximumSpacing(sp, alen, b_side, inc0, inc1)
        return True, u""
    except System.Exception as ex:
        return False, u"{0!s} (regla: «{1}»)".format(ex, rule_name or u"(vacía)")


def _create_from_curves_no_hooks(doc, curves_list, host, norm, bar_type, style, o0, o1):
    ct = clr.GetClrType(Line).BaseType
    n = len(curves_list)
    arr = System.Array.CreateInstance(ct, n)
    for i in range(n):
        arr[i] = curves_list[i]
    return Rebar.CreateFromCurves(
        doc, style, bar_type, None, None, host, norm, arr, o0, o1, True, True
    )


def _hook_orient_for_create(rebar, end):
    e = int(end)
    try:
        o = rebar.GetHookOrientation(e)
        if o is not None:
            return o
    except System.Exception:
        pass
    try:
        o = rebar.GetTerminationOrientation(e)
        if o is not None:
            for name in (u"Left", u"Right"):
                v = int(o) == int(getattr(RebarHookOrientation, name))
                if v:
                    return getattr(RebarHookOrientation, name, RebarHookOrientation.Left)
    except System.Exception:
        pass
    return RebarHookOrientation.Left


def _iter_ids(ilist):
    if ilist is None:
        return []
    out = []
    try:
        n = int(ilist.Count)
    except System.Exception:
        n = 0
    for i in range(n):
        try:
            eid = ilist[i]
            if eid is not None and eid != ElementId.InvalidElementId:
                out.append(eid)
        except System.Exception:
            pass
    return out


def _add_pata_90_fin_a_rebar(doc, rebar, largo_pata_mm, invertir, pos_idx=0):
    """
    Añade un tramo al final del boceto (90° con el eje) y sustituye el `Rebar` (sin ganchos).
    """
    mpo = MultiplanarOption.IncludeAllMultiplanarCurves
    try:
        host = doc.GetElement(rebar.GetHostId())
    except System.Exception:
        host = None
    if host is None:
        return False, u"Host inválido.", None
    le = _mm_to_internal(largo_pata_mm)
    try:
        crvs = rebar.GetCenterlineCurves(
            False, False, False, mpo, int(pos_idx)
        )
    except System.Exception as ex:
        return False, u"GetCenterlineCurves: {0!s}".format(ex), None
    if crvs is None or int(crvs.Count) < 1:
        return False, u"Sin curvas de eje (pos. {0}).".format(pos_idx), None

    chain = [crvs[i] for i in range(crvs.Count)]
    c_last = chain[-1]
    T = _tangent_start_curve(c_last)
    if T is None:
        return False, u"Tangente nula en el último tramo.", None

    norm = _rebar_normal(rebar)
    B = _perp_in_plane(norm, T)
    if B is None:
        return False, u"No se pudo calcular la perpendicular in-plane (90°).", None
    p_end = c_last.GetEndPoint(1)
    if isinstance(host, Wall):
        B = _B_pata_90_hacia_nucleo_muro(B, p_end, host)
    else:
        if invertir:
            B = B.Negate()
    p_tip = p_end - B.Multiply(le)
    leg = Line.CreateBound(p_end, p_tip)
    new_chain = chain + [leg]

    bar_type = doc.GetElement(rebar.GetTypeId())
    if not isinstance(bar_type, RebarBarType):
        return False, u"RebarBarType no resuelto.", None

    try:
        style = rebar.Style
    except System.Exception:
        style = RebarStyle.Standard

    o0 = _hook_orient_for_create(rebar, 0)
    o1 = _hook_orient_for_create(rebar, 1)
    orig_rebar_id = rebar.Id

    def _try_shape_nseg_como_malla_l():
        """
        Misma pauta que horizont. L+135: RebarShape n-seg + pata; evita
        «Rebar shape / ganchos» al Commit. Depende de ``rebar_extender_l_ganchos_135_rps``.
        """
        try:
            import rebar_extender_l_ganchos_135_rps as _rfx
        except System.Exception:  # noqa: BLE001
            return None
        try:
            return _rfx._try_create_l_from_rebar_shape_2seg(
                doc, new_chain, host, norm, bar_type, style, o0, o1
            )
        except System.Exception:  # noqa: BLE001
            return None

    t = Transaction(
        doc, u"BIMTools: pata 90° fin boceto — nueva + layout (sin borrar malla)"
    )
    t.Start()
    new_rb = None
    try:
        new_rb = _try_shape_nseg_como_malla_l()
        if new_rb is None:
            new_rb = _create_from_curves_no_hooks(
                doc, new_chain, host, norm, bar_type, style, o0, o1
            )
        if new_rb is None:
            t.RollBack()
            return False, u"CreateFromCurves devolvió None.", None
        ok_lay, err_lay = _copy_layout_rebar_shape_driven(rebar, new_rb)
        if not ok_lay:
            t.RollBack()
            return False, u"Layout: {0}".format(err_lay or u"?"), None
        t.Commit()
    except System.Exception as ex:
        t.RollBack()
        return False, u"{0!s}".format(ex), None

    out_eid = new_rb.Id
    t2 = Transaction(
        doc, u"BIMTools: pata 90° fin boceto — eliminar rebar malla (original)"
    )
    t2.Start()
    try:
        if orig_rebar_id is not None and orig_rebar_id != ElementId.InvalidElementId:
            _old = doc.GetElement(orig_rebar_id)
            if _old is not None:
                doc.Delete(orig_rebar_id)
        t2.Commit()
    except System.Exception as ex2:
        try:
            t2.RollBack()
        except System.Exception:  # noqa: BLE001
            pass
        _nr = doc.GetElement(out_eid) if out_eid is not None else None
        if _nr is not None and isinstance(_nr, Rebar):
            return (
                True,
                u"Pata {0} mm; nuevo id {1}. [AVISO: no se eliminó malla (id {2}): {3!s}]".format(
                    int(round(float(largo_pata_mm))),
                    int(_eid(new_rb.Id)),
                    _eid(orig_rebar_id),
                    ex2,
                ),
                _nr,
            )
        return False, u"Delete malla+post pata: {0!s}".format(ex2), None

    new_rb = doc.GetElement(out_eid)
    if new_rb is None or not isinstance(new_rb, Rebar):
        return (
            False,
            u"Rebar pata 90 no resuelto (id {0}).".format(_eid(out_eid)),
            None,
        )
    return (
        True,
        u"Pata {0} mm; nuevo id {1}.".format(
            int(round(float(largo_pata_mm))), int(_eid(new_rb.Id))
        ),
        new_rb,
    )


def _run(doc, ar):
    t = Transaction(
        doc, u"BIMTools: quitar area reinf. (vert. ext./int. pata 90° fin boceto)"
    )
    t.Start()
    try:
        new_ids = AreaReinforcement.RemoveAreaReinforcementSystem(doc, ar)
    except System.Exception as ex:
        t.RollBack()
        return 0, 0, 0, 0, [u"RemoveAreaReinforcementSystem: {0!s}".format(ex)], []
    t.Commit()
    try:
        doc.Regenerate()
    except System.Exception:
        pass

    created = _iter_ids(new_ids)
    n_ok = n_skip = n_fail = 0
    log = []
    rebar_finales = []
    for eid in created:
        r = doc.GetElement(eid)
        if r is None:
            continue
        if not isinstance(r, Rebar):
            n_skip += 1
            rebar_finales.append(eid)
            continue
        host = doc.GetElement(r.GetHostId())
        cara = _face_cara_ext_o_int(r, host)
        if cara is None:
            n_skip += 1
            rebar_finales.append(eid)
            continue
        if not _rebar_es_vertical_por_criterio(r, host, 0):
            n_skip += 1
            rebar_finales.append(eid)
            continue
        Lp = _largo_pata_mm_resuelto(doc, host)
        invertir = bool(INVERTIR_DIRECCION_PATA)
        if cara == u"int":
            invertir = not invertir
        ok, msg, new_rb = _add_pata_90_fin_a_rebar(
            doc, r, Lp, invertir, INDICE_POSICION
        )
        if ok:
            n_ok += 1
            if new_rb is not None:
                rebar_finales.append(new_rb.Id)
            else:
                rebar_finales.append(eid)
            log.append(
                u"OK id {0} ({3}, pata ≈{1} mm): {2}".format(
                    int(_eid(eid)),
                    int(round(Lp)),
                    msg or u"",
                    cara,
                )
            )
        else:
            n_fail += 1
            rebar_finales.append(eid)
            log.append(
                u"FALLO id {0}: {1}".format(int(_eid(eid)), msg or u"")
            )
    return n_ok, n_skip, n_fail, len(created), log, rebar_finales


def run(uidoc):
    doc = uidoc.Document
    ids = uidoc.Selection.GetElementIds()
    if ids is None or ids.Count == 0:
        TaskDialog.Show(
            u"Arearein: vert. ext./int. pata 90° (fin boceto)",
            u"Selecciona un Area Reinforcement (un elemento).",
        )
        return
    if int(ids.Count) != 1:
        TaskDialog.Show(
            u"Arearein: vert. ext./int. pata 90° (fin boceto)",
            u"Selecciona un solo Area Reinforcement.",
        )
        return
    eid = ids[0]
    ar = doc.GetElement(eid)
    if not isinstance(ar, AreaReinforcement):
        TaskDialog.Show(
            u"Arearein: vert. ext./int. pata 90° (fin boceto)",
            u"El elemento no es un Area Reinforcement.",
        )
        return
    n_ok, n_skip, n_fail, n_created, log, _rebar_ids_vis = _run(doc, ar)
    summary = (
        u"Rebar creados al quitar el área: {3}. "
        u"Pata 90° (vert. cara ext. o int., fin)={0}, "
        u"omitidas (no criterio)={1}, error={2}."
    ).format(n_ok, n_skip, n_fail, n_created)
    if _LARGO_IMPORT_ERR and largo_pata_mm_desde_espesor_host is None:
        summary += u"\n(Aviso: no se importó rebar_extender_l_ganchos_135_rps; " u"largo fijo/ espesor local.)"
    tail = u"\n".join(log[:50])
    if len(log) > 50:
        tail += u"\n..."
    print(summary)
    print(tail)
    TaskDialog.Show(
        u"Arearein: vert. ext./int. — pata 90° fin boceto (int. opuesta a ext.)",
        u"{0}\n\n{1}".format(summary, tail),
    )


if __name__ == u"__main__":
    run(__revit__.ActiveUIDocument)  # noqa: F821
