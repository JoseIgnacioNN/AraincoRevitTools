# -*- coding: utf-8 -*-
"""
1) Seleccionar un **Area Reinforcement**.
2) Quitar el «area system» (``RebarInSystem`` → ``Rebar``) con
   ``AreaReinforcement.RemoveAreaReinforcementSystem``.
3) Entre las barras creadas, quedarse con las **verticales** (muro: tangente
   domina eje Z global; losa: sentido contrario a «horizontales en planta» del
   script ``arearein_exterior_h_l135_rps``).
4) **Alargar** el trazado eje a lo largo de la barra: ``mm inicio = L'×k0`` y
   ``mm fin = L'×k1``, con *L'* = *L* (tabla) + **25 mm fijos** adicionales, donde
   *L* sale de
   :func:`bimtools_rebar_hook_lengths.traslape_mm_from_nominal_diameter_mm` (tabla
   traslape/empotramiento G25 / G35 / G45 / base BIMTools) según el diámetro
   nominal del ``RebarBarType`` activo en cada barra.

Revit 2024+ | RPS (IronPython 3.4). Misma resolución de ``sys.path`` que
``arearein_exterior_h_l135_rps`` (incl. ``RPS_SCRIPTS_DIR_OVERRIDE``).
"""

from __future__ import print_function

import os
import sys
import clr

import System

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
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

# --- Import tabla empotramiento / traslape por diámetro (mismo dir. que este script) ---
RPS_SCRIPTS_DIR_OVERRIDE = u""

try:
    from bimtools_rebar_hook_lengths import traslape_mm_from_nominal_diameter_mm
except System.Exception as ex:
    traslape_mm_from_nominal_diameter_mm = None
    _EMBED_IMPORT_ERR = u"{0!s}".format(ex)
else:
    _EMBED_IMPORT_ERR = u""

# Grado de hormigón: G25, G35, G45, o None (tabla base BIMTools)
CONCRETE_GRADE = u"G35"
# Suma fija (mm) al largo de empotramiento/traslape de tabla (mismos criterio BIMTools).
EMBED_EXTRA_TABLE_MM = 25.0
# Extensión fija (mm) en el **pie** (menor *Z*) con muro unido en cara inferior
# (empalme sin pata L ni gancho 135 en ese extremo). Coherente con el criterio BIMTools.
EMBED_MURO_INFERIOR_MM = 25.0

# Tolerancia |t·Z| mínima para tratar la barra como **vertical** en muro
# (barras en el plano del muro; las «horizontales» tienen |t·Z| bajo)
MURO_VERT_MIN_ABS_TZ = 0.45
# Tolerancia «en planta» (losa) — igual criterio que arearein_exterior
HORIZONTAL_MAX_ABS_TZ = 0.02
# Losa: «vertical en planta» = NO horizontal según criterio X/Y
PLAN_PREFER_X_FOR_HORIZONTAL = True
# (k0, k1) multiplican L de tabla en el **inicio** (primer endpoint) y **fin** del eje
# L = traslape_mm_from_nominal_diameter_mm(ø). p.ej. (1, 0) = solo hacia inicio, (0.5, 0.5) = reparte
K_EMPORT_INICIO = 1.0
K_EMPORT_FIN = 0.0


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
        guess = os.path.join(home, u"CustomRevitExtensions", u"BIMTools.extension", u"scripts")
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

# Reintento de import si el path aún no existía
if traslape_mm_from_nominal_diameter_mm is None:
    try:
        from bimtools_rebar_hook_lengths import traslape_mm_from_nominal_diameter_mm
    except System.Exception as ex2:
        traslape_mm_from_nominal_diameter_mm = None
        if not _EMBED_IMPORT_ERR:
            _EMBED_IMPORT_ERR = u"{0!s}".format(ex2)


def _tangent_first_segment(rebar, pos_idx=0):
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
    t, txy, _ = _tangent_first_segment(rebar, pos_idx)
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
    t, txy, _ = _tangent_first_segment(rebar, pos_idx)
    if t is None:
        return False
    if host is not None and isinstance(host, Wall):
        return abs(float(t.Z)) >= float(MURO_VERT_MIN_ABS_TZ)
    if host is not None and isinstance(host, Floor):
        return not _rebar_horizontal_en_plano(
            rebar, PLAN_PREFER_X_FOR_HORIZONTAL, HORIZONTAL_MAX_ABS_TZ, pos_idx
        )
    return abs(float(t.Z)) >= float(MURO_VERT_MIN_ABS_TZ)


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


def _mm_to_internal(mm):
    return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)


def _nominal_diameter_mm_from_rebar(rebar, doc):
    try:
        bt = doc.GetElement(rebar.GetTypeId())
    except System.Exception:
        bt = None
    if not isinstance(bt, RebarBarType):
        return None
    try:
        d = bt.BarModelDiameter
        return float(UnitUtils.ConvertFromInternalUnits(d, UnitTypeId.Millimeters))
    except System.Exception:
        return None


def _tangent_start_curve(crv):
    if crv is None:
        return None
    p0 = crv.GetEndPoint(0)
    p1 = crv.GetEndPoint(1)
    v = p1 - p0
    if v.GetLength() < 1e-12:
        return None
    return v.Normalize()


def _tangent_at_end_of_curve(crv):
    """Tangente **saliente** en el nodo 1 (dirección de marcha de la polilínea)."""
    return _tangent_start_curve(crv)


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


def _extender_rebar_por_eje_mm(
    doc,
    rebar,
    mm_inicio,
    mm_fin,
    pos_idx=0,
):
    """
    Sustituye el rebar alargando en la primera/última curva (mm positivos, según
    sentido: inicio hacia atrás, fin hacia adelante a lo largo del trazado).
    """
    mpo = MultiplanarOption.IncludeAllMultiplanarCurves
    try:
        crvs = rebar.GetCenterlineCurves(
            False, False, False, mpo, int(pos_idx)
        )
    except System.Exception as ex:
        return False, u"GetCenterlineCurves: {0!s}".format(ex), None
    if crvs is None or int(crvs.Count) < 1:
        return False, u"Sin curvas de eje (pos. {0}).".format(pos_idx), None

    la = max(0.0, float(mm_inicio))
    lb = max(0.0, float(mm_fin))
    if la < 1e-6 and lb < 1e-6:
        return False, u"K_EMPORT_INICIO / K_EMPORT_FIN anulan la extensión (0 mm).", None

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
    new_p0 = p0s - t0.Multiply(a)
    new_p1e = p1e + t1.Multiply(b)

    if int(crvs.Count) == 1:
        c_new = Line.CreateBound(new_p0, new_p1e)
        new_chain = [c_new]
    else:
        c_first_new = Line.CreateBound(new_p0, p0e)
        c_last_new = Line.CreateBound(p1s, new_p1e)
        new_chain = [c_first_new] + chain[1:-1] + [c_last_new]
        for i in range(len(new_chain) - 1):
            e_prev = new_chain[i].GetEndPoint(1)
            s_next = new_chain[i + 1].GetEndPoint(0)
            d = (e_prev - s_next).GetLength()
            if d > 0.01:
                return (
                    False,
                    u"Polilínea no consecutiva (gap {0:,.4f} ft) — use barra con un tramo o "
                    u"ajuste la malla a retos.".format(d),
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
    except System.Exception:
        style = RebarStyle.Standard
    norm = _rebar_normal(rebar)
    o0 = _hook_orient_for_create(rebar, 0)
    o1 = _hook_orient_for_create(rebar, 1)

    t = Transaction(doc, u"BIMTools: rebar eje + empotramiento (tabla diámetro)")
    t.Start()
    try:
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
        try:
            doc.Delete(rebar.Id)
        except System.Exception as ex2:
            t.RollBack()
            return False, u"Delete rebar: {0!s}".format(ex2), None
        t.Commit()
    except System.Exception as ex:
        t.RollBack()
        return False, u"{0!s}".format(ex), None
    return (
        True,
        u"Ext. inicio {0} mm, fin {1} mm; nuevo id {2}.".format(
            int(round(la)), int(round(lb)), int(_eid(new_rb.Id))
        ),
        new_rb,
    )


def _eid(ei):
    try:
        return int(ei.Value)
    except System.Exception:
        try:
            return int(ei.IntegerValue)
        except System.Exception:
            return 0


def _rebar_eje_p_start_p_end(rebar, pos_idx=0):
    u"""
    Origen y fin del boceto de eje: primer GetEndPoint(0) y último GetEndPoint(1);
    mismos criterios que al extender por inicio/fin.
    """
    mpo = MultiplanarOption.IncludeAllMultiplanarCurves
    try:
        crvs = rebar.GetCenterlineCurves(
            False, False, False, mpo, int(pos_idx)
        )
    except System.Exception:
        return None, None
    if crvs is None or int(crvs.Count) < 1:
        return None, None
    try:
        c0 = crvs[0]
        cN = crvs[crvs.Count - 1]
        return c0.GetEndPoint(0), cN.GetEndPoint(1)
    except System.Exception:
        return None, None


def extender_vertical_cabeza_empotramiento_por_diam(
    doc, rebar, pos_idx, concrete_grade=None
):
    u"""
    Extiende el boceto en el extremo de **mayor Z** (cabeza del trazado vertical) con
    Sobre el tramo a la **cabeza** (mayor *Z*) se aplica
    *L* = valor de tabla + ``EMBED_EXTRA_TABLE_MM`` (25 mm fijos). La tabla es
    ``traslape_mm_from_nominal_diameter_mm(ø, grado)`` (G25 / G35 / G45 o base
    BIMTools). Sustituye el Rebar vía _extender_rebar_por_eje_mm. Uso: muro en
    cara superior, sin pata L ni 135 en cabeza: desarrollo en eje.
    """
    if traslape_mm_from_nominal_diameter_mm is None or doc is None or rebar is None:
        return False, u"Empotramiento: módulo tabla o rebar no válido.", None
    d_mm = _nominal_diameter_mm_from_rebar(rebar, doc)
    if d_mm is None or d_mm <= 0.0:
        return False, u"Ø nominal no leído (tipo).", None
    g = CONCRETE_GRADE if concrete_grade is None else concrete_grade
    try:
        L = traslape_mm_from_nominal_diameter_mm(d_mm, g)
    except System.Exception as ex:
        return False, u"Tabla (ø): {0!s}".format(ex), None
    if L is None or L != L or float(L) < 0.0:
        return False, u"L de tabla (empotram.) inválida.", None
    Lf = max(0.0, float(L)) + float(EMBED_EXTRA_TABLE_MM)
    if Lf < 0.1:
        return True, u"L=0, sin extender (empotram. cabeza).", rebar
    p0, p1 = _rebar_eje_p_start_p_end(rebar, int(pos_idx))
    if p0 is None or p1 is None:
        return False, u"Sin curvas de boceto.", None
    try:
        eps = UnitUtils.ConvertToInternalUnits(0.5, UnitTypeId.Millimeters)
    except System.Exception:
        eps = 1.0e-4
    if float(p1.Z) >= float(p0.Z) - float(eps):
        return _extender_rebar_por_eje_mm(
            doc, rebar, 0.0, Lf, int(pos_idx)
        )
    return _extender_rebar_por_eje_mm(doc, rebar, Lf, 0.0, int(pos_idx))


def extender_vertical_pie_emp_muro_inf_mm(doc, rebar, pos_idx, mm_fijo=None):
    u"""
    Alarga el boceto en el extremo de **menor Z** (pie hacia muro apilado en la
    cara inferior) con *mm* fijos (por defecto ``EMBED_MURO_INFERIOR_MM``),
    cuando en ese extremo no corresponde pata L ni gancho 135 (unión con muro bajo).
    Misma sustitución que :func:`extender_vertical_cabeza_empotramiento_por_diam`
    pero hacia el encuentro inferior.
    """
    if doc is None or rebar is None:
        return False, u"Empotramiento pie: doc o rebar no válido.", None
    try:
        m = float(EMBED_MURO_INFERIOR_MM if mm_fijo is None else mm_fijo)
    except System.Exception:
        m = float(EMBED_MURO_INFERIOR_MM)
    m = max(0.0, m)
    if m < 0.1:
        return True, u"mm pie muro inf. = 0, sin extender.", rebar
    p0, p1 = _rebar_eje_p_start_p_end(rebar, int(pos_idx))
    if p0 is None or p1 is None:
        return False, u"Sin curvas de boceto.", None
    try:
        eps = UnitUtils.ConvertToInternalUnits(0.5, UnitTypeId.Millimeters)
    except System.Exception:
        eps = 1.0e-4
    if float(p1.Z) >= float(p0.Z) - float(eps):
        return _extender_rebar_por_eje_mm(doc, rebar, m, 0.0, int(pos_idx))
    return _extender_rebar_por_eje_mm(doc, rebar, 0.0, m, int(pos_idx))


def _run_area_to_rebar_empotramiento(doc, ar):
    if traslape_mm_from_nominal_diameter_mm is None:
        return (
            0,
            0,
            0,
            0,
            [u"Import bimtools_rebar_hook_lengths: {0}".format(_EMBED_IMPORT_ERR)],
        )

    t = Transaction(doc, u"BIMTools: quitar area reinforcement system (vert. emp.)")
    t.Start()
    try:
        new_ids = AreaReinforcement.RemoveAreaReinforcementSystem(doc, ar)
    except System.Exception as ex:
        t.RollBack()
        return 0, 0, 0, 0, [u"RemoveAreaReinforcementSystem: {0!s}".format(ex)]
    t.Commit()
    try:
        doc.Regenerate()
    except System.Exception:
        pass

    created = _iter_ids(new_ids)
    n_ok = n_skip = n_fail = 0
    log = []
    for eid in created:
        r = doc.GetElement(eid)
        if r is None or not isinstance(r, Rebar):
            n_skip += 1
            continue
        host = doc.GetElement(r.GetHostId())
        if not _rebar_es_vertical_por_criterio(r, host, 0):
            n_skip += 1
            continue
        d_mm = _nominal_diameter_mm_from_rebar(r, doc)
        if d_mm is None or d_mm <= 0:
            n_fail += 1
            log.append(
                u"FALLO id {0}: no se lee diámetro de tipo.".format(int(_eid(eid)))
            )
            continue
        try:
            L = traslape_mm_from_nominal_diameter_mm(
                d_mm, CONCRETE_GRADE
            )
        except System.Exception as ex:
            n_fail += 1
            log.append(
                u"FALLO id {0}: tabla mm (ø {1}): {2}".format(
                    int(_eid(eid)), d_mm, ex
                )
            )
            continue
        if L is None or L < 0:
            n_fail += 1
            log.append(
                u"FALLO id {0}: L empotramiento/traslape inválido (ø {1} mm).".format(
                    int(_eid(eid)), d_mm
                )
            )
            continue
        Lp = float(L) + float(EMBED_EXTRA_TABLE_MM)
        mi = Lp * float(K_EMPORT_INICIO)
        mf = Lp * float(K_EMPORT_FIN)
        ok, msg, _ = _extender_rebar_por_eje_mm(doc, r, mi, mf, 0)
        if ok:
            n_ok += 1
            log.append(
                u"OK id {0} ø{1} mm: Ltab={2} +{3}→ L'={4} → +inicio {5} +fin {6} | {7}".format(
                    int(_eid(eid)),
                    int(round(d_mm)),
                    int(round(float(L))),
                    int(round(float(EMBED_EXTRA_TABLE_MM))),
                    int(round(Lp)),
                    int(round(mi)),
                    int(round(mf)),
                    msg or u"",
                )
            )
        else:
            n_fail += 1
            log.append(
                u"FALLO id {0} (ø{1} mm): {2}".format(
                    int(_eid(eid)), int(round(d_mm)), msg or u""
                )
            )
    return n_ok, n_skip, n_fail, len(created), log


def run(uidoc):
    if traslape_mm_from_nominal_diameter_mm is None:
        TaskDialog.Show(
            u"Arearein: verticales + empotramiento",
            u"No se pudo importar bimtools_rebar_hook_lengths.py.\n{0}\n"
            u"Colóquelo en la carpeta scripts del add-in.".format(
                _EMBED_IMPORT_ERR or u"(sin detalle)"
            ),
        )
        return
    doc = uidoc.Document
    ids = uidoc.Selection.GetElementIds()
    if ids is None or ids.Count == 0:
        TaskDialog.Show(
            u"Arearein: verticales + empotramiento",
            u"Selecciona un Area Reinforcement (un elemento).",
        )
        return
    if int(ids.Count) != 1:
        TaskDialog.Show(
            u"Arearein: verticales + empotramiento",
            u"Selecciona un solo Area Reinforcement.",
        )
        return
    eid = ids[0]
    ar = doc.GetElement(eid)
    if not isinstance(ar, AreaReinforcement):
        TaskDialog.Show(
            u"Arearein: verticales + empotramiento",
            u"El elemento no es un Area Reinforcement.",
        )
        return
    n_ok, n_skip, n_fail, n_created, log = _run_area_to_rebar_empotramiento(doc, ar)
    summary = (
        u"Rebar creados al quitar el área: {3}. Grado: {4}. "
        u"Alargadas (tabla)={0}, omitidas (no vert.)={1}, error={2}."
    ).format(
        n_ok, n_skip, n_fail, n_created, (CONCRETE_GRADE or u"base")
    )
    tail = u"\n".join(log[:50])
    if len(log) > 50:
        tail += u"\n..."
    print(summary)
    print(tail)
    TaskDialog.Show(
        u"Arearein: verticales + empotramiento (tabla diámetro)",
        u"{0}\n\n{1}".format(summary, tail),
    )


if __name__ == u"__main__":
    run(__revit__.ActiveUIDocument)  # noqa: F821
