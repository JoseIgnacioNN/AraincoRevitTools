# -*- coding: utf-8 -*-
"""
1) Seleccionar un Area Reinforcement.
2) Quitar el «area system» (Rebar vía ``AreaReinforcement.RemoveAreaReinforcementSystem``).
3) De las nuevas barras, filtrar **solo la cara interior** (muro: opuesto a ``Wall.Orientation``;
   losa: capa opuesta a la definida con ``EXTERIOR_IS_TOP_LAYER`` en este archivo).
4) De esas, **horizontales en planta** (mismos criterios que :mod:`arearein_exterior_h_l135_rps`).
5) Aplicar L + ganchos 135° con :mod:`rebar_extender_l_ganchos_135_rps`: **pata L en el extremo
   final** del boceto (``pata_en_extremo_final=True``) y **sentido de pata** opuesto a la cara
   exterior (``invertir=not l135.INVERTIR_DIRECCION_PATA``).

Revit 2024 + RPS. Mismo ``sys.path`` / ``RPS_SCRIPTS_DIR_OVERRIDE`` que el script de cara exterior.
   Para ambas caras a la vez: :mod:`arearein_ambas_caras_h_l135_rps`.
"""

from __future__ import print_function

import os
import sys
import clr

import System  # requerido para except System.Exception en IronPython

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    ElementId,
    Floor,
    Transaction,
    Wall,
    XYZ,
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB.Structure import (
    AreaReinforcement,
    Rebar,
)

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

try:
    import rebar_extender_l_ganchos_135_rps as l135
except System.Exception as ex:
    l135 = None
    _L135_IMPORT_ERR = u"{0!s}".format(ex)
else:
    _L135_IMPORT_ERR = u""

PLAN_PREFER_X_FOR_HORIZONTAL = True
HORIZONTAL_MAX_ABS_TZ = 0.02
# Debe matchear con arearein_exterior_h_l135_rps: si allí "exterior" = arriba, aquí interior = abajo
EXTERIOR_IS_TOP_LAYER = True
MEDIA_CAPA_TOL = 0.0
# Muro: interior = centro de barra al lado de **-Orientation** (cara opuesta a la exterior)
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


def _tangent_first_segment_xy(rebar):
    from Autodesk.Revit.DB.Structure import MultiplanarOption
    mpo = MultiplanarOption.IncludeAllMultiplanarCurves
    try:
        crvs = rebar.GetCenterlineCurves(
            False, False, False, mpo, 0
        )
    except System.Exception:
        return None, None
    if crvs is None or crvs.Count < 1:
        return None, None
    c0 = crvs[0]
    try:
        p0 = c0.GetEndPoint(0)
        p1 = c0.GetEndPoint(1)
        v = p1 - p0
        if v.GetLength() < 1e-12:
            return None, None
        t = v.Normalize()
    except System.Exception:
        return None, None
    txy = XYZ(t.X, t.Y, 0.0)
    return t, txy


def _tangent_dominant_segment_xy(rebar):
    from Autodesk.Revit.DB.Structure import MultiplanarOption
    mpo = MultiplanarOption.IncludeAllMultiplanarCurves
    try:
        crvs = rebar.GetCenterlineCurves(False, False, False, mpo, 0)
    except System.Exception:
        return None, None
    if crvs is None or crvs.Count < 1:
        return None, None
    best_c = None
    best_len = -1.0
    for i in range(int(crvs.Count)):
        try:
            c = crvs[i]
            p0 = c.GetEndPoint(0)
            p1 = c.GetEndPoint(1)
            v = p1 - p0
            ln = float(v.GetLength())
            if ln > best_len:
                best_len = ln
                best_c = c
        except System.Exception:
            continue
    if best_c is None or best_len < 1e-12:
        return None, None
    try:
        p0 = best_c.GetEndPoint(0)
        p1 = best_c.GetEndPoint(1)
        v = p1 - p0
        t = v.Normalize()
    except System.Exception:
        return None, None
    txy = XYZ(t.X, t.Y, 0.0)
    return t, txy


def _rebar_interior_cara_inf_losa(rebar, host, exterior_is_top, tol):
    """
    Losa: capa interior = opuesta a la de ``arearein_exterior_h_l135_rps`` (si exterior=top → interior=mid inferior).
    """
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
    """
    Muro: ``Wall.Orientation`` apunta a la cara exterior. La interior es el semiespacio con
    ``(bb_rebar - bb_muro) · Orientation < 0``.
    """
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


def _rebar_horizontal_en_plano(rebar, prefer_x, max_abs_tz, host=None):
    t, txy = _tangent_dominant_segment_xy(rebar)
    if t is None or txy is None:
        t, txy = _tangent_first_segment_xy(rebar)
    if t is None or txy is None:
        return False
    if abs(float(t.Z)) > float(max_abs_tz):
        return False
    if host is not None and isinstance(host, Wall):
        return True
    if txy.GetLength() < 1e-9:
        return True
    txy = txy.Normalize()
    ex, ey = abs(float(txy.X)), abs(float(txy.Y))
    if prefer_x:
        return ex >= ey
    return ey > ex


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


def _invertir_pata_respecto_exterior():
    """
    Misma pata L que en arearein_exterior_h_l135_rps sería l135.INVERTIR_DIRECCION_PATA;
    la cara interior exige el sentido opuesto.
    """
    if l135 is None:
        return False
    return not bool(l135.INVERTIR_DIRECCION_PATA)


def _run_area_to_l(doc, ar):
    msgs = []
    t = Transaction(doc, u"BIMTools: quitar area reinforcement system (cara int.)")
    t.Start()
    try:
        new_ids = AreaReinforcement.RemoveAreaReinforcementSystem(doc, ar)
    except System.Exception as ex:
        t.RollBack()
        return 0, 0, 0, [u"RemoveAreaReinforcementSystem: {0!s}".format(ex)], 0
    t.Commit()
    try:
        doc.Regenerate()
    except System.Exception:
        pass

    created = _iter_ids(new_ids)
    n_ok = n_skip = n_fail = 0
    invert = _invertir_pata_respecto_exterior()
    for eid in created:
        r = doc.GetElement(eid)
        if r is None or not isinstance(r, Rebar):
            n_skip += 1
            continue
        host = doc.GetElement(r.GetHostId())
        if not _rebar_solo_cara_interior(r, host):
            n_skip += 1
            continue
        if not _rebar_horizontal_en_plano(
            r, PLAN_PREFER_X_FOR_HORIZONTAL, HORIZONTAL_MAX_ABS_TZ, host
        ):
            n_skip += 1
            continue
        largo_p = l135.largo_pata_mm_desde_espesor_host(doc, host)
        ok, msg, _nrb = l135.extender_l_asignar_ganchos_135_y_reemplazar(
            doc,
            r,
            largo_p,
            invert,
            l135.INDICE_POSICION,
            True,
        )
        if ok:
            n_ok += 1
            msgs.append(u"OK id {0}: {1}".format(eid.IntegerValue, msg or u""))
        else:
            n_fail += 1
            msgs.append(u"FALLO id {0}: {1}".format(eid.IntegerValue, msg or u""))
    return n_ok, n_skip, n_fail, msgs, len(created)


def run(uidoc):
    if l135 is None:
        TaskDialog.Show(
            u"Area Reinforcement (interior) → L + 135°",
            u"No se pudo importar rebar_extender_l_ganchos_135_rps.py: {0}\n"
            u"Colóquelo en el mismo directorio que este script.".format(_L135_IMPORT_ERR),
        )
        return
    doc = uidoc.Document
    ids = uidoc.Selection.GetElementIds()
    if ids is None or ids.Count == 0:
        TaskDialog.Show(
            u"Area Reinforcement (interior) → L + 135°",
            u"Selecciona un Area Reinforcement (un elemento).",
        )
        return
    if int(ids.Count) != 1:
        TaskDialog.Show(
            u"Area Reinforcement (interior) → L + 135°",
            u"Selecciona un solo Area Reinforcement.",
        )
        return
    eid = ids[0]
    ar = doc.GetElement(eid)
    if not isinstance(ar, AreaReinforcement):
        TaskDialog.Show(
            u"Area Reinforcement (interior) → L + 135°",
            u"El elemento no es un Area Reinforcement.",
        )
        return
    n_ok, n_skip, n_fail, log, n_created = _run_area_to_l(doc, ar)
    summary = (
        u"Rebar creados tras quitar el área: {3}. "
        u"OK L+135° (cara interior)={0}, omitidas={1}, error={2}."
    ).format(n_ok, n_skip, n_fail, n_created)
    tail = u"\n".join(log[:40])
    if len(log) > 40:
        tail += u"\n..."
    print(summary)
    print(tail)
    TaskDialog.Show(
        u"Area Reinforcement (interior) → L + 135°",
        u"{0}\n\n{1}".format(summary, tail),
    )


if __name__ == u"__main__":
    run(__revit__.ActiveUIDocument)  # noqa: F821
