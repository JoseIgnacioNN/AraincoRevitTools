# -*- coding: utf-8 -*-
"""
1) Seleccionar un **Area Reinforcement**.
2) **Una sola** conversión: ``RemoveAreaReinforcementSystem`` (quitar el sistema de área).
3) Sobre **cada** Rebar creada:
   - si es **horizontal en planta** (criterio compartido con los scripts de una cara);
   - y su centro respecto al host encaja con **cara exterior** → L + 135° como
     :mod:`arearein_exterior_h_l135_rps` (``INVERTIR_DIRECCION_PATA`` de
     :mod:`rebar_extender_l_ganchos_135_rps`);
   - o encaja con **cara interior** → L + 135° con pata al **revés** (como
     :mod:`arearein_interior_h_l135_rps`);
   - si no encaja con ninguna cara, se omite.

La prioridad de clasificación es **exterior primero, luego interior** (casi no se solapan).

Revit 2024 + RPS. Mismos ajustes de ruta (``RPS_SCRIPTS_DIR_OVERRIDE``) y constantes
horizontales que en ``arearein_exterior_h_l135_rps``/``arearein_interior_h_l135_rps``.
"""

from __future__ import print_function

import os
import sys
import clr

import System  # requerido para except System.Exception en IronPython

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import Transaction
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

_L135_ERR = u""
_AREX_ERR = u""
_ARIN_ERR = u""
l135 = None
arex = None
arin = None
try:
    import rebar_extender_l_ganchos_135_rps as l135
except System.Exception as ex:
    l135 = None
    _L135_ERR = u"{0!s}".format(ex)
try:
    import arearein_exterior_h_l135_rps as arex
except System.Exception as ex:
    arex = None
    _AREX_ERR = u"{0!s}".format(ex)
try:
    import arearein_interior_h_l135_rps as arin
except System.Exception as ex:
    arin = None
    _ARIN_ERR = u"{0!s}".format(ex)


def _import_error_message():
    parts = []
    if l135 is None:
        parts.append(u"rebar_extender_l_ganchos_135_rps: {0}".format(_L135_ERR or u"—"))
    if arex is None:
        parts.append(u"arearein_exterior_h_l135_rps: {0}".format(_AREX_ERR or u"—"))
    if arin is None:
        parts.append(u"arearein_interior_h_l135_rps: {0}".format(_ARIN_ERR or u"—"))
    if not parts:
        return u""
    return u"\n".join(parts)


def _run_area_ambas_caras(doc, ar):
    if l135 is None or arex is None or arin is None:
        return 0, 0, 0, 0, 0, [u"Import: {0}".format(_import_error_message())], 0

    msgs = []
    t = Transaction(doc, u"BIMTools: quitar area reinforcement system (ext+int L+135)")
    t.Start()
    try:
        new_ids = AreaReinforcement.RemoveAreaReinforcementSystem(doc, ar)
    except System.Exception as ex:
        t.RollBack()
        return 0, 0, 0, 0, 0, [u"RemoveAreaReinforcementSystem: {0!s}".format(ex)], 0
    t.Commit()
    try:
        doc.Regenerate()
    except System.Exception:
        pass

    created = arex._iter_ids(new_ids)
    n_ok_ex = n_ok_in = n_skip = n_fail = 0
    inv_ext = bool(l135.INVERTIR_DIRECCION_PATA)
    inv_int = not inv_ext
    px, tz = arex.PLAN_PREFER_X_FOR_HORIZONTAL, arex.HORIZONTAL_MAX_ABS_TZ

    for eid in created:
        r = doc.GetElement(eid)
        if r is None or not isinstance(r, Rebar):
            n_skip += 1
            continue
        host = doc.GetElement(r.GetHostId())
        if not arex._rebar_horizontal_en_plano(r, px, tz, host):
            n_skip += 1
            continue
        if arex._rebar_solo_cara_exterior(r, host):
            invert, cara = inv_ext, u"ext"
        elif arin._rebar_solo_cara_interior(r, host):
            invert, cara = inv_int, u"int"
        else:
            n_skip += 1
            continue
        largo_p = l135.largo_pata_mm_desde_espesor_host(doc, host)
        ok, msg, _nrb = l135.extender_l_asignar_ganchos_135_y_reemplazar(
            doc,
            r,
            largo_p,
            invert,
            l135.INDICE_POSICION,
            cara == u"int",
        )
        if ok:
            if cara == u"ext":
                n_ok_ex += 1
            else:
                n_ok_in += 1
            msgs.append(
                u"OK [{0}] id {1}: {2}".format(
                    cara, eid.IntegerValue, (msg or u"").strip()
                )
            )
        else:
            n_fail += 1
            msgs.append(
                u"FALLO [{0}] id {1}: {2}".format(
                    cara, eid.IntegerValue, (msg or u"").strip()
                )
            )
    n_ok = n_ok_ex + n_ok_in
    return n_ok, n_ok_ex, n_ok_in, n_skip, n_fail, msgs, len(created)


def run(uidoc):
    err = _import_error_message()
    if err:
        TaskDialog.Show(
            u"Area Reinforcement (ext+int) → L + 135°",
            u"No se pudo importar módulos necesarios. Coloque los .py en la misma carpeta.\n\n{0}".format(
                err
            ),
        )
        return
    doc = uidoc.Document
    ids = uidoc.Selection.GetElementIds()
    if ids is None or ids.Count == 0:
        TaskDialog.Show(
            u"Area Reinforcement (ext+int) → L + 135°",
            u"Selecciona un Area Reinforcement (un elemento).",
        )
        return
    if int(ids.Count) != 1:
        TaskDialog.Show(
            u"Area Reinforcement (ext+int) → L + 135°",
            u"Selecciona un solo Area Reinforcement.",
        )
        return
    eid = ids[0]
    ar = doc.GetElement(eid)
    if not isinstance(ar, AreaReinforcement):
        TaskDialog.Show(
            u"Area Reinforcement (ext+int) → L + 135°",
            u"El elemento no es un Area Reinforcement.",
        )
        return
    n_ok, n_ex, n_in, n_skip, n_fail, log, n_created = _run_area_ambas_caras(doc, ar)
    summary = (
        u"Rebar tras quitar el área: {0}. "
        u"OK total={1} (exterior={2}, interior={3}), omitidas={4}, error={5}."
    ).format(n_created, n_ok, n_ex, n_in, n_skip, n_fail)
    tail = u"\n".join(log[:50])
    if len(log) > 50:
        tail += u"\n..."
    print(summary)
    print(tail)
    TaskDialog.Show(
        u"Area Reinforcement (ext+int) → L + 135°",
        u"{0}\n\n{1}".format(summary, tail),
    )


if __name__ == u"__main__":
    run(__revit__.ActiveUIDocument)  # noqa: F821
