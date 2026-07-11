# -*- coding: utf-8 -*-
"""
RevitPythonShell (RPS) — Revit 2025 | IronPython 2.7 / 3.x

1. Recoge todas las ``Rebar`` libres (no forman parte de un Area Reinforcement).
2. Se queda con las que tienen Host Category = Floor (losa).
3. Rellena ``Armadura_Nivel`` con el nombre del nivel de la losa host.
4. Selecciona en el modelo las barras procesadas.

IMPORTANTE: en RPS usar File > Run script (o execfile / #load).
No pegar el archivo en la consola interactiva (>>>): las lineas en blanco
cierran los bloques y producen "expected an indented block".
"""

from __future__ import print_function

import os
import sys

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    Floor,
    StorageType,
    Transaction,
)
from Autodesk.Revit.DB.Structure import Rebar, RebarHostCategory
from Autodesk.Revit.UI import TaskDialog

# ── Documento activo (RPS / pyRevit) ─────────────────────────────────────────
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

ARMADURA_NIVEL_PARAM = u"Armadura_Nivel"
_TX_NAME = u"Arainco: Armadura_Nivel en Rebar Floor"
_HOST_CATEGORY_PARAM_NAMES = (
    u"Host Category",
    u"Categoría del anfitrión",
    u"Categoria del anfitrion",
    u"Categoría de anfitrión",
)
_FLOOR_LABEL_ALIASES = frozenset((
    u"floor",
    u"floors",
    u"suelo",
    u"suelos",
    u"losa",
    u"losas",
))


# ── sys.path → scripts/ (para importar conjunto_guid si está disponible) ─────
def _script_containing_dir():
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except NameError:
        return None


def _ensure_scripts_on_path():
    candidates = []
    d0 = _script_containing_dir()
    if d0:
        candidates.append(d0)
    try:
        home = os.path.expanduser(u"~")
        guess = os.path.join(
            home, u"CustomRevitExtensions", u"BIMTools.extension", u"scripts",
        )
        if os.path.isdir(guess):
            candidates.append(guess)
    except Exception:
        pass
    for p in candidates:
        if p and p not in sys.path and os.path.isdir(p):
            sys.path.insert(0, p)


_ensure_scripts_on_path()

try:
    from conjunto_guid import stamp_armadura_nivel as _stamp_nivel_ext
except Exception:
    _stamp_nivel_ext = None


# ── Helpers ──────────────────────────────────────────────────────────────────
def _as_text(value):
    if value is None:
        return u""
    try:
        return unicode(value)
    except NameError:
        return str(value)
    except Exception:
        try:
            return str(value)
        except Exception:
            return u""


def _norm_param_def_name(name):
    if name is None:
        return u""
    try:
        return _as_text(name).replace(u"\u00A0", u" ").strip()
    except Exception:
        return u""


def _find_element_parameter(element, param_name):
    if element is None or not param_name:
        return None
    target = _norm_param_def_name(param_name).lower()
    try:
        p = element.LookupParameter(param_name)
        if p is not None:
            return p
    except Exception:
        pass
    try:
        for p in element.Parameters:
            if p is None:
                continue
            try:
                dn = _norm_param_def_name(p.Definition.Name).lower()
            except Exception:
                continue
            if dn == target:
                return p
    except Exception:
        pass
    return None


def _set_string(element, param_name, valor):
    if not valor:
        return False
    p = _find_element_parameter(element, param_name)
    if p is None or p.IsReadOnly:
        return False
    try:
        if p.StorageType == StorageType.String:
            p.Set(valor)
            return True
    except Exception:
        pass
    try:
        p.Set(valor)
        return True
    except Exception:
        pass
    try:
        p.SetValueString(valor)
        return True
    except Exception:
        return False


def stamp_armadura_nivel(element, valor):
    if _stamp_nivel_ext is not None:
        try:
            return bool(_stamp_nivel_ext(element, valor))
        except Exception:
            pass
    return _set_string(element, ARMADURA_NIVEL_PARAM, valor)


def _element_id_int(eid):
    if eid is None:
        return None
    try:
        v = getattr(eid, "Value", None)
        if v is not None:
            return int(v)
    except Exception:
        pass
    try:
        return int(eid.IntegerValue)
    except Exception:
        pass
    return None


def _get_host_element(rebar):
    """Host estructural de una ``Rebar`` libre vía ``GetHostId()``."""
    if rebar is None:
        return None
    try:
        hid = rebar.GetHostId()
    except Exception:
        return None
    if hid is None or hid == ElementId.InvalidElementId:
        return None
    try:
        return doc.GetElement(hid)
    except Exception:
        return None


def _host_category_is_floor(rebar):
    """
    True si Host Category = Floor.

    Orden: propiedad API ``HostCategory``, parámetro de instancia, host Floor.
    """
    if rebar is None:
        return False

    # 1) API Rebar.HostCategory (Revit 2025)
    try:
        hc = rebar.HostCategory
        if hc is not None and int(hc) == int(RebarHostCategory.Floor):
            return True
        if hc is not None and int(hc) != int(RebarHostCategory.Floor):
            # Valor conocido distinto de Floor → no es losa
            if int(hc) != int(RebarHostCategory.Other):
                return False
    except Exception:
        pass

    # 2) Parámetro de instancia «Host Category»
    for param_name in _HOST_CATEGORY_PARAM_NAMES:
        try:
            p = rebar.LookupParameter(param_name)
            if p is None or not p.HasValue:
                continue
            # En schedules el valor es entero RebarHostCategory
            try:
                if p.StorageType == StorageType.Integer:
                    if int(p.AsInteger()) == int(RebarHostCategory.Floor):
                        return True
                    continue
            except Exception:
                pass
            s = p.AsString()
            if not s:
                try:
                    s = p.AsValueString()
                except Exception:
                    s = None
            if s and _as_text(s).strip().lower() in _FLOOR_LABEL_ALIASES:
                return True
        except Exception:
            continue

    # 3) Host real es Floor
    host = _get_host_element(rebar)
    return isinstance(host, Floor)


def _nivel_losa_como_string(floor):
    """Nombre del nivel de la losa host."""
    if floor is None or not isinstance(floor, Floor):
        return None
    lid = None
    try:
        lid = floor.LevelId
        if lid is None or lid == ElementId.InvalidElementId:
            lid = None
    except Exception:
        lid = None
    if lid is None:
        for bip_name in (
            u"INSTANCE_REFERENCE_LEVEL_PARAM",
            u"LEVEL_PARAM",
            u"SCHEDULE_LEVEL_PARAM",
        ):
            try:
                bip = getattr(BuiltInParameter, bip_name, None)
                if bip is None:
                    continue
                p = floor.get_Parameter(bip)
                if p is None or not p.HasValue or p.StorageType != StorageType.ElementId:
                    continue
                eid = p.AsElementId()
                if eid is not None and eid != ElementId.InvalidElementId:
                    lid = eid
                    break
            except Exception:
                pass
    if lid is None:
        return None
    try:
        level = doc.GetElement(lid)
        if level is None or level.Name is None:
            return None
        return _as_text(level.Name)
    except Exception:
        return None


def collect_rebar_libres_floor(document):
    """
    ``Rebar`` libres (no Area Reinforcement) con Host Category = Floor.

    Las barras de Area Reinforcement son ``RebarInSystem``; al filtrar solo
    ``Rebar`` quedan fuera del sistema de área.
    """
    out = []
    for el in FilteredElementCollector(document).OfClass(Rebar):
        if el is None or not isinstance(el, Rebar):
            continue
        if not _host_category_is_floor(el):
            continue
        out.append(el)
    return out


def _select_ids(ids):
    if not ids:
        return 0
    sel = List[ElementId]()
    for eid in ids:
        if eid is None or eid == ElementId.InvalidElementId:
            continue
        sel.Add(eid)
    if sel.Count < 1:
        return 0
    try:
        uidoc.Selection.SetElementIds(sel)
        return int(sel.Count)
    except Exception:
        return 0


def run():
    rebars = collect_rebar_libres_floor(doc)
    stats = {
        u"n_candidatas": len(rebars),
        u"n_nivel_ok": 0,
        u"n_sin_host_floor": 0,
        u"n_sin_nivel": 0,
        u"n_param_fail": 0,
        u"por_nivel": {},
        u"ids_ok": [],
    }

    t = Transaction(doc, _TX_NAME)
    t.Start()
    try:
        for barra in rebars:
            host = _get_host_element(barra)
            if not isinstance(host, Floor):
                stats[u"n_sin_host_floor"] += 1
                continue
            nivel = _nivel_losa_como_string(host)
            if not nivel:
                stats[u"n_sin_nivel"] += 1
                continue
            if stamp_armadura_nivel(barra, nivel):
                stats[u"n_nivel_ok"] += 1
                stats[u"ids_ok"].append(barra.Id)
                stats[u"por_nivel"][nivel] = stats[u"por_nivel"].get(nivel, 0) + 1
            else:
                stats[u"n_param_fail"] += 1
        t.Commit()
    except Exception:
        if t.HasStarted():
            t.RollBack()
        raise

    n_sel = _select_ids(stats[u"ids_ok"])
    stats[u"n_seleccionadas"] = n_sel
    return stats


def _format_resumen(stats):
    lines = [
        u"Rebar libres con Host Category = Floor: {0}".format(
            stats.get(u"n_candidatas") or 0,
        ),
        u"Armadura_Nivel escrito: {0}".format(stats.get(u"n_nivel_ok") or 0),
        u"Seleccionadas en modelo: {0}".format(stats.get(u"n_seleccionadas") or 0),
    ]
    if stats.get(u"n_sin_host_floor"):
        lines.append(
            u"Sin host Floor resoluble: {0}".format(stats[u"n_sin_host_floor"]),
        )
    if stats.get(u"n_sin_nivel"):
        lines.append(
            u"Sin nivel de losa: {0}".format(stats[u"n_sin_nivel"]),
        )
    if stats.get(u"n_param_fail"):
        lines.append(
            u"Parámetro Armadura_Nivel no escribible: {0}".format(
                stats[u"n_param_fail"],
            ),
        )
    por_nivel = stats.get(u"por_nivel") or {}
    if por_nivel:
        lines.append(u"Por nivel:")
        for nombre in sorted(por_nivel.keys(), key=lambda s: s.lower()):
            lines.append(u"  · {0}: {1}".format(nombre, por_nivel[nombre]))
    return u"\n".join(lines)


# ── Ejecución ────────────────────────────────────────────────────────────────
def main():
    try:
        stats = run()
        msg = _format_resumen(stats)
        print(msg)
        TaskDialog.Show(u"Arainco: Armadura_Nivel Rebar Floor", msg)
    except Exception as ex:
        err = u"Error: {0}".format(_as_text(ex))
        print(err)
        TaskDialog.Show(u"Arainco: Armadura_Nivel Rebar Floor", err)
        raise


main()
