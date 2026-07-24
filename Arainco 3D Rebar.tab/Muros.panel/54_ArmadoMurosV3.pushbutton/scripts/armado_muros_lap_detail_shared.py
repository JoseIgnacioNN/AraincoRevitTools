# -*- coding: utf-8 -*-
"""Detail Component de empalme (copia mínima portable para Armado Muros 34)."""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import BuiltInCategory, FilteredElementCollector, FamilySymbol

_LAP_DETAIL_DEFAULT_FAMILY_NAME = u"EST_D_DEATIL ITEM_EMPALME"
_LAP_DETAIL_DEFAULT_TYPE_NAME = u"Empalme"
_LAP_DETAIL_ALT_FAMILY_NAMES = (
    u"EST_D_DEATIL ITEM_EMPALME",
    u"EST_D_DETAIL ITEM_EMPALME",
)


def _norm_name(s):
    try:
        t = unicode(s or u"")
    except Exception:
        t = u""
    try:
        t = t.replace(u"\u00A0", u" ")
    except Exception:
        pass
    t = u" ".join([p for p in t.strip().lower().split() if p])
    return t


def _find_fixed_lap_detail_symbol_id(doc):
    if doc is None:
        return None, u"No hay documento activo."
    fam_target = unicode(_LAP_DETAIL_DEFAULT_FAMILY_NAME or u"").strip().lower()
    fam_alt_targets = set()
    try:
        for nm in _LAP_DETAIL_ALT_FAMILY_NAMES:
            t = unicode(nm or u"").strip().lower()
            if t:
                fam_alt_targets.add(t)
    except Exception:
        pass
    if fam_target:
        fam_alt_targets.add(fam_target)
    typ_target = unicode(_LAP_DETAIL_DEFAULT_TYPE_NAME or u"").strip().lower()
    fam_alt_targets = set([_norm_name(x) for x in fam_alt_targets if _norm_name(x)])
    fam_target = _norm_name(fam_target)
    typ_target = _norm_name(typ_target)
    try:
        syms = list(
            FilteredElementCollector(doc)
            .OfClass(FamilySymbol)
            .OfCategory(BuiltInCategory.OST_DetailComponents)
        )
    except Exception:
        syms = []
    if not syms:
        return None, u"No hay Detail Components en el proyecto."
    for sym in syms:
        if sym is None:
            continue
        fam = u""
        typ = u""
        try:
            fam = _norm_name(getattr(sym, "FamilyName", None))
        except Exception:
            fam = u""
        if not fam:
            try:
                if sym.Family is not None:
                    fam = _norm_name(sym.Family.Name)
            except Exception:
                fam = u""
        try:
            typ = _norm_name(getattr(sym, "Name", None))
        except Exception:
            typ = u""
        fam_ok = (fam in fam_alt_targets) if fam_alt_targets else (fam == fam_target)
        if fam_ok and typ == typ_target:
            try:
                return sym.Id, None
            except Exception:
                break
    for sym in syms:
        if sym is None:
            continue
        try:
            fam = _norm_name(getattr(sym, "FamilyName", None))
        except Exception:
            fam = u""
        if not fam:
            try:
                if sym.Family is not None:
                    fam = _norm_name(sym.Family.Name)
            except Exception:
                fam = u""
        fam_ok = (fam in fam_alt_targets) if fam_alt_targets else (fam == fam_target)
        if fam_ok:
            try:
                return sym.Id, (
                    u"Detail Component fijo: no se encontró tipo exacto '{0}', se usó otro tipo de la familia '{1}'."
                    .format(_LAP_DETAIL_DEFAULT_TYPE_NAME, _LAP_DETAIL_DEFAULT_FAMILY_NAME)
                )
            except Exception:
                pass
    return None, (
        u"No se encontró Detail Component fijo '{0} : {1}'."
        .format(_LAP_DETAIL_DEFAULT_FAMILY_NAME, _LAP_DETAIL_DEFAULT_TYPE_NAME)
    )
