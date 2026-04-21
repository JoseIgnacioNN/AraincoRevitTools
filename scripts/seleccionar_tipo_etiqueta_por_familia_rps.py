# -*- coding: utf-8 -*-
"""
Resuelve un tipo de etiqueta (FamilySymbol) por nombre de familia y nombre de tipo.

Ejecutable en RevitPythonShell (RPS): edita FAMILY_NAME y TYPE_NAME abajo y ejecuta.

Tras encontrar el símbolo:
- Lo activa en el documento si hace falta (requisito habitual para colocar etiquetas).
- Deja listo su uso con IndependentTag.Create: registra TAG_SYMBOL_ID y TAG_FAMILY_SYMBOL
  en el módulo __main__ de la sesión RPS (y en las variables globales de este script).

Flujo típico: ejecuta este script; luego ejecuta etiquetar_area_reinforcement_rps.py
(sin pegar IDs a mano) en la misma sesión de consola.

Si no encuentra la familia o el tipo, lista candidatos en la consola de RPS.
"""

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import Family, FamilySymbol, FilteredElementCollector, Transaction
from Autodesk.Revit.UI import TaskDialog

# Tras un run exitoso de main(), quedan rellenados (y en __main__ para otros scripts RPS).
TAG_FAMILY_SYMBOL = None
TAG_SYMBOL_ID = None

# --- Editar antes de ejecutar ---
FAMILY_NAME = u"EST_A_STRUCTURAL AREA REINFORCEMENT TAG_PLANTA_MALLA"
TYPE_NAME = u"Malla Inferior"

# True = comparar nombres sin distinguir mayúsculas/minúsculas
CASE_INSENSITIVE = True


def _get_doc():
    try:
        return doc
    except NameError:
        return __revit__.ActiveUIDocument.Document


def _norm(s):
    if s is None:
        return u""
    if CASE_INSENSITIVE:
        return s.strip().lower()
    return s.strip()


def _find_family(document, family_name_wanted):
    wanted = _norm(family_name_wanted)
    matches = []
    for fam in FilteredElementCollector(document).OfClass(Family):
        if _norm(fam.Name) == wanted:
            matches.append(fam)
    return matches


def _symbol_names_in_family(document, family):
    names = []
    for sid in family.GetFamilySymbolIds():
        sym = document.GetElement(sid)
        if sym is not None:
            names.append(sym.Name)
    return sorted(names)


def _find_symbol_in_family(document, family, type_name_wanted):
    wanted = _norm(type_name_wanted)
    for sid in family.GetFamilySymbolIds():
        sym = document.GetElement(sid)
        if sym is None:
            continue
        if not isinstance(sym, FamilySymbol):
            continue
        if _norm(sym.Name) == wanted:
            return sym
    return None


def _ensure_symbol_active(document, sym):
    if sym.IsActive:
        return True, None
    trans = Transaction(document, u"Activar tipo de etiqueta")
    trans.Start()
    try:
        sym.Activate()
        document.Regenerate()
        trans.Commit()
        return True, None
    except Exception as ex:
        trans.RollBack()
        return False, str(ex)


def main():
    global TAG_FAMILY_SYMBOL, TAG_SYMBOL_ID
    TAG_FAMILY_SYMBOL = None
    TAG_SYMBOL_ID = None
    try:
        import __main__ as _main

        _main.TAG_FAMILY_SYMBOL = None
        _main.TAG_SYMBOL_ID = None
    except Exception:
        pass

    document = _get_doc()
    fn = FAMILY_NAME.strip()
    tn = TYPE_NAME.strip()

    if not fn or not tn:
        msg = u"Define FAMILY_NAME y TYPE_NAME no vacíos en el script."
        print(msg)
        try:
            TaskDialog.Show(u"Tipo de etiqueta", msg)
        except Exception:
            pass
        return

    families = _find_family(document, fn)
    if not families:
        print(u"No se encontró ninguna familia con nombre: {!r}".format(fn))
        print(u"Familias cargadas (primeras 80, orden alfabético):")
        all_names = sorted(f.Name for f in FilteredElementCollector(document).OfClass(Family))
        for n in all_names[:80]:
            print(u"  - {}".format(n))
        if len(all_names) > 80:
            print(u"  ... (+{} más)".format(len(all_names) - 80))
        return

    if len(families) > 1:
        print(u"Aviso: hay {} familias que coinciden con {!r}; se usa la primera.".format(
            len(families), fn))

    family = families[0]
    sym = _find_symbol_in_family(document, family, tn)
    if sym is None:
        print(u"Familia encontrada: {!r}".format(family.Name))
        print(u"No hay tipo con nombre: {!r}".format(tn))
        print(u"Tipos disponibles en esa familia:")
        for n in _symbol_names_in_family(document, family):
            print(u"  - {}".format(n))
        return

    ok_act, err_act = _ensure_symbol_active(document, sym)
    if not ok_act:
        print(u"Aviso: no se pudo activar el tipo de etiqueta: {}".format(err_act))

    TAG_FAMILY_SYMBOL = sym
    TAG_SYMBOL_ID = sym.Id
    try:
        import __main__ as _main

        _main.TAG_FAMILY_SYMBOL = sym
        _main.TAG_SYMBOL_ID = sym.Id
    except Exception:
        pass

    eid = sym.Id.IntegerValue
    resumen = (
        u"FamilySymbol OK (activo para colocar etiquetas)\n"
        u"  Familia: {}\n"
        u"  Tipo:    {}\n"
        u"  ElementId: {}  →  IndependentTag.Create(..., ElementId({}), ...)".format(
            family.Name, sym.Name, eid, eid)
    )
    print(resumen)
    print(
        u"\nEn esta sesión RPS: ya está en __main__.TAG_SYMBOL_ID; "
        u"etiquetar_area_reinforcement_rps.py lo tomará solo si TAG_SYMBOL_ID del script sigue None."
    )
    print(u"Copia manual si lo necesitas: TAG_SYMBOL_ID = ElementId({})".format(eid))
    try:
        TaskDialog.Show(u"Tipo de etiqueta", resumen)
    except Exception:
        pass


main()
