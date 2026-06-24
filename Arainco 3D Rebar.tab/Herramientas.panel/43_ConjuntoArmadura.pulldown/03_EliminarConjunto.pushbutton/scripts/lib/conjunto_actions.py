# -*- coding: utf-8 -*-
"""Operaciones sobre conjuntos de armadura por ``Armadura_Conjunto_GUID``."""

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from System.Collections.Generic import List
from Autodesk.Revit.DB import (
    BuiltInCategory,
    DetailCurve,
    ElementId,
    FamilyInstance,
    FilledRegion,
    TextNote,
    Transaction,
)
from Autodesk.Revit.DB.Structure import Rebar
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import (
    TaskDialog,
    TaskDialogCommonButtons,
    TaskDialogResult,
)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

from lib.corrida_guid import (
    ARMADURA_CONJUNTO_GUID_PARAM,
    collect_corrida_por_conjunto_guid,
    get_armadura_conjunto_guid,
)

_DIALOG_BASE = u"Arainco: Conjunto armadura"
_PICK_PROMPT = (
    u"Selecciona una barra (Rebar), empalme o croquis de despiece con "
    u"«Armadura_Conjunto_GUID» (p. ej. creados por Armado Muros o Armado Columnas)."
)


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _guid_snippet(guid, max_len=72):
    if not guid:
        return u""
    s = _as_unicode(guid).strip()
    if len(s) <= max_len:
        return s
    return s[: max_len - 1] + u"…"


class _FiltroCorridaReferencia(ISelectionFilter):
    def AllowElement(self, elem):
        if isinstance(elem, Rebar):
            return True
        if isinstance(elem, (DetailCurve, TextNote, FilledRegion)):
            return True
        if isinstance(elem, FamilyInstance):
            try:
                cat = elem.Category
                if cat is not None:
                    return int(cat.Id.IntegerValue) == int(
                        BuiltInCategory.OST_DetailComponents,
                    )
            except Exception:
                pass
        return False

    def AllowReference(self, reference, position):
        return False


def _element_id_list(ids):
    sel = List[ElementId]()
    for eid in ids or []:
        if eid is None:
            continue
        if isinstance(eid, ElementId):
            sel.Add(eid)
        else:
            try:
                sel.Add(ElementId(int(eid)))
            except Exception:
                pass
    return sel


def _summarize_counts(corrida):
    rebar_ids = (corrida or {}).get(u"rebar_ids") or []
    empalme_ids = (corrida or {}).get(u"empalme_ids") or []
    lienzo_ids = (corrida or {}).get(u"lienzo_ids") or []
    all_ids = (corrida or {}).get(u"all_ids") or []
    return len(rebar_ids), len(empalme_ids), len(lienzo_ids), len(all_ids)


def pick_conjunto_referencia(uidoc):
    """
    Pide una referencia con GUID de conjunto.

    Retorna ``(doc, guid, corrida)`` o ``None`` si cancela o falla.
    """
    if uidoc is None:
        TaskDialog.Show(_DIALOG_BASE, u"No hay documento activo.")
        return None
    doc = uidoc.Document
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            _FiltroCorridaReferencia(),
            _PICK_PROMPT,
        )
    except OperationCanceledException:
        return None
    except Exception as ex:
        TaskDialog.Show(_DIALOG_BASE, u"Error al seleccionar:\n{0}".format(ex))
        return None

    if ref is None:
        return None

    elem = doc.GetElement(ref.ElementId)
    guid = get_armadura_conjunto_guid(elem)
    if not guid:
        TaskDialog.Show(
            _DIALOG_BASE,
            u"El elemento elegido no tiene valor en «{0}».\n\n"
            u"Solo aplica a barras, empalmes o croquis de despiece con ese "
            u"parámetro vinculado.".format(ARMADURA_CONJUNTO_GUID_PARAM),
        )
        return None

    corrida = collect_corrida_por_conjunto_guid(doc, guid)
    all_ids = corrida.get(u"all_ids") or []
    if not all_ids:
        TaskDialog.Show(
            _DIALOG_BASE,
            u"No se encontraron elementos con GUID «{0}».".format(
                _guid_snippet(guid),
            ),
        )
        return None

    return doc, guid, corrida


def select_conjunto_en_modelo(uidoc, corrida):
    ids = (corrida or {}).get(u"all_ids") or []
    sel = _element_id_list(ids)
    if sel.Count < 1:
        return 0
    try:
        uidoc.Selection.SetElementIds(sel)
        uidoc.ShowElements(sel)
        return int(sel.Count)
    except Exception as ex:
        TaskDialog.Show(_DIALOG_BASE, u"Error al seleccionar:\n{0}".format(ex))
        return 0


def hide_conjunto_en_vista(doc, view, corrida):
    ids = (corrida or {}).get(u"all_ids") or []
    sel = _element_id_list(ids)
    if view is None:
        TaskDialog.Show(_DIALOG_BASE, u"No hay vista activa.")
        return 0
    if sel.Count < 1:
        return 0
    t = Transaction(doc, u"Arainco: Ocultar conjunto armadura")
    try:
        t.Start()
        view.HideElements(sel)
        t.Commit()
        return int(sel.Count)
    except Exception as ex:
        if t.HasStarted():
            try:
                t.RollBack()
            except Exception:
                pass
        TaskDialog.Show(_DIALOG_BASE, u"Error al ocultar:\n{0}".format(ex))
        return 0


def _confirm_delete_conjunto(guid, corrida):
    """Alerta nativa Revit: el usuario debe confirmar Sí antes de eliminar."""
    n_rebar, n_emp, n_lienzo, n_total = _summarize_counts(corrida)
    td = TaskDialog(_DIALOG_BASE)
    td.MainInstruction = u"¿Eliminar el conjunto de armadura?"
    td.MainContent = (
        u"El conjunto seleccionado será eliminado del modelo.\n\n"
        u"GUID: {0}\n\n"
        u"  · Barras estructurales: {1}\n"
        u"  · Empalmes (detail): {2}\n"
        u"  · Croquis de despiece: {3}\n"
        u"  · Total a eliminar: {4}\n\n"
        u"Esta acción no se puede deshacer fuera de Revit "
        u"(usa Deshacer inmediatamente si te equivocas).\n\n"
        u"¿Confirmas que deseas eliminar este conjunto?"
    ).format(
        _guid_snippet(guid), n_rebar, n_emp, n_lienzo, n_total,
    )
    td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
    try:
        return int(td.Show()) == int(TaskDialogResult.Yes)
    except Exception:
        return td.Show() == TaskDialogResult.Yes


def delete_conjunto(doc, uidoc, guid, corrida):
    ids = list((corrida or {}).get(u"all_ids") or [])
    if not ids or not guid:
        return 0

    n_rebar, n_emp, n_lienzo, n_total = _summarize_counts(corrida)
    if not _confirm_delete_conjunto(guid, corrida):
        return 0

    deleted = 0
    t = Transaction(doc, u"Arainco: Eliminar conjunto armadura")
    try:
        t.Start()
        for eid in ids:
            try:
                doc.Delete(eid)
                deleted += 1
            except Exception:
                pass
        t.Commit()
    except Exception as ex:
        if t.HasStarted():
            try:
                t.RollBack()
            except Exception:
                pass
        TaskDialog.Show(_DIALOG_BASE, u"Error al eliminar:\n{0}".format(ex))
        return 0

    try:
        uidoc.Selection.SetElementIds(List[ElementId]())
    except Exception:
        pass

    TaskDialog.Show(
        _DIALOG_BASE,
        u"Conjunto eliminado: {0} elemento(s) borrados "
        u"({1} barras, {2} empalmes, {3} croquis).\n"
        u"GUID: {4}".format(
            deleted, n_rebar, n_emp, n_lienzo, _guid_snippet(guid),
        ),
    )
    return deleted


def run_seleccionar(uiapp):
    uidoc = uiapp.ActiveUIDocument if uiapp is not None else None
    picked = pick_conjunto_referencia(uidoc)
    if not picked:
        return
    _doc, guid, corrida = picked
    n = select_conjunto_en_modelo(uidoc, corrida)
    if n < 1:
        return
    n_rebar, n_emp, n_lienzo, n_total = _summarize_counts(corrida)
    TaskDialog.Show(
        _DIALOG_BASE,
        u"Conjunto seleccionado ({0} elemento(s)).\n\n"
        u"GUID: {1}\n"
        u"  · Barras: {2}\n"
        u"  · Empalmes: {3}\n"
        u"  · Croquis: {4}".format(
            n_total, _guid_snippet(guid), n_rebar, n_emp, n_lienzo,
        ),
    )


def run_ocultar(uiapp):
    uidoc = uiapp.ActiveUIDocument if uiapp is not None else None
    picked = pick_conjunto_referencia(uidoc)
    if not picked:
        return
    doc, guid, corrida = picked
    view = uidoc.ActiveView if uidoc is not None else None
    n = hide_conjunto_en_vista(doc, view, corrida)
    if n < 1:
        return
    n_rebar, n_emp, n_lienzo, n_total = _summarize_counts(corrida)
    TaskDialog.Show(
        _DIALOG_BASE,
        u"Conjunto oculto en la vista activa ({0} elemento(s)).\n\n"
        u"GUID: {1}\n"
        u"  · Barras: {2}\n"
        u"  · Empalmes: {3}\n"
        u"  · Croquis: {4}".format(
            n_total, _guid_snippet(guid), n_rebar, n_emp, n_lienzo,
        ),
    )


def run_eliminar(uiapp):
    uidoc = uiapp.ActiveUIDocument if uiapp is not None else None
    picked = pick_conjunto_referencia(uidoc)
    if not picked:
        return
    doc, guid, corrida = picked
    delete_conjunto(doc, uidoc, guid, corrida)
