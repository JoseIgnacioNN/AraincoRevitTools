# -*- coding: utf-8 -*-
"""
Cota de anchura combinada: muros (Wall) y zapatas de muro (Wall Foundation).

Reutiliza la lógica de cota_muro_anchura.py y cota_fundacion_anchura.py.
La interfaz inicial permite elegir qué categorías cotar.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import Transaction, ViewPlan, ViewSheet, Wall, WallFoundation
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

try:
    from Autodesk.Revit.Exceptions import OperationCanceledException
except Exception:
    OperationCanceledException = Exception

from cota_fundacion_anchura import _cotar_anchura_wall_foundation
from cota_muro_anchura import _cotar_anchura_muro

_TITULO = u"Arainco: Cotas anchura muro y fundación"
_TX_COTAR = u"Arainco: Cotas anchura muro y fundación"


def _element_id_key(elem):
    try:
        return int(elem.Id.IntegerValue)
    except Exception:
        return None


def _desde_preseleccion(uidoc, doc, incluir_fundacion, incluir_muro):
    wfs = []
    walls = []
    vistos_wf = set()
    vistos_wall = set()
    try:
        for eid in uidoc.Selection.GetElementIds():
            el = doc.GetElement(eid)
            if el is None:
                continue
            if incluir_fundacion and isinstance(el, WallFoundation):
                key = _element_id_key(el)
                if key is not None and key in vistos_wf:
                    continue
                if key is not None:
                    vistos_wf.add(key)
                wfs.append(el)
            elif incluir_muro and isinstance(el, Wall):
                key = _element_id_key(el)
                if key is not None and key in vistos_wall:
                    continue
                if key is not None:
                    vistos_wall.add(key)
                walls.append(el)
    except Exception:
        pass
    return wfs, walls


def _filtro_seleccion(incluir_fundacion, incluir_muro):
    class _FiltroCombinado(ISelectionFilter):
        def AllowElement(self, elem):
            try:
                if elem is None:
                    return False
                if incluir_fundacion and isinstance(elem, WallFoundation):
                    return True
                if incluir_muro and isinstance(elem, Wall):
                    return True
                return False
            except Exception:
                return False

        def AllowReference(self, ref, pt):
            return False

    return _FiltroCombinado()


def _mensaje_picker(incluir_fundacion, incluir_muro):
    if incluir_fundacion and incluir_muro:
        return (
            u"Seleccione muros y/o zapatas de muro (Wall Foundation). "
            u"Finalizar en la cinta o Esc para cancelar."
        )
    if incluir_fundacion:
        return (
            u"Seleccione zapatas de muro (Wall Foundation). "
            u"Finalizar en la cinta o Esc para cancelar."
        )
    return (
        u"Seleccione muros (Wall). "
        u"Finalizar en la cinta o Esc para cancelar."
    )


def _pick_elementos(uidoc, incluir_fundacion, incluir_muro):
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            _filtro_seleccion(incluir_fundacion, incluir_muro),
            _mensaje_picker(incluir_fundacion, incluir_muro),
        )
    except OperationCanceledException:
        return [], []
    except Exception:
        return [], []
    if not refs:
        return [], []

    doc = uidoc.Document
    wfs = []
    walls = []
    vistos_wf = set()
    vistos_wall = set()
    for ref in refs:
        try:
            el = doc.GetElement(ref.ElementId)
        except Exception:
            continue
        if el is None:
            continue
        if incluir_fundacion and isinstance(el, WallFoundation):
            key = _element_id_key(el)
            if key is not None and key in vistos_wf:
                continue
            if key is not None:
                vistos_wf.add(key)
            wfs.append(el)
        elif incluir_muro and isinstance(el, Wall):
            key = _element_id_key(el)
            if key is not None and key in vistos_wall:
                continue
            if key is not None:
                vistos_wall.add(key)
            walls.append(el)
    return wfs, walls


def _obtener_elementos(uidoc, doc, incluir_fundacion, incluir_muro):
    wfs, walls = _desde_preseleccion(uidoc, doc, incluir_fundacion, incluir_muro)
    if wfs or walls:
        return wfs, walls
    return _pick_elementos(uidoc, incluir_fundacion, incluir_muro)


def _validar_vista(view):
    if isinstance(view, ViewSheet):
        return False, u"Abra una vista de modelo, no una lámina."
    if not isinstance(view, ViewPlan):
        return (
            False,
            u"La vista activa debe ser una planta (ViewPlan).\n"
            u"Abra una planta donde se vean los elementos y vuelva a ejecutar.",
        )
    return True, u""


def ejecutar_cotas(uidoc, incluir_fundacion, incluir_muro):
    """
    Crea cotas de anchura según las categorías indicadas.

    Returns:
        (ok, mensaje) — ok True si se creó al menos una cota.
    """
    if not incluir_fundacion and not incluir_muro:
        return False, u"Marque al menos una categoría para cotar."

    doc = uidoc.Document
    view = doc.ActiveView

    ok_vista, msg_vista = _validar_vista(view)
    if not ok_vista:
        return False, msg_vista

    wfs, walls = _obtener_elementos(uidoc, doc, incluir_fundacion, incluir_muro)
    if not wfs and not walls:
        return False, u""

    creadas = 0
    errores = []

    with Transaction(doc, _TX_COTAR) as t:
        t.Start()
        try:
            if incluir_fundacion:
                for wf in wfs:
                    try:
                        wf_id = wf.Id.IntegerValue
                    except Exception:
                        wf_id = u"?"
                    n, errs = _cotar_anchura_wall_foundation(doc, view, wf)
                    creadas += n
                    for err in errs:
                        errores.append(
                            u"Fundación Id {0}: {1}".format(wf_id, err)
                        )

            if incluir_muro:
                for wall in walls:
                    try:
                        wall_id = wall.Id.IntegerValue
                    except Exception:
                        wall_id = u"?"
                    n, errs = _cotar_anchura_muro(doc, view, wall)
                    creadas += n
                    for err in errs:
                        errores.append(u"Muro Id {0}: {1}".format(wall_id, err))

            t.Commit()
        except Exception as ex:
            try:
                t.RollBack()
            except Exception:
                pass
            return False, u"Error inesperado al crear las cotas:\n{}".format(ex)

    if creadas == 0:
        msg = u"No se pudo crear ninguna cota de anchura."
        if errores:
            msg += u"\n\nDetalles:\n" + u"\n".join(errores)
        return False, msg

    if errores:
        return True, u"\n".join(errores)

    partes = []
    if incluir_fundacion and wfs:
        partes.append(u"{0} fundación(es)".format(len(wfs)))
    if incluir_muro and walls:
        partes.append(u"{0} muro(s)".format(len(walls)))
    resumen = u", ".join(partes) if partes else u"elementos"
    return True, u"Se crearon {0} cota(s) de anchura ({1}).".format(creadas, resumen)


def resumen_vista(uidoc):
    """Texto de estado para la UI según la vista activa."""
    if uidoc is None:
        return u"No hay documento activo."
    view = uidoc.ActiveView
    ok, msg = _validar_vista(view)
    if not ok:
        return msg
    try:
        vname = view.Name
    except Exception:
        vname = u"Vista"
    return (
        u"Vista: {0} · Seleccione las categorías y pulse «Cotar anchura».".format(
            vname
        )
    )
