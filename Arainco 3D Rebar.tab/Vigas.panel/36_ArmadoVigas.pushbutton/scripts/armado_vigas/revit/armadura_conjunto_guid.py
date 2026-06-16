# -*- coding: utf-8 -*-
"""
GUID de corrida (``Armadura_Conjunto_GUID``) para elementos creados por Armado vigas.

Una ejecución de «Colocar armadura» comparte un mismo GUID en Rebar, etiquetas,
detail items de traslape y cotas, cuando el parámetro existe y es escribible.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import ElementId, FilteredElementCollector, IndependentTag

try:
    from conjunto_guid import (
        finalizar_armadura_conjunto_guid_ejecucion,
        iniciar_armadura_conjunto_guid_ejecucion,
        stamp_armadura_conjunto_guid,
        stamp_armadura_conjunto_guid_en_rebars,
    )
except Exception:
    finalizar_armadura_conjunto_guid_ejecucion = None
    iniciar_armadura_conjunto_guid_ejecucion = None
    stamp_armadura_conjunto_guid = None
    stamp_armadura_conjunto_guid_en_rebars = None


def iniciar_corrida_conjunto_guid():
    if iniciar_armadura_conjunto_guid_ejecucion is None:
        return None
    try:
        return iniciar_armadura_conjunto_guid_ejecucion()
    except Exception:
        return None


def finalizar_corrida_conjunto_guid():
    if finalizar_armadura_conjunto_guid_ejecucion is None:
        return
    try:
        finalizar_armadura_conjunto_guid_ejecucion()
    except Exception:
        pass


def stamp_elemento_si_permite(element, conjunto_guid=None):
    if element is None or stamp_armadura_conjunto_guid is None:
        return False
    try:
        return bool(stamp_armadura_conjunto_guid(element, conjunto_guid=conjunto_guid))
    except Exception:
        return False


def _rebar_ids_int(rebars):
    out = set()
    for rb in rebars or []:
        if rb is None:
            continue
        try:
            out.add(int(rb.Id.IntegerValue))
        except Exception:
            pass
    return out


def _tag_referencia_rebar(tag, rebar_ids_int):
    if tag is None or not rebar_ids_int:
        return False
    invalid = ElementId.InvalidElementId
    try:
        for tid in tag.GetTaggedLocalElementIds():
            try:
                if int(tid.IntegerValue) in rebar_ids_int:
                    return True
            except Exception:
                continue
    except Exception:
        pass
    try:
        for leid in tag.GetTaggedElementIds():
            try:
                link_inst = leid.LinkInstanceId
                if (
                    link_inst is not None
                    and link_inst != invalid
                    and int(link_inst.IntegerValue) >= 0
                ):
                    continue
            except Exception:
                pass
            for attr in (u"LinkedElementId", u"HostElementId"):
                try:
                    eid = getattr(leid, attr, None)
                    if eid is None:
                        continue
                    if int(eid.IntegerValue) in rebar_ids_int:
                        return True
                except Exception:
                    continue
    except Exception:
        pass
    return False


def stamp_etiquetas_rebar_en_vista(document, view, rebars, conjunto_guid=None):
    """Etiquetas ``IndependentTag`` asociadas a los Rebar de la corrida."""
    if document is None or view is None or stamp_armadura_conjunto_guid is None:
        return 0
    rebar_ids = _rebar_ids_int(rebars)
    if not rebar_ids:
        return 0
    n = 0
    try:
        tags = (
            FilteredElementCollector(document, view.Id)
            .OfClass(IndependentTag)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        tags = []
    for tag in tags or []:
        if tag is None:
            continue
        try:
            if not _tag_referencia_rebar(tag, rebar_ids):
                continue
        except Exception:
            continue
        if stamp_elemento_si_permite(tag, conjunto_guid):
            n += 1
    return n


def stamp_elementos_auxiliares(elements, conjunto_guid=None):
    n = 0
    for el in elements or []:
        if stamp_elemento_si_permite(el, conjunto_guid):
            n += 1
    return n


def aplicar_conjunto_guid_elementos_creados(
    document,
    view,
    rebars_longitudinales_y_estribos,
    rebars_laterales=None,
    lap_result=None,
    conjunto_guid=None,
):
    """
    Aplica ``Armadura_Conjunto_GUID`` a todo lo creado en la transacción de colocación.
    """
    if stamp_armadura_conjunto_guid is None:
        return 0

    todos_rebars = list(rebars_longitudinales_y_estribos or [])
    if rebars_laterales:
        todos_rebars.extend(list(rebars_laterales))

    n = 0
    if stamp_armadura_conjunto_guid_en_rebars is not None:
        try:
            n += int(
                stamp_armadura_conjunto_guid_en_rebars(
                    todos_rebars, conjunto_guid=conjunto_guid,
                )
                or 0
            )
        except Exception:
            pass

    aux = []
    if lap_result:
        aux.extend(list(lap_result.get(u"elements_created") or []))
    n += stamp_elementos_auxiliares(aux, conjunto_guid)

    if view is not None:
        n += stamp_etiquetas_rebar_en_vista(
            document, view, todos_rebars, conjunto_guid=conjunto_guid,
        )

    return n
