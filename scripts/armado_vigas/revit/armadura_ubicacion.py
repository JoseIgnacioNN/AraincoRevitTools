# -*- coding: utf-8 -*-
"""
Parámetros de instancia en barras creadas por Armado vigas.

- ``Armadura_Ubicacion``: superior ``F'`` · inferior ``F`` · laterales ``Lateral``
- ``Armadura_Capa``: ``(1ºC.)``, ``(2ºC.)``, … según capa 1-based del modelo
- ``Armadura_En Lamina``: número de lámina desde el parámetro de vista ``Sheet Number``
- ``Armadura_Eje``: eje de la elevación/sección activa (parámetro de vista ``Armadura_Eje``)
- ``Armadura_Arainco``: Yes
- ``Armadura_Malla``: No
- ``Armadura_Nivel``: nombre del nivel de referencia del host (viga)
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import BuiltInParameter, ElementId, StorageType

ARMADURA_UBICACION_PARAM = u"Armadura_Ubicacion"
ARMADURA_UBICACION_INFERIOR = u"F"
ARMADURA_UBICACION_SUPERIOR = u"F'"
ARMADURA_UBICACION_LATERAL = u"Lateral"
ARMADURA_CAPA_PARAM = u"Armadura_Capa"
ARMADURA_EN_LAMINA_PARAM = u"Armadura_En Lamina"
ARMADURA_EJE_PARAM = u"Armadura_Eje"
ARMADURA_ARAINCO_PARAM = u"Armadura_Arainco"
ARMADURA_MALLA_PARAM = u"Armadura_Malla"
ARMADURA_NIVEL_PARAM = u"Armadura_Nivel"
SHEET_NUMBER_VIEW_PARAM = u"Sheet Number"


def _parametro_como_texto(param):
    if param is None:
        return u""
    try:
        if not param.HasValue:
            return u""
        s = param.AsString()
        if s is not None and unicode(s).strip():
            return unicode(s).strip()
        vs = param.AsValueString()
        if vs is not None and unicode(vs).strip():
            return unicode(vs).strip()
        return u""
    except Exception:
        return u""


def leer_sheet_number_desde_vista(view):
    """
    Lee ``Sheet Number`` de la vista activa (parámetro de instancia en Vistas).

    Si la vista no tiene el parámetro o está vacío, devuelve cadena vacía.
    """
    if view is None:
        return u""
    try:
        p = view.LookupParameter(SHEET_NUMBER_VIEW_PARAM)
        return _parametro_como_texto(p)
    except Exception:
        return u""


def _valor_ubicacion(es_cara_inferior):
    return ARMADURA_UBICACION_INFERIOR if es_cara_inferior else ARMADURA_UBICACION_SUPERIOR


def stamp_armadura_ubicacion(rebar, es_cara_inferior=False, valor=None):
    """Escribe ``Armadura_Ubicacion`` si el parámetro existe y es escribible.

    ``valor``: texto explícito (p. ej. ``Lateral``). Si es ``None``, usa
    ``F`` / ``F'`` según ``es_cara_inferior``.
    """
    if rebar is None:
        return False
    if valor is None:
        valor = _valor_ubicacion(bool(es_cara_inferior))
    else:
        try:
            valor = unicode(valor or u"").strip()
        except Exception:
            valor = u""
    if not valor:
        return False
    try:
        p = rebar.LookupParameter(ARMADURA_UBICACION_PARAM)
        if p is None or p.IsReadOnly:
            return False
        p.Set(valor)
        return True
    except Exception:
        return False


def aplicar_armadura_ubicacion_laterales(rebars_laterales):
    """Estampa ``Armadura_Ubicacion = Lateral`` en rebars del alma."""
    if not rebars_laterales:
        return 0
    n = 0
    for rb in rebars_laterales:
        if stamp_armadura_ubicacion(rb, valor=ARMADURA_UBICACION_LATERAL):
            n += 1
    return n


def armadura_capa_valor_desde_layer(layer_num):
    """Capa 1 → ``(1ºC.)``, capa 2 → ``(2ºC.)``, etc."""
    try:
        n = int(layer_num)
    except Exception:
        n = 1
    return u"({0}ºC.)".format(max(1, n))


def _rebar_element_id_int(rebar):
    try:
        return int(rebar.Id.IntegerValue)
    except Exception:
        return None


def stamp_armadura_capa(rebar, layer_num=1):
    """Escribe ``Armadura_Capa`` si el parámetro existe y es escribible."""
    if rebar is None:
        return False
    valor = armadura_capa_valor_desde_layer(layer_num)
    try:
        p = rebar.LookupParameter(ARMADURA_CAPA_PARAM)
        if p is None or p.IsReadOnly:
            return False
        p.Set(valor)
        return True
    except Exception:
        return False


def _layer_num_for_rebar(rebar, layer_by_id, default=1):
    rid = _rebar_element_id_int(rebar)
    if rid is None or not layer_by_id:
        return default
    try:
        return int(layer_by_id.get(rid, default))
    except Exception:
        return default


def aplicar_armadura_ubicacion_longitudinales(rebars_by_side):
    """
    Aplica ``Armadura_Ubicacion`` a rebars longitudinales por cara.

    ``rebars_by_side``: ``{"sup": [...], "inf": [...], "layer_by_id": {id: capa}}``
    """
    if not rebars_by_side:
        return 0
    n = 0
    for rb in rebars_by_side.get(u"sup") or []:
        if stamp_armadura_ubicacion(rb, es_cara_inferior=False):
            n += 1
    for rb in rebars_by_side.get(u"inf") or []:
        if stamp_armadura_ubicacion(rb, es_cara_inferior=True):
            n += 1
    return n


def aplicar_armadura_capa_longitudinales(rebars_by_side):
    """
    Aplica ``Armadura_Capa`` a rebars longitudinales sup/inf.

    Usa ``layer_by_id`` en ``rebars_by_side`` (capa 1-based al crear la barra).
    """
    if not rebars_by_side:
        return 0
    layer_by_id = rebars_by_side.get(u"layer_by_id") or {}
    n = 0
    for rb in (rebars_by_side.get(u"sup") or []) + (rebars_by_side.get(u"inf") or []):
        layer_num = _layer_num_for_rebar(rb, layer_by_id)
        if stamp_armadura_capa(rb, layer_num):
            n += 1
    return n


def stamp_armadura_en_lamina(rebar, sheet_number=u""):
    """Escribe ``Armadura_En Lamina`` si el parámetro existe y es escribible."""
    if rebar is None:
        return False
    try:
        valor = unicode(sheet_number or u"").strip()
    except Exception:
        valor = u""
    try:
        p = rebar.LookupParameter(ARMADURA_EN_LAMINA_PARAM)
        if p is None or p.IsReadOnly:
            return False
        p.Set(valor)
        return True
    except Exception:
        return False


def aplicar_armadura_en_lamina(rebars, view, rebars_laterales=None):
    """
    Aplica ``Armadura_En Lamina`` a todas las barras creadas en la corrida.

    ``rebars``: longitudinales + estribos/confinamiento.
    ``rebars_laterales``: barras laterales (lista aparte en colocar).
    """
    sheet_number = leer_sheet_number_desde_vista(view)
    todos = list(rebars or [])
    if rebars_laterales:
        todos.extend(list(rebars_laterales))
    if not todos:
        return 0
    n = 0
    for rb in todos:
        if stamp_armadura_en_lamina(rb, sheet_number):
            n += 1
    return n


def leer_armadura_eje_desde_vista(view):
    """
    Lee ``Armadura_Eje`` de la vista activa (elevación/sección que se está armando).

    Si la vista no tiene el parámetro o está vacío, devuelve cadena vacía.
    """
    if view is None:
        return u""
    try:
        p = view.LookupParameter(ARMADURA_EJE_PARAM)
        return _parametro_como_texto(p)
    except Exception:
        return u""


def stamp_armadura_eje(rebar, eje_valor=u""):
    """Escribe ``Armadura_Eje`` si el parámetro existe y es escribible."""
    if rebar is None:
        return False
    try:
        valor = unicode(eje_valor or u"").strip()
    except Exception:
        valor = u""
    if not valor:
        return False
    try:
        p = rebar.LookupParameter(ARMADURA_EJE_PARAM)
        if p is None or p.IsReadOnly:
            return False
        p.Set(valor)
        return True
    except Exception:
        return False


def aplicar_armadura_eje(rebars, view, rebars_laterales=None):
    """
    Aplica ``Armadura_Eje`` (eje de la elevación activa) a todas las barras
    creadas en la corrida.

    ``rebars``: longitudinales + estribos/confinamiento.
    ``rebars_laterales``: barras laterales (lista aparte en colocar).
    """
    eje_valor = leer_armadura_eje_desde_vista(view)
    if not eje_valor:
        return 0
    todos = list(rebars or [])
    if rebars_laterales:
        todos.extend(list(rebars_laterales))
    if not todos:
        return 0
    n = 0
    for rb in todos:
        if stamp_armadura_eje(rb, eje_valor):
            n += 1
    return n


def _set_yes_no_param(element, param_name, yes=True):
    """Escribe Yes/No (Integer 0/1, bool o texto)."""
    if element is None or not param_name:
        return False
    try:
        p = element.LookupParameter(param_name)
    except Exception:
        p = None
    if p is None or p.IsReadOnly:
        return False
    try:
        st = p.StorageType
        if st == StorageType.Integer:
            p.Set(1 if yes else 0)
            return True
    except Exception:
        pass
    if yes:
        candidates = (1, True, u"1", u"Yes", u"yes", u"YES", u"Sí", u"SI")
    else:
        candidates = (0, False, u"0", u"No", u"no", u"NO")
    for val in candidates:
        try:
            p.Set(val)
            return True
        except Exception:
            continue
    try:
        p.SetValueString(u"Yes" if yes else u"No")
        return True
    except Exception:
        return False


def stamp_armadura_arainco(rebar, yes=True):
    """``Armadura_Arainco`` = Yes (herramientas Arainco/Bizards)."""
    return _set_yes_no_param(rebar, ARMADURA_ARAINCO_PARAM, yes=yes)


def stamp_armadura_malla(rebar, yes=False):
    """``Armadura_Malla`` = No en armado de vigas (no es malla de muro/losa)."""
    return _set_yes_no_param(rebar, ARMADURA_MALLA_PARAM, yes=yes)


def stamp_armadura_nivel(rebar, nivel_nombre=u""):
    """Escribe ``Armadura_Nivel`` (nombre de nivel)."""
    if rebar is None:
        return False
    try:
        valor = unicode(nivel_nombre or u"").strip()
    except Exception:
        valor = u""
    if not valor:
        return False
    try:
        p = rebar.LookupParameter(ARMADURA_NIVEL_PARAM)
        if p is None or p.IsReadOnly:
            return False
        p.Set(valor)
        return True
    except Exception:
        return False


def _rebar_document(rebar):
    try:
        return rebar.Document
    except Exception:
        return None


def _rebar_host_element(rebar, doc=None):
    if rebar is None:
        return None
    document = doc or _rebar_document(rebar)
    if document is None:
        return None
    try:
        hid = rebar.GetHostId()
    except Exception:
        return None
    if hid is None or hid == ElementId.InvalidElementId:
        return None
    try:
        return document.GetElement(hid)
    except Exception:
        return None


def _nivel_nombre_desde_element_id(doc, eid):
    if doc is None or eid is None:
        return u""
    try:
        if eid == ElementId.InvalidElementId:
            return u""
    except Exception:
        pass
    try:
        el = doc.GetElement(eid)
    except Exception:
        return u""
    if el is None:
        return u""
    try:
        name = el.Name
        if name:
            return unicode(name).strip()
    except Exception:
        pass
    return u""


def nivel_nombre_desde_host_viga(host, doc=None):
    """
    Nivel de referencia del host de viga (Structural Framing).

    Orden: Reference Level · Level · Schedule Level · LevelId.
    """
    if host is None:
        return u""
    document = doc
    if document is None:
        try:
            document = host.Document
        except Exception:
            document = None
    if document is None:
        return u""

    for bip_name in (
        u"INSTANCE_REFERENCE_LEVEL_PARAM",
        u"LEVEL_PARAM",
        u"SCHEDULE_LEVEL_PARAM",
        u"INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM",
    ):
        try:
            bip = getattr(BuiltInParameter, bip_name, None)
            if bip is None:
                continue
            p = host.get_Parameter(bip)
            if p is None or not p.HasValue:
                continue
            if p.StorageType == StorageType.ElementId:
                name = _nivel_nombre_desde_element_id(document, p.AsElementId())
                if name:
                    return name
        except Exception:
            continue

    for n in (u"Reference Level", u"Nivel de referencia", u"Level", u"Nivel"):
        try:
            p = host.LookupParameter(n)
            if p is None or not p.HasValue:
                continue
            if p.StorageType == StorageType.ElementId:
                name = _nivel_nombre_desde_element_id(document, p.AsElementId())
                if name:
                    return name
            txt = _parametro_como_texto(p)
            if txt:
                return txt
        except Exception:
            continue

    try:
        lid = host.LevelId
        name = _nivel_nombre_desde_element_id(document, lid)
        if name:
            return name
    except Exception:
        pass
    return u""


def stamp_armadura_nivel_desde_host(rebar, doc=None):
    """``Armadura_Nivel`` desde el host estructural del Rebar (viga)."""
    if rebar is None:
        return False
    document = doc or _rebar_document(rebar)
    host = _rebar_host_element(rebar, document)
    name = nivel_nombre_desde_host_viga(host, document)
    if not name:
        return False
    return stamp_armadura_nivel(rebar, name)


def aplicar_marca_parametros_armado_vigas(
    rebars, view=None, rebars_laterales=None, document=None
):
    """
    Estampa en **toda** la corrida de Armado vigas:

    - ``Armadura_Eje`` ← vista activa
    - ``Armadura_Arainco`` = Yes
    - ``Armadura_Malla`` = No
    - ``Armadura_Nivel`` ← nivel de referencia del host viga

    Returns:
        dict con contadores ``eje``, ``arainco``, ``malla``, ``nivel``.
    """
    todos = list(rebars or [])
    if rebars_laterales:
        todos.extend(list(rebars_laterales))
    stats = {
        u"eje": 0,
        u"arainco": 0,
        u"malla": 0,
        u"nivel": 0,
        u"total": len(todos),
    }
    if not todos:
        return stats

    eje_valor = leer_armadura_eje_desde_vista(view)
    for rb in todos:
        if rb is None:
            continue
        if document is not None:
            try:
                rb2 = document.GetElement(rb.Id)
                if rb2 is not None:
                    rb = rb2
            except Exception:
                pass

        if eje_valor and stamp_armadura_eje(rb, eje_valor):
            stats[u"eje"] += 1
        if stamp_armadura_arainco(rb, yes=True):
            stats[u"arainco"] += 1
        if stamp_armadura_malla(rb, yes=False):
            stats[u"malla"] += 1
        if stamp_armadura_nivel_desde_host(rb, doc=document):
            stats[u"nivel"] += 1
    return stats
