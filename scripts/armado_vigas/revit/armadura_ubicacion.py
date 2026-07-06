# -*- coding: utf-8 -*-
"""
Parámetros de instancia en barras creadas por Armado vigas.

- ``Armadura_Ubicacion``: superior ``F'`` · inferior ``F``
- ``Armadura_Capa``: ``(1ºC.)``, ``(2ºC.)``, … según capa 1-based del modelo
- ``Armadura_En Lamina``: número de lámina desde el parámetro de vista ``Sheet Number``
"""

from __future__ import print_function

ARMADURA_UBICACION_PARAM = u"Armadura_Ubicacion"
ARMADURA_UBICACION_INFERIOR = u"F"
ARMADURA_UBICACION_SUPERIOR = u"F'"
ARMADURA_CAPA_PARAM = u"Armadura_Capa"
ARMADURA_EN_LAMINA_PARAM = u"Armadura_En Lamina"
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


def stamp_armadura_ubicacion(rebar, es_cara_inferior=False):
    """Escribe ``Armadura_Ubicacion`` si el parámetro existe y es escribible."""
    if rebar is None:
        return False
    valor = _valor_ubicacion(bool(es_cara_inferior))
    try:
        p = rebar.LookupParameter(ARMADURA_UBICACION_PARAM)
        if p is None or p.IsReadOnly:
            return False
        p.Set(valor)
        return True
    except Exception:
        return False


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
