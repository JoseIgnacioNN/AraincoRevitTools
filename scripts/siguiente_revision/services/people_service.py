# -*- coding: utf-8 -*-
"""
PeopleService — carga y mapeo de personas desde personas.json.

Lee el archivo de directorio de personas y construye:
- Lista de nombres para mostrar en el combo (nombre completo o abreviación si no hay nombre).
- Mapa display→sheet_value para convertir el ítem seleccionado al texto que va a la lámina.
"""

from __future__ import print_function

import codecs
import json
import os

try:
    unicode
except NameError:
    unicode = str

from siguiente_revision.constants import (
    ISSUES_DIR,
    PERSONAS_FILE_NAME,
    PERSONA_ROL_MODELADOR,
    PERSONA_ROL_INGENIERO,
    DIBUJO_INICIALES,
)

PERSONAS_FILE = os.path.join(ISSUES_DIR, PERSONAS_FILE_NAME)


def _normalize_rol(val):
    s = (val or u"").strip()
    if s == PERSONA_ROL_INGENIERO:
        return PERSONA_ROL_INGENIERO
    return PERSONA_ROL_MODELADOR


def load_display_map(target_rol):
    """
    Carga personas.json y devuelve (display_list, display_to_sheet_map) para el rol dado.

    - display_list: nombres ordenados para el combo (nombre completo, o abreviación si no hay nombre).
    - display_to_sheet_map: dict {display_label: sheet_value} donde sheet_value es la abreviación
      (si existe) o el nombre completo.

    Si no se puede leer el archivo, devuelve listas vacías.
    """
    items_order = []
    display_to_sheet = {}
    if target_rol not in (PERSONA_ROL_MODELADOR, PERSONA_ROL_INGENIERO):
        return items_order, display_to_sheet
    if not os.path.isfile(PERSONAS_FILE):
        return items_order, display_to_sheet
    try:
        with codecs.open(PERSONAS_FILE, "r", "utf-8-sig") as f:
            data = json.load(f)
    except Exception:
        return items_order, display_to_sheet
    if not isinstance(data, list):
        return items_order, display_to_sheet
    seen = set()
    for p in data:
        if not isinstance(p, dict):
            continue
        if _normalize_rol(p.get("rol", u"")) != target_rol:
            continue
        abr = (p.get("abreviacion") or u"").strip()
        nom = (p.get("nombre") or u"").strip()
        display = nom if nom else abr
        if not display:
            continue
        sheet_val = abr if abr else nom
        key = display.lower()
        if key in seen:
            continue
        seen.add(key)
        items_order.append(display)
        display_to_sheet[display] = sheet_val
    try:
        items_order.sort(key=lambda s: s.lower())
    except Exception:
        items_order.sort()
    return items_order, display_to_sheet


def display_to_sheet_value(selected_display, mapping):
    """
    Convierte el ítem seleccionado del combo (nombre mostrado) al texto para
    los parámetros de lámina (abreviación).

    Si el label no está en el mapa, se devuelve el label sin cambio.
    """
    label = unicode(selected_display or u"").strip()
    if not label:
        return u""
    m = mapping or {}
    return unicode(m.get(label, label)).strip()


def fallback_items():
    """
    Devuelve la lista de iniciales de respaldo cuando personas.json no está disponible.
    El mapa es identidad (display == sheet_value).
    """
    items = list(DIBUJO_INICIALES)
    return items, {x: x for x in items}
