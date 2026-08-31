# -*- coding: utf-8 -*-
"""Carga de listas de personas desde personas.json (editable)."""

from __future__ import print_function

import codecs
import json
import os

try:
    unicode
except NameError:
    unicode = str

from laminas_por_categoria.constants import (
    PERSONA_ROLES,
    PERSONAS_FALLBACK,
    PERSONAS_FILE_NAME,
)


def personas_json_path():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), PERSONAS_FILE_NAME)


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except Exception:
        return str(text)


def _clean_items(raw):
    items = []
    seen = set()
    if not isinstance(raw, (list, tuple)):
        return items
    for entry in raw:
        name = _as_unicode(entry).strip()
        if not name:
            continue
        key = name.lower()
        if key in seen:
            continue
        seen.add(key)
        items.append(name)
    return items


def _parse_role_block(raw, fallback_block):
    default_index = int(fallback_block.get(u"default_index", 0) or 0)
    items = list(fallback_block.get(u"items") or ())
    if isinstance(raw, (list, tuple)):
        parsed = _clean_items(raw)
        if parsed:
            items = parsed
    elif isinstance(raw, dict):
        parsed = _clean_items(raw.get(u"items") or raw.get(u"nombres"))
        if parsed:
            items = parsed
        try:
            if raw.get(u"default_index") is not None:
                default_index = int(raw.get(u"default_index"))
        except Exception:
            pass
    items_sorted = list(items)
    try:
        items_sorted.sort(key=lambda s: _as_unicode(s).lower())
    except Exception:
        items_sorted.sort()
    if default_index < 0:
        default_index = 0
    if items_sorted and default_index >= len(items_sorted):
        default_index = 0
    return items_sorted, default_index


def load_personas():
    """
    Devuelve dict rol → (items_ordenados, default_index).

    Si el JSON no existe o un rol está vacío/inválido, usa el respaldo del .dyn.
    """
    data = None
    path = personas_json_path()
    if os.path.isfile(path):
        try:
            with codecs.open(path, "r", "utf-8-sig") as f:
                data = json.load(f)
        except Exception:
            data = None
    if not isinstance(data, dict):
        data = {}

    result = {}
    for role in PERSONA_ROLES:
        fallback = PERSONAS_FALLBACK[role]
        result[role] = _parse_role_block(data.get(role), fallback)
    return result
