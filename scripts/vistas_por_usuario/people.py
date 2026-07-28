# -*- coding: utf-8 -*-
"""
Directorio de personas — misma fuente que Siguiente Revisión / BIMIssue.

Usa ``Y:\\00_SERVIDOR DE INCIDENCIAS\\personas.json``.
El código de clasificación de vistas es la **abreviación** del JSON
(convención Python con puntos, p. ej. ``J.N.N.``).

Incluye caché corta en AppDomain para no releer la red en cada apertura de UI.
"""

from __future__ import print_function

import os
import time

try:
    unicode
except NameError:
    unicode = str

_CACHE_KEY = u"BIMTools.VistasPorUsuario.ModeladoresCache"
_CACHE_TTL_S = 120.0


def _cache_get():
    try:
        from System import AppDomain

        payload = AppDomain.CurrentDomain.GetData(_CACHE_KEY)
        if not payload or not isinstance(payload, tuple) or len(payload) != 3:
            return None
        ts, items, mapping = payload
        if (time.time() - float(ts)) > _CACHE_TTL_S:
            return None
        return list(items), dict(mapping)
    except Exception:
        return None


def _cache_set(items, mapping):
    try:
        from System import AppDomain

        AppDomain.CurrentDomain.SetData(
            _CACHE_KEY, (time.time(), list(items), dict(mapping))
        )
    except Exception:
        pass


def invalidate_modeladores_cache():
    """Invalida la caché tras editar personas.json desde la UI."""
    try:
        from System import AppDomain

        AppDomain.CurrentDomain.SetData(_CACHE_KEY, None)
    except Exception:
        pass


def personas_paths():
    """
    (issues_dir, personas_file) — misma ruta que Siguiente Revisión.

    Si el paquete no está disponible, construye la ruta canónica en Y:.
    """
    try:
        from siguiente_revision.services.people_service import PERSONAS_FILE
        from siguiente_revision.constants import ISSUES_DIR

        return ISSUES_DIR, PERSONAS_FILE
    except Exception:
        issues = u"Y:\\00_SERVIDOR DE INCIDENCIAS"
        return issues, os.path.join(issues, u"personas.json")


def load_modeladores():
    """
    Devuelve (display_list, display_to_code_map) para rol Modelador.

    - display_list: nombres para el combo
    - display_to_code_map: nombre → abreviación (clasificación de vistas)

    Si personas.json no está disponible o no tiene modeladores, usa el
    respaldo de iniciales con puntos de Siguiente Revisión.
    """
    cached = _cache_get()
    if cached is not None:
        return cached

    try:
        from siguiente_revision.services import people_service
        from siguiente_revision.constants import PERSONA_ROL_MODELADOR
    except Exception:
        items, mapping = _fallback_modeladores()
        _cache_set(items, mapping)
        return items, mapping

    items, mapping = people_service.load_display_map(PERSONA_ROL_MODELADOR)
    if not items:
        items, mapping = people_service.fallback_items()
    _cache_set(items, mapping)
    return items, mapping


def display_to_code(selected_display, mapping):
    """Nombre mostrado del combo → abreviación para Subclasificacion / nombres."""
    try:
        from siguiente_revision.services import people_service

        return people_service.display_to_sheet_value(selected_display, mapping)
    except Exception:
        label = unicode(selected_display or u"").strip()
        if not label:
            return u""
        m = mapping or {}
        return unicode(m.get(label, label)).strip()


def _fallback_modeladores():
    """Respaldo local si el paquete siguiente_revision no está en sys.path."""
    initials = (
        u"J.N.N.",
        u"P.C.C.",
        u"A.S.A.",
        u"C.M.H.",
        u"S.J.M.",
        u"R.P.C.",
        u"C.O.C.",
        u"H.C.P.",
        u"L.B.M.",
        u"J.N.O.",
        u"T.M.M.",
        u"B.L.A.",
        u"X.X.X.",
    )
    return list(initials), {x: x for x in initials}
