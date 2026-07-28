# -*- coding: utf-8 -*-
"""ElementId → int compatible Revit 2024–2026 (Value / IntegerValue)."""

from __future__ import print_function

try:
    from Autodesk.Revit.DB import ElementId
except Exception:
    ElementId = None


def revit_version_year(doc):
    """Año principal de Revit (2024, 2025, …) desde ``Application.VersionNumber``."""
    if doc is None:
        return 0
    try:
        app = doc.Application
    except Exception:
        return 0
    try:
        raw = app.VersionNumber
    except Exception:
        return 0
    try:
        s = str(raw).strip()
    except Exception:
        return 0
    if not s:
        pass
    else:
        for part in s.replace(u",", u".").split(u"."):
            part = part.strip()
            if len(part) == 4 and part.isdigit():
                return int(part)
        try:
            y = int(float(s))
            if y >= 2024:
                return y
        except Exception:
            pass
    try:
        app = doc.Application
        for attr in (u"VersionName", u"VersionBuild"):
            try:
                raw = getattr(app, attr, None)
            except Exception:
                raw = None
            if not raw:
                continue
            for token in str(raw).replace(u",", u" ").split():
                token = token.strip()
                if len(token) == 4 and token.isdigit():
                    y = int(token)
                    if y >= 2024:
                        return y
    except Exception:
        pass
    return 0


def element_id_to_int(eid):
    """
    ID numérico de ``ElementId``.

    Revit 2026+: ``Value``; versiones anteriores: ``IntegerValue``.
    """
    if eid is None:
        return None
    if ElementId is not None:
        try:
            if eid == ElementId.InvalidElementId:
                return None
        except Exception:
            pass
    try:
        return int(eid.Value)
    except Exception:
        pass
    try:
        return int(eid.IntegerValue)
    except Exception:
        pass
    try:
        return int(eid)
    except Exception:
        return None


def wall_id_int(wall):
    """``wall.Id`` → int o ``None``."""
    if wall is None:
        return None
    try:
        return element_id_to_int(wall.Id)
    except Exception:
        return None


def normalize_muro_id_key(key):
    """
    Clave estable ``int`` (Python) para dicts ``por_muro_id``.

    Evita fallos de lookup en Revit 2025+ (``Value`` / ``System.Int64`` vs
    ``IntegerValue`` / ``System.Int32`` en IronPython o pythonnet).
    """
    if key is None:
        return None
    try:
        if isinstance(key, bool):
            return int(key)
        if isinstance(key, int):
            return int(key)
    except Exception:
        pass
    v = element_id_to_int(key)
    if v is not None:
        return int(v)
    try:
        return int(key)
    except Exception:
        return None


def normalize_muro_id_dict(mapping):
    """Reindexa un dict ``{wall_id: …}`` con ``normalize_muro_id_key``."""
    if not mapping:
        return mapping
    out = {}
    for k, v in mapping.items():
        nk = normalize_muro_id_key(k)
        if nk is not None:
            out[nk] = v
    return out


def element_id_from_int(value):
    """``int`` → ``ElementId`` (Int64 en Revit 2024+)."""
    if ElementId is None:
        return None
    if value is None:
        return ElementId.InvalidElementId
    try:
        return ElementId(int(value))
    except Exception:
        pass
    return ElementId.InvalidElementId
