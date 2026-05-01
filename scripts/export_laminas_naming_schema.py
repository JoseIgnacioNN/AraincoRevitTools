# -*- coding: utf-8 -*-
"""
Persistencia de la receta «Nombre Personalizado» (exportar láminas) en el documento Revit.

Los datos se guardan en :class:`ProjectInfo` mediante Extensible Storage, de modo que
viajan con el modelo (worksharing local, ACC / nube, etc.) tras sincronizar.

Formato JSON interno (campo ``payload``): ``{"v":1,"segments":[{Key,Prefix,Suffix,Separator},...]}``
"""

from __future__ import print_function

import json

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.DB.ExtensibleStorage import Entity, Schema, SchemaBuilder

try:
    from Autodesk.Revit.DB.ExtensibleStorage import AccessLevel
except Exception:
    AccessLevel = None

from System import Guid, String as ClrString

_SCHEMA_GUID = Guid("c4d3e2f1-a5b6-4789-a012-fedcba987654")
_SCHEMA_NAME = "BIMToolsExportLaminasNaming"
_VENDOR_ID = "BIMT"
_FIELD_PAYLOAD = "payload"
_JSON_VERSION = 1

_schema_cached = None


def _ensure_schema():
    global _schema_cached
    if _schema_cached is not None:
        return _schema_cached
    sch = Schema.Lookup(_SCHEMA_GUID)
    if sch is not None:
        _schema_cached = sch
        return sch
    sb = SchemaBuilder(_SCHEMA_GUID)
    sb.SetSchemaName(_SCHEMA_NAME)
    try:
        sb.SetVendorId(_VENDOR_ID)
    except Exception:
        pass
    if AccessLevel is not None:
        try:
            sb.SetReadAccessLevel(AccessLevel.Public)
            sb.SetWriteAccessLevel(AccessLevel.Public)
        except Exception:
            pass
    try:
        sb.AddSimpleField(_FIELD_PAYLOAD, ClrString)
    except Exception:
        from System import Type

        sb.AddSimpleField(_FIELD_PAYLOAD, Type.GetType("System.String"))
    sch = sb.Finish()
    _schema_cached = sch
    return sch


def _project_info(document):
    if document is None:
        return None
    try:
        return document.ProjectInformation
    except Exception:
        return None


def load_recipe_payload(document):
    """
    Lee el JSON crudo del proyecto. Devuelve ``None`` si no hay dato o es inválido.
    """
    pinfo = _project_info(document)
    if pinfo is None:
        return None
    try:
        sch = _ensure_schema()
    except Exception:
        return None
    try:
        ent = pinfo.GetEntity(sch)
        if ent is None or not ent.IsValid():
            return None
    except Exception:
        return None
    try:
        fld = sch.GetField(_FIELD_PAYLOAD)
    except Exception:
        fld = None
    raw = None
    if fld is not None:
        try:
            raw = ent.Get[str](fld)
        except Exception:
            try:
                raw = ent.Get[ClrString](fld)
            except Exception:
                pass
    if raw is None:
        try:
            raw = ent.Get[str](_FIELD_PAYLOAD)
        except Exception:
            try:
                raw = ent.Get[ClrString](_FIELD_PAYLOAD)
            except Exception:
                return None
    if not raw:
        return None
    try:
        u = unicode(raw).strip()
    except Exception:
        u = u"" if raw is None else str(raw).decode("utf-8", "replace")
    if not u:
        return None
    try:
        return json.loads(u)
    except Exception:
        return None


def _segments_from_payload(obj):
    if not isinstance(obj, dict):
        return None
    try:
        segs = obj.get("segments")
    except Exception:
        return None
    if segs is None:
        try:
            segs = obj.get(u"segments")
        except Exception:
            segs = None
    if not isinstance(segs, list):
        return None
    out = []
    for item in segs:
        if not isinstance(item, dict):
            continue
        try:
            k = item.get(u"Key", item.get("Key", u""))
            k = unicode(k) if k is not None else u""
        except Exception:
            k = u""
        if not k:
            continue
        try:
            pre = unicode(item.get(u"Prefix", item.get("Prefix", u"")) or u"")
            suf = unicode(item.get(u"Suffix", item.get("Suffix", u"")) or u"")
            sep = unicode(item.get(u"Separator", item.get("Separator", u"")) or u"")
        except Exception:
            pre = suf = sep = u""
        out.append(
            {u"Key": k, u"Prefix": pre, u"Suffix": suf, u"Separator": sep}
        )
    return out


def load_recipe_segments(document):
    """
    Devuelve lista de dicts {Key, Prefix, Suffix, Separator} o lista vacía.
    """
    obj = load_recipe_payload(document)
    if obj is None:
        return []
    segs = _segments_from_payload(obj)
    return segs if segs is not None else []


def save_recipe_segments(document, segments_dict_list):
    """
    Persiste la receta en ``ProjectInformation``. ``segments_dict_list`` es la lista
    de segmentos (mismas claves que usa ``evaluate_naming_recipe``).

    :returns: True si el documento se actualizó (transacción confirmada).
    """
    pinfo = _project_info(document)
    if pinfo is None:
        return False
    try:
        sch = _ensure_schema()
    except Exception:
        return False
    try:
        payload_obj = {
            "v": _JSON_VERSION,
            "segments": [
                {
                    u"Key": unicode(s.get(u"Key", u"")),
                    u"Prefix": unicode(s.get(u"Prefix") or u""),
                    u"Suffix": unicode(s.get(u"Suffix") or u""),
                    u"Separator": unicode(s.get(u"Separator") or u""),
                }
                for s in segments_dict_list or []
            ],
        }
    except Exception:
        return False
    try:
        json_text = json.dumps(payload_obj, ensure_ascii=False, separators=(",", ":"))
        if not isinstance(json_text, unicode):
            try:
                json_text = unicode(json_text, "utf-8", "replace")
            except Exception:
                json_text = unicode(str(json_text))
    except Exception:
        return False

    t = Transaction(document, u"BIMTools: receta nombres exportación")
    try:
        t.Start()
    except Exception:
        return False
    ok = False
    try:
        try:
            old = pinfo.GetEntity(sch)
            if old is not None and old.IsValid():
                pinfo.DeleteEntity(sch)
        except Exception:
            try:
                pinfo.DeleteEntity(sch)
            except Exception:
                pass
        ent = Entity(sch)
        try:
            fld = sch.GetField(_FIELD_PAYLOAD)
            ent.Set[str](fld, json_text)
        except Exception:
            try:
                ent.Set[str](_FIELD_PAYLOAD, json_text)
            except Exception:
                try:
                    ent.Set(_FIELD_PAYLOAD, json_text)
                except Exception:
                    try:
                        ent.Set(u"payload", json_text)
                    except Exception:
                        raise
        pinfo.SetEntity(ent)
        t.Commit()
        ok = True
    except Exception:
        try:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
        except Exception:
            pass
        ok = False
    return ok
