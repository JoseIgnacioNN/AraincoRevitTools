# -*- coding: utf-8 -*-
"""
Marcadores DetailCurve de cotas de confinamiento (columnas).

Cada línea de detalle almacena el ``ElementId`` de la cota padre. El DMU
``confinement_dim_updater_dmu`` elimina los marcadores cuando se borra la cota.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import CurveElement, DetailCurve, ElementId, FilteredElementCollector
from Autodesk.Revit.DB.ExtensibleStorage import Entity, Schema, SchemaBuilder

try:
    from Autodesk.Revit.DB.ExtensibleStorage import AccessLevel
except Exception:
    AccessLevel = None

from System import Guid, Int32

CONFINEMENT_DIM_LINK_SCHEMA_GUID = Guid("d4e5f6a7-b8c9-4012-d345-6789abcdef01")
_SCHEMA_NAME = "BIMToolsConfinementDimMarker"
_VENDOR_ID = "BIMT"

_schema_cached = None


def element_id_to_int(eid):
    """Revit 2026+: ``ElementId.Value``; versiones anteriores: ``IntegerValue``."""
    if eid is None:
        return None
    try:
        return int(eid.Value)
    except Exception:
        try:
            return int(eid.IntegerValue)
        except Exception:
            return None


def _ensure_schema():
    global _schema_cached
    if _schema_cached is not None:
        return _schema_cached
    sch = Schema.Lookup(CONFINEMENT_DIM_LINK_SCHEMA_GUID)
    if sch is not None:
        _schema_cached = sch
        return sch
    sb = SchemaBuilder(CONFINEMENT_DIM_LINK_SCHEMA_GUID)
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
    sb.AddSimpleField("dim", Int32)
    sb.AddSimpleField("vw", Int32)
    sch = sb.Finish()
    _schema_cached = sch
    return sch


def set_confinement_dim_marker_link(marker_detail_curve, dimension_id, view_id):
    """Registra vínculo marcador → cota de confinamiento."""
    if marker_detail_curve is None:
        return False
    dm = element_id_to_int(dimension_id)
    vw = element_id_to_int(view_id)
    if dm is None or vw is None:
        return False
    try:
        sch = _ensure_schema()
    except Exception:
        return False
    try:
        ent = Entity(sch)
        try:
            ent.Set[int]("dim", dm)
            ent.Set[int]("vw", vw)
        except Exception:
            ent.Set("dim", dm)
            ent.Set("vw", vw)
        marker_detail_curve.SetEntity(ent)
        return True
    except Exception:
        return False


def get_confinement_dim_marker_link(marker_detail_curve):
    if marker_detail_curve is None:
        return None
    try:
        sch = _ensure_schema()
    except Exception:
        return None
    ent = None
    try:
        ent = marker_detail_curve.GetEntity(sch)
        if ent is None or not ent.IsValid():
            return None
        dm = int(ent.Get[int]("dim"))
        vw = int(ent.Get[int]("vw"))
    except Exception:
        try:
            if ent is None or not ent.IsValid():
                return None
            dm = int(ent.Get[Int32]("dim"))
            vw = int(ent.Get[Int32]("vw"))
        except Exception:
            return None
    return {
        "dim": ElementId(dm),
        "vw": ElementId(vw),
    }


def iter_confinement_dim_markers(document):
    """Itera DetailCurve con schema de marcador de cota de confinamiento."""
    if document is None:
        return
    try:
        _ensure_schema()
    except Exception:
        return
    try:
        col = FilteredElementCollector(document).OfClass(CurveElement)
        for el in col:
            if not isinstance(el, DetailCurve):
                continue
            try:
                link = get_confinement_dim_marker_link(el)
            except Exception:
                link = None
            if link is not None:
                yield el, link
    except Exception:
        return


def find_confinement_dim_markers_for_dim_ids(document, dim_id_ints):
    """
    ``dim_id_ints``: ids enteros de cotas borradas.
    Retorna lista de DetailCurve cuyo vínculo apunta a alguna de esas cotas.
    """
    if not dim_id_ints:
        return []
    want = set(int(x) for x in dim_id_ints)
    out = []
    seen = set()
    for el, link in iter_confinement_dim_markers(document):
        did = element_id_to_int(link.get("dim"))
        if did is None or did not in want:
            continue
        eid = element_id_to_int(el.Id)
        if eid is None or eid in seen:
            continue
        seen.add(eid)
        out.append(el)
    return out
