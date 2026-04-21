# -*- coding: utf-8 -*-
"""
Vínculo Detail de empalme — **zapata de muro** (``enfierrado_wall_foundation``).

Mismos campos que :mod:`lap_detail_link_schema` (ra, rb, dim) pero **GUID distinto** para que el
DMU **no** recoloque el tramo del símbolo al cambiar el armado: la posición queda la de la
creación (eje de stock entre tramos), coherente con conjuntos ``MaximumSpacing`` donde el
centro predominante puede ser la primera barra del layout.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import ElementId
from Autodesk.Revit.DB.ExtensibleStorage import Entity, Schema, SchemaBuilder

try:
    from Autodesk.Revit.DB.ExtensibleStorage import AccessLevel
except Exception:
    AccessLevel = None

from System import Guid, Int32

from embed_anchorage_link_schema import element_id_to_int

LAP_DETAIL_WALL_FOUNDATION_LINK_SCHEMA_GUID = Guid(
    "c3d4e5f6-a7b8-4902-d345-6789abcdef01"
)
_SCHEMA_NAME = "BIMToolsLapDetailLinkWallFoundation"
_VENDOR_ID = "BIMT"

_schema_cached = None


def _ensure_schema():
    global _schema_cached
    if _schema_cached is not None:
        return _schema_cached
    sch = Schema.Lookup(LAP_DETAIL_WALL_FOUNDATION_LINK_SCHEMA_GUID)
    if sch is not None:
        _schema_cached = sch
        return sch
    sb = SchemaBuilder(LAP_DETAIL_WALL_FOUNDATION_LINK_SCHEMA_GUID)
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
        sb.AddSimpleField("ra", Int32)
        sb.AddSimpleField("rb", Int32)
        sb.AddSimpleField("dim", Int32)
    except Exception:
        from System import Type

        t_int = Type.GetType("System.Int32")
        sb.AddSimpleField("ra", t_int)
        sb.AddSimpleField("rb", t_int)
        sb.AddSimpleField("dim", t_int)
    sch = sb.Finish()
    _schema_cached = sch
    return sch


def set_wall_foundation_lap_detail_rebar_link(
    detail_inst, rebar_tail_id, rebar_head_id, dimension_id=None
):
    if detail_inst is None:
        return False
    try:
        sch = _ensure_schema()
    except Exception:
        return False
    try:
        a = int(rebar_tail_id.Value)
    except Exception:
        try:
            a = int(rebar_tail_id.IntegerValue)
        except Exception:
            return False
    try:
        b = int(rebar_head_id.Value)
    except Exception:
        try:
            b = int(rebar_head_id.IntegerValue)
        except Exception:
            return False
    d = 0
    if dimension_id is not None and dimension_id != ElementId.InvalidElementId:
        try:
            d = int(dimension_id.Value)
        except Exception:
            try:
                d = int(dimension_id.IntegerValue)
            except Exception:
                d = 0
    try:
        ent = Entity(sch)
        try:
            ent.Set[int]("ra", a)
            ent.Set[int]("rb", b)
            ent.Set[int]("dim", d)
        except Exception:
            ent.Set("ra", a)
            ent.Set("rb", b)
            ent.Set("dim", d)
        detail_inst.SetEntity(ent)
        return True
    except Exception:
        return False


def get_wall_foundation_lap_detail_rebar_link(detail_inst):
    """
    Retorna dict { 'ra': ElementId, 'rb': ElementId, 'dim': ElementId|None } o None.
    """
    if detail_inst is None:
        return None
    try:
        sch = _ensure_schema()
    except Exception:
        return None
    try:
        ent = detail_inst.GetEntity(sch)
        if ent is None or not ent.IsValid():
            return None
        a = int(ent.Get[int]("ra"))
        b = int(ent.Get[int]("rb"))
        d = int(ent.Get[int]("dim"))
    except Exception:
        try:
            a = int(ent.Get[Int32]("ra"))
            b = int(ent.Get[Int32]("rb"))
            d = int(ent.Get[Int32]("dim"))
        except Exception:
            return None
    return {
        "ra": ElementId(a),
        "rb": ElementId(b),
        "dim": ElementId(d) if d else None,
    }


def iter_wall_foundation_lap_linked_detail_instances(document):
    if document is None:
        return
    try:
        sch = _ensure_schema()
    except Exception:
        return
    try:
        from Autodesk.Revit.DB import FamilyInstance, FilteredElementCollector
        from Autodesk.Revit.DB import BuiltInCategory

        bic = BuiltInCategory.OST_DetailComponents
        col = (
            FilteredElementCollector(document)
            .OfCategory(bic)
            .OfClass(FamilyInstance)
        )
        for el in col:
            try:
                link = get_wall_foundation_lap_detail_rebar_link(el)
            except Exception:
                link = None
            if link is not None:
                yield el, link
    except Exception:
        return


def find_wall_foundation_lap_details_touching_rebar_ids(document, id_ints):
    if not id_ints:
        return []
    want = set(int(x) for x in id_ints)
    out = []
    for inst, link in iter_wall_foundation_lap_linked_detail_instances(document):
        ia = element_id_to_int(link["ra"])
        ib = element_id_to_int(link["rb"])
        if (ia is not None and ia in want) or (ib is not None and ib in want):
            out.append((inst, link))
    return out
