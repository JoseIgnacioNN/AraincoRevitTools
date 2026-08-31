# -*- coding: utf-8 -*-
"""
Cota de empotramiento: el vínculo vive en la DetailCurve del **marcador** (línea corta).

``endi`` indica qué extremo de la barra (0 inicio / 1 fin) define la posición del marcador;
al cambiar dimensiones del armado, el DMU elimina marcador y cota y los vuelve a crear; el mm
de empotramiento en el esquema se actualiza con la distancia medida extremo–cara en el modelo.
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

from System import Guid, Int32, String as CString, Type

EMBED_ANCHORAGE_SCHEMA_GUID = Guid("c3d4e5f6-a7b8-4901-c234-56789abcdef0")
_SCHEMA_NAME = "BIMToolsEmbedAnchorageDim"
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
    sch = Schema.Lookup(EMBED_ANCHORAGE_SCHEMA_GUID)
    if sch is not None:
        _schema_cached = sch
        return sch
    sb = SchemaBuilder(EMBED_ANCHORAGE_SCHEMA_GUID)
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
    t_str = Type.GetType("System.String")
    sb.AddSimpleField("rb", Int32)
    sb.AddSimpleField("dim", Int32)
    sb.AddSimpleField("vw", Int32)
    sb.AddSimpleField("amm", Int32)
    sb.AddSimpleField("endi", Int32)
    sb.AddSimpleField("fst", t_str)
    sch = sb.Finish()
    _schema_cached = sch
    return sch


def set_embed_anchorage_link(
    marker_detail_curve,
    rebar_id,
    dimension_id,
    view_id,
    face_stable_repr,
    anchorage_mm,
    end_index,
):
    if marker_detail_curve is None:
        return False
    try:
        sch = _ensure_schema()
    except Exception:
        return False
    try:
        rb = int(rebar_id.Value)
    except Exception:
        try:
            rb = int(rebar_id.IntegerValue)
        except Exception:
            return False
    try:
        dm = int(dimension_id.Value)
    except Exception:
        try:
            dm = int(dimension_id.IntegerValue)
        except Exception:
            return False
    try:
        vw = int(view_id.Value)
    except Exception:
        try:
            vw = int(view_id.IntegerValue)
        except Exception:
            return False
    try:
        amm = int(round(float(anchorage_mm)))
    except Exception:
        amm = 0
    try:
        endi = int(end_index)
    except Exception:
        endi = 0
    if endi not in (0, 1):
        endi = 0
    fst = face_stable_repr or u""
    try:
        fst = unicode(fst)
    except Exception:
        fst = str(fst)
    try:
        ent = Entity(sch)
        try:
            ent.Set[int]("rb", rb)
            ent.Set[int]("dim", dm)
            ent.Set[int]("vw", vw)
            ent.Set[int]("amm", amm)
            ent.Set[int]("endi", endi)
            ent.Set[str]("fst", fst)
        except Exception:
            ent.Set("rb", rb)
            ent.Set("dim", dm)
            ent.Set("vw", vw)
            ent.Set("amm", amm)
            ent.Set("endi", endi)
            ent.Set("fst", fst)
        marker_detail_curve.SetEntity(ent)
        return True
    except Exception:
        return False


def get_embed_anchorage_link(marker_detail_curve):
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
        rb = int(ent.Get[int]("rb"))
        dm = int(ent.Get[int]("dim"))
        vw = int(ent.Get[int]("vw"))
        amm = int(ent.Get[int]("amm"))
        endi = int(ent.Get[int]("endi"))
        fst = ent.Get[str]("fst")
    except Exception:
        try:
            if ent is None or not ent.IsValid():
                return None
            rb = int(ent.Get[Int32]("rb"))
            dm = int(ent.Get[Int32]("dim"))
            vw = int(ent.Get[Int32]("vw"))
            amm = int(ent.Get[Int32]("amm"))
            endi = int(ent.Get[Int32]("endi"))
            fst = ent.Get[CString]("fst")
        except Exception:
            return None
    if fst is None:
        fst = u""
    try:
        fst = unicode(fst)
    except Exception:
        fst = str(fst)
    return {
        "rb": ElementId(rb),
        "dim": ElementId(dm),
        "vw": ElementId(vw),
        "amm": amm,
        "endi": endi,
        "fst": fst,
    }


def iter_embed_anchorage_markers(document):
    if document is None:
        return
    try:
        _ensure_schema()
    except Exception:
        return
    try:
        # DetailLine hereda DetailCurve; OfClass(DetailCurve) a veces no recoge todo según versión.
        col = FilteredElementCollector(document).OfClass(CurveElement)
        for el in col:
            if not isinstance(el, DetailCurve):
                continue
            try:
                link = get_embed_anchorage_link(el)
            except Exception:
                link = None
            if link is not None:
                yield el, link
    except Exception:
        return


def find_embed_anchorage_touching_rebar_ids(document, id_ints):
    if not id_ints:
        return []
    want = set(int(x) for x in id_ints)
    out = []
    for el, link in iter_embed_anchorage_markers(document):
        rid = element_id_to_int(link["rb"])
        if rid is None:
            continue
        if rid in want:
            out.append((el, link))
    return out


def update_embed_anchorage_dim_id(marker_detail_curve, new_dimension_id):
    if marker_detail_curve is None:
        return False
    link = get_embed_anchorage_link(marker_detail_curve)
    if link is None:
        return False
    try:
        nd = int(new_dimension_id.Value)
    except Exception:
        try:
            nd = int(new_dimension_id.IntegerValue)
        except Exception:
            return False
    return set_embed_anchorage_link(
        marker_detail_curve,
        link["rb"],
        ElementId(nd),
        link["vw"],
        link["fst"],
        link["amm"],
        link["endi"],
    )
