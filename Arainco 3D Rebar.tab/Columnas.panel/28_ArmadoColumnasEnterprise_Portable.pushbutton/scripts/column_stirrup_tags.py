# -*- coding: utf-8 -*-
u"""
Etiquetas de confinamiento (estribos/trabas) en Armado Columnas.

Familias (mismo criterio que **Armado vigas**); orientación **vertical** (columnas).

- ``EST_A_STRUCTURAL REBAR TAG_CONFINAMIENTO_VIGA`` — estribo o traba individual.
- ``EST_A_STRUCTURAL REBAR TAG_CONFINAMIENTO_VIGA_MULTI HOST`` — varios estribos
  interiores o varias trabas del mismo tramo.

Posición: a la **derecha** del pilar, **500 mm** de offset fijo respecto al borde
visible **más lejano** del análisis (columna y estribos/trabas). Cotas y etiquetas
comparten la misma X vertical. Por cada lote (tramo continuo o tercio **L/3**) se crea
una **cota alineada vertical** (tipo «Linear - Confinamiento»); estribos y trabas
comparten el **mismo punto de inserción** de cabecera de etiqueta (sin escalonar).
"""

from __future__ import print_function

import os
import sys

import clr

clr.AddReference("RevitAPI")

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    DimensionStyleType,
    DimensionType,
    ElementId,
    ElementType,
    FilteredElementCollector,
    IndependentTag,
    Line,
    Options,
    Reference,
    ReferenceArray,
    StorageType,
    SubTransaction,
    TagMode,
    TagOrientation,
    ViewDrafting,
    ViewSection,
    ViewType,
    XYZ,
)
from Autodesk.Revit.DB.Structure import Rebar

CONFINEMENT_TAG_FAMILY = u"EST_A_STRUCTURAL REBAR TAG_CONFINAMIENTO_VIGA"
CONFINEMENT_MULTIHOST_TAG_FAMILY = (
    u"EST_A_STRUCTURAL REBAR TAG_CONFINAMIENTO_VIGA_MULTI HOST"
)

STIRRUP_TAG_OFFSET_FROM_COLUMN_RIGHT_MM = 500.0
CONFINEMENT_DIM_MARKER_LENGTH_MM = 80.0
_FIXED_DIMSTYLE_NAME = u"Linear - Confinamiento"
_FIXED_DIMSTYLE_CACHE = {}

try:
    from enfierrado_shaft_hashtag import (
        _collect_rebar_tag_symbol_map,
        _primary_rebar_shape_tag_key,
        _rebar_reference_candidates_for_tag,
        _rebar_shape_name_candidates,
    )
except Exception:
    _collect_rebar_tag_symbol_map = None
    _primary_rebar_shape_tag_key = None
    _rebar_reference_candidates_for_tag = None
    _rebar_shape_name_candidates = None

try:
    from confinement_dim_link_schema import set_confinement_dim_marker_link
except Exception:
    set_confinement_dim_marker_link = None

try:
    from confinement_dim_updater_dmu import ensure_confinement_dim_link_updater_registered
except Exception:
    ensure_confinement_dim_link_updater_registered = None


def _load_viga_etiquetar_module():
    u"""``armado_vigas.revit.etiquetar_confinamiento`` (pushbutton o ``scripts/``)."""
    mod_name = u"armado_vigas.revit.etiquetar_confinamiento"
    if mod_name in sys.modules:
        return sys.modules[mod_name]
    try:
        from armado_vigas.revit import etiquetar_confinamiento as mod
        return mod
    except Exception:
        pass
    here = os.path.dirname(os.path.abspath(__file__))
    candidates = []
    try:
        candidates.append(os.path.dirname(here))
    except Exception:
        pass
    cursor = here
    for _ in range(12):
        candidates.append(
            os.path.join(
                cursor,
                u"BIMTools.tab",
                u"Armadura.panel",
                u"36_ArmadoVigas.pushbutton",
                u"scripts",
            )
        )
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    seen = set()
    for scripts_root in candidates:
        if not scripts_root:
            continue
        try:
            nd = os.path.normpath(scripts_root)
        except Exception:
            nd = scripts_root
        if nd in seen:
            continue
        seen.add(nd)
        mod_path = os.path.join(
            nd, u"armado_vigas", u"revit", u"etiquetar_confinamiento.py",
        )
        if not os.path.isfile(mod_path):
            continue
        if nd not in sys.path:
            sys.path.insert(0, nd)
        try:
            from armado_vigas.revit import etiquetar_confinamiento as mod
            return mod
        except Exception:
            continue
    return None


_VIGA_ETQ = _load_viga_etiquetar_module()
_TAGS_MOD = None
if _VIGA_ETQ is not None:
    try:
        _TAGS_MOD = _VIGA_ETQ._load_cabezal_tags_module()
    except Exception:
        _TAGS_MOD = None


def _mm_to_ft(mm):
    return float(mm) / 304.8


def _is_section_or_elevation_view(view):
    if view is None:
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass
    try:
        if isinstance(view, ViewDrafting):
            return False
    except Exception:
        pass
    try:
        if isinstance(view, ViewSection):
            return True
    except Exception:
        pass
    try:
        vt = view.ViewType
        return vt == ViewType.Section or vt == ViewType.Elevation
    except Exception:
        return False


def _dot_xyz(a, b):
    return (
        float(a.X) * float(b.X)
        + float(a.Y) * float(b.Y)
        + float(a.Z) * float(b.Z)
    )


def _view_local_xy(view, pt):
    o = view.Origin
    d = XYZ(
        float(pt.X) - float(o.X),
        float(pt.Y) - float(o.Y),
        float(pt.Z) - float(o.Z),
    )
    return _dot_xyz(d, view.RightDirection), _dot_xyz(d, view.UpDirection)


def _view_local_to_world(view, ref_world, x_loc, y_loc):
    x_ref, y_ref = _view_local_xy(view, ref_world)
    rd = view.RightDirection
    up = view.UpDirection
    return ref_world + rd.Multiply(float(x_loc - x_ref)) + up.Multiply(
        float(y_loc - y_ref)
    )


def _bbox_corners_world(bb):
    if bb is None:
        return []
    try:
        tr = bb.Transform
        mn, mx = bb.Min, bb.Max
        out = []
        for x in (float(mn.X), float(mx.X)):
            for y in (float(mn.Y), float(mx.Y)):
                for z in (float(mn.Z), float(mx.Z)):
                    out.append(tr.OfPoint(XYZ(x, y, z)))
        return out
    except Exception:
        try:
            return [bb.Min, bb.Max]
        except Exception:
            return []


def _bbox_extents_view_local_x(view, bb):
    min_x = None
    max_x = None
    for pt in _bbox_corners_world(bb):
        x_loc, _ = _view_local_xy(view, pt)
        if min_x is None or x_loc < min_x:
            min_x = x_loc
        if max_x is None or x_loc > max_x:
            max_x = x_loc
    return min_x, max_x


def _element_bounding_box(elem, view):
    if elem is None:
        return None
    try:
        bb = elem.get_BoundingBox(view)
        if bb is not None:
            return bb
    except Exception:
        pass
    try:
        return elem.get_BoundingBox(None)
    except Exception:
        return None


def _max_view_local_right(view, elements):
    u"""Mayor coord. local X (borde derecho en vista) entre varios elementos."""
    max_r = None
    for el in elements or []:
        _, max_x = _bbox_extents_view_local_x(view, _element_bounding_box(el, view))
        if max_x is None:
            continue
        if max_r is None or max_x > max_r:
            max_r = max_x
    return float(max_r) if max_r is not None else 0.0


def _confinement_annotation_x_dim(view, column, rebars=None):
    u"""
    X común de cotas y etiquetas: borde derecho más lejano + offset.

    Incluye la columna y todos los estribos/trabas del confinamiento para que
    pilares con cambio de sección queden alineados en una sola vertical.
    """
    elements = []
    if column is not None:
        elements.append(column)
    for rb in rebars or []:
        if rb is not None:
            elements.append(rb)
    x_right = _max_view_local_right(view, elements)
    return x_right + _mm_to_ft(STIRRUP_TAG_OFFSET_FROM_COLUMN_RIGHT_MM)


def _column_min_view_local_left(view, col):
    min_x, _ = _bbox_extents_view_local_x(view, _element_bounding_box(col, view))
    return float(min_x) if min_x is not None else 0.0


def _column_reference_xyz(col, z_ft):
    bb = None
    try:
        bb = col.get_BoundingBox(None)
    except Exception:
        pass
    if bb is None:
        return XYZ(0.0, 0.0, float(z_ft))
    try:
        return XYZ(
            0.5 * (float(bb.Min.X) + float(bb.Max.X)),
            0.5 * (float(bb.Min.Y) + float(bb.Max.Y)),
            float(z_ft),
        )
    except Exception:
        return XYZ(0.0, 0.0, float(z_ft))


def _rebar_center(rebar, view):
    try:
        bb = rebar.get_BoundingBox(view)
        if bb is not None:
            return (bb.Min + bb.Max) * 0.5
    except Exception:
        pass
    try:
        bb = rebar.get_BoundingBox(None)
        if bb is not None:
            return (bb.Min + bb.Max) * 0.5
    except Exception:
        pass
    return None


def _mean_view_local_y(view, ref_world, rebars):
    ys = []
    for rb in rebars:
        if not isinstance(rb, Rebar):
            continue
        p = _rebar_center(rb, view)
        if p is None:
            continue
        _, y_mid = _view_local_xy(view, p)
        if y_mid is not None:
            ys.append(float(y_mid))
    if not ys:
        return None
    return sum(ys) / float(len(ys))


def _tag_type_id_from_map_value(val):
    u"""``FamilySymbol`` (muros) o ``ElementId`` (enfierrado_shaft_hashtag) → ``ElementId``."""
    if val is None:
        return None
    try:
        if isinstance(val, ElementId):
            return val
    except Exception:
        pass
    try:
        tid = val.Id
        if tid is not None:
            return tid
    except Exception:
        pass
    return None


def _rebar_element_id(rebar):
    u"""``Rebar`` o ``ElementId`` → ``ElementId``."""
    if rebar is None:
        return None
    try:
        if isinstance(rebar, ElementId):
            return rebar
    except Exception:
        pass
    try:
        return rebar.Id
    except Exception:
        return None


def _collect_confinement_tag_maps(document):
    u"""Mapas tipo → símbolo/id (vigas; respaldo enfierrado)."""
    if _VIGA_ETQ is not None:
        try:
            tag_map = _VIGA_ETQ._collect_tag_symbol_map(
                document, CONFINEMENT_TAG_FAMILY, _TAGS_MOD,
            )
            mh_map = _VIGA_ETQ._collect_tag_symbol_map(
                document, CONFINEMENT_MULTIHOST_TAG_FAMILY, _TAGS_MOD,
            )
            return tag_map or {}, mh_map or {}
        except Exception:
            pass
    if _collect_rebar_tag_symbol_map is None:
        return {}, {}
    return (
        _collect_rebar_tag_symbol_map(document, CONFINEMENT_TAG_FAMILY) or {},
        _collect_rebar_tag_symbol_map(
            document, CONFINEMENT_MULTIHOST_TAG_FAMILY,
        ) or {},
    )


def _resolve_tag_symbol(document, rebar, tag_map, family_name=None):
    fam = family_name or CONFINEMENT_TAG_FAMILY
    if _TAGS_MOD is not None:
        try:
            sym, lbl = _TAGS_MOD._resolve_tag_symbol_for_rebar(
                document, fam, tag_map, rebar,
            )
            if sym is not None:
                return sym, lbl
        except Exception:
            pass
    if _VIGA_ETQ is not None:
        try:
            sym, lbl = _VIGA_ETQ._resolve_tag_symbol_viga(
                document, rebar, tag_map, _TAGS_MOD,
            )
            if sym is not None:
                return sym, lbl
        except Exception:
            pass
    if not tag_map or rebar is None:
        return None, None
    keys = []
    if _primary_rebar_shape_tag_key is not None:
        pk = _primary_rebar_shape_tag_key(document, rebar)
        if pk:
            keys.append(pk)
    if _rebar_shape_name_candidates is not None:
        keys.extend(_rebar_shape_name_candidates(document, rebar))
    seen = set()
    for k in keys:
        if not k or k in seen:
            continue
        seen.add(k)
        sym = tag_map.get(k)
        if sym is not None:
            return sym, k
    return None, None


def _aplicar_estilo_tag_columna(tag, head):
    u"""Columnas: vertical y sin leader (familias vigas; estilo distinto a vigas)."""
    if tag is None or head is None:
        return
    try:
        tag.TagOrientation = TagOrientation.Vertical
    except Exception:
        pass
    try:
        tag.HasLeader = False
    except Exception:
        pass
    try:
        tag.TagHeadPosition = head
    except Exception:
        pass


def _referencias_tag_rebar(document, rebar, view):
    u"""Referencias para ``IndependentTag`` (vigas / muros / enfierrado)."""
    if _VIGA_ETQ is not None:
        try:
            refs = _VIGA_ETQ._referencias_tag_rebar_viga(
                document, rebar, view, _TAGS_MOD,
            )
            if refs:
                return refs
        except Exception:
            pass
    refs = []
    seen = set()

    def _add_ref(r):
        if r is None:
            return
        try:
            k = r.ConvertToStableRepresentation(document)
        except Exception:
            try:
                k = unicode(r)
            except Exception:
                k = id(r)
        if k in seen:
            return
        seen.add(k)
        refs.append(r)

    if _rebar_reference_candidates_for_tag is not None:
        try:
            for r in _rebar_reference_candidates_for_tag(document, view, rebar) or []:
                _add_ref(r)
        except Exception:
            pass
    if refs:
        return refs

    try:
        subs = rebar.GetSubelements() if hasattr(rebar, "GetSubelements") else None
    except Exception:
        subs = None
    if subs:
        for sub in subs:
            if sub is None:
                continue
            try:
                if hasattr(sub, "GetReference"):
                    _add_ref(sub.GetReference())
            except Exception:
                continue

    try:
        npos = int(getattr(rebar, "NumberOfBarPositions", 0))
    except Exception:
        npos = 0
    if npos > 0:
        for idx in (0, int(npos / 2), max(0, npos - 1)):
            try:
                if hasattr(rebar, "GetReferenceToBarPosition"):
                    _add_ref(rebar.GetReferenceToBarPosition(idx))
                elif hasattr(rebar, "GetReferenceForBarPosition"):
                    _add_ref(rebar.GetReferenceForBarPosition(idx))
            except Exception:
                continue

    try:
        _add_ref(Reference(rebar))
    except Exception:
        pass

    try:
        opts = Options()
        opts.ComputeReferences = True
        opts.View = view
        opts.IncludeNonVisibleObjects = True
        geom = rebar.get_Geometry(opts)
        if geom is not None:
            for go in geom:
                if go is None:
                    continue
                try:
                    rgo = getattr(go, "Reference", None)
                    if rgo is not None:
                        _add_ref(rgo)
                except Exception:
                    pass
    except Exception:
        pass
    return refs


def _activate_tag_symbol(document, sym):
    if sym is None:
        return
    try:
        if not sym.IsActive:
            sym.Activate()
            try:
                document.Regenerate()
            except Exception:
                pass
    except Exception:
        pass


def _crear_tag_sin_leader(document, view, rebar, tag_type_id, head):
    if head is None:
        return None, u"sin punto de inserción"
    if _VIGA_ETQ is not None:
        try:
            sym_el = document.GetElement(tag_type_id)
            if sym_el is not None:
                _activate_tag_symbol(document, sym_el)
            tag, err = _VIGA_ETQ._crear_tag_confinamiento_sin_leader(
                document, view, rebar, tag_type_id, head, _TAGS_MOD,
            )
            if tag is not None:
                _aplicar_estilo_tag_columna(tag, head)
                return tag, None
            if err:
                return None, err
        except Exception as ex:
            try:
                return None, unicode(ex)
            except Exception:
                return None, str(ex)

    refs = _referencias_tag_rebar(document, rebar, view)
    if not refs:
        return None, u"sin referencia API"
    orient = TagOrientation.Vertical
    try:
        sym_el = document.GetElement(tag_type_id)
        _activate_tag_symbol(document, sym_el)
    except Exception:
        pass
    last_ex = None
    for ref in refs:
        try:
            tag = IndependentTag.Create(
                document,
                tag_type_id,
                view.Id,
                ref,
                False,
                orient,
                head,
            )
            if tag is not None:
                _aplicar_estilo_tag_columna(tag, head)
                return tag, None
        except Exception as ex:
            last_ex = ex
    for ref in refs:
        try:
            tag = IndependentTag.Create(
                document,
                view.Id,
                ref,
                False,
                TagMode.TM_ADDBY_CATEGORY,
                orient,
                head,
            )
            if tag is not None:
                try:
                    tag.SetTypeId(tag_type_id)
                except Exception:
                    pass
                _aplicar_estilo_tag_columna(tag, head)
                return tag, None
        except Exception as ex:
            last_ex = ex
    if last_ex is not None:
        try:
            return None, unicode(last_ex)
        except Exception:
            return None, str(last_ex)
    return None, u"no se pudo crear IndependentTag"


def _refs_multihost_extra(document, view, rebar_ids_extra):
    if _TAGS_MOD is not None and hasattr(_TAGS_MOD, u"_refs_multihost_para_rebars"):
        try:
            return _TAGS_MOD._refs_multihost_para_rebars(
                document, view, rebar_ids_extra,
            ) or []
        except Exception:
            pass
    if _rebar_reference_candidates_for_tag is None:
        return []
    refs = []
    seen = set()
    for rid in rebar_ids_extra or []:
        rb = document.GetElement(rid) if document is not None else None
        if not isinstance(rb, Rebar):
            continue
        cand = _rebar_reference_candidates_for_tag(document, view, rb) or []
        if not cand:
            continue
        ref0 = cand[0]
        try:
            key = ref0.ConvertToStableRepresentation(document)
        except Exception:
            key = id(ref0)
        if key in seen:
            continue
        seen.add(key)
        refs.append(ref0)
    return refs


def _crear_tag_multihost(document, view, rebar_ids, tag_map, head):
    u"""Una ``IndependentTag`` multihost (``AddReferences`` en SubTransaction, patrón muros)."""
    ids = list(rebar_ids or [])
    if not ids:
        return False, u"grupo multihost vacío"
    primary = document.GetElement(ids[-1])
    if not isinstance(primary, Rebar):
        return False, u"barra principal no encontrada"
    sym, shape_lbl = _resolve_tag_symbol(
        document, primary, tag_map, CONFINEMENT_MULTIHOST_TAG_FAMILY,
    )
    tag_type_id = _tag_type_id_from_map_value(sym)
    if tag_type_id is None:
        return False, u"sin tipo multihost para shape «{0}»".format(shape_lbl or u"?")
    _activate_tag_symbol(document, document.GetElement(tag_type_id))
    tag, err = _crear_tag_sin_leader(document, view, primary, tag_type_id, head)
    if tag is None:
        return False, err or u"multihost no creado"
    if len(ids) < 2:
        return True, None

    add_fn = getattr(tag, u"AddReferences", None)
    if add_fn is None:
        try:
            document.Delete(tag.Id)
        except Exception:
            pass
        return False, (
            u"Revit no expone AddReferences en esta versión; "
            u"use etiqueta MULTI HOST manual."
        )

    extra_refs = _refs_multihost_extra(document, view, ids[:-1])
    if not extra_refs:
        try:
            document.Delete(tag.Id)
        except Exception:
            pass
        return False, u"sin referencias para hosts adicionales multihost"

    refs_add = List[Reference]()
    for ref in extra_refs:
        refs_add.Add(ref)

    try:
        document.Regenerate()
    except Exception:
        pass

    st = SubTransaction(document)
    try:
        st.Start()
    except Exception:
        try:
            document.Delete(tag.Id)
        except Exception:
            pass
        return False, u"no se pudo abrir SubTransaction para multihost"
    try:
        add_fn(refs_add)
        st.Commit()
        _aplicar_estilo_tag_columna(tag, head)
        return True, None
    except Exception as ex_mh:
        try:
            st.RollBack()
        except Exception:
            pass
        try:
            document.Delete(tag.Id)
        except Exception:
            pass
        try:
            return False, u"Multihost confinamiento (AddReferences): {0}".format(
                unicode(ex_mh)
            )
        except Exception:
            return False, u"Multihost confinamiento (AddReferences): {0}".format(
                str(ex_mh)
            )


def _segment_view_local_y_extent(view, rebars):
    u"""Extensión vertical del lote en coordenadas locales de la vista (min, max)."""
    y_min = None
    y_max = None
    for rb in rebars or []:
        if not isinstance(rb, Rebar):
            continue
        bb = None
        try:
            bb = rb.get_BoundingBox(view)
        except Exception:
            pass
        if bb is None:
            try:
                bb = rb.get_BoundingBox(None)
            except Exception:
                bb = None
        for pt in _bbox_corners_world(bb):
            _, y_loc = _view_local_xy(view, pt)
            if y_min is None or y_loc < y_min:
                y_min = float(y_loc)
            if y_max is None or y_loc > y_max:
                y_max = float(y_loc)
    if y_min is not None and y_max is not None and (y_max - y_min) > _mm_to_ft(5.0):
        return y_min, y_max
    for rb in rebars or []:
        if not isinstance(rb, Rebar):
            continue
        y0, y1 = _rebar_array_extent_view_local_y(view, rb)
        if y0 is None or y1 is None:
            continue
        lo = min(float(y0), float(y1))
        hi = max(float(y0), float(y1))
        if y_min is None or lo < y_min:
            y_min = lo
        if y_max is None or hi > y_max:
            y_max = hi
    if y_min is None or y_max is None:
        return None, None
    return y_min, y_max


def _rebar_array_extent_view_local_y(view, rebar):
    u"""Extremos del reparto del set en Y local de vista (respaldo si bbox es plano)."""
    if rebar is None:
        return None, None
    try:
        n_pos = int(rebar.NumberOfBarPositions)
    except Exception:
        n_pos = 1
    if n_pos < 1:
        n_pos = 1
    origins = []
    for idx in (0, max(0, n_pos - 1)):
        try:
            tr = rebar.GetBarPositionTransform(int(idx))
            if tr is not None:
                origins.append(tr.Origin)
        except Exception:
            pass
    try:
        acc = rebar.GetShapeDrivenAccessor()
        alen = float(acc.ArrayLength)
    except Exception:
        alen = 0.0
    if len(origins) >= 2 and alen > 1e-9:
        o0, o1 = origins[0], origins[-1]
        try:
            delta = o1.Subtract(o0)
            ln = float(delta.GetLength())
            if ln > 1e-9:
                axis = delta.Multiply(1.0 / ln)
                o0 = o0.Add(axis.Multiply(-0.5 * alen))
                o1 = o0.Add(axis.Multiply(alen))
                origins = [o0, o1]
        except Exception:
            pass
    elif len(origins) == 1 and alen > 1e-9:
        o0 = origins[0]
        try:
            path = rebar.GetDistributionPath()
            if path is not None and int(path.Count) >= 1:
                p0 = path[0].GetEndPoint(0)
                p1 = path[int(path.Count) - 1].GetEndPoint(1)
                delta = p1.Subtract(p0)
                ln = float(delta.GetLength())
                if ln > 1e-9:
                    axis = delta.Multiply(1.0 / ln)
                    origins = [o0, o0.Add(axis.Multiply(alen))]
        except Exception:
            pass
    if not origins:
        return None, None
    ys = []
    for pt in origins:
        _, y_loc = _view_local_xy(view, pt)
        ys.append(float(y_loc))
    if not ys:
        return None, None
    return min(ys), max(ys)


def _create_horizontal_dim_marker_detailcurve(doc, view, center_world, length_mm):
    u"""DetailCurve corto horizontal (``RightDirection``) con referencia acotable."""
    if doc is None or view is None or center_world is None:
        return None, None
    try:
        rd = view.RightDirection
        if rd is None or float(rd.GetLength()) < 1e-12:
            return None, None
        rd = rd.Normalize()
        half = 0.5 * _mm_to_ft(float(length_mm))
        p0 = center_world.Add(rd.Multiply(-half))
        p1 = center_world.Add(rd.Multiply(half))
        ln = Line.CreateBound(p0, p1)
        dc = doc.Create.NewDetailCurve(view, ln)
        if dc is None:
            return None, None
        try:
            return dc, dc.GeometryCurve.Reference
        except Exception:
            return dc, None
    except Exception:
        return None, None


def _dimension_type_name_candidates(dt):
    u"""Nombres visibles del tipo de cota (UI Revit)."""
    out = []
    if dt is None:
        return out
    for attr in (u"Name",):
        try:
            v = getattr(dt, attr, None)
            if v:
                out.append(unicode(v).strip())
        except Exception:
            pass
    for bip in (
        BuiltInParameter.SYMBOL_NAME_PARAM,
        BuiltInParameter.ALL_MODEL_TYPE_NAME,
    ):
        try:
            p = dt.get_Parameter(bip)
            if p is not None and p.HasValue and p.StorageType == StorageType.String:
                s = unicode(p.AsString() or u"").strip()
                if s:
                    out.append(s)
        except Exception:
            pass
    try:
        fn = getattr(dt, u"FamilyName", None)
        if fn:
            out.append(unicode(fn).strip())
    except Exception:
        pass
    seen = set()
    uniq = []
    for s in out:
        k = s.lower()
        if k and k not in seen:
            seen.add(k)
            uniq.append(s)
    return uniq


def _is_linear_dimension_type(dt):
    if dt is None:
        return False
    try:
        return dt.StyleType == DimensionStyleType.Linear
    except Exception:
        return True


def _collect_linear_dimension_types(doc):
    out = []
    seen = set()
    try:
        for dt in FilteredElementCollector(doc).OfClass(DimensionType):
            if dt is None or not _is_linear_dimension_type(dt):
                continue
            try:
                k = int(dt.Id.IntegerValue)
            except Exception:
                k = id(dt)
            if k in seen:
                continue
            seen.add(k)
            out.append(dt)
    except Exception:
        pass
    if out:
        return out
    try:
        for et in FilteredElementCollector(doc).OfClass(ElementType):
            if not isinstance(et, DimensionType):
                continue
            if not _is_linear_dimension_type(et):
                continue
            try:
                k = int(et.Id.IntegerValue)
            except Exception:
                k = id(et)
            if k in seen:
                continue
            seen.add(k)
            out.append(et)
    except Exception:
        pass
    return out


def _get_fixed_dimension_type_id(doc):
    u"""Tipo de cota «Linear - Confinamiento» si existe; si no, ``None``."""
    if doc is None:
        return None
    try:
        cache_key = id(doc)
    except Exception:
        cache_key = None
    if cache_key is not None and cache_key in _FIXED_DIMSTYLE_CACHE:
        cached = _FIXED_DIMSTYLE_CACHE.get(cache_key)
        if cached is not None:
            return cached
    found_id = None
    target = unicode(_FIXED_DIMSTYLE_NAME).strip().lower()
    for dt in _collect_linear_dimension_types(doc):
        for nm in _dimension_type_name_candidates(dt):
            if nm.strip().lower() == target:
                found_id = dt.Id
                break
        if found_id is not None:
            break
    if found_id is None:
        for dt in _collect_linear_dimension_types(doc):
            for nm in _dimension_type_name_candidates(dt):
                if target in nm.strip().lower():
                    found_id = dt.Id
                    break
            if found_id is not None:
                break
    if cache_key is not None and found_id is not None:
        _FIXED_DIMSTYLE_CACHE[cache_key] = found_id
    return found_id


def _try_apply_fixed_dimension_type(doc, dim):
    u"""Aplica «Linear - Confinamiento»; no delegar en estilos de shaft/vigas."""
    if doc is None or dim is None:
        return
    dim_type_id = _get_fixed_dimension_type_id(doc)
    if dim_type_id is None:
        return
    try:
        dim.ChangeTypeId(dim_type_id)
    except Exception:
        return
    try:
        if int(dim.GetTypeId().IntegerValue) != int(dim_type_id.IntegerValue):
            return
    except Exception:
        pass


def _create_vertical_confinement_dimension(doc, view, ref_world, x_dim, y_min, y_max):
    u"""
    Cota alineada vertical entre ``y_min`` y ``y_max`` en X ``x_dim`` (coords. vista).

    Returns:
        (bool, unicode|None): éxito y mensaje de error.
    """
    if doc is None or view is None or ref_world is None:
        return False, u"sin documento/vista"
    if y_min is None or y_max is None:
        return False, u"sin extensión vertical del lote"
    try:
        y_lo = float(y_min)
        y_hi = float(y_max)
    except Exception:
        return False, u"extensión vertical inválida"
    if abs(y_hi - y_lo) < _mm_to_ft(10.0):
        return False, u"altura de lote demasiado corta para cota"
    if y_lo > y_hi:
        y_lo, y_hi = y_hi, y_lo

    pt_bot = _view_local_to_world(view, ref_world, float(x_dim), y_lo)
    pt_top = _view_local_to_world(view, ref_world, float(x_dim), y_hi)
    dc_bot, ref_bot = _create_horizontal_dim_marker_detailcurve(
        doc, view, pt_bot, CONFINEMENT_DIM_MARKER_LENGTH_MM,
    )
    dc_top, ref_top = _create_horizontal_dim_marker_detailcurve(
        doc, view, pt_top, CONFINEMENT_DIM_MARKER_LENGTH_MM,
    )
    if ref_bot is None or ref_top is None:
        for dc in (dc_bot, dc_top):
            if dc is not None:
                try:
                    doc.Delete(dc.Id)
                except Exception:
                    pass
        return False, u"sin referencias de marcadores para cota de confinamiento"

    try:
        dim_line = Line.CreateBound(pt_bot, pt_top)
        ra = ReferenceArray()
        ra.Append(ref_bot)
        ra.Append(ref_top)
        dim = doc.Create.NewDimension(view, dim_line, ra)
    except Exception as ex:
        dim = None
        err = ex
    else:
        err = None

    if dim is None:
        for dc in (dc_bot, dc_top):
            if dc is not None:
                try:
                    doc.Delete(dc.Id)
                except Exception:
                    pass
        try:
            return False, unicode(err)
        except Exception:
            return False, str(err)

    _try_apply_fixed_dimension_type(doc, dim)
    if ensure_confinement_dim_link_updater_registered is not None:
        try:
            ensure_confinement_dim_link_updater_registered()
        except Exception:
            pass
    if set_confinement_dim_marker_link is not None:
        for dc in (dc_bot, dc_top):
            if dc is not None:
                try:
                    set_confinement_dim_marker_link(dc, dim.Id, view.Id)
                except Exception:
                    pass
    return True, None


def create_confinement_dimensions_for_column(
    doc,
    view,
    column,
    stirrup_rebars_ordered,
    errors=None,
    stirrup_policy=None,
    val_a=None,
    val_b=None,
    sel_a_text=None,
    sel_b_text=None,
):
    u"""
    Una cota alineada vertical por lote de confinamiento (continuo o L/3).

    Returns:
        int: cantidad de cotas creadas.
    """
    if doc is None or view is None or column is None:
        return 0
    rebars = [rb for rb in (stirrup_rebars_ordered or []) if rb is not None]
    if not rebars or not _is_section_or_elevation_view(view):
        return 0

    try:
        ref_world = _column_reference_xyz(column, 0.0)
        loc = column.Location
        if hasattr(loc, u"Curve") and loc.Curve is not None:
            p0 = loc.Curve.Evaluate(0.0, True)
            p1 = loc.Curve.Evaluate(1.0, True)
            z_ref = 0.5 * (float(p0.Z) + float(p1.Z))
            ref_world = _column_reference_xyz(column, z_ref)
    except Exception:
        ref_world = _column_reference_xyz(column, 0.0)

    x_dim = _confinement_annotation_x_dim(view, column, rebars)

    n_per_seg = len(rebars)
    n_segments = 1
    try:
        from column_stirrup_creator import (
            STIRRUP_POLICY_THIRDS_L3,
            build_stirrup_rect_and_tie_defs,
        )
        if build_stirrup_rect_and_tie_defs is not None:
            _rd, _td = build_stirrup_rect_and_tie_defs(
                int(val_a or 4), int(val_b or 4),
                sel_a_text or u"", sel_b_text or u"",
            )
            n_per_seg = len(_rd) + len(_td)
            if n_per_seg >= 1:
                use_l3 = (
                    stirrup_policy == STIRRUP_POLICY_THIRDS_L3
                    and len(rebars) >= 3 * n_per_seg
                )
                n_segments = 3 if use_l3 else 1
    except Exception:
        pass

    n_dims = 0
    for seg in range(n_segments):
        i0 = seg * n_per_seg
        seg_rebars = rebars[i0: i0 + n_per_seg]
        if not seg_rebars:
            break
        y_min, y_max = _segment_view_local_y_extent(view, seg_rebars)
        ok, err = _create_vertical_confinement_dimension(
            doc, view, ref_world, x_dim, y_min, y_max,
        )
        if ok:
            n_dims += 1
        elif errors is not None and err:
            errors.append(
                u"Cota confinamiento lote {0}: {1}".format(seg + 1, err)
            )
    return n_dims


def _build_tag_jobs(rebars_ordered, val_a, val_b, sel_a_text, sel_b_text, stirrup_policy):
    try:
        from column_stirrup_creator import (
            STIRRUP_POLICY_THIRDS_L3,
            build_stirrup_rect_and_tie_defs,
        )
    except Exception:
        STIRRUP_POLICY_THIRDS_L3 = u"thirds_l3"
        build_stirrup_rect_and_tie_defs = None

    if build_stirrup_rect_and_tie_defs is None:
        return [{"mode": u"conf_tag", u"rebars": list(rebars_ordered or [])}]

    rect_defs, tie_defs = build_stirrup_rect_and_tie_defs(
        int(val_a), int(val_b), sel_a_text, sel_b_text,
    )
    n_rect = len(rect_defs)
    n_tie = len(tie_defs)
    n_per_seg = n_rect + n_tie
    if n_per_seg < 1:
        return []

    rebars = [rb for rb in (rebars_ordered or []) if rb is not None]
    use_l3 = (
        stirrup_policy == STIRRUP_POLICY_THIRDS_L3
        and len(rebars) >= 3 * n_per_seg
    )
    n_segments = 3 if use_l3 else 1

    jobs = []
    for seg in range(n_segments):
        base = seg * n_per_seg
        seg_rebars = rebars[base: base + n_per_seg]
        if len(seg_rebars) < n_per_seg:
            break
        seg_idx = seg

        rect_stirrups = list(seg_rebars[0:n_rect])
        if len(rect_stirrups) >= 2:
            jobs.append({
                u"mode": u"conf_tag_multihost",
                u"rebars": rect_stirrups,
                u"seg": seg_idx,
                u"kind": u"stirrup",
            })
        elif len(rect_stirrups) == 1:
            jobs.append({
                u"mode": u"conf_tag",
                u"rebars": rect_stirrups,
                u"seg": seg_idx,
                u"kind": u"stirrup",
            })
        ties = list(seg_rebars[n_rect:n_per_seg])
        if len(ties) >= 2:
            jobs.append({
                u"mode": u"conf_tag_multihost",
                u"rebars": ties,
                u"seg": seg_idx,
                u"kind": u"tie",
            })
        elif len(ties) == 1:
            jobs.append({
                u"mode": u"conf_tag",
                u"rebars": ties,
                u"seg": seg_idx,
                u"kind": u"tie",
            })
    return jobs


def create_stirrup_tags_for_column(
    doc,
    view,
    column,
    stirrup_rebars_ordered,
    errors=None,
    stirrup_policy=None,
    val_a=None,
    val_b=None,
    sel_a_text=None,
    sel_b_text=None,
):
    u"""
    Etiqueta estribos/trabas de confinamiento de una columna en ``view`` y crea
    una cota alineada vertical por lote (continuo o L/3).

    Returns:
        int: cantidad de etiquetas creadas.
    """
    if doc is None or view is None or column is None:
        return 0
    rebars = [rb for rb in (stirrup_rebars_ordered or []) if rb is not None]
    if not rebars:
        return 0
    if not _is_section_or_elevation_view(view):
        return 0

    if _collect_rebar_tag_symbol_map is None and _VIGA_ETQ is None:
        if errors is not None:
            errors.append(
                u"Etiquetas estribos: módulo Armado vigas o enfierrado_shaft_hashtag "
                u"no disponible."
            )
        return 0

    tag_map, mh_tag_map = _collect_confinement_tag_maps(doc)
    if not tag_map:
        if errors is not None:
            errors.append(
                u"Etiquetas estribos: cargue tipos OST_RebarTags para familia «{0}».".format(
                    CONFINEMENT_TAG_FAMILY,
                )
            )
        return 0

    try:
        ref_world = _column_reference_xyz(column, 0.0)
        loc = column.Location
        if hasattr(loc, u"Curve") and loc.Curve is not None:
            p0 = loc.Curve.Evaluate(0.0, True)
            p1 = loc.Curve.Evaluate(1.0, True)
            z_ref = 0.5 * (float(p0.Z) + float(p1.Z))
            ref_world = _column_reference_xyz(column, z_ref)
    except Exception:
        ref_world = _column_reference_xyz(column, 0.0)

    _, y_ref = _view_local_xy(view, ref_world)
    x_dim = _confinement_annotation_x_dim(view, column, rebars)

    jobs = _build_tag_jobs(
        rebars,
        val_a if val_a is not None else 4,
        val_b if val_b is not None else 4,
        sel_a_text or u"",
        sel_b_text or u"",
        stirrup_policy,
    )
    if not jobs:
        return 0

    n_per_seg = len(rebars)
    n_segments = 1
    try:
        from column_stirrup_creator import (
            STIRRUP_POLICY_THIRDS_L3,
            build_stirrup_rect_and_tie_defs,
        )
        if build_stirrup_rect_and_tie_defs is not None:
            _rd, _td = build_stirrup_rect_and_tie_defs(
                int(val_a or 4), int(val_b or 4),
                sel_a_text or u"", sel_b_text or u"",
            )
            n_per_seg = len(_rd) + len(_td)
            if n_per_seg >= 1:
                use_l3 = (
                    stirrup_policy == STIRRUP_POLICY_THIRDS_L3
                    and len(rebars) >= 3 * n_per_seg
                )
                n_segments = 3 if use_l3 else 1
    except Exception:
        pass

    seg_layout = {}
    for seg in range(n_segments):
        i0 = seg * n_per_seg
        i1 = i0 + n_per_seg
        seg_rebars = rebars[i0:i1]
        if not seg_rebars:
            break
        y_lot = _mean_view_local_y(view, ref_world, seg_rebars)
        y_loc = float(y_lot) if y_lot is not None else y_ref
        y_min, y_max = _segment_view_local_y_extent(view, seg_rebars)
        seg_layout[seg] = {
            u"x_dim": float(x_dim),
            u"y_loc": float(y_loc),
            u"y_min": y_min,
            u"y_max": y_max,
        }

    n_ok = 0

    for job in jobs:
        mode = job.get(u"mode")
        group = [rb for rb in (job.get(u"rebars") or []) if isinstance(rb, Rebar)]
        if not group:
            continue
        seg = int(job.get(u"seg", 0) or 0)
        layout = seg_layout.get(seg) or {}
        x_head = float(layout.get(u"x_dim", x_dim))
        y_loc = float(layout.get(u"y_loc", y_ref))
        head = _view_local_to_world(view, ref_world, x_head, y_loc)

        if mode == u"conf_tag_multihost":
            if not mh_tag_map:
                if errors is not None:
                    errors.append(
                        u"Etiquetas multihost: cargue familia «{0}».".format(
                            CONFINEMENT_MULTIHOST_TAG_FAMILY,
                        )
                    )
                continue
            ids = [_rebar_element_id(rb) for rb in group]
            ids = [i for i in ids if i is not None]
            ok, err = _crear_tag_multihost(doc, view, ids, mh_tag_map, head)
            if ok:
                n_ok += 1
            elif errors is not None:
                errors.append(
                    u"Etiqueta multihost estribos ({0} barras): {1}".format(
                        len(group), err or u"?"
                    )
                )
            continue

        for rb in group:
            sym, shape_lbl = _resolve_tag_symbol(
                doc, rb, tag_map, CONFINEMENT_TAG_FAMILY,
            )
            tag_type_id = _tag_type_id_from_map_value(sym)
            if tag_type_id is None:
                if errors is not None:
                    try:
                        rid = _rebar_element_id(rb).IntegerValue
                    except Exception:
                        rid = u"?"
                    errors.append(
                        u"Etiqueta estribo Id {0}: sin tipo para RebarShape «{1}».".format(
                            rid, shape_lbl or u"?",
                        )
                    )
                continue
            tag, err = _crear_tag_sin_leader(doc, view, rb, tag_type_id, head)
            if tag is not None:
                n_ok += 1
            elif errors is not None:
                try:
                    rid = _rebar_element_id(rb).IntegerValue
                except Exception:
                    rid = u"?"
                errors.append(
                    u"Etiqueta estribo Id {0}: {1}".format(rid, err or u"?")
                )

    for seg, layout in sorted(seg_layout.items()):
        ok, err = _create_vertical_confinement_dimension(
            doc,
            view,
            ref_world,
            layout.get(u"x_dim", x_dim),
            layout.get(u"y_min"),
            layout.get(u"y_max"),
        )
        if not ok and errors is not None and err:
            errors.append(
                u"Cota confinamiento lote {0}: {1}".format(int(seg) + 1, err)
            )

    return n_ok
