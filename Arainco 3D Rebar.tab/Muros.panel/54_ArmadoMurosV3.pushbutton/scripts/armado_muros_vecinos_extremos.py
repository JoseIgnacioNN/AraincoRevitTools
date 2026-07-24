# -*- coding: utf-8 -*-
"""
Muros vecinos en extremos — criterio Armado muros nodo (``wall_node_section``).

Expone detección de muros laterales en inicio/fin de ``LocationCurve`` para el boceto
de Area Reinforcement en Mallas muros lineales. No incluye suelos ni lógica de barras.
"""

from __future__ import print_function

import os
import sys

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import Wall

_WALL_NODE_MOD = None

# Caché de sesión: (doc_hash, host_id) -> lista de ElementId vecinos Wall
_MUROS_VECINOS_CACHE = {}
# (doc_hash, host_id, extremo) -> lista de ElementId
_VECINOS_EXTREMO_CACHE = {}
# (doc_hash, host_id) -> lista de ElementId cara lateral / T
_VECINOS_CARA_CACHE = {}


def clear_vecinos_caches():
    """Invalida cachés de vecinos (nueva sesión UI / cambio de selección)."""
    _MUROS_VECINOS_CACHE.clear()
    _VECINOS_EXTREMO_CACHE.clear()
    _VECINOS_CARA_CACHE.clear()


def _doc_cache_key(doc):
    try:
        return int(doc.GetHashCode())
    except Exception:
        try:
            return id(doc)
        except Exception:
            return 0


def _host_id_int(host):
    try:
        wns = _load_wall_node_section()
        if wns is not None:
            return int(wns._element_id_to_int(host.Id))
    except Exception:
        pass
    try:
        return int(host.Id.IntegerValue)
    except Exception:
        try:
            return int(host.Id.Value)
        except Exception:
            return None


def _walls_from_ids(doc, ids):
    out = []
    for eid in ids or []:
        try:
            el = doc.GetElement(eid)
        except Exception:
            el = None
        if el is not None and isinstance(el, Wall):
            out.append(el)
    return out


def _pushbutton_dir():
    here = os.path.dirname(os.path.abspath(__file__))
    try:
        import bootstrap_paths
        return bootstrap_paths.pin_local_scripts_first()
    except Exception:
        if here and here not in sys.path:
            sys.path.insert(0, here)
        return here


def _load_wall_node_section():
    global _WALL_NODE_MOD
    if _WALL_NODE_MOD is not None:
        return _WALL_NODE_MOD
    _pushbutton_dir()
    try:
        import wall_node_boolean_section_rps as wns
    except Exception:
        wns = None
    _WALL_NODE_MOD = wns
    return wns


def muros_vecinos_en_extremos(doc, host):
    """
    Muros en encuentro en los extremos del ``host`` (L, T, esquina, join).

    Mismo criterio que ``_elementos_para_union`` en Armado muros nodo, filtrado a
    ``Wall`` (sin forjados ni fundaciones).

    :returns: lista de instancias ``Wall`` (puede estar vacía).
    """
    if doc is None or host is None or not isinstance(host, Wall):
        return []

    hid = _host_id_int(host)
    cache_key = (_doc_cache_key(doc), hid)
    if hid is not None and cache_key in _MUROS_VECINOS_CACHE:
        return _walls_from_ids(doc, _MUROS_VECINOS_CACHE[cache_key])

    wns = _load_wall_node_section()
    if wns is None:
        return []

    try:
        wall_line, _curve_orig = wns._location_as_line(host)
    except Exception:
        return []
    if wall_line is None:
        return []

    try:
        elementos, _tol = wns._elementos_para_union(
            doc, host, wall_line, section_plane=None,
        )
    except Exception:
        return []

    host_i = None
    try:
        host_i = wns._element_id_to_int(host.Id)
    except Exception:
        pass

    out = []
    seen = set()
    id_list = []
    for el in elementos or []:
        if el is None or not isinstance(el, Wall):
            continue
        try:
            eid = wns._element_id_to_int(el.Id)
        except Exception:
            eid = None
        if eid is not None and host_i is not None and eid == host_i:
            continue
        if eid is not None:
            if eid in seen:
                continue
            seen.add(eid)
        out.append(el)
        try:
            id_list.append(el.Id)
        except Exception:
            pass

    if hid is not None:
        _MUROS_VECINOS_CACHE[cache_key] = id_list
    return out


def _tol_extremo_default():
    wns = _load_wall_node_section()
    if wns is None:
        return 0.25
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId
        return UnitUtils.ConvertToInternalUnits(80.0, UnitTypeId.Millimeters)
    except Exception:
        return 0.25


def vecino_en_extremo_muro(doc, host, extremo, neighbor):
    """
    True si ``neighbor`` participa en el encuentro en el extremo ``inicio``/``fin`` del host.
    """
    if doc is None or host is None or neighbor is None:
        return False
    if extremo not in (u"inicio", u"fin"):
        return False
    wns = _load_wall_node_section()
    if wns is None:
        return False
    try:
        wall_line, _co = wns._location_as_line(host)
    except Exception:
        return False
    if wall_line is None:
        return False
    tol_end = _tol_extremo_default()
    if not wns._es_muro_lateral_en_extremos(doc, host, wall_line, neighbor, tol_end):
        return False
    try:
        e0 = wall_line.GetEndPoint(0)
        e1 = wall_line.GetEndPoint(1)
        station = e0 if extremo == u"inicio" else e1
    except Exception:
        return False
    tol_curve = wns._tol_extremo_curva_muros(host, neighbor, tol_end)
    ol = neighbor.Location
    from Autodesk.Revit.DB import LocationCurve
    if not isinstance(ol, LocationCurve):
        return False
    oc = ol.Curve
    if oc is None:
        return False
    try:
        for p in (oc.GetEndPoint(0), oc.GetEndPoint(1)):
            if station.DistanceTo(p) <= tol_curve:
                return True
        if wns._dist_point_to_curve(station, oc) <= tol_curve:
            return True
    except Exception:
        pass
    par_lim = 0.10
    try:
        om = wns._midpoint_curve(oc)
        if om is not None:
            t, sep = wns._param_01_y_sep_eje_muro(e0, e1, om)
            if extremo == u"inicio":
                if t <= par_lim and sep <= tol_curve:
                    return True
            else:
                if t >= 1.0 - par_lim and sep <= tol_curve:
                    return True
    except Exception:
        pass
    if wns._esta_unido_por_join_geometry(doc, host, neighbor):
        try:
            for p in (oc.GetEndPoint(0), oc.GetEndPoint(1)):
                if station.DistanceTo(p) <= tol_curve * 1.5:
                    return True
        except Exception:
            pass
    return False


def vecinos_en_extremo(doc, host, extremo):
    """Vecinos del host filtrados al extremo indicado."""
    if doc is None or host is None or not isinstance(host, Wall):
        return []
    if extremo not in (u"inicio", u"fin"):
        return []
    hid = _host_id_int(host)
    cache_key = (_doc_cache_key(doc), hid, extremo)
    if hid is not None and cache_key in _VECINOS_EXTREMO_CACHE:
        return _walls_from_ids(doc, _VECINOS_EXTREMO_CACHE[cache_key])

    out = []
    id_list = []
    for w in muros_vecinos_en_extremos(doc, host):
        if vecino_en_extremo_muro(doc, host, extremo, w):
            out.append(w)
            try:
                id_list.append(w.Id)
            except Exception:
                pass
    if hid is not None:
        _VECINOS_EXTREMO_CACHE[cache_key] = id_list
    return out


def vecinos_cara_lateral_o_t(doc, host):
    """
    Muros vecinos en encuentro T o en cara lateral a mitad de tramo (no en extremos).

    Complementa ``vecinos_en_extremo`` para el boceto de elevación.
    """
    if doc is None or host is None or not isinstance(host, Wall):
        return []

    hid = _host_id_int(host)
    cache_key = (_doc_cache_key(doc), hid)
    if hid is not None and cache_key in _VECINOS_CARA_CACHE:
        return _walls_from_ids(doc, _VECINOS_CARA_CACHE[cache_key])

    wns = _load_wall_node_section()
    if wns is None:
        return []

    try:
        wall_line, _co = wns._location_as_line(host)
    except Exception:
        return []
    if wall_line is None:
        return []

    tol_end = _tol_extremo_default()
    try:
        e0 = wall_line.GetEndPoint(0)
        e1 = wall_line.GetEndPoint(1)
    except Exception:
        return []

    extremo_ids = set()
    for ex in (u"inicio", u"fin"):
        for w in vecinos_en_extremo(doc, host, ex):
            try:
                eid = wns._element_id_to_int(w.Id)
            except Exception:
                eid = None
            if eid is not None:
                extremo_ids.add(eid)

    out = []
    seen = set()
    id_list = []
    from Autodesk.Revit.DB import LocationCurve

    for w in muros_vecinos_en_extremos(doc, host):
        try:
            eid = wns._element_id_to_int(w.Id)
        except Exception:
            eid = None
        if eid is not None and eid in extremo_ids:
            continue
        if eid is not None:
            if eid in seen:
                continue
            seen.add(eid)
        ol = w.Location
        if not isinstance(ol, LocationCurve):
            continue
        oc = ol.Curve
        if oc is None:
            continue
        if not wns._es_muro_lateral_en_extremos(doc, host, wall_line, w, tol_end):
            continue
        if not wns._es_muro_encuentro_cara_lateral_o_t(host, wall_line, w, oc, tol_end):
            continue
        try:
            om = wns._midpoint_curve(oc)
            if om is not None:
                t, _sep = wns._param_01_y_sep_eje_muro(e0, e1, om)
                if t < 0.04 or t > 0.96:
                    continue
        except Exception:
            pass
        out.append(w)
        try:
            id_list.append(w.Id)
        except Exception:
            pass

    if hid is not None:
        _VECINOS_CARA_CACHE[cache_key] = id_list
    return out


def vecino_principal_encuentro_l(doc, host, extremo):
    """
    Primer muro vecino clasificado como encuentro L en ``extremo``, o ``None``.
    """
    try:
        import armado_muros_cabezal_encuentro_l as enc_l
    except Exception:
        enc_l = None
    if enc_l is None:
        return None
    best = None
    best_d = None
    geom_host = None
    try:
        from armado_muros_cabezal import _wall_longitudinal_at_extremo
        geom_host = _wall_longitudinal_at_extremo(host, extremo)
    except Exception:
        geom_host = None
    station = geom_host[u"pt_extremo"] if geom_host else None
    for w in vecinos_en_extremo(doc, host, extremo):
        kind = enc_l.clasificar_encuentro_en_extremo(doc, host, w, extremo)
        if kind != enc_l.CABEZAL_ENC_TIPO_L:
            continue
        if station is None:
            return w
        try:
            d = float(station.DistanceTo(
                enc_l.cabezal_encuentro_l_p_join(doc, host, w, extremo),
            ))
        except Exception:
            d = 0.0
        if best is None or best_d is None or d < best_d:
            best = w
            best_d = d
    return best
