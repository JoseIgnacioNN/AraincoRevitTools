# -*- coding: utf-8 -*-
"""
Muros vecinos en extremos — uniones de muro + p_join (Armado Muros V3).

Criterio principal: ``LocationCurve.get_ElementsAtJoin`` (unión de muro en
punta: L, T, cadena). Eso es lo que Revit usa en esquinas; **no** es
*Unir geometría* (``JoinGeometryUtils``).

Complemento: Join Geometry, intersección de sólidos y bbox cercano; se
confirma con ejes no paralelos / ``p_join`` si no hay unión de muro.

Excluye apilados (muro sobre/bajo) y uniones longitudinales paralelas.
Solo ``Wall``. No barre ``Floor`` ni forjados.
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
    Muros unidos en los extremos del ``host`` (L, T, esquina).

    Primero uniones de muro (``ElementsAtJoin``); luego Join Geometry,
    intersección de sólidos y bbox cercano. Cada candidato se confirma
    con ``_pjoin_valido_en_extremo``. Excluye apilados y muros paralelos.
    **No** consulta ``Floor`` / forjados.

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
        from Autodesk.Revit.DB import (
            ElementIntersectsElementFilter,
            FilteredElementCollector,
            UnitTypeId,
            UnitUtils,
        )
    except Exception:
        return []

    try:
        tol_end = UnitUtils.ConvertToInternalUnits(80.0, UnitTypeId.Millimeters)
    except Exception:
        tol_end = _tol_extremo_default()

    out = []
    seen = set()
    host_i = hid

    def _try_add(w2):
        if w2 is None or not isinstance(w2, Wall):
            return
        try:
            eid = wns._element_id_to_int(w2.Id)
        except Exception:
            eid = None
        if eid is not None and host_i is not None and eid == host_i:
            return
        if eid is not None and eid in seen:
            return
        if not _muro_unido_por_pjoin_en_algun_extremo(
            doc, host, w2, wall_line=wall_line, tol_end=tol_end,
        ):
            return
        if eid is not None:
            seen.add(eid)
        out.append(w2)

    # 1) Uniones de muro en puntas (ElementsAtJoin) — criterio Revit de esquina
    try:
        for jid in wns._coleccion_ids_wall_joins_extremos(host):
            el = doc.GetElement(jid)
            if el is not None and isinstance(el, Wall):
                _try_add(el)
    except Exception:
        pass

    # 2) Unir geometría (Join Geometry; no es lo mismo que unión de muro)
    try:
        for jid in wns._coleccion_ids_unidas(doc, host):
            el = doc.GetElement(jid)
            if el is not None and isinstance(el, Wall):
                _try_add(el)
    except Exception:
        pass

    # 3) Muros que intersectan el sólido del host
    try:
        xf = ElementIntersectsElementFilter(host)
        for eid in (
            FilteredElementCollector(doc)
            .OfClass(Wall)
            .WherePasses(xf)
            .ToElementIds()
        ):
            _try_add(doc.GetElement(eid))
    except Exception:
        pass

    # 4) Bbox cercano: L/T a inglete suelen no solapar volumen ni Unir geom.
    try:
        from Autodesk.Revit.DB import BoundingBoxIntersectsFilter
        pad = max(float(tol_end) * 4.0, 1.0)
        try:
            pad = max(pad, abs(float(host.Width)) * 2.0)
        except Exception:
            pass
        olh = wns._outline_host_inflado(host, pad)
        if olh is not None:
            bf = BoundingBoxIntersectsFilter(olh)
            for eid in (
                FilteredElementCollector(doc)
                .OfClass(Wall)
                .WherePasses(bf)
                .ToElementIds()
            ):
                _try_add(doc.GetElement(eid))
    except Exception:
        pass

    id_list = []
    for w in out:
        try:
            id_list.append(w.Id)
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


def _tol_z_apilamiento():
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId
        return UnitUtils.ConvertToInternalUnits(40.0, UnitTypeId.Millimeters)
    except Exception:
        return 0.12


def _encuentro_l_mod():
    try:
        import armado_muros_cabezal_encuentro_l as enc_l
        return enc_l
    except Exception:
        return None


def _es_apilado_sobre_o_bajo(host, neighbor):
    wns = _load_wall_node_section()
    if wns is None:
        return False
    tol_z = _tol_z_apilamiento()
    try:
        if wns._muro_apilado_bajo_muro_principal(host, neighbor, tol_z):
            return True
        if wns._muro_apilado_sobre_muro_principal(host, neighbor, tol_z):
            return True
    except Exception:
        pass
    return False


def _dist_xy(a, b):
    """Distancia en planta (ignora Z: bases distintas no anulan el encuentro)."""
    if a is None or b is None:
        return 1e30
    try:
        dx = float(a.X) - float(b.X)
        dy = float(a.Y) - float(b.Y)
        return (dx * dx + dy * dy) ** 0.5
    except Exception:
        return 1e30


def _candidato_geometrico_en_extremo(doc, host, extremo, neighbor, wall_line=None, tol_end=None):
    """
    Proximidad del extremo del host al eje del vecino (sin p_join).
    Usado como prefiltro antes de validar intersección de ejes.
    """
    if doc is None or host is None or neighbor is None:
        return False
    if extremo not in (u"inicio", u"fin"):
        return False
    wns = _load_wall_node_section()
    if wns is None:
        return False
    try:
        if wns._esta_unido_por_wall_join_en_extremo(host, neighbor, extremo):
            return True
    except Exception:
        pass
    if wall_line is None:
        try:
            wall_line, _co = wns._location_as_line(host)
        except Exception:
            return False
    if wall_line is None:
        return False
    if tol_end is None:
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
            if _dist_xy(station, p) <= tol_curve:
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
    if wns._esta_unido_por_join_geometry(doc, host, neighbor) or wns._esta_unido_por_wall_join(
        host, neighbor
    ):
        try:
            for p in (oc.GetEndPoint(0), oc.GetEndPoint(1)):
                if _dist_xy(station, p) <= tol_curve * 1.5:
                    return True
        except Exception:
            pass
    return False


def _pjoin_valido_en_extremo(doc, host, neighbor, extremo, wall_line=None, tol_end=None):
    """
    True si ``neighbor`` se une al ``host`` en ``extremo``: unión de muro
    (``ElementsAtJoin``) o, si no hay, intersección de ejes (p_join).
    Excluye apilados y muros paralelos longitudinales.
    """
    if doc is None or host is None or neighbor is None:
        return False
    if extremo not in (u"inicio", u"fin"):
        return False
    if _es_apilado_sobre_o_bajo(host, neighbor):
        return False

    enc_l = _encuentro_l_mod()
    if enc_l is None:
        return False
    if enc_l._dot_dirs_wall(host, neighbor) > 0.92:
        return False

    wns = _load_wall_node_section()
    wall_join_end = False
    if wns is not None:
        try:
            wall_join_end = bool(
                wns._esta_unido_por_wall_join_en_extremo(host, neighbor, extremo)
            )
        except Exception:
            wall_join_end = False

    # Unión de muro en esa punta + ejes no paralelos = encuentro (L/T).
    if wall_join_end:
        return True

    if not _candidato_geometrico_en_extremo(
        doc, host, extremo, neighbor, wall_line=wall_line, tol_end=tol_end,
    ):
        return False

    try:
        from armado_muros_cabezal import _wall_longitudinal_at_extremo
    except Exception:
        return False
    geom_h = _wall_longitudinal_at_extremo(host, extremo)
    if geom_h is None:
        return False
    station = geom_h.get(u"pt_extremo")
    if station is None:
        return False

    lc_h = enc_l.location_curve_wall(host) if enc_l.location_curve_wall else None
    lc_n = enc_l.location_curve_wall(neighbor) if enc_l.location_curve_wall else None
    if lc_h is None or lc_n is None:
        return False
    o1, d1 = enc_l._line_dir_xy(lc_h)
    o2, d2 = enc_l._line_dir_xy(lc_n)
    if enc_l._intersect_lines_xy(o1, d1, o2, d2) is None:
        return False

    try:
        p_join = enc_l.cabezal_encuentro_l_p_join(doc, host, neighbor, extremo)
    except Exception:
        return False
    if p_join is None:
        return False

    if tol_end is None:
        tol_end = _tol_extremo_default()
    tol_curve = tol_end
    if wns is not None:
        try:
            tol_curve = wns._tol_extremo_curva_muros(host, neighbor, tol_end)
        except Exception:
            tol_curve = tol_end
    try:
        if _dist_xy(station, p_join) > float(tol_curve) * 1.5:
            return False
    except Exception:
        return False
    return True


def _muro_unido_por_pjoin_en_algun_extremo(doc, host, neighbor, wall_line=None, tol_end=None):
    for extremo in (u"inicio", u"fin"):
        if _pjoin_valido_en_extremo(
            doc, host, neighbor, extremo,
            wall_line=wall_line, tol_end=tol_end,
        ):
            return True
    return False


def vecino_en_extremo_muro(doc, host, extremo, neighbor):
    """
    True si ``neighbor`` participa en el encuentro en el extremo ``inicio``/``fin`` del host
    (intersección de ejes p_join, criterio Armado Muros V3).
    """
    return _pjoin_valido_en_extremo(doc, host, neighbor, extremo)


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

    No se usa para el boceto de elevación: ahí solo se dibujan encuentros en
    puntas (``vecinos_en_extremo``). Disponible para otros consumidores.
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
