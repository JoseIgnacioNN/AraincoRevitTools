# -*- coding: utf-8 -*-
"""
Etiquetas Structural Rebar Tag para barras de malla (Armado muros lineales).

Familia ``EST_A_STRUCTURAL REBAR TAG_MALLA``:
- barras verticales → tipo ``Vertical``
- barras horizontales → tipo ``Horizontal``

Una etiqueta **multihost** por orientación (vertical / horizontal), anclada a la barra
de la **cara interior**; la barra de la cara exterior se agrega con ``AddReferences``.

La orientación y el tipo de etiqueta se resuelven **solo** con el parámetro de instancia
``Armadura_Malla_Orientacion`` (``V.`` → tipo ``Vertical``; ``H.`` → tipo ``Horizontal``).
No se usa geometría ni ``RebarShape`` para elegir el tipo.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    ElementId,
    IndependentTag,
    Reference,
    SubTransaction,
    TagMode,
    TagOrientation,
    Transaction,
)
from Autodesk.Revit.DB.Structure import Rebar

try:
    unicode
except NameError:
    unicode = str

MALLA_REBAR_TAG_FAMILY_NAME = u"EST_A_STRUCTURAL REBAR TAG_MALLA"
MALLA_TAG_TYPE_VERTICAL = u"Vertical"
MALLA_TAG_TYPE_HORIZONTAL = u"Horizontal"
_MAX_TAG_WARNINGS = 12


def _resultado_vacio():
    return {
        u"n_ok": 0,
        u"n_fail": 0,
        u"n_skip": 0,
        u"messages": [],
    }


def _append_msg(result, msg):
    if result is None or not msg:
        return
    msgs = result.setdefault(u"messages", [])
    if len(msgs) < _MAX_TAG_WARNINGS:
        msgs.append(msg)


def _norm_label(s):
    try:
        t = unicode(s).strip().lower()
    except Exception:
        try:
            t = str(s or u"").strip().lower()
        except Exception:
            return u""
    for ch in (u"\xa0", u"\u200b", u"\ufeff"):
        t = t.replace(ch, u"")
    return u" ".join(t.split())


def _symbol_labels(sym):
    labels = []
    try:
        from armado_muros_cabezal_tags import _symbol_type_labels
        labels.extend(_symbol_type_labels(sym) or [])
    except Exception:
        pass
    try:
        nm = getattr(sym, "Name", None)
        if nm:
            labels.append(unicode(nm))
    except Exception:
        pass
    out = []
    seen = set()
    for raw in labels:
        k = _norm_label(raw)
        if k and k not in seen:
            seen.add(k)
            out.append(k)
    return out


def _symbol_es_tipo_malla(sym, tipo):
    """
    ``tipo``: ``u"vertical"`` | ``u"horizontal"``.
    Coincidencia por nombre de tipo/símbolo dentro de la familia malla.
    """
    if sym is None or not tipo:
        return False
    want = _norm_label(tipo)
    if not want:
        return False
    other = u"horizontal" if want == u"vertical" else u"vertical"
    for lab in _symbol_labels(sym):
        if other in lab and want not in lab:
            return False
        if lab == want or want in lab.split():
            return True
    return False


def _activar_symbol(sym):
    if sym is None:
        return sym
    try:
        if not sym.IsActive:
            sym.Activate()
    except Exception:
        pass
    return sym


def _resolve_tag_symbols(document):
    sym_v = None
    sym_h = None
    try:
        from armado_muros_cabezal_tags import (
            _collect_tag_symbol_map,
            _family_symbols_rebar_tag,
            _lookup_tag_symbol,
        )
    except Exception as ex:
        return None, None, unicode(ex)

    for sym in _family_symbols_rebar_tag(document, MALLA_REBAR_TAG_FAMILY_NAME):
        if sym_v is None and _symbol_es_tipo_malla(sym, u"vertical"):
            sym_v = _activar_symbol(sym)
        if sym_h is None and _symbol_es_tipo_malla(sym, u"horizontal"):
            sym_h = _activar_symbol(sym)

    tag_map = _collect_tag_symbol_map(document, MALLA_REBAR_TAG_FAMILY_NAME)
    if sym_v is None and tag_map:
        sym_v = _lookup_tag_symbol(tag_map, MALLA_TAG_TYPE_VERTICAL)
        sym_v = _activar_symbol(sym_v)
    if sym_h is None and tag_map:
        sym_h = _lookup_tag_symbol(tag_map, MALLA_TAG_TYPE_HORIZONTAL)
        sym_h = _activar_symbol(sym_h)

    if not tag_map and sym_v is None and sym_h is None:
        return None, None, (
            u"no hay tipos OST_RebarTags para familia «{0}».".format(
                MALLA_REBAR_TAG_FAMILY_NAME,
            )
        )
    if sym_v is None and sym_h is None:
        return None, None, (
            u"familia «{0}» sin tipos «{1}» ni «{2}».".format(
                MALLA_REBAR_TAG_FAMILY_NAME,
                MALLA_TAG_TYPE_VERTICAL,
                MALLA_TAG_TYPE_HORIZONTAL,
            )
        )
    return sym_v, sym_h, None


def _head_pos_centroide_muro(wall, view):
    try:
        from armado_muros_etiqueta_malla import (
            _proyectar_punto_plano_vista,
            centroide_geometria_muro,
        )
    except Exception:
        return None
    if wall is None:
        return None
    c = centroide_geometria_muro(wall, view)
    if c is None:
        return None
    return _proyectar_punto_plano_vista(c, view)


def _crear_tag_malla_rebar(document, view, rebar, tag_symbol_id, head_pos):
    if document is None or view is None or rebar is None or tag_symbol_id is None:
        return None, u"parámetros inválidos"
    if head_pos is None:
        return None, u"sin punto de inserción"
    try:
        from armado_muros_cabezal_tags import _referencias_tag_rebar
    except Exception:
        return None, u"módulo referencias tag no disponible"
    refs = _referencias_tag_rebar(document, rebar, view)
    if not refs:
        return None, u"sin referencia API"
    orient = TagOrientation.Horizontal
    add_leader = False
    last_ex = None
    for ref in refs:
        try:
            tag = IndependentTag.Create(
                document,
                tag_symbol_id,
                view.Id,
                ref,
                add_leader,
                orient,
                head_pos,
            )
            if tag is not None:
                try:
                    tag.ChangeTypeId(tag_symbol_id)
                except Exception:
                    try:
                        tag.SetTypeId(tag_symbol_id)
                    except Exception:
                        pass
                _aplicar_estilo_tag_malla(tag, head_pos)
                return tag, None
        except Exception as ex:
            last_ex = ex
            tag = None
    for ref in refs:
        try:
            tag = IndependentTag.Create(
                document,
                view.Id,
                ref,
                add_leader,
                TagMode.TM_ADDBY_CATEGORY,
                orient,
                head_pos,
            )
            if tag is not None:
                try:
                    tag.ChangeTypeId(tag_symbol_id)
                except Exception:
                    try:
                        tag.SetTypeId(tag_symbol_id)
                    except Exception:
                        pass
                _aplicar_estilo_tag_malla(tag, head_pos)
                return tag, None
        except Exception as ex:
            last_ex = ex
    if last_ex is not None:
        try:
            return None, unicode(last_ex)
        except Exception:
            return None, str(last_ex)
    return None, u"no se pudo crear IndependentTag"


def _crear_tag_malla_multihost_rebar(
    document,
    view,
    primary_rebar,
    tag_symbol_id,
    head_pos,
    extra_rebars=None,
):
    """
    Etiqueta multihost: host principal = barra interior; hosts extra vía ``AddReferences``.
    """
    tag, err = _crear_tag_malla_rebar(
        document, view, primary_rebar, tag_symbol_id, head_pos,
    )
    if tag is None:
        return None, err

    extras = []
    try:
        pid = primary_rebar.Id
    except Exception:
        pid = None
    for rb in extra_rebars or []:
        if rb is None or not isinstance(rb, Rebar):
            continue
        try:
            if pid is not None and rb.Id == pid:
                continue
        except Exception:
            pass
        extras.append(rb)
    if not extras:
        return tag, None

    add_fn = getattr(tag, "AddReferences", None)
    if add_fn is None:
        return tag, None

    extra_refs = []
    try:
        from armado_muros_cabezal_tags import _refs_multihost_para_rebars
        extra_refs = _refs_multihost_para_rebars(
            document, view, [rb.Id for rb in extras],
        )
    except Exception:
        extra_refs = []

    if not extra_refs:
        return tag, None

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
        return tag, None
    try:
        add_fn(refs_add)
        st.Commit()
        _aplicar_estilo_tag_malla(tag, head_pos)
        return tag, None
    except Exception as ex_mh:
        try:
            st.RollBack()
        except Exception:
            pass
        try:
            msg = unicode(ex_mh)
        except Exception:
            msg = str(ex_mh)
        return tag, u"Multihost malla (AddReferences): {0}".format(msg)


def _aplicar_estilo_tag_malla(tag, head):
    if tag is None or head is None:
        return
    try:
        tag.TagOrientation = TagOrientation.Horizontal
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


def _orient_rebar_malla_desde_param(rebar):
    """
    ``u"vertical"`` | ``u"horizontal"`` | ``None`` según ``Armadura_Malla_Orientacion``.

    ``V.`` → vertical; ``H.`` → horizontal.
    """
    if rebar is None:
        return None
    try:
        from armado_muros_rebar_params import get_armadura_malla_orientacion
        orient = get_armadura_malla_orientacion(rebar)
        if orient in (u"vertical", u"horizontal"):
            return orient
    except Exception:
        pass
    return None


def _agrupar_rebars_malla_por_cara_orient(
    document,
    host,
    rebar_ids,
    result=None,
):
    """
    ``{ vertical: {interior: Rebar, exterior: Rebar}, horizontal: {...} }``.

    Agrupa por ``Armadura_Malla_Orientacion`` (``V.`` / ``H.``) y cara int/ext del muro.
    """
    out = {u"vertical": {}, u"horizontal": {}}
    por_orient = {u"vertical": [], u"horizontal": []}

    for eid in rebar_ids or []:
        try:
            rb = document.GetElement(eid)
        except Exception:
            continue
        if rb is None or not isinstance(rb, Rebar):
            continue
        orient = _orient_rebar_malla_desde_param(rb)
        if orient in por_orient:
            por_orient[orient].append(rb)
        elif result is not None:
            result[u"n_skip"] += 1
            _append_msg(
                result,
                u"Rebar {0}: sin «Armadura_Malla_Orientacion» (V./H.).".format(
                    _rebar_id_int(rb),
                ),
            )

    for orient in (u"vertical", u"horizontal"):
        rbs = por_orient.get(orient) or []
        if not rbs:
            continue
        if len(rbs) == 1:
            rb = rbs[0]
            cara = _cara_rebar_en_muro(rb, host)
            if cara:
                out[orient][cara] = rb
            else:
                out[orient][u"interior"] = rb
            continue
        scored = []
        for rb in rbs:
            off = _offset_rebar_along_wall_normal(rb, host)
            if off is None:
                off = 0.0
            scored.append((off, rb))
        scored.sort(key=lambda item: item[0])
        out[orient][u"interior"] = scored[0][1]
        out[orient][u"exterior"] = scored[-1][1]
        if len(scored) > 2:
            for off, rb in scored[1:-1]:
                cara = _cara_rebar_en_muro(rb, host)
                if cara and cara not in out[orient]:
                    out[orient][cara] = rb

    return out


def _bb_center_xyz(elem):
    try:
        from arearein_exterior_h_l135_rps import _bb_center_xyz as _fn
        return _fn(elem)
    except Exception:
        pass
    try:
        bb = elem.get_BoundingBox(None) if elem is not None else None
        if bb is None:
            return None
        return (bb.Min + bb.Max).Multiply(0.5)
    except Exception:
        return None


def _offset_rebar_along_wall_normal(rebar, host):
    """Proyección del centro de barra sobre ``Wall.Orientation`` (ft)."""
    if rebar is None or host is None:
        return None
    try:
        from Autodesk.Revit.DB import Wall
        if not isinstance(host, Wall):
            return None
        ori = host.Orientation
        if ori is None or ori.GetLength() < 1e-12:
            return None
        ori = ori.Normalize()
        pr = _bb_center_xyz(rebar)
        pw = _bb_center_xyz(host)
        if pr is None or pw is None:
            return None
        return float((pr - pw).DotProduct(ori))
    except Exception:
        return None


def _cara_rebar_en_muro(rebar, host):
    if rebar is None or host is None:
        return None
    try:
        from arearein_exterior_h_l135_rps import _rebar_solo_cara_exterior
        if _rebar_solo_cara_exterior(rebar, host):
            return u"exterior"
    except Exception:
        pass
    try:
        from arearein_interior_h_l135_rps import _rebar_solo_cara_interior
        if _rebar_solo_cara_interior(rebar, host):
            return u"interior"
    except Exception:
        pass
    try:
        from armado_muros_verticales_embed_colision import (
            _rebar_es_vertical_exterior,
            _rebar_es_vertical_interior,
        )
        if _rebar_es_vertical_exterior(rebar, host):
            return u"exterior"
        if _rebar_es_vertical_interior(rebar, host):
            return u"interior"
    except Exception:
        pass
    off = _offset_rebar_along_wall_normal(rebar, host)
    if off is not None:
        if off >= 0.0:
            return u"exterior"
        return u"interior"
    return None


def _params_dict_for_wall(params_por_muro_id, wid):
    if not params_por_muro_id or wid is None:
        return None
    keys = [wid]
    try:
        keys.append(int(wid))
    except Exception:
        pass
    try:
        keys.append(unicode(wid))
    except Exception:
        try:
            keys.append(str(wid))
        except Exception:
            pass
    seen = set()
    for key in keys:
        if key in seen:
            continue
        seen.add(key)
        try:
            tup = params_por_muro_id.get(key)
            if tup is not None:
                return tup[0]
        except Exception:
            continue
    return None


def _vista_permite_tags_malla(view):
    try:
        from armado_muros_cabezal_tags import _vista_permite_rebar_tags
        return bool(_vista_permite_rebar_tags(view))
    except Exception:
        pass
    try:
        from armado_muros_etiqueta_malla import es_vista_elevacion_seccion
        return bool(es_vista_elevacion_seccion(view))
    except Exception:
        return False


def etiquetar_rebars_malla_muro(
    document,
    view,
    wall,
    rebar_ids,
    sym_vertical,
    sym_horizontal,
    params_dict=None,
    muro_contencion=False,
    result=None,
):
    """
    Etiqueta malla de un muro (sin transacción): una etiqueta multihost por orientación,
    anclada a la barra **interior** y con la **exterior** como host adicional.

    Tipo de etiqueta según ``Armadura_Malla_Orientacion`` (``V.`` / ``H.``).
    """
    res = result if result is not None else _resultado_vacio()
    if document is None or view is None or wall is None or not rebar_ids:
        res[u"n_skip"] += len(rebar_ids or [])
        return res
    if sym_vertical is None and sym_horizontal is None:
        res[u"n_fail"] += 1
        _append_msg(res, u"Etiquetas malla rebar: sin símbolos Vertical/Horizontal.")
        return res

    head = _head_pos_centroide_muro(wall, view)
    if head is None:
        try:
            wid = int(wall.Id.Value)
        except Exception:
            try:
                wid = int(wall.Id.IntegerValue)
            except Exception:
                wid = u"?"
        res[u"n_fail"] += 1
        _append_msg(
            res,
            u"Muro {0}: etiqueta rebar malla — sin centroide.".format(wid),
        )
        return res

    host = wall
    grupos = _agrupar_rebars_malla_por_cara_orient(
        document, host, rebar_ids, result=res,
    )

    for orient in (u"vertical", u"horizontal"):
        hosts = grupos.get(orient) or {}
        rb_int = hosts.get(u"interior")
        rb_ext = hosts.get(u"exterior")
        if rb_int is None:
            if rb_ext is not None:
                res[u"n_skip"] += 1
                try:
                    wid = int(host.Id.Value)
                except Exception:
                    try:
                        wid = int(host.Id.IntegerValue)
                    except Exception:
                        wid = u"?"
                _append_msg(
                    res,
                    u"Muro {0}: sin barra {1} interior para etiqueta multihost "
                    u"(exterior id {2}).".format(
                        wid,
                        orient,
                        _rebar_id_int(rb_ext),
                    ),
                )
            elif not hosts:
                res[u"n_skip"] += 1
                try:
                    wid = int(host.Id.Value)
                except Exception:
                    try:
                        wid = int(host.Id.IntegerValue)
                    except Exception:
                        wid = u"?"
                _append_msg(
                    res,
                    u"Muro {0}: sin par {1} interior/exterior para etiqueta multihost.".format(
                        wid, orient,
                    ),
                )
            continue

        es_vert = orient == u"vertical"
        sym = sym_vertical if es_vert else sym_horizontal
        tipo_nombre = (
            MALLA_TAG_TYPE_VERTICAL if es_vert else MALLA_TAG_TYPE_HORIZONTAL
        )
        if sym is None:
            res[u"n_fail"] += 1
            _append_msg(
                res,
                u"Muro {0}: sin tipo de etiqueta «{1}».".format(
                    _rebar_id_int(rb_int),
                    tipo_nombre,
                ),
            )
            continue

        rb_ext = hosts.get(u"exterior")
        extras = [rb_ext] if rb_ext is not None else []
        tag, err = _crear_tag_malla_multihost_rebar(
            document, view, rb_int, sym.Id, head, extras,
        )
        if tag is not None:
            res[u"n_ok"] += 1
            if err:
                _append_msg(
                    res,
                    u"Rebar {0} ({1}): {2}.".format(
                        _rebar_id_int(rb_int),
                        tipo_nombre,
                        err,
                    ),
                )
        else:
            res[u"n_fail"] += 1
            _append_msg(
                res,
                u"Rebar {0} ({1}): {2}.".format(
                    _rebar_id_int(rb_int),
                    tipo_nombre,
                    err or u"fallo",
                ),
            )
    return res


def _rebar_id_int(rebar):
    if rebar is None:
        return u"?"
    try:
        return int(rebar.Id.Value)
    except Exception:
        try:
            return int(rebar.Id.IntegerValue)
        except Exception:
            return u"?"


def etiquetar_rebars_malla_en_vista(
    document,
    view,
    rebars_por_muro_id,
    params_por_muro_id=None,
    muro_contencion=False,
    walls_by_id=None,
):
    """
    Etiqueta malla en ``rebars_por_muro_id`` (una transacción por muro).

    Multihost interior + exterior por ``Armadura_Malla_Orientacion`` (``V.`` / ``H.``).
    """
    res = _resultado_vacio()
    if document is None or view is None or not rebars_por_muro_id:
        return res
    if not _vista_permite_tags_malla(view):
        _append_msg(
            res,
            u"Etiquetas rebar malla: use planta, alzado o sección (no plantilla ni 3D).",
        )
        return res

    sym_v, sym_h, err_sym = _resolve_tag_symbols(document)
    if err_sym:
        res[u"n_fail"] += 1
        _append_msg(res, u"Etiquetas rebar malla: {0}.".format(err_sym))
        return res

    try:
        document.Regenerate()
    except Exception:
        pass

    wall_map = walls_by_id or {}
    for wid, eid_list in rebars_por_muro_id.items():
        if not eid_list:
            continue
        wall = wall_map.get(wid)
        if wall is None:
            try:
                wall = wall_map.get(int(wid))
            except Exception:
                pass
        if wall is None:
            try:
                wall = document.GetElement(ElementId(int(wid)))
            except Exception:
                wall = None
        if wall is None:
            res[u"n_fail"] += len(eid_list)
            _append_msg(res, u"Muro {0}: no encontrado para etiquetar.".format(wid))
            continue

        params_dict = _params_dict_for_wall(params_por_muro_id, wid)
        t = Transaction(
            document,
            u"Arainco: Etiquetas rebar malla muro {0}".format(wid),
        )
        try:
            t.Start()
            etiquetar_rebars_malla_muro(
                document,
                view,
                wall,
                eid_list,
                sym_v,
                sym_h,
                params_dict=params_dict,
                muro_contencion=muro_contencion,
                result=res,
            )
            t.Commit()
        except Exception as ex:
            try:
                if t.HasStarted():
                    t.RollBack()
            except Exception:
                pass
            res[u"n_fail"] += len(eid_list)
            _append_msg(
                res,
                u"Muro {0}: etiquetas rebar malla — {1}.".format(wid, unicode(ex)),
            )
    return res
