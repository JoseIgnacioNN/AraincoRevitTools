# -*- coding: utf-8 -*-
"""
Reubicar / aplicar etiquetas de muro (OST_WallTags) por encima de etiquetas de malla.

Tras crear ``EST_A_STRUCTURAL REBAR TAG_MALLA``, coloca
``EST_A_WALL TAG_ELEVACION_MHA`` / ``Espesor Muro`` con cabeza desplazada en
``View.UpDirection`` por encima de las cabezas de malla.

Si ya hay wall tags mono-host: delete + recreate con ese tipo.
Si no hay: crea la etiqueta de espesor.
Multi-host se omite con aviso.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    FamilySymbol,
    FilteredElementCollector,
    IndependentTag,
    Reference,
    TagMode,
    TagOrientation,
    XYZ,
)

try:
    unicode
except NameError:
    unicode = str

# Familia/tipo de etiqueta de espesor (elevación MHA) — imagen Type Selector.
WALL_TAG_FAMILY_NAME = u"EST_A_WALL TAG_ELEVACION_MHA"
WALL_TAG_TYPE_NAME = u"Espesor Muro"

# Calibración imagen (escala 1:50, altura de texto en papel 2.5 mm).
# Offset un poco más corto para acercar M.H.A. al bloque V./H. sin solapar.
_WALL_TAG_TEXT_HEIGHT_PAPER_MM = 2.5
_WALL_TAG_SCALE_REFERENCE = 50
_WALL_TAG_OFFSET_TEXT_HEIGHTS = 1.5  # → 3.75 mm papel → 187.5 mm modelo @ 1:50
_WALL_TAG_OFFSET_TEXT_HEIGHTS_SINGLE = 1.25
_MAX_WARNINGS = 12


def _resultado_vacio():
    return {
        u"n_ok": 0,
        u"n_skip_multihost": 0,
        u"n_fail": 0,
        u"n_skip_no_tag": 0,
        u"messages": [],
    }


def _append_msg(result, msg):
    if result is None or not msg:
        return
    msgs = result.setdefault(u"messages", [])
    if len(msgs) < _MAX_WARNINGS:
        msgs.append(msg)


def _eid_int(eid):
    if eid is None:
        return None
    try:
        return int(eid.IntegerValue)
    except Exception:
        try:
            return int(eid.Value)
        except Exception:
            return None


def _wall_id_int(wall):
    if wall is None:
        return None
    try:
        return _eid_int(wall.Id)
    except Exception:
        return None


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


def _symbol_family_name(sym):
    try:
        fam = sym.Family
        if fam is not None:
            return unicode(fam.Name or u"")
    except Exception:
        pass
    return u""


def _symbol_type_name(sym):
    try:
        nm = getattr(sym, "Name", None)
        if nm:
            return unicode(nm)
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import BuiltInParameter, StorageType
        for bip_name in (u"SYMBOL_NAME_PARAM", u"ALL_MODEL_TYPE_NAME"):
            bip = getattr(BuiltInParameter, bip_name, None)
            if bip is None:
                continue
            p = sym.get_Parameter(bip)
            if p is None or not p.HasValue:
                continue
            if p.StorageType != StorageType.String:
                continue
            raw = p.AsString()
            if raw:
                return unicode(raw)
    except Exception:
        pass
    return u""


def _activar_symbol(sym):
    if sym is None:
        return None
    try:
        if not sym.IsActive:
            sym.Activate()
    except Exception:
        pass
    return sym


def resolve_wall_espesor_tag_symbol(doc):
    """
    ``FamilySymbol`` de ``EST_A_WALL TAG_ELEVACION_MHA`` / ``Espesor Muro``.

    :returns: ``(symbol, err_msg)``.
    """
    if doc is None:
        return None, u"documento nulo"
    want_fam = _norm_label(WALL_TAG_FAMILY_NAME)
    want_typ = _norm_label(WALL_TAG_TYPE_NAME)
    candidates = []

    def _scan(col):
        for sym in col or []:
            if sym is None or not isinstance(sym, FamilySymbol):
                continue
            if _norm_label(_symbol_family_name(sym)) != want_fam:
                continue
            candidates.append(sym)

    try:
        col = (
            FilteredElementCollector(doc)
            .OfClass(FamilySymbol)
            .OfCategory(BuiltInCategory.OST_WallTags)
        )
        _scan(col)
    except Exception:
        pass
    if not candidates:
        try:
            _scan(FilteredElementCollector(doc).OfClass(FamilySymbol))
        except Exception:
            pass
    if not candidates:
        return None, (
            u"familia «{0}» no encontrada (categoría Wall Tags)."
            .format(WALL_TAG_FAMILY_NAME)
        )
    exact = []
    fuzzy = []
    for sym in candidates:
        tn = _norm_label(_symbol_type_name(sym))
        if tn == want_typ:
            exact.append(sym)
        elif want_typ and want_typ in tn:
            fuzzy.append(sym)
    pick = None
    if exact:
        pick = exact[0]
    elif fuzzy:
        pick = fuzzy[0]
    if pick is None:
        return None, (
            u"familia «{0}» sin tipo «{1}»."
            .format(WALL_TAG_FAMILY_NAME, WALL_TAG_TYPE_NAME)
        )
    return _activar_symbol(pick), None


def _default_snap_espesor(type_id):
    return {
        u"type_id": type_id,
        u"has_leader": False,
        u"orientation": TagOrientation.Horizontal,
    }


def _iter_net_collection(coll):
    """Itera ICollection / ISet de .NET (IronPython a veces no hace ``for x in set``)."""
    if coll is None:
        return
    try:
        for item in coll:
            yield item
        return
    except Exception:
        pass
    try:
        n = int(coll.Count)
        for i in range(n):
            try:
                yield coll[i]
            except Exception:
                try:
                    yield coll.get_Item(i)
                except Exception:
                    pass
        return
    except Exception:
        pass
    try:
        en = coll.GetEnumerator()
        while en.MoveNext():
            yield en.Current
    except Exception:
        pass


def _as_element_id(obj):
    """Normaliza ``ElementId`` / ``LinkElementId`` → ``ElementId`` local si hay."""
    if obj is None:
        return None
    try:
        if isinstance(obj, ElementId):
            if obj == ElementId.InvalidElementId:
                return None
            return obj
    except Exception:
        pass
    # LinkElementId u objetos duck-typed
    for attr in (u"HostElementId", u"LinkedElementId", u"ElementId"):
        try:
            eid = getattr(obj, attr, None)
        except Exception:
            eid = None
        if eid is None:
            continue
        try:
            if eid != ElementId.InvalidElementId:
                # Puede no pasar isinstance en algunos wrappers
                k = _eid_int(eid)
                if k is not None:
                    return eid if isinstance(eid, ElementId) else ElementId(int(k))
        except Exception:
            try:
                k = _eid_int(eid)
                if k is not None:
                    return ElementId(int(k))
            except Exception:
                pass
    # Entero crudo
    k = _eid_int(obj)
    if k is not None:
        try:
            return ElementId(int(k))
        except Exception:
            pass
    return None


def _is_wall_tag_category(tag):
    try:
        cat = tag.Category
        if cat is None:
            return False
        return int(cat.Id.IntegerValue) == int(BuiltInCategory.OST_WallTags)
    except Exception:
        return False


def _tag_in_view(tag, view_id, view=None):
    if tag is None or not isinstance(tag, IndependentTag):
        return False
    if view_id is None:
        return True
    try:
        oid = tag.OwnerViewId
    except Exception:
        return True
    try:
        if oid == view_id:
            return True
    except Exception:
        pass
    # Vista dependiente ↔ primaria: la etiqueta puede vivir en la otra
    if view is not None:
        try:
            prim = view.GetPrimaryViewId()
            if prim is not None and prim != ElementId.InvalidElementId:
                if oid == prim or view_id == prim:
                    return True
        except Exception:
            pass
    return False


def _tagged_ids_of_independent_tag(tag):
    """ElementIds etiquetados por un IndependentTag (API variable entre versiones)."""
    out = []
    seen = set()
    if tag is None:
        return out

    def _push(eid):
        eid = _as_element_id(eid)
        if eid is None:
            return
        k = _eid_int(eid)
        if k is None or k in seen:
            return
        seen.add(k)
        out.append(eid)

    # Legacy singular
    for attr in (u"TaggedLocalElementId", u"TaggedElementId"):
        try:
            _push(getattr(tag, attr, None))
        except Exception:
            pass

    # Elementos locales (API reciente)
    try:
        els = tag.GetTaggedLocalElements()
        for el in _iter_net_collection(els):
            if el is None:
                continue
            try:
                _push(el.Id)
            except Exception:
                pass
    except Exception:
        pass

    for getter in (
        lambda: tag.GetTaggedLocalElementIds(),
        lambda: tag.GetTaggedElementIds(),
    ):
        try:
            ids = getter()
        except Exception:
            ids = None
        if ids is None:
            continue
        for eid in _iter_net_collection(ids):
            _push(eid)

    try:
        refs = tag.GetTaggedReferences()
    except Exception:
        refs = None
    if refs is not None:
        for ref in _iter_net_collection(refs):
            if ref is None:
                continue
            try:
                _push(ref.ElementId)
            except Exception:
                pass
    return out


def _collect_independent_tags_in_view(doc, view, category):
    """``IndependentTag`` de ``category`` en la vista (OwnerViewId)."""
    out = []
    seen = set()
    if doc is None or view is None or category is None:
        return out
    try:
        view_id = view.Id
    except Exception:
        return out

    def _add(tag):
        if tag is None or not isinstance(tag, IndependentTag):
            return
        if not _tag_in_view(tag, view_id, view=view):
            return
        try:
            cat = tag.Category
            if cat is None or int(cat.Id.IntegerValue) != int(category):
                return
        except Exception:
            return
        try:
            k = _eid_int(tag.Id)
        except Exception:
            k = None
        if k is not None and k in seen:
            return
        if k is not None:
            seen.add(k)
        out.append(tag)

    # 1) Collector acotado a la vista
    try:
        col = (
            FilteredElementCollector(doc, view_id)
            .OfClass(IndependentTag)
            .OfCategory(category)
            .WhereElementIsNotElementType()
        )
        for tag in col:
            _add(tag)
    except Exception:
        pass

    # 1b) También tags del primario si la activa es dependiente
    try:
        prim = view.GetPrimaryViewId()
        if prim is not None and prim != ElementId.InvalidElementId and prim != view_id:
            col = (
                FilteredElementCollector(doc, prim)
                .OfClass(IndependentTag)
                .OfCategory(category)
                .WhereElementIsNotElementType()
            )
            for tag in col:
                _add(tag)
    except Exception:
        pass

    # 2) Documento + OwnerView / ElementOwnerViewFilter
    if not out:
        try:
            from Autodesk.Revit.DB import ElementOwnerViewFilter
            col = (
                FilteredElementCollector(doc)
                .OfClass(IndependentTag)
                .OfCategory(category)
                .WherePasses(ElementOwnerViewFilter(view_id))
                .WhereElementIsNotElementType()
            )
            for tag in col:
                _add(tag)
        except Exception:
            try:
                col = (
                    FilteredElementCollector(doc)
                    .OfClass(IndependentTag)
                    .WhereElementIsNotElementType()
                )
                for tag in col:
                    _add(tag)
            except Exception:
                pass
    return out


def collect_wall_tags_for_wall(doc, view, wall):
    """
    Etiquetas ``OST_WallTags`` en ``view`` que etiquetan ``wall``.

    Retorna lista de ``(tag, tagged_ids, is_multihost)``.
    Prioridad: ``GetDependentElements`` del muro (fiable para mono-host).
    """
    hits = []
    seen_tag = set()
    wid = _wall_id_int(wall)
    if wid is None or doc is None or view is None or wall is None:
        return hits
    try:
        view_id = view.Id
    except Exception:
        view_id = None
    try:
        wall_eid = wall.Id
    except Exception:
        wall_eid = None

    def _append_hit(tag, tagged, is_mh):
        try:
            tid = _eid_int(tag.Id)
        except Exception:
            tid = None
        if tid is not None and tid in seen_tag:
            return
        if tid is not None:
            seen_tag.add(tid)
        hits.append((tag, tagged, is_mh))

    # A) Dependientes del muro (recomendado Autodesk; suele omitir multi-host)
    try:
        from Autodesk.Revit.DB import ElementClassFilter
        dep = wall.GetDependentElements(ElementClassFilter(IndependentTag))
        for eid in _iter_net_collection(dep):
            try:
                tag = doc.GetElement(eid)
            except Exception:
                continue
            if not isinstance(tag, IndependentTag):
                continue
            if not _is_wall_tag_category(tag):
                continue
            if not _tag_in_view(tag, view_id, view=view):
                continue
            tagged = _tagged_ids_of_independent_tag(tag)
            # Dependiente del muro → tratable como mono-host a efectos de borrado
            _append_hit(tag, tagged, False)
    except Exception:
        pass

    # B) Collector en vista + ids etiquetados
    for tag in _collect_independent_tags_in_view(
        doc, view, BuiltInCategory.OST_WallTags,
    ):
        tagged = _tagged_ids_of_independent_tag(tag)
        tagged_ints = []
        hit_wall = False
        for eid in tagged:
            k = _eid_int(eid)
            if k is not None:
                tagged_ints.append(k)
            if k == wid:
                hit_wall = True
            elif wall_eid is not None:
                try:
                    if eid == wall_eid:
                        hit_wall = True
                        if k is None:
                            tagged_ints.append(wid)
                except Exception:
                    pass
        if not hit_wall:
            continue
        is_mh = len(set(tagged_ints)) > 1
        _append_hit(tag, tagged, is_mh)

    return hits


def delete_wall_tags_for_wall(doc, view, wall, errores=None):
    """
    Elimina etiquetas de muro del muro en la vista (mono-host / dependientes).

    Multi-host detectado por ids: no se borra (aviso). Retorna
    ``(n_deleted, n_skip_multihost, style_snap_or_None)``.
    """
    n_del = 0
    n_mh = 0
    style = None
    hits = collect_wall_tags_for_wall(doc, view, wall)
    wid = _wall_id_int(wall)
    for tag, _tagged, is_mh in hits:
        if is_mh:
            n_mh += 1
            msg = (
                u"Etiqueta muro id {0}: multi-host — no eliminada."
                .format(wid)
            )
            if errores is not None:
                errores.append(msg)
            continue
        if style is None:
            style = _snapshot_wall_tag(tag)
        try:
            tag_id = tag.Id
        except Exception:
            tag_id = None
        if tag_id is None:
            continue
        # Desbloquear si está pineada
        try:
            if getattr(tag, "Pinned", False):
                tag.Pinned = False
        except Exception:
            pass
        try:
            doc.Delete(tag_id)
            n_del += 1
        except Exception as ex_del:
            try:
                msg = u"Etiqueta muro id {0}: delete — {1}".format(
                    wid, unicode(ex_del),
                )
            except Exception:
                msg = u"Etiqueta muro id {0}: delete falló.".format(wid)
            if errores is not None:
                errores.append(msg)
    return n_del, n_mh, style


def _rebar_ids_int_set(rebars_por_muro_id, wall_id):
    out = set()
    if not rebars_por_muro_id or wall_id is None:
        return out
    for key in (wall_id, int(wall_id) if wall_id is not None else None, str(wall_id)):
        if key is None:
            continue
        lst = rebars_por_muro_id.get(key)
        if lst is None:
            continue
        for eid in lst or []:
            k = _eid_int(eid) if not isinstance(eid, int) else int(eid)
            if k is None:
                try:
                    k = int(eid)
                except Exception:
                    k = None
            if k is not None:
                out.add(k)
        if out:
            break
    return out


def _mesh_tag_heads_for_wall(doc, view, wall, rebars_por_muro_id):
    """
    ``TagHeadPosition`` de etiquetas de malla (OST_RebarTags) asociadas a las
    rebars del muro. Si no hay, lista vacía.
    """
    heads = []
    wid = _wall_id_int(wall)
    want = _rebar_ids_int_set(rebars_por_muro_id, wid)
    if not want:
        return heads
    for tag in _collect_independent_tags_in_view(
        doc, view, BuiltInCategory.OST_RebarTags,
    ):
        tagged = _tagged_ids_of_independent_tag(tag)
        hit = False
        for eid in tagged:
            k = _eid_int(eid)
            if k is not None and k in want:
                hit = True
                break
        if not hit:
            continue
        try:
            hp = tag.TagHeadPosition
        except Exception:
            hp = None
        if hp is not None:
            heads.append(hp)
    return heads


def _view_up(view):
    if view is None:
        return None
    try:
        up = view.UpDirection
        if up is None:
            return None
        n = up.Normalize()
        if n is None:
            return None
        return n
    except Exception:
        return None


def _paper_offset_mm(n_mesh_heads=0):
    """
    Offset en mm de papel según nº de cabezas de malla.

    Referencia: 2.5 mm texto @ 1:50; doble malla → 1.5 × 2.5 = 3.75 mm papel.
    """
    try:
        n = int(n_mesh_heads or 0)
    except Exception:
        n = 0
    if n == 1:
        heights = _WALL_TAG_OFFSET_TEXT_HEIGHTS_SINGLE
    else:
        # 0 (fallback), 2+ → bloque V+H como en la imagen
        heights = _WALL_TAG_OFFSET_TEXT_HEIGHTS
    return float(_WALL_TAG_TEXT_HEIGHT_PAPER_MM) * float(heights)


def _delta_above_mesh_ft(view, n_mesh_heads=0):
    """
    Offset modelo (pies) en ``View.UpDirection``.

    ``modelo_mm = papel_mm × Scale`` (texto fijo en papel).
    Ejemplos con doble malla (3.75 mm papel): 1:50 → 187.5 mm; 1:75 → 281 mm; 1:100 → 375 mm.
    """
    paper_mm = _paper_offset_mm(n_mesh_heads)
    try:
        from armado_muros_cabezal_tags import (
            _mm_to_internal,
            _view_scale_denominator,
        )
        sd = _view_scale_denominator(view)
        if sd is None or int(sd) <= 0:
            sd = _WALL_TAG_SCALE_REFERENCE
        return _mm_to_internal(paper_mm * float(sd))
    except Exception:
        # Fallback 1:50
        return (paper_mm * float(_WALL_TAG_SCALE_REFERENCE)) / 304.8


def _head_above_mesh(wall, view, mesh_heads):
    """
    Punto de cabeza: máximo mesh head proyectado en Up + delta calibrado.
    Fallback: centroide de malla + delta.
    """
    up = _view_up(view)
    if up is None:
        return None
    n_heads = 0
    try:
        n_heads = len(mesh_heads or [])
    except Exception:
        n_heads = 0
    delta = _delta_above_mesh_ft(view, n_mesh_heads=n_heads)
    base = None
    if mesh_heads:
        best = None
        best_proj = None
        for hp in mesh_heads:
            if hp is None:
                continue
            try:
                proj = hp.X * up.X + hp.Y * up.Y + hp.Z * up.Z
            except Exception:
                continue
            if best is None or proj > best_proj:
                best = hp
                best_proj = proj
        base = best
    if base is None:
        try:
            from armado_muros_malla_rebar_tags import _head_pos_centroide_muro
            base = _head_pos_centroide_muro(wall, view)
        except Exception:
            base = None
        # Sin cabezas leídas: asumir doble malla (V+H) como en la imagen
        if n_heads < 1:
            delta = _delta_above_mesh_ft(view, n_mesh_heads=2)
    if base is None:
        return None
    try:
        return XYZ(
            base.X + up.X * delta,
            base.Y + up.Y * delta,
            base.Z + up.Z * delta,
        )
    except Exception:
        return None


def _snapshot_wall_tag(tag):
    """Dict con type_id, has_leader, orientation."""
    snap = {
        u"type_id": None,
        u"has_leader": False,
        u"orientation": TagOrientation.Horizontal,
    }
    if tag is None:
        return snap
    try:
        snap[u"type_id"] = tag.GetTypeId()
    except Exception:
        snap[u"type_id"] = None
    try:
        snap[u"has_leader"] = bool(tag.HasLeader)
    except Exception:
        snap[u"has_leader"] = False
    try:
        snap[u"orientation"] = tag.TagOrientation
    except Exception:
        snap[u"orientation"] = TagOrientation.Horizontal
    return snap


def _crear_wall_tag(doc, view, wall, snap, head_pos):
    """Crea IndependentTag de muro; retorna (tag, err)."""
    if doc is None or view is None or wall is None or head_pos is None:
        return None, u"parámetros inválidos"
    type_id = snap.get(u"type_id") if snap else None
    if type_id is None or type_id == ElementId.InvalidElementId:
        return None, u"sin tipo de etiqueta de muro"
    try:
        ref = Reference(wall)
    except Exception as ex:
        try:
            return None, unicode(ex)
        except Exception:
            return None, str(ex)
    has_leader = bool(snap.get(u"has_leader")) if snap else False
    orient = snap.get(u"orientation") if snap else TagOrientation.Horizontal
    if orient is None:
        orient = TagOrientation.Horizontal
    last_ex = None

    def _finish(tag):
        if tag is None:
            return None
        try:
            tag.ChangeTypeId(type_id)
        except Exception:
            try:
                tag.SetTypeId(type_id)
            except Exception:
                pass
        try:
            tag.TagHeadPosition = head_pos
        except Exception:
            pass
        try:
            tag.HasLeader = has_leader
        except Exception:
            pass
        try:
            tag.TagOrientation = orient
        except Exception:
            pass
        return tag

    try:
        tag = IndependentTag.Create(
            doc,
            type_id,
            view.Id,
            ref,
            has_leader,
            orient,
            head_pos,
        )
        tag = _finish(tag)
        if tag is not None:
            return tag, None
    except Exception as ex:
        last_ex = ex
    try:
        tag = IndependentTag.Create(
            doc,
            view.Id,
            ref,
            has_leader,
            TagMode.TM_ADDBY_CATEGORY,
            orient,
            head_pos,
        )
        tag = _finish(tag)
        if tag is not None:
            return tag, None
    except Exception as ex:
        last_ex = ex
    if last_ex is not None:
        try:
            return None, unicode(last_ex)
        except Exception:
            return None, str(last_ex)
    return None, u"no se pudo crear IndependentTag de muro"


def rebase_wall_tags_above_mesh(
    doc,
    uidoc,
    walls,
    rebars_por_muro_id=None,
    errores=None,
):
    """
    En la vista activa, por cada muro seleccionado:

    1. Comprueba si ya tiene ``OST_WallTags``.
    2. Si tiene (mono-host): las elimina.
    3. Coloca ``EST_A_WALL TAG_ELEVACION_MHA`` / ``Espesor Muro`` por encima
       de las etiquetas de malla.

    Multi-host: no se elimina (aviso). Debe llamarse dentro de la txn padre.
    """
    result = _resultado_vacio()
    result[u"n_deleted"] = 0
    if doc is None or uidoc is None or not walls:
        return result
    view = None
    try:
        view = uidoc.ActiveView
    except Exception:
        view = None
    if view is None:
        return result
    try:
        if getattr(view, "IsTemplate", False):
            return result
    except Exception:
        pass

    sym, err_sym = resolve_wall_espesor_tag_symbol(doc)
    if sym is None:
        result[u"n_fail"] = int(result[u"n_fail"]) + 1
        msg = u"Etiqueta espesor muro: {0}".format(err_sym or u"?")
        _append_msg(result, msg)
        if errores is not None:
            errores.append(msg)
        return result
    try:
        type_id = sym.Id
    except Exception:
        type_id = None
    if type_id is None or type_id == ElementId.InvalidElementId:
        msg = u"Etiqueta espesor muro: Id de tipo inválido."
        result[u"n_fail"] = int(result[u"n_fail"]) + 1
        _append_msg(result, msg)
        if errores is not None:
            errores.append(msg)
        return result

    snap_base = _default_snap_espesor(type_id)

    for wall in walls:
        if wall is None:
            continue
        wid = _wall_id_int(wall)

        mesh_heads = _mesh_tag_heads_for_wall(
            doc, view, wall, rebars_por_muro_id,
        )
        new_head = _head_above_mesh(wall, view, mesh_heads)
        if new_head is None:
            result[u"n_fail"] = int(result[u"n_fail"]) + 1
            msg = u"Etiqueta muro id {0}: sin punto de inserción.".format(wid)
            _append_msg(result, msg)
            if errores is not None:
                errores.append(msg)
            continue

        # 1–2: detectar y eliminar etiquetas existentes en esta vista
        n_del, n_mh, style = delete_wall_tags_for_wall(
            doc, view, wall, errores=errores,
        )
        result[u"n_deleted"] = int(result.get(u"n_deleted", 0)) + int(n_del)
        if n_mh:
            result[u"n_skip_multihost"] = (
                int(result.get(u"n_skip_multihost", 0)) + int(n_mh)
            )
            for _i in range(int(n_mh)):
                _append_msg(
                    result,
                    u"Etiqueta muro id {0}: multi-host — no eliminada.".format(wid),
                )

        snap = dict(snap_base)
        if style:
            if style.get(u"has_leader") is not None:
                snap[u"has_leader"] = bool(style.get(u"has_leader"))
            if style.get(u"orientation") is not None:
                snap[u"orientation"] = style.get(u"orientation")

        # 3: nueva etiqueta Espesor Muro en posición sobre malla
        new_tag, err_c = _crear_wall_tag(doc, view, wall, snap, new_head)
        if new_tag is None:
            result[u"n_fail"] = int(result[u"n_fail"]) + 1
            msg = u"Etiqueta muro id {0}: {1} — {2}".format(
                wid,
                u"recreate" if n_del else u"create",
                err_c or u"?",
            )
            _append_msg(result, msg)
            if errores is not None:
                errores.append(msg)
            continue
        result[u"n_ok"] = int(result[u"n_ok"]) + 1

    return result


def resumen_para_embed(wall_res):
    """Claves para fusionar en ``embed_res`` / mensaje OK."""
    wall_res = wall_res or {}
    return {
        u"n_wall_tags_rebase": int(wall_res.get(u"n_ok", 0) or 0),
        u"n_wall_tags_rebase_deleted": int(wall_res.get(u"n_deleted", 0) or 0),
        u"n_wall_tags_rebase_skip_multihost": int(
            wall_res.get(u"n_skip_multihost", 0) or 0,
        ),
        u"n_wall_tags_rebase_fail": int(wall_res.get(u"n_fail", 0) or 0),
        u"messages": list(wall_res.get(u"messages") or []),
    }
