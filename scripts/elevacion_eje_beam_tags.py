# -*- coding: utf-8 -*-
"""
Etiquetas de viga en Elevación Eje.

Por cada ViewSection creada:
  - Vigas (Structural Framing) visibles con Material for Model Behavior = Concrete
  - Paralelas al plano de la vista (eje ⟂ ViewDirection)
  - IndependentTag: EST_A_STRUCTURAL FRAMING TAG_ELEVACION / Tag Viga
  - Cabeza en la cara superior del bbox (máximo en View.UpDirection)
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FamilySymbol,
    FilteredElementCollector,
    IndependentTag,
    Reference,
    StorageType,
    TagMode,
    TagOrientation,
    XYZ,
)

from elevacion_eje_collect import (
    _eid_int,
    _vector_unitario,
    recoger_concrete_en_vista,
    viga_paralela_a_vista,
)

BEAM_TAG_FAMILY_NAME = u"EST_A_STRUCTURAL FRAMING TAG_ELEVACION"
BEAM_TAG_TYPE_NAME = u"Tag Viga"


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _norm_label(s):
    t = _as_unicode(s).strip().lower()
    for ch in (u"\xa0", u"\u200b", u"\ufeff"):
        t = t.replace(ch, u"")
    return u" ".join(t.split())


def _symbol_family_name(sym):
    try:
        fam = sym.Family
        if fam is not None:
            return _as_unicode(fam.Name or u"")
    except Exception:
        pass
    return u""


def _symbol_type_name(sym):
    try:
        nm = getattr(sym, "Name", None)
        if nm:
            return _as_unicode(nm)
    except Exception:
        pass
    for bip_name in (u"SYMBOL_NAME_PARAM", u"ALL_MODEL_TYPE_NAME"):
        try:
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
                return _as_unicode(raw)
        except Exception:
            continue
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


def resolve_beam_tag_symbol(document):
    """
    ``FamilySymbol`` EST_A_STRUCTURAL FRAMING TAG_ELEVACION / Tag Viga.

    Returns:
        (symbol, None) o (None, mensaje_error)
    """
    if document is None:
        return None, u"documento nulo"
    want_fam = _norm_label(BEAM_TAG_FAMILY_NAME)
    want_typ = _norm_label(BEAM_TAG_TYPE_NAME)
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
            FilteredElementCollector(document)
            .OfClass(FamilySymbol)
            .OfCategory(BuiltInCategory.OST_StructuralFramingTags)
        )
        _scan(col)
    except Exception:
        pass
    if not candidates:
        try:
            _scan(FilteredElementCollector(document).OfClass(FamilySymbol))
        except Exception:
            pass
    if not candidates:
        return None, (
            u"familia «{0}» no encontrada (Structural Framing Tags)."
            .format(BEAM_TAG_FAMILY_NAME)
        )
    exact = []
    fuzzy = []
    for sym in candidates:
        tn = _norm_label(_symbol_type_name(sym))
        if tn == want_typ:
            exact.append(sym)
        elif want_typ and want_typ in tn:
            fuzzy.append(sym)
    pick = exact[0] if exact else (fuzzy[0] if fuzzy else None)
    if pick is None:
        return None, (
            u"familia «{0}» sin tipo «{1}»."
            .format(BEAM_TAG_FAMILY_NAME, BEAM_TAG_TYPE_NAME)
        )
    return _activar_symbol(pick), None


def _esquinas_bbox(bb):
    mn, mx = bb.Min, bb.Max
    return (
        XYZ(mn.X, mn.Y, mn.Z),
        XYZ(mx.X, mn.Y, mn.Z),
        XYZ(mn.X, mx.Y, mn.Z),
        XYZ(mx.X, mx.Y, mn.Z),
        XYZ(mn.X, mn.Y, mx.Z),
        XYZ(mx.X, mn.Y, mx.Z),
        XYZ(mn.X, mx.Y, mx.Z),
        XYZ(mx.X, mx.Y, mx.Z),
    )


def _punto_cara_superior_bbox(beam, view):
    """
    Centro de la cara superior del bbox: vértices con mayor proyección
    sobre ``View.UpDirection`` (fallback: +Z mundo).
    """
    if beam is None:
        return None
    bb = None
    try:
        bb = beam.get_BoundingBox(view)
    except Exception:
        bb = None
    if bb is None:
        try:
            bb = beam.get_BoundingBox(None)
        except Exception:
            bb = None
    if bb is None:
        return None

    up = None
    try:
        if view is not None:
            up = _vector_unitario(view.UpDirection)
    except Exception:
        up = None
    if up is None:
        up = XYZ(0.0, 0.0, 1.0)

    corners = _esquinas_bbox(bb)
    try:
        dmax = max(float(p.DotProduct(up)) for p in corners)
    except Exception:
        try:
            return XYZ(
                (bb.Min.X + bb.Max.X) * 0.5,
                (bb.Min.Y + bb.Max.Y) * 0.5,
                bb.Max.Z,
            )
        except Exception:
            return None

    tops = []
    for p in corners:
        try:
            if abs(float(p.DotProduct(up)) - dmax) <= 1e-6:
                tops.append(p)
        except Exception:
            continue
    if not tops:
        return None
    try:
        n = float(len(tops))
        return XYZ(
            sum(p.X for p in tops) / n,
            sum(p.Y for p in tops) / n,
            sum(p.Z for p in tops) / n,
        )
    except Exception:
        return tops[0]


def recoger_vigas_concrete_paralelas(document, view, material_cache=None):
    """Vigas Concrete visibles en ``view`` y paralelas a su plano."""
    packed = recoger_concrete_en_vista(document, view, material_cache)
    return list(packed.get(u"vigas") or [])


def _crear_beam_tag(document, view, beam, type_id, head_pos):
    if document is None or view is None or beam is None or head_pos is None:
        return None, u"parámetros inválidos"
    if type_id is None or type_id == ElementId.InvalidElementId:
        return None, u"sin tipo de etiqueta de viga"
    try:
        ref = Reference(beam)
    except Exception as ex:
        return None, _as_unicode(ex)
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
            tag.HasLeader = False
        except Exception:
            pass
        try:
            tag.TagOrientation = orient
        except Exception:
            pass
        return tag

    try:
        tag = IndependentTag.Create(
            document,
            type_id,
            view.Id,
            ref,
            False,
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
            document,
            view.Id,
            ref,
            False,
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
        return None, _as_unicode(last_ex)
    return None, u"no se pudo crear IndependentTag de viga"


def etiquetar_vigas_concrete_paralelas(
    document,
    view,
    symbol=None,
    vigas=None,
    tagged_hosts=None,
):
    """
    Crea etiquetas Tag Viga en ``view``.

    Debe llamarse dentro de una Transaction abierta (crop ya activo).

    Args:
        vigas: lista precargada (opcional)
        tagged_hosts: set de host Id int ya etiquetados en la vista

    Returns:
        dict ``n_ok``, ``n_skip``, ``n_fail``, ``error_symbol``
    """
    result = {
        u"n_ok": 0,
        u"n_skip": 0,
        u"n_fail": 0,
        u"error_symbol": None,
    }
    if document is None or view is None:
        result[u"n_fail"] = 1
        return result

    if symbol is None:
        symbol, err = resolve_beam_tag_symbol(document)
        if symbol is None:
            result[u"error_symbol"] = err or u"Tipo de etiqueta no encontrado."
            result[u"n_fail"] = 1
            return result
    try:
        type_id = symbol.Id
    except Exception:
        type_id = None
    if type_id is None:
        result[u"error_symbol"] = u"Id de tipo de etiqueta inválido."
        result[u"n_fail"] = 1
        return result

    if vigas is None:
        vigas = recoger_vigas_concrete_paralelas(document, view)
    if not vigas:
        return result

    already = tagged_hosts if tagged_hosts is not None else set()

    for beam in vigas:
        if beam is None:
            continue
        bid = _eid_int(getattr(beam, u"Id", None))
        if bid is not None and bid in already:
            result[u"n_skip"] = int(result[u"n_skip"]) + 1
            continue
        if not viga_paralela_a_vista(beam, view):
            continue
        head = _punto_cara_superior_bbox(beam, view)
        if head is None:
            result[u"n_fail"] = int(result[u"n_fail"]) + 1
            continue
        tag, err = _crear_beam_tag(document, view, beam, type_id, head)
        if tag is None:
            result[u"n_fail"] = int(result[u"n_fail"]) + 1
        else:
            result[u"n_ok"] = int(result[u"n_ok"]) + 1
            if bid is not None:
                already.add(bid)
    return result
