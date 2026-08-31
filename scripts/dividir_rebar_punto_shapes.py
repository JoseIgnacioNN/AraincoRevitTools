# -*- coding: utf-8 -*-
"""
Asignación de RebarShape tras dividir barras.

Reglas (ampliables): nombre de shape original → extremos / intermedios.

Caso 03 (N tramos resultantes, N ≥ 2):
  - Primera y última → Shape «02»
  - Tramos intermedios (si N ≥ 3) → Shape «01»
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("System")

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BuiltInParameter,
    Curve,
    ElementId,
    FilteredElementCollector,
    StorageType,
)
from Autodesk.Revit.DB.Structure import (
    Rebar,
    RebarHookOrientation,
    RebarShape,
)

# original_key -> {ends, middle}
# ends = primera y última barra; middle = tramos entre ellas (si hay ≥ 3).
SPLIT_SHAPE_RULES = {
    u"03": {u"ends": u"02", u"middle": u"01"},
}


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def normalize_shape_key(name):
    """Normaliza «03», «3», «Shape 03», «M_03» → «03» si hay dígitos."""
    s = _as_unicode(name).strip()
    if not s:
        return u""
    low = s.lower()
    for pref in (u"shape ", u"forma ", u"form ", u"rebar shape "):
        if low.startswith(pref):
            s = s[len(pref) :].strip()
            low = s.lower()
            break
    digits = u"".join(ch for ch in s if ch in u"0123456789")
    if digits:
        # Si hay varios grupos, preferir los últimos 1–2 dígitos tipicos (01..99)
        if len(digits) > 2:
            digits = digits[-2:]
        try:
            n = int(digits)
            return u"{0:02d}".format(n)
        except Exception:
            return digits.zfill(2)
    return s


def rebar_shape_display_name(sh):
    """
    Nombre visible del ``RebarShape`` en el Navegador de proyectos.

    ``Element.Name`` suele estar vacío; usar ``SYMBOL_NAME_PARAM`` /
    ``ALL_MODEL_TYPE_NAME`` (mismo criterio que fundación / estribos).
    """
    if sh is None:
        return u""
    try:
        n = _as_unicode(getattr(sh, u"Name", None)).strip()
        if n:
            return n
    except Exception:
        pass
    for bip_name in (u"SYMBOL_NAME_PARAM", u"ALL_MODEL_TYPE_NAME"):
        try:
            bip = getattr(BuiltInParameter, bip_name, None)
            if bip is None:
                continue
            p = sh.get_Parameter(bip)
            if p is None or not p.HasValue:
                continue
            if p.StorageType == StorageType.String:
                s = _as_unicode(p.AsString()).strip()
                if s:
                    return s
            else:
                s = _as_unicode(p.AsValueString()).strip()
                if s:
                    return s
        except Exception:
            continue
    return u""


def list_rebar_shape_labels(doc, limit=12):
    """Etiquetas visibles de shapes en el documento (diagnóstico)."""
    labels = []
    if doc is None:
        return labels
    try:
        for sh in FilteredElementCollector(doc).OfClass(RebarShape):
            sn = rebar_shape_display_name(sh)
            if sn and sn not in labels:
                labels.append(sn)
            if len(labels) >= int(limit):
                break
    except Exception:
        pass
    return labels


def get_rebar_shape_name(doc, rebar):
    if doc is None or rebar is None:
        return u""
    sid = None
    try:
        sid = rebar.GetShapeId()
    except Exception:
        try:
            sid = rebar.RebarShapeId
        except Exception:
            sid = None
    if sid is not None:
        try:
            if sid != ElementId.InvalidElementId:
                sh = doc.GetElement(sid)
                if sh is not None:
                    nm = rebar_shape_display_name(sh)
                    if nm:
                        return nm
        except Exception:
            pass
    # Fallback parámetro de instancia
    for bip_name in (u"REBAR_SHAPE", u"REBAR_SHAPE_PARAM"):
        bip = getattr(BuiltInParameter, bip_name, None)
        if bip is None:
            continue
        try:
            p = rebar.get_Parameter(bip)
            if p is not None and p.HasValue:
                v = p.AsValueString() or p.AsString()
                if v:
                    return _as_unicode(v).strip()
        except Exception:
            pass
    try:
        p = rebar.LookupParameter(u"Shape")
        if p is not None and p.HasValue:
            v = p.AsValueString() or p.AsString()
            if v:
                return _as_unicode(v).strip()
    except Exception:
        pass
    return u""


def find_rebar_shape_by_name(doc, name):
    """
    Busca ``RebarShape`` por nombre mostrado (exacto / sin mayúsculas /
    clave normalizada / solo dígitos).
    """
    if doc is None:
        return None
    want = _as_unicode(name).strip()
    key = normalize_shape_key(name)
    if not want and not key:
        return None
    try:
        want_lower = want.lower()
    except Exception:
        want_lower = want
    match_casefold = None
    match_norm = None
    match_digits = None
    try:
        shapes = list(FilteredElementCollector(doc).OfClass(RebarShape))
    except Exception:
        return None
    for sh in shapes:
        if sh is None:
            continue
        sn = rebar_shape_display_name(sh)
        if not sn:
            continue
        if sn == want or sn == key:
            return sh
        try:
            sn_low = sn.lower()
        except Exception:
            sn_low = sn
        if match_casefold is None and want_lower and sn_low == want_lower:
            match_casefold = sh
        nk = normalize_shape_key(sn)
        if key and nk == key:
            if match_norm is None:
                match_norm = sh
            # Preferir etiqueta corta «02» sobre nombres largos con mismos dígitos
            if sn == key or sn.strip() == key:
                return sh
        if key and match_digits is None and nk == key:
            match_digits = sh
    return match_casefold or match_norm or match_digits


def resolve_split_shape_rule(original_shape_name):
    """
    Devuelve ``{ends, middle}`` según ``SPLIT_SHAPE_RULES``, o ``None``.

    Acepta reglas legacy ``(A, B)`` interpretadas como ends=A, middle=B.
    """
    key = normalize_shape_key(original_shape_name)
    if not key:
        return None
    rule = SPLIT_SHAPE_RULES.get(key)
    if rule is None:
        for k, v in SPLIT_SHAPE_RULES.items():
            if normalize_shape_key(k) == key:
                rule = v
                break
    if rule is None:
        return None
    if isinstance(rule, dict):
        ends = _as_unicode(rule.get(u"ends") or u"").strip()
        middle = _as_unicode(rule.get(u"middle") or ends).strip()
    elif isinstance(rule, (tuple, list)) and len(rule) >= 2:
        ends = _as_unicode(rule[0]).strip()
        middle = _as_unicode(rule[1]).strip() or ends
    else:
        return None
    if not ends:
        return None
    if not middle:
        middle = ends
    return {u"ends": ends, u"middle": middle}


def target_shapes_for_split(original_shape_name):
    """
    Compatibilidad: ``(ends, middle)`` o ``None``.

    Preferir :func:`target_shape_names_for_pieces` para N tramos.
    """
    rule = resolve_split_shape_rule(original_shape_name)
    if rule is None:
        return None
    return (rule[u"ends"], rule[u"middle"])


def target_shape_names_for_pieces(original_shape_name, n_pieces):
    """
    Lista de nombres de shape para ``n_pieces`` tramos (orden inicio→fin).

    Regla 03: [02, 02] | [02, 01, 02] | [02, 01, …, 01, 02]
    """
    try:
        n = int(n_pieces)
    except Exception:
        return None
    if n < 1:
        return None
    rule = resolve_split_shape_rule(original_shape_name)
    if rule is None:
        return None
    ends = rule[u"ends"]
    middle = rule[u"middle"]
    if n == 1:
        return [ends]
    names = []
    for i in range(n):
        if i == 0 or i == n - 1:
            names.append(ends)
        else:
            names.append(middle)
    return names


def set_rebar_shape(doc, rebar, shape_or_name):
    """
    Fuerza ``RebarShape`` y verifica el nombre resultante.

    Returns:
        (ok, mensaje, nombre_final)
    """
    if doc is None or rebar is None or shape_or_name is None:
        return False, u"Parámetros incompletos.", u""
    shape = shape_or_name
    target_name = u""
    if isinstance(shape, RebarShape):
        target_name = rebar_shape_display_name(shape)
    else:
        target_name = _as_unicode(shape_or_name).strip()
        shape = find_rebar_shape_by_name(doc, shape_or_name)
    if shape is None:
        avail = list_rebar_shape_labels(doc, limit=10)
        hint = u""
        if avail:
            hint = u" Disponibles: {0}.".format(u", ".join(avail))
        return (
            False,
            u"No se encontró RebarShape «{0}» en el documento.{1}".format(
                target_name, hint
            ),
            u"",
        )
    if not target_name:
        target_name = rebar_shape_display_name(shape)
    target_key = normalize_shape_key(target_name)

    last_err = u""
    # 1) ShapeDrivenAccessor.SetRebarShapeId
    try:
        acc = rebar.GetShapeDrivenAccessor()
        if acc is not None:
            acc.SetRebarShapeId(shape.Id)
    except Exception as ex:
        last_err = _as_unicode(ex)
        try:
            acc = rebar.GetShapeDrivenAccessor()
            fn = getattr(acc, u"SetRebarShapeId", None)
            if fn is not None:
                fn(shape.Id)
                last_err = u""
        except Exception as ex2:
            last_err = _as_unicode(ex2)

    try:
        doc.Regenerate()
    except Exception:
        pass

    final = get_rebar_shape_name(doc, rebar)
    final_key = normalize_shape_key(final)
    if final_key and target_key and final_key == target_key:
        return True, u"", final
    if final and target_name and final.strip() == target_name.strip():
        return True, u"", final

    # 2) Reintento: a veces hace falta un segundo Set tras Regenerate
    try:
        rebar.GetShapeDrivenAccessor().SetRebarShapeId(shape.Id)
        doc.Regenerate()
    except Exception as ex:
        if not last_err:
            last_err = _as_unicode(ex)

    final = get_rebar_shape_name(doc, rebar)
    final_key = normalize_shape_key(final)
    if final_key and target_key and final_key == target_key:
        return True, u"", final

    detail = last_err or u"sin detalle API"
    return (
        False,
        u"No se pudo fijar Shape «{0}» (quedó «{1}»). {2}".format(
            target_name or target_key, final or u"?", detail
        ),
        final,
    )


def create_rebar_with_shape(
    doc,
    rebar_shape,
    bar_type,
    start_hook,
    end_hook,
    host,
    norm,
    curves_list,
    start_orient,
    end_orient,
):
    """
    Intenta ``CreateFromCurvesAndShape``; si falla, ``None``.
    """
    if doc is None or rebar_shape is None or bar_type is None or host is None:
        return None
    if not curves_list:
        return None
    curves_clean = []
    for c in curves_list:
        if c is None:
            continue
        try:
            curves_clean.append(c.Clone())
        except Exception:
            curves_clean.append(c)
    if not curves_clean:
        return None

    lst = List[Curve]()
    for c in curves_clean:
        lst.Add(c)

    norms = []
    if norm is not None:
        try:
            norms.append(norm.Normalize())
        except Exception:
            norms.append(norm)
        try:
            norms.append(norm.Normalize().Negate())
        except Exception:
            pass
    if not norms:
        from Autodesk.Revit.DB import XYZ

        norms = [XYZ.BasisZ]

    orients = (
        (start_orient, end_orient),
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
    )
    hook_pairs = ((start_hook, end_hook), (None, None))
    invalid = ElementId.InvalidElementId

    for h0, h1 in hook_pairs:
        for so, eo in orients:
            for nvec in norms:
                # Overload con end treatments
                try:
                    rb = Rebar.CreateFromCurvesAndShape(
                        doc,
                        rebar_shape,
                        bar_type,
                        h0,
                        h1,
                        host,
                        nvec,
                        lst,
                        so,
                        eo,
                        0.0,
                        0.0,
                        invalid,
                        invalid,
                    )
                    if rb is not None:
                        return rb
                except Exception:
                    pass
                try:
                    rb = Rebar.CreateFromCurvesAndShape(
                        doc,
                        rebar_shape,
                        bar_type,
                        h0,
                        h1,
                        host,
                        nvec,
                        lst,
                        so,
                        eo,
                    )
                    if rb is not None:
                        return rb
                except Exception:
                    pass
    return None


def apply_split_shape_rules(doc, original_rebar, rebar_a, rebar_b):
    """
    Aplica reglas de shape a los tramos A y B según la barra original.

    Returns:
        dict: applied, original, target_a, target_b, ok_a, ok_b, final_a, final_b, errors
    """
    return apply_split_shape_rules_to_list(doc, original_rebar, [rebar_a, rebar_b])


def apply_split_shape_rules_to_list(doc, original_rebar, rebars):
    """
    Aplica la regla de shape a N tramos.

    Para shape «03»: extremos → «02», intermedios → «01».
    """
    result = {
        u"applied": False,
        u"original": u"",
        u"target_a": u"",
        u"target_b": u"",
        u"targets": [],
        u"ok_a": False,
        u"ok_b": False,
        u"final_a": u"",
        u"final_b": u"",
        u"ok_all": False,
        u"finals": [],
        u"errors": [],
    }
    rebars = [rb for rb in (rebars or []) if rb is not None]
    if not rebars:
        result[u"errors"].append(u"Sin tramos para asignar shape.")
        return result
    orig_name = get_rebar_shape_name(doc, original_rebar)
    result[u"original"] = orig_name
    names = target_shape_names_for_pieces(orig_name, len(rebars))
    if not names:
        result[u"errors"].append(
            u"Sin regla de shape para original «{0}» (clave {1}).".format(
                orig_name or u"?", normalize_shape_key(orig_name) or u"?"
            )
        )
        return result
    result[u"targets"] = list(names)
    result[u"target_a"] = names[0]
    result[u"target_b"] = names[-1] if len(names) > 1 else names[0]
    result[u"applied"] = True

    finals = []
    oks = []
    for i, rb in enumerate(rebars):
        name = names[i] if i < len(names) else names[-1]
        ok, err, final = set_rebar_shape(doc, rb, name)
        oks.append(bool(ok))
        finals.append(final)
        if not ok and err:
            result[u"errors"].append(u"T{0}→{1}: {2}".format(i + 1, name, err))
    result[u"finals"] = finals
    result[u"ok_all"] = all(oks) if oks else False
    result[u"ok_a"] = oks[0] if oks else False
    result[u"ok_b"] = oks[1] if len(oks) > 1 else (oks[0] if oks else False)
    result[u"final_a"] = finals[0] if finals else u""
    result[u"final_b"] = finals[1] if len(finals) > 1 else (finals[0] if finals else u"")
    return result
