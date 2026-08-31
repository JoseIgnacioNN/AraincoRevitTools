# -*- coding: utf-8 -*-
"""
Ajuste de patas por cambio de diámetro (tabla BIMTools) en shapes «02» y «03».

- **Shape 02**: actualiza el parámetro de forma **A**.
- **Shape 03**: actualiza **A** y **C** (B = tramo largo, sin cambio).

Tablas G25 / G35 / G45 vía :mod:`bimtools_rebar_hook_lengths`.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("System")

import System

from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    StorageType,
    Transaction,
    TransactionStatus,
    UnitTypeId,
    UnitUtils,
)
from Autodesk.Revit.DB.Structure import Rebar, RebarBarType, RebarShape

from bimtools_rebar_hook_lengths import (
    concrete_grade_from_text,
    hook_length_mm_from_nominal_diameter_mm,
    normalize_concrete_grade,
)

_AD_TYPE_CACHE = u"BIMTools.RebarPataLDmuBarTypeCache"
_AD_GRADE_OVERRIDE = u"BIMTools.RebarPataLDmuConcreteGradeOverride"

_SHAPE_SEGMENTS = {
    u"02": (u"A",),
    u"03": (u"A", u"C"),
}
_GRADE_PARAM_NAME_HINTS = (
    u"dosific",
    u"hormigon",
    u"hormig",
    u"concrete",
    u"grado",
    u"resistencia",
    u"grade",
)

_PATA_MATCH_TOL_MM = 3.0
_TXN = u"Arainco: Pata L por cambio diámetro (DMU)"


def _mm_to_internal(mm):
    return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)


def _doc_cache_key(doc):
    try:
        return int(doc.GetHashCode())
    except Exception:
        return id(doc)


def _type_cache_all():
    try:
        data = System.AppDomain.CurrentDomain.GetData(_AD_TYPE_CACHE)
    except Exception:
        data = None
    if data is None:
        data = {}
        try:
            System.AppDomain.CurrentDomain.SetData(_AD_TYPE_CACHE, data)
        except Exception:
            pass
    return data


def _remember_bar_type(doc, rebar_id_int, type_id_int):
    cache = _type_cache_all()
    dk = _doc_cache_key(doc)
    if dk not in cache:
        cache[dk] = {}
    cache[dk][int(rebar_id_int)] = int(type_id_int)


def _read_previous_bar_type(doc, rebar_id_int):
    cache = _type_cache_all()
    dk = _doc_cache_key(doc)
    sub = cache.get(dk) or {}
    return sub.get(int(rebar_id_int))


def peek_bar_type_id(doc, rebar_id_int):
    """Último ``RebarBarType`` Id conocido para la barra (sin modificar caché)."""
    return _read_previous_bar_type(doc, rebar_id_int)


def remember_bar_type_id(doc, rebar_id_int, type_id_int):
    _remember_bar_type(doc, rebar_id_int, type_id_int)


def bar_type_id_int(rebar):
    """``ElementId.IntegerValue`` del ``RebarBarType`` actual, o ``None``."""
    if rebar is None:
        return None
    try:
        tid = rebar.GetTypeId()
    except Exception:
        return None
    if tid is None or tid == ElementId.InvalidElementId:
        return None
    try:
        return int(tid.IntegerValue)
    except Exception:
        return None


def remember_bar_type_snapshot_for_rebar(doc, rebar):
    """Actualiza caché con el ``RebarBarType`` actual (tras procesar o sembrar)."""
    if doc is None or rebar is None:
        return
    try:
        rid = int(rebar.Id.IntegerValue)
        tid = bar_type_id_int(rebar)
    except Exception:
        return
    if tid is None:
        return
    remember_bar_type_id(doc, rid, tid)


def rebar_bar_type_change_detected(doc, rebar, updater_data=None, element_id=None):
    """
    ``True`` si el evento corresponde a cambio de ``RebarBarType`` (diámetro).

    Criterios (cualquiera basta):
    - ``UpdaterData.IsChangeTriggered(..., GetChangeTypeElementType())``
    - ``ElementId`` de tipo distinto al último registrado en caché de sesión
    """
    if rebar is None:
        return False
    try:
        rid = int(rebar.Id.IntegerValue)
    except Exception:
        return False
    tid = bar_type_id_int(rebar)
    if tid is None:
        return False
    prev = peek_bar_type_id(doc, rid)
    if prev is not None and int(prev) != int(tid):
        return True
    if updater_data is not None and element_id is not None:
        try:
            from Autodesk.Revit.DB import Element

            ct = Element.GetChangeTypeElementType()
            if updater_data.IsChangeTriggered(element_id, ct):
                return True
        except Exception:
            pass
    return False


def seed_bar_type_cache_if_unknown(doc, rebar):
    """Registra el tipo actual sin disparar ajuste (primera vez que se ve la barra)."""
    if doc is None or rebar is None:
        return
    try:
        rid = int(rebar.Id.IntegerValue)
    except Exception:
        return
    if peek_bar_type_id(doc, rid) is not None:
        return
    tid = bar_type_id_int(rebar)
    if tid is not None:
        remember_bar_type_id(doc, rid, tid)


def seed_all_rebar_bar_types_in_document(doc):
    """Precarga caché RebarBarType de todas las barras del documento (apertura / registro DMU)."""
    if doc is None:
        return 0
    try:
        from Autodesk.Revit.DB import FilteredElementCollector

        n = 0
        for rb in FilteredElementCollector(doc).OfClass(Rebar):
            if rb is None:
                continue
            before = peek_bar_type_id(doc, int(rb.Id.IntegerValue))
            seed_bar_type_cache_if_unknown(doc, rb)
            after = peek_bar_type_id(doc, int(rb.Id.IntegerValue))
            if before is None and after is not None:
                n += 1
        return n
    except Exception:
        return 0


def _grade_override_all():
    try:
        data = System.AppDomain.CurrentDomain.GetData(_AD_GRADE_OVERRIDE)
    except Exception:
        data = None
    if data is None:
        data = {}
        try:
            System.AppDomain.CurrentDomain.SetData(_AD_GRADE_OVERRIDE, data)
        except Exception:
            pass
    return data


def set_document_concrete_grade_override(doc, concrete_grade):
    """Fija G25/G35/G45 para el documento (sesión Revit). ``None`` = auto-detectar."""
    overrides = _grade_override_all()
    dk = _doc_cache_key(doc)
    g = normalize_concrete_grade(concrete_grade)
    if g is None:
        overrides.pop(dk, None)
    else:
        overrides[dk] = g


def get_document_concrete_grade_override(doc):
    overrides = _grade_override_all()
    return overrides.get(_doc_cache_key(doc))


def _param_text_value(param):
    if param is None:
        return None
    try:
        st = param.StorageType
        if st == StorageType.String:
            return param.AsString()
        if st == StorageType.Integer:
            return unicode(param.AsInteger())
        if st == StorageType.Double:
            try:
                return param.AsValueString()
            except Exception:
                return unicode(param.AsDouble())
        if st == StorageType.ElementId:
            return None
    except Exception:
        return None
    try:
        return param.AsValueString()
    except Exception:
        return None


def _grade_from_element_parameters(element, name_hints=True):
    if element is None:
        return None
    try:
        params = element.Parameters
    except Exception:
        return None
    if params is None:
        return None
    hinted = []
    other = []
    try:
        it = params.GetEnumerator()
        while it.MoveNext():
            p = it.Current
            if p is None or not p.HasValue:
                continue
            try:
                pname = p.Definition.Name or u""
            except Exception:
                pname = u""
            txt = _param_text_value(p)
            g = concrete_grade_from_text(txt)
            if g is None:
                continue
            pl = pname.lower()
            if name_hints and any(h in pl for h in _GRADE_PARAM_NAME_HINTS):
                hinted.append(g)
            else:
                other.append(g)
    except Exception:
        return None
    if hinted:
        return hinted[0]
    if other:
        return other[0]
    return None


def _grade_from_host_material(doc, host):
    if doc is None or host is None:
        return None
    try:
        mid = host.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)
    except Exception:
        mid = None
    if mid is None:
        return None
    try:
        mat_id = mid.AsElementId()
    except Exception:
        return None
    if mat_id is None or mat_id == ElementId.InvalidElementId:
        return None
    try:
        mat = doc.GetElement(mat_id)
    except Exception:
        mat = None
    if mat is None:
        return None
    try:
        return concrete_grade_from_text(mat.Name)
    except Exception:
        return None


def resolve_concrete_grade_for_rebar(doc, rebar, concrete_grade=None):
    """
    Resuelve dosificación G25/G35/G45 para tablas de pata.

    Orden: argumento explícito → override por documento (AppDomain) → parámetros
    del Rebar → host → ``ProjectInformation`` → material estructural del host.
    Si no hay coincidencia, ``None`` (tabla base / legacy G25).
    """
    g = normalize_concrete_grade(concrete_grade)
    if g is not None:
        return g
    if doc is not None:
        g = get_document_concrete_grade_override(doc)
        if g is not None:
            return g
    if rebar is not None:
        g = _grade_from_element_parameters(rebar, name_hints=True)
        if g is not None:
            return g
    host = None
    if doc is not None and rebar is not None:
        try:
            host = doc.GetElement(rebar.GetHostId())
        except Exception:
            host = None
    if host is not None:
        g = _grade_from_element_parameters(host, name_hints=True)
        if g is not None:
            return g
        g = _grade_from_host_material(doc, host)
        if g is not None:
            return g
    if doc is not None:
        try:
            pi = doc.ProjectInformation
        except Exception:
            pi = None
        if pi is not None:
            g = _grade_from_element_parameters(pi, name_hints=True)
            if g is not None:
                return g
    return None


def _grade_label(concrete_grade):
    g = normalize_concrete_grade(concrete_grade)
    return g if g is not None else u"G25"


def _rebar_shape_visible_name(shape):
    if shape is None:
        return u""
    for bip in (
        BuiltInParameter.SYMBOL_NAME_PARAM,
        BuiltInParameter.ALL_MODEL_TYPE_NAME,
    ):
        try:
            p = shape.get_Parameter(bip)
            if p is not None and p.HasValue:
                s = (p.AsString() or u"").strip()
                if s:
                    return s
        except Exception:
            continue
    try:
        return (getattr(shape, u"Name", None) or u"").strip()
    except Exception:
        return u""


def _shape_digits_from_name(name):
    if not name:
        return None
    key = name.strip()
    if key in _SHAPE_SEGMENTS:
        return key
    try:
        key_low = key.lower()
    except Exception:
        key_low = key
    if key_low in _SHAPE_SEGMENTS:
        return key_low
    digits = u"".join(ch for ch in key if ch in u"0123456789")
    if digits in _SHAPE_SEGMENTS:
        return digits
    return None


def shape_digits_from_rebar(doc, rebar):
    """``02``, ``03`` o ``None`` según el ``RebarShape`` de la barra."""
    if doc is None or rebar is None:
        return None
    sid = None
    try:
        sid = rebar.GetShapeId()
    except Exception:
        sid = None
    if sid is None:
        try:
            sid = rebar.RebarShapeId
        except Exception:
            sid = None
    if sid is None or sid == ElementId.InvalidElementId:
        return None
    try:
        shape = doc.GetElement(sid)
    except Exception:
        shape = None
    if not isinstance(shape, RebarShape):
        return None
    return _shape_digits_from_name(_rebar_shape_visible_name(shape))


def segments_for_shape(shape_digits):
    """Letras de segmento a actualizar (``A`` o ``A``+``C``)."""
    if shape_digits is None:
        return ()
    return _SHAPE_SEGMENTS.get(shape_digits, ())


def _nominal_diameter_mm(doc, rebar):
    try:
        bt = doc.GetElement(rebar.GetTypeId())
    except Exception:
        bt = None
    if not isinstance(bt, RebarBarType):
        return None
    try:
        d = bt.BarNominalDiameter
        return float(UnitUtils.ConvertFromInternalUnits(d, UnitTypeId.Millimeters))
    except Exception:
        try:
            d = bt.BarModelDiameter
            return float(UnitUtils.ConvertFromInternalUnits(d, UnitTypeId.Millimeters))
        except Exception:
            return None


def target_pata_tabla_mm(diameter_mm, concrete_grade=None):
    """Largo de pata (mm) según tabla BIMTools para el Ø y la dosificación."""
    try:
        v = hook_length_mm_from_nominal_diameter_mm(diameter_mm, concrete_grade)
        if v is None:
            return None
        return float(v)
    except Exception:
        return None


def _lookup_segment_param(doc, rebar, letter):
    if rebar is None or not letter:
        return None
    want = letter.strip().upper()
    if doc is not None:
        sid = None
        try:
            sid = rebar.GetShapeId()
        except Exception:
            sid = None
        if sid is None:
            try:
                sid = rebar.RebarShapeId
            except Exception:
                sid = None
        if sid is not None and sid != ElementId.InvalidElementId:
            try:
                shape = doc.GetElement(sid)
                defn = (
                    shape.GetRebarShapeDefinition()
                    if shape is not None
                    else None
                )
                if defn is not None:
                    for pid in defn.GetParameters():
                        p = None
                        try:
                            p = rebar.get_Parameter(pid)
                        except Exception:
                            p = None
                        if p is None:
                            try:
                                pe = doc.GetElement(pid)
                                if pe is not None:
                                    p = rebar.LookupParameter(pe.Name)
                            except Exception:
                                p = None
                        if p is None:
                            continue
                        try:
                            pname = (p.Definition.Name or u"").strip().upper()
                        except Exception:
                            pname = u""
                        if pname == want:
                            return p
            except Exception:
                pass
    for name in (want, want.lower()):
        try:
            p = rebar.LookupParameter(name)
            if p is not None:
                return p
        except Exception:
            pass
        try:
            plist = rebar.GetParameters(name)
            if plist is not None and int(plist.Count) > 0:
                return plist[0]
        except Exception:
            pass
    return None


def read_segment_mm(rebar, letter, doc=None):
    p = _lookup_segment_param(doc, rebar, letter)
    if p is None:
        return None
    try:
        if not p.HasValue:
            return None
    except Exception:
        return None
    try:
        if p.StorageType != StorageType.Double:
            return None
        return float(
            UnitUtils.ConvertFromInternalUnits(
                float(p.AsDouble()), UnitTypeId.Millimeters
            )
        )
    except Exception:
        return None


def _set_segment_mm(doc, rebar, letter, mm):
    p = _lookup_segment_param(doc, rebar, letter)
    if p is None:
        return False, u"Sin parámetro {0}.".format(letter)
    try:
        if p.IsReadOnly:
            return False, u"Parámetro {0} de solo lectura.".format(letter)
    except Exception:
        pass
    try:
        p.Set(_mm_to_internal(mm))
        return True, u""
    except Exception as ex:
        return False, u"{0}: {1!s}".format(letter, ex)


def _segments_need_update(rebar, letters, target_tabla_mm, doc=None):
    if not letters or target_tabla_mm is None:
        return False
    tol = float(_PATA_MATCH_TOL_MM)
    tgt = float(target_tabla_mm)
    for letter in letters:
        cur = read_segment_mm(rebar, letter, doc=doc)
        if cur is None:
            return True
        if abs(float(cur) - tgt) > tol:
            return True
    return False


def apply_pata_l_for_diameter_change(doc, rebar, concrete_grade=None, pos_idx=0):
    """
    Ajusta segmentos A (shape «02») o A+C («03») según tabla para el Ø actual.

    Solo debe invocarse cuando ya se confirmó cambio de ``RebarBarType``.

    Returns:
        (changed, message, rebar_out)
    """
    if doc is None or rebar is None or not isinstance(rebar, Rebar):
        return False, u"No es Rebar.", rebar

    shape_key = shape_digits_from_rebar(doc, rebar)
    if shape_key not in _SHAPE_SEGMENTS:
        return False, u"Shape no admitido (solo 02 / 03).", rebar

    letters = segments_for_shape(shape_key)
    if not letters:
        return False, u"Shape sin segmentos definidos.", rebar

    d_mm = _nominal_diameter_mm(doc, rebar)
    if d_mm is None or float(d_mm) <= 0.0:
        return False, u"Diámetro no resuelto.", rebar

    grade_eff = resolve_concrete_grade_for_rebar(doc, rebar, concrete_grade)
    target_tabla = target_pata_tabla_mm(d_mm, grade_eff)
    if target_tabla is None or float(target_tabla) < 0.1:
        return False, u"Largo tabla inválido.", rebar

    if not _segments_need_update(rebar, letters, target_tabla, doc=doc):
        return False, u"Segmentos ya coinciden con tabla.", rebar

    txn = Transaction(doc, _TXN)
    try:
        if txn.Start() != TransactionStatus.Started:
            return False, u"No se pudo abrir transacción.", rebar
    except Exception:
        return False, u"Transacción falló.", rebar

    try:
        for letter in letters:
            ok, err = _set_segment_mm(doc, rebar, letter, target_tabla)
            if not ok:
                txn.RollBack()
                return False, err or u"Error al escribir {0}.".format(letter), rebar
        try:
            doc.Regenerate()
        except Exception:
            pass
        txn.Commit()
    except Exception as ex:
        try:
            if txn.GetStatus() == TransactionStatus.Started:
                txn.RollBack()
        except Exception:
            pass
        return False, u"{0!s}".format(ex), rebar

    seg_txt = u", ".join(letters)
    rid_out = int(rebar.Id.IntegerValue)
    msg = (
        u"Shape {0}: {1} → {2:.0f} mm (Ø {3:.0f} mm, {4}). id {5}."
    ).format(
        shape_key,
        seg_txt,
        float(target_tabla),
        float(d_mm),
        _grade_label(grade_eff),
        rid_out,
    )
    return True, msg, rebar


def process_rebars_for_diameter_pata_l(doc, rebar_element_ids, concrete_grade=None):
    """
    Procesa un lote de ids de ``Rebar`` tras cambio de diámetro.

    Returns:
        dict con contadores ``ok``, ``skip``, ``fail``, ``messages``.
    """
    res = {u"ok": 0, u"skip": 0, u"fail": 0, u"messages": []}
    if doc is None or not rebar_element_ids:
        return res
    for eid in rebar_element_ids:
        try:
            el = doc.GetElement(eid)
        except Exception:
            el = None
        if el is None or not isinstance(el, Rebar):
            res[u"skip"] += 1
            continue
        changed, msg, _rb = apply_pata_l_for_diameter_change(
            doc, el, concrete_grade=concrete_grade
        )
        remember_bar_type_snapshot_for_rebar(doc, el)
        if changed:
            res[u"ok"] += 1
            if msg:
                res[u"messages"].append(msg)
        elif msg and (
            u"Sin cambio" in msg
            or u"no admitido" in msg
            or u"ya coinciden" in msg
        ):
            res[u"skip"] += 1
        else:
            res[u"skip"] += 1
    return res
