# -*- coding: utf-8 -*-
"""
Núcleo compartido: sincronizar tipo de IndependentTag (Rebar Tags) con RebarShape.

Usado por:
- rebar_shape_tag_updater_dmu.py (DMU + ExternalEvent)

Configuración DMU: ajusta REBAR_TAG_SYNC_DEFAULT_FAMILY_NAMES (tupla de nombres de familia).
"""

from __future__ import print_function

import unicodedata

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FailureProcessingResult,
    Family,
    FamilySymbol,
    FilteredElementCollector,
    IFailuresPreprocessor,
    IndependentTag,
    Reference,
    StorageType,
    SubTransaction,
    TagOrientation,
    Transaction,
    TransactionStatus,
    XYZ,
)

# Editar según proyecto: nombres exactos de familia(s) Structural Rebar Tags (navegador).
# El DMU usa esta lista; el botón puede inferir familia desde la vista si no defines fallback.
REBAR_TAG_SYNC_DEFAULT_FAMILY_NAMES = (u"EST_A_STRUCTURAL REBAR TAG",)


def _to_python_str(s):
    if s is None:
        return u""
    try:
        t = s.ToString()
    except Exception:
        try:
            t = str(s)
        except Exception:
            return u""
    if isinstance(t, bytes):
        try:
            return t.decode("utf-8", "replace")
        except Exception:
            return u""
    try:
        return type(u"")(t)
    except Exception:
        return u""


def normalize_label(s):
    t = _to_python_str(s).strip()
    if not t:
        return u""
    t = unicodedata.normalize("NFKC", t)
    t = u" ".join(t.split())
    return t


def _is_ascii_digit_string(s):
    return bool(s) and all(u"0" <= c <= u"9" for c in s)


def comparison_keys(label_raw):
    base = normalize_label(label_raw)
    if not base:
        return []
    seen = set()
    out = []

    def add(k):
        if not k:
            return
        if k not in seen:
            seen.add(k)
            out.append(k)

    add(base)
    add(base.lower())
    for pref in (u"Shape ", u"shape ", u"Forma ", u"forma ", u"Form ", u"form "):
        if base.startswith(pref):
            rest = normalize_label(base[len(pref) :])
            add(rest)
            add(rest.lower())
    for sep in (u":", u" - ", u"-", u" – "):
        if sep in base:
            tail = normalize_label(base.split(sep)[-1])
            add(tail)
            add(tail.lower())
    if _is_ascii_digit_string(base):
        add(base.zfill(2))
        add(base.zfill(2).lower())
        try:
            n = int(base)
            add(type(u"")(n))
            add(type(u"")(n).zfill(2))
        except Exception:
            pass
    return out


def family_names_match(fam_name, expected):
    a = normalize_label(_to_python_str(fam_name)).lower()
    b = normalize_label(_to_python_str(expected)).lower()
    return bool(a) and a == b


def is_rebar_category(el, bic=None):
    if bic is None:
        bic = BuiltInCategory
    try:
        if el is None or el.Category is None:
            return False
        return int(el.Category.Id.IntegerValue) == int(bic.OST_Rebar)
    except Exception:
        return False


def is_rebar_tag_category(el, bic=None):
    if bic is None:
        bic = BuiltInCategory
    try:
        if el is None or el.Category is None:
            return False
        return int(el.Category.Id.IntegerValue) == int(bic.OST_RebarTags)
    except Exception:
        return False


def element_type_ids_equal(a, b):
    try:
        return int(a.IntegerValue) == int(b.IntegerValue)
    except Exception:
        return False


class TagSyncFailurePreprocessor(IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        return FailureProcessingResult.Continue


def rebar_shape_name_candidates(doc, rebar):
    seen = set()
    out = []

    def push_raw(raw):
        n = normalize_label(raw)
        if n and n not in seen:
            seen.add(n)
            out.append(n)

    try:
        sid = rebar.GetShapeId()
    except Exception:
        try:
            sid = rebar.RebarShapeId
        except Exception:
            sid = None
    if sid is None or sid == ElementId.InvalidElementId:
        return out
    try:
        if int(sid.IntegerValue) < 0:
            return out
    except Exception:
        pass

    sh = doc.GetElement(sid)
    if sh is None:
        return out

    push_raw(_to_python_str(getattr(sh, u"Name", None)))

    for bip_name in (u"ALL_MODEL_TYPE_NAME", u"SYMBOL_NAME_PARAM"):
        bip = getattr(BuiltInParameter, bip_name, None)
        if bip is None:
            continue
        p = sh.get_Parameter(bip)
        if p is None or not p.HasValue:
            continue
        try:
            push_raw(p.AsString())
        except Exception:
            pass

    for bip_name in (u"REBAR_SHAPE", u"REBAR_SHAPE_ID"):
        bip = getattr(BuiltInParameter, bip_name, None)
        if bip is None:
            continue
        p = rebar.get_Parameter(bip)
        if p is None or not p.HasValue:
            continue
        if p.StorageType == StorageType.String:
            try:
                push_raw(p.AsString())
            except Exception:
                pass
        elif p.StorageType == StorageType.ElementId:
            try:
                eid = p.AsElementId()
                el2 = doc.GetElement(eid)
                if el2 is not None:
                    push_raw(_to_python_str(getattr(el2, u"Name", None)))
            except Exception:
                pass

    return out


def norm_family_name(s):
    return normalize_label(s).lower()


def find_families_by_name(document, family_name_wanted):
    wanted = norm_family_name(family_name_wanted)
    if not wanted:
        return []
    matches = []
    for fam in FilteredElementCollector(document).OfClass(Family):
        if norm_family_name(fam.Name) == wanted:
            matches.append(fam)
    return matches


def family_rebar_tag_symbols_merged(doc, fam):
    by_id = {}
    try:
        fam_fid = int(fam.Id.IntegerValue)
    except Exception:
        fam_fid = None
    try:
        for sid in fam.GetFamilySymbolIds():
            sym = doc.GetElement(sid)
            if sym is None:
                continue
            try:
                by_id[int(sym.Id.IntegerValue)] = sym
            except Exception:
                pass
    except Exception:
        pass
    if fam_fid is not None:
        try:
            try:
                fs_clr = clr.GetClrType(FamilySymbol)
            except Exception:
                fs_clr = FamilySymbol
            for sym in FilteredElementCollector(doc).OfClass(fs_clr):
                try:
                    sf = sym.Family
                    if sf is None:
                        continue
                    if int(sf.Id.IntegerValue) != fam_fid:
                        continue
                except Exception:
                    continue
                try:
                    k = int(sym.Id.IntegerValue)
                except Exception:
                    continue
                if k not in by_id:
                    by_id[k] = sym
        except Exception:
            pass
    return list(by_id.values())


def type_symbol_label_strings(el):
    seen = set()
    out = []

    def push_raw(raw):
        n = normalize_label(raw)
        if not n or n in seen:
            return
        seen.add(n)
        out.append(n)

    push_raw(_to_python_str(getattr(el, u"Name", None)))
    for bip_name in (u"SYMBOL_NAME_PARAM", u"ALL_MODEL_TYPE_NAME"):
        bip = getattr(BuiltInParameter, bip_name, None)
        if bip is None:
            continue
        try:
            p = el.get_Parameter(bip)
            if p is None or not p.HasValue:
                continue
            push_raw(p.AsString())
        except Exception:
            pass
    return out


def symbol_map_add_keys_for_sym(out, sym):
    try:
        sid = sym.Id
    except Exception:
        return
    try:
        for label in type_symbol_label_strings(sym):
            for key in comparison_keys(label):
                if key not in out:
                    out[key] = sid
    except Exception:
        pass


def symbol_map_from_family(doc, fam):
    if fam is None:
        return {}
    try:
        fam_ref = doc.GetElement(fam.Id)
        if fam_ref is not None:
            fam = fam_ref
    except Exception:
        pass
    bic = BuiltInCategory
    rebar_tag_cat = int(bic.OST_RebarTags)
    candidates = family_rebar_tag_symbols_merged(doc, fam)
    for require_rebar_tag_cat in (True, False):
        out = {}
        for sym in candidates:
            try:
                if require_rebar_tag_cat:
                    cat = sym.Category
                    if cat is None:
                        continue
                    if int(cat.Id.IntegerValue) != rebar_tag_cat:
                        continue
                symbol_map_add_keys_for_sym(out, sym)
            except Exception:
                continue
        if out:
            return out
    return {}


def symbol_map_from_family_names(doc, family_names):
    """Une mapas de todas las familias cuyo nombre coincide (lista de strings)."""
    combined = {}
    for name in family_names:
        if not normalize_label(name):
            continue
        for fam in find_families_by_name(doc, name):
            part = symbol_map_from_family(doc, fam)
            for k, v in part.items():
                if k not in combined:
                    combined[k] = v
    return combined


def lookup_tag_type_id(symbol_map, shape_label):
    if not symbol_map or not normalize_label(shape_label):
        return None
    for key in comparison_keys(_to_python_str(shape_label)):
        tid = symbol_map.get(key)
        if tid is not None:
            return tid
    return None


def tag_rebar_int_if_match(tag, rebar_set, invalid_element_id):
    try:
        tagged = tag.GetTaggedLocalElementIds()
        for tid in tagged:
            try:
                ti = int(tid.IntegerValue)
            except Exception:
                continue
            if ti in rebar_set:
                return ti
    except Exception:
        pass
    try:
        for leid in tag.GetTaggedElementIds():
            try:
                link_inst = leid.LinkInstanceId
                if (
                    link_inst is not None
                    and link_inst != invalid_element_id
                    and int(link_inst.IntegerValue) >= 0
                ):
                    continue
            except Exception:
                pass
            for attr in (u"LinkedElementId", u"HostElementId"):
                try:
                    eid = getattr(leid, attr, None)
                    if eid is None:
                        continue
                    ti = int(eid.IntegerValue)
                    if ti in rebar_set:
                        return ti
                except Exception:
                    continue
    except Exception:
        pass
    return None


def collect_tag_rebar_pairs(doc, rebar_ints):
    if not rebar_ints:
        return []
    rebar_set = set(int(x) for x in rebar_ints)
    invalid = ElementId.InvalidElementId
    out = []
    try:
        try:
            it_clr = clr.GetClrType(IndependentTag)
        except Exception:
            it_clr = IndependentTag
        coll = (
            FilteredElementCollector(doc)
            .OfClass(it_clr)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        return out
    for el in coll:
        ti = tag_rebar_int_if_match(el, rebar_set, invalid)
        if ti is not None:
            out.append((el, ti))
    return out


def activate_types_without_new_transaction(document, activate_ids_int_set):
    for iv in activate_ids_int_set:
        sym = document.GetElement(ElementId(iv))
        if sym is None:
            continue
        try:
            if bool(sym.IsActive):
                continue
        except Exception:
            continue
        try:
            sym.Activate()
        except Exception as ex:
            print(
                u"  [aviso] Activate ElementId={}: {}".format(iv, _to_python_str(ex))
            )


def document_has_open_transaction(document):
    try:
        return bool(document.IsModifiable)
    except Exception:
        return False


def recreate_rebar_independent_tag(doc, tag, rebar, type_id, fallback_old_id):
    if tag is None or not tag.IsValidObject:
        try:
            tag = doc.GetElement(fallback_old_id)
        except Exception:
            return False
    if tag is None or not tag.IsValidObject:
        return False
    try:
        view_id = tag.OwnerViewId
        if view_id is None or view_id == ElementId.InvalidElementId:
            return False
    except Exception:
        return False
    add_leader = False
    try:
        add_leader = bool(tag.HasLeader)
    except Exception:
        pass
    orient = TagOrientation.Horizontal
    try:
        orient = tag.TagOrientation
    except Exception:
        pass
    head = None
    try:
        head = tag.TagHeadPosition
    except Exception:
        pass
    view = doc.GetElement(view_id)
    if head is None and view is not None:
        try:
            bb = rebar.get_BoundingBox(view)
            if bb is None:
                bb = rebar.get_BoundingBox(None)
            if bb is not None and bb.Min is not None and bb.Max is not None:
                head = XYZ(
                    (bb.Min.X + bb.Max.X) * 0.5,
                    (bb.Min.Y + bb.Max.Y) * 0.5,
                    (bb.Min.Z + bb.Max.Z) * 0.5,
                )
        except Exception:
            pass
    if head is None:
        return False
    st = SubTransaction(doc)
    try:
        st.Start()
    except Exception:
        return False
    try:
        old_id = tag.Id
        doc.Delete(old_id)
        IndependentTag.Create(
            doc, type_id, view_id, Reference(rebar), add_leader, orient, head
        )
        if st.GetStatus() == TransactionStatus.Started:
            st.Commit()
        return True
    except Exception:
        if st.GetStatus() == TransactionStatus.Started:
            try:
                st.RollBack()
            except Exception:
                pass
        return False


def try_set_tag_type(doc, tag, rebar, type_id, no_activate=False):
    if tag is None or not tag.IsValidObject or type_id is None:
        return False, u"argumentos"
    try:
        tid_tag = tag.Id
    except Exception:
        return False, u"id tag"
    new_sym = doc.GetElement(type_id)
    if new_sym is None:
        return False, u"tipo inexistente"
    if no_activate:
        try:
            if not bool(new_sym.IsActive):
                new_sym.Activate()
        except Exception as ex:
            return False, _to_python_str(ex)
    else:
        if not bool(getattr(new_sym, u"IsActive", True)):
            try:
                new_sym.Activate()
            except Exception as ex:
                return False, _to_python_str(ex)
    try:
        cur = tag.GetTypeId()
        if element_type_ids_equal(cur, type_id):
            return True, None
    except Exception:
        pass
    unpinned = False
    try:
        if bool(tag.Pinned):
            tag.Pinned = False
            unpinned = True
    except Exception:
        pass
    try:
        tag.ChangeTypeId(type_id)
    except Exception:
        pass
    finally:
        if unpinned:
            try:
                tag.Pinned = True
            except Exception:
                pass
    tag_after = doc.GetElement(tid_tag)
    if tag_after is not None and element_type_ids_equal(tag_after.GetTypeId(), type_id):
        return True, None
    if recreate_rebar_independent_tag(doc, tag_after or tag, rebar, type_id, tid_tag):
        return True, None
    return False, u"ChangeTypeId y recreación fallaron"


def plan_sync_operations_from_rebar_ints(doc, rebar_ints, symbol_map):
    """
    Lista de (tag_id_int, rebar_id_int, type_id) y conjunto de ids de tipo a activar.
    Solo operaciones donde hace falta cambiar tipo y hay match en symbol_map.
    """
    if not symbol_map or not rebar_ints:
        return [], set()
    bic = BuiltInCategory
    pairs = collect_tag_rebar_pairs(doc, rebar_ints)
    ops = []
    activate = set()
    for tag, rebar_i in pairs:
        if tag is None:
            continue
        try:
            tid_tag = tag.Id
        except Exception:
            continue
        tag = doc.GetElement(tid_tag)
        if tag is None or not tag.IsValidObject:
            continue
        try:
            if getattr(tag, u"IsOrphaned", False):
                continue
        except Exception:
            pass
        rebar = doc.GetElement(ElementId(rebar_i))
        if rebar is None or not is_rebar_category(rebar, bic):
            continue
        type_id = None
        for shape_label in rebar_shape_name_candidates(doc, rebar):
            type_id = lookup_tag_type_id(symbol_map, shape_label)
            if type_id is not None:
                break
        if type_id is None:
            continue
        try:
            cur = tag.GetTypeId()
            if element_type_ids_equal(cur, type_id):
                continue
        except Exception:
            pass
        try:
            ops.append((int(tid_tag.IntegerValue), rebar_i, type_id))
            activate.add(int(type_id.IntegerValue))
        except Exception:
            continue
    return ops, activate


def apply_sync_operations(doc, ops, activate_ids):
    """Ejecuta operaciones dentro de transacción ya iniciada."""
    activate_types_without_new_transaction(doc, activate_ids)
    try:
        doc.Regenerate()
    except Exception:
        pass
    n_ok = 0
    n_fail = 0
    for tag_id_int, rebar_id_int, type_id in ops:
        tag = doc.GetElement(ElementId(tag_id_int))
        rebar = doc.GetElement(ElementId(rebar_id_int))
        if tag is None or rebar is None:
            n_fail += 1
            continue
        ok, err = try_set_tag_type(doc, tag, rebar, type_id, no_activate=True)
        if ok:
            n_ok += 1
        else:
            n_fail += 1
            print(u"  [fallo] tag id={}: {}".format(tag_id_int, err))
    return n_ok, n_fail


def execute_sync_with_transaction(doc, txn_name, ops, activate_ids):
    """
    Abre Transaction o SubTransaction y aplica apply_sync_operations.
    Devuelve (n_ok, n_fail) o (0, 0) si no hay ops.
    """
    if not ops:
        return 0, 0
    txn = Transaction(doc, txn_name)
    try:
        opt = txn.GetFailureHandlingOptions()
        opt.SetFailuresPreprocessor(TagSyncFailurePreprocessor())
        txn.SetFailureHandlingOptions(opt)
    except Exception:
        pass
    txn_started = False
    st = None
    sub_started = False
    try:
        txn.Start()
        txn_started = True
    except Exception as ex_outer:
        print(
            u"  [diag] Transaction.Start: {} — probando SubTransaction.".format(
                _to_python_str(ex_outer)
            )
        )
        try:
            st = SubTransaction(doc)
            st.Start()
            sub_started = True
        except Exception as ex_sub:
            if document_has_open_transaction(doc):
                print(
                    u"  [diag] SubTransaction.Start: {} — continuando sin envoltorio.".format(
                        _to_python_str(ex_sub)
                    )
                )
            else:
                print(
                    u"  [error] No se pudo abrir transacción: {}".format(
                        _to_python_str(ex_sub)
                    )
                )
                return 0, len(ops)

    n_ok = n_fail = 0
    try:
        n_ok, n_fail = apply_sync_operations(doc, ops, activate_ids)
    finally:
        if txn_started and txn.GetStatus() == TransactionStatus.Started:
            try:
                txn.Commit()
            except Exception:
                try:
                    txn.RollBack()
                except Exception:
                    pass
        elif sub_started and st is not None:
            try:
                if st.GetStatus() == TransactionStatus.Started:
                    st.Commit()
            except Exception:
                try:
                    if st.GetStatus() == TransactionStatus.Started:
                        st.RollBack()
                except Exception:
                    pass
    return n_ok, n_fail


def apply_tag_sync_for_rebar_ints(doc, rebar_ints, symbol_map, txn_name):
    """
    API principal para DMU: rebar_ints = iterable de ids enteros de Rebar;
    symbol_map = dict de claves de comparación -> ElementId de tipo de etiqueta.
    """
    if not rebar_ints or not symbol_map:
        return 0, 0
    try:
        doc.Regenerate()
    except Exception:
        pass
    ints = []
    for x in rebar_ints:
        try:
            ints.append(int(x))
        except Exception:
            pass
    ops, activate = plan_sync_operations_from_rebar_ints(doc, ints, symbol_map)
    return execute_sync_with_transaction(doc, txn_name, ops, activate)


to_python_str = _to_python_str  # alias público para diagnósticos / DMU
