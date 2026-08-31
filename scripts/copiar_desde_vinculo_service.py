# -*- coding: utf-8 -*-
"""
Servicio — copiar elementos desde un modelo vinculado al documento host.

Revit 2024+ | pyRevit / IronPython.
Usa ``GetTotalTransform()`` del vínculo para conservar la posición en el host.
"""

from __future__ import print_function

try:
    unicode
except NameError:
    unicode = str

import clr

clr.AddReference("RevitAPI")
clr.AddReference("System")

from System.Collections.Generic import List
from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    CopyPasteOptions,
    DuplicateTypeAction,
    ElementId,
    ElementTransformUtils,
    FailureProcessingResult,
    FailureResolutionType,
    FailureSeverity,
    FilteredElementCollector,
    IDuplicateTypeNamesHandler,
    IFailuresPreprocessor,
    RevitLinkInstance,
    Transaction,
    TransactionGroup,
)

# Lista blanca de categorías copiables (orden de presentación en UI).
_ALLOWED_BUILTIN_NAMES = (
    u"OST_StructuralColumns",
    u"OST_StructuralFraming",
    u"OST_Walls",
    u"OST_Floors",
    u"OST_Grids",
    u"OST_Levels",
    u"OST_StructuralFoundation",
)

_ALLOWED_DISPLAY_LABELS = {
    u"OST_StructuralColumns": u"Structural Columns",
    u"OST_StructuralFraming": u"Structural Framing",
    u"OST_Walls": u"Walls",
    u"OST_Floors": u"Floors",
    u"OST_Grids": u"Grids",
    u"OST_Levels": u"Levels",
    u"OST_StructuralFoundation": u"Structural Foundations",
}


def _safe_builtin_category(name):
    try:
        return getattr(BuiltInCategory, name)
    except Exception:
        return None


def _build_allowed_builtin_map():
    """BuiltInCategory → (orden, etiqueta fija en inglés)."""
    mapping = {}
    for index, name in enumerate(_ALLOWED_BUILTIN_NAMES):
        bic = _safe_builtin_category(name)
        if bic is None:
            continue
        mapping[bic] = (index, _ALLOWED_DISPLAY_LABELS.get(name) or name)
    return mapping


_ALLOWED_BUILTIN = _build_allowed_builtin_map()
_TRANSACTION_NAME = u"Arainco: Copiar desde vínculo"


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _group_id(kind, label):
    return u"{0}\x00{1}".format(kind, label)


class _UseDestinationTypes(IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        return DuplicateTypeAction.UseDestinationTypes


def _copy_options():
    opts = CopyPasteOptions()
    try:
        opts.SetDuplicateTypeNamesHandler(_UseDestinationTypes())
    except Exception:
        pass
    return opts


def _element_id_key(eid):
    if eid is None:
        return None
    try:
        return int(eid.Value)
    except Exception:
        try:
            return int(eid.IntegerValue)
        except Exception:
            return None


def _selectable_category_info(element):
    """
    Si el elemento pertenece a una categoría permitida, devuelve
    ``(orden, etiqueta)``; si no, ``None``.
    """
    try:
        cat = element.Category
        if cat is None:
            return None
        bic = cat.BuiltInCategory
        meta = _ALLOWED_BUILTIN.get(bic)
        if meta is None:
            return None
        return meta
    except Exception:
        return None


def list_loaded_revit_links(doc):
    """
    Vínculos Revit cargados en el documento host.

    Returns:
        list[dict]: ``link_id_int``, ``name``, ``link_doc_title``, ``instance``
    """
    out = []
    if doc is None:
        return out
    try:
        links = (
            FilteredElementCollector(doc)
            .OfClass(RevitLinkInstance)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        return out

    seen = set()
    for inst in links or []:
        if inst is None:
            continue
        key = _element_id_key(inst.Id)
        if key is None or key in seen:
            continue
        seen.add(key)

        link_doc = None
        try:
            link_doc = inst.GetLinkDocument()
        except Exception:
            link_doc = None
        if link_doc is None:
            continue

        name = u""
        try:
            name = _as_unicode(inst.Name).strip()
        except Exception:
            pass

        doc_title = u""
        try:
            doc_title = _as_unicode(link_doc.Title).strip()
        except Exception:
            pass

        out.append(
            {
                u"link_id_int": key,
                u"name": name or doc_title or u"Vínculo {0}".format(key),
                u"link_doc_title": doc_title,
                u"instance": inst,
            }
        )

    out.sort(key=lambda item: (item.get(u"name") or u"").lower())
    return out


def collect_link_categories(link_doc):
    """
    Agrupa instancias del vínculo por las categorías permitidas.

    Siempre devuelve las 7 categorías de la lista blanca (con count 0 si
    no hay instancias), en el orden fijo de presentación.
    """
    # Plantilla fija: una entrada por categoría permitida.
    slots = []
    for index, name in enumerate(_ALLOWED_BUILTIN_NAMES):
        label = _ALLOWED_DISPLAY_LABELS.get(name) or name
        slots.append(
            {
                u"kind": u"modelo",
                u"label": label,
                u"group_id": _group_id(u"modelo", label),
                u"sort_order": index,
                u"element_ids": [],
                u"count": 0,
            }
        )
    by_label = dict((item[u"label"], item) for item in slots)

    if link_doc is None:
        return slots

    try:
        elems = (
            FilteredElementCollector(link_doc)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        return slots

    for el in elems or []:
        info = _selectable_category_info(el)
        if info is None:
            continue
        _order, label = info
        bucket = by_label.get(label)
        if bucket is None:
            continue
        try:
            bucket[u"element_ids"].append(el.Id)
        except Exception:
            continue

    for item in slots:
        item[u"count"] = len(item[u"element_ids"])
    return slots


def summarize_category_groups(groups):
    total = 0
    with_elements = 0
    for grp in groups or []:
        count = grp.get(u"count", 0)
        total += count
        if count > 0:
            with_elements += 1
    return {
        u"total": total,
        u"categories": len(groups or []),
        u"categories_with_elements": with_elements,
        u"model_count": total,
        u"annotation_count": 0,
        u"model_categories": with_elements,
        u"annotation_categories": 0,
    }


def _to_id_list(element_ids):
    id_list = List[ElementId]()
    for eid in element_ids or []:
        if eid is None:
            continue
        if isinstance(eid, ElementId):
            id_list.Add(eid)
        else:
            try:
                id_list.Add(ElementId(int(eid)))
            except Exception:
                pass
    return id_list


def _failure_description(failure_msg):
    if failure_msg is None:
        return u""
    for getter in (
        u"GetDescriptionText",
        u"GetAdditionalDescription",
    ):
        try:
            fn = getattr(failure_msg, getter, None)
            if fn is None:
                continue
            text = _as_unicode(fn()).strip()
            if text:
                return text
        except Exception:
            continue
    try:
        return _as_unicode(failure_msg).strip()
    except Exception:
        return u""


def _failure_severity_label(severity):
    try:
        if severity == FailureSeverity.Warning:
            return u"Warning"
        if severity == FailureSeverity.Error:
            return u"Error"
        if severity == FailureSeverity.DocumentCorruption:
            return u"DocumentCorruption"
    except Exception:
        pass
    return u"Other"


def _iter_failure_msgs(failures_accessor):
    if failures_accessor is None:
        return
    try:
        fmsgs = failures_accessor.GetFailureMessages()
    except Exception:
        return
    if fmsgs is None:
        return
    try:
        n = int(fmsgs.Count)
    except Exception:
        n = 0
    for i in range(n):
        f = None
        try:
            f = fmsgs.get_Item(i)
        except Exception:
            try:
                f = fmsgs[i]
            except Exception:
                f = None
        if f is not None:
            yield f


def _failure_is_zero_height_column(failure_msg):
    """Error típico al copiar columnas: altura 0.0 por offsets/niveles."""
    desc = _failure_description(failure_msg).lower()
    if not desc:
        return False
    if u"column height is not 0.0" in desc:
        return True
    if u"change offset value" in desc and u"column" in desc:
        return True
    if u"altura" in desc and (u"columna" in desc or u"column" in desc) and u"0" in desc:
        return True
    return False


def _try_resolve_delete_elements(failures_accessor, fma):
    """
    Aplica la resolución «Delete Element(s)» (o resolución por defecto).
    Evita DeleteError, que deja elementos inválidos en el modelo.
    """
    if failures_accessor is None or fma is None:
        return False

    def _permitted(rt):
        try:
            return bool(failures_accessor.IsFailureResolutionPermitted(fma, rt))
        except Exception:
            pass
        try:
            if hasattr(fma, u"HasResolutionOfType"):
                return bool(fma.HasResolutionOfType(rt))
        except Exception:
            pass
        return False

    rt_del = getattr(FailureResolutionType, u"DeleteElements", None)
    if rt_del is not None and _permitted(rt_del):
        try:
            failures_accessor.SetCurrentResolutionType(fma, rt_del)
            failures_accessor.ResolveFailure(fma)
            return True
        except Exception:
            pass

    try:
        has_res = False
        try:
            if hasattr(fma, u"HasResolutions"):
                has_res = bool(fma.HasResolutions())
        except Exception:
            has_res = False
        if has_res:
            try:
                if hasattr(failures_accessor, u"IsFailureResolutionPermitted"):
                    if not bool(failures_accessor.IsFailureResolutionPermitted(fma)):
                        return False
            except (Exception, TypeError):
                pass
            failures_accessor.ResolveFailure(fma)
            return True
    except Exception:
        pass

    # Respaldo: borrar ids fallidos directamente.
    try:
        ids = fma.GetFailingElementIds()
    except Exception:
        ids = None
    if ids is not None:
        try:
            if int(ids.Count) > 0:
                failures_accessor.DeleteElements(ids)
                return True
        except Exception:
            pass
    return False


def _collect_failing_element_ids(failure_msgs):
    """Une todos los ElementId fallidos de una lista de FailureMessage."""
    id_list = List[ElementId]()
    seen = set()
    for fma in failure_msgs or []:
        for getter in (u"GetFailingElementIds", u"GetAdditionalElementIds"):
            try:
                ids = getattr(fma, getter)()
            except Exception:
                ids = None
            if ids is None:
                continue
            try:
                n = int(ids.Count)
            except Exception:
                n = 0
            for i in range(n):
                try:
                    eid = ids.get_Item(i)
                except Exception:
                    try:
                        eid = ids[i]
                    except Exception:
                        eid = None
                key = _element_id_key(eid)
                if key is None or key in seen:
                    continue
                seen.add(key)
                id_list.Add(eid)
    return id_list, len(seen)


_ZERO_HEIGHT_TOL_FT = 1.0 / 304.8  # ~1 mm
_MIN_COLUMN_HEIGHT_FT = 100.0 / 304.8  # 100 mm si no hay dato en el link


def _structural_column_height_ft(doc, el):
    """Altura vertical de columna estructural (ft); None si no se puede calcular."""
    if el is None or doc is None:
        return None
    try:
        p_base = el.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM)
        p_top = el.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM)
        p_base_off = el.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM)
        p_top_off = el.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM)
        if p_base is None or p_top is None:
            raise Exception(u"no level params")
        base_lvl = doc.GetElement(p_base.AsElementId())
        top_lvl = doc.GetElement(p_top.AsElementId())
        if base_lvl is None or top_lvl is None:
            raise Exception(u"missing levels")
        z0 = float(base_lvl.Elevation)
        z1 = float(top_lvl.Elevation)
        if p_base_off is not None:
            try:
                z0 += float(p_base_off.AsDouble())
            except Exception:
                pass
        if p_top_off is not None:
            try:
                z1 += float(p_top_off.AsDouble())
            except Exception:
                pass
        return abs(z1 - z0)
    except Exception:
        pass
    try:
        bb = el.get_BoundingBox(None)
        if bb is not None:
            return abs(float(bb.Max.Z) - float(bb.Min.Z))
    except Exception:
        pass
    return None


def _instance_origin(el):
    """Punto de ubicación de una instancia (o centro de bbox)."""
    if el is None:
        return None
    try:
        loc = el.Location
        if loc is not None and hasattr(loc, u"Point") and loc.Point is not None:
            return loc.Point
    except Exception:
        pass
    try:
        loc = el.Location
        if loc is not None and hasattr(loc, u"Curve") and loc.Curve is not None:
            return loc.Curve.Evaluate(0.5, True)
    except Exception:
        pass
    try:
        bb = el.get_BoundingBox(None)
        if bb is not None:
            return (bb.Min + bb.Max) * 0.5
    except Exception:
        pass
    return None


def _is_structural_column(el):
    try:
        cat = el.Category
        if cat is None:
            return False
        return cat.BuiltInCategory == BuiltInCategory.OST_StructuralColumns
    except Exception:
        return False


def _fix_column_height(host_doc, host_el, desired_height_ft):
    """
    Corrige una columna con altura ~0: mismo nivel base/tope y offset superior
    = offset base + altura deseada (tomada del vínculo).
    """
    if host_el is None or host_doc is None:
        return False
    h = float(desired_height_ft or 0.0)
    if h <= _ZERO_HEIGHT_TOL_FT:
        h = _MIN_COLUMN_HEIGHT_FT

    p_base = host_el.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM)
    p_top = host_el.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM)
    p_base_off = host_el.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM)
    p_top_off = host_el.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM)
    if p_base is None or p_top is None or p_top_off is None:
        return False
    if p_top.IsReadOnly or p_top_off.IsReadOnly:
        return False

    try:
        base_id = p_base.AsElementId()
    except Exception:
        return False
    if base_id is None:
        return False

    base_off = 0.0
    if p_base_off is not None:
        try:
            if p_base_off.HasValue:
                base_off = float(p_base_off.AsDouble())
        except Exception:
            base_off = 0.0

    try:
        p_top.Set(base_id)
        p_top_off.Set(base_off + h)
    except Exception:
        return False
    return True


def _repair_zero_height_columns(host_doc, link_doc, transform, source_ids, copied_ids):
    """
    Corrige columnas copiadas con altura ~0 usando la altura del origen en el vínculo.

    Returns:
        (n_fixed, n_unfixed)
    """
    sources = []
    for sid in source_ids or []:
        try:
            el = link_doc.GetElement(sid) if link_doc is not None else None
        except Exception:
            el = None
        if el is None or not _is_structural_column(el):
            continue
        pt = _instance_origin(el)
        if pt is not None and transform is not None:
            try:
                pt = transform.OfPoint(pt)
            except Exception:
                pass
        h = _structural_column_height_ft(link_doc, el)
        if h is None or h <= _ZERO_HEIGHT_TOL_FT:
            try:
                bb = el.get_BoundingBox(None)
                if bb is not None:
                    h = abs(float(bb.Max.Z) - float(bb.Min.Z))
            except Exception:
                h = None
        sources.append((pt, h, el))

    n_fixed = 0
    n_unfixed = 0
    for cid in copied_ids or []:
        try:
            hel = host_doc.GetElement(cid)
        except Exception:
            hel = None
        if hel is None or not _is_structural_column(hel):
            continue
        hh = _structural_column_height_ft(host_doc, hel)
        if hh is not None and hh > _ZERO_HEIGHT_TOL_FT:
            continue

        desired = None
        hpt = _instance_origin(hel)
        if hpt is not None and sources:
            best_h = None
            best_d = None
            for spt, sh, _sel in sources:
                if spt is None:
                    continue
                try:
                    d = hpt.DistanceTo(spt)
                except Exception:
                    continue
                if best_d is None or d < best_d:
                    best_d = d
                    best_h = sh
            desired = best_h
        if desired is None:
            for _spt, sh, _sel in sources:
                if sh is not None and sh > _ZERO_HEIGHT_TOL_FT:
                    desired = sh
                    break

        if _fix_column_height(host_doc, hel, desired):
            n_fixed += 1
        else:
            n_unfixed += 1

    if n_fixed or n_unfixed:
        try:
            print(
                u"[CopiarDesdeVinculo] Reparación columnas: {0} corregida(s), {1} sin corregir.".format(
                    n_fixed, n_unfixed
                )
            )
        except Exception:
            pass
    return n_fixed, n_unfixed


class _CopyFromLinkFailuresPreprocessor(IFailuresPreprocessor):
    """
    Instrumenta fallos durante CopyElements.

    - Warning: DeleteWarning.
    - Error: borrar elementos fallidos en lote (DeleteElements / ResolveFailure).
    - Nunca DeleteError a ciegas ni Continue con errores pendientes (eso abre
      el diálogo modal de Revit con solo Cancel).
    """

    def __init__(self, records, category_label=u""):
        self.records = records if records is not None else []
        self.category_label = category_label or u""
        self.deleted_via_error = 0

    def set_category(self, label):
        self.category_label = label or u""

    def PreprocessFailures(self, failures_accessor):
        if failures_accessor is None:
            return FailureProcessingResult.Continue

        msgs = list(_iter_failure_msgs(failures_accessor))
        if not msgs:
            return FailureProcessingResult.Continue

        warnings = []
        errors = []
        for f in msgs:
            try:
                sev = f.GetSeverity()
            except Exception:
                continue
            if sev == FailureSeverity.Warning:
                warnings.append(f)
            elif sev == FailureSeverity.Error:
                errors.append(f)

        cleared = False

        for f in warnings:
            desc = _failure_description(f)
            rec = {
                u"severity": u"Warning",
                u"description": desc,
                u"category": self.category_label,
                u"handled": False,
                u"resolution": u"",
            }
            try:
                failures_accessor.DeleteWarning(f)
                rec[u"handled"] = True
                rec[u"resolution"] = u"DeleteWarning"
                cleared = True
            except Exception as ex:
                rec[u"handle_error"] = _as_unicode(ex)
            self.records.append(rec)
            try:
                print(
                    u"[CopiarDesdeVinculo] WARNING [{0}] {1}".format(
                        self.category_label or u"-", desc
                    )
                )
            except Exception:
                pass

        # Errores: borrar todos los ids fallidos de una vez, luego ResolveFailure.
        if errors:
            id_list, n_ids = _collect_failing_element_ids(errors)
            batch_ok = False
            if n_ids > 0:
                try:
                    failures_accessor.DeleteElements(id_list)
                    batch_ok = True
                    cleared = True
                    self.deleted_via_error += n_ids
                    try:
                        print(
                            u"[CopiarDesdeVinculo] ERROR batch DeleteElements: {0} id(s) [{1}]".format(
                                n_ids, self.category_label or u"-"
                            )
                        )
                    except Exception:
                        pass
                except Exception as ex:
                    try:
                        print(
                            u"[CopiarDesdeVinculo] ERROR batch DeleteElements falló: {0}".format(
                                _as_unicode(ex)
                            )
                        )
                    except Exception:
                        pass

            unresolved = 0
            for f in errors:
                desc = _failure_description(f)
                is_zh = _failure_is_zero_height_column(f)
                rec = {
                    u"severity": u"Error",
                    u"description": desc,
                    u"category": self.category_label,
                    u"handled": False,
                    u"resolution": u"",
                }
                resolved = False
                if batch_ok:
                    # Tras borrar ids, resolver cada mensaje de la cola.
                    try:
                        resolved = _try_resolve_delete_elements(failures_accessor, f)
                    except Exception:
                        resolved = False
                    if not resolved:
                        # Elementos ya borrados: marcar manejado para no abrir modal.
                        resolved = True
                    rec[u"resolution"] = u"DeleteElements (batch)"
                else:
                    resolved = _try_resolve_delete_elements(failures_accessor, f)
                    if resolved:
                        rec[u"resolution"] = u"DeleteElements"
                        self.deleted_via_error += 1
                if resolved:
                    rec[u"handled"] = True
                    cleared = True
                    if is_zh:
                        rec[u"resolution"] = u"DeleteElements (column height 0)"
                else:
                    unresolved += 1
                    rec[u"resolution"] = u"unresolved"
                self.records.append(rec)
                try:
                    print(
                        u"[CopiarDesdeVinculo] ERROR [{0}] handled={1} · {2}".format(
                            self.category_label or u"-",
                            rec[u"handled"],
                            desc,
                        )
                    )
                except Exception:
                    pass

            if unresolved > 0:
                # Nunca Continue con errores: abre el diálogo modal (solo Cancel).
                try:
                    print(
                        u"[CopiarDesdeVinculo] Rollback categoría «{0}»: {1} error(es) sin resolver.".format(
                            self.category_label or u"-", unresolved
                        )
                    )
                except Exception:
                    pass
                return FailureProcessingResult.ProceedWithRollBack

        if cleared:
            return FailureProcessingResult.ProceedWithCommit
        return FailureProcessingResult.Continue


def _attach_failures_preprocessor(txn, preprocessor):
    if txn is None or preprocessor is None:
        return False
    try:
        opts = txn.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(preprocessor)
        try:
            opts.SetClearAfterRollback(True)
        except Exception:
            pass
        try:
            opts.SetForcedModalHandling(False)
        except Exception:
            pass
        txn.SetFailureHandlingOptions(opts)
        return True
    except Exception:
        return False


def _dispose_scope(obj):
    if obj is None:
        return
    try:
        obj.Dispose()
    except Exception:
        pass


def _copy_one_category(link_doc, host_doc, transform, opts, id_list):
    """
    Copia una categoría y repara columnas con altura ~0 antes del Commit.

    Returns:
        (n_copiados, n_columnas_corregidas, n_columnas_sin_corregir)
    """
    source_ids = []
    try:
        for i in range(int(id_list.Count)):
            try:
                source_ids.append(id_list.get_Item(i))
            except Exception:
                try:
                    source_ids.append(id_list[i])
                except Exception:
                    pass
    except Exception:
        try:
            source_ids = list(id_list)
        except Exception:
            source_ids = []

    copied = ElementTransformUtils.CopyElements(
        link_doc, id_list, host_doc, transform, opts
    )
    copied_ids = list(copied or [])
    n_fixed, n_unfixed = _repair_zero_height_columns(
        host_doc, link_doc, transform, source_ids, copied_ids
    )
    return len(copied_ids), n_fixed, n_unfixed


def copy_categories_from_link(
    host_doc, link_instance, category_groups, progress_callback=None
):
    """
    Copia al host las instancias indicadas, aplicando la transformación del vínculo.

    Una transacción por categoría (bajo TransactionGroup) para que un warning/error
    no revierta categorías ya copiadas. Los warnings se instrumentan y se silencian.
    """
    result = {
        u"copied": 0,
        u"requested": 0,
        u"errors": [],
        u"warnings": [],
        u"warning_count": 0,
        u"warning_unique_count": 0,
        u"deleted_invalid": 0,
        u"by_category": [],
        u"committed": False,
    }
    if host_doc is None or link_instance is None:
        result[u"errors"].append(u"Documento host o vínculo no válido.")
        return result

    link_doc = None
    try:
        link_doc = link_instance.GetLinkDocument()
    except Exception:
        link_doc = None
    if link_doc is None:
        result[u"errors"].append(u"El vínculo no está cargado.")
        return result

    transform = None
    try:
        transform = link_instance.GetTotalTransform()
    except Exception as ex:
        result[u"errors"].append(
            u"No se pudo leer la transformación del vínculo: {0}".format(
                _as_unicode(ex)
            )
        )
        return result

    work_groups = []
    for grp in category_groups or []:
        ids = grp.get(u"element_ids") or []
        if ids:
            work_groups.append(grp)

    if not work_groups:
        result[u"errors"].append(u"No hay categorías con elementos para copiar.")
        return result

    total_steps = max(1, len(work_groups))
    opts = _copy_options()
    failure_records = []
    preprocessor = _CopyFromLinkFailuresPreprocessor(failure_records)

    tg = TransactionGroup(host_doc, _TRANSACTION_NAME)
    try:
        tg.Start()
    except Exception as ex:
        _dispose_scope(tg)
        result[u"errors"].append(
            u"No se pudo iniciar TransactionGroup: {0}".format(_as_unicode(ex))
        )
        return result

    requested = 0
    copied_total = 0

    try:
        for index, grp in enumerate(work_groups):
            ids = grp.get(u"element_ids") or []
            label = grp.get(u"label") or u"(categoría)"
            if progress_callback is not None:
                try:
                    progress_callback(index + 1, total_steps, label)
                except Exception:
                    pass

            requested += len(ids)
            id_list = _to_id_list(ids)
            cat_result = {
                u"label": label,
                u"requested": int(id_list.Count),
                u"copied": 0,
                u"error": u"",
                u"warnings": [],
                u"warning_count": 0,
            }
            if id_list.Count < 1:
                result[u"by_category"].append(cat_result)
                continue

            preprocessor.set_category(label)
            records_before = len(failure_records)
            tx_name = u"{0} — {1}".format(_TRANSACTION_NAME, label)
            t = Transaction(host_doc, tx_name)
            committed_cat = False
            n_fixed = 0
            n_unfixed = 0
            try:
                _attach_failures_preprocessor(t, preprocessor)
                t.Start()
                n, n_fixed, n_unfixed = _copy_one_category(
                    link_doc, host_doc, transform, opts, id_list
                )
                t.Commit()
                committed_cat = True
                cat_result[u"copied"] = n
                copied_total += n
                if n_fixed > 0:
                    result[u"warnings"].append(
                        u"{0}: {1} columna(s) con altura ~0 corregida(s) (offsets/niveles).".format(
                            label, n_fixed
                        )
                    )
                    result[u"warning_count"] = int(
                        result.get(u"warning_count", 0) or 0
                    ) + n_fixed
                    cat_result[u"warning_count"] = int(
                        cat_result.get(u"warning_count", 0) or 0
                    ) + n_fixed
                if n_unfixed > 0:
                    result[u"warnings"].append(
                        u"{0}: {1} columna(s) con altura ~0 no se pudieron corregir.".format(
                            label, n_unfixed
                        )
                    )
            except Exception as ex:
                try:
                    if t.HasStarted():
                        t.RollBack()
                except Exception:
                    pass
                cat_result[u"error"] = _as_unicode(ex)
                result[u"errors"].append(
                    u"{0}: {1}".format(label, _as_unicode(ex))
                )
                try:
                    print(
                        u"[CopiarDesdeVinculo] ROLLBACK categoría «{0}»: {1}".format(
                            label, _as_unicode(ex)
                        )
                    )
                except Exception:
                    pass
            finally:
                _dispose_scope(t)

            # Asociar warnings/errores capturados en esta categoría.
            cat_unique = []
            cat_seen = set()
            cat_warn_n = 0
            cat_deleted = 0
            for rec in failure_records[records_before:]:
                desc = rec.get(u"description") or u""
                sev = rec.get(u"severity") or u"Warning"
                resolution = _as_unicode(rec.get(u"resolution") or u"")
                line = u"[{0}] {1}".format(sev, desc) if desc else sev
                if resolution:
                    line = u"{0} → {1}".format(line, resolution)
                if u"DeleteElements" in resolution:
                    cat_deleted += 1
                if sev == u"Warning" or (
                    sev == u"Error" and rec.get(u"handled")
                ):
                    cat_warn_n += 1
                    if line not in cat_seen:
                        cat_seen.add(line)
                        cat_unique.append(line)
                    flat = u"{0}: {1}".format(label, desc or u"(sin texto)")
                    if resolution:
                        flat = u"{0} [{1}]".format(flat, resolution)
                    result[u"warnings"].append(flat)
                elif sev == u"Error" and not rec.get(u"handled"):
                    result[u"errors"].append(
                        u"{0}: {1}".format(label, desc or u"(error sin resolver)")
                    )

            cat_result[u"warning_count"] = cat_warn_n
            cat_result[u"warnings"] = cat_unique
            cat_result[u"deleted_invalid"] = cat_deleted
            result[u"warning_count"] = int(result.get(u"warning_count", 0) or 0) + cat_warn_n
            result[u"deleted_invalid"] = int(result.get(u"deleted_invalid", 0) or 0) + cat_deleted

            if not committed_cat and not cat_result[u"error"]:
                cat_result[u"error"] = u"Transacción no confirmada (posibles warnings/errores)."
                result[u"errors"].append(
                    u"{0}: {1}".format(label, cat_result[u"error"])
                )

            result[u"by_category"].append(cat_result)

        result[u"requested"] = requested
        result[u"copied"] = copied_total

        if copied_total > 0:
            try:
                tg.Assimilate()
                result[u"committed"] = True
            except Exception as ex:
                try:
                    tg.RollBack()
                except Exception:
                    pass
                result[u"committed"] = False
                result[u"errors"].append(
                    u"No se pudo asimilar TransactionGroup: {0}".format(
                        _as_unicode(ex)
                    )
                )
        else:
            try:
                tg.RollBack()
            except Exception:
                pass
            result[u"committed"] = False
            if not result[u"errors"]:
                result[u"errors"].append(u"Ningún elemento se copió.")
    except Exception as ex:
        try:
            tg.RollBack()
        except Exception:
            pass
        result[u"committed"] = False
        result[u"errors"].append(_as_unicode(ex))
    finally:
        _dispose_scope(tg)

    # Deduplicar mensajes de warning (muestras) y fijar conteos.
    all_warn_msgs = list(result.get(u"warnings") or [])
    seen = set()
    unique = []
    for w in all_warn_msgs:
        key = _as_unicode(w)
        if key in seen:
            continue
        seen.add(key)
        unique.append(key)
    result[u"warnings"] = unique
    result[u"warning_unique_count"] = len(unique)
    if not result.get(u"warning_count"):
        result[u"warning_count"] = len(all_warn_msgs)

    try:
        print(
            u"[CopiarDesdeVinculo] Fin · copiados={0}/{1} · warnings={2} ({3} únicos) · errores={4} · committed={5}".format(
                result.get(u"copied", 0),
                result.get(u"requested", 0),
                result.get(u"warning_count", 0),
                result.get(u"warning_unique_count", 0),
                len(result.get(u"errors") or []),
                result.get(u"committed", False),
            )
        )
        for w in result.get(u"warnings") or []:
            print(u"[CopiarDesdeVinculo]   WARN  {0}".format(_as_unicode(w)))
        for e in result.get(u"errors") or []:
            print(u"[CopiarDesdeVinculo]   ERR   {0}".format(_as_unicode(e)))
    except Exception:
        pass

    return result


def run_copy_in_transaction(
    host_doc,
    link_instance,
    category_groups,
    transaction_name=None,
    progress_callback=None,
):
    """
    Ejecuta la copia con TransactionGroup + 1 Transaction por categoría.

    ``transaction_name`` se mantiene por compatibilidad (el nombre interno
    sigue siendo ``Arainco: Copiar desde vínculo``).
    """
    _ = transaction_name  # API estable; el prefijo Arainco es fijo.
    return copy_categories_from_link(
        host_doc,
        link_instance,
        category_groups,
        progress_callback=progress_callback,
    )
