# -*- coding: utf-8 -*-
"""
Crear Vistas de Revisión — Plantas estructurales (Piso / Cielo) por nivel.

Requisitos:
- Revit 2024+ / API compatible con ejecución bajo pyRevit (IronPython / CPython según entorno).
- Tipos de vista (ViewFamilyType) con nombre exacto en el proyecto:
  - "Structural Plan (Piso)"
  - "Structural Plan (Cielo)"
- Nombre de vista: "REVISION_" + nombre del nivel + "_PISO" o + "_CIELO" según el tipo creado.
- Rango PISO: corte y tope al nivel asociado +1500 mm; fondo y profundidad Unlimited
  solo si el nivel asociado es el de menor elevación del proyecto; resto de PISO:
  fondo y profundidad al nivel asociado −1500 mm.
- Rango CIELO: corte al nivel asociado +1000 mm; tope al nivel superior +300 mm
  (o asociado +4000 mm si no hay nivel usable); fondo y profundidad de vista al
  nivel asociado +0 mm (la profundidad no puede coincidir con el tope).
- Parámetros de instancia: Clasificacion, Section Filter, Subclasificacion, Zona.
- Una sola transacción «Crear Vistas de Revisión»: SubTransactions por vista (crear+rango),
  Regenerate, SubTransactions por vista (parámetros), Commit final.

Ejecución en pyRevit: el documento activo se toma de __revit__.ActiveUIDocument.
En RPS: definir `doc` (y opcionalmente `uidoc`) antes de llamar a `main()`.
"""

from __future__ import print_function

import clr
import unicodedata

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
try:
    clr.AddReference("System")
except Exception:
    pass

from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    FilterElement,
    Level,
    PlanViewPlane,
    PlanViewRange,
    StorageType,
    SubTransaction,
    Transaction,
    TransactionStatus,
    UnitTypeId,
    UnitUtils,
    View,
    ViewFamily,
    ViewFamilyType,
    ViewPlan,
)

# Nombres exactos de ViewFamilyType en la plantilla/proyecto
VFT_NAME_PISO = "Structural Plan (Piso)"
VFT_NAME_CIELO = "Structural Plan (Cielo)"

# Prefijo de nomenclatura para vistas creadas (underscore tras REVISION)
NAME_PREFIX = "REVISION_"

# Sufijos de nombre según tipo de planta (deben ser únicos en el documento).
SUFFIX_PISO = "_PISO"
SUFFIX_CIELO = "_CIELO"

# Parámetros de instancia al crear cada vista (nombres según aparecen en Revit)
REVISION_VIEW_INSTANCE_PARAMS = (
    ("Clasificacion", "03_REVISION"),
    ("Section Filter", "03_REVISION"),
    ("Subclasificacion", "03_REVISION"),
    ("Zona", "General"),
)


def _normalize_compare_name(value):
    """
    Texto comparable para nombres de tipo: str() (pythonnet puede devolver System.String),
    strip y normalización Unicode NFC (evita caracteres visualmente iguales distintos en binario).
    """
    if value is None:
        return ""
    s = str(value).strip()
    try:
        return unicodedata.normalize("NFC", s)
    except Exception:
        return s


def _view_family_type_display_name(vft):
    """
    Nombre del ViewFamilyType como lo expone Revit a la API.
    En pyRevit 5 (CPython + pythonnet), vft.Name a veces no compara bien con str literales;
    por eso se fuerza str() y se usan parámetros estándar como respaldo.
    """
    if vft is None:
        return ""
    try:
        n = vft.Name
        if n is not None:
            s = _normalize_compare_name(n)
            if s:
                return s
    except Exception:
        pass
    for bip in (
        BuiltInParameter.ALL_MODEL_TYPE_NAME,
        BuiltInParameter.SYMBOL_NAME_PARAM,
    ):
        try:
            p = vft.get_Parameter(bip)
            if p and p.HasValue:
                s = _normalize_compare_name(p.AsString())
                if s:
                    return s
        except Exception:
            continue
    return ""


def _iter_view_family_types(doc):
    """Recorre tipos de vista; WhereElementIsElementType acota a tipos editables (como en UI de tipos)."""
    col = FilteredElementCollector(doc)
    try:
        col = col.WhereElementIsElementType()
    except Exception:
        pass
    for vft in col.OfClass(ViewFamilyType):
        yield vft


def _find_view_family_type_by_name(doc, exact_name, view_families=None):
    """
    Busca ViewFamilyType por nombre de tipo (igual que el desplegable de «New Structural Plan»).

    view_families: iterable de ViewFamily para filtrar (p. ej. solo StructuralPlan).
    Si el primer intento falla, se reintenta sin filtro de familia por si «Cielo» fuera otro family.
    """
    target = _normalize_compare_name(exact_name)
    if not target:
        return None

    def search(filter_families):
        for vft in _iter_view_family_types(doc):
            try:
                if vft is None:
                    continue
                if filter_families is not None:
                    ok = False
                    for vf in filter_families:
                        if vft.ViewFamily == vf:
                            ok = True
                            break
                    if not ok:
                        continue
                if _view_family_type_display_name(vft) == target:
                    return vft
            except Exception:
                continue
        return None

    if view_families is not None:
        found = search(view_families)
        if found is not None:
            return found
    return search(None)


def _structural_plan_type_names_for_diagnostics(doc, limit=40):
    """Lista nombres de tipos StructuralPlan que ve la API (útil si falla la coincidencia)."""
    names = []
    for vft in _iter_view_family_types(doc):
        try:
            if vft is None or vft.ViewFamily != ViewFamily.StructuralPlan:
                continue
            n = _view_family_type_display_name(vft)
            if n:
                names.append(n)
        except Exception:
            continue
    names = sorted(set(names))
    if limit and len(names) > limit:
        return names[:limit], len(names)
    return names, len(names)


def _mm_to_internal(mm_value):
    """
    Convierte milímetros a unidades internas de la API (pies).

    UnitTypeId.Millimeters es un ForgeTypeId (identificador tipado de Autodesk)
    que indica a UnitUtils qué unidad de entrada usar en la conversión. No
    obstante, los desfases del rango de vista (PlanViewRange.SetOffset) siempre
    reciben un double en unidades internas (pies), de ahí esta conversión.
    """
    return UnitUtils.ConvertToInternalUnits(float(mm_value), UnitTypeId.Millimeters)


def _collect_levels_sorted(doc):
    """Todos los Level del documento, ordenados por elevación (ascendente)."""
    levels = list(FilteredElementCollector(doc).OfClass(Level))
    levels.sort(key=lambda lv: lv.Elevation)
    return levels


def _existing_view_names(doc):
    """Conjunto de nombres de vista ya usados (todas las vistas)."""
    names = set()
    for v in FilteredElementCollector(doc).OfClass(View):
        try:
            if v and v.Name:
                names.add(_normalize_compare_name(v.Name))
        except Exception:
            continue
    return names


def _find_parameter_on_element(element, param_name):
    """
    Localiza el parámetro por nombre visible (LookupParameter + iteración Parameters),
    igual que en otros scripts BIMTools cuando LookupParameter falla.
    """
    if element is None or not param_name:
        return None
    cand = str(param_name).strip()
    try:
        p = element.LookupParameter(cand)
        if p is not None:
            return p
    except Exception:
        pass
    target_norm = _normalize_compare_name(cand)
    target_lower = target_norm.lower()
    iterators = []
    try:
        if hasattr(element, "GetOrderedParameters"):
            iterators.append(element.GetOrderedParameters())
    except Exception:
        pass
    try:
        iterators.append(element.Parameters)
    except Exception:
        pass
    for it in iterators:
        if it is None:
            continue
        try:
            for p in it:
                if p is None:
                    continue
                try:
                    dname = p.Definition.Name
                except Exception:
                    continue
                if _normalize_compare_name(dname) == target_norm:
                    return p
                try:
                    if str(dname).strip().lower() == target_lower:
                        return p
                except Exception:
                    pass
        except Exception:
            continue
    return None


def _storage_type_label(p):
    """Etiqueta legible de StorageType del parámetro (para mensajes de error)."""
    if p is None:
        return "?"
    try:
        st = p.StorageType
        if st == StorageType.String:
            return "String"
        if st == StorageType.Double:
            return "Double"
        if st == StorageType.Integer:
            return "Integer"
        if st == StorageType.ElementId:
            return "ElementId"
        return str(int(st))
    except Exception:
        return "?"


def _param_display_value(p):
    """Valor mostrable actual del parámetro (para comprobar si la escritura surtió efecto)."""
    if p is None:
        return ""
    try:
        s = p.AsString()
        if s is not None and str(s).strip():
            return str(s).strip()
    except Exception:
        pass
    try:
        v = p.AsValueString()
        if v is not None and str(v).strip():
            return str(v).strip()
    except Exception:
        pass
    return ""


def _suggest_parameter_names_on_element(element, param_name, max_suggestions=15):
    """
    Lista nombres de parámetros en el elemento que podrían corresponder al buscado
    (subcadena o palabras comunes), útil si el nombre no coincide exactamente.
    """
    if element is None or not param_name:
        return []
    target = _normalize_compare_name(param_name).lower()
    words = [w for w in target.replace(".", " ").split() if len(w) > 2]
    found = []
    iterators = []
    try:
        if hasattr(element, "GetOrderedParameters"):
            iterators.append(element.GetOrderedParameters())
    except Exception:
        pass
    try:
        iterators.append(element.Parameters)
    except Exception:
        pass
    for it in iterators:
        if it is None:
            continue
        try:
            for q in it:
                if q is None:
                    continue
                try:
                    dn = str(q.Definition.Name or "")
                except Exception:
                    continue
                dnl = dn.lower()
                if target in dnl or dnl in target:
                    found.append(dn)
                elif words and all(w in dnl for w in words):
                    found.append(dn)
        except Exception:
            continue
    out = []
    seen = set()
    for n in found:
        k = n.lower()
        if k not in seen:
            seen.add(k)
            out.append(n)
        if len(out) >= max_suggestions:
            break
    return out


def _view_type_has_parameter(element, doc, param_name):
    """
    True si el parámetro existe en el ElementType de la vista pero no se buscó ahí.
    Indica binding a «tipo» en lugar de «instancia».
    """
    if element is None or doc is None or not param_name:
        return False, ""
    try:
        tid = element.GetTypeId()
        if tid is None or tid == ElementId.InvalidElementId:
            return False, ""
        te = doc.GetElement(tid)
        if te is None:
            return False, ""
        p = _find_parameter_on_element(te, param_name)
        if p is None:
            return False, ""
        tname = ""
        try:
            tname = str(te.Name or "")
        except Exception:
            pass
        return True, tname
    except Exception:
        return False, ""


def _list_parameter_names_on_view(view, limit=50):
    """Lista nombres de parámetros de la vista (muestra en diagnóstico)."""
    if view is None or limit <= 0:
        return []
    names = []
    iterators = []
    try:
        if hasattr(view, "GetOrderedParameters"):
            iterators.append(view.GetOrderedParameters())
    except Exception:
        pass
    try:
        iterators.append(view.Parameters)
    except Exception:
        pass
    for it in iterators:
        if it is None:
            continue
        try:
            for q in it:
                if q is None:
                    continue
                try:
                    dn = str(q.Definition.Name or "").strip()
                    if dn:
                        names.append(dn)
                except Exception:
                    continue
        except Exception:
            continue
    try:
        names = sorted(set(names), key=lambda x: x.lower())
    except Exception:
        names = sorted(set(names))
    return names[:limit]


def _find_filter_element_id_by_name(doc, filter_name):
    """
    Id de FilterElement (filtro de vista) por nombre.
    Útil si «Section Filter» u otro parámetro almacena ElementId en lugar de texto.
    """
    if doc is None:
        return ElementId.InvalidElementId
    target = _normalize_compare_name(filter_name)
    if not target:
        return ElementId.InvalidElementId
    try:
        for fe in FilteredElementCollector(doc).OfClass(FilterElement):
            try:
                if fe is None:
                    continue
                if _normalize_compare_name(fe.Name) == target:
                    return fe.Id
            except Exception:
                continue
    except Exception:
        pass
    return ElementId.InvalidElementId


def _regenerate_doc_safe(doc):
    try:
        doc.Regenerate()
    except Exception:
        pass


def _subtransaction_run(doc, fn):
    """
    Ejecuta fn() dentro de una SubTransaction (requiere transacción padre abierta).
    Devuelve el valor de fn(); hace Commit si fn termina sin error, si no RollBack.
    Así no se acumulan ElementId si la subtransacción revierte la creación.
    """
    st = SubTransaction(doc)
    st.Start()
    try:
        out = fn()
        st.Commit()
        return out
    except Exception:
        try:
            st.RollBack()
        except Exception:
            pass
        raise


def _set_parameter_value_robust(element, doc, param_name, value):
    """
    Escribe valor en parámetro de instancia: texto (String / SetValueString),
    o ElementId si el parámetro almacena un FilterElement con ese nombre.

    Devuelve (ok, mensaje). Si ok es False, mensaje explica la causa (solo lectura,
    no encontrado, tipo de almacenamiento, excepción de la API, valor no reflejado).
    """
    if element is None:
        return False, "referencia al elemento es nula"

    p = _find_parameter_on_element(element, param_name)
    if p is None:
        on_type, type_name = _view_type_has_parameter(element, doc, param_name)
        if on_type:
            return False, (
                "no aparece en la instancia de vista pero sí en el tipo de vista «{}». "
                "Project Parameters / Shared Parameters deben estar enlazados a "
                "**Instancia** en la categoría **Vistas** para poder automatizarlo "
                "por vista."
            ).format(type_name or "?")
        sug = _suggest_parameter_names_on_element(element, param_name)
        msg = (
            "no existe parámetro con nombre visible {!r} en esta vista "
            "(LookupParameter + recorrido Parameters)."
        ).format(param_name)
        if sug:
            msg += " Sugerencias de nombres parecidos: {}.".format(", ".join(sug))
        return False, msg

    if p.IsReadOnly:
        return False, "el parámetro existe pero IsReadOnly=True (no editable por API)."

    sval = str(value)
    st_label = _storage_type_label(p)
    reasons = []

    def _check_applied():
        """Confirma que el valor quedó reflejado (tolerancia a mayúsculas / texto parcial)."""
        try:
            got = _param_display_value(p)
            if not sval.strip():
                return True
            if not got:
                return False
            if got == sval or sval in got or got in sval:
                return True
            try:
                if got.strip().lower() == sval.strip().lower():
                    return True
            except Exception:
                pass
            return False
        except Exception:
            return True

    try:
        p.SetValueString(sval)
        if _check_applied():
            return True, ""
        reasons.append(
            "SetValueString no arrojó error pero el valor leído tras asignar es {!r} "
            "(esperado relacionado con {!r}); StorageType={}.".format(
                _param_display_value(p), sval, st_label
            )
        )
    except Exception as ex:
        reasons.append("SetValueString: {}".format(ex))

    try:
        if p.StorageType == StorageType.String:
            p.Set(sval)
            if _check_applied():
                return True, ""
            reasons.append(
                "Set(String) no dejó valor reconocible (leído={!r}).".format(
                    _param_display_value(p)
                )
            )
    except Exception as ex:
        reasons.append("Set(String): {}".format(ex))

    try:
        from System import String as ClrString

        if p.StorageType == StorageType.String:
            p.Set(ClrString(sval))
            if _check_applied():
                return True, ""
            reasons.append(
                "Set(ClrString) no dejó valor reconocible (leído={!r}).".format(
                    _param_display_value(p)
                )
            )
    except Exception as ex:
        reasons.append("Set(ClrString): {}".format(ex))

    if doc is not None and p.StorageType == StorageType.ElementId:
        fid = _find_filter_element_id_by_name(doc, sval)
        if fid is not None and fid != ElementId.InvalidElementId:
            try:
                p.Set(fid)
                if _check_applied() or _param_display_value(p):
                    return True, ""
                reasons.append(
                    "Set(ElementId) ejecutado pero el valor mostrado no coincide "
                    "(leído={!r}).".format(_param_display_value(p))
                )
            except Exception as ex:
                reasons.append("Set(ElementId): {}".format(ex))
        else:
            reasons.append(
                "StorageType=ElementId: no hay FilterElement con nombre exacto {!r}.".format(
                    sval
                )
            )
        try:
            p.SetValueString(sval)
            if _check_applied():
                return True, ""
        except Exception as ex:
            reasons.append("SetValueString (ElementId): {}".format(ex))

    if st_label not in ("String", "ElementId"):
        reasons.append(
            "StorageType={}: no hay ruta de asignación probada para este tipo "
            "(puede ser lista desplegable numérica o unidad).".format(st_label)
        )

    return False, " ".join(reasons) if reasons else "motivo desconocido"


def _apply_revision_instance_parameters(view, doc, report_errors, attach_param_catalog):
    """
    Rellena Clasificacion, Section Filter, Subclasificacion y Zona.
    attach_param_catalog: si True, ante el primer fallo añade muestra de nombres
    de parámetros en esa vista.

    Devuelve True si se añadió la línea de catálogo de parámetros al reporte.
    """
    catalog_line_added = False

    for pname, pval in REVISION_VIEW_INSTANCE_PARAMS:
        ok, detail = _set_parameter_value_robust(view, doc, pname, pval)
        if ok:
            continue
        vname = ""
        try:
            vname = view.Name or ""
        except Exception:
            pass
        report_errors.append(
            'Vista "{}": «{}» (valor deseado {!r}) — {}'.format(
                vname, pname, pval, detail
            )
        )
        if attach_param_catalog and not catalog_line_added:
            sample = _list_parameter_names_on_view(view, limit=45)
            if sample:
                report_errors.append(
                    'Muestra de parámetros en la vista "{}": {}.'.format(
                        vname, "; ".join(sample)
                    )
                )
            catalog_line_added = True
    return catalog_line_added


def _is_lowest_project_level(level, levels_sorted):
    """
    True si el nivel comparte la elevación mínima del proyecto (≤ tolerancia en pies).
    Así todas las vistas PISO ancladas a ese nivel reciben el rango «unlimited» inferior.
    """
    if level is None or not levels_sorted:
        return False
    try:
        min_elev = min(lv.Elevation for lv in levels_sorted)
        tol = 1e-4
        return abs(float(level.Elevation) - float(min_elev)) <= tol
    except Exception:
        return False


# Holgura mínima entre planos (pies internos vía mm). Revit exige
# Top ≥ Cut > Bottom ≥ View Depth; 1 mm evita igualdad rechazada en algunas versiones.
_VIEW_RANGE_MIN_GAP_MM = 1.0


def _plane_elevation(level, offset):
    return float(level.Elevation) + float(offset)


def _set_view_plane(vr, plane, level_id, offset):
    vr.SetLevelId(plane, level_id)
    try:
        is_unlim = level_id == PlanViewRange.Unlimited
    except Exception:
        is_unlim = False
    if is_unlim:
        try:
            vr.SetOffset(plane, 0.0)
        except Exception:
            pass
        return
    vr.SetOffset(plane, float(offset))


def _eid_is_valid(eid):
    if eid is None:
        return False
    try:
        if eid == ElementId.InvalidElementId:
            return False
    except Exception:
        pass
    try:
        return int(eid.Value) > 0
    except Exception:
        try:
            return int(eid.IntegerValue) > 0
        except Exception:
            return False


def _vft_default_template_param(vft):
    if vft is None:
        return None
    try:
        p = vft.get_Parameter(BuiltInParameter.DEFAULT_VIEW_TEMPLATE)
        if p is not None and not p.IsReadOnly:
            return p
    except Exception:
        pass
    try:
        p = vft.LookupParameter(u"View Template applied to new views")
        if p is not None and not p.IsReadOnly:
            return p
    except Exception:
        pass
    return None


def _get_vft_default_template_id(vft):
    try:
        if hasattr(vft, "DefaultTemplateId"):
            tid = vft.DefaultTemplateId
            if _eid_is_valid(tid):
                return tid
    except Exception:
        pass
    p = _vft_default_template_param(vft)
    if p is None:
        return None
    try:
        return p.AsElementId()
    except Exception:
        return None


def _set_vft_default_template_id(vft, eid):
    try:
        if hasattr(vft, "DefaultTemplateId"):
            vft.DefaultTemplateId = eid
            return True
    except Exception:
        pass
    p = _vft_default_template_param(vft)
    if p is None:
        return False
    try:
        p.Set(eid)
        return True
    except Exception:
        return False


def _detach_view_template(view):
    """Quita la plantilla para poder editar el rango de vista."""
    if view is None:
        return
    invalid = ElementId.InvalidElementId
    try:
        tid = view.ViewTemplateId
        if _eid_is_valid(tid):
            view.ViewTemplateId = invalid
    except Exception:
        pass
    try:
        p = view.get_Parameter(BuiltInParameter.VIEW_TEMPLATE_ID)
        if p is not None and not p.IsReadOnly:
            p.Set(invalid)
    except Exception:
        pass


def _create_plan_view(doc, vft, level_id, view_name=None):
    """
    ViewPlan.Create sin la plantilla por defecto del tipo.

    Structural Plan (Cielo) suele tener plantilla RCP; al crear una planta
    estructural (mira hacia abajo) Revit aplica un rango invertido y lanza
    «Plan view range is not valid» ya en Create, antes de SetViewRange.
    """
    old_tid = _get_vft_default_template_id(vft)
    paused = False
    invalid = ElementId.InvalidElementId
    try:
        if _eid_is_valid(old_tid):
            paused = _set_vft_default_template_id(vft, invalid)
        try:
            view = ViewPlan.Create(doc, vft.Id, level_id)
        except Exception as ex:
            label = view_name or u""
            raise Exception(
                u'No se pudo crear la planta «{}»: {}'.format(label, ex)
            )
        _detach_view_template(view)
        return view
    finally:
        if paused and old_tid is not None:
            try:
                _set_vft_default_template_id(vft, old_tid)
            except Exception:
                pass


def _commit_view_range(view, vr):
    try:
        view.SetViewRange(vr)
    except Exception as ex:
        vname = u""
        try:
            vname = view.Name or u""
        except Exception:
            pass
        raise Exception(
            u'Rango de vista no válido en «{}»: {}'.format(vname, ex)
        )


def _level_above(level, levels_sorted, min_delta=None):
    """
    Siguiente nivel por encima de `level` en la lista ordenada por elevación.
    Si `min_delta` (pies internos) se indica, exige Elevation > level + min_delta.
    Devuelve None si `level` es el más alto (o no hay ninguno lo bastante alto).
    """
    if level is None or not levels_sorted:
        return None
    tol = 1e-4  # pies — tolerancia para comparar elevación
    try:
        extra = float(min_delta) if min_delta is not None else 0.0
    except Exception:
        extra = 0.0
    threshold = float(level.Elevation) + max(tol, extra)
    for i, lv in enumerate(levels_sorted):
        if lv.Id == level.Id:
            j = i + 1
            while j < len(levels_sorted):
                try:
                    if float(levels_sorted[j].Elevation) > threshold:
                        return levels_sorted[j]
                except Exception:
                    pass
                j += 1
            return None
    return None


def _level_above_for_cielo(level, levels_sorted, offset_cut, offset_top_on_upper):
    """
    Nivel superior usable como Top de Cielo: Top (nivel + 300 mm) debe quedar
    por encima del Cut (asociado + 1000 mm). Omite niveles intermedios demasiado
    cercanos (p. ej. NPT / cara superior de losa).
    """
    gap = _mm_to_internal(_VIEW_RANGE_MIN_GAP_MM)
    min_delta = float(offset_cut) - float(offset_top_on_upper) + gap
    if min_delta < 1e-4:
        min_delta = 1e-4
    return _level_above(level, levels_sorted, min_delta=min_delta)


def _apply_view_range_piso(
    view,
    assoc_level,
    offset_pos_1500,
    offset_neg_1500,
    is_lowest_project_level,
):
    """
    Vistas tipo Piso (Structural Plan Piso), nivel asociado = GenLevel de la vista.

    Común: Cut = nivel asociado + 1500 mm; Top = Cut + 1 mm (Revit exige Top ≥ Cut).

    Si is_lowest_project_level: Bottom y View Depth = Unlimited (PlanViewRange.Unlimited).
    Si no: mismo criterio anterior al modo unlimited — Bottom y View Depth al nivel
    asociado con offset −1500 mm.
    """
    _detach_view_template(view)
    vr = view.GetViewRange()
    lid = assoc_level.Id
    gap = _mm_to_internal(_VIEW_RANGE_MIN_GAP_MM)
    cut_off = float(offset_pos_1500)

    _set_view_plane(vr, PlanViewPlane.CutPlane, lid, cut_off)
    _set_view_plane(vr, PlanViewPlane.TopClipPlane, lid, cut_off + gap)

    if is_lowest_project_level:
        unlim = PlanViewRange.Unlimited
        # View Depth primero: Bottom Unlimited exige profundidad ≤ fondo.
        _set_view_plane(vr, PlanViewPlane.ViewDepthPlane, unlim, 0.0)
        _set_view_plane(vr, PlanViewPlane.BottomClipPlane, unlim, 0.0)
    else:
        _set_view_plane(vr, PlanViewPlane.BottomClipPlane, lid, offset_neg_1500)
        _set_view_plane(vr, PlanViewPlane.ViewDepthPlane, lid, offset_neg_1500)

    _commit_view_range(view, vr)


def _plan_looks_up(view):
    """True si la planta mira hacia arriba (RCP / CeilingPlan)."""
    try:
        z = float(view.ViewDirection.Z)
        if z > 0.1:
            return True
        if z < -0.1:
            return False
    except Exception:
        pass
    try:
        if str(view.ViewType) == "CeilingPlan":
            return True
    except Exception:
        pass
    try:
        vft = view.Document.GetElement(view.GetTypeId())
        if vft is not None and vft.ViewFamily == ViewFamily.CeilingPlan:
            return True
    except Exception:
        pass
    return False


def _gen_level_id_and_elevation(view, assoc_level):
    try:
        gen = view.GenLevel
        if gen is not None:
            return gen.Id, float(gen.Elevation)
    except Exception:
        pass
    return assoc_level.Id, float(assoc_level.Elevation)


def _cielo_top_offset_from_assoc(
    assoc_elev, level_above, offset_cut, offset_top_300, offset_fallback
):
    """Desfase de Top respecto al nivel asociado (un solo LevelId en el rango)."""
    gap = _mm_to_internal(_VIEW_RANGE_MIN_GAP_MM)
    cut_z = float(assoc_elev) + float(offset_cut)
    if level_above is not None:
        top_z = float(level_above.Elevation) + float(offset_top_300)
        if top_z > cut_z + gap:
            return top_z - float(assoc_elev)
    fb = float(offset_fallback)
    if float(assoc_elev) + fb > cut_z + gap:
        return fb
    return float(offset_cut) + gap


def _write_view_range_same_level(view, level_id, top_off, cut_off, bot_off, depth_off):
    """
    Escribe un rango con los cuatro planos en el mismo nivel.

    Primero abre el Top (desfase alto) para que SetLevelId no deje Top < Cut.
    Si depth_off <= bot_off: planta que mira hacia abajo.
    Si depth_off >= top_off: planta que mira hacia arriba (RCP).
    """
    _detach_view_template(view)
    gap = _mm_to_internal(_VIEW_RANGE_MIN_GAP_MM)
    top_off = float(top_off)
    cut_off = float(cut_off)
    bot_off = float(bot_off)
    depth_off = float(depth_off)
    if top_off < cut_off + gap:
        top_off = cut_off + gap

    vr = view.GetViewRange()
    hi = max(top_off, depth_off, cut_off) + _mm_to_internal(2000.0)
    try:
        top_lid = vr.GetLevelId(PlanViewPlane.TopClipPlane)
        if top_lid != PlanViewRange.Unlimited:
            vr.SetOffset(PlanViewPlane.TopClipPlane, hi)
    except Exception:
        pass

    _set_view_plane(vr, PlanViewPlane.TopClipPlane, level_id, hi)
    _set_view_plane(vr, PlanViewPlane.CutPlane, level_id, cut_off)

    if depth_off <= bot_off + 1e-9:
        try:
            d_lid = vr.GetLevelId(PlanViewPlane.ViewDepthPlane)
            if d_lid != PlanViewRange.Unlimited:
                vr.SetOffset(
                    PlanViewPlane.ViewDepthPlane,
                    min(depth_off, bot_off) - _mm_to_internal(500.0),
                )
        except Exception:
            pass
        _set_view_plane(vr, PlanViewPlane.ViewDepthPlane, level_id, depth_off)
        _set_view_plane(vr, PlanViewPlane.BottomClipPlane, level_id, bot_off)
    else:
        _set_view_plane(vr, PlanViewPlane.BottomClipPlane, level_id, bot_off)
        _set_view_plane(vr, PlanViewPlane.ViewDepthPlane, level_id, depth_off)

    _set_view_plane(vr, PlanViewPlane.TopClipPlane, level_id, top_off)
    _commit_view_range(view, vr)


def _apply_view_range_cielo(
    view,
    assoc_level,
    level_above,
    offset_cut_1000,
    offset_top_depth_300,
    offset_fallback_top_depth_4000,
):
    """
    Vistas tipo Cielo.

    Todos los planos se anclan al nivel asociado (el Top del nivel superior se
    expresa como desfase). Así se evita «Plan view range is not valid» al mezclar
    LevelId distintos.

    Planta hacia abajo: View Depth = Bottom (asociado + 0).
    Planta hacia arriba (RCP): View Depth = Top.
    """
    _detach_view_template(view)
    lid, assoc_elev = _gen_level_id_and_elevation(view, assoc_level)
    cut_off = float(offset_cut_1000)
    top_off = _cielo_top_offset_from_assoc(
        assoc_elev,
        level_above,
        cut_off,
        offset_top_depth_300,
        offset_fallback_top_depth_4000,
    )
    looks_up = _plan_looks_up(view)

    if looks_up:
        attempts = (
            (top_off, cut_off, 0.0, top_off),
            (top_off, cut_off, cut_off, top_off),
            (top_off, cut_off, 0.0, 0.0),
        )
    else:
        attempts = (
            (top_off, cut_off, 0.0, 0.0),
            (top_off, cut_off, 0.0, top_off),
            (top_off, cut_off, cut_off, top_off),
        )

    last_ex = None
    for top_a, cut_a, bot_a, dep_a in attempts:
        try:
            _write_view_range_same_level(view, lid, top_a, cut_a, bot_a, dep_a)
            return
        except Exception as ex:
            last_ex = ex

    vname = u""
    try:
        vname = view.Name or u""
    except Exception:
        pass
    raise Exception(
        u'Rango de vista no válido en «{}» (Cielo, looks_up={}, '
        u'top={:.1f} mm, cut={:.1f} mm): {}'.format(
            vname,
            looks_up,
            top_off * 304.8,
            cut_off * 304.8,
            last_ex,
        )
    )


def main(doc=None):
    """
    Crea por cada nivel dos vistas (Piso y Cielo) con rango de vista configurado.

    Nombres: REVISION_ + nivel + _PISO / _CIELO.
    Si ya existe una vista con ese nombre, no se duplica (se omite).
    Rellena parámetros en SubTransactions dentro de la misma transacción principal.
    """
    try:
        from pyrevit import forms
    except Exception:
        forms = None

    if doc is None:
        try:
            doc = __revit__.ActiveUIDocument.Document  # noqa: F821
        except Exception:
            msg = "No hay documento: pasa `doc` o ejecuta en pyRevit/RPS con contexto válido."
            if forms:
                forms.alert(msg, title="Crear vistas de revisión")
            else:
                print(msg)
            return

    structural_families = (ViewFamily.StructuralPlan,)

    vft_piso = _find_view_family_type_by_name(doc, VFT_NAME_PISO, structural_families)
    vft_cielo = _find_view_family_type_by_name(doc, VFT_NAME_CIELO, structural_families)
    missing = []
    if vft_piso is None:
        missing.append(VFT_NAME_PISO)
    if vft_cielo is None:
        missing.append(VFT_NAME_CIELO)
    if missing:
        sample, total = _structural_plan_type_names_for_diagnostics(doc)
        extra = ""
        if sample:
            lines = "\n".join("  • {}".format(n) for n in sample)
            if total > len(sample):
                lines += "\n  … (+{} más)".format(total - len(sample))
            extra = (
                "\n\nNombres de tipo Structural Plan que devuelve la API en este documento:\n"
                + lines
            )
        msg = (
            "No se encontraron los ViewFamilyType con nombre coincidente:\n\n"
            + "\n".join("  • {}".format(m) for m in missing)
            + "\n\nComprueba mayúsculas, espacios o caracteres raros; compara con la lista inferior."
            + extra
        )
        if forms:
            forms.alert(msg, title="Error — tipos de vista")
        else:
            print(msg)
        return

    levels = _collect_levels_sorted(doc)
    if not levels:
        msg = "No hay niveles (Level) en el proyecto."
        if forms:
            forms.alert(msg, title="Crear vistas de revisión")
        else:
            print(msg)
        return

    # Desfases en unidades internas (pies)
    off_piso_1500 = _mm_to_internal(1500)
    off_piso_neg_1500 = _mm_to_internal(-1500)
    off_cut_cielo = _mm_to_internal(1000)
    off_top_depth = _mm_to_internal(300)
    off_fallback = _mm_to_internal(4000)

    used = _existing_view_names(doc)
    created = []
    skipped = []
    param_issues = []
    # Ids de vistas creadas en esta ejecución (parámetros en subtransacciones posteriores)
    created_view_ids = []

    txn_title = "Arainco: Crear vistas de revisión"
    param_catalog_shown = [False]

    try:
        with Transaction(doc, txn_title) as txn:
            txn.Start()
            try:
                for level in levels:
                    try:
                        lvl_name = str(level.Name) if level.Name is not None else ""
                    except Exception:
                        lvl_name = ""

                    name_piso = NAME_PREFIX + lvl_name + SUFFIX_PISO
                    name_cielo = NAME_PREFIX + lvl_name + SUFFIX_CIELO
                    key_piso = _normalize_compare_name(name_piso)
                    key_cielo = _normalize_compare_name(name_cielo)

                    if key_piso in used:
                        skipped.append(name_piso + " (Piso — ya existía)")
                    else:
                        piso_lowest = _is_lowest_project_level(level, levels)

                        def _make_piso(
                            lev=level,
                            nm=name_piso,
                            lowest=piso_lowest,
                        ):
                            vp = _create_plan_view(doc, vft_piso, lev.Id, nm)
                            vp.Name = nm
                            _apply_view_range_piso(
                                vp,
                                lev,
                                off_piso_1500,
                                off_piso_neg_1500,
                                lowest,
                            )
                            _regenerate_doc_safe(doc)
                            return vp.Id

                        vid_p = _subtransaction_run(doc, _make_piso)
                        created_view_ids.append((vid_p, name_piso))
                        used.add(key_piso)
                        created.append(name_piso)

                    if key_cielo in used:
                        skipped.append(name_cielo + " (Cielo — ya existía)")
                    else:
                        lvl_up = _level_above_for_cielo(
                            level, levels, off_cut_cielo, off_top_depth
                        )

                        def _make_cielo(
                            lev=level,
                            lup=lvl_up,
                            nm=name_cielo,
                        ):
                            vc = _create_plan_view(doc, vft_cielo, lev.Id, nm)
                            vc.Name = nm
                            _apply_view_range_cielo(
                                vc,
                                lev,
                                lup,
                                off_cut_cielo,
                                off_top_depth,
                                off_fallback,
                            )
                            _regenerate_doc_safe(doc)
                            return vc.Id

                        vid_c = _subtransaction_run(doc, _make_cielo)
                        created_view_ids.append((vid_c, name_cielo))
                        used.add(key_cielo)
                        created.append(name_cielo)

                _regenerate_doc_safe(doc)

                for vid, _vname in created_view_ids:

                    def _parametrize_one(eid=vid, pcs=param_catalog_shown):
                        v = doc.GetElement(eid)
                        if v is None:
                            eid_num = None
                            try:
                                eid_num = int(eid.Value)
                            except Exception:
                                try:
                                    eid_num = int(eid.IntegerValue)
                                except Exception:
                                    eid_num = eid
                            param_issues.append(
                                "ElementId={}: GetElement devolvió None "
                                "(no se puede asignar parámetros).".format(eid_num)
                            )
                            return None
                        added = _apply_revision_instance_parameters(
                            v,
                            doc,
                            param_issues,
                            attach_param_catalog=not pcs[0],
                        )
                        if added:
                            pcs[0] = True
                        _regenerate_doc_safe(doc)
                        return True

                    _subtransaction_run(doc, _parametrize_one)

                txn.Commit()
            except Exception:
                if txn.GetStatus() == TransactionStatus.Started:
                    txn.RollBack()
                raise
    except Exception as ex:
        import traceback

        err = "{}\n\n{}".format(ex, traceback.format_exc())
        if forms:
            forms.alert(
                "Error en la transacción (creación y/o parámetros):\n\n{}".format(err),
                title=txn_title,
            )
        else:
            print(err)
        return

    lines = ["Creadas: {}.".format(len(created))]
    if created:
        lines.append("")
        lines.extend(created)
    if skipped:
        lines.append("")
        lines.append("Omitidas (nombre ya usado):")
        lines.extend(skipped)
    if param_issues:
        lines.append("")
        lines.append("Parámetros (detalle abajo en ventana aparte si hay fallos):")
        lines.extend(param_issues)
    summary = "\n".join(lines)
    if forms:
        forms.alert(summary, title="Crear Vistas de Revisión")
    else:
        print(summary)

    if param_issues:
        detail = "\n".join(param_issues)
        if len(detail) > 14000:
            detail = detail[:13900] + "\n\n[... mensaje truncado ...]"
        diag = (
            "No se pudieron escribir uno o más parámetros de instancia.\n\n"
            "QUÉ FALLÓ (por orden):\n\n"
            + detail
            + "\n\n"
            "Comprueba en Gestor de parámetros que estos campos estén enlazados a "
            "**Instancia** en **Vistas (Views)**, no solo al tipo de vista."
        )
        if forms:
            forms.alert(diag, title="Crear vistas — diagnóstico parámetros")
        else:
            print(diag)


if __name__ == "__main__":
    main()
