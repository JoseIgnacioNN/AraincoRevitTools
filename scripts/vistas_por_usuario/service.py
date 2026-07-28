# -*- coding: utf-8 -*-
"""
Servicio Revit — creación de vistas por usuario (02_TRABAJO).

Equivalente Python del grafo Dynamo VistasPorUsuario_script.dyn:
- Valida que no existan vistas con Subclasificacion = abreviación del modelador.
- Duplica plantillas de vista maestras y tipos Detail / Building Section.
- Crea plantas Cielo y Piso por nivel seleccionado.
- Asigna rango de vista, escala y parámetros de instancia.

El código de usuario es la abreviación de personas.json (con puntos, p. ej. J.N.N.),
misma fuente que Siguiente Revisión.

Requisitos: Revit 2024+, tipos de vista con nombres definidos en constants.py.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    ElementParameterFilter,
    FilteredElementCollector,
    OverrideGraphicSettings,
    ParameterFilterElement,
    ParameterFilterRuleFactory,
    ParameterFilterUtilities,
    SharedParameterElement,
    StorageType,
    SubTransaction,
    Transaction,
    TransactionStatus,
    View,
    ViewDuplicateOption,
    ViewFamily,
    ViewPlan,
)

try:
    from System.Collections.Generic import List as DotNetList
except Exception:
    DotNetList = None

from crear_vistas_revision_estructural import (
    _apply_view_range_cielo,
    _apply_view_range_piso,
    _collect_levels_sorted,
    _find_view_family_type_by_name,
    _is_lowest_project_level,
    _level_above,
    _mm_to_internal,
    _normalize_compare_name,
    _regenerate_doc_safe,
    _set_parameter_value_robust,
    _structural_plan_type_names_for_diagnostics,
    _subtransaction_run,
    _view_family_type_display_name,
    _iter_view_family_types,
)

from vistas_por_usuario.constants import (
    CLASIFICACION,
    MASTER_TEMPLATE_SEEDS,
    PARAM_DETAIL_SECTION_TYPE,
    PARAM_PLAN,
    TRANSACTION_TITLE,
    VFT_NAME_CIELO,
    VFT_NAME_DETAIL_FAMILY,
    VFT_NAME_PISO,
    VFT_NAME_SECTION_FAMILY,
    ZONA_DEFAULT,
)


class VistasPorUsuarioError(Exception):
    """Error de validación o configuración previo a la transacción."""


class VistasPorUsuarioRequest(object):
    """Entrada del formulario WPF."""

    def __init__(self, usuario_code, scale, levels, usuario_display=None):
        self.usuario_code = (usuario_code or u"").strip()
        self.usuario_display = usuario_display or self.usuario_code
        self.scale = int(scale)
        self.levels = list(levels or [])


class VistasPorUsuarioResult(object):
    def __init__(self):
        self.created = []
        self.skipped = []
        self.warnings = []
        self.templates_created = []
        self.types_created = []
        self.filters_created = []


def _get_param_string(element, doc, param_name):
    try:
        from crear_vistas_revision_estructural import (
            _find_parameter_on_element,
            _param_display_value,
        )

        p = _find_parameter_on_element(element, param_name)
        if p is None:
            return u""
        return _param_display_value(p)
    except Exception:
        return u""


def _existing_view_names(doc):
    names = set()
    for v in FilteredElementCollector(doc).OfClass(View):
        try:
            if v and v.Name:
                names.add(_normalize_compare_name(v.Name))
        except Exception:
            continue
    return names


def _collect_non_template_views(doc):
    views = []
    for v in FilteredElementCollector(doc).OfClass(View):
        try:
            if v is None:
                continue
            if getattr(v, "IsTemplate", False):
                continue
            views.append(v)
        except Exception:
            continue
    return views


def validate_user_views_not_exist(doc, usuario_code):
    """
    True si puede continuar (no hay vistas de ese usuario).
    Equivalente al filtro Subclasificacion == usuario en el .dyn.
    """
    code = _normalize_compare_name(usuario_code)
    if not code:
        return False, u"Seleccione un usuario válido."

    matches = []
    for v in _collect_non_template_views(doc):
        sub = _normalize_compare_name(_get_param_string(v, doc, u"Subclasificacion"))
        if sub == code:
            matches.append(v)

    if matches:
        return False, (
            u"Las vistas de esta Usuario ya fueron generadas. "
            u"Para vistas adicionales, duplicar de las ya existentes."
        )
    return True, None


def _norm_tpl_name(value):
    """Nombre comparable: NFC + strip + minúsculas + espacios colapsados."""
    s = _normalize_compare_name(value)
    if not s:
        return u""
    try:
        parts = s.lower().split()
        return u" ".join(parts)
    except Exception:
        return s.lower()


def _view_name_candidates(view):
    """Posibles nombres visibles (Name / Title / VIEW_NAME)."""
    names = []
    if view is None:
        return names
    for getter in (
        lambda: view.Name,
        lambda: getattr(view, "Title", None),
    ):
        try:
            raw = getter()
            n = _normalize_compare_name(raw)
            if n and n not in names:
                names.append(n)
        except Exception:
            pass
    try:
        from Autodesk.Revit.DB import BuiltInParameter

        p = view.get_Parameter(BuiltInParameter.VIEW_NAME)
        if p is not None:
            n = _normalize_compare_name(p.AsString())
            if n and n not in names:
                names.append(n)
    except Exception:
        pass
    return names


def _view_display_name(view):
    cands = _view_name_candidates(view)
    return cands[0] if cands else u""


def _is_view_template(view):
    """
    Detecta plantilla. En CPython/pythonnet IsTemplate a veces no se comporta
    como bool nativo: comparar de varias formas.
    """
    if view is None:
        return False
    try:
        flag = view.IsTemplate
    except Exception:
        return False
    try:
        if flag is True or flag is False:
            return bool(flag)
    except Exception:
        pass
    try:
        import System

        return bool(System.Convert.ToBoolean(flag))
    except Exception:
        pass
    try:
        return str(flag).strip().lower() in (u"true", u"1", u"yes")
    except Exception:
        return False


def _element_id_int(eid):
    if eid is None:
        return None
    try:
        return int(eid.Value)
    except Exception:
        pass
    try:
        return int(eid.IntegerValue)
    except Exception:
        return None


def _iter_all_views(doc):
    """Todas las View del documento (OfClass + OST_Views)."""
    seen = set()
    collectors = []
    try:
        collectors.append(FilteredElementCollector(doc).OfClass(View))
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import BuiltInCategory

        collectors.append(
            FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_Views)
            .WhereElementIsNotElementType()
        )
    except Exception:
        pass
    for col in collectors:
        try:
            for v in col:
                if v is None:
                    continue
                try:
                    key = v.Id
                except Exception:
                    continue
                kid = _element_id_int(key)
                if kid is not None:
                    if kid in seen:
                        continue
                    seen.add(kid)
                yield v
        except Exception:
            continue


def _iter_view_templates(doc):
    for v in _iter_all_views(doc):
        if _is_view_template(v):
            yield v


def _view_matches_seed(view, seed_norm):
    if not seed_norm:
        return False
    for n in _view_name_candidates(view):
        if _norm_tpl_name(n) == seed_norm:
            return True
    return False


def _find_view_template_by_exact_name(doc, exact_name):
    """
    Localiza plantilla por nombre.
    1) Preferir IsTemplate=True
    2) Si el nombre coincide exactamente, aceptar aunque IsTemplate falle
       (quirk pythonnet / API): el dialogo de Revit lista esas semillas.
    """
    target = _norm_tpl_name(exact_name)
    if not target:
        return None
    name_hits = []
    for v in _iter_all_views(doc):
        if not _view_matches_seed(v, target):
            continue
        if _is_view_template(v):
            return v
        name_hits.append(v)
    if name_hits:
        return name_hits[0]
    return None


def _find_view_template_master(doc, seed_name):
    """
    Plantilla semilla del proyecto.
    Equivalente Dynamo: Views -> Name == seed (IsTemplate preferido).

    Returns:
        (view_or_None, diagnostic_message_or_None)
    """
    target = _norm_tpl_name(seed_name)
    if not target:
        return None, u"nombre semilla vacío"

    hit = _find_view_template_by_exact_name(doc, seed_name)
    if hit is not None:
        return hit, None

    fuzzy = []
    for v in _iter_all_views(doc):
        for n in _view_name_candidates(v):
            vn = _norm_tpl_name(n)
            if not vn:
                continue
            if target in vn or vn in target:
                label = n
                if _is_view_template(v):
                    label = u"{} [template]".format(n)
                if label not in fuzzy:
                    fuzzy.append(label)

    if fuzzy:
        shown = u", ".join(u"\u00ab{}\u00bb".format(n) for n in fuzzy[:10])
        return None, (
            u"No hay coincidencia exacta con \u00ab{}\u00bb. "
            u"Candidatos: {}"
        ).format(seed_name, shown)

    # Diagnóstico: cuantas vistas / templates ve la API
    n_views = 0
    n_tpl = 0
    sample = []
    for v in _iter_all_views(doc):
        n_views += 1
        if _is_view_template(v):
            n_tpl += 1
            if len(sample) < 15:
                sample.append(_view_display_name(v) or u"?")
    extra = u"API ve {} vistas, {} con IsTemplate=True.".format(n_views, n_tpl)
    if sample:
        extra += u" Ejemplos template: " + u", ".join(sample)
    return None, (
        u"no hay vista/plantilla llamada \u00ab{}\u00bb. {}"
    ).format(seed_name, extra)


def _template_seed_names_for_diagnostics(doc, limit=20):
    names = []
    # Preferir IsTemplate; si ninguno, listar nombres que parezcan semillas
    for v in _iter_view_templates(doc):
        try:
            names.append(_view_display_name(v))
        except Exception:
            continue
    if not names:
        keys = (
            u"architectural",
            u"structural",
            u"foundation",
            u"detail",
            u"section",
            u"ceiling",
            u"reflected",
        )
        for v in _iter_all_views(doc):
            dn = _view_display_name(v)
            low = dn.lower()
            if any(k in low for k in keys):
                mark = u" [T]" if _is_view_template(v) else u""
                names.append(dn + mark)
    names.sort(key=lambda s: s.lower())
    return names[:limit], len(names)


def _find_view_by_name(doc, exact_name):
    target = _norm_tpl_name(exact_name)
    for v in _iter_all_views(doc):
        if _view_matches_seed(v, target):
            return v
    return None


def _copy_view_template_via_copy_elements(doc, source_view, new_name):
    """
    Alternativa a View.Duplicate: ElementTransformUtils.CopyElements.
    En algunos proyectos Duplicate lanza «View cannot be duplicated» sobre plantillas.
    """
    from Autodesk.Revit.DB import (
        CopyPasteOptions,
        ElementId,
        ElementTransformUtils,
        Transform,
        XYZ,
    )
    from System.Collections.Generic import List

    ids = List[ElementId]()
    ids.Add(source_view.Id)
    errors = []

    # Misma documentación: mismo documento + traslación nula
    try:
        copied = ElementTransformUtils.CopyElements(doc, ids, XYZ.Zero)
        if copied is not None and copied.Count > 0:
            dup = doc.GetElement(list(copied)[0])
            if dup is not None:
                dup.Name = new_name
                _regenerate_doc_safe(doc)
                return dup, None
        errors.append(u"CopyElements(XYZ) sin resultado")
    except Exception as ex:
        errors.append(u"CopyElements(XYZ): {}".format(ex))

    try:
        opts = CopyPasteOptions()
        copied = ElementTransformUtils.CopyElements(
            doc, ids, doc, Transform.Identity, opts
        )
        if copied is not None and copied.Count > 0:
            dup = doc.GetElement(list(copied)[0])
            if dup is not None:
                dup.Name = new_name
                _regenerate_doc_safe(doc)
                return dup, None
        errors.append(u"CopyElements(doc→doc) sin resultado")
    except Exception as ex:
        errors.append(u"CopyElements(doc→doc): {}".format(ex))

    return None, u"; ".join(errors)


def _duplicate_view_template(doc, source_view, new_name):
    """
    Equivalente a ViewTemplates.Duplicate del .dyn.
    Orden: Duplicate → CopyElements → semilla «…_00».
    """
    if source_view is None:
        return None, u"plantilla origen no encontrada"

    existing = _find_view_template_by_exact_name(doc, new_name)
    if existing is not None:
        return existing, None

    collision = _find_view_by_name(doc, new_name)
    if collision is not None and collision.Id != source_view.Id:
        return None, (
            u"Ya existe una vista llamada \u00ab{}\u00bb "
            u"(IsTemplate={})"
        ).format(new_name, _is_view_template(collision))

    sources = [source_view]
    try:
        seed_name = _view_display_name(source_view)
        alt = _find_view_template_by_exact_name(doc, seed_name + u"_00")
        if alt is not None and alt.Id != source_view.Id:
            sources.append(alt)
    except Exception:
        pass

    errors = []
    for src in sources:
        src_label = _view_display_name(src)

        # 1) View.Duplicate (como Dynamo)
        try:
            dup_id = src.Duplicate(ViewDuplicateOption.Duplicate)
            if dup_id is not None and _element_id_int(dup_id) not in (None, -1):
                dup = doc.GetElement(dup_id)
                if dup is not None:
                    try:
                        dup.Name = new_name
                        _regenerate_doc_safe(doc)
                        renamed = _find_view_template_by_exact_name(doc, new_name)
                        return (renamed if renamed is not None else dup), None
                    except Exception as ex:
                        errors.append(
                            u"Duplicate+rename «{}»: {}".format(src_label, ex)
                        )
                        try:
                            doc.Delete(dup_id)
                        except Exception:
                            pass
                else:
                    errors.append(u"Duplicate «{}»: GetElement None".format(src_label))
            else:
                errors.append(u"Duplicate «{}»: Id inválido".format(src_label))
        except Exception as ex:
            errors.append(u"Duplicate «{}»: {}".format(src_label, ex))

        # 2) CopyElements (mismo documento)
        dup, err = _copy_view_template_via_copy_elements(doc, src, new_name)
        if dup is not None:
            renamed = _find_view_template_by_exact_name(doc, new_name)
            return (renamed if renamed is not None else dup), None
        if err:
            errors.append(u"CopyElements «{}»: {}".format(src_label, err))

    return None, u" | ".join(errors) if errors else u"no se pudo copiar la plantilla"


def _find_vft_by_family_and_name_hint(doc, view_families, name_hint):
    hint = _normalize_compare_name(name_hint).lower()
    for vft in _iter_view_family_types(doc):
        try:
            if vft is None:
                continue
            ok_family = False
            for vf in view_families:
                if vft.ViewFamily == vf:
                    ok_family = True
                    break
            if not ok_family:
                continue
            dn = _view_family_type_display_name(vft).lower()
            if hint in dn:
                return vft
        except Exception:
            continue
    return None


def _find_vft_by_exact_name(doc, exact_name, view_families=None):
    return _find_view_family_type_by_name(doc, exact_name, view_families)


def _get_or_duplicate_vft(doc, source_vft, new_name):
    if source_vft is None:
        return None, u"tipo origen no encontrado"
    existing = _find_vft_by_exact_name(doc, new_name, None)
    if existing is not None:
        return existing, None
    try:
        dup = source_vft.Duplicate(new_name)
        _regenerate_doc_safe(doc)
        return dup, None
    except Exception as ex:
        return None, str(ex)


def _set_associated_level(view, level):
    if view is None or level is None:
        return False
    try:
        p = view.LookupParameter(u"Associated Level")
        if p is not None and not p.IsReadOnly:
            p.Set(level.Id)
            return True
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import BuiltInParameter

        p = view.get_Parameter(BuiltInParameter.PLAN_VIEW_LEVEL_ID)
        if p is not None and not p.IsReadOnly:
            p.Set(level.Id)
            return True
    except Exception:
        pass
    return False


def _apply_param_list(element, doc, param_list, usuario_code, report):
    for pname, pval in param_list:
        if pval is None:
            val = usuario_code
        else:
            val = pval
        ok, detail = _set_parameter_value_robust(element, doc, pname, val)
        if not ok:
            label = u""
            try:
                label = element.Name or u""
            except Exception:
                pass
            report.append(u'«{}» en "{}": {}'.format(pname, label, detail))


def _sheet_titles_for_levels(levels):
    if len(levels) == 2:
        return [u"PLANTA FUNDACION", u"PLANTA RADIER"]
    titles = []
    for lv in levels:
        try:
            titles.append(str(lv.Name or u""))
        except Exception:
            titles.append(u"")
    return titles


def _template_names(usuario_code):
    return {
        u"cielo": u"{}_{}_STRUCTURAL PLAN (CIELO)".format(CLASIFICACION, usuario_code),
        u"piso": u"{}_{}_STRUCTURAL PLAN (PISO)".format(CLASIFICACION, usuario_code),
        u"detail": u"{}_{}_DETAIL".format(CLASIFICACION, usuario_code),
        u"section": u"{}_{}_BUILDING SECTION".format(CLASIFICACION, usuario_code),
    }


def _resolve_view_family_types(doc, result):
    structural = (ViewFamily.StructuralPlan,)
    vft_cielo = _find_vft_by_exact_name(doc, VFT_NAME_CIELO, structural)
    vft_piso = _find_vft_by_exact_name(doc, VFT_NAME_PISO, structural)
    missing = []
    if vft_cielo is None:
        missing.append(VFT_NAME_CIELO)
    if vft_piso is None:
        missing.append(VFT_NAME_PISO)
    if missing:
        sample, total = _structural_plan_type_names_for_diagnostics(doc)
        extra = u""
        if sample:
            lines = u"\n".join(u"  • {}".format(n) for n in sample)
            if total > len(sample):
                lines += u"\n  … (+{} más)".format(total - len(sample))
            extra = (
                u"\n\nTipos Structural Plan en este documento:\n" + lines
            )
        raise VistasPorUsuarioError(
            u"No se encontraron ViewFamilyType:\n"
            + u"\n".join(u"  • {}".format(m) for m in missing)
            + extra
        )
    vft_detail = _find_vft_by_family_and_name_hint(
        doc, (ViewFamily.Detail, ViewFamily.Drafting), VFT_NAME_DETAIL_FAMILY
    )
    if vft_detail is None:
        vft_detail = _find_vft_by_family_and_name_hint(doc, (ViewFamily.Detail,), u"detail")
    vft_section = _find_vft_by_family_and_name_hint(
        doc, (ViewFamily.Section, ViewFamily.Elevation), VFT_NAME_SECTION_FAMILY
    )
    if vft_section is None:
        vft_section = _find_vft_by_family_and_name_hint(
            doc, (ViewFamily.Section,), u"section"
        )
    if vft_detail is None:
        result.warnings.append(
            u"No se encontró tipo base Detail; se omitirá duplicar Detail ({})".format(
                VFT_NAME_DETAIL_FAMILY
            )
        )
    if vft_section is None:
        result.warnings.append(
            u"No se encontró tipo base Building Section; se omitirá duplicar sección."
        )
    return vft_cielo, vft_piso, vft_detail, vft_section


PARAM_SECTION_FILTER = u"Section Filter"


def _find_parameter_filter_by_name(doc, filter_name):
    target = _normalize_compare_name(filter_name)
    for f in FilteredElementCollector(doc).OfClass(ParameterFilterElement):
        try:
            if _normalize_compare_name(f.Name) == target:
                return f
        except Exception:
            continue
    return None


def _find_section_filter_param_id(doc, cat_ids):
    """ElementId del parámetro «Section Filter» (shared / filtrable / muestra View)."""
    # SharedParameterElement
    try:
        for spe in FilteredElementCollector(doc).OfClass(SharedParameterElement):
            try:
                if spe and _normalize_compare_name(spe.Name) == _normalize_compare_name(
                    PARAM_SECTION_FILTER
                ):
                    return spe.Id
            except Exception:
                continue
    except Exception:
        pass

    # Filtrables en las categorías del filtro
    try:
        for p_id in ParameterFilterUtilities.GetFilterableParametersInCommon(doc, cat_ids):
            p_elem = doc.GetElement(p_id)
            if p_elem is None:
                continue
            try:
                if _normalize_compare_name(p_elem.Name) == _normalize_compare_name(
                    PARAM_SECTION_FILTER
                ):
                    return p_id
            except Exception:
                continue
    except Exception:
        pass

    # Muestra en una View (como Dynamo ParameterByName sobre View)
    try:
        for v in FilteredElementCollector(doc).OfClass(View):
            try:
                if v is None or getattr(v, "IsTemplate", False):
                    continue
                p = v.LookupParameter(PARAM_SECTION_FILTER)
                if p is not None:
                    return p.Id
            except Exception:
                continue
    except Exception:
        pass
    return None


def _create_not_equals_string_rule(param_id, value_text):
    """CreateNotEqualsRule compatible Revit 2023+ y overload antiguo."""
    try:
        return ParameterFilterRuleFactory.CreateNotEqualsRule(param_id, value_text)
    except TypeError:
        return ParameterFilterRuleFactory.CreateNotEqualsRule(
            param_id, value_text, True
        )


def _ensure_user_section_filter(doc, usuario_code, warnings):
    """
    Equivalente Dynamo ParameterFilterElement.ByRules:
    - name = código usuario
    - categoría Sections
    - regla Section Filter NotEquals usuario
    """
    if not usuario_code:
        return None

    if DotNetList is None:
        warnings.append(u"No se pudo crear filtro: List[ElementId] no disponible.")
        return None

    cat_ids = DotNetList[ElementId]()
    cat_ids.Add(ElementId(BuiltInCategory.OST_Sections))

    param_id = _find_section_filter_param_id(doc, cat_ids)
    if param_id is None:
        # Algunas plantillas tienen Section Filter solo en Views
        cat_ids2 = DotNetList[ElementId]()
        cat_ids2.Add(ElementId(BuiltInCategory.OST_Views))
        param_id = _find_section_filter_param_id(doc, cat_ids2)
        if param_id is not None:
            cat_ids = cat_ids2
        else:
            warnings.append(
                u"No se encontró el parámetro «{}» para crear el filtro «{}»."
                .format(PARAM_SECTION_FILTER, usuario_code)
            )
            return None

    filter_name = usuario_code
    rule = _create_not_equals_string_rule(param_id, usuario_code)
    elem_filter = ElementParameterFilter(rule)

    existing = _find_parameter_filter_by_name(doc, filter_name)
    if existing is not None:
        try:
            existing.SetElementFilter(elem_filter)
            existing.SetCategories(cat_ids)
        except Exception as ex:
            warnings.append(
                u"Filtro «{}» existente no actualizable: {}".format(filter_name, ex)
            )
        return existing

    try:
        return ParameterFilterElement.Create(doc, filter_name, cat_ids, elem_filter)
    except Exception as ex:
        warnings.append(
            u"No se pudo crear ParameterFilter «{}»: {}".format(filter_name, ex)
        )
        return None


def _apply_filter_to_view_template(view_template, filter_elem, warnings):
    """
    Equivalente Dynamo View.SetFilterOverrides:
    añade el filtro a la plantilla y oculta elementos que cumplen la regla
    (Section Filter ≠ usuario), igual que el patrón de filtro por eje.
    """
    if view_template is None or filter_elem is None:
        return False
    fid = filter_elem.Id
    label = _view_display_name(view_template) or u"?"

    try:
        already = False
        try:
            already = view_template.GetFilters().Contains(fid)
        except Exception:
            already = fid in list(view_template.GetFilters())
        if not already:
            view_template.AddFilter(fid)
    except Exception as ex:
        warnings.append(
            u"Plantilla «{}»: no se pudo AddFilter: {}".format(label, ex)
        )
        return False

    try:
        view_template.SetIsFilterEnabled(fid, True)
    except Exception:
        pass

    # Ocultar elementos con Section Filter ≠ usuario
    try:
        view_template.SetFilterVisibility(fid, False)
    except Exception as ex:
        warnings.append(
            u"Plantilla «{}»: SetFilterVisibility: {}".format(label, ex)
        )

    try:
        view_template.SetFilterOverrides(fid, OverrideGraphicSettings())
    except Exception:
        pass

    return True


def _apply_user_filters_to_templates(doc, usuario_code, templates, tpl_names, result):
    """Crea filtro por usuario y lo asigna a plantillas 02_TRABAJO_{usuario}_*."""
    filt = _ensure_user_section_filter(doc, usuario_code, result.warnings)
    if filt is None:
        return

    applied = []
    for key, tpl in templates.items():
        if tpl is None:
            continue
        expected = _norm_tpl_name(tpl_names.get(key) or u"")
        actual = _norm_tpl_name(_view_display_name(tpl))
        # No modificar semillas compartidas (fallback sin copia por usuario)
        if not expected or actual != expected:
            result.warnings.append(
                u"Filtro «{}» no aplicado a «{}» "
                u"(no es la plantilla por-usuario «{}»)."
                .format(usuario_code, _view_display_name(tpl), tpl_names.get(key))
            )
            continue
        if _apply_filter_to_view_template(tpl, filt, result.warnings):
            applied.append(_view_display_name(tpl))

    if applied:
        result.filters_created.append(usuario_code)
        result.filters_created.extend(
            u"  → {}".format(n) for n in applied
        )


def _set_default_view_template_on_type(view_family_type, template_view, warnings):
    """
    Asigna «View Template applied to new views» en el ViewFamilyType
    (BuiltInParameter.DEFAULT_VIEW_TEMPLATE), como en Type Properties de Revit.
    """
    if view_family_type is None or template_view is None:
        return False

    type_label = u"?"
    tpl_label = u"?"
    try:
        type_label = _view_family_type_display_name(view_family_type) or u"?"
    except Exception:
        pass
    try:
        tpl_label = _view_display_name(template_view) or u"?"
    except Exception:
        pass

    # 1) Propiedad API si existe (algunas versiones)
    try:
        if hasattr(view_family_type, "DefaultTemplateId"):
            view_family_type.DefaultTemplateId = template_view.Id
            return True
    except Exception:
        pass

    # 2) BuiltInParameter.DEFAULT_VIEW_TEMPLATE
    try:
        from Autodesk.Revit.DB import BuiltInParameter

        p = view_family_type.get_Parameter(BuiltInParameter.DEFAULT_VIEW_TEMPLATE)
        if p is not None and not p.IsReadOnly:
            p.Set(template_view.Id)
            return True
    except Exception as ex:
        warnings.append(
            u'Tipo «{}» — DEFAULT_VIEW_TEMPLATE: {}'.format(type_label, ex)
        )

    # 3) Por nombre visible (UI)
    try:
        p = view_family_type.LookupParameter(u"View Template applied to new views")
        if p is not None and not p.IsReadOnly:
            p.Set(template_view.Id)
            return True
    except Exception as ex:
        warnings.append(
            u'Tipo «{}» — LookupParameter View Template: {}'.format(type_label, ex)
        )

    warnings.append(
        u'No se pudo asignar plantilla «{}» como default en tipo «{}» '
        u'(View Template applied to new views).'
        .format(tpl_label, type_label)
    )
    return False


def _duplicate_templates_and_types(doc, usuario_code, vft_detail, vft_section, result):
    """
    Como el .dyn:
    1) Busca plantilla semilla por nombre exacto (MASTER_TEMPLATE_SEEDS).
    2) ViewTemplates.Duplicate → 02_TRABAJO_{usuario}_…
    3) Asigna parámetros de clasificación a la plantilla nueva.
    4) Duplica tipos Detail / Building Section.
    """
    tpl_names = _template_names(usuario_code)
    templates = {}
    sample_names, total_tpl = _template_seed_names_for_diagnostics(doc)
    missing_plan_seeds = []
    plan_fail_details = []

    for key, seed_name in MASTER_TEMPLATE_SEEDS.items():
        master, diag = _find_view_template_master(doc, seed_name)
        if master is None:
            extra = u""
            if diag:
                extra = u" Motivo: {}.".format(diag)
            if sample_names:
                lines = u"\n".join(u"  • {}".format(n) for n in sample_names)
                if total_tpl > len(sample_names):
                    lines += u"\n  … (+{} más)".format(total_tpl - len(sample_names))
                extra += u"\nVistas/plantillas detectadas (muestra):\n" + lines
            msg = (
                u"Plantilla semilla no encontrada: «{}» "
                u"(necesaria para crear «{}»).{}"
            ).format(seed_name, tpl_names[key], extra)
            result.warnings.append(msg)
            if key in (u"cielo", u"piso"):
                missing_plan_seeds.append(seed_name)
                plan_fail_details.append(
                    u"«{}»: {}".format(seed_name, diag or u"no encontrada")
                )
            continue

        dup, err = _duplicate_view_template(doc, master, tpl_names[key])
        if dup is None:
            result.warnings.append(
                u"No se pudo duplicar plantilla «{}» desde «{}»: {}. "
                u"Se asignará la plantilla semilla compartida a las vistas."
                .format(tpl_names[key], seed_name, err)
            )
            # Respaldo: usar la semilla maestra (como Dynamo no pudo copiar)
            templates[key] = master
            continue

        # Parámetros de instancia en la plantilla (mismo set que el .dyn)
        if key in (u"cielo", u"piso"):
            _apply_param_list(dup, doc, PARAM_PLAN, usuario_code, result.warnings)
        else:
            _apply_param_list(
                dup, doc, PARAM_DETAIL_SECTION_TYPE, usuario_code, result.warnings
            )

        templates[key] = dup
        result.templates_created.append(tpl_names[key])

    if missing_plan_seeds and not (
        templates.get(u"cielo") and templates.get(u"piso")
    ):
        detail = u""
        if plan_fail_details:
            detail = u"\n\nDetalle:\n" + u"\n".join(
                u"  • {}".format(d) for d in plan_fail_details
            )
        if sample_names:
            detail += u"\n\nDetectadas por la API (muestra):\n" + u"\n".join(
                u"  • {}".format(n) for n in sample_names[:15]
            )
        raise VistasPorUsuarioError(
            u"No se pudieron crear las plantillas de planta (Cielo/Piso).\n"
            u"Semillas requeridas:\n"
            + u"\n".join(u"  • {}".format(s) for s in missing_plan_seeds)
            + detail
        )

    detail_type_name = u"Detail ({})".format(usuario_code)
    section_type_name = u"Building Section ({})".format(usuario_code)

    if vft_detail is not None:
        dup_d, err = _get_or_duplicate_vft(doc, vft_detail, detail_type_name)
        if dup_d is None:
            result.warnings.append(u"Detail type: {}".format(err))
        else:
            _apply_param_list(dup_d, doc, PARAM_DETAIL_SECTION_TYPE, usuario_code, result.warnings)
            # Default template en el tipo (UI: View Template applied to new views)
            tpl_detail = templates.get(u"detail")
            if tpl_detail is not None:
                _set_default_view_template_on_type(dup_d, tpl_detail, result.warnings)
            else:
                result.warnings.append(
                    u"Tipo «{}»: no hay plantilla DETAIL para asignar por defecto."
                    .format(detail_type_name)
                )
            result.types_created.append(detail_type_name)

    if vft_section is not None:
        dup_s, err = _get_or_duplicate_vft(doc, vft_section, section_type_name)
        if dup_s is None:
            result.warnings.append(u"Building Section type: {}".format(err))
        else:
            _apply_param_list(dup_s, doc, PARAM_DETAIL_SECTION_TYPE, usuario_code, result.warnings)
            tpl_section = templates.get(u"section")
            if tpl_section is not None:
                _set_default_view_template_on_type(dup_s, tpl_section, result.warnings)
            else:
                result.warnings.append(
                    u"Tipo «{}»: no hay plantilla BUILDING SECTION para asignar por defecto."
                    .format(section_type_name)
                )
            result.types_created.append(section_type_name)

    # Filtros Dynamo: ParameterFilter (Section Filter ≠ usuario) en plantillas
    _apply_user_filters_to_templates(doc, usuario_code, templates, tpl_names, result)

    return templates


def _apply_plan_view_metadata(view, doc, level, usuario_code, scale, sheet_title, template_view, warnings):
    try:
        view.Scale = int(scale)
    except Exception:
        warnings.append(u'No se pudo asignar escala {} a «{}»'.format(scale, view.Name))

    _apply_param_list(view, doc, PARAM_PLAN, usuario_code, warnings)

    if not _set_associated_level(view, level):
        warnings.append(
            u'Vista «{}» — no se pudo asignar Associated Level.'.format(view.Name)
        )

    ok_title, detail_title = _set_parameter_value_robust(
        view, doc, u"Title on Sheet", sheet_title
    )
    if not ok_title:
        warnings.append(
            u'Vista «{}» — Title on Sheet: {}'.format(view.Name, detail_title)
        )

    # Asignación de plantilla (API ViewTemplateId + BIP si existe)
    if template_view is not None:
        applied = False
        try:
            if not _is_view_template(view):
                view.ViewTemplateId = template_view.Id
                applied = True
        except Exception as ex:
            warnings.append(
                u'ViewTemplateId falló en «{}»: {}'.format(view.Name, ex)
            )
        if not applied:
            try:
                from Autodesk.Revit.DB import BuiltInParameter

                p = view.get_Parameter(BuiltInParameter.VIEW_TEMPLATE_ID)
                if p is not None and not p.IsReadOnly:
                    p.Set(template_view.Id)
                    applied = True
            except Exception:
                pass
        if not applied:
            warnings.append(
                u'No se pudo aplicar plantilla «{}» a «{}».'
                .format(_view_display_name(template_view), view.Name)
            )


def create_user_views(doc, request):
    """
    Ejecuta la creación completa. Devuelve VistasPorUsuarioResult.
    Lanza VistasPorUsuarioError si falla validación previa.
    """
    if not request.usuario_code:
        raise VistasPorUsuarioError(u"Seleccione un usuario.")

    ok, msg = validate_user_views_not_exist(doc, request.usuario_code)
    if not ok:
        raise VistasPorUsuarioError(msg)

    levels_all = _collect_levels_sorted(doc)
    if not levels_all:
        raise VistasPorUsuarioError(u"No hay niveles en el proyecto.")

    selected = list(request.levels)
    if not selected:
        raise VistasPorUsuarioError(u"Seleccione al menos un nivel.")

    # Ordenar niveles seleccionados por elevación (ElementId comparable 2024+)
    sel_ids = set(lv.Id for lv in selected)
    levels = [lv for lv in levels_all if lv.Id in sel_ids]
    if not levels:
        raise VistasPorUsuarioError(u"Los niveles seleccionados no son válidos.")

    result = VistasPorUsuarioResult()
    vft_cielo, vft_piso, vft_detail, vft_section = _resolve_view_family_types(doc, result)

    off_piso_1500 = _mm_to_internal(1500)
    off_piso_neg_1500 = _mm_to_internal(-1500)
    off_cut_cielo = _mm_to_internal(1000)
    off_top_depth = _mm_to_internal(300)
    off_fallback = _mm_to_internal(4000)

    used = _existing_view_names(doc)
    sheet_titles = _sheet_titles_for_levels(levels)
    usuario = request.usuario_code

    txn = Transaction(doc, TRANSACTION_TITLE)
    txn.Start()
    try:
        templates = _duplicate_templates_and_types(
            doc, usuario, vft_detail, vft_section, result
        )

        for idx, level in enumerate(levels):
            try:
                lvl_name = str(level.Name or u"")
            except Exception:
                lvl_name = u""

            if len(levels) == 2:
                sheet_cielo = sheet_titles[idx] if idx < len(sheet_titles) else lvl_name
                sheet_piso = sheet_titles[idx] if idx < len(sheet_titles) else lvl_name
            else:
                sheet_cielo = u"PLANTA ESTRUCTURA CIELO " + lvl_name
                sheet_piso = u"PLANTA ESTRUCTURA PISO " + lvl_name

            name_cielo = u"{}_CIELO_{}".format(usuario, lvl_name)
            name_piso = u"{}_PISO_{}".format(usuario, lvl_name)
            key_cielo = _normalize_compare_name(name_cielo)
            key_piso = _normalize_compare_name(name_piso)

            if key_piso in used:
                result.skipped.append(name_piso + u" (ya existía)")
            else:
                piso_lowest = _is_lowest_project_level(level, levels_all)

                def _make_piso(lev=level, nm=name_piso, lowest=piso_lowest):
                    vp = ViewPlan.Create(doc, vft_piso.Id, lev.Id)
                    vp.Name = nm
                    _apply_view_range_piso(
                        vp, lev, off_piso_1500, off_piso_neg_1500, lowest
                    )
                    _regenerate_doc_safe(doc)
                    return vp

                vp = _subtransaction_run(doc, _make_piso)
                _apply_plan_view_metadata(
                    vp,
                    doc,
                    level,
                    usuario,
                    request.scale,
                    sheet_piso,
                    templates.get(u"piso"),
                    result.warnings,
                )
                used.add(key_piso)
                result.created.append(name_piso)

            if key_cielo in used:
                result.skipped.append(name_cielo + u" (ya existía)")
            else:
                lvl_up = _level_above(level, levels_all)

                def _make_cielo(lev=level, lup=lvl_up, nm=name_cielo):
                    vc = ViewPlan.Create(doc, vft_cielo.Id, lev.Id)
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
                    return vc

                vc = _subtransaction_run(doc, _make_cielo)
                _apply_plan_view_metadata(
                    vc,
                    doc,
                    level,
                    usuario,
                    request.scale,
                    sheet_cielo,
                    templates.get(u"cielo"),
                    result.warnings,
                )
                used.add(key_cielo)
                result.created.append(name_cielo)

        _regenerate_doc_safe(doc)
        txn.Commit()
    except Exception:
        if txn.GetStatus() == TransactionStatus.Started:
            txn.RollBack()
        raise

    return result


def format_result_message(result):
    lines = [u"Vistas creadas: {}.".format(len(result.created))]
    if result.created:
        lines.append(u"")
        lines.extend(result.created)
    if result.templates_created:
        lines.append(u"")
        lines.append(u"Plantillas duplicadas:")
        lines.extend(result.templates_created)
    if result.types_created:
        lines.append(u"")
        lines.append(u"Tipos duplicados:")
        lines.extend(result.types_created)
    if result.skipped:
        lines.append(u"")
        lines.append(u"Omitidas:")
        lines.extend(result.skipped)
    if result.warnings:
        lines.append(u"")
        lines.append(u"Advertencias:")
        lines.extend(result.warnings[:15])
        if len(result.warnings) > 15:
            lines.append(u"… (+{} más)".format(len(result.warnings) - 15))
    return u"\n".join(lines)


def format_success_dialog(result, usuario_display=None, usuario_code=None):
    """
    Texto para diálogo WPF de éxito.

    Returns:
        (instruction, content) — instrucción breve + detalle de vistas/tipos.
    """
    who = (usuario_display or u"").strip()
    code = (usuario_code or u"").strip()
    if who and code and who != code:
        who_label = u"{0} ({1})".format(who, code)
    else:
        who_label = who or code or u"el usuario seleccionado"

    n_views = len(result.created or [])
    n_types = len(result.types_created or [])
    instruction = (
        u"Creación exitosa para {0}: {1} vista(s) y {2} tipo(s) "
        u"Detail/Sección."
    ).format(who_label, n_views, n_types)

    lines = []
    if result.created:
        lines.append(u"Vistas creadas:")
        for name in result.created:
            lines.append(u"  • " + name)
    if result.types_created:
        if lines:
            lines.append(u"")
        lines.append(u"Tipos Detail / Building Section:")
        for name in result.types_created:
            lines.append(u"  • " + name)
    if result.filters_created:
        if lines:
            lines.append(u"")
        lines.append(u"Filtros de vista (Section Filter):")
        for name in result.filters_created:
            if name.startswith(u"  →"):
                lines.append(name)
            else:
                lines.append(u"  • " + name)
    if result.templates_created:
        if lines:
            lines.append(u"")
        lines.append(u"Plantillas de vista:")
        for name in result.templates_created:
            lines.append(u"  • " + name)
    if result.skipped:
        if lines:
            lines.append(u"")
        lines.append(u"Omitidas (ya existían):")
        for name in result.skipped:
            lines.append(u"  • " + name)

    content = u"\n".join(lines) if lines else u"No se generaron elementos nuevos."
    return instruction, content
