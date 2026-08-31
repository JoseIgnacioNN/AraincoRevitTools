# -*- coding: utf-8 -*-
"""
Servicio Revit — creación de vistas por categoría (01_ENTREGABLE).

Equivalente Python del grafo Dynamo VistasPorCategoria_script.dyn:
- Valida que no existan vistas con Section Filter = {categoria}_{zona}.
- Permite varias categorías en una sola ejecución (misma zona, niveles y escala).
- Duplica plantillas de vista maestras y tipos Detail / Building Section.
- Crea plantas Cielo y Piso por nivel seleccionado
  (en 01_LO: cuatro plantas por nivel — Cielo/Piso × inferior/superior).
- Asigna rango de vista, escala y parámetros de instancia.

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
    _create_plan_view,
    _find_view_family_type_by_name,
    _is_lowest_project_level,
    _level_above_for_cielo,
    _mm_to_internal,
    _normalize_compare_name,
    _regenerate_doc_safe,
    _set_parameter_value_robust,
    _structural_plan_type_names_for_diagnostics,
    _subtransaction_run,
    _view_family_type_display_name,
    _iter_view_family_types,
)

from vistas_por_categoria.constants import (
    ARMADURA_UBICACION_INFERIOR,
    ARMADURA_UBICACION_SUPERIOR,
    CATEGORIA_LO,
    CLASIFICACION,
    DISCIPLINE_DETAIL_SECTION,
    LO_PLANS_PER_LEVEL,
    MASTER_TEMPLATE_SEEDS,
    PARAM_ARMADURA_UBICACION,
    TRANSACTION_TITLE,
    VFT_NAME_CIELO,
    VFT_NAME_DETAIL_FAMILY,
    VFT_NAME_PISO,
    VFT_NAME_SECTION_FAMILY,
    VIEW_SUFFIX_CIELO,
    VIEW_SUFFIX_INFERIOR,
    VIEW_SUFFIX_PISO,
    VIEW_SUFFIX_SUPERIOR,
    ZONA_DEFAULT,
)


class VistasPorCategoriaError(Exception):
    """Error de validación o configuración previo a la transacción."""


def _normalize_categorias(categorias):
    """Lista de (codigo, etiqueta) sin duplicados. Acepta un código o pares."""
    try:
        basestring
    except NameError:
        basestring = str

    items = []
    seen = set()
    if categorias is None:
        return items
    if isinstance(categorias, basestring):
        raw = [(categorias, categorias)]
    else:
        raw = list(categorias or [])

    for item in raw:
        code = u""
        display = u""
        if isinstance(item, (list, tuple)):
            if len(item) >= 1:
                code = _normalize_compare_name(item[0])
            if len(item) >= 2 and item[1]:
                display = _normalize_compare_name(item[1])
            else:
                display = code
        else:
            code = _normalize_compare_name(item)
            display = code
        if not code or code in seen:
            continue
        seen.add(code)
        items.append((code, display))
    return items


class VistasPorCategoriaRequest(object):
    """Entrada del formulario WPF (una o varias categorías, misma zona)."""

    def __init__(self, categorias, zona, scale, levels):
        # categorias: lista de (codigo, etiqueta) o un código unicode
        self.categorias = _normalize_categorias(categorias)
        first = self.categorias[0] if self.categorias else (u"", u"")
        self.categoria_code = first[0]
        self.categoria_display = first[1]
        self.zona = _normalize_compare_name(zona) or ZONA_DEFAULT
        self.scale = int(scale)
        self.levels = list(levels or [])

    @property
    def section_filter(self):
        return u"{}_{}".format(self.categoria_code, self.zona)


class VistasPorCategoriaResult(object):
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
    """
    Vistas reales (no plantillas). Usa _is_view_template: en pythonnet
    getattr(IsTemplate) a veces no discrimina bien y las plantillas 01_ENTREGABLE_*
    se contaban como vistas y bloqueaban la validación.
    """
    views = []
    for v in FilteredElementCollector(doc).OfClass(View):
        try:
            if v is None:
                continue
            if _is_view_template(v):
                continue
            views.append(v)
        except Exception:
            continue
    return views


def validate_categoria_views_not_exist(doc, categoria_code, zona):
    """
    True si puede continuar (no hay vistas de esa categoría+zona).
    Equivalente Dynamo: Section Filter / pareja Subclasificacion+Zona.

    No considera View Templates (solo vistas de proyecto).
    """
    cat = _normalize_compare_name(categoria_code)
    zon = _normalize_compare_name(zona or ZONA_DEFAULT)
    if not cat:
        return False, u"Seleccione una categoría válida."
    if not zon:
        return False, u"Indique un nombre de zona."

    section_key = _normalize_compare_name(
        u"{}_{}".format(categoria_code, zona or ZONA_DEFAULT)
    )
    matches = []
    for v in _collect_non_template_views(doc):
        sf = _normalize_compare_name(_get_param_string(v, doc, u"Section Filter"))
        if sf and sf == section_key:
            matches.append(v)
            continue
        sub = _normalize_compare_name(_get_param_string(v, doc, u"Subclasificacion"))
        zona_v = _normalize_compare_name(_get_param_string(v, doc, u"Zona"))
        if sub == cat and zona_v == zon:
            matches.append(v)

    if matches:
        return False, (
            u"Las vistas de esta Categoría Zona ya fueron generadas. "
            u"Para vistas adicionales, duplicar de las ya existentes."
        )
    return True, None


def _param_plan(categoria_code, zona):
    section_filter = u"{}_{}".format(categoria_code, zona)
    return (
        (u"Clasificacion", CLASIFICACION),
        (u"Subclasificacion", categoria_code),
        (u"Zona", zona),
        (u"Section Filter", section_filter),
    )


def _param_detail_section(categoria_code, zona):
    section_filter = u"{}_{}".format(categoria_code, zona)
    return (
        (u"Clasificacion", CLASIFICACION),
        (u"Subclasificacion", categoria_code),
        (u"Zona", zona),
        (u"Discipline", DISCIPLINE_DETAIL_SECTION),
        (u"Section Filter", section_filter),
    )


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


def _apply_param_list(element, doc, param_list, report):
    for pname, pval in param_list:
        ok, detail = _set_parameter_value_robust(element, doc, pname, pval)
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


def _is_lo_categoria(categoria_code):
    """True si la categoría es plantas losas (armadura inferior/superior)."""
    return _normalize_compare_name(categoria_code) == _normalize_compare_name(
        CATEGORIA_LO
    )


def plans_per_level_for_categoria(categoria_code):
    """Plantas estructurales creadas por cada nivel de la categoría."""
    if _is_lo_categoria(categoria_code):
        return LO_PLANS_PER_LEVEL
    return 2


def _plan_view_names(categoria, lvl_name, zona):
    """
    Nombres de planta por nivel (no LO).
    LO usa ``_iter_lo_plan_view_jobs`` (4 plantas: Cielo/Piso × inferior/superior).
    """
    name_cielo = u"{}_CIELO_{}_{}".format(categoria, lvl_name, zona)
    name_piso = u"{}_PISO_{}_{}".format(categoria, lvl_name, zona)
    return name_cielo, name_piso


def _lo_plan_template_key(plan_key, arm_suffix):
    """Clave interna de plantilla LO (p. ej. ``cielo_SUPERIOR``)."""
    return u"{}_{}".format(plan_key, arm_suffix)


def _lo_armadura_ubicacion_filter_name(categoria_code, zona, arm_suffix):
    """Nombre del filtro Rebar por tipo de armadura (superior / inferior)."""
    return u"{}_{}_ARM_UBIC_{}".format(categoria_code, zona, arm_suffix)


def _lo_plan_sheet_title(plan_token, arm_suffix, levels, sheet_titles, idx, lvl_name):
    if len(levels) == 2:
        base = sheet_titles[idx] if idx < len(sheet_titles) else lvl_name
        return u"{} {} {}".format(base, plan_token, arm_suffix)
    return u"PLANTA ESTRUCTURA {} {} {}".format(plan_token, arm_suffix, lvl_name)


def _iter_lo_plan_view_jobs(categoria, lvl_name, zona):
    """
    Cuatro plantas LO por nivel: Cielo/Piso con armadura inferior y superior.

    Yields dicts con vft_key, plan_token, arm_suffix, arm_value, name, sheet_title
    (sheet_title se rellena fuera con contexto de nivel).
    """
    for plan_token, vft_key in (
        (VIEW_SUFFIX_CIELO, u"cielo"),
        (VIEW_SUFFIX_PISO, u"piso"),
    ):
        for arm_suffix, arm_value in (
            (VIEW_SUFFIX_INFERIOR, ARMADURA_UBICACION_INFERIOR),
            (VIEW_SUFFIX_SUPERIOR, ARMADURA_UBICACION_SUPERIOR),
        ):
            name = u"{}_{}_{}_{}_{}".format(
                categoria, lvl_name, zona, plan_token, arm_suffix
            )
            yield {
                u"vft_key": vft_key,
                u"plan_token": plan_token,
                u"arm_suffix": arm_suffix,
                u"arm_value": arm_value,
                u"template_key": _lo_plan_template_key(vft_key, arm_suffix),
                u"name": name,
            }


def _plan_sheet_titles(categoria, levels, sheet_titles, idx, lvl_name):
    """Title on Sheet para el par Cielo/Piso (categorías distintas de LO)."""
    if len(levels) == 2:
        base = sheet_titles[idx] if idx < len(sheet_titles) else lvl_name
        return base, base
    return (
        u"PLANTA ESTRUCTURA CIELO " + lvl_name,
        u"PLANTA ESTRUCTURA PISO " + lvl_name,
    )


def _has_plan_templates(templates, categoria_code):
    """True si existen plantillas de planta mínimas para la categoría."""
    if _is_lo_categoria(categoria_code):
        return all(
            templates.get(_lo_plan_template_key(plan_key, arm_suffix))
            for plan_key in (u"cielo", u"piso")
            for arm_suffix in (VIEW_SUFFIX_INFERIOR, VIEW_SUFFIX_SUPERIOR)
        )
    return bool(templates.get(u"cielo") and templates.get(u"piso"))


def _template_names(categoria_code, zona):
    names = {
        u"detail": u"{}_{}_DETAIL_{}".format(CLASIFICACION, categoria_code, zona),
        u"section": u"{}_{}_BUILDING SECTION_{}".format(
            CLASIFICACION, categoria_code, zona
        ),
    }
    if _is_lo_categoria(categoria_code):
        for plan_key, label in ((u"cielo", u"CIELO"), (u"piso", u"PISO")):
            base = u"{}_{}_STRUCTURAL PLAN ({})_{}".format(
                CLASIFICACION, categoria_code, label, zona
            )
            for arm_suffix in (VIEW_SUFFIX_INFERIOR, VIEW_SUFFIX_SUPERIOR):
                names[_lo_plan_template_key(plan_key, arm_suffix)] = u"{}_{}".format(
                    base, arm_suffix
                )
        return names

    names[u"cielo"] = u"{}_{}_STRUCTURAL PLAN (CIELO)_{}".format(
        CLASIFICACION, categoria_code, zona
    )
    names[u"piso"] = u"{}_{}_STRUCTURAL PLAN (PISO)_{}".format(
        CLASIFICACION, categoria_code, zona
    )
    return names


def _resolve_view_family_types(doc, result):
    """Resuelve ViewFamilyType Cielo/Piso/Detail/Section (mismo patrón para todas las categorías)."""
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
        raise VistasPorCategoriaError(
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


def _ensure_section_filter(doc, section_filter_key, warnings):
    """
    Equivalente Dynamo ParameterFilterElement.ByRules:
    - name = {categoria}_{zona}
    - categoría Sections
    - regla Section Filter NotEquals clave
    """
    if not section_filter_key:
        return None

    if DotNetList is None:
        warnings.append(u"No se pudo crear filtro: List[ElementId] no disponible.")
        return None

    cat_ids = DotNetList[ElementId]()
    cat_ids.Add(ElementId(BuiltInCategory.OST_Sections))

    param_id = _find_section_filter_param_id(doc, cat_ids)
    if param_id is None:
        cat_ids2 = DotNetList[ElementId]()
        cat_ids2.Add(ElementId(BuiltInCategory.OST_Views))
        param_id = _find_section_filter_param_id(doc, cat_ids2)
        if param_id is not None:
            cat_ids = cat_ids2
        else:
            warnings.append(
                u"No se encontró el parámetro «{}» para crear el filtro «{}»."
                .format(PARAM_SECTION_FILTER, section_filter_key)
            )
            return None

    filter_name = section_filter_key
    rule = _create_not_equals_string_rule(param_id, section_filter_key)
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


def _apply_parameter_filter_hidden(view_or_template, filter_elem, warnings):
    """Añade filtro a vista o plantilla y oculta elementos que cumplen la regla."""
    if view_or_template is None or filter_elem is None:
        return False
    fid = filter_elem.Id
    label = _view_display_name(view_or_template) or u"?"

    try:
        already = False
        try:
            already = view_or_template.GetFilters().Contains(fid)
        except Exception:
            already = fid in list(view_or_template.GetFilters())
        if not already:
            view_or_template.AddFilter(fid)
    except Exception as ex:
        warnings.append(
            u"«{}»: no se pudo AddFilter: {}".format(label, ex)
        )
        return False

    try:
        view_or_template.SetIsFilterEnabled(fid, True)
    except Exception:
        pass

    try:
        view_or_template.SetFilterVisibility(fid, False)
    except Exception as ex:
        warnings.append(
            u"«{}»: SetFilterVisibility: {}".format(label, ex)
        )

    try:
        view_or_template.SetFilterOverrides(fid, OverrideGraphicSettings())
    except Exception:
        pass

    return True


def _apply_filter_to_view_template(view_template, filter_elem, warnings):
    """Añade el filtro a la plantilla y oculta elementos que cumplen la regla."""
    return _apply_parameter_filter_hidden(view_template, filter_elem, warnings)


def _apply_filters_to_templates(doc, section_filter_key, templates, tpl_names, result):
    """Crea filtro por categoría/zona y lo asigna a plantillas 01_ENTREGABLE_*."""
    filt = _ensure_section_filter(doc, section_filter_key, result.warnings)
    if filt is None:
        return

    applied = []
    for key, tpl in templates.items():
        if tpl is None:
            continue
        expected = _norm_tpl_name(tpl_names.get(key) or u"")
        actual = _norm_tpl_name(_view_display_name(tpl))
        if not expected or actual != expected:
            result.warnings.append(
                u"Filtro «{}» no aplicado a «{}» "
                u"(no es la plantilla por-categoría «{}»)."
                .format(
                    section_filter_key,
                    _view_display_name(tpl),
                    tpl_names.get(key),
                )
            )
            continue
        if _apply_filter_to_view_template(tpl, filt, result.warnings):
            applied.append(_view_display_name(tpl))

    if applied:
        result.filters_created.append(section_filter_key)
        result.filters_created.extend(
            u"  → {}".format(n) for n in applied
        )


def _find_armadura_ubicacion_param_id(doc, cat_ids):
    """ElementId del parámetro «Armadura_Ubicacion» en categoría Rebar."""
    try:
        for spe in FilteredElementCollector(doc).OfClass(SharedParameterElement):
            try:
                if spe and _normalize_compare_name(spe.Name) == _normalize_compare_name(
                    PARAM_ARMADURA_UBICACION
                ):
                    return spe.Id
            except Exception:
                continue
    except Exception:
        pass

    try:
        for p_id in ParameterFilterUtilities.GetFilterableParametersInCommon(doc, cat_ids):
            p_elem = doc.GetElement(p_id)
            if p_elem is None:
                continue
            try:
                if _normalize_compare_name(p_elem.Name) == _normalize_compare_name(
                    PARAM_ARMADURA_UBICACION
                ):
                    return p_id
            except Exception:
                continue
    except Exception:
        pass

    try:
        rebar = (
            FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_Rebar)
            .WhereElementIsNotElementType()
            .FirstElement()
        )
        if rebar is not None:
            p = rebar.LookupParameter(PARAM_ARMADURA_UBICACION)
            if p is not None:
                return p.Id
    except Exception:
        pass
    return None


def _ensure_armadura_ubicacion_filter(doc, filter_name, ubicacion_value, warnings):
    """
    Crea/actualiza filtro Rebar: Armadura_Ubicacion NotEquals ``ubicacion_value``.
    Con visibilidad apagada en plantilla, solo quedan visibles los rebars con ese valor.
    """
    if not filter_name or not ubicacion_value:
        return None

    if DotNetList is None:
        warnings.append(
            u"No se pudo crear filtro «{}»: List[ElementId] no disponible.".format(
                filter_name
            )
        )
        return None

    cat_ids = DotNetList[ElementId]()
    cat_ids.Add(ElementId(BuiltInCategory.OST_Rebar))

    param_id = _find_armadura_ubicacion_param_id(doc, cat_ids)
    if param_id is None:
        warnings.append(
            u"No se encontró «{}» para el filtro «{}».".format(
                PARAM_ARMADURA_UBICACION, filter_name
            )
        )
        return None

    rule = _create_not_equals_string_rule(param_id, ubicacion_value)
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


def _apply_lo_armadura_filters_to_templates(
    doc, categoria_code, zona, templates, tpl_names, result
):
    """
    En 01_LO: filtro Rebar Armadura_Ubicacion en plantillas superior/inferior.
    Superior → NotEquals F' · Inferior → NotEquals F (visibilidad apagada).
    """
    assignments = (
        (VIEW_SUFFIX_SUPERIOR, ARMADURA_UBICACION_SUPERIOR),
        (VIEW_SUFFIX_INFERIOR, ARMADURA_UBICACION_INFERIOR),
    )
    for arm_suffix, arm_value in assignments:
        filter_name = _lo_armadura_ubicacion_filter_name(
            categoria_code, zona, arm_suffix
        )
        filt = _ensure_armadura_ubicacion_filter(
            doc, filter_name, arm_value, result.warnings
        )
        if filt is None:
            continue

        applied = []
        for plan_key in (u"cielo", u"piso"):
            tpl_key = _lo_plan_template_key(plan_key, arm_suffix)
            tpl = templates.get(tpl_key)
            if tpl is None:
                continue
            expected = _norm_tpl_name(tpl_names.get(tpl_key) or u"")
            actual = _norm_tpl_name(_view_display_name(tpl))
            if not expected or actual != expected:
                result.warnings.append(
                    u"Filtro «{}» no aplicado a «{}» "
                    u"(no es la plantilla por-categoría «{}»)."
                    .format(
                        filter_name,
                        _view_display_name(tpl),
                        tpl_names.get(tpl_key),
                    )
                )
                continue
            if _apply_filter_to_view_template(tpl, filt, result.warnings):
                applied.append(_view_display_name(tpl))

        if applied:
            result.filters_created.append(
                u"{} (Rebar «{}» ≠ «{}»)".format(
                    filter_name, PARAM_ARMADURA_UBICACION, arm_value
                )
            )
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

    try:
        if hasattr(view_family_type, "DefaultTemplateId"):
            view_family_type.DefaultTemplateId = template_view.Id
            return True
    except Exception:
        pass

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


def _duplicate_templates_and_types(doc, categoria_code, zona, vft_detail, vft_section, result):
    """
    Como el .dyn (igual para LO, RP, PG, etc.):
    1) Busca plantilla semilla por nombre exacto (MASTER_TEMPLATE_SEEDS).
    2) ViewTemplates.Duplicate → 01_ENTREGABLE_{categoria}_…_{zona}
    3) Asigna parámetros de clasificación a la plantilla nueva.
    4) Duplica tipos Detail / Building Section.
    """
    section_filter = u"{}_{}".format(categoria_code, zona)
    params_plan = _param_plan(categoria_code, zona)
    params_detail = _param_detail_section(categoria_code, zona)
    tpl_names = _template_names(categoria_code, zona)
    templates = {}
    sample_names, total_tpl = _template_seed_names_for_diagnostics(doc)
    missing_plan_seeds = []
    plan_fail_details = []

    for key, seed_name in MASTER_TEMPLATE_SEEDS.items():
        if key in (u"cielo", u"piso") and _is_lo_categoria(categoria_code):
            master, diag = _find_view_template_master(doc, seed_name)
            if master is None:
                extra = u""
                if diag:
                    extra = u" Motivo: {}.".format(diag)
                if sample_names:
                    lines = u"\n".join(u"  • {}".format(n) for n in sample_names)
                    if total_tpl > len(sample_names):
                        lines += u"\n  … (+{} más)".format(
                            total_tpl - len(sample_names)
                        )
                    extra += u"\nVistas/plantillas detectadas (muestra):\n" + lines
                for arm_suffix in (VIEW_SUFFIX_INFERIOR, VIEW_SUFFIX_SUPERIOR):
                    tpl_key = _lo_plan_template_key(key, arm_suffix)
                    msg = (
                        u"Plantilla semilla no encontrada: «{}» "
                        u"(necesaria para crear «{}»).{}"
                    ).format(seed_name, tpl_names.get(tpl_key), extra)
                    result.warnings.append(msg)
                missing_plan_seeds.append(seed_name)
                plan_fail_details.append(
                    u"«{}»: {}".format(seed_name, diag or u"no encontrada")
                )
                continue

            for arm_suffix in (VIEW_SUFFIX_INFERIOR, VIEW_SUFFIX_SUPERIOR):
                tpl_key = _lo_plan_template_key(key, arm_suffix)
                target_name = tpl_names.get(tpl_key)
                dup, err = _duplicate_view_template(doc, master, target_name)
                if dup is None:
                    result.warnings.append(
                        u"No se pudo duplicar plantilla «{}» desde «{}»: {}. "
                        u"Se asignará la plantilla semilla compartida a las vistas."
                        .format(target_name, seed_name, err)
                    )
                    templates[tpl_key] = master
                    continue

                _apply_param_list(dup, doc, params_plan, result.warnings)
                templates[tpl_key] = dup
                result.templates_created.append(target_name)
            continue

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
            templates[key] = master
            continue

        if key in (u"cielo", u"piso"):
            _apply_param_list(dup, doc, params_plan, result.warnings)
        else:
            _apply_param_list(dup, doc, params_detail, result.warnings)

        templates[key] = dup
        result.templates_created.append(tpl_names[key])

    if missing_plan_seeds and not _has_plan_templates(templates, categoria_code):
        detail = u""
        if plan_fail_details:
            detail = u"\n\nDetalle:\n" + u"\n".join(
                u"  • {}".format(d) for d in plan_fail_details
            )
        if sample_names:
            detail += u"\n\nDetectadas por la API (muestra):\n" + u"\n".join(
                u"  • {}".format(n) for n in sample_names[:15]
            )
        raise VistasPorCategoriaError(
            u"No se pudieron crear las plantillas de planta (Cielo/Piso).\n"
            u"Semillas requeridas:\n"
            + u"\n".join(u"  • {}".format(s) for s in missing_plan_seeds)
            + detail
        )

    detail_type_name = u"Detail ({})".format(section_filter)
    section_type_name = u"Building Section ({})".format(section_filter)

    if vft_detail is not None:
        dup_d, err = _get_or_duplicate_vft(doc, vft_detail, detail_type_name)
        if dup_d is None:
            result.warnings.append(u"Detail type: {}".format(err))
        else:
            _apply_param_list(dup_d, doc, params_detail, result.warnings)
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
            _apply_param_list(dup_s, doc, params_detail, result.warnings)
            tpl_section = templates.get(u"section")
            if tpl_section is not None:
                _set_default_view_template_on_type(dup_s, tpl_section, result.warnings)
            else:
                result.warnings.append(
                    u"Tipo «{}»: no hay plantilla BUILDING SECTION para asignar por defecto."
                    .format(section_type_name)
                )
            result.types_created.append(section_type_name)

    _apply_filters_to_templates(doc, section_filter, templates, tpl_names, result)
    if _is_lo_categoria(categoria_code):
        _apply_lo_armadura_filters_to_templates(
            doc, categoria_code, zona, templates, tpl_names, result
        )

    return templates


def _apply_view_template_to_view(view, template_view, warnings):
    """Asigna plantilla de vista si ``view`` no es plantilla."""
    if view is None or template_view is None:
        return False
    if _is_view_template(view):
        return False

    applied = False
    try:
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
    return applied


def _apply_plan_view_metadata(
    view, doc, level, categoria_code, zona, scale, sheet_title, template_view, warnings,
    apply_template=True,
):
    try:
        view.Scale = int(scale)
    except Exception:
        warnings.append(u'No se pudo asignar escala {} a «{}»'.format(scale, view.Name))

    _apply_param_list(view, doc, _param_plan(categoria_code, zona), warnings)

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

    if template_view is not None and apply_template:
        _apply_view_template_to_view(view, template_view, warnings)


def _create_lo_plan_views_for_level(
    doc,
    categoria,
    zona,
    scale,
    level,
    levels,
    levels_all,
    idx,
    sheet_titles,
    vft_cielo,
    vft_piso,
    templates,
    used,
    result,
    offsets,
):
    """Cuatro plantas LO: Cielo/Piso × armadura inferior/superior."""
    off_piso_1500 = offsets[0]
    off_piso_neg_1500 = offsets[1]
    off_cut_cielo = offsets[2]
    off_top_depth = offsets[3]
    off_fallback = offsets[4]

    try:
        lvl_name = str(level.Name or u"")
    except Exception:
        lvl_name = u""

    piso_lowest = _is_lowest_project_level(level, levels_all)
    lvl_up = _level_above_for_cielo(
        level, levels_all, off_cut_cielo, off_top_depth
    )

    for job in _iter_lo_plan_view_jobs(categoria, lvl_name, zona):
        name = job[u"name"]
        key = _normalize_compare_name(name)
        if key in used:
            result.skipped.append(name + u" (ya existía)")
            continue

        vft_key = job[u"vft_key"]
        plan_token = job[u"plan_token"]
        arm_suffix = job[u"arm_suffix"]
        template_key = job[u"template_key"]
        sheet_title = _lo_plan_sheet_title(
            plan_token, arm_suffix, levels, sheet_titles, idx, lvl_name
        )

        if vft_key == u"piso":
            def _make_piso(
                lev=level,
                nm=name,
                lowest=piso_lowest,
            ):
                vp = _create_plan_view(doc, vft_piso, lev.Id, nm)
                vp.Name = nm
                _apply_view_range_piso(
                    vp, lev, off_piso_1500, off_piso_neg_1500, lowest
                )
                _regenerate_doc_safe(doc)
                return vp

            view = _subtransaction_run(doc, _make_piso)
        else:
            def _make_cielo(
                lev=level,
                lup=lvl_up,
                nm=name,
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
                return vc

            view = _subtransaction_run(doc, _make_cielo)

        _apply_plan_view_metadata(
            view,
            doc,
            level,
            categoria,
            zona,
            scale,
            sheet_title,
            templates.get(template_key),
            result.warnings,
        )

        used.add(key)
        result.created.append(name)


def _create_plan_views_for_categoria(
    doc,
    categoria,
    zona,
    scale,
    levels,
    levels_all,
    vft_cielo,
    vft_piso,
    vft_detail,
    vft_section,
    used,
    sheet_titles,
    result,
    offsets,
):
    """Crea plantillas, tipos y plantas Cielo/Piso de una categoría (TX abierta).

    En 01_LO: cuatro plantas por nivel (Cielo/Piso × inferior/superior).
    Resto: dos plantas por nivel (Cielo y Piso).
    """
    templates = _duplicate_templates_and_types(
        doc, categoria, zona, vft_detail, vft_section, result
    )
    off_piso_1500 = offsets[0]
    off_piso_neg_1500 = offsets[1]
    off_cut_cielo = offsets[2]
    off_top_depth = offsets[3]
    off_fallback = offsets[4]

    if _is_lo_categoria(categoria):
        for idx, level in enumerate(levels):
            _create_lo_plan_views_for_level(
                doc,
                categoria,
                zona,
                scale,
                level,
                levels,
                levels_all,
                idx,
                sheet_titles,
                vft_cielo,
                vft_piso,
                templates,
                used,
                result,
                offsets,
            )
        return

    for idx, level in enumerate(levels):
        try:
            lvl_name = str(level.Name or u"")
        except Exception:
            lvl_name = u""

        sheet_cielo, sheet_piso = _plan_sheet_titles(
            categoria, levels, sheet_titles, idx, lvl_name
        )
        name_cielo, name_piso = _plan_view_names(categoria, lvl_name, zona)
        key_cielo = _normalize_compare_name(name_cielo)
        key_piso = _normalize_compare_name(name_piso)

        if key_piso in used:
            result.skipped.append(name_piso + u" (ya existía)")
        else:
            piso_lowest = _is_lowest_project_level(level, levels_all)

            def _make_piso(lev=level, nm=name_piso, lowest=piso_lowest):
                vp = _create_plan_view(doc, vft_piso, lev.Id, nm)
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
                categoria,
                zona,
                scale,
                sheet_piso,
                templates.get(u"piso"),
                result.warnings,
            )
            used.add(key_piso)
            result.created.append(name_piso)

        if key_cielo in used:
            result.skipped.append(name_cielo + u" (ya existía)")
        else:
            lvl_up = _level_above_for_cielo(
                level, levels_all, off_cut_cielo, off_top_depth
            )

            def _make_cielo(lev=level, lup=lvl_up, nm=name_cielo):
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
                return vc

            vc = _subtransaction_run(doc, _make_cielo)
            _apply_plan_view_metadata(
                vc,
                doc,
                level,
                categoria,
                zona,
                scale,
                sheet_cielo,
                templates.get(u"cielo"),
                result.warnings,
            )
            used.add(key_cielo)
            result.created.append(name_cielo)


def create_categoria_views(doc, request):
    """
    Ejecuta la creación completa (una o varias categorías, misma zona).
    Devuelve VistasPorCategoriaResult.
    Lanza VistasPorCategoriaError si falla validación previa.
    """
    cats = list(request.categorias or [])
    if not cats and request.categoria_code:
        cats = [(request.categoria_code, request.categoria_display)]
    if not cats:
        raise VistasPorCategoriaError(u"Seleccione al menos una categoría.")

    zona = request.zona or ZONA_DEFAULT

    levels_all = _collect_levels_sorted(doc)
    if not levels_all:
        raise VistasPorCategoriaError(u"No hay niveles en el proyecto.")

    selected = list(request.levels)
    if not selected:
        raise VistasPorCategoriaError(u"Seleccione al menos un nivel.")

    sel_ids = set(lv.Id for lv in selected)
    levels = [lv for lv in levels_all if lv.Id in sel_ids]
    if not levels:
        raise VistasPorCategoriaError(u"Los niveles seleccionados no son válidos.")

    result = VistasPorCategoriaResult()
    to_create = []
    for code, _display in cats:
        ok, _msg = validate_categoria_views_not_exist(doc, code, zona)
        if ok:
            to_create.append(code)
        else:
            result.skipped.append(
                u"{} / zona {} (ya generadas)".format(code, zona)
            )

    if not to_create:
        if len(cats) == 1:
            raise VistasPorCategoriaError(
                u"Las vistas de esta Categoría Zona ya fueron generadas. "
                u"Para vistas adicionales, duplicar de las ya existentes."
            )
        raise VistasPorCategoriaError(
            u"Las vistas de las categorías seleccionadas ya fueron generadas "
            u"para la zona {}. Para vistas adicionales, duplicar de las ya "
            u"existentes.".format(zona)
        )

    vft_cielo, vft_piso, vft_detail, vft_section = _resolve_view_family_types(
        doc, result
    )

    offsets = (
        _mm_to_internal(1500),
        _mm_to_internal(-1500),
        _mm_to_internal(1000),
        _mm_to_internal(300),
        _mm_to_internal(4000),
    )
    used = _existing_view_names(doc)
    sheet_titles = _sheet_titles_for_levels(levels)
    scale = request.scale

    txn = Transaction(doc, TRANSACTION_TITLE)
    txn.Start()
    try:
        for categoria in to_create:
            _create_plan_views_for_categoria(
                doc,
                categoria,
                zona,
                scale,
                levels,
                levels_all,
                vft_cielo,
                vft_piso,
                vft_detail,
                vft_section,
                used,
                sheet_titles,
                result,
                offsets,
            )
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


def format_success_dialog(
    result,
    categoria_display=None,
    categoria_code=None,
    zona=None,
    categorias=None,
):
    """
    Texto para diálogo WPF de éxito.

    Returns:
        (instruction, content) — instrucción breve + detalle de vistas/tipos.
    """
    zon = (zona or ZONA_DEFAULT).strip()
    items = _normalize_categorias(categorias)
    if not items:
        who = (categoria_display or u"").strip()
        code = (categoria_code or u"").strip()
        if who or code:
            items = [(code or who, who or code)]

    if len(items) > 1:
        codes = [c[0] for c in items if c and c[0]]
        who_label = u"{0} categorías ({1}) / zona {2}".format(
            len(codes), u", ".join(codes), zon
        )
    else:
        who = (items[0][1] if items else u"").strip()
        code = (items[0][0] if items else u"").strip()
        if not who:
            who = (categoria_display or u"").strip()
        if not code:
            code = (categoria_code or u"").strip()
        if who and code and who != code:
            who_label = u"{0} / zona {1}".format(who, zon)
        else:
            who_label = u"{0} / zona {1}".format(who or code or u"categoría", zon)

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
        lines.append(u"Filtros de vista:")
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

