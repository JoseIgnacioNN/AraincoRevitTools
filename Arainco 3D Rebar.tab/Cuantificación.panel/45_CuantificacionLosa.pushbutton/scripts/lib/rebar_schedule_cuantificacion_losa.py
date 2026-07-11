# -*- coding: utf-8 -*-
"""
Cuantificación de armadura en losas: tablas Rebar por nivel, malla y ubicación.

Filtros fijos por tabla:
  - Host Category = Floor (entero RebarHostCategory.Floor = 5 en la API)
  - Armadura_Nivel = nombre del nivel
  - Armadura_Malla = Yes (malla) o No (sin malla)
  - Armadura_Ubicacion = F' (superior) o F (inferior)

Por cada nivel se generan 4 tablas (Malla/Sin Malla × Superior/Inferior).
Los campos de filtrado (Armadura_Nivel, Armadura_Malla, Armadura_Ubicacion,
Host Category) se mantienen ocultos en la vista del cuadro.
"""

from __future__ import print_function

from Autodesk.Revit.DB import (
    BuiltInCategory,
    Category,
    FilteredElementCollector,
    Level,
    ScheduleFilter,
    ScheduleFilterType,
    StorageType,
    ViewDuplicateOption,
    ViewSchedule,
)
from Autodesk.Revit.DB.Structure import RebarHostCategory

HOST_CATEGORY_PARAM = u"Host Category"
HOST_CATEGORY_ALIASES = (
    HOST_CATEGORY_PARAM,
    u"Categoría del anfitrión",
    u"Categoria del anfitrion",
    u"Categoría de anfitrión",
)
# En UI se ve «Floor»; en API el filtro es el entero de RebarHostCategory.
HOST_CATEGORY_FLOOR_LABEL = u"Floor"
HOST_CATEGORY_FLOOR_VALUE = int(RebarHostCategory.Floor)  # 5

ARMADURA_NIVEL_PARAM = u"Armadura_Nivel"
ARMADURA_MALLA_PARAM = u"Armadura_Malla"
ARMADURA_UBICACION_PARAM = u"Armadura_Ubicacion"

# Parámetro de instancia de la vista/tabla (organización del navegador).
SUBCLASIFICACION_PARAM = u"Subclasificacion"
SUBCLASIFICACION_ALIASES = (
    SUBCLASIFICACION_PARAM,
    u"Subclasificación",
)

TEMPLATE_SCHEDULE_NAME = u"Arainco Plantilla Cuantificación Losa"
SCHEDULE_NAME_PREFIX = u"Arainco Cuantificación Losa - "

# Sufijos de nombre y valor Yes/No (entero) para Armadura_Malla.
# Sin malla: suffix vacío → el nombre solo lleva Superior/Inferior.
MALLA_VARIANTS = (
    {u"suffix": u"Malla", u"yes": True, u"int_value": 1, u"label": u"Yes"},
    {u"suffix": u"", u"yes": False, u"int_value": 0, u"label": u"No"},
)

# Superior = F' · Inferior = F (Armadura_Ubicacion en losas).
UBICACION_VARIANTS = (
    {u"suffix": u"Superior", u"value": u"F'"},
    {u"suffix": u"Inferior", u"value": u"F"},
)

# Segmentos con Rounded* (no hay I; salta de H a J). La API no crea fórmulas.
ROUNDED_LETTERS = (u"A", u"B", u"C", u"D", u"E", u"F", u"G", u"H", u"J")

# Campos nativos (sin calculados). Tras Quantity el usuario inserta Rounded* en plantilla.
# Rebar Number no se incluye en estas tablas.
NATIVE_FIELD_SPECS = (
    (u"Bar Diameter", (u"Bar Diameter", u"Diámetro de barra", u"Diametro de barra")),
    (u"Shape", (u"Shape", u"Forma")),
    (u"Quantity", (u"Quantity", u"Cantidad", u"Count", u"Recuento")),
    (u"K", (u"K",)),
    (u"O", (u"O",)),
    (u"R", (u"R",)),
    (u"A", (u"A",)),
    (u"B", (u"B",)),
    (u"C", (u"C",)),
    (u"D", (u"D",)),
    (u"E", (u"E",)),
    (u"F", (u"F",)),
    (u"G", (u"G",)),
    (u"H", (u"H",)),
    (u"J", (u"J",)),
    (ARMADURA_NIVEL_PARAM, (ARMADURA_NIVEL_PARAM,)),
    (ARMADURA_MALLA_PARAM, (ARMADURA_MALLA_PARAM,)),
    (ARMADURA_UBICACION_PARAM, (ARMADURA_UBICACION_PARAM,)),
    (HOST_CATEGORY_PARAM, HOST_CATEGORY_ALIASES),
)


def rounded_field_name(letter):
    return u"Rounded {0}".format(letter)


def rounded_formula(letter):
    return u"roundup({0} / 10 mm) * 10 mm".format(letter)



def _normalize_name(name):
    if name is None:
        return u""
    try:
        s = unicode(name).strip().lower()
    except NameError:
        s = str(name).strip().lower()
    for src, dst in (
        (u"á", u"a"),
        (u"í", u"i"),
        (u"é", u"e"),
        (u"ó", u"o"),
        (u"ú", u"u"),
        (u"ñ", u"n"),
    ):
        s = s.replace(src, dst)
    return s


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _schedulable_field_name(defn, doc, sched_field):
    try:
        return sched_field.GetName(doc)
    except Exception:
        pass
    try:
        pid = sched_field.ParameterId
        if pid is None or pid == pid.InvalidElementId:
            return u""
        from Autodesk.Revit.DB import SharedParameterElement

        spe = doc.GetElement(pid)
        if isinstance(spe, SharedParameterElement):
            pd = spe.GetDefinition()
            if pd is not None:
                return pd.Name
    except Exception:
        pass
    return u""


def _find_schedulable_field(defn, doc, param_names):
    targets = {_normalize_name(n) for n in param_names}
    for sf in defn.GetSchedulableFields():
        name = _schedulable_field_name(defn, doc, sf)
        if _normalize_name(name) in targets:
            return sf, name
    return None, None


def _field_display_name(field):
    for attr in ("ColumnHeading", "GetName"):
        try:
            if attr == "GetName":
                val = field.GetName()
            else:
                val = getattr(field, attr, None)
            if val is not None and _as_unicode(val).strip():
                return _as_unicode(val).strip()
        except Exception:
            continue
    return u""


def _find_schedule_field(defn, doc, param_names):
    targets = {_normalize_name(n) for n in param_names}
    for i in range(defn.GetFieldCount()):
        field = defn.GetField(i)
        heading = _field_display_name(field)
        if _normalize_name(heading) in targets:
            return field
        try:
            sf = field.GetSchedulableField()
        except Exception:
            continue
        name = _schedulable_field_name(defn, doc, sf)
        if _normalize_name(name) in targets:
            return field
    return None


def _ensure_field(defn, doc, param_names):
    existing = _find_schedule_field(defn, doc, param_names)
    if existing is not None:
        return existing
    sf, _ = _find_schedulable_field(defn, doc, param_names)
    if sf is None:
        return None
    try:
        return defn.AddField(sf)
    except Exception:
        return None


def _hide_field(field):
    """Oculta un campo del cuadro (sigue disponible para filtros)."""
    if field is None:
        return
    try:
        field.IsHidden = True
    except Exception:
        pass


# Campos solo para filtrado: deben existir en la definición pero ocultos en la vista.
FILTER_ONLY_FIELD_SPECS = (
    (ARMADURA_NIVEL_PARAM, (ARMADURA_NIVEL_PARAM,)),
    (ARMADURA_MALLA_PARAM, (ARMADURA_MALLA_PARAM,)),
    (ARMADURA_UBICACION_PARAM, (ARMADURA_UBICACION_PARAM,)),
    (HOST_CATEGORY_PARAM, HOST_CATEGORY_ALIASES),
)


def _hide_filter_only_fields(defn, doc):
    """Oculta Armadura_Nivel, Armadura_Malla, Armadura_Ubicacion y Host Category."""
    for _label, aliases in FILTER_ONLY_FIELD_SPECS:
        field = _ensure_field(defn, doc, aliases)
        _hide_field(field)


def _lookup_param(element, names):
    if element is None:
        return None
    for name in names:
        try:
            p = element.LookupParameter(name)
            if p is not None:
                return p
        except Exception:
            continue
    try:
        targets = {_normalize_name(n) for n in names}
        for p in element.Parameters:
            try:
                dn = p.Definition.Name if p.Definition is not None else u""
            except Exception:
                continue
            if _normalize_name(dn) in targets:
                return p
    except Exception:
        pass
    return None


def set_subclasificacion_host_category(schedule, host_category_label=None):
    """
    Rellena el parámetro de vista «Subclasificacion» con el Host Category
    (p. ej. Floor). Usado en propiedades de la tabla / organización del navegador.

    Returns:
        tuple: (ok: bool, detail: str)
    """
    if schedule is None:
        return False, u"Sin tabla."
    value = host_category_label or HOST_CATEGORY_FLOOR_LABEL
    return _set_view_string_param(schedule, SUBCLASIFICACION_ALIASES, value)


def set_armadura_nivel_on_schedule(schedule, level_name):
    """
    Rellena el parámetro de vista «Armadura_Nivel» con el nombre del nivel
    de esa tabla.

    Returns:
        tuple: (ok: bool, detail: str)
    """
    if schedule is None:
        return False, u"Sin tabla."
    if not level_name:
        return False, u"Nivel vacío."
    return _set_view_string_param(schedule, (ARMADURA_NIVEL_PARAM,), level_name)


def _set_view_string_param(element, param_names, value):
    """Escribe un parámetro de instancia de vista (string / SetValueString)."""
    p = _lookup_param(element, param_names)
    label = param_names[0] if param_names else u"?"
    if p is None:
        return False, u"No se encontró el parámetro «{0}» en la tabla.".format(label)
    if p.IsReadOnly:
        return False, u"«{0}» es de solo lectura.".format(label)

    try:
        if p.StorageType == StorageType.String:
            p.Set(value)
            return True, u""
    except Exception as ex:
        last = _as_unicode(ex)
    else:
        last = u""

    try:
        p.SetValueString(value)
        return True, u""
    except Exception as ex:
        last = _as_unicode(ex) or last

    return False, last or u"No se pudo escribir «{0}».".format(label)


def apply_schedule_identity_params(schedule, level_name):
    """
    Subclasificacion = Floor y Armadura_Nivel = nombre del nivel en la vista tabla.

    Returns:
        list of (param_label, detail) for failures only.
    """
    warnings = []
    ok, detail = set_subclasificacion_host_category(schedule, HOST_CATEGORY_FLOOR_LABEL)
    if not ok:
        warnings.append((SUBCLASIFICACION_PARAM, detail))
    ok, detail = set_armadura_nivel_on_schedule(schedule, level_name)
    if not ok:
        warnings.append((ARMADURA_NIVEL_PARAM, detail))
    return warnings


def _rounded_name_targets():
    return {_normalize_name(rounded_field_name(L)) for L in ROUNDED_LETTERS}


def contar_campos_rounded(defn):
    """Cuántos campos Rounded A..H / J hay en la definición del cuadro."""
    targets = _rounded_name_targets()
    n = 0
    for i in range(defn.GetFieldCount()):
        name = _normalize_name(_field_display_name(defn.GetField(i)))
        if name in targets:
            n += 1
    return n


def listar_rounded_faltantes(defn):
    """Nombres Rounded* que aún no están en el cuadro."""
    targets_have = set()
    wanted = _rounded_name_targets()
    for i in range(defn.GetFieldCount()):
        name = _normalize_name(_field_display_name(defn.GetField(i)))
        if name in wanted:
            targets_have.add(name)
    missing = []
    for letter in ROUNDED_LETTERS:
        n = rounded_field_name(letter)
        if _normalize_name(n) not in targets_have:
            missing.append(n)
    return missing


def _clear_filters(defn):
    while defn.GetFilterCount() > 0:
        defn.RemoveFilter(0)


def _add_native_fields(defn, doc):
    """Agrega campos nativos en orden (sin Rebar Number)."""
    missing = []
    for _label, aliases in NATIVE_FIELD_SPECS:
        if _find_schedule_field(defn, doc, aliases) is not None:
            continue
        field = _ensure_field(defn, doc, aliases)
        if field is None:
            missing.append(aliases[0])
    return missing


def _try_add_filter(defn, field, filter_type, value):
    """Intenta AddFilter; devuelve (ok, error)."""
    try:
        defn.AddFilter(ScheduleFilter(field.FieldId, filter_type, value))
        return True, u""
    except Exception as ex:
        return False, _as_unicode(ex)


def _add_host_category_floor_filter(defn, host_field):
    """
    Filtra Host Category = Floor.

    El parámetro es entero (RebarHostCategory); la cadena «Floor» no es válida
    para ScheduleFilter.Equal.
    """
    # Preferir Equal con el enum entero.
    ok, err = _try_add_filter(
        defn, host_field, ScheduleFilterType.Equal, HOST_CATEGORY_FLOOR_VALUE
    )
    if ok:
        return True, u""

    # Respaldo: algunos builds aceptan string Contains/Equal.
    for ftype, val in (
        (ScheduleFilterType.Equal, HOST_CATEGORY_FLOOR_LABEL),
        (ScheduleFilterType.Contains, HOST_CATEGORY_FLOOR_LABEL),
    ):
        ok2, err2 = _try_add_filter(defn, host_field, ftype, val)
        if ok2:
            return True, u""
        err = err2 or err

    return False, err or u"Filtro Host Category = Floor no válido."


def _add_armadura_nivel_filter(defn, nivel_field, level_name):
    ok, err = _try_add_filter(
        defn, nivel_field, ScheduleFilterType.Equal, level_name
    )
    if ok:
        return True, u""
    ok2, err2 = _try_add_filter(
        defn, nivel_field, ScheduleFilterType.Contains, level_name
    )
    if ok2:
        return True, u""
    return False, err2 or err or u"Filtro Armadura_Nivel no válido."


def _add_armadura_malla_filter(defn, malla_field, malla_yes):
    """
    Filtra Armadura_Malla Yes/No.

    Parámetro Yes/No: preferir entero 1/0; respaldos con cadenas Yes/No.
    """
    int_value = 1 if malla_yes else 0
    label = u"Yes" if malla_yes else u"No"
    ok, err = _try_add_filter(
        defn, malla_field, ScheduleFilterType.Equal, int_value
    )
    if ok:
        return True, u""
    for ftype, val in (
        (ScheduleFilterType.Equal, label),
        (ScheduleFilterType.Contains, label),
        (ScheduleFilterType.Equal, u"Sí" if malla_yes else u"No"),
        (ScheduleFilterType.Equal, u"Si" if malla_yes else u"No"),
    ):
        ok2, err2 = _try_add_filter(defn, malla_field, ftype, val)
        if ok2:
            return True, u""
        err = err2 or err
    return False, err or u"Filtro Armadura_Malla no válido."


def _add_armadura_ubicacion_filter(defn, ubicacion_field, ubicacion_value):
    """Filtra Armadura_Ubicacion = F (inferior) o F' (superior)."""
    ok, err = _try_add_filter(
        defn, ubicacion_field, ScheduleFilterType.Equal, ubicacion_value
    )
    if ok:
        return True, u""
    ok2, err2 = _try_add_filter(
        defn, ubicacion_field, ScheduleFilterType.Contains, ubicacion_value
    )
    if ok2:
        return True, u""
    return False, err2 or err or u"Filtro Armadura_Ubicacion no válido."


def _apply_losa_filters(defn, doc, level_name, malla_yes, ubicacion_value):
    host_field = _ensure_field(defn, doc, HOST_CATEGORY_ALIASES)
    if host_field is None:
        return False, u"No se encontró el parámetro «{0}» en el cuadro.".format(
            HOST_CATEGORY_PARAM
        )

    nivel_field = _ensure_field(defn, doc, (ARMADURA_NIVEL_PARAM,))
    if nivel_field is None:
        return False, u"No se encontró el parámetro «{0}» en el cuadro.".format(
            ARMADURA_NIVEL_PARAM
        )

    malla_field = _ensure_field(defn, doc, (ARMADURA_MALLA_PARAM,))
    if malla_field is None:
        return False, u"No se encontró el parámetro «{0}» en el cuadro.".format(
            ARMADURA_MALLA_PARAM
        )

    ubicacion_field = _ensure_field(defn, doc, (ARMADURA_UBICACION_PARAM,))
    if ubicacion_field is None:
        return False, u"No se encontró el parámetro «{0}» en el cuadro.".format(
            ARMADURA_UBICACION_PARAM
        )

    _clear_filters(defn)
    ok, err = _add_host_category_floor_filter(defn, host_field)
    if not ok:
        return False, u"No se pudo filtrar Host Category: {0}".format(err)
    ok, err = _add_armadura_nivel_filter(defn, nivel_field, level_name)
    if not ok:
        return False, u"No se pudo filtrar Armadura_Nivel: {0}".format(err)
    ok, err = _add_armadura_malla_filter(defn, malla_field, malla_yes)
    if not ok:
        return False, u"No se pudo filtrar Armadura_Malla: {0}".format(err)
    ok, err = _add_armadura_ubicacion_filter(defn, ubicacion_field, ubicacion_value)
    if not ok:
        return False, u"No se pudo filtrar Armadura_Ubicacion: {0}".format(err)

    _hide_filter_only_fields(defn, doc)
    return True, u""


def listar_niveles(doc):
    """Niveles del proyecto ordenados por elevación (ascendente)."""
    levels = list(FilteredElementCollector(doc).OfClass(Level))
    try:
        levels.sort(key=lambda lv: (lv.Elevation, _as_unicode(lv.Name)))
    except Exception:
        levels.sort(key=lambda lv: _as_unicode(getattr(lv, "Name", u"")))
    return levels


def schedule_name_for_variant(level_name, malla_suffix, ubicacion_suffix):
    """
    Nombre de tabla.
    - Con malla: ``… - {Nivel} - Malla - Superior|Inferior``
    - Sin malla: ``… - {Nivel} - Armadura - Superior|Inferior``
    """
    if malla_suffix:
        return u"{0}{1} - {2} - {3}".format(
            SCHEDULE_NAME_PREFIX, level_name, malla_suffix, ubicacion_suffix
        )
    return u"{0}{1} - Armadura - {2}".format(
        SCHEDULE_NAME_PREFIX, level_name, ubicacion_suffix
    )


def _find_schedule_by_name(doc, name):
    target = _normalize_name(name)
    for vs in FilteredElementCollector(doc).OfClass(ViewSchedule):
        try:
            if vs.IsTemplate:
                continue
            if _normalize_name(vs.Name) == target:
                return vs
        except Exception:
            continue
    return None


def _is_rebar_schedule(schedule):
    try:
        cat_id = schedule.Definition.CategoryId
        cat = Category.GetCategory(schedule.Document, BuiltInCategory.OST_Rebar)
        return cat is not None and cat_id == cat.Id
    except Exception:
        return False


def asegurar_plantilla(doc):
    """
    Crea o reutiliza la plantilla de cuantificación (sin filtro por nivel).

    Returns:
        tuple: (ViewSchedule|None, created: bool, missing_fields: list, message: str)
    """
    cat = Category.GetCategory(doc, BuiltInCategory.OST_Rebar)
    if cat is None:
        return None, False, [], u"No se encontró la categoría Structural Rebar."

    existing = _find_schedule_by_name(doc, TEMPLATE_SCHEDULE_NAME)
    created = False
    if existing is not None:
        schedule = existing
        if not _is_rebar_schedule(schedule):
            return (
                None,
                False,
                [],
                u"Ya existe una vista «{0}» que no es un cuadro de Rebar.".format(
                    TEMPLATE_SCHEDULE_NAME
                ),
            )
    else:
        try:
            schedule = ViewSchedule.CreateSchedule(doc, cat.Id)
        except Exception as ex:
            return None, False, [], u"No se pudo crear la plantilla: {0}".format(ex)
        created = True
        try:
            schedule.Name = TEMPLATE_SCHEDULE_NAME
        except Exception:
            pass

    defn = schedule.Definition
    missing = _add_native_fields(defn, doc)

    # Plantilla: solo filtro Host Category = Floor (sin nivel).
    host_field = _ensure_field(defn, doc, HOST_CATEGORY_ALIASES)
    if host_field is None:
        if created:
            try:
                doc.Delete(schedule.Id)
            except Exception:
                pass
        return (
            None,
            False,
            missing,
            u"El parámetro «{0}» no está disponible en cuadros de armadura.".format(
                HOST_CATEGORY_PARAM
            ),
        )

    nivel_ok = _ensure_field(defn, doc, (ARMADURA_NIVEL_PARAM,))
    if nivel_ok is None:
        if created:
            try:
                doc.Delete(schedule.Id)
            except Exception:
                pass
        return (
            None,
            False,
            missing,
            u"El parámetro «{0}» no está disponible en cuadros de armadura.\n"
            u"Verifique que esté vinculado a la categoría Rebar.".format(
                ARMADURA_NIVEL_PARAM
            ),
        )

    malla_ok = _ensure_field(defn, doc, (ARMADURA_MALLA_PARAM,))
    if malla_ok is None:
        if created:
            try:
                doc.Delete(schedule.Id)
            except Exception:
                pass
        return (
            None,
            False,
            missing,
            u"El parámetro «{0}» no está disponible en cuadros de armadura.\n"
            u"Verifique que esté vinculado a la categoría Rebar.".format(
                ARMADURA_MALLA_PARAM
            ),
        )

    # Visible en plantilla; el filtro F / F' se aplica en cada tabla por variante.
    ubicacion_ok = _ensure_field(defn, doc, (ARMADURA_UBICACION_PARAM,))
    if ubicacion_ok is None:
        if created:
            try:
                doc.Delete(schedule.Id)
            except Exception:
                pass
        return (
            None,
            False,
            missing,
            u"El parámetro «{0}» no está disponible en cuadros de armadura.\n"
            u"Verifique que esté vinculado a la categoría Rebar.".format(
                ARMADURA_UBICACION_PARAM
            ),
        )

    _clear_filters(defn)
    ok, err = _add_host_category_floor_filter(defn, host_field)
    if not ok:
        if created:
            try:
                doc.Delete(schedule.Id)
            except Exception:
                pass
        return None, False, missing, u"No se pudo filtrar la plantilla: {0}".format(err)

    _hide_filter_only_fields(defn, doc)

    action = u"Creada" if created else u"Actualizada"
    return (
        schedule,
        created,
        missing,
        u"{0} plantilla «{1}».".format(action, TEMPLATE_SCHEDULE_NAME),
    )


def _duplicate_from_template(doc, template, level_name, malla_variant, ubicacion_variant):
    try:
        new_id = template.Duplicate(ViewDuplicateOption.Duplicate)
    except Exception as ex:
        return None, u"No se pudo duplicar la plantilla: {0}".format(ex)

    schedule = doc.GetElement(new_id)
    if schedule is None:
        return None, u"La duplicación de la plantilla no devolvió una vista."

    malla_suffix = malla_variant[u"suffix"]
    ubicacion_suffix = ubicacion_variant[u"suffix"]
    malla_yes = bool(malla_variant[u"yes"])
    ubicacion_value = ubicacion_variant[u"value"]
    target_name = schedule_name_for_variant(level_name, malla_suffix, ubicacion_suffix)
    try:
        schedule.Name = target_name
    except Exception:
        try:
            schedule.Name = u"{0} ({1})".format(target_name, schedule.Id.IntegerValue)
        except Exception:
            pass

    ok, err = _apply_losa_filters(
        schedule.Definition, doc, level_name, malla_yes, ubicacion_value
    )
    if not ok:
        try:
            doc.Delete(schedule.Id)
        except Exception:
            pass
        return None, err

    apply_schedule_identity_params(schedule, level_name)
    return schedule, u""


def crear_o_actualizar_cuadros_por_nivel(doc, levels=None):
    """
    Asegura plantilla y 4 tablas por Level seleccionado
    (Malla/Sin Malla × Superior/Inferior).

    Args:
        doc: Document
        levels: lista opcional de ``Level``. Si es None, usa todos los del proyecto.

    Los campos Rounded* no se pueden crear por API. Se agregan una vez en la
    plantilla (UI) y se heredan al duplicar. Si la plantilla tiene más Rounded*
    que una tabla existente, esa tabla se regenera desde la plantilla.

    Returns:
        dict: ok, message, template, created, updated, regenerated, failed,
        levels, missing_fields, missing_rounded, open_template
    """
    from lib.cuantificacion_progress import CuantificacionLosaProgress

    result = {
        u"ok": False,
        u"message": u"",
        u"template": None,
        u"created": [],
        u"updated": [],
        u"regenerated": [],
        u"failed": [],
        u"levels": [],
        u"missing_fields": [],
        u"missing_rounded": [],
        u"open_template": False,
    }

    if levels is None:
        levels = listar_niveles(doc)
    else:
        levels = list(levels or [])
    result[u"levels"] = [_as_unicode(lv.Name) for lv in levels if lv is not None]
    if not levels:
        result[u"message"] = u"No hay niveles seleccionados."
        return result

    expected = len(levels) * len(MALLA_VARIANTS) * len(UBICACION_VARIANTS)
    total_steps = expected + 1  # +1 plantilla
    identity_warn = []
    missing = []
    missing_rounded = []
    tpl_msg = u""
    n_rounded_tpl = 0

    with CuantificacionLosaProgress(total_steps) as pb:
        pb.step(u"Plantilla…")
        template, _tpl_created, missing, tpl_msg = asegurar_plantilla(doc)
        result[u"missing_fields"] = missing
        if template is None:
            result[u"message"] = tpl_msg
            return result
        result[u"template"] = template
        set_subclasificacion_host_category(template, HOST_CATEGORY_FLOOR_LABEL)

        missing_rounded = listar_rounded_faltantes(template.Definition)
        result[u"missing_rounded"] = missing_rounded
        result[u"open_template"] = False
        n_rounded_tpl = contar_campos_rounded(template.Definition)

        for lv in levels:
            level_name = _as_unicode(lv.Name) if lv is not None else u""
            if not level_name:
                result[u"failed"].append((u"(sin nombre)", u"Nivel sin nombre."))
                for _ in range(len(MALLA_VARIANTS) * len(UBICACION_VARIANTS)):
                    pb.step(u"(nivel sin nombre)")
                continue

            for malla_variant in MALLA_VARIANTS:
                for ubicacion_variant in UBICACION_VARIANTS:
                    malla_suffix = malla_variant[u"suffix"]
                    ubicacion_suffix = ubicacion_variant[u"suffix"]
                    malla_yes = bool(malla_variant[u"yes"])
                    ubicacion_value = ubicacion_variant[u"value"]
                    target = schedule_name_for_variant(
                        level_name, malla_suffix, ubicacion_suffix
                    )
                    pb.step(target)
                    existing = _find_schedule_by_name(doc, target)

                    if existing is not None:
                        n_rounded_ex = contar_campos_rounded(existing.Definition)
                        if n_rounded_tpl > n_rounded_ex:
                            try:
                                doc.Delete(existing.Id)
                            except Exception as ex:
                                result[u"failed"].append(
                                    (target, u"No se pudo regenerar: {0}".format(ex))
                                )
                                continue
                            schedule, err = _duplicate_from_template(
                                doc,
                                template,
                                level_name,
                                malla_variant,
                                ubicacion_variant,
                            )
                            if schedule is None:
                                result[u"failed"].append((target, err))
                            else:
                                result[u"regenerated"].append(
                                    _as_unicode(schedule.Name)
                                )
                            continue

                        ok, err = _apply_losa_filters(
                            existing.Definition,
                            doc,
                            level_name,
                            malla_yes,
                            ubicacion_value,
                        )
                        if ok:
                            for plabel, detail in apply_schedule_identity_params(
                                existing, level_name
                            ):
                                if len(identity_warn) < 5:
                                    identity_warn.append(
                                        u"{0} «{1}»: {2}".format(
                                            target, plabel, detail
                                        )
                                    )
                            result[u"updated"].append(target)
                        else:
                            result[u"failed"].append((target, err))
                        continue

                    schedule, err = _duplicate_from_template(
                        doc, template, level_name, malla_variant, ubicacion_variant
                    )
                    if schedule is None:
                        result[u"failed"].append((target, err))
                    else:
                        result[u"created"].append(_as_unicode(schedule.Name))

    n_ok = (
        len(result[u"created"])
        + len(result[u"updated"])
        + len(result[u"regenerated"])
    )
    n_fail = len(result[u"failed"])

    lines = [
        tpl_msg,
        u"Niveles seleccionados: {0} | Tablas esperadas: {1} "
        u"(nivel × malla × superior/inferior) | "
        u"Creadas: {2} | Actualizadas: {3} | Regeneradas: {4} | Fallidas: {5}".format(
            len(levels),
            expected,
            len(result[u"created"]),
            len(result[u"updated"]),
            len(result[u"regenerated"]),
            n_fail,
        ),
    ]
    if identity_warn:
        lines.append(
            u"Aviso parámetros de vista: {0}".format(u" | ".join(identity_warn))
        )
    if missing:
        lines.append(
            u"Campos no encontrados en la plantilla: {0}".format(u", ".join(missing))
        )
    if result[u"failed"]:
        for name, err in result[u"failed"][:5]:
            lines.append(u"· {0}: {1}".format(name, err))
        if len(result[u"failed"]) > 5:
            lines.append(u"· … y {0} más.".format(len(result[u"failed"]) - 5))

    result[u"message"] = u"\n".join(lines)
    result[u"ok"] = n_ok > 0
    return result
