# -*- coding: utf-8 -*-
"""
Cuantificación Tipo 2 — una tabla de Rebar por Host Category.

Duplica la vista «Plantilla Cuantificacion» y aplica un filtro
Host Category = <categoría> por cada Host Category presente en los
Structural Rebar del proyecto.

Revit 2024+ | pyRevit / IronPython.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    Category,
    FilteredElementCollector,
    ScheduleFilter,
    ScheduleFilterType,
    ScheduleSortGroupField,
    ScheduleSortOrder,
    StorageType,
    Transaction,
    ViewDuplicateOption,
    ViewSchedule,
)
from Autodesk.Revit.DB.Structure import Rebar, RebarHostCategory
from Autodesk.Revit.UI import TaskDialog

_TOOL_TITLE = u"Arainco: Cuantificación Tipo 2"
_TRANSACTION_NAME = u"Arainco: Cuantificación Tipo 2"

TEMPLATE_SCHEDULE_NAME = u"Plantilla Cuantificacion"
SCHEDULE_NAME_PREFIX = u"Arainco Cuantificación Tipo 2 - "

HOST_CATEGORY_PARAM = u"Host Category"
HOST_CATEGORY_ALIASES = (
    HOST_CATEGORY_PARAM,
    u"Categoría del anfitrión",
    u"Categoria del anfitrion",
    u"Categoría de anfitrión",
)

SUBCLASIFICACION_ALIASES = (
    u"Subclasificacion",
    u"Subclasificación",
)

CLASIFICACION_ALIASES = (
    u"Clasificacion",
    u"Clasificación",
)
CLASIFICACION_VALUE = u"Armadura Tipo 2"

ARMADURA_NIVEL_PARAM = u"Armadura_Nivel"
BAR_DIAMETER_ALIASES = (
    u"Bar Diameter",
    u"Diámetro de barra",
    u"Diametro de barra",
)

# Columnas ocultas en todas las tablas Tipo 2 (si existen en la plantilla).
FLOOR_HIDDEN_FIELD_SPECS = (
    (u"Armadura_Marca", (u"Armadura_Marca",)),
    (
        u"Shape",
        (
            u"Shape",
            u"Forma",
        ),
    ),
    (
        u"ShapeQuantity",
        (
            u"ShapeQuantity",
            u"Shape Quantity",
            u"Cantidad de forma",
        ),
    ),
    (u"Rounded A", (u"Rounded A",)),
    (u"Rounded B", (u"Rounded B",)),
    (u"Rounded C", (u"Rounded C",)),
    (u"Rounded D", (u"Rounded D",)),
    (u"Rounded E", (u"Rounded E",)),
    (u"Rounded F", (u"Rounded F",)),
    (u"Rounded G", (u"Rounded G",)),
    (u"Rounded H", (u"Rounded H",)),
    (u"Rounded J", (u"Rounded J",)),
    (u"A", (u"A",)),
    (u"B", (u"B",)),
    (u"C", (u"C",)),
    (u"D", (u"D",)),
    (u"E", (u"E",)),
    (u"F", (u"F",)),
    (u"G", (u"G",)),
    (u"H", (u"H",)),
    (u"J", (u"J",)),
    (u"K", (u"K",)),
    (u"O", (u"O",)),
    (u"R", (u"R",)),
    (
        u"LARGO PARCIAL",
        (
            u"LARGO PARCIAL",
            u"Largo Parcial",
            u"Largo parcial",
        ),
    ),
)


def _pbar_enabled():
    try:
        from pyrevit import forms as _forms  # noqa: F401

        return True
    except Exception:
        return False


class _Tipo2Progress(object):
    """Context manager no-op si pyRevit ProgressBar no está disponible."""

    def __init__(self, total, title_prefix=None):
        self._total = max(1, int(total or 1))
        self._index = 0
        self._pb = None
        self._open = False
        self._title_prefix = title_prefix or _TOOL_TITLE

    def __enter__(self):
        if not _pbar_enabled():
            return self
        try:
            from pyrevit import forms as _pyrevit_forms

            self._pb = _pyrevit_forms.ProgressBar(
                title=self._title(0),
                cancellable=False,
            )
            try:
                from System.Windows.Media import Color, SolidColorBrush

                self._pb.Resources[u"pyRevitAccentBrush"] = SolidColorBrush(
                    Color.FromRgb(91, 192, 222),
                )
            except Exception:
                pass
            self._pb.__enter__()
            self._open = True
        except Exception:
            self._pb = None
            self._open = False
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._open and self._pb is not None:
            try:
                self._pb.__exit__(exc_type, exc_val, exc_tb)
            except Exception:
                pass
        self._open = False
        self._pb = None
        return False

    def _title(self, index):
        return u"{0} {1}/{2}".format(
            self._title_prefix,
            int(index) + 1,
            int(self._total),
        )

    def step(self, phase_label):
        """Avanza un paso y actualiza título de la barra."""
        if self._pb is None:
            return
        i = int(self._index)
        if i >= self._total:
            i = self._total - 1
        self._index = i + 1
        label = phase_label or u""
        base = u"{0} — {1}".format(self._title(i), label) if label else self._title(i)
        try:
            if hasattr(self._pb, u"update_progress"):
                try:
                    self._pb.update_progress(i + 1, max_value=self._total)
                except TypeError:
                    try:
                        self._pb.update_progress(i + 1, max=self._total)
                    except Exception:
                        pass
        except Exception:
            pass
        try:
            self._pb.title = base
        except Exception:
            pass

# Nombre de miembro RebarHostCategory → etiqueta UI (inglés Revit).
_ENUM_MEMBER_LABELS = {
    u"Floor": u"Floor",
    u"Wall": u"Wall",
    u"StructuralFraming": u"Structural Framing",
    u"StructuralFoundation": u"Structural Foundation",
    u"Stairs": u"Stairs",
    u"StructuralColumn": u"Structural Column",
    u"Part": u"Part",
    u"FabricArea": u"Fabric Area",
    u"FabricSheet": u"Fabric Sheet",
    u"Rebar": u"Rebar",
    u"AreaReinforcement": u"Area Reinforcement",
    u"PathReinforcement": u"Path Reinforcement",
}


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


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


def _mostrar_aviso(uiapp, instruction, content=u"", ok_text=u"Entendido"):
    try:
        from bimtools_instruction_dialog import show_message_dialog
        from revit_wpf_window_position import revit_main_hwnd

        hwnd = revit_main_hwnd(uiapp) if uiapp is not None else None
        show_message_dialog(
            _TOOL_TITLE,
            instruction=_as_unicode(instruction),
            content=_as_unicode(content) if content else None,
            ok_text=ok_text,
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass
    msg = _as_unicode(instruction)
    if content:
        msg = u"{0}\n\n{1}".format(msg, _as_unicode(content))
    try:
        TaskDialog.Show(_TOOL_TITLE, msg)
    except Exception:
        pass


def _rebar_host_category_none_value():
    try:
        return int(getattr(RebarHostCategory, u"None"))
    except Exception:
        return 0


def _catalog_host_categories():
    """
    Catálogo (value:int → label:str) desde RebarHostCategory.
    Omite None / valores nulos.
    """
    catalog = {}
    none_val = _rebar_host_category_none_value()

    def _add(member_name, label):
        try:
            raw = getattr(RebarHostCategory, member_name)
            value = int(raw)
        except Exception:
            return
        if value == none_val:
            return
        catalog[value] = label

    for member_name, label in _ENUM_MEMBER_LABELS.items():
        _add(member_name, label)

    # Miembros adicionales del enum no listados arriba.
    names = []
    try:
        from System import Enum

        names = list(Enum.GetNames(RebarHostCategory))
    except Exception:
        names = []
    for member_name in names:
        mn = _as_unicode(member_name)
        if mn in (u"None", u""):
            continue
        try:
            value = int(getattr(RebarHostCategory, mn))
        except Exception:
            continue
        if value in catalog or value == none_val:
            continue
        label = _ENUM_MEMBER_LABELS.get(mn)
        if not label:
            spaced = []
            buf = u""
            for ch in mn:
                if ch.isupper() and buf:
                    spaced.append(buf)
                    buf = ch
                else:
                    buf += ch
            if buf:
                spaced.append(buf)
            label = u" ".join(spaced) if spaced else mn
        catalog[value] = label

    return catalog

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


def _read_host_category_value(rebar):
    """Devuelve int de Host Category o None si no aplica."""
    p = _lookup_param(rebar, HOST_CATEGORY_ALIASES)
    if p is None or not p.HasValue:
        return None
    try:
        if p.StorageType == StorageType.Integer:
            return int(p.AsInteger())
    except Exception:
        pass
    # Respaldo: algunos builds exponen string; mapear por etiqueta.
    text = u""
    try:
        text = p.AsString() or u""
    except Exception:
        pass
    if not text:
        try:
            text = p.AsValueString() or u""
        except Exception:
            text = u""
    text_n = _normalize_name(text)
    if not text_n:
        return None
    catalog = _catalog_host_categories()
    for value, label in catalog.items():
        if _normalize_name(label) == text_n:
            return value
    return None


def listar_host_categories_en_rebars(doc):
    """
    Host Categories presentes en Structural Rebar del proyecto.

    Returns:
        list of dict: [{u"value": int, u"label": str}, ...] ordenado por label.
    """
    catalog = _catalog_host_categories()
    none_val = _rebar_host_category_none_value()

    found = {}
    rebars = (
        FilteredElementCollector(doc)
        .OfClass(Rebar)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    for rb in rebars or []:
        value = _read_host_category_value(rb)
        if value is None or value == none_val:
            continue
        label = catalog.get(value)
        if not label:
            label = u"Host Category {0}".format(value)
        found[value] = label

    items = [
        {u"value": v, u"label": found[v]}
        for v in found
    ]
    items.sort(key=lambda it: _normalize_name(it[u"label"]))
    return items


def schedule_name_for_host(host_label):
    return u"{0}{1}".format(SCHEDULE_NAME_PREFIX, host_label)


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


def encontrar_plantilla(doc):
    """
    Localiza «Plantilla Cuantificacion» (debe existir en el proyecto).

    Returns:
        (ViewSchedule|None, mensaje_error)
    """
    schedule = _find_schedule_by_name(doc, TEMPLATE_SCHEDULE_NAME)
    if schedule is None:
        return (
            None,
            u"No se encontró la tabla «{0}».\n"
            u"Créela o cópiela al proyecto antes de usar esta herramienta.".format(
                TEMPLATE_SCHEDULE_NAME
            ),
        )
    if not _is_rebar_schedule(schedule):
        return (
            None,
            u"La vista «{0}» existe pero no es un cuadro de Structural Rebar.".format(
                TEMPLATE_SCHEDULE_NAME
            ),
        )
    return schedule, u""


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


def _find_schedulable_field(defn, doc, param_names):
    targets = {_normalize_name(n) for n in param_names}
    for sf in defn.GetSchedulableFields():
        name = _schedulable_field_name(defn, doc, sf)
        if _normalize_name(name) in targets:
            return sf, name
    return None, None


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
    if field is None:
        return
    try:
        field.IsHidden = True
    except Exception:
        pass


def _clear_filters(defn):
    while defn.GetFilterCount() > 0:
        defn.RemoveFilter(0)


def _try_add_filter(defn, field, filter_type, value):
    try:
        defn.AddFilter(ScheduleFilter(field.FieldId, filter_type, value))
        return True, u""
    except Exception as ex:
        return False, _as_unicode(ex)


def _add_host_category_filter(defn, host_field, host_value, host_label):
    """Filtra Host Category; prioriza entero RebarHostCategory."""
    ok, err = _try_add_filter(
        defn, host_field, ScheduleFilterType.Equal, int(host_value)
    )
    if ok:
        return True, u""
    for ftype, val in (
        (ScheduleFilterType.Equal, host_label),
        (ScheduleFilterType.Contains, host_label),
    ):
        ok2, err2 = _try_add_filter(defn, host_field, ftype, val)
        if ok2:
            return True, u""
        err = err2 or err
    return False, err or u"Filtro Host Category = {0} no válido.".format(host_label)


def _set_view_string_param(element, param_names, value):
    p = _lookup_param(element, param_names)
    if p is None:
        return False, u"Parámetro no encontrado."
    if p.IsReadOnly:
        return False, u"Parámetro de solo lectura."
    text = _as_unicode(value)
    try:
        if p.StorageType == StorageType.String:
            p.Set(text)
            return True, u""
    except Exception as ex:
        err = _as_unicode(ex)
    else:
        err = u""
    try:
        p.SetValueString(text)
        return True, u""
    except Exception as ex2:
        return False, err or _as_unicode(ex2)


def apply_schedule_identity_params(schedule, host_label):
    """
    Clasificacion = «Armadura Tipo 2»;
    Subclasificacion = Host Category (p. ej. Floor, Wall).
    """
    _set_view_string_param(schedule, CLASIFICACION_ALIASES, CLASIFICACION_VALUE)
    _set_view_string_param(schedule, SUBCLASIFICACION_ALIASES, host_label)


def _clear_sort_group_fields(defn):
    try:
        defn.ClearSortGroupFields()
        return
    except Exception:
        pass
    try:
        while defn.GetSortGroupFieldCount() > 0:
            defn.RemoveSortGroupField(0)
    except Exception:
        pass


def _apply_sort_group_chain(defn, field_specs):
    """
    Aplica Sort/Group en orden y desactiva itemizar.

    ``field_specs``: lista de (field, show_header, show_blank).
    """
    _clear_sort_group_fields(defn)
    try:
        for field, show_header, show_blank in field_specs:
            sg = ScheduleSortGroupField(field.FieldId, ScheduleSortOrder.Ascending)
            sg.ShowHeader = bool(show_header)
            sg.ShowBlankLine = bool(show_blank)
            defn.AddSortGroupField(sg)
    except Exception as ex:
        return False, u"No se pudo aplicar Sort/Group: {0}".format(_as_unicode(ex))
    try:
        defn.IsItemized = False
    except Exception as ex:
        return False, u"No se pudo desactivar Itemize: {0}".format(_as_unicode(ex))
    return True, u""


def _hide_floor_columns(defn, doc):
    """Oculta columnas de detalle en tablas Tipo 2 (si están presentes)."""
    for _label, aliases in FLOOR_HIDDEN_FIELD_SPECS:
        field = _find_schedule_field(defn, doc, aliases)
        _hide_field(field)


def _apply_nivel_diametro_sort_grouping(defn, doc):
    """
    Sort/Group (sin itemizar):
      1. Armadura_Nivel (piso; header + blank)
      2. Bar Diameter
    """
    nivel_field = _ensure_field(defn, doc, (ARMADURA_NIVEL_PARAM,))
    if nivel_field is None:
        return False, u"No se encontró el parámetro «{0}» en el cuadro.".format(
            ARMADURA_NIVEL_PARAM
        )
    diam_field = _ensure_field(defn, doc, BAR_DIAMETER_ALIASES)
    if diam_field is None:
        return False, u"No se encontró el parámetro «Bar Diameter» en el cuadro."

    ok, err = _apply_sort_group_chain(
        defn,
        (
            (nivel_field, True, True),
            (diam_field, False, False),
        ),
    )
    if not ok:
        return False, err
    # Nivel solo como agrupación; diámetro sigue visible como columna.
    _hide_field(nivel_field)
    return True, u""


def apply_sort_grouping_for_host(defn, doc, host_label):
    """Organización: nivel → diámetro; oculta las mismas columnas en todas las tablas."""
    ok, err = _apply_nivel_diametro_sort_grouping(defn, doc)
    if not ok:
        return False, err
    _hide_floor_columns(defn, doc)
    return True, u""


def apply_host_category_filter(defn, doc, host_value, host_label):
    """Limpia filtros, aplica Host Category y Sort/Group según anfitrión."""
    host_field = _ensure_field(defn, doc, HOST_CATEGORY_ALIASES)
    if host_field is None:
        return False, u"El parámetro «{0}» no está disponible en el cuadro.".format(
            HOST_CATEGORY_PARAM
        )
    _clear_filters(defn)
    ok, err = _add_host_category_filter(defn, host_field, host_value, host_label)
    if not ok:
        return False, u"No se pudo filtrar Host Category: {0}".format(err)
    _hide_field(host_field)

    ok, err = apply_sort_grouping_for_host(defn, doc, host_label)
    if not ok:
        return False, err
    return True, u""


def _rename_schedule(schedule, target_name):
    try:
        schedule.Name = target_name
    except Exception:
        try:
            schedule.Name = u"{0} ({1})".format(
                target_name, schedule.Id.IntegerValue
            )
        except Exception:
            pass


def _duplicate_for_host(doc, template, host_value, host_label):
    try:
        new_id = template.Duplicate(ViewDuplicateOption.Duplicate)
    except Exception as ex:
        return None, u"No se pudo duplicar la plantilla: {0}".format(ex)

    schedule = doc.GetElement(new_id)
    if schedule is None:
        return None, u"La duplicación de la plantilla no devolvió una vista."

    target_name = schedule_name_for_host(host_label)
    _rename_schedule(schedule, target_name)

    ok, err = apply_host_category_filter(
        schedule.Definition, doc, host_value, host_label
    )
    if not ok:
        try:
            doc.Delete(schedule.Id)
        except Exception:
            pass
        return None, err

    apply_schedule_identity_params(schedule, host_label)
    return schedule, u""


def crear_tablas_por_host_category(doc, hosts=None):
    """
    Crea o actualiza una tabla por Host Category.

    Returns:
        dict con created / updated / failed / hosts
    """
    result = {
        u"created": [],
        u"updated": [],
        u"failed": [],
        u"hosts": [],
    }

    template, err = encontrar_plantilla(doc)
    if template is None:
        result[u"failed"].append((TEMPLATE_SCHEDULE_NAME, err))
        return result

    if hosts is None:
        hosts = listar_host_categories_en_rebars(doc)
    result[u"hosts"] = list(hosts)

    if not hosts:
        result[u"failed"].append(
            (
                u"(ninguna)",
                u"No hay Structural Rebar con Host Category en el proyecto.",
            )
        )
        return result

    with _Tipo2Progress(len(hosts), title_prefix=_TOOL_TITLE) as pb:
        for host in hosts:
            host_value = host[u"value"]
            host_label = host[u"label"]
            target = schedule_name_for_host(host_label)
            pb.step(host_label)
            existing = _find_schedule_by_name(doc, target)

            if existing is not None:
                ok, err = apply_host_category_filter(
                    existing.Definition, doc, host_value, host_label
                )
                if ok:
                    apply_schedule_identity_params(existing, host_label)
                    result[u"updated"].append(_as_unicode(existing.Name))
                else:
                    result[u"failed"].append((target, err))
                continue

            schedule, err = _duplicate_for_host(
                doc, template, host_value, host_label
            )
            if schedule is None:
                result[u"failed"].append((target, err))
            else:
                result[u"created"].append(_as_unicode(schedule.Name))

    return result


def _format_result_summary(result):
    lines = []
    hosts = result.get(u"hosts") or []
    if hosts:
        labels = [h[u"label"] for h in hosts]
        lines.append(
            u"Host Category detectadas ({0}): {1}".format(
                len(labels), u", ".join(labels)
            )
        )
    created = result.get(u"created") or []
    updated = result.get(u"updated") or []
    failed = result.get(u"failed") or []
    if created:
        lines.append(u"Creadas ({0}):".format(len(created)))
        for name in created:
            lines.append(u"  · {0}".format(name))
    if updated:
        lines.append(u"Actualizadas ({0}):".format(len(updated)))
        for name in updated:
            lines.append(u"  · {0}".format(name))
    if failed:
        lines.append(u"Fallidas ({0}):".format(len(failed)))
        for name, err in failed:
            lines.append(u"  · {0}: {1}".format(name, err))
    if not lines:
        return u"Sin cambios."
    return u"\n".join(lines)


def run(uiapp):
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        _mostrar_aviso(uiapp, u"No hay documento activo.")
        return

    doc = uidoc.Document
    if doc.IsFamilyDocument:
        _mostrar_aviso(
            uiapp,
            u"Esta herramienta solo funciona en un proyecto (no en familias).",
        )
        return

    template, err = encontrar_plantilla(doc)
    if template is None:
        _mostrar_aviso(uiapp, err)
        return

    hosts = listar_host_categories_en_rebars(doc)
    if not hosts:
        _mostrar_aviso(
            uiapp,
            u"No hay Structural Rebar con Host Category en el proyecto.",
        )
        return

    t = Transaction(doc, _TRANSACTION_NAME)
    t.Start()
    try:
        result = crear_tablas_por_host_category(doc, hosts=hosts)
        t.Commit()
    except Exception as ex:
        try:
            if t.HasStarted():
                t.RollBack()
        except Exception:
            pass
        _mostrar_aviso(
            uiapp,
            u"Error al generar las tablas.",
            _as_unicode(ex),
        )
        return

    n_ok = len(result.get(u"created") or []) + len(result.get(u"updated") or [])
    n_fail = len(result.get(u"failed") or [])
    if n_fail and not n_ok:
        instruction = u"No se pudo generar ninguna tabla."
    elif n_fail:
        instruction = u"Tablas generadas con advertencias."
    else:
        instruction = u"Tablas de cuantificación Tipo 2 listas."

    _mostrar_aviso(uiapp, instruction, _format_result_summary(result))
