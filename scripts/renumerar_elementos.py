# -*- coding: utf-8 -*-
"""
Arainco: Renumerar elementos — lógica de negocio.

Renumera elementos en el orden de selección (habitaciones, áreas, espacios,
puertas, muros, ventanas, estacionamientos, niveles, ejes y viewports).
Revit 2024+ | IronPython (pyRevit).
"""

from __future__ import print_function

try:
    unicode
except NameError:
    unicode = str

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    Category,
    FilteredElementCollector,
    Grid,
    Level,
    OverrideGraphicSettings,
    TemporaryViewMode,
    Transaction,
    TransactionGroup,
    View3D,
    ViewPlan,
    ViewSection,
    ViewSheet,
    Viewport,
    ViewType,
)
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

BIC = BuiltInCategory

ALLOWED_VIEW_CLASSES = (View3D, ViewPlan, ViewSection, ViewSheet)

DUPE_SWEEP = u"sweep"
DUPE_ALERT = u"alert"
DUPE_SKIP = u"skip"

DUPE_LABELS = (
    (DUPE_SWEEP, u"Barrer (desplaza el que ocupa el número)"),
    (DUPE_ALERT, u"Alertar y omitir (no asigna duplicados)"),
    (DUPE_SKIP, u"Omitir si ya tiene número"),
)

KEY_ROOMS = u"rooms"
KEY_AREAS = u"areas"
KEY_SPACES = u"spaces"
KEY_DOORS = u"doors"
KEY_DOORS_BY_ROOM = u"doors_by_room"
KEY_WALLS = u"walls"
KEY_WINDOWS = u"windows"
KEY_PARKING = u"parking"
KEY_LEVELS = u"levels"
KEY_GRIDS = u"grids"
KEY_VIEWPORTS = u"viewports"

_SPATIAL_SUBCATS = {
    BIC.OST_Rooms: (u"OST_RoomReference", u"OST_RoomInteriorFill"),
    BIC.OST_Areas: (u"OST_AreaReference", u"OST_AreaInteriorFill"),
    BIC.OST_MEPSpaces: (u"OST_MEPSpaceReference", u"OST_MEPSpaceInteriorFill"),
}

TX_GROUP = u"Arainco: Renumerar elementos"
TX_ITEM = u"Arainco: Renumerar {0}"
TX_HANDLES = u"Arainco: Asas de selección"
TX_UNMARK = u"Arainco: Restaurar gráficos"
TX_DOOR = u"Arainco: Renumerar puerta"


def _u(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except Exception:
        return str(text)


class RenumberOption(object):
    def __init__(self, key, label, bic, by_bic=None):
        self.key = key
        self.label = label
        self.bic = bic
        self.by_bic = by_bic

    @property
    def is_doors_by_room(self):
        return self.key == KEY_DOORS_BY_ROOM


class RenumberStats(object):
    def __init__(self):
        self.renumbered = 0
        self.skipped_session = 0
        self.skipped_dupe = 0
        self.skipped_already_numbered = 0
        self.displaced = 0
        self.failed = 0
        self.alerts = []

    def add_alert(self, message):
        msg = _u(message).strip()
        if msg and msg not in self.alerts:
            self.alerts.append(msg)


def list_options_for_view(view):
    """Opciones de categoría según el tipo de vista activa."""
    if isinstance(view, ViewSheet):
        return [
            RenumberOption(KEY_VIEWPORTS, u"Viewports (n.º de detalle)", BIC.OST_Viewports),
        ]
    options = [
        RenumberOption(KEY_ROOMS, u"Habitaciones", BIC.OST_Rooms),
        RenumberOption(KEY_SPACES, u"Espacios MEP", BIC.OST_MEPSpaces),
        RenumberOption(KEY_DOORS, u"Puertas", BIC.OST_Doors),
        RenumberOption(
            KEY_DOORS_BY_ROOM,
            u"Puertas por habitación",
            BIC.OST_Doors,
            by_bic=BIC.OST_Rooms,
        ),
        RenumberOption(KEY_WALLS, u"Muros", BIC.OST_Walls),
        RenumberOption(KEY_WINDOWS, u"Ventanas", BIC.OST_Windows),
        RenumberOption(KEY_PARKING, u"Estacionamientos", BIC.OST_Parking),
        RenumberOption(KEY_LEVELS, u"Niveles", BIC.OST_Levels),
        RenumberOption(KEY_GRIDS, u"Ejes", BIC.OST_Grids),
    ]
    try:
        if view.ViewType == ViewType.AreaPlan:
            options.insert(
                1,
                RenumberOption(KEY_AREAS, u"Áreas", BIC.OST_Areas),
            )
    except Exception:
        pass
    return options


def is_allowed_view(view):
    return isinstance(view, ALLOWED_VIEW_CLASSES)


def allowed_view_names():
    return u"planta, área, 3D, sección o lámina"


# ---------------------------------------------------------------------------
# Incremento alfanumérico (misma lógica que pyRevit; padding opcional)
# ---------------------------------------------------------------------------

def _inc_or_dec_string(str_id, shift, refit=False):
    if not str_id:
        return u"1" if shift >= 0 else u"0"
    if shift == 0:
        return str_id
    next_str = u""
    index = len(str_id) - 1
    carry = shift
    while index >= 0:
        this_char = str_id[index]
        if this_char.isdigit():
            char_range = (u"0", u"9")
        elif this_char.isalpha():
            if this_char.islower():
                char_range = (u"a", u"z")
            else:
                char_range = (u"A", u"Z")
        else:
            next_str += this_char
            index -= 1
            continue
        direction = int(carry / abs(carry)) if carry != 0 else 1
        start_char, end_char = char_range if direction > 0 else char_range[::-1]
        char_steps = abs(ord(end_char) - ord(start_char)) + 1
        dist = abs(ord(this_char) - ord(start_char))
        offset = (dist + abs(carry)) % char_steps
        next_char = chr(ord(start_char) + (offset * direction))
        next_str += next_char
        carry = int((dist + abs(carry)) / char_steps) * direction
        index -= 1
        if refit and index == -1:
            if carry > 0:
                str_id = start_char + str_id
                if start_char.isalpha():
                    carry -= 1
                index = 0
            elif direction == -1:
                if next_str.endswith(start_char):
                    next_str = next_str[:-1]
                else:
                    while next_str.endswith(end_char):
                        next_str = next_str[:-1]
    return next_str[::-1]


def increment_number(number, preserve_padding=True):
    """Incrementa el identificador. Si preserve_padding, mantiene ceros a la izquierda."""
    raw = _u(number).strip()
    if not raw:
        return u"1"
    return _inc_or_dec_string(raw, 1, refit=not preserve_padding)


def extend_counter(number):
    """Añade un nivel al identificador: 101 → 101A."""
    raw = _u(number).strip()
    if not raw:
        return u"A"
    if raw[-1].isdigit():
        return raw + u"A"
    return raw + u"1"


# ---------------------------------------------------------------------------
# Lectura / escritura del número
# ---------------------------------------------------------------------------

def get_number(element):
    if element is None:
        return None
    if hasattr(element, u"Number"):
        try:
            val = element.Number
            if val is not None:
                return _u(val)
        except Exception:
            pass
    param = None
    try:
        if isinstance(element, (Level, Grid)):
            param = element.Parameter[BuiltInParameter.DATUM_TEXT]
        elif isinstance(element, Viewport):
            param = element.Parameter[BuiltInParameter.VIEWPORT_DETAIL_NUMBER]
        else:
            param = element.Parameter[BuiltInParameter.ALL_MODEL_MARK]
    except Exception:
        param = None
    if param is None:
        return None
    try:
        val = param.AsString()
        if val is None:
            return None
        return _u(val)
    except Exception:
        return None


def set_number(element, new_number):
    value = _u(new_number)
    if element is None:
        return False
    if hasattr(element, u"Number"):
        try:
            element.Number = value
            return True
        except Exception:
            return False
    param = None
    try:
        if isinstance(element, (Level, Grid)):
            param = element.Parameter[BuiltInParameter.DATUM_TEXT]
        elif isinstance(element, Viewport):
            param = element.Parameter[BuiltInParameter.VIEWPORT_DETAIL_NUMBER]
        else:
            param = element.Parameter[BuiltInParameter.ALL_MODEL_MARK]
    except Exception:
        param = None
    if param is None:
        return False
    try:
        if param.IsReadOnly:
            return False
        param.Set(value)
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Mapa número → ids (varios ocupantes por número)
# ---------------------------------------------------------------------------

def _add_mapping(mapping, number, eid):
    key = _u(number).strip()
    if not key:
        return
    bucket = mapping.get(key)
    if bucket is None:
        mapping[key] = [eid]
    elif eid not in bucket:
        bucket.append(eid)


def _remove_mapping(mapping, number, eid):
    key = _u(number).strip()
    if not key or key not in mapping:
        return
    mapping[key] = [x for x in mapping[key] if x != eid]
    if not mapping[key]:
        mapping.pop(key)


def _first_other_occupant(mapping, number, target_id):
    key = _u(number).strip()
    for eid in mapping.get(key) or []:
        if eid != target_id:
            return eid
    return None


def collect_number_map(doc, views, builtin_cat):
    mapping = {}
    if builtin_cat == BIC.OST_Viewports:
        for view in views:
            if isinstance(view, ViewSheet):
                for vpid in view.GetAllViewports():
                    vp = doc.GetElement(vpid)
                    _add_mapping(mapping, get_number(vp), vpid)
                return mapping
    try:
        col = (
            FilteredElementCollector(doc)
            .OfCategory(builtin_cat)
            .WhereElementIsNotElementType()
        )
        for el in col:
            _add_mapping(mapping, get_number(el), el.Id)
    except Exception:
        pass
    return mapping


def find_replacement_number(existing_number, mapping, preserve_padding):
    replaced = increment_number(existing_number, preserve_padding=preserve_padding)
    guard = 0
    while replaced in mapping and guard < 10000:
        replaced = increment_number(replaced, preserve_padding=preserve_padding)
        guard += 1
    return replaced


# ---------------------------------------------------------------------------
# Gráficos temporales
# ---------------------------------------------------------------------------

def _category_from_bic(doc, bic):
    try:
        return Category.GetCategory(doc, bic)
    except Exception:
        return None


def _toggle_spatial_handles(doc, views, bicat, state):
    ref_name, fill_name = _SPATIAL_SUBCATS.get(bicat, (None, None))
    ref_bic = getattr(BIC, ref_name, None) if ref_name else None
    fill_bic = getattr(BIC, fill_name, None) if fill_name else None
    rr_cat = _category_from_bic(doc, ref_bic) if ref_bic else None
    rr_int = _category_from_bic(doc, fill_bic) if fill_bic else None
    enabled_temp = []
    t = Transaction(doc, TX_HANDLES)
    t.Start()
    try:
        if state and bicat != BIC.OST_Viewports:
            for view in views:
                temp_on = False
                try:
                    if view.CanEnableTemporaryViewPropertiesMode():
                        temp_on = bool(
                            view.EnableTemporaryViewPropertiesMode(view.Id)
                        )
                except Exception:
                    temp_on = False
                if not temp_on:
                    continue
                enabled_temp.append(view)
                for sub in (rr_cat, rr_int):
                    if sub is None:
                        continue
                    try:
                        sub.set_Visible(view, state)
                    except Exception:
                        try:
                            sub.Visible[view] = state
                        except Exception:
                            pass
        if not state:
            for view in views:
                try:
                    view.DisableTemporaryViewMode(
                        TemporaryViewMode.TemporaryViewProperties
                    )
                except Exception:
                    pass
        t.Commit()
    except Exception:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
    return enabled_temp


def mark_element_as_renumbered(views, element):
    ogs = OverrideGraphicSettings()
    ogs.SetHalftone(True)
    ogs.SetSurfaceTransparency(100)
    for view in views:
        try:
            view.SetElementOverrides(element.Id, ogs)
        except Exception:
            pass


def unmark_elements(doc, views, saved_overrides):
    t = Transaction(doc, TX_UNMARK)
    t.Start()
    try:
        for elid, view_dict in saved_overrides.items():
            for view in views:
                try:
                    if view.Id in view_dict:
                        ogs = view_dict[view.Id]
                    else:
                        ogs = OverrideGraphicSettings()
                    view.SetElementOverrides(elid, ogs)
                except Exception:
                    pass
        t.Commit()
    except Exception:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()


def get_open_allowed_views(uidoc, doc, for_viewports):
    if for_viewports:
        return [doc.ActiveView]
    views = []
    try:
        for ui_view in uidoc.GetOpenUIViews():
            view = doc.GetElement(ui_view.ViewId)
            if isinstance(view, ALLOWED_VIEW_CLASSES):
                views.append(view)
    except Exception:
        pass
    if not views and doc.ActiveView is not None:
        views = [doc.ActiveView]
    return views


# ---------------------------------------------------------------------------
# Renumerar un elemento
# ---------------------------------------------------------------------------

def _displace_occupant(doc, mapping, occupant_id, preserve_padding, stats):
    occupant = doc.GetElement(occupant_id)
    if occupant is None:
        return False
    current = get_number(occupant)
    if not current:
        return False
    replaced = find_replacement_number(current, mapping, preserve_padding)
    if not set_number(occupant, replaced):
        return False
    _remove_mapping(mapping, current, occupant_id)
    _add_mapping(mapping, replaced, occupant_id)
    stats.displaced += 1
    return True


def renumber_element(
    doc,
    views,
    target,
    new_number,
    mapping,
    dupe_mode,
    preserve_padding,
    stats,
):
    """Aplica new_number respetando el modo de duplicados. Devuelve True si numeró."""
    wanted = _u(new_number).strip()
    if not wanted:
        stats.failed += 1
        return False

    occupant_id = _first_other_occupant(mapping, wanted, target.Id)
    if occupant_id is not None:
        if dupe_mode == DUPE_ALERT:
            stats.skipped_dupe += 1
            stats.add_alert(
                u"Número duplicado «{0}»: elemento omitido.".format(wanted)
            )
            return False
        if dupe_mode == DUPE_SKIP:
            existing = get_number(target)
            if existing:
                stats.skipped_already_numbered += 1
                return False
            if not _displace_occupant(
                doc, mapping, occupant_id, preserve_padding, stats
            ):
                stats.failed += 1
                return False
        else:
            if not _displace_occupant(
                doc, mapping, occupant_id, preserve_padding, stats
            ):
                stats.failed += 1
                return False

    existing = get_number(target)
    _remove_mapping(mapping, existing, target.Id)
    if not set_number(target, wanted):
        if existing:
            _add_mapping(mapping, existing, target.Id)
        stats.failed += 1
        return False
    _add_mapping(mapping, wanted, target.Id)
    mark_element_as_renumbered(views, target)
    stats.renumbered += 1
    return True


# ---------------------------------------------------------------------------
# Selección
# ---------------------------------------------------------------------------

class _CategoryFilter(ISelectionFilter):
    def __init__(self, category_id):
        self._cid = category_id

    def AllowElement(self, elem):
        try:
            cat = elem.Category
            return cat is not None and cat.Id == self._cid
        except Exception:
            return False

    def AllowReference(self, reference, position):
        return False


def _category_id(doc, bic):
    cat = _category_from_bic(doc, bic)
    if cat is None:
        return None
    try:
        return cat.Id
    except Exception:
        return None


def pick_element_by_category(uidoc, doc, bic, message):
    cid = _category_id(doc, bic)
    if cid is None:
        return None
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            _CategoryFilter(cid),
            _u(message),
        )
    except OperationCanceledException:
        return None
    except Exception:
        return None
    if ref is None:
        return None
    try:
        return doc.GetElement(ref.ElementId)
    except Exception:
        return None


def _save_overrides(views, element, saved):
    if element.Id in saved:
        return
    saved[element.Id] = {}
    for view in views:
        try:
            saved[element.Id][view.Id] = view.GetElementOverrides(element.Id)
        except Exception:
            pass


def _warning_bar(title):
    try:
        from pyrevit import forms

        return forms.WarningBar(title=_u(title))
    except Exception:
        return None


class _NullBar(object):
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False


def _bar_or_null(title):
    bar = _warning_bar(title)
    if bar is None:
        return _NullBar()
    return bar


# ---------------------------------------------------------------------------
# Puertas / habitaciones
# ---------------------------------------------------------------------------

def get_door_rooms(door):
    try:
        phase = door.Document.GetElement(door.CreatedPhaseId)
        return door.FromRoom[phase], door.ToRoom[phase]
    except Exception:
        return None, None


def doors_for_room(doc, room_id):
    result = []
    try:
        col = (
            FilteredElementCollector(doc)
            .OfCategory(BIC.OST_Doors)
            .WhereElementIsNotElementType()
        )
        for door in col:
            from_room, to_room = get_door_rooms(door)
            if from_room is not None and from_room.Id == room_id:
                result.append(door)
            elif to_room is not None and to_room.Id == room_id:
                result.append(door)
    except Exception:
        pass
    return result


# ---------------------------------------------------------------------------
# Rutinas principales
# ---------------------------------------------------------------------------

def pick_and_renumber(
    uidoc,
    doc,
    option,
    starting_number,
    dupe_mode,
    preserve_padding,
    skip_session,
    stats,
):
    views = get_open_allowed_views(
        uidoc, doc, for_viewports=(option.bic == BIC.OST_Viewports)
    )
    index = _u(starting_number).strip()
    spatial = option.bic in _SPATIAL_SUBCATS
    tg = TransactionGroup(doc, TX_GROUP)
    tg.Start()
    try:
        if spatial:
            _toggle_spatial_handles(doc, views, option.bic, True)
        mapping = collect_number_map(doc, views, option.bic)
        saved = {}
        processed = set()
        bar_title = (
            u"Seleccione {0} en orden. ESC termina. Siguiente: {1}"
        ).format(option.label.lower(), index)
        with _bar_or_null(bar_title):
            while True:
                prompt = (
                    u"Seleccione {0} (siguiente: {1}). ESC para terminar."
                ).format(option.label.lower(), index)
                picked = pick_element_by_category(
                    uidoc, doc, option.bic, prompt
                )
                if picked is None:
                    break
                if skip_session and picked.Id.IntegerValue in processed:
                    stats.skipped_session += 1
                    continue
                t = Transaction(doc, TX_ITEM.format(option.label))
                t.Start()
                try:
                    _save_overrides(views, picked, saved)
                    applied = renumber_element(
                        doc,
                        views,
                        picked,
                        index,
                        mapping,
                        dupe_mode,
                        preserve_padding,
                        stats,
                    )
                    t.Commit()
                except Exception:
                    if t.HasStarted() and not t.HasEnded():
                        t.RollBack()
                    stats.failed += 1
                    applied = False
                if applied:
                    processed.add(picked.Id.IntegerValue)
                    index = increment_number(index, preserve_padding=preserve_padding)
        unmark_elements(doc, views, saved)
        if spatial:
            _toggle_spatial_handles(doc, views, option.bic, False)
        tg.Assimilate()
    except Exception:
        try:
            if tg.HasStarted() and not tg.HasEnded():
                tg.RollBack()
        except Exception:
            pass
        raise
    return stats


def door_by_room_renumber(
    uidoc,
    doc,
    option,
    dupe_mode,
    preserve_padding,
    skip_session,
    stats,
):
    views = get_open_allowed_views(uidoc, doc, for_viewports=False)
    tg = TransactionGroup(doc, u"Arainco: Renumerar puertas por habitación")
    tg.Start()
    try:
        _toggle_spatial_handles(doc, views, option.bic, True)
        _toggle_spatial_handles(doc, views, option.by_bic, True)
        mapping = collect_number_map(doc, views, option.bic)
        saved = {}
        processed = set()
        with _bar_or_null(u"Seleccione pares puerta + habitación. ESC termina."):
            while True:
                picked_door = pick_element_by_category(
                    uidoc,
                    doc,
                    option.bic,
                    u"Seleccione una puerta. ESC para terminar.",
                )
                if picked_door is None:
                    break
                if skip_session and picked_door.Id.IntegerValue in processed:
                    stats.skipped_session += 1
                    continue
                from_room, to_room = get_door_rooms(picked_door)
                both = from_room is not None and to_room is not None
                none = from_room is None and to_room is None
                if both or none:
                    picked_room = pick_element_by_category(
                        uidoc,
                        doc,
                        option.by_bic,
                        u"Seleccione la habitación asociada. ESC cancela este par.",
                    )
                    if picked_room is None:
                        continue
                else:
                    picked_room = from_room or to_room
                room_doors = doors_for_room(doc, picked_room.Id)
                room_number = get_number(picked_room)
                if not room_number:
                    stats.failed += 1
                    stats.add_alert(
                        u"La habitación no tiene número; puerta omitida."
                    )
                    continue
                t = Transaction(doc, TX_DOOR)
                t.Start()
                try:
                    _save_overrides(views, picked_door, saved)
                    door_count = len(room_doors)
                    if door_count <= 1:
                        wanted = room_number
                    else:
                        room_door_numbers = [
                            get_number(x) for x in room_doors if get_number(x)
                        ]
                        wanted = extend_counter(room_number)
                        guard = 0
                        while wanted in room_door_numbers and guard < 500:
                            wanted = increment_number(
                                wanted, preserve_padding=preserve_padding
                            )
                            guard += 1
                    applied = renumber_element(
                        doc,
                        views,
                        picked_door,
                        wanted,
                        mapping,
                        dupe_mode,
                        preserve_padding,
                        stats,
                    )
                    t.Commit()
                except Exception:
                    if t.HasStarted() and not t.HasEnded():
                        t.RollBack()
                    stats.failed += 1
                    applied = False
                if applied:
                    processed.add(picked_door.Id.IntegerValue)
        unmark_elements(doc, views, saved)
        _toggle_spatial_handles(doc, views, option.bic, False)
        _toggle_spatial_handles(doc, views, option.by_bic, False)
        tg.Assimilate()
    except Exception:
        try:
            if tg.HasStarted() and not tg.HasEnded():
                tg.RollBack()
        except Exception:
            pass
        raise
    return stats


def format_summary(stats, option_label):
    lines = [
        u"Categoría: {0}".format(option_label),
        u"Numerados: {0}".format(stats.renumbered),
    ]
    if stats.displaced:
        lines.append(u"Desplazados (barrido): {0}".format(stats.displaced))
    if stats.skipped_dupe:
        lines.append(u"Omitidos por duplicado: {0}".format(stats.skipped_dupe))
    if stats.skipped_already_numbered:
        lines.append(
            u"Omitidos (ya tenían número): {0}".format(
                stats.skipped_already_numbered
            )
        )
    if stats.skipped_session:
        lines.append(
            u"Omitidos (ya numerados en esta sesión): {0}".format(
                stats.skipped_session
            )
        )
    if stats.failed:
        lines.append(u"Fallidos: {0}".format(stats.failed))
    if stats.alerts:
        lines.append(u"")
        lines.append(u"Avisos:")
        for alert in stats.alerts[:12]:
            lines.append(u"• {0}".format(alert))
        if len(stats.alerts) > 12:
            lines.append(
                u"• … y {0} más.".format(len(stats.alerts) - 12)
            )
    if stats.renumbered < 1 and stats.failed < 1 and stats.skipped_dupe < 1:
        instruction = u"No se numeró ningún elemento."
    else:
        instruction = u"Renumeración finalizada: {0} elemento(s).".format(
            stats.renumbered
        )
    return instruction, u"\n".join(lines)
