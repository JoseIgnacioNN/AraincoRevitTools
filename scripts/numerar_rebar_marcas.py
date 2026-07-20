# -*- coding: utf-8 -*-
"""
Numerar marcas de Rebar en el documento completo.

Agrupa barras estructurales (``Rebar`` y ``RebarInSystem`` de Area Reinforcement)
por fingerprint:

  - diámetro nominal (mm)
  - shape (nombre de ``RebarShape``)
  - longitudes de segmentos A/B/C/… redondeadas a 10 mm

No entran ganchos, end treatments, capa/cara ni longitud total aparte de los
segmentos de forma. La secuencia de tramos se normaliza si la barra se lee
invertida.

Asigna marca ``{ø}{nº}`` por diámetro (serie desde 01) y escribe en
``Armadura_Marca``.

En **cada corrida** reevalúa todas las barras del documento y reasigna índices
desde cero por diámetro (orden estable por ``ElementId`` mínimo del grupo).
No conserva marcas previas: el plan completo se recalcula siempre.

Muestra ``pyrevit.forms.ProgressBar`` (acento BIMTools) durante el análisis
de fingerprints y durante la escritura de marcas.

Revit 2024+ | pyRevit / RPS (importable).
"""

from __future__ import print_function

import math

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    SpecTypeId,
    StorageType,
    Transaction,
    UnitUtils,
    UnitTypeId,
)
from Autodesk.Revit.DB.Structure import (
    AreaReinforcement,
    Rebar,
    RebarBarType,
    RebarInSystem,
)
from Autodesk.Revit.UI import TaskDialog

TOOL_TITLE = u"Arainco: Numerar marcas Rebar"
TRANSACTION_NAME = u"Arainco: Numerar marcas Rebar"
ARMADURA_MARCA_PARAM = u"Armadura_Marca"
PROGRESS_ACCENT_RGB = (91, 192, 222)


def _pbar_enabled():
    try:
        from pyrevit import forms as _forms  # noqa: F401
    except Exception:
        return False
    return True


class NumerarRebarProgress(object):
    """Context manager no-op si pyRevit ProgressBar no está disponible."""

    def __init__(self, total, title_prefix=None):
        self._total = max(1, int(total or 1))
        self._index = 0
        self._pb = None
        self._open = False
        self._title_prefix = title_prefix or TOOL_TITLE

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

                r, g, b = PROGRESS_ACCENT_RGB
                self._pb.Resources[u"pyRevitAccentBrush"] = SolidColorBrush(
                    Color.FromRgb(r, g, b),
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

    def step(self, phase_label=None):
        """Avanza un paso y actualiza título de la barra."""
        if self._pb is None:
            return
        i = int(self._index)
        if i >= self._total:
            i = self._total - 1
        self._index = i + 1
        label = phase_label or u""
        base = (
            u"{0} — {1}".format(self._title(i), label) if label else self._title(i)
        )
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

# Parámetros de segmento de forma (longitudes) a considerar en el matching.
SEGMENT_PARAM_NAMES = (
    u"A",
    u"B",
    u"C",
    u"D",
    u"E",
    u"F",
    u"G",
    u"H",
    u"J",
    u"K",
    u"O",
    u"R",
)

ROUND_STEP_MM = 10.0


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _element_id_int(eid):
    if eid is None:
        return -1
    try:
        if eid == ElementId.InvalidElementId:
            return -1
    except Exception:
        pass
    try:
        return int(eid.Value)
    except Exception:
        try:
            return int(eid.IntegerValue)
        except Exception:
            return -1


def _roundup_10mm(mm):
    """``ceil(mm / 10) * 10``; valores <= 0 se normalizan a 0."""
    try:
        v = float(mm)
    except Exception:
        return 0
    if v <= 0.0:
        return 0
    return int(math.ceil(v / ROUND_STEP_MM) * int(ROUND_STEP_MM))


def _format_index(index):
    """Padding mínimo 2 dígitos; desde 100 sin cero extra."""
    n = int(index)
    if n < 100:
        return u"{:02d}".format(n)
    return u"{}".format(n)


def format_mark(diameter_mm, index):
    return u"{}{}".format(int(diameter_mm), _format_index(index))


def _get_armadura_marca(bar):
    """Valor actual de ``Armadura_Marca`` (texto) o cadena vacía."""
    if bar is None:
        return u""
    try:
        p = bar.LookupParameter(ARMADURA_MARCA_PARAM)
    except Exception:
        p = None
    if p is None:
        return u""
    try:
        st = p.StorageType
        if st == StorageType.String:
            return _as_unicode(p.AsString() or u"").strip()
    except Exception:
        pass
    try:
        return _as_unicode(p.AsValueString() or u"").strip()
    except Exception:
        return u""


def _assign_indices_full(group_items, diameter_mm):
    """
    Asigna índice secuencial 1..n por grupo, ignorando marcas existentes.

    ``group_items``: lista ``(min_id, fp, members)`` ya ordenada por min_id.

    Returns:
        list of (fp, members, mark, index)
    """
    result = []
    for index, (_min_id, fp, members) in enumerate(group_items, start=1):
        mark = format_mark(diameter_mm, index)
        result.append((fp, members, mark, index))
    return result


def _bar_diameter_mm(bar_type):
    """Diámetro nominal entero (mm) desde ``RebarBarType``."""
    if bar_type is None or not isinstance(bar_type, RebarBarType):
        return None
    for attr in ("BarNominalDiameter", "BarModelDiameter", "BarDiameter"):
        try:
            raw = getattr(bar_type, attr, None)
            if raw is None:
                continue
            mm = UnitUtils.ConvertFromInternalUnits(float(raw), UnitTypeId.Millimeters)
            if mm > 0:
                return int(round(mm))
        except Exception:
            continue
    return None


def _is_length_param(param):
    if param is None:
        return False
    try:
        if param.StorageType != StorageType.Double:
            return False
    except Exception:
        return False
    try:
        dt = param.Definition.GetDataType()
        if dt is not None and SpecTypeId.Length is not None:
            return dt == SpecTypeId.Length
    except Exception:
        pass
    return False


def _param_as_mm(param):
    """Lee un Double como mm; no exige SpecTypeId.Length (falla en algunos entornos)."""
    if param is None:
        return None
    try:
        if param.StorageType != StorageType.Double:
            return None
    except Exception:
        return None
    try:
        raw = float(param.AsDouble())
        return UnitUtils.ConvertFromInternalUnits(raw, UnitTypeId.Millimeters)
    except Exception:
        return None


def _param_definition_name(param):
    if param is None:
        return None
    try:
        return _as_unicode(param.Definition.Name).strip()
    except Exception:
        return None


def _iter_bar_parameters(bar):
    """Parámetros de instancia (ordenados si la API lo permite)."""
    if bar is None:
        return
    try:
        if hasattr(bar, "GetOrderedParameters"):
            coll = bar.GetOrderedParameters()
            if coll is not None:
                for p in coll:
                    yield p
                return
    except Exception:
        pass
    try:
        for p in bar.Parameters:
            yield p
    except Exception:
        pass


_SEGMENT_NAME_SET = set(SEGMENT_PARAM_NAMES)


def _bar_length_mm(bar):
    """Longitud de barra (mm) si está disponible — útil en ``RebarInSystem``."""
    try:
        p = bar.get_Parameter(BuiltInParameter.REBAR_BAR_LENGTH)
        mm = _param_as_mm(p)
        if mm is not None:
            return mm
    except Exception:
        pass
    for name in (u"Bar Length", u"Longitud de barra", u"Longitud"):
        try:
            mm = _param_as_mm(bar.LookupParameter(name))
            if mm is not None:
                return mm
        except Exception:
            continue
    return None


def _segments_from_shape_definition(doc, bar):
    """Lee tramos vía ``RebarShapeDefinition.GetParameters`` (Ids de shared params)."""
    parts = {}
    if doc is None or bar is None:
        return parts
    try:
        sid = None
        try:
            sid = bar.GetShapeId()
        except Exception:
            try:
                sid = bar.GetShapeId(0)
            except Exception:
                sid = getattr(bar, "RebarShapeId", None)
        if sid is None or _element_id_int(sid) < 0:
            return parts
        shape = doc.GetElement(sid)
        if shape is None:
            return parts
        defn = shape.GetRebarShapeDefinition()
        if defn is None:
            return parts
        for pid in defn.GetParameters():
            try:
                p = bar.get_Parameter(pid)
            except Exception:
                p = None
            if p is None:
                try:
                    # Algunas versiones exponen el shared param solo por nombre.
                    pe = doc.GetElement(pid)
                    pname = pe.Name if pe is not None else None
                    if pname:
                        p = bar.LookupParameter(pname)
                except Exception:
                    p = None
            name = _param_definition_name(p)
            if not name:
                continue
            key = name.upper() if len(name) == 1 else name
            if key not in _SEGMENT_NAME_SET and name not in _SEGMENT_NAME_SET:
                continue
            use_name = key if key in _SEGMENT_NAME_SET else name
            mm = _param_as_mm(p)
            if mm is None:
                continue
            parts[use_name] = _roundup_10mm(mm)
    except Exception:
        pass
    return parts


def _segment_fingerprint(bar, doc=None):
    """
    Tupla ordenada ``(nombre, mm_redondeado)`` de segmentos A…R (+ longitud).

    Estrategia (misma idea que ``enfierrado_shaft_hashtag``):
    1) Parámetros de la definición de forma
    2) Recorrer parámetros de instancia por nombre de letra
    3) LookupParameter por nombre
    4) Longitud total de barra (mallas / respaldo)
    """
    parts = {}

    # 1) Shape definition
    if doc is not None:
        parts.update(_segments_from_shape_definition(doc, bar))

    # 2) Iterar parámetros de instancia (con y sin filtro Length)
    for require_length in (True, False):
        found_any = False
        for param in _iter_bar_parameters(bar):
            if param is None:
                continue
            try:
                if param.StorageType != StorageType.Double:
                    continue
            except Exception:
                continue
            if require_length and not _is_length_param(param):
                continue
            name = _param_definition_name(param)
            if not name:
                continue
            key = name.upper() if len(name) == 1 else name
            if key not in _SEGMENT_NAME_SET and name not in _SEGMENT_NAME_SET:
                continue
            use_name = key if key in _SEGMENT_NAME_SET else name
            mm = _param_as_mm(param)
            if mm is None:
                continue
            parts[use_name] = _roundup_10mm(mm)
            found_any = True
        if found_any and require_length:
            break

    # 3) Lookup directo por letra (por si el iterador no los lista)
    for name in SEGMENT_PARAM_NAMES:
        if name in parts:
            continue
        try:
            p = bar.LookupParameter(name)
        except Exception:
            p = None
        if p is None:
            try:
                plist = bar.GetParameters(name)
                if plist is not None and int(plist.Count) > 0:
                    p = plist[0]
            except Exception:
                p = None
        mm = _param_as_mm(p)
        if mm is None:
            continue
        parts[name] = _roundup_10mm(mm)

    # 4) Longitud total
    bar_len = _bar_length_mm(bar)
    if bar_len is not None:
        parts[u"__L__"] = _roundup_10mm(bar_len)

    ordered = []
    for name in SEGMENT_PARAM_NAMES:
        if name in parts:
            ordered.append((name, parts[name]))
    if u"__L__" in parts:
        ordered.append((u"__L__", parts[u"__L__"]))
    return tuple(ordered)


def _is_bar_element(el):
    return isinstance(el, (Rebar, RebarInSystem))


def _owned_by_area_reinforcement(doc, ris):
    """True si el ``RebarInSystem`` pertenece a un ``AreaReinforcement``."""
    if not isinstance(ris, RebarInSystem):
        return False
    try:
        sys_id = ris.SystemId
    except Exception:
        return False
    if _element_id_int(sys_id) < 0:
        return False
    try:
        owner = doc.GetElement(sys_id)
    except Exception:
        return False
    return isinstance(owner, AreaReinforcement)


def _get_shape_name(doc, rebar):
    """Nombre de ``RebarShape``; cadena vacía si no hay."""
    sid = None
    try:
        sid = rebar.GetShapeId()
    except Exception:
        try:
            sid = rebar.GetShapeId(0)
        except Exception:
            try:
                sid = rebar.RebarShapeId
            except Exception:
                sid = None
    if sid is None or _element_id_int(sid) < 0:
        return u""
    try:
        shape = doc.GetElement(sid) if doc is not None else None
    except Exception:
        shape = None
    if shape is None:
        return u""
    try:
        return _as_unicode(shape.Name).strip()
    except Exception:
        return u""


def _segment_lengths_tuple(segment_fp):
    """
    Solo longitudes de segmentos de forma (A/B/C/…), sin ``__L__``.

    Normaliza barra invertida: min(secuencia, secuencia invertida).
    """
    named_vals = []
    for item in segment_fp or ():
        try:
            name, val = item
        except Exception:
            continue
        if name == u"__L__":
            continue
        named_vals.append(int(val))
    vals = tuple(named_vals)
    if not vals:
        return vals
    return min(vals, tuple(reversed(vals)))


def build_fingerprint(doc, rebar):
    """
    Fingerprint = diámetro + nombre de shape + longitudes de segmentos.

    Returns:
        (fingerprint_tuple, diameter_mm) o (None, None) si no se puede clasificar.
    """
    if rebar is None or not _is_bar_element(rebar):
        return None, None

    type_id = rebar.GetTypeId()
    bar_type = doc.GetElement(type_id) if type_id else None
    diam = _bar_diameter_mm(bar_type)
    if diam is None or diam <= 0:
        return None, None

    shape_name = _get_shape_name(doc, rebar)
    segment_fp = _segment_fingerprint(rebar, doc)
    segs = _segment_lengths_tuple(segment_fp)

    fp = (
        int(diam),
        shape_name,
        segs,
    )
    return fp, diam


def format_fingerprint_debug(fp):
    """Texto corto del fingerprint para diagnósticos / resumen."""
    if not fp:
        return u"(vacío)"
    try:
        diam, shape_name, segs = fp
    except Exception:
        return _as_unicode(fp)
    seg_txt = u"+".join(_as_unicode(v) for v in (segs or ()))
    if not seg_txt:
        seg_txt = u"—"
    shape_txt = shape_name if shape_name else u"—"
    return u"ø{} shape={} segs=({})".format(diam, shape_txt, seg_txt)


def collect_rebars(doc):
    """
    Barras del documento: ``Rebar`` libres + ``RebarInSystem`` de Area Reinforcement.

    No incluye ``RebarInSystem`` de Path Reinforcement ni el elemento
    ``AreaReinforcement`` en sí.
    """
    result = []
    seen = set()

    for el in FilteredElementCollector(doc).OfClass(Rebar):
        if el is None or not isinstance(el, Rebar):
            continue
        eid = _element_id_int(el.Id)
        if eid < 0 or eid in seen:
            continue
        seen.add(eid)
        result.append(el)

    for el in FilteredElementCollector(doc).OfClass(RebarInSystem):
        if el is None or not isinstance(el, RebarInSystem):
            continue
        if not _owned_by_area_reinforcement(doc, el):
            continue
        eid = _element_id_int(el.Id)
        if eid < 0 or eid in seen:
            continue
        seen.add(eid)
        result.append(el)

    return result


def build_numbering_plan(doc, progress=None, rebars=None):
    """
    Calcula grupos y marcas sin escribir (reevaluación completa).

    Toma todas las barras, agrupa por fingerprint y asigna índices desde 01
    por diámetro, sin conservar marcas previas.

    Args:
        doc: Document
        progress: ``NumerarRebarProgress`` opcional (un ``step`` por barra).
        rebars: lista opcional ya recolectada (evita doble ``collect_rebars``).

    Returns:
        dict con keys:
          assignments: list of (rebar, mark, diam, index)
          groups_by_diam: {diam: [(fp, [rebar,...], mark, index), ...]}
          skipped: list of (rebar, reason)
          total_rebars, total_rebar, total_rebar_in_system,
          total_groups, total_assigned,
          bars_already_ok, bars_to_update
    """
    if rebars is None:
        rebars = collect_rebars(doc)
    n_rebar = sum(1 for b in rebars if isinstance(b, Rebar))
    n_ris = sum(1 for b in rebars if isinstance(b, RebarInSystem))
    skipped = []
    # diam -> fp -> list[rebar]
    buckets = {}

    for rb in rebars:
        if progress is not None:
            try:
                progress.step(u"Analizando barras…")
            except Exception:
                pass
        fp, diam = build_fingerprint(doc, rb)
        if fp is None:
            skipped.append((rb, u"Sin diámetro / tipo de barra válido"))
            continue
        by_fp = buckets.setdefault(diam, {})
        by_fp.setdefault(fp, []).append(rb)

    assignments = []
    groups_by_diam = {}
    bars_already_ok = 0
    bars_to_update = 0

    for diam in sorted(buckets.keys()):
        fp_map = buckets[diam]
        group_items = []
        for fp, members in fp_map.items():
            min_id = min(_element_id_int(m.Id) for m in members)
            group_items.append((min_id, fp, members))
        group_items.sort(key=lambda t: t[0])

        assigned = _assign_indices_full(group_items, diam)
        groups_by_diam[diam] = []
        for fp, members, mark, index in assigned:
            groups_by_diam[diam].append((fp, members, mark, index))
            for rb in members:
                assignments.append((rb, mark, diam, index))
                current = _get_armadura_marca(rb)
                if current == mark:
                    bars_already_ok += 1
                else:
                    bars_to_update += 1

    return {
        u"assignments": assignments,
        u"groups_by_diam": groups_by_diam,
        u"skipped": skipped,
        u"total_rebars": len(rebars),
        u"total_rebar": n_rebar,
        u"total_rebar_in_system": n_ris,
        u"total_groups": sum(len(v) for v in groups_by_diam.values()),
        u"total_assigned": len(assignments),
        u"bars_already_ok": bars_already_ok,
        u"bars_to_update": bars_to_update,
    }


def _set_armadura_marca(bar, text):
    """Escribe la marca en ``Armadura_Marca`` (sobrescribe)."""
    if bar is None:
        return False
    try:
        p = bar.LookupParameter(ARMADURA_MARCA_PARAM)
    except Exception:
        p = None
    if p is None or p.IsReadOnly:
        return False
    valor = _as_unicode(text)
    try:
        st = p.StorageType
        if st == StorageType.String:
            p.Set(valor)
            return True
    except Exception:
        pass
    try:
        p.SetValueString(valor)
        return True
    except Exception:
        pass
    try:
        p.Set(valor)
        return True
    except Exception:
        return False


def apply_numbering(doc, plan, progress=None):
    """
    Escribe ``Armadura_Marca`` en transacción.

    Sobrescribe siempre el valor previo (haya o no marca anterior).
    Returns (ok_count, fail_count, skipped_unchanged).
    ``skipped_unchanged`` queda en 0 (se mantiene por compatibilidad de firma).
    """
    assignments = plan.get(u"assignments") or []
    if not assignments:
        return 0, 0, 0

    ok = 0
    fail = 0
    t = Transaction(doc, TRANSACTION_NAME)
    t.Start()
    try:
        for rb, mark, _diam, _idx in assignments:
            if progress is not None:
                try:
                    progress.step(
                        u"Escribiendo {}…".format(ARMADURA_MARCA_PARAM)
                    )
                except Exception:
                    pass
            try:
                if _set_armadura_marca(rb, mark):
                    ok += 1
                else:
                    fail += 1
            except Exception:
                fail += 1
        t.Commit()
    except Exception:
        try:
            t.RollBack()
        except Exception:
            pass
        raise
    return ok, fail, 0


def _summary_text(plan):
    lines = []
    lines.append(
        u"Barras en documento: {}  (Rebar: {}  |  RebarInSystem área: {})".format(
            plan.get(u"total_rebars", 0),
            plan.get(u"total_rebar", 0),
            plan.get(u"total_rebar_in_system", 0),
        )
    )
    lines.append(
        u"Asignables: {}  |  Grupos: {}  |  Omitidas: {}".format(
            plan.get(u"total_assigned", 0),
            plan.get(u"total_groups", 0),
            len(plan.get(u"skipped") or []),
        )
    )
    lines.append(
        u"Reevaluación completa: barras ya OK {}  |  a actualizar {}".format(
            plan.get(u"bars_already_ok", 0),
            plan.get(u"bars_to_update", 0),
        )
    )
    lines.append(u"")
    lines.append(u"Grupos por diámetro (marca → cantidad de sets):")
    groups_by_diam = plan.get(u"groups_by_diam") or {}
    for diam in sorted(groups_by_diam.keys()):
        groups = groups_by_diam[diam]
        parts = []
        for _fp, members, mark, _idx in groups:
            parts.append(u"{}×{}".format(mark, len(members)))
        preview = u", ".join(parts[:12])
        if len(parts) > 12:
            preview += u", … (+{})".format(len(parts) - 12)
        lines.append(u"  ø{} mm: {}".format(diam, preview))
    lines.append(u"")
    lines.append(
        u"Matching = ø + nombre RebarShape + segmentos (10 mm). "
        u"Sin ganchos, sin capa, sin L total."
    )
    lines.append(u"Ejemplos de fingerprint (hasta 8 grupos):")
    shown = 0
    for diam in sorted(groups_by_diam.keys()):
        for fp, members, mark, _idx in groups_by_diam[diam]:
            if shown >= 8:
                break
            lines.append(
                u"  {} ({} sets): {}".format(
                    mark, len(members), format_fingerprint_debug(fp)
                )
            )
            shown += 1
        if shown >= 8:
            break
    lines.append(u"")
    lines.append(
        u"Cada corrida reevalúa todas las barras, sobrescribe {} "
        u"y reasigna desde cero por diámetro.".format(ARMADURA_MARCA_PARAM)
    )
    return u"\n".join(lines)


def _show_ok_cancel(uiapp, instruction, content):
    try:
        from bimtools_instruction_dialog import show_ok_cancel_dialog
        from revit_wpf_window_position import revit_main_hwnd

        hwnd = revit_main_hwnd(uiapp) if uiapp is not None else None
        return show_ok_cancel_dialog(
            TOOL_TITLE,
            instruction=instruction,
            content=content,
            ok_text=u"Numerar",
            cancel_text=u"Cancelar",
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
    except Exception:
        pass
    try:
        from Autodesk.Revit.UI import TaskDialogCommonButtons, TaskDialogResult

        r = TaskDialog.Show(
            TOOL_TITLE,
            instruction + u"\n\n" + content,
            TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel,
        )
        return r == TaskDialogResult.Ok
    except Exception:
        return True


def _show_message(uiapp, instruction, content=u""):
    try:
        from bimtools_instruction_dialog import show_message_dialog
        from revit_wpf_window_position import revit_main_hwnd

        hwnd = revit_main_hwnd(uiapp) if uiapp is not None else None
        show_message_dialog(
            TOOL_TITLE,
            instruction=instruction,
            content=content,
            ok_text=u"Entendido",
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass
    try:
        msg = instruction
        if content:
            msg = instruction + u"\n\n" + content
        TaskDialog.Show(TOOL_TITLE, msg)
    except Exception:
        print(instruction)
        if content:
            print(content)


def run(revit_app):
    """Punto de entrada pyRevit: ``run(__revit__)``."""
    uidoc = revit_app.ActiveUIDocument
    if uidoc is None:
        _show_message(revit_app, u"No hay documento activo.")
        return
    doc = uidoc.Document
    if doc is None:
        _show_message(revit_app, u"No hay documento activo.")
        return

    rebars = collect_rebars(doc)
    n_bars = len(rebars)
    if n_bars > 0:
        with NumerarRebarProgress(
            n_bars,
            title_prefix=u"Arainco: Analizando marcas Rebar",
        ) as pb:
            plan = build_numbering_plan(doc, progress=pb, rebars=rebars)
    else:
        plan = build_numbering_plan(doc, rebars=rebars)

    if plan[u"total_assigned"] <= 0:
        _show_message(
            revit_app,
            u"No se encontraron barras numerables en el documento.",
            content=_summary_text(plan),
        )
        return

    accepted = _show_ok_cancel(
        revit_app,
        u"¿Reasignar marcas a todas las barras? "
        u"(sobrescribe Armadura_Marca existente; reevaluación completa)",
        _summary_text(plan),
    )
    if not accepted:
        return

    assignments = plan.get(u"assignments") or []
    try:
        if assignments:
            with NumerarRebarProgress(
                len(assignments),
                title_prefix=u"Arainco: Escribiendo Armadura_Marca",
            ) as pb:
                ok, fail, _unchanged = apply_numbering(doc, plan, progress=pb)
        else:
            ok, fail, _unchanged = apply_numbering(doc, plan)
    except Exception as ex:
        _show_message(
            revit_app,
            u"Error al escribir {}.".format(ARMADURA_MARCA_PARAM),
            content=_as_unicode(ex),
        )
        return

    result_lines = [
        u"Marcas escritas (sobrescritura): {}".format(ok),
        u"Fallos: {}".format(fail),
        u"Grupos reevaluados: {}".format(plan.get(u"total_groups", 0)),
        u"Omitidas (sin clasificar): {}".format(len(plan[u"skipped"])),
    ]
    _show_message(
        revit_app,
        u"Reevaluación completa de marcas Rebar finalizada.",
        content=u"\n".join(result_lines),
    )
