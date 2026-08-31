# -*- coding: utf-8 -*-
"""
Arainco: Sección Detail — visibilidad del símbolo por vista.

1. Selecciona el símbolo de sección (marcador/viewer) de una vista Detail.
2. Obtiene Building Sections con el mismo «Section Filter» que el marcador.
3. Hide in view del mismo marcador (Id de la ViewSection) en cada candidata,
   excepto en la vista activa. No abre un collector por vista.

Revit 2024+ | pyRevit (importable via run).
"""

from __future__ import division, print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BuiltInParameter,
    CheckoutStatus,
    ElementId,
    FailureProcessingResult,
    FailureSeverity,
    FilteredElementCollector,
    IFailuresPreprocessor,
    StorageType,
    SubTransaction,
    Transaction,
    View,
    ViewFamily,
    ViewFamilyType,
    ViewSection,
    ViewType,
    WorksharingUtils,
)
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

# Cachés de una ejecución (se limpian al empezar a ocultar).
_VFT_IS_BUILDING_SECTION = {}
_CANDIDATES_BY_SF = {}
_ES_BUILDING_SECTION_FN = []
_HIDE_FAILURES_PREPROCESSOR = None
_CAPTURED_REVIT_FAILURES = []

_PERM_FAILURE_MARKERS = (
    u"permission",
    u"permiso",
    u"owned by",
    u"ocupado",
    u"checked out",
    u"borrowed",
    u"cannot be edited",
    u"can't be edited",
    u"can't edit",
    u"cannot edit",
    u"no se puede editar",
    u"cannot be hidden",
    u"can't be hidden",
    u"cannot hide",
    u"can't hide",
    u"no se puede ocultar",
    u"not permitted",
    u"no permitido",
    u"workset",
    u"subproyecto",
    u"relinquish",
    u"another user",
    u"otro usuario",
    u"last edited by",
    u"editado por",
    u"no permission",
    u"sin permiso",
    u"not editable",
    u"no editable",
    u"element is not editable",
)

_TOOL_TITLE = u"Arainco: Sección Detail solo vista"
_TXN_HIDE = u"Arainco: Sección Detail solo vista"
_PROGRESS_ACCENT_RGB = (91, 192, 222)
_PROMPT = (
    u"Selecciona uno o varios símbolos de sección Detail "
    u"(cabeza / línea de corte). Finish para continuar / Esc cancela."
)


def _pbar_enabled():
    try:
        from pyrevit import forms as _forms  # noqa: F401
        return True
    except Exception:
        return False


class _ToolProgress(object):
    """``forms.ProgressBar`` de pyRevit; no-op si no está disponible."""

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

                r, g, b = _PROGRESS_ACCENT_RGB
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
        if self._pb is None:
            return
        i = int(self._index)
        if i >= self._total:
            i = self._total - 1
        self._index = i + 1
        label = _as_unicode(phase_label).strip() if phase_label else u""
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


def _short_view_label(view):
    try:
        name = _as_unicode(view.Name).strip()
    except Exception:
        name = u""
    if not name:
        return u"Vista"
    if len(name) > 48:
        return name[:47] + u"…"
    return name


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _reset_run_caches():
    _VFT_IS_BUILDING_SECTION.clear()
    _CANDIDATES_BY_SF.clear()
    del _CAPTURED_REVIT_FAILURES[:]


def _eid_int(eid):
    try:
        return int(eid.IntegerValue)
    except Exception:
        try:
            return int(eid.Value)
        except Exception:
            return -1


def _mostrar_aviso(uiapp, instruction, content=u"", ok_text=u"Entendido"):
    try:
        from bimtools_instruction_dialog import show_message_dialog
        from revit_wpf_window_position import revit_main_hwnd

        hwnd = revit_main_hwnd(uiapp) if uiapp is not None else None
        show_message_dialog(
            _TOOL_TITLE,
            instruction=_as_unicode(instruction),
            content=_as_unicode(content) if content else None,
            ok_text=_as_unicode(ok_text),
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass
    try:
        body = _as_unicode(instruction)
        extra = _as_unicode(content).strip()
        if extra:
            body = body + u"\n\n" + extra
        TaskDialog.Show(_TOOL_TITLE, body)
    except Exception:
        pass


def _es_vista_detail(doc, view):
    """True si ``view`` es una ViewSection de familia Detail (no plantilla)."""
    if view is None or doc is None:
        return False
    try:
        if not isinstance(view, ViewSection):
            return False
    except Exception:
        return False
    try:
        if getattr(view, "IsTemplate", False):
            return False
    except Exception:
        pass
    try:
        if view.ViewType == ViewType.Detail:
            return True
    except Exception:
        pass
    try:
        vft = doc.GetElement(view.GetTypeId())
        if vft is not None and isinstance(vft, ViewFamilyType):
            return vft.ViewFamily == ViewFamily.Detail
    except Exception:
        pass
    return False


def _view_from_marker_element(doc, element):
    """
    Obtiene la View asociada al símbolo/marcador seleccionado en una vista.

    El clic sobre la cabeza o la línea de corte suele devolver la propia
    ``ViewSection``. También se contempla viewer → ID_PARAM.
    """
    if doc is None or element is None:
        return None

    if isinstance(element, ViewSection):
        return element

    if isinstance(element, View):
        return element

    try:
        param = element.get_Parameter(BuiltInParameter.ID_PARAM)
    except Exception:
        param = None
    if param is not None:
        try:
            ref_id = param.AsElementId()
        except Exception:
            ref_id = None
        if ref_id is not None and _eid_int(ref_id) >= 0:
            try:
                ref_el = doc.GetElement(ref_id)
            except Exception:
                ref_el = None
            if isinstance(ref_el, View):
                return ref_el

    return None


def _resolve_detail_section(doc, element):
    """
    A partir del elemento clicado (símbolo o ViewSection), devuelve la
    ViewSection Detail asociada, o None.
    """
    if doc is None or element is None:
        return None

    if _es_vista_detail(doc, element):
        return element

    view = _view_from_marker_element(doc, element)
    if view is not None and _es_vista_detail(doc, view):
        return view

    return None


class _DetailSectionSymbolFilter(ISelectionFilter):
    """Permite el símbolo de sección (cabeza/línea) que representa un Detail."""

    def __init__(self, doc):
        self._doc = doc

    def AllowElement(self, element):
        return _resolve_detail_section(self._doc, element) is not None

    def AllowReference(self, reference, point):
        return False


def _collect_detail_sections_from_ids(doc, eids):
    sections = []
    seen = set()
    if eids is None:
        return sections
    for eid in eids:
        try:
            el = doc.GetElement(eid)
        except Exception:
            el = None
        section = _resolve_detail_section(doc, el)
        if section is None:
            continue
        iid = _eid_int(section.Id)
        if iid in seen:
            continue
        seen.add(iid)
        sections.append(section)
    return sections


def _get_preselected_detail_sections(uidoc, doc):
    try:
        eids = uidoc.Selection.GetElementIds()
    except Exception:
        return []
    return _collect_detail_sections_from_ids(doc, eids)


def _pick_detail_section_symbols(uidoc, doc):
    refs = list(
        uidoc.Selection.PickObjects(
            ObjectType.Element,
            _DetailSectionSymbolFilter(doc),
            _PROMPT,
        )
    )
    sections = []
    seen = set()
    for pref in refs:
        try:
            el = doc.GetElement(pref.ElementId)
        except Exception:
            el = None
        section = _resolve_detail_section(doc, el)
        if section is None:
            continue
        iid = _eid_int(section.Id)
        if iid in seen:
            continue
        seen.add(iid)
        sections.append(section)
    return sections


def _view_display_name(view):
    name = u""
    try:
        name = _as_unicode(view.Name).strip()
    except Exception:
        name = u""
    if not name:
        name = u"(sin nombre)"
    try:
        vt = _as_unicode(view.ViewType)
    except Exception:
        vt = u""
    if vt:
        return u"{0}  [{1}]".format(name, vt)
    return name


def _es_building_section_view(view):
    """True si la vista es Building Section (no plantilla / no Detail)."""
    if view is None:
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass
    try:
        vt = view.ViewType
        if vt == ViewType.Detail:
            return False
        if vt != ViewType.Section:
            return False
    except Exception:
        return False

    tid = -1
    try:
        tid = _eid_int(view.GetTypeId())
    except Exception:
        tid = -1
    if tid >= 0 and tid in _VFT_IS_BUILDING_SECTION:
        return _VFT_IS_BUILDING_SECTION[tid]

    ok = _es_building_section_uncached(view)
    if tid >= 0:
        _VFT_IS_BUILDING_SECTION[tid] = ok
    return ok


def _es_building_section_uncached(view):
    if not _ES_BUILDING_SECTION_FN:
        fn = None
        try:
            from filtro_armadura_eje import es_vista_building_section as fn
        except Exception:
            fn = None
        _ES_BUILDING_SECTION_FN.append(fn)
    fn = _ES_BUILDING_SECTION_FN[0] if _ES_BUILDING_SECTION_FN else None
    if fn is not None:
        try:
            return bool(fn(view))
        except Exception:
            pass
    try:
        vft = view.Document.GetElement(view.GetTypeId()) if view.Document else None
        if vft is not None and isinstance(vft, ViewFamilyType):
            if vft.ViewFamily == ViewFamily.Detail:
                return False
            return vft.ViewFamily == ViewFamily.Section
    except Exception:
        pass
    return False


def _leer_section_filter(doc, view):
    """Texto de «Section Filter» en la vista, o None. Sin imports en el bucle."""
    if view is None:
        return None
    try:
        param = view.LookupParameter(u"Section Filter")
    except Exception:
        param = None
    if param is None:
        return None
    try:
        if hasattr(param, "HasValue") and not param.HasValue:
            return None
    except Exception:
        pass
    try:
        storage = param.StorageType
    except Exception:
        storage = None
    if storage == StorageType.String:
        try:
            s = param.AsString()
            if s and _as_unicode(s).strip():
                return _as_unicode(s).strip()
        except Exception:
            pass
    elif storage == StorageType.ElementId:
        try:
            eid = param.AsElementId()
        except Exception:
            eid = None
        if eid is not None and _eid_int(eid) >= 0 and doc is not None:
            try:
                el = doc.GetElement(eid)
            except Exception:
                el = None
            if el is not None:
                try:
                    n = _as_unicode(getattr(el, "Name", None)).strip()
                    if n:
                        return n
                except Exception:
                    pass
    try:
        vs = param.AsValueString()
        if vs and _as_unicode(vs).strip():
            return _as_unicode(vs).strip()
    except Exception:
        pass
    try:
        s = param.AsString()
        if s and _as_unicode(s).strip():
            return _as_unicode(s).strip()
    except Exception:
        pass
    return None


def _section_filters_iguales(a, b):
    if not a or not b:
        return False
    try:
        return _as_unicode(a).strip().lower() == _as_unicode(b).strip().lower()
    except Exception:
        return False


def _building_sections_mismo_filter(doc, section):
    """
    Building Sections con el mismo «Section Filter» que el marcador.

    Un collector de ViewSection (documento, no por vista). Caché por filtro.
    """
    found = []
    sf_key = _leer_section_filter(doc, section)
    if not sf_key:
        return found, None

    cache_key = _as_unicode(sf_key).strip().lower()
    section_iid = _eid_int(section.Id) if section is not None else -1
    cached = _CANDIDATES_BY_SF.get(cache_key)
    if cached is not None:
        if section_iid < 0:
            return list(cached), sf_key
        return [v for v in cached if _eid_int(v.Id) != section_iid], sf_key

    try:
        views = list(FilteredElementCollector(doc).OfClass(ViewSection))
    except Exception:
        return found, sf_key

    all_for_sf = []
    seen = set()
    n_views = len(views)
    with _ToolProgress(n_views, u"Arainco: Buscando candidatas") as pb:
        for view in views:
            pb.step(_short_view_label(view))
            try:
                if view.IsTemplate:
                    continue
            except Exception:
                pass
            try:
                if view.ViewType != ViewType.Section:
                    continue
            except Exception:
                continue
            iid = _eid_int(view.Id)
            if iid < 0 or iid in seen:
                continue
            if not _es_building_section_view(view):
                continue
            sf_view = _leer_section_filter(doc, view)
            if not _section_filters_iguales(sf_key, sf_view):
                continue
            seen.add(iid)
            all_for_sf.append(view)

    def _sort_key(v):
        try:
            return _as_unicode(v.Name).lower()
        except Exception:
            return u""

    all_for_sf.sort(key=_sort_key)
    _CANDIDATES_BY_SF[cache_key] = all_for_sf
    if section_iid < 0:
        return list(all_for_sf), sf_key
    return [v for v in all_for_sf if _eid_int(v.Id) != section_iid], sf_key


def _marker_view_name(marker):
    """Nombre de la vista a la que apunta el marcador (``marker.Name``)."""
    if marker is None:
        return u""
    try:
        return _as_unicode(marker.Name).strip()
    except Exception:
        return u""


def _marker_type_info(doc, marker):
    """(familia, tipo) del ViewFamilyType del marcador."""
    tipo_nombre = u"Sin Tipo"
    familia_nombre = u"Desconocida"
    if doc is None or marker is None:
        return familia_nombre, tipo_nombre
    try:
        type_id = marker.GetTypeId()
    except Exception:
        type_id = None
    if type_id is None or _eid_int(type_id) < 0:
        return familia_nombre, tipo_nombre
    try:
        marker_type = doc.GetElement(type_id)
    except Exception:
        marker_type = None
    if marker_type is None:
        return familia_nombre, tipo_nombre
    try:
        raw = _as_unicode(getattr(marker_type, "Name", None)).strip()
        if raw:
            tipo_nombre = raw
    except Exception:
        pass
    try:
        if hasattr(marker_type, "FamilyName"):
            raw_f = _as_unicode(marker_type.FamilyName).strip()
            if raw_f:
                familia_nombre = raw_f
    except Exception:
        pass
    return familia_nombre, tipo_nombre


def _es_misma_vista(view_a, view_b):
    if view_a is None or view_b is None:
        return False
    try:
        return view_a.Id == view_b.Id
    except Exception:
        return False


def _hit_from_marker(doc, marker):
    familia, tipo = _marker_type_info(doc, marker)
    return {
        u"id": _eid_int(marker.Id),
        u"name": _marker_view_name(marker) or u"(sin nombre)",
        u"familia": familia,
        u"tipo": tipo,
    }


def _failure_description(failure):
    if failure is None:
        return u""
    for getter in (u"GetDescriptionText", u"GetDefaultResolutionCaption"):
        try:
            fn = getattr(failure, getter, None)
            if fn is None:
                continue
            text = _as_unicode(fn()).strip()
            if text:
                return text
        except Exception:
            pass
    return u""


def _clear_captured_revit_failures():
    del _CAPTURED_REVIT_FAILURES[:]


def _take_captured_revit_failures():
    out = []
    seen = set()
    for text in _CAPTURED_REVIT_FAILURES:
        t = _as_unicode(text).strip()
        if not t:
            continue
        key = t.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(t)
    del _CAPTURED_REVIT_FAILURES[:]
    return out


def _record_revit_failure(text):
    t = _as_unicode(text).strip()
    if t:
        _CAPTURED_REVIT_FAILURES.append(t)


def _join_reasons(*parts):
    out = []
    seen = set()
    for part in parts:
        if part is None:
            continue
        if isinstance(part, (list, tuple)):
            items = part
        else:
            items = [part]
        for raw in items:
            t = _as_unicode(raw).strip()
            if not t:
                continue
            key = t.lower()
            if key in seen:
                continue
            seen.add(key)
            out.append(t)
    return u"; ".join(out)


def _failure_looks_permission_or_hide(failure):
    text = _failure_description(failure).lower()
    if not text:
        return False
    for marker in _PERM_FAILURE_MARKERS:
        if marker in text:
            return True
    return False


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


class _HidePermissionFailuresPreprocessor(IFailuresPreprocessor):
    """
    Intercepta avisos/errores de permiso u ocultar: guarda el texto original
    de Revit y los quita de la cola para que no salga el diálogo modal ni se
    aborte el lote. El usuario los ve en el informe de la herramienta.
    """

    def PreprocessFailures(self, failures_accessor):
        if failures_accessor is None:
            return FailureProcessingResult.Continue
        msgs = list(_iter_failure_msgs(failures_accessor))
        handled = False
        for f in msgs:
            try:
                sev = f.GetSeverity()
            except Exception:
                continue
            is_perm = _failure_looks_permission_or_hide(f)
            if sev == FailureSeverity.Warning:
                if is_perm:
                    _record_revit_failure(_failure_description(f))
                try:
                    if is_perm:
                        failures_accessor.DeleteWarning(f)
                        handled = True
                except Exception:
                    pass
                continue
            if sev != FailureSeverity.Error:
                continue
            if not is_perm:
                continue
            _record_revit_failure(_failure_description(f))
            try:
                failures_accessor.DeleteError(f)
                handled = True
            except Exception:
                pass
        if handled:
            return FailureProcessingResult.ProceedWithCommit
        return FailureProcessingResult.Continue


def _attach_hide_failures_preprocessor(txn):
    global _HIDE_FAILURES_PREPROCESSOR
    if txn is None:
        return
    try:
        if _HIDE_FAILURES_PREPROCESSOR is None:
            _HIDE_FAILURES_PREPROCESSOR = _HidePermissionFailuresPreprocessor()
        opts = txn.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(_HIDE_FAILURES_PREPROCESSOR)
        try:
            opts.SetClearAfterRollback(True)
        except Exception:
            pass
        try:
            opts.SetForcedModalHandling(False)
        except Exception:
            pass
        txn.SetFailureHandlingOptions(opts)
    except Exception:
        pass


def _checkout_owned_by_other(doc, eid):
    if doc is None or eid is None:
        return False, u""
    try:
        if not getattr(doc, "IsWorkshared", False):
            return False, u""
    except Exception:
        return False, u""
    try:
        status = WorksharingUtils.GetCheckoutStatus(doc, eid)
    except Exception:
        return False, u""
    owned = False
    try:
        owned = status == CheckoutStatus.OwnedByOtherUser
    except Exception:
        try:
            owned = _as_unicode(status).lower().find(u"other") >= 0
        except Exception:
            owned = False
    if not owned:
        return False, u""
    owner = u""
    try:
        info = WorksharingUtils.GetWorksharingTooltipInfo(doc, eid)
        owner = _as_unicode(getattr(info, "Owner", None)).strip()
    except Exception:
        owner = u""
    if owner:
        return True, u"ocupado por {0}".format(owner)
    return True, u"ocupado por otro usuario"


def _vista_editable_para_ocultar(doc, view):
    """(ok, motivo). Hide in view modifica la Building Section candidata."""
    if view is None:
        return False, u"sin vista"
    owned, why = _checkout_owned_by_other(doc, view.Id)
    if owned:
        return False, why or u"vista sin permiso de edición"
    return True, u""


def _marcador_se_puede_ocultar(view, marker):
    if view is None or marker is None:
        return False, u"sin marcador"
    try:
        if hasattr(marker, "CanBeHidden") and not marker.CanBeHidden(view):
            return False, u"Revit no permite ocultarlo en esta vista"
    except Exception:
        pass
    return True, u""


def _try_hide_ids_in_view(doc, view, ids):
    """
    SubTransaction: HideElements.

    Returns:
        (ok_api, textos_revit)
        ok_api True si la llamada no lanzó. El caller confirma con IsHidden.
    """
    if view is None or ids is None:
        return False, []
    try:
        if ids.Count < 1:
            return False, []
    except Exception:
        return False, []
    _clear_captured_revit_failures()
    st = SubTransaction(doc)
    try:
        st.Start()
    except Exception:
        return False, _take_captured_revit_failures()
    try:
        view.HideElements(ids)
        captured = _take_captured_revit_failures()
        st.Commit()
        return True, captured
    except Exception as ex:
        captured = _take_captured_revit_failures()
        try:
            st.RollBack()
        except Exception:
            pass
        if not captured:
            captured = [_as_unicode(ex)]
        return False, captured


def _marcadores_a_ocultar(section):
    """
    El marcador de sección es la propia ViewSection seleccionada
    (mismo ElementId / nombre en todas las vistas).
    """
    if section is None:
        return []
    return [section]


def _hide_markers_in_view(doc, view, markers):
    """
    Hide in view del Id conocido. Sin collector por vista.

    Permisos / no ocultable: se omiten y no abortan el lote.
    Returns:
        (hidden_list, already_list, skipped_list)
        skipped_list: [(marker, motivo), ...]
    """
    from System.Collections.Generic import List

    already = []
    skipped = []
    pending = []
    if view is None or not markers:
        return [], already, skipped

    ok_view, why_view = _vista_editable_para_ocultar(doc, view)
    if not ok_view:
        for marker in markers:
            if marker is None:
                continue
            skipped.append((marker, why_view))
        return [], already, skipped

    for marker in markers:
        if marker is None:
            continue
        try:
            if marker.IsHidden(view):
                already.append(marker)
                continue
        except Exception:
            pass
        ok_m, why_m = _marcador_se_puede_ocultar(view, marker)
        if not ok_m:
            skipped.append((marker, why_m))
            continue
        pending.append(marker)
    if not pending:
        return [], already, skipped

    ids = List[ElementId]()
    for marker in pending:
        try:
            ids.Add(marker.Id)
        except Exception:
            pass
    if ids.Count == 0:
        return [], already, skipped

    _api_ok, captured = _try_hide_ids_in_view(doc, view, ids)

    hidden = []
    still = []
    for marker in pending:
        try:
            if marker.IsHidden(view):
                hidden.append(marker)
                continue
        except Exception:
            pass
        still.append(marker)

    if still:
        for marker in still:
            one = List[ElementId]()
            try:
                one.Add(marker.Id)
            except Exception:
                skipped.append(
                    (
                        marker,
                        _join_reasons(captured) or u"no se pudo ocultar (permisos)",
                    )
                )
                continue
            _one_ok, one_captured = _try_hide_ids_in_view(doc, view, one)
            try:
                if marker.IsHidden(view):
                    hidden.append(marker)
                    continue
            except Exception:
                pass
            skipped.append(
                (
                    marker,
                    _join_reasons(one_captured, captured)
                    or u"aviso de Revit: no se pudo ocultar (permisos)",
                )
            )
        return hidden, already, skipped

    return hidden, already, skipped


def _ocultar_marcadores_en_candidatas(doc, sections, active_view):
    """
    Candidatas por Section Filter; Hide in view del marcador seleccionado
    (mismo Id) salvo en la vista activa. No recorre OST_Viewers por vista.
    """
    _reset_run_caches()
    reports = []
    jobs = []

    for section in sections:
        nombre = _marker_view_name(section)
        markers = _marcadores_a_ocultar(section)
        candidates, sf_key = _building_sections_mismo_filter(doc, section)
        item = {
            u"section": section,
            u"nombre": nombre,
            u"sf_key": sf_key,
            u"n_candidatas": len(candidates),
            u"hidden": [],
            u"kept_active": [],
            u"already": [],
            u"skipped": [],
            u"errors": [],
        }
        reports.append(item)
        if not sf_key or not nombre or not markers:
            continue
        hits_sel = [_hit_from_marker(doc, m) for m in markers]
        for view in candidates:
            if _es_misma_vista(view, active_view):
                item[u"kept_active"].append((view, hits_sel))
                continue
            jobs.append((view, markers, item))

    if not jobs:
        return reports

    t = Transaction(doc, _TXN_HIDE)
    _attach_hide_failures_preprocessor(t)
    t.Start()
    try:
        with _ToolProgress(len(jobs), u"Arainco: Hide in view") as pb:
            for view, markers, item in jobs:
                pb.step(_short_view_label(view))
                hidden, already, skipped = _hide_markers_in_view(
                    doc, view, markers
                )
                if already:
                    item[u"already"].append(
                        (view, [_hit_from_marker(doc, m) for m in already])
                    )
                if hidden:
                    item[u"hidden"].append(
                        (view, [_hit_from_marker(doc, m) for m in hidden])
                    )
                if skipped:
                    hits_skip = []
                    reasons = []
                    seen_r = set()
                    for marker, reason in skipped:
                        hits_skip.append(_hit_from_marker(doc, marker))
                        r = _as_unicode(reason).strip()
                        key = r.lower()
                        if key and key not in seen_r:
                            seen_r.add(key)
                            reasons.append(r)
                    item[u"skipped"].append(
                        (view, hits_skip, u"; ".join(reasons))
                    )
        t.Commit()
    except Exception:
        try:
            t.RollBack()
        except Exception:
            pass
        raise
    return reports


def _format_marker_hit(hit):
    return (
        u"      ID marcador: {0} | Apunta a: '{1}'\n"
        u"      Familia: {2} | Tipo: {3}"
    ).format(
        hit.get(u"id", -1),
        hit.get(u"name", u""),
        hit.get(u"familia", u""),
        hit.get(u"tipo", u""),
    )


def _append_view_hits(lines, view, hits, suffix=u""):
    lines.append(
        u"  - {0}{1}".format(_view_display_name(view), suffix or u"")
    )
    for hit in hits:
        lines.append(_format_marker_hit(hit))


def _section_label(section):
    name = u""
    try:
        name = _as_unicode(section.Name).strip()
    except Exception:
        name = u""
    if not name:
        name = u"(sin nombre)"
    return u"{0}  (Id {1})".format(name, _eid_int(section.Id))


def _build_report(reports, active_view):
    lines = []
    active_name = u""
    if active_view is not None:
        try:
            active_name = _as_unicode(active_view.Name).strip()
        except Exception:
            active_name = u""
    if active_name:
        lines.append(u"Vista activa: {0}".format(active_name))
        lines.append(u"")

    lines.append(
        u"Símbolos de sección Detail seleccionados: {0}".format(len(reports))
    )
    lines.append(u"")

    for item in reports:
        section = item.get(u"section")
        sf_key = item.get(u"sf_key")
        nombre = item.get(u"nombre") or u""
        n_candidatas = item.get(u"n_candidatas") or 0
        hidden = item.get(u"hidden") or []
        kept = item.get(u"kept_active") or []
        already = item.get(u"already") or []
        skipped = item.get(u"skipped") or []
        errors = item.get(u"errors") or []

        lines.append(u"• {0}".format(_section_label(section)))
        if not sf_key:
            lines.append(
                u"  Sin «Section Filter»: no se buscan Building Sections."
            )
            lines.append(u"")
            continue
        if not nombre:
            lines.append(
                u"  El marcador no tiene nombre: no se puede comparar."
            )
            lines.append(u"")
            continue

        lines.append(u"  Nombre del marcador: {0}".format(nombre))
        lines.append(u"  Section Filter: {0}".format(sf_key))
        lines.append(
            u"  Building Sections candidatas (mismo Section Filter): "
            u"{0}".format(n_candidatas)
        )
        if n_candidatas == 0:
            lines.append(u"  No hay Building Sections con ese Section Filter.")
        elif not hidden and not kept and not already and not skipped and not errors:
            lines.append(
                u"  El marcador no estaba visible (o no se pudo ocultar) "
                u"en esas Building Sections."
            )
        else:
            if hidden:
                lines.append(
                    u"  Hide in view aplicado ({0}):".format(len(hidden))
                )
                for view, hits in hidden:
                    _append_view_hits(lines, view, hits)
            if kept:
                lines.append(
                    u"  Vista activa: el marcador se mantiene visible "
                    u"({0}):".format(len(kept))
                )
                for view, hits in kept:
                    _append_view_hits(lines, view, hits, u"  ← activa")
            if already:
                lines.append(
                    u"  Ya estaba oculto ({0}):".format(len(already))
                )
                for view, hits in already:
                    _append_view_hits(lines, view, hits)
            if skipped:
                lines.append(
                    u"  Avisos de Revit (permisos / no ocultable), se omitió "
                    u"y se continuó ({0}):".format(len(skipped))
                )
                for view, hits, reason in skipped:
                    suffix = u""
                    if reason:
                        suffix = u"  — {0}".format(reason)
                    _append_view_hits(lines, view, hits, suffix)
            if errors:
                lines.append(u"  No se pudo ocultar:")
                for view, msg in errors:
                    lines.append(
                        u"  - {0}: {1}".format(_view_display_name(view), msg)
                    )
        lines.append(u"")

    return u"\n".join(lines).strip()


def run(uiapp):
    """Entrada desde el pushbutton."""
    if uiapp is None:
        return
    try:
        uidoc = uiapp.ActiveUIDocument
    except Exception:
        uidoc = None
    if uidoc is None:
        _mostrar_aviso(uiapp, u"No hay documento activo.")
        return
    doc = uidoc.Document
    if doc is None:
        _mostrar_aviso(uiapp, u"No hay documento activo.")
        return

    try:
        active_view = uidoc.ActiveView
    except Exception:
        active_view = None

    sections = _get_preselected_detail_sections(uidoc, doc)
    if not sections:
        try:
            sections = _pick_detail_section_symbols(uidoc, doc)
        except OperationCanceledException:
            return
        except Exception as ex:
            _mostrar_aviso(
                uiapp,
                u"No se pudo completar la selección.",
                content=_as_unicode(ex),
            )
            return

    if not sections:
        _mostrar_aviso(
            uiapp,
            u"No hay símbolos de sección Detail seleccionados.",
            content=(
                u"Seleccione la cabeza o la línea de corte de una o varias "
                u"secciones de tipo Detail e intente de nuevo."
            ),
        )
        return

    try:
        reports = _ocultar_marcadores_en_candidatas(doc, sections, active_view)
    except Exception as ex:
        _mostrar_aviso(
            uiapp,
            u"No se pudo aplicar Hide in view.",
            content=_as_unicode(ex),
        )
        return

    report = _build_report(reports, active_view)
    n_hidden = 0
    n_skipped = 0
    for item in reports:
        n_hidden += len(item.get(u"hidden") or [])
        n_skipped += len(item.get(u"skipped") or [])
    if n_hidden:
        instruction = (
            u"Marcador oculto (Hide in view) en {0} Building Section(s). "
            u"En la vista activa se mantiene visible."
        ).format(n_hidden)
        if n_skipped:
            instruction = instruction + (
                u" {0} vista(s) con aviso de Revit (permisos); el texto "
                u"original está en el informe."
            ).format(n_skipped)
    elif n_skipped:
        instruction = (
            u"No se ocultó el marcador: las candidatas están ocupadas "
            u"o Revit no permite ocultar el visor. El resto no se detuvo."
        )
    else:
        instruction = (
            u"No se ocultó el marcador en ninguna Building Section "
            u"(sin coincidencia, ya oculto o solo vista activa)."
        )
    _mostrar_aviso(
        uiapp,
        instruction,
        content=report,
        ok_text=u"Entendido",
    )
