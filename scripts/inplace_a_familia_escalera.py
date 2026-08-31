# -*- coding: utf-8 -*-
"""
In-Place → Familia Escalera — convierte un Model In-Place en familia cargable
categoría Escalera (FreeForm), la coloca en la misma posición, valida host de
rebar y elimina el in-place original.

Revit 2025+ | pyRevit | IronPython 2.7 / 3.x / CPython 3
"""

from __future__ import print_function

try:
    unicode
except NameError:
    unicode = str

import math
import os
import re
import tempfile

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    ElementTransformUtils,
    Family,
    FamilyInstance,
    FamilySource,
    FilteredElementCollector,
    FreeFormElement,
    GeometryInstance,
    IFamilyLoadOptions,
    Level,
    Line,
    LocationPoint,
    Options,
    SaveAsOptions,
    Solid,
    SolidUtils,
    Transaction,
    TransactionGroup,
    Transform,
    ViewDetailLevel,
    XYZ,
)
from Autodesk.Revit.DB.Structure import RebarHostData, StructuralType
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler, TaskDialog
from Autodesk.Revit.UI.Events import IdlingEventArgs
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

from System import AppDomain, EventHandler
from System.Windows import RoutedEventHandler, WindowState
from System.Windows.Controls import ComboBoxItem
from System.Windows.Input import Key, KeyEventHandler
from System.Windows.Markup import XamlReader

from bimtools_ui_tokens import BTN_MANUAL
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_wpf_shell import build_simple_tool_xaml
from revit_wpf_window_position import (
    bind_center_wpf_on_revit_monitor,
    position_wpf_window_center_on_monitor,
    revit_main_hwnd,
)

_APPDOMAIN_WINDOW_KEY = u"Arainco_InPlaceAFamiliaEscalera_UI"
_APPDOMAIN_EVENT_KEY = u"Arainco_InPlaceAFamiliaEscalera_ExtEvent"
_APPDOMAIN_HANDLER_KEY = u"Arainco_InPlaceAFamiliaEscalera_Handler"
_APPDOMAIN_HANDLER_MODULE = u"Arainco_InPlaceAFamiliaEscalera_HandlerModule"
_APPDOMAIN_HANDLER_VERSION = u"Arainco_InPlaceAFamiliaEscalera_HandlerVersion"
_APPDOMAIN_PENDING_CONVERT = u"Arainco_InPlaceAFamiliaEscalera_PendingConvert"
_APPDOMAIN_IDLING_HANDLER = u"Arainco_InPlaceAFamiliaEscalera_IdlingHandler"

# Subir este valor fuerza recrear ExternalEvent (evita handler viejo sin Comments).
_TOOL_CODE_VERSION = u"2026-03-20-comments-v4"

_TOOL_TITLE = u"Arainco: In-Place a familia escalera"
_TRANSACTION_GROUP = u"Arainco: In-Place a familia escalera"
_ALREADY_RUNNING = u"La herramienta ya esta en ejecucion."

# Categorías / modos ofrecidos en el combo (extensible).
# Nota Revit: OST_Stairs es categoría de sistema; no se puede asignar a una
# familia cargable creada desde Generic Model. Para el flujo Escalera usamos
# Modelo genérico + FAMILY_CAN_HOST_REBAR (host válido de armadura).
_CATEGORY_MODE_GENERIC_HOST_REBAR = u"generic_host_rebar"
_CATEGORY_CHOICES = (
    # (etiqueta UI, mode, valor Comments)
    (u"Escalera (Modelo genérico + armadura)", _CATEGORY_MODE_GENERIC_HOST_REBAR, u"Escalera"),
)

_CATEGORY_COMMENTS_BY_MODE = dict(
    (mode, comments) for _label, mode, comments in _CATEGORY_CHOICES
)

_MIN_SOLID_VOLUME = 1.0e-9
_TEMPLATE_NAMES = (
    u"Metric Generic Model.rft",
    u"Modelo genérico métrico.rft",
    u"Generic Model.rft",
    u"Modelo generico metrico.rft",
)


def _eid_int(eid_or_elem):
    if eid_or_elem is None:
        return None
    try:
        eid = eid_or_elem.Id
    except Exception:
        eid = eid_or_elem
    try:
        return int(eid.Value)
    except Exception:
        try:
            return int(eid.IntegerValue)
        except Exception:
            return None


def _solid_stats(solids):
    n = len(solids) if solids else 0
    vol = 0.0
    for s in solids or []:
        try:
            vol += float(s.Volume)
        except Exception:
            pass
    return n, vol


def _element_solid_stats(elem):
    solids = _extract_solids(elem) if elem is not None else []
    return _solid_stats(solids)


def _find_family_by_name(doc, name):
    target = _as_unicode(name)
    for f in FilteredElementCollector(doc).OfClass(Family):
        try:
            if _as_unicode(f.Name) == target:
                return f
        except Exception:
            continue
    return None


def _safe_elem_name(elem):
    try:
        return _as_unicode(elem.Name)
    except Exception:
        return u"(sin Name)"


def _resolve_unique_loadable_name(doc, desired_name):
    """
    Nombre libre de colisión con cualquier Family del proyecto.

    El Model In-Place deja una Family IsInPlace con el mismo nombre; LoadFamily
    no puede sobrescribirla con un .rfa cargable, así que hay que usar otro nombre.
    """
    base = _sanitize_family_name(desired_name)
    if not base:
        base = u"Escalera_InPlace"

    existing = _find_family_by_name(doc, base)
    if existing is None:
        return base

    # Conflicto (casi siempre la Family in-place del mismo elemento).
    candidates = [
        base + u"_Cargable",
        base + u"_GM",
    ]
    for cand in candidates:
        if _find_family_by_name(doc, cand) is None:
            return cand

    for i in range(2, 200):
        cand = u"{0}_Cargable_{1}".format(base, i)
        if _find_family_by_name(doc, cand) is None:
            return cand
    return u"{0}_Cargable_X".format(base)


def _select_and_show(uidoc, elem):
    if uidoc is None or elem is None:
        return False
    try:
        from System.Collections.Generic import List as ClrList

        ids = ClrList[ElementId]()
        ids.Add(elem.Id)
        uidoc.Selection.SetElementIds(ids)
    except Exception:
        return False
    try:
        uidoc.ShowElements(elem.Id)
    except Exception:
        try:
            uidoc.ShowElements(elem)
        except Exception:
            pass
    return True


def _post_verify_instance(doc, uidoc, new_fi):
    """Comprueba que la instancia existe, tiene sólidos y es host válido."""
    if new_fi is None:
        return False, u"Instancia nueva es None."

    live = doc.GetElement(new_fi.Id) if new_fi.Id is not None else None
    if live is None:
        return False, u"La instancia no existe en el documento tras el commit."

    n_solids, vol = _element_solid_stats(live)
    if n_solids < 1 or vol <= _MIN_SOLID_VOLUME:
        return False, u"Instancia sin sólidos visibles (n={0}, vol={1}).".format(
            n_solids, vol
        )

    try:
        if not bool(RebarHostData.IsValidHost(live)):
            return False, u"IsValidHost=False tras commit."
    except Exception as ex:
        return False, u"IsValidHost error: {0}".format(_as_unicode(ex))

    _select_and_show(uidoc, live)
    return True, u"OK"


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except Exception:
        return str(text)


def _mostrar_aviso(uiapp, instruction, content=u"", ok_text=u"Entendido"):
    hwnd = None
    try:
        if uiapp is not None:
            hwnd = revit_main_hwnd(uiapp)
    except Exception:
        pass
    try:
        from bimtools_instruction_dialog import show_message_dialog

        show_message_dialog(
            _TOOL_TITLE,
            instruction,
            content=content,
            ok_text=ok_text,
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass
    try:
        body = instruction
        if content:
            body = instruction + u"\n\n" + content
        TaskDialog.Show(_TOOL_TITLE, body)
    except Exception:
        pass


def _attach_revit_owner(win, uiapp):
    if win is None or uiapp is None:
        return
    try:
        from System.Windows.Interop import WindowInteropHelper

        hwnd = revit_main_hwnd(uiapp)
        if hwnd is not None:
            WindowInteropHelper(win).Owner = hwnd
    except Exception:
        pass


def _prepare_window(win, uiapp):
    if win is None:
        return
    try:
        hwnd = revit_main_hwnd(uiapp)
        bind_center_wpf_on_revit_monitor(win, hwnd)
        position_wpf_window_center_on_monitor(win, hwnd)
    except Exception:
        pass
    _attach_revit_owner(win, uiapp)


def _window_is_alive(win):
    if win is None:
        return False
    try:
        _ = win.Title
    except Exception:
        return False
    try:
        if hasattr(win, "IsLoaded") and (not win.IsLoaded):
            return False
    except Exception:
        pass
    return True


def _find_window_by_tool_title():
    try:
        from System.Windows import Application

        app = Application.Current
        if app is None:
            return None
        for w in app.Windows:
            try:
                txt = w.FindName(u"TxtTitle")
                if txt is not None and _as_unicode(txt.Text) == _TOOL_TITLE:
                    if _window_is_alive(w):
                        return w
            except Exception:
                continue
    except Exception:
        return None
    return None


def _get_active_window():
    try:
        win = AppDomain.CurrentDomain.GetData(_APPDOMAIN_WINDOW_KEY)
    except Exception:
        win = None
    if _window_is_alive(win):
        return win
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, None)
    except Exception:
        pass
    return _find_window_by_tool_title()


def _set_active_window(win):
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, win)
    except Exception:
        pass


def _clear_active_window():
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, None)
    except Exception:
        pass


def _activate_existing(win, uiapp):
    try:
        if win.WindowState == WindowState.Minimized:
            win.WindowState = WindowState.Normal
    except Exception:
        pass
    try:
        win.Activate()
        win.Focus()
    except Exception:
        pass
    _mostrar_aviso(uiapp, _ALREADY_RUNNING)


def _resolve_manual_path():
    candidates = []
    try:
        import bimtools_paths

        pb = bimtools_paths.get_pushbutton_dir()
        if pb:
            candidates.append(os.path.join(pb, u"manual_usuario.html"))
    except Exception:
        pass
    try:
        ext_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        for tab_name in os.listdir(ext_dir):
            if not tab_name.endswith(u".tab"):
                continue
            panel = os.path.join(ext_dir, tab_name, u"Modelado.panel")
            if not os.path.isdir(panel):
                continue
            for pb_name in os.listdir(panel):
                if u"InPlaceAFamiliaEscalera" not in pb_name:
                    continue
                candidates.append(
                    os.path.join(panel, pb_name, u"manual_usuario.html")
                )
    except Exception:
        pass
    seen = set()
    for path in candidates:
        try:
            ap = os.path.normpath(os.path.abspath(path))
        except Exception:
            continue
        if ap in seen:
            continue
        seen.add(ap)
        if os.path.isfile(ap):
            return ap
    return None


def _open_manual(uiapp):
    path = _resolve_manual_path()
    if not path:
        _mostrar_aviso(
            uiapp,
            u"No se encontró manual_usuario.html.",
            content=u"Debe estar en la carpeta del pushbutton de la herramienta.",
        )
        return
    try:
        os.startfile(path)
    except Exception as ex:
        _mostrar_aviso(
            uiapp,
            u"No se pudo abrir el manual.",
            content=_as_unicode(ex),
        )


def _sanitize_family_name(name):
    raw = _as_unicode(name).strip()
    if not raw:
        raw = u"Escalera_InPlace"
    cleaned = re.sub(r'[<>:"/\\|?*]', u"_", raw)
    cleaned = cleaned.strip(u" .")
    if not cleaned:
        cleaned = u"Escalera_InPlace"
    if len(cleaned) > 120:
        cleaned = cleaned[:120]
    return cleaned


# ── Geometría / conversión ───────────────────────────────────────────────────


def _is_model_in_place(elem):
    if elem is None or not isinstance(elem, FamilyInstance):
        return False
    try:
        sym = elem.Symbol
        if sym is None:
            return False
        fam = sym.Family
        if fam is None:
            return False
        return bool(fam.IsInPlace)
    except Exception:
        return False


class _InPlaceSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return _is_model_in_place(elem)

    def AllowReference(self, reference, point):
        return False


def _collect_solids_from_geometry(geom_elem, solids_out):
    if geom_elem is None:
        return
    for go in geom_elem:
        if go is None:
            continue
        solid = go if isinstance(go, Solid) else None
        if solid is not None:
            try:
                if solid.Volume > _MIN_SOLID_VOLUME:
                    solids_out.append(solid)
            except Exception:
                pass
            continue
        if isinstance(go, GeometryInstance):
            try:
                inst_geom = go.GetInstanceGeometry()
            except Exception:
                inst_geom = None
            _collect_solids_from_geometry(inst_geom, solids_out)


def _extract_solids(elem):
    opts = Options()
    opts.ComputeReferences = False
    opts.IncludeNonVisibleObjects = False
    try:
        opts.DetailLevel = ViewDetailLevel.Fine
    except Exception:
        pass
    try:
        geom = elem.get_Geometry(opts)
    except Exception:
        geom = None
    solids = []
    _collect_solids_from_geometry(geom, solids)
    # Copias independientes (evita referencias vivas al documento).
    clones = []
    identity = Transform.Identity
    for solid in solids:
        try:
            cloned = SolidUtils.CreateTransformed(solid, identity)
            if cloned is not None and cloned.Volume > _MIN_SOLID_VOLUME:
                clones.append(cloned)
        except Exception:
            continue
    return clones


def _instance_transform(elem, solids=None):
    """
    Transform de colocación para el in-place.

    Muchos Model In-Place reportan GetTransform() = Identity en (0,0,0) aunque
    la geometría esté en coordenadas de proyecto. En ese caso se usa el centro
    del bbox / sólidos como origen de familia.
    """
    tf = None
    try:
        tf = elem.GetTransform()
    except Exception:
        tf = None

    if tf is None:
        tf = Transform.Identity
        loc = getattr(elem, "Location", None)
        if isinstance(loc, LocationPoint):
            try:
                tf.Origin = loc.Point
            except Exception:
                pass

    geo_origin = None
    bb = None
    try:
        bb = elem.get_BoundingBox(None)
    except Exception:
        bb = None
    if bb is not None:
        try:
            geo_origin = (bb.Min + bb.Max) * 0.5
        except Exception:
            geo_origin = None

    if geo_origin is None and solids:
        # Centroid aproximado por vértices de aristas (Solid no tiene GetBoundingBox).
        try:
            pts = []
            for s in solids:
                try:
                    for edge in s.Edges:
                        c = edge.AsCurve()
                        if c is None:
                            continue
                        pts.append(c.GetEndPoint(0))
                        pts.append(c.GetEndPoint(1))
                except Exception:
                    continue
            if pts:
                sx = sy = sz = 0.0
                for p in pts:
                    sx += p.X
                    sy += p.Y
                    sz += p.Z
                n = float(len(pts))
                geo_origin = XYZ(sx / n, sy / n, sz / n)
        except Exception:
            geo_origin = None

    try:
        origin = tf.Origin
        origin_is_zero = origin.DistanceTo(XYZ.Zero) < 1.0e-6
    except Exception:
        origin_is_zero = True
        origin = XYZ.Zero

    if geo_origin is not None:
        try:
            geo_far = geo_origin.DistanceTo(XYZ.Zero) > 1.0
        except Exception:
            geo_far = False
        # Transform inútil (origen 0) con geometría lejos → anclar al centro geométrico.
        if origin_is_zero and geo_far:
            fixed = Transform.Identity
            try:
                fixed.BasisX = tf.BasisX
                fixed.BasisY = tf.BasisY
                fixed.BasisZ = tf.BasisZ
            except Exception:
                pass
            fixed.Origin = geo_origin
            return fixed
        # Origen del transform lejos del sólido → preferir centro geométrico (misma rotación).
        try:
            if origin.DistanceTo(geo_origin) > 50.0:  # ~15 m
                fixed = Transform.Identity
                try:
                    fixed.BasisX = tf.BasisX
                    fixed.BasisY = tf.BasisY
                    fixed.BasisZ = tf.BasisZ
                except Exception:
                    pass
                fixed.Origin = geo_origin
                return fixed
        except Exception:
            pass

    return tf


def _find_level_for_instance(doc, elem, origin):
    try:
        lid = elem.LevelId
        if lid is not None and lid != ElementId.InvalidElementId:
            lv = doc.GetElement(lid)
            if isinstance(lv, Level):
                return lv
    except Exception:
        pass
    for bip in (
        BuiltInParameter.FAMILY_LEVEL_PARAM,
        BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
        BuiltInParameter.SCHEDULE_LEVEL_PARAM,
    ):
        try:
            p = elem.get_Parameter(bip)
            if p is None or p.AsElementId() is None:
                continue
            eid = p.AsElementId()
            lv = doc.GetElement(eid)
            if isinstance(lv, Level):
                return lv
        except Exception:
            continue
    levels = list(
        FilteredElementCollector(doc).OfClass(Level).ToElements()
    )
    if not levels:
        return None
    z = 0.0
    try:
        z = float(origin.Z)
    except Exception:
        pass
    best = None
    best_dz = None
    for lv in levels:
        try:
            elev = float(lv.Elevation)
        except Exception:
            continue
        dz = abs(elev - z)
        if best is None or dz < best_dz:
            best = lv
            best_dz = dz
    return best


def _find_generic_model_template(app):
    version = u""
    try:
        version = _as_unicode(app.VersionNumber).strip()
    except Exception:
        version = u""

    search_roots = []
    try:
        lib_paths = app.GetLibraryPaths()
        if lib_paths is not None:
            for key in lib_paths.Keys:
                try:
                    search_roots.append(_as_unicode(lib_paths[key]))
                except Exception:
                    pass
    except Exception:
        pass

    program_data = os.environ.get(u"ProgramData", u"C:\\ProgramData")
    if version:
        search_roots.extend(
            [
                os.path.join(
                    program_data,
                    u"Autodesk",
                    u"RVT {0}".format(version),
                    u"Family Templates",
                ),
                os.path.join(
                    program_data,
                    u"Autodesk",
                    u"Revit {0}".format(version),
                    u"Family Templates",
                ),
            ]
        )

    seen = set()
    for root in search_roots:
        if not root:
            continue
        try:
            ap = os.path.normpath(os.path.abspath(root))
        except Exception:
            continue
        if ap in seen or not os.path.isdir(ap):
            continue
        seen.add(ap)
        for dirpath, _dirnames, filenames in os.walk(ap):
            lower_map = dict((f.lower(), f) for f in filenames)
            for wanted in _TEMPLATE_NAMES:
                key = wanted.lower()
                if key in lower_map:
                    return os.path.join(dirpath, lower_map[key])
    return None


class _OverwriteFamilyLoadOptions(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        try:
            overwriteParameterValues.Value = True
        except Exception:
            try:
                overwriteParameterValues = True
            except Exception:
                pass
        return True

    def OnSharedFamilyFound(
        self, sharedFamily, familyInUse, source, overwriteParameterValues
    ):
        try:
            source.Value = FamilySource.Family
        except Exception:
            try:
                source = FamilySource.Family
            except Exception:
                pass
        try:
            overwriteParameterValues.Value = True
        except Exception:
            try:
                overwriteParameterValues = True
            except Exception:
                pass
        return True


def _comments_for_mode(category_mode):
    try:
        return _CATEGORY_COMMENTS_BY_MODE.get(category_mode) or u"Escalera"
    except Exception:
        return u"Escalera"


def _find_comments_parameter(elem):
    """Localiza Comments / Comentarios en instancia (y respaldo por nombre)."""
    if elem is None:
        return None
    try:
        p = elem.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if p is not None:
            return p
    except Exception:
        pass
    for name in (u"Comments", u"Comentarios", u"Comment"):
        try:
            p = elem.LookupParameter(name)
            if p is not None:
                return p
        except Exception:
            continue
    return None


def _read_comments(elem):
    p = _find_comments_parameter(elem)
    if p is None:
        return None
    try:
        return _as_unicode(p.AsString())
    except Exception:
        try:
            return _as_unicode(p.AsValueString())
        except Exception:
            return None


def _set_instance_comments(elem, comments):
    """Escribe la categoría de flujo en Comments de instancia."""
    if elem is None:
        return False
    text = _as_unicode(comments).strip()
    if not text:
        return False

    p = _find_comments_parameter(elem)
    if p is None:
        return False
    if p.IsReadOnly:
        return False

    set_ok = False
    try:
        p.Set(text)
        set_ok = True
    except Exception:
        try:
            from System import String

            p.Set(String(text))
            set_ok = True
        except Exception:
            return False

    if not set_ok:
        return False

    read_back = _read_comments(elem)
    return _as_unicode(read_back).strip() == text


def _ensure_comments_after_commit(doc, fi, comments):
    """Si Comments quedó vacío tras el commit, lo escribe en una transacción corta."""
    current = _read_comments(fi)
    wanted = _as_unicode(comments).strip()
    if current is not None and _as_unicode(current).strip() == wanted:
        return True
    t = Transaction(doc, u"Arainco: Comments categoría Escalera")
    t.Start()
    try:
        ok = _set_instance_comments(fi, wanted)
        t.Commit()
        return ok
    except Exception:
        try:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
        except Exception:
            pass
        return False


def _set_can_host_rebar(fam_doc):
    """Activa «Can Host Rebar» en OwnerFamily (Modelo genérico)."""
    owner = fam_doc.OwnerFamily
    p = owner.get_Parameter(BuiltInParameter.FAMILY_CAN_HOST_REBAR)
    if p is None:
        raise Exception(
            u"La plantilla no expone el parámetro «Can Host Rebar». "
            u"Use Metric Generic Model."
        )
    if p.IsReadOnly:
        raise Exception(u"No se pudo activar «Can Host Rebar» (parámetro solo lectura).")
    p.Set(1)


def _create_family_document(app, solids_world, instance_tf, category_mode, family_name):
    template = _find_generic_model_template(app)
    if not template:
        raise Exception(
            u"No se encontró la plantilla Metric Generic Model (.rft). "
            u"Verifique la instalación de Revit."
        )

    fam_doc = app.NewFamilyDocument(template)
    if fam_doc is None:
        raise Exception(u"No se pudo crear el documento de familia.")

    inv = instance_tf.Inverse
    created = 0
    t = Transaction(fam_doc, u"Arainco: Geometría FreeForm Escalera")
    t.Start()
    try:
        if category_mode != _CATEGORY_MODE_GENERIC_HOST_REBAR:
            raise Exception(u"Modo de categoría no soportado.")
        _set_can_host_rebar(fam_doc)
        for solid in solids_world:
            local = SolidUtils.CreateTransformed(solid, inv)
            if local is None or local.Volume <= _MIN_SOLID_VOLUME:
                continue
            FreeFormElement.Create(fam_doc, local)
            created += 1
        if created < 1:
            raise Exception(u"No se pudo crear geometría FreeForm válida.")
        t.Commit()
    except Exception:
        try:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
        except Exception:
            pass
        try:
            fam_doc.Close(False)
        except Exception:
            pass
        raise

    tmp_dir = tempfile.mkdtemp(prefix=u"arainco_inplace_stair_")
    safe_name = _sanitize_family_name(family_name)
    rfa_path = os.path.join(tmp_dir, safe_name + u".rfa")
    opts = SaveAsOptions()
    opts.OverwriteExistingFile = True
    try:
        fam_doc.SaveAs(rfa_path, opts)
    finally:
        try:
            fam_doc.Close(False)
        except Exception:
            pass

    return rfa_path, safe_name


def _activate_first_symbol(doc, family):
    if bool(family.IsInPlace):
        raise Exception(
            u"La familia '{0}' es in-place; no se puede usar como familia cargable.".format(
                _safe_elem_name(family)
            )
        )
    symbols = list(family.GetFamilySymbolIds())
    if not symbols:
        raise Exception(u"La familia cargada no tiene tipos.")
    sym = doc.GetElement(symbols[0])
    if sym is None:
        raise Exception(u"No se pudo obtener el tipo de familia.")
    if not sym.IsActive:
        sym.Activate()
    return sym


def _apply_instance_rotation(doc, fi, instance_tf):
    """Alinea la rotación en planta (eje Z) con el transform del in-place."""
    try:
        bx = instance_tf.BasisX
        angle = math.atan2(bx.Y, bx.X)
    except Exception:
        return
    if abs(angle) < 1.0e-9:
        return
    loc = fi.Location
    if not isinstance(loc, LocationPoint):
        return
    try:
        origin = loc.Point
    except Exception:
        return
    try:
        axis = Line.CreateBound(origin, origin + XYZ.BasisZ)
    except Exception:
        return
    try:
        loc.Rotate(axis, angle)
    except Exception:
        try:
            ElementTransformUtils.RotateElement(doc, fi.Id, axis, angle)
        except Exception:
            pass


def _try_load_family(doc, rfa_path):
    """Intenta LoadFamily; nunca acepta una Family IsInPlace."""
    fam_ref = clr.Reference[Family]()
    loaded = False
    load_ex = None
    try:
        loaded = bool(doc.LoadFamily(rfa_path, _OverwriteFamilyLoadOptions(), fam_ref))
    except Exception as ex:
        load_ex = ex
        loaded = False
    family = None
    try:
        family = fam_ref.Value
    except Exception:
        family = None

    if family is not None:
        try:
            if bool(family.IsInPlace):
                family = None
                loaded = False
        except Exception:
            pass

    return loaded, family, load_ex


def _load_place_and_replace(
    doc, app, inplace_elem, solids, instance_tf, category_mode, family_name
):
    inplace_id = inplace_elem.Id
    desired = _sanitize_family_name(family_name)
    unique_name = _resolve_unique_loadable_name(doc, desired)

    rfa_path, safe_name = _create_family_document(
        app, solids, instance_tf, category_mode, unique_name
    )

    tg = TransactionGroup(doc, _TRANSACTION_GROUP)
    tg.Start()
    new_fi = None
    try:
        t = Transaction(doc, u"Arainco: Cargar y colocar familia escalera")
        t.Start()
        try:
            loaded, family, load_ex = _try_load_family(doc, rfa_path)

            if (not loaded) or family is None:
                # Solo reutilizar por nombre si es cargable (nunca in-place).
                cand = _find_family_by_name(doc, safe_name)
                if cand is not None and (not bool(cand.IsInPlace)):
                    family = cand
                else:
                    detail = _as_unicode(load_ex) if load_ex else u"LoadFamily=False"
                    if cand is not None and bool(cand.IsInPlace):
                        detail += (
                            u"; existe familia in-place homónima Id={0} "
                            u"(no usable como cargable)"
                        ).format(_eid_int(cand))
                    raise Exception(
                        u"No se pudo cargar la familia «{0}». {1}".format(
                            safe_name, detail
                        )
                    )

            symbol = _activate_first_symbol(doc, family)
            origin = instance_tf.Origin
            level = _find_level_for_instance(doc, inplace_elem, origin)

            if level is not None:
                new_fi = doc.Create.NewFamilyInstance(
                    origin, symbol, level, StructuralType.NonStructural
                )
            else:
                new_fi = doc.Create.NewFamilyInstance(
                    origin, symbol, StructuralType.NonStructural
                )

            try:
                doc.Regenerate()
            except Exception:
                pass

            _apply_instance_rotation(doc, new_fi, instance_tf)

            comments_val = _comments_for_mode(category_mode)
            if not _set_instance_comments(new_fi, comments_val):
                raise Exception(
                    u"No se pudo escribir Comments='{0}' en la instancia.".format(
                        comments_val
                    )
                )

            if not RebarHostData.IsValidHost(new_fi):
                raise Exception(
                    u"La instancia creada no es un host válido de armadura. "
                    u"No se eliminó el Model In-Place."
                )

            n_solids, vol = _element_solid_stats(new_fi)
            if n_solids < 1 or vol <= _MIN_SOLID_VOLUME:
                # Aún sin sólidos tras Regenerate: no borrar in-place.
                raise Exception(
                    u"La instancia nueva no tiene sólidos (n={0}, vol={1}). "
                    u"Revisar origen/transform de familia. In-place NO eliminado.".format(
                        n_solids, vol
                    )
                )

            doc.Delete(inplace_id)
            t.Commit()
        except Exception:
            try:
                if t.HasStarted() and not t.HasEnded():
                    t.RollBack()
            except Exception:
                pass
            raise

        tg.Assimilate()
    except Exception:
        try:
            tg.RollBack()
        except Exception:
            pass
        raise
    finally:
        try:
            if os.path.isfile(rfa_path):
                os.remove(rfa_path)
        except Exception:
            pass
        try:
            tmp_dir = os.path.dirname(rfa_path)
            if os.path.isdir(tmp_dir):
                os.rmdir(tmp_dir)
        except Exception:
            pass

    comments_val = _comments_for_mode(category_mode)
    try:
        live_fi = doc.GetElement(new_fi.Id) if new_fi is not None else None
        _ensure_comments_after_commit(doc, live_fi or new_fi, comments_val)
    except Exception:
        pass

    return new_fi, comments_val


def convert_inplace_to_stair_family(uiapp, inplace_id, category_mode, family_name):
    """
    Convierte el Model In-Place indicado.
    Devuelve (ok, mensaje).
    """
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        return False, u"No hay documento activo."
    doc = uidoc.Document
    if doc.IsFamilyDocument:
        return False, u"Abra un proyecto (no un documento de familia)."

    elem = doc.GetElement(inplace_id)
    if not _is_model_in_place(elem):
        return False, u"El elemento ya no es un Model In-Place válido."

    solids = _extract_solids(elem)
    if not solids:
        return False, u"El Model In-Place no tiene sólidos válidos."

    instance_tf = _instance_transform(elem, solids)
    app = doc.Application

    comments_val = _comments_for_mode(category_mode)
    try:
        new_fi, comments_val = _load_place_and_replace(
            doc,
            app,
            elem,
            solids,
            instance_tf,
            category_mode,
            family_name,
        )
    except Exception as ex:
        return False, _as_unicode(ex)

    verify_ok, verify_msg = _post_verify_instance(doc, uidoc, new_fi)
    try:
        live = doc.GetElement(new_fi.Id)
        cmt = _read_comments(live)
        if _as_unicode(cmt).strip() != _as_unicode(comments_val).strip():
            _ensure_comments_after_commit(doc, live, comments_val)
    except Exception:
        pass

    name = family_name
    try:
        name = _as_unicode(new_fi.Symbol.Family.Name)
    except Exception:
        name = _sanitize_family_name(family_name)

    if not verify_ok:
        return False, u"La conversión no dejó una instancia usable: {0}.".format(
            verify_msg
        )

    return True, (
        u"Familia «{0}» creada y colocada. Comments={1}. Model In-Place eliminado."
    ).format(name, comments_val)


# ── ExternalEvent ────────────────────────────────────────────────────────────


def _set_pending_convert(inplace_id, category_mode, family_name):
    try:
        AppDomain.CurrentDomain.SetData(
            _APPDOMAIN_PENDING_CONVERT,
            (inplace_id, category_mode, family_name),
        )
    except Exception:
        pass


def _take_pending_convert():
    try:
        payload = AppDomain.CurrentDomain.GetData(_APPDOMAIN_PENDING_CONVERT)
    except Exception:
        payload = None
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_PENDING_CONVERT, None)
    except Exception:
        pass
    if payload is None:
        return None, None, None
    try:
        return payload[0], payload[1], payload[2]
    except Exception:
        return None, None, None


def _peek_pending_convert():
    try:
        payload = AppDomain.CurrentDomain.GetData(_APPDOMAIN_PENDING_CONVERT)
    except Exception:
        return None
    return payload


class _ToolHandler(IExternalEventHandler):
    def __init__(self):
        self.action = None  # u"pick" | u"convert"
        self.ctrl = None
        self.convert_inplace_id = None
        self.convert_category = None
        self.convert_family_name = None

    def clear_convert_args(self):
        self.convert_inplace_id = None
        self.convert_category = None
        self.convert_family_name = None

    def Execute(self, uiapp):
        ctrl = self.ctrl
        action = self.action
        self.action = None

        # Siempre resolver funciones del módulo cargado ahora (no del handler viejo).
        try:
            import inplace_a_familia_escalera as _mod
            run_pending = _mod._run_pending_convert_if_any
            peek = _mod._peek_pending_convert
        except Exception:
            run_pending = _run_pending_convert_if_any
            peek = _peek_pending_convert

        if action == u"convert":
            self.clear_convert_args()
            run_pending(uiapp)
            return

        if action == u"pick":
            try:
                if ctrl is None:
                    return
                ctrl.execute_pick(uiapp)
            except Exception as ex:
                try:
                    _mostrar_aviso(
                        uiapp,
                        u"Error en la herramienta.",
                        content=_as_unicode(ex),
                    )
                except Exception:
                    pass
                if ctrl is not None:
                    try:
                        ctrl.set_status(u"Error.")
                    except Exception:
                        pass
                    try:
                        ctrl.show_window()
                    except Exception:
                        pass
            return

        if peek() is not None:
            run_pending(uiapp)

    def GetName(self):
        return _TOOL_TITLE


def _execute_convert_after_ui_closed(uiapp, inplace_id, category_mode, family_name):
    """Convierte con la UI ya cerrada."""
    if inplace_id is None:
        _mostrar_aviso(uiapp, u"No hay Model In-Place seleccionado.")
        return
    if not family_name:
        _mostrar_aviso(uiapp, u"Indique un nombre de familia.")
        return
    try:
        ok, msg = convert_inplace_to_stair_family(
            uiapp,
            inplace_id,
            category_mode,
            family_name,
        )
    except Exception as ex:
        _mostrar_aviso(
            uiapp,
            u"No se pudo convertir.",
            content=_as_unicode(ex),
        )
        return

    if ok:
        _mostrar_aviso(uiapp, u"Conversión completada.", content=msg)
    else:
        _mostrar_aviso(uiapp, u"No se pudo convertir.", content=msg)


def _run_pending_convert_if_any(uiapp, source=None):
    """Consume el pending y ejecuta la conversión (ExternalEvent o Idling)."""
    if _peek_pending_convert() is None:
        return False
    inplace_id, category, family_name = _take_pending_convert()
    try:
        _execute_convert_after_ui_closed(
            uiapp, inplace_id, category, family_name
        )
    except Exception as ex:
        try:
            _mostrar_aviso(
                uiapp,
                u"Error en la herramienta.",
                content=_as_unicode(ex),
            )
        except Exception:
            pass
    return True


def _arm_idling_convert(uiapp):
    """Fallback fiable: corre en el próximo ciclo Idle de Revit."""
    if uiapp is None:
        return

    # Quitar handler previo si quedó colgado
    try:
        old = AppDomain.CurrentDomain.GetData(_APPDOMAIN_IDLING_HANDLER)
        if old is not None:
            try:
                uiapp.Idling -= old
            except Exception:
                pass
    except Exception:
        pass

    def _on_idle(sender, args):
        try:
            h = AppDomain.CurrentDomain.GetData(_APPDOMAIN_IDLING_HANDLER)
            if h is not None:
                try:
                    uiapp.Idling -= h
                except Exception:
                    pass
            AppDomain.CurrentDomain.SetData(_APPDOMAIN_IDLING_HANDLER, None)
        except Exception:
            pass
        try:
            import inplace_a_familia_escalera as _mod

            _mod._run_pending_convert_if_any(uiapp)
        except Exception:
            _run_pending_convert_if_any(uiapp)

    try:
        handler = EventHandler[IdlingEventArgs](_on_idle)
    except Exception:
        # IronPython a veces acepta el callable directo
        handler = _on_idle

    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_IDLING_HANDLER, handler)
        uiapp.Idling += handler
    except Exception:
        pass


def _get_or_create_external_event(force_new=False):
    """
    Recrea el ExternalEvent si cambió el módulo/versión del código.
    Sin esto, Raise=Accepted ejecuta un handler viejo (p. ej. sin Comments).
    """
    try:
        ev = AppDomain.CurrentDomain.GetData(_APPDOMAIN_EVENT_KEY)
        handler = AppDomain.CurrentDomain.GetData(_APPDOMAIN_HANDLER_KEY)
        mod = AppDomain.CurrentDomain.GetData(_APPDOMAIN_HANDLER_MODULE)
        ver = AppDomain.CurrentDomain.GetData(_APPDOMAIN_HANDLER_VERSION)
    except Exception:
        ev = None
        handler = None
        mod = None
        ver = None

    reuse = False
    if (
        (not force_new)
        and ev is not None
        and handler is not None
        and mod == __name__
        and ver == _TOOL_CODE_VERSION
        and hasattr(handler, "clear_convert_args")
        and hasattr(handler, "Execute")
    ):
        reuse = True

    if reuse:
        return ev, handler

    handler = _ToolHandler()
    ev = ExternalEvent.Create(handler)
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_EVENT_KEY, ev)
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_HANDLER_KEY, handler)
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_HANDLER_MODULE, __name__)
        AppDomain.CurrentDomain.SetData(
            _APPDOMAIN_HANDLER_VERSION, _TOOL_CODE_VERSION
        )
    except Exception:
        pass
    return ev, handler


# ── UI ───────────────────────────────────────────────────────────────────────


_BODY_XAML = u"""
<StackPanel>
  <TextBlock Text="Model In-Place" Style="{StaticResource Label}" Margin="0,0,0,6"/>
  <Grid Margin="0,0,0,12">
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="Auto"/>
    </Grid.ColumnDefinitions>
    <TextBlock x:Name="TxtInPlace" Grid.Column="0" TextWrapping="Wrap"
               VerticalAlignment="Center" Foreground="#95B8CC"
               Text="Ninguno seleccionado."/>
    <Button x:Name="BtnPick" Grid.Column="1" Content="Seleccionar"
            Style="{StaticResource BtnSelectOutline}" MinWidth="110"
            Margin="10,0,0,0" Padding="10,4"
            ToolTip="Seleccionar un Model In-Place en el modelo"/>
  </Grid>

  <TextBlock Text="Modo / categoría" Style="{StaticResource Label}" Margin="0,0,0,6"/>
  <ComboBox x:Name="CmbCategory" Style="{StaticResource Combo}"
            Margin="0,0,0,12" MinHeight="30"/>

  <TextBlock Text="Nombre de familia" Style="{StaticResource Label}" Margin="0,0,0,6"/>
  <TextBox x:Name="TxtFamilyName" Style="{StaticResource BimToolsTextBoxDark}"
           MinHeight="30" VerticalContentAlignment="Center" Padding="10,4"/>
</StackPanel>
"""

_FOOTER_LEADING_XAML = (
    u'<Button x:Name="BtnManual" Content="Manual" '
    u'Style="{{StaticResource BtnSelectOutline}}" '
    u'Background="{bg}" MinWidth="96" Padding="8,2" '
    u'ToolTip="Abrir manual de usuario" VerticalAlignment="Center"/>'
).format(bg=BTN_MANUAL)

_FOOTER_ACTIONS_XAML = u"""
<Button x:Name="BtnClose" Content="Cerrar" Margin="0,0,8,0"
        Style="{StaticResource BtnSelectOutline}" MinWidth="100"/>
<Button x:Name="BtnConvert" Content="Convertir" IsDefault="True"
        Style="{StaticResource BtnPrimary}" MinWidth="120"
        ToolTip="Crear familia Escalera, colocarla y eliminar el Model In-Place"/>
"""


def _build_xaml():
    return build_simple_tool_xaml(
        title=_TOOL_TITLE,
        styles_xml=BIMTOOLS_DARK_STYLES_XML,
        body_xaml=_BODY_XAML,
        footer_leading_xaml=_FOOTER_LEADING_XAML,
        footer_actions_xaml=_FOOTER_ACTIONS_XAML,
        footer_hint_xaml=(
            u"Revit no permite familia cargable de categoría Escalera (sistema). "
            u"Se crea Modelo genérico con «Can Host Rebar». Se elimina el in-place "
            u"tras una conversión correcta."
        ),
        width=520,
        min_width=440,
        height=0,
        min_height=0,
        resize_mode=u"NoResize",
        size_to_content_height=True,
    )


class _ToolController(object):
    def __init__(self, revit, win, ext_event, handler):
        self.revit = revit
        self.win = win
        self.ext_event = ext_event
        self.handler = handler
        self.inplace_id = None
        self._busy = False

        self.txt_inplace = win.FindName(u"TxtInPlace")
        self.txt_name = win.FindName(u"TxtFamilyName")
        self.cmb = win.FindName(u"CmbCategory")
        self.txt_status = win.FindName(u"TxtStatus")
        self.txt_subtitle = win.FindName(u"TxtSubtitle")

        if self.txt_subtitle is not None:
            try:
                self.txt_subtitle.Text = (
                    u"Convierte un Model In-Place en familia cargable (Modelo genérico "
                    u"con hospedaje de armadura) en la misma posición."
                )
            except Exception:
                pass

        self._fill_categories()
        self.set_status(u"Seleccione un Model In-Place.")

    def _fill_categories(self):
        if self.cmb is None:
            return
        try:
            self.cmb.Items.Clear()
        except Exception:
            pass
        for label, mode, _comments in _CATEGORY_CHOICES:
            item = ComboBoxItem()
            item.Content = label
            item.Tag = mode
            try:
                self.cmb.Items.Add(item)
            except Exception:
                pass
        try:
            if self.cmb.Items.Count > 0:
                self.cmb.SelectedIndex = 0
        except Exception:
            pass

    def set_status(self, text):
        if self.txt_status is None:
            return
        try:
            self.txt_status.Text = _as_unicode(text)
        except Exception:
            pass

    def hide_window(self):
        try:
            self.win.Hide()
        except Exception:
            pass

    def show_window(self):
        try:
            self.win.Show()
            self.win.Activate()
        except Exception:
            pass

    def _selected_category(self):
        try:
            item = self.cmb.SelectedItem
            if item is not None and item.Tag is not None:
                return item.Tag
        except Exception:
            pass
        return _CATEGORY_MODE_GENERIC_HOST_REBAR

    def _family_name_from_ui(self):
        name = u""
        try:
            if self.txt_name is not None:
                name = _as_unicode(self.txt_name.Text)
        except Exception:
            name = u""
        return _sanitize_family_name(name)

    def _update_inplace_label(self, elem):
        if self.txt_inplace is None:
            return
        if elem is None:
            try:
                self.txt_inplace.Text = u"Ninguno seleccionado."
            except Exception:
                pass
            return
        label = u"Model In-Place"
        try:
            cat = elem.Category
            if cat is not None:
                label = _as_unicode(cat.Name)
        except Exception:
            pass
        try:
            name = _as_unicode(elem.Name)
            if name:
                label = u"{0}: {1}".format(label, name)
        except Exception:
            pass
        try:
            label = u"{0} (Id {1})".format(label, int(elem.Id.Value))
        except Exception:
            try:
                label = u"{0} (Id {1})".format(label, int(elem.Id.IntegerValue))
            except Exception:
                pass
        try:
            self.txt_inplace.Text = label
        except Exception:
            pass

        # Nombre familia por defecto
        default_name = u"Escalera_InPlace"
        try:
            fam = elem.Symbol.Family
            raw = _as_unicode(fam.Name).strip()
            if raw:
                default_name = raw
        except Exception:
            pass
        try:
            if self.txt_name is not None:
                current = _as_unicode(self.txt_name.Text).strip()
                if (not current) or current.startswith(u"Escalera_"):
                    self.txt_name.Text = _sanitize_family_name(default_name)
        except Exception:
            pass

    def request_pick(self):
        if self._busy:
            return
        self.handler.ctrl = self
        self.handler.action = u"pick"
        self.hide_window()
        self.ext_event.Raise()

    def request_convert(self):
        if self._busy:
            return
        if self.inplace_id is None:
            _mostrar_aviso(
                self.revit,
                u"Seleccione un Model In-Place antes de convertir.",
            )
            return
        name = self._family_name_from_ui()
        if not name:
            _mostrar_aviso(self.revit, u"Indique un nombre de familia.")
            return

        category = self._selected_category()
        inplace_id = self.inplace_id
        self._busy = True

        _set_pending_convert(inplace_id, category, name)
        self.handler.ctrl = None
        self.handler.action = u"convert"
        self.handler.convert_inplace_id = inplace_id
        self.handler.convert_category = category
        self.handler.convert_family_name = name

        # Principal: ExternalEvent fresco (force_new al abrir la herramienta)
        try:
            self.ext_event.Raise()
        except Exception:
            pass

        # Respaldo: Idling (no abortar si falla el armado)
        try:
            _arm_idling_convert(self.revit)
        except Exception:
            pass

        try:
            self.win.Close()
        except Exception:
            try:
                self.hide_window()
            except Exception:
                pass

    def execute_pick(self, uiapp):
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            _mostrar_aviso(uiapp, u"No hay documento activo.")
            self.show_window()
            return
        doc = uidoc.Document
        if doc.IsFamilyDocument:
            _mostrar_aviso(
                uiapp,
                u"Abra un proyecto (no un documento de familia).",
            )
            self.show_window()
            return
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element,
                _InPlaceSelectionFilter(),
                u"Seleccione un Model In-Place.",
            )
        except Exception:
            self.set_status(u"Selección cancelada.")
            self.show_window()
            return

        elem = doc.GetElement(ref.ElementId)
        if not _is_model_in_place(elem):
            _mostrar_aviso(
                uiapp,
                u"El elemento seleccionado no es un Model In-Place.",
            )
            self.show_window()
            return

        self.inplace_id = elem.Id
        self._update_inplace_label(elem)
        self.set_status(u"Model In-Place listo. Pulse Convertir.")
        self.show_window()


def run(revit):
    """Punto de entrada pyRevit."""
    existing = _get_active_window()
    if existing is not None:
        _activate_existing(existing, revit)
        return

    uidoc = None
    try:
        uidoc = revit.ActiveUIDocument
    except Exception:
        uidoc = None
    if uidoc is None:
        _mostrar_aviso(revit, u"No hay documento activo.")
        return
    try:
        if uidoc.Document.IsFamilyDocument:
            _mostrar_aviso(
                revit,
                u"Abra un proyecto (no un documento de familia).",
            )
            return
    except Exception:
        pass

    xaml = _build_xaml()
    win = XamlReader.Parse(xaml)
    _prepare_window(win, revit)

    ext_event, handler = _get_or_create_external_event(force_new=True)
    ctrl = _ToolController(revit, win, ext_event, handler)

    def _on_pick(sender, args):
        ctrl.request_pick()

    def _on_convert(sender, args):
        ctrl.request_convert()

    def _on_manual(sender, args):
        _open_manual(revit)

    def _on_close(sender, args):
        try:
            win.Close()
        except Exception:
            pass

    def _on_closed(sender, args):
        _clear_active_window()
        try:
            handler.ctrl = None
        except Exception:
            pass
        # No tocar action ni pending convert: Raise ya quedó encolado.

    def _on_key(sender, args):
        if args.Key == Key.Escape:
            try:
                win.Close()
            except Exception:
                pass
            args.Handled = True

    btn_pick = win.FindName(u"BtnPick")
    btn_convert = win.FindName(u"BtnConvert")
    btn_manual = win.FindName(u"BtnManual")
    btn_close = win.FindName(u"BtnClose")

    if btn_pick is not None:
        btn_pick.Click += RoutedEventHandler(_on_pick)
    if btn_convert is not None:
        btn_convert.Click += RoutedEventHandler(_on_convert)
    if btn_manual is not None:
        btn_manual.Click += RoutedEventHandler(_on_manual)
    if btn_close is not None:
        btn_close.Click += RoutedEventHandler(_on_close)

    win.Closed += EventHandler(_on_closed)
    win.KeyDown += KeyEventHandler(_on_key)

    _set_active_window(win)
    try:
        win.Show()
    except Exception as ex:
        _clear_active_window()
        _mostrar_aviso(
            revit,
            u"No se pudo abrir el formulario.",
            content=_as_unicode(ex),
        )
