# -*- coding: utf-8 -*-
"""
Dividir y Traslapar — core (traslape por diámetro, tabla BIMTools).

Camino geométrico 2024+: proyectar punto → partir centerline → extender ±L/2
→ CreateFromCurves ×2 → borrar original.

Revit 2024+ | pyRevit | IronPython
"""

from __future__ import print_function

import clr
import os
import tempfile

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

import System
from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BooleanOperationsType,
    BooleanOperationsUtils,
    BuiltInCategory,
    BuiltInFailures,
    CurveLoop,
    ElementId,
    FailureProcessingResult,
    FailureSeverity,
    FilteredElementCollector,
    GeometryInstance,
    GeometryCreationUtilities,
    IFailuresPreprocessor,
    Line,
    Mesh,
    Options,
    Plane,
    PlanarFace,
    SketchPlane,
    Solid,
    StorageType,
    Transaction,
    TransactionGroup,
    UnitTypeId,
    UnitUtils,
    UV,
    View,
    ViewDetailLevel,
    ViewSchedule,
    ViewSection,
    ViewSheet,
    ViewType,
    XYZ,
)
from Autodesk.Revit.DB.Structure import (
    MultiplanarOption,
    Rebar,
    RebarBarType,
    RebarHookOrientation,
    RebarPresentationMode,
    RebarStyle,
)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

from bimtools_rebar_hook_lengths import traslape_mm_from_nominal_diameter_mm
from dividir_rebar_punto_geom import (
    build_plan_polyline_from_frame_mm,
    build_spans_mm,
    ft_to_mm,
    mm_to_internal,
    normalize_lap_mode,
    piece_intervals_with_lap,
    project_xyz_mm_to_uv,
    set_span_length_mm,
    split_distances_with_lap,
    validate_cut_with_lap,
    validate_cuts_with_lap,
)

_TRANSACTION_NAME = u"Arainco: Dividir y Traslapar"
_PREVIEW_TG_NAME = u"Arainco: Vista previa cortes (temporal)"
_PREVIEW_TX_NAME = u"Arainco: Marca temporal de corte"
_MIN_PIECE_MM = 100.0
_ARMADURA_CONJUNTO_GUID_PARAM = u"Armadura_Conjunto_GUID"
# Prisma de sonda ~20×20×20 mm centrado en el startpoint del tramo.
_HOST_PROBE_HALF_MM = 10.0
_HOST_PROBE_HEIGHT_MM = 20.0
_PROGRESS_ACCENT_RGB = (91, 192, 222)
_PROGRESS_PHASES = 7


def _pbar_enabled():
    try:
        from pyrevit import forms as _forms  # noqa: F401
    except Exception:
        return False
    return True


class DividirRebarProgress(object):
    """ProgressBar pyRevit (acento BIMTools); no-op si no está disponible."""

    def __init__(self, total=None, title_prefix=None):
        self._total = max(1, int(total or _PROGRESS_PHASES))
        self._index = 0
        self._pb = None
        self._open = False
        self._title_prefix = title_prefix or _TRANSACTION_NAME

    def __enter__(self):
        if not _pbar_enabled() or self._total < 1:
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

    def step(self, phase_label):
        """Avanza una fase y actualiza título de la barra."""
        if self._pb is None:
            return
        i = int(self._index)
        self._index = i + 1
        base = u"{0} — {1}".format(self._title(i), phase_label or u"")
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


def _progress_step(progress, label):
    if progress is None:
        return
    try:
        progress.step(label)
    except Exception:
        pass
_HOST_INTERSECT_VOL_TOL_FT3 = 1e-12
_OUTSIDE_HOST_FAILURE_IDS = None
_MARKER_PNG_CACHE = {}
_CROSS_HALF_MM = 120.0
_LAP_TICK_MM = 40.0

_CATS_CONCRETE_IN_VIEW = (
    BuiltInCategory.OST_Walls,  # muros
    BuiltInCategory.OST_StructuralFraming,  # vigas
    BuiltInCategory.OST_StructuralFoundation,  # fundaciones
    BuiltInCategory.OST_StructuralColumns,  # columnas
    BuiltInCategory.OST_Floors,  # losas
)
_CONTEXT_MAX_ELEMENTS = 250
_CONTEXT_MIN_EDGE_MM = 1.0
# Cara de corte en sección/alzado: normal casi paralela a ViewDirection.
_SECTION_CUT_FACE_DOT_MIN = 0.65


# ---------------------------------------------------------------------------
# Warning silencioso: rebar outside of its host
# ---------------------------------------------------------------------------


def _outside_host_failure_ids():
    """FailureDefinitionId de rebar/container fuera del host (según versión API)."""
    global _OUTSIDE_HOST_FAILURE_IDS
    if _OUTSIDE_HOST_FAILURE_IDS is not None:
        return _OUTSIDE_HOST_FAILURE_IDS
    ids = []
    try:
        rf = BuiltInFailures.RebarFailures
    except Exception:
        rf = None
    if rf is not None:
        for attr in (
            u"OutSideOfHost",
            u"RebarOutSideOfHost",
            u"RebarContainerOutSideOfHostWarning",
        ):
            try:
                fid = getattr(rf, attr, None)
                if fid is not None:
                    ids.append(fid)
            except Exception:
                pass
    _OUTSIDE_HOST_FAILURE_IDS = ids
    return _OUTSIDE_HOST_FAILURE_IDS


def _failure_message_looks_outside_host(fmsg):
    """Respaldo por texto si el FailureDefinitionId no está en la API."""
    try:
        desc = fmsg.GetDescriptionText() or u""
    except Exception:
        return False
    try:
        low = _as_unicode(desc).lower()
    except Exception:
        return False
    markers = (
        u"completamente fuera",
        u"completely outside",
        u"outside of its host",
        u"fuera de su anfitrión",
        u"fuera de su host",
        u"outside of the host",
    )
    for m in markers:
        if m in low:
            return True
    return False


class _RebarOutsideHostWarningSwallower(IFailuresPreprocessor):
    """Elimina warnings de rebar fuera del host; no toca errores ni otros avisos."""

    def _iter_failure_msgs(self, failures_accessor):
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

    def PreprocessFailures(self, failures_accessor):
        if failures_accessor is None:
            return FailureProcessingResult.Continue
        known = _outside_host_failure_ids()
        for f in self._iter_failure_msgs(failures_accessor):
            try:
                if f.GetSeverity() != FailureSeverity.Warning:
                    continue
            except Exception:
                continue
            delete = False
            try:
                fid = f.GetFailureDefinitionId()
            except Exception:
                fid = None
            if fid is not None and known:
                for kid in known:
                    try:
                        if fid == kid:
                            delete = True
                            break
                    except Exception:
                        pass
            if not delete:
                delete = _failure_message_looks_outside_host(f)
            if delete:
                try:
                    failures_accessor.DeleteWarning(f)
                except Exception:
                    pass
        return FailureProcessingResult.Continue


def _attach_rebar_outside_host_swallower(txn):
    """Adjunta el preprocessor a una ``Transaction`` (antes de Start)."""
    if txn is None or not isinstance(txn, Transaction):
        return False
    try:
        opts = txn.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(_RebarOutsideHostWarningSwallower())
        txn.SetFailureHandlingOptions(opts)
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------


def internal_to_mm(ft):
    try:
        return float(UnitUtils.ConvertFromInternalUnits(float(ft), UnitTypeId.Millimeters))
    except Exception:
        return ft_to_mm(ft)


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _exception_text(ex):
    try:
        return _as_unicode(ex)
    except Exception:
        return u"Error desconocido."


def _element_id_int(eid):
    if eid is None:
        return 0
    try:
        return int(eid.IntegerValue)
    except Exception:
        try:
            return int(eid.Value)
        except Exception:
            return 0


# ---------------------------------------------------------------------------
# Selección / lectura
# ---------------------------------------------------------------------------


class _RebarSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, Rebar)

    def AllowReference(self, reference, point):
        return False


def pick_rebar(uidoc):
    if uidoc is None:
        return None, u"No hay documento activo."
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            _RebarSelectionFilter(),
            u"Seleccione una armadura estructural (Rebar).",
        )
    except Exception as ex:
        msg = _exception_text(ex)
        if u"cancel" in msg.lower() or u"OperationCanceled" in msg:
            return None, u"Selección cancelada."
        return None, msg
    if ref is None:
        return None, u"Selección cancelada."
    el = uidoc.Document.GetElement(ref.ElementId)
    if not isinstance(el, Rebar):
        return None, u"El elemento no es una Rebar."
    return el, None


class _RebarIdSelectionFilter(ISelectionFilter):
    """Solo permite la Rebar cuyo Id coincide (para clic de corte)."""

    def __init__(self, rebar_id):
        self._id_int = _element_id_int(rebar_id)

    def AllowElement(self, element):
        if not isinstance(element, Rebar):
            return False
        return _element_id_int(element.Id) == self._id_int

    def AllowReference(self, reference, point):
        return True


class _RebarPickFilterWithPreview(ISelectionFilter):
    """Filtro de la barra + preview de hover mientras dura PickObjects."""

    def __init__(self, rebar_id, preview):
        self._id_int = _element_id_int(rebar_id)
        self._preview = preview

    def AllowElement(self, element):
        if not isinstance(element, Rebar):
            return False
        return _element_id_int(element.Id) == self._id_int

    def AllowReference(self, reference, point):
        if self._preview is not None and point is not None:
            try:
                self._preview.update_hover(point)
            except Exception:
                pass
        return True


def _is_selection_cancelled(ex):
    """True si el usuario canceló / abortó el pick (Esc, Cancelar, etc.)."""
    try:
        import Autodesk.Revit.Exceptions as _RvtEx

        if isinstance(ex, _RvtEx.OperationCanceledException):
            return True
    except Exception:
        pass
    msg = _exception_text(ex).lower()
    return (
        u"cancel" in msg
        or u"abort" in msg
        or u"operationcanceled" in msg
        or u"cancelad" in msg
        or u"abortad" in msg
    )


def _object_type_point_on_element():
    try:
        return getattr(ObjectType, u"PointOnElement", None)
    except Exception:
        return None


def _global_point_from_ref(ref):
    if ref is None:
        return None
    try:
        gp = ref.GlobalPoint
        if gp is not None:
            return gp
    except Exception:
        pass
    return None


def _marker_png_path(kind=u"cut"):
    """PNG en %TEMP% para TemporaryGraphicsManager (cruz / extremo traslape)."""
    global _MARKER_PNG_CACHE
    key = _as_unicode(kind or u"cut")
    cached = _MARKER_PNG_CACHE.get(key)
    if cached and os.path.isfile(cached):
        return cached
    try:
        clr.AddReference(u"System.Drawing")
        from System.Drawing import Bitmap, Color, Pen, SolidBrush
        from System.Drawing.Drawing2D import SmoothingMode
        from System.Drawing.Imaging import ImageFormat
    except Exception:
        return None

    size = 48 if key == u"cut" else 28
    path = os.path.join(
        tempfile.gettempdir(),
        u"Arainco_dividir_rebar_marker_{0}.png".format(key),
    )
    try:
        bmp = Bitmap(size, size)
        g = System.Drawing.Graphics.FromImage(bmp)
        g.SmoothingMode = SmoothingMode.AntiAlias
        g.Clear(Color.FromArgb(0, 0, 0, 0))
        if key == u"cut":
            fill = SolidBrush(Color.FromArgb(230, 251, 191, 36))
            edge = Pen(Color.FromArgb(255, 185, 28, 28), 2.5)
            cross = Pen(Color.FromArgb(255, 15, 23, 42), 3.0)
        else:
            fill = SolidBrush(Color.FromArgb(200, 56, 189, 248))
            edge = Pen(Color.FromArgb(255, 14, 116, 144), 1.5)
            cross = None
        m = 3 if key == u"cut" else 2
        g.FillEllipse(fill, m, m, size - 2 * m, size - 2 * m)
        g.DrawEllipse(edge, m, m, size - 2 * m, size - 2 * m)
        if cross is not None:
            c = size / 2.0
            r = size / 4.5
            g.DrawLine(cross, c - r, c - r, c + r, c + r)
            g.DrawLine(cross, c - r, c + r, c + r, c - r)
        g.Dispose()
        bmp.Save(path, ImageFormat.Png)
        bmp.Dispose()
        _MARKER_PNG_CACHE[key] = path
        return path
    except Exception:
        return None


def _ortho_normal_to(tangent):
    if tangent is None:
        return XYZ.BasisX
    try:
        t = XYZ(tangent.X, tangent.Y, tangent.Z)
        if t.GetLength() < 1e-9:
            return XYZ.BasisX
        t = t.Normalize()
    except Exception:
        return XYZ.BasisX
    ref = XYZ.BasisZ
    try:
        if abs(t.DotProduct(ref)) > 0.92:
            ref = XYZ.BasisX
    except Exception:
        ref = XYZ.BasisX
    try:
        n = t.CrossProduct(ref)
        if n.GetLength() < 1e-9:
            n = t.CrossProduct(XYZ.BasisY)
        if n.GetLength() < 1e-9:
            return XYZ.BasisX
        return n.Normalize()
    except Exception:
        return XYZ.BasisX


def _tangent_at_distance(curves, dist):
    remaining = float(dist)
    for c in curves or []:
        leng = _curve_length(c)
        if remaining <= leng + 1e-12:
            try:
                p0, p1 = _curve_endpoints(c)
                if p0 is None or p1 is None:
                    return XYZ.BasisX
                v = p1 - p0
                if v.GetLength() < 1e-12:
                    return XYZ.BasisX
                return v.Normalize()
            except Exception:
                return XYZ.BasisX
        remaining -= leng
    if curves:
        try:
            p0, p1 = _curve_endpoints(curves[-1])
            if p0 is not None and p1 is not None:
                v = p1 - p0
                if v.GetLength() > 1e-12:
                    return v.Normalize()
        except Exception:
            pass
    return XYZ.BasisX


class _CutPickPreview(object):
    """
    Marcas temporales en la vista mientras el usuario define cortes.

    Preferencia: TemporaryGraphicsManager (sin undo).
    Respaldo: ModelCurve en TransactionGroup con RollBack al terminar.
    """

    def __init__(self, doc, view, curves=None, lap_mm=None):
        self._doc = doc
        self._view = view
        self._curves = list(curves or [])
        self._lap_ft = None
        if lap_mm is not None:
            try:
                lv = float(lap_mm)
                if lv > 0:
                    self._lap_ft = mm_to_internal(lv)
            except Exception:
                self._lap_ft = None
        self._tg = None
        self._tg_indices = []
        self._hover_indices = []
        self._hover_last = None
        self._hover_min_ft = mm_to_internal(40.0)
        self._tg_ok = False
        self._model_ids = []
        self._tg_group = None
        self._use_model = False
        self._total_ft = 0.0
        for c in self._curves:
            self._total_ft += float(_curve_length(c) or 0.0)
        try:
            from Autodesk.Revit.DB import TemporaryGraphicsManager

            self._tg = TemporaryGraphicsManager.GetTemporaryGraphicsManager(doc)
            self._tg_ok = self._tg is not None
        except Exception:
            self._tg = None
            self._tg_ok = False
        if not self._tg_ok:
            self._start_model_group()

    def _start_model_group(self):
        if self._doc is None or self._use_model:
            return
        try:
            self._tg_group = TransactionGroup(self._doc, _PREVIEW_TG_NAME)
            self._tg_group.Start()
            self._use_model = True
        except Exception:
            self._tg_group = None
            self._use_model = False

    def _view_id_for_tg(self):
        try:
            if self._view is not None:
                return self._view.Id
        except Exception:
            pass
        try:
            return ElementId.InvalidElementId
        except Exception:
            return ElementId(-1)

    def _add_tmp_icon(self, xyz, kind=u"cut"):
        if not self._tg_ok or xyz is None:
            return False
        path = _marker_png_path(kind)
        if not path:
            return False
        try:
            from Autodesk.Revit.DB import InCanvasControlData

            data = InCanvasControlData(path, xyz)
            idx = self._tg.AddControl(data, self._view_id_for_tg())
            self._tg_indices.append(int(idx))
            return True
        except Exception:
            return False

    def _add_model_cross(self, xyz, tangent):
        if self._doc is None or xyz is None:
            return False
        if not self._use_model:
            self._start_model_group()
        if not self._use_model:
            return False
        half = mm_to_internal(_CROSS_HALF_MM)
        n = _ortho_normal_to(tangent)
        try:
            b = tangent.Normalize().CrossProduct(n).Normalize()
        except Exception:
            b = XYZ.BasisY
        try:
            plane = Plane.CreateByOriginAndBasis(xyz, n, b)
            sp = SketchPlane.Create(self._doc, plane)
        except Exception:
            try:
                plane = Plane.CreateByNormalAndOrigin(n, xyz)
                sp = SketchPlane.Create(self._doc, plane)
            except Exception:
                return False
        ok_any = False
        t = Transaction(self._doc, _PREVIEW_TX_NAME)
        try:
            t.Start()
            for direction in (n, b):
                try:
                    p0 = xyz - direction.Multiply(half)
                    p1 = xyz + direction.Multiply(half)
                    if p0.DistanceTo(p1) < 1e-9:
                        continue
                    ln = Line.CreateBound(p0, p1)
                    mc = self._doc.Create.NewModelCurve(ln, sp)
                    if mc is not None:
                        self._model_ids.append(mc.Id)
                        ok_any = True
                except Exception:
                    continue
            if self._lap_ft and self._lap_ft > 0 and tangent is not None:
                try:
                    half_lap = 0.5 * float(self._lap_ft)
                    tick = mm_to_internal(_LAP_TICK_MM)
                    tdir = tangent.Normalize()
                    for sign in (-1.0, 1.0):
                        mid = xyz + tdir.Multiply(sign * half_lap)
                        q0 = mid - n.Multiply(tick)
                        q1 = mid + n.Multiply(tick)
                        ln2 = Line.CreateBound(q0, q1)
                        mc2 = self._doc.Create.NewModelCurve(ln2, sp)
                        if mc2 is not None:
                            self._model_ids.append(mc2.Id)
                            ok_any = True
                except Exception:
                    pass
            t.Commit()
        except Exception:
            try:
                if t.HasStarted():
                    t.RollBack()
            except Exception:
                pass
            return False
        return ok_any

    def add_projected(self, point_on, cut_dist_ft, tangent=None):
        if point_on is None:
            return
        tan = tangent or _tangent_at_distance(self._curves, cut_dist_ft)
        drew = self._add_tmp_icon(point_on, u"cut")
        if self._lap_ft and self._curves and self._total_ft > 0:
            half = 0.5 * float(self._lap_ft)
            d0 = max(0.0, float(cut_dist_ft) - half)
            d1 = min(self._total_ft, float(cut_dist_ft) + half)
            p0 = _point_at_distance(self._curves, d0)
            p1 = _point_at_distance(self._curves, d1)
            if p0 is not None:
                drew = self._add_tmp_icon(p0, u"lap") or drew
            if p1 is not None:
                drew = self._add_tmp_icon(p1, u"lap") or drew
        if not drew:
            self._add_model_cross(point_on, tan)

    def add_from_pick(self, pick_xyz):
        if pick_xyz is None or not self._curves:
            if pick_xyz is not None:
                self.add_projected(pick_xyz, 0.0, XYZ.BasisX)
            return
        proj = project_point_on_polyline(self._curves, pick_xyz)
        if not proj.get(u"ok"):
            self.add_projected(pick_xyz, 0.0, XYZ.BasisX)
            return
        pt = proj.get(u"point_on") or pick_xyz
        dist = float(proj.get(u"cut_dist") or 0.0)
        tan = _tangent_at_distance(self._curves, dist)
        self.add_projected(pt, dist, tan)

    def add_from_cut_mm(self, cut_mm):
        if not self._curves:
            return
        try:
            dist = mm_to_internal(float(cut_mm))
        except Exception:
            return
        pt = _point_at_distance(self._curves, dist)
        if pt is None:
            return
        tan = _tangent_at_distance(self._curves, dist)
        self.add_projected(pt, dist, tan)

    def _clear_hover(self):
        if self._tg is None:
            self._hover_indices = []
            self._hover_last = None
            return
        for idx in list(self._hover_indices):
            try:
                self._tg.RemoveControl(int(idx))
            except Exception:
                pass
        self._hover_indices = []
        self._hover_last = None

    def update_hover(self, pick_xyz):
        """Marca temporal que sigue el cursor sobre la barra (durante PickObjects)."""
        if pick_xyz is None or not self._tg_ok:
            return
        pt = pick_xyz
        if self._curves:
            try:
                proj = project_point_on_polyline(self._curves, pick_xyz)
                if proj.get(u"ok") and proj.get(u"point_on") is not None:
                    pt = proj[u"point_on"]
            except Exception:
                pt = pick_xyz
        if self._hover_last is not None:
            try:
                if self._hover_last.DistanceTo(pt) < self._hover_min_ft:
                    return
            except Exception:
                pass
        self._clear_hover()
        path = _marker_png_path(u"cut")
        if not path:
            return
        try:
            from Autodesk.Revit.DB import InCanvasControlData

            data = InCanvasControlData(path, pt)
            idx = self._tg.AddControl(data, self._view_id_for_tg())
            self._hover_indices.append(int(idx))
            self._hover_last = pt
        except Exception:
            pass

    def clear(self):
        self._clear_hover()
        if self._tg is not None and self._tg_indices:
            for idx in list(self._tg_indices):
                try:
                    self._tg.RemoveControl(int(idx))
                except Exception:
                    pass
            self._tg_indices = []
        if self._tg_group is not None:
            try:
                self._tg_group.RollBack()
            except Exception:
                try:
                    self._tg_group.Assimilate()
                except Exception:
                    pass
            self._tg_group = None
            self._model_ids = []
            self._use_model = False


def pick_cut_point_on_rebar(uidoc, rebar, prompt=None):
    """
    Punto de corte: un solo clic sobre la barra (no PickPoint genérico).

    Un único ``PickObject`` para no encadenar varios prompts (fluidez).
    Preferir ``PointOnElement`` si existe; si no, ``Element`` + ``GlobalPoint``.
    """
    if uidoc is None:
        return None, u"No hay documento activo."
    if not isinstance(rebar, Rebar):
        return None, u"Barra inválida."

    prompt = prompt or (
        u"Clic sobre la barra en el punto de corte (Esc cancela)."
    )
    filt = _RebarIdSelectionFilter(rebar.Id)

    ot = _object_type_point_on_element()
    if ot is None:
        ot = ObjectType.Element

    try:
        ref = uidoc.Selection.PickObject(ot, filt, prompt)
    except Exception as ex:
        msg = _exception_text(ex)
        if _is_selection_cancelled(ex):
            return None, u"Selección cancelada."
        if ot != ObjectType.Element:
            try:
                ref = uidoc.Selection.PickObject(ObjectType.Element, filt, prompt)
            except Exception as ex2:
                if _is_selection_cancelled(ex2):
                    return None, u"Selección cancelada."
                return None, _exception_text(ex2)
        else:
            return None, msg

    pt = _global_point_from_ref(ref)
    if pt is not None:
        return pt, None
    return (
        None,
        u"No se obtuvo el punto de clic. Haga clic directamente sobre la barra "
        u"visible en la vista activa.",
    )


def pick_cut_points_on_rebar(
    uidoc, rebar, prompt=None, lap_mm=None, existing_cuts_mm=None
):
    """
    Varios puntos de corte sobre la barra, con marcas temporales en la vista.

    Usa ``PickObjects`` para mostrar **Finalizar** / **Cancelar** en la barra
    de opciones. Mientras se elige, una marca sigue el cursor sobre la barra;
    los cortes ya existentes también se dibujan.

    Returns:
        (lista_XYZ, None) o (None, mensaje_error). Lista vacía si Finalizar sin puntos.
    """
    if uidoc is None:
        return None, u"No hay documento activo."
    if not isinstance(rebar, Rebar):
        return None, u"Barra inválida."

    doc = uidoc.Document
    view = None
    try:
        view = uidoc.ActiveView
    except Exception:
        view = None

    curves = _centerline_curves(rebar, 0, True, True)
    if not curves:
        curves = _centerline_curves(rebar, 0, True, False)

    preview = _CutPickPreview(doc, view, curves=curves, lap_mm=lap_mm)
    for mm in existing_cuts_mm or []:
        try:
            preview.add_from_cut_mm(mm)
        except Exception:
            pass
    try:
        uidoc.RefreshActiveView()
    except Exception:
        pass

    prompt = prompt or (
        u"Clic sobre la barra en cada punto de corte. "
        u"Pulse Finalizar en la barra de opciones (arriba o abajo) cuando termine."
    )
    filt = _RebarPickFilterWithPreview(rebar.Id, preview)
    ot = _object_type_point_on_element()
    if ot is None:
        ot = ObjectType.Element

    refs = None
    try:
        try:
            refs = uidoc.Selection.PickObjects(ot, filt, prompt)
        except Exception as ex:
            if _is_selection_cancelled(ex):
                return None, u"Selección cancelada."
            if ot != ObjectType.Element:
                try:
                    refs = uidoc.Selection.PickObjects(
                        ObjectType.Element, filt, prompt
                    )
                except Exception as ex2:
                    if _is_selection_cancelled(ex2):
                        return None, u"Selección cancelada."
                    return None, _exception_text(ex2)
            else:
                return None, _exception_text(ex)

        points = []
        if refs is not None:
            for ref in refs:
                pt = _global_point_from_ref(ref)
                if pt is not None:
                    points.append(pt)
                    try:
                        preview.add_from_pick(pt)
                    except Exception:
                        pass
        try:
            uidoc.RefreshActiveView()
        except Exception:
            pass
        return points, None
    finally:
        try:
            preview.clear()
        except Exception:
            pass
        try:
            uidoc.RefreshActiveView()
        except Exception:
            pass


def pick_point(uidoc, prompt=None):
    """Compat: preferir ``pick_cut_point_on_rebar`` cuando haya Rebar."""
    if uidoc is None:
        return None, u"No hay documento activo."
    try:
        from Autodesk.Revit.UI.Selection import ObjectSnapTypes

        pt = uidoc.Selection.PickPoint(
            ObjectSnapTypes.Endpoints
            | ObjectSnapTypes.Nearest
            | ObjectSnapTypes.Intersections,
            prompt or u"Indique el punto de corte sobre la barra.",
        )
    except Exception as ex:
        if _is_selection_cancelled(ex):
            return None, u"Selección cancelada."
        msg = _exception_text(ex)
        return (
            None,
            u"PickPoint no disponible (¿falta plano de trabajo?). {0}".format(msg),
        )
    if pt is None:
        return None, u"Selección cancelada."
    return pt, None


def _shape_driven_accessor(rebar):
    try:
        return rebar.GetShapeDrivenAccessor()
    except Exception:
        return None


def _call_bool_flag(obj, name):
    """
    Lee un flag bool de la API que puede ser método (Revit) o propiedad.

    Importante: ``bool(rebar.IsRebarFreeForm)`` es siempre True en IronPython
    porque el *método* es truthy; hay que invocar ``IsRebarFreeForm()``.
    """
    if obj is None or not name:
        return None
    try:
        attr = getattr(obj, name, None)
    except Exception:
        return None
    if attr is None:
        return None
    try:
        if callable(attr):
            return bool(attr())
        return bool(attr)
    except Exception:
        return None


def _is_free_form(rebar):
    """True solo si la API indica free-form (shape-driven → False)."""
    ff = _call_bool_flag(rebar, u"IsRebarFreeForm")
    if ff is True:
        return True
    if ff is False:
        return False
    sd = _call_bool_flag(rebar, u"IsRebarShapeDriven")
    if sd is True:
        return False
    if sd is False:
        return True
    # Respaldo: si hay ShapeDrivenAccessor, no es free-form.
    return _shape_driven_accessor(rebar) is None


def _cantidad_posiciones(rebar):
    best = 1
    for getter in (
        lambda: int(rebar.NumberOfBarPositions),
        lambda: int(rebar.GetNumberOfBarPositions()),
        lambda: int(rebar.Quantity),
    ):
        try:
            n = int(getter())
            if n > best:
                best = n
        except Exception:
            pass
    return best


def _layout_rule_name(rebar, acc=None):
    try:
        r = rebar.LayoutRule
        if r is not None:
            s = r.ToString() or u""
            if s:
                return s
    except Exception:
        pass
    if acc is None:
        acc = _shape_driven_accessor(rebar)
    if acc is not None:
        try:
            r = acc.GetLayoutRule()
            if r is not None:
                s = r.ToString() or u""
                if s:
                    return s
        except Exception:
            pass
    return u""


def _centerline_curves(rebar, bar_index=0, suppress_hooks=True, suppress_bends=True):
    bi = int(bar_index)
    for getter in (
        lambda: rebar.GetTransformedCenterlineCurves(
            False,
            bool(suppress_hooks),
            bool(suppress_bends),
            MultiplanarOption.IncludeAllMultiplanarCurves,
            bi,
        ),
        lambda: rebar.GetCenterlineCurves(
            False,
            bool(suppress_hooks),
            bool(suppress_bends),
            MultiplanarOption.IncludeAllMultiplanarCurves,
            bi,
        ),
    ):
        try:
            curves = getter()
            if curves is not None and curves.Count > 0:
                return [curves[i] for i in range(curves.Count)]
        except Exception:
            pass
    return []


def _clone_curve(crv):
    if crv is None:
        return None
    try:
        return crv.Clone()
    except Exception:
        return crv


def _curve_endpoints(crv):
    try:
        return crv.GetEndPoint(0), crv.GetEndPoint(1)
    except Exception:
        return None, None


def _curve_length(crv):
    try:
        return float(crv.Length)
    except Exception:
        p0, p1 = _curve_endpoints(crv)
        if p0 is None or p1 is None:
            return 0.0
        return p0.DistanceTo(p1)


# ---------------------------------------------------------------------------
# Host por Armadura_Conjunto_GUID + colisión del startpoint
# ---------------------------------------------------------------------------


def _get_armadura_conjunto_guid(element):
    """Lee ``Armadura_Conjunto_GUID`` o ``None``."""
    if element is None:
        return None
    p = None
    try:
        p = element.LookupParameter(_ARMADURA_CONJUNTO_GUID_PARAM)
    except Exception:
        p = None
    if p is None:
        try:
            target = _ARMADURA_CONJUNTO_GUID_PARAM.lower()
            for q in element.Parameters:
                try:
                    dn = _as_unicode(q.Definition.Name).replace(u"\u00A0", u" ").strip()
                except Exception:
                    continue
                if dn.lower() == target:
                    p = q
                    break
        except Exception:
            p = None
    if p is None:
        return None
    try:
        if p.StorageType == StorageType.String:
            val = p.AsString()
        else:
            val = p.AsValueString()
    except Exception:
        val = None
    if val is None:
        return None
    try:
        s = _as_unicode(val).strip()
    except Exception:
        s = u""
    return s or None


def _geometry_options_fine():
    opts = Options()
    try:
        opts.ComputeReferences = False
    except Exception:
        pass
    try:
        opts.DetailLevel = ViewDetailLevel.Fine
    except Exception:
        pass
    return opts


def _iter_element_solids(elem):
    """Sólidos con volumen del elemento (incluye GeometryInstance)."""
    if elem is None:
        return
    try:
        geo_elem = elem.get_Geometry(_geometry_options_fine())
    except Exception:
        return
    if geo_elem is None:
        return

    def _vol_ok(solid):
        if solid is None or not isinstance(solid, Solid):
            return False
        try:
            return float(solid.Volume) > 1e-12
        except Exception:
            return False

    items = []
    try:
        for g in geo_elem:
            items.append(g)
    except Exception:
        try:
            n = int(geo_elem.Size)
        except Exception:
            try:
                n = int(geo_elem.Count)
            except Exception:
                n = 0
        for i in range(n):
            try:
                items.append(geo_elem[i])
            except Exception:
                pass

    for g in items:
        if g is None:
            continue
        if _vol_ok(g):
            yield g
        elif isinstance(g, GeometryInstance):
            try:
                inst = g.GetInstanceGeometry()
            except Exception:
                inst = None
            if inst is None:
                continue
            try:
                for sg in inst:
                    if _vol_ok(sg):
                        yield sg
            except Exception:
                pass


def _build_probe_prism_at_point(xyz):
    """Prisma vertical pequeño centrado en ``xyz`` para probar colisión con hosts."""
    if xyz is None:
        return None
    half = mm_to_internal(_HOST_PROBE_HALF_MM)
    hgt = mm_to_internal(_HOST_PROBE_HEIGHT_MM)
    if half <= 0 or hgt <= 0:
        return None
    z0 = float(xyz.Z) - 0.5 * hgt
    px = float(xyz.X)
    py = float(xyz.Y)
    p1 = XYZ(px - half, py - half, z0)
    p2 = XYZ(px + half, py - half, z0)
    p3 = XYZ(px + half, py + half, z0)
    p4 = XYZ(px - half, py + half, z0)
    try:
        loop = CurveLoop.Create(
            [
                Line.CreateBound(p1, p2),
                Line.CreateBound(p2, p3),
                Line.CreateBound(p3, p4),
                Line.CreateBound(p4, p1),
            ]
        )
    except Exception:
        return None
    try:
        sol = GeometryCreationUtilities.CreateExtrusionGeometry(
            [loop], XYZ.BasisZ, hgt
        )
    except Exception:
        return None
    if sol is None:
        return None
    try:
        if float(sol.Volume) < 1e-15:
            return None
    except Exception:
        return None
    return sol


def _solids_intersect_volume(solid_a, solid_b):
    if solid_a is None or solid_b is None:
        return False
    try:
        if float(solid_a.Volume) <= 1e-12 or float(solid_b.Volume) <= 1e-12:
            return False
    except Exception:
        return False
    try:
        inter = BooleanOperationsUtils.ExecuteBooleanOperation(
            solid_a, solid_b, BooleanOperationsType.Intersect
        )
    except Exception:
        return False
    if inter is None:
        return False
    try:
        return float(inter.Volume) > _HOST_INTERSECT_VOL_TOL_FT3
    except Exception:
        return False


def _bbox_contains_point(host, xyz, pad_ft=None):
    if host is None or xyz is None:
        return False
    try:
        bb = host.get_BoundingBox(None)
    except Exception:
        bb = None
    if bb is None:
        return False
    pad = float(pad_ft) if pad_ft is not None else mm_to_internal(5.0)
    try:
        return (
            float(bb.Min.X) - pad <= float(xyz.X) <= float(bb.Max.X) + pad
            and float(bb.Min.Y) - pad <= float(xyz.Y) <= float(bb.Max.Y) + pad
            and float(bb.Min.Z) - pad <= float(xyz.Z) <= float(bb.Max.Z) + pad
        )
    except Exception:
        return False


def _host_contains_point(host, xyz, probe_solid=None):
    """
    True si el startpoint colisiona con el volumen del host.

    Primero bbox (rápido); luego intersección del prisma sonda con sólidos.
    """
    if host is None or xyz is None:
        return False
    if not _bbox_contains_point(host, xyz):
        return False
    probe = probe_solid
    if probe is None:
        probe = _build_probe_prism_at_point(xyz)
    if probe is None:
        # Sin prisma: bbox basta como aproximación débil
        return True
    for sd in _iter_element_solids(host):
        if _solids_intersect_volume(probe, sd):
            return True
    return False


def _collect_rebars_por_conjunto_guid(doc, guid):
    """
    Todas las ``Rebar`` del documento con el mismo ``Armadura_Conjunto_GUID``.

    Incluye la barra original si comparte ese GUID.
    """
    target = _as_unicode(guid).strip() if guid else u""
    if not target or doc is None:
        return []
    out = []
    try:
        rebars = (
            FilteredElementCollector(doc)
            .OfClass(Rebar)
            .WhereElementIsNotElementType()
        )
    except Exception:
        return []
    for rb in rebars:
        if not isinstance(rb, Rebar):
            continue
        try:
            gid = _get_armadura_conjunto_guid(rb)
        except Exception:
            continue
        if gid and gid == target:
            out.append(rb)
    return out


def _unique_hosts_from_rebars(doc, rebars):
    """
    Hosts de las rebars dadas, sin duplicados (por ElementId).

    Orden: primera aparición al recorrer ``rebars``.
    """
    hosts = []
    seen = set()
    if doc is None:
        return hosts
    for rb in rebars or []:
        if rb is None:
            continue
        try:
            hid = rb.GetHostId()
        except Exception:
            hid = None
        if hid is None:
            continue
        try:
            key = _element_id_int(hid)
        except Exception:
            key = 0
        if key <= 0 or key in seen:
            continue
        try:
            host = doc.GetElement(hid)
        except Exception:
            host = None
        if host is None:
            continue
        try:
            if not host.IsValidObject:
                continue
        except Exception:
            pass
        seen.add(key)
        hosts.append(host)
    return hosts


def _host_candidates_from_original_rebar(doc, original_rebar):
    """
    Pipeline de hosts candidatos para tramos divididos:

    1. GUID de la barra original
    2. Todas las rebars con ese GUID
    3. Hosts de esas rebars, sin duplicados

    Returns:
        (conjunto_guid_or_None, peer_rebars, hosts_unicos)
    """
    guid = _get_armadura_conjunto_guid(original_rebar)
    if not guid:
        return None, [], []
    peers = _collect_rebars_por_conjunto_guid(doc, guid)
    hosts = _unique_hosts_from_rebars(doc, peers)
    return guid, peers, hosts


def _chain_start_point(chain):
    """Startpoint del tramo producto: extremo 0 de la primera curva."""
    if not chain:
        return None
    p0, _p1 = _curve_endpoints(chain[0])
    return p0


def _bbox_center(host):
    try:
        bb = host.get_BoundingBox(None)
        if bb is None:
            return None
        return XYZ(
            0.5 * (float(bb.Min.X) + float(bb.Max.X)),
            0.5 * (float(bb.Min.Y) + float(bb.Max.Y)),
            0.5 * (float(bb.Min.Z) + float(bb.Max.Z)),
        )
    except Exception:
        return None


def _resolve_host_for_chain(chain, candidates, fallback):
    """
    Compara el startpoint del tramo producto contra la lista depurada de hosts.

    Varios hits → el más cercano al centroide del bbox. Sin hits → ``fallback``
    (host de la barra original).
    """
    start = _chain_start_point(chain)
    if start is None:
        return fallback
    if not candidates:
        return fallback
    probe = _build_probe_prism_at_point(start)
    hits = []
    for h in candidates:
        if h is None:
            continue
        try:
            if _host_contains_point(h, start, probe_solid=probe):
                hits.append(h)
        except Exception:
            continue
    if not hits:
        return fallback
    if len(hits) == 1:
        return hits[0]
    best = None
    best_d = None
    for h in hits:
        c = _bbox_center(h)
        if c is None:
            continue
        try:
            d = float(start.DistanceTo(c))
        except Exception:
            continue
        if best_d is None or d < best_d:
            best_d = d
            best = h
    return best or hits[0] or fallback


def _is_line_curve(crv):
    return isinstance(crv, Line) or (crv is not None and crv.GetType().Name == u"Line")


def _diameter_mm(bar_type):
    if bar_type is None:
        return None
    for attr in ("BarNominalDiameter", "BarModelDiameter", "BarDiameter"):
        try:
            val = getattr(bar_type, attr)
            if val is not None:
                return internal_to_mm(float(val))
        except Exception:
            pass
    try:
        p = bar_type.LookupParameter(u"Bar Diameter")
        if p is not None:
            return internal_to_mm(float(p.AsDouble()))
    except Exception:
        pass
    return None


def _rebar_normal(rebar):
    try:
        acc = _shape_driven_accessor(rebar)
        if acc is not None:
            n = acc.Normal
            if n is not None and n.GetLength() > 1e-12:
                return n.Normalize()
    except Exception:
        pass
    return XYZ.BasisZ


def _hook_type(doc, hook_id):
    if doc is None or hook_id is None:
        return None
    try:
        if hook_id == ElementId.InvalidElementId:
            return None
    except Exception:
        pass
    try:
        return doc.GetElement(hook_id)
    except Exception:
        return None


def _is_in_group(rebar):
    try:
        gid = rebar.GroupId
        if gid is not None and gid != ElementId.InvalidElementId:
            return True
    except Exception:
        pass
    return False


def layout_label(rule_name, n_pos=1):
    """Etiqueta corta para UI / mensaje de resultado."""
    rule = rule_name or u""
    n = int(n_pos or 1)
    if rule == u"MaximumSpacing" or u"MaximumSpacing" in rule:
        if n > 1:
            return u"Maximum Spacing ×{0}".format(n)
        return u"Maximum Spacing"
    if rule in (u"Number", u"FixedNumber") or (
        n > 1 and rule not in (u"Single", u"")
    ):
        return u"Fixed Number ×{0}".format(n)
    if n > 1:
        return u"Fixed Number ×{0}".format(n)
    if rule == u"Single" or not rule:
        return u"Single"
    return rule


def check_eligibility(doc, rebar):
    """
    Shape-driven, no free-form, no group; layout Single, FixedNumber o
    MaximumSpacing; centerline solo líneas (sin arcos).

    Returns:
        (ok, mensaje, curves_o_None)
    """
    if not isinstance(rebar, Rebar):
        return False, u"El elemento no es una Rebar.", None
    if _is_in_group(rebar):
        return False, u"La barra pertenece a un Group; no se puede dividir.", None
    if _is_free_form(rebar):
        return False, u"Rebar free-form no soportada en esta versión de la herramienta.", None
    rule = _layout_rule_name(rebar)
    allowed = (
        u"Single",
        u"FixedNumber",
        u"Number",
        u"MaximumSpacing",
        u"",
    )
    if rule and rule not in allowed:
        return (
            False,
            u"Layout «{0}» no soportado. Se admiten Single, Fixed Number "
            u"y Maximum Spacing.".format(rule),
            None,
        )
    curves = _centerline_curves(rebar, 0, True, True)
    if not curves:
        curves = _centerline_curves(rebar, 0, True, False)
    if not curves:
        return False, u"No se pudo leer la centerline de la barra.", None
    for c in curves:
        if not _is_line_curve(c):
            return False, u"La barra tiene tramos curvos/arco; no soportado en el MVP.", None
    try:
        acc = _shape_driven_accessor(rebar)
        if acc is not None and hasattr(acc, u"IsMultiplanar"):
            if bool(acc.IsMultiplanar):
                return False, u"Barras multiplanares no soportadas en el MVP.", None
    except Exception:
        pass
    host = doc.GetElement(rebar.GetHostId()) if doc is not None else None
    if host is None:
        return False, u"La barra no tiene host válido.", None
    bar_type = doc.GetElement(rebar.GetTypeId()) if doc is not None else None
    if not isinstance(bar_type, RebarBarType):
        return False, u"No se pudo obtener RebarBarType.", None
    return True, u"", curves


def lap_mm_for_rebar(doc, rebar, concrete_grade=None):
    bar_type = doc.GetElement(rebar.GetTypeId())
    if not isinstance(bar_type, RebarBarType):
        return None, None
    d_mm = _diameter_mm(bar_type)
    if d_mm is None or d_mm <= 0:
        return None, None
    grade = concrete_grade
    if grade in (u"", u"base", u"BASE", u"Base"):
        grade = None
    lap = traslape_mm_from_nominal_diameter_mm(d_mm, grade)
    return d_mm, lap


def project_point_on_polyline(curves, point):
    """
    Proyecta ``point`` al punto más cercano de la polilínea de líneas.

    Returns:
        dict con keys: ok, message, cut_dist, total_len, seg_index, point_on
        o ok=False.
    """
    if not curves or point is None:
        return {u"ok": False, u"message": u"Curvas o punto inválidos."}
    best_dist = None
    best = None
    cum = 0.0
    lengths = []
    for i, c in enumerate(curves):
        leng = _curve_length(c)
        lengths.append(leng)
        try:
            res = c.Project(point)
        except Exception:
            res = None
        if res is None:
            # fallback: extremos
            p0, p1 = _curve_endpoints(c)
            if p0 is None or p1 is None:
                cum += leng
                continue
            d0 = point.DistanceTo(p0)
            d1 = point.DistanceTo(p1)
            if d0 <= d1:
                xyz, param_along = p0, 0.0
                dist = d0
            else:
                xyz, param_along = p1, leng
                dist = d1
        else:
            try:
                xyz = res.XYZPoint
            except Exception:
                xyz = None
            if xyz is None:
                cum += leng
                continue
            dist = point.DistanceTo(xyz)
            try:
                # Parameter is raw; use Distance from start for lines
                p0, _p1 = _curve_endpoints(c)
                if p0 is not None:
                    param_along = p0.DistanceTo(xyz)
                else:
                    param_along = 0.5 * leng
            except Exception:
                param_along = 0.5 * leng
        if best_dist is None or dist < best_dist:
            best_dist = dist
            best = {
                u"seg_index": i,
                u"param_along": max(0.0, min(leng, float(param_along))),
                u"point_on": xyz,
                u"cum_before": cum,
            }
        cum += leng
    total = cum
    if best is None or total <= 0:
        return {u"ok": False, u"message": u"No se pudo proyectar el punto sobre la barra."}
    cut_dist = float(best[u"cum_before"]) + float(best[u"param_along"])
    # clamp to open interval
    eps = max(1e-6, total * 1e-9)
    if cut_dist <= eps or cut_dist >= total - eps:
        return {
            u"ok": False,
            u"message": u"El punto proyecta sobre un extremo de la barra; elija un punto interior.",
        }
    return {
        u"ok": True,
        u"message": u"",
        u"cut_dist": cut_dist,
        u"total_len": total,
        u"seg_index": best[u"seg_index"],
        u"point_on": best[u"point_on"],
        u"proj_dist": best_dist,
        u"lengths": lengths,
    }


def _point_at_distance(curves, dist):
    """Punto a distancia ``dist`` a lo largo de la polilínea (pies)."""
    remaining = float(dist)
    for c in curves:
        leng = _curve_length(c)
        if remaining <= leng + 1e-12:
            p0, p1 = _curve_endpoints(c)
            if p0 is None or p1 is None:
                return None
            if leng < 1e-12:
                return p0
            t = max(0.0, min(1.0, remaining / leng))
            return p0 + (p1 - p0).Multiply(t)
        remaining -= leng
    # clamp to end
    if not curves:
        return None
    _p0, p1 = _curve_endpoints(curves[-1])
    return p1


def _subchain_by_distance(curves, d0, d1):
    """Extrae sub-polilínea entre distancias [d0, d1] como lista de Line."""
    d0 = float(d0)
    d1 = float(d1)
    if d1 - d0 < 1e-9:
        return []
    out = []
    cum = 0.0
    for c in curves:
        leng = _curve_length(c)
        seg_a = cum
        seg_b = cum + leng
        cum = seg_b
        if seg_b <= d0 + 1e-12 or seg_a >= d1 - 1e-12:
            continue
        a = max(d0, seg_a)
        b = min(d1, seg_b)
        if b - a < 1e-9:
            continue
        p0, p1 = _curve_endpoints(c)
        if p0 is None or p1 is None or leng < 1e-12:
            continue
        t0 = (a - seg_a) / leng
        t1 = (b - seg_a) / leng
        qa = p0 + (p1 - p0).Multiply(t0)
        qb = p0 + (p1 - p0).Multiply(t1)
        if qa.DistanceTo(qb) < 1e-9:
            continue
        try:
            out.append(Line.CreateBound(qa, qb))
        except Exception:
            continue
    return out


def build_split_curve_chains(curves, cut_dist, lap_ft, lap_mode=None):
    """
    Construye dos cadenas de Line con solape según ``lap_mode``.

    Returns:
        (ok, msg, curves_a, curves_b, meta)
    """
    ok_m, msg_m, chains, meta = build_multi_split_curve_chains(
        curves, [cut_dist], lap_ft, lap_mode=lap_mode
    )
    if not ok_m or not chains or len(chains) < 2:
        return False, msg_m, None, None, None
    return True, u"", chains[0], chains[1], meta


def build_multi_split_curve_chains(curves, cuts_ft, lap_ft, lap_mode=None):
    """
    N cortes → N+1 cadenas según ``lap_mode`` (default simétrico ±lap/2).

    Returns:
        (ok, msg, list_of_curve_lists, meta)
    """
    mode = normalize_lap_mode(lap_mode)
    total = 0.0
    for c in curves:
        total += _curve_length(c)
    min_piece_ft = mm_to_internal(_MIN_PIECE_MM)
    ok, msg, sorted_cuts = validate_cuts_with_lap(
        total, cuts_ft, lap_ft, min_piece_ft, lap_mode=mode
    )
    if not ok:
        return False, msg, None, None
    try:
        intervals = piece_intervals_with_lap(
            total, sorted_cuts, lap_ft, lap_mode=mode
        )
    except Exception as ex:
        return False, _exception_text(ex), None, None
    chains = []
    lens_ft = []
    for d0, d1 in intervals:
        ch = _subchain_by_distance(curves, d0, d1)
        if not ch:
            return (
                False,
                u"No se pudo generar un tramo tras el corte.",
                None,
                None,
            )
        chains.append(ch)
        lens_ft.append(max(0.0, float(d1) - float(d0)))
    meta = {
        u"total_ft": total,
        u"cuts_ft": list(sorted_cuts),
        u"lap_ft": float(lap_ft),
        u"half_ft": 0.5 * float(lap_ft),
        u"lap_mode": mode,
        u"lens_ft": lens_ft,
        u"lens_mm": [internal_to_mm(x) for x in lens_ft],
        u"cuts_mm": [internal_to_mm(x) for x in sorted_cuts],
        u"overlap_mm": internal_to_mm(lap_ft),
        u"n_pieces": len(chains),
        u"intervals_ft": [(float(a), float(b)) for a, b in intervals],
    }
    # Compat 1 corte
    if len(sorted_cuts) == 1 and len(lens_ft) >= 2:
        meta[u"cut_ft"] = sorted_cuts[0]
        meta[u"len_a_ft"] = lens_ft[0]
        meta[u"len_b_ft"] = lens_ft[1]
        meta[u"len_a_mm"] = internal_to_mm(lens_ft[0])
        meta[u"len_b_mm"] = internal_to_mm(lens_ft[1])
    return True, u"", chains, meta


def suggest_bar_orientation(curves):
    """
    'horizontal' | 'vertical' según el eje dominante de la centerline.
    """
    if not curves:
        return u"horizontal"
    dx = dy = dz = 0.0
    for c in curves:
        try:
            p0, p1 = _curve_endpoints(c)
            if p0 is None or p1 is None:
                continue
            dx += abs(float(p1.X) - float(p0.X))
            dy += abs(float(p1.Y) - float(p0.Y))
            dz += abs(float(p1.Z) - float(p0.Z))
        except Exception:
            continue
    if dz >= dx and dz >= dy:
        return u"vertical"
    return u"horizontal"


def _geometry_options_for_view(view):
    opts = Options()
    try:
        opts.ComputeReferences = False
    except Exception:
        pass
    try:
        opts.DetailLevel = ViewDetailLevel.Fine
    except Exception:
        pass
    try:
        opts.IncludeNonVisibleObjects = True
    except Exception:
        pass
    if view is not None:
        try:
            opts.View = view
        except Exception:
            pass
    return opts


def _is_section_elevation_view(view):
    """Vista 2D ortográfica tipo sección o alzado (incluye ViewSection de elevación eje)."""
    if view is None:
        return False
    try:
        if isinstance(view, ViewSection):
            return True
    except Exception:
        pass
    try:
        vt = view.ViewType
        return vt == ViewType.Section or vt == ViewType.Elevation
    except Exception:
        return False


def _texto_material_indica_hormigon(text):
    try:
        s = _as_unicode(text).strip().lower()
    except Exception:
        return False
    if not s:
        return False
    keys = (
        u"hormigon",
        u"hormigón",
        u"concrete",
        u"h°",
        u"h.",
    )
    return any(k in s for k in keys)


def _mat_o_texto_sugiere_hormigon(material, param, _document):
    if param is not None and param.HasValue:
        for attr in (u"AsString", u"AsValueString"):
            try:
                t = getattr(param, attr)()
                if t and _texto_material_indica_hormigon(t):
                    return True
            except Exception:
                pass
    if material is None:
        return False
    try:
        n = material.Name
    except Exception:
        n = None
    if n and _texto_material_indica_hormigon(n):
        return True
    try:
        from Autodesk.Revit.DB.Structure import StructuralMaterialType

        if hasattr(material, u"StructuralMaterialType"):
            if material.StructuralMaterialType == StructuralMaterialType.Concrete:
                return True
    except Exception:
        pass
    return False


def _wall_has_concrete_structure_material(wall):
    """Muro con capa estructural de hormigón (compound structure o nombre de tipo)."""
    if wall is None:
        return False
    try:
        from Autodesk.Revit.DB import MaterialFunctionAssignment, Wall

        if not isinstance(wall, Wall):
            return False
        wt = wall.WallType
        if wt is None:
            return False
        if _texto_material_indica_hormigon(wt.Name or u""):
            return True
        cs = wt.GetCompoundStructure()
        if cs is None:
            return False
        doc = wall.Document
        if doc is None:
            return False
        try:
            n_layers = int(cs.LayerCount)
        except Exception:
            n_layers = 0
        for i in range(n_layers):
            try:
                fn = cs.GetLayerFunction(i)
            except Exception:
                fn = None
            if fn != MaterialFunctionAssignment.Structure:
                continue
            try:
                mid = cs.GetMaterialId(i)
            except Exception:
                mid = None
            if mid is None or mid == ElementId.InvalidElementId:
                continue
            mat = doc.GetElement(mid)
            if mat is not None and _texto_material_indica_hormigon(mat.Name or u""):
                return True
    except Exception:
        pass
    return False


def _element_is_concrete_for_elevation_canvas(elem):
    """
    Hormigón estructural visible en alzado/sección: muros, vigas, fundaciones,
    columnas y losas (mismo criterio que contorno / Armado Vigas).
    """
    if elem is None:
        return False
    try:
        from Autodesk.Revit.DB.Structure import StructuralMaterialType

        sm = elem.StructuralMaterialType
        if sm == StructuralMaterialType.Concrete:
            return True
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import BuiltInParameter

        p = elem.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)
        if p is not None and p.HasValue:
            try:
                vs = p.AsValueString()
                if vs and _texto_material_indica_hormigon(vs):
                    return True
            except Exception:
                pass
            try:
                s = p.AsString()
                if s and _texto_material_indica_hormigon(s):
                    return True
            except Exception:
                pass
    except Exception:
        pass
    for key in (u"Structural Material", u"Material estructural"):
        try:
            p = elem.LookupParameter(key)
            if p and p.HasValue:
                try:
                    vs = p.AsValueString()
                except Exception:
                    vs = None
                if vs and _texto_material_indica_hormigon(vs):
                    return True
        except Exception:
            pass
    try:
        from Autodesk.Revit.DB import BuiltInParameter, Floor, FloorType

        doc0 = elem.Document
        if doc0 is not None and elem is not None:
            tid = elem.GetTypeId()
            if tid is not None and tid != ElementId.InvalidElementId:
                et = doc0.GetElement(tid)
                if et is not None:
                    p2 = et.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)
                    if p2 is not None and p2.HasValue:
                        if p2.StorageType == StorageType.ElementId:
                            mid = p2.AsElementId()
                            if mid is not None and mid != ElementId.InvalidElementId:
                                m = doc0.GetElement(mid)
                                if m is not None and _mat_o_texto_sugiere_hormigon(
                                    m, p2, doc0
                                ):
                                    return True
                        for attr in (u"AsString", u"AsValueString"):
                            try:
                                t = getattr(p2, attr)()
                                if t and _texto_material_indica_hormigon(t):
                                    return True
                            except Exception:
                                pass
        _bip_fs = getattr(BuiltInParameter, u"FLOOR_PARAM_IS_STRUCTURAL", None)
        if _bip_fs is not None and doc0 is not None and isinstance(elem, Floor):
            p_st = elem.get_Parameter(_bip_fs)
            if p_st is not None and p_st.HasValue:
                try:
                    if p_st.AsInteger() == 1:
                        return True
                except Exception:
                    pass
        if doc0 is not None and elem is not None:
            tid = elem.GetTypeId()
            if tid is not None and tid != ElementId.InvalidElementId:
                et = doc0.GetElement(tid)
                if isinstance(et, FloorType):
                    if _bip_fs is not None:
                        p_t = et.get_Parameter(_bip_fs)
                        if p_t is not None and p_t.HasValue:
                            try:
                                if p_t.AsInteger() == 1:
                                    return True
                            except Exception:
                                pass
                    if _texto_material_indica_hormigon(et.Name or u""):
                        return True
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import Wall

        if isinstance(elem, Wall):
            return _wall_has_concrete_structure_material(elem)
    except Exception:
        pass
    return False


def _element_is_structural_concrete(elem):
    """Alias legado → criterio ampliado de canvas de elevación."""
    return _element_is_concrete_for_elevation_canvas(elem)


def _transform_xyz(xyz, tr):
    if xyz is None:
        return None
    if tr is None:
        return xyz
    try:
        return tr.OfPoint(xyz)
    except Exception:
        return xyz


def _xyz_to_mm_tuple(xyz):
    if xyz is None:
        return None
    try:
        return (
            internal_to_mm(float(xyz.X)),
            internal_to_mm(float(xyz.Y)),
            internal_to_mm(float(xyz.Z)),
        )
    except Exception:
        return None


def _curve_tessellate_xyz_mm(crv, transform=None):
    """Puntos discretos de una curva Revit en mm (mundo)."""
    if crv is None:
        return []
    pts = []
    try:
        raw = list(crv.Tessellate())
    except Exception:
        raw = []
    if not raw:
        try:
            raw = [crv.GetEndPoint(0), crv.GetEndPoint(1)]
        except Exception:
            raw = []
    for p in raw:
        pw = _transform_xyz(p, transform)
        t = _xyz_to_mm_tuple(pw)
        if t is not None:
            pts.append(t)
    cleaned = []
    for t in pts:
        if cleaned and _dist3_mm(cleaned[-1], t) < 0.5:
            continue
        cleaned.append(t)
    return cleaned


def _dist3_mm(a, b):
    return (
        (float(a[0]) - float(b[0])) ** 2
        + (float(a[1]) - float(b[1])) ** 2
        + (float(a[2]) - float(b[2])) ** 2
    ) ** 0.5


def _geometry_options_model():
    """Geometría de modelo (sin corte de vista) — respaldo si la vista no devuelve sólidos."""
    return _geometry_options_fine()


def _is_solid_geom(g):
    if g is None:
        return False
    try:
        if isinstance(g, Solid):
            return True
    except Exception:
        pass
    try:
        return type(g).__name__ == u"Solid"
    except Exception:
        return False


def _is_mesh_geom(g):
    if g is None:
        return False
    try:
        if isinstance(g, Mesh):
            return True
    except Exception:
        pass
    try:
        return type(g).__name__ == u"Mesh"
    except Exception:
        return False


def _is_curve_geom(g):
    if g is None:
        return False
    try:
        if hasattr(g, u"GetEndPoint") and hasattr(g, u"Tessellate"):
            return True
    except Exception:
        pass
    return False


def _walk_solid_segments(solid, transform=None):
    """Aristas de caras + bordes del sólido (contorno visible en sección)."""
    if solid is None:
        return
    try:
        if float(solid.Volume) <= 1e-12:
            return
    except Exception:
        pass
    try:
        faces = solid.Faces
    except Exception:
        faces = None
    if faces is not None:
        try:
            face_list = list(faces)
        except Exception:
            face_list = []
        for face in face_list:
            loops = None
            try:
                loops = face.GetEdgesAsCurveLoops()
            except Exception:
                loops = None
            if loops is None:
                continue
            try:
                loop_list = list(loops)
            except Exception:
                loop_list = []
            for loop in loop_list:
                try:
                    curves = list(loop)
                except Exception:
                    curves = []
                for crv in curves:
                    pts = _curve_tessellate_xyz_mm(crv, transform)
                    if len(pts) >= 2:
                        yield pts
    try:
        edges = solid.Edges
    except Exception:
        edges = None
    if edges is not None:
        try:
            edge_list = list(edges)
        except Exception:
            edge_list = []
        for edge in edge_list:
            try:
                ec = edge.AsCurve()
            except Exception:
                ec = None
            pts = _curve_tessellate_xyz_mm(ec, transform)
            if len(pts) >= 2:
                yield pts


def _walk_mesh_segments(mesh, transform=None):
    if mesh is None:
        return
    verts_mm = []
    try:
        nvert = int(mesh.Vertices.Count)
    except Exception:
        nvert = 0
    for i in range(nvert):
        try:
            v = mesh.Vertices[i]
        except Exception:
            continue
        t = _xyz_to_mm_tuple(_transform_xyz(v, transform))
        if t is not None:
            verts_mm.append(t)
    if not verts_mm:
        return
    try:
        nt = int(mesh.NumTriangles)
    except Exception:
        return
    seen = set()
    for ti in range(nt):
        try:
            tri = mesh.get_Triangle(ti)
        except Exception:
            continue
        try:
            ia = int(tri.get_Index(0))
            ib = int(tri.get_Index(1))
            ic = int(tri.get_Index(2))
        except Exception:
            continue
        for a, b in ((ia, ib), (ib, ic), (ic, ia)):
            key = (min(a, b), max(a, b))
            if key in seen:
                continue
            seen.add(key)
            if a < 0 or b < 0 or a >= len(verts_mm) or b >= len(verts_mm):
                continue
            seg = [verts_mm[a], verts_mm[b]]
            if _dist3_mm(seg[0], seg[1]) >= 0.5:
                yield seg


def _transform_normal_xyz(n, tr):
    if n is None:
        return None
    if tr is None:
        return n
    try:
        nx = (
            float(tr.BasisX.X) * float(n.X)
            + float(tr.BasisY.X) * float(n.Y)
            + float(tr.BasisZ.X) * float(n.Z)
        )
        ny = (
            float(tr.BasisX.Y) * float(n.X)
            + float(tr.BasisY.Y) * float(n.Y)
            + float(tr.BasisZ.Y) * float(n.Z)
        )
        nz = (
            float(tr.BasisX.Z) * float(n.X)
            + float(tr.BasisY.Z) * float(n.Y)
            + float(tr.BasisZ.Z) * float(n.Z)
        )
        return XYZ(nx, ny, nz)
    except Exception:
        return n


def _face_world_normal(face, transform=None):
    if face is None:
        return None
    n = None
    try:
        if isinstance(face, PlanarFace):
            n = face.FaceNormal
    except Exception:
        n = None
    if n is None:
        try:
            bb = face.GetBoundingBox()
            if bb is not None:
                u_mid = 0.5 * (float(bb.Min.U) + float(bb.Max.U))
                v_mid = 0.5 * (float(bb.Min.V) + float(bb.Max.V))
                derivs = face.ComputeDerivatives(UV(u_mid, v_mid))
                n = derivs.Normal
        except Exception:
            n = None
    if n is None:
        return None
    n = _transform_normal_xyz(n, transform)
    return _vector_unit(n)


def _walk_solid_section_visible_segments(solid, view_dir, transform=None):
    """
    En sección/alzado: aristas de caras de corte (normal ∥ ViewDirection).
    Si no hay caras de corte, todas las aristas del sólido de vista.
    """
    if solid is None:
        return
    try:
        if float(solid.Volume) <= 1e-12:
            return
    except Exception:
        pass
    vd = _vector_unit(view_dir)
    cut_segs = []
    all_segs = []
    try:
        face_list = list(solid.Faces)
    except Exception:
        face_list = []
    for face in face_list:
        normal = _face_world_normal(face, transform)
        is_cut = False
        if normal is not None and vd is not None:
            is_cut = abs(_vector_dot(normal, vd)) >= _SECTION_CUT_FACE_DOT_MIN
        loops = None
        try:
            loops = face.GetEdgesAsCurveLoops()
        except Exception:
            loops = None
        if loops is None:
            continue
        try:
            loop_list = list(loops)
        except Exception:
            loop_list = []
        for loop in loop_list:
            try:
                curves = list(loop)
            except Exception:
                curves = []
            for crv in curves:
                pts = _curve_tessellate_xyz_mm(crv, transform)
                if len(pts) >= 2:
                    all_segs.append(pts)
                    if is_cut:
                        cut_segs.append(pts)
    target = cut_segs if cut_segs else all_segs
    for pts in target:
        yield pts


def _walk_section_visible_geometry(geom_elem, view_dir, transform=None):
    """Geometría de vista en sección/alzado: cortes + mallas."""
    if geom_elem is None:
        return
    items = []
    try:
        for g in geom_elem:
            items.append(g)
    except Exception:
        try:
            n = int(geom_elem.Size)
        except Exception:
            n = 0
        for i in range(n):
            try:
                items.append(geom_elem[i])
            except Exception:
                pass
    for g in items:
        if g is None:
            continue
        if _is_solid_geom(g):
            for pts in _walk_solid_section_visible_segments(g, view_dir, transform):
                yield pts
            continue
        if _is_mesh_geom(g):
            for pts in _walk_mesh_segments(g, transform):
                yield pts
            continue
        if _is_curve_geom(g) and not _is_solid_geom(g) and not _is_mesh_geom(g):
            pts = _curve_tessellate_xyz_mm(g, transform)
            if len(pts) >= 2:
                yield pts
            continue
        if isinstance(g, GeometryInstance):
            try:
                tr_inst = g.Transform
            except Exception:
                tr_inst = None
            tr = tr_inst
            if transform is not None and tr_inst is not None:
                try:
                    tr = transform.Multiply(tr_inst)
                except Exception:
                    tr = tr_inst
            elif transform is not None:
                tr = transform
            try:
                inst = g.GetInstanceGeometry()
            except Exception:
                inst = None
            if inst is None:
                try:
                    inst = g.GetSymbolGeometry()
                except Exception:
                    inst = None
            for seg in _walk_section_visible_geometry(inst, view_dir, tr):
                yield seg


def _geometry_section_view_segments_mm(elem, opts, view_dir):
    if elem is None or opts is None:
        return []
    try:
        geo = elem.get_Geometry(opts)
    except Exception:
        geo = None
    if geo is None:
        return []
    segs = []
    for seg in _walk_section_visible_geometry(geo, view_dir, None):
        segs.append(seg)
    return segs


def _element_bbox_corners_xyz_mm(elem, view):
    """Esquinas del bounding box del elemento en mm (modelo)."""
    bb = None
    if view is not None:
        try:
            bb = elem.get_BoundingBox(view)
        except Exception:
            bb = None
    if bb is None:
        try:
            bb = elem.get_BoundingBox(None)
        except Exception:
            bb = None
    if bb is None:
        return []
    try:
        mn, mx = bb.Min, bb.Max
        corners_ft = [
            XYZ(float(mn.X), float(mn.Y), float(mn.Z)),
            XYZ(float(mx.X), float(mn.Y), float(mn.Z)),
            XYZ(float(mx.X), float(mx.Y), float(mn.Z)),
            XYZ(float(mn.X), float(mx.Y), float(mn.Z)),
            XYZ(float(mn.X), float(mn.Y), float(mx.Z)),
            XYZ(float(mx.X), float(mn.Y), float(mx.Z)),
            XYZ(float(mx.X), float(mx.Y), float(mx.Z)),
            XYZ(float(mn.X), float(mx.Y), float(mx.Z)),
        ]
    except Exception:
        return []
    out = []
    for c in corners_ft:
        t = _xyz_to_mm_tuple(c)
        if t is not None:
            out.append(t)
    return out


def _element_view_extents_uv_mm(elem, view, origin_mm, u_axis, v_axis):
    """
    Extensión AABB proyectada Right×Up (mm), origen = inicio de barra.

    Mismo criterio que ``armado_vigas.revit.elev_geometry.element_view_extents``.
    """
    corners = _element_bbox_corners_xyz_mm(elem, view)
    if not corners:
        return None
    us = []
    vs = []
    for t in corners:
        u, v = project_xyz_mm_to_uv(t, origin_mm, u_axis, v_axis)
        us.append(float(u))
        vs.append(float(v))
    if not us or not vs:
        return None
    return min(us), max(us), min(vs), max(vs)


def _uv_rect_edge_polylines(u0, u1, v0, v1):
    try:
        u0, u1 = float(u0), float(u1)
        v0, v1 = float(v0), float(v1)
    except Exception:
        return []
    if u1 < u0:
        u0, u1 = u1, u0
    if v1 < v0:
        v0, v1 = v1, v0
    return [
        [[u0, v0], [u1, v0]],
        [[u1, v0], [u1, v1]],
        [[u1, v1], [u0, v1]],
        [[u0, v1], [u0, v0]],
    ]


def _element_bbox_uv_polylines(elem, view, origin_mm, u_axis, v_axis):
    """Rectángulo 2D (UV mm) del bbox del elemento en el plano de la vista."""
    ext = _element_view_extents_uv_mm(elem, view, origin_mm, u_axis, v_axis)
    if ext is None:
        return []
    u0, u1, v0, v1 = ext
    if (u1 - u0) < _CONTEXT_MIN_EDGE_MM and (v1 - v0) < _CONTEXT_MIN_EDGE_MM:
        return []
    return _uv_rect_edge_polylines(u0, u1, v0, v1)


def _element_bbox_segments_mm(elem, view=None, transform=None):
    """Wireframe del bounding box del elemento en mm."""
    if elem is None:
        return []
    bb = None
    try:
        bb = elem.get_BoundingBox(view)
    except Exception:
        bb = None
    if bb is None:
        try:
            bb = elem.get_BoundingBox(None)
        except Exception:
            bb = None
    if bb is None:
        return []
    try:
        mn, mx = bb.Min, bb.Max
        corners = [
            XYZ(float(mn.X), float(mn.Y), float(mn.Z)),
            XYZ(float(mx.X), float(mn.Y), float(mn.Z)),
            XYZ(float(mx.X), float(mx.Y), float(mn.Z)),
            XYZ(float(mn.X), float(mx.Y), float(mn.Z)),
            XYZ(float(mn.X), float(mn.Y), float(mx.Z)),
            XYZ(float(mx.X), float(mn.Y), float(mx.Z)),
            XYZ(float(mx.X), float(mx.Y), float(mx.Z)),
            XYZ(float(mn.X), float(mx.Y), float(mx.Z)),
        ]
    except Exception:
        return []
    edges = (
        (0, 1), (1, 2), (2, 3), (3, 0),
        (4, 5), (5, 6), (6, 7), (7, 4),
        (0, 4), (1, 5), (2, 6), (3, 7),
    )
    out = []
    for i, j in edges:
        try:
            a = _xyz_to_mm_tuple(_transform_xyz(corners[i], transform))
            b = _xyz_to_mm_tuple(_transform_xyz(corners[j], transform))
        except Exception:
            continue
        if a is not None and b is not None and _dist3_mm(a, b) >= 0.5:
            out.append([a, b])
    return out


def _geometry_segments_mm(elem, opts):
    """Listas de puntos mm por segmento de geometría del elemento."""
    if elem is None or opts is None:
        return []
    try:
        geo = elem.get_Geometry(opts)
    except Exception:
        geo = None
    if geo is None:
        return []
    segs = []
    for seg in _walk_geometry_curves(geo, None):
        segs.append(seg)
    return segs


def _walk_geometry_curves(geom_elem, transform=None):
    """Recorre geometría de elemento y devuelve listas de puntos mm por curva."""
    if geom_elem is None:
        return
    items = []
    try:
        for g in geom_elem:
            items.append(g)
    except Exception:
        try:
            n = int(geom_elem.Size)
        except Exception:
            n = 0
        for i in range(n):
            try:
                items.append(geom_elem[i])
            except Exception:
                pass
    for g in items:
        if g is None:
            continue
        if _is_curve_geom(g) and not _is_solid_geom(g) and not _is_mesh_geom(g):
            pts = _curve_tessellate_xyz_mm(g, transform)
            if len(pts) >= 2:
                yield pts
            continue
        if _is_solid_geom(g):
            for pts in _walk_solid_segments(g, transform):
                yield pts
            continue
        if _is_mesh_geom(g):
            for pts in _walk_mesh_segments(g, transform):
                yield pts
            continue
        if isinstance(g, GeometryInstance):
            try:
                tr_inst = g.Transform
            except Exception:
                tr_inst = None
            tr = tr_inst
            if transform is not None and tr_inst is not None:
                try:
                    tr = transform.Multiply(tr_inst)
                except Exception:
                    tr = tr_inst
            elif transform is not None:
                tr = transform
            try:
                inst = g.GetInstanceGeometry()
            except Exception:
                inst = None
            if inst is None:
                try:
                    inst = g.GetSymbolGeometry()
                except Exception:
                    inst = None
            for seg in _walk_geometry_curves(inst, tr):
                yield seg


def _vector_dot(a, b):
    try:
        return float(a.X) * float(b.X) + float(a.Y) * float(b.Y) + float(a.Z) * float(b.Z)
    except Exception:
        return 0.0


def _vector_unit(xyz):
    if xyz is None:
        return None
    try:
        ln = (float(xyz.X) ** 2 + float(xyz.Y) ** 2 + float(xyz.Z) ** 2) ** 0.5
    except Exception:
        return None
    if ln < 1e-12:
        return None
    return XYZ(float(xyz.X) / ln, float(xyz.Y) / ln, float(xyz.Z) / ln)


def _build_view_projection_frame(view, origin_xyz_ft):
    """
    Marco UV alineado con la vista activa (Right/Up en pantalla).

    En sección o alzado usa directamente RightDirection y UpDirection de Revit
    (mismo sistema que el alzado en pantalla).

    Returns:
        (origin_mm, u_axis, v_axis) o None
    """
    if view is None or origin_xyz_ft is None:
        return None
    try:
        n = _vector_unit(view.ViewDirection)
        right = _vector_unit(view.RightDirection)
        up = _vector_unit(view.UpDirection)
    except Exception:
        return None
    if n is None or right is None or up is None:
        return None
    try:
        if _is_section_elevation_view(view):
            u = right
            v = up
            if _vector_dot(v, up) < 0.0:
                v = XYZ(-float(v.X), -float(v.Y), -float(v.Z))
        else:
            u = XYZ(
                float(right.X) - float(n.X) * _vector_dot(right, n),
                float(right.Y) - float(n.Y) * _vector_dot(right, n),
                float(right.Z) - float(n.Z) * _vector_dot(right, n),
            )
            u = _vector_unit(u)
            if u is None:
                u = right
            v = n.CrossProduct(u)
            v = _vector_unit(v)
            if v is None:
                return None
            if _vector_dot(v, up) < 0.0:
                u = XYZ(-float(u.X), -float(u.Y), -float(u.Z))
                v = n.CrossProduct(u)
                v = _vector_unit(v)
        origin_mm = _xyz_to_mm_tuple(origin_xyz_ft)
        if origin_mm is None:
            return None
        u_ax = (float(u.X), float(u.Y), float(u.Z))
        v_ax = (float(v.X), float(v.Y), float(v.Z))
        return origin_mm, u_ax, v_ax
    except Exception:
        return None


def _polyline_uv_from_xyz_mm(points_xyz_mm, origin_mm, u_axis, v_axis):
    uv = []
    for p in points_xyz_mm or []:
        uv.append(project_xyz_mm_to_uv(p, origin_mm, u_axis, v_axis))
    if len(uv) < 2:
        return None
    try:
        du = float(uv[-1][0]) - float(uv[0][0])
        dv = float(uv[-1][1]) - float(uv[0][1])
        seg_len = (du * du + dv * dv) ** 0.5
    except Exception:
        seg_len = 0.0
    if seg_len < _CONTEXT_MIN_EDGE_MM:
        return None
    return [[float(u), float(v)] for u, v in uv]


def _collect_concrete_context_polylines_uv(
    doc, view, origin_mm, u_axis, v_axis, skip_element_id=None
):
    """
    Polilíneas UV (mm) del hormigón visible en ``view`` para el esquema de UI.

    Muros, vigas, fundaciones, columnas y losas de hormigón: silueta AABB
    proyectada Right×Up (como Armado Vigas).
    """
    if doc is None or view is None or origin_mm is None:
        return [], 0, []
    polylines = []
    fill_rects_uv = []
    n_elems = 0
    try:
        skip_int = _element_id_int(skip_element_id) if skip_element_id else None
    except Exception:
        skip_int = None
    for cat in _CATS_CONCRETE_IN_VIEW:
        try:
            coll = (
                FilteredElementCollector(doc, view.Id)
                .OfCategory(cat)
                .WhereElementIsNotElementType()
            )
        except Exception:
            continue
        try:
            elems = list(coll)
        except Exception:
            elems = []
        for elem in elems:
            if elem is None:
                continue
            if skip_int is not None:
                try:
                    if _element_id_int(elem.Id) == skip_int:
                        continue
                except Exception:
                    pass
            if not _element_is_concrete_for_elevation_canvas(elem):
                continue
            n_elems += 1
            if n_elems > _CONTEXT_MAX_ELEMENTS:
                break
            ext = _element_view_extents_uv_mm(
                elem, view, origin_mm, u_axis, v_axis
            )
            if ext is not None:
                u0, u1, v0, v1 = ext
                if (u1 - u0) >= _CONTEXT_MIN_EDGE_MM or (
                    v1 - v0
                ) >= _CONTEXT_MIN_EDGE_MM:
                    fill_rects_uv.append(
                        [float(u0), float(u1), float(v0), float(v1)]
                    )
        if n_elems > _CONTEXT_MAX_ELEMENTS:
            break
    return polylines, n_elems, fill_rects_uv


def _curves_to_xyz_mm(curves):
    """Vértices de la polilínea centerline en mm (mundo Revit)."""
    pts = []
    for c in curves or []:
        p0, p1 = _curve_endpoints(c)
        if p0 is None or p1 is None:
            continue
        a = (internal_to_mm(p0.X), internal_to_mm(p0.Y), internal_to_mm(p0.Z))
        b = (internal_to_mm(p1.X), internal_to_mm(p1.Y), internal_to_mm(p1.Z))
        if not pts:
            pts.append(a)
        else:
            last = pts[-1]
            d_start = (
                (last[0] - a[0]) ** 2 + (last[1] - a[1]) ** 2 + (last[2] - a[2]) ** 2
            ) ** 0.5
            if d_start > 1.0:  # > 1 mm → gap / sentido invertido
                d_end = (
                    (last[0] - b[0]) ** 2
                    + (last[1] - b[1]) ** 2
                    + (last[2] - b[2]) ** 2
                ) ** 0.5
                if d_end < d_start:
                    a, b = b, a
                else:
                    pts.append(a)
        pts.append(b)
    return pts


def prepare_division_session(doc, rebar, concrete_grade=None, view=None):
    """
    Datos para la UI multipunto (sin crear elementos).

    Returns:
        (ok, err, session_dict)
    """
    if doc is None or rebar is None:
        return False, u"Parámetros incompletos.", None
    ok, err, curves = check_eligibility(doc, rebar)
    if not ok:
        return False, err, None
    d_mm, lap_mm = lap_mm_for_rebar(doc, rebar, concrete_grade)
    if d_mm is None or lap_mm is None or lap_mm <= 0:
        return False, u"No se pudo obtener diámetro / traslape de la tabla.", None
    if not curves:
        return False, u"No se pudo leer la centerline.", None
    total_ft = sum(_curve_length(c) for c in curves)
    total_mm = internal_to_mm(total_ft)
    plan = None
    context_polylines_uv = []
    context_fill_rects_uv = []
    context_n_elems = 0
    xyz_mm = _curves_to_xyz_mm(curves)
    origin_xyz_ft = None
    if curves:
        try:
            origin_xyz_ft = _curve_endpoints(curves[0])[0]
        except Exception:
            origin_xyz_ft = None
    frame = None
    if view is not None and origin_xyz_ft is not None:
        frame = _build_view_projection_frame(view, origin_xyz_ft)
    if frame is not None and xyz_mm:
        origin_mm, u_axis, v_axis = frame
        try:
            plan = build_plan_polyline_from_frame_mm(xyz_mm, origin_mm, u_axis, v_axis)
        except Exception:
            plan = None
        if plan:
            try:
                (
                    context_polylines_uv,
                    context_n_elems,
                    context_fill_rects_uv,
                ) = _collect_concrete_context_polylines_uv(
                    doc,
                    view,
                    origin_mm,
                    u_axis,
                    v_axis,
                    skip_element_id=rebar.Id,
                )
            except Exception:
                context_polylines_uv = []
                context_fill_rects_uv = []
                context_n_elems = 0
    if plan is None:
        try:
            from dividir_rebar_punto_geom import build_plan_polyline_mm

            n = _rebar_normal(rebar)
            normal_t = None
            if n is not None:
                try:
                    normal_t = (float(n.X), float(n.Y), float(n.Z))
                except Exception:
                    normal_t = None
            plan = build_plan_polyline_mm(xyz_mm, normal=normal_t)
        except Exception:
            plan = None
    session = {
        u"rebar_id": rebar.Id,
        u"rebar_id_int": _element_id_int(rebar.Id),
        u"diameter_mm": d_mm,
        u"lap_mm": float(lap_mm),
        u"total_mm": float(total_mm),
        u"n_segments": len(curves),
        u"n_positions": _cantidad_posiciones(rebar),
        u"layout": _layout_rule_name(rebar),
        u"concrete_grade": concrete_grade,
        u"context_polylines_uv": context_polylines_uv,
        u"context_fill_rects_uv": context_fill_rects_uv,
        u"context_n_elems": int(context_n_elems),
        u"context_n_polylines": len(context_polylines_uv) + len(context_fill_rects_uv),
        u"view_is_section_elevation": bool(
            view is not None and _is_section_elevation_view(view)
        ),
    }
    if plan and plan.get(u"points_uv"):
        session[u"plan_points_uv"] = plan[u"points_uv"]
        session[u"plan_plane"] = plan.get(u"plane") or u"xy"
        session[u"plan_arc_mm"] = plan.get(u"arc_mm") or []
        session[u"plan_flip_v"] = bool(plan.get(u"flip_v", True))
        plan_tot = float(plan.get(u"total_mm") or 0.0)
        if plan_tot > 1e-6 and abs(plan_tot - float(total_mm)) > 0.5:
            ratio = float(total_mm) / plan_tot
            session[u"plan_arc_mm"] = [
                float(a) * ratio for a in session[u"plan_arc_mm"]
            ]
    return True, None, session


# ---------------------------------------------------------------------------
# Creación Revit
# ---------------------------------------------------------------------------


def _curves_to_list(curves_clean):
    from Autodesk.Revit.DB import Curve

    lst = List[Curve]()
    for c in curves_clean:
        lst.Add(c)
    return lst


def _curves_to_array(curves_clean):
    from Autodesk.Revit.DB import Curve

    n = len(curves_clean)
    arr = System.Array.CreateInstance(Curve, n)
    for i in range(n):
        arr[i] = curves_clean[i]
    return arr


def _dot_safe(a, b):
    try:
        return float(a.DotProduct(b))
    except Exception:
        return None


_ARMADURA_PARAM_PREFIX = u"Armadura_"


def _param_definition_name(param):
    if param is None:
        return u""
    try:
        return _as_unicode(param.Definition.Name).strip()
    except Exception:
        return u""


def _is_armadura_instance_param(param):
    """True si el parámetro tiene nombre con prefijo Armadura_."""
    if param is None:
        return False
    name = _param_definition_name(param)
    if not name:
        return False
    return name.startswith(_ARMADURA_PARAM_PREFIX) or name.lower().startswith(
        u"armadura_"
    )


def _find_param_by_name(element, name):
    if element is None or not name:
        return None
    try:
        p = element.LookupParameter(name)
        if p is not None:
            return p
    except Exception:
        pass
    target = _as_unicode(name).strip().lower()
    try:
        for p in element.Parameters:
            if p is None:
                continue
            try:
                dn = _param_definition_name(p).lower()
            except Exception:
                continue
            if dn == target:
                return p
    except Exception:
        pass
    return None


def _copy_one_param_value(src_param, dst_param):
    """Copia valor según StorageType. Devuelve True si escribió."""
    if src_param is None or dst_param is None:
        return False
    try:
        if dst_param.IsReadOnly:
            return False
    except Exception:
        return False
    try:
        st = src_param.StorageType
    except Exception:
        return False
    try:
        if st == StorageType.String:
            val = src_param.AsString()
            if val is None:
                try:
                    val = src_param.AsValueString()
                except Exception:
                    val = u""
            dst_param.Set(val if val is not None else u"")
            return True
        if st == StorageType.Integer:
            # Incluye Yes/No
            dst_param.Set(int(src_param.AsInteger()))
            return True
        if st == StorageType.Double:
            dst_param.Set(float(src_param.AsDouble()))
            return True
        if st == StorageType.ElementId:
            eid = src_param.AsElementId()
            if eid is None:
                eid = ElementId.InvalidElementId
            dst_param.Set(eid)
            return True
    except Exception:
        pass
    # Respaldo por ValueString (p. ej. algunos shared params)
    try:
        vs = src_param.AsValueString()
        if vs is not None:
            dst_param.SetValueString(_as_unicode(vs))
            return True
    except Exception:
        pass
    return False


def copy_armadura_instance_parameters(src_rebar, dst_rebar):
    """
    Hereda en ``dst_rebar`` todos los parámetros de instancia cuyo nombre
    empieza por ``Armadura_`` (p. ej. ``Armadura_Eje``).

    Returns:
        int: cantidad de parámetros escritos con éxito.
    """
    if src_rebar is None or dst_rebar is None:
        return 0
    n_ok = 0
    try:
        params = list(src_rebar.Parameters)
    except Exception:
        params = []
    seen = set()
    for sp in params:
        if not _is_armadura_instance_param(sp):
            continue
        name = _param_definition_name(sp)
        key = name.lower()
        if not key or key in seen:
            continue
        seen.add(key)
        # Sin valor: igual intentar copiar (vacío / 0) solo si HasValue
        try:
            if not sp.HasValue:
                # Aun así copiar strings vacíos / yes-no en 0 puede ser útil;
                # si no hay valor, saltar para no pisar defaults del destino.
                continue
        except Exception:
            pass
        dp = _find_param_by_name(dst_rebar, name)
        if dp is None:
            continue
        if _copy_one_param_value(sp, dp):
            n_ok += 1
    return n_ok


def _copy_layout(src, dst):
    """Copia layout Single / FixedNumber / MaximumSpacing de ``src`` a ``dst``."""
    a0 = _shape_driven_accessor(src)
    a1 = _shape_driven_accessor(dst)
    if a0 is None or a1 is None:
        return False, u"Layout no copiable (sin ShapeDrivenAccessor)."
    rule_name = _layout_rule_name(src, a0)
    try:
        sp = float(src.MaxSpacing)
    except Exception:
        sp = 0.0
    try:
        alen = float(a0.ArrayLength)
    except Exception:
        try:
            alen = float(a0.GetArrayLength())
        except Exception:
            alen = 0.0
    try:
        b_side = bool(a0.BarsOnNormalSide)
    except Exception:
        b_side = True
    try:
        n0 = a0.Normal
        n1 = a1.Normal
        d = _dot_safe(n0, n1)
        if d is not None and d < 0:
            b_side = not b_side
    except Exception:
        pass
    try:
        inc0 = bool(src.IncludeFirstBar)
        inc1 = bool(src.IncludeLastBar)
    except Exception:
        inc0, inc1 = True, True
    nbars = _cantidad_posiciones(src)

    # Maximum Spacing con ArrayLength ≤ spacing → Single (igual que Revit).
    if rule_name == u"MaximumSpacing":
        if int(nbars) <= 1 or alen < 1e-9 or (sp > 1e-12 and sp >= alen - 1e-9):
            try:
                a1.SetLayoutAsSingle()
                return True, u""
            except Exception as ex:
                return False, _exception_text(ex)

    combos = (
        (bool(b_side), bool(inc0), bool(inc1)),
        (not bool(b_side), bool(inc0), bool(inc1)),
    )
    last_ex = None
    for b_try, i0, i1 in combos:
        try:
            if rule_name == u"Single":
                a1.SetLayoutAsSingle()
            elif rule_name == u"MaximumSpacing":
                a1.SetLayoutAsMaximumSpacing(sp, alen, b_try, i0, i1)
            elif rule_name in (u"Number", u"FixedNumber"):
                a1.SetLayoutAsFixedNumber(nbars, alen, b_try, i0, i1)
            elif rule_name == u"NumberWithSpacing":
                a1.SetLayoutAsNumberWithSpacing(nbars, sp, alen, b_try, i0, i1)
            elif rule_name == u"MinimumClearSpacing":
                a1.SetLayoutAsMinimumClearSpacing(sp, alen, b_try, i0, i1)
            else:
                if int(nbars) <= 1:
                    a1.SetLayoutAsSingle()
                else:
                    a1.SetLayoutAsFixedNumber(nbars, alen, b_try, i0, i1)
            return True, u""
        except Exception as ex:
            last_ex = ex
            continue
    try:
        if int(nbars) > 1:
            a1.SetLayoutAsFixedNumber(nbars, alen, b_side, inc0, inc1)
            return True, u""
        a1.SetLayoutAsSingle()
        return True, u""
    except Exception:
        return False, _exception_text(last_ex) if last_ex is not None else u"layout"


def _copy_bar_included(src, dst, n_pos):
    """Replica barras incluidas/excluidas por posición."""
    n = int(n_pos or 0)
    if n <= 1:
        return
    try:
        n_dst = _cantidad_posiciones(dst)
    except Exception:
        n_dst = n
    for i in range(min(n, n_dst)):
        try:
            dst.SetBarIncluded(bool(src.IsBarIncluded(int(i))), int(i))
        except Exception:
            pass


def _normal_from_chain(curves, fallback):
    if not curves:
        return fallback or XYZ.BasisZ
    try:
        p0, p1 = _curve_endpoints(curves[0])
        if p0 is None or p1 is None:
            return fallback or XYZ.BasisZ
        if len(curves) > 1:
            q0, q1 = _curve_endpoints(curves[1])
            if q0 is None or q1 is None:
                return fallback or XYZ.BasisZ
            v1 = (p1 - p0).Normalize()
            v2 = (q1 - q0).Normalize()
            n = v1.CrossProduct(v2)
            if n.GetLength() > 1e-9:
                return n.Normalize()
    except Exception:
        pass
    return fallback or XYZ.BasisZ


def _prepare_curves(curves):
    out = []
    for c in curves or []:
        cc = _clone_curve(c)
        if cc is not None:
            out.append(cc)
    return out


def _find_longest_segment_index(curves):
    """Índice del segmento más largo de la cadena (vano / tramo principal)."""
    best_i = 0
    best_len = -1.0
    for i, c in enumerate(curves or []):
        if c is None:
            continue
        ln = _curve_length(c)
        if ln > best_len:
            best_len = ln
            best_i = i
    return best_i


def _shape_major_segment_index(rebar_shape):
    """
    ``MajorSegmentIndex`` de la definición de forma, o ``0`` si no aplica.

    El índice indica qué segmento de la forma es el «mayor»; las curvas
    pasadas a CreateFromCurves deben alinearlo con el segmento geométrico
    más largo.
    """
    if rebar_shape is None:
        return 0
    try:
        from Autodesk.Revit.DB.Structure import RebarShapeDefinitionBySegments

        defn = rebar_shape.GetRebarShapeDefinition()
        if defn is not None and isinstance(defn, RebarShapeDefinitionBySegments):
            idx = int(defn.MajorSegmentIndex)
            n = int(defn.NumberOfSegments)
            if 0 <= idx < n:
                return idx
    except Exception:
        pass
    return 0


def _reverse_curve_chain(curves):
    """Invierte el orden de la polilínea (cada curva CreateReversed)."""
    out = []
    for c in reversed(list(curves or [])):
        if c is None:
            continue
        rev = None
        try:
            rev = c.CreateReversed()
        except Exception:
            rev = None
        if rev is None:
            p0, p1 = _curve_endpoints(c)
            if p0 is not None and p1 is not None:
                try:
                    rev = Line.CreateBound(p1, p0)
                except Exception:
                    rev = None
        if rev is not None:
            out.append(rev)
    return out


def _orient_chain_longest_as_major(curves, rebar_shape=None):
    """
    Orienta la cadena para que el segmento más largo coincida con el
    segmento mayor de la forma (``MajorSegmentIndex``, por defecto 0).

    Solo se puede invertir la polilínea completa (continua). Si tras invertir
    el índice del más largo queda más cerca del major, se usa la inversión.

    Returns:
        (curves_orientadas, reversed_bool)
    """
    clist = list(curves or [])
    if len(clist) < 2:
        return clist, False
    major_idx = _shape_major_segment_index(rebar_shape)
    if major_idx < 0:
        major_idx = 0
    if major_idx >= len(clist):
        major_idx = 0

    long_idx = _find_longest_segment_index(clist)
    if long_idx == major_idx:
        return clist, False

    rev = _reverse_curve_chain(clist)
    if not rev or len(rev) != len(clist):
        return clist, False
    long_rev = _find_longest_segment_index(rev)
    if long_rev == major_idx:
        return rev, True
    # Preferir la orientación que acerca el tramo largo al major
    if abs(long_rev - major_idx) < abs(long_idx - major_idx):
        return rev, True
    return clist, False


def _create_from_curves(
    doc,
    curves_list,
    host,
    norm,
    bar_type,
    style,
    start_hook,
    end_hook,
    start_orient,
    end_orient,
    rebar_shape=None,
):
    curves_clean = _prepare_curves(curves_list)
    if not curves_clean:
        raise RuntimeError(u"Sin curvas válidas para CreateFromCurves.")

    # Segmento más largo → segmento mayor de la forma (invertir si hace falta).
    curves_clean, reversed_chain = _orient_chain_longest_as_major(
        curves_clean, rebar_shape
    )
    h0_use = start_hook
    h1_use = end_hook
    so_use = start_orient
    eo_use = end_orient
    if reversed_chain:
        h0_use, h1_use = end_hook, start_hook
        so_use, eo_use = end_orient, start_orient

    # Preferir CreateFromCurvesAndShape si hay shape objetivo (regla de división)
    if rebar_shape is not None:
        try:
            from dividir_rebar_punto_shapes import create_rebar_with_shape

            rb_shaped = create_rebar_with_shape(
                doc,
                rebar_shape,
                bar_type,
                h0_use,
                h1_use,
                host,
                norm,
                curves_clean,
                so_use,
                eo_use,
            )
            if rb_shaped is not None:
                try:
                    if h0_use is not None:
                        rb_shaped.SetHookTypeId(0, h0_use.Id)
                    else:
                        rb_shaped.SetHookTypeId(0, ElementId.InvalidElementId)
                    if h1_use is not None:
                        rb_shaped.SetHookTypeId(1, h1_use.Id)
                    else:
                        rb_shaped.SetHookTypeId(1, ElementId.InvalidElementId)
                except Exception:
                    pass
                return rb_shaped
        except Exception:
            pass

    norms = []
    seen = []

    def _add_n(n):
        if n is None:
            return
        try:
            nn = n.Normalize()
        except Exception:
            nn = n
        for s in seen:
            d = _dot_safe(s, nn)
            if d is not None and abs(d) > 0.999:
                return
        seen.append(nn)
        norms.append(nn)

    _add_n(norm)
    try:
        if norm is not None:
            _add_n(norm.Negate())
    except Exception:
        pass
    _add_n(_normal_from_chain(curves_clean, norm))
    if not norms:
        norms = [XYZ.BasisZ]

    hook_pairs = ((h0_use, h1_use), (None, None))
    if h0_use is None and h1_use is None:
        hook_pairs = ((None, None),)
    orient_pairs = (
        (so_use, eo_use),
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
    )
    flag_pairs = ((False, True), (True, False), (True, True))
    last_err = u""

    for h0, h1 in hook_pairs:
        for so, eo in orient_pairs:
            for nvec in norms:
                for use_exist, create_new in flag_pairs:
                    for container in ("list", "array"):
                        try:
                            if container == "list":
                                payload = _curves_to_list(curves_clean)
                            else:
                                payload = _curves_to_array(curves_clean)
                            rb = Rebar.CreateFromCurves(
                                doc,
                                style,
                                bar_type,
                                h0,
                                h1,
                                host,
                                nvec,
                                payload,
                                so,
                                eo,
                                bool(use_exist),
                                bool(create_new),
                            )
                            if rb is not None:
                                try:
                                    if h0_use is not None:
                                        rb.SetHookTypeId(0, h0_use.Id)
                                    else:
                                        rb.SetHookTypeId(0, ElementId.InvalidElementId)
                                    if h1_use is not None:
                                        rb.SetHookTypeId(1, h1_use.Id)
                                    else:
                                        rb.SetHookTypeId(1, ElementId.InvalidElementId)
                                except Exception:
                                    pass
                                # Forzar shape si se indicó (CreateFromCurves no lo garantiza)
                                if rebar_shape is not None:
                                    try:
                                        from dividir_rebar_punto_shapes import set_rebar_shape

                                        set_rebar_shape(doc, rb, rebar_shape)
                                    except Exception:
                                        pass
                                return rb
                        except Exception as ex:
                            last_err = _exception_text(ex)
    raise RuntimeError(
        u"CreateFromCurves no produjo Rebar válido. {0}".format(last_err)
    )


def read_rebar_summary(doc, rebar, concrete_grade=None):
    ok, err, curves = check_eligibility(doc, rebar)
    d_mm, lap_mm = lap_mm_for_rebar(doc, rebar, concrete_grade)
    if not curves:
        curves = _centerline_curves(rebar, 0, True, True)
    if not curves:
        curves = _centerline_curves(rebar, 0, True, False)
    total_ft = sum(_curve_length(c) for c in curves) if curves else 0.0
    return {
        u"ok": ok,
        u"error": err,
        u"rebar_id": rebar.Id,
        u"rebar_id_int": _element_id_int(rebar.Id),
        u"diameter_mm": d_mm,
        u"lap_mm": lap_mm,
        u"total_mm": internal_to_mm(total_ft) if total_ft else None,
        u"n_segments": len(curves) if curves else 0,
        u"n_positions": _cantidad_posiciones(rebar),
        u"layout": _layout_rule_name(rebar),
    }


def resolve_active_model_view(uidoc):
    """
    Vista gráfica de modelo activa (como 37_RebarUnobscuredVista).

    Prioriza ``ActiveGraphicalView`` frente a ``ActiveView`` (este último puede
    desviarse con el navegador o paneles WPF).
    """
    if uidoc is None:
        return None
    doc = uidoc.Document
    if doc is None:
        return None
    view = None
    try:
        view = getattr(uidoc, u"ActiveGraphicalView", None)
    except Exception:
        view = None
    if view is None:
        try:
            view = uidoc.ActiveView
        except Exception:
            view = None
    if view is None:
        return None
    try:
        resolved = doc.GetElement(view.Id)
        if isinstance(resolved, View):
            view = resolved
    except Exception:
        pass
    if not isinstance(view, View):
        return None
    try:
        if view.IsTemplate:
            return None
    except Exception:
        pass
    if isinstance(view, (ViewSheet, ViewSchedule)):
        return None
    return view


def _presentation_mode_name(mode):
    if mode is None:
        return u""
    try:
        return unicode(mode.ToString() or u"")
    except NameError:
        try:
            return str(mode.ToString() or u"")
        except Exception:
            return u""
    except Exception:
        return u""


def _is_presentation_all(mode):
    try:
        if mode == RebarPresentationMode.All:
            return True
    except Exception:
        pass
    name = _presentation_mode_name(mode)
    return name == u"All" or name == u""


def capture_rebar_presentation(doc, rebar, preferred_view=None):
    """
    Captura ``PresentationMode`` (Show All / Middle / FirstLast / Select) por vista.

    Siempre incluye ``preferred_view`` si es válida. En el resto de vistas de
    modelo solo guarda overrides distintos de All, para no saturar el proyecto.
    En modo Select también captura índices ocultos (``IsBarHidden``).
    """
    snapshots = []
    if doc is None or rebar is None:
        return snapshots

    preferred_id = None
    if preferred_view is not None and isinstance(preferred_view, View):
        try:
            if not preferred_view.IsTemplate and not isinstance(
                preferred_view, (ViewSheet, ViewSchedule)
            ):
                preferred_id = preferred_view.Id
        except Exception:
            preferred_id = None

    views = []
    seen = set()
    if preferred_id is not None:
        try:
            pv = doc.GetElement(preferred_id)
            if isinstance(pv, View):
                views.append(pv)
                seen.add(_element_id_int(preferred_id))
        except Exception:
            pass

    try:
        for v in FilteredElementCollector(doc).OfClass(View):
            try:
                if v is None or bool(v.IsTemplate):
                    continue
            except Exception:
                continue
            if isinstance(v, (ViewSheet, ViewSchedule)):
                continue
            vid = _element_id_int(v.Id)
            if vid is None or vid in seen:
                continue
            seen.add(vid)
            views.append(v)
    except Exception:
        pass

    n_pos = _cantidad_posiciones(rebar)
    for view in views:
        try:
            mode = rebar.GetPresentationMode(view)
        except Exception:
            continue
        is_preferred = False
        try:
            is_preferred = preferred_id is not None and view.Id == preferred_id
        except Exception:
            is_preferred = False
        if not is_preferred and _is_presentation_all(mode):
            continue

        mode_name = _presentation_mode_name(mode)
        hidden = []
        if mode_name == u"Select" or u"Select" in mode_name:
            for i in range(int(n_pos)):
                try:
                    if bool(rebar.IsBarHidden(view, int(i))):
                        hidden.append(int(i))
                except Exception:
                    pass

        snapshots.append(
            {
                u"view_id": view.Id,
                u"mode": mode,
                u"mode_name": mode_name,
                u"hidden": hidden,
            }
        )
    return snapshots


def apply_rebar_presentation(doc, rebars, snapshots):
    """
    Aplica PresentationMode capturado a barras nuevas.

    Debe ejecutarse en Transaction abierta, preferible tras layout + Regenerate.
    Devuelve cuántas aplicaciones (barra×vista) tuvieron éxito.
    """
    if doc is None or not rebars or not snapshots:
        return 0
    n_ok = 0
    for snap in snapshots:
        try:
            view = doc.GetElement(snap.get(u"view_id"))
        except Exception:
            view = None
        if not isinstance(view, View):
            continue
        try:
            if bool(view.IsTemplate):
                continue
        except Exception:
            pass
        mode = snap.get(u"mode")
        if mode is None:
            continue
        hidden = list(snap.get(u"hidden") or [])
        for rb in rebars:
            if not isinstance(rb, Rebar):
                continue
            try:
                rb.SetPresentationMode(view, mode)
            except Exception:
                continue
            for i in hidden:
                try:
                    rb.SetBarHiddenStatus(view, int(i), True)
                except Exception:
                    pass
            n_ok += 1
    return n_ok


def apply_unobscured_to_rebars(doc, rebar_ids, view):
    """
    Aplica View Unobscured (+ sólido si aplica) a rebars ya existentes.

    Debe ejecutarse en una Transaction abierta. Devuelve cuántas barras
    quedaron con ``IsUnobscuredInView`` = True.
    """
    if doc is None or not rebar_ids or view is None:
        return 0
    try:
        view = doc.GetElement(view.Id)
    except Exception:
        return 0
    if not isinstance(view, View):
        return 0
    try:
        if view.IsTemplate:
            return 0
    except Exception:
        pass

    rebars = []
    for eid in rebar_ids:
        try:
            el = doc.GetElement(eid)
        except Exception:
            el = None
        if isinstance(el, Rebar):
            rebars.append(el)
    if not rebars:
        return 0

    try:
        doc.Regenerate()
    except Exception:
        pass

    try:
        from bimtools_rebar_3d_visibility import apply_reinforcement_unobscured_in_view

        apply_reinforcement_unobscured_in_view(
            doc, rebars, view, unobscured=True, solid_in_view=True
        )
    except Exception:
        for rb in rebars:
            try:
                rb.SetUnobscuredInView(view, True)
            except Exception:
                pass
            try:
                rb.SetSolidInView(view, True)
            except Exception:
                pass

    try:
        hide_ids = List[ElementId]()
        for rb in rebars:
            hide_ids.Add(rb.Id)
        if hide_ids.Count > 0:
            view.UnhideElements(hide_ids)
    except Exception:
        pass

    try:
        doc.Regenerate()
    except Exception:
        pass

    n_ok = 0
    for rb in rebars:
        try:
            # Re-fetch por si el proxy quedó obsoleto
            rb2 = doc.GetElement(rb.Id)
            if rb2 is None:
                continue
            if bool(rb2.IsUnobscuredInView(view)):
                n_ok += 1
            else:
                # Segundo intento directo
                try:
                    rb2.SetUnobscuredInView(view, True)
                except Exception:
                    pass
                try:
                    if bool(rb2.IsUnobscuredInView(view)):
                        n_ok += 1
                except Exception:
                    pass
        except Exception:
            pass
    return n_ok


def cut_mm_from_pick(doc, rebar, pick_xyz):
    """
    Proyecta un clic 3D sobre la centerline y devuelve la distancia en mm.

    Returns:
        (ok, message_or_None, cut_mm_or_None)
    """
    if doc is None or rebar is None or pick_xyz is None:
        return False, u"Parámetros incompletos.", None
    if not isinstance(rebar, Rebar):
        return False, u"Barra inválida.", None
    curves = _centerline_curves(rebar, 0, True, True)
    if not curves:
        curves = _centerline_curves(rebar, 0, True, False)
    if not curves:
        return False, u"No se pudo leer la centerline.", None
    proj = project_point_on_polyline(curves, pick_xyz)
    if not proj.get(u"ok"):
        return False, proj.get(u"message") or u"Proyección fallida.", None
    return True, None, internal_to_mm(proj[u"cut_dist"])


def divide_rebar_at_point(doc, rebar, pick_xyz, concrete_grade=None, view=None):
    """
    Divide ``rebar`` en el punto proyectado con traslape de tabla (1 corte).

    Preferir ``divide_rebar_at_cuts`` para N cortes desde la UI.
    """
    if doc is None or rebar is None or pick_xyz is None:
        return False, u"Parámetros incompletos.", None, None
    ok, err, _curves = check_eligibility(doc, rebar)
    if not ok:
        return False, err, None, None
    ok_p, msg_p, cut_mm = cut_mm_from_pick(doc, rebar, pick_xyz)
    if not ok_p or cut_mm is None:
        return False, msg_p or u"Proyección fallida.", None, None
    return divide_rebar_at_cuts(
        doc, rebar, [cut_mm], concrete_grade=concrete_grade, view=view
    )


def divide_rebar_at_cuts(
    doc,
    rebar,
    cuts_mm,
    concrete_grade=None,
    view=None,
    lap_mode=None,
    place_lap_dims=True,
    lap_dim_prefer_above=False,
    progress=None,
):
    """
    Divide ``rebar`` en N cortes (mm desde el inicio de la centerline) con traslape.

    ``place_lap_dims``: si False, coloca Detail de empalme sin cotas.
    ``lap_dim_prefer_above``: cotas hacia Up de la vista (sobre barras).
    ``progress``: ``DividirRebarProgress`` opcional (fases durante el proceso).

    Returns:
        (ok, mensaje, ids_nuevos_list_or_None, meta_or_None)
    """
    if doc is None or rebar is None:
        return False, u"Parámetros incompletos.", None, None
    if not cuts_mm:
        return False, u"Indique al menos un punto de corte.", None, None

    _progress_step(progress, u"Preparando")

    ok, err, _curves = check_eligibility(doc, rebar)
    if not ok:
        return False, err, None, None

    d_mm, lap_mm = lap_mm_for_rebar(doc, rebar, concrete_grade)
    if d_mm is None or lap_mm is None or lap_mm <= 0:
        return False, u"No se pudo obtener diámetro / traslape de la tabla.", None, None

    curves = _centerline_curves(rebar, 0, True, True)
    if not curves:
        curves = _centerline_curves(rebar, 0, True, False)
    if not curves:
        return False, u"No se pudo leer la centerline.", None, None

    cuts_ft = [mm_to_internal(float(c)) for c in cuts_mm]
    lap_ft = mm_to_internal(lap_mm)
    ok_s, msg_s, chains, meta = build_multi_split_curve_chains(
        curves, cuts_ft, lap_ft, lap_mode=lap_mode
    )
    if not ok_s or not chains:
        return False, msg_s or u"No se pudieron generar los tramos.", None, None

    fallback_host = doc.GetElement(rebar.GetHostId())
    # 1) GUID original → 2) rebars del conjunto → 3) hosts únicos (sin duplicados)
    conjunto_guid, peer_rebars, host_candidates = _host_candidates_from_original_rebar(
        doc, rebar
    )
    hosts_resolved = []
    n_host_reassigned = 0
    fallback_host_id = _element_id_int(
        fallback_host.Id if fallback_host is not None else None
    )
    bar_type = doc.GetElement(rebar.GetTypeId())
    try:
        style = rebar.Style
    except Exception:
        style = RebarStyle.Standard
    norm = _rebar_normal(rebar)
    hook_start = _hook_type(doc, rebar.GetHookTypeId(0))
    hook_end = _hook_type(doc, rebar.GetHookTypeId(1))
    try:
        so0 = rebar.GetHookOrientation(0)
        so1 = rebar.GetHookOrientation(1)
    except Exception:
        so0 = RebarHookOrientation.Right
        so1 = RebarHookOrientation.Left

    target_view = None
    target_view_id = None
    if view is not None and isinstance(view, View):
        try:
            if not view.IsTemplate and not isinstance(view, (ViewSheet, ViewSchedule)):
                target_view_id = view.Id
                target_view = doc.GetElement(target_view_id)
        except Exception:
            target_view = None
            target_view_id = None

    old_id = rebar.Id
    n_pos = _cantidad_posiciones(rebar)
    presentation_snaps = capture_rebar_presentation(doc, rebar, target_view)
    tag_infos = []
    try:
        from dividir_rebar_punto_tags import capture_rebar_tag_infos

        tag_infos = capture_rebar_tag_infos(doc, old_id)
    except Exception:
        tag_infos = []

    shape_targets = None
    shape_objs = []
    orig_shape_name = u""
    try:
        from dividir_rebar_punto_shapes import (
            find_rebar_shape_by_name,
            get_rebar_shape_name,
            target_shape_names_for_pieces,
        )

        orig_shape_name = get_rebar_shape_name(doc, rebar)
        shape_targets = target_shape_names_for_pieces(orig_shape_name, len(chains))
        if shape_targets:
            cache = {}
            for name in shape_targets:
                if name not in cache:
                    cache[name] = find_rebar_shape_by_name(doc, name)
            shape_objs = [cache.get(n) for n in shape_targets]
    except Exception:
        shape_targets = None
        shape_objs = [None] * len(chains)

    while len(shape_objs) < len(chains):
        shape_objs.append(None)

    n_pieces = len(chains)
    shape_info = {}
    new_ids = []
    layout_notes = []
    n_armadura_params = 0
    n_presentation = 0

    # Segmentos de detail [C−L/2, C+L/2] sobre la centerline original (antes de borrar).
    lap_segments = []
    lap_detail_info = {
        u"n_ok": 0,
        u"n_fail": 0,
        u"n_dims_ok": 0,
        u"n_dims_fail": 0,
        u"ids": [],
        u"dim_ids": [],
        u"errors": [],
        u"warning": None,
    }
    _place_lap_details = None
    _view_accepts_details = None
    _expand_lap_for_presentation = None
    try:
        from dividir_rebar_punto_lap_detail import (
            build_lap_segments_from_cuts,
            expand_lap_segments_for_presentation,
            place_lap_details_for_segments,
            view_accepts_detail_components,
        )

        _place_lap_details = place_lap_details_for_segments
        _view_accepts_details = view_accepts_detail_components
        _expand_lap_for_presentation = expand_lap_segments_for_presentation
        sorted_cuts_ft = list(meta.get(u"cuts_ft") or cuts_ft)
        lap_segments = build_lap_segments_from_cuts(
            curves,
            sorted_cuts_ft,
            lap_ft,
            _point_at_distance,
            lap_mode=meta.get(u"lap_mode"),
        )
        if sorted_cuts_ft and not lap_segments:
            lap_detail_info[u"n_fail"] = len(sorted_cuts_ft)
            lap_detail_info[u"errors"].append(
                u"No se pudieron construir segmentos de empalme en la centerline."
            )
    except Exception as ex_lap_prep:
        n_cuts_prep = len(list(meta.get(u"cuts_ft") or cuts_ft or []))
        lap_detail_info[u"n_fail"] = max(1, n_cuts_prep)
        lap_detail_info[u"errors"].append(_as_unicode(ex_lap_prep))
        lap_segments = []
        _expand_lap_for_presentation = None

    t = Transaction(doc, _TRANSACTION_NAME)
    try:
        _attach_rebar_outside_host_swallower(t)
    except Exception:
        pass
    t.Start()
    try:
        new_rebars = []
        _progress_step(progress, u"Creando tramos")
        for i, chain in enumerate(chains):
            is_first = i == 0
            is_last = i == n_pieces - 1
            h0 = hook_start if is_first else None
            h1 = hook_end if is_last else None
            o0 = so0 if is_first else RebarHookOrientation.Right
            o1 = so1 if is_last else RebarHookOrientation.Left
            host_i = _resolve_host_for_chain(
                chain, host_candidates, fallback_host
            )
            if host_i is None:
                t.RollBack()
                return (
                    False,
                    u"No hay host válido para el tramo {0}.".format(i + 1),
                    None,
                    None,
                )
            host_i_id = _element_id_int(host_i.Id)
            hosts_resolved.append(host_i_id)
            if fallback_host_id and host_i_id and host_i_id != fallback_host_id:
                n_host_reassigned += 1
            rb = _create_from_curves(
                doc,
                chain,
                host_i,
                norm,
                bar_type,
                style,
                h0,
                h1,
                o0,
                o1,
                rebar_shape=shape_objs[i] if i < len(shape_objs) else None,
            )
            if rb is None:
                t.RollBack()
                return False, u"CreateFromCurves devolvió None (tramo {0}).".format(i + 1), None, None
            new_rebars.append(rb)

        _progress_step(progress, u"Layout y parámetros")
        rule_src = _layout_rule_name(rebar)
        need_layout = n_pos > 1 or rule_src == u"MaximumSpacing"
        if need_layout:
            for i, rb_new in enumerate(new_rebars):
                ok_lay, err_lay = _copy_layout(rebar, rb_new)
                if ok_lay:
                    try:
                        doc.Regenerate()
                    except Exception:
                        pass
                    _copy_bar_included(rebar, rb_new, n_pos)
                else:
                    layout_notes.append(
                        u"T{0}: {1}".format(i + 1, err_lay or u"layout no copiado")
                    )

        for rb_new in new_rebars:
            n_armadura_params = max(
                n_armadura_params, int(copy_armadura_instance_parameters(rebar, rb_new))
            )

        try:
            from dividir_rebar_punto_shapes import apply_split_shape_rules_to_list

            shape_info = apply_split_shape_rules_to_list(doc, rebar, new_rebars) or {}
            try:
                doc.Regenerate()
            except Exception:
                pass
        except Exception as ex_sh:
            shape_info = {
                u"applied": False,
                u"original": orig_shape_name,
                u"errors": [_as_unicode(ex_sh)],
            }

        if shape_targets is not None and any(s is None for s in shape_objs):
            errs = list(shape_info.get(u"errors") or [])
            missing = sorted(
                set(
                    shape_targets[i]
                    for i, s in enumerate(shape_objs)
                    if s is None and i < len(shape_targets)
                )
            )
            errs.append(
                u"Regla {0}→{1} activa, pero falta RebarShape en el documento: {2}.".format(
                    orig_shape_name or u"?",
                    u"/".join(shape_targets),
                    u", ".join(missing) if missing else u"?",
                )
            )
            shape_info[u"errors"] = errs
            shape_info[u"applied"] = True
            shape_info[u"original"] = orig_shape_name or u""
            shape_info[u"targets"] = list(shape_targets)
            shape_info[u"target_a"] = shape_targets[0]
            shape_info[u"target_b"] = shape_targets[-1]

        _progress_step(progress, u"Presentación y empalmes")
        # Tras layout/shape + Regenerate: heredar Show Middle / FirstLast / etc.
        if presentation_snaps:
            try:
                doc.Regenerate()
            except Exception:
                pass
            n_presentation = apply_rebar_presentation(
                doc, new_rebars, presentation_snaps
            )

        new_ids = [rb.Id for rb in new_rebars]

        # Detail de empalme por corte (misma transacción = un solo Undo).
        if lap_segments and _place_lap_details is not None:
            detail_view = target_view
            try:
                if detail_view is None or (
                    _view_accepts_details is not None
                    and not _view_accepts_details(detail_view)
                ):
                    try:
                        av = doc.ActiveView
                        if av is not None and (
                            _view_accepts_details is None
                            or _view_accepts_details(av)
                        ):
                            detail_view = av
                    except Exception:
                        pass
                if detail_view is None or (
                    _view_accepts_details is not None
                    and not _view_accepts_details(detail_view)
                ):
                    lap_detail_info[u"n_fail"] = len(lap_segments)
                    lap_detail_info[u"errors"].append(
                        u"La vista activa no admite Detail Components "
                        u"(use planta, alzado o sección; no 3D ni lámina)."
                    )
                else:
                    segs_to_place = list(lap_segments)
                    # Desplazar empalme a barras representadas (Show Middle, etc.).
                    # Usar la original: aún existe y ya tiene PresentationMode.
                    if _expand_lap_for_presentation is not None:
                        try:
                            segs_to_place = _expand_lap_for_presentation(
                                rebar,
                                detail_view,
                                lap_segments,
                                source_bar_index=0,
                            ) or segs_to_place
                        except Exception:
                            segs_to_place = list(lap_segments)
                    placed = _place_lap_details(
                        doc,
                        detail_view,
                        segs_to_place,
                        place_dims=bool(place_lap_dims),
                        prefer_dims_above=bool(lap_dim_prefer_above),
                    ) or {}
                    lap_detail_info.update(placed)
            except Exception as ex_lap_place:
                lap_detail_info[u"n_fail"] = len(lap_segments)
                lap_detail_info[u"errors"].append(_as_unicode(ex_lap_place))

        _progress_step(progress, u"Finalizando división")
        doc.Delete(old_id)
        t.Commit()
    except Exception as ex:
        try:
            t.RollBack()
        except Exception:
            pass
        return False, _exception_text(ex), None, None

    n_unobscured = 0
    n_tags = 0
    n_mra = 0
    annotate_avisos = []
    used_default_tags = False
    _progress_step(progress, u"Visibilidad en vista")
    if target_view_id is not None:
        try:
            target_view = doc.GetElement(target_view_id)
        except Exception:
            target_view = None
        if isinstance(target_view, View):
            t2 = Transaction(doc, u"Arainco: View Unobscured tras división")
            t2.Start()
            try:
                n_unobscured = apply_unobscured_to_rebars(doc, new_ids, target_view)
                t2.Commit()
            except Exception:
                try:
                    t2.RollBack()
                except Exception:
                    pass
                n_unobscured = 0

    _progress_step(progress, u"Etiquetas y MRA")
    # Etiquetas + MRA «Recorrido Barras» (siempre; soft-fail).
    t3 = Transaction(doc, u"Arainco: Etiquetar y MRA tras división")
    t3.Start()
    try:
        from dividir_rebar_punto_tags import annotate_divided_rebars

        rebars2 = [doc.GetElement(eid) for eid in new_ids]
        ann_view = None
        if target_view_id is not None:
            try:
                ann_view = doc.GetElement(target_view_id)
            except Exception:
                ann_view = None
        if not isinstance(ann_view, View):
            ann_view = view if isinstance(view, View) else None
        ann = annotate_divided_rebars(
            doc, ann_view, rebars2, tag_infos=tag_infos
        ) or {}
        n_tags = int(ann.get(u"n_tags") or 0)
        n_mra = int(ann.get(u"n_mra") or 0)
        annotate_avisos = list(ann.get(u"avisos") or [])
        used_default_tags = bool(ann.get(u"used_default_tags"))
        t3.Commit()
    except Exception:
        try:
            t3.RollBack()
        except Exception:
            pass
        n_tags = 0
        n_mra = 0
        annotate_avisos = []
        used_default_tags = False

    meta = dict(meta or {})
    meta[u"diameter_mm"] = d_mm
    meta[u"lap_mm"] = lap_mm
    meta[u"concrete_grade"] = concrete_grade
    meta[u"cuts_mm"] = list(meta.get(u"cuts_mm") or [])
    meta[u"new_ids"] = [_element_id_int(eid) for eid in new_ids]
    meta[u"n_positions"] = n_pos
    meta[u"unobscured_count"] = int(n_unobscured)
    meta[u"unobscured"] = int(n_unobscured) > 0
    meta[u"presentation_applied"] = int(n_presentation)
    meta[u"presentation_modes"] = [
        s.get(u"mode_name") for s in (presentation_snaps or []) if s.get(u"mode_name")
    ]
    meta[u"armadura_params_copied"] = int(n_armadura_params)
    meta[u"tags_created"] = int(n_tags)
    meta[u"mra_created"] = int(n_mra)
    meta[u"annotate_avisos"] = list(annotate_avisos)
    meta[u"used_default_tags"] = bool(used_default_tags)
    meta[u"shape_rule"] = shape_info
    meta[u"conjunto_guid"] = conjunto_guid
    meta[u"n_peer_rebars"] = len(peer_rebars or [])
    meta[u"hosts_resolved"] = list(hosts_resolved)
    meta[u"n_host_candidates"] = len(host_candidates or [])
    meta[u"host_candidate_ids"] = [
        _element_id_int(h.Id) for h in (host_candidates or []) if h is not None
    ]
    meta[u"n_host_reassigned"] = int(n_host_reassigned)
    meta[u"fallback_host_id"] = fallback_host_id
    lap_ids_int = []
    for eid in lap_detail_info.get(u"ids") or []:
        try:
            lap_ids_int.append(_element_id_int(eid))
        except Exception:
            pass
    dim_ids_int = []
    for eid in lap_detail_info.get(u"dim_ids") or []:
        try:
            dim_ids_int.append(_element_id_int(eid))
        except Exception:
            pass
    meta[u"lap_detail"] = {
        u"n_ok": int(lap_detail_info.get(u"n_ok") or 0),
        u"n_fail": int(lap_detail_info.get(u"n_fail") or 0),
        u"n_dims_ok": int(lap_detail_info.get(u"n_dims_ok") or 0),
        u"n_dims_fail": int(lap_detail_info.get(u"n_dims_fail") or 0),
        u"ids": lap_ids_int,
        u"dim_ids": dim_ids_int,
        u"errors": list(lap_detail_info.get(u"errors") or []),
        u"warning": lap_detail_info.get(u"warning"),
        u"n_segments": len(lap_segments),
    }
    if len(meta.get(u"cuts_mm") or []) == 1:
        meta[u"cut_mm"] = meta[u"cuts_mm"][0]

    parts = []
    for i, eid in enumerate(new_ids):
        lm = (meta.get(u"lens_mm") or [0])[i] if i < len(meta.get(u"lens_mm") or []) else 0
        parts.append(u"Id {0} ({1:.0f} mm)".format(_element_id_int(eid), float(lm)))
    msg = (
        u"Dividida ø{0:.0f} mm · traslape {1:.0f} mm · {2} corte(s) → {3}"
    ).format(d_mm, lap_mm, len(meta.get(u"cuts_mm") or []), u", ".join(parts))
    layout_rule = _layout_rule_name(rebar)
    if n_pos > 1 or layout_rule == u"MaximumSpacing":
        msg = u"{0} · layout {1}".format(msg, layout_label(layout_rule, n_pos))
    if shape_info.get(u"applied"):
        finals = shape_info.get(u"finals") or []
        if shape_info.get(u"ok_all"):
            msg = u"{0} · Shape {1}→{2}".format(
                msg,
                shape_info.get(u"original") or u"?",
                u"/".join([f or u"?" for f in finals]) if finals else u"?",
            )
        else:
            targets = shape_info.get(u"targets") or [
                shape_info.get(u"target_a") or u"?",
                shape_info.get(u"target_b") or u"?",
            ]
            msg = u"{0} · Shape objetivo {1} (quedó {2})".format(
                msg,
                u"/".join([t or u"?" for t in targets]),
                u"/".join([f or u"?" for f in finals]) if finals else u"?",
            )
        if shape_info.get(u"errors"):
            msg = u"{0} [shape: {1}]".format(
                msg, u"; ".join(shape_info.get(u"errors") or [])
            )
    elif shape_info.get(u"errors"):
        msg = u"{0} [shape: {1}]".format(
            msg, u"; ".join(shape_info.get(u"errors") or [])
        )
    if n_armadura_params > 0:
        msg = u"{0} · Armadura_* ×{1}".format(msg, n_armadura_params)
    if n_host_reassigned > 0:
        msg = u"{0} · host reasignado ×{1}".format(msg, n_host_reassigned)
    if n_unobscured > 0:
        msg = u"{0} · Unobscured ×{1}".format(msg, n_unobscured)
    elif target_view_id is not None:
        msg = u"{0} · aviso: Unobscured no quedó activo en la vista".format(msg)
    if n_presentation > 0:
        mode_names = []
        seen_modes = set()
        for s in presentation_snaps or []:
            mn = s.get(u"mode_name") or u""
            if mn and mn not in seen_modes:
                seen_modes.add(mn)
                mode_names.append(mn)
        mode_txt = u"/".join(mode_names) if mode_names else u"heredada"
        msg = u"{0} · presentación {1}".format(msg, mode_txt)
    elif presentation_snaps:
        msg = u"{0} · aviso: no se pudo heredar PresentationMode".format(msg)
    if n_tags > 0:
        msg = u"{0} · etiquetas ×{1}".format(msg, n_tags)
    elif tag_infos:
        msg = u"{0} · aviso: no se recrearon etiquetas".format(msg)
    elif annotate_avisos:
        tag_av = None
        for av in annotate_avisos:
            try:
                al = _as_unicode(av).lower()
            except Exception:
                al = u""
            if u"etiqueta" in al:
                tag_av = av
                break
        if tag_av:
            msg = u"{0} · aviso: {1}".format(msg, tag_av)
    if n_mra > 0:
        msg = u"{0} · MRA recorrido ×{1}".format(msg, n_mra)
    else:
        mra_av = None
        for av in annotate_avisos:
            try:
                al = _as_unicode(av).lower()
            except Exception:
                al = u""
            if u"mra" in al or u"multi-rebar" in al or u"recorrido" in al:
                mra_av = av
                break
        if mra_av:
            msg = u"{0} · aviso: {1}".format(msg, mra_av)
        elif target_view_id is not None:
            msg = u"{0} · aviso: MRA «Recorrido Barras» no aplicado".format(msg)
    n_lap_ok = int(lap_detail_info.get(u"n_ok") or 0)
    n_lap_fail = int(lap_detail_info.get(u"n_fail") or 0)
    n_dims_ok = int(lap_detail_info.get(u"n_dims_ok") or 0)
    n_dims_fail = int(lap_detail_info.get(u"n_dims_fail") or 0)
    lap_errs = list(lap_detail_info.get(u"errors") or [])
    lap_warn = lap_detail_info.get(u"warning")
    if n_lap_ok > 0:
        msg = u"{0} · empalmes Detail ×{1}".format(msg, n_lap_ok)
        if n_dims_ok > 0:
            msg = u"{0} · acotados ×{1}".format(msg, n_dims_ok)
        if n_dims_fail > 0:
            dim_brief = None
            for e in lap_errs:
                try:
                    et = _as_unicode(e)
                except Exception:
                    et = u""
                if not et:
                    continue
                el = et.lower()
                if (
                    u"cota" in el
                    or u"left/right" in el
                    or u"acota" in el
                ):
                    dim_brief = et
                    break
            msg = u"{0} · aviso: {1} cota(s) de empalme fallaron{2}".format(
                msg,
                n_dims_fail,
                u" ({0})".format(dim_brief) if dim_brief else u"",
            )
    if n_lap_fail > 0:
        brief = lap_errs[0] if lap_errs else u"no colocados"
        if n_lap_ok > 0:
            msg = u"{0} · aviso: {1} empalme(s) fallaron ({2})".format(
                msg, n_lap_fail, brief
            )
        else:
            msg = u"{0} · aviso: empalmes no colocados ({1})".format(msg, brief)
    elif lap_warn and n_lap_ok > 0:
        msg = u"{0} · aviso empalme: {1}".format(msg, lap_warn)
    if layout_notes:
        msg = u"{0} (aviso: {1})".format(msg, u"; ".join(layout_notes))
    else:
        msg = u"{0}.".format(msg)
    return True, msg, new_ids, meta


def _apply_unobscured_in_view(doc, rebars, view):
    """Compat: delega en ``apply_unobscured_to_rebars``."""
    if not rebars:
        return False
    ids = []
    for rb in rebars:
        try:
            ids.append(rb.Id)
        except Exception:
            pass
    return apply_unobscured_to_rebars(doc, ids, view) > 0
