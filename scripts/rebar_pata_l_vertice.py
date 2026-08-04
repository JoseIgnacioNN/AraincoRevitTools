# -*- coding: utf-8 -*-
"""
Pata L en vértice de Structural Rebar (host losa).

Flujo:
1. Seleccionar un Rebar con host Floor.
2. Seleccionar un extremo de la barra (clic cerca del vértice).
3. Generar pata L geométrica en ese extremo con largo = espesor losa − 50 mm.

La barra nueva hereda la configuración de la original (layout, ganchos,
orientación/rotación, Style, parámetros de instancia, MoveBarInSet,
barras incluidas, **representación** PresentationMode/Unobscured/Solid,
**IndependentTag** y **MRA** MultiReferenceAnnotation) **excepto** el
RebarShape (cambia a la forma L, p. ej. «02»).

- Revit 2025+
- pyRevit / IronPython (``revit.Transaction`` + ``revit.TransactionGroup``)
- No edita in-place: CreateFromCurves (+ shape 02/03) y borra el original.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

import codecs
import datetime
import os
import traceback

import System
from Autodesk.Revit.DB import (
    BuiltInCategory,
    DimensionStyleType,
    ElementCategoryFilter,
    ElementClassFilter,
    ElementId,
    Floor,
    FilteredElementCollector,
    IndependentTag,
    MultiReferenceAnnotation,
    MultiReferenceAnnotationOptions,
    MultiReferenceAnnotationType,
    Options,
    Plane,
    Reference,
    SketchPlane,
    StorageType,
    TagMode,
    TagOrientation,
    Transaction,
    TransactionGroup,
    UnitTypeId,
    UnitUtils,
    View,
    View3D,
    ViewDetailLevel,
    XYZ,
)
from Autodesk.Revit.DB.Structure import (
    MultiplanarOption,
    Rebar,
    RebarStyle,
)
from System.Collections.Generic import List as ClrList
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

try:
    from Autodesk.Revit.UI.Selection import ObjectSnapTypes
except Exception:
    ObjectSnapTypes = None

_DIALOG_TITLE = u"Arainco: Pata L en vértice"
# Undo unificado (TransactionGroup.Assimilate)
_TRANSACTION_GROUP = u"Arainco: Pata L en vértice Rebar"
_TRANSACTION_NAME = u"Arainco: Pata L en vértice Rebar"
_TRANSACTION_SKETCH = u"Arainco: Plano de trabajo vista (Pata L vértice)"
_PATA_RESTA_MM = 50.0
_PATA_MIN_MM = 10.0
_PATA_FALLBACK_MM = 150.0
# Tol. para asociar el clic al extremo de la barra (mm)
_VERTEX_PICK_TOL_MM = 3000.0
_DIAG_DIALOG_LINES = 24


# ---------------------------------------------------------------------------
# Transacciones pyRevit 2025+ (TransactionGroup → un Undo; rollback limpio)
# ---------------------------------------------------------------------------


def _pyrevit_revit_mod():
    """``pyrevit.revit`` si está disponible (Transaction / TransactionGroup)."""
    try:
        from pyrevit import revit as _rv

        return _rv
    except Exception:
        return None


class _transaction_group_scope(object):
    """
    ``with revit.TransactionGroup(name, doc, assimilate=True)``.

    Excepción → ``RollBack`` del grupo. Éxito + assimilate → un solo Deshacer.
    """

    def __init__(self, doc, name, assimilate=True):
        self._doc = doc
        self._name = name
        self._assimilate = bool(assimilate)
        self._cm = None
        self._tg = None
        self._owns = False

    def __enter__(self):
        _rv = _pyrevit_revit_mod()
        if _rv is not None:
            self._cm = _rv.TransactionGroup(
                self._name, self._doc, assimilate=self._assimilate
            )
            return self._cm.__enter__()
        self._tg = TransactionGroup(self._doc, self._name)
        self._tg.Start()
        self._owns = True
        return self

    def __exit__(self, exc_type, exc, tb):
        if self._cm is not None:
            return self._cm.__exit__(exc_type, exc, tb)
        if not self._owns or self._tg is None:
            return False
        try:
            if exc_type is not None:
                try:
                    self._tg.RollBack()
                except Exception:
                    pass
            else:
                try:
                    if self._assimilate:
                        self._tg.Assimilate()
                    else:
                        self._tg.Commit()
                except Exception:
                    try:
                        self._tg.RollBack()
                    except Exception:
                        pass
                    raise
        finally:
            try:
                self._tg.Dispose()
            except Exception:
                pass
            self._tg = None
            self._owns = False
        return False


class _transaction_scope(object):
    """
    ``with revit.Transaction(name, doc)``.

    Excepción → RollBack (pyRevit o DB.Transaction), no deja el doc bloqueado.
    """

    def __init__(self, doc, name):
        self._doc = doc
        self._name = name
        self._cm = None
        self._txn = None
        self._owns = False

    def __enter__(self):
        _rv = _pyrevit_revit_mod()
        if _rv is not None:
            self._cm = _rv.Transaction(self._name, self._doc)
            return self._cm.__enter__()
        self._txn = Transaction(self._doc, self._name)
        self._txn.Start()
        self._owns = True
        return self

    def __exit__(self, exc_type, exc, tb):
        if self._cm is not None:
            return self._cm.__exit__(exc_type, exc, tb)
        if not self._owns or self._txn is None:
            return False
        try:
            if exc_type is not None:
                try:
                    self._txn.RollBack()
                except Exception:
                    pass
            else:
                try:
                    self._txn.Commit()
                except Exception:
                    try:
                        self._txn.RollBack()
                    except Exception:
                        pass
                    raise
        finally:
            try:
                self._txn.Dispose()
            except Exception:
                pass
            self._txn = None
            self._owns = False
        return False


def _collector_rebar_tags(doc):
    """
    IndependentTag de armadura: filtros nativos (categoría + clase).

    Evita dos barridos completos OfClass / OfCategory por separado.
    """
    if doc is None:
        return None
    try:
        return (
            FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_RebarTags)
            .OfClass(IndependentTag)
            .WhereElementIsNotElementType()
        )
    except Exception:
        try:
            return (
                FilteredElementCollector(doc)
                .OfClass(IndependentTag)
                .WhereElementIsNotElementType()
            )
        except Exception:
            return None


def _collector_mra(doc):
    if doc is None:
        return None
    try:
        return (
            FilteredElementCollector(doc)
            .OfClass(MultiReferenceAnnotation)
            .WhereElementIsNotElementType()
        )
    except Exception:
        try:
            return FilteredElementCollector(doc).OfClass(MultiReferenceAnnotation)
        except Exception:
            return None


def _collector_views(doc):
    if doc is None:
        return None
    try:
        return (
            FilteredElementCollector(doc)
            .OfClass(View)
            .WhereElementIsNotElementType()
        )
    except Exception:
        try:
            return FilteredElementCollector(doc).OfClass(View)
        except Exception:
            return None


def _as_unicode(val):
    if val is None:
        return u""
    try:
        return unicode(val)
    except NameError:
        try:
            return str(val)
        except Exception:
            return u""


def _exception_text(ex):
    if ex is None:
        return u""
    base = _as_unicode(ex)
    try:
        tb = traceback.format_exc()
        if tb and u"NoneType: None" not in tb and len(tb) > 40:
            # Solo colas cortas: el log completo ya tiene el stack si se pide
            return base
    except Exception:
        pass
    return base


class _DiagSession(object):
    """Instrumentación: consola pyRevit + archivo scripts/_diag_logs/."""

    def __init__(self, prefix=u"pata_l_vertice"):
        self._lines = []
        self._path = None
        self._prefix = prefix or u"pata_l_vertice"
        self._open_log()

    def _open_log(self):
        try:
            base = os.path.dirname(os.path.abspath(__file__))
            log_dir = os.path.join(base, u"_diag_logs")
            if not os.path.isdir(log_dir):
                os.makedirs(log_dir)
            stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            self._path = os.path.join(
                log_dir, u"{0}_{1}.log".format(self._prefix, stamp)
            )
            self.log(u"=== Arainco: Pata L en vértice — DIAG ===")
        except Exception as ex:
            self._path = None
            self._lines.append(
                u"[diag] No se pudo crear log: {0}".format(_exception_text(ex))
            )

    def log(self, msg):
        try:
            line = _as_unicode(msg)
        except Exception:
            line = u"?"
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        entry = u"[{0}] {1}".format(ts, line)
        self._lines.append(entry)
        try:
            print(entry)
        except Exception:
            pass
        if not self._path:
            return
        try:
            with codecs.open(self._path, "a", "utf-8") as fh:
                fh.write(entry + u"\n")
        except Exception:
            pass

    def step(self, name):
        self.log(u"--- {0} ---".format(name))

    def ex(self, label, ex):
        self.log(u"{0}: {1}".format(label, _exception_text(ex)))
        try:
            tb = traceback.format_exc()
            if tb and u"NoneType: None" not in tb:
                for ln in tb.strip().splitlines()[-12:]:
                    self.log(u"  | {0}".format(ln))
        except Exception:
            pass

    def path(self):
        return self._path

    def tail(self, n=None):
        n = n or _DIAG_DIALOG_LINES
        chunk = self._lines[-n:] if len(self._lines) > n else self._lines
        return u"\n".join(chunk)

    def summary_for_dialog(self, headline=None):
        parts = []
        if headline:
            parts.append(_as_unicode(headline))
        if self._path:
            parts.append(u"Log:\n{0}".format(self._path))
        parts.append(u"Últimas líneas:\n{0}".format(self.tail()))
        return u"\n\n".join(parts)


def _mm_from_internal(length_int):
    return float(
        UnitUtils.ConvertFromInternalUnits(float(length_int), UnitTypeId.Millimeters)
    )


def _internal_from_mm(mm):
    return float(UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters))


def _mostrar_aviso(uiapp, instruction, content=u""):
    try:
        from bimtools_instruction_dialog import show_message_dialog
        from revit_wpf_window_position import revit_main_hwnd

        hwnd = revit_main_hwnd(uiapp) if uiapp is not None else None
        if show_message_dialog(
            _DIALOG_TITLE,
            instruction=instruction,
            content=content,
            ok_text=u"Entendido",
            hwnd_revit=hwnd,
            uiapp=uiapp,
        ):
            return
    except Exception:
        pass
    msg = _as_unicode(instruction)
    c = _as_unicode(content).strip()
    if c:
        msg = u"{0}\n\n{1}".format(msg, c)
    try:
        TaskDialog.Show(_DIALOG_TITLE, msg)
    except Exception:
        pass


class _FiltroRebar(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, Rebar)

    def AllowReference(self, reference, position):
        return False


class _FiltroPuntoEnRebar(ISelectionFilter):
    """Solo la barra ya elegida; permite referencias para PointOnElement/Subelement."""

    def __init__(self, rebar_id):
        self._rebar_id = rebar_id

    def AllowElement(self, elem):
        if not isinstance(elem, Rebar):
            return False
        try:
            return elem.Id == self._rebar_id
        except Exception:
            return False

    def AllowReference(self, reference, position):
        return True


def _is_selection_cancelled(ex):
    if ex is None:
        return False
    if isinstance(ex, OperationCanceledException):
        return True
    try:
        name = type(ex).__name__ or u""
        if u"OperationCanceled" in name or u"Cancelled" in name:
            return True
    except Exception:
        pass
    msg = _as_unicode(ex).lower()
    return (u"cancel" in msg) or (u"abort" in msg)


def _global_point_from_ref(ref):
    if ref is None:
        return None
    for attr in (u"GlobalPoint", u"UVPoint"):
        try:
            pt = getattr(ref, attr, None)
            if pt is not None and hasattr(pt, u"X"):
                # UVPoint no es XYZ; solo GlobalPoint
                if attr == u"GlobalPoint":
                    return pt
        except Exception:
            pass
    try:
        return ref.GlobalPoint
    except Exception:
        return None


def _object_type_point_on_element():
    try:
        return ObjectType.PointOnElement
    except Exception:
        return None


def _object_type_subelement():
    try:
        return ObjectType.Subelement
    except Exception:
        return None


def _ensure_view_sketch_plane(doc, view):
    """Activa SketchPlane de la vista para que PickPoint pueda abrirse."""
    if doc is None or view is None:
        return False
    try:
        if view.SketchPlane is not None:
            return True
    except Exception:
        pass
    try:
        with _transaction_scope(doc, _TRANSACTION_SKETCH):
            plane = Plane.CreateByNormalAndOrigin(view.ViewDirection, view.Origin)
            sketch_plane = SketchPlane.Create(doc, plane)
            view.SketchPlane = sketch_plane
        return True
    except Exception:
        return False


def _is_free_form(rebar):
    try:
        if rebar.IsRebarFreeForm():
            return True
    except Exception:
        pass
    try:
        fn = getattr(rebar, u"IsRebarFreeForm", None)
        if fn is not None and callable(fn) and bool(fn()):
            return True
    except Exception:
        pass
    return False


def _centerline_curves(rebar, pos_idx=0):
    if rebar is None:
        return []
    for mpo_name in (
        u"IncludeAllMultiplanarCurves",
        u"IncludeOnlyPlanarCurves",
    ):
        mpo = getattr(MultiplanarOption, mpo_name, None)
        if mpo is None:
            continue
        try:
            raw = rebar.GetCenterlineCurves(False, False, False, mpo, int(pos_idx))
            if raw is not None and int(raw.Count) > 0:
                return [raw[i] for i in range(int(raw.Count))]
        except Exception:
            pass
    try:
        raw = rebar.GetCenterlineCurves(False, False, False)
        if raw is not None and int(raw.Count) > 0:
            return [raw[i] for i in range(int(raw.Count))]
    except Exception:
        pass
    return []


def _centerline_curves_for_tag(rebar, pos_idx=0):
    """
    Centerline para ancla de etiqueta (mismo criterio que
    ``56_DividirRebarPuntoTraslape`` → ``dividir_rebar_punto_tags``).

    GetCenterlineCurves(adjust=False, setHooks=True, setCranks=True, …)
    da el tramo físico/principal más estable para TagHeadPosition.
    """
    if rebar is None:
        return []
    mpo = getattr(MultiplanarOption, u"IncludeAllMultiplanarCurves", None)
    if mpo is not None:
        try:
            raw = rebar.GetCenterlineCurves(
                False, True, True, mpo, int(pos_idx)
            )
            if raw is not None and int(raw.Count) > 0:
                return [raw[i] for i in range(int(raw.Count))]
        except Exception:
            pass
    # Fallback al centerline genérico de la tool
    return _centerline_curves(rebar, pos_idx)


def _bar_index_for_tag(rebar, view=None):
    """Índice Show Middle del set (como 56): medio, no 0 por defecto en layout."""
    if rebar is None:
        return 0
    try:
        n = int(rebar.NumberOfBarPositions)
    except Exception:
        return 0
    if n <= 1:
        return 0
    try:
        # Unset / Middle típicos
        from Autodesk.Revit.DB.Structure import RebarPresentationMode

        mode = None
        try:
            if view is not None and hasattr(rebar, u"GetPresentationMode"):
                mode = rebar.GetPresentationMode(view)
        except Exception:
            mode = None
        if mode is not None:
            try:
                if mode == RebarPresentationMode.FirstLast:
                    return 0
                if mode == RebarPresentationMode.Select:
                    # primera incluida
                    for i in range(n):
                        try:
                            if rebar.DoesBarExistAtPosition(i) and (
                                not hasattr(rebar, u"IsBarExcluded")
                                or not rebar.IsBarExcluded(i)
                            ):
                                return i
                        except Exception:
                            continue
            except Exception:
                pass
    except Exception:
        pass
    return int(n // 2)


def _dedupe_xyz(points, tol_mm=1.0):
    out = []
    tol = _internal_from_mm(tol_mm)
    for p in points or []:
        if p is None:
            continue
        dup = False
        for q in out:
            try:
                if float(p.DistanceTo(q)) <= tol:
                    dup = True
                    break
            except Exception:
                continue
        if not dup:
            out.append(p)
    return out


def _ordered_vertices(curves):
    """Vértices a lo largo del centerline (inicio de cada tramo + fin final)."""
    pts = []
    if not curves:
        return pts
    for i, c in enumerate(curves):
        try:
            p0 = c.GetEndPoint(0)
            p1 = c.GetEndPoint(1)
        except Exception:
            continue
        if i == 0:
            pts.append(p0)
        pts.append(p1)
    return _dedupe_xyz(pts, tol_mm=2.0)


def _rebar_from_selection(uidoc):
    doc = uidoc.Document
    ids = uidoc.Selection.GetElementIds()
    if ids is None or ids.Count != 1:
        return None
    el = doc.GetElement(ids[0])
    return el if isinstance(el, Rebar) else None


def _pick_rebar(uidoc):
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            _FiltroRebar(),
            u"1/2 — Seleccione un Structural Rebar (host losa).",
        )
    except OperationCanceledException:
        return None
    except Exception:
        return None
    if ref is None:
        return None
    return uidoc.Document.GetElement(ref.ElementId)


def _try_pick_object(uidoc, object_type, filt, prompt):
    try:
        return uidoc.Selection.PickObject(object_type, filt, prompt), None
    except Exception as ex:
        if _is_selection_cancelled(ex):
            return None, u"cancel"
        return None, _as_unicode(ex)


def _pick_point_on_rebar(uidoc, rebar):
    """
    Punto sobre la barra (clic en el rebar).

    Preferir PointOnElement / Subelement / Element + GlobalPoint.
    PickPoint genérico exige plano de trabajo y suele fallar en vista 3D.
    Returns: (XYZ|None, err_msg|None). err_msg u"cancel" si el usuario cancela.
    """
    if uidoc is None or rebar is None:
        return None, u"Documento o barra inválidos."

    filt = _FiltroPuntoEnRebar(rebar.Id)
    prompt = (
        u"2/2 — Clic sobre la barra, cerca del extremo donde desea la pata L "
        u"(Esc cancela)."
    )

    attempts = []
    ot_poe = _object_type_point_on_element()
    if ot_poe is not None:
        attempts.append(ot_poe)
    ot_sub = _object_type_subelement()
    if ot_sub is not None:
        attempts.append(ot_sub)
    attempts.append(ObjectType.Element)

    last_err = u""
    for ot in attempts:
        ref, err = _try_pick_object(uidoc, ot, filt, prompt)
        if err == u"cancel":
            return None, u"cancel"
        if ref is not None:
            pt = _global_point_from_ref(ref)
            if pt is not None:
                return pt, None
            last_err = u"No se obtuvo GlobalPoint del clic."
        elif err:
            last_err = err

    # Último recurso: PickPoint con SketchPlane de la vista
    doc = uidoc.Document
    view = None
    try:
        view = uidoc.ActiveView
    except Exception:
        view = None
    if view is not None:
        _ensure_view_sketch_plane(doc, view)
    prompt_pt = (
        u"2/2 — Indique un punto cerca del extremo de la barra "
        u"(snap a extremos; Esc cancela)."
    )
    try:
        if ObjectSnapTypes is not None:
            try:
                snaps = (
                    ObjectSnapTypes.Endpoints
                    | ObjectSnapTypes.Nearest
                    | ObjectSnapTypes.Intersections
                )
                pt = uidoc.Selection.PickPoint(snaps, prompt_pt)
                if pt is not None:
                    return pt, None
            except Exception as ex:
                if _is_selection_cancelled(ex):
                    return None, u"cancel"
                last_err = _as_unicode(ex)
        pt = uidoc.Selection.PickPoint(prompt_pt)
        if pt is not None:
            return pt, None
    except Exception as ex:
        if _is_selection_cancelled(ex):
            return None, u"cancel"
        last_err = _as_unicode(ex)

    return (
        None,
        last_err
        or u"No se pudo capturar el punto. Clic directo sobre la barra visible "
        u"en la vista.",
    )


def _closest_end_vertex(vertices, point, max_tol_mm=None):
    """
    Elige el extremo del trazado (índice 0 o último) más cercano al clic.

    No considera vértices intermedios: un clic cerca de un extremo define la pata.
    """
    if not vertices or point is None or len(vertices) < 1:
        return None, None
    if max_tol_mm is None:
        max_tol_mm = _VERTEX_PICK_TOL_MM
    max_d = _internal_from_mm(max_tol_mm)
    candidates = [0]
    if len(vertices) > 1:
        candidates.append(len(vertices) - 1)
    best_i = None
    best_d = None
    for i in candidates:
        try:
            d = float(point.DistanceTo(vertices[i]))
        except Exception:
            continue
        if best_d is None or d < best_d:
            best_d = d
            best_i = i
    if best_i is None or best_d is None or best_d > max_d:
        return None, best_d
    return best_i, best_d


def _get_floor_host(doc, rebar):
    try:
        host = doc.GetElement(rebar.GetHostId())
    except Exception:
        host = None
    if isinstance(host, Floor):
        return host
    return None


def _pata_largo_mm(doc, floor):
    try:
        from rebar_extender_l_ganchos_135_rps import largo_pata_mm_desde_espesor_host

        return float(
            largo_pata_mm_desde_espesor_host(
                doc,
                floor,
                resta_mm=_PATA_RESTA_MM,
                fallback_mm=_PATA_FALLBACK_MM,
                min_largo_mm=_PATA_MIN_MM,
            )
        )
    except Exception:
        pass
    try:
        from area_rein_losa_sketch_pata import pata_largo_mm_floor

        return float(pata_largo_mm_floor(doc, floor))
    except Exception:
        return max(_PATA_MIN_MM, _PATA_FALLBACK_MM)


def _avg_vertex_z(vertices):
    if not vertices:
        return None
    zs = []
    for v in vertices:
        try:
            zs.append(float(v.Z))
        except Exception:
            pass
    if not zs:
        return None
    return sum(zs) / float(len(zs))


def _floor_mid_z(floor):
    try:
        bb = floor.get_BoundingBox(None)
        if bb is not None:
            return 0.5 * (float(bb.Min.Z) + float(bb.Max.Z))
    except Exception:
        pass
    return None


def _pata_dir_xyz(rebar, floor, vertices):
    """
    Dirección unitaria de la pata: hacia el interior de la losa.

    Inferior (Z < mitad losa) → +Z; superior → −Z.
    """
    bar_z = _avg_vertex_z(vertices)
    mid_z = _floor_mid_z(floor)
    if bar_z is not None and mid_z is not None:
        if bar_z >= mid_z:
            return XYZ(0.0, 0.0, -1.0)
        return XYZ(0.0, 0.0, 1.0)
    # Respaldo: normal shape-driven (si es vertical horizontalmente…) o +Z
    try:
        n = rebar.GetShapeDrivenAccessor().Normal
        if n is not None and float(n.GetLength()) > 1e-12:
            n = n.Normalize()
            if abs(float(n.Z)) > 0.7:
                # Normal de plano horizontal ≈ Z: pata entra en la losa
                if float(n.Z) >= 0.0:
                    return n
                return n.Negate()
    except Exception:
        pass
    return XYZ(0.0, 0.0, 1.0)


# Parámetros de instancia que no se copian (geometría / forma / derivados de layout).
_SKIP_PARAM_NAMES = frozenset(
    (
        u"Rebar Number",
        u"Bar Count",
        u"Quantity",
        u"Total Bar Length",
        u"Max Rebar Length",
        u"Bar Length",
        u"Actual Bar Length",
        u"Rebar Shape",
        u"Shape",
        u"Shape Image",
        u"Array Length",
        u"Maximum Spacing",
        u"Spacing",
        u"Number of Bar Positions",
        u"Layout Rule",
        u"Forma de armadura",
        u"Forma",
        u"Imagen de forma",
        u"Longitud de barra",
        u"Longitud total de barra",
        u"Cantidad",
        u"Espaciado",
        u"Espaciado máximo",
        u"Longitud de la matriz",
        u"Regla de disposición",
    )
)


def _param_value_tuple(p):
    """(storage_type_int, value) o None si no se puede leer."""
    if p is None:
        return None
    try:
        if not p.HasValue:
            return None
    except Exception:
        pass
    try:
        st = p.StorageType
    except Exception:
        return None
    try:
        if st == StorageType.String:
            return (int(st), p.AsString() or u"")
        if st == StorageType.Integer:
            return (int(st), int(p.AsInteger()))
        if st == StorageType.Double:
            return (int(st), float(p.AsDouble()))
        if st == StorageType.ElementId:
            return (int(st), p.AsElementId())
    except Exception:
        return None
    return None


def _set_param_value(dp, st_int, val):
    if dp is None or dp.IsReadOnly:
        return False
    try:
        if int(st_int) == int(StorageType.String):
            dp.Set(val if val is not None else u"")
            return True
        if int(st_int) == int(StorageType.Integer):
            dp.Set(int(val))
            return True
        if int(st_int) == int(StorageType.Double):
            dp.Set(float(val))
            return True
        if int(st_int) == int(StorageType.ElementId):
            dp.Set(val)
            return True
    except Exception:
        return False
    return False


def _snapshot_hooks(rebar):
    hooks = []
    for end in (0, 1):
        item = {u"end": end, u"type_id": None, u"orient": None, u"rot_deg": None}
        try:
            item[u"type_id"] = rebar.GetHookTypeId(int(end))
        except Exception:
            pass
        try:
            item[u"orient"] = rebar.GetHookOrientation(int(end))
        except Exception:
            pass
        for meth in (
            u"GetTerminationRotationAngle",
            u"GetHookRotationAngle",
        ):
            try:
                fn = getattr(rebar, meth, None)
                if fn is not None and callable(fn):
                    item[u"rot_deg"] = float(fn(int(end)))
                    break
            except Exception:
                pass
        hooks.append(item)
    return hooks


def _apply_hooks(dst, hooks):
    if dst is None or not hooks:
        return
    for item in hooks:
        end = int(item.get(u"end", 0))
        tid = item.get(u"type_id")
        if tid is not None:
            try:
                dst.SetHookTypeId(end, tid)
            except Exception:
                pass
        orient = item.get(u"orient")
        if orient is not None:
            for meth in (u"SetTerminationOrientation", u"SetHookOrientation"):
                try:
                    fn = getattr(dst, meth, None)
                    if fn is not None and callable(fn):
                        fn(end, orient)
                        break
                except Exception:
                    pass
        rot = item.get(u"rot_deg")
        if rot is not None:
            for meth in (
                u"SetTerminationRotationAngle",
                u"SetHookRotationAngle",
            ):
                try:
                    fn = getattr(dst, meth, None)
                    if fn is not None and callable(fn):
                        fn(end, float(rot))
                        break
                except Exception:
                    pass


def _snapshot_instance_params(rebar):
    out = []
    try:
        params = rebar.Parameters
    except Exception:
        return out
    if params is None:
        return out
    for p in params:
        if p is None:
            continue
        try:
            if p.IsReadOnly:
                continue
        except Exception:
            continue
        try:
            name = p.Definition.Name
        except Exception:
            continue
        if not name or name in _SKIP_PARAM_NAMES:
            continue
        # Evitar shape-driven / A B C de tramos numéricos cortos típicos de forma
        try:
            if len(name) <= 2 and name.isalpha() and name.isupper():
                continue
        except Exception:
            pass
        val = _param_value_tuple(p)
        if val is None:
            continue
        out.append((name, val[0], val[1]))
    return out


def _apply_instance_params(dst, param_rows):
    if dst is None or not param_rows:
        return 0
    n_ok = 0
    for name, st_int, val in param_rows:
        try:
            dp = dst.LookupParameter(name)
        except Exception:
            dp = None
        if dp is None:
            continue
        if _set_param_value(dp, st_int, val):
            n_ok += 1
    return n_ok


def _cantidad_posiciones(rebar):
    try:
        n = int(rebar.Quantity)
        if n > 0:
            return n
    except Exception:
        pass
    try:
        n = int(rebar.NumberOfBarPositions)
        if n > 0:
            return n
    except Exception:
        pass
    return 1


def _snapshot_moved_bar_transforms(rebar):
    out = []
    n = _cantidad_posiciones(rebar)
    for i in range(max(1, n)):
        mt = None
        try:
            mt = rebar.GetMovedBarTransform(int(i))
        except Exception:
            mt = None
        out.append(mt)
    return out


def _apply_moved_bar_transforms(dst, transforms):
    if dst is None or not transforms:
        return
    n = min(len(transforms), _cantidad_posiciones(dst))
    for i in range(n):
        mt = transforms[i]
        if mt is None:
            continue
        try:
            if bool(mt.IsIdentity):
                continue
        except Exception:
            pass
        try:
            dst.MoveBarInSet(int(i), mt)
        except Exception:
            try:
                acc = dst.GetShapeDrivenAccessor()
                if acc is not None and hasattr(acc, u"MoveBarInSet"):
                    acc.MoveBarInSet(int(i), mt)
            except Exception:
                pass


def _snapshot_bars_included(rebar):
    flags = []
    n = _cantidad_posiciones(rebar)
    for i in range(max(1, n)):
        try:
            flags.append(bool(rebar.IsBarIncluded(int(i))))
        except Exception:
            flags.append(True)
    return flags


def _apply_bars_included(dst, flags):
    if dst is None or not flags:
        return
    n = min(len(flags), _cantidad_posiciones(dst))
    for i in range(n):
        try:
            dst.SetBarIncluded(bool(flags[i]), int(i))
        except Exception:
            pass


def _element_id_int(eid):
    """ElementId / LinkElementId → int, o None.

    Revit 2025+: preferir ``ElementId.Value`` (IntegerValue deprecado).
    """
    if eid is None:
        return None
    try:
        if eid == ElementId.InvalidElementId:
            return None
    except Exception:
        pass
    # ElementId: Value (2025+) luego IntegerValue (compat)
    for attr in (u"Value", u"IntegerValue"):
        try:
            v = getattr(eid, attr, None)
            if v is not None:
                return int(v)
        except Exception:
            continue
    # LinkElementId (API reciente en GetTagged*)
    for attr in (u"HostElementId", u"LinkedElementId"):
        try:
            sub = getattr(eid, attr, None)
            if sub is not None:
                n = _element_id_int(sub)
                if n is not None:
                    return n
        except Exception:
            continue
    try:
        return int(eid)
    except Exception:
        return None


def _iter_id_collection(ids):
    if ids is None:
        return
    try:
        for x in ids:
            yield x
        return
    except Exception:
        pass
    try:
        n = int(ids.Count)
    except Exception:
        try:
            n = int(ids.Size)
        except Exception:
            return
    for i in range(n):
        try:
            yield ids.get_Item(i)
        except Exception:
            try:
                yield ids[i]
            except Exception:
                pass


def _tag_links_rebar(tag, rebar_id_int, diag=None, log_match_detail=False):
    if tag is None or rebar_id_int is None:
        return False
    want = int(rebar_id_int)
    found_ids = []
    for gname, getter in (
        (u"GetTaggedLocalElementIds", lambda: tag.GetTaggedLocalElementIds()),
        (u"GetTaggedElementIds", lambda: tag.GetTaggedElementIds()),
    ):
        try:
            ids = getter()
        except Exception as ex:
            if log_match_detail and diag is not None:
                diag.ex(u"tag_links.{0}".format(gname), ex)
            ids = None
        for leid in _iter_id_collection(ids):
            n = _element_id_int(leid)
            if n is not None:
                found_ids.append((gname, n, type(leid).__name__))
            if n == want:
                if log_match_detail and diag is not None:
                    diag.log(
                        u"tag_links MATCH via {0} id={1} type={2}".format(
                            gname, n, type(leid).__name__
                        )
                    )
                return True
    try:
        refs = tag.GetTaggedReferences()
        for r in _iter_id_collection(refs):
            try:
                n = _element_id_int(r.ElementId)
                if n is not None:
                    found_ids.append((u"GetTaggedReferences", n, u"Reference"))
                if n == want:
                    if log_match_detail and diag is not None:
                        diag.log(u"tag_links MATCH via GetTaggedReferences id={0}".format(n))
                    return True
            except Exception:
                pass
    except Exception as ex:
        if log_match_detail and diag is not None:
            diag.ex(u"tag_links.GetTaggedReferences", ex)
    try:
        n = _element_id_int(tag.TaggedLocalElementId)
        if n is not None:
            found_ids.append((u"TaggedLocalElementId", n, u"prop"))
        if n == want:
            if log_match_detail and diag is not None:
                diag.log(u"tag_links MATCH via TaggedLocalElementId id={0}".format(n))
            return True
    except Exception:
        pass
    if log_match_detail and diag is not None:
        diag.log(
            u"tag_links NO match want={0} found={1}".format(
                want, found_ids[:8] if found_ids else u"[]"
            )
        )
    return False


def _tag_belongs_to_mra(doc, tag, diag=None):
    """True si la etiqueta es dependiente de un MultiReferenceAnnotation (MRA)."""
    if tag is None:
        return False
    try:
        mid = tag.MultiReferenceAnnotationId
        n = _element_id_int(mid)
        if n is None:
            return False
        if doc is not None:
            try:
                el = doc.GetElement(mid)
                if el is None:
                    if diag is not None:
                        diag.log(
                            u"tag MRA id={0} resuelve Element=None → no MRA".format(n)
                        )
                    return False
                if isinstance(el, MultiReferenceAnnotation):
                    return True
                if diag is not None:
                    diag.log(
                        u"tag MultiReferenceAnnotationId={0} no es MRA (es {1})".format(
                            n, type(el).__name__
                        )
                    )
                return False
            except Exception as ex:
                if diag is not None:
                    diag.ex(u"tag_belongs_to_mra.GetElement", ex)
        return True
    except Exception as ex:
        if diag is not None:
            diag.ex(u"tag_belongs_to_mra", ex)
        return False


def _pack_tag_info(doc, tag, rebar=None, diag=None):
    """
    Captura etiquetado del rebar origen (patrón 56_DividirRebarPuntoTraslape):

    - familia + type seed
    - orient / leader / rotation
    - ancla en mitad del segmento principal + ``head_offset`` relativo
    """
    info = {
        u"type_id": None,
        u"type_name": u"",
        u"family_id": None,
        u"family_name": u"",
        u"view_id": None,
        u"head": None,
        u"head_offset": None,
        u"anchor": None,
        u"bar_index": 0,
        u"orient": TagOrientation.Horizontal,
        u"leader": True,
        u"rotation": None,
        u"tag_id": None,
    }
    try:
        info[u"type_id"] = tag.GetTypeId()
    except Exception as ex:
        if diag is not None:
            diag.ex(u"pack_tag.type_id", ex)
    try:
        info[u"view_id"] = tag.OwnerViewId
    except Exception as ex:
        if diag is not None:
            diag.ex(u"pack_tag.view_id", ex)
    try:
        info[u"head"] = tag.TagHeadPosition
    except Exception as ex:
        if diag is not None:
            diag.ex(u"pack_tag.head", ex)
    try:
        info[u"orient"] = tag.TagOrientation
    except Exception:
        info[u"orient"] = TagOrientation.Horizontal
    try:
        info[u"leader"] = bool(tag.HasLeader)
    except Exception:
        info[u"leader"] = True
    try:
        info[u"rotation"] = float(tag.RotationAngle)
    except Exception:
        info[u"rotation"] = None
    try:
        info[u"tag_id"] = tag.Id
    except Exception:
        info[u"tag_id"] = None

    # Familia del tipo de etiqueta origen (clave de herencia)
    fam = None
    type_el = None
    try:
        if doc is not None and info.get(u"type_id") is not None:
            type_el = doc.GetElement(info[u"type_id"])
    except Exception:
        type_el = None
    if type_el is not None:
        try:
            info[u"type_name"] = _as_unicode(getattr(type_el, u"Name", u"") or u"")
        except Exception:
            pass
        try:
            fam = getattr(type_el, u"Family", None)
        except Exception:
            fam = None
    if fam is None and doc is not None:
        try:
            from rebar_tag_shape_sync_core import family_of_independent_tag

            fam = family_of_independent_tag(doc, tag)
        except Exception:
            fam = _family_of_tag_type(doc, info.get(u"type_id"))
    if fam is not None:
        try:
            info[u"family_id"] = fam.Id
            info[u"family_name"] = _as_unicode(getattr(fam, u"Name", u"") or u"")
        except Exception:
            pass

    # Ancla + offset (como capture_rebar_tag_infos de 56)
    view = None
    try:
        if doc is not None and info.get(u"view_id") is not None:
            view = doc.GetElement(info[u"view_id"])
    except Exception:
        view = None
    if rebar is not None:
        try:
            bi = _bar_index_for_tag(rebar, view)
            info[u"bar_index"] = int(bi)
            anchor = _midpoint_segmento_principal_rebar(
                rebar, view=None, diag=None, bar_index=bi
            )
            info[u"anchor"] = anchor
            head = info.get(u"head")
            if head is not None and anchor is not None:
                try:
                    info[u"head_offset"] = XYZ(
                        float(head.X) - float(anchor.X),
                        float(head.Y) - float(anchor.Y),
                        float(head.Z) - float(anchor.Z),
                    )
                except Exception:
                    info[u"head_offset"] = None
        except Exception as ex:
            if diag is not None:
                diag.ex(u"pack_tag.anchor/offset", ex)

    if info.get(u"type_id") is None and info.get(u"family_id") is None:
        if diag is not None:
            diag.log(
                u"pack_tag RECHAZADO tag={0}: sin type ni family".format(
                    _element_id_int(info.get(u"tag_id"))
                )
            )
        return None
    if diag is not None:
        try:
            hx = float(info[u"head"].X) if info.get(u"head") is not None else None
            hy = float(info[u"head"].Y) if info.get(u"head") is not None else None
            hz = float(info[u"head"].Z) if info.get(u"head") is not None else None
        except Exception:
            hx = hy = hz = None
        diag.log(
            u"pack_tag OK tag={0} type={1} «{2}» family={3} «{4}» "
            u"view={5} leader={6} bar_idx={7} head=({8},{9},{10})".format(
                _element_id_int(info.get(u"tag_id")),
                _element_id_int(info.get(u"type_id")),
                info.get(u"type_name") or u"",
                _element_id_int(info.get(u"family_id")),
                info.get(u"family_name") or u"",
                _element_id_int(info.get(u"view_id")),
                info.get(u"leader"),
                info.get(u"bar_index"),
                hx,
                hy,
                hz,
            )
        )
    return info


def _snapshot_independent_tags(doc, rebar, diag=None):
    """IndependentTag del rebar (excluye las hijas de MRA)."""
    out = []
    seen = set()
    rid = _element_id_int(rebar.Id) if rebar is not None else None
    if diag is not None:
        diag.step(u"SNAPSHOT IndependentTag rebar={0}".format(rid))
    if doc is None or rebar is None or rid is None:
        if diag is not None:
            diag.log(u"snapshot tags: doc/rebar/rid inválido")
        return out

    stats = {
        u"dep_class": 0,
        u"dep_cat": 0,
        u"skip_mra": 0,
        u"skip_pack": 0,
        u"ok": 0,
        u"coll_class_scan": 0,
        u"coll_cat_scan": 0,
        u"coll_link_match": 0,
    }

    def _add_tag(tag, source):
        if tag is None or not isinstance(tag, IndependentTag):
            return
        tid = _element_id_int(getattr(tag, u"Id", None))
        if _tag_belongs_to_mra(doc, tag, diag=diag):
            stats[u"skip_mra"] += 1
            if diag is not None:
                diag.log(
                    u"  skip MRA tag id={0} source={1}".format(tid, source)
                )
            return
        if tid is not None and tid in seen:
            if diag is not None:
                diag.log(u"  skip duplicate tag id={0}".format(tid))
            return
        info = _pack_tag_info(doc, tag, rebar=rebar, diag=diag)
        if info is None:
            stats[u"skip_pack"] += 1
            return
        if tid is not None:
            seen.add(tid)
        out.append(info)
        stats[u"ok"] += 1
        if diag is not None:
            diag.log(u"  ADD tag id={0} source={1}".format(tid, source))

    # 1) Dependientes del Rebar
    for filt_name, filt in (
        (u"ElementClassFilter(IndependentTag)", ElementClassFilter(IndependentTag)),
        (
            u"ElementCategoryFilter(OST_RebarTags)",
            ElementCategoryFilter(BuiltInCategory.OST_RebarTags),
        ),
    ):
        try:
            deps = rebar.GetDependentElements(filt)
            n_deps = 0
            for did in _iter_id_collection(deps):
                n_deps += 1
                try:
                    el = doc.GetElement(did)
                except Exception as ex:
                    if diag is not None:
                        diag.ex(u"GetDependentElements GetElement", ex)
                    continue
                _add_tag(el, filt_name)
            if diag is not None:
                diag.log(u"GetDependentElements[{0}] count={1}".format(filt_name, n_deps))
            if u"Class" in filt_name:
                stats[u"dep_class"] = n_deps
            else:
                stats[u"dep_cat"] = n_deps
        except Exception as ex:
            if diag is not None:
                diag.ex(u"GetDependentElements[{0}]".format(filt_name), ex)

    # 2) Collector nativo (categoría OST_RebarTags ∩ IndependentTag)
    #    Un solo barrido API en lugar de OfClass + OfCategory separados.
    try:
        coll = _collector_rebar_tags(doc)
        n_scan = 0
        n_match = 0
        if coll is not None:
            for tag in coll:
                n_scan += 1
                if not isinstance(tag, IndependentTag):
                    continue
                linked = _tag_links_rebar(
                    tag,
                    rid,
                    diag=diag,
                    log_match_detail=False,
                )
                if linked:
                    n_match += 1
                    _add_tag(tag, u"OfCategory+OfClass(RebarTags)")
        if diag is not None:
            diag.log(
                u"Collector[RebarTags∩IndependentTag] scan={0} link_match={1}".format(
                    n_scan, n_match
                )
            )
        stats[u"coll_class_scan"] = n_scan
        stats[u"coll_link_match"] = n_match
        stats[u"coll_cat_scan"] = n_scan
    except Exception as ex:
        if diag is not None:
            diag.ex(u"collector rebar tags", ex)

    if diag is not None:
        diag.log(u"snapshot tags RESULT count={0} stats={1}".format(len(out), stats))
        if not out:
            try:
                sample_n = 0
                coll_s = _collector_rebar_tags(doc)
                if coll_s is not None:
                    for tag in coll_s:
                        if not isinstance(tag, IndependentTag):
                            continue
                        sample_n += 1
                        if sample_n > 5:
                            break
                        tid = _element_id_int(tag.Id)
                        diag.log(u"sample tag id={0} MRA={1}".format(
                            tid,
                            _tag_belongs_to_mra(doc, tag, diag=None),
                        ))
                        _tag_links_rebar(tag, rid, diag=diag, log_match_detail=True)
            except Exception as ex:
                diag.ex(u"sample tags", ex)
    return out


def _project_point_to_view(p, view):
    if p is None or view is None:
        return p
    try:
        vd = view.ViewDirection
        if vd is None or float(vd.GetLength()) < 1e-12:
            return p
        vd = vd.Normalize()
        vo = view.Origin
        if vo is None:
            return p
        d = float((p - vo).DotProduct(vd))
        return p - vd.Multiply(d)
    except Exception:
        return p


def _referencias_tag_rebar(doc, rebar, view=None, diag=None, preferred_bar_index=None):
    """
    Referencias para IndependentTag.Create — orden de 56_Dividir:

    1) GetReferenceToBarPosition (barra representada)
    2) subelementos
    3) extremos del set
    4) Reference(rebar)

    Sin geometry collector masivo (en Pata L generaba 27 refs y
    anclajes difíciles de ver).
    """
    refs = []
    seen = set()
    sources = []

    def _add(r, src):
        if r is None:
            return
        try:
            key = r.ConvertToStableRepresentation(doc)
        except Exception:
            try:
                key = u"{0}".format(_element_id_int(r.ElementId))
            except Exception:
                key = id(r)
        if key in seen:
            return
        seen.add(key)
        refs.append(r)
        sources.append(src)

    def _ref_at(idx):
        try:
            if hasattr(rebar, u"GetReferenceToBarPosition"):
                return rebar.GetReferenceToBarPosition(int(idx))
        except Exception:
            pass
        try:
            if hasattr(rebar, u"GetReferenceForBarPosition"):
                return rebar.GetReferenceForBarPosition(int(idx))
        except Exception:
            pass
        return None

    if preferred_bar_index is None:
        preferred_bar_index = _bar_index_for_tag(rebar, view)
    try:
        bi = int(preferred_bar_index)
    except Exception:
        bi = 0

    _add(_ref_at(bi), u"BarPosition_{0}".format(bi))

    try:
        subs = rebar.GetSubelements() if hasattr(rebar, u"GetSubelements") else None
        if subs:
            for sub in subs:
                if sub is None:
                    continue
                try:
                    if hasattr(sub, u"GetReference"):
                        _add(sub.GetReference(), u"subelement")
                except Exception:
                    pass
    except Exception as ex:
        if diag is not None:
            diag.ex(u"refs.GetSubelements", ex)

    try:
        npos = int(rebar.NumberOfBarPositions)
    except Exception:
        npos = 0
    if npos > 0:
        for idx in (0, max(0, npos - 1)):
            if int(idx) == bi:
                continue
            _add(_ref_at(idx), u"BarPosition_{0}".format(idx))
    try:
        _add(Reference(rebar), u"Reference(rebar)")
    except Exception as ex:
        if diag is not None:
            diag.ex(u"ref.Reference(rebar)", ex)

    if diag is not None:
        diag.log(
            u"refs TOTAL n={0} preferred_bar={1} sources={2}".format(
                len(refs), bi, u",".join(sources[:8]) if sources else u"-"
            )
        )
    return refs


def _crear_etiqueta_rebar(
    doc,
    view,
    rebar,
    type_id,
    head,
    orient,
    add_leader,
    diag=None,
    rotation=None,
    bar_index=None,
):
    """
    Crea IndependentTag como en 56:

    - tipo ya resuelto por shape (llamador)
    - head en mitad tramo principal (+ offset opcional)
    - refs simples por posición de barra
    - fuerza TagOrientation + TagHeadPosition tras Create
    """
    if diag is not None:
        diag.step(
            u"CREATE IndependentTag type={0} view={1} rebar={2}".format(
                _element_id_int(type_id),
                _element_id_int(getattr(view, u"Id", None)),
                _element_id_int(getattr(rebar, u"Id", None)),
            )
        )
    if view is None or type_id is None or rebar is None or doc is None:
        msg = u"Datos incompletos para etiqueta."
        if diag is not None:
            diag.log(msg)
        return None, msg
    try:
        if bool(view.IsTemplate) or isinstance(view, View3D):
            msg = u"Vista no admite IndependentTag (plantilla/3D)."
            if diag is not None:
                diag.log(msg)
            return None, msg
    except Exception:
        pass

    if bar_index is None:
        bar_index = _bar_index_for_tag(rebar, view)
    try:
        bi = int(bar_index)
    except Exception:
        bi = 0

    # Head crudo (sin proyectar) — patrón 56
    p = head
    if p is None:
        p = _midpoint_segmento_principal_rebar(
            rebar, view=view, diag=diag, bar_index=bi
        )
    if p is None:
        try:
            bb = rebar.get_BoundingBox(view) or rebar.get_BoundingBox(None)
            if bb is not None:
                p = XYZ(
                    (bb.Min.X + bb.Max.X) * 0.5,
                    (bb.Min.Y + bb.Max.Y) * 0.5,
                    (bb.Min.Z + bb.Max.Z) * 0.5,
                )
        except Exception:
            p = None
    if p is None:
        msg = u"Sin punto de inserción para etiqueta."
        if diag is not None:
            diag.log(msg)
        return None, msg
    if diag is not None:
        try:
            diag.log(
                u"insert point=({0:.4f},{1:.4f},{2:.4f}) leader={3} bar={4}".format(
                    float(p.X), float(p.Y), float(p.Z), bool(add_leader), bi
                )
            )
        except Exception:
            pass

    try:
        sym = doc.GetElement(type_id)
        if sym is not None and hasattr(sym, u"IsActive") and not sym.IsActive:
            sym.Activate()
        if diag is not None and sym is not None:
            diag.log(
                u"tag type «{0}» active={1}".format(
                    _as_unicode(getattr(sym, u"Name", u"")),
                    getattr(sym, u"IsActive", u"?"),
                )
            )
    except Exception as ex:
        if diag is not None:
            diag.ex(u"Activate tag type", ex)

    if orient is None:
        orient = TagOrientation.Horizontal

    refs = _referencias_tag_rebar(
        doc, rebar, view, diag=diag, preferred_bar_index=bi
    )
    if not refs:
        msg = u"Sin referencia API en rebar para etiquetar."
        if diag is not None:
            diag.log(msg)
        return None, msg

    last_ex = u""
    for ri, ref in enumerate(refs):
        tag = None
        try:
            tag = IndependentTag.Create(
                doc, type_id, view.Id, ref, bool(add_leader), orient, p
            )
        except Exception as ex:
            last_ex = _as_unicode(ex)
            tag = None
            if diag is not None and ri < 4:
                diag.log(u"FAIL Create A ref#{0}: {1}".format(ri, last_ex))
        if tag is None:
            try:
                tag = IndependentTag.Create(
                    doc,
                    view.Id,
                    ref,
                    bool(add_leader),
                    TagMode.TM_ADDBY_CATEGORY,
                    orient,
                    p,
                )
                if tag is not None:
                    try:
                        tag.ChangeTypeId(type_id)
                    except Exception:
                        try:
                            tag.SetTypeId(type_id)
                        except Exception:
                            pass
            except Exception as ex:
                last_ex = _as_unicode(ex)
                tag = None
                if diag is not None and ri < 4:
                    diag.log(u"FAIL Create B ref#{0}: {1}".format(ri, last_ex))
        if tag is None:
            continue
        # Forzar como 56 (Create a veces reescribe)
        try:
            tag.TagOrientation = orient
        except Exception:
            pass
        _force_tag_head_position(tag, p, diag=diag, label=u"TagHead post-Create")
        if rotation is not None:
            try:
                tag.RotationAngle = float(rotation)
            except Exception:
                pass
        if diag is not None:
            diag.log(
                u"OK Create tagId={0} ref#{1}".format(
                    _element_id_int(tag.Id), ri
                )
            )
        return tag, None

    msg = last_ex or u"IndependentTag.Create falló con todas las referencias."
    if diag is not None:
        diag.log(u"CREATE FAILED: {0}".format(msg))
    return None, msg


def _family_of_tag_type(doc, type_id):
    if doc is None or type_id is None:
        return None
    try:
        el = doc.GetElement(type_id)
    except Exception:
        return None
    if el is None:
        return None
    try:
        fam = getattr(el, u"Family", None)
        if fam is not None:
            return fam
    except Exception:
        pass
    try:
        # tipo de IndependentTag → FamilySymbol
        tid = el.GetTypeId() if hasattr(el, u"GetTypeId") else None
        if tid is not None:
            el2 = doc.GetElement(tid)
            if el2 is not None and getattr(el2, u"Family", None) is not None:
                return el2.Family
    except Exception:
        pass
    return None


def _resolve_tag_family(doc, family_id=None, family_name=None, seed_type_id=None):
    """Obtiene la Family de etiqueta capturada (id, nombre o type seed)."""
    if doc is None:
        return None
    fam = None
    if family_id is not None:
        try:
            fam = doc.GetElement(family_id)
        except Exception:
            fam = None
        if fam is not None:
            try:
                # debe ser Family, no symbol
                if hasattr(fam, u"GetFamilySymbolIds"):
                    return fam
            except Exception:
                pass
            # a veces llega el Id de FamilySymbol → .Family
            try:
                f2 = getattr(fam, u"Family", None)
                if f2 is not None:
                    return f2
            except Exception:
                pass
            fam = None
    if family_name:
        try:
            from rebar_tag_shape_sync_core import find_families_by_name

            matches = find_families_by_name(doc, family_name) or []
            if matches:
                return matches[0]
        except Exception:
            pass
    return _family_of_tag_type(doc, seed_type_id)


def _tag_type_id_for_rebar_shape(
    doc,
    rebar,
    seed_type_id=None,
    diag=None,
    family_id=None,
    family_name=None,
    allow_seed_fallback=True,
):
    """
    Tipo de etiqueta en la **familia capturada del origen** cuyo nombre
    coincide con el RebarShape **actual** del rebar (p. ej. tras pata L).

    No asume que el type_id origen sea válido: solo define la familia.
    Si no hay match por shape y ``allow_seed_fallback``, devuelve seed.
    """
    if doc is None or rebar is None:
        return seed_type_id if allow_seed_fallback else None
    try:
        from rebar_tag_shape_sync_core import (
            lookup_tag_type_id,
            rebar_shape_name_candidates,
            symbol_map_from_family,
        )
    except Exception as ex:
        if diag is not None:
            diag.ex(u"import rebar_tag_shape_sync_core", ex)
        return seed_type_id if allow_seed_fallback else None

    shapes = []
    try:
        shapes = list(rebar_shape_name_candidates(doc, rebar) or [])
    except Exception as ex:
        if diag is not None:
            diag.ex(u"rebar_shape_name_candidates", ex)
    if diag is not None:
        diag.log(u"rebar shape candidates (nuevo)={0}".format(shapes))

    fam = _resolve_tag_family(
        doc,
        family_id=family_id,
        family_name=family_name,
        seed_type_id=seed_type_id,
    )
    if fam is None:
        if diag is not None:
            diag.log(
                u"tag family no resuelta (family_id={0} name=«{1}» seed={2})".format(
                    _element_id_int(family_id),
                    family_name or u"",
                    _element_id_int(seed_type_id),
                )
            )
        return seed_type_id if allow_seed_fallback else None

    fam_name = _as_unicode(getattr(fam, u"Name", u"?") or u"?")
    try:
        sm = symbol_map_from_family(doc, fam)
    except Exception as ex:
        if diag is not None:
            diag.ex(u"symbol_map_from_family", ex)
        return seed_type_id if allow_seed_fallback else None
    if not sm:
        if diag is not None:
            diag.log(u"familia «{0}»: symbol_map vacío".format(fam_name))
        return seed_type_id if allow_seed_fallback else None

    if diag is not None:
        try:
            # muestra nombres de tipo en familia (diagnóstico shape vs type)
            sample = []
            for k in list(sm.keys())[:12]:
                sample.append(_as_unicode(k))
            diag.log(
                u"familia etiqueta «{0}» tipos-keys muestra={1} n_keys={2}".format(
                    fam_name, sample, len(sm)
                )
            )
        except Exception:
            pass

    for label in shapes:
        try:
            tid = lookup_tag_type_id(sm, label)
        except Exception:
            tid = None
        if tid is not None:
            if diag is not None:
                tname = u""
                try:
                    tname = _as_unicode(getattr(doc.GetElement(tid), u"Name", u""))
                except Exception:
                    pass
                diag.log(
                    u"tag type RESUELTO por shape «{0}» → Id={1} Name=«{2}» "
                    u"familia=«{3}» (seed origen={4})".format(
                        label,
                        _element_id_int(tid),
                        tname,
                        fam_name,
                        _element_id_int(seed_type_id),
                    )
                )
            return tid

    if diag is not None:
        diag.log(
            u"sin tipo para shapes {0} en familia «{1}»; "
            u"fallback seed={2} allow={3}".format(
                shapes,
                fam_name,
                _element_id_int(seed_type_id),
                bool(allow_seed_fallback),
            )
        )
    return seed_type_id if allow_seed_fallback else None


def _sync_tag_type_to_rebar_shape(
    doc, tag, rebar, diag=None, family_id=None, family_name=None, seed_type_id=None
):
    """Alinea el tipo de la IndependentTag al RebarShape (misma familia origen)."""
    if doc is None or tag is None or rebar is None:
        return False
    seed = seed_type_id
    if seed is None:
        try:
            seed = tag.GetTypeId()
        except Exception:
            seed = None
    if family_id is None or not family_name:
        # Si no viene del snapshot, inferir familia del type actual
        try:
            from rebar_tag_shape_sync_core import family_of_independent_tag

            fam = family_of_independent_tag(doc, tag)
            if fam is not None:
                if family_id is None:
                    family_id = fam.Id
                if not family_name:
                    family_name = _as_unicode(getattr(fam, u"Name", u"") or u"")
        except Exception:
            pass
    tid = _tag_type_id_for_rebar_shape(
        doc,
        rebar,
        seed_type_id=seed,
        diag=diag,
        family_id=family_id,
        family_name=family_name,
        allow_seed_fallback=True,
    )
    if tid is None:
        return False
    try:
        cur = tag.GetTypeId()
        if _element_id_int(cur) == _element_id_int(tid):
            if diag is not None:
                diag.log(u"tag type ya coincide con shape nuevo")
            return True
    except Exception:
        pass
    try:
        from rebar_tag_shape_sync_core import try_set_tag_type

        ok, err = try_set_tag_type(doc, tag, rebar, tid, no_activate=True)
        if diag is not None:
            if ok:
                diag.log(
                    u"tag type sync OK → {0}".format(_element_id_int(tid))
                )
            else:
                diag.log(u"tag type sync FAIL: {0}".format(err or u"?"))
        return bool(ok)
    except Exception as ex:
        if diag is not None:
            diag.ex(u"try_set_tag_type", ex)
        try:
            tag.ChangeTypeId(tid)
            if diag is not None:
                diag.log(u"tag ChangeTypeId fallback OK")
            return True
        except Exception as ex2:
            if diag is not None:
                diag.ex(u"ChangeTypeId fallback", ex2)
            return False


def _verify_tag_alive(doc, tag, rebar, diag=None):
    """Comprueba que la etiqueta sigue viva y enlazada tras recreate/sync."""
    if tag is None or doc is None:
        return False
    try:
        tid = tag.Id
        el = doc.GetElement(tid)
    except Exception as ex:
        if diag is not None:
            diag.ex(u"verify GetElement tag", ex)
        return False
    if el is None:
        if diag is not None:
            diag.log(u"verify: tag eliminada del documento")
        return False
    try:
        if getattr(el, u"IsOrphaned", False):
            if diag is not None:
                diag.log(u"verify: tag IsOrphaned=True")
            return False
    except Exception:
        pass
    if rebar is not None:
        rid = _element_id_int(rebar.Id)
        if not _tag_links_rebar(el, rid, diag=None, log_match_detail=False):
            if diag is not None:
                diag.log(u"verify: tag ya no enlaza rebar {0}".format(rid))
            return False
    if diag is not None:
        try:
            v = doc.GetElement(el.OwnerViewId)
            vn = _as_unicode(getattr(v, u"Name", u"?")) if v else u"?"
            tn = u""
            try:
                tn = _as_unicode(getattr(doc.GetElement(el.GetTypeId()), u"Name", u""))
            except Exception:
                pass
            diag.log(
                u"verify OK tag={0} view=«{1}» type=«{2}» orphaned=False".format(
                    _element_id_int(el.Id), vn, tn
                )
            )
        except Exception:
            diag.log(u"verify OK tag={0}".format(_element_id_int(el.Id)))
    return True


def _vista_permite_tag_mra(view):
    """Planta/alzado/sección de modelo (no plantilla, no 3D)."""
    if view is None:
        return False
    try:
        if not isinstance(view, View):
            return False
    except Exception:
        return False
    try:
        if bool(view.IsTemplate):
            return False
    except Exception:
        pass
    try:
        if isinstance(view, View3D):
            return False
    except Exception:
        pass
    return True


def _resolve_target_view(doc, preferred_view, fallback_view_id, diag=None, label=u"view"):
    """
    Vista de colocación: preferir ``preferred_view`` (vista activa);
    si no es válida, usar la vista origen del snapshot.
    """
    if preferred_view is not None and _vista_permite_tag_mra(preferred_view):
        if diag is not None:
            diag.log(
                u"{0} target=ACTIVE «{1}» id={2}".format(
                    label,
                    _as_unicode(getattr(preferred_view, u"Name", u"?")),
                    _element_id_int(getattr(preferred_view, u"Id", None)),
                )
            )
        return preferred_view
    if doc is not None and fallback_view_id is not None:
        try:
            v = doc.GetElement(fallback_view_id)
        except Exception:
            v = None
        if v is not None and _vista_permite_tag_mra(v):
            if diag is not None:
                diag.log(
                    u"{0} target=ORIGEN «{1}» id={2} "
                    u"(activa no válida para tag/MRA)".format(
                        label,
                        _as_unicode(getattr(v, u"Name", u"?")),
                        _element_id_int(fallback_view_id),
                    )
                )
            return v
    if diag is not None:
        diag.log(u"{0} target=None (sin vista válida)".format(label))
    return None


def _midpoint_segmento_principal_rebar(
    rebar, view=None, diag=None, bar_index=None
):
    """
    Mitad del **segmento principal** del centerline (curva más larga).

    Mismo criterio que 56_DividirRebarPuntoTraslape
    (``_rebar_tag_anchor_xyz``): punto sobre el tramo mayor, sin offset
    artificial. Tras pata L, el tramo largo es el cuerpo; la pata corta no.
    """
    if rebar is None:
        return None

    if bar_index is None:
        bar_index = _bar_index_for_tag(rebar, view)
    try:
        pos_idx = int(bar_index)
    except Exception:
        pos_idx = 0

    curves = _centerline_curves_for_tag(rebar, pos_idx)
    if not curves and pos_idx != 0:
        curves = _centerline_curves_for_tag(rebar, 0)
    if not curves:
        if diag is not None:
            diag.log(u"mid principal: sin centerline curves")
        return None

    best = None
    best_len = -1.0
    best_i = 0
    for i, c in enumerate(curves):
        if c is None:
            continue
        ln = 0.0
        try:
            ln = float(c.Length)
        except Exception:
            try:
                ln = float(c.GetEndPoint(0).DistanceTo(c.GetEndPoint(1)))
            except Exception:
                ln = 0.0
        if ln > best_len:
            best_len = ln
            best = c
            best_i = i

    if best is None:
        return None

    mid = None
    try:
        mid = best.Evaluate(0.5, True)
    except Exception:
        try:
            mid = XYZ(
                (best.GetEndPoint(0).X + best.GetEndPoint(1).X) * 0.5,
                (best.GetEndPoint(0).Y + best.GetEndPoint(1).Y) * 0.5,
                (best.GetEndPoint(0).Z + best.GetEndPoint(1).Z) * 0.5,
            )
        except Exception:
            mid = None
    if mid is None:
        return None

    # No proyectar al plano de vista (56 no lo hace: el Z de la barra es correcto).
    p = mid

    if diag is not None:
        try:
            diag.log(
                u"mid principal: curve#{0}/{1} L={2:.1f} mm "
                u"pos={3} pt=({4:.3f},{5:.3f},{6:.3f})".format(
                    best_i,
                    len(curves),
                    float(_mm_from_internal(best_len)),
                    pos_idx,
                    float(p.X),
                    float(p.Y),
                    float(p.Z),
                )
            )
        except Exception:
            diag.log(u"mid principal: OK (sin detalle coords)")
    return p


def _tag_head_for_new_rebar(rebar, tag_info, view=None, diag=None):
    """
    Cabeza de etiqueta sobre el rebar nuevo.

    Por diseño (Pata L): **mitad del tramo principal**.
    Se ignora el ``head`` absoluto del origen (podría quedar fuera del
    crop). ``head_offset`` del origen solo se aplica si es pequeño y
    mantiene legibilidad; por defecto se usan 0 = sobre el tramo.
    """
    bi = _bar_index_for_tag(rebar, view)
    anchor = _midpoint_segmento_principal_rebar(
        rebar, view=view, diag=diag, bar_index=bi
    )
    if anchor is None:
        return None
    # Siempre ancla del tramo principal (pedido explícito del usuario).
    if diag is not None:
        diag.log(u"tag head = mitad tramo principal (offset origen no aplica)")
    return anchor


def _force_tag_head_position(tag, head, diag=None, label=u"force TagHead"):
    """Fuerza TagHeadPosition (API a veces ignora el punto de Create)."""
    if tag is None or head is None:
        return False
    try:
        tag.TagHeadPosition = head
        if diag is not None:
            try:
                h = tag.TagHeadPosition
                diag.log(
                    u"{0} OK=({1:.3f},{2:.3f},{3:.3f})".format(
                        label, float(h.X), float(h.Y), float(h.Z)
                    )
                )
            except Exception:
                diag.log(u"{0} OK".format(label))
        return True
    except Exception as ex:
        if diag is not None:
            diag.ex(label, ex)
        return False


def _rebar_point_en_vista(rebar, view):
    """Compat: ancla de etiqueta (mitad tramo principal)."""
    return _midpoint_segmento_principal_rebar(rebar, view, diag=None)


def _mra_principal_bar_dir(rebar, bar_index=0):
    """Dirección 3D del tramo principal (curva más larga) de una posición."""
    curves = _centerline_curves_for_tag(rebar, bar_index)
    if not curves:
        curves = _centerline_curves(rebar, bar_index)
    if not curves:
        return None
    best = None
    best_len = -1.0
    for c in curves:
        if c is None:
            continue
        try:
            ln = float(c.Length)
        except Exception:
            ln = 0.0
        if ln > best_len:
            best_len = ln
            best = c
    if best is None:
        return None
    try:
        d = best.GetEndPoint(1) - best.GetEndPoint(0)
        if float(d.GetLength()) > 1e-9:
            return d.Normalize()
    except Exception:
        pass
    try:
        p0 = best.Evaluate(0.0, True)
        p1 = best.Evaluate(1.0, True)
        d = p1 - p0
        if float(d.GetLength()) > 1e-9:
            return d.Normalize()
    except Exception:
        pass
    return None


def _mra_project_dir_on_view(v, vd, fallback=None):
    """Proyecta dirección al plano de vista (⊥ vd)."""
    if v is None:
        return fallback
    try:
        proj = v - vd.Multiply(float(v.DotProduct(vd)))
        if float(proj.GetLength()) > 1e-9:
            return proj.Normalize()
    except Exception:
        pass
    return fallback


def _mra_curve_mid(rebar, pos_idx):
    """Punto medio del tramo principal en la posición ``pos_idx``."""
    p = _midpoint_segmento_principal_rebar(
        rebar, view=None, diag=None, bar_index=pos_idx
    )
    if p is not None:
        return p
    curves = _centerline_curves(rebar, pos_idx)
    if not curves:
        return None
    try:
        return curves[0].Evaluate(0.5, True)
    except Exception:
        return None


def _mra_spacing_dir_in_view(rebar, view, diag=None):
    """
    Dirección de cota MRA = distribución del set en el plano de vista.

    Criterio (``multi_rebar_annotation_seleccion_rps`` /
    ``area_rein_losa_sketch``):
    1) pos0 → posN (mids tramo principal)
    2) bar_dir × ViewDirection  (⊥ a la barra → refs lineal válidas)
    3) RightDirection / UpDirection

    Importante: NUNCA usar la dirección del tramo principal de la barra
    como DimensionLineDirection (Rev it: «references must be
    perpendicular to the dimension line»).
    """
    try:
        vd = view.ViewDirection.Normalize()
        rd = view.RightDirection.Normalize()
        vup = view.UpDirection.Normalize()
    except Exception as ex:
        if diag is not None:
            diag.ex(u"mra view vectors", ex)
        return None, None

    spacing = None
    method = u"?"

    try:
        n = int(rebar.NumberOfBarPositions)
    except Exception:
        n = 1

    if n > 1:
        try:
            p0 = _mra_curve_mid(rebar, 0)
            pn = _mra_curve_mid(rebar, n - 1)
            if p0 is not None and pn is not None:
                v = pn - p0
                if float(v.GetLength()) > 1e-6:
                    spacing = _mra_project_dir_on_view(v.Normalize(), vd, None)
                    if spacing is not None:
                        method = u"array pos0→posN n={0}".format(n)
        except Exception as ex:
            if diag is not None:
                diag.ex(u"mra spacing array", ex)

    if spacing is None:
        bar_dir = _mra_principal_bar_dir(rebar, _bar_index_for_tag(rebar, view))
        if bar_dir is not None:
            try:
                # Barra ≈ vertical a la vista (entra en pantalla) → Up
                if abs(float(bar_dir.DotProduct(vd))) > 0.8:
                    spacing = vup
                    method = u"bar∥vd → UpDirection"
                else:
                    cross = bar_dir.CrossProduct(vd)
                    if float(cross.GetLength()) > 1e-9:
                        spacing = _mra_project_dir_on_view(
                            cross.Normalize(), vd, None
                        )
                        if spacing is not None:
                            method = u"barDir×vd"
            except Exception as ex:
                if diag is not None:
                    diag.ex(u"mra barDir×vd", ex)

    if spacing is None:
        # Preferir right si no es // a la barra
        bar_dir = _mra_principal_bar_dir(rebar, 0)
        spacing = rd
        method = u"fallback RightDirection"
        if bar_dir is not None:
            try:
                # Si Right es casi // al tramo, usar Up
                bar_pl = _mra_project_dir_on_view(bar_dir, vd, None)
                if bar_pl is not None:
                    if abs(float(bar_pl.DotProduct(rd))) > 0.9:
                        spacing = vup
                        method = u"fallback UpDirection (Right∥bar)"
            except Exception:
                pass

    # Offset lateral = perpendicular a la distribución en el plano
    try:
        perp = spacing.CrossProduct(vd)
        if float(perp.GetLength()) < 1e-9:
            perp = vup
        else:
            perp = perp.Normalize()
    except Exception:
        perp = vup

    # Si spacing residual // bar proyectada, rotar 90° en plano
    try:
        bar_dir = _mra_principal_bar_dir(rebar, 0)
        bar_pl = _mra_project_dir_on_view(bar_dir, vd, None) if bar_dir else None
        if bar_pl is not None and abs(float(spacing.DotProduct(bar_pl))) > 0.85:
            spacing = perp
            perp = spacing.CrossProduct(vd)
            if float(perp.GetLength()) > 1e-9:
                perp = perp.Normalize()
            method = method + u" + rot90 (era∥bar)"
            if diag is not None:
                diag.log(u"mra: DimensionLineDirection era // bar → rotada 90°")
    except Exception:
        pass

    if diag is not None:
        try:
            diag.log(
                u"mra spacing method={0} dir=({1:.3f},{2:.3f},{3:.3f})".format(
                    method,
                    float(spacing.X),
                    float(spacing.Y),
                    float(spacing.Z),
                )
            )
        except Exception:
            diag.log(u"mra spacing method={0}".format(method))
    return spacing, perp


def _recrear_etiquetas(doc, tag_infos, new_rebar, diag=None, active_view=None):
    n = 0
    errs = []
    if diag is not None:
        diag.step(
            u"RECREATE tags n_src={0} new_rebar={1} active_view={2}".format(
                len(tag_infos or []),
                _element_id_int(getattr(new_rebar, u"Id", None)),
                _element_id_int(getattr(active_view, u"Id", None))
                if active_view is not None
                else None,
            )
        )
    if doc is None or new_rebar is None or not tag_infos:
        if diag is not None:
            diag.log(
                u"skip recreate: doc={0} rebar={1} infos={2}".format(
                    doc is not None,
                    new_rebar is not None,
                    len(tag_infos or []),
                )
            )
        return n, errs
    try:
        doc.Regenerate()
        if diag is not None:
            diag.log(u"Regenerate pre-tag OK")
    except Exception as ex:
        if diag is not None:
            diag.ex(u"Regenerate pre-tag", ex)

    # Una sola etiqueta por vista de destino (evitar duplicar si varios orígenes)
    created_views = set()

    for i, info in enumerate(tag_infos):
        if diag is not None:
            diag.log(
                u"recreate[{0}] family=«{1}» seed_type={2} «{3}» "
                u"view_origen={4} src_tag={5}".format(
                    i,
                    info.get(u"family_name") or u"",
                    _element_id_int(info.get(u"type_id")),
                    info.get(u"type_name") or u"",
                    _element_id_int(info.get(u"view_id")),
                    _element_id_int(info.get(u"tag_id")),
                )
            )
        view = _resolve_target_view(
            doc,
            active_view,
            info.get(u"view_id"),
            diag=diag,
            label=u"tag[{0}]".format(i),
        )
        if view is None:
            errs.append(u"Etiqueta: sin vista activa/origen válida.")
            continue
        vkey = _element_id_int(view.Id)
        if vkey in created_views:
            if diag is not None:
                diag.log(u"tag[{0}] skip: ya hay etiqueta en vista {1}".format(i, vkey))
            continue

        # Ancla = mitad tramo principal (mismo patrón que 56:_rebar_tag_anchor_xyz)
        head = _tag_head_for_new_rebar(new_rebar, info, view, diag=diag)
        if head is None:
            head = _midpoint_segmento_principal_rebar(new_rebar, view, diag=diag)

        seed_type = info.get(u"type_id")
        # Tipo por shape nuevo en familia capturada (como resolve_tag_type_id_for_rebar de 56)
        type_id = _tag_type_id_for_rebar_shape(
            doc,
            new_rebar,
            seed_type_id=seed_type,
            diag=diag,
            family_id=info.get(u"family_id"),
            family_name=info.get(u"family_name"),
            allow_seed_fallback=True,
        )
        if diag is not None:
            t_new = u""
            try:
                if type_id is not None:
                    t_new = _as_unicode(
                        getattr(doc.GetElement(type_id), u"Name", u"") or u""
                    )
            except Exception:
                pass
            diag.log(
                u"tag[{0}] create type Id={1} «{2}» "
                u"(shape nuevo; familia «{3}»)".format(
                    i,
                    _element_id_int(type_id),
                    t_new,
                    info.get(u"family_name") or u"",
                )
            )

        bi = _bar_index_for_tag(new_rebar, view)
        leader = info.get(u"leader")
        if leader is None:
            leader = True

        tag, err = _crear_etiqueta_rebar(
            doc,
            view,
            new_rebar,
            type_id,
            head,
            info.get(u"orient", TagOrientation.Horizontal),
            leader,
            diag=diag,
            rotation=info.get(u"rotation"),
            bar_index=bi,
        )
        # Seed origen solo si el tipo de shape falló (último recurso)
        if tag is None and type_id is not None and seed_type is not None:
            if _element_id_int(type_id) != _element_id_int(seed_type):
                if diag is not None:
                    diag.log(
                        u"recreate[{0}]: reintento con type seed origen "
                        u"(último recurso)".format(i)
                    )
                tag, err = _crear_etiqueta_rebar(
                    doc,
                    view,
                    new_rebar,
                    seed_type,
                    head,
                    info.get(u"orient", TagOrientation.Horizontal),
                    leader,
                    diag=diag,
                    rotation=info.get(u"rotation"),
                    bar_index=bi,
                )
        if tag is not None:
            # NO sync_recreate (56 no lo hace post-create: destruye posición)
            # Solo re-forzar cabeza = mitad tramo principal
            head2 = _midpoint_segmento_principal_rebar(
                new_rebar, view, diag=None, bar_index=bi
            )
            if head2 is None:
                head2 = head
            if head2 is not None:
                try:
                    tag.TagOrientation = info.get(
                        u"orient", TagOrientation.Horizontal
                    )
                except Exception:
                    pass
                _force_tag_head_position(
                    tag, head2, diag=diag, label=u"TagHead FINAL mid principal"
                )
            try:
                doc.Regenerate()
            except Exception:
                pass

            ok_v = _verify_tag_alive(doc, tag, new_rebar, diag=diag)
            if ok_v:
                n += 1
                if vkey is not None:
                    created_views.add(vkey)
                if diag is not None:
                    diag.log(
                        u"recreate[{0}] SUCCESS id={1} en vista activa/destino".format(
                            i, _element_id_int(getattr(tag, u"Id", None))
                        )
                    )
            else:
                errs.append(u"Etiqueta creada pero no verificable / huérfana.")
                if diag is not None:
                    diag.log(u"recreate[{0}] FAIL verify".format(i))
        elif err:
            errs.append(err)
            if diag is not None:
                diag.log(u"recreate[{0}] FAIL: {1}".format(i, err))
    if diag is not None:
        diag.log(u"RECREATE done ok={0} err={1}".format(n, len(errs)))
    return n, errs


def _mra_references_rebar(doc, mra, rebar_id_int):
    if doc is None or mra is None or rebar_id_int is None:
        return False
    dim = None
    try:
        dim = doc.GetElement(mra.DimensionId)
    except Exception:
        dim = None
    if dim is None:
        return False
    try:
        refs = dim.References
    except Exception:
        refs = None
    if refs is None:
        return False
    for r in _iter_id_collection(refs):
        if r is None:
            continue
        try:
            if _element_id_int(r.ElementId) == int(rebar_id_int):
                return True
        except Exception:
            continue
    return False


def _dimension_line_data(doc, dim):
    """(origin XYZ, direction XYZ) aproximados desde la cota MRA."""
    if dim is None:
        return None, None
    origin = None
    direction = None
    try:
        crv = dim.Curve
        if crv is not None:
            origin = crv.Evaluate(0.5, True)
            d = crv.GetEndPoint(1) - crv.GetEndPoint(0)
            if float(d.GetLength()) > 1e-9:
                direction = d.Normalize()
    except Exception:
        pass
    if origin is None:
        try:
            origin = dim.Origin
        except Exception:
            origin = None
    return origin, direction


def _snapshot_mra(doc, rebar, diag=None):
    """MultiReferenceAnnotation (MRA) ligadas al rebar."""
    out = []
    rid = _element_id_int(rebar.Id) if rebar is not None else None
    if diag is not None:
        diag.step(u"SNAPSHOT MRA rebar={0}".format(rid))
    if doc is None or rid is None:
        return out
    try:
        coll = _collector_mra(doc)
    except Exception as ex:
        if diag is not None:
            diag.ex(u"MRA collector", ex)
        return out
    if coll is None:
        return out
    n_scan = 0
    for mra in coll:
        n_scan += 1
        if mra is None:
            continue
        if not _mra_references_rebar(doc, mra, rid):
            continue
        info = {
            u"type_id": None,
            u"view_id": None,
            u"tag_head": None,
            u"dim_origin": None,
            u"dim_dir": None,
            u"tag_leader": False,
            u"mra_id": None,
        }
        try:
            info[u"mra_id"] = mra.Id
        except Exception:
            pass
        try:
            info[u"type_id"] = mra.GetTypeId()
        except Exception:
            pass
        try:
            info[u"view_id"] = mra.OwnerViewId
        except Exception:
            pass
        try:
            tag = doc.GetElement(mra.TagId)
        except Exception:
            tag = None
        if tag is not None:
            try:
                info[u"tag_head"] = tag.TagHeadPosition
            except Exception:
                pass
            try:
                info[u"tag_leader"] = bool(tag.HasLeader)
            except Exception:
                pass
        try:
            dim = doc.GetElement(mra.DimensionId)
        except Exception:
            dim = None
        o, d = _dimension_line_data(doc, dim)
        info[u"dim_origin"] = o
        info[u"dim_dir"] = d
        if info.get(u"type_id") is None:
            if diag is not None:
                diag.log(
                    u"MRA skip id={0}: type_id nulo".format(
                        _element_id_int(info.get(u"mra_id"))
                    )
                )
            continue
        out.append(info)
        if diag is not None:
            tname = u""
            try:
                te = doc.GetElement(info[u"type_id"])
                tname = _as_unicode(getattr(te, u"Name", u"")) if te else u""
            except Exception:
                pass
            diag.log(
                u"MRA capture id={0} type={1} «{2}» view={3}".format(
                    _element_id_int(info.get(u"mra_id")),
                    _element_id_int(info.get(u"type_id")),
                    tname,
                    _element_id_int(info.get(u"view_id")),
                )
            )
    if diag is not None:
        diag.log(u"snapshot MRA RESULT n={0} scanned={1}".format(len(out), n_scan))
    return out


def _recrear_mra(doc, mra_infos, new_rebar, diag=None, active_view=None):
    """
    Recrea MRA en la **vista activa** (si es válida); usa tipo MRA y
    leader del snapshot; recalcula origen/dirección sobre el rebar nuevo.
    """
    n = 0
    errs = []
    if diag is not None:
        diag.step(
            u"RECREATE MRA n_src={0} active_view={1}".format(
                len(mra_infos or []),
                _element_id_int(getattr(active_view, u"Id", None))
                if active_view is not None
                else None,
            )
        )
    if doc is None or new_rebar is None or not mra_infos:
        return n, errs

    done_types_views = set()  # (type_id_int, view_id_int)

    for i, info in enumerate(mra_infos or []):
        view = _resolve_target_view(
            doc,
            active_view,
            info.get(u"view_id"),
            diag=diag,
            label=u"mra[{0}]".format(i),
        )
        if view is None:
            errs.append(u"MRA: sin vista activa/origen válida.")
            continue

        type_id = info.get(u"type_id")
        try:
            mrat = doc.GetElement(type_id) if type_id is not None else None
        except Exception as ex:
            if diag is not None:
                diag.ex(u"mra GetElement type", ex)
            mrat = None

        # Validar tipo (isinstance estricto falla a veces en IronPython)
        is_mrat = False
        if mrat is not None:
            try:
                is_mrat = isinstance(mrat, MultiReferenceAnnotationType)
            except Exception:
                is_mrat = False
            if not is_mrat:
                try:
                    is_mrat = u"MultiReferenceAnnotationType" in type(mrat).__name__
                except Exception:
                    is_mrat = False
        if not is_mrat:
            msg = u"MRA: tipo inválido id={0}".format(_element_id_int(type_id))
            errs.append(msg)
            if diag is not None:
                diag.log(msg + u" class={0}".format(type(mrat).__name__ if mrat else None))
            continue

        key = (_element_id_int(type_id), _element_id_int(view.Id))
        if key in done_types_views:
            if diag is not None:
                diag.log(u"mra[{0}] skip duplicate type+view".format(i))
            continue

        try:
            opts = MultiReferenceAnnotationOptions(mrat)
        except Exception:
            try:
                opts = MultiReferenceAnnotationOptions()
                opts.MultiReferenceAnnotationType = mrat.Id
            except Exception as ex:
                if diag is not None:
                    diag.ex(u"mra options", ex)
                errs.append(u"MRA: no se pudo crear Options.")
                continue
        try:
            opts.DimensionStyleType = DimensionStyleType.Linear
        except Exception:
            pass

        try:
            vd = view.ViewDirection.Normalize()
        except Exception as ex:
            if diag is not None:
                diag.ex(u"mra ViewDirection", ex)
            errs.append(u"MRA: ViewDirection inválida.")
            continue

        # Recalcular geometría en la vista de destino
        spacing_dir, perp = _mra_spacing_dir_in_view(new_rebar, view, diag=diag)
        if spacing_dir is None:
            try:
                spacing_dir = view.RightDirection.Normalize()
            except Exception:
                spacing_dir = XYZ.BasisX
        if perp is None:
            try:
                perp = view.UpDirection.Normalize()
            except Exception:
                perp = XYZ.BasisY

        # Ancla MRA = mitad del tramo principal (como tag / area_rein)
        p_mid = _midpoint_segmento_principal_rebar(
            new_rebar, view=None, diag=diag, bar_index=_bar_index_for_tag(new_rebar, view)
        )
        if p_mid is None:
            p_mid = info.get(u"dim_origin") or info.get(u"tag_head")
        if p_mid is None:
            errs.append(u"MRA: sin punto de ancla en el rebar.")
            if diag is not None:
                diag.log(u"mra[{0}] FAIL sin p_mid".format(i))
            continue

        # Offset lateral a lo largo de la distribución (area_rein_losa_sketch)
        try:
            off = UnitUtils.ConvertToInternalUnits(150.0, UnitTypeId.Millimeters)
        except Exception:
            off = 150.0 / 304.8

        # Candidatos DimensionLineDirection (reintentar si Create rechaza dirección)
        dir_candidates = []
        try:
            dir_candidates.append((spacing_dir, perp, u"spacing"))
            # Invertidos
            dir_candidates.append(
                (spacing_dir.Negate(), perp.Negate(), u"spacing_neg")
            )
            dir_candidates.append((perp, spacing_dir, u"perp_as_dim"))
            dir_candidates.append(
                (perp.Negate(), spacing_dir.Negate(), u"perp_neg")
            )
            try:
                rd = view.RightDirection.Normalize()
                vup = view.UpDirection.Normalize()
                dir_candidates.append((rd, vup, u"view_Right"))
                dir_candidates.append((vup, rd, u"view_Up"))
            except Exception:
                pass
        except Exception:
            dir_candidates = [(spacing_dir, perp, u"spacing")]

        # Dedup por dirección aproximada
        seen_dirs = set()
        unique_dirs = []
        for d, p, lab in dir_candidates:
            if d is None:
                continue
            try:
                key = (
                    round(float(d.X), 3),
                    round(float(d.Y), 3),
                    round(float(d.Z), 3),
                )
            except Exception:
                key = lab
            if key in seen_dirs:
                continue
            seen_dirs.add(key)
            unique_dirs.append((d, p if p is not None else spacing_dir, lab))

        ids = ClrList[ElementId]()
        ids.Add(new_rebar.Id)

        mra = None
        last_ex = None
        used_label = u""
        p_line = p_mid
        for dim_dir, off_dir, lab in unique_dirs:
            try:
                # offset a lo largo de dim (distribución) o de off_dir
                try:
                    p_line = p_mid + dim_dir.Multiply(float(off))
                except Exception:
                    p_line = p_mid
                try:
                    opts.DimensionPlaneNormal = vd
                    opts.DimensionLineDirection = dim_dir
                    opts.DimensionLineOrigin = p_line
                    opts.TagHeadPosition = p_line
                    opts.TagHasLeader = bool(info.get(u"tag_leader", False))
                except Exception as ex:
                    if diag is not None:
                        diag.ex(u"mra config opts[{0}]".format(lab), ex)
                    continue
                try:
                    opts.SetElementsToDimension(ids)
                except Exception as ex:
                    if diag is not None:
                        diag.ex(u"mra SetElements[{0}]".format(lab), ex)
                    continue
                try:
                    if hasattr(opts, u"ElementsMatchReferenceCategory"):
                        if not opts.ElementsMatchReferenceCategory(ids):
                            if diag is not None:
                                diag.log(
                                    u"mra[{0}] {1}: ElementsMatch fail".format(
                                        i, lab
                                    )
                                )
                            continue
                except Exception:
                    pass
                try:
                    mra = MultiReferenceAnnotation.Create(doc, view.Id, opts)
                except Exception as ex:
                    last_ex = ex
                    if diag is not None:
                        diag.log(
                            u"mra[{0}] Create FAIL [{1}]: {2}".format(
                                i, lab, _exception_text(ex)
                            )
                        )
                    mra = None
                    continue
                if mra is not None:
                    used_label = lab
                    break
            except Exception as ex:
                last_ex = ex
                continue

        if mra is None:
            msg = u"MRA Create: {0}".format(
                _exception_text(last_ex) if last_ex is not None else u"None"
            )
            errs.append(msg)
            if diag is not None:
                diag.log(u"mra[{0}] todos los intentos fallaron".format(i))
            continue

        n += 1
        done_types_views.add(key)
        if diag is not None:
            diag.log(
                u"mra[{0}] SUCCESS id={1} dir={2} view=«{3}»".format(
                    i,
                    _element_id_int(mra.Id),
                    used_label,
                    _as_unicode(getattr(view, u"Name", u"?")),
                )
            )
        try:
            mtag = doc.GetElement(mra.TagId)
            if mtag is not None:
                mtag.TagHeadPosition = p_line
        except Exception:
            pass
    if diag is not None:
        diag.log(u"MRA recreated={0} src={1} errs={2}".format(n, len(mra_infos or []), errs[:4]))
    return n, errs


def _snapshot_presentation(doc, rebar):
    """PresentationMode + Unobscured/Solid en vistas no plantilla."""
    out = []
    if doc is None or rebar is None:
        return out
    try:
        views = _collector_views(doc)
    except Exception:
        return out
    if views is None:
        return out
    for view in views:
        if view is None:
            continue
        try:
            if bool(view.IsTemplate):
                continue
        except Exception:
            pass
        try:
            if not isinstance(view, View):
                continue
        except Exception:
            continue
        item = {
            u"view_id": view.Id,
            u"mode": None,
            u"unobscured": None,
            u"solid": None,
        }
        try:
            item[u"mode"] = rebar.GetPresentationMode(view)
        except Exception:
            item[u"mode"] = None
        try:
            item[u"unobscured"] = bool(rebar.IsUnobscuredInView(view))
        except Exception:
            item[u"unobscured"] = None
        try:
            if isinstance(view, View3D):
                item[u"solid"] = bool(rebar.IsSolidInView(view))
        except Exception:
            item[u"solid"] = None
        if (
            item[u"mode"] is None
            and item[u"unobscured"] is None
            and item[u"solid"] is None
        ):
            continue
        out.append(item)
    return out


def _apply_presentation(doc, dst, pres_rows):
    if doc is None or dst is None or not pres_rows:
        return 0
    n_ok = 0
    for item in pres_rows:
        try:
            view = doc.GetElement(item.get(u"view_id"))
        except Exception:
            view = None
        if view is None or not isinstance(view, View):
            continue
        try:
            if bool(view.IsTemplate):
                continue
        except Exception:
            pass
        mode = item.get(u"mode")
        if mode is not None:
            try:
                dst.SetPresentationMode(view, mode)
                n_ok += 1
            except Exception:
                pass
        unob = item.get(u"unobscured")
        if unob is not None:
            try:
                dst.SetUnobscuredInView(view, bool(unob))
            except Exception:
                pass
        sol = item.get(u"solid")
        if sol is not None:
            try:
                if isinstance(view, View3D):
                    dst.SetSolidInView(view, bool(sol))
            except Exception:
                pass
    return n_ok


def _rebar_get_style(rebar):
    """Lee RebarStyle (algunas builds IronPython no exponen `.Style`)."""
    if rebar is None:
        return RebarStyle.Standard
    for name in (u"Style", u"get_Style"):
        try:
            attr = getattr(rebar, name, None)
        except Exception:
            attr = None
        if attr is None:
            continue
        try:
            return attr() if callable(attr) else attr
        except Exception:
            continue
    return RebarStyle.Standard


def _rebar_set_style(rebar, style, diag=None):
    """Asigna Style si la API lo permite; no falla la herencia."""
    if rebar is None or style is None:
        return False
    try:
        rebar.Style = style
        if diag is not None:
            diag.log(u"Style set via property = {0}".format(style))
        return True
    except Exception:
        pass
    try:
        setter = getattr(rebar, u"set_Style", None)
        if setter is not None and callable(setter):
            setter(style)
            if diag is not None:
                diag.log(u"Style set via set_Style = {0}".format(style))
            return True
    except Exception as ex:
        if diag is not None:
            diag.log(
                u"Style omitido (API no escribible): {0}".format(
                    _exception_text(ex)
                )
            )
        return False
    if diag is not None:
        diag.log(u"Style omitido: Rebar sin propiedad Style en esta build.")
    return False


def _snapshot_rebar_config(doc, rebar, diag=None):
    """Captura configuración a heredar (antes de borrar el original)."""
    cfg = {
        u"style": None,
        u"hooks": [],
        u"params": [],
        u"moved": [],
        u"included": [],
        u"presentation": [],
        u"tags": [],
        u"mra": [],
    }
    if rebar is None:
        if diag is not None:
            diag.log(u"snapshot: rebar is None")
        return cfg
    if diag is not None:
        diag.step(
            u"SNAPSHOT config rebar={0}".format(_element_id_int(rebar.Id))
        )
    cfg[u"style"] = _rebar_get_style(rebar)
    cfg[u"hooks"] = _snapshot_hooks(rebar)
    cfg[u"params"] = _snapshot_instance_params(rebar)
    cfg[u"moved"] = _snapshot_moved_bar_transforms(rebar)
    cfg[u"included"] = _snapshot_bars_included(rebar)
    cfg[u"presentation"] = _snapshot_presentation(doc, rebar)
    cfg[u"tags"] = _snapshot_independent_tags(doc, rebar, diag=diag)
    cfg[u"mra"] = _snapshot_mra(doc, rebar)
    if diag is not None:
        diag.log(
            u"snapshot summary: tags={0} mra={1} pres_views={2} hooks={3} params={4} style={5}".format(
                len(cfg.get(u"tags") or []),
                len(cfg.get(u"mra") or []),
                len(cfg.get(u"presentation") or []),
                len(cfg.get(u"hooks") or []),
                len(cfg.get(u"params") or []),
                cfg.get(u"style"),
            )
        )
    return cfg


def _apply_rebar_config(doc, src_for_layout, dst, cfg, diag=None, active_view=None):
    """
    Aplica layout + hooks + params + representación + etiquetas + MRA.

    ``src_for_layout`` solo se usa si aún existe en el doc; si es None, se
    asume que el layout ya lo copió ``extend_rebar_pata_l``.

    Tags y MRA se recrean en ``active_view`` (vista donde se ejecuta la
    herramienta) cuando es válida. El type de etiqueta se elige en la
    familia capturada del origen según el RebarShape **nuevo**.
    """
    if dst is None or cfg is None:
        if diag is not None:
            diag.log(u"apply_config: dst/cfg nulo")
        return
    if diag is not None:
        diag.step(
            u"APPLY config → new_rebar={0} active_view={1}".format(
                _element_id_int(getattr(dst, u"Id", None)),
                _element_id_int(getattr(active_view, u"Id", None))
                if active_view is not None
                else None,
            )
        )
    # Layout (si el original sigue vivo: reaplicación más robusta)
    if src_for_layout is not None:
        try:
            from rebar_extender_l_ganchos_135_rps import (
                _copy_layout_rebar_shape_driven as _copy_layout_ext,
            )

            _copy_layout_ext(src_for_layout, dst)
            if diag is not None:
                diag.log(u"layout: rebar_extender ok")
        except Exception as ex:
            if diag is not None:
                diag.ex(u"layout rebar_extender", ex)
            try:
                from area_rein_losa_sketch_pata import _copy_layout_shape_driven

                _copy_layout_shape_driven(src_for_layout, dst)
                if diag is not None:
                    diag.log(u"layout: area_rein_losa_sketch_pata ok")
            except Exception as ex2:
                if diag is not None:
                    diag.ex(u"layout sketch_pata", ex2)

    _rebar_set_style(dst, cfg.get(u"style"), diag=diag)

    _apply_hooks(dst, cfg.get(u"hooks") or [])
    try:
        if doc is not None:
            doc.Regenerate()
    except Exception as ex:
        if diag is not None:
            diag.ex(u"Regenerate post-hooks", ex)

    _apply_instance_params(dst, cfg.get(u"params") or [])
    _apply_bars_included(dst, cfg.get(u"included") or [])
    _apply_moved_bar_transforms(dst, cfg.get(u"moved") or [])

    # Segunda pasada de hooks/params por si layout regeneró valores por defecto
    try:
        if doc is not None:
            doc.Regenerate()
    except Exception:
        pass
    _apply_hooks(dst, cfg.get(u"hooks") or [])
    _apply_instance_params(dst, cfg.get(u"params") or [])

    # Representación (antes de etiquetas/MRA: Middle/Select etc. en cada vista)
    n_pres = _apply_presentation(doc, dst, cfg.get(u"presentation") or [])
    if diag is not None:
        diag.log(u"presentation applied views≈{0}".format(n_pres))
    try:
        if doc is not None:
            doc.Regenerate()
    except Exception as ex:
        if diag is not None:
            diag.ex(u"Regenerate post-presentation", ex)

    n_tags, tag_errs = _recrear_etiquetas(
        doc, cfg.get(u"tags") or [], dst, diag=diag, active_view=active_view
    )
    n_mra, mra_errs = _recrear_mra(
        doc, cfg.get(u"mra") or [], dst, diag=diag, active_view=active_view
    )

    # Verificar tags tras recreate (siguen vivos y apuntan al rebar nuevo)
    if diag is not None and n_tags > 0:
        try:
            want = _element_id_int(dst.Id)
            n_linked = 0
            coll = _collector_rebar_tags(doc)
            if coll is None:
                coll = (
                    FilteredElementCollector(doc)
                    .OfClass(IndependentTag)
                    .WhereElementIsNotElementType()
                )
            for tag in coll:
                if not isinstance(tag, IndependentTag):
                    continue
                if not _tag_links_rebar(tag, want, diag=None, log_match_detail=False):
                    continue
                n_linked += 1
            diag.log(
                u"post-check IndependentTag linked to new rebar: {0}".format(
                    n_linked
                )
            )
        except Exception as ex:
            diag.ex(u"post-check tags", ex)

    all_errs = list(tag_errs or [])[:6]
    for me in (mra_errs or [])[:4]:
        if me and me not in all_errs:
            all_errs.append(me)

    # Notas en cfg para resumen al usuario
    cfg[u"_n_pres"] = n_pres
    cfg[u"_n_tags"] = n_tags
    cfg[u"_n_mra"] = n_mra
    cfg[u"_n_tags_src"] = len(cfg.get(u"tags") or [])
    cfg[u"_n_mra_src"] = len(cfg.get(u"mra") or [])
    cfg[u"_tag_errs"] = all_errs
    if diag is not None:
        diag.log(
            u"APPLY DONE tags {0}/{1} mra {2}/{3} errs={4}".format(
                n_tags,
                cfg[u"_n_tags_src"],
                n_mra,
                cfg[u"_n_mra_src"],
                cfg[u"_tag_errs"],
            )
        )


def _aplicar_pata_l(
    doc, rebar, pata_start, pata_end, dir_xyz, largo_mm, diag=None, active_view=None
):
    from area_rein_losa_sketch_pata import extend_rebar_pata_l

    if diag is None:
        diag = _DiagSession()
    avisos = []
    if diag is not None:
        diag.step(u"INICIO pata L")
        diag.log(
            u"pata_start={0} pata_end={1} largo_mm={2:g} dir=({3:.3f},{4:.3f},{5:.3f})".format(
                bool(pata_start),
                bool(pata_end),
                float(largo_mm),
                float(dir_xyz.X) if dir_xyz else 0.0,
                float(dir_xyz.Y) if dir_xyz else 0.0,
                float(dir_xyz.Z) if dir_xyz else 0.0,
            )
        )
        diag.log(u"rebar src id={0}".format(_element_id_int(rebar.Id)))
        if active_view is not None:
            diag.log(
                u"active_view «{0}» id={1}".format(
                    _as_unicode(getattr(active_view, u"Name", u"?")),
                    _element_id_int(getattr(active_view, u"Id", None)),
                )
            )

    cfg = _snapshot_rebar_config(doc, rebar, diag=diag)
    try:
        avisos.append(
            u"Capturado: {0} etiqueta(s), {1} MRA.".format(
                len(cfg.get(u"tags") or []),
                len(cfg.get(u"mra") or []),
            )
        )
    except Exception:
        pass

    new_rb = None
    try:
        # Un solo Undo (Assimilate) + tx nativa pyRevit; excepción → RollBack limpio
        with _transaction_group_scope(doc, _TRANSACTION_GROUP, assimilate=True):
            with _transaction_scope(doc, _TRANSACTION_NAME):
                if diag is not None:
                    diag.step(u"TX: extend_rebar_pata_l")
                # extend_rebar_pata_l crea el nuevo, copia layout básico y borra el original
                new_rb = extend_rebar_pata_l(
                    doc,
                    rebar,
                    bool(pata_start),
                    bool(pata_end),
                    dir_xyz,
                    float(largo_mm),
                    plane=None,
                    avisos=avisos,
                )
                if new_rb is None or new_rb is rebar:
                    raise RuntimeError(
                        u"; ".join(avisos)
                        if avisos
                        else u"No se pudo crear la pata L."
                    )
                try:
                    # Comparación por ints (Value/IntegerValue) — sin lógica geom.
                    if int(_element_id_int(new_rb.Id) or -1) == int(
                        _element_id_int(rebar.Id) or -2
                    ):
                        raise RuntimeError(
                            u"; ".join(avisos)
                            if avisos
                            else u"La pata L no se aplicó."
                        )
                except RuntimeError:
                    raise
                except Exception:
                    pass

                if diag is not None:
                    diag.log(
                        u"nuevo rebar id={0} avisos_extend={1}".format(
                            _element_id_int(new_rb.Id),
                            u" | ".join(avisos[-5:]),
                        )
                    )
                    try:
                        from rebar_tag_shape_sync_core import (
                            rebar_shape_name_candidates,
                        )

                        shapes = list(
                            rebar_shape_name_candidates(doc, new_rb) or []
                        )
                        diag.log(
                            u"nuevo rebar shape candidates={0}".format(shapes)
                        )
                    except Exception as ex:
                        diag.ex(u"post-create shape candidates", ex)
                    diag.step(u"TX: apply_rebar_config (tags/MRA/…)")

                # Original ya eliminado → heredar snapshot (shape nueva se mantiene)
                _apply_rebar_config(
                    doc, None, new_rb, cfg, diag=diag, active_view=active_view
                )

        if diag is not None:
            diag.log(u"TX Group Commit/Assimilate OK")
        parts = list(avisos or [])
        try:
            n_t = int(cfg.get(u"_n_tags") or 0)
            n_m = int(cfg.get(u"_n_mra") or 0)
            n_p = int(cfg.get(u"_n_pres") or 0)
            n_ts = int(cfg.get(u"_n_tags_src") or 0)
            n_ms = int(cfg.get(u"_n_mra_src") or 0)
            parts.append(
                u"Heredado: {0} vista(s) presentación · "
                u"{1}/{2} etiqueta(s) · {3}/{4} MRA.".format(
                    n_p, n_t, n_ts, n_m, n_ms
                )
            )
            for te in cfg.get(u"_tag_errs") or []:
                if te:
                    parts.append(u"Tag/MRA: {0}".format(te))
            if n_ts > 0 and n_t <= 0:
                parts.append(u"Aviso: no se recreó ninguna etiqueta capturada.")
            if n_ms > 0 and n_m <= 0:
                parts.append(u"Aviso: no se recreó ningún MRA capturado.")
            if n_ts <= 0:
                parts.append(
                    u"Aviso: no se capturó ninguna IndependentTag del rebar origen."
                )
        except Exception:
            pass
        if diag is not None:
            try:
                parts.append(diag.summary_for_dialog())
            except Exception:
                if diag.path():
                    parts.append(u"Log: {0}".format(diag.path()))
        return True, u"\n".join([p for p in parts if p]), new_rb
    except Exception as ex:
        if diag is not None:
            diag.ex(u"ROLLBACK / error", ex)
        msg = _as_unicode(ex)
        if diag is not None:
            try:
                msg = diag.summary_for_dialog(msg)
            except Exception:
                pass
        return False, msg, None


def run(uiapp):
    """Entrada principal (uiapp = __revit__)."""
    diag = _DiagSession()
    if uiapp is None:
        return
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        _mostrar_aviso(uiapp, u"No hay documento activo.")
        return
    doc = uidoc.Document
    if doc is None:
        _mostrar_aviso(uiapp, u"No hay documento activo.")
        return

    active_view = None
    try:
        active_view = uidoc.ActiveView
    except Exception:
        active_view = None
    if active_view is None:
        try:
            active_view = doc.ActiveView
        except Exception:
            active_view = None

    rebar = _rebar_from_selection(uidoc)
    if rebar is None:
        rebar = _pick_rebar(uidoc)
    if rebar is None:
        return

    if not isinstance(rebar, Rebar):
        _mostrar_aviso(uiapp, u"La selección no es un Structural Rebar.")
        return

    if _is_free_form(rebar):
        _mostrar_aviso(
            uiapp,
            u"Rebar free-form no soportado.",
            u"Use una barra shape-driven.",
        )
        return

    floor = _get_floor_host(doc, rebar)
    if floor is None:
        _mostrar_aviso(
            uiapp,
            u"El host del Rebar no es una losa (Floor).",
            u"Esta herramienta genera la pata L con largo = espesor de losa − 50 mm.",
        )
        return

    curves = _centerline_curves(rebar, 0)
    if not curves:
        _mostrar_aviso(uiapp, u"No se pudo leer el centerline de la barra.")
        return

    vertices = _ordered_vertices(curves)
    if len(vertices) < 2:
        _mostrar_aviso(uiapp, u"La barra no tiene vértices suficientes.")
        return

    pick_pt, pick_err = _pick_point_on_rebar(uidoc, rebar)
    if pick_err == u"cancel":
        return
    if pick_pt is None:
        _mostrar_aviso(
            uiapp,
            u"No se pudo seleccionar el extremo de la barra.",
            pick_err or u"",
        )
        return

    v_idx, dist = _closest_end_vertex(vertices, pick_pt)
    if v_idx is None:
        d_mm = _mm_from_internal(dist) if dist is not None else None
        extra = (
            u"Distancia al extremo más cercano ≈ {0:g} mm "
            u"(tol. {1:g} mm). Haga clic más cerca del extremo."
        ).format(float(d_mm), float(_VERTEX_PICK_TOL_MM)) if d_mm is not None else (
            u"Haga clic sobre la barra, cerca del extremo con pata L."
        )
        _mostrar_aviso(
            uiapp,
            u"No hay un extremo de la barra cerca del punto seleccionado.",
            extra,
        )
        return

    pata_start = int(v_idx) == 0
    pata_end = (not pata_start) and int(v_idx) == len(vertices) - 1
    if not pata_start and not pata_end:
        pata_start = True
        pata_end = False

    largo_mm = _pata_largo_mm(doc, floor)
    dir_xyz = _pata_dir_xyz(rebar, floor, vertices)

    ok, detail, new_rb = _aplicar_pata_l(
        doc,
        rebar,
        pata_start,
        pata_end,
        dir_xyz,
        largo_mm,
        diag=diag,
        active_view=active_view,
    )
    if not ok:
        _mostrar_aviso(
            uiapp,
            u"No se pudo generar la pata L.",
            detail or u"",
        )
        return

    extremo = u"inicio" if pata_start else u"final"
    try:
        new_id = _element_id_int(new_rb.Id) if new_rb is not None else 0
        if new_id is None:
            new_id = 0
    except Exception:
        new_id = 0
    dist_mm = _mm_from_internal(dist) if dist is not None else 0.0
    content_parts = [
        u"Extremo: {0} del trazado.".format(extremo),
        u"Largo pata: {0:g} mm (espesor losa − {1:g} mm).".format(
            float(largo_mm), _PATA_RESTA_MM
        ),
        u"Vértice a ≈ {0:g} mm del clic.".format(float(dist_mm)),
    ]
    if new_id:
        content_parts.append(u"Nuevo Rebar Id: {0}.".format(new_id))
    if detail:
        content_parts.append(detail)
    _mostrar_aviso(
        uiapp,
        u"Pata L aplicada correctamente.",
        u"\n".join(content_parts),
    )


def run_pyrevit(__revit__):
    run(__revit__)
