# -*- coding: utf-8 -*-
# Ejecutar en RPS: File > Run script (no pegar línea a línea).
"""
RPS (Revit 2024+): Agregar contornos de 2 Shaft Openings al boundary (sketch) de un AreaReinforcement.

Flujo:
1) Seleccionar 1 elemento: AreaReinforcement (sin filtros de clase/categoría en PickObject; se valida después)
2) Seleccionar 2 elementos: Shaft Openings (se valida después; usa BoundaryCurves/BoundaryRect)
3) Extraer y almacenar las curvas de ambos shafts
4) Editar el sketch/boundary del AreaReinforcement y agregar esas curvas
"""

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from System.Collections.Generic import List
from Autodesk.Revit.DB import (
    Arc,
    Curve,
    ElementClassFilter,
    ElementId,
    Line,
    Plane,
    Sketch,
    SketchEditScope,
    SketchPlane,
    Transaction,
    XYZ,
)
from Autodesk.Revit.DB.Structure import AreaReinforcement
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType


# ── Document / UIDocument (pyRevit usa __revit__; RPS puede predefinir doc/uidoc) ─
try:
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
except NameError:
    try:
        doc  # noqa: F401
    except NameError:
        doc = None
    try:
        uidoc  # noqa: F401
    except NameError:
        uidoc = None


def _revit_to_foreground():
    """
    Trae Revit al primer plano antes de iniciar selecciones (RPS).
    En algunos entornos Windows, PickObject puede fallar si Revit no está activo.
    """
    try:
        import ctypes

        uiapp = getattr(uidoc, "Application", None)
        hwnd_raw = getattr(uiapp, "MainWindowHandle", 0) if uiapp is not None else 0
        try:
            hwnd = int(hwnd_raw.ToInt64())  # IntPtr -> int
        except Exception:
            try:
                hwnd = int(hwnd_raw)
            except Exception:
                hwnd = 0

        if not hwnd:
            return

        user32 = ctypes.windll.user32
        try:
            kernel32 = ctypes.windll.kernel32
        except Exception:
            kernel32 = None
        SW_RESTORE = 9

        # 1) Asegurar que la ventana no esté minimizada
        try:
            user32.ShowWindowAsync(hwnd, SW_RESTORE)
        except Exception:
            try:
                user32.ShowWindow(hwnd, SW_RESTORE)
            except Exception:
                pass

        # 1.5) Truco típico para permitir SetForegroundWindow: simular ALT
        # (Windows a veces bloquea el foco si el "input" viene de otra ventana/app)
        try:
            VK_MENU = 0x12  # ALT
            KEYEVENTF_KEYUP = 0x0002
            user32.keybd_event(VK_MENU, 0, 0, 0)
            user32.keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0)
        except Exception:
            pass

        # 2) Workaround de restricciones de foco: conectar threads temporalmente
        try:
            GetForegroundWindow = user32.GetForegroundWindow
            GetWindowThreadProcessId = user32.GetWindowThreadProcessId
            AttachThreadInput = user32.AttachThreadInput
            GetCurrentThreadId = kernel32.GetCurrentThreadId if kernel32 else None

            fg = GetForegroundWindow()
            if fg:
                fg_tid = GetWindowThreadProcessId(fg, 0)
            else:
                fg_tid = 0
            cur_tid = GetCurrentThreadId() if GetCurrentThreadId else 0

            if fg_tid and fg_tid != cur_tid:
                AttachThreadInput(fg_tid, cur_tid, True)
                try:
                    user32.SetForegroundWindow(hwnd)
                    user32.BringWindowToTop(hwnd)
                    user32.SetActiveWindow(hwnd)
                    try:
                        user32.SetFocus(hwnd)
                    except Exception:
                        pass
                finally:
                    AttachThreadInput(fg_tid, cur_tid, False)
            else:
                user32.SetForegroundWindow(hwnd)
                user32.BringWindowToTop(hwnd)
                user32.SetActiveWindow(hwnd)
                try:
                    user32.SetFocus(hwnd)
                except Exception:
                    pass
        except Exception:
            # 3) Fallback simple
            try:
                user32.SetForegroundWindow(hwnd)
            except Exception:
                pass

        # 4) Fallback adicional: SwitchToThisWindow (deprecated pero útil)
        try:
            user32.SwitchToThisWindow(hwnd, True)
        except Exception:
            pass
    except Exception:
        pass


def _minimize_foreground_if_not_revit():
    """
    Si la ventana en primer plano NO es Revit, pero pertenece al MISMO proceso,
    la minimiza. Esto suele ser la consola/ventana de RPS que queda "encima" de Revit.
    """
    try:
        import ctypes

        uiapp = getattr(uidoc, "Application", None)
        hwnd_raw = getattr(uiapp, "MainWindowHandle", 0) if uiapp is not None else 0
        try:
            revit_hwnd = int(hwnd_raw.ToInt64())
        except Exception:
            try:
                revit_hwnd = int(hwnd_raw)
            except Exception:
                revit_hwnd = 0

        if not revit_hwnd:
            return

        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32

        fg = user32.GetForegroundWindow()
        if not fg:
            return

        # Si ya es Revit, no hacemos nada.
        if int(fg) == int(revit_hwnd):
            return

        # Si la ventana foreground pertenece al mismo proceso (Revit), es muy probable que sea RPS.
        pid_buf = ctypes.c_ulong(0)
        user32.GetWindowThreadProcessId(fg, ctypes.byref(pid_buf))
        fg_pid = int(pid_buf.value)
        cur_pid = int(kernel32.GetCurrentProcessId())

        if fg_pid and fg_pid == cur_pid:
            SW_MINIMIZE = 6
            try:
                user32.ShowWindowAsync(fg, SW_MINIMIZE)
            except Exception:
                try:
                    user32.ShowWindow(fg, SW_MINIMIZE)
                except Exception:
                    pass
    except Exception:
        pass


def _prepare_revit_for_pick():
    """
    Orden requerido:
    1) Llevar Revit al frente
    2) Si algo del mismo proceso queda encima (RPS), minimizarlo
    3) Llevar Revit al frente nuevamente
    """
    _revit_to_foreground()
    _minimize_foreground_if_not_revit()
    _revit_to_foreground()


def _hide_rps_window_if_available():
    """
    Patrón RPS correcto: quitar Topmost y ocultar __window__ antes de PickObjects/PickObject.
    Retorna una función restore() segura para llamar en finally.
    """
    w = globals().get("__window__", None)
    if w is None:
        def _noop():
            return
        return _noop

    prev_topmost = None
    try:
        prev_topmost = bool(getattr(w, "Topmost", False))
    except Exception:
        prev_topmost = None

    try:
        w.Topmost = False
    except Exception:
        pass
    try:
        w.Hide()
    except Exception:
        pass

    def _restore():
        try:
            w.Show()
        except Exception:
            pass
        try:
            if prev_topmost is not None:
                w.Topmost = prev_topmost
        except Exception:
            pass
        try:
            w.Activate()
        except Exception:
            pass

    return _restore


def _dot(a, b):
    return a.X * b.X + a.Y * b.Y + a.Z * b.Z


def _project_point_to_plane(pt, plane):
    """Proyecta un XYZ al plano (por proyección ortogonal)."""
    if pt is None or plane is None:
        return None
    n = plane.Normal
    o = plane.Origin
    if n is None or o is None:
        return None
    v = pt - o
    dist = _dot(v, n)
    return pt - (n * dist)


def _project_curve_to_plane(curve, plane):
    """Proyecta Line/Arc al plano. Retorna None si no se puede."""
    if curve is None or not curve.IsBound or plane is None:
        return None
    try:
        if isinstance(curve, Line):
            p0 = _project_point_to_plane(curve.GetEndPoint(0), plane)
            p1 = _project_point_to_plane(curve.GetEndPoint(1), plane)
            if p0 is None or p1 is None:
                return None
            if p0.DistanceTo(p1) < 1e-6:
                return None
            return Line.CreateBound(p0, p1)
        if isinstance(curve, Arc):
            p0 = _project_point_to_plane(curve.GetEndPoint(0), plane)
            p1 = _project_point_to_plane(curve.GetEndPoint(1), plane)
            c = _project_point_to_plane(curve.Center, plane)
            if p0 is None or p1 is None or c is None:
                return None
            # Nota: esto asume que el arco original era coplanar; para shafts típicos (en planta) funciona.
            return Arc.Create(p0, p1, c)
    except Exception:
        return None
    return None


def _curvas_desde_shaft(opening):
    """
    Obtiene las curvas del contorno de un shaft opening (BoundaryCurves o BoundaryRect).
    Retorna lista[Curve] (puede ser vacía).
    """
    curvas = []
    # Rectangular boundary
    try:
        is_rect = getattr(opening, "IsRectBoundary", False)
        if is_rect:
            rect = getattr(opening, "BoundaryRect", None)
            if rect is not None:
                mn = getattr(rect, "Min", None) or getattr(rect, "Minimum", None)
                mx = getattr(rect, "Max", None) or getattr(rect, "Maximum", None)
                if mn is not None and mx is not None:
                    z = mn.Z
                    pts = [
                        XYZ(mn.X, mn.Y, z),
                        XYZ(mx.X, mn.Y, z),
                        XYZ(mx.X, mx.Y, z),
                        XYZ(mn.X, mx.Y, z),
                    ]
                    for i in range(4):
                        c = Line.CreateBound(pts[i], pts[(i + 1) % 4])
                        if c is not None:
                            curvas.append(c)
            return curvas
    except Exception:
        pass

    # Curved / non-rect boundary
    try:
        boundary = getattr(opening, "BoundaryCurves", None)
        if boundary is not None:
            for c in boundary:
                if c is not None and c.IsBound:
                    curvas.append(c)
    except Exception:
        pass
    return curvas


class _SilenceSketchFailures(object):
    """Failure preprocessor simple para evitar prompts al hacer commit del SketchEditScope."""

    def PreprocessFailures(self, failuresAccessor):
        try:
            fails = failuresAccessor.GetFailureMessages()
            if fails:
                for f in fails:
                    try:
                        failuresAccessor.DeleteWarning(f)
                    except Exception:
                        pass
        except Exception:
            pass
        # 0 == FailureProcessingResult.Continue
        return 0


def _get_area_reinforcement_sketch_id(area_rein, document):
    """
    Obtiene el SketchId del AreaReinforcement tomando uno de sus boundary CurveElements.
    Usa AreaReinforcement.GetBoundaryCurveIds() -> CurveElement.SketchId
    """
    if area_rein is None or document is None:
        return None

    # Fallback 0: muchas veces el Sketch es dependiente directo del AreaReinforcement
    try:
        # En IronPython suele requerirse el tipo CLR explícito
        dep = area_rein.GetDependentElements(ElementClassFilter(clr.GetClrType(Sketch)))
        if dep:
            # normalmente hay 1; tomamos el primero válido
            for sid in dep:
                if sid and sid != ElementId.InvalidElementId:
                    sk = document.GetElement(sid)
                    if sk is not None and isinstance(sk, Sketch):
                        return sk.Id
    except Exception:
        pass

    # Fallback 0.5: sin filtro (más lento), pero útil si el filtro no captura el Sketch
    try:
        dep_all = area_rein.GetDependentElements(None)
        if dep_all:
            for sid in dep_all:
                if sid and sid != ElementId.InvalidElementId:
                    sk = document.GetElement(sid)
                    if sk is not None and isinstance(sk, Sketch):
                        return sk.Id
    except Exception:
        pass

    try:
        ids = area_rein.GetBoundaryCurveIds()
    except Exception:
        ids = None
    if not ids:
        return None

    for cid in ids:
        try:
            ce = document.GetElement(cid)
            sid = getattr(ce, "SketchId", None)
            if sid and sid != ElementId.InvalidElementId:
                return sid
        except Exception:
            continue
    return None


def _add_curves_to_area_reinforcement_boundary(area_rein, shaft_curves, document):
    """
    IMPORTANTE (Revit 2024): AreaReinforcement NO soporta edición de sketch/boundary via SketchEditScope.
    Además, la sobrecarga Create(document, host, IList<Curve>, ...) exige UN SOLO bucle cerrado (sin huecos),
    por lo que no es posible “agregar shafts” como bucles internos por API.

    Workaround automatizable:
    - Crear ModelCurves con el contorno de los shafts en el mismo plano del boundary del AreaReinforcement,
      dejándolas seleccionadas para que el usuario use "Editar contorno" manualmente.
    """
    if area_rein is None or document is None:
        raise Exception("AreaReinforcement/document inválidos.")
    if not shaft_curves:
        return 0

    # Obtener plano del boundary existente (desde sus CurveElements)
    plane = None
    try:
        b_ids = list(area_rein.GetBoundaryCurveIds() or [])
    except Exception:
        b_ids = []
    if b_ids:
        try:
            ce0 = document.GetElement(b_ids[0])
            gc0 = getattr(ce0, "GeometryCurve", None)
            if gc0 is not None:
                p = gc0.GetEndPoint(0)
                z = getattr(p, "Z", 0.0)
                plane = Plane.CreateByNormalAndOrigin(XYZ(0, 0, 1), XYZ(0, 0, z))
        except Exception:
            plane = None
    if plane is None:
        # fallback: plano horizontal en Z=0
        plane = Plane.CreateByNormalAndOrigin(XYZ(0, 0, 1), XYZ(0, 0, 0))

    sketch_plane = SketchPlane.Create(document, plane)
    created_ids = []
    t = Transaction(document, "RPS: Model lines shafts (para editar boundary)")
    t.Start()
    try:
        for c in shaft_curves:
            pc = _project_curve_to_plane(c, plane) if c is not None else None
            if pc is None:
                continue
            try:
                mc = document.Create.NewModelCurve(pc, sketch_plane)
                if mc is not None:
                    created_ids.append(mc.Id)
            except Exception:
                continue
        t.Commit()
    except Exception:
        if t.HasStarted():
            t.RollBack()
        raise

    # Dejar selección lista (AreaReinforcement + curvas de shaft)
    try:
        to_select = [area_rein.Id] + created_ids
        uidoc.Selection.SetElementIds(List[ElementId](to_select))
    except Exception:
        pass

    return len(created_ids)


def _pick_element(prompt):
    """PickObject sin filtros de clase/categoría; retorna Element o None."""
    try:
        ref = uidoc.Selection.PickObject(ObjectType.Element, prompt)
        if ref is None:
            return None
        return doc.GetElement(ref.ElementId)
    except Exception:
        return None


# ── Main ─────────────────────────────────────────────────────────────────────
if doc is None or uidoc is None:
    print("Ejecuta este script dentro de Revit (RPS).")
else:
    # Selección interactiva (requerida): 3 selecciones por separado (PickObject),
    # manteniendo el patrón RPS correcto (ocultar __window__ para que Revit capture input).
    restore_window = _hide_rps_window_if_available()
    try:
        _prepare_revit_for_pick()
        ar_elem = _pick_element("1/3 Selecciona el STRUCTURAL AreaReinforcement")

        _prepare_revit_for_pick()
        s1 = _pick_element("2/3 Selecciona el primer Shaft Opening")

        _prepare_revit_for_pick()
        s2 = _pick_element("3/3 Selecciona el segundo Shaft Opening")
    except Exception as ex:
        TaskDialog.Show("BIMTools", "Operación cancelada o error en selección:\n{}".format(ex))
        raise SystemExit
    finally:
        restore_window()

    if ar_elem is None or not isinstance(ar_elem, AreaReinforcement):
        TaskDialog.Show("BIMTools", "El primer elemento debe ser un AreaReinforcement (Structural Area Reinforcement).")
    elif s1 is None or s2 is None:
        TaskDialog.Show("BIMTools", "Selección cancelada o inválida.")
    else:
        curvas = []
        curvas.extend(_curvas_desde_shaft(s1) or [])
        curvas.extend(_curvas_desde_shaft(s2) or [])
        if not curvas:
            TaskDialog.Show("BIMTools", "No se pudieron extraer curvas de los Shaft Openings (BoundaryCurves/BoundaryRect).")
        else:
            try:
                n = _add_curves_to_area_reinforcement_boundary(ar_elem, curvas, doc)
                TaskDialog.Show(
                    "BIMTools",
                    "Limitación API Revit 2024:\n"
                    "- No se puede editar el boundary del AreaReinforcement por API (SketchEditScope no soportado).\n"
                    "- No se pueden añadir 'shafts' como huecos por Create(...).\n\n"
                    "Se crearon {} model lines con los contornos de los shafts y quedaron seleccionadas.\n"
                    "Ahora usa 'Editar contorno' del AreaReinforcement manualmente y toma esas líneas."
                    .format(n),
                )
            except Exception as ex:
                TaskDialog.Show("BIMTools", "Error editando boundary:\n{}".format(ex))

