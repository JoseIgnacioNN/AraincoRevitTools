# -*- coding: utf-8 -*-
"""
Unir geometría entre elementos de **material estructural hormigón** (Concrete) en la **vista activa**.

Paquete portable: vive en ``<pushbutton>/scripts/`` (sin depender de scripts/ de la extensión).

- Alcance: ejemplares a unir = **hormigón** de ``FilteredElementCollector(document, view.Id)``.
- Candidatos por **solape de cajas** (``BoundingBoxIntersectsFilter`` + contorno inflado); la API
  de intersección sólida a menudo no devuelve pares con contacto mínimo (columna/viga).
- Criterio de hormigón: ``join_geometry_material_concrete.material_estructural_es_concrete``.
- Tras unir, en pares **forjado + muro/viga/pilar/cimentación**: ``SwitchJoinOrder`` si hace
  falta para que el **forjado sea el recortado** (no el cortante), vía
  ``IsCuttingElementInJoin`` (mismo criterio que el flujo RPS de referencia).
- Barra de progreso: ``pyrevit.forms.ProgressBar`` con acento BIMTools (91,192,222),
  patrón ``armado_muros_lineales`` — fases lectura, candidatos y unión.
- Pantalla de inicio / avisos: diálogo WPF shell BIMTools (``join_geometry_instruction_dialog``).
- Fallos «Can't cut joined element»: ``IFailuresPreprocessor`` aplica la resolución API
  equivalente a *Unjoin elements* (``FailureResolutionType.UnjoinElements`` / ``DetachElements``)
  y devuelve ``ProceedWithCommit`` para no mostrar el diálogo modal de Revit.
- API: ``JoinGeometryUtils`` (Revit 2024+; pyRevit / IronPython).
- Durante lectura / candidatos / unión: se deshabilita la ventana principal de Revit
  (``user32.EnableWindow``) para que no se lancen otros comandos; la barra de pyRevit
  es otra ventana de nivel superior y sigue mostrando el progreso.

Copia portable: ``BIMTools.tab/Modelado.panel/11_UnirGeometriaHormigonVista.pushbutton/scripts/``.
Tras editar aquí, sincronice con esa carpeta (ver ``ESTRUCTURA_PORTABLE.txt``).
"""

from __future__ import print_function

import clr
import System

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BoundingBoxIntersectsFilter,
    ElementCategoryFilter,
    ElementFilter,
    ElementId,
    FailureProcessingResult,
    FailureResolutionType,
    FailureSeverity,
    FilteredElementCollector,
    IFailuresPreprocessor,
    JoinGeometryUtils,
    LogicalOrFilter,
    Outline,
    Transaction,
    UnitUtils,
    UnitTypeId,
    View,
    ViewSchedule,
    ViewSheet,
    XYZ,
)
from System.Collections.Generic import List
try:
    from pyrevit import forms as _pyrevit_forms
except Exception:
    _pyrevit_forms = None

from join_geometry_material_concrete import (
    _CATS_ESCANEO_MATERIAL_ESTRUCTURAL,
    material_estructural_es_concrete,
)

_PBAR_BASE_LECTURA = u"Arainco: Unir geom. hormigón — lectura"
_PBAR_BASE_CANDIDATOS = u"Arainco: Unir geom. hormigón — candidatos"
_PBAR_BASE_UNION = u"Arainco: Unir geom. hormigón — uniendo"


def _builtin_cannot_cut_joined_failure_ids():
    """``FailureDefinitionId`` conocidos para «Can't cut joined element» (por API)."""
    out = []
    try:
        from Autodesk.Revit.DB import BuiltInFailures

        for group_name in (u"CutFailures", u"JoinElementsFailures"):
            grp = getattr(BuiltInFailures, group_name, None)
            if grp is None:
                continue
            for attr in (
                u"CannotCutJoinedElement",
                u"CannotCutJoinedElements",
                u"CannotCutJoinedElementError",
            ):
                x = getattr(grp, attr, None)
                if x is not None:
                    out.append(x)
    except Exception:
        pass
    return out


def _failure_text_cut_joined(fma):
    t = None
    for meth in (u"GetDescriptionText", u"GetDescriptionString", u"GetDescription"):
        try:
            t = getattr(fma, meth)()
            if t:
                break
        except Exception:
            pass
    if not t:
        try:
            t = fma.GetDefaultResolutionCaption()
        except Exception:
            t = None
    if not t:
        return False
    try:
        s = unicode(t).lower()
    except Exception:
        s = str(t).lower()
    return (
        u"can't cut joined" in s
        or u"cannot cut joined" in s
        or u"no se puede cortar" in s
        or u"elemento unido" in s
    )


def _failure_id_matches(fid, known_ids):
    if fid is None or not known_ids:
        return False
    for kid in known_ids:
        try:
            if kid == fid:
                return True
        except Exception:
            pass
    return False


def _resolution_types_unjoin_ui_order():
    """Prioridad como el botón «Unjoin Elements» del diálogo de Revit."""
    out = []
    for name in (u"UnjoinElements", u"DetachElements"):
        t = getattr(FailureResolutionType, name, None)
        if t is not None:
            out.append(t)
    return out


def _failure_has_resolutions(fma):
    if fma is None:
        return False
    try:
        if hasattr(fma, u"HasResolutions") and bool(fma.HasResolutions()):
            return True
    except System.Exception:
        pass
    try:
        if hasattr(fma, u"GetNumberOfResolutions"):
            return int(fma.GetNumberOfResolutions()) > 0
    except System.Exception:
        pass
    return False


def _try_resolve_failure_with_unjoin(failures_accessor, fma):
    """
    Aplica la resolución equivalente al botón *Unjoin Elements* del diálogo.
    Devuelve True si se llamó a ``ResolveFailure``.
    """
    if failures_accessor is None or fma is None:
        return False

    def _permitted(rt):
        try:
            return bool(failures_accessor.IsFailureResolutionPermitted(fma, rt))
        except System.Exception:
            pass
        try:
            if hasattr(fma, u"HasResolutionOfType"):
                return bool(fma.HasResolutionOfType(rt))
        except System.Exception:
            pass
        return False

    for rt in _resolution_types_unjoin_ui_order():
        if not _permitted(rt):
            continue
        try:
            failures_accessor.SetCurrentResolutionType(fma, rt)
            failures_accessor.ResolveFailure(fma)
            return True
        except System.Exception:
            continue

    # Resolución por defecto del fallo (suele ser Unjoin Elements en este error).
    try:
        if _failure_has_resolutions(fma):
            try:
                if hasattr(failures_accessor, u"IsFailureResolutionPermitted"):
                    if not bool(failures_accessor.IsFailureResolutionPermitted(fma)):
                        return False
            except (System.Exception, TypeError):
                pass
            failures_accessor.ResolveFailure(fma)
            return True
    except System.Exception:
        pass
    return False


class _JoinGeomFailuresPreprocessor(IFailuresPreprocessor):
    """
    Trata la cola de fallos al ``Commit`` de un lote de ``JoinGeometry`` / ``SwitchJoinOrder``.

    - **Errores** tipo «Can't cut joined element»: resolución API *Unjoin elements*
      (``UnjoinElements`` / ``DetachElements``) cuando Revit la permite.
    - **Warnings**: se eliminan de la cola.
    - Si se resolvió algún error: ``ProceedWithCommit`` (obligatorio; con ``Continue``
      Revit muestra el diálogo aunque el fallo se haya tratado).
    - Si no hay resolución aplicable: ``DeleteError`` como respaldo.
    """

    def _iter_msgs(self, failures_accessor):
        if failures_accessor is None:
            return
        try:
            fmsgs = failures_accessor.GetFailureMessages()
        except System.Exception:
            return
        if fmsgs is None:
            return
        try:
            n = int(fmsgs.Count)
        except System.Exception:
            n = 0
        for i in range(n):
            f = None
            try:
                f = fmsgs.get_Item(i)
            except System.Exception:
                try:
                    f = fmsgs[i]
                except System.Exception:
                    f = None
            if f is not None:
                yield f

    def PreprocessFailures(self, failures_accessor):
        if failures_accessor is None:
            return FailureProcessingResult.Continue
        _known_cut = _builtin_cannot_cut_joined_failure_ids()
        _msgs = []
        for f in self._iter_msgs(failures_accessor):
            _msgs.append(f)
        resolved_or_cleared = False
        for f in _msgs:
            try:
                sev = f.GetSeverity()
            except System.Exception:
                continue
            if sev == FailureSeverity.Warning:
                try:
                    failures_accessor.DeleteWarning(f)
                    resolved_or_cleared = True
                except System.Exception:
                    pass
                continue
            if sev != FailureSeverity.Error:
                continue
            _fid = None
            try:
                _fid = f.GetFailureDefinitionId()
            except System.Exception:
                pass
            _is_cut_joined = _failure_id_matches(_fid, _known_cut) or _failure_text_cut_joined(
                f
            )
            if _is_cut_joined:
                if _try_resolve_failure_with_unjoin(failures_accessor, f):
                    resolved_or_cleared = True
                    continue
            try:
                failures_accessor.DeleteError(f)
                resolved_or_cleared = True
            except System.Exception:
                pass
        # Continue mostraría el diálogo aunque se haya resuelto el fallo.
        if resolved_or_cleared:
            return FailureProcessingResult.ProceedWithCommit
        return FailureProcessingResult.Continue


def _pbar_phase_title(base_title, total):
    """Título inicial 0/N (mismo patrón que Armado Muros)."""
    try:
        t = max(int(total), 1)
    except Exception:
        t = 1
    return u"{} 0/{}".format(base_title, t)


def _pbar_start(title, count):
    u"""
    ``forms.ProgressBar`` de pyRevit con acento BIMTools (91,192,222),
    igual que ``armado_muros_lineales._ml_pbar_start``.
    """
    if _pyrevit_forms is None or count is None or int(count) < 1:
        return None
    try:
        pb = _pyrevit_forms.ProgressBar(
            title=title,
            cancellable=False,
        )
        try:
            from System.Windows.Media import Color, SolidColorBrush

            r, g, b = (91, 192, 222)
            pb.Resources[u"pyRevitAccentBrush"] = SolidColorBrush(Color.FromRgb(r, g, b))
        except Exception:
            pass
        return pb
    except Exception:
        return None


def _pbar_step(pb, current_index, count, base_title):
    u"""*current_index*: 0…count-1. Actualiza barra y título con contador."""
    if pb is None:
        return
    c = int(count) if count else 0
    if c < 1:
        c = 1
    i = int(current_index) + 1
    try:
        if hasattr(pb, u"update_progress"):
            try:
                pb.update_progress(i, max_value=c)
            except TypeError:
                try:
                    pb.update_progress(i, max=c)
                except Exception:
                    pass
    except Exception:
        pass
    try:
        pb.title = u"{} {}/{}".format(base_title, i, c)
    except Exception:
        pass


def _pbar_enter(pb):
    if pb is None:
        return False
    try:
        pb.__enter__()
        return True
    except Exception:
        return False


def _pbar_exit(pb, pbar_open):
    if not pbar_open or pb is None:
        return
    try:
        pb.__exit__(None, None, None)
    except Exception:
        pass


def _hwnd_to_int(hwnd):
    if hwnd is None:
        return None
    try:
        if hasattr(hwnd, u"ToInt32"):
            return int(hwnd.ToInt32())
    except Exception:
        pass
    try:
        if hasattr(hwnd, u"ToInt64"):
            return int(hwnd.ToInt64())
    except Exception:
        pass
    try:
        return int(hwnd)
    except Exception:
        return None


def _revit_main_window_set_enabled(revit, enable):
    """
    Habilita o deshabilita la ventana principal de Revit (cinta, canvas, etc.).
    *revit*: normalmente ``UIApplication`` (``__revit__`` de pyRevit).
    """
    if revit is None:
        return
    try:
        from revit_wpf_window_position import revit_main_hwnd
    except Exception:
        return
    hwnd = revit_main_hwnd(revit)
    h = _hwnd_to_int(hwnd)
    if h is None or h == 0:
        return
    try:
        import ctypes
    except Exception:
        return
    try:
        ctypes.windll.user32.EnableWindow(h, 1 if enable else 0)
    except Exception:
        pass


class _BloquearComandosRevit(object):
    """
    Deshabilita la ventana principal mientras dura el trabajo; al salir, la restaura.
    Así el usuario no puede activar otras órdenes; la barra de ``ProgressBar`` de
    pyRevit no es hija de ese HWND, así que sigue actualizándose.
    """

    def __init__(self, revit):
        self._revit = revit
        self._touched = False

    def __enter__(self):
        _revit_main_window_set_enabled(self._revit, False)
        self._touched = True
        return self

    def __exit__(self, _exc_type, _exc, _tb):
        if self._touched:
            _revit_main_window_set_enabled(self._revit, True)
        self._touched = False
        return False


def _exc_text(ex):
    try:
        return unicode(ex)
    except Exception:
        try:
            return str(ex)
        except Exception:
            return u""


def _element_id_to_int(eid):
    if eid is None or eid == ElementId.InvalidElementId:
        return None
    try:
        return int(eid.Value)
    except Exception:
        try:
            return int(eid.IntegerValue)
        except Exception:
            return None


def _vista_permitida(view):
    if view is None:
        return False, u"No hay vista activa."
    if isinstance(view, ViewSheet):
        return (
            False,
            u"La vista activa es una hoja. Abra una vista de modelo (planta, 3D, corte, etc.) y vuelva a ejecutar la herramienta.",
        )
    if isinstance(view, ViewSchedule):
        return (
            False,
            u"La vista activa es un cuadro de planificación. Abra una vista de modelo y vuelva a ejecutar la herramienta.",
        )
    if not isinstance(view, View):
        return (
            False,
            u"La vista activa no es una vista de modelo válida para esta herramienta.",
        )
    return True, u""


def _categoria_structural_or_filter():
    """
    OR de ``ElementCategoryFilter`` para las categorías estructurales de escaneo.
    Reduce el ``FilteredElementCollector`` a nivel documento a los tipos relevantes.
    """
    L = List[ElementFilter]()
    for c in _CATS_ESCANEO_MATERIAL_ESTRUCTURAL:
        L.Add(ElementCategoryFilter(c))
    return LogicalOrFilter(L)


def _outline_bbox_inflado(elem, doc):
    """
    Caja mínima del elemento en el modelo, inflada ~30 mm: candidatos con contacto
    aunque la intersección sólida estricta de la API quede vacía.
    """
    if elem is None or doc is None:
        return None
    try:
        bb = elem.get_BoundingBox(None)
        if bb is None:
            return None
        try:
            if not bb.Enabled:
                return None
        except Exception:
            pass
        try:
            pad = UnitUtils.ConvertToInternalUnits(30.0, UnitTypeId.Millimeters)
        except Exception:
            pad = 0.1
        a, b = bb.Min, bb.Max
        p0 = XYZ(a.X - pad, a.Y - pad, a.Z - pad)
        p1 = XYZ(b.X + pad, b.Y + pad, b.Z + pad)
        return Outline(p0, p1)
    except Exception:
        return None


def _recoger_hormigon_en_vista(doc, view, pbar_cats=None):
    """
    Elementos con material estructural concrete, en categorías habituales, visibles en la vista.
    ``pbar_cats``: un paso al terminar cada categoría (1/3 en la UI).
    """
    out = []
    cats = _CATS_ESCANEO_MATERIAL_ESTRUCTURAL
    n_cat = len(cats) if cats else 1
    for i, cat in enumerate(cats):
        try:
            for el in (
                FilteredElementCollector(doc, view.Id)
                .OfCategory(cat)
                .WhereElementIsNotElementType()
            ):
                if material_estructural_es_concrete(el):
                    out.append(el)
        except Exception:
            pass
        if pbar_cats is not None and n_cat > 0:
            _pbar_step(
                pbar_cats, i, n_cat, _PBAR_BASE_LECTURA
            )
    return out


def _pares_unicos_por_caja(doc, view, elements_concrete, pbar_cajas=None):
    """
    Pares (id1, id2) con id1 < id2, si el **BoundingBox** del elemento (inflado) cruza
    otro candidato estructural y el otro Id está en ``allowed`` (hormigón en la vista).
    """
    _ = view
    allowed = set()
    for el in elements_concrete:
        if el and el.Id:
            ni = _element_id_to_int(el.Id)
            if ni is not None:
                allowed.add(ni)
    if not allowed:
        return
    cat_or = _categoria_structural_or_filter()
    processed = set()
    nint = _element_id_to_int
    n_el = max(len(elements_concrete), 1)
    for idx, e in enumerate(elements_concrete):
        if pbar_cajas is not None:
            _pbar_step(
                pbar_cajas, idx, n_el, _PBAR_BASE_CANDIDATOS
            )
        if e is None:
            continue
        ol = _outline_bbox_inflado(e, doc)
        if ol is None:
            continue
        try:
            bf = BoundingBoxIntersectsFilter(ol)
        except Exception:
            continue
        try:
            oids = (
                FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .WherePasses(bf)
                .WherePasses(cat_or)
                .ToElementIds()
            )
        except Exception:
            continue
        eid = e.Id
        ne = nint(eid)
        if ne is None:
            continue
        for oid in oids:
            if oid is None or oid == eid:
                continue
            no = nint(oid)
            if no is None or no not in allowed:
                continue
            a, b = (ne, no) if ne < no else (no, ne)
            key = (a, b)
            if key in processed:
                continue
            processed.add(key)
            yield ElementId(a), ElementId(b)


def _es_floor(elem):
    try:
        c = elem.Category
        if c is None:
            return False
        return int(c.Id.IntegerValue) == int(BuiltInCategory.OST_Floors)
    except Exception:
        return False


def _join_geometry_try_both_orders(doc, a, b):
    """
    ``JoinGeometry`` a veces falla con «cannot be joined» / ``secondElement`` según el
    orden de los argumentos. Prueba (a,b) y, si aplica, (b,a).
    Devuelve ``(True, False)`` si unió con el orden original, ``(True, True)`` si con orden
    invertido, ``(False, None)`` si ambos fallan (mensaje del último intento).
    """
    try:
        JoinGeometryUtils.JoinGeometry(doc, a, b)
        return True, False
    except Exception as ex1:
        msg1 = _exc_text(ex1)
        low = msg1.lower() if msg1 else u""
        reintentar = (
            u"second" in low
            and u"element" in low
            or u"cannot be joined" in low
            or u"no se pueden unir" in low
            or u"elements cannot be joined" in low
        )
        if not reintentar:
            return False, msg1
        try:
            JoinGeometryUtils.JoinGeometry(doc, b, a)
            return True, True
        except Exception as ex2:
            return False, _exc_text(ex2) or msg1


def _switch_forjado_recortado_por_otro(doc, a, b, err_switch):
    """
    Si un elemento es forjado (Floor) y el otro no, y el forjado actúa como **cortante**,
    invierte el orden para que la losa quede recortada (viga/muro/pilar/zapata cortan al forjado).
    """
    if a is None or b is None or doc is None:
        return False
    floor = other = None
    if _es_floor(a) and not _es_floor(b):
        floor, other = a, b
    elif _es_floor(b) and not _es_floor(a):
        floor, other = b, a
    else:
        return False
    try:
        if not JoinGeometryUtils.AreElementsJoined(doc, floor, other):
            return False
    except Exception:
        return False
    try:
        if not bool(
            JoinGeometryUtils.IsCuttingElementInJoin(doc, floor, other)
        ):
            return False
    except Exception:
        return False
    try:
        JoinGeometryUtils.SwitchJoinOrder(doc, floor, other)
        return True
    except Exception as ex:
        if err_switch is not None and len(err_switch) < 8:
            err_switch.append(
                u"Switch forjado: {0}".format(_exc_text(ex))
            )
        return False


def run(revit):
    from join_geometry_instruction_dialog import (
        mostrar_aviso,
        show_selection_instructions,
    )

    uidoc = revit.ActiveUIDocument
    if uidoc is None:
        mostrar_aviso(revit, u"No hay documento activo.")
        return

    doc = uidoc.Document
    view = uidoc.ActiveView
    ok, msg = _vista_permitida(view)
    if not ok:
        mostrar_aviso(revit, msg)
        return

    if not show_selection_instructions(revit):
        return

    with _BloquearComandosRevit(revit):
        n_cat = len(_CATS_ESCANEO_MATERIAL_ESTRUCTURAL) or 1
        _pb1 = _pbar_start(_pbar_phase_title(_PBAR_BASE_LECTURA, n_cat), n_cat)
        _pbar1_open = _pbar_enter(_pb1)
        try:
            elements_concrete = _recoger_hormigon_en_vista(
                doc, view, _pb1 if _pbar1_open else None
            )
        finally:
            _pbar_exit(_pb1, _pbar1_open)

        if not elements_concrete:
            mostrar_aviso(
                revit,
                u"No se encontraron elementos de hormigón en la vista activa.",
                u"Revise muros, forjados, pilares y cimentación con material "
                u"estructural Concrete.",
            )
            return

        n_el = len(elements_concrete)
        _pb2 = _pbar_start(_pbar_phase_title(_PBAR_BASE_CANDIDATOS, n_el), n_el)
        _pbar2_open = _pbar_enter(_pb2)
        try:
            pairs = list(
                _pares_unicos_por_caja(
                    doc, view, elements_concrete, _pb2 if _pbar2_open else None
                )
            )
        finally:
            _pbar_exit(_pb2, _pbar2_open)

        if not pairs:
            mostrar_aviso(
                revit,
                u"No se detectaron pares candidatos a unir.",
                u"Hay ejemplares de hormigón en la vista, pero ninguna caja "
                u"solapada entre sí.",
            )
            return

        ya_unidos = 0
        nuevos = 0
        fallos = 0
        inversiones = 0
        err_switch = []

        n_pairs = len(pairs)
        _pb = _pbar_start(_pbar_phase_title(_PBAR_BASE_UNION, n_pairs), n_pairs)
        _pbar_open = _pbar_enter(_pb)
        tx = Transaction(
            doc, u"Arainco: Unir geometría hormigón (vista activa)"
        )
        try:
            _fho = tx.GetFailureHandlingOptions()
            _fho.SetFailuresPreprocessor(_JoinGeomFailuresPreprocessor())
            try:
                _fho.SetClearAfterRollback(True)
            except System.Exception:
                pass
            tx.SetFailureHandlingOptions(_fho)
        except System.Exception:
            pass
        tx.Start()
        try:
            for idx, (ida, idb) in enumerate(pairs):
                if _pbar_open:
                    _pbar_step(_pb, idx, n_pairs, _PBAR_BASE_UNION)
                a = doc.GetElement(ida)
                b = doc.GetElement(idb)
                if a is None or b is None:
                    fallos += 1
                    continue
                try:
                    already = JoinGeometryUtils.AreElementsJoined(doc, a, b)
                except Exception:
                    try:
                        already = JoinGeometryUtils.AreElementsJoined(
                            doc, a.Id, b.Id
                        )
                    except Exception:
                        already = False
                if already:
                    ya_unidos += 1
                else:
                    ok_join, swap_or_err = _join_geometry_try_both_orders(
                        doc, a, b
                    )
                    if ok_join:
                        nuevos += 1
                    else:
                        fallos += 1
                        continue
                if _switch_forjado_recortado_por_otro(doc, a, b, err_switch):
                    inversiones += 1
            tx.Commit()
        except Exception as ex:
            try:
                tx.RollBack()
            except Exception:
                pass
            mostrar_aviso(
                revit,
                u"No se pudo completar la unión de geometría.",
                _exc_text(ex),
            )
            return
        finally:
            _pbar_exit(_pb, _pbar_open)

    view_name = getattr(view, u"Name", u"?") or u"?"
    if nuevos > 0 or inversiones > 0:
        detalle = u"Vista: {0}. Nuevas uniones: {1}.".format(view_name, nuevos)
        if inversiones > 0:
            detalle += u" Órdenes de corte ajustados: {0}.".format(inversiones)
        mostrar_aviso(revit, u"Proceso completado correctamente.", detalle)
    elif ya_unidos > 0 and fallos == 0:
        mostrar_aviso(
            revit,
            u"Proceso completado correctamente.",
            u"Vista: {0}. Los pares candidatos ya estaban unidos.".format(
                view_name
            ),
        )
    elif fallos > 0 and nuevos == 0:
        mostrar_aviso(
            revit,
            u"No se pudieron crear uniones nuevas.",
            u"Vista: {0}. Pares sin unir: {1}.".format(view_name, fallos),
        )
    else:
        mostrar_aviso(
            revit,
            u"Proceso completado correctamente.",
            u"Vista: {0}.".format(view_name),
        )
