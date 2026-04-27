# -*- coding: utf-8 -*-
"""
Módulo compartido: Malla en Losa (Area Reinforcement en losas).
Crea AreaReinforcement en losas de hormigón armado.
Usado por el botón 08_CrearAreaReinforcementRPS.
Sin dependencias entre botones: cada uno importa este módulo independientemente.
"""

import math
import os
import weakref
import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")

import Autodesk.Revit.DB as _RDB
# IronPython: importar "Family" al namespace global puede colisionar (p. ej. con type/ViewType);
# OfClass() debe recibir System.Type explícito.
_AR_REVIT_FAMILY_CLR_TYPE = clr.GetClrType(_RDB.Family)
_AR_REVIT_FAMILY_SYMBOL_CLR_TYPE = clr.GetClrType(_RDB.FamilySymbol)

from System.Collections.Generic import List
from System.Windows.Markup import XamlReader
from System import Action, EventHandler
from System.Windows import RoutedEventHandler, SizeToContent
from System.Windows.Threading import DispatcherPriority
from System.Windows.Input import Key, KeyBinding, ModifierKeys, ApplicationCommands, CommandBinding
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption
from System import Uri, UriKind
import System

from revit_wpf_window_position import (
    position_wpf_window_top_left_at_active_view,
    revit_main_hwnd,
)

from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    Curve,
    CurveLoop,
    ElementId,
    ElementTypeGroup,
    FilteredElementCollector,
    Floor,
    GeometryInstance,
    IndependentTag,
    JoinGeometryUtils,
    Options,
    Reference,
    Sketch,
    Solid,
    StorageType,
    SubTransaction,
    TagMode,
    TagOrientation,
    Transaction,
    UnitUtils,
    UnitTypeId,
    View3D,
    XYZ,
)
from Autodesk.Revit.DB.Structure import (
    AreaReinforcement,
    AreaReinforcementLayerType,
    AreaReinforcementType,
    RebarBarType,
    RebarHookType,
)
from Autodesk.Revit.UI import TaskDialog, ExternalEvent, IExternalEventHandler
from Autodesk.Revit.UI.Selection import ObjectType

# ── Constantes (rutas desde raíz de la extensión) ───────────────────────────
_script_dir = os.path.dirname(os.path.abspath(__file__))
_ext_root = os.path.dirname(_script_dir)
_panel_dir = os.path.join(_ext_root, "BIMTools.tab", "Armadura.panel")
_tab_dir = os.path.join(_ext_root, "BIMTools.tab")
_LOGO_PATHS = [
    os.path.join(_panel_dir, "08_CrearAreaReinforcementRPS.pushbutton", "empresa_logo.png"),
    os.path.join(_panel_dir, "22_EnfierradoFundacionAislada.pushbutton", "empresa_logo.png"),
    os.path.join(_panel_dir, "08_CrearAreaReinforcementRPS.pushbutton", "logo.png"),
    os.path.join(_tab_dir, "Incidencias.panel", "Incidencias.stack", "01_BIMIssue.pushbutton", "logo.png"),
]

# Misma línea de diseño que Refuerzo Borde Losa (barras_bordes_losa_gancho_empotramiento).
_APPDOMAIN_WINDOW_KEY = "BIMTools.AreaReinforcementLosa.ActiveWindow"
_TOOL_TASK_DIALOG_TITLE = u"BIMTools — Malla en Losa"
_WINDOW_OPEN_MS = 180
_WINDOW_CLOSE_MS = 180


def _task_dialog_show(title, message, wpf_window=None):
    """TaskDialog detrás del WPF si Topmost=True; igual que Refuerzo Borde Losa."""
    if wpf_window is not None:
        try:
            wpf_window.Topmost = False
        except Exception:
            pass
    try:
        TaskDialog.Show(title, message)
    finally:
        if wpf_window is not None:
            try:
                wpf_window.Topmost = True
            except Exception:
                pass


def _wpf_from_window_ref(window_ref):
    win = window_ref() if window_ref else None
    return getattr(win, "_win", None) if win else None


def _get_active_window():
    try:
        win = System.AppDomain.CurrentDomain.GetData(_APPDOMAIN_WINDOW_KEY)
    except Exception:
        return None
    if win is None:
        return None
    try:
        _ = win.Title
    except Exception:
        _clear_active_window()
        return None
    try:
        if hasattr(win, "IsLoaded") and (not win.IsLoaded):
            _clear_active_window()
            return None
    except Exception:
        pass
    return win


def _set_active_window(win):
    try:
        System.AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, win)
    except Exception:
        pass


def _clear_active_window():
    try:
        System.AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, None)
    except Exception:
        pass


_DEFAULT_SPACING_MM = 150
# Offset de recubrimiento (mm) aplicado a las curvas del perímetro de la losa para el Area Reinforcement RPS.
OFFSET_RECUBRIMIENTO_MM = 20.0

# Etiquetado automático en vista de planta (mismo criterio que etiquetar_area_reinforcement_rps.py)
_AR_TAG_ADD_LEADER = False
_AR_TAG_ORIENTATION = TagOrientation.Horizontal
# Si no es None, fuerza ese FamilySymbol y se ignora la lógica malla inferior / doble malla.
_AR_TAG_SYMBOL_ID = None
# Familia y tipos de etiqueta (como en seleccionar_tipo_etiqueta_por_familia_rps.py + IndependentTag.Create con symbol id).
_AR_TAG_FAMILY_NAME = u"EST_A_STRUCTURAL AREA REINFORCEMENT TAG_PLANTA_MALLA"
_AR_TAG_TYPE_MALLA_INFERIOR = u"Malla Inferior"
_AR_TAG_TYPE_DOBLE_MALLA = u"Doble Malla"
_AR_TAG_NORM_CASE_INSENSITIVE = True

# Ancho del diálogo: fila de combos + pads GroupBox/chrome + borde ventana (sin hueco extra a la derecha).
_AR_LOSA_INPUT_COLS_PER_ROW = 2
_AR_LOSA_COMBO_WIDTH_PX = 110
# Columna central @ (márgenes 6+6 + glifo): coherente con Grid Column Auto en XAML.
_AR_LOSA_DIAM_ESP_AT_COL_PX = 28
_AR_LOSA_BLOCK_PAD_H_PX = 16
_AR_LOSA_GROUPBOX_PAD_H_PX = 16
# Debe cuadrar con Padding horizontal del Border raíz del XAML (izq. + der.).
_AR_LOSA_OUTER_PAD_H_PX = 28
# Cabecera «Malla en Losa» + logo + cerrar (mínimo por si el contenido calculado es menor).
_AR_LOSA_WIDTH_TITLE_MIN_PX = 288


def _area_reinforcement_losa_form_width_px(input_cols_per_row=None, combo_width_px=None):
    """
    Ancho horizontal del formulario en px.
    ``input_cols_per_row``: columnas de controles en la fila más ancha (Major/Minor: Diam. + Esp.).
    ``combo_width_px``: ancho fijo de cada ComboBox (debe coincidir con el estilo ``Combo`` en XAML).
    """
    cols = int(input_cols_per_row or _AR_LOSA_INPUT_COLS_PER_ROW)
    cols = max(1, cols)
    c = int(combo_width_px or _AR_LOSA_COMBO_WIDTH_PX)
    row_inner = cols * c + _AR_LOSA_DIAM_ESP_AT_COL_PX + _AR_LOSA_BLOCK_PAD_H_PX
    w = row_inner + _AR_LOSA_GROUPBOX_PAD_H_PX + _AR_LOSA_OUTER_PAD_H_PX
    w = max(w, _AR_LOSA_WIDTH_TITLE_MIN_PX)
    return int((int(w) + 3) // 4 * 4)


def _element_id_int(eid):
    """Revit 2024+: ElementId.Value; versiones anteriores: IntegerValue."""
    if eid is None:
        return None
    try:
        v = getattr(eid, "Value", None)
        if v is not None:
            return int(v)
    except Exception:
        pass
    try:
        return int(eid.IntegerValue)
    except Exception:
        pass
    try:
        return int(eid)
    except Exception:
        return None


# ── Funciones auxiliares ─────────────────────────────────────────────────────
def _obtener_curvas_sketch(floor, document):
    """Obtiene las curvas del perímetro exterior del sketch de la losa (compatible con Create(7 params))."""
    try:
        sketch_id = floor.SketchId
        if sketch_id is None or sketch_id == ElementId.InvalidElementId:
            return None
        sketch = document.GetElement(sketch_id)
        if sketch is None or not isinstance(sketch, Sketch):
            return None
        profile = sketch.Profile
        if profile is None:
            return None
        n_loops = profile.Size
        if n_loops < 1:
            return None
        curve_array = profile.get_Item(0)
        if curve_array is None:
            return None
        curves = []
        n_curves = curve_array.Size
        for j in range(n_curves):
            c = curve_array.get_Item(j)
            if c is not None:
                curves.append(c)
        return curves if curves else None
    except Exception:
        return None


def _obtener_curvas_sketch_con_recubrimiento(floor, document):
    """
    Obtiene las curvas del perímetro exterior del sketch de la losa y les aplica
    un offset hacia el interior (recubrimiento) de OFFSET_RECUBRIMIENTO_MM mm.
    Usado por el botón Crear Area Reinf. RPS.
    """
    curves = _obtener_curvas_sketch(floor, document)
    if not curves or len(curves) < 2:
        return curves
    try:
        offset_internal = UnitUtils.ConvertToInternalUnits(
            OFFSET_RECUBRIMIENTO_MM, UnitTypeId.Millimeters
        )
        curve_list = List[Curve](curves)
        loop = CurveLoop.Create(curve_list)
        if loop is None or not loop.HasPlane():
            return curves
        normal = XYZ(0, 0, 1)

        # CreateViaOffset depende del sentido del CurveLoop (CW/CCW). Para asegurar que el
        # recubrimiento siempre vaya "hacia adentro", probamos ambos signos y elegimos el
        # que reduce el perímetro (longitud total) respecto al original.
        try:
            base_len = loop.GetExactLength()
        except Exception:
            base_len = None

        offset_candidates = []
        for sign in (1.0, -1.0):
            try:
                cl = CurveLoop.CreateViaOffset(loop, sign * offset_internal, normal)
                if cl is None:
                    continue
                try:
                    cand_len = cl.GetExactLength()
                except Exception:
                    cand_len = None
                offset_candidates.append((cl, cand_len))
            except Exception:
                continue

        if not offset_candidates:
            return curves

        if base_len is not None:
            best = None
            best_len = None
            for cl, cand_len in offset_candidates:
                if cand_len is None:
                    continue
                if best is None or cand_len < best_len:
                    best = cl
                    best_len = cand_len
            if best is None:
                best = offset_candidates[0][0]
        else:
            best = offset_candidates[0][0]

        return [c for c in best]
    except Exception:
        return curves


def _obtener_direccion_principal(curves):
    """Obtiene la dirección principal a partir de la primera curva del perfil de la losa."""
    try:
        if curves and len(curves) > 0:
            first = curves[0]
            if hasattr(first, "GetEndPoint"):
                p0 = first.GetEndPoint(0)
                p1 = first.GetEndPoint(1)
                dx = p1.X - p0.X
                dy = p1.Y - p0.Y
                dz = p1.Z - p0.Z
                length = (dx * dx + dy * dy + dz * dz) ** 0.5
                if length > 1e-6:
                    return XYZ(dx / length, dy / length, dz / length)
        return XYZ(1, 0, 0)
    except Exception:
        return XYZ(1, 0, 0)


def _obtener_layout_dir_area_reinforcement(floor, curves):
    """
    Dirección de trazado (layout) para ``AreaReinforcement.Create``:
    1) ``Floor.SpanDirectionAngle`` en losas estructurales (radianes en planta);
       se proyecta al plano del perímetro offset si hay curvas (losas inclinadas).
    2) Si no aplica o la proyección degenera: ``_obtener_direccion_principal(curves)``.
    """
    base = None
    if floor is not None:
        try:
            ang = float(floor.SpanDirectionAngle)
            c = math.cos(ang)
            s = math.sin(ang)
            base = XYZ(c, s, 0.0)
        except Exception:
            base = None
    if base is not None:
        try:
            if curves and len(curves) >= 2:
                curve_list = List[Curve](curves)
                loop = CurveLoop.Create(curve_list)
                if loop is not None and loop.HasPlane():
                    n = loop.GetPlane().Normal
                    dot = (
                        float(base.X) * float(n.X)
                        + float(base.Y) * float(n.Y)
                        + float(base.Z) * float(n.Z)
                    )
                    vx = float(base.X) - dot * float(n.X)
                    vy = float(base.Y) - dot * float(n.Y)
                    vz = float(base.Z) - dot * float(n.Z)
                    L = (vx * vx + vy * vy + vz * vz) ** 0.5
                    if L > 1e-9:
                        return XYZ(vx / L, vy / L, vz / L)
        except Exception:
            pass
        try:
            return XYZ(float(base.X), float(base.Y), float(base.Z))
        except Exception:
            pass
    return _obtener_direccion_principal(curves) if curves else XYZ(1, 0, 0)


def _get_first_rebar_hook_type_id(document):
    """Obtiene el ID del primer RebarHookType (o InvalidElementId si no hay)."""
    from Autodesk.Revit.DB import ElementId, FilteredElementCollector
    from Autodesk.Revit.DB.Structure import RebarHookType
    try:
        for elem in FilteredElementCollector(document).OfClass(RebarHookType):
            if elem:
                return elem.Id
    except Exception:
        pass
    return ElementId.InvalidElementId


def _losa_collect_joined_element_ids(document, floor):
    """``ElementId`` de elementos con *Join Geometry* respecto a la losa."""
    if document is None or floor is None:
        return []
    out = []
    try:
        raw = JoinGeometryUtils.GetJoinedElements(document, floor)
    except Exception:
        return []
    if raw is None:
        return []
    try:
        for eid in raw:
            if eid is None:
                continue
            ni = _element_id_int(eid)
            if ni is not None:
                out.append(eid)
    except Exception:
        pass
    return out


def _losa_es_muro_viga_o_columna(element):
    """True si el elemento es muro, viga estructural o columna (incl. arquitectónica)."""
    if element is None:
        return False
    try:
        cat = element.Category
    except Exception:
        cat = None
    if cat is None:
        return False
    try:
        cid = int(cat.Id.IntegerValue)
    except Exception:
        return False
    for bic in (
        BuiltInCategory.OST_Walls,
        BuiltInCategory.OST_StructuralFraming,
        BuiltInCategory.OST_StructuralColumns,
        BuiltInCategory.OST_Columns,
    ):
        try:
            if cid == int(bic):
                return True
        except Exception:
            continue
    return False


def _losa_switch_join_order_losa_recortada(document, floor, other, errores):
    """
    Tras «Unir geometría», asegura que muros/vigas/columnas sean los cortantes
    y la losa el elemento recortado: si la API indica que la losa corta al otro,
    invierte el orden (firstElement = losa, secondElement = otro).
    """
    if document is None or floor is None or other is None:
        return
    if not _losa_es_muro_viga_o_columna(other):
        return
    try:
        if not JoinGeometryUtils.AreElementsJoined(document, floor, other):
            return
    except Exception:
        return
    try:
        floor_corta_otro = bool(
            JoinGeometryUtils.IsCuttingElementInJoin(document, floor, other)
        )
    except Exception:
        return
    if not floor_corta_otro:
        return
    try:
        JoinGeometryUtils.SwitchJoinOrder(document, floor, other)
    except Exception as ex:
        if errores is not None:
            try:
                oid = other.Id
                errores.append(
                    u"Join: SwitchJoinOrder Id {0}: {1}".format(
                        _element_id_int(oid), str(ex))
                )
            except Exception:
                errores.append(u"Join: no se pudo invertir orden de unión.")


def _losa_unjoin_all(document, floor, other_ids, errores):
    """Desune la losa de cada elemento en ``other_ids``."""
    if not other_ids:
        return
    for oid in other_ids:
        try:
            oth = document.GetElement(oid)
        except Exception:
            oth = None
        if oth is None:
            continue
        try:
            JoinGeometryUtils.UnjoinGeometry(document, floor, oth)
        except Exception as ex:
            if errores is not None:
                try:
                    errores.append(
                        u"Join: desunir Id {0}: {1}".format(_element_id_int(oid), str(ex))
                    )
                except Exception:
                    errores.append(u"Join: desunir falló para un elemento unido.")


def _losa_rejoin_all(document, floor, other_ids, errores):
    """Restaura *Join Geometry* con los mismos elementos."""
    if not other_ids:
        return
    for oid in other_ids:
        try:
            oth = document.GetElement(oid)
        except Exception:
            oth = None
        if oth is None:
            continue
        already_joined = False
        try:
            already_joined = bool(
                JoinGeometryUtils.AreElementsJoined(document, floor, oth)
            )
        except Exception:
            pass
        if already_joined:
            _losa_switch_join_order_losa_recortada(document, floor, oth, errores)
            continue
        try:
            JoinGeometryUtils.JoinGeometry(document, floor, oth)
        except Exception as ex:
            if errores is not None:
                try:
                    errores.append(
                        u"Join: reunir Id {0}: {1}".format(_element_id_int(oid), str(ex))
                    )
                except Exception:
                    errores.append(u"Join: no se pudo restaurar unión geométrica.")
            continue
        _losa_switch_join_order_losa_recortada(document, floor, oth, errores)


# ── Gancho desde espesor de losa (antes de crear Area Reinforcement) ─────────
# Hook Length (mm) = espesor_losa − _HOOK_RESTA_MM (ej. 180 mm → gancho 120 mm).
_HOOK_RESTA_MM = 60
_HOOK_TOLERANCIA_MM = 0.5


def _obtener_espesor_losa_mm(floor):
    """
    Obtiene el espesor de una losa en mm.
    Prioriza FLOOR_ATTR_THICKNESS_PARAM (instancia) y LookupParameter('Default Thickness')
    en instancia y tipo como fallback (el espesor suele estar en el FloorType).
    """
    if floor is None:
        return None
    # 1) BuiltInParameter en la instancia (igual que verificar_o_crear_hook_desde_losa)
    try:
        param = floor.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)
        if param and param.HasValue:
            return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters)
    except Exception:
        pass
    # 2) LookupParameter "Default Thickness" en la instancia
    try:
        param = floor.LookupParameter("Default Thickness")
        if param and param.HasValue:
            return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters)
    except Exception:
        pass
    # 3) LookupParameter en el tipo (común en losas estructurales)
    try:
        type_id = floor.GetTypeId()
        if type_id and type_id != ElementId.InvalidElementId:
            floor_type = floor.Document.GetElement(type_id)
            if floor_type:
                for pname in ("Default Thickness", "Thickness", "Espesor"):
                    param = floor_type.LookupParameter(pname)
                    if param and param.HasValue:
                        return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters)
    except Exception:
        pass
    return None


def _obtener_rebar_bar_types(document):
    """Obtiene todos los RebarBarType del documento."""
    try:
        return list(FilteredElementCollector(document).OfClass(RebarBarType))
    except Exception:
        return []


def _obtener_hook_length_mm(bar_type, hook_type):
    """Obtiene el Hook Length en mm desde la tabla Hook Lengths del RebarBarType."""
    try:
        largo_interno = bar_type.GetHookLength(hook_type.Id)
        return UnitUtils.ConvertFromInternalUnits(largo_interno, UnitTypeId.Millimeters)
    except Exception:
        return None


def _buscar_hook_por_largo(document, largo_target_mm):
    """
    Busca un RebarHookType cuyo Hook Length sea igual a largo_target_mm
    en TODOS los RebarBarType. Solo retorna un gancho si todos los bar types
    tienen el largo correcto para ese hook.
    """
    bar_types = _obtener_rebar_bar_types(document)
    if not bar_types:
        return None
    for ht in FilteredElementCollector(document).OfClass(RebarHookType):
        if ht is None:
            continue
        todos_coinciden = True
        for bar_type in bar_types:
            try:
                largo_mm = _obtener_hook_length_mm(bar_type, ht)
                if largo_mm is None or abs(largo_mm - largo_target_mm) > _HOOK_TOLERANCIA_MM:
                    todos_coinciden = False
                    break
            except Exception:
                todos_coinciden = False
                break
        if todos_coinciden:
            return ht
    return None


def _crear_hook_desde_largo(document, largo_mm, en_transaccion=True):
    """
    Crea un RebarHookType con el Hook Length indicado.
    Debe ejecutarse dentro de una Transaction si en_transaccion=False.
    Retorna el RebarHookType creado.
    """
    bar_types = _obtener_rebar_bar_types(document)
    if not bar_types:
        raise Exception("No hay RebarBarType en el documento.")
    nombres_existentes = []
    for ht in FilteredElementCollector(document).OfClass(RebarHookType):
        try:
            if ht and ht.Name:
                nombres_existentes.append(ht.Name)
        except Exception:
            pass
    angulo_rad = math.radians(90.0)
    multiplicador = 12.0
    t = Transaction(document, "Crear Rebar Hook desde espesor losa") if en_transaccion else None
    if t:
        t.Start()
    try:
        hook_type = RebarHookType.Create(document, angulo_rad, multiplicador)
        largo_interno = UnitUtils.ConvertToInternalUnits(largo_mm, UnitTypeId.Millimeters)
        for bt in bar_types:
            bt.SetAutoCalcHookLengths(hook_type.Id, False)
            bt.SetHookLength(hook_type.Id, largo_interno)
        # Nombre opcional: si falla (ej. caracteres no admitidos), el gancho sigue siendo válido
        largo_str = "{} mm".format(int(round(largo_mm)))
        nombre_base = "Rebar Hook - 90 - {}".format(largo_str)
        nombre_final = nombre_base
        if nombre_base in nombres_existentes:
            contador = 1
            while nombre_final in nombres_existentes:
                nombre_final = "{} ({})".format(nombre_base, contador)
                contador += 1
        try:
            hook_type.Name = nombre_final
        except Exception:
            pass
        if t:
            t.Commit()
        return hook_type
    except Exception:
        if t:
            t.RollBack()
        raise


def _obtener_o_crear_hook_desde_espesor_losa(document, floor, en_transaccion=True):
    """
    Obtiene o crea un RebarHookType con Hook Length = espesor_losa - 60 mm.
    Debe ejecutarse ANTES de crear el Area Reinforcement.
    Retorna ElementId del gancho o InvalidElementId si no se pudo obtener.
    """
    espesor_mm = _obtener_espesor_losa_mm(floor)
    if espesor_mm is None:
        return ElementId.InvalidElementId
    largo_target = espesor_mm - _HOOK_RESTA_MM
    if largo_target <= 0:
        return ElementId.InvalidElementId
    hook = _buscar_hook_por_largo(document, largo_target)
    if hook:
        return hook.Id
    # Crear nuevo gancho (no tragar excepciones: el llamador debe manejarlas)
    nuevo = _crear_hook_desde_largo(document, largo_target, en_transaccion=en_transaccion)
    return nuevo.Id if nuevo else ElementId.InvalidElementId


def _crear_gancho_por_defecto(document):
    """
    Crea un RebarHookType por defecto (90°, extensión 12× diámetro) si no existe ninguno.
    Debe ejecutarse dentro de una Transaction ya iniciada.
    Retorna ElementId del gancho creado.
    """
    import math
    from Autodesk.Revit.DB import FilteredElementCollector, UnitUtils, UnitTypeId
    from Autodesk.Revit.DB.Structure import RebarBarType, RebarHookType
    angulo_rad = math.radians(90.0)
    multiplicador = 12.0
    hook_type = RebarHookType.Create(document, angulo_rad, multiplicador)
    largo_mm = 50.0
    largo_interno = UnitUtils.ConvertToInternalUnits(largo_mm, UnitTypeId.Millimeters)
    try:
        for bt in FilteredElementCollector(document).OfClass(RebarBarType):
            bt.SetAutoCalcHookLengths(hook_type.Id, False)
            bt.SetHookLength(hook_type.Id, largo_interno)
    except Exception:
        pass
    try:
        hook_type.Name = u"Rebar Hook - 90º - 50.0 mm (por defecto)"
    except Exception:
        pass
    return hook_type.Id


def _asignar_hook_a_area_reinforcement(area_rein, hook_type_id):
    """
    Asigna el RebarHookType al inicio y final de todas las capas del Area Reinforcement.
    El Create(7 params) acepta hookTypeId pero Revit NO lo aplica a las capas;
    hay que asignar explícitamente estos parámetros.
    Usa BuiltInParameter (invariante al idioma) + LookupParameter + fallback por iteración.
    """
    from Autodesk.Revit.DB import BuiltInParameter, ElementId, StorageType
    if not area_rein or not hook_type_id or hook_type_id == ElementId.InvalidElementId:
        return
    asignados = 0
    bip_names = [
        "REBAR_SYSTEM_HOOK_TYPE_MAJOR_TOP", "REBAR_SYSTEM_HOOK_TYPE_MAJOR_BOTTOM",
        "REBAR_SYSTEM_HOOK_TYPE_MINOR_TOP", "REBAR_SYSTEM_HOOK_TYPE_MINOR_BOTTOM",
        "REBAR_SYSTEM_HOOK_TYPE_EXTERIOR_MAJOR", "REBAR_SYSTEM_HOOK_TYPE_EXTERIOR_MINOR",
        "REBAR_SYSTEM_HOOK_TYPE_INTERIOR_MAJOR", "REBAR_SYSTEM_HOOK_TYPE_INTERIOR_MINOR",
        "REBAR_SYSTEM_HOOK_TYPE_TOP_DIR_1", "REBAR_SYSTEM_HOOK_TYPE_TOP_DIR_2",
        "REBAR_SYSTEM_HOOK_TYPE_BOTTOM_DIR_1", "REBAR_SYSTEM_HOOK_TYPE_BOTTOM_DIR_2",
    ]
    for name in bip_names:
        try:
            bip = getattr(BuiltInParameter, name, None)
            if bip is not None:
                p = area_rein.get_Parameter(bip)
                if p and not p.IsReadOnly and p.StorageType == StorageType.ElementId:
                    p.Set(hook_type_id)
                    asignados += 1
        except Exception:
            continue
    if asignados == 0:
        hook_param_names = [
            u"Exterior Major Hook Type", u"Top Major Hook Type",
            u"Exterior Minor Hook Type", u"Top Minor Hook Type",
            u"Interior Major Hook Type", u"Bottom Major Hook Type",
            u"Interior Minor Hook Type", u"Bottom Minor Hook Type",
        ]
        for pname in hook_param_names:
            try:
                p = area_rein.LookupParameter(pname)
                if p and not p.IsReadOnly and p.StorageType == StorageType.ElementId:
                    p.Set(hook_type_id)
                    asignados += 1
            except Exception:
                continue
    if asignados == 0:
        try:
            for p in area_rein.Parameters:
                if p is None or p.IsReadOnly or p.StorageType != StorageType.ElementId:
                    continue
                try:
                    nombre = p.Definition.Name if p.Definition else ""
                    if "ook" in nombre.lower() or "gancho" in nombre.lower():
                        p.Set(hook_type_id)
                        asignados += 1
                except Exception:
                    continue
        except Exception:
            pass


def _mm_to_internal(val_mm):
    """Convierte milímetros a unidades internas de Revit (pies)."""
    try:
        v = float(val_mm)
        try:
            from Autodesk.Revit.DB import UnitUtils, UnitTypeId
            return UnitUtils.ConvertToInternalUnits(v, UnitTypeId.Millimeters)
        except Exception:
            return v / 304.8
    except (TypeError, ValueError):
        return _DEFAULT_SPACING_MM / 304.8


def _is_area_reinforcement_type(elem):
    """Comprueba si elem es AreaReinforcementType."""
    if elem is None:
        return False
    try:
        if isinstance(elem, AreaReinforcementType):
            return True
    except Exception:
        pass
    try:
        tn = getattr(elem, "GetType", None)
        if tn and callable(tn):
            tinfo = tn()
            name = getattr(tinfo, "Name", "") or getattr(tinfo, "FullName", "") or str(tinfo)
            if "AreaReinforcementType" in str(name):
                return True
    except Exception:
        pass
    return False


def _get_default_area_reinforcement_type_id(document):
    """Obtiene un AreaReinforcementType válido del documento."""
    def _valid_id(eid):
        if not eid or eid == ElementId.InvalidElementId:
            return False
        n = _element_id_int(eid)
        return n is not None and n >= 0

    try:
        for elem in FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_AreaRein):
            if elem is None:
                continue
            try:
                if hasattr(elem, "GetTypeId"):
                    tid = elem.GetTypeId()
                    if _valid_id(tid):
                        t = document.GetElement(tid)
                        if t and _is_area_reinforcement_type(t):
                            return tid
            except Exception:
                continue
    except Exception:
        pass

    try:
        default_id = document.GetDefaultElementTypeId(ElementTypeGroup.AreaReinforcementType)
        if _valid_id(default_id):
            t = document.GetElement(default_id)
            if t and _is_area_reinforcement_type(t):
                return default_id
            if t is not None:
                return default_id
    except Exception:
        pass

    try:
        for elem in (FilteredElementCollector(document)
                     .OfCategory(BuiltInCategory.OST_AreaRein)
                     .WhereElementIsElementType()):
            if elem and _is_area_reinforcement_type(elem):
                return elem.Id
    except Exception:
        pass

    try:
        for elem in FilteredElementCollector(document).OfClass(AreaReinforcementType):
            if elem and _is_area_reinforcement_type(elem):
                return elem.Id
    except Exception:
        try:
            rt = clr.GetClrType(AreaReinforcementType)
            for elem in FilteredElementCollector(document).OfClass(rt):
                if elem and _is_area_reinforcement_type(elem):
                    return elem.Id
        except Exception:
            pass

    return None


def _get_rebar_bar_types(document):
    """Obtiene los RebarBarType del documento con display (øXX mm)."""
    result = []
    seen_ids = set()

    def _add(bar_type):
        if not bar_type:
            return
        try:
            eid = _element_id_int(bar_type.Id)
            if eid is None or eid in seen_ids:
                return
            seen_ids.add(eid)
            disp = u"Barra (ID{})".format(eid)
            try:
                diam_ft = bar_type.BarNominalDiameter
                diam_mm = int(round(float(diam_ft) * 304.8))
                disp = u"\u00f8{} mm".format(diam_mm) if diam_mm > 0 else disp
            except Exception:
                pass
            result.append((disp, bar_type))
        except Exception:
            pass

    try:
        ids = list(FilteredElementCollector(document).OfClass(RebarBarType).ToElementIds())
    except Exception:
        try:
            rt = clr.GetClrType(RebarBarType)
            ids = list(FilteredElementCollector(document).OfClass(rt).ToElementIds())
        except Exception:
            return []

    for eid in ids:
        try:
            t = document.GetElement(eid)
            if t:
                _add(t)
        except Exception:
            pass

    def _sort_key(x):
        try:
            return float(x[1].BarNominalDiameter)
        except Exception:
            return 0
    result.sort(key=_sort_key)
    return result


def _aplicar_parametros_malla(area_rein, params_dict, layer_active_dict):
    """Aplica diámetro, espaciado y activación de capas.
    Para losas: Exterior=Top (superior), Interior=Bottom (inferior)."""
    if not area_rein:
        return
    from Autodesk.Revit.DB import ElementId
    from Autodesk.Revit.DB.Structure import AreaReinforcementLayerType

    def _convert_mm(val):
        try:
            v = float(val)
            try:
                from Autodesk.Revit.DB import UnitUtils, UnitTypeId
                return UnitUtils.ConvertToInternalUnits(v, UnitTypeId.Millimeters)
            except Exception:
                return v / 304.8
        except (TypeError, ValueError):
            return 150.0 / 304.8

    layer_config = [
        ("exterior_major", [u"Exterior Major Spacing"], [u"Exterior Major Bar Type", u"Exterior Major Rebar Type"], AreaReinforcementLayerType.TopOrFrontMajor),
        ("exterior_minor", [u"Exterior Minor Spacing"], [u"Exterior Minor Bar Type", u"Exterior Minor Rebar Type"], AreaReinforcementLayerType.TopOrFrontMinor),
        ("interior_major", [u"Interior Major Spacing"], [u"Interior Major Bar Type", u"Interior Major Rebar Type"], AreaReinforcementLayerType.BottomOrBackMajor),
        ("interior_minor", [u"Interior Minor Spacing"], [u"Interior Minor Bar Type", u"Interior Minor Rebar Type"], AreaReinforcementLayerType.BottomOrBackMinor),
    ]

    for layer_key, spacing_names, bar_names, layer_type in layer_config:
        bar_type_id, spacing_mm = params_dict.get(layer_key, (None, "150"))
        is_active = layer_active_dict.get(layer_key, True)

        try:
            area_rein.SetLayerActive(layer_type, is_active)
        except Exception:
            pass
        # Nombres de parámetros Direction: Exterior/Interior y Top/Bottom (para losas)
        dir_param_names = {
            "exterior_major": [u"Exterior Major Direction", u"Top Major Direction", u"Top Mayor Direction"],
            "exterior_minor": [u"Exterior Minor Direction", u"Top Minor Direction"],
            "interior_major": [u"Interior Major Direction", u"Bottom Major Direction", u"Bottom Mayor Direction"],
            "interior_minor": [u"Interior Minor Direction", u"Bottom Minor Direction"],
        }
        for pname in dir_param_names.get(layer_key, []):
            try:
                p = area_rein.LookupParameter(pname)
                if p and not p.IsReadOnly:
                    p.Set(1 if is_active else 0)
                    break
            except Exception:
                continue

        spacing_internal = _convert_mm(spacing_mm)
        for name in spacing_names:
            try:
                p = area_rein.LookupParameter(name)
                if p and not p.IsReadOnly:
                    p.Set(spacing_internal)
            except Exception:
                pass
        for name in bar_names:
            try:
                p = area_rein.LookupParameter(name)
                if p and not p.IsReadOnly and bar_type_id and bar_type_id != ElementId.InvalidElementId:
                    p.Set(bar_type_id)
            except Exception:
                pass

    params_iter = None
    if hasattr(area_rein, "GetOrderedParameters"):
        params_iter = area_rein.GetOrderedParameters()
    elif hasattr(area_rein, "Parameters"):
        params_iter = area_rein.Parameters
    if params_iter:
        for param in params_iter:
            if param is None or param.IsReadOnly:
                continue
            try:
                def_name = (param.Definition.Name or "").lower()
                for layer_key, _, _, _ in layer_config:
                    bar_type_id, spacing_mm = params_dict.get(layer_key, (None, "150"))
                    is_active = layer_active_dict.get(layer_key, True)
                    ext_int = "exterior" if "exterior" in layer_key else "interior"
                    top_bot = "top" if "exterior" in layer_key else "bottom"
                    maj_min = "major" if "major" in layer_key else "minor"
                    # Coincidir exterior/top o interior/bottom
                    matches_layer = (ext_int in def_name or top_bot in def_name) and maj_min in def_name
                    if matches_layer:
                        try:
                            if "spacing" in def_name:
                                param.Set(_convert_mm(spacing_mm))
                            elif ("bar" in def_name or "rebar" in def_name) and "type" in def_name and bar_type_id and bar_type_id != ElementId.InvalidElementId:
                                param.Set(bar_type_id)
                            elif "direction" in def_name:
                                param.Set(1 if is_active else 0)
                        except Exception:
                            pass
                        break
            except Exception:
                continue


# Plantas donde tiene sentido etiquetar refuerzo de área (str(ViewType) en API Revit).
_VISTAS_PLANTA_PARA_ETIQUETA = frozenset(
    ("FloorPlan", "StructuralPlan", "CeilingPlan", "EngineeringPlan")
)


def _es_vista_planta_area_reinforcement(view):
    """True si la vista es tipo planta (arquitectura o estructural, etc.)."""
    if view is None:
        return False
    return str(view.ViewType) in _VISTAS_PLANTA_PARA_ETIQUETA


def _vista_valida_etiqueta_area_reinforcement(view):
    """
    Mismas restricciones que etiquetar_area_reinforcement_rps._view_ok_for_tag.
    str(ViewType) evita conflictos de IronPython con ViewType.
    """
    if view is None:
        return False, u"Vista nula."
    if view.IsTemplate:
        return False, u"La vista activa es una plantilla de vista."
    if str(view.ViewType) == "Perspective":
        return False, u"No se pueden crear etiquetas en vista en perspectiva."
    if isinstance(view, View3D) and not view.IsLocked:
        return False, u"En vista 3D la cámara debe estar bloqueada para etiquetar."
    return True, None


def _ar_tag_bbox_center_xyz(element, view):
    bb = element.get_BoundingBox(view)
    if bb is None:
        bb = element.get_BoundingBox(None)
    if bb is None or bb.Min is None or bb.Max is None:
        return None
    mn, mx = bb.Min, bb.Max
    return XYZ((mn.X + mx.X) * 0.5, (mn.Y + mx.Y) * 0.5, (mn.Z + mx.Z) * 0.5)


def _ar_tag_solid_centroid(solid):
    try:
        return solid.ComputeCentroid()
    except Exception:
        return None


def _ar_tag_volume_weighted_centroid(element):
    if element is None:
        return None
    opts = Options()
    opts.ComputeReferences = False
    try:
        geom_elem = element.get_Geometry(opts)
    except Exception:
        return None
    if geom_elem is None:
        return None
    vol_sum = 0.0
    sx = sy = sz = 0.0
    for obj in geom_elem:
        if obj is None:
            continue
        if isinstance(obj, Solid) and obj.Volume > 1e-12:
            c = _ar_tag_solid_centroid(obj)
            if c is None:
                continue
            v = obj.Volume
            sx += c.X * v
            sy += c.Y * v
            sz += c.Z * v
            vol_sum += v
        elif isinstance(obj, GeometryInstance):
            try:
                inst_geom = obj.GetInstanceGeometry()
                if inst_geom is None:
                    continue
                for g in inst_geom:
                    if isinstance(g, Solid) and g.Volume > 1e-12:
                        c = _ar_tag_solid_centroid(g)
                        if c is None:
                            continue
                        v = g.Volume
                        sx += c.X * v
                        sy += c.Y * v
                        sz += c.Z * v
                        vol_sum += v
            except Exception:
                pass
    if vol_sum < 1e-12:
        return None
    return XYZ(sx / vol_sum, sy / vol_sum, sz / vol_sum)


def _ar_tag_area_reinforcement_host_id(area_rein):
    try:
        hid = area_rein.GetHostId()
        if hid is not None and hid != ElementId.InvalidElementId:
            return hid
    except Exception:
        pass
    return None


def _ar_tag_punto_insercion(document, area_rein, view):
    hid = _ar_tag_area_reinforcement_host_id(area_rein)
    if hid is not None:
        host = document.GetElement(hid)
        if host is not None:
            c = _ar_tag_volume_weighted_centroid(host)
            if c is not None:
                return c
            c = _ar_tag_bbox_center_xyz(host, None)
            if c is not None:
                return c
    return _ar_tag_bbox_center_xyz(area_rein, view)


def _ar_tag_norm(s):
    if s is None:
        return u""
    t = str(s).replace(u"\xa0", u" ").replace(u"\u200b", u"").strip()
    if _AR_TAG_NORM_CASE_INSENSITIVE:
        return t.lower()
    return t


def _ar_tag_safe_element_name(elem):
    """IronPython: .Name en elementos Revit puede fallar de forma críptica; usar con precaución."""
    if elem is None:
        return u""
    try:
        n = elem.Name
        if n is None:
            return u""
        return str(n)
    except Exception:
        return u""


def _ar_tag_family_name_of_symbol(sym):
    """Nombre de familia desde el símbolo (más fiable que enlazar solo por elemento Family)."""
    if sym is None:
        return u""
    try:
        fn = sym.FamilyName
        if fn is not None and str(fn).strip():
            return str(fn)
    except Exception:
        pass
    try:
        fam = sym.Family
        if fam is not None:
            return _ar_tag_safe_element_name(fam)
    except Exception:
        pass
    return u""


def _ar_tag_type_name_candidates(sym):
    """
    Posibles textos de tipo para un FamilySymbol (etiquetas a veces no exponen solo .Name).
    Incluye trozos tras ':' (formato 'Familia : Tipo' en parámetros).
    """
    out = []
    seen_norm = set()

    def _add(raw):
        if raw is None:
            return
        t = str(raw).replace(u"\xa0", u" ").strip()
        if not t:
            return
        k = _ar_tag_norm(t)
        if k and k not in seen_norm:
            seen_norm.add(k)
            out.append(t)
        if u":" in t:
            tail = t.split(u":")[-1].strip()
            k2 = _ar_tag_norm(tail)
            if k2 and k2 not in seen_norm:
                seen_norm.add(k2)
                out.append(tail)

    try:
        _add(sym.Name)
    except Exception:
        pass
    for bip_attr in (
        u"SYMBOL_NAME_PARAM",
        u"ALL_MODEL_TYPE_NAME",
        u"ELEM_FAMILY_AND_TYPE_PARAM",
    ):
        try:
            bip = getattr(BuiltInParameter, bip_attr, None)
            if bip is None:
                continue
            p = sym.get_Parameter(bip)
            if p is None:
                continue
            st = p.StorageType
            if st == StorageType.String:
                _add(p.AsString())
            elif st == StorageType.Integer:
                _add(str(p.AsInteger()))
            else:
                try:
                    _add(p.AsValueString())
                except Exception:
                    pass
        except Exception:
            pass
    return out


def _ar_tag_symbol_type_matches(sym, type_name_wanted):
    wanted = _ar_tag_norm(type_name_wanted)
    if not wanted:
        return False
    for c in _ar_tag_type_name_candidates(sym):
        if _ar_tag_norm(c) == wanted:
            return True
    return False


def _ar_tag_find_symbol_by_family_and_type_strings(document, family_name_wanted, type_name_wanted):
    """
    Recorre FamilySymbol del documento comparando FamilyName (propiedad API) y tipo.
    """
    fam_w = _ar_tag_norm(family_name_wanted)
    if not fam_w:
        return None
    try:
        for sym in FilteredElementCollector(document).OfClass(
            _AR_REVIT_FAMILY_SYMBOL_CLR_TYPE
        ):
            try:
                if _ar_tag_norm(_ar_tag_family_name_of_symbol(sym)) != fam_w:
                    continue
            except Exception:
                continue
            if _ar_tag_symbol_type_matches(sym, type_name_wanted):
                return sym
    except Exception:
        pass
    return None


def _ar_tag_list_types_for_family_name_string(document, family_name_wanted):
    """Lista tipos (candidatos de nombre) de todos los símbolos cuyo FamilyName coincide."""
    fam_w = _ar_tag_norm(family_name_wanted)
    labels = []
    seen = set()
    try:
        for sym in FilteredElementCollector(document).OfClass(
            _AR_REVIT_FAMILY_SYMBOL_CLR_TYPE
        ):
            try:
                if _ar_tag_norm(_ar_tag_family_name_of_symbol(sym)) != fam_w:
                    continue
            except Exception:
                continue
            for c in _ar_tag_type_name_candidates(sym):
                k = _ar_tag_norm(c)
                if k and k not in seen:
                    seen.add(k)
                    labels.append(c)
    except Exception:
        pass
    return sorted(labels, key=lambda x: x.lower())


def _ar_tag_find_families(document, family_name_wanted):
    wanted = _ar_tag_norm(family_name_wanted)
    matches = []
    for fam in FilteredElementCollector(document).OfClass(_AR_REVIT_FAMILY_CLR_TYPE):
        if _ar_tag_norm(_ar_tag_safe_element_name(fam)) == wanted:
            matches.append(fam)
    return matches


def _ar_tag_same_family_id(fam_id_a, fam_id_b):
    try:
        ia = _element_id_int(fam_id_a)
        ib = _element_id_int(fam_id_b)
        if ia is not None and ib is not None:
            return ia == ib
    except Exception:
        pass
    try:
        return fam_id_a == fam_id_b
    except Exception:
        return False


def _ar_tag_find_symbol_via_collector(document, family_id, type_name_wanted):
    """
    Respaldo: en IronPython isinstance(FamilySymbol) a veces falla con símbolos de etiqueta;
    GetFamilySymbolIds() puede no devolver todo en algunos casos.
    """
    try:
        for sym in FilteredElementCollector(document).OfClass(
            _AR_REVIT_FAMILY_SYMBOL_CLR_TYPE
        ):
            try:
                fam = sym.Family
                if fam is None or not _ar_tag_same_family_id(fam.Id, family_id):
                    continue
            except Exception:
                continue
            if _ar_tag_symbol_type_matches(sym, type_name_wanted):
                return sym
    except Exception:
        pass
    return None


def _ar_tag_list_symbol_names_in_family(document, family):
    nombres = []
    try:
        for sid in family.GetFamilySymbolIds():
            el = document.GetElement(sid)
            if el is None:
                continue
            for c in _ar_tag_type_name_candidates(el):
                if c:
                    nombres.append(c)
    except Exception:
        pass
    if not nombres:
        try:
            for sym in FilteredElementCollector(document).OfClass(
                _AR_REVIT_FAMILY_SYMBOL_CLR_TYPE
            ):
                try:
                    if sym.Family is None or not _ar_tag_same_family_id(
                        sym.Family.Id, family.Id
                    ):
                        continue
                except Exception:
                    continue
                for c in _ar_tag_type_name_candidates(sym):
                    if c:
                        nombres.append(c)
        except Exception:
            pass
    return sorted(set(nombres), key=lambda x: x.lower())


def _ar_tag_find_symbol_in_family(document, family, type_name_wanted):
    for sid in family.GetFamilySymbolIds():
        sym = document.GetElement(sid)
        if sym is None:
            continue
        if _ar_tag_symbol_type_matches(sym, type_name_wanted):
            return sym
    return _ar_tag_find_symbol_via_collector(document, family.Id, type_name_wanted)


def _ar_tag_activate_symbol_inplace(sym):
    """Activa el símbolo dentro de la transacción ya abierta (p. ej. SubTransaction de etiquetas)."""
    try:
        if sym.IsActive:
            return True
        sym.Activate()
        return True
    except Exception:
        return False


def _ar_tag_resolve_symbol_id_por_familia_y_tipo(document, family_name, type_name, errores):
    """
    Resuelve FamilySymbol por nombre de familia + tipo.
    Prioridad: escaneo por FamilySymbol.FamilyName (API) y candidatos de nombre de tipo.
    """
    if not family_name or not type_name:
        return None
    sym = _ar_tag_find_symbol_by_family_and_type_strings(
        document, family_name, type_name)
    families = []
    if sym is None:
        families = _ar_tag_find_families(document, family_name)
        if families:
            sym = _ar_tag_find_symbol_in_family(document, families[0], type_name)
    if sym is None:
        if not families:
            families = _ar_tag_find_families(document, family_name)
        candidatos = _ar_tag_list_types_for_family_name_string(document, family_name)
        if not candidatos and families:
            candidatos = _ar_tag_list_symbol_names_in_family(document, families[0])
        extra = u""
        if candidatos:
            muestra = candidatos[:25]
            extra = u" Tipos detectados: {}.".format(u"; ".join(muestra))
            if len(candidatos) > 25:
                extra += u" …"
        fam_label = family_name
        if families:
            fam_label = _ar_tag_safe_element_name(families[0]) or family_name
        if not families and not candidatos:
            errores.append(
                u"Etiqueta: no hay Family ni símbolos con FamilyName {!r}.".format(
                    family_name))
            return None
        errores.append(
            u"Etiqueta: no se resolvió el tipo {!r} en familia {!r}.{}".format(
                type_name, fam_label, extra))
        return None
    if not _ar_tag_activate_symbol_inplace(sym):
        errores.append(
            u"Etiqueta: no se pudo activar el tipo {!r}.".format(type_name))
        return None
    return sym.Id


# Nombres de parámetro Direction en Area Reinforcement (inglés + variantes del UI).
_AR_TAG_P_TOP_MAJOR = (
    u"Top Major Direction",
    u"Exterior Major Direction",
    u"Top Mayor Direction",
)
_AR_TAG_P_TOP_MINOR = (
    u"Top Minor Direction",
    u"Exterior Minor Direction",
)
_AR_TAG_P_BOTTOM_MAJOR = (
    u"Bottom Major Direction",
    u"Interior Major Direction",
    u"Bottom Mayor Direction",
)
_AR_TAG_P_BOTTOM_MINOR = (
    u"Bottom Minor Direction",
    u"Interior Minor Direction",
)


def _ar_tag_direction_layer_on(area_rein, parameter_name_candidates):
    """True si el parámetro de dirección existe y está activado (capa armada en esa dirección)."""
    for pname in parameter_name_candidates:
        try:
            p = area_rein.LookupParameter(pname)
            if p is None:
                continue
            st = p.StorageType
            if st == StorageType.Integer:
                return p.AsInteger() != 0
            if st == StorageType.Double:
                return abs(p.AsDouble()) > 1e-12
            if st == StorageType.String:
                s = (p.AsString() or u"").strip().lower()
                if s in (u"yes", u"sí", u"si", u"1", u"true", u"verdadero"):
                    return True
                if s in (u"no", u"0", u"false", u"falso", u""):
                    return False
                return len(s) > 0
        except Exception:
            continue
    return False


def _ar_tag_tipo_por_direcciones_en_area_rein(area_rein):
    """
    Según parámetros del elemento ya creado:
    - Los cuatro (Top/Bottom Major/Minor Direction) activos → Doble Malla.
    - Top Major y Top Minor desactivados → Malla Inferior.
    - Otros casos → None (etiqueta por categoría).
    """
    t_maj = _ar_tag_direction_layer_on(area_rein, _AR_TAG_P_TOP_MAJOR)
    t_min = _ar_tag_direction_layer_on(area_rein, _AR_TAG_P_TOP_MINOR)
    b_maj = _ar_tag_direction_layer_on(area_rein, _AR_TAG_P_BOTTOM_MAJOR)
    b_min = _ar_tag_direction_layer_on(area_rein, _AR_TAG_P_BOTTOM_MINOR)
    if t_maj and t_min and b_maj and b_min:
        return _AR_TAG_TYPE_DOBLE_MALLA
    if (not t_maj) and (not t_min):
        return _AR_TAG_TYPE_MALLA_INFERIOR
    return None


def _crear_etiqueta_area_reinforcement(document, view, area_rein, tag_symbol_id=None):
    """
    IndependentTag: tag_symbol_id explícito, o _AR_TAG_SYMBOL_ID, o TM_ADDBY_CATEGORY.
    """
    ref = Reference(area_rein)
    pnt = _ar_tag_punto_insercion(document, area_rein, view)
    if pnt is None:
        raise Exception(
            u"No se pudo calcular un punto para la etiqueta (host sin geometría ni bbox)."
        )
    effective_id = tag_symbol_id
    if effective_id is None or effective_id == ElementId.InvalidElementId:
        if _AR_TAG_SYMBOL_ID is not None and _AR_TAG_SYMBOL_ID != ElementId.InvalidElementId:
            effective_id = _AR_TAG_SYMBOL_ID
    if effective_id is not None and effective_id != ElementId.InvalidElementId:
        return IndependentTag.Create(
            document,
            effective_id,
            view.Id,
            ref,
            _AR_TAG_ADD_LEADER,
            _AR_TAG_ORIENTATION,
            pnt,
        )
    return IndependentTag.Create(
        document,
        view.Id,
        ref,
        _AR_TAG_ADD_LEADER,
        TagMode.TM_ADDBY_CATEGORY,
        _AR_TAG_ORIENTATION,
        pnt,
    )


def _ocultar_rebar_in_system_de_area_reinforcement_en_vista(document, view, area_reinforcements):
    """
    Oculta en ``view`` los RebarInSystem generados por cada AreaReinforcement
    (Hide in View → Elements), de modo que solo quede visible el Area Reinforcement
    para etiquetado. Requiere ``document.Regenerate()`` previo para que existan los
    RebarInSystem (si ``ReinforcementSettings.HostStructuralRebar`` es false,
    ``GetRebarInSystemIds`` puede devolver vacío).
    """
    if view is None or document is None or not area_reinforcements:
        return
    try:
        if getattr(view, "IsTemplate", False):
            return
    except Exception:
        pass
    try:
        from System.Collections.Generic import List as ClrList
        from Autodesk.Revit.DB import ElementCategoryFilter

        ids_hide = ClrList[ElementId]()
        for ar in area_reinforcements:
            if ar is None:
                continue
            n_antes = int(ids_hide.Count)
            try:
                sys_ids = ar.GetRebarInSystemIds()
            except Exception:
                sys_ids = None
            if sys_ids is not None:
                try:
                    n = int(sys_ids.Count)
                except Exception:
                    n = 0
                for i in range(n):
                    try:
                        eid = sys_ids[i]
                        if eid is not None and eid != ElementId.InvalidElementId:
                            ids_hide.Add(eid)
                    except Exception:
                        continue
            if int(ids_hide.Count) == n_antes:
                try:
                    flt = ElementCategoryFilter(BuiltInCategory.OST_RebarInSystem)
                    dep = ar.GetDependentElements(flt)
                    if dep is not None:
                        try:
                            nd = int(dep.Count)
                        except Exception:
                            nd = 0
                        for j in range(nd):
                            try:
                                eid = dep[j]
                                if eid is not None and eid != ElementId.InvalidElementId:
                                    ids_hide.Add(eid)
                            except Exception:
                                continue
                except Exception:
                    pass
        if ids_hide.Count < 1:
            return
        try:
            view.HideElements(ids_hide)
        except Exception:
            pass
    except Exception:
        pass


# ── ExternalEvent Handlers ───────────────────────────────────────────────────
class ColocarAreaReinforcementHandler(IExternalEventHandler):
    """Ejecuta la creación de Area Reinforcement en losas."""

    def __init__(self, window_ref, get_area_type_fn, aplicar_parametros_fn, get_curvas_fn,
                 crear_gancho_fn, asignar_hook_fn):
        self._window_ref = window_ref
        self._get_area_type = get_area_type_fn
        self._aplicar_parametros = aplicar_parametros_fn
        self._get_curvas = get_curvas_fn
        self._crear_gancho = crear_gancho_fn
        self._asignar_hook = asignar_hook_fn
        self.floor_ids = []
        self.params_dict = {}
        self.layer_active_dict = {}
        self.area_reinforcement_type_id = None
        self.asignar_ganchos = True

    def Execute(self, uiapp):
        from System.Collections.Generic import List
        from Autodesk.Revit.DB import Curve, ElementId, Floor, Transaction, XYZ
        from Autodesk.Revit.DB.Structure import AreaReinforcement
        _wpf = _wpf_from_window_ref(self._window_ref)
        try:
            doc = uiapp.ActiveUIDocument.Document
            uidoc = uiapp.ActiveUIDocument
            # Vista al inicio del handler (ExternalEvent); evita depender del estado tras el bucle.
            vista_etiqueta = uidoc.ActiveView
            if not self.floor_ids:
                _task_dialog_show("Malla en Losa - Error", u"No hay losas seleccionadas.", _wpf)
                return
            if not self.params_dict or not any(
                pid and pid != ElementId.InvalidElementId
                for pid, _ in self.params_dict.values()
            ):
                _task_dialog_show("Malla en Losa - Error", u"Selecciona al menos un tipo de barra válido.", _wpf)
                return
            area_type_id = self.area_reinforcement_type_id or (self._get_area_type(doc) if self._get_area_type else None)
            if not area_type_id or area_type_id == ElementId.InvalidElementId:
                _task_dialog_show(
                    "Malla en Losa - Error",
                    u"No hay AreaReinforcementType en el proyecto. Crea uno manualmente.",
                    _wpf,
                )
                return
            first_bar_id = None
            for pid, _ in self.params_dict.values():
                if pid and pid != ElementId.InvalidElementId:
                    first_bar_id = pid
                    break
            if not first_bar_id and self.params_dict:
                first_bar_id = list(self.params_dict.values())[0][0]
            creados = 0
            errores = []
            area_rein_creados = []
            etiquetas_ok = 0
            trans = Transaction(doc, "Area Reinforcement en losas")
            try:
                trans.Start()
                for floor_id in self.floor_ids:
                    floor = doc.GetElement(floor_id)
                    if not floor or not isinstance(floor, Floor):
                        errores.append(u"ID {}: no es losa válida".format(_element_id_int(floor_id)))
                        continue
                    joined_ids = []
                    try:
                        joined_ids = _losa_collect_joined_element_ids(doc, floor)
                        if joined_ids:
                            _losa_unjoin_all(doc, floor, joined_ids, errores)
                            try:
                                doc.Regenerate()
                            except Exception:
                                pass
                            floor = doc.GetElement(floor_id)
                            if floor is None or not isinstance(floor, Floor):
                                errores.append(
                                    u"ID {}: losa inválida tras desunir geometría".format(
                                        _element_id_int(floor_id)))
                                continue
                        # Obtener o crear gancho ANTES de crear el Area Reinforcement (espesor losa - 60 mm).
                        rebar_hook_type_id = ElementId.InvalidElementId
                        if self.asignar_ganchos:
                            sub = SubTransaction(doc)
                            try:
                                sub.Start()
                                rebar_hook_type_id = _obtener_o_crear_hook_desde_espesor_losa(doc, floor, en_transaccion=False)
                                # No usar el «primer gancho del proyecto»: suele ser uno viejo (p. ej. 140 mm) y ignora la regla espesor−60.
                                if not rebar_hook_type_id or rebar_hook_type_id == ElementId.InvalidElementId:
                                    rebar_hook_type_id = self._crear_gancho(doc) if self._crear_gancho else ElementId.InvalidElementId
                                sub.Commit()
                            except Exception as ex_sub:
                                if sub.HasStarted():
                                    try:
                                        sub.RollBack()
                                    except Exception:
                                        pass
                                rebar_hook_type_id = ElementId.InvalidElementId
                                errores.append(u"ID {}: gancho - {}".format(_element_id_int(floor_id), str(ex_sub)))
                                continue
                        curves = self._get_curvas(floor, doc)
                        if not curves or len(curves) == 0:
                            errores.append(u"ID {}: la losa no tiene sketch válido o no se pudieron obtener las curvas".format(_element_id_int(floor_id)))
                            continue
                        layout_dir = _obtener_layout_dir_area_reinforcement(floor, curves)
                        curve_list = List[Curve](curves)
                        try:
                            area_rein = AreaReinforcement.Create(
                                doc, floor, curve_list, layout_dir,
                                area_type_id, first_bar_id, rebar_hook_type_id
                            )
                            if self.asignar_ganchos and self._asignar_hook and rebar_hook_type_id and rebar_hook_type_id != ElementId.InvalidElementId:
                                self._asignar_hook(area_rein, rebar_hook_type_id)
                            if area_rein and self._aplicar_parametros:
                                self._aplicar_parametros(
                                    area_rein,
                                    self.params_dict,
                                    self.layer_active_dict
                                )
                            creados += 1
                            area_rein_creados.append(area_rein)
                        except Exception as ex:
                            errores.append(u"ID {}: {}".format(_element_id_int(floor_id), str(ex)))
                    finally:
                        if joined_ids:
                            try:
                                fl = doc.GetElement(floor_id)
                                if fl is not None and isinstance(fl, Floor):
                                    _losa_rejoin_all(doc, fl, joined_ids, errores)
                            except Exception:
                                pass
                            try:
                                doc.Regenerate()
                            except Exception:
                                pass
                if area_rein_creados:
                    try:
                        doc.Regenerate()
                    except Exception:
                        pass
                    _ocultar_rebar_in_system_de_area_reinforcement_en_vista(
                        doc, vista_etiqueta, area_rein_creados
                    )
                # Etiquetado en la misma transacción que la creación (un solo Commit).
                if area_rein_creados and _es_vista_planta_area_reinforcement(vista_etiqueta):
                    ok_v, _ = _vista_valida_etiqueta_area_reinforcement(vista_etiqueta)
                    if ok_v:
                        try:
                            sym_override = None
                            if (
                                _AR_TAG_SYMBOL_ID is not None
                                and _AR_TAG_SYMBOL_ID != ElementId.InvalidElementId
                            ):
                                sym_override = _AR_TAG_SYMBOL_ID
                            sym_por_tipo = {}
                            for ar in area_rein_creados:
                                if ar is None:
                                    continue
                                try:
                                    if sym_override is not None:
                                        sid_use = sym_override
                                    else:
                                        tipo_et = _ar_tag_tipo_por_direcciones_en_area_rein(
                                            ar)
                                        sid_use = None
                                        if tipo_et:
                                            if tipo_et not in sym_por_tipo:
                                                sym_por_tipo[tipo_et] = (
                                                    _ar_tag_resolve_symbol_id_por_familia_y_tipo(
                                                        doc,
                                                        _AR_TAG_FAMILY_NAME,
                                                        tipo_et,
                                                        errores,
                                                    ))
                                            sid_use = sym_por_tipo[tipo_et]
                                    tag = _crear_etiqueta_area_reinforcement(
                                        doc,
                                        vista_etiqueta,
                                        ar,
                                        sid_use,
                                    )
                                    if tag is not None:
                                        etiquetas_ok += 1
                                except Exception as ex_et:
                                    errores.append(
                                        u"Etiqueta (ref. {}): {}".format(
                                            _element_id_int(ar.Id), str(ex_et)))
                        except Exception as ex_et_all:
                            errores.append(
                                u"Etiquetas [{}]: {}".format(
                                    type(ex_et_all).__name__, str(ex_et_all)))
                trans.Commit()
                # Sin mensaje en formulario ni TaskDialog de éxito (la vista refleja el resultado).
                # Cerrar la ventana al finalizar si así se configuró (p.ej. botón RPS).
                try:
                    win = self._window_ref() if self._window_ref else None
                    if win and getattr(win, "_close_on_finish", False) and hasattr(win, "_win") and win._win:
                        try:
                            disp = getattr(win._win, "Dispatcher", None)
                            if disp:
                                disp.Invoke(lambda: win._close_with_fade())
                            else:
                                win._close_with_fade()
                        except Exception:
                            try:
                                win._close_with_fade()
                            except Exception:
                                try:
                                    win._win.Close()
                                except Exception:
                                    pass
                except Exception:
                    pass
            except Exception as ex:
                if trans.HasStarted():
                    try:
                        trans.RollBack()
                    except Exception:
                        pass
                _task_dialog_show("Malla en Losa - Error", u"Error:\n\n{}".format(str(ex)), _wpf)
        except Exception as ex:
            _task_dialog_show("Malla en Losa - Error", u"Error:\n\n{}".format(str(ex)), _wpf)
        finally:
            try:
                win = self._window_ref() if self._window_ref else None
                if win and hasattr(win, "_win"):
                    win._win.Activate()
            except Exception:
                pass

    def GetName(self):
        return "ColocarAreaReinforcementLosa"


class SeleccionarLosaHandler(IExternalEventHandler):
    """Ejecuta la selección de losas en contexto API de Revit."""

    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        from Autodesk.Revit.UI.Selection import ObjectType
        from Autodesk.Revit.DB import BuiltInCategory, Floor
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        win = self._window_ref()
        if not win:
            return
        try:
            win._document = doc
            refs = list(uidoc.Selection.PickObjects(
                ObjectType.Element,
                u"Selecciona una o más losas. Finaliza con Finish o Cancel."
            ))
            if not refs:
                pass
            else:
                floor_ids = []
                for ref in refs:
                    elem = doc.GetElement(ref.ElementId)
                    if elem and elem.Category and _element_id_int(elem.Category.Id) == int(BuiltInCategory.OST_Floors):
                        if isinstance(elem, Floor):
                            floor_ids.append(ref.ElementId)
                if floor_ids:
                    win._floor_ids = floor_ids
                    win._actualizar_info_losas(doc, floor_ids)
                else:
                    pass
        except Exception as ex:
            err = str(ex).lower()
            if "cancel" not in err and "operation" not in err:
                _task_dialog_show("Malla en Losa - Error", str(ex), win._win)
        finally:
            try:
                win._win.Show()
                win._win.Activate()
            except Exception:
                pass

    def GetName(self):
        return "SeleccionarLosa"


# ── XAML — Misma línea de diseño que Refuerzo Borde Losa ─────────────────────
XAML = u"""
<Window
    x:Name="ArLosaWin"
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Arainco - Malla en Losa"
    SizeToContent="Height"
    WindowStartupLocation="Manual"
    Background="Transparent"
    AllowsTransparency="True"
    FontFamily="Segoe UI"
    WindowStyle="None"
    ResizeMode="NoResize"
    Topmost="True"
    UseLayoutRounding="True">

  <!--
    Apertura: ScaleTransform en ArLosaRootScale (origen 0,0). Animar Window.Height en host Revit suele dejar ~1 px de alto.
  -->
  <Window.Resources>
    <Storyboard x:Key="ArLosaOpenGrowStoryboard">
      <DoubleAnimation Storyboard.TargetName="ArLosaRootScale" Storyboard.TargetProperty="ScaleX"
                       From="0" To="1" Duration="0:0:0.18" FillBehavior="HoldEnd">
        <DoubleAnimation.EasingFunction>
          <QuadraticEase EasingMode="EaseOut"/>
        </DoubleAnimation.EasingFunction>
      </DoubleAnimation>
      <DoubleAnimation Storyboard.TargetName="ArLosaRootScale" Storyboard.TargetProperty="ScaleY"
                       From="0" To="1" Duration="0:0:0.18" FillBehavior="HoldEnd">
        <DoubleAnimation.EasingFunction>
          <QuadraticEase EasingMode="EaseOut"/>
        </DoubleAnimation.EasingFunction>
      </DoubleAnimation>
      <DoubleAnimation Storyboard.TargetName="ArLosaWin" Storyboard.TargetProperty="Opacity"
                       From="0" To="1" Duration="0:0:0.18" FillBehavior="HoldEnd">
        <DoubleAnimation.EasingFunction>
          <QuadraticEase EasingMode="EaseOut"/>
        </DoubleAnimation.EasingFunction>
      </DoubleAnimation>
    </Storyboard>
""" + BIMTOOLS_DARK_STYLES_XML + u"""
  </Window.Resources>

  <Border x:Name="ArLosaRootChrome" CornerRadius="10" Background="#0A1A2F" Padding="12"
          BorderBrush="#1A3A4D" BorderThickness="1"
          HorizontalAlignment="Stretch" ClipToBounds="True" RenderTransformOrigin="0,0">
    <Border.Effect>
      <DropShadowEffect Color="#000000" BlurRadius="16" ShadowDepth="0" Opacity="0.35"/>
    </Border.Effect>
    <Border.RenderTransform>
      <ScaleTransform x:Name="ArLosaRootScale" ScaleX="0" ScaleY="0"/>
    </Border.RenderTransform>
  <Grid HorizontalAlignment="Stretch">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
    </Grid.RowDefinitions>

    <Border x:Name="TitleBar" Grid.Row="0" Background="#0E1B32" CornerRadius="6" Padding="10,8" Margin="0,0,0,10"
            BorderBrush="#21465C" BorderThickness="1" HorizontalAlignment="Stretch">
      <Grid HorizontalAlignment="Stretch">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <Image x:Name="ImgLogo" Width="40" Height="40" Grid.Column="0"
               Stretch="Uniform" Margin="0,0,10,0" VerticalAlignment="Center"/>
        <StackPanel Grid.Column="1" VerticalAlignment="Center" Margin="0,0,8,0">
          <TextBlock Text="Malla en Losa"
                     FontSize="15" FontWeight="SemiBold" Foreground="#E8F4F8"
                     TextWrapping="Wrap"/>
        </StackPanel>
        <Button x:Name="BtnClose" Grid.Column="2"
                Style="{StaticResource BtnCloseX_MinimalNoBg}"
                VerticalAlignment="Center" HorizontalAlignment="Right" ToolTip="Cerrar"/>
      </Grid>
    </Border>

    <StackPanel Grid.Row="1" Margin="0,0,0,6" HorizontalAlignment="Stretch">
      <Button x:Name="BtnSeleccionar" Content="Seleccionar losas en modelo"
              Style="{StaticResource BtnSelectOutline}"
              HorizontalAlignment="Stretch"/>
    </StackPanel>

    <GroupBox Grid.Row="2" Style="{StaticResource GbParams}" Margin="0" HorizontalAlignment="Left">
      <GroupBox.Header>
        <CheckBox x:Name="ChkMallaSuperior" IsChecked="True" Content="Malla superior"
                  Foreground="#E8F4F8" FontWeight="SemiBold" FontSize="11" VerticalAlignment="Center"/>
      </GroupBox.Header>
      <StackPanel x:Name="PanelSuperiorContent" IsEnabled="True">
        <Border Background="#0E1B32" CornerRadius="4" Padding="8,6" BorderBrush="#1A3A4D" BorderThickness="1" Margin="0,0,0,6" HorizontalAlignment="Left">
          <StackPanel>
            <TextBlock Text="Luz Mayor" Style="{StaticResource LabelSmall}" Margin="0,0,0,4"/>
            <Grid HorizontalAlignment="Left">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="110"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="110"/>
              </Grid.ColumnDefinitions>
              <ComboBox Grid.Column="0" x:Name="CmbExteriorMajorDiametro" Style="{StaticResource Combo}" IsEditable="False">
                <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
              </ComboBox>
              <TextBlock Grid.Column="1" Text="@" FontSize="12" FontWeight="Bold"
                         Foreground="#95B8CC" VerticalAlignment="Center" HorizontalAlignment="Center" Margin="6,0,6,0"/>
              <ComboBox Grid.Column="2" x:Name="CmbExteriorMajorEspaciamiento" Style="{StaticResource Combo}" IsEditable="True">
                <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
              </ComboBox>
            </Grid>
          </StackPanel>
        </Border>
        <Border Background="#0E1B32" CornerRadius="4" Padding="8,6" BorderBrush="#1A3A4D" BorderThickness="1" HorizontalAlignment="Left">
          <StackPanel>
            <TextBlock Text="Luz Menor" Style="{StaticResource LabelSmall}" Margin="0,0,0,4"/>
            <Grid HorizontalAlignment="Left">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="110"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="110"/>
              </Grid.ColumnDefinitions>
              <ComboBox Grid.Column="0" x:Name="CmbExteriorMinorDiametro" Style="{StaticResource Combo}" IsEditable="False">
                <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
              </ComboBox>
              <TextBlock Grid.Column="1" Text="@" FontSize="12" FontWeight="Bold"
                         Foreground="#95B8CC" VerticalAlignment="Center" HorizontalAlignment="Center" Margin="6,0,6,0"/>
              <ComboBox Grid.Column="2" x:Name="CmbExteriorMinorEspaciamiento" Style="{StaticResource Combo}" IsEditable="True">
                <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
              </ComboBox>
            </Grid>
          </StackPanel>
        </Border>
      </StackPanel>
    </GroupBox>

    <GroupBox Grid.Row="3" Style="{StaticResource GbParams}" Margin="0" HorizontalAlignment="Left">
      <GroupBox.Header>
        <CheckBox x:Name="ChkMallaInferior" IsChecked="True" Content="Malla inferior"
                  Foreground="#E8F4F8" FontWeight="SemiBold" FontSize="11" VerticalAlignment="Center"/>
      </GroupBox.Header>
      <StackPanel x:Name="PanelInferiorContent" IsEnabled="True">
        <Border Background="#0E1B32" CornerRadius="4" Padding="8,6" BorderBrush="#1A3A4D" BorderThickness="1" Margin="0,0,0,6" HorizontalAlignment="Left">
          <StackPanel>
            <TextBlock Text="Luz Mayor" Style="{StaticResource LabelSmall}" Margin="0,0,0,4"/>
            <Grid HorizontalAlignment="Left">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="110"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="110"/>
              </Grid.ColumnDefinitions>
              <ComboBox Grid.Column="0" x:Name="CmbInteriorMajorDiametro" Style="{StaticResource Combo}" IsEditable="False">
                <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
              </ComboBox>
              <TextBlock Grid.Column="1" Text="@" FontSize="12" FontWeight="Bold"
                         Foreground="#95B8CC" VerticalAlignment="Center" HorizontalAlignment="Center" Margin="6,0,6,0"/>
              <ComboBox Grid.Column="2" x:Name="CmbInteriorMajorEspaciamiento" Style="{StaticResource Combo}" IsEditable="True">
                <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
              </ComboBox>
            </Grid>
          </StackPanel>
        </Border>
        <Border Background="#0E1B32" CornerRadius="4" Padding="8,6" BorderBrush="#1A3A4D" BorderThickness="1" HorizontalAlignment="Left">
          <StackPanel>
            <TextBlock Text="Luz Menor" Style="{StaticResource LabelSmall}" Margin="0,0,0,4"/>
            <Grid HorizontalAlignment="Left">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="110"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="110"/>
              </Grid.ColumnDefinitions>
              <ComboBox Grid.Column="0" x:Name="CmbInteriorMinorDiametro" Style="{StaticResource Combo}" IsEditable="False">
                <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
              </ComboBox>
              <TextBlock Grid.Column="1" Text="@" FontSize="12" FontWeight="Bold"
                         Foreground="#95B8CC" VerticalAlignment="Center" HorizontalAlignment="Center" Margin="6,0,6,0"/>
              <ComboBox Grid.Column="2" x:Name="CmbInteriorMinorEspaciamiento" Style="{StaticResource Combo}" IsEditable="True">
                <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
              </ComboBox>
            </Grid>
          </StackPanel>
        </Border>
      </StackPanel>
    </GroupBox>

    <StackPanel Grid.Row="4" Margin="0,14,0,0" HorizontalAlignment="Stretch">
      <Button x:Name="BtnColocar" Content="Colocar armaduras"
              Style="{StaticResource BtnPrimary}"
              HorizontalAlignment="Stretch"/>
    </StackPanel>

    <Border Grid.Row="5" Background="Transparent"/>
  </Grid>
  </Border>
</Window>
"""


# ── Ventana principal ───────────────────────────────────────────────────────
class AreaReinforcementLosaWindow(object):
    def __init__(self, revit, close_on_finish=False):
        self._document = None
        self._floor_ids = []
        self._rebar_type_ids = {}
        self._area_reinforcement_type_id = None
        self._revit = revit
        self._close_on_finish = bool(close_on_finish)
        self._win = XamlReader.Parse(XAML)
        self._form_width_px = float(_area_reinforcement_losa_form_width_px())
        self._win.Width = self._form_width_px
        self._win.MinWidth = self._form_width_px
        self._win.MaxWidth = self._form_width_px
        self._open_grow_storyboard_started = False
        self._colocar_handler = ColocarAreaReinforcementHandler(
            weakref.ref(self),
            _get_default_area_reinforcement_type_id,
            _aplicar_parametros_malla,
            _obtener_curvas_sketch_con_recubrimiento,
            _crear_gancho_por_defecto,
            _asignar_hook_a_area_reinforcement,
        )
        self._colocar_event = ExternalEvent.Create(self._colocar_handler)
        self._seleccion_handler = SeleccionarLosaHandler(weakref.ref(self))
        self._seleccion_event = ExternalEvent.Create(self._seleccion_handler)
        self._is_closing_with_fade = False
        self._base_top = None
        self._setup_ui()
        self._wire_commands()
        self._wire_lifecycle_handlers()
        self._wire_open_grow_storyboard_completed()

    def _wire_open_grow_storyboard_completed(self):
        try:
            sb = self._win.TryFindResource("ArLosaOpenGrowStoryboard")
            if sb is not None:
                sb.Completed += EventHandler(self._on_open_grow_storyboard_completed)
        except Exception:
            pass

    def _on_open_grow_storyboard_completed(self, sender, args):
        try:
            self._win.MinWidth = self._form_width_px
            self._win.MaxWidth = self._form_width_px
        except Exception:
            pass

    def _position_win_top_left_active_view(self):
        try:
            uidoc = self._revit.ActiveUIDocument if self._revit else None
            hwnd = None
            if self._revit is not None:
                try:
                    hwnd = revit_main_hwnd(self._revit.Application)
                except Exception:
                    hwnd = None
            position_wpf_window_top_left_at_active_view(self._win, uidoc, hwnd)
        except Exception:
            pass

    def _begin_open_grow_storyboard(self):
        """Tamaño con SizeToContent; esquina sup. izq. del área de vista; ScaleTransform 0→1 + opacidad."""
        if getattr(self, "_open_grow_storyboard_started", False):
            return
        self._open_grow_storyboard_started = True
        try:
            from System import TimeSpan
            from System.Windows import Duration
            from System.Windows.Media import ScaleTransform

            try:
                self._win.BeginAnimation(self._win.OpacityProperty, None)
            except Exception:
                pass
            sc = self._win.FindName("ArLosaRootScale")
            if sc is not None:
                try:
                    sc.BeginAnimation(ScaleTransform.ScaleXProperty, None)
                    sc.BeginAnimation(ScaleTransform.ScaleYProperty, None)
                except Exception:
                    pass
                try:
                    sc.ScaleX = 0.0
                    sc.ScaleY = 0.0
                except Exception:
                    pass

            fw = float(self._form_width_px)
            self._win.Width = fw
            try:
                self._win.SizeToContent = SizeToContent.Height
            except Exception:
                pass
            try:
                self._win.UpdateLayout()
            except Exception:
                pass

            self._position_win_top_left_active_view()
            try:
                self._base_top = float(self._win.Top)
            except Exception:
                pass

            sb = self._win.TryFindResource("ArLosaOpenGrowStoryboard")
            n = 0
            if sb is not None:
                try:
                    n = int(sb.Children.Count)
                except Exception:
                    n = 0
            if sb is None or n < 3:
                try:
                    if sc is not None:
                        sc.ScaleX = 1.0
                        sc.ScaleY = 1.0
                    self._win.Opacity = 1.0
                except Exception:
                    pass
                self._on_open_grow_storyboard_completed(None, None)
                return
            dur = Duration(TimeSpan.FromMilliseconds(float(_WINDOW_OPEN_MS)))
            for i in range(n):
                try:
                    sb.Children[i].Duration = dur
                except Exception:
                    pass
            sb.Begin(self._win, True)
        except Exception:
            try:
                self._win.Width = float(self._form_width_px)
                self._win.SizeToContent = SizeToContent.Height
                self._win.Opacity = 1.0
                sc = self._win.FindName("ArLosaRootScale")
                if sc is not None:
                    sc.ScaleX = 1.0
                    sc.ScaleY = 1.0
            except Exception:
                pass
            self._on_open_grow_storyboard_completed(None, None)

    def _actualizar_info_losa(self, floor, document):
        pass

    def _actualizar_info_losas(self, document, floor_ids):
        pass

    def _setup_ui(self):
        btn_sel = self._win.FindName("BtnSeleccionar")
        btn_col = self._win.FindName("BtnColocar")
        if btn_col:
            btn_col.Click += RoutedEventHandler(self._on_colocar)
        if btn_sel:
            btn_sel.Click += RoutedEventHandler(self._on_seleccionar)
        chk_sup = self._win.FindName("ChkMallaSuperior")
        chk_inf = self._win.FindName("ChkMallaInferior")
        if chk_sup:
            chk_sup.Checked += RoutedEventHandler(self._on_chk_malla_superior_changed)
            chk_sup.Unchecked += RoutedEventHandler(self._on_chk_malla_superior_changed)
        if chk_inf:
            chk_inf.Checked += RoutedEventHandler(self._on_chk_malla_inferior_changed)
            chk_inf.Unchecked += RoutedEventHandler(self._on_chk_malla_inferior_changed)
        try:
            from System.Windows.Input import MouseButtonEventHandler

            btn_close = self._win.FindName("BtnClose")
            title_bar = self._win.FindName("TitleBar")
            if title_bar is not None:
                def _on_titlebar_down(sender, e):
                    try:
                        self._win.DragMove()
                    except Exception:
                        pass

                title_bar.MouseLeftButtonDown += MouseButtonEventHandler(_on_titlebar_down)
            if btn_close is not None:
                def _on_close_click(sender, e):
                    try:
                        self._close_with_fade()
                    except Exception:
                        pass

                btn_close.Click += RoutedEventHandler(_on_close_click)

                def _on_close_down(sender, e):
                    try:
                        e.Handled = True
                    except Exception:
                        pass

                btn_close.MouseLeftButtonDown += MouseButtonEventHandler(_on_close_down)
        except Exception:
            pass
        self._win.Loaded += RoutedEventHandler(self._on_window_loaded)

    def _on_chk_malla_superior_changed(self, sender, args):
        try:
            chk = self._win.FindName("ChkMallaSuperior")
            enabled = chk.IsChecked == True if chk else True
            panel = self._win.FindName("PanelSuperiorContent")
            if panel:
                panel.IsEnabled = enabled
                panel.Opacity = 1.0 if enabled else 0.35
            else:
                for name in ("CmbExteriorMajorDiametro", "CmbExteriorMajorEspaciamiento",
                            "CmbExteriorMinorDiametro", "CmbExteriorMinorEspaciamiento"):
                    ctrl = self._win.FindName(name)
                    if ctrl:
                        ctrl.IsEnabled = enabled
                        ctrl.Opacity = 1.0 if enabled else 0.35
        except Exception:
            pass

    def _on_chk_malla_inferior_changed(self, sender, args):
        try:
            chk = self._win.FindName("ChkMallaInferior")
            enabled = chk.IsChecked == True if chk else True
            panel = self._win.FindName("PanelInferiorContent")
            if panel:
                panel.IsEnabled = enabled
                panel.Opacity = 1.0 if enabled else 0.35
            else:
                for name in ("CmbInteriorMajorDiametro", "CmbInteriorMajorEspaciamiento",
                            "CmbInteriorMinorDiametro", "CmbInteriorMinorEspaciamiento"):
                    ctrl = self._win.FindName(name)
                    if ctrl:
                        ctrl.IsEnabled = enabled
                        ctrl.Opacity = 1.0 if enabled else 0.35
        except Exception:
            pass

    def _sync_malla_checkboxes(self):
        """Sincroniza el estado de los paneles con los checkboxes."""
        try:
            self._on_chk_malla_superior_changed(None, None)
            self._on_chk_malla_inferior_changed(None, None)
        except Exception:
            pass

    def _wire_lifecycle_handlers(self):
        try:
            def _on_closed(sender, args):
                try:
                    _clear_active_window()
                except Exception:
                    pass

            self._win.Closed += RoutedEventHandler(_on_closed)
        except Exception:
            pass

    def _wire_commands(self):
        def _on_close_cmd(sender, e):
            try:
                self._close_with_fade()
            except Exception:
                pass

        try:
            self._win.CommandBindings.Add(
                CommandBinding(ApplicationCommands.Close, RoutedEventHandler(_on_close_cmd))
            )
            self._win.InputBindings.Add(
                KeyBinding(ApplicationCommands.Close, Key.Escape, ModifierKeys.None)
            )
        except Exception:
            pass

    def _close_with_fade(self):
        """Inverso de ArLosaOpenGrowStoryboard: escala 1→0 y opacidad 1→0 (misma duración / EaseIn que fundación aislada)."""
        if getattr(self, "_is_closing_with_fade", False):
            return
        self._is_closing_with_fade = True
        try:
            from System import TimeSpan
            from System.Windows import Duration
            from System.Windows.Media import ScaleTransform
            from System.Windows.Media.Animation import DoubleAnimation, QuadraticEase, EasingMode

            try:
                self._win.BeginAnimation(self._win.OpacityProperty, None)
                self._win.BeginAnimation(self._win.TopProperty, None)
                self._win.BeginAnimation(self._win.WidthProperty, None)
                self._win.BeginAnimation(self._win.HeightProperty, None)
            except Exception:
                pass

            sc = self._win.FindName("ArLosaRootScale")
            if sc is not None:
                try:
                    sc.BeginAnimation(ScaleTransform.ScaleXProperty, None)
                    sc.BeginAnimation(ScaleTransform.ScaleYProperty, None)
                except Exception:
                    pass

            dur = Duration(TimeSpan.FromMilliseconds(float(_WINDOW_CLOSE_MS)))
            ease_in = QuadraticEase()
            ease_in.EasingMode = EasingMode.EaseIn

            def _da(from_v, to_v):
                a = DoubleAnimation()
                a.From = float(from_v)
                a.To = float(to_v)
                a.Duration = dur
                a.EasingFunction = ease_in
                return a

            try:
                sx0 = float(sc.ScaleX) if sc is not None else 1.0
                sy0 = float(sc.ScaleY) if sc is not None else 1.0
            except Exception:
                sx0 = sy0 = 1.0
            try:
                op0 = float(self._win.Opacity)
            except Exception:
                op0 = 1.0

            ax = _da(sx0, 0.0)
            ay = _da(sy0, 0.0)
            op_anim = _da(op0, 0.0)

            def _on_done(sender, args):
                try:
                    self._win.Close()
                except Exception:
                    pass

            op_anim.Completed += _on_done
            if sc is not None:
                sc.BeginAnimation(ScaleTransform.ScaleXProperty, ax)
                sc.BeginAnimation(ScaleTransform.ScaleYProperty, ay)
            self._win.BeginAnimation(self._win.OpacityProperty, op_anim)
        except Exception:
            self._is_closing_with_fade = False
            try:
                self._win.Close()
            except Exception:
                pass

    def _show_with_fade(self):
        """
        Muestra con opacidad 0; Width/Height/Opacity se animan desde el Storyboard (Window.Loaded → cola).
        """
        try:
            try:
                self._win.BeginAnimation(self._win.OpacityProperty, None)
                self._win.BeginAnimation(self._win.TopProperty, None)
                self._win.BeginAnimation(self._win.WidthProperty, None)
                self._win.BeginAnimation(self._win.HeightProperty, None)
            except Exception:
                pass

            self._win.Opacity = 0.0
            if not self._win.IsVisible:
                self._win.Show()
            try:
                self._win.UpdateLayout()
            except Exception:
                pass

            try:
                self._base_top = float(self._win.Top)
            except Exception:
                if self._base_top is None:
                    self._base_top = 0.0

            self._is_closing_with_fade = False
            self._win.Activate()
        except Exception:
            try:
                self._win.Opacity = 1.0
            except Exception:
                pass
            try:
                if not self._win.IsVisible:
                    self._win.Show()
                self._is_closing_with_fade = False
                self._win.Activate()
            except Exception:
                pass

    def _on_window_loaded(self, sender, args):
        self._load_logo()
        self._cargar_combos()
        self._sync_malla_checkboxes()
        try:
            self._win.Dispatcher.BeginInvoke(
                Action(self._begin_open_grow_storyboard),
                DispatcherPriority.Loaded,
            )
        except Exception:
            try:
                self._begin_open_grow_storyboard()
            except Exception:
                pass

    def _load_logo(self):
        """Carga logo.png desde esta carpeta o fallback a otras apps BIMTools."""
        try:
            img_ctrl = self._win.FindName("ImgLogo")
            if not img_ctrl:
                return
            for logo_path in _LOGO_PATHS:
                if os.path.exists(logo_path):
                    bmp = BitmapImage()
                    bmp.BeginInit()
                    bmp.UriSource = Uri(logo_path, UriKind.Absolute)
                    bmp.CacheOption = BitmapCacheOption.OnLoad
                    bmp.EndInit()
                    bmp.Freeze()
                    img_ctrl.Source = bmp
                    break
        except Exception:
            pass

    def _cargar_combos(self):
        try:
            d = self._document or self._revit.ActiveUIDocument.Document
        except Exception:
            d = self._document
        self._area_reinforcement_type_id = _get_default_area_reinforcement_type_id(d)
        if not self._area_reinforcement_type_id:
            pass
        bar_types = _get_rebar_bar_types(d)
        self._rebar_type_ids = {}
        for disp, bar_type in bar_types:
            self._rebar_type_ids[str(disp)] = bar_type.Id
        bar_disps = [disp for disp, _ in bar_types]
        espaciamientos = ["100", "150", "200", "250", "300"]
        diam_names = ("CmbExteriorMajorDiametro", "CmbExteriorMinorDiametro", "CmbInteriorMajorDiametro", "CmbInteriorMinorDiametro")
        esp_names = ("CmbExteriorMajorEspaciamiento", "CmbExteriorMinorEspaciamiento", "CmbInteriorMajorEspaciamiento", "CmbInteriorMinorEspaciamiento")
        for name in diam_names:
            cmb = self._win.FindName(name)
            if cmb:
                cmb.ItemsSource = bar_disps
                if bar_types:
                    cmb.SelectedIndex = min(1, len(bar_types) - 1) if len(bar_types) > 1 else 0
        for name in esp_names:
            cmb = self._win.FindName(name)
            if cmb:
                cmb.ItemsSource = espaciamientos
                cmb.SelectedIndex = 1  # 150 mm por defecto

    def _on_seleccionar(self, sender, args):
        self._win.Hide()
        self._seleccion_event.Raise()

    def _on_colocar(self, sender, args):
        from Autodesk.Revit.DB import ElementId
        try:
            if not self._floor_ids:
                _task_dialog_show(
                    "Malla en Losa",
                    u"Primero selecciona una o más losas.",
                    self._win,
                )
                return
            chk_sup = self._win.FindName("ChkMallaSuperior")
            chk_inf = self._win.FindName("ChkMallaInferior")
            if chk_sup and chk_inf and chk_sup.IsChecked != True and chk_inf.IsChecked != True:
                _task_dialog_show(
                    "Malla en Losa",
                    u"Por lo menos una malla debe estar activada.",
                    self._win,
                )
                return
            area_type_id = self._area_reinforcement_type_id
            if not area_type_id:
                try:
                    d = self._document or self._revit.ActiveUIDocument.Document
                    area_type_id = _get_default_area_reinforcement_type_id(d)
                except Exception:
                    pass
            if not area_type_id:
                _task_dialog_show(
                    "Malla en Losa",
                    u"No hay tipo de Area Reinforcement en el proyecto.",
                    self._win,
                )
                return
            rebar_ids = getattr(self, "_rebar_type_ids", {})
            chk_sup = self._win.FindName("ChkMallaSuperior")
            chk_inf = self._win.FindName("ChkMallaInferior")
            malla_superior_activa = chk_sup.IsChecked == True if chk_sup else True
            malla_inferior_activa = chk_inf.IsChecked == True if chk_inf else True
            layer_config = [
                ("exterior_major", "CmbExteriorMajorDiametro", "CmbExteriorMajorEspaciamiento", malla_superior_activa),
                ("exterior_minor", "CmbExteriorMinorDiametro", "CmbExteriorMinorEspaciamiento", malla_superior_activa),
                ("interior_major", "CmbInteriorMajorDiametro", "CmbInteriorMajorEspaciamiento", malla_inferior_activa),
                ("interior_minor", "CmbInteriorMinorDiametro", "CmbInteriorMinorEspaciamiento", malla_inferior_activa),
            ]
            params_dict = {}
            layer_active_dict = {}
            for layer_key, diam_name, esp_name, is_active in layer_config:
                cmb_diam = self._win.FindName(diam_name)
                cmb_esp = self._win.FindName(esp_name)
                bar_id = rebar_ids.get(str(cmb_diam.SelectedItem if cmb_diam else None), ElementId.InvalidElementId)
                esp = (cmb_esp.SelectedItem if cmb_esp and cmb_esp.SelectedItem else None) or (cmb_esp.Text if cmb_esp else None) or "150"
                params_dict[layer_key] = (bar_id, str(esp))
                layer_active_dict[layer_key] = bool(is_active)
            if not any(pid and pid != ElementId.InvalidElementId for pid, _ in params_dict.values()):
                _task_dialog_show(
                    "Malla en Losa",
                    u"Selecciona al menos un diámetro válido.",
                    self._win,
                )
                return
            self._colocar_handler.floor_ids = list(self._floor_ids)
            self._colocar_handler.params_dict = params_dict
            self._colocar_handler.layer_active_dict = layer_active_dict
            self._colocar_handler.area_reinforcement_type_id = area_type_id
            self._colocar_handler.asignar_ganchos = True
            self._colocar_event.Raise()
        except Exception as ex:
            _task_dialog_show(
                "Malla en Losa - Error",
                u"Error:\n\n{}".format(str(ex)),
                self._win,
            )

    def show(self):
        uidoc = self._revit.ActiveUIDocument
        if uidoc is None:
            _task_dialog_show(
                _TOOL_TASK_DIALOG_TITLE,
                u"No hay documento activo.",
                self._win,
            )
            return
        hwnd = None
        try:
            from System.Windows.Interop import WindowInteropHelper

            hwnd = revit_main_hwnd(self._revit.Application)
            if hwnd:
                helper = WindowInteropHelper(self._win)
                helper.Owner = hwnd
        except Exception:
            pass
        position_wpf_window_top_left_at_active_view(self._win, uidoc, hwnd)
        self._document = uidoc.Document
        self._cargar_combos()
        self._show_with_fade()


def run(revit, close_on_finish=False):
    """Punto de entrada: lanza la ventana Malla en Losa."""
    existing = _get_active_window()
    if existing is not None:
        try:
            from System.Windows import WindowState

            if existing.WindowState == WindowState.Minimized:
                existing.WindowState = WindowState.Normal
        except Exception:
            pass
        try:
            existing.Show()
        except Exception:
            pass
        try:
            existing.Activate()
            existing.Focus()
        except Exception:
            pass
        _task_dialog_show(
            _TOOL_TASK_DIALOG_TITLE,
            u"La herramienta ya está en ejecución.",
            existing,
        )
        return

    w = AreaReinforcementLosaWindow(revit, close_on_finish=close_on_finish)
    _set_active_window(w._win)
    w.show()
