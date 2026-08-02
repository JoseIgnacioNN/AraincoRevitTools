# -*- coding: utf-8 -*-
"""
Arainco: Suples losa.

Flujo:
  1. Pick Floor → planta Sketch + contexto
  2. Tab Superior → tipos «Suple en apoyo» / «Suple en Borde»
     (Inferior = placeholder)
  3. Apoyo: Paño 1 + Paño 2 + recorrido → L = ¼·max(luz menor)
  4. Borde: 1 polígono → lm → L=¼·lm → recorrido
     → AR strip (borde−25 mm → +L adentro), Top Major
  5. Apoyo: AR strip 2L × recorrido, Top Major
  6. Remove Area System → Rebar libres + Show Middle / stamps / tags / MRA

Revit 2024+ | IronPython (pyRevit).
"""

from __future__ import print_function

import math
import weakref

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")

from System import AppDomain, EventHandler, TimeSpan
from System.Windows import (
    CornerRadius,
    FontWeights,
    HorizontalAlignment,
    Point as WpfPoint,
    RoutedEventHandler,
    SizeChangedEventHandler,
    TextWrapping,
    Thickness,
    VerticalAlignment,
    Visibility,
    WindowState,
)
from System.Windows import GridLength, GridUnitType
from System.Windows.Controls import (
    Border,
    Button,
    CheckBox,
    Canvas as WpfCanvas,
    ColumnDefinition,
    ComboBox,
    ComboBoxItem,
    Dock,
    DockPanel,
    Grid,
    Orientation,
    StackPanel,
    TextBlock,
)
from System.Windows.Controls import SelectionChangedEventHandler
from System.Windows.Input import (
    Cursors,
    Key,
    KeyEventHandler,
    MouseButton,
    MouseButtonEventHandler,
    MouseEventHandler,
    MouseWheelEventHandler,
)
from System.Windows.Markup import XamlReader
from System.Windows.Media import (
    Color,
    DoubleCollection,
    Matrix,
    MatrixTransform,
    PointCollection,
    SolidColorBrush,
    TranslateTransform,
)
from System.Windows.Shapes import Ellipse as WpfEllipse
from System.Windows.Shapes import Line as WpfLine
from System.Windows.Shapes import Polygon as WpfPolygon
from System.Windows.Shapes import Rectangle as WpfRectangle
from System.Windows.Threading import DispatcherTimer

from Autodesk.Revit.DB import (
    BuiltInCategory,
    Category,
    Floor,
    IndependentTag,
    Reference,
    TagMode,
    TagOrientation,
    Transaction,
)
from Autodesk.Revit.DB.Structure import AreaReinforcement, Rebar, RebarInSystem
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler, TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

from bimtools_ui_tokens import WINDOW_CHROME_TITLE
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from revit_wpf_window_position import (
    bind_center_wpf_on_revit_monitor,
    position_wpf_window_center_on_monitor,
    revit_main_hwnd,
)

try:
    from bimtools_instruction_dialog import show_message_dialog
except Exception:
    show_message_dialog = None

from area_rein_losa_panos import (  # noqa: E402
    _line_edge_intersection,
    _snap_array_length_to_spacing_multiple,
    luz_menor_mm_from_polygon,
    rect_from_two_points_mm,
)
try:
    from conjunto_guid import (
        ARMADURA_UBICACION_SUPERIOR,
        generar_armadura_conjunto_guid,
        stamp_armadura_arainco,
        stamp_armadura_conjunto_guid,
        stamp_armadura_malla,
        stamp_armadura_nivel,
        stamp_armadura_ubicacion,
    )
except Exception:
    ARMADURA_UBICACION_SUPERIOR = u"F'"
    generar_armadura_conjunto_guid = None
    stamp_armadura_arainco = None
    stamp_armadura_conjunto_guid = None
    stamp_armadura_malla = None
    stamp_armadura_nivel = None
    stamp_armadura_ubicacion = None

try:
    from bimtools_rebar_3d_visibility import apply_reinforcement_unobscured_in_view
except Exception:
    apply_reinforcement_unobscured_in_view = None

from area_rein_losa_sketch import (  # noqa: E402
    _CTX_GRID,
    _HUD_SCALE_TAG,
    _LAYER_KEYS,
    _SNAP_PX,
    _SNAP_TAG,
    _append_polyline_snap,
    _append_ring_snap,
    _append_wall_beam_intersection_snap,
    _aplicar_show_middle_barras_area_reinforcement,
    _bar_types_sorted,
    _bbox_mm,
    _brush,
    _build_snap_cell_index,
    _build_wall_beam_geo_mm,
    _count_ctx,
    _layer_cfg_for_keys,
    _loop_to_polyline_mm,
    _nivel_losa_como_string,
    _plane_from_curves,
    _plane_mm_dir_to_xyz,
    _plane_mm_to_xyz,
    _poly_mm_to_curves,
    _aplicar_estilo_tag_rebar_sin_leader,
    _crear_independent_tag_rebar,
    _presentacion_show_middle_en_vista,
    _proyectar_punto_plano_vista,
    _punto_insercion_tag_show_middle,
    _resolver_tag_type_id_por_shape,
    _resolver_vista_para_show_middle,
    _snap_point_mm,
    _vista_ok_para_etiquetas_rebar,
    collect_existing_area_rein_on_floor,
    crear_area_reinforcement,
    obtener_loops_sketch,
    paint_planta_context_layers,
    recolectar_contexto_planta,
)

try:
    from geometria_estribos_viga import (
        _multi_reference_annotation_type_by_name,
        _vista_permite_multi_rebar_annotation,
        crear_multi_rebar_annotations_por_nombre_tipo,
    )
except Exception:
    _multi_reference_annotation_type_by_name = None
    _vista_permite_multi_rebar_annotation = None
    crear_multi_rebar_annotations_por_nombre_tipo = None

_MRA_TYPE_NAME_RECORRIDO_BARRAS = u"Recorrido Barras"
# Offset lateral MRA respecto al array (mm), además de medio ancho del bbox
_MRA_SUPLE_OFFSET_EXTRA_MM = 300.0

# Misma convención que Area Rein. Losa Sketch (63_AreaReinLosaSketch):
# familia TAG_FLOOR, sin leader, tipo = RebarShape (fallback «01»).
# WALL_HORIZONTAL solo si TAG_FLOOR no está cargada en el proyecto.
_SUPLE_REBAR_TAG_FAMILY_NAME = u"EST_A_STRUCTURAL REBAR TAG_FLOOR"
_SUPLE_REBAR_TAG_FAMILY_FALLBACK = u"EST_A_STRUCTURAL REBAR TAG_WALL_HORIZONTAL"
_SUPLE_REBAR_TAG_FALLBACK_TYPE = u"01"

_DIALOG_TITLE = u"Arainco: Suples losa"
_SINGLETON_KEY = u"Arainco.SuplesLosa.ActiveWindow"
_SPACING_OPTS_MM = (100, 125, 150, 200, 250, 300)

_DRAW_IDLE = u"idle"
_DRAW_PANO1 = u"pano1"
_DRAW_PANO2 = u"pano2"
_DRAW_RECORRIDO = u"recorrido"
_DRAW_BORDE_POLY = u"borde_poly"

# Tipos activos en tab Superior (mutuamente exclusivos en UI)
_SUP_TYPE_APOYO = u"apoyo"
_SUP_TYPE_BORDE = u"borde"

_PANO1_COLOR = u"#38bdf8"
_PANO2_COLOR = u"#a78bfa"
_RECORRIDO_COLOR = u"#fbbf24"
_BAR_PREVIEW_COLOR = u"#E8F4F8"
_FACE_SUP = u"#5BC0DE"
_FACE_BORDE = u"#f59e0b"  # acento tipográfico del tipo Borde
# Recubrimiento en el extremo al borde de losa (mm)
_COVER_BORDE_MM = 25.0

_SUP_TYPE_HINTS = {
    _SUP_TYPE_APOYO: (
        u"SUP → Suple en apoyo: P1·P2·recorrido y luego Colocar armadura."
    ),
    _SUP_TYPE_BORDE: (
        u"SUP → Suple en Borde: polígono → recorrido → Colocar armadura (AR)."
    ),
}

# Orden = tabs (Superior primero = default)
_FACE_GROUPS = (
    {
        u"id": u"superior",
        u"title": u"Superior",
        u"pill": u"SUP",
        u"color": _FACE_SUP,
        u"hint": _SUP_TYPE_HINTS[_SUP_TYPE_APOYO],
    },
    {
        u"id": u"inferior",
        u"title": u"Inferior",
        u"pill": u"INF",
        u"color": u"#4ade80",
        u"hint": u"INF → tipos inferiores (placeholder).",
    },
)

_XAML = u"""
<Window
  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  Title="__CHROME__"
  Height="680" Width="920"
  MinHeight="560" MinWidth="780"
  ResizeMode="CanResize"
  WindowStartupLocation="Manual"
  Background="#071018"
  FontFamily="Segoe UI"
  FontSize="12"
  ShowInTaskbar="False">
  <Window.Resources>
__STYLES__
  </Window.Resources>
  <Border Background="#071018" BorderBrush="#21465C" BorderThickness="1" Padding="18">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <StackPanel Grid.Row="0" Margin="0,0,0,8">
        <TextBlock x:Name="TxtTitle" Text="Arainco: Suples losa"
                   Foreground="#E8F4F8" FontSize="18" FontWeight="Bold"/>
        <TextBlock x:Name="TxtSubtitle" Margin="0,6,0,0" Foreground="#95B8CC"
                   FontSize="11" TextWrapping="Wrap"
                   Text="Superior: Suple en apoyo / Suple en Borde · Inferior por definir."/>
      </StackPanel>

      <Grid Grid.Row="1">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="340"/>
        </Grid.ColumnDefinitions>

        <Border Grid.Column="0" Background="#0a1620" BorderBrush="#21465C"
                BorderThickness="1" CornerRadius="4,0,0,4" Padding="0">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Border Grid.Row="0" Background="#0a1620" BorderBrush="#21465C"
                    BorderThickness="0,0,0,1" Padding="8,6,8,4">
              <TextBlock x:Name="TxtCanvasHeader" Foreground="#64748b"
                         FontSize="10" FontWeight="SemiBold"
                         Text="PLANTA · SKETCH (mm)"/>
            </Border>
            <Border Grid.Row="1" Background="#050E18" BorderBrush="Transparent"
                    BorderThickness="0" Padding="8,4,8,8">
              <Border Background="#050E18" BorderBrush="#21465C"
                      BorderThickness="1" CornerRadius="4">
                <Canvas x:Name="CvPlan" ClipToBounds="True" Focusable="True"/>
              </Border>
            </Border>
          </Grid>
        </Border>

        <Border Grid.Column="1" Background="#0a1620" BorderBrush="#21465C"
                BorderThickness="1" CornerRadius="0,4,4,0" Padding="8,8">
          <ScrollViewer VerticalScrollBarVisibility="Auto"
                        HorizontalScrollBarVisibility="Disabled">
            <StackPanel x:Name="PnlSectionRail">
              <Border Background="#0a1620" BorderBrush="#21465C"
                      BorderThickness="1" CornerRadius="4" Padding="10">
                <StackPanel>
                  <TextBlock Text="Capas de suple" Foreground="#E8F4F8"
                             FontSize="12" FontWeight="SemiBold" Margin="0,0,0,6"/>
                  <TextBlock x:Name="TxtLayersHint" Foreground="#64748b" FontSize="10"
                             Margin="0,0,0,8" TextWrapping="Wrap"
                             Text="SUP → Suple en apoyo: P1·P2·recorrido y luego Colocar armadura."/>
                  <StackPanel x:Name="PnlFaces"/>
                </StackPanel>
              </Border>
            </StackPanel>
          </ScrollViewer>
        </Border>
      </Grid>

      <TextBlock Grid.Row="2" x:Name="TxtHint" Foreground="#64748b" FontSize="10"
                 TextWrapping="Wrap" Margin="0,8,0,0"
                 Text="Snap a vértices/aristas · Esc cancela clic · rueda zoom · botón medio pan."/>

      <Grid Grid.Row="3" Margin="0,14,0,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="TxtStatus" Grid.Column="0" VerticalAlignment="Center"
                   Foreground="#64748b" FontSize="10" TextWrapping="Wrap" Margin="0,0,12,0"/>
        <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
          <Button x:Name="BtnCancelar" Content="Cancelar"
                  Style="{StaticResource BtnSelectOutline}" MinWidth="110" Margin="0,0,10,0"/>
          <Button x:Name="BtnCrear" Content="Colocar armadura"
                  Style="{StaticResource BtnPrimary}" MinWidth="150"/>
        </StackPanel>
      </Grid>
    </Grid>
  </Border>
</Window>
""".replace(u"__CHROME__", WINDOW_CHROME_TITLE).replace(
    u"__STYLES__", BIMTOOLS_DARK_STYLES_XML
)


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _element_id_int(eid):
    if eid is None:
        return 0
    try:
        return int(eid.IntegerValue)
    except Exception:
        pass
    try:
        return int(eid.Value)
    except Exception:
        return 0


def _suple_ensure_rebar_tag_visibility(doc, view):
    """Hace visibles anotaciones / OST_RebarTags en la vista."""
    if view is None:
        return
    try:
        if bool(getattr(view, u"AreAnnotationCategoriesHidden", False)):
            view.AreAnnotationCategoriesHidden = False
    except Exception:
        pass
    try:
        cat = Category.GetCategory(doc, BuiltInCategory.OST_RebarTags)
        if cat is None:
            cat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_RebarTags)
        if cat is not None and bool(view.GetCategoryHidden(cat.Id)):
            view.SetCategoryHidden(cat.Id, False)
    except Exception:
        pass


def _suple_element_visible_in_view(doc, view, element_id):
    """True si el elemento entra al collector de visibles de la vista."""
    if doc is None or view is None or element_id is None:
        return False
    try:
        from Autodesk.Revit.DB import (
            ElementId,
            ElementIdSetFilter,
            FilteredElementCollector,
        )
        from System.Collections.Generic import List as ClrList

        ids = ClrList[ElementId]()
        ids.Add(element_id)
        col = (
            FilteredElementCollector(doc, view.Id)
            .WhereElementIsNotElementType()
            .WherePasses(ElementIdSetFilter(ids))
        )
        for _el in col:
            return True
    except Exception:
        pass
    return False


def _suple_ensure_rebar_host_visible(doc, view, barra, do_regenerate=False):
    """
    Tras RemoveAreaSystem, fuerza OST_Rebar + Unobscured/Solid/Show Middle
    para que IndependentTag pueda pintar en la vista activa.

    ``do_regenerate`` default False: el caller debe regenerar en lote
    (p. ej. ``_suple_prepare_rebars_for_annotation``).
    """
    if doc is None or view is None or barra is None:
        return
    try:
        cat = Category.GetCategory(doc, BuiltInCategory.OST_Rebar)
        if cat is not None and bool(view.GetCategoryHidden(cat.Id)):
            view.SetCategoryHidden(cat.Id, False)
    except Exception:
        pass
    try:
        if hasattr(barra, u"SetUnobscuredInView"):
            barra.SetUnobscuredInView(view, True)
    except Exception:
        pass
    try:
        if hasattr(barra, u"SetSolidInView"):
            barra.SetSolidInView(view, True)
    except Exception:
        pass
    try:
        _presentacion_show_middle_en_vista(barra, view)
    except Exception:
        pass
    if do_regenerate:
        try:
            doc.Regenerate()
        except Exception:
            pass


def _suple_prepare_rebars_for_annotation(doc, view, rebars):
    """
    Un solo pase: categoría + unobscured + Show Middle + 1 Regenerate.
    Evita N regenerates por barra al etiquetar / MRA.
    """
    if doc is None or view is None or not rebars:
        return
    try:
        cat = Category.GetCategory(doc, BuiltInCategory.OST_Rebar)
        if cat is not None and bool(view.GetCategoryHidden(cat.Id)):
            view.SetCategoryHidden(cat.Id, False)
    except Exception:
        pass
    if apply_reinforcement_unobscured_in_view is not None:
        try:
            apply_reinforcement_unobscured_in_view(
                doc, rebars, view, unobscured=True
            )
        except Exception:
            pass
    for barra in rebars:
        if barra is None:
            continue
        try:
            if hasattr(barra, u"SetSolidInView"):
                barra.SetSolidInView(view, True)
        except Exception:
            pass
        try:
            _presentacion_show_middle_en_vista(barra, view)
        except Exception:
            pass
    try:
        doc.Regenerate()
    except Exception:
        pass


def _suple_tag_in_view_collector(doc, view, tag_id):
    return _suple_element_visible_in_view(doc, view, tag_id)


def _suple_crear_tag_rebar_libre(doc, view, barra, tag_type_id, point):
    """
    Create para Rebar libres (post RemoveAreaSystem).

    Distinto de Sketch (RebarInSystem): prioriza ``Reference(barra)`` completa;
    las refs de subelemento/posición a menudo crean tags con OwnerView pero
    sin gráficos en el collector de la vista. Prueba sin/con leader.
    """
    if doc is None or view is None or barra is None or tag_type_id is None:
        return None
    if point is None:
        return None

    refs = []
    try:
        refs.append(Reference(barra))
    except Exception:
        pass
    try:
        from enfierrado_shaft_hashtag import _rebar_reference_candidates_for_tag

        for r in _rebar_reference_candidates_for_tag(doc, view, barra) or []:
            if r is not None:
                refs.append(r)
    except Exception:
        pass
    if not refs:
        return None

    try:
        sym = doc.GetElement(tag_type_id)
        if sym is not None and hasattr(sym, u"IsActive") and not bool(sym.IsActive):
            sym.Activate()
    except Exception:
        pass

    def _try_one_tag(add_leader, ref, orient):
        created_local = None
        try:
            created_local = IndependentTag.Create(
                doc,
                tag_type_id,
                view.Id,
                ref,
                add_leader,
                orient,
                point,
            )
        except Exception:
            created_local = None
        if created_local is None:
            try:
                created_local = IndependentTag.Create(
                    doc,
                    view.Id,
                    ref,
                    add_leader,
                    TagMode.TM_ADDBY_CATEGORY,
                    orient,
                    point,
                )
                if created_local is not None:
                    try:
                        created_local.SetTypeId(tag_type_id)
                    except Exception:
                        pass
            except Exception:
                created_local = None
        if created_local is None:
            return None
        try:
            if not add_leader:
                _aplicar_estilo_tag_rebar_sin_leader(created_local, point)
            else:
                created_local.HasLeader = True
                created_local.TagHeadPosition = point
        except Exception:
            pass
        try:
            doc.Regenerate()
        except Exception:
            pass
        if _suple_tag_in_view_collector(doc, view, created_local.Id):
            return created_local
        try:
            doc.Delete(created_local.Id)
        except Exception:
            pass
        return None

    # Camino preferido primero (caso feliz: 1 regenerate)
    try:
        hit = _try_one_tag(False, refs[0], TagOrientation.Horizontal)
        if hit is not None:
            return hit
    except Exception:
        pass
    for add_leader in (False, True):
        for ri, ref in enumerate(refs):
            for orient in (TagOrientation.Horizontal, TagOrientation.Vertical):
                if (
                    not add_leader
                    and ri == 0
                    and orient == TagOrientation.Horizontal
                ):
                    continue  # ya intentado
                hit = _try_one_tag(add_leader, ref, orient)
                if hit is not None:
                    return hit
    return None


def _mostrar_aviso(uiapp, instruction, content=u""):
    try:
        hwnd = revit_main_hwnd(uiapp) if uiapp is not None else None
        if show_message_dialog is not None:
            show_message_dialog(
                _DIALOG_TITLE,
                instruction=_as_unicode(instruction),
                content=_as_unicode(content) if content else None,
                ok_text=u"Entendido",
                hwnd_revit=hwnd,
                uiapp=uiapp,
            )
            return
    except Exception:
        pass
    try:
        TaskDialog.Show(
            _DIALOG_TITLE,
            u"{0}\n{1}".format(
                _as_unicode(instruction), _as_unicode(content)
            ).strip(),
        )
    except Exception:
        pass


def _focus_existing(uiapp):
    try:
        win = AppDomain.CurrentDomain.GetData(_SINGLETON_KEY)
    except Exception:
        win = None
    if win is None:
        return False
    try:
        if not bool(win.IsLoaded):
            return False
    except Exception:
        return False
    try:
        win.WindowState = WindowState.Maximized
    except Exception:
        pass
    try:
        win.Activate()
        win.Focus()
    except Exception:
        pass
    _mostrar_aviso(uiapp, u"La herramienta ya esta en ejecucion.")
    return True


def _register_singleton(win):
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, win)
    except Exception:
        pass


def _unregister_singleton():
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, None)
    except Exception:
        pass


class _FloorSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, Floor)

    def AllowReference(self, reference, position):
        return False


def _pick_floor(uidoc, doc, uiapp):
    try:
        ids = list(uidoc.Selection.GetElementIds())
        for eid in ids:
            el = doc.GetElement(eid)
            if isinstance(el, Floor):
                return el
    except Exception:
        pass
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            _FloorSelectionFilter(),
            u"Seleccione una losa (Floor)",
        )
        if ref is None:
            return None
        el = doc.GetElement(ref.ElementId)
        if isinstance(el, Floor):
            return el
    except Exception:
        return None
    _mostrar_aviso(uiapp, u"El elemento seleccionado no es una losa (Floor).")
    return None


def _centroid_mm(pts):
    if not pts:
        return (0.0, 0.0)
    n = float(len(pts))
    return (
        sum(float(p[0]) for p in pts) / n,
        sum(float(p[1]) for p in pts) / n,
    )


def _unit_mm(dx, dy):
    L = math.hypot(dx, dy)
    if L < 1e-9:
        return None
    return (dx / L, dy / L)


def _ray_ring_first_hit_mm(origin, dir_xy, ring, min_t=0.0):
    """
    Primer hit de la semirrecta origin + t·dir (t ≥ min_t) con el anillo cerrado.

    ``dir_xy`` debe ser unitario. Returns (pt, t_mm) o (None, None).
    """
    if not ring or len(ring) < 2 or dir_xy is None:
        return None, None
    try:
        ox, oy = float(origin[0]), float(origin[1])
        dx, dy = float(dir_xy[0]), float(dir_xy[1])
    except Exception:
        return None, None
    if math.hypot(dx, dy) < 1e-12:
        return None, None
    a = (ox, oy)
    b = (ox + dx, oy + dy)
    best_t = None
    best_pt = None
    n = len(ring)
    # Evitar arista de cierre duplicada si ring[0]==ring[-1]
    last = n - 1
    try:
        if (
            abs(float(ring[0][0]) - float(ring[-1][0])) < 1e-6
            and abs(float(ring[0][1]) - float(ring[-1][1])) < 1e-6
        ):
            last = n - 2
    except Exception:
        pass
    for i in range(max(0, last) + 1):
        c = ring[i]
        d = ring[(i + 1) % n]
        try:
            res = _line_edge_intersection(a, b, c, d)
        except Exception:
            res = None
        if res is None:
            continue
        pt, t, _u = res
        try:
            t = float(t)
        except Exception:
            continue
        if t < float(min_t) - 1e-6:
            continue
        if best_t is None or t < best_t:
            best_t = t
            best_pt = (float(pt[0]), float(pt[1]))
    if best_pt is None:
        return None, None
    return best_pt, best_t


def _stations_exact_spacing_mm(rec_len, esp_mm):
    """
    Estaciones a lo largo del recorrido con paso exacto ``esp_mm``.

    Requiere ``rec_len`` múltiplo de ``esp`` (tras snap). Incluye extremos:
    ``0, e, 2e, …, n·e`` con ``n = rec_len / e``.
    """
    try:
        L = float(rec_len)
        s = float(esp_mm)
    except Exception:
        return [0.0]
    if L < 1.0 or s < 1.0:
        return [0.0]
    n_int = int(math.floor((L / s) + 1e-9))
    if n_int < 1:
        return [0.0]
    return [float(i) * s for i in range(n_int + 1)]


def _snap_recorrido_to_spacing_mm(ax, ay, dx, dy, rec_len, esp_mm):
    """
    Ajusta la luz de distribución a ``floor(L/e)·e`` (Maximum Spacing exacto).

    Centra el exceso en el recorrido (no crece). Devuelve
    ``(ax, ay, rec_len', stations)`` o ``None`` si no hay snap viable.
    """
    snapped = _snap_array_length_to_spacing_multiple(rec_len, esp_mm)
    if snapped is None:
        return None
    try:
        L0 = float(rec_len)
        L1 = float(snapped)
        ax = float(ax)
        ay = float(ay)
        dx = float(dx)
        dy = float(dy)
        esp = float(esp_mm)
    except Exception:
        return None
    excess = L0 - L1
    if excess > 0.5:
        ax = ax + dx * (excess * 0.5)
        ay = ay + dy * (excess * 0.5)
    stations = _stations_exact_spacing_mm(L1, esp)
    return ax, ay, L1, stations


def _snap_suple_layout_to_spacing(layout, esp_mm):
    """
    Copia de ``layout`` con ``rec_len`` / ``origin`` ajustados al espaciamiento.

    No reconstruye ``bars`` (preview las calcula en ``_compute_*_layout``).
    Usado en creación AR para garantizar múltiplo de ``e`` aunque el ítem
    se haya guardado con recorrido crudo.
    """
    if layout is None:
        return None
    try:
        esp = float(esp_mm if esp_mm is not None else layout.get(u"esp") or 150)
        if esp < 1.0:
            esp = 150.0
        ax, ay = layout.get(u"origin")
        dx, dy = layout.get(u"dir")
        rec_len = float(layout.get(u"rec_len") or 0.0)
        ax, ay = float(ax), float(ay)
        dx, dy = float(dx), float(dy)
    except Exception:
        return layout
    snapped = _snap_recorrido_to_spacing_mm(ax, ay, dx, dy, rec_len, esp)
    if snapped is None:
        return layout
    ax2, ay2, rec2, stations = snapped
    out = dict(layout)
    out[u"origin"] = (ax2, ay2)
    out[u"rec_len"] = float(rec2)
    out[u"esp"] = float(esp)
    out[u"n_bars"] = len(stations)
    return out


def _strip_polygon_mm_from_layout(layout):
    """
    Contorno cerrado del strip (mm plano Sketch) para AreaReinforcement.

    - Apoyo: recorrido × ±L (perp).
    - Borde (``inward``): recorrido × [−outer_t … +L] (borde−cover → adentro).

    Major = dirección de barra = perp al recorrido.
    """
    if layout is None:
        return None
    try:
        L = float(layout.get(u"L") or 0.0)
        ax, ay = layout.get(u"origin")
        dx, dy = layout.get(u"dir")
        px, py = layout.get(u"perp")
        rec_len = float(layout.get(u"rec_len") or 0.0)
        ax, ay = float(ax), float(ay)
        dx, dy = float(dx), float(dy)
        px, py = float(px), float(py)
    except Exception:
        return None
    if L < 10.0 or rec_len < 1.0:
        return None
    if layout.get(u"inward"):
        try:
            t_out = float(layout.get(u"outer_t") or 0.0)
        except Exception:
            t_out = 0.0
        if t_out < 0.0:
            t_out = 0.0
        # Interior (+perp·L) → exterior (−perp·outer_t)
        p0 = (ax + px * L, ay + py * L)
        p1 = (ax + dx * rec_len + px * L, ay + dy * rec_len + py * L)
        p2 = (
            ax + dx * rec_len - px * t_out,
            ay + dy * rec_len - py * t_out,
        )
        p3 = (ax - px * t_out, ay - py * t_out)
        return [p0, p1, p2, p3]
    # Apoyo: a±perp*L → b±perp*L (b = a + dir*rec_len)
    p0 = (ax + px * L, ay + py * L)
    p1 = (ax + dx * rec_len + px * L, ay + dy * rec_len + py * L)
    p2 = (ax + dx * rec_len - px * L, ay + dy * rec_len - py * L)
    p3 = (ax - px * L, ay - py * L)
    return [p0, p1, p2, p3]


def _collect_system_rebars(doc, system_rein):
    """Lista de RebarInSystem hijos de Area/Path Reinforcement."""
    out = []
    if doc is None or system_rein is None:
        return out
    try:
        ids = system_rein.GetRebarInSystemIds()
    except Exception:
        ids = None
    if ids is None:
        return out
    try:
        n = int(ids.Count)
    except Exception:
        n = 0
    for i in range(n):
        try:
            el = doc.GetElement(ids[i])
        except Exception:
            el = None
        if el is not None:
            out.append(el)
    return out


def _element_ids_to_rebars(doc, id_list):
    """ElementIds → lista de Rebar / RebarInSystem."""
    out = []
    if doc is None or not id_list:
        return out
    try:
        n = int(id_list.Count)
        iterable = [id_list[i] for i in range(n)]
    except Exception:
        try:
            iterable = list(id_list)
        except Exception:
            return out
    for eid in iterable:
        try:
            el = doc.GetElement(eid)
        except Exception:
            el = None
        if isinstance(el, (Rebar, RebarInSystem)):
            out.append(el)
    return out


def _remove_area_reinforcement_system(doc, area_rein):
    """
    Equivalente UI «Remove Area Reinforcement System».
    Elimina el AR y convierte RebarInSystem → Rebar libres.
    """
    if doc is None or area_rein is None:
        return []
    try:
        new_ids = AreaReinforcement.RemoveAreaReinforcementSystem(doc, area_rein)
    except Exception:
        return []
    return _element_ids_to_rebars(doc, new_ids)


def _rebar_npos_qty(rb):
    """(n_posiciones, quantity) para ranking de sets fragmentados."""
    npos = 0
    qty = 0
    try:
        npos = int(rb.NumberOfBarPositions)
    except Exception:
        npos = 0
    try:
        qty = int(rb.Quantity)
    except Exception:
        qty = npos
    if npos <= 0:
        npos = 1
    if qty <= 0:
        qty = npos
    return npos, qty


def _mid_recorrido_xyz(item, plane):
    """Punto medio del recorrido del suple en XYZ (plano Sketch)."""
    if plane is None:
        return None
    recorrido = None
    layout = None
    try:
        recorrido = (item or {}).get(u"recorrido")
        layout = (item or {}).get(u"layout")
    except Exception:
        pass
    try:
        if recorrido is not None:
            a, b = recorrido[0], recorrido[1]
            mx = 0.5 * (float(a[0]) + float(b[0]))
            my = 0.5 * (float(a[1]) + float(b[1]))
            return _plane_mm_to_xyz((mx, my), plane)
        if layout is not None:
            ax, ay = layout.get(u"origin")
            dx, dy = layout.get(u"dir")
            rec_len = float(layout.get(u"rec_len") or 0.0)
            mx = float(ax) + float(dx) * rec_len * 0.5
            my = float(ay) + float(dy) * rec_len * 0.5
            return _plane_mm_to_xyz((mx, my), plane)
    except Exception:
        return None
    return None


def _pick_suple_annotation_rebar(doc, free_rebars, item, plane, view):
    """
    Con pasadas/aberturas el AR genera varios sets cortados.

    Para MRA se elige un solo set representativo (las etiquetas van a todos):
    el más cercano al centro del recorrido (Show Middle); si no, el de mayor cantidad.
    """
    if not free_rebars:
        return None
    refreshed = []
    for rb in free_rebars:
        try:
            el = doc.GetElement(rb.Id) if doc is not None else rb
        except Exception:
            el = rb
        if isinstance(el, (Rebar, RebarInSystem)):
            refreshed.append(el)
    if not refreshed:
        return None
    if len(refreshed) == 1:
        return refreshed[0]

    mid = _mid_recorrido_xyz(item, plane)
    best = None
    best_dist = None
    if mid is not None and view is not None:
        for rb in refreshed:
            try:
                p = _punto_insercion_tag_show_middle(rb, view)
            except Exception:
                p = None
            if p is None:
                continue
            try:
                d = mid.DistanceTo(p)
            except Exception:
                continue
            if best is None or d < best_dist:
                best = rb
                best_dist = d
    if best is not None:
        return best

    # Fallback: set con más barras / posiciones
    best = refreshed[0]
    best_score = -1
    for rb in refreshed:
        npos, qty = _rebar_npos_qty(rb)
        score = qty * 1000 + npos
        if score > best_score:
            best_score = score
            best = rb
    return best


def _suple_project_dir_on_view(xyz_dir, view):
    """Proyecta un XYZ dirección al plano de vista; None si degenera."""
    if xyz_dir is None or view is None:
        return None
    try:
        from Autodesk.Revit.DB import XYZ

        vd = view.ViewDirection
        if vd is None or vd.GetLength() < 1e-12:
            return None
        vd = vd.Normalize()
        v = xyz_dir
        if v.GetLength() < 1e-12:
            return None
        v = v.Normalize()
        dot = float(v.DotProduct(vd))
        px = float(v.X) - dot * float(vd.X)
        py = float(v.Y) - dot * float(vd.Y)
        pz = float(v.Z) - dot * float(vd.Z)
        proj = XYZ(px, py, pz)
        if proj.GetLength() < 1e-9:
            return None
        return proj.Normalize()
    except Exception:
        return None


def _suple_spacing_dir_xyz(rebar, layout, plane, view):
    """
    Dirección de distribución del set (a lo largo del recorrido) en el plano de vista.

    Preferencia: ``layout['dir']`` del sketch; si no, vector pos0→posN del Rebar;
    si no, ``view.RightDirection``.
    """
    from Autodesk.Revit.DB import XYZ

    vd = None
    rd = None
    try:
        vd = view.ViewDirection.Normalize()
        rd = view.RightDirection.Normalize()
    except Exception:
        pass

    # 1) Layout del suple (recorrido)
    if layout is not None and plane is not None:
        try:
            dx, dy = layout.get(u"dir")
            d3 = _plane_mm_dir_to_xyz(float(dx), float(dy), plane)
            if d3 is not None:
                proj = _suple_project_dir_on_view(d3, view)
                if proj is not None:
                    return proj
        except Exception:
            pass

    # 2) Posiciones del array Rebar
    try:
        from Autodesk.Revit.DB.Structure import MultiplanarOption

        n = 0
        try:
            n = int(rebar.NumberOfBarPositions)
        except Exception:
            n = 0
        if n > 1:
            for mpo_name in (
                u"IncludeAllMultiplanarCurves",
                u"IncludeOnlyPlanarCurves",
            ):
                mpo = getattr(MultiplanarOption, mpo_name, None)
                if mpo is None:
                    continue
                try:
                    cs0 = list(
                        rebar.GetCenterlineCurves(False, False, False, mpo, 0)
                    )
                    csn = list(
                        rebar.GetCenterlineCurves(False, False, False, mpo, n - 1)
                    )
                    if cs0 and csn:
                        c0 = cs0[0].Evaluate(0.5, True)
                        cn = csn[0].Evaluate(0.5, True)
                        v = cn - c0
                        if float(v.GetLength()) > 1e-6:
                            proj = _suple_project_dir_on_view(v.Normalize(), view)
                            if proj is not None:
                                return proj
                except Exception:
                    pass
    except Exception:
        pass

    return rd


def _suple_centro_rebar_mra(rebar, view):
    if rebar is None:
        return None
    try:
        bb = rebar.get_BoundingBox(view)
        if bb is not None:
            return (bb.Min + bb.Max) * 0.5
    except Exception:
        pass
    try:
        bb0 = rebar.get_BoundingBox(None)
        if bb0 is not None:
            return (bb0.Min + bb0.Max) * 0.5
    except Exception:
        pass
    return None


def _suple_mitad_barra_layout_xyz(layout, plane):
    """
    Punto medio de la barra resultante (layout sketch → XYZ).

    Apoyo: centro del tramo ±L. Borde: mitad entre borde−cover y +L.
    """
    if layout is None or plane is None:
        return None
    try:
        L = float(layout.get(u"L") or 0.0)
        ax, ay = layout.get(u"origin")
        dx, dy = layout.get(u"dir")
        px, py = layout.get(u"perp")
        mid_s = float(layout.get(u"rec_len") or 0.0) * 0.5
        cx = float(ax) + float(dx) * mid_s
        cy = float(ay) + float(dy) * mid_s
        if layout.get(u"inward"):
            t_out = float(layout.get(u"outer_t") or 0.0)
            if t_out < 0.0:
                t_out = 0.0
            # Mitad del segmento [−outer_t … +L] sobre perp
            t_mid = 0.5 * (L - t_out)
            mx = cx + float(px) * t_mid
            my = cy + float(py) * t_mid
        else:
            mx, my = cx, cy
        return _plane_mm_to_xyz((mx, my), plane)
    except Exception:
        return None


def _suple_punto_mitad_barra_mra(rebar, view, layout, plane):
    """
    Ancla MRA a la mitad de la barra resultante (no al extremo del bbox).

    Orden: Show Middle centerline → layout sketch → bbox del set.
    """
    try:
        p = _punto_insercion_tag_show_middle(rebar, view)
        if p is not None:
            return p
    except Exception:
        pass
    p_lay = _suple_mitad_barra_layout_xyz(layout, plane)
    if p_lay is not None:
        try:
            return _proyectar_punto_plano_vista(p_lay, view)
        except Exception:
            return p_lay
    return _suple_centro_rebar_mra(rebar, view)


def _suple_crear_mra_uno(doc, view, rebar, layout, plane, mrat_type, avisos):
    """
    MRA «Recorrido Barras» con DimensionLineDirection = distribución (recorrido).

    Tag / origen de cota en la mitad de la barra resultante (Show Middle).
    """
    if avisos is None:
        avisos = []
    if doc is None or view is None or rebar is None or mrat_type is None:
        return False
    try:
        from Autodesk.Revit.DB import (
            DimensionStyleType,
            ElementId,
            MultiReferenceAnnotation,
            MultiReferenceAnnotationOptions,
            UnitTypeId,
            UnitUtils,
        )
        from System.Collections.Generic import List as ClrList
    except Exception as ex:
        avisos.append(u"MRA: imports fallaron ({0}).".format(_as_unicode(ex)))
        return False

    try:
        rid = int(rebar.Id.IntegerValue)
    except Exception:
        rid = 0

    p_mid = _suple_punto_mitad_barra_mra(rebar, view, layout, plane)
    if p_mid is None:
        avisos.append(u"MRA Id {0}: sin mitad de barra.".format(rid))
        return False
    try:
        vd = view.ViewDirection.Normalize()
    except Exception:
        avisos.append(u"MRA Id {0}: ViewDirection inválida.".format(rid))
        return False

    spacing_dir = _suple_spacing_dir_xyz(rebar, layout, plane, view)
    if spacing_dir is None:
        avisos.append(u"MRA Id {0}: sin dirección de distribución.".format(rid))
        return False

    # Perpendicular en plano de vista (lado del array / eje de barra)
    try:
        perp_dir = spacing_dir.CrossProduct(vd)
        if perp_dir.GetLength() < 1e-9:
            perp_dir = view.UpDirection.Normalize()
        else:
            perp_dir = perp_dir.Normalize()
    except Exception:
        try:
            perp_dir = view.UpDirection.Normalize()
        except Exception:
            avisos.append(u"MRA Id {0}: sin dirección de offset.".format(rid))
            return False

    # Solo margen pequeño: no sumar ½ longitud de barra (eso iba al extremo)
    try:
        off_ft = UnitUtils.ConvertToInternalUnits(
            float(_MRA_SUPLE_OFFSET_EXTRA_MM), UnitTypeId.Millimeters
        )
    except Exception:
        off_ft = 300.0 / 304.8

    # Preferir lado “afuera” del borde si hay layout.perp
    side = 1.0
    if layout is not None and plane is not None and layout.get(u"inward"):
        try:
            px, py = layout.get(u"perp")
            # outward = −perp (hacia borde losa)
            out3 = _plane_mm_dir_to_xyz(-float(px), -float(py), plane)
            if out3 is not None:
                out_v = _suple_project_dir_on_view(out3, view)
                if out_v is not None and float(out_v.DotProduct(perp_dir)) < 0.0:
                    side = -1.0
        except Exception:
            pass

    try:
        p_line = p_mid + perp_dir.Multiply(float(off_ft) * side)
    except Exception:
        p_line = p_mid

    try:
        opts = MultiReferenceAnnotationOptions(mrat_type)
    except Exception:
        try:
            opts = MultiReferenceAnnotationOptions()
            opts.MultiReferenceAnnotationType = mrat_type.Id
        except Exception as ex:
            avisos.append(
                u"MRA Id {0}: options ({1}).".format(rid, _as_unicode(ex))
            )
            return False
    try:
        opts.DimensionStyleType = DimensionStyleType.Linear
    except Exception:
        pass
    try:
        opts.DimensionPlaneNormal = vd
        opts.DimensionLineDirection = spacing_dir
        opts.DimensionLineOrigin = p_line
        opts.TagHeadPosition = p_line
        opts.TagHasLeader = False
    except Exception as ex:
        avisos.append(
            u"MRA Id {0}: config ({1}).".format(rid, _as_unicode(ex))
        )
        return False

    ids = ClrList[ElementId]()
    ids.Add(rebar.Id)
    try:
        opts.SetElementsToDimension(ids)
    except Exception as ex:
        avisos.append(
            u"MRA Id {0}: SetElements ({1}).".format(rid, _as_unicode(ex))
        )
        return False
    try:
        if hasattr(opts, u"ElementsMatchReferenceCategory"):
            if not opts.ElementsMatchReferenceCategory(doc):
                avisos.append(
                    u"MRA Id {0}: no válido para el tipo «{1}».".format(
                        rid, _MRA_TYPE_NAME_RECORRIDO_BARRAS
                    )
                )
                return False
    except Exception:
        pass
    try:
        mra = MultiReferenceAnnotation.Create(doc, view.Id, opts)
        if mra is None:
            avisos.append(u"MRA Id {0}: Create retornó None.".format(rid))
            return False
    except Exception as ex:
        avisos.append(
            u"MRA Id {0}: Create falló ({1}).".format(rid, _as_unicode(ex))
        )
        return False
    return True


def _suple_crear_mras_con_layout(doc, view, plane, jobs, avisos):
    """
    ``jobs``: lista de ``(rebar, layout)``.
    Returns cantidad de MRA creados.
    """
    if avisos is None:
        avisos = []
    if doc is None or view is None or not jobs:
        return 0
    if _vista_permite_multi_rebar_annotation is not None:
        try:
            if not _vista_permite_multi_rebar_annotation(view):
                avisos.append(
                    u"MRA: use planta/alzado/sección (no plantilla ni 3D)."
                )
                return 0
        except Exception:
            pass
    mrat_type = None
    if _multi_reference_annotation_type_by_name is not None:
        try:
            mrat_type = _multi_reference_annotation_type_by_name(
                doc, _MRA_TYPE_NAME_RECORRIDO_BARRAS
            )
        except Exception:
            mrat_type = None
    if mrat_type is None:
        avisos.append(
            u"MRA: no existe el tipo «{0}» en el proyecto.".format(
                _MRA_TYPE_NAME_RECORRIDO_BARRAS
            )
        )
        return 0
    n_ok = 0
    for rebar, layout in jobs:
        if rebar is None:
            continue
        try:
            if _suple_crear_mra_uno(
                doc, view, rebar, layout, plane, mrat_type, avisos
            ):
                n_ok += 1
        except Exception as ex:
            avisos.append(u"MRA: {0}".format(_as_unicode(ex)))
    return int(n_ok)


# ---------------------------------------------------------------------------
# ProgressBar (pyRevit — mismo patrón Area Rein. losa / Armado vigas)
# ---------------------------------------------------------------------------


def _pbar_enabled():
    try:
        from pyrevit import forms as _forms  # noqa: F401
    except Exception:
        return False
    return True


class _SuplesLosaCrearProgress(object):
    """Context manager no-op si ``pyrevit.forms.ProgressBar`` no está disponible."""

    def __init__(self, total, title_prefix=None):
        self._total = max(1, int(total or 1))
        self._pb = None
        self._open = False
        self._title_prefix = title_prefix or _DIALOG_TITLE

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

    def _title(self, current):
        cur = max(0, int(current))
        if cur < 1:
            return u"{0} — Colocando 0/{1}…".format(
                self._title_prefix, int(self._total)
            )
        return u"{0} — Colocando {1}/{2}…".format(
            self._title_prefix, cur, int(self._total)
        )

    def update(self, current, label=None):
        if self._pb is None:
            return
        c = max(1, min(int(current), int(self._total)))
        base = self._title(c)
        if label:
            base = u"{0} ({1})".format(base, _as_unicode(label))
        try:
            if hasattr(self._pb, u"update_progress"):
                try:
                    self._pb.update_progress(c, max_value=self._total)
                except TypeError:
                    try:
                        self._pb.update_progress(c, max=self._total)
                    except Exception:
                        pass
        except Exception:
            pass
        try:
            self._pb.title = base
        except Exception:
            pass


# ---------------------------------------------------------------------------
# ExternalEvent — Crear
# ---------------------------------------------------------------------------


class _CrearHandler(IExternalEventHandler):
    def __init__(self, ctrl_ref):
        self._ctrl_ref = ctrl_ref
        # Strong refs while ExternalEvent is queued / running (weakref alone
        # is not enough after the WPF window closes).
        self._pending_ctrl = None

    def GetName(self):
        return u"AraincoSuplesLosaCrear"

    def Execute(self, app):
        ctrl = self._pending_ctrl
        if ctrl is None and self._ctrl_ref:
            ctrl = self._ctrl_ref()
        try:
            if ctrl is not None:
                ctrl._execute_crear()
        except Exception as ex:
            uiapp = getattr(ctrl, u"_uiapp", None) if ctrl else None
            if uiapp is not None:
                _mostrar_aviso(
                    uiapp, u"Error al crear suples.", content=_as_unicode(ex)
                )
        finally:
            self._pending_ctrl = None
            if ctrl is not None:
                try:
                    ctrl._crear_pending = False
                    ctrl._dispose_crear_event()
                except Exception:
                    pass


# ---------------------------------------------------------------------------
# UI / Controller
# ---------------------------------------------------------------------------


class SuplesLosaController(object):
    def __init__(self, uiapp, uidoc, doc, floor, curves_outer, loops_all, plane):
        self._uiapp = uiapp
        self._uidoc = uidoc
        self._doc = doc
        self._floor = floor
        self._curves = list(curves_outer)
        self._loops = loops_all or [curves_outer]
        self._plane = plane
        # Vista donde se lanzó la herramienta (Unobscured al crear)
        self._host_view = None
        try:
            self._host_view = uidoc.ActiveView if uidoc is not None else None
        except Exception:
            self._host_view = None
        self._face_ui = {}
        self._active_face = u"superior"
        self._win = None
        self._ui_cv_plan = None
        self._ui_txt_canvas_header = None
        self._loop_polylines_mm = []
        self._overlays = []
        self._existing_ars = []
        self._sketch_holes = max(0, len(self._loops) - 1)
        self._ctx_geo_cache = None
        self._view_xform = None
        self._view_zoom = 1.0
        self._view_pan_x = 0.0
        self._view_pan_y = 0.0
        self._panning = False
        self._pan_last_x = 0.0
        self._pan_last_y = 0.0
        self._scene_layer = None
        self._hud_layer = None
        self._scene_base = None
        self._scene_matrix_transform = None
        self._last_canvas_cw = 0.0
        self._last_canvas_ch = 0.0
        self._size_redraw_timer = None
        self._size_redraw_pending = False
        self._view_redraw_timer = None
        self._view_redraw_pending = False
        self._snap_verts = []
        self._snap_segs = []
        self._snap_cell_index = None
        self._snap_geo_dirty = True
        self._hover_snap = None
        self._last_snap = None
        self._pick_pt1 = None
        self._draw_mode = _DRAW_IDLE
        self._active_sup_type = _SUP_TYPE_APOYO
        self._sup_type_syncing = False
        self._entre_enabled = True
        self._borde_enabled = False
        self._pano1 = None
        self._pano2 = None
        self._recorrido = None  # ((x,y),(x,y)) draft actual
        self._L_mm = None
        self._suples = []  # apoyo: definidos pendientes de crear
        self._suple_seq = 0
        self._borde_suples = []  # borde: definidos (poly + recorrido)
        self._borde_seq = 0
        self._borde_poly = None  # borrador: {pts, lm_mm, L_mm}
        self._borde_recorrido = None  # borrador: ((x,y),(x,y))
        self._entre_ui = {}
        self._borde_ui = {}
        self._crear_kind = None
        self._bar_types = _bar_types_sorted(doc) or []

        overlays, _w, _b = recolectar_contexto_planta(doc, floor, plane)
        self._overlays = overlays or []
        self._existing_ars = collect_existing_area_rein_on_floor(doc, floor, plane) or []
        self._build_polylines_cache()

        self._win = XamlReader.Parse(_XAML)
        self._wire()
        self._build_face_panels()
        self._handler = _CrearHandler(weakref.ref(self))
        self._crear_event = ExternalEvent.Create(self._handler)
        self._crear_pending = False
        self._set_draw_mode(_DRAW_PANO1)
        nw, nb, npas = _count_ctx(self._overlays)
        self._set_status(
            u"Floor Id {0} · dibuje Paño 1 (2 clics en esquinas opuestas).".format(
                _element_id_int(floor.Id)
            )
        )

        def _on_closed(sender, args):
            _unregister_singleton()
            self._win = None
            self._ui_cv_plan = None
            self._ui_txt_canvas_header = None
            # Creación pendiente: no Dispose del ExternalEvent hasta Execute.
            if getattr(self, u"_crear_pending", False):
                return
            self._dispose_crear_event()

        self._win.Closed += EventHandler(_on_closed)

    # ---- setup --------------------------------------------------------------

    def _build_polylines_cache(self):
        self._loop_polylines_mm = []
        for loop in self._loops:
            pts = _loop_to_polyline_mm(loop, self._plane)
            if pts and len(pts) >= 2:
                self._loop_polylines_mm.append(pts)

    def _ensure_ctx_geo_cache(self):
        cache = getattr(self, u"_ctx_geo_cache", None)
        if cache is not None:
            return cache
        wall_geo, beam_geo, wall_pts, beam_pts = _build_wall_beam_geo_mm(
            self._overlays
        )
        cache = {
            u"wall_geo": wall_geo,
            u"beam_geo": beam_geo,
            u"wall_pts": wall_pts,
            u"beam_pts": beam_pts,
        }
        self._ctx_geo_cache = cache
        return cache

    def _get_cv_plan(self):
        if self._ui_cv_plan is not None:
            return self._ui_cv_plan
        if self._win is not None:
            self._ui_cv_plan = self._win.FindName(u"CvPlan")
        return self._ui_cv_plan

    def _wire(self):
        btn_cancel = self._win.FindName(u"BtnCancelar")
        btn_crear = self._win.FindName(u"BtnCrear")
        cv = self._win.FindName(u"CvPlan")
        self._ui_cv_plan = cv
        self._ui_txt_canvas_header = self._win.FindName(u"TxtCanvasHeader")

        if btn_cancel is not None:
            btn_cancel.Click += RoutedEventHandler(self._on_cancel)
        if btn_crear is not None:
            btn_crear.Click += RoutedEventHandler(self._on_crear_click)
        if cv is not None:
            # Sin Background el Canvas no recibe hit-test en zonas vacías
            # (mismo criterio que Area Rein. Losa Sketch).
            try:
                cv.Background = SolidColorBrush(Color.FromArgb(0, 0, 0, 0))
            except Exception:
                pass
            cv.SizeChanged += SizeChangedEventHandler(self._on_canvas_size)
            cv.MouseWheel += MouseWheelEventHandler(self._on_canvas_wheel)
            cv.MouseDown += MouseButtonEventHandler(self._on_canvas_mouse_down)
            cv.MouseUp += MouseButtonEventHandler(self._on_canvas_mouse_up)
            cv.MouseMove += MouseEventHandler(self._on_canvas_mouse_move)
            try:
                cv.PreviewMouseLeftButtonDown += MouseButtonEventHandler(
                    self._on_canvas_click
                )
            except Exception:
                cv.MouseLeftButtonDown += MouseButtonEventHandler(self._on_canvas_click)
            try:
                cv.KeyDown += KeyEventHandler(self._on_canvas_key)
            except Exception:
                pass
            try:
                cv.Cursor = Cursors.Cross
            except Exception:
                pass

    def _on_cancel(self, sender, args):
        try:
            self._win.Close()
        except Exception:
            pass

    def _set_status(self, text):
        try:
            tb = self._win.FindName(u"TxtStatus")
            if tb is not None:
                tb.Text = _as_unicode(text)
        except Exception:
            pass

    def _update_entre_status_ui(self):
        ui = self._entre_ui or {}
        tb = ui.get(u"txt_info")
        if tb is None:
            return
        n_def = len(self._suples or [])
        parts = [u"Definidos: {0}".format(n_def)]
        if self._pano1 is not None:
            parts.append(
                u"borrador P1 lm={0:.0f}".format(
                    float(self._pano1.get(u"lm_mm") or 0.0)
                )
            )
        elif self._draw_mode != _DRAW_IDLE:
            parts.append(u"borrador…")
        try:
            tb.Text = u" · ".join(parts)
        except Exception:
            pass

    # ---- draw mode / entre losas --------------------------------------------

    def _set_draw_mode(self, mode):
        self._draw_mode = mode or _DRAW_IDLE
        self._pick_pt1 = None
        self._hover_snap = None
        self._mark_snap_geo_dirty()
        if self._entre_enabled:
            labels = {
                _DRAW_PANO1: u"Dibuje Paño 1: 2 clics (esquinas opuestas).",
                _DRAW_PANO2: u"Dibuje Paño 2: 2 clics (esquinas opuestas).",
                _DRAW_RECORRIDO: u"Dibuje el recorrido: 2 clics (línea).",
                _DRAW_IDLE: u"Listo. Ajuste Ø/Esp. y pulse Colocar armadura.",
            }
            self._set_status(labels.get(self._draw_mode, labels[_DRAW_IDLE]))
        elif self._borde_enabled:
            labels_b = {
                _DRAW_BORDE_POLY: (
                    u"Suple en Borde: dibuje el polígono (2 clics, esquinas opuestas)."
                ),
                _DRAW_RECORRIDO: (
                    u"Suple en Borde: dibuje el recorrido (2 clics)."
                ),
                _DRAW_IDLE: (
                    u"Listo. Nuevo polígono o pulse Colocar armadura."
                ),
            }
            self._set_status(
                labels_b.get(self._draw_mode, labels_b[_DRAW_IDLE])
            )
        self._update_entre_status_ui()
        self._update_borde_status_ui()
        self._sync_draw_buttons()
        try:
            self._redraw_canvas()
        except Exception:
            pass

    def _sync_draw_buttons(self):
        ui = self._entre_ui or {}
        for key, mode in (
            (u"btn_p1", _DRAW_PANO1),
            (u"btn_p2", _DRAW_PANO2),
            (u"btn_rec", _DRAW_RECORRIDO),
        ):
            btn = ui.get(key)
            if btn is None:
                continue
            try:
                active = self._draw_mode == mode
                btn.Opacity = 1.0 if active else 0.75
                btn.FontWeight = FontWeights.SemiBold if active else FontWeights.Normal
            except Exception:
                pass
        bui = self._borde_ui or {}
        for key, mode in (
            (u"btn_poly", _DRAW_BORDE_POLY),
            (u"btn_rec", _DRAW_RECORRIDO),
        ):
            btn_b = bui.get(key)
            if btn_b is None:
                continue
            try:
                active_b = self._borde_enabled and self._draw_mode == mode
                btn_b.Opacity = 1.0 if active_b else 0.75
                btn_b.FontWeight = (
                    FontWeights.SemiBold if active_b else FontWeights.Normal
                )
            except Exception:
                pass

    def _mark_snap_geo_dirty(self):
        self._snap_geo_dirty = True

    def _recompute_L(self):
        self._L_mm = None
        if self._pano1 is None or self._pano2 is None:
            return
        lm1 = self._pano1.get(u"lm_mm")
        lm2 = self._pano2.get(u"lm_mm")
        if lm1 is None or lm2 is None:
            return
        self._L_mm = max(float(lm1), float(lm2)) / 4.0

    def _layers_hint_for_active(self):
        face = getattr(self, u"_active_face", u"superior")
        if face != u"superior":
            for g in _FACE_GROUPS:
                if g[u"id"] == face:
                    return g[u"hint"]
            return u""
        tip = getattr(self, u"_active_sup_type", None)
        return _SUP_TYPE_HINTS.get(tip) or (
            u"SUP → active «Suple en apoyo» o «Suple en Borde»."
        )

    def _refresh_layers_hint(self):
        if self._win is None:
            return
        hint = self._win.FindName(u"TxtLayersHint")
        if hint is None:
            return
        try:
            hint.Text = self._layers_hint_for_active()
        except Exception:
            pass

    def _apply_type_panel_state(self, ui, enabled):
        """Sync toggle visual + body enable/opacity for a tipo panel."""
        if not ui:
            return
        chk = ui.get(u"chk")
        if chk is not None:
            try:
                if bool(chk.IsChecked) != bool(enabled):
                    chk.IsChecked = bool(enabled)
            except Exception:
                pass
        parts = ui.get(u"toggle_parts")
        if parts:
            self._sync_face_toggle_visual(parts, bool(enabled))
        body = ui.get(u"body")
        if body is not None:
            try:
                body.IsEnabled = bool(enabled)
                body.Opacity = 1.0 if enabled else 0.45
            except Exception:
                pass

    def _set_active_sup_type(self, type_id):
        """
        Activa un tipo Superior (apoyo / borde) o ninguno (None).

        Mutuamente exclusivo: solo un tipo interactivo a la vez.
        """
        if getattr(self, u"_sup_type_syncing", False):
            return
        if type_id not in (_SUP_TYPE_APOYO, _SUP_TYPE_BORDE, None):
            return
        self._sup_type_syncing = True
        try:
            self._active_sup_type = type_id
            self._entre_enabled = type_id == _SUP_TYPE_APOYO
            self._borde_enabled = type_id == _SUP_TYPE_BORDE
            self._apply_type_panel_state(self._entre_ui, self._entre_enabled)
            self._apply_type_panel_state(self._borde_ui, self._borde_enabled)
            if self._entre_enabled:
                if self._pano1 is None:
                    self._set_draw_mode(_DRAW_PANO1)
                elif self._pano2 is None:
                    self._set_draw_mode(_DRAW_PANO2)
                elif self._recorrido is None:
                    self._set_draw_mode(_DRAW_RECORRIDO)
                else:
                    self._set_draw_mode(_DRAW_IDLE)
            elif self._borde_enabled:
                self._clear_borde_draft()
                self._set_draw_mode(_DRAW_BORDE_POLY)
            else:
                self._set_draw_mode(_DRAW_IDLE)
                self._set_status(
                    u"Ningún tipo Superior activo. Active apoyo o borde."
                )
            self._refresh_layers_hint()
        finally:
            self._sup_type_syncing = False

    def _on_entre_toggle(self, sender, args):
        if getattr(self, u"_sup_type_syncing", False):
            return
        ui = self._entre_ui or {}
        chk = ui.get(u"chk")
        try:
            on = bool(chk.IsChecked) if chk is not None else False
        except Exception:
            on = False
        if on:
            self._set_active_sup_type(_SUP_TYPE_APOYO)
        elif getattr(self, u"_active_sup_type", None) == _SUP_TYPE_APOYO:
            self._set_active_sup_type(None)

    def _on_borde_toggle(self, sender, args):
        if getattr(self, u"_sup_type_syncing", False):
            return
        ui = self._borde_ui or {}
        chk = ui.get(u"chk")
        try:
            on = bool(chk.IsChecked) if chk is not None else False
        except Exception:
            on = False
        if on:
            self._set_active_sup_type(_SUP_TYPE_BORDE)
        elif getattr(self, u"_active_sup_type", None) == _SUP_TYPE_BORDE:
            self._set_active_sup_type(None)

    def _update_borde_status_ui(self):
        ui = self._borde_ui or {}
        tb = ui.get(u"txt_info")
        if tb is None:
            return
        n_def = len(self._borde_suples or [])
        parts = [u"Definidos: {0}".format(n_def)]
        draft = self._borde_poly
        if draft is not None:
            parts.append(
                u"poly lm={0:.0f} L={1:.0f}".format(
                    float(draft.get(u"lm_mm") or 0.0),
                    float(draft.get(u"L_mm") or 0.0),
                )
            )
            if self._borde_recorrido is not None:
                parts.append(u"recorrido ok")
            else:
                parts.append(u"falta recorrido")
        else:
            parts.append(u"1 polígono → recorrido · L=¼·lm")
        try:
            tb.Text = u" · ".join(parts)
        except Exception:
            pass

    def _read_borde_diam_esp(self):
        """Ø/Esp del panel Borde (para cuando se defina la creación)."""
        ui = self._borde_ui or {}
        bar_id = None
        esp = 150
        diam_mm = 0
        cmb_d = ui.get(u"cmb_diam")
        cmb_e = ui.get(u"cmb_esp")
        try:
            it = cmb_d.SelectedItem if cmb_d is not None else None
            if it is not None:
                bar_id = it.Tag
                for dmm, _lab, bt in self._bar_types or []:
                    if bt is not None and bt.Id == bar_id:
                        diam_mm = int(dmm)
                        break
        except Exception:
            pass
        try:
            it = cmb_e.SelectedItem if cmb_e is not None else None
            if it is not None:
                esp = int(it.Tag)
        except Exception:
            pass
        bar_type = None
        if bar_id is not None:
            try:
                bar_type = self._doc.GetElement(bar_id)
            except Exception:
                bar_type = None
        if bar_type is None and self._bar_types:
            bar_type = self._bar_types[0][2]
            diam_mm = int(self._bar_types[0][0])
        return bar_type, esp, diam_mm

    def _on_borde_param_changed(self, sender, args):
        self._update_borde_status_ui()
        try:
            self._redraw_canvas()
        except Exception:
            pass

    def _slab_outer_ring_mm(self):
        """Contorno exterior de la losa (Sketch) en mm del plano."""
        loops = self._loop_polylines_mm or []
        if not loops:
            return None
        return loops[0]

    def _compute_borde_layout(self, poly=None, recorrido=None):
        """
        Layout preview/creación borde: barra ⊥ al recorrido.

        - Hacia adentro del paño: +perp·L (L = ¼·lm).
        - Hacia afuera: hasta el borde de losa menos recubrimiento
          (``_COVER_BORDE_MM``).
        """
        poly = poly if poly is not None else self._borde_poly
        rec = recorrido if recorrido is not None else self._borde_recorrido
        if poly is None or rec is None:
            return None
        pts = poly.get(u"pts") or []
        L = float(poly.get(u"L_mm") or 0.0)
        if L < 10.0 and pts:
            lm = poly.get(u"lm_mm")
            if lm is None:
                lm = luz_menor_mm_from_polygon(pts)
            if lm is not None:
                L = float(lm) / 4.0
        if L < 10.0 or not pts:
            return None
        _bt, esp_mm, diam_mm = self._read_borde_diam_esp()
        esp_mm = float(esp_mm or 150)
        if esp_mm < 1.0:
            esp_mm = 150.0
        a, b = rec[0], rec[1]
        ax, ay = float(a[0]), float(a[1])
        bx, by = float(b[0]), float(b[1])
        ud = _unit_mm(bx - ax, by - ay)
        if ud is None:
            return None
        dx, dy = ud
        # Perpendicular unitario; forzar hacia el interior del polígono
        px, py = -dy, dx
        c1 = _centroid_mm(pts)
        side = (bx - ax) * (c1[1] - ay) - (by - ay) * (c1[0] - ax)
        if side < 0:
            px, py = -px, -py
        rec_len = math.hypot(bx - ax, by - ay)
        if rec_len < 1.0:
            return None
        # Maximum Spacing exacto: rec_len = floor(L/e)·e (centrado)
        snapped = _snap_recorrido_to_spacing_mm(
            ax, ay, dx, dy, rec_len, esp_mm
        )
        if snapped is not None:
            ax, ay, rec_len, stations = snapped
        else:
            # Recorrido < e: una barra al centro
            stations = [rec_len * 0.5]
        slab = self._slab_outer_ring_mm()
        out_dir = (-px, -py)
        cover = float(_COVER_BORDE_MM)
        bars = []
        outer_ts = []
        for s in stations:
            cx = ax + dx * s
            cy = ay + dy * s
            # Extremo interior: +perp·L dentro del paño
            p_in = (cx + px * L, cy + py * L)
            # Extremo exterior: borde losa − recubrimiento (no llega al filo)
            p_out = (cx, cy)
            t_out = 0.0
            if slab is not None:
                hit, t_hit = _ray_ring_first_hit_mm(
                    (cx, cy), out_dir, slab, min_t=0.0
                )
                if hit is not None and t_hit is not None:
                    t_edge = float(t_hit)
                    t_out = t_edge - cover
                    if t_out < 0.0:
                        t_out = 0.0
                    p_out = (cx + out_dir[0] * t_out, cy + out_dir[1] * t_out)
            bars.append((p_out, p_in))
            outer_ts.append(t_out)
        # overhang en la estación media (preview)
        mid_i = int(len(stations) // 2) if stations else 0
        outer_t_mid = outer_ts[mid_i] if outer_ts else 0.0
        return {
            u"L": L,
            u"esp": esp_mm,
            u"diam_mm": diam_mm,
            u"rec_len": rec_len,
            u"n_bars": len(bars),
            u"perp": (px, py),
            u"dir": (dx, dy),
            u"origin": (ax, ay),
            u"inward": True,
            u"to_slab_edge": True,
            u"cover_mm": cover,
            u"outer_t": float(outer_t_mid),
            u"bars": bars,
        }

    def _clear_borde_draft(self):
        self._borde_poly = None
        self._borde_recorrido = None
        self._pick_pt1 = None
        self._hover_snap = None
        self._mark_snap_geo_dirty()

    def _commit_borde_draft(self):
        """Congela polígono + recorrido de borde (L = ¼ · lm)."""
        poly = self._borde_poly
        rec = self._borde_recorrido
        if poly is None or rec is None:
            return None
        pts = poly.get(u"pts") or []
        lm_mm = float(poly.get(u"lm_mm") or 0.0)
        if not pts or lm_mm < 50.0:
            return None
        L_mm = float(poly.get(u"L_mm") or (lm_mm / 4.0))
        layout = self._compute_borde_layout(poly, rec)
        if layout is None:
            return None
        bar_type, esp, diam_mm = self._read_borde_diam_esp()
        bar_id = None
        try:
            if bar_type is not None:
                bar_id = bar_type.Id
        except Exception:
            bar_id = None
        self._borde_seq = int(self._borde_seq or 0) + 1
        item = {
            u"id": self._borde_seq,
            u"label": u"B{0}".format(self._borde_seq),
            u"tipo": _SUP_TYPE_BORDE,
            u"pts": list(pts),
            u"recorrido": rec,
            u"lm_mm": float(lm_mm),
            u"L_mm": float(L_mm),
            u"esp": float(esp or 150),
            u"diam_mm": int(diam_mm or 0),
            u"bar_type_id": bar_id,
            u"layout": layout,
        }
        self._borde_suples.append(item)
        self._clear_borde_draft()
        return item

    def _start_draw_borde_poly(self, sender, args):
        if not self._borde_enabled:
            return
        self._clear_borde_draft()
        self._set_draw_mode(_DRAW_BORDE_POLY)
        n = len(self._borde_suples or [])
        if n > 0:
            self._set_status(
                u"{0} borde(s) en lista. Dibuje el nuevo polígono (2 clics).".format(n)
            )

    def _start_draw_borde_recorrido(self, sender, args):
        if not self._borde_enabled:
            return
        if self._borde_poly is None:
            _mostrar_aviso(
                self._uiapp,
                u"Defina primero el polígono.",
                content=u"Un solo polígono → luego el recorrido.",
            )
            return
        self._borde_recorrido = None
        self._set_draw_mode(_DRAW_RECORRIDO)

    def _clear_draft(self):
        """Limpia solo el borrador en curso (no borra suples ya definidos)."""
        self._pano1 = None
        self._pano2 = None
        self._recorrido = None
        self._L_mm = None
        self._pick_pt1 = None
        self._hover_snap = None
        self._mark_snap_geo_dirty()

    def _commit_draft_suple(self):
        """Congela el borrador actual en la lista de suples a modelar."""
        layout = self._compute_entre_layout()
        if layout is None:
            return None
        bar_type, esp, diam_mm = self._read_diam_esp()
        bar_id = None
        try:
            if bar_type is not None:
                bar_id = bar_type.Id
        except Exception:
            bar_id = None
        self._suple_seq = int(self._suple_seq or 0) + 1
        item = {
            u"id": self._suple_seq,
            u"label": u"S{0}".format(self._suple_seq),
            u"pano1": self._pano1,
            u"pano2": self._pano2,
            u"recorrido": self._recorrido,
            u"L_mm": float(layout.get(u"L") or 0.0),
            u"esp": float(layout.get(u"esp") or esp or 150),
            u"diam_mm": int(layout.get(u"diam_mm") or diam_mm or 0),
            u"bar_type_id": bar_id,
            u"layout": layout,
        }
        self._suples.append(item)
        self._clear_draft()
        return item

    def _start_draw_pano1(self, sender, args):
        if not self._entre_enabled:
            return
        # Nuevo suple: no toca los ya definidos en self._suples
        self._clear_draft()
        self._set_draw_mode(_DRAW_PANO1)
        n = len(self._suples or [])
        if n > 0:
            self._set_status(
                u"{0} suple(s) en lista. Dibuje Paño 1 del nuevo.".format(n)
            )

    def _start_draw_pano2(self, sender, args):
        if not self._entre_enabled:
            return
        if self._pano1 is None:
            _mostrar_aviso(self._uiapp, u"Defina primero el Paño 1.")
            return
        self._pano2 = None
        self._recorrido = None
        self._L_mm = None
        self._set_draw_mode(_DRAW_PANO2)

    def _start_draw_recorrido(self, sender, args):
        if not self._entre_enabled:
            return
        if self._pano1 is None or self._pano2 is None:
            _mostrar_aviso(self._uiapp, u"Defina Paño 1 y Paño 2 antes del recorrido.")
            return
        self._recorrido = None
        self._set_draw_mode(_DRAW_RECORRIDO)

    # ---- crear --------------------------------------------------------------

    def _read_diam_esp(self):
        ui = self._entre_ui or {}
        bar_id = None
        esp = 150
        diam_mm = 0
        cmb_d = ui.get(u"cmb_diam")
        cmb_e = ui.get(u"cmb_esp")
        try:
            it = cmb_d.SelectedItem if cmb_d is not None else None
            if it is not None:
                bar_id = it.Tag
                # Tag = ElementId; diameter from matching bar_types
                for dmm, _lab, bt in self._bar_types or []:
                    if bt is not None and bt.Id == bar_id:
                        diam_mm = int(dmm)
                        break
        except Exception:
            pass
        try:
            it = cmb_e.SelectedItem if cmb_e is not None else None
            if it is not None:
                esp = int(it.Tag)
        except Exception:
            pass
        bar_type = None
        if bar_id is not None:
            try:
                bar_type = self._doc.GetElement(bar_id)
            except Exception:
                bar_type = None
        if bar_type is None and self._bar_types:
            bar_type = self._bar_types[0][2]
            diam_mm = int(self._bar_types[0][0])
        return bar_type, esp, diam_mm

    def _on_entre_param_changed(self, sender, args):
        """Ø / Esp. → actualizar preview de barras en canvas."""
        try:
            self._redraw_canvas()
        except Exception:
            pass
        self._update_entre_status_ui()

    def _compute_entre_layout(self):
        """
        Geometría de preview/creación en mm del plano Sketch.

        Returns dict:
          L, esp, diam_mm, rec_len, n_bars, perp, bars=[((x0,y0),(x1,y1)), ...]
        o None si faltan datos.
        """
        if not self._entre_enabled:
            return None
        if self._pano1 is None or self._pano2 is None or self._recorrido is None:
            return None
        self._recompute_L()
        L = float(self._L_mm or 0.0)
        if L < 10.0:
            return None
        _bt, esp_mm, diam_mm = self._read_diam_esp()
        esp_mm = float(esp_mm or 150)
        if esp_mm < 1.0:
            esp_mm = 150.0
        a, b = self._recorrido[0], self._recorrido[1]
        ax, ay = float(a[0]), float(a[1])
        bx, by = float(b[0]), float(b[1])
        ud = _unit_mm(bx - ax, by - ay)
        if ud is None:
            return None
        dx, dy = ud
        px, py = -dy, dx
        c1 = _centroid_mm(self._pano1.get(u"pts") or [])
        side = (bx - ax) * (c1[1] - ay) - (by - ay) * (c1[0] - ax)
        if side < 0:
            px, py = -px, -py
        rec_len = math.hypot(bx - ax, by - ay)
        if rec_len < 1.0:
            return None
        # Maximum Spacing exacto: rec_len = floor(L/e)·e (centrado)
        snapped = _snap_recorrido_to_spacing_mm(
            ax, ay, dx, dy, rec_len, esp_mm
        )
        if snapped is not None:
            ax, ay, rec_len, stations = snapped
        else:
            stations = [rec_len * 0.5]
        bars = []
        for s in stations:
            cx = ax + dx * s
            cy = ay + dy * s
            p0 = (cx + px * L, cy + py * L)
            p1 = (cx - px * L, cy - py * L)
            bars.append((p0, p1))
        return {
            u"L": L,
            u"esp": esp_mm,
            u"diam_mm": diam_mm,
            u"rec_len": rec_len,
            u"n_bars": len(bars),
            u"perp": (px, py),
            u"dir": (dx, dy),
            u"origin": (ax, ay),
            u"bars": bars,
        }

    def _dispose_crear_event(self):
        try:
            if self._crear_event is not None:
                self._crear_event.Dispose()
        except Exception:
            pass
        self._crear_event = None

    def _on_crear_click(self, sender, args):
        if getattr(self, u"_active_face", u"superior") != u"superior":
            _mostrar_aviso(
                self._uiapp,
                u"Cambie a la tab Superior para crear.",
                content=u"Los tipos de Inferior aún no están definibles.",
            )
            return
        tip = getattr(self, u"_active_sup_type", None)
        if tip == _SUP_TYPE_BORDE:
            if (
                self._borde_poly is not None
                and self._borde_recorrido is not None
            ):
                if self._commit_borde_draft() is None:
                    _mostrar_aviso(
                        self._uiapp, u"No se pudo cerrar el borrador de borde."
                    )
                    return
            if not self._borde_suples:
                _mostrar_aviso(
                    self._uiapp,
                    u"No hay suples de borde definidos.",
                    content=u"Defina 1 polígono y luego el recorrido.",
                )
                return
            self._crear_kind = _SUP_TYPE_BORDE
        elif tip == _SUP_TYPE_APOYO and self._entre_enabled:
            if (
                self._pano1 is not None
                and self._pano2 is not None
                and self._recorrido is not None
            ):
                if self._commit_draft_suple() is None:
                    _mostrar_aviso(
                        self._uiapp, u"No se pudo cerrar el borrador actual."
                    )
                    return
            if not self._suples:
                _mostrar_aviso(
                    self._uiapp,
                    u"No hay suples definidos.",
                    content=u"Defina al menos uno: Paño 1 → Paño 2 → recorrido.",
                )
                return
            self._crear_kind = _SUP_TYPE_APOYO
        else:
            _mostrar_aviso(
                self._uiapp,
                u"Active un tipo Superior.",
                content=u"«Suple en apoyo» o «Suple en Borde».",
            )
            return
        ev = self._crear_event
        if ev is None:
            _mostrar_aviso(self._uiapp, u"ExternalEvent no disponible.")
            return
        try:
            self._crear_pending = True
            self._handler._pending_ctrl = self
            ev.Raise()
        except Exception as ex:
            self._crear_pending = False
            try:
                self._handler._pending_ctrl = None
            except Exception:
                pass
            _mostrar_aviso(
                self._uiapp,
                u"No se pudo iniciar la colocación.",
                content=_as_unicode(ex),
            )
            return
        # Cerrar UI de inmediato; Execute usa el controlador (strong ref).
        try:
            if self._win is not None:
                self._win.Close()
        except Exception:
            pass

    def _suple_layer_cfg(self, bar_type_id, esp_mm):
        """Solo Top Major activo (suple superior); Minor y Bottom OFF."""
        esp = int(esp_mm or 150)
        src = {}
        for key in _LAYER_KEYS:
            src[key] = {
                u"bar_type_id": bar_type_id,
                u"spacing_mm": esp,
            }
        return _layer_cfg_for_keys(src, [u"exterior_major"], spacing_mm=esp)

    def _create_one_suple_area(self, doc, floor, plane, item, nivel_valor=None):
        """
        Crea AR strip → Remove Area System → Rebar libres + stamps.

        Show Middle / Unobscured se aplican en lote en ``_execute_crear``
        (más rápido; mismo resultado).

        Returns (n_bars, err, free_rebars).
        """
        layout = item.get(u"layout")
        if layout is None:
            return 0, u"Sin layout.", []
        bar_type = None
        try:
            bid = item.get(u"bar_type_id")
            if bid is not None:
                bar_type = doc.GetElement(bid)
        except Exception:
            bar_type = None
        if bar_type is None:
            if item.get(u"tipo") == _SUP_TYPE_BORDE:
                bar_type, _e, _d = self._read_borde_diam_esp()
            else:
                bar_type, _e, _d = self._read_diam_esp()
        if bar_type is None:
            try:
                if self._bar_types:
                    bar_type = self._bar_types[0][2]
            except Exception:
                bar_type = None
        if bar_type is None:
            return 0, u"Sin RebarBarType.", []
        L_mm = float(item.get(u"L_mm") or layout.get(u"L") or 0.0)
        if L_mm < 10.0:
            return 0, u"L inválida.", []
        esp_mm = float(item.get(u"esp") or layout.get(u"esp") or 150)
        # Borde: regenerar layout si falta outer_t (p. ej. ítems antiguos)
        if item.get(u"tipo") == _SUP_TYPE_BORDE or layout.get(u"inward"):
            if layout.get(u"outer_t") is None or not layout.get(u"inward"):
                try:
                    lay2 = self._compute_borde_layout(
                        {
                            u"pts": item.get(u"pts") or [],
                            u"lm_mm": item.get(u"lm_mm"),
                            u"L_mm": L_mm,
                        },
                        item.get(u"recorrido"),
                    )
                    if lay2 is not None:
                        layout = lay2
                except Exception:
                    pass
        # Luz de distribución múltiplo de e → paso real = e (Maximum Spacing)
        layout = _snap_suple_layout_to_spacing(layout, esp_mm) or layout
        try:
            item[u"layout"] = layout
        except Exception:
            pass
        pts = _strip_polygon_mm_from_layout(layout)
        if not pts:
            return 0, u"No se pudo construir el contorno del strip.", []
        curves = _poly_mm_to_curves(pts, plane)
        if not curves:
            return 0, u"Contorno AR inválido.", []
        try:
            px, py = layout.get(u"perp")
            major_dir = _plane_mm_dir_to_xyz(float(px), float(py), plane)
        except Exception:
            return 0, u"Dirección Major inválida.", []
        layer_cfg = self._suple_layer_cfg(bar_type.Id, esp_mm)
        area_rein, create_err = crear_area_reinforcement(
            doc, floor, curves, major_dir, layer_cfg
        )
        if area_rein is None:
            return 0, create_err or u"AreaReinforcement.Create falló.", []

        err = None
        # Un Regenerate por AR: necesario antes de RemoveAreaSystem
        try:
            doc.Regenerate()
        except Exception:
            pass

        gid = None
        if generar_armadura_conjunto_guid is not None:
            try:
                gid = generar_armadura_conjunto_guid()
            except Exception:
                gid = None
        if nivel_valor is None:
            try:
                nivel_valor = _nivel_losa_como_string(doc, floor)
            except Exception:
                nivel_valor = None

        # Remove Area System → Rebar libres (stamp solo en libres; los de
        # sistema se descartan con el remove).
        free_rebars = _remove_area_reinforcement_system(doc, area_rein)
        if not free_rebars:
            return 0, (err or u"") + (
                u" RemoveAreaReinforcementSystem no devolvió Rebar."
                if not err
                else u" · RemoveAreaSystem falló."
            ), []

        try:
            self._stamp_suple_elements(
                free_rebars, conjunto_guid=gid, nivel_valor=nivel_valor
            )
        except Exception as ex_stamp2:
            stamp_msg = u"Params libres: {0}".format(_as_unicode(ex_stamp2))
            err = (err + u" · " + stamp_msg) if err else stamp_msg

        n = 0
        for rb in free_rebars:
            try:
                n += int(rb.Quantity)
            except Exception:
                try:
                    n += int(rb.NumberOfBarPositions)
                except Exception:
                    n += 1
        if n <= 0:
            n = len(free_rebars) or 1
        return n, err, free_rebars

    def _stamp_suple_element(self, element, conjunto_guid=None, nivel_valor=None):
        """Stamp Arainco / Malla=No / F' / Nivel / GUID en un elemento."""
        if element is None:
            return
        try:
            if stamp_armadura_arainco is not None:
                stamp_armadura_arainco(element, yes=True)
        except Exception:
            pass
        try:
            if stamp_armadura_malla is not None:
                stamp_armadura_malla(element, yes=False)
        except Exception:
            pass
        try:
            if stamp_armadura_ubicacion is not None:
                stamp_armadura_ubicacion(
                    element, ARMADURA_UBICACION_SUPERIOR or u"F'"
                )
        except Exception:
            pass
        try:
            if stamp_armadura_nivel is not None and nivel_valor:
                stamp_armadura_nivel(element, nivel_valor)
        except Exception:
            pass
        try:
            gid = conjunto_guid
            if not gid and generar_armadura_conjunto_guid is not None:
                gid = generar_armadura_conjunto_guid()
            if stamp_armadura_conjunto_guid is not None and gid:
                stamp_armadura_conjunto_guid(element, conjunto_guid=gid)
        except Exception:
            pass

    def _stamp_suple_elements(self, elements, conjunto_guid=None, nivel_valor=None):
        """Stamp en una lista de elementos (mismo GUID / nivel)."""
        for el in elements or []:
            self._stamp_suple_element(
                el, conjunto_guid=conjunto_guid, nivel_valor=nivel_valor
            )

    def _etiquetar_rebars_show_middle(
        self, doc, view, rebars, avisos, tag_map=None, prepare_hosts=True
    ):
        """
        Etiqueta Rebar libres (post RemoveAreaSystem) en la vista activa.

        Familia ``TAG_FLOOR`` (fallback WALL_HORIZONTAL); tipo = RebarShape.
        ``prepare_hosts=False`` si ya se llamó ``_suple_prepare_rebars_for_annotation``.
        """
        if avisos is None:
            avisos = []
        if doc is None or not rebars:
            return 0

        view = _resolver_vista_para_show_middle(doc, view)
        if view is None:
            avisos.append(u"Etiqueta rebar: no hay vista de modelo válida.")
            return 0
        ok_view, msg_view = _vista_ok_para_etiquetas_rebar(view)
        if not ok_view:
            avisos.append(msg_view or u"Etiqueta rebar: vista no válida.")
            return 0

        try:
            from enfierrado_shaft_hashtag import _collect_rebar_tag_symbol_map
        except Exception as ex:
            avisos.append(
                u"Etiqueta rebar: no se pudo cargar helper ({0}).".format(
                    _as_unicode(ex)
                )
            )
            return 0

        family_used = _SUPLE_REBAR_TAG_FAMILY_NAME
        if tag_map is None:
            try:
                tag_map = _collect_rebar_tag_symbol_map(doc, family_used)
            except Exception as ex_map:
                avisos.append(
                    u"Etiqueta rebar: fallo al leer tipos ({0}).".format(
                        _as_unicode(ex_map)
                    )
                )
                return 0
            if not tag_map:
                family_used = _SUPLE_REBAR_TAG_FAMILY_FALLBACK
                try:
                    tag_map = _collect_rebar_tag_symbol_map(doc, family_used)
                except Exception:
                    tag_map = None
        if not tag_map:
            avisos.append(
                u"Etiqueta rebar: no hay tipos de «{0}» ni «{1}».".format(
                    _SUPLE_REBAR_TAG_FAMILY_NAME,
                    _SUPLE_REBAR_TAG_FAMILY_FALLBACK,
                )
            )
            return 0

        _suple_ensure_rebar_tag_visibility(doc, view)
        if prepare_hosts:
            # Unobscured + Show Middle + 1 Regenerate (no N regenerates por barra)
            _suple_prepare_rebars_for_annotation(doc, view, rebars)

        n_ok = 0
        n_fail = 0
        for barra in rebars:
            if not isinstance(barra, (Rebar, RebarInSystem)):
                try:
                    barra = doc.GetElement(barra.Id)
                except Exception:
                    barra = None
            if not isinstance(barra, (Rebar, RebarInSystem)):
                n_fail += 1
                continue

            tag_type_id = _resolver_tag_type_id_por_shape(
                doc, barra, tag_map, _SUPLE_REBAR_TAG_FALLBACK_TYPE
            )
            if tag_type_id is None:
                n_fail += 1
                continue

            p = None
            try:
                p = _punto_insercion_tag_show_middle(barra, view)
            except Exception:
                p = None
            if p is not None:
                try:
                    p2 = _proyectar_punto_plano_vista(p, view)
                    if p2 is not None:
                        p = p2
                except Exception:
                    pass
            if p is None:
                n_fail += 1
                continue

            created = None
            try:
                created = _suple_crear_tag_rebar_libre(
                    doc, view, barra, tag_type_id, p
                )
            except Exception:
                created = None
            if created is None:
                try:
                    created = _crear_independent_tag_rebar(
                        doc, view, barra, tag_type_id, p
                    )
                    if created is not None:
                        _aplicar_estilo_tag_rebar_sin_leader(created, p)
                except Exception:
                    created = None

            if created is not None:
                n_ok += 1
            else:
                n_fail += 1

        if n_ok == 0 and n_fail > 0:
            avisos.append(
                u"Etiqueta rebar: no se pudo etiquetar {0} set(s) "
                u"(familia «{1}»)."
                .format(n_fail, family_used)
            )
        return n_ok

    def _resolve_tag_view(self, doc):
        """
        Solo la vista activa de UI (``uidoc.ActiveView`` / ``doc.ActiveView``).
        Las etiquetas deben crearse ahí — sin caer a host_view de lanzamiento.
        """
        view = None
        try:
            if self._uidoc is not None:
                view = self._uidoc.ActiveView
        except Exception:
            view = None
        if view is None and doc is not None:
            try:
                view = doc.ActiveView
            except Exception:
                view = None
        resolved = _resolver_vista_para_show_middle(doc, view)
        if resolved is None and view is not None:
            # Plantilla/3D: no etiquetar en host_view alternativa
            return None
        return resolved

    def _execute_crear(self):
        doc = self._doc
        floor = self._floor
        plane = self._plane
        kind = getattr(self, u"_crear_kind", None) or getattr(
            self, u"_active_sup_type", _SUP_TYPE_APOYO
        )
        if kind == _SUP_TYPE_BORDE:
            items = list(self._borde_suples or [])
            tx_name = u"Arainco: Suples losa (suple en borde)"
            done_title = u"Suples en borde creados."
            label_prefix = u"B"
        else:
            kind = _SUP_TYPE_APOYO
            items = list(self._suples or [])
            tx_name = u"Arainco: Suples losa (suple en apoyo)"
            done_title = u"Suples en apoyo creados."
            label_prefix = u"S"
        if doc is None or floor is None or plane is None or not items:
            _mostrar_aviso(self._uiapp, u"Datos incompletos para crear.")
            return

        view = self._resolve_tag_view(doc)
        nivel_valor = None
        try:
            nivel_valor = _nivel_losa_como_string(doc, floor)
        except Exception:
            nivel_valor = None
        # +1 fase final (MRA / etiquetas). ProgressBar con UI WPF ya cerrada.
        n_items = len(items)
        total_steps = max(1, n_items + 1)
        pbar = _SuplesLosaCrearProgress(total_steps)
        pbar.__enter__()

        total_bars = 0
        ok_sets = 0
        n_tags = 0
        n_mra = 0
        errors = []
        avisos_tag = []
        avisos_mra = []
        created_rebars = []
        # (ElementId, layout) representativo por suple — MRA con dir. recorrido
        annotate_jobs = []

        try:
            t = Transaction(doc, tx_name)
            t.Start()
            try:
                for i, item in enumerate(items):
                    label = item.get(u"label") or u"{0}{1}".format(
                        label_prefix, i + 1
                    )
                    pbar.update(i + 1, label)
                    n, err, free_rebars = self._create_one_suple_area(
                        doc,
                        floor,
                        plane,
                        item,
                        nivel_valor=nivel_valor,
                    )
                    if free_rebars:
                        ok_sets += 1
                        total_bars += max(0, int(n or 0))
                        created_rebars.extend(free_rebars)
                        rep = _pick_suple_annotation_rebar(
                            doc, free_rebars, item, plane, view
                        )
                        if rep is None:
                            try:
                                rep = free_rebars[0]
                            except Exception:
                                rep = None
                        if rep is not None:
                            try:
                                annotate_jobs.append(
                                    (rep.Id, item.get(u"layout"))
                                )
                            except Exception:
                                pass
                    if err:
                        errors.append(
                            u"{0}: {1}".format(label, err)
                        )
                try:
                    doc.Regenerate()
                except Exception:
                    pass
                # Vista activa fresca (puede haber cambiado) — aún sin etiquetar
                view = self._resolve_tag_view(doc)

                refreshed = []
                for rb in created_rebars:
                    try:
                        el = doc.GetElement(rb.Id)
                    except Exception:
                        el = rb
                    if isinstance(el, (Rebar, RebarInSystem)):
                        refreshed.append(el)
                created_rebars = refreshed

                # Un pase: Show Middle + unobscured + 1 Regenerate
                if created_rebars and view is not None:
                    _suple_prepare_rebars_for_annotation(
                        doc, view, created_rebars
                    )

                mra_jobs = []
                by_id = {}
                for rb in created_rebars:
                    try:
                        by_id[int(rb.Id.IntegerValue)] = rb
                    except Exception:
                        pass
                for eid, lay in annotate_jobs:
                    el = None
                    try:
                        el = by_id.get(int(eid.IntegerValue))
                    except Exception:
                        el = None
                    if el is None:
                        try:
                            el = doc.GetElement(eid)
                        except Exception:
                            el = None
                    if isinstance(el, (Rebar, RebarInSystem)):
                        mra_jobs.append((el, lay))
                if not mra_jobs and created_rebars:
                    for rb in created_rebars:
                        mra_jobs.append((rb, None))

                # MRA «Recorrido Barras» con dirección = recorrido del suple
                # (RightDirection del helper estribos falla en bordes oblicuos).
                if mra_jobs and view is not None:
                    try:
                        n_mra = int(
                            _suple_crear_mras_con_layout(
                                doc, view, plane, mra_jobs, avisos_mra
                            )
                            or 0
                        )
                    except Exception as ex_mra:
                        avisos_mra.append(_as_unicode(ex_mra))
                        n_mra = 0
                    # Fallback al helper genérico si el layout-aware falla todo
                    if (
                        n_mra <= 0
                        and crear_multi_rebar_annotations_por_nombre_tipo
                        is not None
                    ):
                        try:
                            only_bars = [jb[0] for jb in mra_jobs]
                            n_mra = int(
                                crear_multi_rebar_annotations_por_nombre_tipo(
                                    doc,
                                    view,
                                    only_bars,
                                    avisos_mra,
                                    _MRA_TYPE_NAME_RECORRIDO_BARRAS,
                                )
                                or 0
                            )
                        except Exception as ex_mra2:
                            avisos_mra.append(_as_unicode(ex_mra2))
                    if n_mra <= 0:
                        avisos_mra.append(
                            u"MRA: 0 anotaciones (tipo «{0}», vista planta/"
                            u"alzado/sección).".format(
                                _MRA_TYPE_NAME_RECORRIDO_BARRAS
                            )
                        )
                elif mra_jobs and view is None:
                    avisos_mra.append(
                        u"MRA: sin vista activa válida."
                    )

                t.Commit()
            except Exception as ex:
                try:
                    t.RollBack()
                except Exception:
                    pass
                _mostrar_aviso(
                    self._uiapp,
                    u"Error al crear Area Reinforcement.",
                    content=_as_unicode(ex),
                )
                return

            # Etiquetas en Transaction propia sobre la vista activa
            pbar.update(total_steps, u"etiquetas")
            view_tag = self._resolve_tag_view(doc)
            if created_rebars and view_tag is not None:
                t_tag = Transaction(
                    doc, u"Arainco: Suples losa — etiquetas rebar (vista activa)"
                )
                t_tag.Start()
                try:
                    # Reutilizar elementos ya refrescados si la vista no cambió
                    same_view = False
                    try:
                        same_view = (
                            view is not None
                            and view_tag is not None
                            and view.Id == view_tag.Id
                        )
                    except Exception:
                        same_view = False
                    if same_view:
                        bars_tag = list(created_rebars)
                    else:
                        bars_tag = []
                        for rb in created_rebars:
                            try:
                                el = doc.GetElement(rb.Id)
                            except Exception:
                                el = rb
                            if isinstance(el, (Rebar, RebarInSystem)):
                                bars_tag.append(el)
                    n_tags = int(
                        self._etiquetar_rebars_show_middle(
                            doc,
                            view_tag,
                            bars_tag,
                            avisos_tag,
                            prepare_hosts=not same_view,
                        )
                        or 0
                    )
                    t_tag.Commit()
                except Exception as ex_tag:
                    try:
                        if t_tag.HasStarted():
                            t_tag.RollBack()
                    except Exception:
                        pass
                    avisos_tag.append(
                        u"Etiqueta rebar: {0}".format(_as_unicode(ex_tag))
                    )
                    n_tags = 0
                try:
                    if self._uidoc is not None:
                        self._uidoc.RefreshActiveView()
                except Exception:
                    pass
            elif not created_rebars:
                avisos_tag.append(
                    u"Etiqueta rebar: no hay Rebar libres para etiquetar."
                )
            else:
                avisos_tag.append(
                    u"Etiqueta rebar: sin vista activa válida "
                    u"(abra una planta/alzado/sección y reintente)."
                )
        finally:
            try:
                pbar.__exit__(None, None, None)
            except Exception:
                pass

        if kind == _SUP_TYPE_BORDE:
            self._borde_suples = []
            self._clear_borde_draft()
        else:
            self._suples = []
            self._clear_draft()
        msg = (
            u"{0} set(s) · {1} barras (sin area system) · "
            u"{2} etiqueta(s) · {3} MRA."
        ).format(ok_sets, total_bars, n_tags, n_mra)
        if errors:
            msg = msg + u"\n" + u"\n".join(errors[:8])
        seen_av = set()
        for av in (avisos_tag or []) + (avisos_mra or []):
            if not av or av in seen_av:
                continue
            seen_av.add(av)
            if len(seen_av) > 8:
                break
            msg = msg + u"\n" + av
        _mostrar_aviso(self._uiapp, done_title, content=msg)

    # ---- canvas size / zoom / pan (compact) ---------------------------------

    def _on_canvas_size(self, sender, args):
        self._schedule_size_redraw()

    def _schedule_size_redraw(self):
        self._size_redraw_pending = True
        t = getattr(self, u"_size_redraw_timer", None)
        if t is not None:
            try:
                if not t.IsEnabled:
                    t.Start()
                return
            except Exception:
                self._size_redraw_timer = None
        try:
            t = DispatcherTimer()
            t.Interval = TimeSpan.FromMilliseconds(12)

            def _tick(sender, args):
                try:
                    sender.Stop()
                except Exception:
                    pass
                if getattr(self, u"_size_redraw_pending", False):
                    self._flush_size_redraw()

            t.Tick += _tick
            self._size_redraw_timer = t
            t.Start()
        except Exception:
            self._flush_size_redraw()

    def _flush_size_redraw(self):
        self._size_redraw_pending = False
        cv = self._get_cv_plan()
        if cv is None:
            return
        try:
            cw = float(cv.ActualWidth)
            ch = float(cv.ActualHeight)
        except Exception:
            return
        if cw < 40 or ch < 40:
            return
        if abs(cw - float(self._last_canvas_cw or 0)) < 0.5 and abs(
            ch - float(self._last_canvas_ch or 0)
        ) < 0.5:
            return
        try:
            self._redraw_canvas()
        except Exception:
            pass

    def _apply_view_to_xform(self, cw, ch):
        xf = self._view_xform
        if xf is None:
            return False
        try:
            cw = float(cw)
            ch = float(ch)
            if cw < 40 or ch < 40:
                return False
            min_x = float(xf[u"min_x"])
            max_x = float(xf[u"max_x"])
            min_y = float(xf[u"min_y"])
            max_y = float(xf[u"max_y"])
            fit_scale = float(xf.get(u"fit_scale") or 0.0)
            if fit_scale < 1e-12:
                return False
            zoom = float(self._view_zoom) if self._view_zoom else 1.0
            zoom = max(0.25, min(16.0, zoom))
            self._view_zoom = zoom
            scale = fit_scale * zoom
            bbox_cx = (min_x + max_x) / 2.0
            bbox_cy = (min_y + max_y) / 2.0
            cx_mm = bbox_cx + float(self._view_pan_x or 0.0)
            cy_mm = bbox_cy + float(self._view_pan_y or 0.0)
            xf[u"scale"] = scale
            xf[u"ox"] = cw / 2.0 - (cx_mm - min_x) * scale
            xf[u"oy"] = ch / 2.0 - (max_y - cy_mm) * scale
            xf[u"cw"] = cw
            xf[u"ch"] = ch
            return True
        except Exception:
            return False

    def _update_hud_scale_bar(self, cw=None, ch=None, scale=None):
        hud = getattr(self, u"_hud_layer", None)
        if hud is None:
            return
        xf = self._view_xform
        try:
            if cw is None:
                cw = float(xf[u"cw"]) if xf else 0.0
            if ch is None:
                ch = float(xf[u"ch"]) if xf else 0.0
            if scale is None:
                scale = float(xf[u"scale"]) if xf else 0.0
            cw, ch, scale = float(cw), float(ch), float(scale)
        except Exception:
            return
        if cw < 40 or ch < 40 or scale < 1e-12:
            return
        try:
            remove = [
                c
                for c in list(hud.Children)
                if getattr(c, u"Tag", None) == _HUD_SCALE_TAG
            ]
            for c in remove:
                hud.Children.Remove(c)
        except Exception:
            pass
        try:
            bar_mm, sx, sy = 2000.0, 12.0, ch - 18.0
            bar_px = bar_mm * scale
            sl = WpfLine()
            sl.X1, sl.Y1, sl.X2, sl.Y2 = sx, sy, sx + bar_px, sy
            sl.Stroke = _brush(u"#95B8CC")
            sl.StrokeThickness = 2
            sl.Tag = _HUD_SCALE_TAG
            sl.IsHitTestVisible = False
            hud.Children.Add(sl)
            for xx in (sx, sx + bar_px):
                t = WpfLine()
                t.X1, t.Y1, t.X2, t.Y2 = xx, sy - 4, xx, sy + 4
                t.Stroke = _brush(u"#95B8CC")
                t.StrokeThickness = 1.5
                t.Tag = _HUD_SCALE_TAG
                t.IsHitTestVisible = False
                hud.Children.Add(t)
            stb = TextBlock()
            stb.Text = u"2.00 m · {:.1f} px/m".format(scale * 1000.0)
            stb.Foreground = _brush(u"#64748b")
            stb.FontSize = 10
            stb.Tag = _HUD_SCALE_TAG
            stb.IsHitTestVisible = False
            WpfCanvas.SetLeft(stb, sx)
            WpfCanvas.SetTop(stb, sy - 18)
            hud.Children.Add(stb)
        except Exception:
            pass

    def _apply_scene_view_transform(self):
        scene = getattr(self, u"_scene_layer", None)
        base = getattr(self, u"_scene_base", None)
        if scene is None or base is None:
            return False
        cv = self._get_cv_plan()
        if cv is None:
            return False
        try:
            cw = float(cv.ActualWidth)
            ch = float(cv.ActualHeight)
        except Exception:
            return False
        if abs(float(base.get(u"cw", 0.0)) - cw) > 0.5:
            return False
        if abs(float(base.get(u"ch", 0.0)) - ch) > 0.5:
            return False
        if not self._apply_view_to_xform(cw, ch):
            return False
        xf = self._view_xform
        try:
            zoom = float(self._view_zoom) if self._view_zoom else 1.0
            ox0, oy0 = float(base[u"ox0"]), float(base[u"oy0"])
            ox, oy = float(xf[u"ox"]), float(xf[u"oy"])
            mtx = Matrix(zoom, 0.0, 0.0, zoom, ox - ox0 * zoom, oy - oy0 * zoom)
            mt = getattr(self, u"_scene_matrix_transform", None)
            if mt is None:
                mt = MatrixTransform(mtx)
                self._scene_matrix_transform = mt
                scene.RenderTransform = mt
            else:
                mt.Matrix = mtx
                if scene.RenderTransform is not mt:
                    scene.RenderTransform = mt
        except Exception:
            return False
        try:
            self._update_hud_scale_bar(cw, ch, float(xf[u"scale"]))
        except Exception:
            pass
        return True

    def _schedule_view_redraw(self):
        if self._apply_scene_view_transform():
            if self._hover_snap is not None or self._pick_pt1 is not None:
                try:
                    self._refresh_snap_overlay()
                except Exception:
                    pass
            return
        self._view_redraw_pending = True
        t = getattr(self, u"_view_redraw_timer", None)
        if t is not None:
            try:
                if not t.IsEnabled:
                    t.Start()
                return
            except Exception:
                self._view_redraw_timer = None
        try:
            t = DispatcherTimer()
            t.Interval = TimeSpan.FromMilliseconds(16)

            def _tick(sender, args):
                try:
                    sender.Stop()
                except Exception:
                    pass
                if getattr(self, u"_view_redraw_pending", False):
                    self._view_redraw_pending = False
                    try:
                        self._redraw_canvas(view_only=True)
                    except Exception:
                        pass

            t.Tick += _tick
            self._view_redraw_timer = t
            t.Start()
        except Exception:
            self._view_redraw_pending = False
            try:
                self._redraw_canvas(view_only=True)
            except Exception:
                pass

    def _on_canvas_mouse_down(self, sender, e):
        try:
            if e.ChangedButton != MouseButton.Middle:
                return
        except Exception:
            return
        cv = sender
        if cv is None or self._view_xform is None:
            return
        try:
            pos = e.GetPosition(cv)
            self._pan_last_x = float(pos.X)
            self._pan_last_y = float(pos.Y)
        except Exception:
            return
        self._panning = True
        try:
            cv.CaptureMouse()
            cv.Cursor = Cursors.SizeAll
        except Exception:
            pass
        try:
            e.Handled = True
        except Exception:
            pass

    def _on_canvas_mouse_up(self, sender, e):
        try:
            if e.ChangedButton != MouseButton.Middle:
                return
        except Exception:
            return
        self._panning = False
        cv = sender
        if cv is not None:
            try:
                if cv.IsMouseCaptured:
                    cv.ReleaseMouseCapture()
            except Exception:
                pass
            try:
                cv.Cursor = Cursors.Cross
            except Exception:
                pass
        try:
            e.Handled = True
        except Exception:
            pass

    def _on_canvas_wheel(self, sender, e):
        cv, xf = sender, self._view_xform
        if cv is None or xf is None:
            return
        try:
            delta = int(e.Delta)
        except Exception:
            return
        if delta == 0:
            return
        zoom = max(0.25, min(16.0, float(self._view_zoom or 1.0)))
        try:
            zoom_new = zoom * math.pow(1.06, float(delta) / 120.0)
        except Exception:
            zoom_new = zoom * (1.06 if delta > 0 else 1.0 / 1.06)
        zoom_new = max(0.25, min(16.0, zoom_new))
        if abs(zoom_new - zoom) < 1e-12:
            return
        try:
            pos = e.GetPosition(cv)
            px, py = float(pos.X), float(pos.Y)
            min_x, max_y = float(xf[u"min_x"]), float(xf[u"max_y"])
            scale_old = float(xf[u"scale"])
            ox, oy = float(xf[u"ox"]), float(xf[u"oy"])
            xmm = min_x + (px - ox) / scale_old
            ymm = max_y - (py - oy) / scale_old
            self._view_zoom = zoom_new
            fit_scale = float(xf[u"fit_scale"])
            scale_new = fit_scale * zoom_new
            bbox_cx = (float(xf[u"min_x"]) + float(xf[u"max_x"])) / 2.0
            bbox_cy = (float(xf[u"min_y"]) + float(xf[u"max_y"])) / 2.0
            cw, ch = float(xf[u"cw"]), float(xf[u"ch"])
            ox_new = px - (xmm - min_x) * scale_new
            oy_new = py - (max_y - ymm) * scale_new
            cx_mm = min_x + (cw / 2.0 - ox_new) / scale_new
            cy_mm = max_y - (ch / 2.0 - oy_new) / scale_new
            self._view_pan_x = cx_mm - bbox_cx
            self._view_pan_y = cy_mm - bbox_cy
        except Exception:
            self._view_zoom = zoom_new
        self._schedule_view_redraw()
        try:
            e.Handled = True
        except Exception:
            pass

    # ---- snap / click -------------------------------------------------------

    def _overlay_host(self):
        hud = getattr(self, u"_hud_layer", None)
        return hud if hud is not None else self._get_cv_plan()

    def _canvas_to_mm(self, pos):
        xf = self._view_xform
        if xf is None:
            return None
        try:
            scale = float(xf[u"scale"])
            if scale < 1e-12:
                return None
            return (
                float(xf[u"min_x"]) + (float(pos.X) - float(xf[u"ox"])) / scale,
                float(xf[u"max_y"]) - (float(pos.Y) - float(xf[u"oy"])) / scale,
            )
        except Exception:
            return None

    def _mm_to_px(self, xmm, ymm):
        xf = self._view_xform
        if xf is None:
            return None
        try:
            scale = float(xf[u"scale"])
            return (
                float(xf[u"ox"]) + (float(xmm) - float(xf[u"min_x"])) * scale,
                float(xf[u"oy"]) + (float(xf[u"max_y"]) - float(ymm)) * scale,
            )
        except Exception:
            return None

    def _rebuild_snap_geometry(self):
        verts, segs = [], []
        for pts in self._loop_polylines_mm or []:
            _append_ring_snap(verts, segs, pts, include_midpoints=True)
        for ov in self._overlays or []:
            pts = ov.get(u"pts") or []
            if ov.get(u"kind") == _CTX_GRID or ov.get(u"closed") is False:
                _append_polyline_snap(verts, segs, pts, include_midpoints=True)
            else:
                _append_ring_snap(verts, segs, pts, include_midpoints=True)
        try:
            _append_wall_beam_intersection_snap(verts, self._overlays)
        except Exception:
            pass
        for ar in self._existing_ars or []:
            for ring in ar.get(u"loops") or []:
                _append_ring_snap(verts, segs, ring, include_midpoints=True)
            if not ar.get(u"loops"):
                _append_ring_snap(verts, segs, ar.get(u"pts") or [], include_midpoints=True)
        for pano in (self._pano1, self._pano2):
            if pano is not None:
                _append_ring_snap(
                    verts, segs, pano.get(u"pts") or [], include_midpoints=True
                )
        if self._recorrido is not None:
            _append_polyline_snap(
                verts, segs, [self._recorrido[0], self._recorrido[1]], include_midpoints=True
            )
        for item in self._borde_suples or []:
            _append_ring_snap(
                verts, segs, item.get(u"pts") or [], include_midpoints=True
            )
            rec_i = item.get(u"recorrido")
            if rec_i is not None:
                _append_polyline_snap(
                    verts, segs, [rec_i[0], rec_i[1]], include_midpoints=True
                )
        if self._borde_poly is not None:
            _append_ring_snap(
                verts,
                segs,
                self._borde_poly.get(u"pts") or [],
                include_midpoints=True,
            )
        if self._borde_recorrido is not None:
            _append_polyline_snap(
                verts,
                segs,
                [self._borde_recorrido[0], self._borde_recorrido[1]],
                include_midpoints=True,
            )
        self._snap_verts = verts
        self._snap_segs = segs
        try:
            self._snap_cell_index = _build_snap_cell_index(verts, segs)
        except Exception:
            self._snap_cell_index = None

    def _ensure_snap_geometry(self):
        if not getattr(self, u"_snap_geo_dirty", True):
            return
        try:
            self._rebuild_snap_geometry()
        except Exception:
            self._snap_verts = []
            self._snap_segs = []
            self._snap_cell_index = None
        self._snap_geo_dirty = False

    def _snap_thresh_mm(self):
        xf = self._view_xform
        if xf is None:
            return 0.0
        try:
            scale = float(xf[u"scale"])
            if scale < 1e-12:
                return 0.0
            return float(_SNAP_PX) / scale
        except Exception:
            return 0.0

    def _resolve_snap(self, pt_mm):
        if pt_mm is None:
            return None, None
        self._ensure_snap_geometry()
        thresh = self._snap_thresh_mm()
        if thresh <= 0:
            return (float(pt_mm[0]), float(pt_mm[1])), None
        return _snap_point_mm(
            pt_mm,
            self._snap_verts,
            self._snap_segs,
            thresh,
            getattr(self, u"_snap_cell_index", None),
        )

    def _clear_snap_overlay(self, cv):
        if cv is None:
            return
        try:
            remove = [
                c for c in list(cv.Children) if getattr(c, u"Tag", None) == _SNAP_TAG
            ]
            for c in remove:
                cv.Children.Remove(c)
        except Exception:
            pass

    def _refresh_snap_overlay(self):
        cv = self._overlay_host()
        if cv is None or self._view_xform is None:
            return
        self._clear_snap_overlay(cv)
        hover = self._hover_snap
        # Preview paño / recorrido
        if self._pick_pt1 is not None and hover is not None:
            try:
                if self._draw_mode in (
                    _DRAW_PANO1,
                    _DRAW_PANO2,
                    _DRAW_BORDE_POLY,
                ):
                    pts = rect_from_two_points_mm(
                        self._pick_pt1, (hover[0], hover[1])
                    )
                    if pts and len(pts) >= 4:
                        poly = WpfPolygon()
                        pc = PointCollection()
                        for qx, qy in pts:
                            qpx = self._mm_to_px(qx, qy)
                            if qpx:
                                pc.Add(WpfPoint(qpx[0], qpx[1]))
                        if pc.Count >= 3:
                            poly.Points = pc
                            if self._draw_mode == _DRAW_PANO1:
                                col = _PANO1_COLOR
                            elif self._draw_mode == _DRAW_PANO2:
                                col = _PANO2_COLOR
                            else:
                                col = _FACE_BORDE
                            poly.Fill = _brush(col, 40)
                            poly.Stroke = _brush(col)
                            poly.StrokeThickness = 1.4
                            poly.StrokeDashArray = DoubleCollection()
                            poly.StrokeDashArray.Add(5)
                            poly.StrokeDashArray.Add(3)
                            poly.Tag = _SNAP_TAG
                            poly.IsHitTestVisible = False
                            cv.Children.Add(poly)
                elif self._draw_mode == _DRAW_RECORRIDO:
                    p0 = self._mm_to_px(self._pick_pt1[0], self._pick_pt1[1])
                    p1 = self._mm_to_px(hover[0], hover[1])
                    if p0 and p1:
                        ln = WpfLine()
                        ln.X1, ln.Y1, ln.X2, ln.Y2 = p0[0], p0[1], p1[0], p1[1]
                        ln.Stroke = _brush(_RECORRIDO_COLOR)
                        ln.StrokeThickness = 2.0
                        ln.StrokeDashArray = DoubleCollection()
                        ln.StrokeDashArray.Add(6)
                        ln.StrokeDashArray.Add(3)
                        ln.Tag = _SNAP_TAG
                        ln.IsHitTestVisible = False
                        cv.Children.Add(ln)
                    # Borde: preview barra ⊥ al recorrido (usa paño + dirección)
                    if self._borde_enabled and self._borde_poly is not None:
                        try:
                            lay = self._compute_borde_layout(
                                self._borde_poly,
                                (
                                    self._pick_pt1,
                                    (hover[0], hover[1]),
                                ),
                            )
                            ends = self._mid_bar_ends_mm(lay) if lay else None
                            if ends is not None:
                                q0, q1, _mid, _L = ends
                                s0 = self._mm_to_px(q0[0], q0[1])
                                s1 = self._mm_to_px(q1[0], q1[1])
                                if s0 and s1:
                                    bln = WpfLine()
                                    bln.X1, bln.Y1 = s0[0], s0[1]
                                    bln.X2, bln.Y2 = s1[0], s1[1]
                                    bln.Stroke = _brush(_BAR_PREVIEW_COLOR, 220)
                                    bln.StrokeThickness = 2.0
                                    bln.StrokeDashArray = DoubleCollection()
                                    bln.StrokeDashArray.Add(4)
                                    bln.StrokeDashArray.Add(3)
                                    bln.Tag = _SNAP_TAG
                                    bln.IsHitTestVisible = False
                                    cv.Children.Add(bln)
                        except Exception:
                            pass
            except Exception:
                pass
        if hover is None:
            return
        try:
            xmm, ymm, kind = hover[0], hover[1], hover[2]
        except Exception:
            return
        pxpy = self._mm_to_px(xmm, ymm)
        if pxpy is None:
            return
        px, py = pxpy
        try:
            if kind == u"vertex":
                r = 6.0
                el = WpfEllipse()
                el.Width = el.Height = r * 2.0
                el.Fill = _brush(u"#fbbf24", 200)
                el.Stroke = _brush(u"#E8F4F8")
                el.StrokeThickness = 1.5
                el.Tag = _SNAP_TAG
                el.IsHitTestVisible = False
                WpfCanvas.SetLeft(el, px - r)
                WpfCanvas.SetTop(el, py - r)
                cv.Children.Add(el)
            elif kind == u"edge":
                s = 9.0
                rect = WpfRectangle()
                rect.Width = rect.Height = s
                rect.Fill = _brush(u"#4ade80", 180)
                rect.Stroke = _brush(u"#E8F4F8")
                rect.StrokeThickness = 1.2
                rect.Tag = _SNAP_TAG
                rect.IsHitTestVisible = False
                WpfCanvas.SetLeft(rect, px - s * 0.5)
                WpfCanvas.SetTop(rect, py - s * 0.5)
                cv.Children.Add(rect)
            accent = u"#fbbf24" if kind == u"vertex" else u"#4ade80"
            for dx0, dy0, dx1, dy1 in ((-10, 0, 10, 0), (0, -10, 0, 10)):
                ln = WpfLine()
                ln.X1, ln.Y1, ln.X2, ln.Y2 = px + dx0, py + dy0, px + dx1, py + dy1
                ln.Stroke = _brush(accent)
                ln.StrokeThickness = 1.0
                ln.Tag = _SNAP_TAG
                ln.IsHitTestVisible = False
                cv.Children.Add(ln)
        except Exception:
            pass

    def _on_canvas_mouse_move(self, sender, e):
        if self._view_xform is None:
            return
        cv = sender
        if cv is None:
            return
        try:
            pos = e.GetPosition(cv)
        except Exception:
            return
        if self._panning:
            try:
                mx, my = float(pos.X), float(pos.Y)
                dx = mx - float(self._pan_last_x)
                dy = my - float(self._pan_last_y)
                self._pan_last_x, self._pan_last_y = mx, my
                scale = float(self._view_xform.get(u"scale") or 0.0)
                if scale > 1e-12:
                    self._view_pan_x = float(self._view_pan_x or 0.0) - dx / scale
                    self._view_pan_y = float(self._view_pan_y or 0.0) + dy / scale
                    self._schedule_view_redraw()
            except Exception:
                pass
            return
        raw = self._canvas_to_mm(pos)
        if raw is None:
            return
        snapped, kind = self._resolve_snap(raw)
        if snapped is None:
            return
        self._last_snap = (snapped, kind)
        show = kind is not None or self._pick_pt1 is not None
        if not show:
            if self._hover_snap is not None:
                self._hover_snap = None
                self._refresh_snap_overlay()
            return
        new_hover = (snapped[0], snapped[1], kind)
        prev = self._hover_snap
        if (
            prev is not None
            and abs(prev[0] - new_hover[0]) < 0.05
            and abs(prev[1] - new_hover[1]) < 0.05
            and prev[2] == new_hover[2]
        ):
            return
        self._hover_snap = new_hover
        self._refresh_snap_overlay()

    def _on_canvas_key(self, sender, e):
        try:
            if e.Key == Key.Escape:
                if self._pick_pt1 is not None:
                    self._pick_pt1 = None
                    self._hover_snap = None
                    self._refresh_snap_overlay()
                    self._set_status(u"Clic cancelado. Reintente.")
                    e.Handled = True
        except Exception:
            pass

    def _on_canvas_click(self, sender, e):
        if self._panning:
            return
        try:
            if e.ChangedButton != MouseButton.Left:
                return
        except Exception:
            pass
        if self._draw_mode == _DRAW_IDLE:
            return
        apoyo_draw = self._entre_enabled and self._draw_mode in (
            _DRAW_PANO1,
            _DRAW_PANO2,
            _DRAW_RECORRIDO,
        )
        borde_draw = self._borde_enabled and self._draw_mode in (
            _DRAW_BORDE_POLY,
            _DRAW_RECORRIDO,
        )
        if not apoyo_draw and not borde_draw:
            return
        cv = self._get_cv_plan()
        if cv is None or self._view_xform is None:
            return
        try:
            pos = e.GetPosition(cv)
            e.Handled = True
        except Exception:
            return
        try:
            cv.Focus()
        except Exception:
            pass
        raw = self._canvas_to_mm(pos)
        if raw is None:
            return
        pt, kind = self._resolve_snap(raw)
        if pt is None:
            return
        snap_lbl = u""
        if kind == u"vertex":
            snap_lbl = u" · snap vértice"
        elif kind == u"edge":
            snap_lbl = u" · snap arista"

        if self._borde_enabled and self._draw_mode == _DRAW_BORDE_POLY:
            if self._pick_pt1 is None:
                self._pick_pt1 = pt
                self._set_status(
                    u"Esquina A ({0:.0f}, {1:.0f}){2}. Indique esquina opuesta.".format(
                        pt[0], pt[1], snap_lbl
                    )
                )
                self._mark_snap_geo_dirty()
                return
            pts = rect_from_two_points_mm(self._pick_pt1, pt)
            self._pick_pt1 = None
            if pts is None:
                _mostrar_aviso(self._uiapp, u"Polígono demasiado pequeño.")
                return
            lm = luz_menor_mm_from_polygon(pts)
            if lm is None or lm < 50.0:
                _mostrar_aviso(self._uiapp, u"Luz menor del polígono inválida.")
                return
            # Un solo polígono por set → pasar a recorrido
            self._borde_poly = {
                u"pts": pts,
                u"lm_mm": float(lm),
                u"L_mm": float(lm) / 4.0,
            }
            self._borde_recorrido = None
            self._mark_snap_geo_dirty()
            self._update_borde_status_ui()
            self._set_draw_mode(_DRAW_RECORRIDO)
            self._set_status(
                u"Polígono ok · lm={0:.0f} · L={1:.0f} mm (¼·lm). "
                u"Dibuje el recorrido (2 clics).".format(
                    float(lm), float(lm) / 4.0
                )
            )
            return

        if self._draw_mode in (_DRAW_PANO1, _DRAW_PANO2):
            if self._pick_pt1 is None:
                self._pick_pt1 = pt
                self._set_status(
                    u"Esquina A ({0:.0f}, {1:.0f}){2}. Indique esquina opuesta.".format(
                        pt[0], pt[1], snap_lbl
                    )
                )
                self._mark_snap_geo_dirty()
                return
            pts = rect_from_two_points_mm(self._pick_pt1, pt)
            self._pick_pt1 = None
            if pts is None:
                _mostrar_aviso(self._uiapp, u"Paño demasiado pequeño.")
                return
            lm = luz_menor_mm_from_polygon(pts)
            if lm is None or lm < 50.0:
                _mostrar_aviso(self._uiapp, u"Luz menor del paño inválida.")
                return
            pano = {
                u"pts": pts,
                u"lm_mm": float(lm),
                u"label": u"P1" if self._draw_mode == _DRAW_PANO1 else u"P2",
            }
            if self._draw_mode == _DRAW_PANO1:
                self._pano1 = pano
                self._pano2 = None
                self._recorrido = None
                self._L_mm = None
                self._mark_snap_geo_dirty()
                self._set_draw_mode(_DRAW_PANO2)
            else:
                self._pano2 = pano
                self._recorrido = None
                self._recompute_L()
                self._mark_snap_geo_dirty()
                self._update_entre_status_ui()
                self._set_draw_mode(_DRAW_RECORRIDO)
                self._set_status(
                    u"L = {0:.0f} mm (¼·max lm). Dibuje el recorrido (2 clics).".format(
                        float(self._L_mm or 0.0)
                    )
                )
            return

        if self._draw_mode == _DRAW_RECORRIDO:
            if self._pick_pt1 is None:
                self._pick_pt1 = pt
                self._set_status(
                    u"Inicio recorrido ({0:.0f}, {1:.0f}){2}. Indique el fin.".format(
                        pt[0], pt[1], snap_lbl
                    )
                )
                return
            p0 = self._pick_pt1
            self._pick_pt1 = None
            if math.hypot(pt[0] - p0[0], pt[1] - p0[1]) < 50.0:
                _mostrar_aviso(self._uiapp, u"Recorrido demasiado corto.")
                return

            if self._borde_enabled:
                if self._borde_poly is None:
                    _mostrar_aviso(
                        self._uiapp,
                        u"Defina primero el polígono.",
                    )
                    self._set_draw_mode(_DRAW_BORDE_POLY)
                    return
                self._borde_recorrido = (p0, pt)
                self._mark_snap_geo_dirty()
                item = self._commit_borde_draft()
                if item is None:
                    _mostrar_aviso(
                        self._uiapp, u"No se pudo registrar el suple de borde."
                    )
                    self._set_draw_mode(_DRAW_RECORRIDO)
                    return
                n = len(self._borde_suples)
                self._update_borde_status_ui()
                self._set_draw_mode(_DRAW_BORDE_POLY)
                self._set_status(
                    u"{0} añadido (L={1:.0f} mm). Lista: {2}. "
                    u"Otro polígono o pulse Colocar armadura.".format(
                        item.get(u"label"),
                        float(item.get(u"L_mm") or 0.0),
                        n,
                    )
                )
                return

            self._recorrido = (p0, pt)
            self._recompute_L()
            self._mark_snap_geo_dirty()
            item = self._commit_draft_suple()
            if item is None:
                _mostrar_aviso(self._uiapp, u"No se pudo registrar el suple.")
                self._set_draw_mode(_DRAW_RECORRIDO)
                return
            n = len(self._suples)
            self._update_entre_status_ui()
            self._set_draw_mode(_DRAW_PANO1)
            self._set_status(
                u"{0} añadido (L={1:.0f} mm). Lista: {2}. "
                u"Dibuje Paño 1 de otro o pulse Colocar armadura.".format(
                    item.get(u"label"),
                    float(item.get(u"L_mm") or 0.0),
                    n,
                )
            )
            # _set_draw_mode ya redibuja; no repetir _redraw_canvas

    # ---- face / tipo UI -----------------------------------------------------

    def _parse_hex_rgb(self, hex_color):
        h = (_as_unicode(hex_color) or u"#5BC0DE").lstrip(u"#")
        if len(h) < 6:
            return (0x5B, 0xC0, 0xDE)
        try:
            return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
        except Exception:
            return (0x5B, 0xC0, 0xDE)

    def _sync_face_toggle_visual(self, parts, checked):
        if not parts:
            return
        on = bool(checked)
        try:
            if parts.get(u"thumb_xform") is not None:
                parts[u"thumb_xform"].X = 18.0 if on else 0.0
        except Exception:
            pass
        try:
            if on:
                ar, ag, ab = parts.get(u"accent") or (0x5B, 0xC0, 0xDE)
                if parts.get(u"track_fill") is not None:
                    parts[u"track_fill"].Color = Color.FromRgb(ar, ag, ab)
                if parts.get(u"track_border") is not None:
                    parts[u"track_border"].Color = Color.FromRgb(ar, ag, ab)
            else:
                if parts.get(u"track_fill") is not None:
                    parts[u"track_fill"].Color = Color.FromRgb(18, 38, 54)
                if parts.get(u"track_border") is not None:
                    parts[u"track_border"].Color = Color.FromRgb(33, 70, 92)
        except Exception:
            pass

    def _apply_face_toggle(self, chk, label_text, accent_hex, parts):
        host = StackPanel()
        host.Orientation = Orientation.Horizontal
        host.VerticalAlignment = VerticalAlignment.Center
        track_fill = SolidColorBrush(Color.FromRgb(18, 38, 54))
        track_border = SolidColorBrush(Color.FromRgb(33, 70, 92))
        track = Border()
        track.Width, track.Height = 36.0, 18.0
        track.CornerRadius = CornerRadius(9.0)
        track.Background, track.BorderBrush = track_fill, track_border
        track.BorderThickness = Thickness(1)
        track.Margin = Thickness(0, 0, 8, 0)
        track.ClipToBounds = True
        thumb_xform = TranslateTransform(0.0, 0.0)
        thumb = Border()
        thumb.Width = thumb.Height = 12.0
        thumb.CornerRadius = CornerRadius(6.0)
        thumb.Background = SolidColorBrush(Color.FromRgb(232, 244, 248))
        thumb.HorizontalAlignment = HorizontalAlignment.Left
        thumb.Margin = Thickness(2, 0, 0, 0)
        thumb.RenderTransform = thumb_xform
        track.Child = thumb
        host.Children.Add(track)
        lbl = TextBlock()
        lbl.Text = _as_unicode(label_text or u"")
        lbl.FontSize = 11.0
        try:
            lbl.FontWeight = FontWeights.SemiBold
        except Exception:
            pass
        lbl.Foreground = _brush(u"#E8F4F8")
        host.Children.Add(lbl)
        chk.Content = host
        ar, ag, ab = self._parse_hex_rgb(accent_hex)
        parts.clear()
        parts[u"thumb_xform"] = thumb_xform
        parts[u"track_fill"] = track_fill
        parts[u"track_border"] = track_border
        parts[u"accent"] = (ar, ag, ab)
        try:
            on = bool(chk.IsChecked)
        except Exception:
            on = False
        self._sync_face_toggle_visual(parts, on)

    def _apply_combo_style(self, cmb):
        for key in (u"ComboStretch", u"Combo"):
            try:
                cmb.Style = self._win.FindResource(key)
                return
            except Exception:
                pass

    def _build_entre_losas_panel(self, parent):
        """UI del tipo Suple en apoyo dentro del panel Superior."""
        border = Border()
        border.Background = _brush(u"#0E1B32")
        border.BorderBrush = _brush(_FACE_SUP)
        border.BorderThickness = Thickness(1)
        border.CornerRadius = CornerRadius(4)
        border.Padding = Thickness(8)
        border.Margin = Thickness(0, 4, 0, 0)

        root = StackPanel()
        root.Orientation = Orientation.Vertical

        hdr = DockPanel()
        hdr.LastChildFill = True
        hdr.Margin = Thickness(0, 0, 0, 6)
        chk = CheckBox()
        chk.IsChecked = True
        chk.Cursor = Cursors.Hand
        try:
            chk.Style = self._win.FindResource(u"BimToolsToggleMini")
        except Exception:
            pass
        parts = {}
        self._apply_face_toggle(chk, u"Suple en apoyo", _FACE_SUP, parts)
        chk.Checked += RoutedEventHandler(self._on_entre_toggle)
        chk.Unchecked += RoutedEventHandler(self._on_entre_toggle)
        hdr.Children.Add(chk)
        root.Children.Add(hdr)

        body = StackPanel()
        body.Orientation = Orientation.Vertical

        tip = TextBlock()
        tip.Text = (
            u"Barra recta (sin ganchos) · cara superior · "
            u"L = ¼·max(luz menor P1, P2)."
        )
        tip.Foreground = _brush(u"#64748b")
        tip.FontSize = 9
        tip.TextWrapping = TextWrapping.Wrap
        tip.Margin = Thickness(0, 0, 0, 6)
        body.Children.Add(tip)

        # Ø / Esp
        row = DockPanel()
        row.LastChildFill = True
        row.Margin = Thickness(0, 0, 0, 6)
        lb_o = TextBlock()
        lb_o.Text = u"Ø"
        lb_o.Foreground = _brush(u"#95B8CC")
        lb_o.Width = 14
        lb_o.VerticalAlignment = VerticalAlignment.Center
        DockPanel.SetDock(lb_o, Dock.Left)
        cmb_o = ComboBox()
        self._apply_combo_style(cmb_o)
        cmb_o.Margin = Thickness(0, 0, 8, 0)
        default_bar = None
        for dmm, lab, bt in self._bar_types:
            it = ComboBoxItem()
            it.Content = lab
            it.Tag = bt.Id
            cmb_o.Items.Add(it)
            if dmm == 12:
                default_bar = it
        if default_bar is not None:
            cmb_o.SelectedItem = default_bar
        elif cmb_o.Items.Count > 0:
            cmb_o.SelectedIndex = 0
        DockPanel.SetDock(cmb_o, Dock.Left)
        cmb_o.Width = 90

        lb_e = TextBlock()
        lb_e.Text = u"Esp."
        lb_e.Foreground = _brush(u"#95B8CC")
        lb_e.Width = 28
        lb_e.VerticalAlignment = VerticalAlignment.Center
        DockPanel.SetDock(lb_e, Dock.Left)
        cmb_e = ComboBox()
        self._apply_combo_style(cmb_e)
        for esp in _SPACING_OPTS_MM:
            it = ComboBoxItem()
            it.Content = u"{}".format(esp)
            it.Tag = esp
            cmb_e.Items.Add(it)
            if esp == 150:
                cmb_e.SelectedItem = it
        if cmb_e.SelectedIndex < 0 and cmb_e.Items.Count > 0:
            cmb_e.SelectedIndex = 0
        try:
            cmb_o.SelectionChanged += SelectionChangedEventHandler(
                self._on_entre_param_changed
            )
            cmb_e.SelectionChanged += SelectionChangedEventHandler(
                self._on_entre_param_changed
            )
        except Exception:
            pass

        # Simple horizontal via nested stack
        row_sp = StackPanel()
        row_sp.Orientation = Orientation.Horizontal
        for el in (lb_o, cmb_o, lb_e, cmb_e):
            row_sp.Children.Add(el)
        body.Children.Add(row_sp)

        txt_info = TextBlock()
        txt_info.Foreground = _brush(u"#95B8CC")
        txt_info.FontSize = 10
        txt_info.TextWrapping = TextWrapping.Wrap
        txt_info.Margin = Thickness(0, 6, 0, 6)
        txt_info.Text = u"P1 — · P2 — · L=— · recorrido —"
        body.Children.Add(txt_info)

        def _mk_btn(content, handler):
            b = Button()
            b.Content = content
            b.Style = self._win.FindResource(u"BtnSelectOutline")
            b.Margin = Thickness(0, 0, 0, 4)
            b.HorizontalAlignment = HorizontalAlignment.Stretch
            b.Click += RoutedEventHandler(handler)
            return b

        btn_p1 = _mk_btn(u"Nuevo / Paño 1", self._start_draw_pano1)
        btn_p2 = _mk_btn(u"Paño 2", self._start_draw_pano2)
        btn_rec = _mk_btn(u"Recorrido", self._start_draw_recorrido)
        body.Children.Add(btn_p1)
        body.Children.Add(btn_p2)
        body.Children.Add(btn_rec)

        root.Children.Add(body)
        border.Child = root
        parent.Children.Add(border)

        self._entre_ui = {
            u"chk": chk,
            u"body": body,
            u"toggle_parts": parts,
            u"cmb_diam": cmb_o,
            u"cmb_esp": cmb_e,
            u"txt_info": txt_info,
            u"btn_p1": btn_p1,
            u"btn_p2": btn_p2,
            u"btn_rec": btn_rec,
        }

    def _build_borde_panel(self, parent):
        """UI del tipo Suple en Borde (Superior)."""
        border = Border()
        border.Background = _brush(u"#0E1B32")
        border.BorderBrush = _brush(_FACE_BORDE)
        border.BorderThickness = Thickness(1)
        border.CornerRadius = CornerRadius(4)
        border.Padding = Thickness(8)
        border.Margin = Thickness(0, 8, 0, 0)

        root = StackPanel()
        root.Orientation = Orientation.Vertical

        hdr = DockPanel()
        hdr.LastChildFill = True
        hdr.Margin = Thickness(0, 0, 0, 6)
        chk = CheckBox()
        chk.IsChecked = False
        chk.Cursor = Cursors.Hand
        try:
            chk.Style = self._win.FindResource(u"BimToolsToggleMini")
        except Exception:
            pass
        parts = {}
        self._apply_face_toggle(chk, u"Suple en Borde", _FACE_BORDE, parts)
        chk.Checked += RoutedEventHandler(self._on_borde_toggle)
        chk.Unchecked += RoutedEventHandler(self._on_borde_toggle)
        hdr.Children.Add(chk)
        root.Children.Add(hdr)

        body = StackPanel()
        body.Orientation = Orientation.Vertical
        body.IsEnabled = False
        try:
            body.Opacity = 0.45
        except Exception:
            pass

        tip = TextBlock()
        tip.Text = (
            u"Cara superior · 1 polígono → recorrido · "
            u"barra ⊥: borde−25 mm → L adentro · AR + tags + MRA Recorrido Barras."
        )
        tip.Foreground = _brush(u"#64748b")
        tip.FontSize = 9
        tip.TextWrapping = TextWrapping.Wrap
        tip.Margin = Thickness(0, 0, 0, 6)
        body.Children.Add(tip)

        lb_o = TextBlock()
        lb_o.Text = u"Ø"
        lb_o.Foreground = _brush(u"#95B8CC")
        lb_o.Width = 14
        lb_o.VerticalAlignment = VerticalAlignment.Center
        cmb_o = ComboBox()
        self._apply_combo_style(cmb_o)
        cmb_o.Margin = Thickness(0, 0, 8, 0)
        cmb_o.Width = 90
        default_bar = None
        for dmm, lab, bt in self._bar_types:
            it = ComboBoxItem()
            it.Content = lab
            it.Tag = bt.Id
            cmb_o.Items.Add(it)
            if dmm == 12:
                default_bar = it
        if default_bar is not None:
            cmb_o.SelectedItem = default_bar
        elif cmb_o.Items.Count > 0:
            cmb_o.SelectedIndex = 0

        lb_e = TextBlock()
        lb_e.Text = u"Esp."
        lb_e.Foreground = _brush(u"#95B8CC")
        lb_e.Width = 28
        lb_e.VerticalAlignment = VerticalAlignment.Center
        cmb_e = ComboBox()
        self._apply_combo_style(cmb_e)
        for esp in _SPACING_OPTS_MM:
            it = ComboBoxItem()
            it.Content = u"{}".format(esp)
            it.Tag = esp
            cmb_e.Items.Add(it)
            if esp == 150:
                cmb_e.SelectedItem = it
        if cmb_e.SelectedIndex < 0 and cmb_e.Items.Count > 0:
            cmb_e.SelectedIndex = 0
        try:
            cmb_o.SelectionChanged += SelectionChangedEventHandler(
                self._on_borde_param_changed
            )
            cmb_e.SelectionChanged += SelectionChangedEventHandler(
                self._on_borde_param_changed
            )
        except Exception:
            pass

        row_sp = StackPanel()
        row_sp.Orientation = Orientation.Horizontal
        for el in (lb_o, cmb_o, lb_e, cmb_e):
            row_sp.Children.Add(el)
        body.Children.Add(row_sp)

        txt_info = TextBlock()
        txt_info.Foreground = _brush(u"#95B8CC")
        txt_info.FontSize = 10
        txt_info.TextWrapping = TextWrapping.Wrap
        txt_info.Margin = Thickness(0, 6, 0, 6)
        txt_info.Text = u"Definidos: 0 · 1 polígono → recorrido · L=¼·lm"
        body.Children.Add(txt_info)

        btn_poly = Button()
        btn_poly.Content = u"Nuevo / Polígono"
        btn_poly.Style = self._win.FindResource(u"BtnSelectOutline")
        btn_poly.Margin = Thickness(0, 0, 0, 4)
        btn_poly.HorizontalAlignment = HorizontalAlignment.Stretch
        btn_poly.Click += RoutedEventHandler(self._start_draw_borde_poly)
        body.Children.Add(btn_poly)

        btn_rec = Button()
        btn_rec.Content = u"Recorrido"
        btn_rec.Style = self._win.FindResource(u"BtnSelectOutline")
        btn_rec.Margin = Thickness(0, 0, 0, 4)
        btn_rec.HorizontalAlignment = HorizontalAlignment.Stretch
        btn_rec.Click += RoutedEventHandler(self._start_draw_borde_recorrido)
        body.Children.Add(btn_rec)

        root.Children.Add(body)
        border.Child = root
        parent.Children.Add(border)

        self._borde_ui = {
            u"chk": chk,
            u"body": body,
            u"toggle_parts": parts,
            u"cmb_diam": cmb_o,
            u"cmb_esp": cmb_e,
            u"txt_info": txt_info,
            u"btn_poly": btn_poly,
            u"btn_rec": btn_rec,
        }

    def _set_active_face(self, face_id):
        """Cambia tab Superior/Inferior: visual de tabs + contenido visible."""
        if face_id not in (u"superior", u"inferior"):
            return
        self._active_face = face_id
        for g in _FACE_GROUPS:
            g_id = g[u"id"]
            ui = self._face_ui.get(g_id) or {}
            on = g_id == face_id
            accent = g[u"color"]
            tab = ui.get(u"tab")
            pill = ui.get(u"pill")
            pill_tb = ui.get(u"pill_tb")
            label_tb = ui.get(u"label_tb")
            panel = ui.get(u"panel")
            if tab is not None:
                try:
                    tab.Background = _brush(u"#0E1B32" if on else u"#0a1620")
                    tab.BorderBrush = _brush(accent if on else u"#21465C")
                    tab.BorderThickness = Thickness(1, 1, 1, 2 if on else 1)
                except Exception:
                    pass
            if pill is not None:
                try:
                    pill.BorderBrush = _brush(accent if on else u"#21465C")
                except Exception:
                    pass
            if pill_tb is not None:
                try:
                    pill_tb.Foreground = _brush(accent if on else u"#64748b")
                except Exception:
                    pass
            if label_tb is not None:
                try:
                    label_tb.Foreground = _brush(u"#E8F4F8" if on else u"#95B8CC")
                    label_tb.FontWeight = (
                        FontWeights.SemiBold if on else FontWeights.Normal
                    )
                except Exception:
                    pass
            if panel is not None:
                try:
                    panel.Visibility = (
                        Visibility.Visible if on else Visibility.Collapsed
                    )
                except Exception:
                    pass
        self._refresh_layers_hint()

    def _build_face_panels(self):
        """Tabs Superior / Inferior + un solo panel de contenido activo."""
        pnl = self._win.FindName(u"PnlFaces")
        if pnl is None:
            return
        try:
            pnl.Children.Clear()
        except Exception:
            while pnl.Children.Count > 0:
                pnl.Children.RemoveAt(pnl.Children.Count - 1)
        self._face_ui = {}

        tabs_row = Grid()
        tabs_row.Margin = Thickness(0, 0, 0, 8)
        for _i in range(len(_FACE_GROUPS)):
            cd = ColumnDefinition()
            cd.Width = GridLength(1.0, GridUnitType.Star)
            tabs_row.ColumnDefinitions.Add(cd)

        content_host = StackPanel()
        content_host.Orientation = Orientation.Vertical

        for gi, group in enumerate(_FACE_GROUPS):
            g_id = group[u"id"]
            g_color = group[u"color"]

            tab = Border()
            tab.Cursor = Cursors.Hand
            tab.CornerRadius = CornerRadius(4, 4, 0, 0)
            tab.Padding = Thickness(8, 8, 8, 8)
            tab.Margin = Thickness(0, 0, 4, 0) if gi == 0 else Thickness(4, 0, 0, 0)
            tab.Background = _brush(u"#0a1620")
            tab.BorderBrush = _brush(u"#21465C")
            tab.BorderThickness = Thickness(1)
            tab_inner = StackPanel()
            tab_inner.Orientation = Orientation.Horizontal
            tab_inner.HorizontalAlignment = HorizontalAlignment.Center

            pill = Border()
            pill.Background = _brush(u"#0E1B32")
            pill.BorderBrush = _brush(u"#21465C")
            pill.BorderThickness = Thickness(1)
            pill.CornerRadius = CornerRadius(3)
            pill.Padding = Thickness(6, 2, 6, 2)
            pill.Margin = Thickness(0, 0, 6, 0)
            pill_tb = TextBlock()
            pill_tb.Text = group[u"pill"]
            pill_tb.Foreground = _brush(u"#64748b")
            pill_tb.FontSize = 10
            try:
                pill_tb.FontWeight = FontWeights.Bold
            except Exception:
                pass
            pill.Child = pill_tb

            label_tb = TextBlock()
            label_tb.Text = group[u"title"]
            label_tb.Foreground = _brush(u"#95B8CC")
            label_tb.FontSize = 12
            label_tb.VerticalAlignment = VerticalAlignment.Center

            tab_inner.Children.Add(pill)
            tab_inner.Children.Add(label_tb)
            tab.Child = tab_inner

            panel = Border()
            panel.Background = _brush(u"#0a1620")
            panel.BorderBrush = _brush(g_color)
            panel.BorderThickness = Thickness(1)
            panel.CornerRadius = CornerRadius(0, 4, 4, 4)
            panel.Padding = Thickness(8)
            panel.Visibility = Visibility.Collapsed

            tipos_host = StackPanel()
            tipos_host.Orientation = Orientation.Vertical
            if g_id == u"superior":
                self._build_entre_losas_panel(tipos_host)
                self._build_borde_panel(tipos_host)
                # Default: apoyo activo (borde apagado)
                self._set_active_sup_type(_SUP_TYPE_APOYO)
            else:
                ph = TextBlock()
                ph.Text = (
                    u"Tipos de suple inferior — por definir.\n"
                    u"Use la tab Superior (apoyo / borde)."
                )
                ph.Foreground = _brush(u"#64748b")
                ph.FontSize = 10
                ph.TextWrapping = TextWrapping.Wrap
                tipos_host.Children.Add(ph)
            panel.Child = tipos_host

            def _make_tab_handler(fid):
                def _on(s, e):
                    self._set_active_face(fid)

                return _on

            tab.MouseLeftButtonDown += MouseButtonEventHandler(
                _make_tab_handler(g_id)
            )

            Grid.SetColumn(tab, gi)
            tabs_row.Children.Add(tab)
            content_host.Children.Add(panel)
            self._face_ui[g_id] = {
                u"tab": tab,
                u"pill": pill,
                u"pill_tb": pill_tb,
                u"label_tb": label_tb,
                u"panel": panel,
                u"tipos": tipos_host,
            }

        pnl.Children.Add(tabs_row)
        pnl.Children.Add(content_host)
        self._set_active_face(getattr(self, u"_active_face", u"superior"))

    # ---- redraw -------------------------------------------------------------

    def _mid_bar_ends_mm(self, layout):
        """
        Extremos de la barra de representación (mitad del recorrido).

        - Apoyo (default): simétrica ±L (perp).
        - Borde (``inward``): (borde losa − recubrimiento) → +perp·L (adentro).
        """
        if layout is None:
            return None
        try:
            L = float(layout.get(u"L") or 0.0)
            ax, ay = layout.get(u"origin")
            dx, dy = layout.get(u"dir")
            px, py = layout.get(u"perp")
            mid_s = float(layout.get(u"rec_len") or 0.0) * 0.5
            cx = float(ax) + float(dx) * mid_s
            cy = float(ay) + float(dy) * mid_s
            if layout.get(u"inward"):
                # Afuera: hasta borde − cover (outer_t ya descuenta cover)
                t_out = float(layout.get(u"outer_t") or 0.0)
                if t_out < 0.0:
                    t_out = 0.0
                q0 = (cx - float(px) * t_out, cy - float(py) * t_out)
                q1 = (cx + float(px) * L, cy + float(py) * L)
            else:
                q0 = (cx + float(px) * L, cy + float(py) * L)
                q1 = (cx - float(px) * L, cy - float(py) * L)
            return q0, q1, (cx, cy), L
        except Exception:
            return None

    def _draw_one_mid_bar(self, scene, to_px, layout, label=None, accent=None):
        """Barra ⊥ al recorrido (mitad), largo 2L. ``accent`` = color ticks/label."""
        ends = self._mid_bar_ends_mm(layout)
        if ends is None:
            return
        q0, q1, mid, L = ends
        s0 = to_px(q0[0], q0[1])
        s1 = to_px(q1[0], q1[1])
        if s0 is None or s1 is None:
            return
        accent = accent or _FACE_SUP
        bln = WpfLine()
        bln.X1, bln.Y1, bln.X2, bln.Y2 = s0[0], s0[1], s1[0], s1[1]
        bln.Stroke = _brush(_BAR_PREVIEW_COLOR, 230)
        bln.StrokeThickness = 2.0
        scene.Children.Add(bln)
        for tp in (s0, s1):
            tick = WpfEllipse()
            tick.Width = tick.Height = 5
            tick.Fill = _brush(accent)
            tick.IsHitTestVisible = False
            WpfCanvas.SetLeft(tick, tp[0] - 2.5)
            WpfCanvas.SetTop(tick, tp[1] - 2.5)
            scene.Children.Add(tick)
        t_mid = to_px(mid[0], mid[1])
        if t_mid is None:
            return
        tb = TextBlock()
        prefix = (label + u" · ") if label else u""
        tb.Text = u"{0}Ø{1} @{2:.0f} · L={3:.0f}".format(
            prefix,
            int(layout.get(u"diam_mm") or 0),
            float(layout.get(u"esp") or 0),
            L,
        )
        tb.Foreground = _brush(accent)
        tb.FontSize = 10
        try:
            tb.FontWeight = FontWeights.SemiBold
        except Exception:
            pass
        tb.IsHitTestVisible = False
        WpfCanvas.SetLeft(tb, t_mid[0] + 8)
        WpfCanvas.SetTop(tb, t_mid[1] - 10)
        scene.Children.Add(tb)

    def _draw_mid_layer(self, scene, to_px, add_polygon):
        """
        - Suples ya definidos: 1 barra a mitad del recorrido c/u (permanecen).
        - Borrador en curso: guías de paños/recorrido (temporales).
        """
        for item in self._suples or []:
            try:
                self._draw_one_mid_bar(
                    scene,
                    to_px,
                    item.get(u"layout"),
                    label=item.get(u"label"),
                )
            except Exception:
                pass

        # Suples en Borde confirmados (polígono + recorrido + barra ⊥)
        for item in self._borde_suples or []:
            pts = item.get(u"pts") or []
            if pts:
                add_polygon(
                    pts,
                    _FACE_BORDE,
                    _FACE_BORDE,
                    stroke_w=2.0,
                    dashed=False,
                    fill_a=45,
                )
                try:
                    c = _centroid_mm(pts)
                    pxpy = to_px(c[0], c[1])
                    tb = TextBlock()
                    tb.Text = u"{0} · lm {1:.0f} · L={2:.0f}".format(
                        item.get(u"label") or u"B",
                        float(item.get(u"lm_mm") or 0),
                        float(item.get(u"L_mm") or 0),
                    )
                    tb.Foreground = _brush(_FACE_BORDE)
                    tb.FontSize = 10
                    try:
                        tb.FontWeight = FontWeights.SemiBold
                    except Exception:
                        pass
                    WpfCanvas.SetLeft(tb, pxpy[0] - 36)
                    WpfCanvas.SetTop(tb, pxpy[1] - 8)
                    scene.Children.Add(tb)
                except Exception:
                    pass
            rec_i = item.get(u"recorrido")
            if rec_i is not None:
                try:
                    a, b = rec_i
                    p0 = to_px(a[0], a[1])
                    p1 = to_px(b[0], b[1])
                    ln = WpfLine()
                    ln.X1, ln.Y1, ln.X2, ln.Y2 = p0[0], p0[1], p1[0], p1[1]
                    ln.Stroke = _brush(_FACE_BORDE)
                    ln.StrokeThickness = 2.5
                    scene.Children.Add(ln)
                except Exception:
                    pass
            try:
                lay = item.get(u"layout")
                if lay is None and pts and rec_i is not None:
                    lay = self._compute_borde_layout(
                        {
                            u"pts": pts,
                            u"lm_mm": item.get(u"lm_mm"),
                            u"L_mm": item.get(u"L_mm"),
                        },
                        rec_i,
                    )
                if lay is not None:
                    self._draw_one_mid_bar(
                        scene,
                        to_px,
                        lay,
                        label=item.get(u"label"),
                        accent=_FACE_BORDE,
                    )
            except Exception:
                pass

        # Borrador Borde: polígono → recorrido → barra ⊥ (si hay recorrido)
        if self._borde_poly is not None:
            pts_d = self._borde_poly.get(u"pts") or []
            if pts_d:
                add_polygon(
                    pts_d,
                    _FACE_BORDE,
                    _FACE_BORDE,
                    stroke_w=2.0,
                    dashed=True,
                    fill_a=40,
                )
                try:
                    c = _centroid_mm(pts_d)
                    pxpy = to_px(c[0], c[1])
                    tb = TextBlock()
                    tb.Text = u"poly · lm {0:.0f} · L={1:.0f}".format(
                        float(self._borde_poly.get(u"lm_mm") or 0),
                        float(self._borde_poly.get(u"L_mm") or 0),
                    )
                    tb.Foreground = _brush(_FACE_BORDE)
                    tb.FontSize = 10
                    WpfCanvas.SetLeft(tb, pxpy[0] - 40)
                    WpfCanvas.SetTop(tb, pxpy[1] - 8)
                    scene.Children.Add(tb)
                except Exception:
                    pass
        if self._borde_recorrido is not None:
            try:
                a, b = self._borde_recorrido
                p0 = to_px(a[0], a[1])
                p1 = to_px(b[0], b[1])
                ln = WpfLine()
                ln.X1, ln.Y1, ln.X2, ln.Y2 = p0[0], p0[1], p1[0], p1[1]
                ln.Stroke = _brush(_RECORRIDO_COLOR)
                ln.StrokeThickness = 2.5
                ln.StrokeDashArray = DoubleCollection()
                ln.StrokeDashArray.Add(6)
                ln.StrokeDashArray.Add(3)
                scene.Children.Add(ln)
            except Exception:
                pass
            try:
                lay_d = self._compute_borde_layout()
                if lay_d is not None:
                    self._draw_one_mid_bar(
                        scene, to_px, lay_d, label=u"borr.", accent=_FACE_BORDE
                    )
            except Exception:
                pass

        # Guías del borrador actual (desaparecen al confirmar el recorrido)
        for pano, col, lab in (
            (self._pano1, _PANO1_COLOR, u"P1"),
            (self._pano2, _PANO2_COLOR, u"P2"),
        ):
            if pano is None:
                continue
            pts = pano.get(u"pts") or []
            add_polygon(pts, col, col, stroke_w=2.0, dashed=False, fill_a=55)
            try:
                c = _centroid_mm(pts)
                pxpy = to_px(c[0], c[1])
                tb = TextBlock()
                tb.Text = u"{0} · lm {1:.0f}".format(
                    lab, float(pano.get(u"lm_mm") or 0)
                )
                tb.Foreground = _brush(col)
                tb.FontSize = 10
                try:
                    tb.FontWeight = FontWeights.SemiBold
                except Exception:
                    pass
                WpfCanvas.SetLeft(tb, pxpy[0] - 28)
                WpfCanvas.SetTop(tb, pxpy[1] - 8)
                scene.Children.Add(tb)
            except Exception:
                pass
        if self._recorrido is not None:
            try:
                a, b = self._recorrido
                p0 = to_px(a[0], a[1])
                p1 = to_px(b[0], b[1])
                ln = WpfLine()
                ln.X1, ln.Y1, ln.X2, ln.Y2 = p0[0], p0[1], p1[0], p1[1]
                ln.Stroke = _brush(_RECORRIDO_COLOR)
                ln.StrokeThickness = 2.5
                scene.Children.Add(ln)
                for p in (p0, p1):
                    el = WpfEllipse()
                    el.Width = el.Height = 8
                    el.Fill = _brush(_RECORRIDO_COLOR)
                    el.Stroke = _brush(u"#E8F4F8")
                    el.StrokeThickness = 1
                    WpfCanvas.SetLeft(el, p[0] - 4)
                    WpfCanvas.SetTop(el, p[1] - 4)
                    scene.Children.Add(el)
            except Exception:
                pass

    def _redraw_canvas(self, view_only=False):
        if not view_only:
            self._view_redraw_pending = False
        cv = self._get_cv_plan()
        if cv is None:
            return
        try:
            cw = float(cv.ActualWidth)
            ch = float(cv.ActualHeight)
        except Exception:
            cw = ch = 0.0
        if cw < 40 or ch < 40:
            return
        if view_only and self._apply_scene_view_transform():
            if self._hover_snap is not None or self._pick_pt1 is not None:
                try:
                    self._refresh_snap_overlay()
                except Exception:
                    pass
            return

        self._last_canvas_cw = cw
        self._last_canvas_ch = ch
        all_pts = []
        loop_polylines = list(self._loop_polylines_mm or [])
        for pts in loop_polylines:
            all_pts.extend(pts)
        for ov in self._overlays or []:
            all_pts.extend(ov.get(u"pts") or [])
        for ar in self._existing_ars or []:
            for ring in ar.get(u"loops") or []:
                all_pts.extend(ring)
            if not ar.get(u"loops"):
                all_pts.extend(ar.get(u"pts") or [])
        # Suples definidos: extremos de cada barra media
        for item in self._suples or []:
            try:
                ends = self._mid_bar_ends_mm(item.get(u"layout"))
                if ends is not None:
                    all_pts.append(ends[0])
                    all_pts.append(ends[1])
            except Exception:
                pass
        # Borrador en curso: guías
        for pano in (self._pano1, self._pano2):
            if pano is not None:
                all_pts.extend(pano.get(u"pts") or [])
        if self._recorrido is not None:
            all_pts.append(self._recorrido[0])
            all_pts.append(self._recorrido[1])
        if not all_pts:
            try:
                cv.Children.Clear()
            except Exception:
                pass
            return

        min_x, min_y, max_x, max_y = _bbox_mm(all_pts)
        bw = max(max_x - min_x, 1.0)
        bh = max(max_y - min_y, 1.0)
        pad = 36.0
        fit_scale = min((cw - 2 * pad) / bw, (ch - 2 * pad) / bh)
        if fit_scale < 1e-12:
            fit_scale = 1e-12
        zoom = max(0.25, min(16.0, float(self._view_zoom or 1.0)))
        self._view_zoom = zoom
        scale = fit_scale * zoom
        bbox_cx = (min_x + max_x) / 2.0
        bbox_cy = (min_y + max_y) / 2.0
        cx_mm = bbox_cx + float(self._view_pan_x or 0.0)
        cy_mm = bbox_cy + float(self._view_pan_y or 0.0)
        ox0 = cw / 2.0 - (bbox_cx - min_x) * fit_scale
        oy0 = ch / 2.0 - (max_y - bbox_cy) * fit_scale
        ox = cw / 2.0 - (cx_mm - min_x) * scale
        oy = ch / 2.0 - (max_y - cy_mm) * scale
        self._view_xform = {
            u"min_x": min_x,
            u"max_x": max_x,
            u"min_y": min_y,
            u"max_y": max_y,
            u"ox": ox,
            u"oy": oy,
            u"scale": scale,
            u"fit_scale": fit_scale,
            u"cw": cw,
            u"ch": ch,
        }
        self._scene_base = {
            u"ox0": ox0,
            u"oy0": oy0,
            u"fit_scale": fit_scale,
            u"min_x": min_x,
            u"max_x": max_x,
            u"min_y": min_y,
            u"max_y": max_y,
            u"cw": cw,
            u"ch": ch,
        }
        try:
            cv.Children.Clear()
        except Exception:
            return
        scene = WpfCanvas()
        scene.IsHitTestVisible = False
        hud = WpfCanvas()
        hud.IsHitTestVisible = False
        cv.Children.Add(scene)
        cv.Children.Add(hud)
        self._scene_layer = scene
        self._hud_layer = hud
        self._scene_matrix_transform = None

        def to_px(xmm, ymm):
            return (
                ox0 + (xmm - min_x) * fit_scale,
                oy0 + (max_y - ymm) * fit_scale,
            )

        def _add_polygon(pts, fill_hex, stroke_hex, stroke_w=1.2, dashed=False, fill_a=200):
            if not pts or len(pts) < 3:
                return
            poly = WpfPolygon()
            pc = PointCollection()
            for xmm, ymm in pts:
                px, py = to_px(xmm, ymm)
                pc.Add(WpfPoint(px, py))
            poly.Points = pc
            poly.Fill = _brush(fill_hex, fill_a)
            poly.Stroke = _brush(stroke_hex)
            poly.StrokeThickness = stroke_w
            if dashed:
                try:
                    dashes = DoubleCollection()
                    dashes.Add(4)
                    dashes.Add(3)
                    poly.StrokeDashArray = dashes
                except Exception:
                    pass
            scene.Children.Add(poly)

        hdr = self._ui_txt_canvas_header
        header_text = None
        try:
            nw, nb, _np = _count_ctx(self._overlays)
            n_ar = len(self._existing_ars or [])
            header_text = (
                u"PLANTA · {:.0f}×{:.0f} mm · AR {} · muros {} · vigas {}"
            ).format(bw, bh, n_ar, nw, nb)
        except Exception:
            pass

        paint_planta_context_layers(
            scene=scene,
            hud=hud,
            to_px=to_px,
            add_polygon=_add_polygon,
            loop_polylines=loop_polylines,
            overlays=self._overlays,
            existing_ars=self._existing_ars,
            curves_outer=self._curves,
            plane=self._plane,
            ctx_geo_cache=self._ensure_ctx_geo_cache(),
            ox0=ox0,
            oy0=oy0,
            min_x=min_x,
            max_x=max_x,
            min_y=min_y,
            max_y=max_y,
            fit_scale=fit_scale,
            bw=bw,
            bh=bh,
            cw=cw,
            ch=ch,
            major_xyz=None,
            mid_layer_callback=self._draw_mid_layer,
            header_tb=hdr,
            header_text=header_text,
            context_line_scale=0.5,
        )
        try:
            self._ensure_snap_geometry()
        except Exception:
            self._snap_geo_dirty = False
        try:
            self._apply_scene_view_transform()
        except Exception:
            pass
        try:
            self._refresh_snap_overlay()
        except Exception:
            pass

    def show(self):
        try:
            hwnd = revit_main_hwnd(self._uiapp)
            bind_center_wpf_on_revit_monitor(self._win, hwnd)
            position_wpf_window_center_on_monitor(self._win, hwnd)
        except Exception:
            pass
        try:
            from System.Windows.Interop import WindowInteropHelper

            hwnd = revit_main_hwnd(self._uiapp)
            if hwnd is not None:
                WindowInteropHelper(self._win).Owner = hwnd
        except Exception:
            pass
        try:
            self._win.WindowState = WindowState.Maximized
        except Exception:
            pass
        _register_singleton(self._win)
        self._win.Show()
        try:
            self._win.Activate()
            self._win.Focus()
            self._redraw_canvas()
            cv = self._get_cv_plan()
            if cv is not None:
                cv.Focus()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Entrada
# ---------------------------------------------------------------------------


def run(revit):
    uiapp = revit
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
    if doc.IsFamilyDocument:
        _mostrar_aviso(uiapp, u"Abra un proyecto (no un family document).")
        return
    if _focus_existing(uiapp):
        return
    floor = _pick_floor(uidoc, doc, uiapp)
    if floor is None:
        return
    loops = obtener_loops_sketch(floor, doc)
    if not loops:
        _mostrar_aviso(
            uiapp,
            u"La losa no tiene Sketch válido.",
            content=u"Se requiere Floor.SketchId con Profile (loop exterior).",
        )
        return
    outer = loops[0]
    plane = _plane_from_curves(outer)
    if plane is None:
        _mostrar_aviso(uiapp, u"No se pudo obtener el plano del Sketch.")
        return
    try:
        ctrl = SuplesLosaController(uiapp, uidoc, doc, floor, outer, loops, plane)
        ctrl.show()
    except Exception as ex:
        _unregister_singleton()
        _mostrar_aviso(uiapp, u"Error al abrir la UI.", content=_as_unicode(ex))


def run_pyrevit(revit):
    run(revit)
