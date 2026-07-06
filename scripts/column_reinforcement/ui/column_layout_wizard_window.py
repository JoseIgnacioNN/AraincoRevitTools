# -*- coding: utf-8 -*-
"""Asistente WPF único: rejilla por sección + esquema de troceo."""

import os

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPIUI")

from System import AppDomain, Double
from System.IO import File
from System.Windows import FontWeights, GridLength, GridUnitType, RoutedEventHandler, Size, SizeToContent, SystemParameters, TextAlignment, VerticalAlignment, Visibility
from System.Windows.Markup import XamlReader
from System.Windows.Controls import Canvas, ScrollBarVisibility, TextBlock
from System.Windows.Media import Brushes, Color, RotateTransform, SolidColorBrush
from System.Windows.Shapes import Ellipse, Line as WpfLine, Rectangle

try:
    from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
except Exception:
    BIMTOOLS_DARK_STYLES_XML = u""

try:
    from Autodesk.Revit.UI import TaskDialog
except Exception:
    TaskDialog = None

from column_reinforcement.ui.troceo_scheme_window import (
    TroceoSchemeController,
    TroceoSchemeOutcome,
    _unpack_troceo_row,
)

_WIZARD_SINGLETON_KEY = u"Arainco.column_reinforcement.ColumnLayoutWizardSingleton"

_PREVIEW_COVER_MM_CAP = 45.72
# Diámetro fijo en px de los círculos de barra en la vista en planta (no depende de la cuenta A×B).
_PREVIEW_BAR_DOT_OUTER_PX = 14.0
_PREVIEW_BAR_DOT_INNER_PX = 12.0
# Factor sobre el encaje geométrico (aire mínimo al borde del canvas).
_PREVIEW_USABLE_FRACTION = 0.98
# Franja inferior del canvas (vista en planta): «A» + leyenda, sin invadir el esquema.
_PLAN_PREVIEW_BOTTOM_BAND_PX = 92.0
_BARS_COUNT_MIN = 2
_BARS_COUNT_MAX = 10
# Fase troceo: altura del formulario = alto del área de trabajo (útil en Full HD ~1040 px sin barra).
_TROCEO_WORKAREA_HEIGHT_TRIM_PX = 0.0


def _preview_get_positions(length, count, edge_cover):
    if count < 2:
        return []
    L = float(length)
    ec = float(edge_cover)
    if count == 2:
        return [-L / 2.0 + ec, L / 2.0 - ec]
    span = L - (2.0 * ec)
    if span <= 1e-6:
        span = max(L * 0.5, 1e-3)
        ec = (L - span) / 2.0
    spacing = span / float(count - 1)
    return [-L / 2.0 + ec + (i * spacing) for i in range(count)]


def _preview_perimeter_ij(nx, ny):
    if nx < 2 or ny < 2:
        return []
    out = []
    for ix in range(nx):
        out.append((ix, ny - 1))
    for iy in range(ny - 2, -1, -1):
        out.append((nx - 1, iy))
    for ix in range(nx - 2, -1, -1):
        out.append((ix, 0))
    for iy in range(1, ny - 1):
        out.append((0, iy))
    return out


def _preview_outer_ij(bars_a, bars_b):
    return _preview_perimeter_ij(int(bars_a), int(bars_b))


def _preview_inner_ij(bars_a, bars_b):
    nx = int(bars_a) - 2
    ny = int(bars_b) - 2
    if nx < 2 or ny < 2:
        return []
    return [
        (ix + 1, iy + 1)
        for ix, iy in _preview_perimeter_ij(nx, ny)
    ]


def _preview_bar_corner_radius_px(ia, ib, bars_a, bars_b, inner_on):
    u"""Radio en px del círculo de barra en (ia, ib), coherente con add_dot (diámetros fijos)."""
    ou = set(_preview_outer_ij(bars_a, bars_b))
    inn = set(_preview_inner_ij(bars_a, bars_b)) if inner_on else set()
    if (ia, ib) in ou:
        return 0.5 * float(_PREVIEW_BAR_DOT_OUTER_PX)
    if inner_on and (ia, ib) in inn:
        return 0.5 * float(_PREVIEW_BAR_DOT_INNER_PX)
    return 0.5 * float(_PREVIEW_BAR_DOT_OUTER_PX)


def _preview_plan_concrete_rect_around_steel(
    nom_left, nom_top, nom_w, nom_h, bul, but, bur, bub, gap_g
):
    u"""Expande el rectángulo nominal de cara de hormigón hasta dejar ``gap_g`` px de aire respecto al acero."""
    pad_l = max(0.0, (float(nom_left) + float(gap_g)) - float(bul))
    pad_t = max(0.0, (float(nom_top) + float(gap_g)) - float(but))
    pad_r = max(0.0, float(bur) - (float(nom_left) + float(nom_w) - float(gap_g)))
    pad_b = max(0.0, float(bub) - (float(nom_top) + float(nom_h) - float(gap_g)))
    return (
        float(nom_left) - pad_l,
        float(nom_top) - pad_t,
        float(nom_w) + pad_l + pad_r,
        float(nom_h) + pad_t + pad_b,
    )


def _preview_plan_steel_bbox_px(
    nom_left,
    nom_top,
    draw_w,
    draw_h,
    ss_f,
    sl_f,
    ba,
    bb,
    offs_a,
    offs_b,
    inner_on,
    rect_defs,
    tie_defs,
    px_margin,
):
    u"""Bounding box (px) de barras + estribos + trabas respecto al mismo ``to_px`` que el overlay."""
    from column_stirrup_creator import tie_axis_shift_toward_section_center

    hs = ss_f / 2.0
    hl = sl_f / 2.0
    ro = 0.5 * _PREVIEW_BAR_DOT_OUTER_PX
    ri = 0.5 * _PREVIEW_BAR_DOT_INNER_PX
    px_margin_draw = px_margin + ro

    def to_px(da, db):
        x = float(nom_left) + (float(da) + hs) / ss_f * draw_w
        y = float(nom_top) + (float(db) + hl) / sl_f * draw_h
        return x, y

    xs = []
    ys = []
    for ix, iy in _preview_outer_ij(ba, bb):
        x, y = to_px(offs_a[ix], offs_b[iy])
        xs.extend([x - ro, x + ro])
        ys.extend([y - ro, y + ro])
    if inner_on:
        for ix, iy in _preview_inner_ij(ba, bb):
            x, y = to_px(offs_a[ix], offs_b[iy])
            xs.extend([x - ri, x + ri])
            ys.extend([y - ri, y + ri])

    for idx_a, idx_b, sp_a, sp_b in rect_defs:
        a0 = max(0, min(idx_a, len(offs_a) - 1))
        a1 = max(0, min(idx_a + sp_a, len(offs_a) - 1))
        b0 = max(0, min(idx_b, len(offs_b) - 1))
        b1 = max(0, min(idx_b + sp_b, len(offs_b) - 1))
        x0, y0 = to_px(offs_a[a0], offs_b[b0])
        x1, y1 = to_px(offs_a[a1], offs_b[b1])
        rx = min(x0, x1) - px_margin_draw
        ry = min(y0, y1) - px_margin_draw
        rw = abs(x1 - x0) + 2.0 * px_margin_draw
        rh = abs(y1 - y0) + 2.0 * px_margin_draw
        xs.extend([rx, rx + rw])
        ys.extend([ry, ry + rh])

    if tie_defs and offs_a and offs_b:
        x_edge0, y_edge0 = to_px(offs_a[0], offs_b[0])
        x_edge1, y_edge1 = to_px(offs_a[-1], offs_b[-1])
        x_mid_px, y_mid_px = to_px(0.0, 0.0)
        for idx, is_a in tie_defs:
            if is_a:
                if 0 <= idx < len(offs_a):
                    r0 = float(_preview_bar_corner_radius_px(idx, 0, ba, bb, inner_on))
                    r1 = float(
                        _preview_bar_corner_radius_px(idx, bb - 1, ba, bb, inner_on)
                    )
                    r_t = max(r0, r1)
                    x_px, _ = to_px(offs_a[idx], 0.0)
                    x_px = x_mid_px + tie_axis_shift_toward_section_center(
                        x_px - x_mid_px, r_t, tie_index=idx, bar_count=ba
                    )
                    ylo = y_edge0 - px_margin
                    yhi = y_edge1 + px_margin
                    xs.extend([x_px, x_px])
                    ys.extend([ylo, yhi])
            else:
                if 0 <= idx < len(offs_b):
                    r0 = float(_preview_bar_corner_radius_px(0, idx, ba, bb, inner_on))
                    r1 = float(
                        _preview_bar_corner_radius_px(ba - 1, idx, ba, bb, inner_on)
                    )
                    r_t = max(r0, r1)
                    _, y_px = to_px(0.0, offs_b[idx])
                    y_px = y_mid_px + tie_axis_shift_toward_section_center(
                        y_px - y_mid_px, r_t, tie_index=idx, bar_count=bb
                    )
                    xlo = x_edge0 - px_margin
                    xhi = x_edge1 + px_margin
                    xs.extend([xlo, xhi])
                    ys.extend([y_px, y_px])

    if not xs:
        return None
    return min(xs), min(ys), max(xs), max(ys)


def _preview_brushes():
    from System.Windows.Media import Color, SolidColorBrush

    def br(hex_str):
        h = hex_str.lstrip("#")
        return SolidColorBrush(
            Color.FromRgb(
                int(h[0:2], 16),
                int(h[2:4], 16),
                int(h[4:6], 16),
            )
        )

    return {
        "rect": br("#475569"),
        "outer": br("#22d3ee"),
        "inner": br("#7dd3fc"),
        "muted": br("#94a3b8"),
    }


class StirrupSectionConfig(object):
    """Configuración de estribos para una sección (paso Rejilla).

    ``spacing_mm`` y ``stirrup_bar_type`` quedan como respaldo del pipeline; el espaciamiento
    y el Ø/tipo operativos se eligen por columna en el esquema vertical de troceo.
    En rejilla solo se definen patrones A/B de confinamiento.
    """

    def __init__(self):
        self.skip = False
        self.stirrup_bar_type = None   # Respaldo; lo habitual es None hasta troceo
        self.spacing_mm = 200.0
        self.sel_a_text = u""
        self.sel_b_text = u""


class ColumnLayoutWizardOutcome(object):
    def __init__(
        self,
        cancelled=False,
        already_running=False,
        section_grid_config=None,
        troceo_outcome=None,
        stirrup_configs=None,
        stirrup_spacing_by_column_id=None,
        stirrup_bar_type_by_column_id=None,
        stirrup_policy_by_column_id=None,
        stirrup_spacing_by_column_lot=None,
        stirrup_bar_type_by_column_lot=None,
    ):
        self.cancelled = bool(cancelled)
        self.already_running = bool(already_running)
        self.section_grid_config = section_grid_config
        self.troceo_outcome = troceo_outcome
        self.stirrup_configs = stirrup_configs or {}
        _sp = {}
        for _k, _v in (stirrup_spacing_by_column_id or {}).items():
            try:
                _sp[int(_k)] = float(_v)
            except Exception:
                pass
        self.stirrup_spacing_by_column_id = _sp
        _btmap = {}
        for _k, _v in (stirrup_bar_type_by_column_id or {}).items():
            try:
                _btmap[int(_k)] = _v
            except Exception:
                pass
        self.stirrup_bar_type_by_column_id = _btmap
        _pol = {}
        for _k, _v in (stirrup_policy_by_column_id or {}).items():
            try:
                _pol[int(_k)] = unicode(_v)
            except Exception:
                pass
        self.stirrup_policy_by_column_id = _pol
        _sp_lot = {}
        for _k, _v in (stirrup_spacing_by_column_lot or {}).items():
            try:
                _inner = {}
                for _lk, _lv in (_v or {}).items():
                    _inner[int(_lk)] = float(_lv)
                _sp_lot[int(_k)] = _inner
            except Exception:
                pass
        self.stirrup_spacing_by_column_lot = _sp_lot
        _bt_lot = {}
        for _k, _v in (stirrup_bar_type_by_column_lot or {}).items():
            try:
                _inner = {}
                for _lk, _lv in (_v or {}).items():
                    _inner[int(_lk)] = _lv
                _bt_lot[int(_k)] = _inner
            except Exception:
                pass
        self.stirrup_bar_type_by_column_lot = _bt_lot


def _load_xaml_text():
    here = os.path.dirname(os.path.abspath(__file__))
    xaml_path = os.path.join(here, "column_layout_wizard_window.xaml")
    txt = File.ReadAllText(xaml_path)
    return txt.replace("__BIMTOOLS_DARK_STYLES__", BIMTOOLS_DARK_STYLES_XML)


class ColumnLayoutWizardController(object):
    def __init__(
        self,
        section_meta,
        troceo_rows,
        uiapp,
        uidoc,
        doc,
        default_bar_diam_mm,
    ):
        """
        ``section_meta``: lista de ``(section_key, título_ui)`` en orden de trabajo.
        ``troceo_rows``: filas para ``TroceoSchemeController`` (Z ascendente).
        """
        self._section_meta = list(section_meta or [])
        self._ordered_keys = [t[0] for t in self._section_meta]
        self._troceo_rows = troceo_rows
        self._uiapp = uiapp
        self._uidoc = uidoc
        self._doc = doc
        self._default_bar_diam_mm = float(default_bar_diam_mm)
        self._section_grid_config = {}
        self._stirrup_section_data = {}
        self._stirrup_spacing_by_troceo_slot = {}
        self._stirrup_bar_type_by_troceo_slot = {}
        self._stirrup_policy_by_troceo_slot = {}
        self._section_idx = 0
        self._troceo_ctrl = None
        self._troceo_outcome = None
        self._suppress_bars_slider = False
        self._stirrup_rebar_types = []   # list of (label, RebarBarType)

        self.window = XamlReader.Parse(_load_xaml_text())
        self._v = Visibility

        self._scroll_grid = self.window.FindName("ScrollGridStep")
        self._scroll_troceo = self.window.FindName("ScrollTroceoStep")
        self._wizard_root_grid = self.window.FindName("WizardRootGrid")
        self._grid_step_root = self.window.FindName("GridStepRoot")
        self._footer1 = self.window.FindName("FooterStep1")
        self._footer2 = self.window.FindName("FooterStep2")
        self._step1_badge = self.window.FindName("Step1Badge")
        self._step2_badge = self.window.FindName("Step2Badge")
        self._step1_badge_text = self.window.FindName("Step1BadgeText")
        self._step2_badge_text = self.window.FindName("Step2BadgeText")

        self._tb_title = self.window.FindName("GridSectionTitle")
        self._tb_section_index = self.window.FindName("GridSectionIndex")
        self._tb_step_progress = self.window.FindName("WizardStepProgress")
        self._tb_inner_hint = self.window.FindName("InnerBarsCountHint")
        self._footer_hint = self.window.FindName("FooterStep1Hint")
        self._block_dims = self.window.FindName("BlockSectionDims")
        self._slider_a = self.window.FindName("SliderBarsA")
        self._slider_b = self.window.FindName("SliderBarsB")
        self._tb_bars_a_val = self.window.FindName("TbBarsAValue")
        self._tb_bars_b_val = self.window.FindName("TbBarsBValue")
        self._chk_inner = self.window.FindName("ChkInnerOutline")
        self._preview_canvas = self.window.FindName("GridPreviewCanvas")

        self._btn_prev = self.window.FindName("BtnGridPrev")
        self._btn_cancel = self.window.FindName("BtnWizardCancel")
        self._btn_next = self.window.FindName("BtnGridNext")

        self._btn_troceo_cancel = self.window.FindName("BtnTroceoCancel")
        self._btn_troceo_ok = self.window.FindName("BtnTroceoConfirm")

        self._stirrup_controls_host = self.window.FindName("StirrupControlsHost")
        self._stirrup_diam_combo = None
        self._stirrup_combo_a = None
        self._stirrup_combo_b = None

        self.window.Closed += self._on_window_closed
        self.window.Loaded += RoutedEventHandler(self._on_wizard_loaded)
        if self._btn_cancel is not None:
            self._btn_cancel.Click += RoutedEventHandler(self._on_cancel_step1)
        if self._btn_next is not None:
            self._btn_next.Click += RoutedEventHandler(self._on_grid_next)
        if self._btn_prev is not None:
            self._btn_prev.Click += RoutedEventHandler(self._on_grid_prev)

        self._wire_preview_handlers()
        self._build_stirrup_controls()
        self._ensure_stirrup_rebar_types_loaded()
        self._wire_stirrup_handlers()

        self._load_current_section_into_ui()
        self._refresh_nav_buttons()
        self._refresh_grid_next_label()
        self._update_step_badges(1)

    def _on_wizard_loaded(self, sender, args):
        self._refresh_section_dims_label()
        self._refresh_grid_preview()
        self._apply_rejilla_window_layout()

    def _apply_rejilla_window_layout(self):
        u"""Paso Rejilla: fila central Auto + altura al contenido (sin aire inferior)."""
        win = self.window
        if win is None:
            return
        grid = self._wizard_root_grid
        try:
            if grid is not None and grid.RowDefinitions.Count > 2:
                grid.RowDefinitions[2].Height = GridLength(1.0, GridUnitType.Auto)
        except Exception:
            pass
        try:
            if self._scroll_grid is not None:
                self._scroll_grid.VerticalAlignment = VerticalAlignment.Top
                self._scroll_grid.VerticalScrollBarVisibility = (
                    ScrollBarVisibility.Auto
                )
        except Exception:
            pass
        try:
            win.SizeToContent = SizeToContent.Height
            win.UpdateLayout()
            cap = float(SystemParameters.WorkArea.Height) * 0.94
            actual_h = float(win.ActualHeight)
            min_h = float(win.MinHeight)
            if actual_h > cap:
                win.SizeToContent = SizeToContent.Manual
                win.Height = cap
            elif actual_h < min_h:
                win.SizeToContent = SizeToContent.Manual
                win.Height = min_h
        except Exception:
            pass
        try:
            from revit_wpf_window_position import (
                position_wpf_window_center_work_area,
            )

            position_wpf_window_center_work_area(win)
        except Exception:
            pass

    def _wire_preview_handlers(self):
        try:
            from System import Double
            from System.Windows import RoutedPropertyChangedEventHandler

            if self._slider_a is not None:
                self._slider_a_value_handler = RoutedPropertyChangedEventHandler[Double](
                    self._on_bars_slider_changed
                )
                self._slider_a.ValueChanged += self._slider_a_value_handler
            if self._slider_b is not None:
                self._slider_b_value_handler = RoutedPropertyChangedEventHandler[Double](
                    self._on_bars_slider_changed
                )
                self._slider_b.ValueChanged += self._slider_b_value_handler
        except Exception:
            if self._slider_a is not None:
                self._slider_a.ValueChanged += self._on_bars_slider_changed
            if self._slider_b is not None:
                self._slider_b.ValueChanged += self._on_bars_slider_changed
            self._slider_a_value_handler = None
            self._slider_b_value_handler = None
        if self._chk_inner is not None:
            self._chk_inner.Checked += self._on_preview_inputs_changed
            self._chk_inner.Unchecked += self._on_preview_inputs_changed
        if self._preview_canvas is not None:
            self._preview_canvas.SizeChanged += self._on_preview_size_changed

    # ------------------------------------------------------------------
    # Estribos — población de controles y helpers
    # ------------------------------------------------------------------

    def _build_stirrup_controls(self):
        u"""Crea los controles de estribos programáticamente (mismo patrón que troceo) para evitar problemas con estilos del sistema."""
        host = self._stirrup_controls_host
        if host is None:
            return
        try:
            from System.Windows import HorizontalAlignment, Thickness
            from System.Windows.Controls import ComboBox, StackPanel as _SP

            style_combo = None
            style_item = None
            try:
                style_combo = self.window.TryFindResource(u"ComboStretch")
            except Exception:
                pass
            try:
                style_item = self.window.TryFindResource(u"ComboItem")
            except Exception:
                pass

            def _label(text):
                tb = TextBlock()
                tb.Text = text
                try:
                    from System.Windows.Media import Color, SolidColorBrush
                    tb.Foreground = SolidColorBrush(Color.FromRgb(0xD5, 0xEA, 0xF2))
                except Exception:
                    pass
                tb.FontSize = 11.0
                try:
                    tb.Margin = Thickness(0.0, 0.0, 0.0, 4.0)
                except Exception:
                    pass
                return tb

            def _combo(bottom_margin=10.0):
                cb = ComboBox()
                if style_combo is not None:
                    try:
                        cb.Style = style_combo
                    except Exception:
                        pass
                if style_item is not None:
                    try:
                        cb.ItemContainerStyle = style_item
                    except Exception:
                        pass
                try:
                    cb.Margin = Thickness(0.0, 0.0, 0.0, bottom_margin)
                    cb.HorizontalAlignment = HorizontalAlignment.Stretch
                except Exception:
                    pass
                return cb

            self._stirrup_combo_a = _combo(10.0)
            self._stirrup_combo_b = _combo(0.0)

            for widget in (
                _label(u"Patr\u00f3n lado A"),
                self._stirrup_combo_a,
                _label(u"Patr\u00f3n lado B"),
                self._stirrup_combo_b,
            ):
                try:
                    host.Children.Add(widget)
                except Exception:
                    pass
        except Exception:
            pass

    def _ensure_stirrup_rebar_types_loaded(self):
        u"""Lista (etiqueta, RebarBarType) para el paso Troceo; sin combo en Rejilla."""
        self._stirrup_rebar_types = []
        if self._doc is None:
            return
        try:
            from Autodesk.Revit.DB import FilteredElementCollector, BuiltInParameter
            from Autodesk.Revit.DB.Structure import RebarBarType

            bar_elems = list(
                FilteredElementCollector(self._doc)
                .OfClass(RebarBarType)
                .WhereElementIsElementType()
                .ToElements()
            )
            for bt in bar_elems:
                label = None
                for bip in (
                    BuiltInParameter.SYMBOL_NAME_PARAM,
                    BuiltInParameter.ALL_MODEL_TYPE_NAME,
                ):
                    try:
                        p = bt.get_Parameter(bip)
                        if p and p.HasValue:
                            s = p.AsString()
                            if s:
                                label = s.strip()
                                break
                    except Exception:
                        pass
                if label is None:
                    label = u"RebarBarType_{}".format(bt.Id.IntegerValue)
                self._stirrup_rebar_types.append((label, bt))
            self._stirrup_rebar_types.sort(key=lambda x: x[0])
        except Exception:
            self._stirrup_rebar_types = []

    def _populate_stirrup_diam_combo(self):
        u"""Obsoleto: la rejilla ya no tiene combo de \u00d8; solo se mantiene por compatibilidad."""
        self._ensure_stirrup_rebar_types_loaded()

    def _default_stirrup_spacing_mm_for_column_elem(self, elem):
        u"""Valor inicial en el esquema vertical (mm); el espaciamiento se define por columna ahí."""
        return 200.0

    def _default_stirrup_bar_type_for_column_elem(self, elem):
        u"""RebarBarType por defecto del paso Rejilla para la sección de esta columna."""
        if elem is None:
            return None
        try:
            from column_reinforcement_layout_rps import (
                _canonical_section_mm_key,
                get_column_dimensions,
            )

            w, d, _, _, _, _ = get_column_dimensions(elem)
            sk = _canonical_section_mm_key(w, d)
            cfg = self._stirrup_section_data.get(sk)
            if cfg and not getattr(cfg, "skip", False):
                return getattr(cfg, "stirrup_bar_type", None)
        except Exception:
            pass
        if self._stirrup_rebar_types:
            return self._stirrup_rebar_types[0][1]
        return None

    def _stirrup_bar_type_choices_for_troceo_scheme(self):
        u"""Lista (etiqueta, tipo) si hay estribos activos en alguna secci\u00f3n (no \u00abo omitir \u00bb)."""
        if not self._stirrup_rebar_types:
            return None
        for cfg in self._stirrup_section_data.values():
            if cfg and not getattr(cfg, "skip", True):
                return self._stirrup_rebar_types
        return None

    def _update_stirrup_pattern_combo(self, combo, val):
        if combo is None:
            return
        try:
            current = str(combo.SelectedItem) if combo.SelectedIndex >= 0 else None
        except Exception:
            current = None
        combo.Items.Clear()
        try:
            from column_stirrup_creator import stirrup_pattern_options
            opts = stirrup_pattern_options(int(val))
        except Exception:
            opts = [u"Perimetral Únicamente"]
        for o in opts:
            combo.Items.Add(o)
        restored = False
        if current:
            for i in range(combo.Items.Count):
                try:
                    if str(combo.Items[i]) == current:
                        combo.SelectedIndex = i
                        restored = True
                        break
                except Exception:
                    pass
        if not restored and combo.Items.Count > 0:
            combo.SelectedIndex = 0

    def _update_stirrup_pattern_combos_for_current(self):
        ba, bb = self._preview_bars_ab()
        self._update_stirrup_pattern_combo(self._stirrup_combo_a, int(ba))
        self._update_stirrup_pattern_combo(self._stirrup_combo_b, int(bb))

    def _wire_stirrup_handlers(self):
        for combo in (
            self._stirrup_combo_a,
            self._stirrup_combo_b,
        ):
            if combo is not None:
                try:
                    combo.SelectionChanged += self._on_stirrup_input_changed
                except Exception:
                    pass

    def _on_stirrup_input_changed(self, sender, args):
        self._refresh_grid_preview()

    def _commit_stirrup_for_section(self, sk):
        cfg = StirrupSectionConfig()
        cfg.skip = False
        cfg.stirrup_bar_type = None
        cfg.spacing_mm = 200.0
        try:
            if self._stirrup_combo_a and self._stirrup_combo_a.SelectedIndex >= 0:
                cfg.sel_a_text = str(self._stirrup_combo_a.SelectedItem or u"")
        except Exception:
            cfg.sel_a_text = u""
        try:
            if self._stirrup_combo_b and self._stirrup_combo_b.SelectedIndex >= 0:
                cfg.sel_b_text = str(self._stirrup_combo_b.SelectedItem or u"")
        except Exception:
            cfg.sel_b_text = u""
        self._stirrup_section_data[sk] = cfg

    def _load_stirrup_for_section(self, sk):
        self._update_stirrup_pattern_combos_for_current()
        cfg = self._stirrup_section_data.get(sk)
        if cfg is None:
            return
        try:
            if self._stirrup_combo_a and cfg.sel_a_text:
                for i in range(self._stirrup_combo_a.Items.Count):
                    if str(self._stirrup_combo_a.Items[i]) == cfg.sel_a_text:
                        self._stirrup_combo_a.SelectedIndex = i
                        break
        except Exception:
            pass
        try:
            if self._stirrup_combo_b and cfg.sel_b_text:
                for i in range(self._stirrup_combo_b.Items.Count):
                    if str(self._stirrup_combo_b.Items[i]) == cfg.sel_b_text:
                        self._stirrup_combo_b.SelectedIndex = i
                        break
        except Exception:
            pass

    # ------------------------------------------------------------------

    def _on_bars_slider_changed(self, sender, args):
        if getattr(self, "_suppress_bars_slider", False):
            return
        self._sync_bars_value_labels_only()
        self._update_stirrup_pattern_combos_for_current()
        self._refresh_grid_preview()

    def _sync_bars_value_labels_only(self):
        a, b = self._preview_bars_ab()
        try:
            if self._tb_bars_a_val is not None:
                self._tb_bars_a_val.Text = str(int(a))
            if self._tb_bars_b_val is not None:
                self._tb_bars_b_val.Text = str(int(b))
        except Exception:
            pass

    def _apply_bars_sliders(self, a, b):
        a = max(_BARS_COUNT_MIN, min(_BARS_COUNT_MAX, int(a)))
        b = max(_BARS_COUNT_MIN, min(_BARS_COUNT_MAX, int(b)))
        self._suppress_bars_slider = True
        try:
            if self._slider_a is not None:
                self._slider_a.Value = float(a)
            if self._slider_b is not None:
                self._slider_b.Value = float(b)
        finally:
            self._suppress_bars_slider = False
        self._sync_bars_value_labels_only()

    def _refresh_inner_bars_hint(self):
        if self._tb_inner_hint is None:
            return
        ba, bb = self._preview_bars_ab()
        inner = self._preview_inner_checked()
        n_inner = len(_preview_inner_ij(ba, bb))
        if not inner:
            self._tb_inner_hint.Text = u""
            return
        if n_inner <= 0:
            self._tb_inner_hint.Text = u"Anillo interior: requiere A ≥ 4 y B ≥ 4."
            return
        self._tb_inner_hint.Text = u"+ {0} barra{1} en anillo interior (vista previa).".format(
            n_inner,
            u"s" if n_inner != 1 else u"",
        )

    def _refresh_section_footer_hints(self):
        nsec = len(self._ordered_keys)
        if self._tb_section_index is not None:
            try:
                if nsec > 1:
                    self._tb_section_index.Visibility = self._v.Visible
                    self._tb_section_index.Text = u"Sección {} de {}".format(
                        self._section_idx + 1,
                        nsec,
                    )
                else:
                    self._tb_section_index.Visibility = self._v.Collapsed
            except Exception:
                pass
        if self._footer_hint is not None:
            try:
                if nsec > 1:
                    self._footer_hint.Visibility = self._v.Visible
                    self._footer_hint.Text = (
                        u"{0} secciones distintas en el modelo (agrupadas por tamaño). "
                        u"Use «Siguiente sección →» para cada una."
                    ).format(nsec)
                else:
                    self._footer_hint.Visibility = self._v.Collapsed
            except Exception:
                pass

    def _on_preview_inputs_changed(self, sender, args):
        self._refresh_grid_preview()

    def _on_preview_size_changed(self, sender, args):
        self._refresh_grid_preview()

    def _current_section_mm(self):
        if not self._ordered_keys:
            return 400.0, 600.0
        sk = self._ordered_keys[self._section_idx]
        try:
            return float(sk[0]), float(sk[1])
        except Exception:
            return 400.0, 600.0

    def _preview_bars_ab(self):
        try:
            if self._slider_a is not None and self._slider_b is not None:
                a = int(round(float(self._slider_a.Value)))
                b = int(round(float(self._slider_b.Value)))
            else:
                return 4, 6
        except Exception:
            return 4, 6
        return (
            max(_BARS_COUNT_MIN, min(_BARS_COUNT_MAX, a)),
            max(_BARS_COUNT_MIN, min(_BARS_COUNT_MAX, b)),
        )

    def _preview_inner_checked(self):
        try:
            return bool(self._chk_inner.IsChecked)
        except Exception:
            return False

    def _refresh_section_dims_label(self):
        if self._block_dims is None:
            return
        ss, sl = self._current_section_mm()
        try:
            self._block_dims.Text = u"{}×{}".format(int(round(ss)), int(round(sl)))
        except Exception:
            pass

    def _refresh_grid_preview(self):
        cv = self._preview_canvas
        if cv is None:
            return
        self._refresh_section_dims_label()
        try:
            cw = float(cv.ActualWidth)
            ch = float(cv.ActualHeight)
        except Exception:
            return
        if cw < 8.0 or ch < 8.0:
            return
        br = _preview_brushes()
        cv.Children.Clear()
        ss_f, sl_f = self._current_section_mm()
        ba, bb = self._preview_bars_ab()
        inner_on = self._preview_inner_checked()
        cover = min(
            float(_PREVIEW_COVER_MM_CAP),
            max(12.0, min(ss_f, sl_f) * 0.085),
        )
        try:
            offs_a = _preview_get_positions(ss_f, ba, cover)
            offs_b = _preview_get_positions(sl_f, bb, cover)
        except Exception:
            return
        if len(offs_a) != ba or len(offs_b) != bb:
            return

        margin = 6.0
        usable_w = max(8.0, cw - 2.0 * margin)
        usable_h = max(8.0, ch - 2.0 * margin)
        diagram_h_max = max(48.0, usable_h - _PLAN_PREVIEW_BOTTOM_BAND_PX)
        scale_geo = min(usable_w / ss_f, diagram_h_max / sl_f)
        scale = scale_geo * _PREVIEW_USABLE_FRACTION
        draw_w = ss_f * scale
        draw_h = sl_f * scale
        left = (cw - draw_w) / 2.0
        top = margin + max(0.0, (diagram_h_max - draw_h) / 2.0)

        hs = ss_f / 2.0
        hl = sl_f / 2.0

        def to_px(da, db):
            x = left + (float(da) + hs) / ss_f * draw_w
            y = top + (float(db) + hl) / sl_f * draw_h
            return x, y

        nom_left = left
        nom_top = top
        nom_w = draw_w
        nom_h = draw_h

        edge_to_bar_mm = (
            offs_a[-1] - ss_f / 2.0
            if offs_a[-1] < ss_f / 2.0
            else ss_f / 2.0 - offs_a[0]
        )
        edge_to_bar_mm = abs(edge_to_bar_mm) if abs(edge_to_bar_mm) > 1.0 else 20.0
        px_margin = draw_w / ss_f * edge_to_bar_mm * 0.55

        sel_ov_a = u""
        sel_ov_b = u""
        try:
            if self._stirrup_combo_a and self._stirrup_combo_a.SelectedIndex >= 0:
                sel_ov_a = str(self._stirrup_combo_a.SelectedItem or u"")
        except Exception:
            pass
        try:
            if self._stirrup_combo_b and self._stirrup_combo_b.SelectedIndex >= 0:
                sel_ov_b = str(self._stirrup_combo_b.SelectedItem or u"")
        except Exception:
            pass

        try:
            from column_stirrup_creator import build_stirrup_rect_and_tie_defs

            rect_defs_ov, tie_defs_ov = build_stirrup_rect_and_tie_defs(
                ba, bb, sel_ov_a, sel_ov_b
            )
        except Exception:
            rect_defs_ov = []
            tie_defs_ov = []

        sb = _preview_plan_steel_bbox_px(
            nom_left,
            nom_top,
            draw_w,
            draw_h,
            ss_f,
            sl_f,
            ba,
            bb,
            offs_a,
            offs_b,
            inner_on,
            rect_defs_ov,
            tie_defs_ov,
            px_margin,
        )
        gap_concrete = 4.0
        if sb is not None:
            bul, but, bur, bub = sb
            left_c, top_c, wc, hc = _preview_plan_concrete_rect_around_steel(
                nom_left, nom_top, nom_w, nom_h, bul, but, bur, bub, gap_concrete
            )
        else:
            left_c, top_c, wc, hc = nom_left, nom_top, nom_w, nom_h

        rect = Rectangle()
        rect.Width = wc
        rect.Height = hc
        rect.Stroke = br["rect"]
        rect.StrokeThickness = 2.0
        rect.Fill = Brushes.Transparent
        Canvas.SetLeft(rect, left_c)
        Canvas.SetTop(rect, top_c)
        cv.Children.Add(rect)

        seen = set()

        def add_dot(da, db, is_outer):
            key = (round(float(da), 4), round(float(db), 4))
            if key in seen:
                return
            seen.add(key)
            px, py = to_px(da, db)
            el = Ellipse()
            if is_outer:
                el.Width = _PREVIEW_BAR_DOT_OUTER_PX
                el.Height = _PREVIEW_BAR_DOT_OUTER_PX
                el.Fill = br["outer"]
            else:
                el.Width = _PREVIEW_BAR_DOT_INNER_PX
                el.Height = _PREVIEW_BAR_DOT_INNER_PX
                el.Fill = br["inner"]
            Canvas.SetLeft(el, px - el.Width / 2.0)
            Canvas.SetTop(el, py - el.Height / 2.0)
            cv.Children.Add(el)

        for ix, iy in _preview_outer_ij(ba, bb):
            add_dot(offs_a[ix], offs_b[iy], True)
        if inner_on:
            for ix, iy in _preview_inner_ij(ba, bb):
                add_dot(offs_a[ix], offs_b[iy], False)

        tb_a = TextBlock()
        tb_a.Text = u"A"
        tb_a.FontSize = 10.0
        tb_a.Foreground = br["muted"]
        tb_a.FontWeight = FontWeights.SemiBold
        tb_a.TextAlignment = TextAlignment.Center
        tb_a.Width = wc
        Canvas.SetLeft(tb_a, left_c)
        Canvas.SetTop(tb_a, top_c + hc + 6.0)
        cv.Children.Add(tb_a)

        tb_b = TextBlock()
        tb_b.Text = u"B"
        tb_b.FontSize = 10.0
        tb_b.Foreground = br["muted"]
        tb_b.FontWeight = FontWeights.SemiBold
        try:
            tb_b.LayoutTransform = RotateTransform(-90.0)
            tb_b.Measure(Size(Double.PositiveInfinity, Double.PositiveInfinity))
            bb_w = float(tb_b.DesiredSize.Width)
            bb_h = float(tb_b.DesiredSize.Height)
        except Exception:
            bb_w, bb_h = 14.0, 14.0
        pad_b = 6.0
        Canvas.SetLeft(tb_b, left_c + wc + pad_b)
        Canvas.SetTop(tb_b, top_c + max((hc - bb_h) / 2.0, 0.0))
        cv.Children.Add(tb_b)

        self._draw_stirrup_overlay(
            cv,
            offs_a,
            offs_b,
            ba,
            bb,
            ss_f,
            sl_f,
            left,
            top,
            draw_w,
            draw_h,
            inner_on,
        )
        self._draw_plan_preview_legend(cv, cw, ch, left_c, top_c, wc, hc, margin)
        self._refresh_inner_bars_hint()

    def _draw_plan_preview_legend(
        self, cv, cw, ch, scheme_left, scheme_top, scheme_w, scheme_h, margin
    ):
        u"""Leyenda de colores (estribos y traba) en la franja inferior del mismo canvas, bajo el esquema."""
        try:
            items = (
                (0, 180, 80, u"Estribo Interior"),
                (140, 140, 140, u"Estribo Perimetral"),
                (220, 60, 60, u"Traba"),
            )
            legend_font = 11.0
            gap_swatch = 8.0
            gap_between = 22.0
            sw, sh = 12.0, 12.0
            pad_bg = 8.0
            # Texto más claro que br["muted"] para contraste sobre fondo oscuro
            fg = SolidColorBrush(Color.FromRgb(226, 232, 240))

            line_h = max(sh, legend_font * 1.35)

            scheme_bottom = scheme_top + scheme_h
            # Banda reservada bajo el rectángulo de sección: etiqueta «A» + aire antes de la leyenda
            a_band = 6.0 + 18.0 + 6.0
            min_content_y_below = scheme_bottom + a_band + 6.0

            sizes = []
            total_w = 0.0
            for r, g, b, label in items:
                tb_m = TextBlock()
                tb_m.Text = label
                tb_m.FontSize = legend_font
                tb_m.Foreground = fg
                tb_m.FontWeight = FontWeights.Medium
                tb_m.Measure(Size(Double.PositiveInfinity, Double.PositiveInfinity))
                tw = float(tb_m.DesiredSize.Width)
                w_item = sw + gap_swatch + tw
                sizes.append((r, g, b, label, tw))
                total_w += w_item + gap_between
            if total_w > gap_between:
                total_w -= gap_between

            use_vertical = total_w > cw - 24.0

            def _legend_row_y(legend_outer_h):
                u"""Siempre bajo el rectángulo de sección; el preview reserva altura inferior en el canvas."""
                y = min_content_y_below
                if y + legend_outer_h > ch - margin:
                    y_alt = ch - margin - legend_outer_h
                    if y_alt >= min_content_y_below:
                        y = y_alt
                return y

            if use_vertical:
                block_w = 0.0
                for r, g, b, label, tw in sizes:
                    block_w = max(block_w, sw + gap_swatch + tw)
                n = len(sizes)
                row_gap = 8.0
                block_h = n * line_h + (n - 1) * row_gap
                legend_outer_h = block_h + 2.0 * pad_bg
                row_y = _legend_row_y(legend_outer_h)
                x0 = max(8.0, scheme_left + (scheme_w - block_w) / 2.0)

                bg = Rectangle()
                bg.Width = block_w + 2.0 * pad_bg
                bg.Height = legend_outer_h
                bg.RadiusX = 5.0
                bg.RadiusY = 5.0
                bg.Fill = SolidColorBrush(Color.FromArgb(242, 15, 23, 42))
                bg.Stroke = SolidColorBrush(Color.FromArgb(200, 51, 65, 85))
                bg.StrokeThickness = 1.0
                Canvas.SetLeft(bg, x0 - pad_bg)
                Canvas.SetTop(bg, row_y - pad_bg)
                cv.Children.Add(bg)

                y = row_y
                for idx_item, (r, g, b, label, tw) in enumerate(sizes):
                    sq = Rectangle()
                    sq.Width = sw
                    sq.Height = sh
                    sq.Fill = SolidColorBrush(Color.FromRgb(r, g, b))
                    sq.Stroke = SolidColorBrush(Color.FromRgb(min(255, r + 30), min(255, g + 30), min(255, b + 30)))
                    sq.StrokeThickness = 1.0
                    Canvas.SetLeft(sq, x0)
                    Canvas.SetTop(sq, y + (line_h - sh) / 2.0)
                    cv.Children.Add(sq)

                    tb = TextBlock()
                    tb.Text = label
                    tb.FontSize = legend_font
                    tb.FontWeight = FontWeights.Medium
                    tb.Foreground = fg
                    Canvas.SetLeft(tb, x0 + sw + gap_swatch)
                    Canvas.SetTop(tb, y + (line_h - legend_font) / 2.0)
                    cv.Children.Add(tb)

                    y += line_h
                    if idx_item < n - 1:
                        y += row_gap
                return

            legend_row_h = line_h
            legend_outer_h = legend_row_h + 2.0 * pad_bg
            row_y = _legend_row_y(legend_outer_h)

            x0 = max(8.0, scheme_left + (scheme_w - total_w) / 2.0)

            bg = Rectangle()
            bg.Width = total_w + 2.0 * pad_bg
            bg.Height = legend_outer_h
            bg.RadiusX = 5.0
            bg.RadiusY = 5.0
            bg.Fill = SolidColorBrush(Color.FromArgb(242, 15, 23, 42))
            bg.Stroke = SolidColorBrush(Color.FromArgb(200, 51, 65, 85))
            bg.StrokeThickness = 1.0
            Canvas.SetLeft(bg, x0 - pad_bg)
            Canvas.SetTop(bg, row_y - pad_bg)
            cv.Children.Add(bg)

            x = x0
            for r, g, b, label, tw in sizes:
                sq = Rectangle()
                sq.Width = sw
                sq.Height = sh
                sq.Fill = SolidColorBrush(Color.FromRgb(r, g, b))
                sq.Stroke = SolidColorBrush(Color.FromRgb(min(255, r + 30), min(255, g + 30), min(255, b + 30)))
                sq.StrokeThickness = 1.0
                Canvas.SetLeft(sq, x)
                Canvas.SetTop(sq, row_y + (legend_row_h - sh) / 2.0)
                cv.Children.Add(sq)

                tb = TextBlock()
                tb.Text = label
                tb.FontSize = legend_font
                tb.FontWeight = FontWeights.Medium
                tb.Foreground = fg
                Canvas.SetLeft(tb, x + sw + gap_swatch)
                Canvas.SetTop(tb, row_y + (legend_row_h - legend_font) / 2.0)
                cv.Children.Add(tb)

                x += sw + gap_swatch + tw + gap_between
        except Exception:
            pass

    def _draw_stirrup_overlay(
        self,
        cv,
        offs_a,
        offs_b,
        ba,
        bb,
        ss_f,
        sl_f,
        left,
        top,
        draw_w,
        draw_h,
        inner_on,
    ):
        u"""Dibuja estribos, estribos interiores y trabas sobre el canvas de preview."""
        if not offs_a or not offs_b:
            return

        try:
            sel_a = u""
            sel_b = u""
            try:
                if self._stirrup_combo_a and self._stirrup_combo_a.SelectedIndex >= 0:
                    sel_a = str(self._stirrup_combo_a.SelectedItem or u"")
            except Exception:
                pass
            try:
                if self._stirrup_combo_b and self._stirrup_combo_b.SelectedIndex >= 0:
                    sel_b = str(self._stirrup_combo_b.SelectedItem or u"")
            except Exception:
                pass

            from column_stirrup_creator import (
                build_stirrup_rect_and_tie_defs,
                tie_axis_shift_toward_section_center,
            )
            rect_defs, tie_defs = build_stirrup_rect_and_tie_defs(ba, bb, sel_a, sel_b)

            hs = ss_f / 2.0
            hl = sl_f / 2.0

            def to_px(da, db):
                x = left + (float(da) + hs) / ss_f * draw_w
                y = top + (float(db) + hl) / sl_f * draw_h
                return x, y

            def _br(r, g, b):
                return SolidColorBrush(Color.FromRgb(r, g, b))

            # Margin in px = half the distance from edge to first bar
            edge_to_bar_mm = offs_a[-1] - ss_f / 2.0 if offs_a[-1] < ss_f / 2.0 else ss_f / 2.0 - offs_a[0]
            edge_to_bar_mm = abs(edge_to_bar_mm) if abs(edge_to_bar_mm) > 1.0 else 20.0
            px_margin = draw_w / ss_f * edge_to_bar_mm * 0.55
            # Misma representación de barra que add_dot: tangencia con diámetro fijo en px.
            px_margin_draw = px_margin + 0.5 * float(_PREVIEW_BAR_DOT_OUTER_PX)
            stir_thick = 2.2

            def _preview_stirrup_corner_radius_px(a0, a1, b0, b1, rw, rh):
                u"""Radio de esquina = media de los radios de barra en las 4 esquinas; acotado al rectángulo."""
                rc = [
                    _preview_bar_corner_radius_px(a0, b0, ba, bb, inner_on),
                    _preview_bar_corner_radius_px(a1, b0, ba, bb, inner_on),
                    _preview_bar_corner_radius_px(a1, b1, ba, bb, inner_on),
                    _preview_bar_corner_radius_px(a0, b1, ba, bb, inner_on),
                ]
                cr = sum(rc) / float(len(rc))
                w = max(1.0, float(rw))
                h = max(1.0, float(rh))
                m = min(w, h)
                cr = min(cr, 0.45 * m, w * 0.5 - 1.0, h * 0.5 - 1.0)
                return max(0.0, cr)

            for idx_a, idx_b, sp_a, sp_b in rect_defs:
                a0 = max(0, min(idx_a, len(offs_a) - 1))
                a1 = max(0, min(idx_a + sp_a, len(offs_a) - 1))
                b0 = max(0, min(idx_b, len(offs_b) - 1))
                b1 = max(0, min(idx_b + sp_b, len(offs_b) - 1))
                x0, y0 = to_px(offs_a[a0], offs_b[b0])
                x1, y1 = to_px(offs_a[a1], offs_b[b1])
                rx = min(x0, x1) - px_margin_draw
                ry = min(y0, y1) - px_margin_draw
                rw = abs(x1 - x0) + 2.0 * px_margin_draw
                rh = abs(y1 - y0) + 2.0 * px_margin_draw
                is_outer = (idx_a == 0 and idx_b == 0 and sp_a == ba - 1 and sp_b == bb - 1)
                color = _br(140, 140, 140) if is_outer else _br(0, 180, 80)
                r_el = Rectangle()
                r_el.Width = max(1.0, rw)
                r_el.Height = max(1.0, rh)
                try:
                    _cr = _preview_stirrup_corner_radius_px(a0, a1, b0, b1, r_el.Width, r_el.Height)
                    if _cr > 0.01:
                        r_el.RadiusX = _cr
                        r_el.RadiusY = _cr
                except Exception:
                    pass
                r_el.Stroke = color
                r_el.StrokeThickness = stir_thick
                r_el.Fill = Brushes.Transparent
                Canvas.SetLeft(r_el, rx)
                Canvas.SetTop(r_el, ry)
                cv.Children.Add(r_el)

            if offs_a and offs_b:
                x_edge0, y_edge0 = to_px(offs_a[0], offs_b[0])
                x_edge1, y_edge1 = to_px(offs_a[-1], offs_b[-1])
                x_mid_px, y_mid_px = to_px(0.0, 0.0)

                for idx, is_a in tie_defs:
                    tie_ln = WpfLine()
                    tie_ln.Stroke = _br(220, 60, 60)
                    tie_ln.StrokeThickness = 2.2
                    if is_a:
                        if 0 <= idx < len(offs_a):
                            r0 = float(
                                _preview_bar_corner_radius_px(idx, 0, ba, bb, inner_on)
                            )
                            r1 = float(
                                _preview_bar_corner_radius_px(
                                    idx, bb - 1, ba, bb, inner_on
                                )
                            )
                            r_t = max(r0, r1)
                            x_px, _ = to_px(offs_a[idx], 0.0)
                            x_px = x_mid_px + tie_axis_shift_toward_section_center(
                                x_px - x_mid_px,
                                r_t,
                                tie_index=idx,
                                bar_count=ba,
                            )
                            tie_ln.X1 = x_px
                            tie_ln.Y1 = y_edge0 - px_margin
                            tie_ln.X2 = x_px
                            tie_ln.Y2 = y_edge1 + px_margin
                            cv.Children.Add(tie_ln)
                    else:
                        if 0 <= idx < len(offs_b):
                            r0 = float(
                                _preview_bar_corner_radius_px(0, idx, ba, bb, inner_on)
                            )
                            r1 = float(
                                _preview_bar_corner_radius_px(
                                    ba - 1, idx, ba, bb, inner_on
                                )
                            )
                            r_t = max(r0, r1)
                            _, y_px = to_px(0.0, offs_b[idx])
                            y_px = y_mid_px + tie_axis_shift_toward_section_center(
                                y_px - y_mid_px,
                                r_t,
                                tie_index=idx,
                                bar_count=bb,
                            )
                            tie_ln.X1 = x_edge0 - px_margin
                            tie_ln.Y1 = y_px
                            tie_ln.X2 = x_edge1 + px_margin
                            tie_ln.Y2 = y_px
                            cv.Children.Add(tie_ln)
        except Exception:
            pass

    def _on_window_closed(self, sender, args):
        try:
            AppDomain.CurrentDomain.SetData(_WIZARD_SINGLETON_KEY, None)
        except Exception:
            pass

    def _update_step_badges(self, step):
        active_bg = "#164e63"
        active_bd = "#22d3ee"
        active_fg = "#E8F4F8"
        idle_bg = "#0f172a"
        idle_bd = "#475569"
        idle_fg = "#94a3b8"
        try:
            from System.Windows.Media import SolidColorBrush
            from System.Windows.Media import Color

            def _br(hex_str):
                h = hex_str.lstrip("#")
                return SolidColorBrush(
                    Color.FromRgb(
                        int(h[0:2], 16),
                        int(h[2:4], 16),
                        int(h[4:6], 16),
                    )
                )

            if int(step) == 1:
                self._step1_badge.Background = _br(active_bg)
                self._step1_badge.BorderBrush = _br(active_bd)
                self._step1_badge_text.Foreground = _br(active_fg)
                self._step2_badge.Background = _br(idle_bg)
                self._step2_badge.BorderBrush = _br(idle_bd)
                self._step2_badge_text.Foreground = _br(idle_fg)
            else:
                self._step1_badge.Background = _br(idle_bg)
                self._step1_badge.BorderBrush = _br(idle_bd)
                self._step1_badge_text.Foreground = _br(idle_fg)
                self._step2_badge.Background = _br(active_bg)
                self._step2_badge.BorderBrush = _br(active_bd)
                self._step2_badge_text.Foreground = _br(active_fg)
            if self._tb_step_progress is not None:
                self._tb_step_progress.Text = u"Paso {0} de 2".format(int(step))
        except Exception:
            pass

    def _attach_owner(self):
        try:
            from System.Windows.Interop import WindowInteropHelper

            from revit_wpf_window_position import (
                position_wpf_window_center_work_area,
                revit_main_hwnd,
            )

            hwnd = revit_main_hwnd(self._uiapp)
            if hwnd:
                WindowInteropHelper(self.window).Owner = hwnd
            position_wpf_window_center_work_area(self.window)
        except Exception:
            try:
                from revit_wpf_window_position import (
                    position_wpf_window_center_work_area,
                )

                position_wpf_window_center_work_area(self.window)
            except Exception:
                try:
                    from System.Windows import WindowStartupLocation

                    self.window.WindowStartupLocation = (
                        WindowStartupLocation.CenterScreen
                    )
                except Exception:
                    pass

    def _load_current_section_into_ui(self):
        if not self._ordered_keys:
            return
        sk = self._ordered_keys[self._section_idx]
        title = self._section_meta[self._section_idx][1]
        try:
            self._tb_title.Text = title
        except Exception:
            pass
        cfg = self._section_grid_config.get(sk)
        try:
            if cfg is not None:
                a = max(
                    _BARS_COUNT_MIN,
                    min(_BARS_COUNT_MAX, int(cfg["bars_a"])),
                )
                b = max(
                    _BARS_COUNT_MIN,
                    min(_BARS_COUNT_MAX, int(cfg["bars_b"])),
                )
                self._apply_bars_sliders(a, b)
                self._chk_inner.IsChecked = bool(cfg["include_inner_outline"])
            else:
                s0, L0 = int(sk[0]), int(sk[1])
                if s0 == L0:
                    self._apply_bars_sliders(4, 4)
                else:
                    self._apply_bars_sliders(4, 6)
                self._chk_inner.IsChecked = False
        except Exception:
            pass
        self._refresh_section_dims_label()
        self._load_stirrup_for_section(sk)
        self._refresh_grid_preview()
        self._refresh_section_footer_hints()
        self._apply_rejilla_window_layout()

    def _refresh_nav_buttons(self):
        if self._btn_prev is None:
            return
        try:
            self._btn_prev.Visibility = (
                self._v.Visible
                if self._section_idx > 0
                else self._v.Collapsed
            )
        except Exception:
            pass

    def _refresh_grid_next_label(self):
        if self._btn_next is None or not self._ordered_keys:
            return
        last = self._section_idx >= len(self._ordered_keys) - 1
        try:
            self._btn_next.Content = (
                u"Siguiente: Esquema Vertical →"
                if last
                else u"Siguiente sección →"
            )
        except Exception:
            pass

    def _commit_current_section(self):
        if not self._ordered_keys:
            return False
        sk = self._ordered_keys[self._section_idx]
        a, b = self._preview_bars_ab()
        if (
            a < _BARS_COUNT_MIN
            or a > _BARS_COUNT_MAX
            or b < _BARS_COUNT_MIN
            or b > _BARS_COUNT_MAX
        ):
            if TaskDialog is not None:
                try:
                    TaskDialog.Show(
                        u"Arainco: Armado Columnas",
                        u"Las cantidades A y B deben estar entre {0} y {1}.".format(
                            _BARS_COUNT_MIN,
                            _BARS_COUNT_MAX,
                        ),
                    )
                except Exception:
                    pass
            return False
        try:
            inner = bool(self._chk_inner.IsChecked)
        except Exception:
            inner = False
        self._section_grid_config[sk] = dict(
            bars_a=int(a),
            bars_b=int(b),
            include_inner_outline=inner,
        )
        self._commit_stirrup_for_section(sk)
        return True

    def _on_grid_prev(self, sender, args):
        if not self._commit_current_section():
            return
        if self._section_idx > 0:
            self._section_idx -= 1
            self._load_current_section_into_ui()
            self._refresh_nav_buttons()
            self._refresh_grid_next_label()

    def _on_grid_next(self, sender, args):
        if not self._commit_current_section():
            return
        if self._section_idx < len(self._ordered_keys) - 1:
            self._section_idx += 1
            self._load_current_section_into_ui()
            self._refresh_nav_buttons()
            self._refresh_grid_next_label()
        else:
            self._goto_troceo_step()

    def _apply_troceo_preferred_window_height(self):
        u"""Paso esquema vertical: ventana del asistente maximizada."""
        win = self.window
        if win is None:
            return
        grid = self._wizard_root_grid
        try:
            if grid is not None and grid.RowDefinitions.Count > 2:
                grid.RowDefinitions[2].Height = GridLength(1.0, GridUnitType.Star)
        except Exception:
            pass
        try:
            from column_reinforcement.ui.troceo_scheme_window import (
                apply_troceo_scheme_window_maximized,
            )
        except Exception:
            try:
                from troceo_scheme_window import apply_troceo_scheme_window_maximized
            except Exception:
                apply_troceo_scheme_window_maximized = None
        if apply_troceo_scheme_window_maximized is not None:
            apply_troceo_scheme_window_maximized(win)

    def _on_cancel_step1(self, sender, args):
        try:
            self.window.DialogResult = False
        except Exception:
            pass
        try:
            self.window.Close()
        except Exception:
            pass

    def _goto_troceo_step(self):
        if not self._troceo_rows:
            self._troceo_outcome = TroceoSchemeOutcome(
                skip_no_cut=True,
                columns=[],
                segment_rebar_bar_type_ids=None,
            )
            try:
                self.window.DialogResult = True
            except Exception:
                pass
            try:
                self.window.Close()
            except Exception:
                pass
            return

        try:
            self._scroll_grid.Visibility = self._v.Collapsed
            self._scroll_troceo.Visibility = self._v.Visible
            self._footer1.Visibility = self._v.Collapsed
            self._footer2.Visibility = self._v.Visible
        except Exception:
            pass
        self._apply_troceo_preferred_window_height()
        self._update_step_badges(2)

        _bar_choices = self._stirrup_bar_type_choices_for_troceo_scheme()
        _ubic_line_labels = []
        _ubic_scheme_by_label = {}
        try:
            from column_reinforcement_layout_rps import (
                troceo_fused_line_labels_and_scheme_map,
            )

            _meta_sorted = []
            for _row in self._troceo_rows or []:
                _elem, _z, _eid, _h, _lv, _pl = _unpack_troceo_row(_row)
                if _eid < 0 or _elem is None:
                    continue
                _meta_sorted.append((_elem, float(_z), int(_eid)))
            _meta_sorted.sort(key=lambda t: t[1])
            _cols_fuse = [t[0] for t in _meta_sorted]
            _bt_by_eid = {}
            for _slot_i, (_el, _zmm, eid) in enumerate(_meta_sorted):
                try:
                    _bt = self._stirrup_bar_type_by_troceo_slot.get(int(_slot_i))
                except Exception:
                    _bt = None
                if _bt is not None and eid >= 0:
                    _bt_by_eid[int(eid)] = _bt
            _ubic_line_labels, _ubic_scheme_by_label = troceo_fused_line_labels_and_scheme_map(
                self._doc,
                _cols_fuse,
                dict(self._section_grid_config),
                dict(self._stirrup_section_data),
                _bt_by_eid,
                float(self._default_bar_diam_mm),
            )
            _ubic_line_labels = _ubic_line_labels or []
            _ubic_scheme_by_label = _ubic_scheme_by_label or {}
        except Exception:
            _ubic_line_labels = []
            _ubic_scheme_by_label = {}

        self._troceo_ctrl = TroceoSchemeController(
            self._troceo_rows,
            uiapp=self._uiapp,
            uidoc=self._uidoc,
            doc=self._doc,
            default_bar_diam_mm=self._default_bar_diam_mm,
            parent_window=self.window,
            blocks_host=self.window.FindName("TroceoBlocksHost"),
            diam_host=None,
            btn_confirm=self._btn_troceo_ok,
            btn_cancel=self._btn_troceo_cancel,
            btn_alternate_sel=self.window.FindName("BtnTroceoAlternateSel"),
            scheme_scrollviewer=self.window.FindName("TroceoSchemeScroll"),
            diam_scrollviewer=None,
            embed_notify=self._on_troceo_embed,
            tb_alternate_start=self.window.FindName("TbTroceoAltStart"),
            tb_alternate_step=self.window.FindName("TbTroceoAltStep"),
            column_stirrup_spacing_slot_store=self._stirrup_spacing_by_troceo_slot,
            column_stirrup_spacing_default_cb=self._default_stirrup_spacing_mm_for_column_elem,
            column_stirrup_bar_type_slot_store=(
                self._stirrup_bar_type_by_troceo_slot
                if _bar_choices
                else None
            ),
            column_stirrup_bar_type_choices=_bar_choices,
            column_stirrup_bar_type_default_cb=(
                self._default_stirrup_bar_type_for_column_elem
                if _bar_choices
                else None
            ),
            column_stirrup_policy_slot_store=self._stirrup_policy_by_troceo_slot,
            longitudinal_line_ubicacion_labels=_ubic_line_labels,
            longitudinal_line_scheme_by_label=_ubic_scheme_by_label,
        )

    def _on_troceo_embed(self, action):
        if action == "cancel":
            self._troceo_outcome = TroceoSchemeOutcome(cancelled=True)
            try:
                self.window.DialogResult = False
            except Exception:
                pass
        elif action == "confirm":
            self._troceo_outcome = self._troceo_ctrl.build_outcome_after_confirm()
            try:
                self.window.DialogResult = True
            except Exception:
                pass
        try:
            self.window.Close()
        except Exception:
            pass

    def _troceo_policy_store(self):
        ctrl = self._troceo_ctrl
        if ctrl is None:
            return self._stirrup_policy_by_troceo_slot
        return getattr(ctrl, "_col_stirrup_policy_store", None) or self._stirrup_policy_by_troceo_slot

    def _collect_stirrup_spacing_by_column_lot(self):
        u"""Espaciamiento por columna y lote (0=T1 … 2=T3)."""
        try:
            from column_reinforcement.ui.troceo_scheme_window import (
                STIRRUP_POLICY_THIRDS_L3,
                _stirrup_policy_for_slot,
                _stirrup_slot_lot_store_key,
            )
        except Exception:
            from troceo_scheme_window import (
                STIRRUP_POLICY_THIRDS_L3,
                _stirrup_policy_for_slot,
                _stirrup_slot_lot_store_key,
            )
        try:
            from column_reinforcement_layout_rps import _element_id_iv
        except Exception:
            _element_id_iv = None
        ctrl = self._troceo_ctrl
        slot_store = self._stirrup_spacing_by_troceo_slot
        policy_store = self._troceo_policy_store()
        by_col = {}
        if ctrl is None or not slot_store:
            return by_col
        entries = getattr(ctrl, "_row_entries", None) or []
        for _elem, _z, eid, slot in entries:
            try:
                s = int(slot)
                pol = _stirrup_policy_for_slot(policy_store, s)
                lots = (
                    [0, 1, 2]
                    if pol == STIRRUP_POLICY_THIRDS_L3
                    else [0]
                )
                col_key = None
                if _element_id_iv is not None and _elem is not None:
                    col_key = int(_element_id_iv(_elem))
                if col_key is None or col_key < 0:
                    col_key = int(eid)
                if col_key < 0:
                    continue
                for lot_i in lots:
                    sk = _stirrup_slot_lot_store_key(s, lot_i)
                    if sk not in slot_store and lot_i == 0 and s in slot_store:
                        mm = float(slot_store[s])
                    elif sk in slot_store:
                        mm = float(slot_store[sk])
                    else:
                        continue
                    if col_key not in by_col:
                        by_col[col_key] = {}
                    by_col[col_key][int(lot_i)] = mm
            except Exception:
                pass
        return by_col

    def _collect_stirrup_bar_type_by_column_lot(self):
        u"""RebarBarType por columna y lote."""
        try:
            from column_reinforcement.ui.troceo_scheme_window import (
                STIRRUP_POLICY_THIRDS_L3,
                _stirrup_policy_for_slot,
                _stirrup_slot_lot_store_key,
            )
        except Exception:
            from troceo_scheme_window import (
                STIRRUP_POLICY_THIRDS_L3,
                _stirrup_policy_for_slot,
                _stirrup_slot_lot_store_key,
            )
        try:
            from column_reinforcement_layout_rps import _element_id_iv
        except Exception:
            _element_id_iv = None
        ctrl = self._troceo_ctrl
        slot_store = self._stirrup_bar_type_by_troceo_slot
        policy_store = self._troceo_policy_store()
        by_col = {}
        if ctrl is None or not slot_store:
            return by_col
        entries = getattr(ctrl, "_row_entries", None) or []
        for _elem, _z, eid, slot in entries:
            try:
                s = int(slot)
                pol = _stirrup_policy_for_slot(policy_store, s)
                lots = (
                    [0, 1, 2]
                    if pol == STIRRUP_POLICY_THIRDS_L3
                    else [0]
                )
                col_key = None
                if _element_id_iv is not None and _elem is not None:
                    col_key = int(_element_id_iv(_elem))
                if col_key is None or col_key < 0:
                    col_key = int(eid)
                if col_key < 0:
                    continue
                for lot_i in lots:
                    sk = _stirrup_slot_lot_store_key(s, lot_i)
                    bt = None
                    if sk in slot_store:
                        bt = slot_store[sk]
                    elif lot_i == 0 and s in slot_store:
                        bt = slot_store[s]
                    if bt is None:
                        continue
                    if col_key not in by_col:
                        by_col[col_key] = {}
                    by_col[col_key][int(lot_i)] = bt
            except Exception:
                pass
        return by_col

    def _collect_stirrup_policy_by_column_id(self):
        try:
            from column_reinforcement.ui.troceo_scheme_window import (
                STIRRUP_POLICY_CONTINUOUS,
                _stirrup_policy_for_slot,
            )
        except Exception:
            from troceo_scheme_window import (
                STIRRUP_POLICY_CONTINUOUS,
                _stirrup_policy_for_slot,
            )
        try:
            from column_reinforcement_layout_rps import _element_id_iv
        except Exception:
            _element_id_iv = None
        ctrl = self._troceo_ctrl
        policy_store = self._troceo_policy_store()
        by_col = {}
        if ctrl is None:
            return by_col
        entries = getattr(ctrl, "_row_entries", None) or []
        for _elem, _z, eid, slot in entries:
            try:
                s = int(slot)
                pol = _stirrup_policy_for_slot(policy_store, s)
                if pol == STIRRUP_POLICY_CONTINUOUS:
                    continue
                col_key = None
                if _element_id_iv is not None and _elem is not None:
                    col_key = int(_element_id_iv(_elem))
                if col_key is None or col_key < 0:
                    col_key = int(eid)
                if col_key >= 0:
                    by_col[col_key] = pol
            except Exception:
                pass
        return by_col

    def _collect_stirrup_spacing_by_column_id(self):
        u"""Mapa plano (lote 0 / Completo) para compatibilidad."""
        lot_map = self._collect_stirrup_spacing_by_column_lot()
        by_col = {}
        for col_key, lots in (lot_map or {}).items():
            try:
                if 0 in lots:
                    by_col[int(col_key)] = float(lots[0])
            except Exception:
                pass
        return by_col

    def _collect_stirrup_bar_type_by_column_id(self):
        u"""Mapa plano (lote 0 / Completo) para compatibilidad."""
        lot_map = self._collect_stirrup_bar_type_by_column_lot()
        by_col = {}
        for col_key, lots in (lot_map or {}).items():
            try:
                if 0 in lots:
                    by_col[int(col_key)] = lots[0]
            except Exception:
                pass
        return by_col

    def show_dialog(self):
        try:
            AppDomain.CurrentDomain.SetData(_WIZARD_SINGLETON_KEY, self.window)
        except Exception:
            pass
        self._attach_owner()
        try:
            self.window.Activate()
        except Exception:
            pass
        ok = self.window.ShowDialog()
        if not ok:
            return ColumnLayoutWizardOutcome(
                cancelled=True,
                section_grid_config=dict(self._section_grid_config),
                troceo_outcome=self._troceo_outcome,
                stirrup_configs=dict(self._stirrup_section_data),
                stirrup_spacing_by_column_id={},
                stirrup_bar_type_by_column_id={},
                stirrup_policy_by_column_id={},
                stirrup_spacing_by_column_lot={},
                stirrup_bar_type_by_column_lot={},
            )
        return ColumnLayoutWizardOutcome(
            cancelled=False,
            section_grid_config=dict(self._section_grid_config),
            troceo_outcome=self._troceo_outcome,
            stirrup_configs=dict(self._stirrup_section_data),
            stirrup_spacing_by_column_id=self._collect_stirrup_spacing_by_column_id(),
            stirrup_bar_type_by_column_id=self._collect_stirrup_bar_type_by_column_id(),
            stirrup_policy_by_column_id=self._collect_stirrup_policy_by_column_id(),
            stirrup_spacing_by_column_lot=self._collect_stirrup_spacing_by_column_lot(),
            stirrup_bar_type_by_column_lot=self._collect_stirrup_bar_type_by_column_lot(),
        )


def show_column_layout_wizard_singleton(
    section_meta,
    troceo_rows,
    uiapp,
    uidoc,
    doc,
    default_bar_diam_mm,
):
    """
    Muestra el asistente o enfoca la instancia existente.

    ``section_meta``: ``list`` de ``(section_key, título)``.
    """
    try:
        existing = AppDomain.CurrentDomain.GetData(_WIZARD_SINGLETON_KEY)
        if existing is not None:
            try:
                loaded = getattr(existing, "IsLoaded", False)
            except Exception:
                loaded = False
            if not loaded:
                try:
                    AppDomain.CurrentDomain.SetData(_WIZARD_SINGLETON_KEY, None)
                except Exception:
                    pass
                existing = None
        if existing is not None:
            try:
                existing.Activate()
                existing.Focus()
            except Exception:
                pass
            if TaskDialog is not None:
                try:
                    TaskDialog.Show(
                        u"Arainco: Armado Columnas",
                        u"La herramienta ya esta en ejecucion.",
                    )
                except Exception:
                    pass
            return ColumnLayoutWizardOutcome(already_running=True)
    except Exception:
        pass

    ctrl = ColumnLayoutWizardController(
        section_meta,
        troceo_rows,
        uiapp,
        uidoc,
        doc,
        default_bar_diam_mm=float(default_bar_diam_mm),
    )
    return ctrl.show_dialog()
