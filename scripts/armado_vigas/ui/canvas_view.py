# -*- coding: utf-8 -*-
"""Renderizado WPF del canvas (alzado, estribos, tramos Tn, labels)."""

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

import System

from System.Windows import HorizontalAlignment, TextWrapping, Thickness, VerticalAlignment, FontWeights, TextAlignment
from System.Windows.Controls import (
    Border,
    Button,
    Canvas,
    ColumnDefinition,
    Grid,
    Orientation,
    RowDefinition,
    ScrollBarVisibility,
    ScrollViewer,
    StackPanel,
    TextBlock,
)
from System.Windows import RoutedEventHandler
from System.Windows import GridLength, GridUnitType
from System.Windows.Input import Cursors, Keyboard, ModifierKeys, MouseButtonEventHandler
from System.Windows.Media import DoubleCollection, SolidColorBrush, Color
from System.Windows import Point
from System.Windows.Media import PointCollection
from System.Windows.Shapes import Line, Polyline, Rectangle

from armado_vigas.domain.confinement import ensure_beam_confinement, get_confin_scenarios
from armado_vigas.domain.constants import (
    BAR_COUNT_MIN,
    BAR_COUNT_MAX,
    CAPAS_MAX,
    ESTRIBO_SPACING_DEFAULT_CENT,
    ESTRIBO_SPACING_DEFAULT_EXT,
    LONG_DIAM_OPTS,
)
from armado_vigas.domain.laterales import (
    LATERALES_COUNT_MAX,
    LATERALES_COUNT_MIN,
    suggest_n_laterales_from_beams,
)
from armado_vigas.domain.layers import (
    beam_n_capas_inf,
    beam_n_capas_sup,
    clamp_bar_count,
    ensure_beam_layers,
    first_layer_bar_count,
    is_global_layer_sync_field,
    layer_bar_count,
    layer_keys,
    set_first_layer_bar_count,
    sync_layer_field_all_beams,
)
from armado_vigas.domain.suple_inferior import (
    beam_suple_inf_enabled,
    beam_suple_layer_index,
    ensure_beam_suple_inferior,
    suple_metrics_mm,
)
from armado_vigas.domain.suple_superior import (
    SUPLE_END_PCT,
    beam_suple_sup_enabled,
    beam_suple_sup_side_enabled,
    ensure_beam_suple_superior,
    suple_sup_segments_layout_px,
)
from armado_vigas.domain.stirrups import compute_stirrup_zones, parse_beam_section, section_height_mm
from armado_vigas.domain.tramos import (
    build_session_tramos,
    find_tramo_for_beam,
    format_dual_tramo_summary,
    sort_beams,
    tramo_exceeds_bar_limit,
)
from armado_vigas.ui import layout as lay
from armado_vigas.ui import typography as typo
from armado_vigas.ui import theme as th
from armado_vigas.ui.theme import apply_panel_chrome, make_role_badge
from armado_vigas.ui.section_preview import draw_section_preview, section_meta_lines
from armado_vigas.ui.wpf_controls import (
    accent_soft_brush,
    brush_hex,
    label_small,
    make_capas_stepper,
    make_diam_combo,
    make_spacing_input,
    make_stepper,
    make_string_combo,
    make_yesno_toggle,
)


ESTRIBO_DIAM_OPTS = (8, 10, 12, 16)

_ZONE_ROLE_STYLE = {
    "ext": (u"#fbbf24", u"#101408", u"#2a1f0a"),
    "cent": (u"#34d399", u"#0a1620", u"#0d2430"),
    "uni": (u"#fde68a", u"#0e1412", u"#1a1810"),
}

# Alzado integrado — escala según layout.ELEVATION_HEIGHT_PX (Opción D)
_ELEV_SCALE = lay.ELEVATION_HEIGHT_PX / 136.0
_ELEV_BEAM_TOP = 28.0 * _ELEV_SCALE
_ELEV_BEAM_H = 50.0 * _ELEV_SCALE
_ELEV_BEAM_BOT = _ELEV_BEAM_TOP + _ELEV_BEAM_H
_ELEV_BAR_INSET = 8.0
_ELEV_BAR_SUP_Y = _ELEV_BEAM_TOP + 9.0 * _ELEV_SCALE
_ELEV_BAR_INF_Y = _ELEV_BEAM_BOT - 9.0 * _ELEV_SCALE
_ELEV_BAR_SUPLE_SUP_Y = _ELEV_BAR_SUP_Y + 14.0 * _ELEV_SCALE
_ELEV_BAR_SUPLE_Y = _ELEV_BAR_INF_Y + 11.0 * _ELEV_SCALE
_ELEV_COL_TOP = 4.0
_ELEV_COL_H = lay.ELEVATION_HEIGHT_PX - 10.0
_ELEV_COL_W = 14.0
_ELEV_WALL_W = 18.0
_ELEV_POCKET_D = 12.0
_ELEV_STROKE_CHORD = 2.0
_ELEV_STROKE_BAR = 2.5
_ELEV_DIR_MARKER_LEN_PX = 48.0
_ELEV_REF_SECTION_H_CM = 60.0
_ELEV_COL_PX_PER_MM = _ELEV_COL_W / 300.0
_ELEV_WALL_PX_PER_MM = _ELEV_WALL_W / 200.0
_ELEV_BREAK_AMP = 3.5 * _ELEV_SCALE
_ELEV_SUPPORT_STROKE = 1.2
_ELEV_BEAM_CHORD = brush_hex(u"#e2e8f0", 230)


def _session_bar_diam_opts(session, current_mm=None):
    opts = getattr(session, "bar_diameters_mm", None) or LONG_DIAM_OPTS
    if current_mm is None:
        return opts
    try:
        cur = int(round(float(current_mm)))
    except Exception:
        return opts
    if cur in opts:
        return opts
    return tuple(sorted(set(list(opts) + [cur])))


class ArmadoVigasCanvasView(object):
    def __init__(self, win, callbacks):
        """
        callbacks: dict con claves
          on_status(msg), on_redraw(), on_toggle_empalme(beam_id, face) — Traslapo sup/inf,
          on_select_tramo(tramo_id), on_select_beam(idx, n_selected),
          on_select_stirrup_zone(idx, role)
        """
        self._win = win
        self._cb = callbacks or {}
        self._host = win.FindName(u"PnlCanvasHost") if win else None
        self._scr = win.FindName(u"ScrCanvas") if win else None
        self._cnv_section = win.FindName(u"CnvSectionPreview") if win else None
        self._txt_section = win.FindName(u"TxtSectionMeta") if win else None
        self._txt_section_rail = win.FindName(u"TxtSectionRailHint") if win else None
        self._pnl_section_ctrls = win.FindName(u"PnlSectionCtrls") if win else None
        self._txt_tramo = win.FindName(u"TxtTramoSummary") if win else None
        self._txt_apoyos = win.FindName(u"TxtApoyosSummary") if win else None
        self._txt_sub = win.FindName(u"TxtSubtitle") if win else None
        self._txt_sel = win.FindName(u"TxtSelectionInfo") if win else None

        self.selected_tramo_sup_id = None
        self.selected_tramo_inf_id = None
        self.selected_beam_idx = -1
        self.selected_beam_indices = set()
        self.selected_stirrup_zone = None
        self._layout_meta = {"contentWidthPx": 640.0, "needsScroll": False}
        self._last_beams = []
        self._last_session = None
        self._drawing = False
        self._pending_redraw = False

    def _is_ctrl_click(self, args):
        try:
            mods = Keyboard.Modifiers
        except Exception:
            try:
                mods = args.KeyboardDevice.Modifiers
            except Exception:
                return False
        return (mods & ModifierKeys.Control) == ModifierKeys.Control

    def _is_beam_selected(self, idx):
        try:
            i = int(idx)
        except Exception:
            return False
        return i in self.selected_beam_indices

    def _normalize_beam_selection(self, beams):
        n = len(beams or [])
        valid = {i for i in self.selected_beam_indices if 0 <= i < n}
        self.selected_beam_indices = valid
        if n == 0:
            self.selected_beam_idx = -1
            self.selected_beam_indices = set()
            return
        if not self.selected_beam_indices:
            if 0 <= self.selected_beam_idx < n:
                self.selected_beam_indices = {self.selected_beam_idx}
            else:
                self.selected_beam_indices = {0}
                self.selected_beam_idx = 0
        elif self.selected_beam_idx not in self.selected_beam_indices:
            self.selected_beam_idx = min(self.selected_beam_indices)

    def _default_stirrup_role(self, beam):
        plan = compute_stirrup_zones(beam)
        role = u"cent"
        if plan.get("mode") == u"single" and plan.get("singleKind") == u"merge":
            role = u"uni"
        return role

    def _handle_beam_select(self, idx, args=None, ctrl=None, role=None, update_zone=True, redraw=True):
        try:
            idx = int(idx)
        except Exception:
            return
        beams = self._last_beams or []
        if idx < 0 or idx >= len(beams):
            return
        if ctrl is None:
            ctrl = self._is_ctrl_click(args) if args is not None else False
        if ctrl:
            if idx in self.selected_beam_indices:
                if len(self.selected_beam_indices) > 1:
                    self.selected_beam_indices.discard(idx)
            else:
                self.selected_beam_indices.add(idx)
            self.selected_beam_idx = idx
        else:
            self.selected_beam_indices = {idx}
            self.selected_beam_idx = idx
        if role is not None:
            self.selected_stirrup_zone = {u"idx": idx, u"role": role}
            self._cb.get("on_select_stirrup_zone", lambda _i, _r: None)(idx, role)
        elif update_zone:
            zone_role = self._default_stirrup_role(beams[idx])
            self.selected_stirrup_zone = {u"idx": idx, u"role": zone_role}
        n_sel = len(self.selected_beam_indices)
        self._cb.get("on_select_beam", lambda _i, _n=1: None)(idx, n_sel)
        if redraw:
            self._cb.get("on_redraw", lambda: None)()

    def _is_section_zone_selected(self, idx, role):
        if not self._is_beam_selected(idx):
            return False
        sz = self.selected_stirrup_zone or {}
        return sz.get("role") == role

    def _targets_for_beam_edit(self, beam):
        beams = self._last_beams or []
        if not beam or not beams:
            return []
        if len(self.selected_beam_indices) <= 1:
            return [beam]
        try:
            idx = beams.index(beam)
        except ValueError:
            return [beam]
        if idx not in self.selected_beam_indices:
            return [beam]
        return [beams[i] for i in sorted(self.selected_beam_indices) if 0 <= i < len(beams)]

    def redraw(self, session):
        if self._host is None:
            return
        if self._drawing:
            self._pending_redraw = True
            return
        self._drawing = True
        try:
            self._redraw_impl(session)
        except Exception as ex:
            self._show_canvas_error(ex)
            raise
        finally:
            self._drawing = False
            if self._pending_redraw:
                self._pending_redraw = False
                self.redraw(session)

    def _redraw_impl(self, session):
        if self._host is None:
            return
        self._host.Children.Clear()
        beams = sort_beams(list(session.domain_beams or []))
        apoyos_loaded = bool(getattr(session, "apoyos_loaded", False))

        if not beams:
            self._show_empty(apoyos_loaded)
            return

        for beam in beams:
            ensure_beam_layers(beam)
            ensure_beam_confinement(beam)
            ensure_beam_suple_inferior(beam)
            ensure_beam_suple_superior(beam)

        self._last_session = session
        viewport_w = self._viewport_width()
        layout_result = lay.compute_layout(
            beams,
            viewport_w,
            apoyos=getattr(session, "apoyos", None) if apoyos_loaded else None,
            use_model_positions=apoyos_loaded,
        )
        layouts = layout_result["layouts"]
        self._layout_meta = layout_result
        tramos_sup = list(getattr(session, "tramos_sup", None) or [])
        tramos_inf = list(getattr(session, "tramos_inf", None) or [])
        if not tramos_sup and not tramos_inf:
            tramos_sup, tramos_inf = build_session_tramos(
                beams,
                empalme_beam_ids_sup=session.empalme_beam_ids_sup,
                empalme_beam_ids_inf=session.empalme_beam_ids_inf,
                split_empalme=session.split_empalme,
            )

        if not self.selected_tramo_sup_id and tramos_sup:
            self.selected_tramo_sup_id = tramos_sup[0]["id"]
        if self.selected_tramo_sup_id and not any(
            t["id"] == self.selected_tramo_sup_id for t in tramos_sup
        ):
            self.selected_tramo_sup_id = tramos_sup[0]["id"] if tramos_sup else None

        if not self.selected_tramo_inf_id and tramos_inf:
            self.selected_tramo_inf_id = tramos_inf[0]["id"]
        if self.selected_tramo_inf_id and not any(
            t["id"] == self.selected_tramo_inf_id for t in tramos_inf
        ):
            self.selected_tramo_inf_id = tramos_inf[0]["id"] if tramos_inf else None

        self._normalize_beam_selection(beams)
        if self.selected_stirrup_zone and self.selected_stirrup_zone.get("idx", -1) >= len(beams):
            self.selected_stirrup_zone = None
        if not self.selected_stirrup_zone and beams:
            self.selected_stirrup_zone = {"idx": 0, "role": "cent"}

        self._update_headers(beams, tramos_sup, tramos_inf, apoyos_loaded, layout_result)

        content_w = float(layout_result["contentWidthPx"])
        root = StackPanel()
        root.Width = content_w
        root.Background = brush_hex(u"#0a1620", 0)

        stack = Border()
        stack.Width = content_w
        stack.BorderBrush = brush_hex(u"#21465C")
        stack.BorderThickness = Thickness(1)
        stack.Background = brush_hex(u"#071018", 0)
        stack_inner = StackPanel()
        stack_inner.Children.Add(
            self._build_suple_sup_zone(beams, layouts, content_w)
        )
        stack_inner.Children.Add(
            self._build_elev_stage_option_d(
                beams, layouts, tramos_sup, tramos_inf, session, apoyos_loaded, content_w,
            )
        )
        stack.Child = stack_inner
        root.Children.Add(stack)

        root.Children.Add(self._build_labels(beams, layouts, session, apoyos_loaded, content_w))
        root.Children.Add(self._build_estribo_zone(beams, layouts, tramos_inf, content_w))
        root.Children.Add(self._build_axis_hint(apoyos_loaded, content_w))

        self._host.Children.Add(root)
        self._last_beams = beams
        self._draw_section_rail(beams)

    def _viewport_width(self):
        try:
            if self._scr is not None and self._scr.ActualWidth > 1.0:
                return float(self._scr.ActualWidth) - 24.0
        except Exception:
            pass
        try:
            if self._host is not None and self._host.ActualWidth > 1.0:
                return float(self._host.ActualWidth) - 24.0
        except Exception:
            pass
        return 640.0

    def _show_empty(self, apoyos_loaded):
        if self._txt_tramo:
            self._txt_tramo.Text = u"—"
        if self._txt_apoyos:
            self._txt_apoyos.Text = u"— (seleccione apoyos)" if not apoyos_loaded else u"sin vigas"
        if self._txt_sub:
            self._txt_sub.Text = u"Vigas · columnas · muros · tramos cap-panel"
        if self._txt_sel:
            self._txt_sel.Text = (
                u"Clic viga → selección · Ctrl+clic → multi-selección · Traslape sup/inf en panel derecho · "
                u"Controles Tn en bandas alzado · Cent/Ext L auto en panel sección."
            )
        if self._cnv_section:
            self._cnv_section.Children.Clear()
        if self._txt_section:
            self._txt_section.Text = u""
        if self._pnl_section_ctrls:
            self._pnl_section_ctrls.Children.Clear()

        empty = Border()
        empty.Background = brush_hex(u"#071018")
        empty.BorderBrush = brush_hex(u"#21465C")
        empty.BorderThickness = Thickness(1)
        empty.Padding = Thickness(24)
        empty.MinHeight = 180.0
        tb = TextBlock()
        tb.Text = u"Sin vigas en el lote.\nCierre y vuelva a ejecutar la herramienta."
        tb.Foreground = brush_hex(u"#64748b")
        tb.TextAlignment = TextAlignment.Center
        tb.HorizontalAlignment = HorizontalAlignment.Center
        empty.Child = tb
        self._host.Children.Add(empty)

    def _show_canvas_error(self, ex):
        try:
            msg = unicode(ex)
        except NameError:
            msg = str(ex)
        if self._txt_sel:
            self._txt_sel.Text = u"Error al dibujar canvas: {0}".format(msg)
        err = Border()
        err.Background = brush_hex(u"#071018")
        err.BorderBrush = brush_hex(u"#f87171")
        err.BorderThickness = Thickness(1)
        err.Padding = Thickness(16)
        err.MinHeight = 120.0
        tb = TextBlock()
        tb.Text = u"No se pudo renderizar el canvas.\n{0}".format(msg)
        tb.Foreground = brush_hex(u"#f87171")
        tb.TextWrapping = TextWrapping.Wrap
        err.Child = tb
        self._host.Children.Add(err)

    def _update_headers(self, beams, tramos_sup, tramos_inf, apoyos_loaded, layout_result):
        if self._txt_tramo:
            self._txt_tramo.Text = format_dual_tramo_summary(beams, tramos_sup, tramos_inf)
        if self._txt_apoyos:
            if apoyos_loaded:
                a = lay.collect_apoyos(beams)
                self._txt_apoyos.Text = u"{0} columna(s) · {1} muro(s) · cadena {2}".format(
                    a["cols"], a["walls"], u" → ".join(a["ids"]) if a["ids"] else u"—",
                )
            else:
                self._txt_apoyos.Text = u"— (seleccione apoyos)"
        if self._txt_sub:
            self._txt_sub.Text = u"{0} vigas · sup {1} / inf {2} tramos Tn".format(
                len(beams), len(tramos_sup or []), len(tramos_inf or []),
            )
        if self._txt_sel:
            base = (
                u"Clic viga → selección · Ctrl+clic → multi-selección · Traslape sup/inf en panel derecho · "
                u"Controles Tn en bandas alzado · Cent/Ext L auto en panel sección."
                if apoyos_loaded
                else u"Clic viga → selección · Ctrl+clic → multi-selección · Traslape en panel derecho · Cent/Ext · Suple en canvas."
            )
            if layout_result.get("needsScroll"):
                base += u" · Desplazar horizontalmente (↔) para ver todas las vigas."
            self._txt_sel.Text = base

    def _draw_section_rail(self, beams):
        session = self._last_session
        idx = self.selected_beam_idx if self.selected_beam_idx >= 0 else 0
        beam = beams[idx] if beams else None
        role = None
        if self.selected_stirrup_zone and self.selected_stirrup_zone.get("idx") == idx:
            r = self.selected_stirrup_zone.get("role")
            role = {
                "ext": u"Ext · ini/fin",
                "cent": u"Cent",
                "uni": u"Único",
                "confin": u"Confin.",
                "suple": u"Suple inf.",
                "supleSup": u"Suple sup.",
                "laterales": u"Laterales",
            }.get(r, r)
        if self._txt_section_rail and beam:
            role_txt = role or u"Cent"
            n_sel = len(self.selected_beam_indices)
            if n_sel > 1:
                labels = u", ".join(
                    lay.beam_canvas_label(i) for i in sorted(self.selected_beam_indices)
                )
                self._txt_section_rail.Text = u"{0} · sección · {1} · lote ({2})".format(
                    labels, role_txt, n_sel,
                )
            else:
                self._txt_section_rail.Text = u"{0} · sección · {1}".format(
                    lay.beam_canvas_label(idx), role_txt,
                )
        elif self._txt_section_rail:
            self._txt_section_rail.Text = u"Sección · confinamiento"
        if beam:
            ensure_beam_layers(beam)
            ensure_beam_confinement(beam)
        if self._cnv_section and beam:
            draw_section_preview(
                self._cnv_section,
                beam,
                role,
                laterales_enabled=bool(getattr(session, "lateralesEnabled", False)),
                n_laterales=int(getattr(session, "nLaterales", 1) or 1),
                diam_laterales=int(getattr(session, "diamLaterales", 16) or 16),
            )
        if self._txt_section and beam:
            self._txt_section.Text = section_meta_lines(beam, role)
        if self._pnl_section_ctrls is not None:
            self._pnl_section_ctrls.Children.Clear()
            if beam:
                session = self._last_session
                if session is not None:
                    self._pnl_section_ctrls.Children.Add(
                        self._build_traslape_rail_block(beam, idx, session)
                    )
                self._pnl_section_ctrls.Children.Add(
                    self._build_section_stirrup_stack(beam, idx, session)
                )

    def _build_axis_hint(self, apoyos_loaded, content_w):
        row = Border()
        row.Width = content_w
        row.Height = lay.AXIS_HINT_HEIGHT_PX
        row.Padding = Thickness(8, 2, 8, 0)
        sp = StackPanel()
        sp.Orientation = Orientation.Horizontal
        left = TextBlock()
        left.Text = u"← Apoyo ini." if apoyos_loaded else u"← Izquierda (orden)"
        left.Foreground = brush_hex(u"#64748b")
        left.FontSize = typo.HDR_FONT_PX
        mid = TextBlock()
        mid.Text = u"Flecha alzado = orden tramos · modelo = eje Revit 0→1"
        mid.Foreground = brush_hex(u"#64748b")
        mid.FontSize = typo.META_FONT_PX
        mid.Margin = Thickness(24, 0, 24, 0)
        right = TextBlock()
        right.Text = u"Apoyo fin. →" if apoyos_loaded else u"Derecha (orden) →"
        right.Foreground = brush_hex(u"#64748b")
        right.FontSize = typo.HDR_FONT_PX
        sp.Children.Add(left)
        sp.Children.Add(mid)
        sp.Children.Add(right)
        row.Child = sp
        return row

    def _elev_line(self, cnv, x1, y1, x2, y2, stroke, thickness=1.5, dash=None, zindex=0):
        ln = Line()
        ln.X1 = float(x1)
        ln.Y1 = float(y1)
        ln.X2 = float(x2)
        ln.Y2 = float(y2)
        ln.Stroke = stroke
        ln.StrokeThickness = float(thickness)
        if dash:
            ln.StrokeDashArray = DoubleCollection(dash)
        if zindex:
            Canvas.SetZIndex(ln, zindex)
        cnv.Children.Add(ln)
        return ln

    def _elev_rect(self, cnv, x, y, w, h, fill=None, stroke=None, thickness=1.0, dash=None, zindex=0):
        rect = Rectangle()
        rect.Width = float(w)
        rect.Height = float(h)
        Canvas.SetLeft(rect, float(x))
        Canvas.SetTop(rect, float(y))
        if fill is not None:
            rect.Fill = fill
        if stroke is not None:
            rect.Stroke = stroke
            rect.StrokeThickness = float(thickness)
            if dash:
                rect.StrokeDashArray = DoubleCollection(dash)
        if zindex:
            Canvas.SetZIndex(rect, zindex)
        cnv.Children.Add(rect)
        return rect

    def _elev_is_wall_id(self, apoyo_id):
        return unicode(apoyo_id or u"").startswith(u"M")

    def _elev_apoyo_width_px(self, apoyo_id, session):
        is_wall = self._elev_is_wall_id(apoyo_id)
        default = _ELEV_WALL_W if is_wall else _ELEV_COL_W
        if not apoyo_id:
            return default
        for ap in getattr(session, "apoyos", None) or []:
            if ap.get("id") != apoyo_id:
                continue
            mm = ap.get("thicknessMm") if is_wall else ap.get("widthMm")
            if mm:
                scale = _ELEV_WALL_PX_PER_MM if is_wall else _ELEV_COL_PX_PER_MM
                return max(default, float(mm) * scale)
        return default

    def _elev_apoyo_half_px(self, apoyo_id, session):
        return self._elev_apoyo_width_px(apoyo_id, session) * 0.5

    def _elev_beam_vertical(self, beam):
        w_cm, h_cm = parse_beam_section(beam.get("type"))
        h_px = max(14.0, _ELEV_BEAM_H * (float(h_cm) / _ELEV_REF_SECTION_H_CM))
        top = _ELEV_BEAM_BOT - h_px
        return top, h_px

    def _elev_beam_full_span_px(self, beam, lay_i, content_w, session):
        left_px = lay.pct_to_px(lay_i["leftPct"], content_w)
        width_px = lay.pct_to_px(lay_i["widthPct"], content_w)
        s_ext = self._elev_apoyo_half_px(beam.get("colStart"), session)
        e_ext = self._elev_apoyo_half_px(beam.get("colEnd"), session)
        return left_px - s_ext, width_px + s_ext + e_ext

    def _elev_beam_clear_span_px(self, beam, lay_i, content_w, session):
        left_px = lay.pct_to_px(lay_i["leftPct"], content_w)
        width_px = lay.pct_to_px(lay_i["widthPct"], content_w)
        s_in = self._elev_apoyo_half_px(beam.get("colStart"), session)
        e_in = self._elev_apoyo_half_px(beam.get("colEnd"), session)
        return left_px + s_in, max(4.0, width_px - s_in - e_in)

    def _elev_support_zones(self, chain, content_w, session):
        zones = []
        for pt in chain or []:
            cx = lay.pct_to_px(pt["pct"], content_w)
            half = self._elev_apoyo_half_px(pt.get("id"), session)
            zones.append({"x0": cx - half, "x1": cx + half})
        return zones

    def _elev_split_edge_by_supports(self, x0, x1, zones):
        spans = []
        cursor = float(x0)
        clipped = []
        for z in zones or []:
            zx0 = max(float(x0), float(z["x0"]))
            zx1 = min(float(x1), float(z["x1"]))
            if zx1 > zx0:
                clipped.append((zx0, zx1))
        clipped.sort(key=lambda t: t[0])
        for zx0, zx1 in clipped:
            if cursor < zx0:
                spans.append({"a": cursor, "b": zx0, "dashed": False})
            if zx0 < zx1:
                spans.append({"a": zx0, "b": zx1, "dashed": True})
            cursor = max(cursor, zx1)
        if cursor < float(x1):
            spans.append({"a": cursor, "b": float(x1), "dashed": False})
        if not spans:
            spans.append({"a": float(x0), "b": float(x1), "dashed": False})
        return spans

    def _draw_elevation_horiz_edge(self, cnv, y, x0, x1, zones, stroke, thickness, zindex=3):
        for seg in self._elev_split_edge_by_supports(x0, x1, zones):
            self._elev_line(
                cnv, seg["a"], y, seg["b"], y, stroke, thickness,
                dash=[5.0, 3.0] if seg["dashed"] else None,
                zindex=zindex,
            )

    def _draw_elevation_break_line(self, cnv, x, y, w, kind, stroke, zindex=7):
        n = 4
        step = float(w) / n
        amp = _ELEV_BREAK_AMP
        points = PointCollection()
        points.Add(Point(float(x), float(y)))
        for i in range(n):
            x_mid = float(x) + step * (i + 0.5)
            x_end = float(x) + step * (i + 1)
            if kind == u"bottom":
                points.Add(Point(x_mid, float(y) - amp))
                points.Add(Point(x_end, float(y) + amp))
            else:
                points.Add(Point(x_mid, float(y) + amp))
                points.Add(Point(x_end, float(y) - amp))
        pl = Polyline()
        pl.Points = points
        pl.Stroke = stroke
        pl.StrokeThickness = 1.1
        Canvas.SetZIndex(pl, zindex)
        cnv.Children.Add(pl)

    def _draw_elevation_beam_fill(self, cnv, x, w, top, h_px):
        axis = brush_hex(u"#64748b", 120)
        self._elev_rect(
            cnv, x, top, w, h_px,
            fill=SolidColorBrush(Color.FromArgb(82, 148, 163, 184)),
            zindex=1,
        )
        self._elev_line(
            cnv, x, top + h_px * 0.5, x + w, top + h_px * 0.5,
            axis, 1.0, dash=[4.0, 2.5], zindex=2,
        )

    def _draw_elevation_beam_edges(self, cnv, x, w, top, h_px, zones, selected=False):
        stroke = brush_hex(u"#22d3ee", 240) if selected else _ELEV_BEAM_CHORD
        sw = 2.2 if selected else _ELEV_STROKE_CHORD
        bot = top + h_px
        self._draw_elevation_horiz_edge(cnv, top, x, x + w, zones, stroke, sw, zindex=3)
        self._draw_elevation_horiz_edge(cnv, bot, x, w + x, zones, stroke, sw, zindex=3)
        self._elev_line(cnv, x, top, x, bot, stroke, sw * 0.85, zindex=3)
        self._elev_line(cnv, x + w, top, x + w, bot, stroke, sw * 0.85, zindex=3)

    def _draw_elevation_beam_section_label(self, cnv, x, w, top, beam):
        w_cm, h_cm = parse_beam_section(beam.get("type"))
        lbl = TextBlock()
        lbl.Text = u"V. {0}/{1}".format(int(round(w_cm * 10)), int(round(h_cm * 10)))
        lbl.FontSize = typo.META_FONT_PX
        lbl.FontWeight = FontWeights.SemiBold
        lbl.Foreground = brush_hex(u"#e2e8f0", 210)
        lbl.TextAlignment = TextAlignment.Center
        lbl.Width = w
        Canvas.SetLeft(lbl, x)
        Canvas.SetTop(lbl, top - 14.0)
        Canvas.SetZIndex(lbl, 9)
        cnv.Children.Add(lbl)

    def _draw_elevation_beam_run(self, cnv, run_left, run_w):
        """Legacy — corrida continua sin apoyos."""
        self._draw_elevation_beam_fill(cnv, run_left, run_w, _ELEV_BEAM_TOP, _ELEV_BEAM_H)
        chord = brush_hex(u"#94a3b8", 210)
        self._elev_line(cnv, run_left, _ELEV_BEAM_TOP, run_left + run_w, _ELEV_BEAM_TOP, chord, _ELEV_STROKE_CHORD, zindex=2)
        self._elev_line(
            cnv, run_left, _ELEV_BEAM_BOT, run_left + run_w, _ELEV_BEAM_BOT, chord, _ELEV_STROKE_CHORD, zindex=2,
        )

    def _draw_elevation_direction_marker(self, cnv, left, width, order_idx, axis_reversed=False, y_mid=None):
        """Marca LocationCurve 0 (⊥) → 1 (flecha), corta y centrada en el tramo."""
        y = y_mid if y_mid is not None else (_ELEV_BEAM_TOP + _ELEV_BEAM_H * 0.5)
        cx = left + width * 0.5
        marker_w = min(_ELEV_DIR_MARKER_LEN_PX, max(24.0, width * 0.22))
        half = marker_w * 0.5
        if width <= half * 2.0 + 4.0:
            return
        if axis_reversed:
            start_x, end_x = cx + half, cx - half
        else:
            start_x, end_x = cx - half, cx + half
        stroke = brush_hex(u"#38bdf8", 210)
        tip = end_x
        shaft_end = tip - 7.0 if tip >= start_x else tip + 7.0
        self._elev_line(cnv, start_x, y, shaft_end, y, stroke, 1.6, zindex=3)
        if tip >= start_x:
            self._elev_line(cnv, tip - 7.0, y - 3.5, tip, y, stroke, 1.6, zindex=3)
            self._elev_line(cnv, tip - 7.0, y + 3.5, tip, y, stroke, 1.6, zindex=3)
            tick_x = start_x - 2.0
        else:
            self._elev_line(cnv, tip + 7.0, y - 3.5, tip, y, stroke, 1.6, zindex=3)
            self._elev_line(cnv, tip + 7.0, y + 3.5, tip, y, stroke, 1.6, zindex=3)
            tick_x = start_x + 2.0
        self._elev_line(cnv, tick_x, y - 4.0, tick_x, y + 4.0, stroke, 1.4, zindex=3)
        if order_idx is not None:
            lbl = TextBlock()
            lbl.Text = u"{0}".format(int(order_idx) + 1)
            lbl.FontSize = typo.META_FONT_PX
            lbl.FontWeight = FontWeights.Bold
            lbl.Foreground = stroke
            lbl_x = (start_x + 2.0) if not axis_reversed else (start_x - 10.0)
            Canvas.SetLeft(lbl, lbl_x)
            Canvas.SetTop(lbl, y - 12.0)
            Canvas.SetZIndex(lbl, 4)
            cnv.Children.Add(lbl)

    def _draw_elevation_support(self, cnv, x, pt, idx, session):
        is_wall = self._elev_is_wall_id(pt.get("id"))
        col_w = self._elev_apoyo_width_px(pt.get("id"), session)
        col_top = _ELEV_COL_TOP
        col_h = _ELEV_COL_H
        left = x - col_w * 0.5
        fill = brush_hex(u"#34d399", 72) if is_wall else brush_hex(u"#94a3b8", 96)
        edge = brush_hex(u"#6ee7b7", 200) if is_wall else brush_hex(u"#e2e8f0", 220)

        self._elev_rect(cnv, left, col_top, col_w, col_h, fill=fill, zindex=6)
        self._elev_line(cnv, left, col_top, left, col_top + col_h, edge, _ELEV_SUPPORT_STROKE, zindex=8)
        self._elev_line(
            cnv, left + col_w, col_top, left + col_w, col_top + col_h, edge, _ELEV_SUPPORT_STROKE, zindex=8,
        )
        axis = brush_hex(u"#34d399", 150) if is_wall else brush_hex(u"#94a3b8", 160)
        self._elev_line(
            cnv, x, col_top, x, col_top + col_h, axis, 0.75, dash=[6.0, 4.0], zindex=7,
        )
        self._draw_elevation_break_line(cnv, left, col_top, col_w, u"top", edge, zindex=8)
        self._draw_elevation_break_line(cnv, left, col_top + col_h, col_w, u"bottom", edge, zindex=8)

    def _bar_span_edges(self, i, n, left, width):
        inset_l = _ELEV_BAR_INSET if i == 0 else 0.0
        inset_r = _ELEV_BAR_INSET if i == n - 1 else 0.0
        return left + inset_l, left + width - inset_r

    def _draw_elevation_top_bars(self, cnv, layouts, tramos_sup, content_w):
        for tramo in tramos_sup or []:
            if not tramo.get("beamIndices"):
                continue
            accent = tramo.get("accent") or u"#22d3ee"
            sup_brush = accent_soft_brush(accent, "strokeSel")
            span = lay.tramo_span(layouts, tramo, content_w)
            x0 = lay.pct_to_px(span["leftPct"], content_w)
            x1 = x0 + lay.pct_to_px(span["widthPct"], content_w)
            if x1 > x0:
                self._elev_line(
                    cnv, x0, _ELEV_BAR_SUP_Y, x1, _ELEV_BAR_SUP_Y,
                    sup_brush, _ELEV_STROKE_BAR, zindex=4,
                )
                if x1 - x0 >= 28.0:
                    lbl = TextBlock()
                    lbl.Text = u"T{0}".format(tramo.get("id"))
                    lbl.FontSize = typo.META_FONT_PX
                    lbl.FontWeight = FontWeights.Bold
                    lbl.Foreground = accent_soft_brush(accent, "text")
                    lbl.TextAlignment = TextAlignment.Center
                    lbl.Width = x1 - x0
                    Canvas.SetLeft(lbl, x0)
                    Canvas.SetTop(lbl, _ELEV_BAR_SUP_Y - 14.0)
                    Canvas.SetZIndex(lbl, 5)
                    cnv.Children.Add(lbl)

    def _support_col_specs(self, chain, content_w, session=None):
        specs = []
        for idx, pt in enumerate(chain):
            is_wall = self._elev_is_wall_id(pt.get("id"))
            col_w = self._elev_apoyo_width_px(pt.get("id"), session)
            specs.append({
                "x": lay.pct_to_px(pt["pct"], content_w),
                "half_w": col_w * 0.5,
                "hook": (not is_wall) and idx > 0,
            })
        return specs

    def _draw_inf_bar_segment(self, cnv, seg_x0, seg_x1, specs, inf_solid, inf_hidden):
        cols = [
            c for c in specs
            if (c["x"] - c["half_w"]) < seg_x1 and (c["x"] + c["half_w"]) > seg_x0
        ]
        cols.sort(key=lambda c: c["x"])
        cursor = seg_x0
        for col in cols:
            cx = col["x"]
            hw = col["half_w"]
            cleft = cx - hw
            cright = cx + hw
            bar_left = max(cleft, seg_x0)
            bar_right = min(cright, seg_x1)
            if cursor < bar_left:
                self._elev_line(cnv, cursor, _ELEV_BAR_INF_Y, bar_left, _ELEV_BAR_INF_Y, inf_solid, _ELEV_STROKE_BAR, zindex=4)
            if col["hook"] and bar_left < bar_right:
                drop = _ELEV_BAR_INF_Y - _ELEV_POCKET_D
                self._elev_line(cnv, bar_left, _ELEV_BAR_INF_Y, bar_left, drop, inf_solid, _ELEV_STROKE_BAR, zindex=4)
                self._elev_line(
                    cnv, bar_left, drop, bar_right, drop, inf_hidden, 2.0, dash=[2.5, 2.0], zindex=4,
                )
                self._elev_line(cnv, bar_right, drop, bar_right, _ELEV_BAR_INF_Y, inf_solid, _ELEV_STROKE_BAR, zindex=4)
            cursor = max(cursor, bar_right)
        if cursor < seg_x1:
            self._elev_line(cnv, cursor, _ELEV_BAR_INF_Y, seg_x1, _ELEV_BAR_INF_Y, inf_solid, _ELEV_STROKE_BAR, zindex=4)

    def _draw_elevation_bottom_bars(self, cnv, layouts, tramos_inf, content_w, chain, session=None):
        specs = self._support_col_specs(chain, content_w, session) if chain else []
        for tramo in tramos_inf or []:
            if not tramo.get("beamIndices"):
                continue
            accent = tramo.get("accent") or u"#fb7185"
            inf_solid = accent_soft_brush(accent, "strokeSel")
            inf_hidden = accent_soft_brush(accent, "stroke")
            span = lay.tramo_span(layouts, tramo, content_w)
            x0 = lay.pct_to_px(span["leftPct"], content_w)
            x1 = x0 + lay.pct_to_px(span["widthPct"], content_w)
            if x1 > x0:
                self._draw_inf_bar_segment(cnv, x0, x1, specs, inf_solid, inf_hidden)
                if x1 - x0 >= 28.0:
                    lbl = TextBlock()
                    lbl.Text = u"T{0}".format(tramo.get("id"))
                    lbl.FontSize = typo.META_FONT_PX
                    lbl.FontWeight = FontWeights.Bold
                    lbl.Foreground = accent_soft_brush(accent, "text")
                    lbl.TextAlignment = TextAlignment.Center
                    lbl.Width = x1 - x0
                    Canvas.SetLeft(lbl, x0)
                    Canvas.SetTop(lbl, _ELEV_BAR_INF_Y + 2.0)
                    Canvas.SetZIndex(lbl, 5)
                    cnv.Children.Add(lbl)

    def _draw_elevation_suple_superior_tramos(self, cnv, beams, layouts, content_w):
        """Zonas 25 % extremos + barra suple sup. (fusión consecutiva)."""
        zone_fill = SolidColorBrush(Color.FromArgb(15, 192, 132, 252))
        zone_stroke = th.brush_sem(th.SEM_SUPLE, 64)
        suple_stroke = th.brush_sem(th.SEM_SUPLE, 220)
        suple_merged = brush_hex(u"#e879f9", 235)
        dash = [3.0, 2.0]

        for i, beam in enumerate(beams or []):
            if i >= len(layouts or []):
                break
            ensure_beam_suple_superior(beam)
            if not beam_suple_sup_enabled(beam):
                continue
            lay_i = layouts[i]
            left = lay.pct_to_px(lay_i["leftPct"], content_w)
            width = lay.pct_to_px(lay_i["widthPct"], content_w)
            if width < 4.0:
                continue
            span_w = width * SUPLE_END_PCT
            bar_top = _ELEV_BAR_SUPLE_SUP_Y - 6.0
            bar_h = 14.0
            if beam_suple_sup_side_enabled(beam, "start"):
                self._elev_rect(
                    cnv, left, bar_top, span_w, bar_h,
                    fill=zone_fill, stroke=zone_stroke, thickness=0.6, dash=dash, zindex=3,
                )
            if beam_suple_sup_side_enabled(beam, "end"):
                self._elev_rect(
                    cnv, left + width - span_w, bar_top, span_w, bar_h,
                    fill=zone_fill, stroke=zone_stroke, thickness=0.6, dash=dash, zindex=3,
                )

        segs = suple_sup_segments_layout_px(beams, layouts, content_w, lay.pct_to_px)
        for seg in segs:
            x0 = seg.get("x0", 0)
            x1 = seg.get("x1", 0)
            if x1 - x0 < 2.0:
                continue
            merged = bool(seg.get("merged"))
            stroke = suple_merged if merged else suple_stroke
            w = 3.5 if merged else 3.0
            self._elev_line(
                cnv, x0, _ELEV_BAR_SUPLE_SUP_Y, x1, _ELEV_BAR_SUPLE_SUP_Y,
                stroke, w, zindex=5,
            )
            for cx in (x0, x1):
                self._elev_rect(
                    cnv, cx - 2.5, _ELEV_BAR_SUPLE_SUP_Y - 2.5, 5.0, 5.0,
                    fill=stroke, zindex=6,
                )
            if merged and seg.get("junctionX") is not None:
                jx = seg["junctionX"]
                self._elev_line(
                    cnv, jx, _ELEV_BAR_SUPLE_SUP_Y - 6.0, jx, _ELEV_BAR_SUPLE_SUP_Y + 6.0,
                    brush_hex(u"#fbbf24", 200), 1.0, dash=[2.0, 2.0], zindex=4,
                )

        if segs:
            lbl = TextBlock()
            lbl.Text = u"Suple sup. · capa n+1"
            lbl.FontSize = typo.META_FONT_PX
            lbl.Foreground = suple_stroke
            Canvas.SetLeft(lbl, 4.0)
            Canvas.SetTop(lbl, _ELEV_BAR_SUPLE_SUP_Y - 10.0)
            Canvas.SetZIndex(lbl, 6)
            cnv.Children.Add(lbl)

    def _draw_elevation_suple_inferior_tramos(self, cnv, beams, layouts, content_w):
        """Zonas 10 % extremos + barra suple central 80 % por viga (si activo)."""
        trim_fill = SolidColorBrush(Color.FromArgb(31, 251, 191, 36))
        trim_stroke = brush_hex(u"#fbbf24", 90)
        suple_fill = SolidColorBrush(Color.FromArgb(15, 192, 132, 252))
        suple_stroke = th.brush_sem(th.SEM_SUPLE, 220)
        suple_brush = th.brush_sem(th.SEM_SUPLE, 235)
        dash = [3.0, 2.0]

        for i, beam in enumerate(beams or []):
            if i >= len(layouts or []):
                break
            ensure_beam_suple_inferior(beam)
            if not beam_suple_inf_enabled(beam):
                continue
            lay_i = layouts[i]
            left = lay.pct_to_px(lay_i["leftPct"], content_w)
            width = lay.pct_to_px(lay_i["widthPct"], content_w)
            if width < 4.0:
                continue
            trim_w = width * 0.1
            sup_x0 = left + trim_w
            sup_x1 = left + width - trim_w
            bar_top = _ELEV_BAR_SUPLE_Y - 6.0
            bar_h = 18.0

            self._elev_rect(
                cnv, left, bar_top, trim_w, bar_h,
                fill=trim_fill, stroke=trim_stroke, thickness=0.8, dash=dash, zindex=3,
            )
            self._elev_rect(
                cnv, left + width - trim_w, bar_top, trim_w, bar_h,
                fill=trim_fill, stroke=trim_stroke, thickness=0.8, dash=dash, zindex=3,
            )

            self._elev_rect(
                cnv, sup_x0, bar_top, sup_x1 - sup_x0, bar_h,
                fill=suple_fill, stroke=th.brush_sem(th.SEM_SUPLE, 64), thickness=0.6, zindex=3,
            )
            self._elev_line(
                cnv, sup_x0, _ELEV_BAR_SUPLE_Y, sup_x1, _ELEV_BAR_SUPLE_Y,
                suple_brush, 3.0, zindex=5,
            )
            for cx in (sup_x0, sup_x1):
                self._elev_rect(
                    cnv, cx - 2.5, _ELEV_BAR_SUPLE_Y - 2.5, 5.0, 5.0,
                    fill=suple_stroke, zindex=6,
                )

            layer = beam_suple_layer_index(beam)
            lbl = TextBlock()
            lbl.Text = u"Suple · {0} · capa {1}".format(beam.get("id") or u"V?", layer)
            lbl.FontSize = typo.META_FONT_PX
            lbl.FontWeight = FontWeights.SemiBold
            lbl.Foreground = suple_stroke
            Canvas.SetLeft(lbl, sup_x0)
            Canvas.SetTop(lbl, bar_top - 10.0)
            Canvas.SetZIndex(lbl, 6)
            cnv.Children.Add(lbl)

            len_mm = 0
            try:
                len_mm = int(round(float(beam.get("len") or 0) * 1000.0))
            except Exception:
                pass
            m = suple_metrics_mm(len_mm)
            dim_y = _ELEV_BAR_SUPLE_Y + 10.0
            self._elev_line(cnv, sup_x0, dim_y, sup_x1, dim_y, suple_stroke, 1.0, zindex=4)
            dim_lbl = TextBlock()
            dim_lbl.Text = u"{0} mm".format(m.get("spanMm") or 0)
            dim_lbl.FontSize = typo.META_FONT_PX
            dim_lbl.Foreground = suple_stroke
            Canvas.SetLeft(dim_lbl, (sup_x0 + sup_x1) * 0.5 - 16.0)
            Canvas.SetTop(dim_lbl, dim_y + 2.0)
            Canvas.SetZIndex(dim_lbl, 6)
            cnv.Children.Add(dim_lbl)

    def _elev_active_hint(self):
        parts = []
        if self.selected_tramo_sup_id:
            parts.append(u"T{0} sup".format(self.selected_tramo_sup_id))
        if self.selected_tramo_inf_id:
            parts.append(u"T{0} inf".format(self.selected_tramo_inf_id))
        return u"Tramo: " + u" · ".join(parts) if parts else u"Sin tramo seleccionado"

    def _build_elev_stage_option_d(
        self, beams, layouts, tramos_sup, tramos_inf, session, apoyos_loaded, content_w,
    ):
        """Opción D: bandas Tn con controles + alzado protagonista."""
        stage = StackPanel()
        stage.Width = content_w

        hdr = Grid()
        hdr.Margin = Thickness(lay.FACE_BLOCK_PAD_PX, 6, lay.FACE_BLOCK_PAD_PX, 4)
        col_title = ColumnDefinition()
        col_title.Width = GridLength(1.0, GridUnitType.Star)
        col_hint = ColumnDefinition()
        col_hint.Width = GridLength.Auto
        hdr.ColumnDefinitions.Add(col_title)
        hdr.ColumnDefinitions.Add(col_hint)

        title = TextBlock()
        title.Text = u"Alzado · corrida · fibras · apoyos · marcas empalme"
        title.Foreground = brush_hex(u"#64748b")
        title.FontSize = typo.TITLE_FONT_PX
        title.FontWeight = FontWeights.Bold
        Grid.SetColumn(title, 0)
        hdr.Children.Add(title)

        hint = TextBlock()
        hint.Text = self._elev_active_hint()
        hint.Foreground = brush_hex(u"#5bb8d4")
        hint.FontSize = typo.META_FONT_PX
        hint.FontWeight = FontWeights.SemiBold
        Grid.SetColumn(hint, 1)
        hdr.Children.Add(hint)
        stage.Children.Add(hdr)

        stage.Children.Add(
            self._build_tramo_bands_ctrl_row(tramos_sup, beams, layouts, content_w, u"sup")
        )
        stage.Children.Add(
            self._build_elevation_canvas(
                beams, layouts, tramos_sup, tramos_inf, session, apoyos_loaded, content_w,
            )
        )
        stage.Children.Add(
            self._build_tramo_bands_ctrl_row(tramos_inf, beams, layouts, content_w, u"inf")
        )
        return stage

    def _build_band_collapsed_summary(self, tramo, beams, face, accent):
        ref_beam = beams[tramo["beamIndices"][0]] if tramo.get("beamIndices") else beams[0]
        is_sup = face == u"sup"
        n_c = beam_n_capas_sup(ref_beam) if is_sup else beam_n_capas_inf(ref_beam)
        chips = []
        for layer_num in range(1, n_c + 1):
            k = layer_keys(layer_num)
            qty_f = k["nSup"] if is_sup else k["nInf"]
            diam_f = k["diamSup"] if is_sup else k["diamInf"]
            chips.append(
                u"{0} {1}ø{2}".format(
                    k["label"], ref_beam.get(qty_f) or 2, ref_beam.get(diam_f) or 16,
                )
            )
        suffix = u" · traslapo" if tramo.get("fromEmpalme") else u""
        sp = StackPanel()
        sp.HorizontalAlignment = HorizontalAlignment.Center
        tn = TextBlock()
        tn.Text = u"T{0}{1}".format(tramo.get("id"), suffix)
        tn.FontSize = typo.TITLE_FONT_PX
        tn.FontWeight = FontWeights.Bold
        tn.Foreground = accent_soft_brush(accent, "text")
        tn.TextAlignment = TextAlignment.Center
        sp.Children.Add(tn)
        mini = TextBlock()
        mini.Text = u" · ".join(chips)
        mini.FontSize = typo.META_FONT_PX
        mini.Foreground = brush_hex(u"#e8f4f8", 180)
        mini.TextAlignment = TextAlignment.Center
        mini.TextWrapping = TextWrapping.NoWrap
        mini.TextTrimming = System.Windows.TextTrimming.CharacterEllipsis
        sp.Children.Add(mini)
        return sp

    def _build_band_tint_strip(self, accent, selected):
        """Franja fina de color tramo — sup al pie, inf en cabecera."""
        tint = Border()
        tint.Height = lay.TRAMO_BAND_TINT_HEIGHT_PX
        tint.Background = accent_soft_brush(accent, "strokeSel" if selected else "stroke")
        return tint

    def _build_tramo_band_cell(self, tramo, beams, face, accent, selected):
        """Celda banda Tn: cuerpo neutro + franja tinte (sup abajo · inf arriba)."""
        is_sup = face == u"sup"
        ref_beam = beams[tramo["beamIndices"][0]] if tramo.get("beamIndices") else beams[0]
        n_c = beam_n_capas_sup(ref_beam) if is_sup else beam_n_capas_inf(ref_beam)
        total_h = lay.tramo_band_cell_height_px(selected, n_c if selected else 0)

        cell = Border()
        cell.Margin = Thickness(1, 0, 1, 0)
        cell.Height = total_h
        cell.VerticalAlignment = VerticalAlignment.Top
        cell.Background = th.brush_panel()
        cell.BorderBrush = th.selection_border_brush(selected)
        cell.BorderThickness = Thickness(1)
        if selected:
            cell.Background = th.selection_background_brush(True)
        cell.Cursor = Cursors.Hand

        body = Border()
        body.Padding = Thickness(2, 2, 2, 1 if is_sup else 2)
        body.Background = th.brush_panel(0)
        body.HorizontalAlignment = HorizontalAlignment.Stretch
        body.VerticalAlignment = VerticalAlignment.Stretch

        content = StackPanel()
        content.HorizontalAlignment = HorizontalAlignment.Center
        content.VerticalAlignment = VerticalAlignment.Center
        if selected:
            tramo_beams = self._tramo_beams(tramo, beams)
            owner = tramo_beams[0] if tramo_beams else beams[0]
            cap = self._cap_col(tramo_beams, owner, face, tramo_accent=accent, compact_band=True)
            cap.HorizontalAlignment = HorizontalAlignment.Center
            cap.Margin = Thickness(0)
            content.Children.Add(cap)
        else:
            content.Children.Add(
                self._build_band_collapsed_summary(tramo, beams, face, accent)
            )
        body.Child = content

        tint = self._build_band_tint_strip(accent, selected)
        shell = Grid()
        shell.Height = total_h
        if is_sup:
            row_body = RowDefinition()
            row_body.Height = GridLength(1.0, GridUnitType.Star)
            row_tint = RowDefinition()
            row_tint.Height = GridLength(lay.TRAMO_BAND_TINT_HEIGHT_PX)
            shell.RowDefinitions.Add(row_body)
            shell.RowDefinitions.Add(row_tint)
            Grid.SetRow(body, 0)
            Grid.SetRow(tint, 1)
        else:
            row_tint = RowDefinition()
            row_tint.Height = GridLength(lay.TRAMO_BAND_TINT_HEIGHT_PX)
            row_body = RowDefinition()
            row_body.Height = GridLength(1.0, GridUnitType.Star)
            shell.RowDefinitions.Add(row_tint)
            shell.RowDefinitions.Add(row_body)
            Grid.SetRow(tint, 0)
            Grid.SetRow(body, 1)
        shell.Children.Add(body)
        shell.Children.Add(tint)
        cell.Child = shell
        return cell

    def _build_tramo_bands_ctrl_row(self, tramos, beams, layouts, content_w, face):
        is_sup = face == u"sup"
        sel_id = self.selected_tramo_sup_id if is_sup else self.selected_tramo_inf_id
        accent_default = u"#22d3ee" if is_sup else u"#fb7185"
        tramos = tramos or []

        wrap = Border()
        if is_sup:
            wrap.Margin = Thickness(lay.FACE_BLOCK_PAD_PX, 2, lay.FACE_BLOCK_PAD_PX, 0)
        else:
            wrap.Margin = Thickness(lay.FACE_BLOCK_PAD_PX, 0, lay.FACE_BLOCK_PAD_PX, 2)

        grid = Grid()
        grid.VerticalAlignment = VerticalAlignment.Top
        inner_w = max(100.0, content_w - 2.0 * lay.FACE_BLOCK_PAD_PX)
        grid.Width = inner_w

        for tramo in tramos:
            span = lay.tramo_span(layouts, tramo, content_w)
            col = ColumnDefinition()
            col.Width = GridLength(max(span["widthPct"], 0.5), GridUnitType.Star)
            grid.ColumnDefinitions.Add(col)

        for i, tramo in enumerate(tramos):
            accent = tramo.get("accent") or accent_default
            sel = sel_id == tramo["id"]
            cell = self._build_tramo_band_cell(tramo, beams, face, accent, sel)
            Grid.SetColumn(cell, i)

            tid = tramo["id"]

            def _select_band(sender, args, tramo_id=tid, tramo_face=face):
                if tramo_face == u"sup":
                    self.selected_tramo_sup_id = tramo_id
                else:
                    self.selected_tramo_inf_id = tramo_id
                self._cb.get("on_select_tramo", lambda _t, _f: None)(tramo_id, tramo_face)
                self._cb.get("on_redraw", lambda: None)()

            try:
                cell.MouseLeftButtonUp += MouseButtonEventHandler(_select_band)
            except Exception:
                pass
            grid.Children.Add(cell)

        wrap.Child = grid
        return wrap

    def _build_traslape_rail_block(self, beam, idx, session):
        block = Border()
        block.Margin = Thickness(0, 0, 0, 10)
        block.Padding = Thickness(8, 8, 8, 8)
        block.BorderBrush = brush_hex(u"#fbbf24", 89)
        block.BorderThickness = Thickness(1)
        block.Background = brush_hex(u"#16120a", 128)

        sp = StackPanel()
        hdr = TextBlock()
        hdr.Text = u"Traslape · viga seleccionada"
        hdr.Foreground = brush_hex(u"#fbbf24")
        hdr.FontSize = typo.TITLE_FONT_PX
        hdr.FontWeight = FontWeights.Bold
        hdr.Margin = Thickness(0, 0, 0, 4)
        sp.Children.Add(hdr)

        hint = TextBlock()
        hint.Text = u"{0} · {1} · {2:.1f} m — @ mitad de viga".format(
            beam.get("id"), beam.get("type"), float(beam.get("len") or 0),
        )
        hint.Foreground = brush_hex(u"#64748b")
        hint.FontSize = typo.META_FONT_PX
        hint.TextWrapping = TextWrapping.Wrap
        hint.Margin = Thickness(0, 0, 0, 8)
        sp.Children.Add(hint)

        row = Grid()
        col_s = ColumnDefinition()
        col_s.Width = GridLength(1.0, GridUnitType.Star)
        col_i = ColumnDefinition()
        col_i.Width = GridLength(1.0, GridUnitType.Star)
        row.ColumnDefinitions.Add(col_s)
        row.ColumnDefinitions.Add(col_i)

        bid = beam.get("id")
        on_sup = bid in (session.empalme_beam_ids_sup or set())
        on_inf = bid in (session.empalme_beam_ids_inf or set())

        def _tras_btn(label, active, face_key, accent):
            btn = Button()
            btn.Content = u"{0}{1}".format(label, u" ✓" if active else u"")
            btn.Padding = Thickness(6, 6, 6, 6)
            btn.FontSize = typo.CTRL_FONT_PX
            btn.FontWeight = FontWeights.SemiBold
            btn.Cursor = Cursors.Hand
            if active:
                btn.Background = accent_soft_brush(accent, "fillSel")
                btn.Foreground = accent_soft_brush(accent, "text")
                btn.BorderBrush = accent_soft_brush(accent, "border")
            else:
                btn.Background = brush_hex(u"#071018")
                btn.Foreground = brush_hex(u"#64748b")
                btn.BorderBrush = brush_hex(u"#2d4455")
            btn.BorderThickness = Thickness(1)

            def _click(sender, args, f=face_key):
                self._cb.get("on_toggle_empalme", lambda _b, _f: None)(bid, f)
                self._cb.get("on_select_beam", lambda _i: None)(idx)
                self._cb.get("on_redraw", lambda: None)()

            try:
                btn.Click += RoutedEventHandler(_click)
            except Exception:
                pass
            return btn

        btn_sup = _tras_btn(u"Superior", on_sup, u"sup", u"#22d3ee")
        btn_inf = _tras_btn(u"Inferior", on_inf, u"inf", u"#fb7185")
        Grid.SetColumn(btn_sup, 0)
        Grid.SetColumn(btn_inf, 1)
        row.Children.Add(btn_sup)
        row.Children.Add(btn_inf)
        sp.Children.Add(row)
        block.Child = sp
        return block

    def _draw_empalme_fiber_mark(self, cnv, x, is_sup, accent_hex):
        y = _ELEV_BAR_SUP_Y if is_sup else _ELEV_BAR_INF_Y
        color = accent_soft_brush(accent_hex or (u"#22d3ee" if is_sup else u"#fb7185"), "strokeSel")
        tick = 7.0
        self._elev_line(cnv, x, y - tick, x, y + tick, color, 2.5, zindex=12)
        self._elev_line(cnv, x - 6.0, y, x + 6.0, y, color, 2.0, zindex=12)
        lbl = u"emp sup" if is_sup else u"emp inf"
        tb = TextBlock()
        tb.Text = lbl
        tb.FontSize = typo.META_FONT_PX
        tb.FontWeight = FontWeights.Bold
        tb.Foreground = color
        Canvas.SetLeft(tb, x + 6.0)
        Canvas.SetTop(tb, y - (14.0 if is_sup else -4.0))
        Canvas.SetZIndex(tb, 12)
        cnv.Children.Add(tb)

    def _build_elevation_canvas(self, beams, layouts, tramos_sup, tramos_inf, session, apoyos_loaded, content_w):
        elev_border = Border()
        elev_border.BorderBrush = brush_hex(u"#21465C", 115)
        elev_border.BorderThickness = Thickness(0, 1, 0, 1)
        elev_border.Padding = Thickness(lay.FACE_BLOCK_PAD_PX, 2, lay.FACE_BLOCK_PAD_PX, 2)
        elev_border.Background = th.brush_panel(0)

        cnv = Canvas()
        cnv.Width = content_w
        cnv.Height = lay.ELEVATION_HEIGHT_PX
        cnv.Background = brush_hex(u"#0a1620", 0)
        cnv.ClipToBounds = True

        chain = (
            lay.build_support_chain(
                beams,
                layouts,
                apoyos=getattr(session, "apoyos", None),
                layout_meta=self._layout_meta,
            )
            if apoyos_loaded
            else []
        )
        zones = self._elev_support_zones(chain, content_w, session) if apoyos_loaded else []

        if apoyos_loaded and layouts:
            for i, beam in enumerate(beams):
                lay_i = layouts[i]
                fx, fw = self._elev_beam_full_span_px(beam, lay_i, content_w, session)
                top, h_px = self._elev_beam_vertical(beam)
                self._draw_elevation_beam_fill(cnv, fx, fw, top, h_px)
        elif layouts:
            run_left = lay.pct_to_px(layouts[0]["leftPct"], content_w)
            run_w = lay.pct_to_px(
                layouts[-1]["leftPct"] + layouts[-1]["widthPct"] - layouts[0]["leftPct"],
                content_w,
            )
            self._draw_elevation_beam_run(cnv, run_left, run_w)

        self._draw_elevation_top_bars(cnv, layouts, tramos_sup, content_w)
        self._draw_elevation_suple_superior_tramos(cnv, beams, layouts, content_w)
        self._draw_elevation_bottom_bars(cnv, layouts, tramos_inf, content_w, chain, session)
        self._draw_elevation_suple_inferior_tramos(cnv, beams, layouts, content_w)

        if apoyos_loaded and layouts:
            for i, beam in enumerate(beams):
                lay_i = layouts[i]
                fx, fw = self._elev_beam_full_span_px(beam, lay_i, content_w, session)
                top, h_px = self._elev_beam_vertical(beam)
                self._draw_elevation_beam_edges(
                    cnv, fx, fw, top, h_px, zones, selected=self._is_beam_selected(i),
                )
                self._draw_elevation_beam_section_label(cnv, fx, fw, top, beam)

        for j, pt in enumerate(chain):
            x = lay.pct_to_px(pt["pct"], content_w)
            self._draw_elevation_support(cnv, x, pt, j, session)

        for i, beam in enumerate(beams):
            lay_i = layouts[i]
            left = lay.pct_to_px(lay_i["leftPct"], content_w)
            width = lay.pct_to_px(lay_i["widthPct"], content_w)
            top, h_px = self._elev_beam_vertical(beam)
            self._draw_elevation_direction_marker(
                cnv, left, width, i, bool(beam.get("axisReversed")), y_mid=top + h_px * 0.5,
            )
            empalme_sup = beam.get("id") in (session.empalme_beam_ids_sup or set())
            empalme_inf = beam.get("id") in (session.empalme_beam_ids_inf or set())
            is_sel = self._is_beam_selected(i)

            span = Border()
            span.Width = width
            span.Height = h_px
            Canvas.SetLeft(span, left)
            Canvas.SetTop(span, top)
            span.Background = brush_hex(u"#000000", 0)
            span.BorderBrush = brush_hex(u"#22d3ee") if is_sel else brush_hex(u"#000000", 0)
            span.BorderThickness = Thickness(2 if is_sel else 0)
            span.Cursor = Cursors.Hand
            Canvas.SetZIndex(span, 10)

            if empalme_sup or empalme_inf:
                mid_x = left + width * 0.5
                if empalme_sup:
                    tr_s = find_tramo_for_beam(tramos_sup, i)
                    acc_s = (tr_s or {}).get("accent") or u"#22d3ee"
                    self._draw_empalme_fiber_mark(cnv, mid_x, True, acc_s)
                if empalme_inf:
                    tr_i = find_tramo_for_beam(tramos_inf, i)
                    acc_i = (tr_i or {}).get("accent") or u"#fb7185"
                    self._draw_empalme_fiber_mark(cnv, mid_x, False, acc_i)

            idx_cap = i
            beam_cap = beam

            def _click_beam(sender, args, idx=idx_cap, b=beam_cap):
                self._select_beam_from_elevation(idx, b, args)

            try:
                span.MouseLeftButtonUp += MouseButtonEventHandler(_click_beam)
            except Exception:
                pass
            cnv.Children.Add(span)

        elev_border.Child = cnv
        return elev_border

    def _select_beam_from_elevation(self, idx, beam, args=None):
        role = self._default_stirrup_role(beam)
        self._handle_beam_select(idx, args, role=role, update_zone=True)

    def _build_labels(self, beams, layouts, session, apoyos_loaded, content_w):
        cnv = Canvas()
        cnv.Width = content_w
        cnv.Height = lay.LABELS_HEIGHT_PX
        for i, beam in enumerate(beams):
            lay_i = layouts[i]
            cx = lay.pct_to_px(lay_i["centerPct"], content_w)
            tb = TextBlock()
            tb.TextAlignment = TextAlignment.Center
            tb.Foreground = brush_hex(u"#7eb8d0") if self._is_beam_selected(i) else brush_hex(u"#95B8CC")
            tb.Cursor = Cursors.Hand
            tb.FontSize = typo.LABEL_FONT_PX
            lines = [u"{0}".format(beam.get("id")), u"{0:.1f} m · {1}".format(beam.get("len") or 0, beam.get("type"))]
            if apoyos_loaded:
                lines.append(u"{0}–{1}".format(beam.get("colStart"), beam.get("colEnd")))
            bid = beam.get("id")
            if bid in (session.empalme_beam_ids_sup or set()):
                lines.append(u"Emp sup @ {0:.1f} m".format(float(beam.get("len") or 0) * 0.5))
            if bid in (session.empalme_beam_ids_inf or set()):
                lines.append(u"Emp inf @ {0:.1f} m".format(float(beam.get("len") or 0) * 0.5))
            tb.Text = u"\n".join(lines)
            tb.Width = lay.pct_to_px(lay_i["widthPct"], content_w)
            Canvas.SetLeft(tb, cx - tb.Width * 0.5)
            Canvas.SetTop(tb, 2.0)
            idx_cap = i
            beam_cap = beam

            def _click_label(sender, args, idx=idx_cap, b=beam_cap):
                self._select_beam_from_elevation(idx, b, args)

            try:
                tb.MouseLeftButtonUp += MouseButtonEventHandler(_click_label)
            except Exception:
                pass
            cnv.Children.Add(tb)
        return cnv

    def _build_suple_sup_zone(self, beams, layouts, content_w):
        zone = StackPanel()
        zone.Width = content_w
        hint = label_small(u"Suple sup. · franja sobre alzado")
        zone.Children.Add(hint)
        strip = Canvas()
        strip.Width = content_w
        strip.Height = lay.SUPLE_SUP_CANVAS_SLOT_HEIGHT_PX + 4.0
        for i, beam in enumerate(beams):
            lay_i = layouts[i]
            left = lay.pct_to_px(lay_i["leftPct"], content_w)
            width = lay.pct_to_px(lay_i["widthPct"], content_w)
            slot = self._build_suple_sup_canvas_slot(beam, i, width)
            Canvas.SetLeft(slot, left)
            Canvas.SetTop(slot, 4.0)
            strip.Children.Add(slot)
        zone.Children.Add(strip)
        return zone

    def _build_estribo_zone(self, beams, layouts, tramos, content_w):
        zone = StackPanel()
        zone.Width = content_w
        hint = label_small(u"Suple inf. · franja bajo alzado")
        zone.Children.Add(hint)
        strip = Canvas()
        strip.Width = content_w
        strip.Height = lay.ESTRIBO_ZONE_HEIGHT_PX + 4.0
        for i, beam in enumerate(beams):
            lay_i = layouts[i]
            left = lay.pct_to_px(lay_i["leftPct"], content_w)
            width = lay.pct_to_px(lay_i["widthPct"], content_w)
            tramo = find_tramo_for_beam(tramos, i)
            tramo_beams = self._tramo_beams(tramo, beams) if tramo else [beam]
            slot = self._build_suple_canvas_slot(beam, i, width, tramo_beams)
            Canvas.SetLeft(slot, left)
            Canvas.SetTop(slot, 4.0)
            strip.Children.Add(slot)
        zone.Children.Add(strip)
        return zone

    def _build_section_stirrup_stack(self, beam, idx, session=None):
        """Cent → Ext → Confin. → Laterales apilados bajo preview de sección."""
        plan = compute_stirrup_zones(beam)
        dock = StackPanel()
        dock.MaxWidth = lay.SECTION_CTRL_WIDTH_PX

        dock.Children.Add(self._section_stack_order_hdr())

        n_sel = len(self.selected_beam_indices)
        if n_sel > 1:
            bulk = TextBlock()
            bulk.Text = u"Config. en lote · {0} vigas seleccionadas · cambios aplican a todas.".format(
                n_sel,
            )
            bulk.FontSize = typo.META_FONT_PX
            bulk.Foreground = brush_hex(u"#5bb8d4")
            bulk.FontWeight = FontWeights.SemiBold
            bulk.TextWrapping = TextWrapping.Wrap
            bulk.Margin = Thickness(0, 0, 0, 4)
            dock.Children.Add(bulk)

        if plan.get("mode") == u"single":
            role = u"uni" if plan.get("singleKind") == u"merge" else u"cent"
            z = (plan.get("zones") or [{}])[0]
            titles = {u"cent": u"Cent", u"uni": u"Único"}
            dock.Children.Add(
                self._section_ctrl_zone(beam, idx, role, titles.get(role, role), z, plan)
            )
        else:
            ext_len = plan.get("L_ext_each") or 0
            cent_z = None
            for z in plan.get("zones") or []:
                if z.get("role") == u"cent":
                    cent_z = z
                    break
            dock.Children.Add(
                self._section_cent_ext_pair(beam, idx, cent_z, ext_len, plan)
            )

        dock.Children.Add(self._confin_section_block(beam, idx))
        if session is not None:
            dock.Children.Add(self._laterales_section_block(session, idx))

        note = TextBlock()
        note.Text = u"Suple inf. permanece en la franja bajo el alzado."
        note.FontSize = typo.META_FONT_PX
        note.Foreground = brush_hex(u"#64748b")
        note.TextWrapping = TextWrapping.Wrap
        note.Margin = Thickness(0, 6, 0, 0)
        dock.Children.Add(note)
        return dock

    def _section_stack_order_hdr(self):
        """Cabecera Cent → Ext → Confin. + nota ancho panel."""
        wrap = StackPanel()
        wrap.Margin = Thickness(0, 2, 0, 1)

        top = Border()
        top.BorderBrush = brush_hex(u"#21465C", 140)
        top.BorderThickness = Thickness(0, 1, 0, 0)
        top.Margin = Thickness(0, 0, 0, 0)
        wrap.Children.Add(top)

        row = Grid()
        row.Margin = Thickness(0, 5, 0, 3)
        col_flow = ColumnDefinition()
        col_flow.Width = GridLength(1.0, GridUnitType.Star)
        col_note = ColumnDefinition()
        col_note.Width = GridLength.Auto
        row.ColumnDefinitions.Add(col_flow)
        row.ColumnDefinitions.Add(col_note)

        flow = StackPanel()
        flow.Orientation = Orientation.Horizontal
        flow.VerticalAlignment = VerticalAlignment.Center

        def _flow_part(text, color):
            tb = TextBlock()
            tb.Text = text
            tb.FontSize = typo.META_FONT_PX
            tb.FontWeight = FontWeights.Bold
            tb.Foreground = brush_hex(color)
            tb.Margin = Thickness(0, 0, 0, 0)
            return tb

        flow.Children.Add(_flow_part(u"Cent", u"#34d399"))
        sep1 = TextBlock()
        sep1.Text = u" → "
        sep1.FontSize = typo.META_FONT_PX
        sep1.FontWeight = FontWeights.Bold
        sep1.Foreground = brush_hex(u"#64748b")
        flow.Children.Add(sep1)
        flow.Children.Add(_flow_part(u"Ext", u"#fbbf24"))
        sep2 = TextBlock()
        sep2.Text = u" → "
        sep2.FontSize = typo.META_FONT_PX
        sep2.FontWeight = FontWeights.Bold
        sep2.Foreground = brush_hex(u"#64748b")
        flow.Children.Add(sep2)
        flow.Children.Add(_flow_part(u"Confin.", u"#5bb8d4"))
        Grid.SetColumn(flow, 0)
        row.Children.Add(flow)

        note = TextBlock()
        note.Text = u"ø|@ · {0:.0f}px".format(lay.SECTION_CTRL_WIDTH_PX)
        note.FontSize = typo.META_FONT_PX
        note.Foreground = brush_hex(u"#64748b", 217)
        note.VerticalAlignment = VerticalAlignment.Center
        note.TextAlignment = TextAlignment.Right
        Grid.SetColumn(note, 1)
        row.Children.Add(note)
        wrap.Children.Add(row)
        return wrap

    def _section_zone_len_hint(self, role, zone, plan):
        if not zone or not zone.get("lenMm"):
            return u""
        mm = zone.get("lenMm")
        if role == u"cent" and plan.get("mode") == u"triple":
            return u"L {0} mm · luz libre".format(mm)
        if role == u"ext":
            return u"L {0} mm ×2".format(mm)
        if role == u"uni":
            return u"L {0} mm · único".format(mm)
        return u"L {0} mm".format(mm)

    def _section_field_lbl(self, text):
        lbl = TextBlock()
        lbl.Text = text or u""
        lbl.FontSize = typo.LABEL_FONT_PX
        lbl.FontWeight = FontWeights.Bold
        lbl.Foreground = brush_hex(u"#64748b")
        lbl.HorizontalAlignment = HorizontalAlignment.Center
        lbl.VerticalAlignment = VerticalAlignment.Center
        return lbl

    def _section_at_row(self, beam, role, half=False):
        """Fila compacta ø | combo | @ | stepper | mm (stepper sin estirar)."""
        grid = Grid()
        grid.Margin = Thickness(4, 0, 4, 5) if half else Thickness(6, 0, 6, 5)
        diam_w = 44.0 if half else 48.0
        col_widths = (
            (10.0, False),
            (diam_w, False),
            (10.0, False),
            (None, False),
            (14.0 if half else 18.0, False),
        )
        for w, star in col_widths:
            col = ColumnDefinition()
            if star:
                col.Width = GridLength(1.0, GridUnitType.Star)
            elif w is None:
                col.Width = GridLength.Auto
            else:
                col.Width = GridLength(w)
            grid.ColumnDefinitions.Add(col)

        lbl_o = self._section_field_lbl(u"ø")
        Grid.SetColumn(lbl_o, 0)
        grid.Children.Add(lbl_o)

        if role in (u"ext", u"uni"):
            diam_val = beam.get("estExtDiam") or 10
            spacing_val = beam.get("estExtSpacing") or ESTRIBO_SPACING_DEFAULT_EXT
            diam_cb = make_diam_combo(
                self._win, diam_val, ESTRIBO_DIAM_OPTS,
                lambda v, b=beam: self._set_beam_field(b, "estExtDiam", v),
                compact=True,
            )
            spacing_cb = make_spacing_input(
                self._win, spacing_val,
                lambda v, b=beam: self._set_beam_field(b, "estExtSpacing", v),
                compact=True,
                width=48.0 if half else 52.0,
            )
        else:
            diam_val = beam.get("estCentDiam") or 8
            spacing_val = beam.get("estCentSpacing") or ESTRIBO_SPACING_DEFAULT_CENT
            diam_cb = make_diam_combo(
                self._win, diam_val, ESTRIBO_DIAM_OPTS,
                lambda v, b=beam: self._set_beam_field(b, "estCentDiam", v),
                compact=True,
            )
            spacing_cb = make_spacing_input(
                self._win, spacing_val,
                lambda v, b=beam: self._set_beam_field(b, "estCentSpacing", v),
                compact=True,
                width=48.0 if half else 52.0,
            )

        diam_cb.Width = diam_w
        diam_cb.MinWidth = diam_w
        diam_cb.MaxWidth = diam_w
        Grid.SetColumn(diam_cb, 1)
        grid.Children.Add(diam_cb)

        lbl_a = self._section_field_lbl(u"@")
        Grid.SetColumn(lbl_a, 2)
        grid.Children.Add(lbl_a)

        spacing_cb.HorizontalAlignment = HorizontalAlignment.Left
        Grid.SetColumn(spacing_cb, 3)
        grid.Children.Add(spacing_cb)

        unit = TextBlock()
        unit.Text = u"mm"
        unit.FontSize = typo.META_FONT_PX
        unit.Foreground = brush_hex(u"#64748b")
        unit.VerticalAlignment = VerticalAlignment.Center
        unit.HorizontalAlignment = HorizontalAlignment.Center
        Grid.SetColumn(unit, 4)
        grid.Children.Add(unit)
        return grid

    def _attach_section_zone_select(self, panel, idx, role):
        def _select(sender, args, i=idx, r=role):
            self._handle_beam_select(i, args, role=r, update_zone=True)

        try:
            panel.MouseLeftButtonUp += MouseButtonEventHandler(_select)
        except Exception:
            pass

    def _section_cent_ext_pair(self, beam, idx, cent_z, ext_len, plan):
        """Cent y Ext en la misma fila (~50% cada uno) para evitar aire en steppers."""
        row = Grid()
        row.Margin = Thickness(0, 0, 0, 3)
        col_cent = ColumnDefinition()
        col_cent.Width = GridLength(1.0, GridUnitType.Star)
        col_ext = ColumnDefinition()
        col_ext.Width = GridLength(1.0, GridUnitType.Star)
        row.ColumnDefinitions.Add(col_cent)
        row.ColumnDefinitions.Add(col_ext)

        cent = self._section_ctrl_zone(
            beam, idx, u"cent", u"Cent", cent_z, plan, half=True, merge_edge=u"right",
        )
        ext = self._section_ctrl_zone(
            beam,
            idx,
            u"ext",
            u"Ext",
            {u"lenMm": ext_len},
            plan,
            half=True,
            merge_edge=u"left",
        )
        Grid.SetColumn(cent, 0)
        Grid.SetColumn(ext, 1)
        row.Children.Add(cent)
        row.Children.Add(ext)
        return row

    def _section_ctrl_zone(
        self, beam, idx, role, title, zone, plan, half=False, merge_edge=None,
    ):
        """Bloque compacto: cabecera (título + L) + fila ø|@."""
        sel = self._is_section_zone_selected(idx, role)
        accent, bg, _title_bg = _ZONE_ROLE_STYLE.get(role, (u"#5bb8d4", u"#071018", u"#0d2430"))
        len_hint = self._section_zone_len_hint(role, zone, plan)
        if role == u"ext":
            len_fg = brush_hex(accent, 179)
        elif role in (u"cent", u"uni"):
            len_fg = brush_hex(accent, 166)
        else:
            len_fg = brush_hex(u"#64748b", 230)

        panel = Border()
        panel.Margin = Thickness(0, 0, 0, 0 if half else 3)
        panel.Padding = Thickness(0, 0, 0, 0)
        thick = 2 if sel else 1
        if half and merge_edge == u"right":
            panel.BorderThickness = Thickness(thick, thick, 1, thick)
        elif half and merge_edge == u"left":
            panel.BorderThickness = Thickness(0, thick, thick, thick)
        else:
            panel.BorderThickness = Thickness(thick, thick, thick, thick)
        panel.BorderBrush = brush_hex(accent) if sel else brush_hex(accent, 89)
        panel.Background = brush_hex(bg)
        panel.HorizontalAlignment = HorizontalAlignment.Stretch
        panel.Cursor = Cursors.Hand

        sp = StackPanel()
        bar = Border()
        bar.Height = 2.0
        bar.Background = brush_hex(accent)
        sp.Children.Add(bar)

        if half:
            hdr = StackPanel()
            hdr.Margin = Thickness(4, 3, 4, 2)
            title_tb = TextBlock()
            title_tb.Text = title
            title_tb.FontSize = typo.TITLE_FONT_PX
            title_tb.FontWeight = FontWeights.Bold
            title_tb.Foreground = brush_hex(accent, 230)
            hdr.Children.Add(title_tb)
            if len_hint:
                len_tb = TextBlock()
                len_tb.Text = len_hint
                len_tb.FontSize = typo.META_FONT_PX
                len_tb.Foreground = len_fg
                len_tb.TextWrapping = TextWrapping.Wrap
                len_tb.Margin = Thickness(0, 1, 0, 0)
                hdr.Children.Add(len_tb)
        else:
            hdr = Grid()
            hdr.Margin = Thickness(6, 3, 6, 2)
            col_id = ColumnDefinition()
            col_id.Width = GridLength(1.0, GridUnitType.Star)
            col_len = ColumnDefinition()
            col_len.Width = GridLength.Auto
            hdr.ColumnDefinitions.Add(col_id)
            hdr.ColumnDefinitions.Add(col_len)

            title_tb = TextBlock()
            title_tb.Text = title
            title_tb.FontSize = typo.TITLE_FONT_PX
            title_tb.FontWeight = FontWeights.Bold
            title_tb.Foreground = brush_hex(accent, 230)
            Grid.SetColumn(title_tb, 0)
            hdr.Children.Add(title_tb)

            if len_hint:
                len_tb = TextBlock()
                len_tb.Text = len_hint
                len_tb.FontSize = typo.META_FONT_PX
                len_tb.Foreground = len_fg
                len_tb.TextAlignment = TextAlignment.Right
                len_tb.VerticalAlignment = VerticalAlignment.Center
                Grid.SetColumn(len_tb, 1)
                hdr.Children.Add(len_tb)

        sp.Children.Add(hdr)
        sp.Children.Add(self._section_at_row(beam, role, half=half))
        panel.Child = sp
        self._attach_section_zone_select(panel, idx, role)
        return panel

    def _build_suple_sup_canvas_slot(self, beam, idx, width):
        """Franja canvas: Suple sup. por viga."""
        slot = Border()
        slot.Width = width
        slot.Height = lay.SUPLE_SUP_CANVAS_SLOT_HEIGHT_PX
        slot.ClipToBounds = False
        selected = self._is_beam_selected(idx)
        apply_panel_chrome(slot, selected=selected, padding=5)
        slot.Cursor = Cursors.Hand

        inner = StackPanel()
        hdr = Grid()
        hdr.Margin = Thickness(0, 0, 0, 4)
        col_id = ColumnDefinition()
        col_id.Width = GridLength(1.0, GridUnitType.Star)
        col_badge = ColumnDefinition()
        col_badge.Width = GridLength.Auto
        hdr.ColumnDefinitions.Add(col_id)
        hdr.ColumnDefinitions.Add(col_badge)

        beam_lbl = TextBlock()
        beam_lbl.Text = lay.beam_canvas_label(idx)
        beam_lbl.FontSize = typo.META_FONT_PX
        beam_lbl.FontWeight = FontWeights.SemiBold
        beam_lbl.Foreground = th.brush_accent() if selected else th.brush_fg_lo()
        beam_lbl.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(beam_lbl, 0)
        hdr.Children.Add(beam_lbl)

        badge = make_role_badge(u"Suple sup.", u"suple")
        badge.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(badge, 1)
        hdr.Children.Add(badge)
        inner.Children.Add(hdr)
        inner.Children.Add(self._build_suple_sup_slot_compact_row(beam, idx))
        slot.Child = inner

        def _select_suple_slot(sender, args, i=idx):
            self._handle_beam_select(i, args, role=u"supleSup", update_zone=True)

        try:
            slot.MouseLeftButtonUp += MouseButtonEventHandler(_select_suple_slot)
        except Exception:
            pass
        return slot

    def _build_suple_sup_slot_compact_row(self, beam, idx):
        """Fila compacta Sí/No · Ini/Fin · ø · n — suple superior."""
        ensure_beam_suple_superior(beam)
        enabled = beam_suple_sup_enabled(beam)
        col = StackPanel()

        row = StackPanel()
        row.Orientation = Orientation.Horizontal
        row.HorizontalAlignment = HorizontalAlignment.Center

        toggle = make_yesno_toggle(
            self._win,
            enabled,
            lambda v, b=beam: self._set_beam_suple_field(b, "supleSupEnabled", v),
            compact=True,
        )
        toggle.Margin = Thickness(0, 0, 6, 0)
        row.Children.Add(toggle)

        diam_suple = int(beam.get("diamSupleSup") or 16)
        diam = make_diam_combo(
            self._win,
            diam_suple,
            _session_bar_diam_opts(self._last_session, diam_suple),
            lambda v, b=beam: self._set_beam_suple_field(b, "diamSupleSup", v),
            compact=True,
            enabled=enabled,
        )
        diam.Margin = Thickness(0, 0, 6, 0)
        row.Children.Add(diam)

        row.Children.Add(make_stepper(
            self._win,
            beam.get("nSupleSup") or BAR_COUNT_MIN,
            BAR_COUNT_MIN,
            BAR_COUNT_MAX,
            1,
            lambda v, b=beam: self._set_beam_suple_field(
                b, "nSupleSup", clamp_bar_count(v),
            ),
            compact=True,
            enabled=enabled,
        ))
        col.Children.Add(row)

        ends = StackPanel()
        ends.Orientation = Orientation.Horizontal
        ends.HorizontalAlignment = HorizontalAlignment.Center
        ends.Margin = Thickness(0, 3, 0, 0)
        ends.Children.Add(self._build_suple_sup_end_toggle(
            beam, u"Ini", "supleSupStartEnabled", enabled,
        ))
        ends.Children.Add(self._build_suple_sup_end_toggle(
            beam, u"Fin", "supleSupEndEnabled", enabled,
        ))
        col.Children.Add(ends)
        return col

    def _build_suple_sup_end_toggle(self, beam, label, field, master_enabled):
        """Toggle Sí/No por extremo (inicio/fin) del tramo."""
        wrap = StackPanel()
        wrap.Orientation = Orientation.Horizontal
        wrap.Margin = Thickness(0, 0, 8, 0)
        wrap.VerticalAlignment = VerticalAlignment.Center

        lbl = TextBlock()
        lbl.Text = label
        lbl.FontSize = typo.META_FONT_PX
        lbl.Foreground = th.brush_fg_lo()
        lbl.Margin = Thickness(0, 0, 3, 0)
        lbl.VerticalAlignment = VerticalAlignment.Center
        wrap.Children.Add(lbl)

        on = bool(beam.get(field))
        toggle = make_yesno_toggle(
            self._win,
            on,
            lambda v, b=beam, f=field: self._set_beam_suple_field(b, f, v),
            compact=True,
            enabled=master_enabled,
        )
        wrap.Children.Add(toggle)
        return wrap

    def _build_suple_canvas_slot(self, beam, idx, width, tramo_beams=None):
        """Franja canvas: Suple inf. por viga — chrome neutro + badge."""
        slot = Border()
        slot.Width = width
        slot.Height = lay.SUPLE_CANVAS_SLOT_HEIGHT_PX
        slot.ClipToBounds = False
        selected = self._is_beam_selected(idx)
        apply_panel_chrome(slot, selected=selected, padding=5)
        slot.Cursor = Cursors.Hand

        inner = StackPanel()

        hdr = Grid()
        hdr.Margin = Thickness(0, 0, 0, 4)
        col_id = ColumnDefinition()
        col_id.Width = GridLength(1.0, GridUnitType.Star)
        col_badge = ColumnDefinition()
        col_badge.Width = GridLength.Auto
        hdr.ColumnDefinitions.Add(col_id)
        hdr.ColumnDefinitions.Add(col_badge)

        beam_lbl = TextBlock()
        beam_lbl.Text = lay.beam_canvas_label(idx)
        beam_lbl.FontSize = typo.META_FONT_PX
        beam_lbl.FontWeight = FontWeights.SemiBold
        beam_lbl.Foreground = th.brush_accent() if selected else th.brush_fg_lo()
        beam_lbl.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(beam_lbl, 0)
        hdr.Children.Add(beam_lbl)

        badge = make_role_badge(u"Suple", u"suple")
        badge.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(badge, 1)
        hdr.Children.Add(badge)
        inner.Children.Add(hdr)

        inner.Children.Add(self._build_suple_slot_compact_row(beam, idx))
        slot.Child = inner

        def _select_suple_slot(sender, args, i=idx):
            self._handle_beam_select(i, args, role=u"suple", update_zone=True)

        try:
            slot.MouseLeftButtonUp += MouseButtonEventHandler(_select_suple_slot)
        except Exception:
            pass
        return slot

    def _build_suple_slot_compact_row(self, beam, idx):
        """Fila compacta Sí/No · ø · n (mockup homogeneidad)."""
        ensure_beam_suple_inferior(beam)
        enabled = beam_suple_inf_enabled(beam)
        row = StackPanel()
        row.Orientation = Orientation.Horizontal
        row.HorizontalAlignment = HorizontalAlignment.Center

        toggle = make_yesno_toggle(
            self._win,
            enabled,
            lambda v, b=beam: self._set_beam_suple_field(b, "supleInfEnabled", v),
            compact=True,
        )
        toggle.Margin = Thickness(0, 0, 6, 0)
        row.Children.Add(toggle)

        diam_suple = int(beam.get("diamSupleInf") or 16)
        diam = make_diam_combo(
            self._win,
            diam_suple,
            _session_bar_diam_opts(self._last_session, diam_suple),
            lambda v, b=beam: self._set_beam_suple_field(b, "diamSupleInf", v),
            compact=True,
            enabled=enabled,
        )
        diam.Margin = Thickness(0, 0, 6, 0)
        row.Children.Add(diam)

        row.Children.Add(make_stepper(
            self._win,
            beam.get("nSupleInf") or BAR_COUNT_MIN,
            BAR_COUNT_MIN,
            BAR_COUNT_MAX,
            1,
            lambda v, b=beam: self._set_beam_suple_field(
                b, "nSupleInf", clamp_bar_count(v),
            ),
            compact=True,
            enabled=enabled,
        ))
        return row

    def _stirrup_section_hdr(self, title, accent):
        """Encabezado de bloque (variante E): «Estribos» / «Suple»."""
        row = Grid()
        row.Margin = Thickness(0, 4 if title == u"Suple" else 0, 3, 0)
        row.Height = lay.ESTRIBO_SECTION_HDR_PX - 4.0
        col_txt = ColumnDefinition()
        col_txt.Width = GridLength.Auto
        col_line = ColumnDefinition()
        col_line.Width = GridLength(1.0, GridUnitType.Star)
        row.ColumnDefinitions.Add(col_txt)
        row.ColumnDefinitions.Add(col_line)
        tb = TextBlock()
        tb.Text = (title or u"").upper()
        tb.FontSize = typo.META_FONT_PX
        tb.FontWeight = FontWeights.Bold
        tb.Foreground = brush_hex(accent, 220)
        tb.Margin = Thickness(2, 0, 6, 0)
        tb.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(tb, 0)
        row.Children.Add(tb)
        line = Border()
        line.Height = 1.0
        line.Background = brush_hex(u"#21465C", 128)
        line.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(line, 1)
        row.Children.Add(line)
        return row

    def _confin_section_block(self, beam, idx):
        """Confin. en panel sección (layout compacto ctrl-zone)."""
        ensure_beam_confinement(beam)
        accent = u"#5bb8d4"
        bg = u"#071018"
        sel = self._is_section_zone_selected(idx, u"confin")
        n_bars = first_layer_bar_count(beam)

        wrap = Border()
        wrap.Margin = Thickness(0, 0, 0, 3)
        wrap.Padding = Thickness(0, 0, 0, 0)
        wrap.BorderBrush = brush_hex(accent) if sel else brush_hex(accent, 89)
        wrap.BorderThickness = Thickness(2 if sel else 1)
        wrap.Background = brush_hex(bg)
        wrap.HorizontalAlignment = HorizontalAlignment.Stretch
        wrap.Cursor = Cursors.Hand

        sp = StackPanel()
        bar = Border()
        bar.Height = 2.0
        bar.Background = brush_hex(accent)
        sp.Children.Add(bar)

        hdr = Grid()
        hdr.Margin = Thickness(6, 3, 6, 2)
        col_id = ColumnDefinition()
        col_id.Width = GridLength(1.0, GridUnitType.Star)
        col_len = ColumnDefinition()
        col_len.Width = GridLength.Auto
        hdr.ColumnDefinitions.Add(col_id)
        hdr.ColumnDefinitions.Add(col_len)

        title_tb = TextBlock()
        title_tb.Text = u"Confin."
        title_tb.FontSize = typo.TITLE_FONT_PX
        title_tb.FontWeight = FontWeights.Bold
        title_tb.Foreground = brush_hex(accent, 230)
        Grid.SetColumn(title_tb, 0)
        hdr.Children.Add(title_tb)

        len_tb = TextBlock()
        len_tb.Text = u"{0}b · vinculado preview".format(n_bars)
        len_tb.FontSize = typo.META_FONT_PX
        len_tb.Foreground = brush_hex(accent, 140)
        len_tb.TextAlignment = TextAlignment.Right
        len_tb.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(len_tb, 1)
        hdr.Children.Add(len_tb)
        sp.Children.Add(hdr)

        fields = Grid()
        fields.Margin = Thickness(6, 0, 6, 4)
        col_combo = ColumnDefinition()
        col_combo.Width = GridLength(1.0, GridUnitType.Star)
        fields.ColumnDefinitions.Add(col_combo)

        opts = [s["label"] for s in get_confin_scenarios(beam)]
        combo = make_string_combo(
            self._win,
            opts,
            beam.get("estConfin") or opts[0],
            lambda v, b=beam: self._set_beam_field(b, "estConfin", v),
            compact=True,
        )
        combo.Height = typo.CTRL_HEIGHT_PX
        combo.FontSize = typo.CTRL_FONT_PX
        combo.HorizontalAlignment = HorizontalAlignment.Stretch
        Grid.SetColumn(combo, 0)
        fields.Children.Add(combo)
        sp.Children.Add(fields)

        hint = TextBlock()
        hint.Text = u"Sección transversal · cambio actualiza SVG"
        hint.FontSize = typo.META_FONT_PX
        hint.Foreground = brush_hex(u"#64748b", 217)
        hint.Margin = Thickness(6, 0, 6, 4)
        hint.TextWrapping = TextWrapping.Wrap
        sp.Children.Add(hint)
        wrap.Child = sp

        self._attach_section_zone_select(wrap, idx, u"confin")
        return wrap

    def _laterales_section_block(self, session, idx):
        """Laterales del alma — lote global (checkbox + cantidad + ø)."""
        accent = u"#a78bfa"
        enabled = bool(getattr(session, "lateralesEnabled", False))
        sel = self._is_section_zone_selected(idx, u"laterales")

        wrap = Border()
        wrap.Margin = Thickness(0, 4, 0, 0)
        wrap.Padding = Thickness(0, 0, 0, 0)
        wrap.BorderBrush = brush_hex(accent) if sel else brush_hex(accent, 89)
        wrap.BorderThickness = Thickness(2 if sel else 1)
        wrap.Background = brush_hex(u"#0c0814")
        wrap.HorizontalAlignment = HorizontalAlignment.Stretch
        wrap.Cursor = Cursors.Hand

        sp = StackPanel()
        bar = Border()
        bar.Height = 2.0
        bar.Background = brush_hex(accent)
        sp.Children.Add(bar)

        hdr = Grid()
        hdr.Margin = Thickness(6, 3, 6, 2)
        col_id = ColumnDefinition()
        col_id.Width = GridLength(1.0, GridUnitType.Star)
        col_toggle = ColumnDefinition()
        col_toggle.Width = GridLength.Auto
        hdr.ColumnDefinitions.Add(col_id)
        hdr.ColumnDefinitions.Add(col_toggle)

        title_tb = TextBlock()
        title_tb.Text = u"Laterales"
        title_tb.FontSize = typo.TITLE_FONT_PX
        title_tb.FontWeight = FontWeights.Bold
        title_tb.Foreground = brush_hex(accent, 230)
        Grid.SetColumn(title_tb, 0)
        hdr.Children.Add(title_tb)

        toggle = make_yesno_toggle(
            self._win,
            enabled,
            lambda v, s=session: self._set_session_laterales_field(s, "lateralesEnabled", v),
            compact=True,
        )
        Grid.SetColumn(toggle, 1)
        toggle.VerticalAlignment = VerticalAlignment.Center
        hdr.Children.Add(toggle)
        sp.Children.Add(hdr)

        fields = Grid()
        fields.Margin = Thickness(6, 0, 6, 4)
        fields.IsEnabled = enabled
        col_n_lbl = ColumnDefinition()
        col_n_lbl.Width = GridLength(28.0)
        col_n = ColumnDefinition()
        col_n.Width = GridLength(1.0, GridUnitType.Star)
        col_o_lbl = ColumnDefinition()
        col_o_lbl.Width = GridLength(16.0)
        col_o = ColumnDefinition()
        col_o.Width = GridLength(1.0, GridUnitType.Star)
        fields.ColumnDefinitions.Add(col_n_lbl)
        fields.ColumnDefinitions.Add(col_n)
        fields.ColumnDefinitions.Add(col_o_lbl)
        fields.ColumnDefinitions.Add(col_o)

        n_lbl = self._section_field_lbl(u"n")
        Grid.SetColumn(n_lbl, 0)
        fields.Children.Add(n_lbl)
        n_step = make_stepper(
            self._win,
            int(getattr(session, "nLaterales", 1) or 1),
            LATERALES_COUNT_MIN,
            LATERALES_COUNT_MAX,
            1,
            lambda v, s=session: self._set_session_laterales_field(s, "nLaterales", v),
            compact=True,
        )
        Grid.SetColumn(n_step, 1)
        fields.Children.Add(n_step)

        o_lbl = self._section_field_lbl(u"ø")
        Grid.SetColumn(o_lbl, 2)
        fields.Children.Add(o_lbl)
        diam_lat = int(getattr(session, "diamLaterales", 16) or 16)
        diam_cb = make_diam_combo(
            self._win,
            diam_lat,
            _session_bar_diam_opts(session, diam_lat),
            lambda v, s=session: self._set_session_laterales_field(s, "diamLaterales", v),
            compact=True,
        )
        diam_cb.Height = typo.CTRL_HEIGHT_PX
        Grid.SetColumn(diam_cb, 3)
        fields.Children.Add(diam_cb)
        sp.Children.Add(fields)

        hint = TextBlock()
        hint.Text = u"Cara alma ±ancho · sugerido según altura del lote"
        hint.FontSize = typo.META_FONT_PX
        hint.Foreground = brush_hex(u"#64748b", 217)
        hint.Margin = Thickness(6, 0, 6, 4)
        hint.TextWrapping = TextWrapping.Wrap
        sp.Children.Add(hint)
        wrap.Child = sp

        self._attach_section_zone_select(wrap, idx, u"laterales")
        return wrap

    def _set_session_laterales_field(self, session, field, value):
        if session is None:
            return
        if field == "lateralesEnabled":
            session.lateralesEnabled = bool(value)
        elif field == "nLaterales":
            try:
                session.nLaterales = max(
                    LATERALES_COUNT_MIN,
                    min(LATERALES_COUNT_MAX, int(value)),
                )
            except Exception:
                session.nLaterales = LATERALES_COUNT_MIN
        elif field == "diamLaterales":
            try:
                session.diamLaterales = int(value)
            except Exception:
                pass
        self._cb.get("on_redraw", lambda: None)()

    def refresh_session_laterales_suggestion(self, session):
        if session is None:
            return
        session.nLaterales = suggest_n_laterales_from_beams(session.domain_beams)

    def _suple_inf_global_row(self, beam, idx, tramo_beams, full_width=False):
        """Fila suple inferior: toggle + ø + n por viga (variante E, sin hints)."""
        ensure_beam_suple_inferior(beam)
        enabled = beam_suple_inf_enabled(beam)
        sel = self._is_section_zone_selected(idx, u"suple")
        wrap = Border()
        wrap.Margin = Thickness(0, 3 if full_width else 6, 0, 0)
        wrap.Padding = Thickness(5, 5, 5, 5)
        wrap.Height = lay.SUPLE_INF_ROW_PX - 2.0
        accent = u"#c084fc"
        wrap.BorderBrush = brush_hex(accent) if sel else brush_hex(accent, 100)
        wrap.BorderThickness = Thickness(2 if sel else 1)
        wrap.Background = brush_hex(u"#0c0814")
        wrap.HorizontalAlignment = HorizontalAlignment.Stretch
        wrap.Cursor = Cursors.Hand

        sp = StackPanel()
        bar = Border()
        bar.Height = 2.0
        bar.Margin = Thickness(0, 0, 0, 5)
        try:
            from System.Windows.Media import LinearGradientBrush, GradientStop

            grad = LinearGradientBrush()
            grad.StartPoint = System.Windows.Point(0.0, 0.5)
            grad.EndPoint = System.Windows.Point(1.0, 0.5)
            grad.GradientStops.Add(GradientStop(Color.FromArgb(0, 192, 132, 252), 0.0))
            grad.GradientStops.Add(GradientStop(Color.FromArgb(255, 192, 132, 252), 0.1))
            grad.GradientStops.Add(GradientStop(Color.FromArgb(255, 192, 132, 252), 0.9))
            grad.GradientStops.Add(GradientStop(Color.FromArgb(0, 192, 132, 252), 1.0))
            bar.Background = grad
        except Exception:
            bar.Background = brush_hex(accent)
        sp.Children.Add(bar)

        lbl_style = brush_hex(u"#e8f4f8")
        toggle = make_yesno_toggle(
            self._win,
            enabled,
            lambda v, b=beam: self._set_beam_suple_field(b, "supleInfEnabled", v),
            compact=True,
        )
        toggle.HorizontalAlignment = HorizontalAlignment.Right
        toggle.Margin = Thickness(4, 0, 0, 0)

        grid = Grid()
        grid.Margin = Thickness(0, 1, 0, 0)
        rd0 = RowDefinition()
        rd0.Height = GridLength(lay.ZONE_PANEL_ROW_PX)
        rd1 = RowDefinition()
        rd1.Height = GridLength(lay.ZONE_PANEL_ROW_PX)
        grid.RowDefinitions.Add(rd0)
        grid.RowDefinitions.Add(rd1)

        col_lbl = ColumnDefinition()
        col_lbl.Width = GridLength(lay.SUPLE_LABEL_PX)
        col_ctrl = ColumnDefinition()
        col_ctrl.Width = GridLength(1.0, GridUnitType.Star)
        grid.ColumnDefinitions.Add(col_lbl)
        grid.ColumnDefinitions.Add(col_ctrl)

        row0_lbl = self._zone_field_label(u"Suple inf.")
        row0_lbl.Foreground = lbl_style
        row0_lbl.Width = lay.SUPLE_LABEL_PX
        row0_lbl.FontSize = typo.LABEL_FONT_PX
        try:
            from System.Windows import TextTrimming
            row0_lbl.TextTrimming = TextTrimming.CharacterEllipsis
        except Exception:
            pass
        Grid.SetRow(row0_lbl, 0)
        Grid.SetColumn(row0_lbl, 0)
        grid.Children.Add(row0_lbl)
        Grid.SetRow(toggle, 0)
        Grid.SetColumn(toggle, 1)
        toggle.VerticalAlignment = VerticalAlignment.Center
        grid.Children.Add(toggle)

        param_row = Grid()
        param_row.Margin = Thickness(0, 2, 0, 0)
        col_o_lbl = ColumnDefinition()
        col_o_lbl.Width = GridLength(20.0)
        col_o_ctrl = ColumnDefinition()
        col_o_ctrl.Width = GridLength(1.0, GridUnitType.Star)
        col_n_lbl = ColumnDefinition()
        col_n_lbl.Width = GridLength(18.0)
        col_n_ctrl = ColumnDefinition()
        col_n_ctrl.Width = GridLength.Auto
        param_row.ColumnDefinitions.Add(col_o_lbl)
        param_row.ColumnDefinitions.Add(col_o_ctrl)
        param_row.ColumnDefinitions.Add(col_n_lbl)
        param_row.ColumnDefinitions.Add(col_n_ctrl)

        lbl_o = self._zone_field_label(u"ø")
        lbl_o.Width = 20.0
        lbl_o.Foreground = brush_hex(u"#95b8cc")
        Grid.SetColumn(lbl_o, 0)
        param_row.Children.Add(lbl_o)

        diam_suple = int(beam.get("diamSupleInf") or 16)
        diam = make_diam_combo(
            self._win,
            diam_suple,
            _session_bar_diam_opts(self._last_session, diam_suple),
            lambda v, b=beam: self._set_beam_suple_field(b, "diamSupleInf", v),
            compact=True,
            enabled=enabled,
        )
        diam.Margin = Thickness(2, 0, 8, 0)
        Grid.SetColumn(diam, 1)
        param_row.Children.Add(diam)

        lbl_n = self._zone_field_label(u"n")
        lbl_n.Width = 18.0
        lbl_n.Foreground = brush_hex(u"#95b8cc")
        Grid.SetColumn(lbl_n, 2)
        param_row.Children.Add(lbl_n)

        stepper = make_stepper(
            self._win,
            beam.get("nSupleInf") or BAR_COUNT_MIN,
            BAR_COUNT_MIN,
            BAR_COUNT_MAX,
            1,
            lambda v, b=beam: self._set_beam_suple_field(
                b, "nSupleInf", clamp_bar_count(v),
            ),
            compact=True,
            enabled=enabled,
        )
        Grid.SetColumn(stepper, 3)
        param_row.Children.Add(stepper)

        Grid.SetRow(param_row, 1)
        Grid.SetColumn(param_row, 0)
        Grid.SetColumnSpan(param_row, 2)
        grid.Children.Add(param_row)
        sp.Children.Add(grid)
        wrap.Child = sp

        def _select_suple(sender, args, i=idx):
            self._handle_beam_select(i, args, role=u"suple", update_zone=True)

        try:
            wrap.MouseLeftButtonUp += MouseButtonEventHandler(_select_suple)
        except Exception:
            pass
        return wrap

    def _zone_field_label(self, text):
        lbl = TextBlock()
        lbl.Text = text or u""
        lbl.Width = lay.ZONE_PANEL_LABEL_PX
        lbl.FontSize = typo.LABEL_FONT_PX
        lbl.FontWeight = FontWeights.Bold
        lbl.Foreground = brush_hex(u"#95b8cc")
        lbl.VerticalAlignment = VerticalAlignment.Center
        return lbl

    def _zone_add_field(self, grid, row, label_text, control, label=None):
        lbl = label if label is not None else self._zone_field_label(label_text)
        Grid.SetRow(lbl, row)
        Grid.SetColumn(lbl, 0)
        grid.Children.Add(lbl)
        Grid.SetRow(control, row)
        Grid.SetColumn(control, 1)
        control.VerticalAlignment = VerticalAlignment.Center
        control.HorizontalAlignment = HorizontalAlignment.Stretch
        grid.Children.Add(control)

    def _zone_fields_grid(self, n_field_rows, with_footer):
        grid = Grid()
        grid.Margin = Thickness(0, 2, 0, 0)
        for _ in range(n_field_rows):
            rd = RowDefinition()
            rd.Height = GridLength(lay.ZONE_PANEL_ROW_PX)
            grid.RowDefinitions.Add(rd)
        if with_footer:
            rd_f = RowDefinition()
            rd_f.Height = GridLength(lay.ZONE_PANEL_FOOTER_PX)
            grid.RowDefinitions.Add(rd_f)
        col_lbl = ColumnDefinition()
        col_lbl.Width = GridLength(lay.ZONE_PANEL_LABEL_PX)
        col_ctrl = ColumnDefinition()
        col_ctrl.Width = GridLength(1.0, GridUnitType.Star)
        grid.ColumnDefinitions.Add(col_lbl)
        grid.ColumnDefinitions.Add(col_ctrl)
        return grid

    def _zone_panel(self, beam, idx, role, zone, plan, show_len_hint=False, stacked=False):
        sel = self._is_section_zone_selected(idx, role)
        accent, bg, title_bg = _ZONE_ROLE_STYLE.get(role, (u"#5bb8d4", u"#071018", u"#0d2430"))
        panel = Border()
        panel.Margin = Thickness(0, 0, 0, 4 if stacked else 3)
        panel.Padding = Thickness(5, 3, 5, 4)
        panel.BorderBrush = brush_hex(accent) if sel else brush_hex(accent, 100)
        panel.BorderThickness = Thickness(2 if sel else 1)
        panel.Background = brush_hex(bg)
        panel.Cursor = Cursors.Hand
        if stacked:
            panel.HorizontalAlignment = HorizontalAlignment.Stretch

        titles = {
            "ext": u"Ext · ini/fin" if stacked else u"Ext",
            "cent": u"Cent",
            "uni": u"Único",
        }
        sp = StackPanel()

        bar = Border()
        bar.Height = 2.0
        bar.Background = brush_hex(accent)
        bar.Margin = Thickness(0, 0, 0, 3)
        sp.Children.Add(bar)

        title = TextBlock()
        title.Text = titles.get(role, role)
        title.Foreground = brush_hex(accent)
        title.FontSize = typo.TITLE_FONT_PX
        title.FontWeight = FontWeights.Bold
        title.Padding = Thickness(2, 0, 2, 1)
        title.Background = brush_hex(title_bg)
        title.Margin = Thickness(0, 0, 0, 1)
        sp.Children.Add(title)

        has_len = bool(show_len_hint and zone and zone.get("lenMm"))
        grid = self._zone_fields_grid(2, with_footer=has_len)
        row_idx = 0

        if role in ("ext", "uni"):
            self._zone_add_field(grid, row_idx, u"ø", make_diam_combo(
                self._win, beam.get("estExtDiam") or 10, ESTRIBO_DIAM_OPTS,
                lambda v, b=beam: self._set_beam_field(b, "estExtDiam", v),
                compact=True,
            ))
            row_idx += 1
            self._zone_add_field(grid, row_idx, u"@", make_spacing_input(
                self._win, beam.get("estExtSpacing") or ESTRIBO_SPACING_DEFAULT_EXT,
                lambda v, b=beam: self._set_beam_field(b, "estExtSpacing", v),
                compact=True,
            ))
            row_idx += 1

        if role == "cent":
            self._zone_add_field(grid, row_idx, u"ø", make_diam_combo(
                self._win, beam.get("estCentDiam") or 8, ESTRIBO_DIAM_OPTS,
                lambda v, b=beam: self._set_beam_field(b, "estCentDiam", v),
                compact=True,
            ))
            row_idx += 1
            self._zone_add_field(grid, row_idx, u"@", make_spacing_input(
                self._win, beam.get("estCentSpacing") or ESTRIBO_SPACING_DEFAULT_CENT,
                lambda v, b=beam: self._set_beam_field(b, "estCentSpacing", v),
                compact=True,
            ))
            row_idx += 1

        if has_len:
            hint = TextBlock()
            if role == "ext" and plan.get("mode") == "triple":
                h_mm = section_height_mm(beam.get("type"))
                hint.Text = u"L {0} mm ×2 (ini/fin) · 2·h={1}".format(
                    zone.get("lenMm"), int(h_mm * 2),
                )
            elif role == "cent" and plan.get("mode") == "triple":
                hint.Text = u"L {0} mm".format(zone.get("lenMm"))
            elif role == "uni":
                hint.Text = u"L {0} mm · único".format(zone.get("lenMm"))
            else:
                hint.Text = u"L {0} mm".format(zone.get("lenMm"))
            hint.FontSize = typo.META_FONT_PX
            hint.Foreground = brush_hex(accent, 200) if role == "ext" else brush_hex(u"#64748b")
            hint.LineHeight = 11.0
            hint.VerticalAlignment = VerticalAlignment.Bottom
            Grid.SetRow(hint, row_idx)
            Grid.SetColumn(hint, 0)
            Grid.SetColumnSpan(hint, 2)
            grid.Children.Add(hint)

        sp.Children.Add(grid)
        panel.Child = sp

        def _select(sender, args, i=idx, r=role):
            self._handle_beam_select(i, args, role=r, update_zone=True)

        try:
            panel.MouseLeftButtonUp += MouseButtonEventHandler(_select)
        except Exception:
            pass
        return panel

    def _empalme_ids_for_face(self, session, face):
        if face == u"sup":
            return session.empalme_beam_ids_sup or set()
        return session.empalme_beam_ids_inf or set()

    def _face_chip_text(self, tramos):
        if not tramos:
            return u"—"
        labels = u" · ".join(
            u"T{0}".format(t.get("id")) for t in tramos
        )
        return u"{0} tramo(s): {1}".format(len(tramos), labels)

    def _build_face_header(self, tramos, is_sup):
        accent = u"#22d3ee" if is_sup else u"#fb7185"
        hdr = Grid()
        hdr.Margin = Thickness(lay.FACE_BLOCK_PAD_PX, lay.FACE_BLOCK_PAD_PX, lay.FACE_BLOCK_PAD_PX, 0)

        col_title = ColumnDefinition()
        col_title.Width = GridLength.Auto
        col_chip = ColumnDefinition()
        col_chip.Width = GridLength.Auto
        col_rule = ColumnDefinition()
        col_rule.Width = GridLength(1.0, GridUnitType.Star)
        hdr.ColumnDefinitions.Add(col_title)
        hdr.ColumnDefinitions.Add(col_chip)
        hdr.ColumnDefinitions.Add(col_rule)

        title = TextBlock()
        title.Text = u"Armadura superior" if is_sup else u"Armadura inferior"
        title.Foreground = brush_hex(accent)
        title.FontSize = typo.HDR_FONT_PX
        title.FontWeight = FontWeights.Bold
        Grid.SetColumn(title, 0)
        hdr.Children.Add(title)

        chip = Border()
        chip.Margin = Thickness(8, 0, 0, 0)
        chip.Padding = Thickness(8, 3, 8, 3)
        chip.CornerRadius = System.Windows.CornerRadius(10)
        chip.BorderBrush = brush_hex(accent, 89)
        chip.BorderThickness = Thickness(1)
        chip.Background = brush_hex(accent, 20 if is_sup else 16)
        chip_tb = TextBlock()
        chip_tb.Text = self._face_chip_text(tramos)
        chip_tb.Foreground = brush_hex(accent)
        chip_tb.FontSize = typo.LABEL_FONT_PX
        chip_tb.FontWeight = FontWeights.SemiBold
        chip.Child = chip_tb
        Grid.SetColumn(chip, 1)
        hdr.Children.Add(chip)

        rule = TextBlock()
        rule.Text = (
            u"Fusión: mismo ancho + colinealidad"
            if is_sup
            else u"Fusión: sección ancho×alto"
        )
        rule.Foreground = brush_hex(u"#64748b")
        rule.FontSize = typo.META_FONT_PX
        rule.TextAlignment = TextAlignment.Right
        rule.TextWrapping = TextWrapping.Wrap
        rule.VerticalAlignment = VerticalAlignment.Center
        rule.Margin = Thickness(12, 0, 0, 0)
        Grid.SetColumn(rule, 2)
        hdr.Children.Add(rule)
        return hdr

    def _build_lane_label(self, text, is_sup, dot_color=None):
        row = StackPanel()
        row.Orientation = Orientation.Horizontal
        row.Margin = Thickness(lay.FACE_BLOCK_PAD_PX, lay.LANE_GAP_PX, lay.FACE_BLOCK_PAD_PX, 0)

        if dot_color:
            dot = Border()
            dot.Width = 6.0
            dot.Height = 6.0
            dot.CornerRadius = System.Windows.CornerRadius(3)
            dot.Background = brush_hex(dot_color)
            dot.Margin = Thickness(0, 0, 6, 0)
            dot.VerticalAlignment = VerticalAlignment.Center
            row.Children.Add(dot)

        lbl = TextBlock()
        lbl.Text = text or u""
        lbl.Foreground = brush_hex(u"#64748b")
        lbl.FontSize = typo.LANE_FONT_PX
        lbl.FontWeight = FontWeights.Bold
        lbl.VerticalAlignment = VerticalAlignment.Center
        row.Children.Add(lbl)
        return row

    def _build_empalme_button_row(self, beams, layouts, session, content_w, face):
        """Botones Traslape alineados con vigas dentro del carril dedicado."""
        is_sup = face == u"sup"
        accent = u"#22d3ee" if is_sup else u"#fb7185"
        empalme_set = self._empalme_ids_for_face(session, face)
        row = Canvas()
        row.Width = content_w
        row.Height = lay.TRAMO_EMPALME_ROW_PX - 8.0
        btn_h = row.Height - 2.0
        for i, beam in enumerate(beams):
            lay_i = layouts[i]
            left = lay.pct_to_px(lay_i["leftPct"], content_w)
            width = lay.pct_to_px(lay_i["widthPct"], content_w)
            empalme_on = beam.get("id") in empalme_set
            btn_w = min(max(72.0, width * 0.55), max(48.0, width - 4.0))
            btn = Button()
            btn.Content = u"Traslapo ✓" if empalme_on else u"Traslapo"
            btn.Width = btn_w
            btn.Height = btn_h
            btn.Padding = Thickness(8, 0, 8, 0)
            btn.FontSize = typo.CTRL_FONT_PX
            btn.FontWeight = FontWeights.SemiBold
            btn.Cursor = Cursors.Hand
            btn.BorderThickness = Thickness(1)
            if empalme_on:
                btn.Background = brush_hex(u"#fbbf24", 26)
                btn.Foreground = brush_hex(u"#fbbf24")
                btn.BorderBrush = brush_hex(u"#fbbf24")
                btn.ToolTip = u"Quitar traslape @ mitad · fibra {0}".format(
                    u"superior" if is_sup else u"inferior"
                )
            else:
                btn.Background = brush_hex(u"#071018")
                btn.Foreground = brush_hex(u"#95b8cc")
                btn.BorderBrush = brush_hex(u"#21465C")
                btn.ToolTip = u"Marcar traslape @ mitad · fibra {0}".format(
                    u"superior" if is_sup else u"inferior"
                )
            Canvas.SetLeft(btn, left + (width - btn_w) * 0.5)
            Canvas.SetTop(btn, 1.0)
            beam_id = beam.get("id")
            idx_cap = i

            def _click_empalme(sender, args, i=idx_cap, bid=beam_id, f=face):
                self._handle_beam_select(i, args, update_zone=False, redraw=False)
                self._cb.get("on_toggle_empalme", lambda _b, _f: None)(bid, f)
                self._cb.get("on_redraw", lambda: None)()

            try:
                btn.Click += RoutedEventHandler(_click_empalme)
            except Exception:
                pass
            row.Children.Add(btn)
        return row

    def _build_traslape_strip(self, beams, layouts, session, content_w, face):
        strip = Border()
        strip.Margin = Thickness(lay.FACE_BLOCK_PAD_PX, 0, lay.FACE_BLOCK_PAD_PX, 0)
        strip.Padding = Thickness(0, 4, 0, 4)
        strip.Height = lay.TRAMO_EMPALME_ROW_PX
        strip.Background = brush_hex(u"#071018", 128)
        strip.BorderBrush = brush_hex(u"#21465C", 102)
        strip.BorderThickness = Thickness(1)
        strip.CornerRadius = System.Windows.CornerRadius(4)
        strip.Child = self._build_empalme_button_row(
            beams, layouts, session, content_w, face,
        )
        return strip

    def _build_tramo_bands_row(self, tramos, layouts, content_w, face, accent_default):
        is_sup = face == u"sup"
        sel_id = self.selected_tramo_sup_id if is_sup else self.selected_tramo_inf_id
        bands = Canvas()
        bands.Width = content_w
        bands.Height = lay.TRAMO_BAND_HEIGHT_PX
        bands.Margin = Thickness(lay.FACE_BLOCK_PAD_PX, 0, lay.FACE_BLOCK_PAD_PX, 0)
        for tramo in tramos or []:
            span = lay.tramo_span(layouts, tramo, content_w)
            band = Border()
            band.Width = lay.pct_to_px(span["widthPct"], content_w)
            band.Height = lay.TRAMO_BAND_HEIGHT_PX - 2.0
            Canvas.SetLeft(band, lay.pct_to_px(span["leftPct"], content_w))
            Canvas.SetTop(band, 1.0)
            accent = tramo.get("accent") or accent_default
            sel = sel_id == tramo["id"]
            band.Background = brush_hex(accent, 200 if sel else 140)
            band.BorderBrush = brush_hex(accent if sel else accent, 220)
            band.BorderThickness = Thickness(2 if sel else 1)
            band.CornerRadius = System.Windows.CornerRadius(2)
            band.Cursor = Cursors.Hand
            if tramo_exceeds_bar_limit(tramo):
                band.BorderBrush = brush_hex(u"#fbbf24")
            tid = tramo["id"]

            def _select_band(sender, args, tramo_id=tid, tramo_face=face):
                if tramo_face == u"sup":
                    self.selected_tramo_sup_id = tramo_id
                else:
                    self.selected_tramo_inf_id = tramo_id
                self._cb.get("on_select_tramo", lambda _t, _f: None)(tramo_id, tramo_face)
                self._cb.get("on_redraw", lambda: None)()

            try:
                band.MouseLeftButtonUp += MouseButtonEventHandler(_select_band)
            except Exception:
                pass
            bands.Children.Add(band)
        return bands

    def _build_panel_lane(self, tramos, beams, layouts, content_w, face):
        is_sup = face == u"sup"
        sel_id = self.selected_tramo_sup_id if is_sup else self.selected_tramo_inf_id
        accent_default = u"#22d3ee" if is_sup else u"#fb7185"

        wrap = Border()
        wrap.Margin = Thickness(lay.FACE_BLOCK_PAD_PX, 0, lay.FACE_BLOCK_PAD_PX, lay.FACE_BLOCK_PAD_PX)
        wrap.Padding = Thickness(0, 5, 0, 4)
        wrap.BorderBrush = brush_hex(u"#21465C", 115)
        wrap.BorderThickness = Thickness(0, 1, 0, 0)

        sv = ScrollViewer()
        sv.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto
        sv.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled
        sv.Height = lay.TRAMO_PANEL_LANE_PX
        sv.ClipToBounds = True

        row = StackPanel()
        row.Orientation = Orientation.Horizontal
        row.Margin = Thickness(2, lay.LANE_GAP_PX, 2, 4)

        for tramo in tramos or []:
            panel = self._build_face_tramo_panel(tramo, beams, layouts, content_w, face)
            panel.Width = lay.TRAMO_PANEL_W_PX
            panel.MinWidth = lay.TRAMO_PANEL_W_PX
            panel.MaxWidth = lay.TRAMO_PANEL_W_PX
            panel.Margin = Thickness(0, 0, lay.LANE_GAP_PX, 0)
            panel.ClipToBounds = True
            accent = tramo.get("accent") or accent_default
            if sel_id == tramo["id"]:
                panel.BorderBrush = brush_hex(accent)
                panel.BorderThickness = Thickness(2)
            row.Children.Add(panel)

        sv.Content = row
        wrap.Child = sv
        return wrap

    def _build_face_tramo_zone(self, beams, layouts, tramos, content_w, face, session=None):
        is_sup = face == u"sup"
        accent_default = u"#22d3ee" if is_sup else u"#fb7185"

        zone = StackPanel()
        zone.Width = content_w
        zone.Background = brush_hex(accent_default, 8 if is_sup else 6)
        if not is_sup:
            sep = Border()
            sep.Height = 1.0
            sep.BorderBrush = brush_hex(u"#fb7185", 56)
            sep.BorderThickness = Thickness(0, 1, 0, 0)
            sep.Margin = Thickness(lay.FACE_BLOCK_PAD_PX, 0, lay.FACE_BLOCK_PAD_PX, 0)
            zone.Children.Add(sep)

        zone.Children.Add(self._build_face_header(tramos, is_sup))
        zone.Children.Add(
            self._build_lane_label(
                u"Carril Traslape", is_sup, dot_color=u"#06b6d4" if is_sup else u"#f472b6",
            )
        )
        if session is not None:
            zone.Children.Add(
                self._build_traslape_strip(beams, layouts, session, content_w, face)
            )
        zone.Children.Add(
            self._build_lane_label(u"Bandas Tn", is_sup, dot_color=accent_default)
        )
        zone.Children.Add(
            self._build_tramo_bands_row(tramos, layouts, content_w, face, accent_default)
        )
        zone.Children.Add(
            self._build_lane_label(u"Controladores Tn · scroll horizontal", is_sup)
        )
        zone.Children.Add(
            self._build_panel_lane(tramos, beams, layouts, content_w, face)
        )
        return zone

    def _build_face_tramo_panel(self, tramo, beams, layouts, content_w, face):
        is_sup = face == u"sup"
        side = u"sup" if is_sup else u"inf"
        tramo_beams = self._tramo_beams(tramo, beams)
        owner = tramo_beams[0] if tramo_beams else beams[0]
        for beam in tramo_beams:
            ensure_beam_layers(beam)
        warn = tramo_exceeds_bar_limit(tramo)
        sel_id = self.selected_tramo_sup_id if is_sup else self.selected_tramo_inf_id
        sel = sel_id == tramo["id"]
        accent = tramo.get("accent") or (u"#22d3ee" if is_sup else u"#fb7185")

        panel = Border()
        panel.Padding = Thickness(4, 4, 4, 4)
        panel.HorizontalAlignment = HorizontalAlignment.Center
        panel.BorderBrush = brush_hex(accent) if sel else brush_hex(u"#21465C")
        panel.BorderThickness = Thickness(2 if sel else 1)
        panel.Background = brush_hex(u"#071018")
        panel.Cursor = Cursors.Hand

        sp = StackPanel()
        sp.HorizontalAlignment = HorizontalAlignment.Center
        bar = Border()
        bar.Height = 2.0
        bar.Background = brush_hex(accent)
        bar.Margin = Thickness(0, 0, 0, 3)
        sp.Children.Add(bar)

        title = TextBlock()
        title.Text = tramo.get("label") or u"T{0}".format(tramo.get("id"))
        title.Foreground = brush_hex(u"#e8f4f8")
        title.FontSize = typo.TITLE_FONT_PX
        title.FontWeight = FontWeights.Bold
        title.TextWrapping = TextWrapping.Wrap
        title.Margin = Thickness(0, 0, 0, 2)
        sp.Children.Add(title)

        meta = TextBlock()
        meta.Text = tramo.get("section") or u""
        meta.Foreground = brush_hex(u"#64748b")
        meta.FontSize = typo.META_FONT_PX
        meta.Margin = Thickness(0, 0, 0, 3)
        sp.Children.Add(meta)

        sp.Children.Add(self._cap_col(tramo_beams, owner, side))

        if warn:
            wtb = TextBlock()
            wtb.Text = u"Barra presunta > 12 m"
            wtb.Foreground = brush_hex(u"#fbbf24")
            wtb.FontSize = typo.META_FONT_PX
            wtb.FontWeight = FontWeights.Bold
            wtb.Margin = Thickness(0, 2, 0, 0)
            sp.Children.Add(wtb)

        panel.Child = sp

        tid = tramo["id"]

        def _select(sender, args, tramo_id=tid, tramo_face=face):
            if tramo_face == u"sup":
                self.selected_tramo_sup_id = tramo_id
            else:
                self.selected_tramo_inf_id = tramo_id
            self._cb.get("on_select_tramo", lambda _t, _f: None)(tramo_id, tramo_face)
            self._cb.get("on_redraw", lambda: None)()

        try:
            panel.MouseLeftButtonUp += MouseButtonEventHandler(_select)
        except Exception:
            pass
        return panel

    def _tramo_beams(self, tramo, beams):
        out = []
        for idx in tramo.get("beamIndices") or []:
            if 0 <= int(idx) < len(beams):
                out.append(beams[int(idx)])
        return out

    # ø por capa: solo dentro del mismo tramo Tn. Cantidades/capas: lote completo.
    _TRAMO_LOCAL_FIELDS = (
        "diamSup",
        "diamInf",
        "diamSup2",
        "diamInf2",
        "diamSup3",
        "diamInf3",
    )

    def _all_domain_beams(self, tramo_beams=None):
        if self._last_beams:
            return self._last_beams
        return list(tramo_beams or [])

    def _sync_tramo_local_config(self, tramo_beams, owner):
        if not tramo_beams or owner is None:
            return
        for beam in tramo_beams:
            if beam is owner:
                continue
            for field in self._TRAMO_LOCAL_FIELDS:
                if field in owner:
                    beam[field] = owner[field]

    def _refresh_tramo_beams_layer_state(self, tramo_beams):
        for beam in tramo_beams or []:
            ensure_beam_layers(beam)
            ensure_beam_confinement(beam)

    def _align_section_beam_for_tramo_edit(self, tramo_beams, owner):
        """Preview sección sigue la viga del tramo editado (capas/cantidades)."""
        beams = self._last_beams or []
        sel = self.selected_beam_idx
        if 0 <= sel < len(beams) and beams[sel] in (tramo_beams or []):
            return
        for i, beam in enumerate(beams):
            if beam is owner:
                self.selected_beam_idx = i
                return

    def _set_beam_suple_field(self, beam, field, value):
        if beam is None:
            return
        for b in self._targets_for_beam_edit(beam):
            b[field] = value
            if field.startswith("supleInf") or field == "supleInfEnabled":
                ensure_beam_suple_inferior(b)
            elif field.startswith("supleSup") or field == "supleSupEnabled":
                ensure_beam_suple_superior(b)
                if field == "supleSupEnabled" and value:
                    b["supleSupStartEnabled"] = True
                    b["supleSupEndEnabled"] = True
        self._cb.get("on_redraw", lambda: None)()

    def _set_beam_field(self, beam, field, value):
        if beam is None:
            return
        for b in self._targets_for_beam_edit(beam):
            if field in ("nSup", "nInf"):
                set_first_layer_bar_count(b, value)
                ensure_beam_layers(b)
                ensure_beam_confinement(b)
            else:
                b[field] = value
                if field == "estConfin":
                    ensure_beam_confinement(b)
        self._cb.get("on_redraw", lambda: None)()

    def _cap_col(self, tramo_beams, owner, side, tramo_accent=None, compact_band=False):
        is_sup = side == u"sup" or side == "sup"
        accent = tramo_accent or (u"#22d3ee" if is_sup else u"#f87171")

        wrap = Border()
        wrap.Margin = Thickness(0)
        if compact_band:
            wrap.Padding = Thickness(1, 1, 1, 1)
            wrap.BorderBrush = th.brush_border(140)
            wrap.BorderThickness = Thickness(1)
            wrap.Background = th.brush_panel(0)
        else:
            wrap.Padding = Thickness(2, 2, 2, 2)
            wrap.BorderBrush = brush_hex(accent, 90)
            wrap.BorderThickness = Thickness(1)
            wrap.Background = th.brush_panel(180)

        col = StackPanel()
        beam = owner
        cap_row = StackPanel()
        cap_row.Orientation = Orientation.Horizontal
        cap_row.HorizontalAlignment = HorizontalAlignment.Center
        cap_row.Margin = Thickness(0, 0, 0, 2)
        cap_lbl = label_small(u"Capas")
        cap_lbl.Margin = Thickness(0, 0, 3, 0)
        cap_lbl.VerticalAlignment = VerticalAlignment.Center
        cap_row.Children.Add(cap_lbl)
        n_capas = beam_n_capas_sup(beam) if is_sup else beam_n_capas_inf(beam)
        cap_row.Children.Add(make_capas_stepper(
            self._win,
            n_capas,
            lambda v, tb=tramo_beams, s=side: self._set_capas_side(tb, s, v),
            compact=True,
        ))
        col.Children.Add(cap_row)

        for layer_num in range(1, CAPAS_MAX + 1):
            k = layer_keys(layer_num)
            qty_field = k["nSup"] if is_sup else k["nInf"]
            diam_field = k["diamSup"] if is_sup else k["diamInf"]
            active = layer_num <= n_capas
            if compact_band and not active:
                continue
            row = StackPanel()
            row.Orientation = Orientation.Horizontal
            row.Margin = Thickness(0, 0, 0, 0)
            row.HorizontalAlignment = HorizontalAlignment.Center
            lbl = label_small(u"{0}".format(k["label"]))
            lbl.Width = 16.0
            lbl.Margin = Thickness(0, 0, 2, 0)
            lbl.VerticalAlignment = VerticalAlignment.Center
            lbl.Foreground = accent_soft_brush(accent, "text") if active else brush_hex(u"#64748b")
            if active:
                lbl.FontWeight = FontWeights.SemiBold
            row.Children.Add(lbl)
            row.Children.Add(make_stepper(
                self._win,
                beam.get(qty_field) or BAR_COUNT_MIN,
                BAR_COUNT_MIN,
                BAR_COUNT_MAX,
                1,
                lambda v, tb=tramo_beams, o=owner, f=qty_field: self._set_tramo_beams_field(
                    tb, f, clamp_bar_count(v), confinement=(f in ("nSup", "nInf")), owner=o,
                ),
                compact=True,
                enabled=active,
            ))
            diam_val = int(beam.get(diam_field) or 16)
            row.Children.Add(make_diam_combo(
                self._win,
                diam_val,
                _session_bar_diam_opts(self._last_session, diam_val),
                lambda v, tb=tramo_beams, o=owner, f=diam_field: self._set_tramo_beams_field(
                    tb, f, v, owner=o,
                ),
                compact=True,
                enabled=active,
            ))
            if not compact_band and not active:
                row.Opacity = 0.55
            col.Children.Add(row)

        wrap.Child = col
        return wrap

    def _set_tramo_beams_field(self, tramo_beams, field, value, confinement=False, owner=None):
        owner = owner or ((tramo_beams or [None])[0])
        all_beams = self._all_domain_beams(tramo_beams)
        if is_global_layer_sync_field(field):
            sync_layer_field_all_beams(all_beams, field, value)
        elif owner is not None:
            owner[field] = value
            self._sync_tramo_local_config(tramo_beams, owner)
            self._refresh_tramo_beams_layer_state(tramo_beams)
        self._align_section_beam_for_tramo_edit(tramo_beams, owner)
        self._cb.get("on_redraw", lambda: None)()

    def _set_capas_side(self, tramo_beams, side, n_capas):
        from armado_vigas.domain.constants import CAPAS_MIN, CAPAS_MAX
        field = "nCapasSup" if side == "sup" else "nCapasInf"
        n_val = max(CAPAS_MIN, min(CAPAS_MAX, int(n_capas)))
        owner = (tramo_beams or [None])[0]
        sync_layer_field_all_beams(self._all_domain_beams(tramo_beams), field, n_val)
        self._align_section_beam_for_tramo_edit(tramo_beams, owner)
        self._cb.get("on_redraw", lambda: None)()
