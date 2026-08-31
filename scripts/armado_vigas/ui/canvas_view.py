# -*- coding: utf-8 -*-
"""Renderizado WPF del canvas (alzado, estribos, tramos Tn, labels)."""

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

import System

from System.Windows import HorizontalAlignment, TextWrapping, Thickness, VerticalAlignment, FontWeights, TextAlignment
from System.Windows import (
    GridLength,
    GridUnitType,
    Point,
    ResizeMode,
    RoutedEventHandler,
    Visibility,
    Window,
    WindowStartupLocation,
)
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
from System.Windows.Input import (
    Cursors,
    Key,
    KeyEventHandler,
    Keyboard,
    ModifierKeys,
    Mouse,
    MouseButton,
    MouseButtonEventHandler,
    MouseEventHandler,
    MouseWheelEventHandler,
)
from System.Windows.Media import (
    DoubleCollection,
    SolidColorBrush,
    Color,
    ScaleTransform,
    PathFigure,
    PathGeometry,
    LineSegment,
    QuadraticBezierSegment,
    PenLineJoin,
    PenLineCap,
)
from System.Windows.Media import PointCollection
from System.Windows.Shapes import Line, Path, Polyline, Rectangle

from armado_vigas.domain.bar_ends import (
    BAR_END_MODE_AUTO,
    BAR_END_MODE_EMP,
    BAR_END_MODE_PATA_L,
    empotramiento_mm_for_diam,
    get_bar_end_mode,
)
from armado_vigas.domain.confinement import (
    clear_conf_draft,
    conf_draft_signature,
    ensure_beam_confinement,
    find_confin_def,
    get_conf_draft,
    get_confin_scenarios,
    is_conf_draft_defined,
    set_conf_draft,
    toggle_conf_estribo,
    toggle_conf_perimetral,
    toggle_conf_traba,
)
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
    LATERALES_DIAM_DEFAULT,
    session_n_laterales,
    suggest_n_laterales_from_beams,
)
from armado_vigas.domain.layers import (
    beam_layer_diam_inf,
    beam_layer_diam_sup,
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
    adjacent_beams_for_apoyo,
    beam_suple_sup_enabled,
    beam_suple_sup_side_enabled,
    ensure_beam_suple_superior,
    ensure_session_suple_sup,
    is_apoyo_suple_sup_on,
    apoyo_allows_suple_sup,
    suple_sup_segments_layout_px,
    toggle_apoyo_suple_sup,
    select_apoyo_suple_sup,
    clear_all_suple_sup_apoyos,
    set_apoyo_suple_sup,
    sync_beams_suple_from_apoyo_set,
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
from armado_vigas.ui import rail_cards
from armado_vigas.ui.wpf_controls import (
    accent_soft_brush,
    brush_hex,
    label_small,
    make_bar_count_combo,
    make_capas_combo,
    make_diam_combo,
    make_int_combo,
    make_spacing_input,
    make_string_combo,
    make_yesno_toggle,
)
from armado_vigas.ui.elev_geom_batch import ElevGeomBatch, apply_aliased_render
from armado_vigas.ui.net_ui import freeze_freezable


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
# Suple INF = capa interior (y menor), no fuera del borde inferior.
_ELEV_BAR_SUPLE_Y = _ELEV_BAR_INF_Y - 11.0 * _ELEV_SCALE
_ELEV_COL_TOP = 4.0
_ELEV_COL_H = lay.ELEVATION_HEIGHT_PX - 10.0
_ELEV_COL_W = 14.0
_ELEV_WALL_W = 18.0
_ELEV_POCKET_D = 12.0
_ELEV_REF_SECTION_H_CM = 60.0
_ELEV_COL_PX_PER_MM = _ELEV_COL_W / 300.0
_ELEV_WALL_PX_PER_MM = _ELEV_WALL_W / 200.0
_ELEV_BREAK_AMP = 2.0 * _ELEV_SCALE
_ELEV_STROKE_BAR = 1.0  # legacy; alzado ya no dibuja barras (grosor por ø)

# Empalme Opción C (preview): solape simbólico.
# SUP: solo la barra saliente se desvía, siempre hacia abajo (y+δ en Canvas).
# La entrante permanece en la fibra (colinear).
_ELEV_EMP_STAGGER_DY = 2.0  # fallback px mínimo desacople
# Separación vertical entre fibra de la pareja y el solape desacoplado (eje a eje).
# Es el hueco visible entre las dos barras paralelas del empalme (ver preview alzado).
_ELEV_EMP_PAIR_SEP_MM = 10.0
_ELEV_EMP_LAP_FRAC = 0.12  # lap total ≈ 12 % del ancho de la viga anfitrión
_ELEV_EMP_LAP_MIN_PX = 16.0
_ELEV_EMP_LAP_MAX_PX = 56.0
# Desplazamiento capas pares: k × media lap. k=2 → shift = solape total visual,
# para que la zona de traslapo de 2.ª/4.ª arranque al terminar la de 1.ª/3.ª.
# Mismo valor que ``EMPALME_LAYER_ALT_LAP_K`` en domain (modelado).
try:
    from armado_vigas.domain.constants import EMPALME_LAYER_ALT_LAP_K as _ELEV_EMP_LAYER_ALT_LAP_K
except Exception:
    _ELEV_EMP_LAYER_ALT_LAP_K = 2.0
_ELEV_EMP_LAYER_GAP = 3.5  # fallback mínimo px si no hay escala V
# Separación simbólica entre capas SUP en alzado (eje a eje).
_ELEV_BAR_LAYER_GAP_MM = 50.0
# Radio de filete (px) en esquinas de barras SUP (polilínea suave).
_ELEV_BAR_CORNER_R = 5.0
# Recubrimiento simbólico 1ª capa SUP respecto a la cara superior (mm).
_ELEV_BAR_COVER_SUP_MM = 25.0
# Estirón extremo vs viga/muro transversal no//: +(b/2 − clearance), misma regla de colocación.
_ELEV_BEAM_END_CLEARANCE_MM = 25.0
_ELEV_WALL_END_CLEARANCE_MM = 25.0
# Tolerancia U (mm) para asociar viga unida no// al extremo libre del Tn.
_ELEV_JOIN_END_TOL_MM = 350.0
# Tol. extra (mm) sobre ½ espesor para asociar muro no// solo al extremo cercano.
_ELEV_WALL_END_EXTRA_TOL_MM = 80.0

# Hormigón unificado (viga · columna · muro): fill + contorno propio del material.
# Contorno = tono hormigón (gris cálido), trazo fino tipo lápiz — no accent UI.
_ELEV_CONCRETE_FILL_ARGB = (52, 138, 142, 148)  # relleno hormigón
_ELEV_CONCRETE_EDGE_HEX = u"#9ca3a8"  # arista hormigón (gris neutro-cálido)
_ELEV_CONCRETE_EDGE_A = 175
_ELEV_CONCRETE_AXIS_HEX = u"#787f86"
_ELEV_CONCRETE_AXIS_A = 90
_ELEV_CONCRETE_STROKE = 0.5
_ELEV_CONCRETE_AXIS_STROKE = 0.4
_ELEV_CONCRETE_AXIS_DASH = [4.0, 3.0]
# Selección: resalte ligero (fill un poco más denso + arista más legible + velo suave).
_ELEV_CONCRETE_SEL_HEX = u"#b8c4cc"
_ELEV_CONCRETE_SEL_A = 220
_ELEV_CONCRETE_SEL_STROKE = 1.0
_ELEV_CONCRETE_SEL_FILL_HEX = u"#9aa3ab"
_ELEV_CONCRETE_SEL_FILL_A = 78
_ELEV_BEAM_SEL_WASH_HEX = u"#5bb8d4"
_ELEV_BEAM_SEL_WASH_A = 22
# Hit-target: sin anillo UI (el resalte va en silueta dibujada).
_ELEV_BEAM_SEL_RING_HEX = u"#000000"
_ELEV_BEAM_SEL_RING_A = 0
_ELEV_BEAM_SEL_RING_STROKE = 0.0
# Apoyo con suple SUP activo: velo en silueta (sin recuadro externo).
_ELEV_SUPLE_APOYO_WASH_HEX = u"#a78bfa"
_ELEV_SUPLE_APOYO_WASH_A = 22
_ELEV_SUPLE_APOYO_EDGE_HEX = u"#a78bfa"
_ELEV_SUPLE_APOYO_EDGE_A = 120
_ELEV_SUPLE_APOYO_STROKE = 0.8
# Apoyo seleccionado para configurar (más fuerte que solo ON).
_ELEV_SUPLE_APOYO_SEL_WASH_A = 48
_ELEV_SUPLE_APOYO_SEL_EDGE_A = 210
_ELEV_SUPLE_APOYO_SEL_STROKE = 1.2

# Vigas unidas (detectadas, no armadas): estilo distinto al lote principal.
_ELEV_JOIN_PAR_EDGE = u"#6b7c8a"
_ELEV_JOIN_PAR_FILL_A = 28
_ELEV_JOIN_NPAR_EDGE = u"#c4a574"
_ELEV_JOIN_NPAR_FILL_A = 40
_ELEV_JOIN_STROKE = 0.7
_ELEV_JOIN_STROKE_NPAR = 0.9
_ELEV_JOIN_LABEL_FONT = 8.0
# Losas seleccionadas (contexto de alzado; no definen extremos de viga).
_ELEV_FLOOR_EDGE = u"#8cb4c8"
_ELEV_FLOOR_FILL_A = 52
_ELEV_FLOOR_STROKE = 0.85
_ELEV_FLOOR_LABEL_FONT = 8.0

# Marcador de dirección LocationCurve 0→1 — meta UI, lenguaje del canvas
# (trazo fino + pill, no blanco sólido). Se dibuja bajo silueta de viga.
_ELEV_DIR_MARKER_LEN_PX = 56.0
_ELEV_DIR_MARKER_MIN_W_PX = 36.0
_ELEV_DIR_MARKER_STROKE = 1.1
_ELEV_DIR_MARKER_HEX = u"#8ba3b5"
_ELEV_DIR_MARKER_A = 200
_ELEV_DIR_MARKER_HEAD_L = 7.0
_ELEV_DIR_MARKER_HEAD_H = 4.0
_ELEV_DIR_MARKER_TICK_H = 4.5
_ELEV_DIR_MARKER_FONT_PX = 9.0
_ELEV_DIR_MARKER_BADGE_H = 15.0
_ELEV_DIR_MARKER_BG = u"#0b1624"
_ELEV_DIR_BELOW_GAP_PX = 7.0  # aire entre fondo de viga y trazo de dirección
# Alias legacy (mismos tokens — evita divergencias en helpers).
_ELEV_STROKE_CHORD = _ELEV_CONCRETE_STROKE
_ELEV_SUPPORT_STROKE = _ELEV_CONCRETE_STROKE
_ELEV_BEAM_CHORD = brush_hex(_ELEV_CONCRETE_EDGE_HEX, _ELEV_CONCRETE_EDGE_A)


def _elev_concrete_fill_brush(selected=False):
    # ARGB original (52, 138, 142, 148) → brush cacheado Freezado
    if selected:
        return brush_hex(_ELEV_CONCRETE_SEL_FILL_HEX, _ELEV_CONCRETE_SEL_FILL_A)
    return brush_hex(u"#8a8e94", 52)


def _elev_concrete_edge_brush(selected=False):
    if selected:
        return brush_hex(_ELEV_CONCRETE_SEL_HEX, _ELEV_CONCRETE_SEL_A)
    return brush_hex(_ELEV_CONCRETE_EDGE_HEX, _ELEV_CONCRETE_EDGE_A)


def _elev_concrete_axis_brush():
    return brush_hex(_ELEV_CONCRETE_AXIS_HEX, _ELEV_CONCRETE_AXIS_A)


def _elev_concrete_stroke_w(selected=False):
    return _ELEV_CONCRETE_SEL_STROKE if selected else _ELEV_CONCRETE_STROKE


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
        self._btn_section_zoom = win.FindName(u"BtnSectionZoom") if win else None
        self._pnl_section_ctrls = win.FindName(u"PnlSectionCtrls") if win else None
        self._txt_tramo = win.FindName(u"TxtTramoSummary") if win else None
        self._txt_apoyos = win.FindName(u"TxtApoyosSummary") if win else None
        self._txt_sub = win.FindName(u"TxtSubtitle") if win else None
        self._txt_sel = win.FindName(u"TxtSelectionInfo") if win else None

        self.selected_tramo_sup_id = None
        self.selected_tramo_inf_id = None
        self.selected_tramo_ids_sup = set()
        self.selected_tramo_ids_inf = set()
        self.selected_beam_idx = -1
        self.selected_beam_indices = set()
        self.selected_stirrup_zone = None
        self.rail_card = u"sup"
        self.card_on_sup = True
        self.card_on_inf = True
        self.card_on_lat = True
        self.card_on_conf = True
        self.conf_face = u"sup"
        self._layout_meta = {"contentWidthPx": 640.0, "needsScroll": False}
        self._last_beams = []
        self._last_session = None
        self._drawing = False
        self._pending_redraw = False
        # Caché alzado: si el estado visual no cambió, solo se refresca el rail.
        self._elev_cache_fp = None
        self._paint_frame = None
        self._zoom_chrome_row = None
        self._elev_hdr = None
        self._txt_elev_hint = None
        self._layout_cache_key = None
        self._layout_cache_val = None
        self._rail_paint_gen = 0
        self._rail_cache_fp = None
        # Memo de sub-firmas visuales válido solo durante un paint (elev_fp + rail_fp).
        self._fp_memo = None

        # Zoom de vista del canvas (LayoutTransform) — independiente del zoom-extents del layout.
        self._view_zoom = float(getattr(lay, "VIEW_ZOOM_DEFAULT", 1.0) or 1.0)
        self._zoom_root = None
        self._txt_zoom = None
        self._nav_wired = False
        # Paneo del ScrCanvas (MMB drag / rueda → horizontal).
        self._pan_active = False
        self._pan_last = None
        self._pan_origin_cursor = None
        self._wire_canvas_nav()
        # Precalentar controles reusables del header (evita ~25 ms en el 1er elev).
        try:
            self._build_elev_header()
        except Exception:
            pass

        # CONF dibujo libre: canvas = sección del rail; alzado = selección vigas
        self.conf_draw_mode = u"draw"  # draw | peri | erase
        self._conf_pending = None
        self._conf_hover = None
        self._conf_origin = None  # (x, y) 1.er clic — origen marquee
        self._conf_cursor = None  # (x, y) punta del cursor en preview
        self._conf_geom = {}
        self._conf_geom_by_cnv = {}
        self._conf_canvas = None
        self._conf_beam = None
        self._conf_status_tb = None
        self._conf_wired_cnv_ids = set()
        self._conf_raf_pending = False
        self._conf_zoom_win = None
        self._conf_zoom_canvas = None
        self._conf_zoom_status = None
        self._wire_section_zoom_btn()

    def invalidate_elev_cache(self):
        """Invalida caché de alzado (topología / geometría cambiada fuera de paint)."""
        self._elev_cache_fp = None
        self._paint_frame = None
        self._layout_cache_key = None
        self._layout_cache_val = None
        self._rail_cache_fp = None
        self._fp_memo = None
        self._rail_paint_gen = int(self._rail_paint_gen or 0) + 1

    def _layout_cache_fingerprint(self, beams, viewport_w, viewport_h, session):
        parts = [
            u"vw={0:.1f}".format(float(viewport_w or 0)),
            u"vh={0:.1f}".format(float(viewport_h or 0)),
        ]
        try:
            parts.append(
                u"ap={0}".format(len(getattr(session, u"apoyos", None) or []))
            )
            jf = getattr(session, u"joined_framing", None) or {}
            c = jf.get(u"counts") or {}
            parts.append(
                u"jf={0},{1}".format(c.get(u"all") or 0, c.get(u"not_parallel") or 0)
            )
        except Exception:
            pass
        for i, b in enumerate(beams or []):
            parts.append(
                u"{0}:{1},{2},{3},{4}".format(
                    self._fp_u(b.get(u"id")),
                    self._fp_u(b.get(u"uStart")),
                    self._fp_u(b.get(u"uEnd")),
                    self._fp_u(b.get(u"vMin")),
                    self._fp_u(b.get(u"vMax")),
                )
            )
        return u"|".join(parts)

    def _compute_layout_cached(self, beams, viewport_w, viewport_h, session):
        key = self._layout_cache_fingerprint(beams, viewport_w, viewport_h, session)
        if key == self._layout_cache_key and self._layout_cache_val is not None:
            return self._layout_cache_val
        result = lay.compute_layout(
            beams,
            viewport_w,
            apoyos=getattr(session, "apoyos", None) or [],
            use_model_positions=True,
            joined=getattr(session, "joined_framing", None),
            viewport_h=viewport_h,
        )
        self._layout_cache_key = key
        self._layout_cache_val = result
        return result

    def _schedule_section_rail(self, beams):
        """Construye rail/sección en Background: alzado se ve primero (path crítico)."""
        self._rail_paint_gen = int(self._rail_paint_gen or 0) + 1
        gen = self._rail_paint_gen
        beams_snap = beams
        session = self._last_session
        try:
            tramos_sup = list(getattr(session, "tramos_sup", None) or []) if session else []
            tramos_inf = list(getattr(session, "tramos_inf", None) or []) if session else []
        except Exception:
            tramos_sup, tramos_inf = [], []
        try:
            rail_fp = self._rail_visual_fingerprint(
                beams_snap, session, tramos_sup, tramos_inf
            )
        except Exception:
            rail_fp = None

        def _run():
            if gen != self._rail_paint_gen:
                return
            # Skip si el rail ya refleja el estado (evita 60–100 ms por rebuild idéntico).
            if (
                rail_fp is not None
                and rail_fp == self._rail_cache_fp
                and self._pnl_section_ctrls is not None
            ):
                try:
                    if self._pnl_section_ctrls.Children.Count > 0:
                        return
                except Exception:
                    pass
            try:
                self._draw_section_rail(beams_snap)
                if rail_fp is not None:
                    self._rail_cache_fp = rail_fp
            except Exception:
                pass

        try:
            from System import Action
            from System.Windows.Threading import DispatcherPriority

            disp = None
            if self._win is not None:
                disp = self._win.Dispatcher
            if disp is not None:
                try:
                    disp.BeginInvoke(Action(_run), DispatcherPriority.Background)
                    return
                except Exception:
                    pass
        except Exception:
            pass
        _run()

    def _elev_emp_for_layer_px(self, beams, tramo, session, diam_mm, ext_s, ext_e):
        """Emp. muro // por capa (estirón ext_* ya calculado; no reescanea col/viga)."""
        emp_s = 0.0
        emp_e = 0.0
        try:
            e_s, e_e = self._elev_tramo_wall_parallel_emp_px(
                beams, tramo, session, diam_mm=diam_mm
            )
            if float(ext_s or 0.0) <= 0.5:
                emp_s = float(e_s or 0.0)
            if float(ext_e or 0.0) <= 0.5:
                emp_e = float(e_e or 0.0)
        except Exception:
            emp_s = emp_e = 0.0
        return emp_s, emp_e

    def _is_ctrl_click(self, args):
        try:
            mods = Keyboard.Modifiers
        except Exception:
            try:
                mods = args.KeyboardDevice.Modifiers
            except Exception:
                return False
        return (mods & ModifierKeys.Control) == ModifierKeys.Control

    def _selected_tramo_ids(self, face):
        """Conjunto de Tn seleccionados en la cara (sincroniza con primario)."""
        is_sup = face == u"sup"
        ids_attr = u"selected_tramo_ids_sup" if is_sup else u"selected_tramo_ids_inf"
        primary_attr = u"selected_tramo_sup_id" if is_sup else u"selected_tramo_inf_id"
        ids = set(getattr(self, ids_attr, None) or set())
        primary = getattr(self, primary_attr, None)
        if not ids and primary is not None:
            ids = {primary}
            setattr(self, ids_attr, ids)
        return ids

    def _is_tramo_selected(self, face, tramo_id):
        return tramo_id in self._selected_tramo_ids(face)

    def _wire_tramo_select_click(self, element, face, tramo_id):
        """Clic / Ctrl+clic en pill o banda Tn → selección (solo pestaña INF/CONF)."""

        def _select(sender, args, tramo_face=face, tid=tramo_id):
            if not self._tramo_beam_selection_allowed(tramo_face):
                return
            try:
                from armado_vigas.ui import rail_cards as _rc

                ctrl = False
                try:
                    ctrl = bool(self._is_ctrl_click(args))
                except Exception:
                    pass
                _rc._select_tramo_multi(self, tramo_face, tid, ctrl=ctrl)
                try:
                    args.Handled = True
                except Exception:
                    pass
                return
            except Exception:
                pass
            # Fallback (misma política de pestañas).
            if not self._tramo_beam_selection_allowed(tramo_face):
                return
            if tramo_face == u"sup":
                self.selected_tramo_sup_id = tid
                self.selected_tramo_ids_sup = {tid}
            else:
                self.selected_tramo_inf_id = tid
                self.selected_tramo_ids_inf = {tid}
            self._cb.get("on_select_tramo", lambda _t, _f: None)(tid, tramo_face)
            self._cb.get("on_redraw", lambda: None)()

        handler = MouseButtonEventHandler(_select)
        try:
            from System.Windows import UIElement

            element.AddHandler(UIElement.MouseLeftButtonUpEvent, handler, True)
        except Exception:
            try:
                element.MouseLeftButtonUp += handler
            except Exception:
                pass

    def _is_beam_selected(self, idx):
        try:
            i = int(idx)
        except Exception:
            return False
        return i in self.selected_beam_indices

    def _active_rail_card(self):
        try:
            return getattr(self, u"rail_card", None) or u"sup"
        except Exception:
            return u"sup"

    def _tramo_beam_selection_allowed(self, face=None):
        """
        Selección en alzado según pestaña activa.

        - SUP: solo tramos de cara superior (bandas Tn SUP). Sin hit de silueta de viga.
        - INF: tramos/face inferior + hit de viga.
        - CONF: tramos sup o inf; no abandona la pestaña CONF.
        - LAT: sin selección.
        """
        card = self._active_rail_card()
        if card == u"lat":
            return False
        if card == u"sup":
            # Bandas Tn SUP: face=sup. face=None = hit de viga → no en SUP.
            return face == u"sup"
        if card == u"inf":
            if face is not None and face != u"inf":
                return False
            return True
        if card == u"conf":
            return True
        return False

    def _elev_beam_selection_ui_enabled(self):
        """Resalte y hit de vigas en silueta: INF y CONF (no SUP: solo bandas Tn)."""
        card = self._active_rail_card()
        return card in (u"inf", u"conf")

    def _is_beam_selected_for_elev(self, idx):
        """Selección visual de viga en alzado/labels (solo INF/CONF)."""
        if not self._elev_beam_selection_ui_enabled():
            return False
        return self._is_beam_selected(idx)

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
        if not self._tramo_beam_selection_allowed():
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
        # Zona Ext/Cent/Uni solo en pestaña CONF: no saltar a confinamiento
        # al elegir tramo/viga en alzado desde INF u otra pestaña.
        on_conf = self._active_rail_card() == u"conf"
        if on_conf and role is not None:
            self.selected_stirrup_zone = {u"idx": idx, u"role": role}
            self._cb.get("on_select_stirrup_zone", lambda _i, _r: None)(idx, role)
        elif on_conf and update_zone:
            zone_role = self._default_stirrup_role(beams[idx])
            self.selected_stirrup_zone = {u"idx": idx, u"role": zone_role}
        elif on_conf:
            # Mantener rol de zona; actualizar idx a la primaria multi-sel.
            sz = self.selected_stirrup_zone or {}
            prev_role = sz.get(u"role") or self._default_stirrup_role(beams[idx])
            self.selected_stirrup_zone = {u"idx": idx, u"role": prev_role}
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

    @staticmethod
    def _fp_u(val):
        try:
            return unicode(val)
        except NameError:
            return str(val)
        except Exception:
            return u"?"

    @staticmethod
    def _fp_num(val, nd=2):
        """Número estable para firma (evita float noise / unicode lento)."""
        if val is None:
            return None
        try:
            return round(float(val), int(nd))
        except Exception:
            return val

    def _fp_memo_get(self, key, factory):
        """Cachea sub-firmas solo si hay memo de paint activo."""
        m = self._fp_memo
        if m is None:
            return factory()
        hit = m.get(key)
        if hit is not None or key in m:
            return hit
        val = factory()
        m[key] = val
        return val

    def _beam_elev_fp_key(self, beam):
        """Tupla compacta de campos de viga que afectan alzado/rail."""
        if not beam:
            return ()
        bid = beam.get(u"id")
        return self._fp_memo_get(
            (u"b", bid),
            lambda: self._beam_elev_fp_key_raw(beam),
        )

    def _beam_elev_fp_key_raw(self, beam):
        g = beam.get
        return (
            g(u"id"),
            g(u"type"),
            self._fp_num(g(u"len"), 3),
            self._fp_num(g(u"uStart"), 4),
            self._fp_num(g(u"uEnd"), 4),
            self._fp_num(g(u"vMin"), 4),
            self._fp_num(g(u"vMax"), 4),
            self._fp_num(g(u"sectionDepthMm"), 1),
            self._fp_num(g(u"heightMm"), 1),
            self._fp_num(g(u"widthMm"), 1),
            g(u"nCapasSup"),
            g(u"nCapasInf"),
            g(u"nSup"),
            g(u"nInf"),
            g(u"diamSup"),
            g(u"diamInf"),
            g(u"nSup2"),
            g(u"nInf2"),
            g(u"diamSup2"),
            g(u"diamInf2"),
            g(u"nSup3"),
            g(u"nInf3"),
            g(u"diamSup3"),
            g(u"diamInf3"),
            g(u"supleInfEnabled"),
            g(u"nSupleInf"),
            g(u"diamSupleInf"),
            g(u"supleSupEnabled"),
            g(u"nSupleSup"),
            g(u"diamSupleSup"),
            g(u"supleSupStartEnabled"),
            g(u"supleSupEndEnabled"),
            g(u"estDiamExt"),
            g(u"estDiamCent"),
            g(u"estSpExt"),
            g(u"estSpCent"),
            g(u"estZonasMode"),
            g(u"estConfin"),
            conf_draft_signature(get_conf_draft(beam) if beam else None),
            g(u"colStart"),
            g(u"colEnd"),
            bool(g(u"axisReversed")),
        )

    def _tramo_elev_fp_key(self, tramo):
        if not tramo:
            return ()
        tid = tramo.get(u"id")
        return self._fp_memo_get(
            (u"t", tid),
            lambda: self._tramo_elev_fp_key_raw(tramo),
        )

    def _tramo_elev_fp_key_raw(self, tramo):
        try:
            idxs = tuple(tramo.get(u"beamIndices") or ())
        except Exception:
            idxs = ()
        return (
            tramo.get(u"id"),
            idxs,
            tramo.get(u"edgeStart"),
            tramo.get(u"edgeEnd"),
            bool(tramo.get(u"fromEmpalme")),
            tramo.get(u"empalme"),
            tramo.get(u"accent"),
        )

    def _armado_store_fp_key(self, session, face):
        try:
            store = getattr(session, u"tramo_armado", None) or {}
            face_map = store.get(face) if isinstance(store, dict) else None
            if not face_map:
                return ()
            items = []
            for k in sorted(face_map.keys(), key=lambda x: self._fp_u(x)):
                cfg = face_map.get(k)
                if isinstance(cfg, dict):
                    items.append(
                        (
                            self._fp_u(k),
                            tuple(
                                (self._fp_u(ck), cfg.get(ck))
                                for ck in sorted(cfg.keys(), key=lambda x: self._fp_u(x))
                            ),
                        )
                    )
                else:
                    items.append((self._fp_u(k), cfg))
            return tuple(items)
        except Exception:
            return ()

    def _session_ui_fp_key(self, session):
        if session is None:
            return ()
        sid = id(session)
        return self._fp_memo_get(
            (u"s", sid),
            lambda: self._session_ui_fp_key_raw(session),
        )

    def _session_ui_fp_key_raw(self, session):
        try:
            emp_s = tuple(sorted(getattr(session, u"empalme_beam_ids_sup", None) or set()))
            emp_i = tuple(sorted(getattr(session, u"empalme_beam_ids_inf", None) or set()))
            ssup = tuple(sorted(getattr(session, u"suple_sup_apoyo_ids", None) or set()))
            try:
                ssel = unicode(
                    getattr(session, u"selected_suple_apoyo_id", None) or u""
                )
            except Exception:
                ssel = u""
            scfg = getattr(session, u"suple_sup_cfg_by_apoyo", None) or {}
            try:
                scfg_sig = tuple(
                    sorted(
                        (unicode(k), int((v or {}).get(u"n") or 0), int((v or {}).get(u"diam") or 0))
                        for k, v in scfg.items()
                    )
                )
            except Exception:
                scfg_sig = ()
        except Exception:
            emp_s = emp_i = ssup = ()
            ssel = u""
            scfg_sig = ()
        return (
            emp_s,
            emp_i,
            ssup,
            ssel,
            scfg_sig,
            getattr(session, u"concreteGrade", None),
            getattr(session, u"nLaterales", None),
            getattr(session, u"diamLaterales", None),
            bool(getattr(session, u"lateralesEnabled", True)),
            getattr(session, u"barEndStartSup", None),
            getattr(session, u"barEndEndSup", None),
            getattr(session, u"barEndStartInf", None),
            getattr(session, u"barEndEndInf", None),
            self._armado_store_fp_key(session, u"sup"),
            self._armado_store_fp_key(session, u"inf"),
        )

    def _selection_fp_key(self):
        return self._fp_memo_get(u"sel", self._selection_fp_key_raw)

    def _selection_fp_key_raw(self):
        sz = self.selected_stirrup_zone
        try:
            if isinstance(sz, dict):
                sz_k = (sz.get(u"idx"), sz.get(u"role"))
            else:
                sz_k = sz
        except Exception:
            sz_k = None
        return (
            self.rail_card or u"",
            getattr(self, u"conf_face", u"sup") or u"sup",
            bool(self.card_on_sup),
            bool(self.card_on_inf),
            bool(self.card_on_lat),
            bool(self.card_on_conf),
            tuple(sorted(self.selected_tramo_ids_sup or [])),
            tuple(sorted(self.selected_tramo_ids_inf or [])),
            self.selected_tramo_sup_id,
            self.selected_tramo_inf_id,
            self.selected_beam_idx,
            tuple(sorted(self.selected_beam_indices or [])),
            sz_k,
        )

    def _elev_visual_fingerprint(
        self, beams, layouts, tramos_sup, tramos_inf, session, content_w, layout_result,
    ):
        """Firma barata (tupla) del estado visual del alzado."""
        try:
            eh = float(
                (layout_result or {}).get(u"elevHeightPx") or lay.ELEVATION_HEIGHT_PX
            )
        except Exception:
            eh = float(lay.ELEVATION_HEIGHT_PX)
        layout_keys = []
        for L in layouts or []:
            try:
                layout_keys.append(
                    (self._fp_num(L.get(u"leftPx"), 1), self._fp_num(L.get(u"widthPx"), 1))
                )
            except Exception:
                layout_keys.append((None, None))
        return (
            self._fp_num(content_w, 1),
            self._fp_num(eh, 1),
            bool((layout_result or {}).get(u"modelPositions")),
            self._selection_fp_key(),
            self._session_ui_fp_key(session),
            tuple(self._beam_elev_fp_key(b) for b in (beams or [])),
            tuple(self._tramo_elev_fp_key(t) for t in (tramos_sup or [])),
            tuple(self._tramo_elev_fp_key(t) for t in (tramos_inf or [])),
            tuple(layout_keys),
        )

    def _rail_visual_fingerprint(self, beams, session, tramos_sup, tramos_inf):
        """Firma del rail + preview sección (sin geometry/layout del alzado)."""
        return (
            self._selection_fp_key(),
            self._session_ui_fp_key(session),
            tuple(self._beam_elev_fp_key(b) for b in (beams or [])),
            tuple(self._tramo_elev_fp_key(t) for t in (tramos_sup or [])),
            tuple(self._tramo_elev_fp_key(t) for t in (tramos_inf or [])),
        )

    def _host_has_paint_frame(self):
        try:
            if self._host is None or self._paint_frame is None:
                return False
            if self._host.Children.Count < 1:
                return False
            # El frame cacheado debe seguir en el árbol.
            for child in self._host.Children:
                if child is self._paint_frame:
                    return True
        except Exception:
            return False
        return False

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
        self._fp_memo = {}
        try:
            beams = sort_beams(list(session.domain_beams or []))
            apoyos_loaded = bool(getattr(session, "apoyos_loaded", False))

            if not beams:
                self.invalidate_elev_cache()
                self._host.Children.Clear()
                self._zoom_root = None
                self._show_empty(apoyos_loaded)
                return

            for beam in beams:
                ensure_beam_layers(beam)
                ensure_beam_confinement(beam)
                ensure_beam_suple_inferior(beam)
                ensure_beam_suple_superior(beam)

            self._last_session = session
            viewport_w = self._viewport_width()
            viewport_h = max(1.0, self._viewport_height())

            layout_result = self._compute_layout_cached(
                beams, viewport_w, viewport_h, session
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

            # Mantener sets multi-sel alineados con primarios / tramos vigentes.
            try:
                rail_cards._normalize_tramo_multi(self, u"sup", tramos_sup)
                rail_cards._normalize_tramo_multi(self, u"inf", tramos_inf)
            except Exception:
                pass

            self._normalize_beam_selection(beams)
            if self.selected_stirrup_zone and self.selected_stirrup_zone.get("idx", -1) >= len(beams):
                self.selected_stirrup_zone = None
            if not self.selected_stirrup_zone and beams:
                self.selected_stirrup_zone = {"idx": 0, "role": "cent"}

            # CONF: alzado se mantiene (selección de vigas/tramos). Estribos/trabas
            # se dibujan en el canvas de sección del rail lateral.

            content_w = float(layout_result["contentWidthPx"])
            elev_fp = self._elev_visual_fingerprint(
                beams, layouts, tramos_sup, tramos_inf, session, content_w, layout_result,
            )

            # Fast path: estado visual del alzado intacto.
            # elev_fp ya incluye selección, ø/n, pestaña y armado.
            # Aun así se reprograma el rail (skip interno si rail_fp no cambió):
            # flags de suple/card no deben dejar el panel lateral muerto.
            if (
                elev_fp == self._elev_cache_fp
                and self._host_has_paint_frame()
            ):
                self._last_beams = beams
                # Headers baratos; pueden reflejar needsScroll tras resize.
                self._update_headers(
                    beams, tramos_sup, tramos_inf, apoyos_loaded, layout_result
                )
                self._schedule_section_rail(beams)
                return

            self._host.Children.Clear()
            self._zoom_root = None
            self._paint_frame = None

            self._update_headers(beams, tramos_sup, tramos_inf, apoyos_loaded, layout_result)

            viewport_w = max(1.0, float(viewport_w or self._viewport_width()))
            viewport_h = max(1.0, float(viewport_h or self._viewport_height()))

            # Con zoom-extents, el contenido llena el ancho útil; el marco usa el mismo tamaño.
            frame_w = max(content_w, viewport_w) if not layout_result.get("zoomExtents") else max(content_w, viewport_w * 0.99)
            frame_h = max(120.0, viewport_h)

            root = Grid()
            root.Width = content_w
            root.Background = brush_hex(u"#0a1620", 0)
            root.HorizontalAlignment = HorizontalAlignment.Center
            root.VerticalAlignment = VerticalAlignment.Center
            try:
                rd0 = RowDefinition()
                rd0.Height = GridLength(1.0, GridUnitType.Auto)
                rd1 = RowDefinition()
                rd1.Height = GridLength(1.0, GridUnitType.Auto)
                rd2 = RowDefinition()
                rd2.Height = GridLength(1.0, GridUnitType.Auto)
                root.RowDefinitions.Add(rd0)
                root.RowDefinitions.Add(rd1)
                root.RowDefinitions.Add(rd2)
            except Exception:
                pass

            stack = Border()
            stack.Width = content_w
            stack.BorderBrush = brush_hex(u"#21465C")
            stack.BorderThickness = Thickness(1)
            stack.Background = brush_hex(u"#071018", 0)
            try:
                # Evita que el contorno 1 px del stack recorte el alzado interior.
                stack.ClipToBounds = False
            except Exception:
                pass
            stack.Child = self._build_elev_stage_option_d(
                beams, layouts, tramos_sup, tramos_inf, session, apoyos_loaded, content_w,
            )
            labels = self._build_labels(beams, layouts, session, apoyos_loaded, content_w)
            axis = self._build_axis_hint(apoyos_loaded, content_w)
            try:
                Grid.SetRow(stack, 0)
                Grid.SetRow(labels, 1)
                Grid.SetRow(axis, 2)
            except Exception:
                pass
            root.Children.Add(stack)
            root.Children.Add(labels)
            root.Children.Add(axis)

            # Marco: centra el dibujo; con zoom-extents content ≈ viewport.
            frame = Grid()
            try:
                frame.MinWidth = max(content_w, frame_w)
                frame.MinHeight = max(120.0, frame_h)
            except Exception:
                frame.Width = max(content_w, frame_w)
            frame.Background = brush_hex(u"#0a1620", 0)
            root.HorizontalAlignment = HorizontalAlignment.Center
            root.VerticalAlignment = VerticalAlignment.Center
            frame.Children.Add(root)
            try:
                self._host.HorizontalAlignment = HorizontalAlignment.Stretch
                self._host.VerticalAlignment = VerticalAlignment.Stretch
            except Exception:
                pass
            self._zoom_root = frame
            self._paint_frame = frame
            self._elev_cache_fp = elev_fp
            self._host.Children.Add(frame)
            self._apply_view_zoom(preserve_scroll=False)
            self._last_beams = beams
            # Rail + preview sección en Background: path crítico = alzado visible.
            self._schedule_section_rail(beams)
        finally:
            self._fp_memo = None

    def _wire_canvas_nav(self):
        """Zoom (Ctrl+rueda), paneo horizontal (rueda / MMB drag) — wire una vez."""
        if self._nav_wired or self._scr is None:
            return
        scr = self._scr
        try:
            scr.PreviewMouseWheel += MouseWheelEventHandler(self._on_canvas_mouse_wheel)
        except Exception:
            try:
                scr.MouseWheel += MouseWheelEventHandler(self._on_canvas_mouse_wheel)
            except Exception:
                pass
        try:
            scr.PreviewMouseDown += MouseButtonEventHandler(self._on_canvas_pan_down)
            scr.PreviewMouseMove += MouseEventHandler(self._on_canvas_pan_move)
            scr.PreviewMouseUp += MouseButtonEventHandler(self._on_canvas_pan_up)
            try:
                scr.LostMouseCapture += MouseEventHandler(self._on_canvas_pan_lost)
            except Exception:
                pass
        except Exception:
            try:
                scr.MouseDown += MouseButtonEventHandler(self._on_canvas_pan_down)
                scr.MouseMove += MouseEventHandler(self._on_canvas_pan_move)
                scr.MouseUp += MouseButtonEventHandler(self._on_canvas_pan_up)
            except Exception:
                pass
        self._nav_wired = True

    def _wire_canvas_zoom(self):
        """Alias legacy → navegación completa."""
        self._wire_canvas_nav()

    def _scroll_h_by(self, delta_px):
        """Desplaza horizontalmente el ScrCanvas (delta positivo = contenido a la izquierda)."""
        scr = self._scr
        if scr is None:
            return
        try:
            off = float(scr.HorizontalOffset or 0.0) + float(delta_px)
            if off < 0:
                off = 0.0
            try:
                extent = float(scr.ExtentWidth or 0.0)
                view = float(scr.ViewportWidth or 0.0)
                max_off = max(0.0, extent - view)
                if off > max_off:
                    off = max_off
            except Exception:
                pass
            scr.ScrollToHorizontalOffset(off)
        except Exception:
            pass

    def _scroll_v_by(self, delta_px):
        scr = self._scr
        if scr is None:
            return
        try:
            off = float(scr.VerticalOffset or 0.0) + float(delta_px)
            if off < 0:
                off = 0.0
            try:
                extent = float(scr.ExtentHeight or 0.0)
                view = float(scr.ViewportHeight or 0.0)
                max_off = max(0.0, extent - view)
                if off > max_off:
                    off = max_off
            except Exception:
                pass
            scr.ScrollToVerticalOffset(off)
        except Exception:
            pass

    def _is_middle_button(self, args):
        try:
            return args.ChangedButton == MouseButton.Middle
        except Exception:
            pass
        try:
            return bool(args.MiddleButton) and str(args.MiddleButton).endswith(u"Pressed")
        except Exception:
            return False

    def _on_canvas_pan_down(self, sender, args):
        """Inicia paneo con botón medio (mano)."""
        if not self._is_middle_button(args):
            return
        scr = self._scr
        if scr is None:
            return
        try:
            self._pan_active = True
            self._pan_last = args.GetPosition(scr)
            try:
                self._pan_origin_cursor = scr.Cursor
            except Exception:
                self._pan_origin_cursor = None
            try:
                scr.Cursor = Cursors.SizeAll
            except Exception:
                try:
                    scr.Cursor = Cursors.Hand
                except Exception:
                    pass
            try:
                scr.CaptureMouse()
            except Exception:
                pass
            try:
                args.Handled = True
            except Exception:
                pass
        except Exception:
            self._pan_active = False
            self._pan_last = None

    def _on_canvas_pan_move(self, sender, args):
        if not self._pan_active or self._scr is None or self._pan_last is None:
            return
        try:
            try:
                from System.Windows.Input import MouseButtonState

                if Mouse.MiddleButton != MouseButtonState.Pressed:
                    self._end_canvas_pan()
                    return
            except Exception:
                pass
            pos = args.GetPosition(self._scr)
            dx = float(pos.X) - float(self._pan_last.X)
            dy = float(pos.Y) - float(self._pan_last.Y)
            # Arrastrar a la derecha: el contenido se mueve con el cursor → offset −dx
            if abs(dx) > 0.01:
                self._scroll_h_by(-dx)
            if abs(dy) > 0.01:
                self._scroll_v_by(-dy)
            self._pan_last = pos
            try:
                args.Handled = True
            except Exception:
                pass
        except Exception:
            pass

    def _on_canvas_pan_up(self, sender, args):
        if not self._pan_active:
            return
        self._end_canvas_pan()
        try:
            args.Handled = True
        except Exception:
            pass

    def _on_canvas_pan_lost(self, sender, args):
        self._end_canvas_pan()

    def _end_canvas_pan(self):
        scr = self._scr
        self._pan_active = False
        self._pan_last = None
        if scr is None:
            return
        try:
            if scr.IsMouseCaptured:
                scr.ReleaseMouseCapture()
        except Exception:
            pass
        try:
            scr.Cursor = self._pan_origin_cursor if self._pan_origin_cursor is not None else Cursors.Arrow
        except Exception:
            try:
                scr.Cursor = Cursors.Arrow
            except Exception:
                pass
        self._pan_origin_cursor = None

    def _zoom_limits(self):
        zmin = float(getattr(lay, "VIEW_ZOOM_MIN", 0.25) or 0.25)
        zmax = float(getattr(lay, "VIEW_ZOOM_MAX", 4.0) or 4.0)
        return zmin, zmax

    def _clamp_view_zoom(self, z):
        zmin, zmax = self._zoom_limits()
        try:
            z = float(z)
        except Exception:
            z = 1.0
        if z < zmin:
            return zmin
        if z > zmax:
            return zmax
        return z

    def _view_zoom_label_text(self):
        try:
            return u"{0:.0f}%".format(self._view_zoom * 100.0)
        except Exception:
            return u"100%"

    def _sync_zoom_label(self):
        try:
            if self._txt_zoom is not None:
                self._txt_zoom.Text = self._view_zoom_label_text()
        except Exception:
            pass

    def _apply_view_zoom(self, preserve_scroll=True, pivot_viewport=None):
        """Aplica ScaleTransform (LayoutTransform) al contenido del ScrCanvas."""
        z = self._clamp_view_zoom(self._view_zoom)
        self._view_zoom = z
        root = self._zoom_root
        if root is None:
            self._sync_zoom_label()
            return

        old_z = 1.0
        try:
            lt = root.LayoutTransform
            if lt is not None and hasattr(lt, "ScaleX"):
                old_z = float(lt.ScaleX) or 1.0
        except Exception:
            old_z = 1.0

        scr = self._scr
        old_h = 0.0
        old_v = 0.0
        pivot_x = None
        pivot_y = None
        if preserve_scroll and scr is not None:
            try:
                old_h = float(scr.HorizontalOffset or 0.0)
                old_v = float(scr.VerticalOffset or 0.0)
            except Exception:
                pass
            if pivot_viewport is not None:
                try:
                    pivot_x = float(pivot_viewport.X)
                    pivot_y = float(pivot_viewport.Y)
                except Exception:
                    pivot_x = pivot_y = None
            if pivot_x is None:
                try:
                    pivot_x = float(scr.ViewportWidth or 0.0) * 0.5
                    pivot_y = float(scr.ViewportHeight or 0.0) * 0.5
                except Exception:
                    pivot_x = pivot_y = 0.0

        try:
            st = ScaleTransform(z, z)
            root.LayoutTransform = st
        except Exception:
            try:
                root.RenderTransform = ScaleTransform(z, z)
                root.RenderTransformOrigin = Point(0, 0)
            except Exception:
                pass

        if preserve_scroll and scr is not None and old_z > 1e-6 and pivot_x is not None:
            try:
                ratio = z / old_z
                new_h = (old_h + pivot_x) * ratio - pivot_x
                new_v = (old_v + pivot_y) * ratio - pivot_y
                if new_h < 0:
                    new_h = 0.0
                if new_v < 0:
                    new_v = 0.0
                scr.ScrollToHorizontalOffset(new_h)
                scr.ScrollToVerticalOffset(new_v)
            except Exception:
                pass

        self._sync_zoom_label()

    def set_view_zoom(self, zoom, preserve_scroll=True, pivot_viewport=None, announce=True):
        """Zoom de vista (1.0 = escala del layout actual)."""
        z = self._clamp_view_zoom(zoom)
        self._view_zoom = z
        self._apply_view_zoom(preserve_scroll=preserve_scroll, pivot_viewport=pivot_viewport)
        if announce:
            try:
                self._cb.get("on_status", lambda _m: None)(
                    u"Zoom alzado · {0}".format(self._view_zoom_label_text())
                )
            except Exception:
                pass

    def zoom_view_by(self, factor, pivot_viewport=None):
        step = float(factor or 1.0)
        if step <= 0:
            return
        self.set_view_zoom(self._view_zoom * step, preserve_scroll=True, pivot_viewport=pivot_viewport)

    def zoom_view_in(self, pivot_viewport=None):
        step = float(getattr(lay, "VIEW_ZOOM_STEP", 1.15) or 1.15)
        self.zoom_view_by(step, pivot_viewport=pivot_viewport)

    def zoom_view_out(self, pivot_viewport=None):
        step = float(getattr(lay, "VIEW_ZOOM_STEP", 1.15) or 1.15)
        self.zoom_view_by(1.0 / step, pivot_viewport=pivot_viewport)

    def zoom_view_reset(self):
        self.set_view_zoom(
            float(getattr(lay, "VIEW_ZOOM_DEFAULT", 1.0) or 1.0),
            preserve_scroll=False,
            announce=True,
        )
        try:
            if self._scr is not None:
                self._scr.ScrollToHorizontalOffset(0.0)
                self._scr.ScrollToVerticalOffset(0.0)
        except Exception:
            pass

    def _on_canvas_mouse_wheel(self, sender, args):
        """Rueda: horizontal · Shift+rueda: vertical · Ctrl+rueda: zoom."""
        try:
            mods = Keyboard.Modifiers
            ctrl = (mods & ModifierKeys.Control) == ModifierKeys.Control
            shift = (mods & ModifierKeys.Shift) == ModifierKeys.Shift
        except Exception:
            ctrl = False
            shift = False

        try:
            delta = int(args.Delta)
        except Exception:
            delta = 0
        if delta == 0:
            return

        # Ctrl + rueda → zoom
        if ctrl:
            pivot = None
            try:
                pivot = args.GetPosition(self._scr)
            except Exception:
                pivot = None
            if delta > 0:
                self.zoom_view_in(pivot_viewport=pivot)
            else:
                self.zoom_view_out(pivot_viewport=pivot)
            try:
                args.Handled = True
            except Exception:
                pass
            return

        # Desplazamiento tipo "línea" legible (rueda tipíca 120 unidades).
        step = max(24.0, min(96.0, abs(float(delta)) * 0.55))
        # delta > 0 = rueda hacia arriba → paneo hacia la izquierda del contenido
        # (convención: ver más a la izquierda / inicio de cadena)
        signed = -step if delta > 0 else step

        if shift:
            # Shift + rueda → scroll vertical
            self._scroll_v_by(signed)
        else:
            # Rueda sin modificadores → paneo horizontal (eje principal del alzado)
            self._scroll_h_by(signed)
        try:
            args.Handled = True
        except Exception:
            pass

    def _make_zoom_chrome_btn(self, label, on_click, tooltip=None):
        btn = Button()
        btn.Content = label
        btn.Padding = Thickness(8, 2, 8, 2)
        btn.Margin = Thickness(0, 0, 4, 0)
        btn.FontSize = typo.META_FONT_PX
        btn.FontWeight = FontWeights.SemiBold
        btn.Cursor = Cursors.Hand
        btn.MinWidth = 28.0
        btn.Height = 22.0
        try:
            btn.Foreground = brush_hex(u"#e8f4f8")
            btn.Background = brush_hex(u"#0E1B32")
            btn.BorderBrush = brush_hex(u"#21465C")
            btn.BorderThickness = Thickness(1)
        except Exception:
            pass
        if tooltip:
            try:
                btn.ToolTip = tooltip
            except Exception:
                pass
        if on_click:
            try:
                btn.Click += RoutedEventHandler(lambda s, e: on_click())
            except Exception:
                pass
        return btn

    def _detach_ui_element(self, el):
        """Quita un elemento de su padre visual (para reutilizar el subtree)."""
        if el is None:
            return
        try:
            parent = el.Parent
            if parent is None:
                return
            try:
                parent.Children.Remove(el)
            except Exception:
                try:
                    parent.Child = None
                except Exception:
                    pass
        except Exception:
            pass

    def _build_zoom_chrome(self):
        """Controles − / % / + / 100 % del zoom de alzado (reutiliza el widget)."""
        if self._zoom_chrome_row is not None:
            self._detach_ui_element(self._zoom_chrome_row)
            self._sync_zoom_label()
            return self._zoom_chrome_row

        row = StackPanel()
        row.Orientation = Orientation.Horizontal
        row.VerticalAlignment = VerticalAlignment.Center

        row.Children.Add(
            self._make_zoom_chrome_btn(
                u"−",
                lambda: self.zoom_view_out(),
                tooltip=u"Alejar (Ctrl + rueda)",
            )
        )

        pct = Border()
        pct.Padding = Thickness(6, 2, 6, 2)
        pct.Margin = Thickness(0, 0, 4, 0)
        pct.Background = brush_hex(u"#0b1624", 230)
        pct.BorderBrush = brush_hex(u"#21465C", 180)
        pct.BorderThickness = Thickness(1)
        pct.VerticalAlignment = VerticalAlignment.Center
        try:
            from System.Windows import CornerRadius

            pct.CornerRadius = CornerRadius(3.0)
        except Exception:
            pass
        tb = TextBlock()
        tb.Text = self._view_zoom_label_text()
        tb.Foreground = brush_hex(u"#5bb8d4")
        tb.FontSize = typo.META_FONT_PX
        tb.FontWeight = FontWeights.Bold
        tb.MinWidth = 36.0
        tb.TextAlignment = TextAlignment.Center
        tb.VerticalAlignment = VerticalAlignment.Center
        pct.Child = tb
        self._txt_zoom = tb
        row.Children.Add(pct)

        row.Children.Add(
            self._make_zoom_chrome_btn(
                u"+",
                lambda: self.zoom_view_in(),
                tooltip=u"Acercar (Ctrl + rueda)",
            )
        )
        row.Children.Add(
            self._make_zoom_chrome_btn(
                u"100%",
                lambda: self.zoom_view_reset(),
                tooltip=u"Restaurar zoom 100 %",
            )
        )
        self._zoom_chrome_row = row
        return row

    def _build_elev_header(self):
        """Header del alzado (título + hint + zoom); se reusa entre paints."""
        if self._elev_hdr is not None:
            self._detach_ui_element(self._elev_hdr)
            if self._txt_elev_hint is not None:
                try:
                    self._txt_elev_hint.Text = self._elev_active_hint()
                except Exception:
                    pass
            # Zoom puede vivir dentro del header reusado o reacoplarse.
            try:
                zoom_row = self._build_zoom_chrome()
                if zoom_row.Parent is None:
                    try:
                        Grid.SetColumn(zoom_row, 2)
                    except Exception:
                        pass
                    self._elev_hdr.Children.Add(zoom_row)
            except Exception:
                pass
            return self._elev_hdr

        hdr = Grid()
        hdr.Margin = Thickness(lay.FACE_BLOCK_PAD_PX, 6, lay.FACE_BLOCK_PAD_PX, 4)
        col_title = ColumnDefinition()
        col_title.Width = GridLength(1.0, GridUnitType.Star)
        col_hint = ColumnDefinition()
        col_hint.Width = GridLength.Auto
        col_zoom = ColumnDefinition()
        col_zoom.Width = GridLength.Auto
        hdr.ColumnDefinitions.Add(col_title)
        hdr.ColumnDefinitions.Add(col_hint)
        hdr.ColumnDefinitions.Add(col_zoom)

        title = TextBlock()
        title.Text = u"Alzado · rueda pan H · MMB arrastrar · Ctrl+rueda zoom"
        title.Foreground = brush_hex(u"#64748b")
        title.FontSize = typo.TITLE_FONT_PX
        title.FontWeight = FontWeights.Bold
        title.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(title, 0)
        hdr.Children.Add(title)

        hint = TextBlock()
        hint.Text = self._elev_active_hint()
        hint.Foreground = brush_hex(u"#5bb8d4")
        hint.FontSize = typo.META_FONT_PX
        hint.FontWeight = FontWeights.SemiBold
        hint.VerticalAlignment = VerticalAlignment.Center
        hint.Margin = Thickness(8, 0, 10, 0)
        Grid.SetColumn(hint, 1)
        hdr.Children.Add(hint)
        self._txt_elev_hint = hint

        zoom_row = self._build_zoom_chrome()
        Grid.SetColumn(zoom_row, 2)
        hdr.Children.Add(zoom_row)
        self._elev_hdr = hdr
        return hdr

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

    def _viewport_height(self):
        """Alto útil del ScrCanvas para centrar el bloque de alzado (sin reescalar)."""
        try:
            if self._scr is not None and float(self._scr.ActualHeight or 0.0) > 1.0:
                return float(self._scr.ActualHeight) - 8.0
        except Exception:
            pass
        try:
            if self._host is not None and float(self._host.ActualHeight or 0.0) > 1.0:
                return float(self._host.ActualHeight) - 8.0
        except Exception:
            pass
        return 480.0

    def _show_empty(self, apoyos_loaded):
        if self._txt_tramo:
            self._txt_tramo.Text = u"—"
        if self._txt_apoyos:
            self._txt_apoyos.Text = u"— (seleccione apoyos)" if not apoyos_loaded else u"sin vigas"
        if self._txt_sub:
            self._txt_sub.Text = u"Vigas · columnas · muros · tramos cap-panel"
        if self._txt_sel:
            self._txt_sel.Text = (
                u"Clic viga → selección · Ctrl+clic multi · Ctrl+rueda zoom · "
                u"Emp/Tn en alzado · capas y SUPLE en rail."
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
        grade = u"G25"
        try:
            from armado_vigas.domain.constants import normalize_concrete_grade

            grade = normalize_concrete_grade(
                getattr(self._last_session, "concreteGrade", None)
                if self._last_session is not None
                else None
            )
        except Exception:
            pass

        if self._txt_tramo:
            sum_txt = format_dual_tramo_summary(beams, tramos_sup, tramos_inf)
            self._txt_tramo.Text = u"{0} · {1}".format(sum_txt, grade)
        if self._txt_apoyos:
            if apoyos_loaded:
                a = lay.collect_apoyos(beams)
                n_floors = 0
                try:
                    for ap in getattr(self._last_session, "apoyos", None) or []:
                        try:
                            k = unicode(ap.get("kind") or u"").lower()
                        except Exception:
                            k = u""
                        if k in (u"floor", u"losa", u"slab"):
                            n_floors += 1
                except Exception:
                    n_floors = 0
                if n_floors > 0:
                    base = (
                        u"{0} columna(s) · {1} muro(s) · {2} losa(s) · cadena {3}".format(
                            a["cols"],
                            a["walls"],
                            n_floors,
                            u" → ".join(a["ids"]) if a["ids"] else u"—",
                        )
                    )
                else:
                    base = u"{0} columna(s) · {1} muro(s) · cadena {2}".format(
                        a["cols"], a["walls"], u" → ".join(a["ids"]) if a["ids"] else u"—",
                    )
            else:
                base = u"— (seleccione apoyos)"
            # Vigas unidas (no paralelas al plano de vista = transversales típicas).
            try:
                jf = getattr(self._last_session, "joined_framing", None) or {}
                c = jf.get("counts") or {}
                n_all = int(c.get("all") or 0)
                n_np = int(c.get("not_parallel") or 0)
                n_p = int(c.get("parallel") or 0)
                if n_all > 0:
                    base += u" · unidas {0} ({1} no // vista · {2} //)".format(
                        n_all, n_np, n_p,
                    )
            except Exception:
                pass
            self._txt_apoyos.Text = base
        if self._txt_sub:
            self._txt_sub.Text = u"{0} vigas · sup {1} / inf {2} tramos Tn · {3}".format(
                len(beams), len(tramos_sup or []), len(tramos_inf or []), grade,
            )
        if self._txt_sel:
            card = self._active_rail_card()
            if card == u"conf":
                base = (
                    u"CONF · clic viga/zona · Ctrl+clic multi-viga · ø/@ y dibujo E en lote · "
                    u"rueda pan H · Ctrl+rueda zoom."
                )
            else:
                base = (
                    u"Clic viga → selección · Ctrl+clic multi · rueda pan H · MMB arrastrar · "
                    u"Ctrl+rueda zoom · Emp pills + bandas Tn · rail SUP/INF/LAT/CONF."
                )
            if layout_result.get("needsScroll"):
                base += u" · Contenido más ancho que el panel — pan horizontal o barra inferior."
            self._txt_sel.Text = base

    def _conf_workspace_active(self):
        """True: pestaña CONF + toggle ON → rail sección es interactivo (135°)."""
        return (
            (getattr(self, u"rail_card", None) or u"") == u"conf"
            and bool(getattr(self, u"card_on_conf", True))
        )

    def _resolve_conf_preview_beam(self, beams, session):
        """Viga activa en CONF: siempre la de selección de alzado (no owner de tramo).

        Cada viga lleva su propio estConfDraft / ø@ Ext·Cent; no se hereda del Tn.
        """
        if not beams:
            return None
        idx = self.selected_beam_idx if self.selected_beam_idx >= 0 else 0
        if idx >= len(beams):
            idx = 0
        return beams[idx]

    def _conf_edit_targets(self):
        """Vigas del lote multi-seleccionado en CONF (diám/esp/dibujo).

        Devuelve todas las vigas de ``selected_beam_indices``; si no hay set,
        la primaria ``_conf_beam`` / ``selected_beam_idx``.
        """
        beams = self._last_beams or []
        out = []
        if beams:
            try:
                for i in sorted(self.selected_beam_indices or []):
                    try:
                        ii = int(i)
                    except Exception:
                        continue
                    if 0 <= ii < len(beams):
                        out.append(beams[ii])
            except Exception:
                out = []
        if out:
            return out
        beam = self._conf_beam
        if beam is None and beams:
            try:
                idx = int(self.selected_beam_idx) if self.selected_beam_idx >= 0 else 0
            except Exception:
                idx = 0
            if 0 <= idx < len(beams):
                beam = beams[idx]
        return [beam] if beam is not None else []

    def _apply_conf_to_targets(self, apply_fn):
        """Aplica dibujo a la viga primaria y propaga el draft al lote multi-sel."""
        targets = self._conf_edit_targets()
        if not targets:
            return None
        primary = self._conf_beam
        if primary is None or primary not in targets:
            primary = targets[0]
        out = None
        try:
            out = apply_fn(primary)
        except Exception:
            return out
        if len(targets) <= 1:
            return out
        try:
            d = get_conf_draft(primary)
        except Exception:
            return out
        for b in targets:
            if b is primary:
                continue
            try:
                set_conf_draft(b, d)
            except Exception:
                pass
        return out

    def _conf_status_text(self):
        mode = getattr(self, u"conf_draw_mode", u"draw") or u"draw"
        p = self._conf_pending
        h = self._conf_hover
        d = get_conf_draft(self._conf_beam) if self._conf_beam else {}
        if mode == u"peri":
            return u"Clic en hormigón = toggle perimetral · peri={0}".format(
                bool(d.get(u"perimetral"))
            )
        if mode == u"erase":
            return u"Clic en estribo / traba / columna para borrar"
        if p is None and h is not None:
            return u"Snap → [{0}] · 1.er clic ancla (suelte)".format(h)
        if p is not None and h is not None and h != p:
            return u"Preview E({0}–{1}) · mueva cursor · 2.º clic cierra".format(min(p, h), max(p, h))
        if p is not None and h == p:
            return u"Preview T[{0}] · 2.º clic cierra traba".format(p)
        if p is not None:
            return u"Ancla [{0}] · mueva cursor (marquee) · 2.º clic cierra · Esc cancela".format(p)
        nE = len(d.get(u"pairs") or [])
        nT = len(d.get(u"ties") or [])
        if nE or nT or d.get(u"perimetral"):
            draft_s = u" · draft {0}E/{1}T".format(nE, nT)
            if d.get(u"perimetral"):
                draft_s += u"+P"
            return u"1.er clic ancla · mueva (sin mantener) · 2.º clic cierra{0}".format(draft_s)
        return u"Sin conf. · 1.er clic ancla · mueva cursor · 2.º clic cierra"

    def _wire_section_zoom_btn(self):
        btn = getattr(self, u"_btn_section_zoom", None)
        if btn is None or getattr(self, u"_section_zoom_btn_wired", False):
            return
        try:
            btn.Click += RoutedEventHandler(lambda s, e: self.open_conf_section_zoom())
            self._section_zoom_btn_wired = True
        except Exception:
            pass

    def _sync_section_zoom_btn(self, conf_active):
        btn = getattr(self, u"_btn_section_zoom", None)
        if btn is None:
            return
        try:
            btn.Visibility = Visibility.Visible if conf_active else Visibility.Collapsed
        except Exception:
            try:
                from System.Windows import Visibility as _Vis

                btn.Visibility = _Vis.Visible if conf_active else _Vis.Collapsed
            except Exception:
                pass

    def open_conf_section_zoom(self):
        """Ventana ampliada del canvas de sección (dibujo E/T interactivo)."""
        if not self._conf_workspace_active():
            try:
                self._cb.get(u"on_status", lambda _m: None)(
                    u"Active la pestaña CONF para ampliar la sección."
                )
            except Exception:
                pass
            return
        win = getattr(self, u"_conf_zoom_win", None)
        if win is not None:
            try:
                if win.IsVisible:
                    win.Activate()
                    self._repaint_conf_canvas_only()
                    return
            except Exception:
                self._conf_zoom_win = None
                self._conf_zoom_canvas = None

        win = Window()
        win.Title = u"Arainco · Sección conf. (zoom)"
        win.Width = 580
        win.Height = 680
        win.MinWidth = 480
        win.MinHeight = 520
        win.WindowStartupLocation = WindowStartupLocation.CenterOwner
        try:
            win.ResizeMode = ResizeMode.CanResize
        except Exception:
            pass
        try:
            if self._win is not None:
                win.Owner = self._win
        except Exception:
            pass
        win.Background = brush_hex(u"#0a1620")

        root = Grid()
        root.Margin = Thickness(12)
        for h in (GridLength.Auto, GridLength(1.0, GridUnitType.Star), GridLength.Auto, GridLength.Auto):
            rd = RowDefinition()
            rd.Height = h
            root.RowDefinitions.Add(rd)

        title = TextBlock()
        title.Text = u"Sección ampliada · dibujo estribos / trabas 135°"
        title.Foreground = brush_hex(u"#e2e8f0")
        title.FontSize = 14.0
        title.FontWeight = FontWeights.Bold
        title.Margin = Thickness(0, 0, 0, 8)
        Grid.SetRow(title, 0)
        root.Children.Add(title)

        bdr = Border()
        bdr.Background = brush_hex(u"#071018")
        bdr.BorderBrush = brush_hex(u"#21465C")
        bdr.BorderThickness = Thickness(1)
        bdr.Padding = Thickness(6)
        Grid.SetRow(bdr, 1)

        cnv = Canvas()
        cnv.Width = 520.0
        cnv.Height = 520.0
        cnv.HorizontalAlignment = HorizontalAlignment.Center
        cnv.VerticalAlignment = VerticalAlignment.Center
        try:
            cnv.SnapsToDevicePixels = True
        except Exception:
            pass
        bdr.Child = cnv
        root.Children.Add(bdr)

        status = TextBlock()
        status.Foreground = brush_hex(u"#22d3ee")
        status.FontSize = 11.0
        status.TextWrapping = TextWrapping.Wrap
        status.Margin = Thickness(0, 8, 0, 6)
        status.Text = self._conf_status_text()
        Grid.SetRow(status, 2)
        root.Children.Add(status)

        tools = StackPanel()
        tools.Orientation = Orientation.Horizontal
        tools.HorizontalAlignment = HorizontalAlignment.Right
        Grid.SetRow(tools, 3)

        # Solo Dibujar / Limpiar / Cerrar (Perimetral y Borrar no van en zoom).
        try:
            self.set_conf_draw_mode(u"draw")
        except Exception:
            self.conf_draw_mode = u"draw"

        btn_draw = Button()
        btn_draw.Content = u"Dibujar"
        btn_draw.Padding = Thickness(10, 5, 10, 5)
        btn_draw.Margin = Thickness(0, 0, 8, 0)
        btn_draw.Cursor = Cursors.Hand
        btn_draw.Background = brush_hex(u"#164e63")
        btn_draw.BorderBrush = brush_hex(u"#22d3ee")
        btn_draw.BorderThickness = Thickness(1)
        btn_draw.Foreground = brush_hex(u"#e2e8f0")
        try:
            btn_draw.ToolTip = u"Clic 1 ancla · clic 2 cierra estribo/traba 135°"
        except Exception:
            pass

        def _draw_mode(sender, args):
            self.set_conf_draw_mode(u"draw")
            try:
                status.Text = self._conf_status_text()
            except Exception:
                pass

        try:
            btn_draw.Click += RoutedEventHandler(_draw_mode)
        except Exception:
            pass
        tools.Children.Add(btn_draw)

        btn_clear = Button()
        btn_clear.Content = u"Limpiar"
        btn_clear.Padding = Thickness(10, 5, 10, 5)
        btn_clear.Margin = Thickness(0, 0, 8, 0)
        btn_clear.Cursor = Cursors.Hand
        btn_clear.Background = th.brush_input()
        btn_clear.BorderBrush = th.brush_border()
        btn_clear.BorderThickness = Thickness(1)
        btn_clear.Foreground = brush_hex(u"#e2e8f0")
        try:
            btn_clear.ToolTip = u"Quitar todos los estribos/trabas del draft"
        except Exception:
            pass

        def _clear(sender, args):
            self.clear_conf_draw()
            try:
                self.set_conf_draw_mode(u"draw")
            except Exception:
                pass
            try:
                status.Text = self._conf_status_text()
            except Exception:
                pass

        try:
            btn_clear.Click += RoutedEventHandler(_clear)
        except Exception:
            pass
        tools.Children.Add(btn_clear)

        btn_close = Button()
        btn_close.Content = u"Cerrar"
        btn_close.Padding = Thickness(12, 5, 12, 5)
        btn_close.Cursor = Cursors.Hand
        btn_close.Background = th.brush_input()
        btn_close.BorderBrush = th.brush_border()
        btn_close.BorderThickness = Thickness(1)
        btn_close.Foreground = brush_hex(u"#e2e8f0")

        def _close(sender, args):
            try:
                win.Close()
            except Exception:
                pass

        try:
            btn_close.Click += RoutedEventHandler(_close)
        except Exception:
            pass
        tools.Children.Add(btn_close)
        root.Children.Add(tools)

        win.Content = root
        self._conf_zoom_win = win
        self._conf_zoom_canvas = cnv
        self._conf_zoom_status = status
        self._conf_canvas = cnv

        def _on_closed(sender, args):
            self._conf_zoom_win = None
            self._conf_zoom_canvas = None
            self._conf_zoom_status = None
            try:
                if self._cnv_section is not None:
                    self._conf_canvas = self._cnv_section
            except Exception:
                pass
            try:
                self._repaint_conf_canvas_only()
            except Exception:
                pass

        try:
            from System import EventHandler
            from System import EventArgs

            win.Closed += EventHandler(_on_closed)
        except Exception:
            try:
                win.Closed += _on_closed
            except Exception:
                pass
        try:
            win.KeyDown += KeyEventHandler(self._on_conf_key)
        except Exception:
            pass

        self._wire_conf_canvas(cnv)
        self._repaint_conf_canvas_only()
        try:
            win.Show()
        except Exception:
            try:
                win.ShowDialog()
            except Exception:
                pass

    def _activate_conf_canvas(self, cnv):
        """Activa el canvas (rail o zoom) y su geom. de hit-test."""
        if cnv is None:
            return
        self._conf_canvas = cnv
        by = getattr(self, u"_conf_geom_by_cnv", None) or {}
        g = by.get(id(cnv))
        if g is not None:
            self._conf_geom = g

    def _section_preferred_tramos(self, session):
        """Tn SUP/INF activos (alzado) para resolver armado en sección/CONF."""
        t_sup = t_inf = None
        if session is None:
            return None, None
        try:
            sid = getattr(self, u"selected_tramo_sup_id", None)
            iid = getattr(self, u"selected_tramo_inf_id", None)
            for t in list(getattr(session, u"tramos_sup", None) or []):
                if sid is not None and t.get(u"id") == sid:
                    t_sup = t
                    break
            for t in list(getattr(session, u"tramos_inf", None) or []):
                if iid is not None and t.get(u"id") == iid:
                    t_inf = t
                    break
        except Exception:
            pass
        return t_sup, t_inf

    def _section_arm_kwargs(self, session, beams, beam_idx):
        """Kwargs comunes de armado SUP/INF para ``draw_section_preview``."""
        t_sup, t_inf = self._section_preferred_tramos(session)
        return dict(
            session=session,
            beams=beams or [],
            beam_idx=beam_idx if beam_idx is not None and beam_idx >= 0 else None,
            preferred_tramo_sup=t_sup,
            preferred_tramo_inf=t_inf,
            preferred_tramo_sup_id=getattr(self, u"selected_tramo_sup_id", None),
            preferred_tramo_inf_id=getattr(self, u"selected_tramo_inf_id", None),
        )

    def _paint_conf_on_canvas(self, cnv, snap_r=22.0, out_geom=None):
        """Pinta viga activa en un canvas CONF (rail o zoom)."""
        if cnv is None:
            return None
        beam = self._conf_beam
        if beam is None:
            beams = self._last_beams or []
            beam = self._resolve_conf_preview_beam(beams, self._last_session)
            self._conf_beam = beam
        if beam is None:
            return None
        ensure_beam_confinement(beam)
        session = self._last_session
        if session is None:
            try:
                from armado_vigas.revit.session import SESSION as _S

                session = _S
            except Exception:
                session = None
        laterales_on = bool(getattr(self, "card_on_lat", True)) and bool(
            getattr(session, "lateralesEnabled", True) if session is not None else True
        )
        geom = out_geom if out_geom is not None else {}
        if out_geom is None:
            geom = {}
        beams = self._last_beams or []
        idx = self.selected_beam_idx if self.selected_beam_idx >= 0 else None
        arm_kw = self._section_arm_kwargs(session, beams, idx)
        meta = draw_section_preview(
            cnv,
            beam,
            role_label=u"Sección",
            laterales_enabled=laterales_on,
            n_laterales=session_n_laterales(session, 0),
            diam_laterales=int(getattr(session, "diamLaterales", LATERALES_DIAM_DEFAULT) or LATERALES_DIAM_DEFAULT) if session is not None else int(LATERALES_DIAM_DEFAULT),
            interactive=True,
            pending_bar=self._conf_pending,
            hover_bar=self._conf_hover,
            conf_draw_mode=getattr(self, u"conf_draw_mode", u"draw") or u"draw",
            snap_r=float(snap_r),
            out_geom=geom,
            show_footer=True,
            cursor_xy=getattr(self, u"_conf_cursor", None),
            origin_xy=getattr(self, u"_conf_origin", None),
            **arm_kw
        )
        if not hasattr(self, u"_conf_geom_by_cnv") or self._conf_geom_by_cnv is None:
            self._conf_geom_by_cnv = {}
        self._conf_geom_by_cnv[id(cnv)] = geom
        try:
            cnv.Cursor = Cursors.Cross
        except Exception:
            pass
        self._wire_conf_canvas(cnv)
        return meta

    def _wire_conf_canvas(self, cnv):
        if cnv is None:
            return
        wired = getattr(self, u"_conf_wired_cnv_ids", None)
        if wired is None:
            wired = set()
            self._conf_wired_cnv_ids = wired
        if id(cnv) in wired:
            return
        wired.add(id(cnv))
        handler = MouseButtonEventHandler(self._on_conf_click)
        try:
            from System.Windows import UIElement

            cnv.AddHandler(UIElement.PreviewMouseLeftButtonDownEvent, handler, True)
        except Exception:
            pass
        try:
            from System.Windows import UIElement

            cnv.AddHandler(UIElement.PreviewMouseLeftButtonUpEvent, handler, True)
        except Exception:
            try:
                cnv.MouseLeftButtonUp += handler
            except Exception:
                pass
        try:
            from System.Windows import UIElement

            # Preview: sigue al cursor aunque el mouse esté sobre hijos (puntos/estribos).
            cnv.AddHandler(
                UIElement.PreviewMouseMoveEvent,
                MouseEventHandler(self._on_conf_mouse_move),
                True,
            )
        except Exception:
            try:
                cnv.MouseMove += MouseEventHandler(self._on_conf_mouse_move)
            except Exception:
                pass
        try:
            cnv.MouseLeave += MouseEventHandler(self._on_conf_mouse_leave)
        except Exception:
            pass
        if not getattr(self, u"_conf_key_wired", False) and self._win is not None:
            try:
                self._win.KeyDown += KeyEventHandler(self._on_conf_key)
                self._conf_key_wired = True
            except Exception:
                pass

    def _update_conf_status_ui(self):
        """Status del dibujo CONF bajo el preview de sección (rail) y zoom."""
        if not self._conf_workspace_active():
            return
        msg = self._conf_status_text()
        if self._txt_section is not None:
            try:
                meta = u""
                if self._conf_beam is not None:
                    meta = section_meta_lines(self._conf_beam, u"Confin.")
                self._txt_section.Text = (msg + (u"\n" + meta if meta else u"")).strip()
            except Exception:
                try:
                    self._txt_section.Text = msg
                except Exception:
                    pass
        if self._conf_status_tb is not None:
            try:
                self._conf_status_tb.Text = msg
            except Exception:
                pass
        if getattr(self, u"_conf_zoom_status", None) is not None:
            try:
                self._conf_zoom_status.Text = msg
            except Exception:
                pass
        if self._txt_section_rail is not None:
            try:
                bid = u""
                if self._conf_beam is not None and isinstance(self._conf_beam, dict):
                    bid = unicode(self._conf_beam.get(u"id") or u"")
                self._txt_section_rail.Text = u"Sección CONF · dibujo 135°{0}".format(
                    (u" · " + bid) if bid else u""
                )
            except Exception:
                pass

    def _conf_nearest_bar(self, mx, my):
        geom = self._conf_geom or {}
        hits = geom.get(u"hits") or []
        max_r = float(geom.get(u"snapR") or 12.0)
        by_i = {}
        for h in hits:
            try:
                d = ((mx - float(h[u"x"])) ** 2 + (my - float(h[u"y"])) ** 2) ** 0.5
            except Exception:
                continue
            i = int(h[u"i"])
            if d <= max_r and (i not in by_i or d < by_i[i]):
                by_i[i] = d
        if not by_i:
            return None
        best_i = None
        best_d = max_r + 1.0
        for i, d in by_i.items():
            if d < best_d:
                best_d = d
                best_i = i
        return best_i

    def _conf_hit_pair(self, mx, my):
        for p in (self._conf_geom or {}).get(u"pairHits") or []:
            try:
                if (
                    mx >= float(p[u"rx"])
                    and mx <= float(p[u"rx"]) + float(p[u"rw"])
                    and my >= float(p[u"ry"])
                    and my <= float(p[u"ry"]) + float(p[u"rh"])
                ):
                    return p
            except Exception:
                continue
        return None

    def _conf_inside_concrete(self, mx, my):
        g = self._conf_geom or {}
        try:
            return (
                mx >= float(g[u"ox"])
                and mx <= float(g[u"ox"]) + float(g[u"secW"])
                and my >= float(g[u"oy"])
                and my <= float(g[u"oy"]) + float(g[u"secH"])
            )
        except Exception:
            return False

    def _on_conf_mouse_move(self, sender, args):
        if not self._conf_workspace_active():
            return
        try:
            self._activate_conf_canvas(sender)
        except Exception:
            pass
        if self._conf_canvas is None:
            return
        try:
            pos = args.GetPosition(self._conf_canvas)
            mx, my = float(pos.X), float(pos.Y)
        except Exception:
            return
        next_h = self._conf_nearest_bar(mx, my)
        hover_changed = next_h != self._conf_hover
        self._conf_hover = next_h

        # Con ancla activa: seguir el cursor para la guía (rubber-band).
        cursor_changed = False
        if self._conf_pending is not None:
            prev = getattr(self, u"_conf_cursor", None)
            # Cuantizar ~1 px: el área segmentada debe seguir el cursor con fluidez (Lab).
            q = (int(round(mx)), int(round(my)))
            pq = None
            if prev is not None:
                try:
                    pq = (int(round(float(prev[0]))), int(round(float(prev[1]))))
                except Exception:
                    pq = None
            if q != pq:
                self._conf_cursor = (mx, my)
                cursor_changed = True
        else:
            self._conf_cursor = None

        if hover_changed or cursor_changed:
            self._schedule_conf_repaint()
        try:
            self._conf_canvas.Cursor = (
                Cursors.Hand if next_h is not None
                else (Cursors.Cross if getattr(self, u"conf_draw_mode", u"draw") != u"peri" else Cursors.Hand)
            )
        except Exception:
            pass

    def _clear_conf_marquee(self):
        """Cancela preview de dibujo (ancla + marquee)."""
        self._conf_pending = None
        self._conf_origin = None
        self._conf_cursor = None

    def _try_close_conf_by_marquee(self, mx, my):
        """Si el marquee cubre columnas, cierra E/T sin snap en barra."""
        if self._conf_pending is None:
            return False
        oxoy = getattr(self, u"_conf_origin", None)
        if oxoy is None:
            return False
        try:
            ox, oy = float(oxoy[0]), float(oxoy[1])
            cx, cy = float(mx), float(my)
        except Exception:
            return False
        geom = self._conf_geom or {}
        first_sup = geom.get(u"first_sup") or []
        first_inf = geom.get(u"first_inf") or []
        n = int(geom.get(u"n") or 0)
        if not first_sup or not first_inf or n < 1:
            return False
        cols = []
        try:
            L = min(ox, cx) - 4.0
            R = max(ox, cx) + 4.0
            T = min(oy, cy)
            B = max(oy, cy)
            for i in range(n):
                bx = float(first_sup[i]["x"])
                y0 = min(float(first_sup[i]["y"]), float(first_inf[i]["y"]))
                y1 = max(float(first_sup[i]["y"]), float(first_inf[i]["y"]))
                if bx < L or bx > R:
                    continue
                if B < y0 or T > y1:
                    continue
                cols.append(i)
        except Exception:
            return False
        if not cols:
            return False
        a = int(self._conf_pending)
        others = [c for c in cols if c != a]
        if others:
            b = others[-1]
        elif a in cols:
            b = a
        else:
            b = cols[-1]
        self._clear_conf_marquee()
        try:
            if a == b:
                self._apply_conf_to_targets(lambda bm, k=a: toggle_conf_traba(bm, k))
            else:
                self._apply_conf_to_targets(
                    lambda bm, i0=a, i1=b: toggle_conf_estribo(bm, i0, i1)
                )
        except Exception:
            return False
        try:
            import time as _time

            self._conf_ignore_clicks_until = _time.time() + 0.35
        except Exception:
            self._conf_ignore_clicks_until = 0.0
        self._after_conf_draft_change()
        return True

    def _on_conf_mouse_leave(self, sender, args):
        changed = False
        if self._conf_hover is not None:
            self._conf_hover = None
            changed = True
        # Si hay ancla, no borrar el cursor: el marquee sigue en la última punta.
        if self._conf_pending is None and getattr(self, u"_conf_cursor", None) is not None:
            self._conf_cursor = None
            changed = True
        if changed:
            self._schedule_conf_repaint()

    def _schedule_conf_repaint(self):
        """Coalesce repaints de hover/rubber (~1 frame)."""
        try:
            import time as _time

            now = _time.time()
            self._conf_repaint_wanted = True
            last = float(getattr(self, u"_conf_repaint_last", 0) or 0)
            # Ya hay un flush encolado / hace <30 ms → no saturar UI.
            if getattr(self, u"_conf_repaint_scheduled", False):
                return
            if now - last < 0.03:
                self._conf_repaint_scheduled = True
                disp = None
                try:
                    if self._win is not None:
                        disp = self._win.Dispatcher
                    elif self._conf_canvas is not None:
                        disp = self._conf_canvas.Dispatcher
                except Exception:
                    disp = None
                if disp is not None:
                    from System import Action
                    from System.Windows.Threading import DispatcherPriority

                    def _flush():
                        self._conf_repaint_scheduled = False
                        if not getattr(self, u"_conf_repaint_wanted", False):
                            return
                        self._conf_repaint_wanted = False
                        self._conf_repaint_last = _time.time()
                        self._repaint_conf_canvas_only()

                    try:
                        disp.BeginInvoke(Action(_flush), DispatcherPriority.Background)
                        return
                    except Exception:
                        self._conf_repaint_scheduled = False
            self._conf_repaint_wanted = False
            self._conf_repaint_last = now
            self._repaint_conf_canvas_only()
        except Exception:
            self._repaint_conf_canvas_only()

    def _on_conf_key(self, sender, args):
        try:
            if args.Key == Key.Escape and self._conf_pending is not None:
                self._clear_conf_marquee()
                self._repaint_conf_canvas_only()
                args.Handled = True
        except Exception:
            pass

    def _on_conf_click(self, sender, args):
        if not self._conf_workspace_active():
            return
        try:
            self._activate_conf_canvas(sender)
        except Exception:
            pass
        beam = self._conf_beam
        cnv = self._conf_canvas
        if beam is None or cnv is None:
            return
        try:
            pos = args.GetPosition(cnv)
            mx, my = float(pos.X), float(pos.Y)
        except Exception:
            return
        mode = getattr(self, u"conf_draw_mode", u"draw") or u"draw"
        bi = self._conf_nearest_bar(mx, my)
        self._conf_hover = bi
        # Antireentrada: Down+Up del mismo gesto, o residual post-commit.
        try:
            import time as _time

            now = _time.time()
            if now < float(getattr(self, u"_conf_ignore_clicks_until", 0) or 0):
                try:
                    args.Handled = True
                except Exception:
                    pass
                return
            # Dedup Down/Up del mismo pick (~180 ms, misma barra o misma celda vacía).
            last_t = float(getattr(self, u"_conf_last_click_t", 0) or 0)
            last_k = getattr(self, u"_conf_last_click_key", None)
            key = (mode, bi if bi is not None else u"none", int(mx // 4), int(my // 4))
            if last_k == key and (now - last_t) < 0.18:
                try:
                    args.Handled = True
                except Exception:
                    pass
                return
            self._conf_last_click_t = now
            self._conf_last_click_key = key
        except Exception:
            pass

        if mode == u"erase":
            ph = self._conf_hit_pair(mx, my)
            if ph is not None and bi is None:
                self._apply_conf_to_targets(
                    lambda b, i0=ph[u"i0"], i1=ph[u"i1"]: toggle_conf_estribo(b, i0, i1)
                )
                self._after_conf_draft_change()
                try:
                    args.Handled = True
                except Exception:
                    pass
                return
            if bi is not None:
                d = get_conf_draft(beam)
                if bi in (d.get(u"ties") or []):
                    self._apply_conf_to_targets(lambda b, k=bi: toggle_conf_traba(b, k))
                else:
                    for p in list(d.get(u"pairs") or []):
                        if int(p[0]) == bi or int(p[1]) == bi:
                            self._apply_conf_to_targets(
                                lambda b, i0=p[0], i1=p[1]: toggle_conf_estribo(b, i0, i1)
                            )
                            break
                self._after_conf_draft_change()
                try:
                    args.Handled = True
                except Exception:
                    pass
            return

        if mode == u"peri":
            if bi is None and self._conf_inside_concrete(mx, my):
                self._apply_conf_to_targets(toggle_conf_perimetral)
                self._after_conf_draft_change()
                try:
                    args.Handled = True
                except Exception:
                    pass
            return

        # Dibujar: 1.er clic ancla · mover (sin mantener) · 2.º clic cierra
        if bi is None:
            if self._conf_pending is not None:
                # Intentar cerrar por columnas bajo el marquee (como el mock)
                closed = self._try_close_conf_by_marquee(mx, my)
                if not closed:
                    self._clear_conf_marquee()
                    self._repaint_conf_canvas_only()
            try:
                args.Handled = True
            except Exception:
                pass
            return

        if self._conf_pending is None:
            # 1.er clic: ancla; el botón se suelta y el marquee sigue al mover
            self._conf_pending = int(bi)
            self._conf_origin = (mx, my)
            self._conf_cursor = (mx, my)
            self._repaint_conf_canvas_only()
            try:
                args.Handled = True
            except Exception:
                pass
            return

        a = int(self._conf_pending)
        b = int(bi)
        self._clear_conf_marquee()
        try:
            if a == b:
                self._apply_conf_to_targets(lambda bm, k=a: toggle_conf_traba(bm, k))
            else:
                self._apply_conf_to_targets(
                    lambda bm, i0=a, i1=b: toggle_conf_estribo(bm, i0, i1)
                )
        except Exception:
            pass
        try:
            import time as _time

            self._conf_ignore_clicks_until = _time.time() + 0.35
        except Exception:
            self._conf_ignore_clicks_until = 0.0
        self._after_conf_draft_change()
        try:
            args.Handled = True
        except Exception:
            pass

    def _after_conf_draft_change(self):
        """Actualiza canvas conf (rail) + elev si hay conf dibujable."""
        try:
            self._rail_cache_fp = None
            # Elevación puede mostrar/ocultar conf una vez definido el draft.
            self._elev_cache_fp = None
        except Exception:
            pass
        self._repaint_conf_canvas_only()
        beams = self._last_beams or []
        try:
            rail_cards.populate_section_rail(self, beams, self._last_session)
        except Exception:
            pass
        try:
            # Refresco alzado (conf en elev solo si draft definido)
            self._cb.get(u"on_redraw", lambda: None)()
        except Exception:
            pass
        try:
            conf = find_confin_def(self._conf_beam) or {}
            d = get_conf_draft(self._conf_beam) if self._conf_beam else {}
            msg = u"CONF · {0} · pairs={1} ties={2}".format(
                conf.get(u"label") or u"",
                list(d.get(u"pairs") or []),
                list(d.get(u"ties") or []),
            )
            self._cb.get(u"on_status", lambda _m: None)(msg)
        except Exception:
            pass

    def _repaint_conf_canvas_only(self):
        """Repinta sección del rail y, si está abierta, la ventana zoom."""
        if not self._conf_workspace_active():
            return
        beam = self._conf_beam
        if beam is None:
            beams = self._last_beams or []
            beam = self._resolve_conf_preview_beam(beams, self._last_session)
            self._conf_beam = beam
        if beam is None:
            return
        ensure_beam_confinement(beam)
        try:
            # Preferir geom del canvas activo (zoom si abierto).
            targets = []
            rail = self._cnv_section
            zoom = getattr(self, u"_conf_zoom_canvas", None)
            if rail is not None:
                targets.append((rail, 22.0))
            if zoom is not None:
                targets.append((zoom, 36.0))
            if not targets and self._conf_canvas is not None:
                targets.append((self._conf_canvas, 22.0))
            active = self._conf_canvas
            for cnv, snap in targets:
                geom = {}
                self._paint_conf_on_canvas(cnv, snap_r=snap, out_geom=geom)
            if active is not None:
                self._activate_conf_canvas(active)
            elif zoom is not None:
                self._activate_conf_canvas(zoom)
            elif rail is not None:
                self._activate_conf_canvas(rail)
            self._update_conf_status_ui()
        except Exception:
            pass

    def set_conf_draw_mode(self, mode):
        if mode not in (u"draw", u"peri", u"erase"):
            return
        self.conf_draw_mode = mode
        self._clear_conf_marquee()
        if self._conf_workspace_active():
            self._repaint_conf_canvas_only()
        try:
            beams = self._last_beams or []
            rail_cards.populate_section_rail(self, beams, self._last_session)
        except Exception:
            pass

    def clear_conf_draw(self, beam=None):
        """Limpia conf. de la viga dada o de todo el lote seleccionado en CONF."""
        if beam is not None:
            targets = [beam]
        else:
            targets = self._conf_edit_targets()
        if not targets:
            return
        for b in targets:
            try:
                clear_conf_draft(b)
            except Exception:
                pass
        self._clear_conf_marquee()
        self._after_conf_draft_change()

    def _draw_section_rail(self, beams):
        """Sección + rail mockup (SUP/INF/LAT/CONF + SUPLE).

        Mismo dibujo en las 4 pestañas (estilo SUP). Solo CONF añade interacción
        de dibujo 135° (snap / rubber-band / clics).
        """
        rail_cards.ensure_rail_state(self)
        session = self._last_session
        idx = self.selected_beam_idx if self.selected_beam_idx >= 0 else 0
        beam = beams[idx] if beams else None

        # Preview: siempre la viga seleccionada en alzado (como SUP).
        preview_beam = beam
        conf_interactive = self._conf_workspace_active()
        if conf_interactive and beams:
            conf_beam = self._resolve_conf_preview_beam(beams, session)
            if conf_beam is not None:
                preview_beam = conf_beam
            new_id = None
            try:
                new_id = unicode(preview_beam.get(u"id") or u"") if preview_beam else None
            except Exception:
                new_id = None
            if new_id != getattr(self, u"_conf_preview_id", None):
                self._clear_conf_marquee()
                self._conf_hover = None
            self._conf_preview_id = new_id
            self._conf_beam = preview_beam
        else:
            self._conf_beam = preview_beam
            self._conf_canvas = None

        role = None
        if self.rail_card == u"lat":
            role = u"Laterales"
        elif self.rail_card == u"conf":
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
            if not role:
                role = u"Confin."
        elif self.rail_card == u"sup":
            role = u"Sup."
        elif self.rail_card == u"inf":
            role = u"Inf."

        if self._txt_section_rail and preview_beam:
            role_txt = role or self.rail_card.upper()
            n_sel = len(self.selected_beam_indices)
            if conf_interactive:
                self._txt_section_rail.Text = u"{0} · sección · dibujo 135°".format(
                    lay.beam_canvas_label(idx),
                )
            elif n_sel > 1:
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
            self._txt_section_rail.Text = u"Sección · preview"

        laterales_on = bool(getattr(self, "card_on_lat", True)) and bool(
            getattr(session, "lateralesEnabled", True) if session is not None else True
        )
        n_lat = session_n_laterales(session, 0)
        d_lat = int(getattr(session, "diamLaterales", LATERALES_DIAM_DEFAULT) or LATERALES_DIAM_DEFAULT) if session is not None else int(LATERALES_DIAM_DEFAULT)

        try:
            self._sync_section_zoom_btn(conf_interactive)
            if not conf_interactive:
                zwin = getattr(self, u"_conf_zoom_win", None)
                if zwin is not None:
                    try:
                        zwin.Close()
                    except Exception:
                        self._conf_zoom_win = None
                        self._conf_zoom_canvas = None
        except Exception:
            pass

        # Mismo paint en las 4 pestañas; interactive solo en CONF.
        if self._cnv_section and preview_beam:
            arm_kw = self._section_arm_kwargs(session, beams, idx)
            if conf_interactive:
                ensure_beam_confinement(preview_beam)
                self._conf_canvas = self._cnv_section
                geom = {}
                draw_section_preview(
                    self._cnv_section,
                    preview_beam,
                    role_label=role or u"Sección",
                    laterales_enabled=laterales_on,
                    n_laterales=n_lat,
                    diam_laterales=d_lat,
                    interactive=True,
                    pending_bar=self._conf_pending,
                    hover_bar=self._conf_hover,
                    conf_draw_mode=getattr(self, u"conf_draw_mode", u"draw") or u"draw",
                    snap_r=22.0,
                    out_geom=geom,
                    show_footer=True,
                    cursor_xy=getattr(self, u"_conf_cursor", None),
                    origin_xy=getattr(self, u"_conf_origin", None),
                    **arm_kw
                )
                if not hasattr(self, u"_conf_geom_by_cnv") or self._conf_geom_by_cnv is None:
                    self._conf_geom_by_cnv = {}
                self._conf_geom_by_cnv[id(self._cnv_section)] = geom
                self._conf_geom = geom
                try:
                    self._cnv_section.Cursor = Cursors.Cross
                except Exception:
                    pass
                self._wire_conf_canvas(self._cnv_section)
                # Mantener ventana zoom sincronizada
                zc = getattr(self, u"_conf_zoom_canvas", None)
                if zc is not None:
                    zgeom = {}
                    self._paint_conf_on_canvas(zc, snap_r=36.0, out_geom=zgeom)
            else:
                self._conf_canvas = None
                draw_section_preview(
                    self._cnv_section,
                    preview_beam,
                    role_label=role or u"Sección",
                    laterales_enabled=laterales_on,
                    n_laterales=n_lat,
                    diam_laterales=d_lat,
                    interactive=False,
                    show_footer=True,
                    **arm_kw
                )
                try:
                    self._cnv_section.Cursor = Cursors.Arrow
                except Exception:
                    pass

        if conf_interactive:
            self._update_conf_status_ui()
        elif self._txt_section and preview_beam:
            self._txt_section.Text = section_meta_lines(preview_beam, role)

        rail_cards.populate_section_rail(self, beams, session)

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

    def _elev_line(self, cnv, x1, y1, x2, y2, stroke, thickness=0.9, dash=None, zindex=0):
        batch = getattr(self, u"_elev_geom_batch", None)
        if batch is not None:
            batch.add_line(x1, y1, x2, y2, stroke, thickness, dash, zindex)
            return None
        ln = Line()
        ln.X1 = float(x1)
        ln.Y1 = float(y1)
        ln.X2 = float(x2)
        ln.Y2 = float(y2)
        ln.Stroke = stroke
        ln.StrokeThickness = float(thickness)
        try:
            ln.SnapsToDevicePixels = True
            ln.StrokeStartLineCap = PenLineCap.Flat
            ln.StrokeEndLineCap = PenLineCap.Flat
        except Exception:
            pass
        if dash:
            ln.StrokeDashArray = DoubleCollection(dash)
        if zindex:
            Canvas.SetZIndex(ln, zindex)
        cnv.Children.Add(ln)
        return ln

    def _elev_rect(self, cnv, x, y, w, h, fill=None, stroke=None, thickness=0.9, dash=None, zindex=0):
        batch = getattr(self, u"_elev_geom_batch", None)
        if batch is not None:
            batch.add_rect(x, y, w, h, fill, stroke, thickness, dash, zindex)
            return None
        rect = Rectangle()
        rect.Width = float(w)
        rect.Height = float(h)
        Canvas.SetLeft(rect, float(x))
        Canvas.SetTop(rect, float(y))
        try:
            rect.SnapsToDevicePixels = True
        except Exception:
            pass
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

    def _elev_mm_to_px_u(self, mm):
        """Convierte mm longitudinales a px del alzado (eje U horizontal)."""
        try:
            mm_f = abs(float(mm))
        except (TypeError, ValueError):
            return 0.0
        sc = self._elev_px_per_ft_u() or self._elev_px_per_ft_v()
        if sc:
            return (mm_f / 304.8) * float(sc)
        # Fallback suave si aún no hay meta de escala (≈ 1 cm → 1 px).
        return max(2.0, mm_f * 0.08)

    def _elev_empotramiento_px_for_diam(self, diam_mm):
        """Longitud de empotramiento en px del alzado según Ø y dosificación."""
        try:
            from armado_vigas.domain.concrete_lengths import session_concrete_grade

            grade = session_concrete_grade(getattr(self, u"_last_session", None))
        except Exception:
            grade = None
        emp_mm, _ = empotramiento_mm_for_diam(diam_mm, concrete_grade=grade)
        px = self._elev_mm_to_px_u(emp_mm)
        # Mínimo legible; el máximo evita tapar todo el tramo en zoom lejano.
        return max(6.0, min(float(px or 0.0), 280.0))

    def _elev_mm_to_px_v(self, mm):
        """Convierte mm de sección a px de alzado (preferencia: peralte dibujado).

        Usa ``h_px / h_mm`` de las vigas del canvas cuando está disponible, para
        que 20/25/50 mm se lean proporcionales al hormigón pintado (no al scale U
        del vano horizontal).
        """
        try:
            mm_f = abs(float(mm))
        except (TypeError, ValueError):
            return 0.0
        ppm = getattr(self, u"_elev_px_per_mm_section", None)
        try:
            if ppm and float(ppm) > 1e-9:
                return mm_f * float(ppm)
        except Exception:
            pass
        sc = self._elev_px_per_ft_v() or self._elev_px_per_ft_u()
        if sc:
            return (mm_f / 304.8) * float(sc)
        try:
            return (mm_f / 10.0) * (_ELEV_BEAM_H / max(1.0, float(_ELEV_REF_SECTION_H_CM)))
        except Exception:
            return max(2.0, mm_f * 0.12)

    def _elev_update_section_scale(self, beams):
        """Guarda px/mm a partir del peralte promedio dibujado vs nominal."""
        ratios = []
        for beam in beams or []:
            try:
                _top, h_px = self._elev_beam_vertical(beam)
            except Exception:
                continue
            try:
                h_px = float(h_px)
            except Exception:
                continue
            if h_px < 4.0:
                continue
            h_mm = None
            for key in (u"sectionDepthMm", u"heightMm"):
                try:
                    v = beam.get(key)
                    if v is not None and float(v) > 1.0:
                        h_mm = float(v)
                        break
                except Exception:
                    pass
            if h_mm is None:
                try:
                    h_mm = float(section_height_mm(beam.get(u"type")))
                except Exception:
                    h_mm = None
            if h_mm is None or h_mm < 1.0:
                continue
            ratios.append(h_px / h_mm)
        if ratios:
            self._elev_px_per_mm_section = sum(ratios) / float(len(ratios))
        else:
            # Fallback silueta de referencia.
            try:
                self._elev_px_per_mm_section = float(_ELEV_BEAM_H) / (
                    float(_ELEV_REF_SECTION_H_CM) * 10.0
                )
            except Exception:
                self._elev_px_per_mm_section = None

    def _elev_layer_gap_px(self):
        """Separación entre capas SUP (50 mm) en px de canvas."""
        gap = self._elev_mm_to_px_v(_ELEV_BAR_LAYER_GAP_MM)
        return max(float(_ELEV_EMP_LAYER_GAP), float(gap or 0.0))

    def _elev_stagger_dy_px(self, n_capas=1):
        """Offset Y del solape desacoplado respecto a la fibra de la pareja.

        Controla la separación visible entre las dos barras paralelas del empalme
        (continua arriba / desacoplada abajo). Independiente del gap entre capas.
        """
        dy = self._elev_mm_to_px_v(_ELEV_EMP_PAIR_SEP_MM)
        return max(float(_ELEV_EMP_STAGGER_DY), float(dy or 0.0))

    def _elev_sync_bar_rows(self, beams):
        """Alinea Y de fibras a la silueta proyectada de las vigas.

        1.ª capa SUP: offset fijo 25 mm desde la cara superior (escala V del alzado).
        Capas siguientes: gap 50 mm (eje a eje).
        """
        self._elev_update_section_scale(beams)
        tops = []
        bots = []
        for beam in beams or []:
            top, h = self._elev_beam_vertical(beam)
            tops.append(top)
            bots.append(top + h)
        layer_gap = self._elev_layer_gap_px()
        if tops:
            top_avg = sum(tops) / float(len(tops))
            bot_avg = sum(bots) / float(len(bots))
            h_avg = max(8.0, bot_avg - top_avg)
            cover_px = self._elev_mm_to_px_v(_ELEV_BAR_COVER_SUP_MM)
            # No empujar la fibra más allá de ~40 % del peralte pintado.
            cover_px = max(2.0, min(cover_px, h_avg * 0.40))
            self._elev_bar_sup_y = top_avg + cover_px
            # INF: simétrico con el mismo offset simbólico (recubrimiento).
            self._elev_bar_inf_y = bot_avg - cover_px
            self._elev_bar_suple_sup_y = self._elev_bar_sup_y + layer_gap
            # Hacia el interior del peralte (arriba en canvas), no fuera del fondo.
            self._elev_bar_suple_y = self._elev_bar_inf_y - layer_gap
        else:
            cover_px = self._elev_mm_to_px_v(_ELEV_BAR_COVER_SUP_MM)
            cover_px = max(2.0, cover_px)
            self._elev_bar_sup_y = _ELEV_BEAM_TOP + cover_px
            self._elev_bar_inf_y = _ELEV_BEAM_BOT - cover_px
            self._elev_bar_suple_sup_y = self._elev_bar_sup_y + layer_gap
            self._elev_bar_suple_y = self._elev_bar_inf_y - layer_gap

    def _bar_y_sup(self):
        return getattr(self, u"_elev_bar_sup_y", _ELEV_BAR_SUP_Y)

    def _bar_y_inf(self):
        return getattr(self, u"_elev_bar_inf_y", _ELEV_BAR_INF_Y)

    def _bar_y_suple_sup(self):
        return getattr(self, u"_elev_bar_suple_sup_y", _ELEV_BAR_SUPLE_SUP_Y)

    def _bar_y_suple_inf(self):
        return getattr(self, u"_elev_bar_suple_y", _ELEV_BAR_SUPLE_Y)

    def _elev_apoyo_entry(self, apoyo_id, session):
        for ap in getattr(session, "apoyos", None) or []:
            if ap.get("id") == apoyo_id:
                return ap
        return None

    def _elev_px_per_ft_u(self):
        meta = self._layout_meta or {}
        try:
            s = float(meta.get("pxPerFtU") or 0)
            if s > 1e-12:
                return s
        except (TypeError, ValueError):
            pass
        return None

    def _elev_px_per_ft_v(self):
        meta = self._layout_meta or {}
        try:
            s = float(meta.get("pxPerFtV") or 0)
            if s > 1e-12:
                return s
        except (TypeError, ValueError):
            pass
        return self._elev_px_per_ft_u()

    def _elev_model_v_range(self):
        meta = self._layout_meta or {}
        try:
            v_min = float(meta["modelVMin"])
            v_max = float(meta["modelVMax"])
            return v_min, v_max
        except Exception:
            return None

    def _elev_content_used_height_px(self):
        """Altura del bloque geométrico (V) a escala, sin pads de framing."""
        v_scale = self._elev_px_per_ft_v()
        vr = self._elev_model_v_range()
        if v_scale is None or vr is None:
            elev_h = float((self._layout_meta or {}).get("elevHeightPx") or lay.ELEVATION_HEIGHT_PX)
            pad = float((self._layout_meta or {}).get("elevPadPx") or lay.ELEVATION_PAD_PX)
            return max(4.0, elev_h - pad * 2.0)
        v_min, v_max = vr
        return max(4.0, (float(v_max) - float(v_min)) * float(v_scale))

    def _elev_v_origin_y(self):
        """Offset Y que centra el bloque V en el canvas de elevación (escala fija)."""
        elev_h = float((self._layout_meta or {}).get("elevHeightPx") or lay.ELEVATION_HEIGHT_PX)
        used = self._elev_content_used_height_px()
        return max(0.0, (float(elev_h) - float(used)) * 0.5)

    def _elev_v_to_top_h(self, v_top, v_bot):
        """Convierte rango V del modelo a (top_px, height_px) centrado en el canvas."""
        v_scale = self._elev_px_per_ft_v()
        vr = self._elev_model_v_range()
        if v_scale is None or vr is None or v_top is None or v_bot is None:
            return None
        v_min, v_max = vr
        try:
            vt = float(v_top)
            vb = float(v_bot)
        except (TypeError, ValueError):
            return None
        if vt < vb:
            vt, vb = vb, vt
        # Bloque V centrado en elev_h; Y↓: V alto → y pequeño dentro del bloque.
        origin_y = self._elev_v_origin_y()
        y_top = origin_y + (float(v_max) - vt) * float(v_scale)
        y_bot = origin_y + (float(v_max) - vb) * float(v_scale)
        h = max(4.0, y_bot - y_top)
        return y_top, h

    def _elev_model_u_to_x(self, u_ft):
        """Escalar U del modelo (pies) → X canvas en px."""
        meta = self._layout_meta or {}
        scale = self._elev_px_per_ft_u()
        try:
            u_min = float(meta.get("modelUMin"))
        except Exception:
            return None
        if scale is None:
            return None
        try:
            return float(lay.CANVAS_SIDE_PAD_PX) + (float(u_ft) - u_min) * float(scale)
        except Exception:
            return None

    def _joined_rec_rect_px(self, rec):
        """(left, top, width, height) de viga unida en px, o None."""
        if not rec:
            return None
        # Preferir AABB proyectado del sólido.
        u0 = rec.get("solidUMin")
        u1 = rec.get("solidUMax")
        if u0 is None or u1 is None:
            u0 = rec.get("uStart") if rec.get("uStart") is not None else rec.get("uMin")
            u1 = rec.get("uEnd") if rec.get("uEnd") is not None else rec.get("uMax")
        v0 = rec.get("vMin")
        v1 = rec.get("vMax")
        if u0 is None or u1 is None:
            return None
        try:
            uf0, uf1 = float(u0), float(u1)
        except (TypeError, ValueError):
            return None
        if uf1 < uf0:
            uf0, uf1 = uf1, uf0
        x0 = self._elev_model_u_to_x(uf0)
        x1 = self._elev_model_u_to_x(uf1)
        if x0 is None or x1 is None:
            return None
        left = min(x0, x1)
        width = max(3.0, abs(x1 - x0))
        mapped = self._elev_v_to_top_h(v1, v0) if (v0 is not None and v1 is not None) else None
        if mapped is None:
            # Fallback: alinear con la primera viga principal del canvas.
            top = _ELEV_BEAM_TOP
            h_px = max(6.0, _ELEV_BEAM_H * 0.85)
        else:
            top, h_px = mapped
        return left, top, width, max(4.0, h_px)

    def _draw_elevation_joined_beams(self, cnv, session, content_w):
        """Siluetas de vigas unidas a la selección (// y no // a la vista)."""
        jf = getattr(session, "joined_framing", None) or {}
        recs = list(jf.get("all") or [])
        if not recs:
            return
        # No // primero (destacan), luego paralelas (fondo más sutil).
        recs_sorted = sorted(
            recs, key=lambda r: (1 if r.get("parallelToView") else 0, r.get("id") or u""),
        )
        for rec in recs_sorted:
            rect = self._joined_rec_rect_px(rec)
            if rect is None:
                continue
            left, top, width, h_px = rect
            # Recortar a content_w con un margen de trazo.
            if left + width < 0 or left > float(content_w) + 2.0:
                continue
            parallel = bool(rec.get("parallelToView"))
            edge_hex = _ELEV_JOIN_PAR_EDGE if parallel else _ELEV_JOIN_NPAR_EDGE
            fill_a = _ELEV_JOIN_PAR_FILL_A if parallel else _ELEV_JOIN_NPAR_FILL_A
            sw = _ELEV_JOIN_STROKE if parallel else _ELEV_JOIN_STROKE_NPAR
            fill = brush_hex(edge_hex, fill_a)
            edge = brush_hex(edge_hex, 210 if not parallel else 170)
            dash = None if not parallel else [3.0, 2.5]

            self._elev_rect(
                cnv, left, top, width, h_px, fill=fill, zindex=2,
            )
            # Contorno (completo; dash en lineas horizontales/verticales).
            bot = top + h_px
            self._elev_line(cnv, left, top, left + width, top, edge, sw, dash=dash, zindex=5)
            self._elev_line(cnv, left, bot, left + width, bot, edge, sw, dash=dash, zindex=5)
            self._elev_line(cnv, left, top, left, bot, edge, sw, dash=dash, zindex=5)
            self._elev_line(cnv, left + width, top, left + width, bot, edge, sw, dash=dash, zindex=5)

            # Cruz diagonal interior solo en no paralelas (lectura “sección”).
            if not parallel and width >= 8.0 and h_px >= 8.0:
                self._elev_line(
                    cnv, left + 1.0, top + 1.0, left + width - 1.0, bot - 1.0,
                    brush_hex(edge_hex, 120), 0.55, zindex=5,
                )
                self._elev_line(
                    cnv, left + 1.0, bot - 1.0, left + width - 1.0, top + 1.0,
                    brush_hex(edge_hex, 120), 0.55, zindex=5,
                )

            lbl = TextBlock()
            tag = u"//" if parallel else u"×"
            lbl.Text = u"{0} {1}".format(rec.get("id") or u"V?", tag)
            lbl.FontSize = _ELEV_JOIN_LABEL_FONT
            lbl.FontWeight = FontWeights.SemiBold
            lbl.Foreground = brush_hex(edge_hex, 230)
            lbl.TextAlignment = TextAlignment.Center
            try:
                lbl.Width = max(20.0, width)
            except Exception:
                pass
            Canvas.SetLeft(lbl, left)
            Canvas.SetTop(lbl, top - 11.0 if parallel else top + h_px + 1.0)
            Canvas.SetZIndex(lbl, 6)
            cnv.Children.Add(lbl)

    def _elev_is_floor_apoyo(self, ap):
        if not ap:
            return False
        try:
            k = unicode(ap.get("kind") or u"").lower()
        except Exception:
            k = u""
        if k in (u"floor", u"losa", u"slab"):
            return True
        try:
            aid = unicode(ap.get("id") or u"")
            if aid.startswith(u"L"):
                return True
        except Exception:
            pass
        return False

    def _draw_elevation_floors(self, cnv, session, content_w):
        """Siluetas de losas seleccionadas (banda horizontal en alzado)."""
        apoyos = list(getattr(session, "apoyos", None) or [])
        if not apoyos:
            return
        edge_hex = _ELEV_FLOOR_EDGE
        fill = brush_hex(edge_hex, _ELEV_FLOOR_FILL_A)
        edge = brush_hex(edge_hex, 210)
        sw = _ELEV_FLOOR_STROKE
        for ap in apoyos:
            if not self._elev_is_floor_apoyo(ap):
                continue
            # Reutilizar proyección U/V del modelo (mismo helper que vigas unidas).
            rect = self._joined_rec_rect_px(ap)
            if rect is None:
                continue
            left, top, width, h_px = rect
            if left + width < 0 or left > float(content_w) + 2.0:
                continue
            # Altura mínima legible (espesor pequeño en escala).
            h_px = max(3.0, float(h_px))
            self._elev_rect(cnv, left, top, width, h_px, fill=fill, zindex=3)
            bot = top + h_px
            self._elev_line(
                cnv, left, top, left + width, top, edge, sw, zindex=6,
            )
            self._elev_line(
                cnv, left, bot, left + width, bot, edge, sw, zindex=6,
            )
            self._elev_line(cnv, left, top, left, bot, edge, sw, zindex=6)
            self._elev_line(
                cnv, left + width, top, left + width, bot, edge, sw, zindex=6,
            )
            # Etiqueta + espesor mm si cabe.
            try:
                th = ap.get("thicknessMm") or ap.get("heightMm")
                name = unicode(ap.get("id") or u"L")
                if th:
                    txt = u"{0} · {1}".format(name, int(th))
                else:
                    txt = name
                lbl = TextBlock()
                lbl.Text = txt
                lbl.FontSize = _ELEV_FLOOR_LABEL_FONT
                lbl.FontWeight = FontWeights.SemiBold
                lbl.Foreground = brush_hex(edge_hex, 230)
                lbl.TextAlignment = TextAlignment.Center
                try:
                    lbl.Width = max(24.0, width)
                except Exception:
                    pass
                Canvas.SetLeft(lbl, left)
                Canvas.SetTop(lbl, top - 11.0)
                Canvas.SetZIndex(lbl, 7)
                cnv.Children.Add(lbl)
            except Exception:
                pass

    def _elev_is_wall_id(self, apoyo_id):
        if unicode(apoyo_id or u"").startswith(u"M"):
            return True
        return False

    def _elev_apoyo_is_wall(self, apoyo_id, session):
        """True si el apoyo es muro (por kind o id M-*)."""
        ap = self._elev_apoyo_entry(apoyo_id, session)
        if ap is not None:
            try:
                if unicode(ap.get("kind") or u"").lower() == u"wall":
                    return True
            except Exception:
                pass
        return self._elev_is_wall_id(apoyo_id)

    def _elev_apoyo_width_px(self, apoyo_id, session):
        is_wall = self._elev_is_wall_id(apoyo_id)
        default = _ELEV_WALL_W if is_wall else _ELEV_COL_W
        if not apoyo_id:
            return default
        ap = self._elev_apoyo_entry(apoyo_id, session)
        if not ap:
            return default
        # Preferir escala del layout (misma que longitudes de viga)
        scale_u = self._elev_px_per_ft_u()
        mm = ap.get("widthMm") or ap.get("thicknessMm")
        if scale_u and mm:
            return max(4.0, (float(mm) / 304.8) * scale_u)
        # Rango U real
        try:
            u0 = ap.get("uMin")
            u1 = ap.get("uMax")
            if scale_u and u0 is not None and u1 is not None:
                return max(4.0, abs(float(u1) - float(u0)) * scale_u)
        except (TypeError, ValueError):
            pass
        if mm:
            sc = _ELEV_WALL_PX_PER_MM if is_wall else _ELEV_COL_PX_PER_MM
            return max(default, float(mm) * sc)
        return default

    def _elev_apoyo_half_px(self, apoyo_id, session):
        return self._elev_apoyo_width_px(apoyo_id, session) * 0.5

    def _elev_wall_half_thickness_px(self, apoyo_id, session):
        """Media espesor de muro no // (px U) — legacy; preferir stretch helper."""
        if not apoyo_id or not self._elev_is_wall_id(apoyo_id):
            return 0.0
        ap = self._elev_apoyo_entry(apoyo_id, session)
        if not ap:
            return 0.0
        if bool(ap.get("parallelToView")):
            return 0.0
        mm = ap.get("thicknessMm") or ap.get("widthMm")
        if not mm:
            return 0.0
        try:
            th = float(mm)
        except (TypeError, ValueError):
            return 0.0
        if th < 1.0:
            return 0.0
        scale_u = self._elev_px_per_ft_u()
        if scale_u:
            return max(0.0, (th * 0.5 / 304.8) * float(scale_u))
        try:
            return max(0.0, (th * 0.5) * float(_ELEV_WALL_PX_PER_MM))
        except Exception:
            return max(0.0, th * 0.5 * 0.08)

    def _elev_wall_stretch_from_width_px(self, width_mm):
        """Estirón muro +(ancho/2 − 25 mm) en px U del alzado."""
        try:
            w = float(width_mm)
        except (TypeError, ValueError):
            return 0.0
        stretch_mm = 0.5 * w - float(_ELEV_WALL_END_CLEARANCE_MM)
        if stretch_mm <= 1e-6:
            return 0.0
        scale_u = self._elev_px_per_ft_u()
        if scale_u:
            return max(0.0, (stretch_mm / 304.8) * float(scale_u))
        try:
            return max(0.0, stretch_mm * float(_ELEV_WALL_PX_PER_MM))
        except Exception:
            return max(0.0, stretch_mm * 0.08)

    def _elev_apoyo_u_ft(self, apoyo_id, session):
        """U (ft) de referencia del apoyo en la vista, o None."""
        ap = self._elev_apoyo_entry(apoyo_id, session)
        if not ap:
            return None
        u = ap.get("uView")
        if u is not None:
            try:
                return float(u)
            except (TypeError, ValueError):
                pass
        u0, u1 = ap.get("uMin"), ap.get("uMax")
        if u0 is not None and u1 is not None:
            try:
                return 0.5 * (float(u0) + float(u1))
            except (TypeError, ValueError):
                pass
        return None

    def _elev_wall_apoyo_at_beam_end(self, beam, side, apoyo_id, session, parallel=False):
        """
        True si ``apoyo_id`` es muro (// o no// según ``parallel``) asociado a ese extremo.

        · No //: centro U cerca del extremo (tol. ½ espesor + margen).
        · //: el extremo cae en el rango U del muro (cubre / empotra).
        """
        if not apoyo_id or not self._elev_apoyo_is_wall(apoyo_id, session):
            return False
        ap = self._elev_apoyo_entry(apoyo_id, session)
        if not ap:
            return False
        is_par = bool(ap.get("parallelToView"))
        if parallel and not is_par:
            return False
        if (not parallel) and is_par:
            return False
        u_end = self._elev_beam_end_u_ft(beam, side)
        if u_end is None:
            return False

        if parallel:
            # Muro // a la vista: suele abarcar un tramo en U; no usar solo el mid.
            u0 = ap.get("uMin")
            u1 = ap.get("uMax")
            try:
                th_mm = float(ap.get("thicknessMm") or ap.get("widthMm") or 200.0)
            except (TypeError, ValueError):
                th_mm = 200.0
            margin_ft = (0.5 * th_mm + float(_ELEV_WALL_END_EXTRA_TOL_MM)) / 304.8
            if u0 is not None and u1 is not None:
                try:
                    lo = min(float(u0), float(u1)) - margin_ft
                    hi = max(float(u0), float(u1)) + margin_ft
                    if lo <= float(u_end) <= hi:
                        # Debe estar más cerca de este extremo que del opuesto
                        # (evita emp. en ambos lados por un muro largo).
                        other = u"end" if side == u"start" else u"start"
                        u_other = self._elev_beam_end_u_ft(beam, other)
                        if u_other is None:
                            return True
                        mid = 0.5 * (float(u0) + float(u1))
                        return abs(float(u_end) - mid) <= abs(float(u_other) - mid) + 1e-9
                except (TypeError, ValueError):
                    pass
            # Fallback: apoyo ya asignado a colStart/colEnd.
            return True

        u_ap = self._elev_apoyo_u_ft(apoyo_id, session)
        if u_ap is None:
            return False
        try:
            th_mm = float(ap.get("thicknessMm") or ap.get("widthMm") or 200.0)
        except (TypeError, ValueError):
            th_mm = 200.0
        tol_ft = (0.5 * th_mm + float(_ELEV_WALL_END_EXTRA_TOL_MM)) / 304.8
        d_this = abs(float(u_end) - float(u_ap))
        if d_this > tol_ft:
            return False
        other = u"end" if side == u"start" else u"start"
        u_other = self._elev_beam_end_u_ft(beam, other)
        if u_other is not None:
            d_other = abs(float(u_other) - float(u_ap))
            if d_other + 1e-9 < d_this:
                return False
        return True

    def _elev_find_wall_at_tramo_end(self, beams, tramo, side, session, parallel):
        """
        Muro en extremo libre: primero colStart/colEnd; si no, busca en apoyos.
        """
        idxs = tramo.get("beamIndices") or []
        if not idxs or not beams:
            return None
        bi = idxs[0] if side == u"start" else idxs[-1]
        if not (0 <= bi < len(beams)):
            return None
        beam = beams[bi]
        aid = beam.get("colStart") if side == u"start" else beam.get("colEnd")
        if self._elev_wall_apoyo_at_beam_end(
            beam, side, aid, session, parallel=parallel
        ):
            return self._elev_apoyo_entry(aid, session)
        # Buscar cualquier muro // / no// de la sesión que cubra el extremo.
        u_end = self._elev_beam_end_u_ft(beam, side)
        if u_end is None:
            return None
        best = None
        best_d = 1e9
        for ap in getattr(session, "apoyos", None) or []:
            try:
                if unicode(ap.get("kind") or u"").lower() != u"wall":
                    continue
            except Exception:
                continue
            is_par = bool(ap.get("parallelToView"))
            if parallel and not is_par:
                continue
            if (not parallel) and is_par:
                continue
            aid2 = ap.get("id")
            if self._elev_wall_apoyo_at_beam_end(
                beam, side, aid2, session, parallel=parallel
            ):
                u_ap = self._elev_apoyo_u_ft(aid2, session)
                try:
                    d = abs(float(u_end) - float(u_ap)) if u_ap is not None else 0.0
                except (TypeError, ValueError):
                    d = 0.0
                if d < best_d:
                    best_d = d
                    best = ap
        return best

    def _elev_tramo_wall_end_extensions_px(self, beams, tramo, session):
        """Extensiones start/end (px) por +(ancho/2−25) si hay muro no// en extremo libre."""
        ext_s = 0.0
        ext_e = 0.0
        if tramo.get("edgeStart") != u"half":
            ap = self._elev_find_wall_at_tramo_end(
                beams, tramo, u"start", session, parallel=False
            )
            if ap is not None:
                w = ap.get("thicknessMm") or ap.get("widthMm")
                ext_s = self._elev_wall_stretch_from_width_px(w)
        if tramo.get("edgeEnd") != u"half":
            ap = self._elev_find_wall_at_tramo_end(
                beams, tramo, u"end", session, parallel=False
            )
            if ap is not None:
                w = ap.get("thicknessMm") or ap.get("widthMm")
                ext_e = self._elev_wall_stretch_from_width_px(w)
        return float(ext_s or 0.0), float(ext_e or 0.0)

    def _elev_tramo_wall_parallel_emp_px(self, beams, tramo, session, diam_mm=16):
        """Extensiones emp. (px) si hay muro // a la vista en extremo libre."""
        emp_s = 0.0
        emp_e = 0.0
        emp_px = self._elev_empotramiento_px_for_diam(diam_mm)
        if tramo.get("edgeStart") != u"half":
            ap = self._elev_find_wall_at_tramo_end(
                beams, tramo, u"start", session, parallel=True
            )
            if ap is not None:
                emp_s = float(emp_px or 0.0)
        if tramo.get("edgeEnd") != u"half":
            ap = self._elev_find_wall_at_tramo_end(
                beams, tramo, u"end", session, parallel=True
            )
            if ap is not None:
                emp_e = float(emp_px or 0.0)
        return float(emp_s or 0.0), float(emp_e or 0.0)

    def _draw_elev_emp_mark(self, cnv, x_bar_end, x_tip, y, stroke=None):
        """Marca visual de empotramiento: solo trazo horizontal (magnitud)."""
        if abs(float(x_tip) - float(x_bar_end)) < 1.0:
            return
        stroke = stroke or brush_hex(u"#94a3b8", 200)
        self._elev_line(
            cnv, x_bar_end, y, x_tip, y, stroke, 1.35, zindex=7,
        )

    def _elev_beam_stretch_from_width_px(self, width_mm):
        """Estirón +(b/2 − 25 mm) en px U del alzado."""
        try:
            w = float(width_mm)
        except (TypeError, ValueError):
            return 0.0
        stretch_mm = 0.5 * w - float(_ELEV_BEAM_END_CLEARANCE_MM)
        if stretch_mm <= 1e-6:
            return 0.0
        scale_u = self._elev_px_per_ft_u()
        if scale_u:
            return max(0.0, (stretch_mm / 304.8) * float(scale_u))
        try:
            return max(0.0, stretch_mm * float(_ELEV_WALL_PX_PER_MM))
        except Exception:
            return max(0.0, stretch_mm * 0.08)

    def _elev_joined_npar_recs(self, session):
        jf = getattr(session, "joined_framing", None) or {}
        return list(jf.get("not_parallel") or [])

    def _elev_beam_end_u_ft(self, beam, side):
        """U (ft) del extremo izquierdo (start) o derecho (end) de la viga en vista."""
        if not beam:
            return None
        u0 = beam.get("uStart")
        u1 = beam.get("uEnd")
        if u0 is None or u1 is None:
            u0 = beam.get("solidUMin")
            u1 = beam.get("solidUMax")
        if u0 is None or u1 is None:
            return None
        try:
            uf0, uf1 = float(u0), float(u1)
        except (TypeError, ValueError):
            return None
        return min(uf0, uf1) if side == u"start" else max(uf0, uf1)

    def _elev_joined_npar_at_beam_end(self, beam, side, session):
        """
        Viga unida no // a la vista asociada al extremo libre.

        Prioriza ``sourceBeamIdInts``; si no hay match, cercanía en U al extremo.
        """
        if beam is None or session is None:
            return None
        recs = self._elev_joined_npar_recs(session)
        if not recs:
            return None
        host_id = beam.get("elementIdInt")
        by_host = []
        for rec in recs:
            if bool(rec.get("parallelToView")):
                continue
            sources = rec.get("sourceBeamIdInts") or []
            if host_id is not None and host_id in sources:
                by_host.append(rec)
        if by_host:
            # Preferir la de mayor ancho de sección (empuje más claro en preview).
            def _w(r):
                try:
                    return float(r.get("sectionWidthMm") or r.get("widthMm") or 0.0)
                except (TypeError, ValueError):
                    return 0.0

            by_host.sort(key=_w, reverse=True)
            return by_host[0]

        u_end = self._elev_beam_end_u_ft(beam, side)
        if u_end is None:
            return None
        tol_ft = float(_ELEV_JOIN_END_TOL_MM) / 304.8
        best = None
        best_d = tol_ft
        for rec in recs:
            if bool(rec.get("parallelToView")):
                continue
            u0 = rec.get("solidUMin")
            u1 = rec.get("solidUMax")
            if u0 is None or u1 is None:
                u0 = rec.get("uMin")
                u1 = rec.get("uMax")
            if u0 is None or u1 is None:
                continue
            try:
                mid = 0.5 * (float(u0) + float(u1))
                d = abs(mid - float(u_end))
            except (TypeError, ValueError):
                continue
            if d <= best_d:
                best_d = d
                best = rec
        return best

    def _elev_apoyo_is_column(self, apoyo_id, session):
        """True si el apoyo es columna estructural."""
        if not apoyo_id:
            return False
        ap = self._elev_apoyo_entry(apoyo_id, session)
        if not ap:
            return False
        try:
            return unicode(ap.get("kind") or u"").lower() == u"column"
        except Exception:
            return False

    def _elev_column_apoyo_at_beam_end(self, beam, side, apoyo_id, session):
        """True si ``apoyo_id`` es columna asociada al extremo (cercanía U)."""
        if not apoyo_id or not self._elev_apoyo_is_column(apoyo_id, session):
            return False
        ap = self._elev_apoyo_entry(apoyo_id, session)
        if not ap:
            return False
        u_end = self._elev_beam_end_u_ft(beam, side)
        if u_end is None:
            return False
        u_ap = self._elev_apoyo_u_ft(apoyo_id, session)
        if u_ap is None:
            aid = beam.get("colStart") if side == u"start" else beam.get("colEnd")
            return unicode(aid or u"") == unicode(apoyo_id)
        try:
            w_mm = float(ap.get("widthMm") or 300.0)
        except (TypeError, ValueError):
            w_mm = 300.0
        tol_ft = (0.5 * w_mm + float(_ELEV_WALL_END_EXTRA_TOL_MM)) / 304.8
        d_this = abs(float(u_end) - float(u_ap))
        if d_this > tol_ft:
            return False
        other = u"end" if side == u"start" else u"start"
        u_other = self._elev_beam_end_u_ft(beam, other)
        if u_other is not None:
            d_other = abs(float(u_other) - float(u_ap))
            if d_other + 1e-9 < d_this:
                return False
        return True

    def _elev_find_column_at_tramo_end(self, beams, tramo, side, session):
        """Columna en extremo libre: primero colStart/colEnd; si no, busca en apoyos."""
        idxs = tramo.get("beamIndices") or []
        if not idxs or not beams:
            return None
        bi = idxs[0] if side == u"start" else idxs[-1]
        if not (0 <= bi < len(beams)):
            return None
        beam = beams[bi]
        aid = beam.get("colStart") if side == u"start" else beam.get("colEnd")
        if self._elev_column_apoyo_at_beam_end(beam, side, aid, session):
            return self._elev_apoyo_entry(aid, session)
        u_end = self._elev_beam_end_u_ft(beam, side)
        if u_end is None:
            return None
        best = None
        best_d = 1e9
        for ap in getattr(session, "apoyos", None) or []:
            try:
                if unicode(ap.get("kind") or u"").lower() != u"column":
                    continue
            except Exception:
                continue
            aid2 = ap.get("id")
            if self._elev_column_apoyo_at_beam_end(beam, side, aid2, session):
                u_ap = self._elev_apoyo_u_ft(aid2, session)
                try:
                    d = abs(float(u_end) - float(u_ap)) if u_ap is not None else 0.0
                except (TypeError, ValueError):
                    d = 0.0
                if d < best_d:
                    best_d = d
                    best = ap
        return best

    def _elev_tramo_column_end_extensions_px(self, beams, tramo, session):
        """Extensiones start/end (px) por +(dim/2−25) si hay columna en extremo libre.

        ``widthMm`` del alzado = proyección U ≈ lado de sección según eje de viga.
        """
        ext_s = 0.0
        ext_e = 0.0
        if tramo.get("edgeStart") != u"half":
            ap = self._elev_find_column_at_tramo_end(
                beams, tramo, u"start", session
            )
            if ap is not None:
                ext_s = self._elev_beam_stretch_from_width_px(ap.get("widthMm"))
        if tramo.get("edgeEnd") != u"half":
            ap = self._elev_find_column_at_tramo_end(
                beams, tramo, u"end", session
            )
            if ap is not None:
                ext_e = self._elev_beam_stretch_from_width_px(ap.get("widthMm"))
        return float(ext_s or 0.0), float(ext_e or 0.0)

    def _elev_tramo_beam_end_extensions_px(self, beams, tramo, session):
        """Extensiones start/end (px) por +(b/2−25) si hay viga no// en extremo libre.

        Misma regla post-fusión / pre-troceo de colocación. No aplica en ``half``.
        """
        ext_s = 0.0
        ext_e = 0.0
        idxs = tramo.get("beamIndices") or []
        if not idxs or not beams or session is None:
            return ext_s, ext_e
        if tramo.get("edgeStart") != u"half":
            i0 = idxs[0]
            if 0 <= i0 < len(beams):
                rec = self._elev_joined_npar_at_beam_end(
                    beams[i0], u"start", session
                )
                if rec is not None:
                    w = rec.get("sectionWidthMm") or rec.get("widthMm")
                    ext_s = self._elev_beam_stretch_from_width_px(w)
        if tramo.get("edgeEnd") != u"half":
            i1 = idxs[-1]
            if 0 <= i1 < len(beams):
                rec = self._elev_joined_npar_at_beam_end(
                    beams[i1], u"end", session
                )
                if rec is not None:
                    w = rec.get("sectionWidthMm") or rec.get("widthMm")
                    ext_e = self._elev_beam_stretch_from_width_px(w)
        return float(ext_s or 0.0), float(ext_e or 0.0)

    def _elev_apply_tramo_end_adjust_px(
        self, x0, x1, beams, tramo, session, diam_mm=16,
    ):
        """Aplica estirón muro/viga/columna y emp. muro // a extremos libres del Tn.

        Returns:
            ``(x0, x1, ext_s, ext_e, emp_s, emp_e)`` —
            ``ext_*`` → pata L; ``emp_*`` → empotramiento (sin pata L).
        """
        ext_s = 0.0
        ext_e = 0.0
        emp_s = 0.0
        emp_e = 0.0
        try:
            w_s, w_e = self._elev_tramo_wall_end_extensions_px(
                beams, tramo, session
            )
            ext_s = max(ext_s, float(w_s or 0.0))
            ext_e = max(ext_e, float(w_e or 0.0))
        except Exception:
            pass
        try:
            c_s, c_e = self._elev_tramo_column_end_extensions_px(
                beams, tramo, session
            )
            ext_s = max(ext_s, float(c_s or 0.0))
            ext_e = max(ext_e, float(c_e or 0.0))
        except Exception:
            pass
        try:
            b_s, b_e = self._elev_tramo_beam_end_extensions_px(
                beams, tramo, session
            )
            ext_s = max(ext_s, float(b_s or 0.0))
            ext_e = max(ext_e, float(b_e or 0.0))
        except Exception:
            pass
        try:
            # Emp. muro // solo si ese extremo no tiene ya estirón pata L.
            e_s, e_e = self._elev_tramo_wall_parallel_emp_px(
                beams, tramo, session, diam_mm=diam_mm
            )
            if float(ext_s or 0.0) <= 0.5:
                emp_s = float(e_s or 0.0)
            if float(ext_e or 0.0) <= 0.5:
                emp_e = float(e_e or 0.0)
        except Exception:
            emp_s = emp_e = 0.0
        x0 = x0 - max(0.0, ext_s) - max(0.0, emp_s)
        x1 = x1 + max(0.0, ext_e) + max(0.0, emp_e)
        return (
            x0,
            x1,
            float(ext_s or 0.0),
            float(ext_e or 0.0),
            float(emp_s or 0.0),
            float(emp_e or 0.0),
        )

    def _elev_beam_vertical(self, beam):
        # Proyección real sobre Up de la vista
        mapped = self._elev_v_to_top_h(beam.get("vMax"), beam.get("vMin"))
        if mapped is not None:
            return mapped
        # Fallback por sección nominal (depth)
        h_mm = beam.get("sectionDepthMm") or beam.get("heightMm")
        if h_mm:
            try:
                v_scale = self._elev_px_per_ft_v() or self._elev_px_per_ft_u()
                if v_scale:
                    h_px = max(12.0, (float(h_mm) / 304.8) * v_scale)
                    top = _ELEV_BEAM_BOT - h_px
                    return top, h_px
            except (TypeError, ValueError):
                pass
        w_cm, h_cm = parse_beam_section(beam.get("type"))
        h_px = max(14.0, _ELEV_BEAM_H * (float(h_cm) / _ELEV_REF_SECTION_H_CM))
        top = _ELEV_BEAM_BOT - h_px
        return top, h_px

    def _elev_clamp_span_px(self, left, width, content_w):
        """Asegura contorno derecho/izquierdo visible (stroke no se recorta)."""
        try:
            cw = float(content_w or 0.0)
            left = float(left)
            width = float(width)
        except (TypeError, ValueError):
            return left, width
        if cw <= 2.0:
            return left, width
        # Margen de trazo + 1 px; coincide con pad lateral del layout.
        stroke_pad = max(1.0, float(_ELEV_CONCRETE_STROKE) * 0.5 + 0.75)
        min_left = stroke_pad
        max_right = cw - stroke_pad
        right = left + max(0.0, width)
        if left < min_left:
            right = max(right, min_left + 1.0)
            left = min_left
        if right > max_right:
            right = max_right
        width = max(1.0, right - left)
        return left, width

    def _elev_beam_model_span_px(self, beam, lay_i, content_w, session):
        """
        Span horizontal de silueta en alzado.

        Usa el mismo layout (uStart/uEnd → leftPct/widthPct) que las bandas Tn
        para que viga y tramos compartan eje X. El sólido AABB suele sobresalir
        del eje y desalineaba bandas vs hormigón.
        """
        left_px = lay.pct_to_px(lay_i["leftPct"], content_w)
        width_px = lay.pct_to_px(lay_i["widthPct"], content_w)
        return self._elev_clamp_span_px(left_px, width_px, content_w)

    def _elev_beam_full_span_px(self, beam, lay_i, content_w, session):
        """Silueta horizontal alineada a bandas Tn / labels."""
        return self._elev_beam_model_span_px(beam, lay_i, content_w, session)

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
                dash=[3.0, 2.5] if seg["dashed"] else None,
                zindex=zindex,
            )

    def _draw_elevation_break_line(self, cnv, x, y, w, kind, stroke, zindex=7, thickness=None):
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
        pl.StrokeThickness = float(
            thickness if thickness is not None else _ELEV_CONCRETE_STROKE
        )
        Canvas.SetZIndex(pl, zindex)
        cnv.Children.Add(pl)

    def _draw_elevation_beam_fill(self, cnv, x, w, top, h_px, selected=False):
        self._elev_rect(
            cnv, x, top, w, h_px,
            fill=_elev_concrete_fill_brush(selected=selected),
            zindex=1,
        )
        if selected:
            # Velo suave de foco (no anillo externo).
            self._elev_rect(
                cnv, x, top, w, h_px,
                fill=brush_hex(_ELEV_BEAM_SEL_WASH_HEX, _ELEV_BEAM_SEL_WASH_A),
                zindex=1,
            )
        # Eje a media altura: un poco más visible si está seleccionada.
        axis_a = 120 if selected else 90
        self._elev_line(
            cnv, x, top + h_px * 0.5, x + w, top + h_px * 0.5,
            brush_hex(_ELEV_CONCRETE_EDGE_HEX, axis_a),
            _ELEV_CONCRETE_AXIS_STROKE + (0.15 if selected else 0.0),
            dash=list(_ELEV_CONCRETE_AXIS_DASH),
            zindex=2,
        )

    def _draw_elevation_beam_edges(self, cnv, x, w, top, h_px, zones, selected=False):
        # Contorno hormigón; seleccionado = arista un poco más legible (sin accent neon).
        stroke = _elev_concrete_edge_brush(selected)
        sw = _elev_concrete_stroke_w(selected)
        bot = top + h_px
        self._draw_elevation_horiz_edge(cnv, top, x, x + w, zones, stroke, sw, zindex=3)
        self._draw_elevation_horiz_edge(cnv, bot, x, w + x, zones, stroke, sw, zindex=3)
        self._elev_line(cnv, x, top, x, bot, stroke, sw, zindex=3)
        self._elev_line(cnv, x + w, top, x + w, bot, stroke, sw, zindex=3)

    def _draw_elevation_beam_section_label(self, cnv, x, w, top, beam, selected=False):
        w_cm, h_cm = parse_beam_section(beam.get("type"))
        lbl = TextBlock()
        lbl.Text = u"V. {0}/{1}".format(int(round(w_cm * 10)), int(round(h_cm * 10)))
        lbl.FontSize = typo.META_FONT_PX
        lbl.FontWeight = FontWeights.SemiBold if not selected else FontWeights.Bold
        lbl.Foreground = brush_hex(
            _ELEV_CONCRETE_EDGE_HEX if not selected else _ELEV_CONCRETE_SEL_HEX,
            230 if selected else 210,
        )
        lbl.TextAlignment = TextAlignment.Center
        lbl.Width = w
        Canvas.SetLeft(lbl, x)
        Canvas.SetTop(lbl, top - 14.0)
        Canvas.SetZIndex(lbl, 9)
        cnv.Children.Add(lbl)

    def _draw_elevation_beam_run(self, cnv, run_left, run_w):
        """Legacy — corrida continua sin apoyos."""
        self._draw_elevation_beam_fill(cnv, run_left, run_w, _ELEV_BEAM_TOP, _ELEV_BEAM_H)
        stroke = _elev_concrete_edge_brush(False)
        sw = _elev_concrete_stroke_w(False)
        self._elev_line(
            cnv, run_left, _ELEV_BEAM_TOP, run_left + run_w, _ELEV_BEAM_TOP,
            stroke, sw, zindex=2,
        )
        self._elev_line(
            cnv, run_left, _ELEV_BEAM_BOT, run_left + run_w, _ELEV_BEAM_BOT,
            stroke, sw, zindex=2,
        )

    def _draw_elevation_direction_marker(
        self, cnv, left, width, order_idx, axis_reversed=False, y_mid=None, badge_below=True,
    ):
        """Dirección 0→1: trazo + chevron + pill bajo la silueta de viga."""
        y = y_mid if y_mid is not None else (_ELEV_BEAM_TOP + _ELEV_BEAM_H * 0.5)
        cx = left + width * 0.5
        marker_w = min(
            _ELEV_DIR_MARKER_LEN_PX,
            max(_ELEV_DIR_MARKER_MIN_W_PX, width * 0.18),
        )
        half = marker_w * 0.5
        if width <= half * 2.0 + 8.0:
            return

        to_right = not axis_reversed
        if to_right:
            start_x, end_x = cx - half, cx + half
        else:
            start_x, end_x = cx + half, cx - half

        stroke = brush_hex(_ELEV_DIR_MARKER_HEX, _ELEV_DIR_MARKER_A)
        sw = _ELEV_DIR_MARKER_STROKE
        head_l = _ELEV_DIR_MARKER_HEAD_L
        head_h = _ELEV_DIR_MARKER_HEAD_H
        tick_h = _ELEV_DIR_MARKER_TICK_H
        sign = 1.0 if to_right else -1.0
        tip = end_x
        shaft_end = tip - sign * (head_l - 1.0)

        # Eje: trazo fino bajo la viga.
        self._elev_line(cnv, start_x, y, shaft_end, y, stroke, sw, zindex=4)

        # Origen 0 — tick discreto.
        self._elev_line(
            cnv, start_x, y - tick_h, start_x, y + tick_h, stroke, sw, zindex=4,
        )

        # Punta = chevron abierto.
        self._elev_line(
            cnv, tip - sign * head_l, y - head_h, tip, y, stroke, sw, zindex=4,
        )
        self._elev_line(
            cnv, tip - sign * head_l, y + head_h, tip, y, stroke, sw, zindex=4,
        )

        if order_idx is None:
            return

        n_txt = u"{0}".format(int(order_idx) + 1)
        badge = Border()
        badge.Height = _ELEV_DIR_MARKER_BADGE_H
        badge_w = max(18.0, 8.0 + 7.5 * len(n_txt))
        badge.Width = badge_w
        badge.Padding = Thickness(0)
        badge.Background = brush_hex(_ELEV_DIR_MARKER_BG, 220)
        badge.BorderBrush = brush_hex(_ELEV_DIR_MARKER_HEX, 180)
        badge.BorderThickness = Thickness(1)
        try:
            from System.Windows import CornerRadius

            badge.CornerRadius = CornerRadius(4.0)
        except Exception:
            pass
        tb = TextBlock()
        tb.Text = n_txt
        tb.FontSize = _ELEV_DIR_MARKER_FONT_PX
        tb.FontWeight = FontWeights.SemiBold
        tb.Foreground = brush_hex(_ELEV_DIR_MARKER_HEX, 235)
        tb.VerticalAlignment = VerticalAlignment.Center
        tb.TextAlignment = TextAlignment.Center
        tb.HorizontalAlignment = HorizontalAlignment.Center
        badge.Child = tb

        Canvas.SetLeft(badge, cx - badge_w * 0.5)
        if badge_below:
            # Bajo el trazo (todo el conjunto queda bajo la viga).
            Canvas.SetTop(badge, y + 3.0)
        else:
            Canvas.SetTop(badge, y - _ELEV_DIR_MARKER_BADGE_H - 2.0)
        Canvas.SetZIndex(badge, 5)
        cnv.Children.Add(badge)

    def _draw_elevation_support(self, cnv, x, pt, idx, session):
        col_w = self._elev_apoyo_width_px(pt.get("id"), session)
        ap = self._elev_apoyo_entry(pt.get("id"), session)
        # Losas: dibujo sin hit de suple.
        if ap is not None and not apoyo_allows_suple_sup(ap):
            # Redibuja silueta simple (misma columna visual) y sale sin hit-test.
            pass
        mapped = None
        if ap is not None:
            mapped = self._elev_v_to_top_h(ap.get("vMax"), ap.get("vMin"))
        if mapped is not None:
            col_top, col_h = mapped
        else:
            col_top = _ELEV_COL_TOP
            col_h = _ELEV_COL_H
        left = x - col_w * 0.5
        fill = _elev_concrete_fill_brush()
        edge = _elev_concrete_edge_brush(False)
        sw = _elev_concrete_stroke_w(False)
        is_wall = self._elev_is_wall_id(pt.get("id"))
        aid = pt.get("id")
        try:
            ensure_session_suple_sup(session)
            suple_on = is_apoyo_suple_sup_on(session, aid)
        except Exception:
            suple_on = False
        try:
            sel_id = getattr(session, u"selected_suple_apoyo_id", None)
            is_sel = bool(
                suple_on
                and aid
                and sel_id
                and unicode(aid) == unicode(sel_id)
            )
        except Exception:
            is_sel = False
        allow_hit = ap is None or apoyo_allows_suple_sup(ap)
        if suple_on and allow_hit:
            # Arista teñida; más fuerte si es el apoyo de configuración activa.
            edge_a = (
                _ELEV_SUPLE_APOYO_SEL_EDGE_A
                if is_sel
                else _ELEV_SUPLE_APOYO_EDGE_A
            )
            sw = float(
                _ELEV_SUPLE_APOYO_SEL_STROKE
                if is_sel
                else _ELEV_SUPLE_APOYO_STROKE
            )
            edge = brush_hex(_ELEV_SUPLE_APOYO_EDGE_HEX, edge_a)

        self._elev_rect(cnv, left, col_top, col_w, col_h, fill=fill, zindex=6)
        if suple_on and allow_hit:
            # Velo suave; más denso si está seleccionado para editar.
            wash_a = (
                _ELEV_SUPLE_APOYO_SEL_WASH_A
                if is_sel
                else _ELEV_SUPLE_APOYO_WASH_A
            )
            self._elev_rect(
                cnv, left, col_top, col_w, col_h,
                fill=brush_hex(_ELEV_SUPLE_APOYO_WASH_HEX, wash_a),
                zindex=6,
            )
        self._elev_line(cnv, left, col_top, left, col_top + col_h, edge, sw, zindex=8)
        self._elev_line(
            cnv, left + col_w, col_top, left + col_w, col_top + col_h, edge, sw, zindex=8,
        )
        self._elev_line(
            cnv, x, col_top, x, col_top + col_h,
            (
                brush_hex(_ELEV_SUPLE_APOYO_EDGE_HEX, 140 if is_sel else 100)
                if (suple_on and allow_hit)
                else _elev_concrete_axis_brush()
            ),
            _ELEV_CONCRETE_AXIS_STROKE + (0.2 if is_sel else (0.1 if suple_on else 0.0)),
            dash=list(_ELEV_CONCRETE_AXIS_DASH),
            zindex=7,
        )
        # Líneas de corte solo en columnas; muros con contorno recto (sin break).
        if is_wall:
            self._elev_line(
                cnv, left, col_top, left + col_w, col_top, edge, sw, zindex=8,
            )
            self._elev_line(
                cnv, left, col_top + col_h, left + col_w, col_top + col_h, edge, sw, zindex=8,
            )
        else:
            self._draw_elevation_break_line(
                cnv, left, col_top, col_w, u"top", edge, zindex=8, thickness=sw,
            )
            self._draw_elevation_break_line(
                cnv, left, col_top + col_h, col_w, u"bottom", edge, zindex=8, thickness=sw,
            )

        # Etiqueta de apoyo con dimensión proyectada
        try:
            mm = None
            if ap is not None:
                mm = ap.get("widthMm") or ap.get("thicknessMm")
            mark = u"{0}".format(aid or u"")
            if mm and col_w >= 18.0:
                mark = u"{0}\n{1}".format(aid or u"", int(mm)) if aid else u"{0}".format(int(mm))
            if mark and col_w >= 12.0:
                tb = TextBlock()
                if is_sel:
                    tb.Text = u"★ {0}".format(aid or u"")
                elif suple_on:
                    tb.Text = u"{0}".format(aid or u"")
                else:
                    tb.Text = mark
                    if mm and not aid and col_w >= 18.0:
                        tb.Text = u"{0}".format(int(mm))
                    elif mm and aid and col_w >= 18.0:
                        # Misma convención anterior: ancho si hay hueco
                        tb.Text = u"{0}".format(int(mm))
                tb.FontSize = typo.META_FONT_PX
                tb.FontWeight = (
                    FontWeights.Bold if is_sel
                    else (FontWeights.SemiBold if suple_on else FontWeights.Normal)
                )
                tb.Foreground = (
                    th.brush_sem(th.SEM_SUPLE, 250)
                    if is_sel
                    else (
                        th.brush_sem(th.SEM_SUPLE, 200)
                        if suple_on
                        else brush_hex(_ELEV_CONCRETE_EDGE_HEX, 210)
                    )
                )
                tb.TextAlignment = TextAlignment.Center
                tb.Width = max(col_w, 18.0)
                try:
                    from System.Windows import TextWrapping

                    tb.TextWrapping = TextWrapping.Wrap
                except Exception:
                    pass
                Canvas.SetLeft(tb, left)
                Canvas.SetTop(tb, col_top + 2.0)
                Canvas.SetZIndex(tb, 9)
                cnv.Children.Add(tb)
        except Exception:
            pass

        # Hit: clic = seleccionar (y activar si off); Ctrl+clic = quitar suple.
        if allow_hit and aid and self._suple_sup_apoyo_selection_allowed():
            hit = Border()
            hit.Width = max(10.0, col_w)
            hit.Height = max(10.0, col_h)
            Canvas.SetLeft(hit, left)
            Canvas.SetTop(hit, col_top)
            hit.Background = brush_hex(u"#000000", 1)
            hit.BorderThickness = Thickness(0)
            hit.Cursor = Cursors.Hand
            Canvas.SetZIndex(hit, 20)
            try:
                hit.ToolTip = (
                    u"Clic: seleccionar apoyo (config n·ø en rail) · "
                    u"si no tiene suple, lo activa · "
                    u"Ctrl+clic: quitar · {0}".format(aid)
                )
            except Exception:
                pass

            def _click_apoyo(sender, args, apoyo_id=aid):
                self._select_suple_sup_apoyo_from_canvas(apoyo_id, args)

            try:
                hit.MouseLeftButtonUp += MouseButtonEventHandler(_click_apoyo)
            except Exception:
                pass
            cnv.Children.Add(hit)

    def _suple_sup_apoyo_selection_allowed(self):
        """Definir suple SUP por apoyo: solo con la pestaña SUP activa."""
        try:
            return self._active_rail_card() == u"sup"
        except Exception:
            return False

    def _elev_confin_zone_style(self, role):
        """Color y grosor de ticks Ext · Cent · Uni en alzado (sin velo de zona)."""
        r = (role or u"cent").lower()
        if r == u"ext":
            return th.SEM_EXT, 185, 1.0
        if r == u"uni":
            return u"#fde68a", 185, 1.05
        return th.SEM_CENT, 165, 0.95

    def _elev_confin_spacing_mm(self, beam, role):
        r = (role or u"cent").lower()
        if r == u"ext":
            try:
                return max(50, int(beam.get("estExtSpacing") or ESTRIBO_SPACING_DEFAULT_EXT))
            except Exception:
                return int(ESTRIBO_SPACING_DEFAULT_EXT)
        try:
            return max(50, int(beam.get("estCentSpacing") or ESTRIBO_SPACING_DEFAULT_CENT))
        except Exception:
            return int(ESTRIBO_SPACING_DEFAULT_CENT)

    def _elev_confin_diam_mm(self, beam, role):
        r = (role or u"cent").lower()
        if r == u"ext":
            try:
                return int(beam.get("estExtDiam") or 10)
            except Exception:
                return 10
        try:
            return int(beam.get("estCentDiam") or 8)
        except Exception:
            return 8

    def _draw_elevation_confinement(self, cnv, beams, layouts, content_w, session=None):
        """
        Preview de estribos Ext/Cent (o único) sobre la silueta de cada viga.

        · Solo si la viga tiene conf. definido en CONF (perimetral / E / T)
        · marcas verticales a @ mm (sin velo/achurado de zona) · etiqueta ø@ · clic → rail CONF
        """
        if not bool(getattr(self, u"card_on_conf", True)):
            return
        beams = beams or []
        layouts = layouts or []
        session = session or getattr(self, u"_last_session", None)
        sz = getattr(self, u"selected_stirrup_zone", None) or {}
        try:
            sel_idx = int(sz.get(u"idx")) if sz.get(u"idx") is not None else -1
        except Exception:
            sel_idx = -1
        sel_role = (sz.get(u"role") or u"").lower()
        conf_rail = getattr(self, u"rail_card", None) == u"conf"

        for i, beam in enumerate(beams):
            if i >= len(layouts):
                break
            try:
                ensure_beam_confinement(beam)
            except Exception:
                pass
            # Sin definición en pestaña CONF (pairs/ties/perimetral) → no dibujar.
            if not is_conf_draft_defined(beam):
                continue
            try:
                plan = compute_stirrup_zones(beam)
            except Exception:
                plan = None
            zones = list((plan or {}).get(u"zones") or [])
            if not zones:
                continue

            left, width = self._elev_beam_full_span_px(
                beam, layouts[i], content_w, session
            )
            if width < 6.0:
                continue
            top, h_px = self._elev_beam_vertical(beam)
            try:
                h_px = float(h_px)
                top = float(top)
            except Exception:
                continue
            if h_px < 6.0:
                continue
            bot = top + h_px
            y_tick0 = top + 2.0
            y_tick1 = bot - 2.0
            if y_tick1 - y_tick0 < 4.0:
                y_tick0, y_tick1 = top, bot

            try:
                len_mm = max(1.0, float(beam.get(u"len") or 0.0) * 1000.0)
            except Exception:
                len_mm = 1.0
            px_per_mm = float(width) / len_mm

            for z in zones:
                role = (z.get(u"role") or u"cent").lower()
                try:
                    frac_s = float(z.get(u"fracStart") or 0.0)
                    frac_l = float(z.get(u"fracLen") or 0.0)
                except Exception:
                    continue
                if frac_l <= 1e-6:
                    continue
                x0 = left + frac_s * width
                zw = max(2.0, frac_l * width)
                x0 = max(left, x0)
                x1 = min(left + width, x0 + zw)
                zw = max(2.0, x1 - x0)

                hex_c, edge_a, tick_sw = self._elev_confin_zone_style(role)
                beam_sel = conf_rail and self._is_beam_selected(i)
                role_sel = (
                    beam_sel
                    and (
                        sel_role == role
                        or (role == u"uni" and sel_role in (u"uni", u"cent"))
                    )
                )
                # Resalte si la viga está en multi-sel (label/ticks); rol primario más fuerte.
                is_sel = role_sel or (beam_sel and not sel_role)
                if is_sel:
                    edge_a = min(255, edge_a + (50 if role_sel else 28))
                    tick_sw = float(tick_sw) + (0.3 if role_sel else 0.15)

                # Solo dibujo de estribos (ticks); sin rectángulo/velo de zona.
                sp_mm = self._elev_confin_spacing_mm(beam, role)
                step_px = max(3.5, float(sp_mm) * px_per_mm)
                x_tick = x0 + min(step_px * 0.5, zw * 0.25)
                n_draw = 0
                max_ticks = 80
                stroke_tick = brush_hex(hex_c, min(255, edge_a + 20))
                while x_tick <= x0 + zw - 0.5 and n_draw < max_ticks:
                    self._elev_line(
                        cnv, x_tick, y_tick0, x_tick, y_tick1,
                        stroke_tick, float(tick_sw), zindex=4,
                    )
                    x_tick += step_px
                    n_draw += 1
                if n_draw == 0 and zw >= 3.0:
                    cx = x0 + zw * 0.5
                    self._elev_line(
                        cnv, cx, y_tick0, cx, y_tick1,
                        stroke_tick, float(tick_sw), zindex=4,
                    )

                if zw >= 36.0 and h_px >= 14.0:
                    diam = self._elev_confin_diam_mm(beam, role)
                    role_lbl = (
                        u"Ext" if role == u"ext"
                        else (u"Único" if role == u"uni" else u"Cent")
                    )
                    lbl = TextBlock()
                    lbl.Text = u"{0} ø{1}@{2}".format(
                        role_lbl, int(diam), int(sp_mm),
                    )
                    lbl.FontSize = typo.META_FONT_PX
                    lbl.FontWeight = (
                        FontWeights.SemiBold if is_sel else FontWeights.Normal
                    )
                    lbl.Foreground = brush_hex(hex_c, 230 if is_sel else 200)
                    lbl.TextAlignment = TextAlignment.Center
                    lbl.Width = max(28.0, zw - 4.0)
                    Canvas.SetLeft(lbl, x0 + 2.0)
                    Canvas.SetTop(lbl, top + h_px * 0.5 - 6.0)
                    Canvas.SetZIndex(lbl, 5)
                    cnv.Children.Add(lbl)

                try:
                    # Solo en pestaña CONF: en SUP/LAT/INF no se seleccionan vigas
                    # por zonas Ext/Cent (el hit robaba clics sobre la silueta).
                    if conf_rail:
                        hit = Border()
                        hit.Width = zw
                        hit.Height = h_px
                        Canvas.SetLeft(hit, x0)
                        Canvas.SetTop(hit, top)
                        hit.Background = brush_hex(u"#000000", 1)
                        hit.BorderThickness = Thickness(0)
                        hit.Cursor = Cursors.Hand
                        Canvas.SetZIndex(hit, 12)
                        try:
                            hit.ToolTip = (
                                u"Clic: zona {0} · Ctrl+clic: multi-viga · estribos".format(
                                    u"Ext" if role == u"ext"
                                    else (u"Único" if role == u"uni" else u"Cent")
                                )
                            )
                        except Exception:
                            pass

                        def _click_zone(sender, args, bi=i, r=role):
                            self._select_stirrup_zone_from_elevation(bi, r, args)

                        try:
                            hit.MouseLeftButtonUp += MouseButtonEventHandler(
                                _click_zone
                            )
                        except Exception:
                            pass
                        cnv.Children.Add(hit)
                except Exception:
                    pass

    def _select_stirrup_zone_from_elevation(self, beam_idx, role, args=None):
        """Selecciona zona Ext/Cent/Uni desde el alzado (pestaña CONF).

        Clic simple: esa viga. Ctrl+clic: multi-selección de vigas (ø/@ y dibujo
        se aplican al lote).
        """
        # No navega automáticamente a CONF: solo actúa si ya está en esa pestaña.
        if self._active_rail_card() != u"conf":
            return
        try:
            idx = int(beam_idx)
        except Exception:
            return
        r = (role or u"cent").lower()
        if r not in (u"ext", u"cent", u"uni", u"confin"):
            r = u"cent"
        # Reusa multi-sel de vigas + zona (no forzar índices a una sola viga).
        self._handle_beam_select(idx, args, role=r, update_zone=True, redraw=True)

    def _select_suple_sup_apoyo_from_canvas(self, apoyo_id, args=None):
        """
        Selección de apoyo para configurar n·ø (y activación de suple SUP).

        · Clic: si off → ON + selecciona; si on → solo selecciona (rail).
        · Ctrl+clic: quita el apoyo del set (OFF).
        """
        if not self._suple_sup_apoyo_selection_allowed():
            try:
                self._cb.get("on_status", lambda m: None)(
                    u"Seleccione la pestaña SUP para definir suples por apoyo."
                )
            except Exception:
                pass
            return
        session = getattr(self, u"_last_session", None)
        if session is None or not apoyo_id:
            return
        beams = list(getattr(session, u"domain_beams", None) or [])
        apoyos = list(getattr(session, u"apoyos", None) or [])
        ensure_session_suple_sup(session)
        ap = self._elev_apoyo_entry(apoyo_id, session)
        if ap is not None and not apoyo_allows_suple_sup(ap):
            try:
                self._cb.get("on_status", lambda m: None)(
                    u"Las losas no definen suple superior."
                )
            except Exception:
                pass
            return

        ctrl = False
        try:
            from System.Windows.Input import Keyboard, ModifierKeys

            ctrl = bool(
                (Keyboard.Modifiers & ModifierKeys.Control)
                == ModifierKeys.Control
            )
        except Exception:
            ctrl = False

        was_on = is_apoyo_suple_sup_on(session, apoyo_id)
        if ctrl and was_on:
            set_apoyo_suple_sup(
                session, apoyo_id, False, beams=beams, apoyos=apoyos
            )
            msg = u"Suple SUP · apoyo {0} · OFF".format(apoyo_id)
        else:
            select_apoyo_suple_sup(
                session,
                apoyo_id,
                beams=beams,
                apoyos=apoyos,
                activate_if_off=True,
            )
            n_a, d_a = (2, 16)
            try:
                from armado_vigas.domain.suple_superior import (
                    get_apoyo_suple_sup_arm,
                )

                n_a, d_a = get_apoyo_suple_sup_arm(session, apoyo_id)
            except Exception:
                pass
            adj = adjacent_beams_for_apoyo(beams, apoyo_id)
            parts = []
            for rec in adj[:4]:
                b = rec.get(u"beam") or {}
                parts.append(
                    u"{0} {1}→{2} mm".format(
                        b.get("id") or u"V?",
                        rec.get(u"view_side") or u"",
                        int(rec.get(u"span_mm") or 0),
                    )
                )
            if was_on:
                msg = u"Suple SUP · editando {0} · n={1} · ø{2}".format(
                    apoyo_id, n_a, d_a
                )
            else:
                msg = u"Suple SUP · {0} ON · n={1} · ø{2}".format(
                    apoyo_id, n_a, d_a
                )
            if parts:
                msg += u" · L/3 " + u"; ".join(parts)

        try:
            self._cb.get("on_status", lambda m: None)(msg)
        except Exception:
            pass
        try:
            self._cb.get("on_redraw", lambda: None)()
        except Exception:
            try:
                self.redraw(session)
            except Exception:
                pass

    def _toggle_suple_sup_apoyo_from_canvas(self, apoyo_id):
        """Compat: toggle clásico (preferir ``_select_suple_sup_apoyo_from_canvas``)."""
        if not self._suple_sup_apoyo_selection_allowed():
            try:
                self._cb.get("on_status", lambda m: None)(
                    u"Seleccione la pestaña SUP para definir suples por apoyo."
                )
            except Exception:
                pass
            return
        session = getattr(self, u"_last_session", None)
        if session is None or not apoyo_id:
            return
        beams = list(getattr(session, u"domain_beams", None) or [])
        apoyos = list(getattr(session, u"apoyos", None) or [])
        ensure_session_suple_sup(session)
        ap = self._elev_apoyo_entry(apoyo_id, session)
        if ap is not None and not apoyo_allows_suple_sup(ap):
            try:
                self._cb.get("on_status", lambda m: None)(
                    u"Las losas no definen suple superior."
                )
            except Exception:
                pass
            return
        on = toggle_apoyo_suple_sup(
            session, apoyo_id, beams=beams, apoyos=apoyos
        )
        if on:
            try:
                session.selected_suple_apoyo_id = apoyo_id
            except Exception:
                pass
        adj = adjacent_beams_for_apoyo(beams, apoyo_id)
        parts = []
        for rec in adj[:6]:
            b = rec.get(u"beam") or {}
            parts.append(
                u"{0} {1}→{2} mm".format(
                    b.get("id") or u"V?",
                    rec.get(u"view_side") or u"",
                    int(rec.get(u"span_mm") or 0),
                )
            )
        msg = u"Suple SUP · apoyo {0} · {1}".format(
            apoyo_id, u"ON · L/3" if on else u"OFF"
        )
        if parts and on:
            msg += u" · " + u"; ".join(parts)
        try:
            self._cb.get("on_status", lambda m: None)(msg)
        except Exception:
            pass
        try:
            self._cb.get("on_redraw", lambda: None)()
        except Exception:
            try:
                self.redraw(session)
            except Exception:
                pass

    def _bar_span_edges(self, i, n, left, width):
        inset_l = _ELEV_BAR_INSET if i == 0 else 0.0
        inset_r = _ELEV_BAR_INSET if i == n - 1 else 0.0
        return left + inset_l, left + width - inset_r

    def _sup_tramo_owner_beam(self, beams, tramo):
        """Viga de referencia de armado SUP para un Tn (primera índice del tramo)."""
        idxs = tramo.get("beamIndices") or []
        if not idxs or not beams:
            return None
        i0 = idxs[0]
        if 0 <= i0 < len(beams):
            return beams[i0]
        return None

    def _emplame_lap_half_px(self, layouts, beam_idx, content_w, diam_mm=None):
        """Media longitud de solape (½ traslape tabla Ø/dosificación) en px U."""
        try:
            d = float(diam_mm if diam_mm is not None else 16)
        except Exception:
            d = 16.0
        if d <= 1e-9:
            d = 16.0
        try:
            from armado_vigas.domain.concrete_lengths import (
                lap_mm_for_diameter,
                session_concrete_grade,
            )

            grade = session_concrete_grade(getattr(self, u"_last_session", None))
            L = lap_mm_for_diameter(d, grade)
            if L is not None and float(L) > 1e-6:
                half = 0.5 * float(self._elev_mm_to_px_u(float(L)) or 0.0)
                # Mínimo legible; sin techo artificial para que G25/G35/G45 se distingan.
                return max(8.0, half)
        except Exception:
            pass
        # Fallback simbólico si no hay tablas.
        if beam_idx is None or not layouts or beam_idx < 0 or beam_idx >= len(layouts):
            return max(_ELEV_EMP_LAP_MIN_PX, 20.0) * 0.5
        try:
            w = lay.pct_to_px(layouts[beam_idx]["widthPct"], content_w)
        except Exception:
            w = 80.0
        total = max(
            _ELEV_EMP_LAP_MIN_PX,
            min(_ELEV_EMP_LAP_MAX_PX, float(w) * _ELEV_EMP_LAP_FRAC),
        )
        return total * 0.5

    def _neighbor_tramo_at_half(self, tramos, tramo, edge):
        """Tramo adyacente en el mismo nudo de empalme (half)."""
        idxs = tramo.get("beamIndices") or []
        if not idxs:
            return None
        bi = idxs[0] if edge == u"start" else idxs[-1]
        try:
            from armado_vigas.domain.tramos import find_tramo_half

            # start(half) ↔ tramo anterior (edgeEnd@misma viga); end ↔ siguiente.
            part = 1 if edge == u"start" else 2
            nb = find_tramo_half(tramos, bi, part)
            if nb is not None and nb is not tramo and nb.get("id") != tramo.get("id"):
                return nb
        except Exception:
            pass
        return None

    def _half_junction_max_diam_mm(
        self,
        tramos,
        beams,
        tramo,
        edge,
        face,
        layer,
        session,
        own_diam,
        es_cara_inferior,
    ):
        """Mayor Ø de tramos adyacentes en un empalme (para largo de solape)."""
        try:
            d_own = float(own_diam or 16)
        except Exception:
            d_own = 16.0
        nb = self._neighbor_tramo_at_half(tramos, tramo, edge)
        if nb is None:
            return d_own
        owner_nb = self._sup_tramo_owner_beam(beams, nb)
        d_nb = d_own
        try:
            from armado_vigas.domain.tramo_armado import tramo_layer_diam

            d_nb = float(
                tramo_layer_diam(
                    session, face, nb, layer, es_cara_inferior, owner_nb
                )
                or d_own
            )
        except Exception:
            if owner_nb is not None:
                try:
                    if es_cara_inferior:
                        d_nb = float(beam_layer_diam_inf(owner_nb, layer) or d_own)
                    else:
                        d_nb = float(beam_layer_diam_sup(owner_nb, layer) or d_own)
                except Exception:
                    d_nb = d_own
        return max(d_own, d_nb)

    def _tramo_half_junction(self, layouts, tramo, edge, content_w, diam_mm=None):
        """Centro X (px) de la viga de empalme en un extremo «half» del Tn.

        edge: ``"start"`` | ``"end"``
        """
        idxs = tramo.get("beamIndices") or []
        if not idxs:
            return None, None, None
        bi = idxs[0] if edge == u"start" else idxs[-1]
        if bi < 0 or bi >= len(layouts or []):
            return None, None, None
        try:
            mid_x = lay.pct_to_px(layouts[bi]["centerPct"], content_w)
        except Exception:
            left = lay.pct_to_px(layouts[bi]["leftPct"], content_w)
            w = lay.pct_to_px(layouts[bi]["widthPct"], content_w)
            mid_x = left + w * 0.5
        lap_h = self._emplame_lap_half_px(layouts, bi, content_w, diam_mm=diam_mm)
        return mid_x, lap_h, bi

    def _empalme_layer_j_mid(self, j_mid, lap_h, layer):
        """Centro de empalme por capa (anti-congestión longitudinal).

        - **Capas impares (1, 3…)**: centro de la viga de empalme (sin desfase).
        - **Capas pares (2, 4…)**: desfase + k·lap_h con k=2 ⇒ un solape total
          a la derecha del centro, para que no se apilen los nudos.
        """
        if j_mid is None:
            return None
        try:
            mid = float(j_mid)
        except Exception:
            return j_mid
        try:
            layer_i = max(1, int(layer))
        except Exception:
            layer_i = 1
        # 1.ª y 3.ª (impares): traslapo en el centro de la viga.
        if (layer_i % 2) == 1:
            return mid
        try:
            lap = abs(float(lap_h or 0.0))
        except Exception:
            lap = 0.0
        if lap < 0.5:
            try:
                lap = max(8.0, self._elev_mm_to_px_v(150.0))
            except Exception:
                lap = 16.0
        # k=2: mid_2 = mid + 2·lap_h → zonas [mid−lap, mid+lap] y [mid+lap, mid+3lap]
        # se tocan en el borde sin superponer el cuerpo del solape.
        shift = lap * float(_ELEV_EMP_LAYER_ALT_LAP_K)
        return mid + shift

    def _sup_bar_stagger_polyline_pts(
        self,
        x0,
        x1,
        y,
        edge_start,
        edge_end,
        lap_s,
        lap_e,
        dy,
        layer=1,
        j_mid_s=None,
        j_mid_e=None,
        face=u"sup",
    ):
        """Polilínea empalme: desacople 45° en saliente + desfase por capa.

        Alternancia longitudinal (anti-congestión)::

          capa 1 y 3: empalme en **centro de viga**
          capa 2 (pares): empalme **desplazado** un solape completo (+ 2·lap)
                        para no superponer el nudo de 1.ª/3.ª

        Solo la **saliente** (`edgeEnd=half`) se desacopla a 45°; la entrante
        queda en fibra.

        ``face``: ``sup`` → stagger hacia abajo (+y); ``inf`` → hacia arriba (−y).
        """
        has_s = edge_start == u"half" and lap_s is not None and lap_s > 0.5
        has_e = edge_end == u"half" and lap_e is not None and lap_e > 0.5
        d = abs(float(dy or 0.0))
        if d < 0.5:
            d = 3.5
        # SUP: interior de sección hacia abajo; INF: hacia arriba.
        try:
            toward = -1.0 if unicode(face or u"sup").lower() == u"inf" else 1.0
        except Exception:
            toward = 1.0
        y_off = y + toward * d

        # Centros base (mitad de viga de empalme); fallback al borde del Tn.
        base_s = j_mid_s if j_mid_s is not None else x0
        base_e = j_mid_e if j_mid_e is not None else x1
        j_s = self._empalme_layer_j_mid(base_s, lap_s, layer) if has_s else None
        j_e = self._empalme_layer_j_mid(base_e, lap_e, layer) if has_e else None

        def _saliente_45(x_left, j_mid, lap):
            """Fibra → 45° que aterriza en punta de la entrante → solape past mid."""
            out = []
            other_tip = j_mid - lap
            x_diag0 = other_tip - d
            x_diag1 = other_tip
            x_end = j_mid + lap
            if x_left < x_diag0 - 0.5:
                out.append((x_left, y))
                out.append((x_diag0, y))
            else:
                out.append((x_left, y))
            out.append((x_diag1, y_off))
            if x_end > x_diag1 + 0.25:
                out.append((x_end, y_off))
            return out

        if has_s and has_e:
            # Tramo estrecho entre dos half: solape en ambos extremos con el mid
            # de cada nudo ya desplazado por capa.
            body_left = (j_s - lap_s) if j_s is not None else (x0 - lap_s)
            j_use = j_e if j_e is not None else x1
            pts = _saliente_45(body_left, j_use, lap_e)
            if pts:
                pts[0] = (body_left, y)
        elif has_s:
            # Entrante: fibra colinear; arranca en el fin desplazado del solape.
            tip_l = (j_s - lap_s) if j_s is not None else (x0 - lap_s)
            pts = [(tip_l, y), (x1, y)]
        elif has_e:
            j_use = j_e if j_e is not None else x1
            pts = _saliente_45(x0, j_use, lap_e)
        else:
            pts = [(x0, y), (x1, y)]

        cleaned = []
        for p in pts:
            try:
                px, py = float(p[0]), float(p[1])
            except Exception:
                continue
            if cleaned:
                lx, ly = cleaned[-1]
                if abs(px - lx) < 0.25 and abs(py - ly) < 0.25:
                    continue
                if px + 0.25 < lx:
                    px = lx
            cleaned.append((px, py))
        if len(cleaned) < 2:
            return [(float(x0), float(y)), (float(x1), float(y))]
        return cleaned

    def _elev_polyline(self, cnv, points, stroke, thickness=0.9, zindex=6, dash=None, radius=None):
        """Traza polilínea SUP: Path continuo con filetes suaves en esquinas.

        ``radius`` en px (por defecto ``_ELEV_BAR_CORNER_R``). Esquinas colineales
        se mantienen rectas; juntas angulosas reciben Bezier cuadrática.
        """
        if not points or len(points) < 2:
            return
        pts = []
        for p in points:
            try:
                x, y = float(p[0]), float(p[1])
            except Exception:
                continue
            if pts and abs(x - pts[-1][0]) < 0.2 and abs(y - pts[-1][1]) < 0.2:
                continue
            pts.append((x, y))
        if len(pts) < 2:
            return

        r = float(radius if radius is not None else _ELEV_BAR_CORNER_R)

        def _len(ax, ay, bx, by):
            dx, dy = bx - ax, by - ay
            return (dx * dx + dy * dy) ** 0.5

        def _norm(ax, ay, bx, by):
            L = _len(ax, ay, bx, by)
            if L < 1e-6:
                return 0.0, 0.0, 0.0
            return (bx - ax) / L, (by - ay) / L, L

        def _emit_path(fig):
            geom = PathGeometry()
            geom.Figures.Add(fig)
            freeze_freezable(geom)
            path = Path()
            path.Data = geom
            path.Stroke = stroke
            path.StrokeThickness = float(thickness)
            path.Fill = None
            try:
                path.StrokeStartLineCap = PenLineCap.Round
                path.StrokeEndLineCap = PenLineCap.Round
                path.StrokeLineJoin = PenLineJoin.Round
                path.SnapsToDevicePixels = False
            except Exception:
                pass
            if dash:
                path.StrokeDashArray = DoubleCollection(dash)
            if zindex:
                Canvas.SetZIndex(path, zindex)
            cnv.Children.Add(path)
            return path

        # Recta simple o radio despreciable: Path con caps redondos.
        if len(pts) == 2 or r < 0.5:
            fig = PathFigure()
            fig.StartPoint = Point(pts[0][0], pts[0][1])
            fig.IsClosed = False
            for i in range(1, len(pts)):
                fig.Segments.Add(LineSegment(Point(pts[i][0], pts[i][1]), True))
            return _emit_path(fig)

        fig = PathFigure()
        fig.StartPoint = Point(pts[0][0], pts[0][1])
        fig.IsClosed = False
        segs = fig.Segments
        n = len(pts)

        for i in range(1, n - 1):
            x0, y0 = pts[i - 1]
            x1, y1 = pts[i]
            x2, y2 = pts[i + 1]
            ux, uy, L0 = _norm(x0, y0, x1, y1)
            vx, vy, L1 = _norm(x1, y1, x2, y2)
            if L0 < 0.5 or L1 < 0.5:
                segs.Add(LineSegment(Point(x1, y1), True))
                continue
            # Sin giro (casi colineales): sin filete.
            cross = abs(ux * vy - uy * vx)
            if cross < 0.02:
                segs.Add(LineSegment(Point(x1, y1), True))
                continue
            trim = min(r, L0 * 0.45, L1 * 0.45)
            if trim < 0.6:
                segs.Add(LineSegment(Point(x1, y1), True))
                continue
            ax = x1 - ux * trim
            ay = y1 - uy * trim
            bx = x1 + vx * trim
            by = y1 + vy * trim
            segs.Add(LineSegment(Point(ax, ay), True))
            segs.Add(QuadraticBezierSegment(Point(x1, y1), Point(bx, by), True))

        segs.Add(LineSegment(Point(pts[-1][0], pts[-1][1]), True))
        return _emit_path(fig)

    def _draw_elevation_top_bars(self, cnv, beams, layouts, tramos_sup, content_w, session=None):
        """Preview armadura SUP por Tn — empalme 45° + desfase longitudinal por capa.

        Capa 1 y 3: solape en centro de viga. Capa 2 (pares): solape desplazado.
        Solo la saliente se desacopla a 45°.
        Extremos libres: estirón +(ancho/2−25 mm) + pata L si muro/viga no //
        o columna; emp. según Ø si muro // a la vista.
        empotramiento según Ø si muro // a la vista.
        """
        beams = beams or []
        session = session or getattr(self, u"_last_session", None)
        layer_gap = self._elev_layer_gap_px()

        for tramo in tramos_sup or []:
            if not tramo.get("beamIndices"):
                continue
            owner = self._sup_tramo_owner_beam(beams, tramo)
            if owner is None:
                continue
            try:
                ensure_beam_layers(owner)
            except Exception:
                pass

            accent = tramo.get("accent") or u"#22d3ee"
            warn = tramo_exceeds_bar_limit(tramo)
            span = lay.tramo_span(layouts, tramo, content_w)
            x0_raw = lay.pct_to_px(span["leftPct"], content_w)
            x1_raw = x0_raw + lay.pct_to_px(span["widthPct"], content_w)

            edge_s = tramo.get("edgeStart")
            edge_e = tramo.get("edgeEnd")

            n_capas = beam_n_capas_sup(owner)
            dy = self._elev_stagger_dy_px(n_capas)
            base_y = self._bar_y_sup()
            parts = []
            mode_s = get_bar_end_mode(session, u"sup", u"start")
            mode_e = get_bar_end_mode(session, u"sup", u"end")

            # Estirón no// (pata L) + emp. muro // según Ø de 1.ª capa (ajuste span).
            diam0 = 16
            try:
                from armado_vigas.domain.tramo_armado import tramo_layer_diam

                diam0 = int(
                    tramo_layer_diam(
                        session, u"sup", tramo, 1, False, owner
                    )
                    or beam_layer_diam_sup(owner, 1)
                    or 16
                )
            except Exception:
                try:
                    diam0 = int(beam_layer_diam_sup(owner, 1) or 16)
                except Exception:
                    diam0 = 16
            x0, x1, ext_s, ext_e, emp_s0, emp_e0 = self._elev_apply_tramo_end_adjust_px(
                x0_raw, x1_raw, beams, tramo, session, diam_mm=diam0
            )
            if x1 - x0 < 2.0:
                continue
            if x1 - x0 < 2.0 and edge_s != u"half" and edge_e != u"half":
                continue

            lap_s = None
            lap_e = None
            j_s = None
            j_e = None
            bi_s = None
            bi_e = None
            if edge_s == u"half":
                d_j = self._half_junction_max_diam_mm(
                    tramos_sup, beams, tramo, u"start", u"sup", 1,
                    session, diam0, False,
                )
                j_s, lap_s, bi_s = self._tramo_half_junction(
                    layouts, tramo, u"start", content_w, diam_mm=d_j,
                )
            if edge_e == u"half":
                d_j = self._half_junction_max_diam_mm(
                    tramos_sup, beams, tramo, u"end", u"sup", 1,
                    session, diam0, False,
                )
                j_e, lap_e, bi_e = self._tramo_half_junction(
                    layouts, tramo, u"end", content_w, diam_mm=d_j,
                )

            for layer in range(1, n_capas + 1):
                y = base_y + (layer - 1) * layer_gap
                try:
                    from armado_vigas.domain.tramo_armado import (
                        tramo_layer_bar_count,
                        tramo_layer_diam,
                    )

                    n_bars = tramo_layer_bar_count(
                        session, u"sup", tramo, layer, False, owner
                    )
                    diam = int(
                        tramo_layer_diam(
                            session, u"sup", tramo, layer, False, owner
                        )
                        or 16
                    )
                except Exception:
                    try:
                        n_bars = layer_bar_count(owner, layer, False)
                    except Exception:
                        n_bars = 2
                    try:
                        diam = int(beam_layer_diam_sup(owner, layer) or 16)
                    except Exception:
                        diam = 16
                # Trazo fino (preview): mismo peso que suples SUP/INF.
                sw = self._elev_bar_stroke_w(diam)
                if warn:
                    stroke = brush_hex(u"#fbbf24", 235 if layer == 1 else 190)
                elif layer == 1:
                    stroke = accent_soft_brush(accent, "strokeSel")
                else:
                    stroke = accent_soft_brush(accent, "stroke")

                # Traslape @ empalme: mayor Ø de tramos adyacentes (por capa).
                if edge_s == u"half" and bi_s is not None:
                    d_j = self._half_junction_max_diam_mm(
                        tramos_sup, beams, tramo, u"start", u"sup", layer,
                        session, diam, False,
                    )
                    lap_s = self._emplame_lap_half_px(
                        layouts, bi_s, content_w, diam_mm=d_j
                    )
                if edge_e == u"half" and bi_e is not None:
                    d_j = self._half_junction_max_diam_mm(
                        tramos_sup, beams, tramo, u"end", u"sup", layer,
                        session, diam, False,
                    )
                    lap_e = self._emplame_lap_half_px(
                        layouts, bi_e, content_w, diam_mm=d_j
                    )

                # Emp. muro // solo por Ø de capa (estirón ext_* reutilizado).
                emp_s, emp_e = self._elev_emp_for_layer_px(
                    beams, tramo, session, diam, ext_s, ext_e
                )
                lx0 = x0_raw - max(0.0, ext_s) - max(0.0, emp_s)
                lx1 = x1_raw + max(0.0, ext_e) + max(0.0, emp_e)
                # Empotramiento UI forzado (si no hay ya emp. por muro //).
                if edge_s != u"half" and mode_s == BAR_END_MODE_EMP and emp_s <= 0.5:
                    lx0 = lx0 - self._elev_empotramiento_px_for_diam(diam)
                if edge_e != u"half" and mode_e == BAR_END_MODE_EMP and emp_e <= 0.5:
                    lx1 = lx1 + self._elev_empotramiento_px_for_diam(diam)

                pts = self._sup_bar_stagger_polyline_pts(
                    lx0, lx1, y, edge_s, edge_e, lap_s, lap_e, dy,
                    layer=layer, j_mid_s=j_s, j_mid_e=j_e,
                )
                self._elev_polyline(cnv, pts, stroke, sw, zindex=6)
                # Empotramiento muro //: marca horizontal según Ø (sin pata L).
                if edge_s != u"half" and float(emp_s or 0.0) > 0.5:
                    x_anch = x0_raw - max(0.0, ext_s)
                    self._draw_elev_emp_mark(cnv, x_anch, lx0, y, stroke)
                if edge_e != u"half" and float(emp_e or 0.0) > 0.5:
                    x_anch = x1_raw + max(0.0, ext_e)
                    self._draw_elev_emp_mark(cnv, x_anch, lx1, y, stroke)
                # Pata L: UI forzada o colisión muro/viga/columna (no en emp. muro //).
                pata_s = (
                    edge_s != u"half"
                    and emp_s <= 0.5
                    and (
                        mode_s == BAR_END_MODE_PATA_L
                        or float(ext_s or 0.0) > 0.5
                    )
                )
                pata_e = (
                    edge_e != u"half"
                    and emp_e <= 0.5
                    and (
                        mode_e == BAR_END_MODE_PATA_L
                        or float(ext_e or 0.0) > 0.5
                    )
                )
                if pata_s:
                    self._draw_elev_pata_l_mark(
                        cnv, lx0, y, u"start", u"sup", stroke,
                        diam_mm=diam, session=session,
                    )
                if pata_e:
                    self._draw_elev_pata_l_mark(
                        cnv, lx1, y, u"end", u"sup", stroke,
                        diam_mm=diam, session=session,
                    )
                parts.append(u"{0}ø{1}".format(n_bars, diam))

            mid_label_w = max(24.0, x1 - x0)
            if mid_label_w >= 28.0 and parts:
                lbl = TextBlock()
                body = u" + ".join(parts)
                lbl.Text = u"T{0} · {1}".format(tramo.get("id"), body)
                if warn:
                    lbl.Text = lbl.Text + u" · >12 m"
                lbl.FontSize = typo.META_FONT_PX
                lbl.FontWeight = FontWeights.Bold
                if warn:
                    lbl.Foreground = brush_hex(u"#fbbf24", 240)
                else:
                    lbl.Foreground = accent_soft_brush(accent, "text")
                lbl.TextAlignment = TextAlignment.Center
                lbl.Width = mid_label_w
                Canvas.SetLeft(lbl, x0)
                Canvas.SetTop(lbl, base_y - 12.0)
                Canvas.SetZIndex(lbl, 7)
                cnv.Children.Add(lbl)

    def _elev_hook_mm_for_diam(self, diam_mm, session=None):
        """Largo de pata L / gancho (mm) según Ø y dosificación (tablas G25/G35/G45)."""
        try:
            d = float(diam_mm or 16)
        except (TypeError, ValueError):
            d = 16.0
        if d <= 1e-9:
            d = 16.0
        try:
            from armado_vigas.domain.concrete_lengths import (
                hook_mm_for_diameter,
                session_concrete_grade,
            )

            sess = session or getattr(self, u"_last_session", None)
            return max(
                0.0,
                float(hook_mm_for_diameter(d, session_concrete_grade(sess)) or 0.0),
            )
        except Exception:
            pass
        try:
            from geometria_empotramiento_extremos import _hook_mm_desde_diametro
            from armado_vigas.domain.concrete_lengths import session_concrete_grade

            sess = session or getattr(self, u"_last_session", None)
            return max(
                0.0,
                float(_hook_mm_desde_diametro(d, session_concrete_grade(sess)) or 0.0),
            )
        except Exception:
            # Respaldo típico: 12·Ø
            return max(50.0, 12.0 * d)

    def _elev_pata_l_px_for_diam(self, diam_mm, session=None):
        """Largo de pata L en px de alzado (eje V / peralte)."""
        hook_mm = self._elev_hook_mm_for_diam(diam_mm, session=session)
        px = self._elev_mm_to_px_v(hook_mm)
        # Legible pero acotado al peralte dibujado.
        return max(4.0, min(float(px or 0.0), 120.0))

    def _draw_elev_pata_l_mark(
        self, cnv, x, y, side, face, stroke=None, diam_mm=16, session=None,
    ):
        """
        Gancho pata L en alzado: pata hacia el interior de la sección.

        El doblez en el extremo libre se dibuja con un filete ligero (Bezier)
        entre el tramo horizontal y la pata vertical.
        """
        stroke = stroke or brush_hex(u"#94a3b8", 200)
        drop = float(self._elev_pata_l_px_for_diam(diam_mm, session=session) or 0.0)
        if drop < 1.0:
            return
        # SUP: pata hacia abajo (+y); INF: hacia arriba (−y).
        y_tip = y + drop if face == u"sup" else y - drop
        # Radio ligero del doblez (∝ largo de pata, techo legible).
        r = max(2.2, min(drop * 0.22, 8.0, float(_ELEV_BAR_CORNER_R) + 2.0))
        approach = max(r * 1.35, 3.5)
        # Enfoque desde el vano (no hacia fuera): simula el doblez sin «Z».
        if side == u"start":
            x_in = float(x) + approach
        else:
            x_in = float(x) - approach
        pts = [
            (x_in, float(y)),
            (float(x), float(y)),
            (float(x), float(y_tip)),
        ]
        self._elev_polyline(
            cnv, pts, stroke, thickness=1.1, zindex=7, radius=r,
        )

    def _draw_elev_free_end_marks(
        self, cnv, x0, x1, y, edge_s, edge_e, face, session, stroke=None, diam_mm=16,
    ):
        """
        Preview en alzado de extremos libres (INF / marcas sueltas).

        Empotramiento: estira con longitud de tabla según Ø.
        Pata L: gancho. Auto: sin marca.
        """
        if x1 - x0 < 4.0:
            return
        stroke = stroke or brush_hex(u"#94a3b8", 200)
        for side, edge, x in (
            (u"start", edge_s, x0),
            (u"end", edge_e, x1),
        ):
            if edge == u"half":
                continue
            mode = get_bar_end_mode(session, face, side)
            if mode == BAR_END_MODE_AUTO:
                continue
            outward = -1.0 if side == u"start" else 1.0
            if mode == BAR_END_MODE_PATA_L:
                self._draw_elev_pata_l_mark(
                    cnv, x, y, side, face, stroke,
                    diam_mm=diam_mm, session=session,
                )
            elif mode == BAR_END_MODE_EMP:
                ext = self._elev_empotramiento_px_for_diam(diam_mm)
                self._elev_line(
                    cnv, x, y, x + outward * ext, y, stroke, 1.2, zindex=7,
                )

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
                self._elev_line(cnv, cursor, self._bar_y_inf(), bar_left, self._bar_y_inf(), inf_solid, _ELEV_STROKE_BAR, zindex=4)
            if col["hook"] and bar_left < bar_right:
                drop = self._bar_y_inf() - _ELEV_POCKET_D
                self._elev_line(cnv, bar_left, self._bar_y_inf(), bar_left, drop, inf_solid, _ELEV_STROKE_BAR, zindex=4)
                self._elev_line(
                    cnv, bar_left, drop, bar_right, drop, inf_hidden, 2.0, dash=[2.5, 2.0], zindex=4,
                )
                self._elev_line(cnv, bar_right, drop, bar_right, self._bar_y_inf(), inf_solid, _ELEV_STROKE_BAR, zindex=4)
            cursor = max(cursor, bar_right)
        if cursor < seg_x1:
            self._elev_line(cnv, cursor, self._bar_y_inf(), seg_x1, self._bar_y_inf(), inf_solid, _ELEV_STROKE_BAR, zindex=4)

    def _draw_elevation_bottom_bars(self, cnv, beams, layouts, tramos_inf, content_w, session=None):
        """Preview armadura INF por Tn — mismas reglas post-fusión que SUP.

        Capas + empalme 45° + desfase por capa.
        Extremos libres: estirón +(ancho/2−25 mm) + pata L si muro/viga no //
        o columna; emp. según Ø si muro // a la vista.
        """
        beams = beams or []
        session = session or getattr(self, u"_last_session", None)
        layer_gap = self._elev_layer_gap_px()

        for tramo in tramos_inf or []:
            if not tramo.get("beamIndices"):
                continue
            owner = self._sup_tramo_owner_beam(beams, tramo)
            if owner is None:
                continue
            try:
                ensure_beam_layers(owner)
            except Exception:
                pass

            accent = tramo.get("accent") or u"#fb7185"
            warn = tramo_exceeds_bar_limit(tramo)
            span = lay.tramo_span(layouts, tramo, content_w)
            x0_raw = lay.pct_to_px(span["leftPct"], content_w)
            x1_raw = x0_raw + lay.pct_to_px(span["widthPct"], content_w)

            edge_s = tramo.get("edgeStart")
            edge_e = tramo.get("edgeEnd")

            n_capas = beam_n_capas_inf(owner)
            dy = self._elev_stagger_dy_px(n_capas)
            base_y = self._bar_y_inf()
            parts = []
            mode_s = get_bar_end_mode(session, u"inf", u"start")
            mode_e = get_bar_end_mode(session, u"inf", u"end")

            # Estirón no// (pata L) + emp. muro // según Ø de 1.ª capa (ajuste span).
            diam0 = 16
            try:
                from armado_vigas.domain.tramo_armado import tramo_layer_diam

                diam0 = int(
                    tramo_layer_diam(
                        session, u"inf", tramo, 1, True, owner
                    )
                    or beam_layer_diam_inf(owner, 1)
                    or 16
                )
            except Exception:
                try:
                    diam0 = int(beam_layer_diam_inf(owner, 1) or 16)
                except Exception:
                    diam0 = 16
            x0, x1, ext_s, ext_e, emp_s0, emp_e0 = self._elev_apply_tramo_end_adjust_px(
                x0_raw, x1_raw, beams, tramo, session, diam_mm=diam0
            )
            if x1 - x0 < 2.0:
                continue
            if x1 - x0 < 2.0 and edge_s != u"half" and edge_e != u"half":
                continue

            lap_s = None
            lap_e = None
            j_s = None
            j_e = None
            bi_s = None
            bi_e = None
            if edge_s == u"half":
                d_j = self._half_junction_max_diam_mm(
                    tramos_inf, beams, tramo, u"start", u"inf", 1,
                    session, diam0, True,
                )
                j_s, lap_s, bi_s = self._tramo_half_junction(
                    layouts, tramo, u"start", content_w, diam_mm=d_j,
                )
            if edge_e == u"half":
                d_j = self._half_junction_max_diam_mm(
                    tramos_inf, beams, tramo, u"end", u"inf", 1,
                    session, diam0, True,
                )
                j_e, lap_e, bi_e = self._tramo_half_junction(
                    layouts, tramo, u"end", content_w, diam_mm=d_j,
                )

            for layer in range(1, n_capas + 1):
                # Capas hacia el interior de la sección (arriba en canvas).
                y = base_y - (layer - 1) * layer_gap
                try:
                    from armado_vigas.domain.tramo_armado import (
                        tramo_layer_bar_count,
                        tramo_layer_diam,
                    )

                    n_bars = tramo_layer_bar_count(
                        session, u"inf", tramo, layer, True, owner
                    )
                    diam = int(
                        tramo_layer_diam(
                            session, u"inf", tramo, layer, True, owner
                        )
                        or 16
                    )
                except Exception:
                    try:
                        n_bars = layer_bar_count(owner, layer, True)
                    except Exception:
                        n_bars = 2
                    try:
                        diam = int(beam_layer_diam_inf(owner, layer) or 16)
                    except Exception:
                        diam = 16
                if edge_s == u"half" and bi_s is not None:
                    d_j = self._half_junction_max_diam_mm(
                        tramos_inf, beams, tramo, u"start", u"inf", layer,
                        session, diam, True,
                    )
                    lap_s = self._emplame_lap_half_px(
                        layouts, bi_s, content_w, diam_mm=d_j
                    )
                if edge_e == u"half" and bi_e is not None:
                    d_j = self._half_junction_max_diam_mm(
                        tramos_inf, beams, tramo, u"end", u"inf", layer,
                        session, diam, True,
                    )
                    lap_e = self._emplame_lap_half_px(
                        layouts, bi_e, content_w, diam_mm=d_j
                    )
                sw = self._elev_bar_stroke_w(diam)
                if warn:
                    stroke = brush_hex(u"#fbbf24", 235 if layer == 1 else 190)
                elif layer == 1:
                    stroke = accent_soft_brush(accent, "strokeSel")
                else:
                    stroke = accent_soft_brush(accent, "stroke")

                # Emp. muro // por capa (estirón ext_* reutilizado).
                emp_s, emp_e = self._elev_emp_for_layer_px(
                    beams, tramo, session, diam, ext_s, ext_e
                )
                lx0 = x0_raw - max(0.0, ext_s) - max(0.0, emp_s)
                lx1 = x1_raw + max(0.0, ext_e) + max(0.0, emp_e)
                if edge_s != u"half" and mode_s == BAR_END_MODE_EMP and emp_s <= 0.5:
                    lx0 = lx0 - self._elev_empotramiento_px_for_diam(diam)
                if edge_e != u"half" and mode_e == BAR_END_MODE_EMP and emp_e <= 0.5:
                    lx1 = lx1 + self._elev_empotramiento_px_for_diam(diam)

                pts = self._sup_bar_stagger_polyline_pts(
                    lx0, lx1, y, edge_s, edge_e, lap_s, lap_e, dy,
                    layer=layer, j_mid_s=j_s, j_mid_e=j_e, face=u"inf",
                )
                self._elev_polyline(cnv, pts, stroke, sw, zindex=6)
                if edge_s != u"half" and float(emp_s or 0.0) > 0.5:
                    x_anch = x0_raw - max(0.0, ext_s)
                    self._draw_elev_emp_mark(cnv, x_anch, lx0, y, stroke)
                if edge_e != u"half" and float(emp_e or 0.0) > 0.5:
                    x_anch = x1_raw + max(0.0, ext_e)
                    self._draw_elev_emp_mark(cnv, x_anch, lx1, y, stroke)
                pata_s = (
                    edge_s != u"half"
                    and emp_s <= 0.5
                    and (
                        mode_s == BAR_END_MODE_PATA_L
                        or float(ext_s or 0.0) > 0.5
                    )
                )
                pata_e = (
                    edge_e != u"half"
                    and emp_e <= 0.5
                    and (
                        mode_e == BAR_END_MODE_PATA_L
                        or float(ext_e or 0.0) > 0.5
                    )
                )
                if pata_s:
                    self._draw_elev_pata_l_mark(
                        cnv, lx0, y, u"start", u"inf", stroke,
                        diam_mm=diam, session=session,
                    )
                if pata_e:
                    self._draw_elev_pata_l_mark(
                        cnv, lx1, y, u"end", u"inf", stroke,
                        diam_mm=diam, session=session,
                    )
                parts.append(u"{0}ø{1}".format(n_bars, diam))

            mid_label_w = max(24.0, x1 - x0)
            if mid_label_w >= 28.0 and parts:
                lbl = TextBlock()
                body = u" + ".join(parts)
                lbl.Text = u"T{0} · {1}".format(tramo.get("id"), body)
                if warn:
                    lbl.Text = lbl.Text + u" · >12 m"
                lbl.FontSize = typo.META_FONT_PX
                lbl.FontWeight = FontWeights.Bold
                if warn:
                    lbl.Foreground = brush_hex(u"#fbbf24", 240)
                else:
                    lbl.Foreground = accent_soft_brush(accent, "text")
                lbl.TextAlignment = TextAlignment.Center
                lbl.Width = mid_label_w
                Canvas.SetLeft(lbl, x0)
                Canvas.SetTop(lbl, base_y + 2.0)
                Canvas.SetZIndex(lbl, 7)
                cnv.Children.Add(lbl)

    def _elev_beam_face_bar_y(self, beam, face):
        """
        Y de 1.ª capa longitudinal en la silueta de esta viga.

        ``face``: ``"sup"`` | ``"inf"``. Usa el peralte dibujado, no el promedio
        global del lote (evita fibras fuera en vigas de distinta altura).
        """
        top, h_px = self._elev_beam_vertical(beam)
        try:
            h_px = float(h_px)
            top = float(top)
        except Exception:
            return self._bar_y_sup() if face != u"inf" else self._bar_y_inf()
        if h_px < 4.0:
            return self._bar_y_sup() if face != u"inf" else self._bar_y_inf()
        bot = top + h_px
        cover = self._elev_mm_to_px_v(_ELEV_BAR_COVER_SUP_MM)
        cover = max(2.0, min(float(cover or 0.0), h_px * 0.35))
        if face == u"inf":
            return bot - cover
        return top + cover

    def _suple_y_sup_for_beam(self, beam):
        """Y de suple SUP = capa (n_capas_sup + 1) dentro de la silueta."""
        layer_gap = self._elev_layer_gap_px()
        try:
            n_capas = max(1, int(beam_n_capas_sup(beam) or 1))
        except Exception:
            n_capas = 1
        base = self._elev_beam_face_bar_y(beam, u"sup")
        # Capas hacia el interior (y+ en canvas).
        y = float(base) + float(n_capas) * float(layer_gap)
        top, h_px = self._elev_beam_vertical(beam)
        try:
            bot = float(top) + float(h_px)
            cover = max(2.0, min(self._elev_mm_to_px_v(_ELEV_BAR_COVER_SUP_MM), float(h_px) * 0.35))
            y_lo = float(top) + cover
            y_hi = bot - cover
            if y < y_lo:
                y = y_lo
            if y > y_hi:
                y = y_hi
        except Exception:
            pass
        return y

    def _suple_y_inf_for_beam(self, beam):
        """Y de suple INF = capa (n_capas_inf + 1) hacia el interior de la viga."""
        layer_gap = self._elev_layer_gap_px()
        try:
            n_capas = max(1, int(beam_n_capas_inf(beam) or 1))
        except Exception:
            n_capas = 1
        base = self._elev_beam_face_bar_y(beam, u"inf")
        # Misma convención que barras INF: capa k en base − (k−1)·gap (y− = interior).
        # Suple = capa n_capas+1 → base − n_capas·gap.
        y = float(base) - float(n_capas) * float(layer_gap)
        top, h_px = self._elev_beam_vertical(beam)
        try:
            bot = float(top) + float(h_px)
            cover = max(2.0, min(self._elev_mm_to_px_v(_ELEV_BAR_COVER_SUP_MM), float(h_px) * 0.35))
            y_lo = float(top) + cover
            # No salir por el fondo: como máximo la 1.ª capa INF.
            y_hi = float(base)
            if y < y_lo:
                y = y_lo
            if y > y_hi:
                y = y_hi
            # Seguridad: siempre dentro del hormigón dibujado.
            if y > bot - 1.0:
                y = bot - max(cover, 2.0)
            if y < float(top) + 1.0:
                y = float(top) + max(cover, 2.0)
        except Exception:
            # Fallback: no fuera del promedio INF global.
            try:
                y = min(y, float(self._bar_y_inf()))
            except Exception:
                pass
        return y

    def _elev_bar_stroke_w(self, diam_mm):
        """Espesor de fibra longitudinal en alzado (SUP/INF/suples, por Ø).

        Misma curva: ø16 ≈ 0.9 px, techo ~1.5 px.
        """
        try:
            d = float(diam_mm or 16)
        except Exception:
            d = 16.0
        return max(0.7, min(1.5, 0.55 + d / 28.0))

    def _elev_suple_stroke_w(self, diam_mm):
        """Alias histórico: mismo espesor que SUP/INF."""
        return self._elev_bar_stroke_w(diam_mm)

    def _elev_suple_sup_fake_tramo(self, seg):
        """
        Tramo sintético para reutilizar el ajuste de extremos de barras SUP.

        Solo el extremo de **apoyo** es libre (estirón / emp / pata L). El corte
        interior L/3 es ``half`` (sin extensión). Tramos ``merged``: ambos
        extremos son cortes de vano → sin estirón de apoyo en canvas.
        """
        idxs = list((seg or {}).get("indices") or [])
        typ = (seg or {}).get("type") or (
            u"merged" if (seg or {}).get("merged") else u""
        )
        if typ == u"start" and idxs:
            return {
                u"beamIndices": [idxs[0]],
                u"edgeStart": u"free",
                u"edgeEnd": u"half",
            }
        if typ == u"end" and idxs:
            return {
                u"beamIndices": [idxs[0]],
                u"edgeStart": u"half",
                u"edgeEnd": u"free",
            }
        if typ == u"merged" and len(idxs) >= 2:
            return {
                u"beamIndices": list(idxs),
                u"edgeStart": u"half",
                u"edgeEnd": u"half",
            }
        return None

    def _elev_suple_sup_apply_real_ends_px(
        self, x0_raw, x1_raw, beams, seg, session, diam_mm,
    ):
        """
        Magnitud real en px: L/3 base + estirón no// + emp. muro // (Ø).

        Returns:
            ``(x0, x1, ext_s, ext_e, emp_s, emp_e)``
        """
        tramo = self._elev_suple_sup_fake_tramo(seg)
        if tramo is None or session is None:
            return (
                float(x0_raw),
                float(x1_raw),
                0.0,
                0.0,
                0.0,
                0.0,
            )
        try:
            return self._elev_apply_tramo_end_adjust_px(
                float(x0_raw),
                float(x1_raw),
                beams,
                tramo,
                session,
                diam_mm=diam_mm,
            )
        except Exception:
            return (
                float(x0_raw),
                float(x1_raw),
                0.0,
                0.0,
                0.0,
                0.0,
            )

    def _elev_suple_sup_side_seg(self, beam_index, view_side, left, width):
        """Segmento L/3 de un lado de viga (coords px de vano completo)."""
        span_w = float(width) * float(SUPLE_END_PCT)
        left = float(left)
        width = float(width)
        if view_side == u"start":
            return {
                u"type": u"start",
                u"indices": [beam_index],
                u"x0": left,
                u"x1": left + span_w,
                u"merged": False,
            }
        return {
            u"type": u"end",
            u"indices": [beam_index],
            u"x0": left + width - span_w,
            u"x1": left + width,
            u"merged": False,
        }

    def _draw_elevation_suple_superior_tramos(self, cnv, beams, layouts, content_w):
        """
        Fibra suple SUP a magnitud real (misma post-fusión que barras SUP):

        L/3 (+ fusión) + estirón no// (muro/viga/columna) + emp. muro //
        + marca pata L cuando hay estirón.
        """
        from armado_vigas.domain.suple_superior import (
            beam_suple_sup_layer_index,
            get_apoyo_suple_sup_arm,
            resolve_suple_sup_arm_for_spec,
        )

        zone_fill = SolidColorBrush(Color.FromArgb(18, 192, 132, 252))
        zone_stroke = th.brush_sem(th.SEM_SUPLE, 72)
        suple_stroke = th.brush_sem(th.SEM_SUPLE, 230)
        suple_merged = brush_hex(u"#e879f9", 240)
        dash = [3.0, 2.0]
        beams = beams or []
        session = getattr(self, u"_last_session", None)
        try:
            ensure_session_suple_sup(session)
            sync_beams_suple_from_apoyo_set(session, beams)
        except Exception:
            pass

        # Zonas: L/3 + estirones/emp del extremo libre (magnitud real).
        for i, beam in enumerate(beams):
            if i >= len(layouts or []):
                break
            ensure_beam_suple_superior(beam)
            if not beam_suple_sup_enabled(beam):
                continue
            left, width = self._elev_beam_full_span_px(
                beam, layouts[i], content_w, session
            )
            if width < 4.0:
                continue
            y = self._suple_y_sup_for_beam(beam)
            bar_top = y - 5.0
            for side in (u"start", u"end"):
                if not beam_suple_sup_side_enabled(beam, side):
                    continue
                try:
                    aid = (
                        beam.get(u"colStart")
                        if side == u"start"
                        else beam.get(u"colEnd")
                    )
                    _, diam_z = get_apoyo_suple_sup_arm(session, aid)
                except Exception:
                    try:
                        diam_z = int(beam.get("diamSupleSup") or 16)
                    except Exception:
                        diam_z = 16
                side_seg = self._elev_suple_sup_side_seg(i, side, left, width)
                try:
                    x0_z = float(side_seg["x0"])
                    x1_z = float(side_seg["x1"])
                except Exception:
                    continue
                zx0, zx1, _, _, _, _ = self._elev_suple_sup_apply_real_ends_px(
                    x0_z, x1_z, beams, side_seg, session, diam_z
                )
                zw = max(2.0, float(zx1) - float(zx0))
                self._elev_rect(
                    cnv, zx0, bar_top, zw, 12.0,
                    fill=zone_fill, stroke=zone_stroke, thickness=0.55,
                    dash=dash, zindex=3,
                )

        segs = suple_sup_segments_layout_px(
            beams, layouts, content_w, lay.pct_to_px, session=session
        )
        best_lbl = None
        best_w = 0.0
        for seg in segs:
            try:
                x0_raw = float(seg.get("x0") or 0)
                x1_raw = float(seg.get("x1") or 0)
            except Exception:
                continue
            if x1_raw - x0_raw < 2.0:
                continue
            idxs = list(seg.get("indices") or [])
            if not idxs or idxs[0] >= len(beams):
                continue
            ref = beams[idxs[0]]
            ensure_beam_suple_superior(ref)
            try:
                n_bars, diam = resolve_suple_sup_arm_for_spec(
                    session, seg, fallback_beam=ref
                )
            except Exception:
                try:
                    diam = int(ref.get("diamSupleSup") or 16)
                except Exception:
                    diam = 16
                try:
                    n_bars = int(ref.get("nSupleSup") or 2)
                except Exception:
                    n_bars = 2
            y = self._suple_y_sup_for_beam(ref)
            if len(idxs) >= 2 and idxs[1] < len(beams):
                y = 0.5 * (
                    self._suple_y_sup_for_beam(beams[idxs[0]])
                    + self._suple_y_sup_for_beam(beams[idxs[1]])
                )
            merged = bool(seg.get("merged"))
            x0, x1, ext_s, ext_e, emp_s, emp_e = (
                self._elev_suple_sup_apply_real_ends_px(
                    x0_raw, x1_raw, beams, seg, session, diam
                )
            )
            if x1 - x0 < 2.0:
                continue
            stroke = suple_merged if merged else suple_stroke
            sw = self._elev_suple_stroke_w(diam)
            if merged:
                sw = min(2.1, sw + 0.25)
            # Resalte: trazo del apoyo en edición (selected_suple_apoyo_id).
            try:
                seg_aid = seg.get(u"apoyo_id")
                ssel = getattr(session, u"selected_suple_apoyo_id", None)
                is_sel_seg = bool(
                    seg_aid
                    and ssel
                    and unicode(seg_aid) == unicode(ssel)
                )
            except Exception:
                is_sel_seg = False
            if is_sel_seg:
                sw = min(2.6, float(sw) + 0.55)
                stroke = brush_hex(u"#e9d5ff", 250)
            self._elev_line(cnv, x0, y, x1, y, stroke, sw, zindex=7)

            # Emp. muro // (sin pata L): trazo horizontal ancla → tip.
            if float(emp_s or 0.0) > 0.5:
                x_anch = x0_raw - max(0.0, float(ext_s or 0.0))
                self._draw_elev_emp_mark(cnv, x_anch, x0, y, stroke)
            if float(emp_e or 0.0) > 0.5:
                x_anch = x1_raw + max(0.0, float(ext_e or 0.0))
                self._draw_elev_emp_mark(cnv, x_anch, x1, y, stroke)

            # Pata L: estirón no// (muro/viga/columna); no si hay emp. //.
            pata_s = float(emp_s or 0.0) <= 0.5 and float(ext_s or 0.0) > 0.5
            pata_e = float(emp_e or 0.0) <= 0.5 and float(ext_e or 0.0) > 0.5
            if pata_s:
                self._draw_elev_pata_l_mark(
                    cnv, x0, y, u"start", u"sup", stroke,
                    diam_mm=diam, session=session,
                )
            if pata_e:
                self._draw_elev_pata_l_mark(
                    cnv, x1, y, u"end", u"sup", stroke,
                    diam_mm=diam, session=session,
                )

            if merged and seg.get("junctionX") is not None:
                try:
                    jx = float(seg["junctionX"])
                except Exception:
                    jx = None
                if jx is not None:
                    self._elev_line(
                        cnv, jx, y - 5.0, jx, y + 5.0,
                        brush_hex(u"#fbbf24", 180), 0.9,
                        dash=[2.0, 2.0], zindex=6,
                    )
            seg_w = x1 - x0
            if seg_w >= best_w:
                best_w = seg_w
                layer = beam_suple_sup_layer_index(ref)
                aid = seg.get("apoyo_id") or u""
                has_ext = (
                    float(ext_s or 0.0) > 0.5 or float(ext_e or 0.0) > 0.5
                )
                has_emp = (
                    float(emp_s or 0.0) > 0.5 or float(emp_e or 0.0) > 0.5
                )
                best_lbl = (
                    0.5 * (x0 + x1),
                    y,
                    n_bars,
                    diam,
                    layer,
                    stroke,
                    aid,
                    has_ext,
                    has_emp,
                )

        if best_lbl is not None:
            mid, y, n_bars, diam, layer, stroke, aid, has_ext, has_emp = (
                best_lbl
            )
            lbl = TextBlock()
            if has_emp and has_ext:
                base = u"Suple L/3+est.+emp"
            elif has_emp:
                base = u"Suple L/3+emp"
            elif has_ext:
                base = u"Suple L/3+estirón"
            else:
                base = u"Suple L/3"
            lbl.Text = u"{0} · {1}ø{2}".format(base, int(n_bars), int(diam))
            if aid:
                lbl.Text = u"{0} @{1} · {2}ø{3}".format(
                    base, aid, int(n_bars), int(diam),
                )
            lbl.FontSize = typo.META_FONT_PX
            lbl.FontWeight = FontWeights.SemiBold
            lbl.Foreground = stroke
            Canvas.SetLeft(lbl, mid - 52.0)
            Canvas.SetTop(lbl, y - 12.0)
            Canvas.SetZIndex(lbl, 8)
            cnv.Children.Add(lbl)

    def _draw_elevation_suple_inferior_tramos(self, cnv, beams, layouts, content_w):
        """Zonas 10 % extremos + fibra suple INF central 80 % por viga."""
        from armado_vigas.domain.suple_inferior import (
            SUPLE_TRIM_PCT_EACH_END,
            SUPLE_SPAN_PCT,
        )

        trim_fill = SolidColorBrush(Color.FromArgb(28, 251, 191, 36))
        trim_stroke = brush_hex(u"#fbbf24", 95)
        suple_fill = SolidColorBrush(Color.FromArgb(16, 192, 132, 252))
        suple_stroke = th.brush_sem(th.SEM_SUPLE, 230)
        dash = [3.0, 2.0]
        beams = beams or []
        session = getattr(self, u"_last_session", None)

        for i, beam in enumerate(beams):
            if i >= len(layouts or []):
                break
            ensure_beam_suple_inferior(beam)
            if not beam_suple_inf_enabled(beam):
                continue
            left, width = self._elev_beam_full_span_px(
                beam, layouts[i], content_w, session
            )
            if width < 4.0:
                continue
            trim_pct = float(SUPLE_TRIM_PCT_EACH_END)
            trim_w = width * trim_pct
            sup_x0 = left + trim_w
            sup_x1 = left + width - trim_w
            if (sup_x1 - sup_x0) < width * float(SUPLE_SPAN_PCT) * 0.5:
                mid = left + width * 0.5
                half = width * float(SUPLE_SPAN_PCT) * 0.5
                sup_x0, sup_x1 = mid - half, mid + half
            y = self._suple_y_inf_for_beam(beam)
            # Zona guía fina centrada en la fibra (siempre dentro vía y clamp).
            zone_h = 6.0
            bar_top = y - zone_h * 0.5
            try:
                diam = int(beam.get("diamSupleInf") or 16)
            except Exception:
                diam = 16
            try:
                n_bars = int(beam.get("nSupleInf") or 2)
            except Exception:
                n_bars = 2
            sw = self._elev_suple_stroke_w(diam)
            tw = max(2.0, trim_w)

            self._elev_rect(
                cnv, left, bar_top, tw, zone_h,
                fill=trim_fill, stroke=trim_stroke, thickness=0.55,
                dash=dash, zindex=3,
            )
            self._elev_rect(
                cnv, left + width - tw, bar_top, tw, zone_h,
                fill=trim_fill, stroke=trim_stroke, thickness=0.55,
                dash=dash, zindex=3,
            )
            self._elev_rect(
                cnv, sup_x0, bar_top, max(2.0, sup_x1 - sup_x0), zone_h,
                fill=suple_fill, stroke=th.brush_sem(th.SEM_SUPLE, 70),
                thickness=0.55, zindex=3,
            )
            self._elev_line(cnv, sup_x0, y, sup_x1, y, suple_stroke, sw, zindex=7)

            layer = beam_suple_layer_index(beam)
            mid = 0.5 * (sup_x0 + sup_x1)
            lbl = TextBlock()
            lbl.Text = u"Suple inf · {0}ø{1} · c{2}".format(
                int(n_bars), int(diam), int(layer),
            )
            lbl.FontSize = typo.META_FONT_PX
            lbl.FontWeight = FontWeights.SemiBold
            lbl.Foreground = suple_stroke
            Canvas.SetLeft(lbl, mid - 42.0)
            # Etiqueta justo sobre la fibra (hacia interior), no bajo el borde.
            Canvas.SetTop(lbl, y - 12.0)
            Canvas.SetZIndex(lbl, 8)
            cnv.Children.Add(lbl)

            try:
                len_mm = int(round(float(beam.get("len") or 0) * 1000.0))
            except Exception:
                len_mm = 0
            m = suple_metrics_mm(len_mm)
            # Cota ligera bajo la fibra pero short; no es la barra de armado.
            dim_y = y + 8.0
            self._elev_line(
                cnv, sup_x0, dim_y, sup_x1, dim_y,
                brush_hex(u"#a78bfa", 120), 0.7, zindex=5,
            )
            dim_lbl = TextBlock()
            dim_lbl.Text = u"{0} mm".format(int(m.get("spanMm") or 0))
            dim_lbl.FontSize = typo.META_FONT_PX
            dim_lbl.Foreground = brush_hex(u"#a78bfa", 170)
            Canvas.SetLeft(dim_lbl, mid - 18.0)
            Canvas.SetTop(dim_lbl, dim_y + 1.0)
            Canvas.SetZIndex(dim_lbl, 7)
            cnv.Children.Add(dim_lbl)

    def _elev_active_hint(self):
        card = self._active_rail_card()
        if card == u"lat":
            return u"LAT · global · sin selección de vigas/Tn en alzado"
        parts = []
        for face, label in ((u"sup", u"sup"), (u"inf", u"inf")):
            if not self._tramo_beam_selection_allowed(face):
                continue
            ids = sorted(self._selected_tramo_ids(face))
            if not ids:
                continue
            if len(ids) == 1:
                parts.append(u"T{0} {1}".format(ids[0], label))
            else:
                parts.append(
                    u"{0} {1}".format(
                        u"+".join(u"T{0}".format(i) for i in ids),
                        label,
                    )
                )
        if card == u"sup":
            base = u"SUP · clic bandas Tn sup · apoyos (suple L/3)"
            return (base + u" · " + u" · ".join(parts)) if parts else base
        if card == u"conf":
            n_v = 0
            try:
                n_v = len(self.selected_beam_indices or set())
            except Exception:
                n_v = 0
            if n_v > 1:
                base = u"CONF · {0} vigas · Ctrl+clic multi · ø/@ y E en lote".format(n_v)
            else:
                base = u"CONF · clic viga/zona · Ctrl+clic multi"
            return (base + u" · " + u" · ".join(parts)) if parts else base
        return u"Tramo: " + u" · ".join(parts) if parts else u"Sin tramo seleccionado"

    def _build_elev_stage_option_d(
        self, beams, layouts, tramos_sup, tramos_inf, session, apoyos_loaded, content_w,
    ):
        """Opción D: empalme por viga (sobre/bajo silueta) + bandas Tn + alzado."""
        stage = StackPanel()
        stage.Width = content_w

        stage.Children.Add(self._build_elev_header())

        # SUP: bandas Tn (resultado) → pills empalme (definición) → silueta
        stage.Children.Add(
            self._build_tramo_bands_ctrl_row(tramos_sup, beams, layouts, content_w, u"sup")
        )
        stage.Children.Add(
            self._build_empalme_pills_row(beams, layouts, session, content_w, u"sup")
        )
        stage.Children.Add(
            self._build_elevation_canvas(
                beams, layouts, tramos_sup, tramos_inf, session, apoyos_loaded, content_w,
            )
        )
        # INF: silueta → pills empalme → bandas Tn
        stage.Children.Add(
            self._build_empalme_pills_row(beams, layouts, session, content_w, u"inf")
        )
        stage.Children.Add(
            self._build_tramo_bands_ctrl_row(tramos_inf, beams, layouts, content_w, u"inf")
        )
        return stage

    def _tramo_band_owner_and_capas(self, beams, tramo, face):
        """Viga owner + nº de capas de la cara para el pill/banda Tn."""
        owner = None
        try:
            owner = self._sup_tramo_owner_beam(beams, tramo)
        except Exception:
            owner = None
        if owner is not None:
            try:
                ensure_beam_layers(owner)
            except Exception:
                pass
        n_capas = 1
        try:
            if owner is not None:
                if face == u"inf":
                    n_capas = int(beam_n_capas_inf(owner) or 1)
                else:
                    n_capas = int(beam_n_capas_sup(owner) or 1)
        except Exception:
            n_capas = 1
        return owner, max(1, min(int(lay.TRAMO_BAND_MAX_CAPAS), n_capas))

    def _tramo_band_layer_specs(self, tramo, face, beams):
        """Lista de (layer, n_bars, diam) para el chip multi-línea."""
        session = getattr(self, u"_last_session", None)
        owner, n_capas = self._tramo_band_owner_and_capas(beams, tramo, face)
        es_inf = face == u"inf"
        specs = []
        try:
            from armado_vigas.domain.tramo_armado import (
                tramo_layer_bar_count,
                tramo_layer_diam,
            )
        except Exception:
            tramo_layer_bar_count = None
            tramo_layer_diam = None
        for layer in range(1, n_capas + 1):
            n_bars = 2
            diam = 16
            if tramo_layer_bar_count is not None:
                try:
                    n_bars = int(
                        tramo_layer_bar_count(
                            session, face, tramo, layer, es_inf, owner
                        )
                        or 2
                    )
                except Exception:
                    n_bars = 2
                try:
                    diam = int(
                        tramo_layer_diam(
                            session, face, tramo, layer, es_inf, owner
                        )
                        or 16
                    )
                except Exception:
                    diam = 16
            specs.append((layer, n_bars, diam))
        return specs

    def _build_band_pill_label(
        self, tramo, accent, selected, selectable=True, face=None, beams=None,
    ):
        """Pill multi-línea (opción C): título Tn[·tras] + filas n×ø por capa."""
        face = face or (u"inf" if tramo.get(u"es_cara_inferior") else u"sup")
        suffix = u" ·tras" if tramo.get("fromEmpalme") else u""
        specs = self._tramo_band_layer_specs(tramo, face, beams or [])
        n_capas = max(1, len(specs))
        multi = n_capas > 1

        pill = Border()
        try:
            from System.Windows import CornerRadius

            pill.CornerRadius = CornerRadius(4.0)
        except Exception:
            pass
        # Padding vertical alineado a layout (no comprimir texto de capas).
        pad_v = max(4.0, float(lay.TRAMO_BAND_PILL_PAD_V_PX) * 0.5)
        pill.Padding = Thickness(7, pad_v - 1.0, 7, pad_v)
        pill.MinHeight = lay.tramo_band_pill_height_px(n_capas)
        try:
            pill.ClipToBounds = False
        except Exception:
            pass
        pill.VerticalAlignment = VerticalAlignment.Center
        pill.HorizontalAlignment = HorizontalAlignment.Center
        pill.BorderThickness = Thickness(1)
        if selected:
            pill.Background = brush_hex(accent, 255)
            pill.BorderBrush = brush_hex(accent, 255)
            fg = brush_hex(u"#0b1220", 255)
        else:
            pill.Background = brush_hex(u"#0b1624", 235)
            pill.BorderBrush = brush_hex(accent, 210)
            fg = brush_hex(u"#f8fafc", 245)

        stack = StackPanel()
        stack.Orientation = Orientation.Vertical
        stack.HorizontalAlignment = HorizontalAlignment.Center
        try:
            stack.ClipToBounds = False
        except Exception:
            pass

        title = TextBlock()
        title.Text = u"T{0}{1}".format(tramo.get("id"), suffix)
        title.FontSize = lay.TRAMO_BAND_PILL_TITLE_FONT_PX
        title.FontWeight = FontWeights.Bold
        title.Foreground = fg
        title.TextAlignment = TextAlignment.Center
        title.HorizontalAlignment = HorizontalAlignment.Center
        title.Height = lay.TRAMO_BAND_PILL_TITLE_H_PX
        title.Margin = Thickness(0, 0, 0, lay.TRAMO_BAND_PILL_TITLE_GAP_PX)
        try:
            title.LineHeight = lay.TRAMO_BAND_PILL_TITLE_H_PX
        except Exception:
            pass
        stack.Children.Add(title)

        for layer, n_bars, diam in specs:
            row = TextBlock()
            if multi:
                row.Text = u"C{0} {1}ø{2}".format(int(layer), int(n_bars), int(diam))
            else:
                row.Text = u"{0}ø{1}".format(int(n_bars), int(diam))
            row.FontSize = lay.TRAMO_BAND_PILL_LAYER_FONT_PX
            row.FontWeight = FontWeights.SemiBold
            row.Foreground = fg
            row.TextAlignment = TextAlignment.Center
            row.HorizontalAlignment = HorizontalAlignment.Center
            row.Height = lay.TRAMO_BAND_PILL_LAYER_H_PX
            try:
                row.LineHeight = lay.TRAMO_BAND_PILL_LAYER_H_PX
            except Exception:
                pass
            stack.Children.Add(row)

        pill.Child = stack
        try:
            tip_parts = [
                u"{0}ø{1}".format(n, d) if not multi else u"C{0} {1}ø{2}".format(i, n, d)
                for i, n, d in specs
            ]
            tip_body = u" + ".join(tip_parts) if tip_parts else u"—"
            if selectable:
                pill.ToolTip = (
                    u"T{0} · {1} · clic: seleccionar · Ctrl+clic: multi".format(
                        tramo.get("id"), tip_body
                    )
                )
            else:
                pill.ToolTip = u"T{0} · {1} · no seleccionable en esta pestaña".format(
                    tramo.get("id"), tip_body
                )
        except Exception:
            pass
        return pill

    def _build_band_trazo(self, accent, selected, width):
        """Trazo fino a lo largo del tramo."""
        line = Border()
        line.Height = (
            lay.TRAMO_BAND_TRAZO_H_SEL_PX if selected else lay.TRAMO_BAND_TRAZO_H_PX
        )
        line.Width = max(4.0, float(width))
        line.Background = brush_hex(accent, 240 if selected else 170)
        line.BorderThickness = Thickness(0)
        line.HorizontalAlignment = HorizontalAlignment.Stretch
        line.VerticalAlignment = VerticalAlignment.Center
        try:
            line.SnapsToDevicePixels = True
        except Exception:
            pass
        return line

    def _build_tramo_band_cell(
        self,
        tramo,
        beams,
        face,
        accent,
        selected,
        cell_w=None,
        selectable=True,
        row_n_capas=1,
    ):
        """Trazo fino + pill multi-línea Tn (opción C)."""
        is_sup = face == u"sup"
        n_capas = max(1, int(row_n_capas or 1))
        total_h = lay.tramo_band_cell_height_px(False, n_capas)
        pill_h = lay.tramo_band_pill_height_px(n_capas)
        trazo_h = max(float(lay.TRAMO_BAND_TRAZO_SLOT_PX), total_h - pill_h)
        w = max(4.0, float(cell_w or 40.0))

        # Hit target transparente (sin relleno/borde de panel).
        cell = Border()
        cell.Margin = Thickness(0)
        cell.Height = total_h
        cell.Width = w
        cell.Background = brush_hex(u"#000000", 1)
        cell.BorderThickness = Thickness(0)
        try:
            cell.ClipToBounds = False
            cell.Cursor = Cursors.Hand if selectable else Cursors.Arrow
        except Exception:
            try:
                cell.Cursor = Cursors.Hand if selectable else Cursors.Arrow
            except Exception:
                cell.Cursor = Cursors.Hand if selectable else None
        cell.VerticalAlignment = VerticalAlignment.Top
        cell.HorizontalAlignment = HorizontalAlignment.Left

        # Stack natural: el pill no se mide a la baja dentro de una fila fija.
        shell = StackPanel()
        shell.Width = w
        shell.Orientation = Orientation.Vertical
        try:
            shell.ClipToBounds = False
        except Exception:
            pass

        pill = self._build_band_pill_label(
            tramo, accent, selected, selectable=selectable, face=face, beams=beams,
        )
        pill.HorizontalAlignment = HorizontalAlignment.Center
        pill.VerticalAlignment = VerticalAlignment.Center
        # Altura reservada = fila de capas de la cara (alineación de trazos).
        pill.MinHeight = pill_h

        trazo = self._build_band_trazo(accent, selected, w)
        trazo.VerticalAlignment = VerticalAlignment.Center
        trazo.Margin = Thickness(0, 1 if is_sup else 0, 0, 0 if is_sup else 1)

        if is_sup:
            # SUP: pill arriba, trazo al pie (hacia la viga).
            pill_slot = Border()
            pill_slot.Height = pill_h
            pill_slot.Width = w
            pill_slot.Background = brush_hex(u"#000000", 0)
            try:
                pill_slot.ClipToBounds = False
            except Exception:
                pass
            pill_slot.Child = pill
            shell.Children.Add(pill_slot)
            trazo_slot = Border()
            trazo_slot.Height = trazo_h
            trazo_slot.Width = w
            trazo_slot.Child = trazo
            trazo.VerticalAlignment = VerticalAlignment.Bottom
            trazo_slot.VerticalAlignment = VerticalAlignment.Bottom
            shell.Children.Add(trazo_slot)
        else:
            # INF: trazo arriba, pill abajo.
            trazo_slot = Border()
            trazo_slot.Height = trazo_h
            trazo_slot.Width = w
            trazo.VerticalAlignment = VerticalAlignment.Top
            trazo_slot.Child = trazo
            shell.Children.Add(trazo_slot)
            pill_slot = Border()
            pill_slot.Height = pill_h
            pill_slot.Width = w
            pill_slot.Background = brush_hex(u"#000000", 0)
            try:
                pill_slot.ClipToBounds = False
            except Exception:
                pass
            pill.VerticalAlignment = VerticalAlignment.Center
            pill_slot.Child = pill
            shell.Children.Add(pill_slot)

        cell.Child = shell
        return cell

    def _build_tramo_bands_ctrl_row(self, tramos, beams, layouts, content_w, face):
        """Bandas Tn (opción C) en el mismo eje X que la silueta de viga."""
        is_sup = face == u"sup"
        sel_ids = self._selected_tramo_ids(face)
        accent_default = u"#22d3ee" if is_sup else u"#fb7185"
        tramos = tramos or []
        can_sel = self._tramo_beam_selection_allowed(face)

        # Alto de fila = peor caso de capas de la cara (alineación del trazo).
        row_n_capas = 1
        for tramo in tramos:
            try:
                _, nc = self._tramo_band_owner_and_capas(beams, tramo, face)
                if nc > row_n_capas:
                    row_n_capas = nc
            except Exception:
                pass

        band_h = lay.tramo_band_cell_height_px(False, row_n_capas)
        wrap = Border()
        wrap.Margin = Thickness(0, 1 if is_sup else 0, 0, 0 if is_sup else 1)
        wrap.Width = content_w
        wrap.Background = brush_hex(u"#000000", 0)

        host = Canvas()
        host.Width = content_w
        host.Height = band_h
        host.ClipToBounds = False

        for tramo in tramos:
            span = lay.tramo_span(layouts, tramo, content_w)
            left = lay.pct_to_px(span["leftPct"], content_w)
            width = max(4.0, lay.pct_to_px(span["widthPct"], content_w))
            gap = 1.0
            cell_left = left + gap * 0.5
            cell_w = max(4.0, width - gap)

            accent = tramo.get("accent") or accent_default
            # Resalte de selección solo si la pestaña admite elegir tramos.
            sel = bool(can_sel and (tramo["id"] in sel_ids))
            cell = self._build_tramo_band_cell(
                tramo,
                beams,
                face,
                accent,
                sel,
                cell_w=cell_w,
                selectable=can_sel,
                row_n_capas=row_n_capas,
            )
            Canvas.SetLeft(cell, cell_left)
            Canvas.SetTop(cell, 0.0)

            if can_sel:
                self._wire_tramo_select_click(cell, face, tramo["id"])
            host.Children.Add(cell)

        wrap.Child = host
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
            def _on(v, f=face_key):
                # Alinear estado al valor del toggle
                now = bid in (
                    (session.empalme_beam_ids_sup if f == u"sup" else session.empalme_beam_ids_inf)
                    or set()
                )
                if bool(v) != bool(now):
                    self._cb.get("on_toggle_empalme", lambda _b, _f: None)(bid, f)
                self._cb.get("on_select_beam", lambda _i: None)(idx)
                self._cb.get("on_redraw", lambda: None)()

            row = StackPanel()
            row.Orientation = Orientation.Horizontal
            row.Margin = Thickness(0, 0, 4, 0)
            tog = make_yesno_toggle(
                self._win, active, _on, compact=True, label=label,
            )
            row.Children.Add(tog)
            return row

        btn_sup = _tras_btn(u"Superior", on_sup, u"sup", u"#22d3ee")
        btn_inf = _tras_btn(u"Inferior", on_inf, u"inf", u"#fb7185")
        Grid.SetColumn(btn_sup, 0)
        Grid.SetColumn(btn_inf, 1)
        row.Children.Add(btn_sup)
        row.Children.Add(btn_inf)
        sp.Children.Add(row)
        block.Child = sp
        return block

    def _build_elevation_canvas(self, beams, layouts, tramos_sup, tramos_inf, session, apoyos_loaded, content_w):
        elev_border = Border()
        elev_border.BorderBrush = brush_hex(u"#21465C", 115)
        elev_border.BorderThickness = Thickness(0, 1, 0, 1)
        # Sin padding horizontal: el dibujo usa todo content_w; pad horizontal
        # recortaba el extremo derecho (ClipToBounds del Border).
        elev_border.Padding = Thickness(0, 2, 0, 2)
        elev_border.Background = th.brush_panel(0)
        try:
            elev_border.ClipToBounds = False
        except Exception:
            pass

        elev_h = float((self._layout_meta or {}).get("elevHeightPx") or lay.ELEVATION_HEIGHT_PX)
        cnv = Canvas()
        cnv.Width = content_w
        cnv.Height = elev_h
        cnv.Background = brush_hex(u"#0a1620", 0)
        # False: el borde derecho del contorno (stroke centrado) no se recorta.
        cnv.ClipToBounds = False
        apply_aliased_render(cnv)

        # Alinea fibras SUP/INF al alto de las siluetas (preview de barras).
        self._elev_sync_bar_rows(beams)

        has_apoyos = bool(getattr(session, "apoyos", None))
        chain = (
            lay.build_support_chain(
                beams,
                layouts,
                apoyos=getattr(session, "apoyos", None),
                layout_meta=self._layout_meta,
            )
            if has_apoyos
            else []
        )
        zones = self._elev_support_zones(chain, content_w, session) if has_apoyos else []

        # Batch StreamGeometry: un Path por estilo (vs. 1 Line/Rect por primitiva).
        # Coordenadas y reglas de negocio no cambian — solo el cruce Python↔.NET.
        batch = ElevGeomBatch()
        prev_batch = getattr(self, u"_elev_geom_batch", None)
        self._elev_geom_batch = batch
        try:
            # Silueta de hormigón.
            if layouts:
                for i, beam in enumerate(beams):
                    lay_i = layouts[i]
                    fx, fw = self._elev_beam_full_span_px(beam, lay_i, content_w, session)
                    top, h_px = self._elev_beam_vertical(beam)
                    sel = self._is_beam_selected_for_elev(i)
                    self._draw_elevation_beam_fill(cnv, fx, fw, top, h_px, selected=sel)
                    self._draw_elevation_beam_edges(
                        cnv, fx, fw, top, h_px, zones, selected=sel,
                    )
                    self._draw_elevation_beam_section_label(
                        cnv, fx, fw, top, beam, selected=sel,
                    )

            # Vigas unidas detectadas (// y no // a la vista activa).
            self._draw_elevation_joined_beams(cnv, session, content_w)

            # Losas seleccionadas (contexto alzado; no entran en cadena de extremos).
            self._draw_elevation_floors(cnv, session, content_w)

            for j, pt in enumerate(chain):
                x = lay.pct_to_px(pt["pct"], content_w)
                self._draw_elevation_support(cnv, x, pt, j, session)

            # Confinamiento / estribos Ext·Cent (preview en alzado).
            self._draw_elevation_confinement(cnv, beams, layouts, content_w, session)

            # Preview armadura longitudinal SUP (por Tn y capas) — análisis pre-modelado.
            self._draw_elevation_top_bars(
                cnv, beams, layouts, tramos_sup, content_w, session=session
            )
            # Preview INF: mismas reglas post-fusión (estirón / pata L / emp. muro //).
            self._draw_elevation_bottom_bars(
                cnv, beams, layouts, tramos_inf, content_w, session=session
            )
            # Suples: capa n+1 — SUP extremos 25 % / fusión; INF central 80 %.
            self._draw_elevation_suple_superior_tramos(cnv, beams, layouts, content_w)
            self._draw_elevation_suple_inferior_tramos(cnv, beams, layouts, content_w)

            for i, beam in enumerate(beams):
                lay_i = layouts[i]
                # Alinear marcador con silueta dibujada (span completo del hormigón).
                fx, fw = self._elev_beam_full_span_px(beam, lay_i, content_w, session)
                top, h_px = self._elev_beam_vertical(beam)
                y_under = top + h_px + _ELEV_DIR_BELOW_GAP_PX
                self._draw_elevation_direction_marker(
                    cnv,
                    fx,
                    fw,
                    i,
                    bool(beam.get("axisReversed")),
                    y_mid=y_under,
                    badge_below=True,
                )

                span = Border()
                span.Width = fw
                span.Height = h_px
                Canvas.SetLeft(span, fx)
                Canvas.SetTop(span, top)
                span.Background = brush_hex(u"#000000", 0)
                # Sin borde de selección: un anillo cyan se leía como “contorno raro”.
                span.BorderBrush = brush_hex(u"#000000", 0)
                span.BorderThickness = Thickness(0)
                Canvas.SetZIndex(span, 10)

                # Hit de viga solo en INF / CONF.
                if self._elev_beam_selection_ui_enabled():
                    span.Cursor = Cursors.Hand
                    idx_cap = i
                    beam_cap = beam

                    def _click_beam(sender, args, idx=idx_cap, b=beam_cap):
                        self._select_beam_from_elevation(idx, b, args)

                    try:
                        span.MouseLeftButtonUp += MouseButtonEventHandler(_click_beam)
                    except Exception:
                        pass
                    cnv.Children.Add(span)
                # En SUP/LAT: sin hit de viga (apoyos / flujo global siguen activos).
        finally:
            try:
                batch.flush(cnv)
            except Exception:
                pass
            self._elev_geom_batch = prev_batch

        elev_border.Child = cnv
        return elev_border

    def _select_beam_from_elevation(self, idx, beam, args=None):
        """Selecciona viga del alzado sin abrir ni enfocar la card CONF."""
        if not self._elev_beam_selection_ui_enabled():
            return
        # Sin rol de estribo: no asociar clic de viga/tramo a Ext/Cent/Uni.
        self._handle_beam_select(idx, args, role=None, update_zone=False)

    def _build_labels(self, beams, layouts, session, apoyos_loaded, content_w):
        cnv = Canvas()
        cnv.Width = content_w
        cnv.Height = lay.LABELS_HEIGHT_PX
        beam_hit = self._elev_beam_selection_ui_enabled()
        for i, beam in enumerate(beams):
            lay_i = layouts[i]
            cx = lay.pct_to_px(lay_i["centerPct"], content_w)
            tb = TextBlock()
            tb.TextAlignment = TextAlignment.Center
            tb.Foreground = (
                brush_hex(u"#7eb8d0")
                if self._is_beam_selected_for_elev(i)
                else brush_hex(u"#95B8CC")
            )
            tb.Cursor = Cursors.Hand if beam_hit else Cursors.Arrow
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
            if beam_hit:
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
        # Panel legacy: solo lectura del label del dibujo
        from armado_vigas.domain.confinement import conf_draft_label, get_conf_draft

        d = get_conf_draft(beam)
        lbl = conf_draft_label(d)
        combo = make_string_combo(
            self._win,
            [lbl] if lbl else ([opts[0]] if opts else [u"Dibujo libre"]),
            lbl or (opts[0] if opts else u"Dibujo libre"),
            lambda v, b=beam: None,
            compact=True,
        )
        try:
            combo.IsEnabled = False
        except Exception:
            pass
        combo.Height = typo.CTRL_HEIGHT_PX
        combo.FontSize = typo.CTRL_FONT_PX
        combo.HorizontalAlignment = HorizontalAlignment.Stretch
        Grid.SetColumn(combo, 0)
        fields.Children.Add(combo)
        sp.Children.Add(fields)

        hint = TextBlock()
        hint.Text = u"Editar en pestaña CONF · preview de sección del rail (135°)"
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
        n_cb = make_int_combo(
            self._win,
            session_n_laterales(session, 0),
            LATERALES_COUNT_MIN,
            LATERALES_COUNT_MAX,
            lambda v, s=session: self._set_session_laterales_field(s, "nLaterales", v),
            compact=True,
            stretch=True,
        )
        Grid.SetColumn(n_cb, 1)
        fields.Children.Add(n_cb)

        o_lbl = self._section_field_lbl(u"ø")
        Grid.SetColumn(o_lbl, 2)
        fields.Children.Add(o_lbl)
        diam_lat = int(getattr(session, "diamLaterales", LATERALES_DIAM_DEFAULT) or LATERALES_DIAM_DEFAULT)
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
        if session is None:
            return set()
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

    def _build_empalme_define_pill(self, active, face, cell_w, beam_label=None):
        """Pill clicable: la viga define (o no) empalme @ mitad en esa cara.

        Activo → trocea la corrida y genera Tn; inactivo → barra continúa.
        """
        is_sup = face == u"sup"
        face_acc = u"#22d3ee" if is_sup else u"#fb7185"
        emp_acc = getattr(th, "SEM_EMPALME", None) or u"#fbbf24"
        w = max(24.0, float(cell_w or 40.0))
        short = w < 52.0

        if active:
            text = u"E" if short else u"Empalme"
            bg = brush_hex(emp_acc, 230)
            border = brush_hex(emp_acc, 255)
            fg = brush_hex(u"#1a1305", 255)
        else:
            text = u"·" if short else u"—"
            bg = brush_hex(u"#0b1624", 210)
            border = brush_hex(face_acc, 90)
            fg = brush_hex(u"#64748b", 230)

        pill = Border()
        try:
            from System.Windows import CornerRadius

            pill.CornerRadius = CornerRadius(4.0)
        except Exception:
            pass
        pill.Padding = Thickness(6 if not short else 4, 1, 6 if not short else 4, 1)
        pill.MinHeight = lay.TRAMO_EMPALME_PILL_H_PX
        pill.Height = lay.TRAMO_EMPALME_PILL_H_PX
        pill.MaxWidth = max(28.0, w - 2.0)
        try:
            pill.MinWidth = min(28.0 if short else 36.0, max(24.0, w - 4.0))
        except Exception:
            pass
        pill.VerticalAlignment = VerticalAlignment.Center
        pill.HorizontalAlignment = HorizontalAlignment.Center
        pill.BorderThickness = Thickness(1.25 if active else 1)
        pill.Background = bg
        pill.BorderBrush = border
        pill.Cursor = Cursors.Hand

        tb = TextBlock()
        tb.Text = text
        tb.FontSize = lay.TRAMO_EMPALME_PILL_FONT_PX
        tb.FontWeight = FontWeights.Bold
        tb.Foreground = fg
        tb.VerticalAlignment = VerticalAlignment.Center
        tb.TextAlignment = TextAlignment.Center
        try:
            tb.HorizontalAlignment = HorizontalAlignment.Center
        except Exception:
            pass
        pill.Child = tb

        try:
            face_lbl = u"SUP" if is_sup else u"INF"
            state = u"define empalme @ mitad · trocea Tn" if active else u"sin empalme · corrida continua"
            bl = beam_label or u"viga"
            pill.ToolTip = u"{0} · {1} · {2}".format(bl, face_lbl, state)
        except Exception:
            pass
        return pill

    def _build_empalme_pills_row(self, beams, layouts, session, content_w, face):
        """Fila de pills de empalme alineadas al eje X de cada viga dibujada."""
        is_sup = face == u"sup"
        empalme_set = self._empalme_ids_for_face(session, face) if session is not None else set()
        row_h = float(lay.TRAMO_EMPALME_ROW_PX)

        wrap = Border()
        wrap.Margin = Thickness(0, 0 if is_sup else 1, 0, 1 if is_sup else 0)
        wrap.Width = content_w
        wrap.Background = brush_hex(u"#000000", 0)
        wrap.BorderThickness = Thickness(0)

        host = Canvas()
        host.Width = content_w
        host.Height = row_h
        host.ClipToBounds = False

        beams = beams or []
        layouts = layouts or []
        for i, beam in enumerate(beams):
            if i >= len(layouts):
                break
            lay_i = layouts[i]
            left = lay.pct_to_px(lay_i["leftPct"], content_w)
            width = max(4.0, lay.pct_to_px(lay_i["widthPct"], content_w))
            gap = 2.0
            cell_left = left + gap * 0.5
            cell_w = max(4.0, width - gap)
            beam_id = beam.get("id")
            active = beam_id in empalme_set
            try:
                beam_lbl = lay.beam_canvas_label(i)
            except Exception:
                beam_lbl = u"V{0}".format(i + 1)

            cell = Border()
            cell.Width = cell_w
            cell.Height = row_h
            cell.Background = brush_hex(u"#000000", 1)
            cell.BorderThickness = Thickness(0)
            cell.Cursor = Cursors.Hand

            pill = self._build_empalme_define_pill(active, face, cell_w, beam_label=beam_lbl)
            # Centrar pill en la celda (Canvas no centra hijos).
            holder = Grid()
            holder.Width = cell_w
            holder.Height = row_h
            pill.HorizontalAlignment = HorizontalAlignment.Center
            pill.VerticalAlignment = VerticalAlignment.Center
            holder.Children.Add(pill)
            cell.Child = holder

            idx_cap = i

            def _toggle(sender, args, bi=idx_cap, bid=beam_id, f=face):
                try:
                    if args is not None:
                        args.Handled = True
                except Exception:
                    pass
                # Selección de viga solo INF/CONF (en SUP no aplica).
                if self._tramo_beam_selection_allowed():
                    self._handle_beam_select(bi, update_zone=False, redraw=False)
                self._cb.get("on_toggle_empalme", lambda _b, _f: None)(bid, f)
                self._cb.get("on_redraw", lambda: None)()

            try:
                cell.MouseLeftButtonUp += MouseButtonEventHandler(_toggle)
            except Exception:
                pass

            Canvas.SetLeft(cell, cell_left)
            Canvas.SetTop(cell, 0.0)
            host.Children.Add(cell)

        wrap.Child = host
        return wrap

    def _build_empalme_button_row(self, beams, layouts, session, content_w, face):
        """Compat: fila de pills empalme (antes toggles Sí/No)."""
        row = self._build_empalme_pills_row(beams, layouts, session, content_w, face)
        try:
            # strip legacy expectaba canvas interno; devolver host si Border
            if hasattr(row, "Child") and row.Child is not None:
                return row.Child
        except Exception:
            pass
        return row

    def _build_traslape_strip(self, beams, layouts, session, content_w, face):
        strip = Border()
        strip.Margin = Thickness(lay.FACE_BLOCK_PAD_PX, 0, lay.FACE_BLOCK_PAD_PX, 0)
        strip.Padding = Thickness(0, 2, 0, 2)
        strip.Height = lay.TRAMO_EMPALME_ROW_PX + 4.0
        strip.Background = brush_hex(u"#071018", 128)
        strip.BorderBrush = brush_hex(u"#21465C", 102)
        strip.BorderThickness = Thickness(1)
        try:
            strip.CornerRadius = System.Windows.CornerRadius(4)
        except Exception:
            pass
        strip.Child = self._build_empalme_pills_row(
            beams, layouts, session, content_w, face,
        )
        return strip

    def _build_tramo_bands_row(self, tramos, layouts, content_w, face, accent_default):
        is_sup = face == u"sup"
        sel_ids = self._selected_tramo_ids(face)
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
            sel = tramo["id"] in sel_ids
            band.Background = brush_hex(accent, 200 if sel else 140)
            band.BorderBrush = brush_hex(accent if sel else accent, 220)
            band.BorderThickness = Thickness(2 if sel else 1)
            band.CornerRadius = System.Windows.CornerRadius(2)
            can_sel = self._tramo_beam_selection_allowed(face)
            try:
                band.Cursor = Cursors.Hand if can_sel else Cursors.Arrow
            except Exception:
                band.Cursor = Cursors.Hand
            if tramo_exceeds_bar_limit(tramo):
                band.BorderBrush = brush_hex(u"#fbbf24")
            if can_sel:
                self._wire_tramo_select_click(band, face, tramo["id"])
            bands.Children.Add(band)
        return bands

    def _build_panel_lane(self, tramos, beams, layouts, content_w, face):
        """Controladores Tn anclados bajo/en la **misma X** que cada banda.

        Antes: StackPanel en orden de id (scroll propio) → desfase visual vs
        bandas/siluetas y se editaba el panel de un extremo creyendo el otro.
        """
        is_sup = face == u"sup"
        sel_ids = self._selected_tramo_ids(face)
        accent_default = u"#22d3ee" if is_sup else u"#fb7185"

        wrap = Border()
        wrap.Margin = Thickness(0, 2, 0, 4)
        wrap.Padding = Thickness(0)
        wrap.BorderBrush = brush_hex(u"#21465C", 115)
        wrap.BorderThickness = Thickness(0, 1, 0, 0)
        wrap.Width = content_w

        host = Canvas()
        host.Width = content_w
        host.Height = lay.TRAMO_PANEL_LANE_PX + 8.0
        host.ClipToBounds = False

        for tramo in tramos or []:
            span = lay.tramo_span(layouts, tramo, content_w)
            left = lay.pct_to_px(span["leftPct"], content_w)
            width = max(8.0, lay.pct_to_px(span["widthPct"], content_w))
            # Ancho del panel: al menos legible, acotado al span y al máximo de diseño.
            pw = max(96.0, min(float(lay.TRAMO_PANEL_W_PX), width - 4.0))
            if pw > width and width >= 72.0:
                pw = max(72.0, width - 2.0)
            panel = self._build_face_tramo_panel(tramo, beams, layouts, content_w, face)
            panel.Width = pw
            panel.MinWidth = pw
            panel.MaxWidth = pw
            panel.ClipToBounds = True
            accent = tramo.get("accent") or accent_default
            if tramo["id"] in sel_ids:
                panel.BorderBrush = brush_hex(accent)
                panel.BorderThickness = Thickness(2)
            # Centrado bajo la banda del mismo Tn.
            px = left + max(0.0, (width - pw) * 0.5)
            Canvas.SetLeft(panel, px)
            Canvas.SetTop(panel, 4.0)
            host.Children.Add(panel)

        wrap.Child = host
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
        sel = self._is_tramo_selected(face, tramo["id"])
        accent = tramo.get("accent") or (u"#22d3ee" if is_sup else u"#fb7185")

        panel = Border()
        panel.Padding = Thickness(4, 4, 4, 4)
        panel.HorizontalAlignment = HorizontalAlignment.Center
        panel.BorderBrush = brush_hex(accent) if sel else brush_hex(u"#21465C")
        panel.BorderThickness = Thickness(2 if sel else 1)
        panel.Background = brush_hex(u"#071018")
        can_sel = self._tramo_beam_selection_allowed(face)
        try:
            panel.Cursor = Cursors.Hand if can_sel else Cursors.Arrow
        except Exception:
            panel.Cursor = Cursors.Hand
        try:
            if can_sel:
                panel.ToolTip = u"Clic: seleccionar · Ctrl+clic: multi-selección"
            else:
                panel.ToolTip = u"Esta banda no es seleccionable en la pestaña actual"
        except Exception:
            pass

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

        sp.Children.Add(self._cap_col(tramo_beams, owner, side, tramo=tramo, face=face))

        if warn:
            wtb = TextBlock()
            wtb.Text = u"Barra presunta > 12 m"
            wtb.Foreground = brush_hex(u"#fbbf24")
            wtb.FontSize = typo.META_FONT_PX
            wtb.FontWeight = FontWeights.Bold
            wtb.Margin = Thickness(0, 2, 0, 0)
            sp.Children.Add(wtb)

        panel.Child = sp
        if can_sel:
            self._wire_tramo_select_click(panel, face, tramo["id"])
        return panel

    def _tramo_beams(self, tramo, beams):
        out = []
        for idx in tramo.get("beamIndices") or []:
            if 0 <= int(idx) < len(beams):
                out.append(beams[int(idx)])
        return out

    # ø y n por tramo Tn (session.tramo_armado). Capas: lote completo.
    _TRAMO_LOCAL_FIELDS = (
        "diamSup",
        "diamInf",
        "diamSup2",
        "diamInf2",
        "diamSup3",
        "diamInf3",
        "nSup",
        "nInf",
        "nSup2",
        "nInf2",
        "nSup3",
        "nInf3",
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
        targets = self._targets_for_beam_edit(beam) or [beam]
        for b in targets:
            if b is None:
                continue
            b[field] = value
            if field.startswith("supleInf") or field == "supleInfEnabled":
                ensure_beam_suple_inferior(b)
                # Asegura bool real (IronPython / None residual).
                if field == "supleInfEnabled":
                    b[u"supleInfEnabled"] = bool(value)
            elif field.startswith("supleSup") or field == "supleSupEnabled":
                ensure_beam_suple_superior(b)
                if field == "supleSupEnabled":
                    b[u"supleSupEnabled"] = bool(value)
                    if value:
                        b["supleSupStartEnabled"] = True
                        b["supleSupEndEnabled"] = True
        try:
            # El fingerprint de alzado/rail depende de estos flags.
            self.invalidate_elev_cache()
        except Exception:
            self._elev_cache_fp = None
            self._rail_cache_fp = None
        self._cb.get("on_redraw", lambda: None)()

    # Campos de confinamiento que invalidan alzado al editarse (multi-sel OK).
    _CONF_PER_BEAM_FIELDS = frozenset((
        u"estConfDraft",
        u"estConfin",
        u"estExtDiam",
        u"estExtSpacing",
        u"estCentDiam",
        u"estCentSpacing",
        u"estZonasMode",
    ))

    def _set_beam_field(self, beam, field, value):
        if beam is None:
            return
        # En CONF (y resto), ø/espaciado/zonas/draft siguen la multi-sel de alzado.
        targets = self._targets_for_beam_edit(beam) or [beam]
        for b in targets:
            if field in ("nSup", "nInf"):
                set_first_layer_bar_count(b, value)
                ensure_beam_layers(b)
                ensure_beam_confinement(b)
            else:
                b[field] = value
                if field == "estConfin":
                    ensure_beam_confinement(b)
                elif field == "estConfDraft":
                    from armado_vigas.domain.confinement import set_conf_draft

                    set_conf_draft(b, value)
                elif field == "estZonasMode":
                    from armado_vigas.domain.stirrups import (
                        ensure_beam_stirrup_zone_mode,
                        normalize_stirrup_zone_mode,
                    )

                    b[field] = normalize_stirrup_zone_mode(value)
                    ensure_beam_stirrup_zone_mode(b)
        try:
            if field in self._CONF_PER_BEAM_FIELDS or field.startswith(u"est"):
                self.invalidate_elev_cache()
        except Exception:
            pass
        self._cb.get("on_redraw", lambda: None)()

    def _cap_col(
        self, tramo_beams, owner, side, tramo_accent=None, compact_band=False,
        tramo=None, face=None,
    ):
        is_sup = side == u"sup" or side == "sup"
        accent = tramo_accent or (u"#22d3ee" if is_sup else u"#f87171")
        face = face or (u"sup" if is_sup else u"inf")
        session = getattr(self, u"_last_session", None)

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
        cap_row.Children.Add(make_capas_combo(
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
            # ø/n desde cfg del tramo (no viga compartida).
            n_show = beam.get(qty_field) or BAR_COUNT_MIN
            d_show = int(beam.get(diam_field) or 16)
            if tramo is not None and session is not None:
                try:
                    from armado_vigas.domain.tramo_armado import (
                        owner_display_value,
                        resolve_first_layer_n_linked,
                    )

                    if layer_num == 1:
                        n_show = resolve_first_layer_n_linked(
                            session, face, tramo, owner
                        )
                    else:
                        n_raw = owner_display_value(
                            session, face, tramo, qty_field, owner
                        )
                        if n_raw is not None:
                            n_show = clamp_bar_count(n_raw)
                    d_raw = owner_display_value(
                        session, face, tramo, diam_field, owner
                    )
                    if d_raw is not None:
                        d_show = int(d_raw)
                except Exception:
                    pass
            row.Children.Add(make_bar_count_combo(
                self._win,
                n_show,
                lambda v, tb=tramo_beams, o=owner, f=qty_field, tr=tramo, fc=face: (
                    self._set_tramo_beams_field(
                        tb, f, clamp_bar_count(v),
                        confinement=(f in ("nSup", "nInf")),
                        owner=o,
                        target_tramo=tr,
                        face=fc,
                    )
                ),
                compact=True,
                enabled=active,
            ))
            row.Children.Add(make_diam_combo(
                self._win,
                d_show,
                _session_bar_diam_opts(session, d_show),
                lambda v, tb=tramo_beams, o=owner, f=diam_field, tr=tramo, fc=face: (
                    self._set_tramo_beams_field(
                        tb, f, v, owner=o, target_tramo=tr, face=fc,
                    )
                ),
                compact=True,
                enabled=active,
            ))
            if not compact_band and not active:
                row.Opacity = 0.55
            col.Children.Add(row)

        wrap.Child = col
        return wrap

    def _set_tramo_beams_field(
        self, tramo_beams, field, value, confinement=False, owner=None,
        target_tramo=None, face=None,
    ):
        """Aplica n/ø a tramo(s); desde panel de un Tn escribe solo ese Tn.

        Desde rail multi-selección sigue aplicando a la selección activa.
        """
        owner = owner or ((tramo_beams or [None])[0])
        all_beams = self._all_domain_beams(tramo_beams)
        if is_global_layer_sync_field(field):
            sync_layer_field_all_beams(all_beams, field, value)
        else:
            try:
                from armado_vigas.ui import rail_cards as _rc

                face = face or (
                    u"inf" if (field or u"").startswith((u"nInf", u"diamInf")) else u"sup"
                )
                session = getattr(self, u"_last_session", None)
                tramos = list(
                    getattr(session, u"tramos_inf" if face == u"inf" else u"tramos_sup", None)
                    or []
                )
                force = None
                if target_tramo is not None:
                    # Panel de un Tn: solo ese tramo (no toda la selección actual).
                    force = [target_tramo]
                _rc._apply_layer_field(
                    self, tramo_beams, owner, field, value,
                    face=face, tramos=tramos, force_tramos=force,
                )
                return
            except Exception:
                if owner is not None:
                    if field in (u"nSup", u"nInf"):
                        set_first_layer_bar_count(owner, value)
                    else:
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
