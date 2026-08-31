# -*- coding: utf-8 -*-
"""Laboratorio interactivo: dibujar estribos E(i–j) y trabas [k] en sección.

Prototipo de evaluación (no escribe a SESSION ni a Colocar).
Fuente de verdad del borrador: perimetral / pairs / ties indexados en 1ª capa.
"""

from __future__ import print_function

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")

from System.Windows import (
    FontWeights,
    GridLength,
    GridUnitType,
    HorizontalAlignment,
    TextAlignment,
    TextWrapping,
    Thickness,
    VerticalAlignment,
    Window,
    WindowStartupLocation,
)
from System.Windows.Controls import (
    Border,
    Button,
    Canvas,
    ColumnDefinition,
    ComboBox,
    ComboBoxItem,
    Grid,
    Orientation,
    ScrollViewer,
    StackPanel,
    TextBlock,
    TextBox,
)
from System.Windows.Input import (
    Key,
    KeyEventHandler,
    MouseButtonEventHandler,
    MouseEventHandler,
    Cursors,
)
from System.Windows.Media import Color, DoubleCollection, FontFamily, SolidColorBrush
from System.Windows.Shapes import Ellipse, Line, Rectangle

from armado_vigas.ui.section_preview import (
    _add_dot,
    _distribute_points,
    _fit_section_rect,
)
from armado_vigas.ui.stirrup_draw_135 import (
    COLOR_ESTRIBO,
    COLOR_TRABA,
    STIRRUP_THICK,
    TIE_THICK,
    draw_estribo_bbox,
    draw_traba_vertical_column,
    stirrup_margin_px,
)
from armado_vigas.ui.wpf_controls import brush_hex

# --- constantes de lab ---
_LAB_W = 440.0
_LAB_H = 420.0
_SECTION_B_CM = 25.0
_SECTION_H_CM = 50.0
_HIT_PAD = 12.0  # px snap alrededor de cada barra
_SNAP_R = 20.0
_PAIR_PAD = 5.0


def _pair_bounds(first_sup, first_inf, i0, i1, pad=_PAIR_PAD):
    a, b = min(int(i0), int(i1)), max(int(i0), int(i1))
    xs = [
        first_sup[a]["x"], first_sup[b]["x"],
        first_inf[a]["x"], first_inf[b]["x"],
    ]
    ys = [
        first_sup[a]["y"], first_sup[b]["y"],
        first_inf[a]["y"], first_inf[b]["y"],
    ]
    return (
        min(xs) - pad,
        min(ys) - pad,
        max(xs) - min(xs) + pad * 2.0,
        max(ys) - min(ys) + pad * 2.0,
        a,
        b,
    )


def _norm_pair(i0, i1):
    a, b = int(i0), int(i1)
    if a == b:
        return None
    return [min(a, b), max(a, b)]


def draft_to_text(draft, n_bars):
    return u"\n".join([
        u"n_barras_1a = {0}".format(n_bars),
        u"modo        = dibujo libre",
        u"perimetral  = {0}".format(bool(draft.get(u"perimetral"))),
        u"estribos E  = {0}".format(list(draft.get(u"pairs") or [])),
        u"trabas      = {0}".format(list(draft.get(u"ties") or [])),
    ])


class ConfinDesignerLab(object):
    """Canvas interactivo de evaluación de estribos/trabas."""

    def __init__(self):
        self.n_bars = 4
        self.mode = u"draw"  # draw | erase | peri
        self.draft = {
            u"perimetral": False,
            u"pairs": [],
            u"ties": [],
        }
        self._pending_pair = None  # índice clic 1/2
        self._hover_bar = None
        self._mouse = None  # (x, y) o None
        self._first_sup = []
        self._first_inf = []
        self._bar_pts_sup = []
        self._win = None
        self._canvas = None
        self._txt_mode = None
        self._txt_draft = None
        self._txt_hint = None
        self._txt_snap = None
        self._cmb_n = None
        self._btn_modes = {}
        self._nav_wired = False

    def show(self, owner=None):
        self._build_window()
        if owner is not None:
            try:
                self._win.Owner = owner
            except Exception:
                pass
        self._paint()
        self._win.Show()
        return self._win

    def _build_window(self):
        win = Window()
        win.Title = u"Lab · Estribos / Trabas (evaluación)"
        win.Width = 720
        win.Height = 580
        win.MinWidth = 640
        win.MinHeight = 480
        win.WindowStartupLocation = WindowStartupLocation.CenterScreen
        win.Background = brush_hex(u"#0a1620")

        root = Grid()
        root.Margin = Thickness(12)
        for _ in range(2):
            cd = ColumnDefinition()
            if _ == 0:
                cd.Width = GridLength(1.0, GridUnitType.Star)
            else:
                cd.Width = GridLength(260, GridUnitType.Pixel)
            root.ColumnDefinitions.Add(cd)

        # --- canvas column ---
        left = StackPanel()
        left.Margin = Thickness(0, 0, 12, 0)

        title = TextBlock()
        title.Text = u"Sección interactiva · 1ª capa (índices 0…n−1)"
        title.Foreground = brush_hex(u"#e2e8f0")
        title.FontSize = 14
        title.FontWeight = FontWeights.Bold
        title.Margin = Thickness(0, 0, 0, 8)
        left.Children.Add(title)

        tools = StackPanel()
        tools.Orientation = Orientation.Horizontal
        tools.Margin = Thickness(0, 0, 0, 8)

        tools.Children.Add(self._lbl(u"n bars"))
        cmb = ComboBox()
        cmb.Width = 56
        cmb.Margin = Thickness(4, 0, 10, 0)
        for n in range(2, 9):
            it = ComboBoxItem()
            it.Content = unicode(n)
            it.Tag = n
            cmb.Items.Add(it)
            if n == self.n_bars:
                cmb.SelectedItem = it
        try:
            from System.Windows.Controls import SelectionChangedEventHandler

            def _on_n(sender, args):
                try:
                    sel = cmb.SelectedItem
                    if sel is not None and sel.Tag is not None:
                        self._set_n_bars(int(sel.Tag))
                except Exception:
                    pass

            cmb.SelectionChanged += SelectionChangedEventHandler(_on_n)
        except Exception:
            pass
        self._cmb_n = cmb
        tools.Children.Add(cmb)

        for mid, label in (
            (u"draw", u"Dibujar"),
            (u"peri", u"Perim."),
            (u"erase", u"Borrar"),
        ):
            btn = self._tool_btn(label, mid)
            tools.Children.Add(btn)
            self._btn_modes[mid] = btn

        btn_clear = self._tool_btn(u"Limpiar", u"clear", accent=False)
        try:
            from System.Windows import RoutedEventHandler as _REH

            btn_clear.Click += _REH(lambda s, e: self._clear_draft())
        except Exception:
            pass
        tools.Children.Add(btn_clear)

        left.Children.Add(tools)

        self._txt_mode = TextBlock()
        self._txt_mode.Foreground = brush_hex(u"#5bb8d4")
        self._txt_mode.FontSize = 11
        self._txt_mode.FontWeight = FontWeights.SemiBold
        self._txt_mode.Margin = Thickness(0, 0, 0, 6)
        left.Children.Add(self._txt_mode)

        bdr = Border()
        bdr.Background = brush_hex(u"#071018")
        bdr.BorderBrush = brush_hex(u"#21465C")
        bdr.BorderThickness = Thickness(1)
        bdr.Padding = Thickness(4)
        bdr.HorizontalAlignment = HorizontalAlignment.Left

        cnv = Canvas()
        cnv.Width = _LAB_W
        cnv.Height = _LAB_H
        cnv.Background = brush_hex(u"#0a1620")
        try:
            cnv.Cursor = Cursors.Cross
        except Exception:
            pass
        # MouseMove para snap + rubber-band.
        try:
            cnv.MouseMove += MouseEventHandler(self._on_canvas_move)
            cnv.MouseLeave += MouseEventHandler(self._on_canvas_leave)
        except Exception:
            pass
        self._canvas = cnv
        bdr.Child = cnv
        left.Children.Add(bdr)

        self._txt_hint = TextBlock()
        self._txt_hint.TextWrapping = TextWrapping.Wrap
        self._txt_hint.Foreground = brush_hex(u"#94a3b8")
        self._txt_hint.FontSize = 10
        self._txt_hint.Margin = Thickness(0, 8, 0, 0)
        self._txt_hint.Text = self._hint_for_mode()
        left.Children.Add(self._txt_hint)

        self._txt_snap = TextBlock()
        self._txt_snap.TextWrapping = TextWrapping.Wrap
        self._txt_snap.Foreground = brush_hex(u"#5bb8d4")
        self._txt_snap.FontSize = 11
        self._txt_snap.FontWeight = FontWeights.SemiBold
        self._txt_snap.Margin = Thickness(0, 4, 0, 0)
        self._txt_snap.Text = u""
        left.Children.Add(self._txt_snap)

        Grid.SetColumn(left, 0)
        root.Children.Add(left)

        # --- right panel ---
        right = StackPanel()
        right.Margin = Thickness(4, 0, 0, 0)

        rh = TextBlock()
        rh.Text = u"Borrador (draft)"
        rh.Foreground = brush_hex(u"#e2e8f0")
        rh.FontWeight = FontWeights.Bold
        rh.FontSize = 12
        rh.Margin = Thickness(0, 0, 0, 6)
        right.Children.Add(rh)

        tb = TextBox()
        tb.IsReadOnly = True
        tb.AcceptsReturn = True
        tb.TextWrapping = TextWrapping.Wrap
        tb.Height = 140
        try:
            tb.FontFamily = FontFamily(u"Consolas")
        except Exception:
            pass
        tb.FontSize = 11
        tb.Foreground = brush_hex(u"#e2e8f0")
        tb.Background = brush_hex(u"#071018")
        tb.BorderBrush = brush_hex(u"#21465C")
        tb.Padding = Thickness(8)
        self._txt_draft = tb
        right.Children.Add(tb)

        note = TextBlock()
        note.Text = (
            u"Prototipo: no modifica Colocar ni SESSION.\n"
            u"Confinamiento solo por dibujo libre "
            u"(pairs / ties / perimetral). Sin catálogo."
        )
        note.Foreground = brush_hex(u"#64748b")
        note.FontSize = 9.5
        note.Margin = Thickness(0, 12, 0, 0)
        note.TextWrapping = TextWrapping.Wrap
        right.Children.Add(note)

        Grid.SetColumn(right, 1)
        root.Children.Add(right)

        win.Content = root
        self._win = win
        try:
            win.KeyDown += KeyEventHandler(self._on_key)
            win.Focusable = True
        except Exception:
            pass
        self._refresh_mode_chrome()

    def _on_key(self, sender, args):
        try:
            if args.Key == Key.Escape and self._pending_pair is not None:
                self._pending_pair = None
                self._paint()
                try:
                    args.Handled = True
                except Exception:
                    pass
        except Exception:
            pass

    def _on_canvas_move(self, sender, args):
        cnv = self._canvas
        if cnv is None:
            return
        try:
            pos = args.GetPosition(cnv)
            mx, my = float(pos.X), float(pos.Y)
        except Exception:
            return
        self._mouse = (mx, my)
        hover = self._nearest_bar(mx, my)
        need = hover != self._hover_bar or self._pending_pair is not None
        self._hover_bar = hover
        if need:
            self._paint()

    def _on_canvas_leave(self, sender, args):
        self._mouse = None
        if self._hover_bar is not None or self._pending_pair is not None:
            self._hover_bar = None
            self._paint()

    def _nearest_bar(self, mx, my):
        best = None
        best_d = _SNAP_R
        for i, pt in enumerate(self._first_sup or []):
            for p in (pt, (self._first_inf or [None])[i] if i < len(self._first_inf or []) else None):
                if not p:
                    continue
                d = ((mx - float(p["x"])) ** 2 + (my - float(p["y"])) ** 2) ** 0.5
                if d <= best_d:
                    best_d = d
                    best = i
        return best

    def _lbl(self, text):
        tb = TextBlock()
        tb.Text = text
        tb.Foreground = brush_hex(u"#94a3b8")
        tb.FontSize = 11
        tb.VerticalAlignment = VerticalAlignment.Center
        tb.Margin = Thickness(0, 0, 2, 0)
        return tb

    def _tool_btn(self, content, mode_id, accent=True):
        btn = Button()
        btn.Content = content
        btn.MinWidth = 64
        btn.Height = 28
        btn.Margin = Thickness(0, 0, 6, 0)
        btn.Padding = Thickness(8, 2, 8, 2)
        btn.Foreground = brush_hex(u"#e2e8f0")
        btn.Background = brush_hex(u"#122636")
        btn.BorderBrush = brush_hex(u"#21465C")
        try:
            from System.Windows import RoutedEventHandler as _REH

            if mode_id == u"clear":
                pass  # wired by caller
            else:
                btn.Click += _REH(lambda s, e, m=mode_id: self._set_mode(m))
        except Exception:
            pass
        return btn

    def _set_mode(self, mode):
        self.mode = mode
        self._pending_pair = None
        self._refresh_mode_chrome()
        self._paint()

    def _set_n_bars(self, n):
        n = max(2, min(8, int(n)))
        if n == self.n_bars:
            return
        self.n_bars = n
        # Poda indices fuera de rango
        self.draft[u"pairs"] = [
            p for p in (self.draft.get(u"pairs") or [])
            if p and max(int(p[0]), int(p[1])) < n
        ]
        self.draft[u"ties"] = [
            t for t in (self.draft.get(u"ties") or []) if 0 <= int(t) < n
        ]
        self._pending_pair = None
        self._paint()

    def _clear_draft(self):
        self.draft = {u"perimetral": False, u"pairs": [], u"ties": []}
        self._pending_pair = None
        self._paint()

    def _hint_for_mode(self):
        if self.mode == u"draw":
            return (
                u"Clic ancla · doble clic cierra: distinta col. → ESTRIBO; "
                u"misma col. → TRABA. Esc cancela."
            )
        if self.mode == u"peri":
            return u"Clic en el hormigón (sin barra) alterna el estribo exterior."
        if self.mode == u"erase":
            return u"Clic en estribo, traba o barra asociada para quitar."
        return u""

    def _refresh_snap_status(self):
        tb = self._txt_snap
        if tb is None:
            return
        h = self._hover_bar
        p = self._pending_pair
        if self.mode == u"draw":
            if p is None and h is not None:
                tb.Text = u"Snap → [{0}] · clic ancla".format(h)
            elif p is not None and h is not None and h != p:
                pr = _norm_pair(p, h)
                tb.Text = u"ESTRIBO E({0}–{1}) · doble clic cierra".format(pr[0], pr[1])
            elif p is not None and h == p:
                tb.Text = u"TRABA [{0}] · doble clic cierra".format(p)
            elif p is not None:
                tb.Text = u"A=[{0}] · doble clic: otra col.=estribo · misma=traba".format(p)
            else:
                tb.Text = u"Clic ancla · doble clic cierra estribo/traba"
        else:
            tb.Text = u""

    def _refresh_mode_chrome(self):
        labels = {
            u"draw": u"Dibujar · clic + doble clic · estribo/traba",
            u"peri": u"Perimetral ON/OFF",
            u"erase": u"Borrar",
        }
        if self._txt_mode is not None:
            self._txt_mode.Text = labels.get(self.mode, self.mode)
        if self._txt_hint is not None:
            self._txt_hint.Text = self._hint_for_mode()
        self._refresh_snap_status()
        for mid, btn in (self._btn_modes or {}).items():
            try:
                if mid == self.mode:
                    btn.Background = brush_hex(u"#164e63")
                    btn.BorderBrush = brush_hex(u"#22d3ee")
                else:
                    btn.Background = brush_hex(u"#122636")
                    btn.BorderBrush = brush_hex(u"#21465C")
            except Exception:
                pass

    def _update_side_panel(self):
        if self._txt_draft is not None:
            self._txt_draft.Text = draft_to_text(self.draft, self.n_bars)

    def _paint(self):
        cnv = self._canvas
        if cnv is None:
            return
        cnv.Children.Clear()
        self._bar_pts_sup = []

        cw, ch = _LAB_W, _LAB_H
        ox, oy, sec_w, sec_h = _fit_section_rect(
            _SECTION_B_CM, _SECTION_H_CM, cw, ch, pad_x=24.0, pad_top=16.0, label_h=28.0
        )
        cover = max(10.0, min(sec_w, sec_h) * 0.12)
        inner_x = ox + cover
        inner_y = oy + cover
        inner_w = sec_w - cover * 2.0
        inner_h = sec_h - cover * 2.0
        st_inset = 4.0
        st_x = inner_x + st_inset
        st_y = inner_y + st_inset
        st_w = inner_w - st_inset * 2.0
        st_h = inner_h - st_inset * 2.0
        scale = sec_h / max(_SECTION_H_CM, 1.0)

        # hormigón
        outer = Rectangle()
        outer.Width = sec_w
        outer.Height = sec_h
        Canvas.SetLeft(outer, ox)
        Canvas.SetTop(outer, oy)
        outer.RadiusX = 2.0
        outer.RadiusY = 2.0
        outer.Stroke = brush_hex(u"#5bb8d4")
        outer.StrokeThickness = 1.4
        outer.Fill = brush_hex(u"#0a1620", 230)
        if self.mode == u"peri":
            outer.Cursor = Cursors.Hand
            try:
                outer.MouseLeftButtonUp += MouseButtonEventHandler(
                    lambda s, e: self._toggle_peri()
                )
            except Exception:
                pass
        cnv.Children.Add(outer)

        # cover dash
        inner = Rectangle()
        inner.Width = inner_w
        inner.Height = inner_h
        Canvas.SetLeft(inner, inner_x)
        Canvas.SetTop(inner, inner_y)
        inner.Stroke = brush_hex(u"#94a3b8", 70)
        inner.StrokeThickness = 0.8
        try:
            from System.Windows.Media import DoubleCollection

            inner.StrokeDashArray = DoubleCollection([3.0, 2.5])
        except Exception:
            pass
        inner.Fill = SolidColorBrush(Color.FromArgb(0, 0, 0, 0))
        if self.mode == u"peri":
            try:
                inner.MouseLeftButtonUp += MouseButtonEventHandler(
                    lambda s, e: self._toggle_peri()
                )
            except Exception:
                pass
        cnv.Children.Add(inner)

        bar_pad = max(6.0, cover * 0.5)
        bar_x0 = st_x + bar_pad
        bar_x1 = st_x + st_w - bar_pad
        sup_y = st_y + bar_pad
        inf_y = st_y + st_h - bar_pad
        n = self.n_bars
        first_sup = _distribute_points(n, bar_x0, bar_x1, sup_y)
        first_inf = _distribute_points(n, bar_x0, bar_x1, inf_y)
        self._first_sup = first_sup
        self._first_inf = first_inf
        self._bar_pts_sup = [{"x": p["x"], "y": p["y"], "i": i} for i, p in enumerate(first_sup)]

        r_est = 1.4
        draft = self.draft
        m_stir = stirrup_margin_px(bar_r=3.5, pad=4.0)

        # Perimetral / E-pairs: estribo 135° (estilo Armado Muros V3)
        if draft.get(u"perimetral"):
            all_pts = list(first_sup) + list(first_inf)
            draw_estribo_bbox(
                cnv, all_pts, margin=m_stir,
                brush=brush_hex(COLOR_ESTRIBO), thick=STIRRUP_THICK,
            )

        for pair in draft.get(u"pairs") or []:
            if len(pair) < 2:
                continue
            i0, i1 = int(pair[0]), int(pair[1])
            if i0 >= n or i1 >= n:
                continue
            a, b = min(i0, i1), max(i0, i1)
            pts = [first_sup[a], first_sup[b], first_inf[a], first_inf[b]]
            hit = draw_estribo_bbox(
                cnv, pts, margin=m_stir,
                brush=brush_hex(COLOR_ESTRIBO), thick=STIRRUP_THICK,
            )
            # Capa hit invisible para borrar (bbox)
            if hit is not None:
                rx, ry, rw, rh = hit
                inv = Rectangle()
                inv.Width = max(rw, 4.0)
                inv.Height = max(rh, 4.0)
                Canvas.SetLeft(inv, rx)
                Canvas.SetTop(inv, ry)
                inv.Fill = SolidColorBrush(Color.FromArgb(1, 0, 0, 0))
                inv.Stroke = SolidColorBrush(Color.FromArgb(0, 0, 0, 0))
                inv.Tag = (u"pair", i0, i1)
                if self.mode == u"erase":
                    inv.Cursor = Cursors.Hand
                    try:
                        inv.MouseLeftButtonUp += MouseButtonEventHandler(
                            lambda s, e, aa=i0, bb=i1: self._erase_pair(aa, bb)
                        )
                    except Exception:
                        pass
                cnv.Children.Add(inv)
            try:
                tag = TextBlock()
                tag.Text = u"ESTRIBO E({0}–{1})".format(a, b)
                tag.FontSize = 9
                tag.FontWeight = FontWeights.SemiBold
                tag.Foreground = brush_hex(u"#86efac")
                if hit is not None:
                    Canvas.SetLeft(tag, hit[0] + 3)
                    Canvas.SetTop(tag, hit[1] + 2)
                cnv.Children.Add(tag)
            except Exception:
                pass

        for idx in draft.get(u"ties") or []:
            ti = int(idx)
            if ti >= n:
                continue
            ln = draw_traba_vertical_column(
                cnv,
                first_sup[ti]["x"], first_sup[ti]["y"],
                first_inf[ti]["x"], first_inf[ti]["y"],
                margin=m_stir,
                brush=brush_hex(COLOR_TRABA),
                thick=TIE_THICK,
            )
            if ln is not None:
                ln.Tag = (u"tie", ti)
                if self.mode == u"erase":
                    ln.Cursor = Cursors.Hand
                    try:
                        ln.MouseLeftButtonUp += MouseButtonEventHandler(
                            lambda s, e, k=ti: self._erase_tie(k)
                        )
                    except Exception:
                        pass
            try:
                tag = TextBlock()
                tag.Text = u"TRABA [{0}]".format(ti)
                tag.FontSize = 9
                tag.FontWeight = FontWeights.SemiBold
                tag.Foreground = brush_hex(u"#fdba74")
                x = (first_sup[ti]["x"] + first_inf[ti]["x"]) * 0.5
                Canvas.SetLeft(tag, x + 3)
                Canvas.SetTop(tag, (first_sup[ti]["y"] + first_inf[ti]["y"]) * 0.5 - 6)
                cnv.Children.Add(tag)
            except Exception:
                pass

        # Rubber-band E-pair (preview polígono A→B)
        self._draw_rubber_band(cnv, first_sup, first_inf, n)

        # Aros de snap + barras
        for i in range(n):
            is_a = self._pending_pair == i
            is_h = self._hover_bar == i
            for pt, fill, stroke, r0 in (
                (first_sup[i], u"#22d3ee", u"#0891b2", 4.5),
                (first_inf[i], u"#f87171", u"#b91c1c", 4.0),
            ):
                # aro snap
                if self.mode == u"draw" or is_a or is_h:
                    ring = Ellipse()
                    rr = _SNAP_R * 2.0
                    ring.Width = rr
                    ring.Height = rr
                    Canvas.SetLeft(ring, pt["x"] - _SNAP_R)
                    Canvas.SetTop(ring, pt["y"] - _SNAP_R)
                    ring.Stroke = brush_hex(
                        u"#22d3ee" if is_h or is_a else u"#5bb8d4",
                        200 if (is_h or is_a) else 40,
                    )
                    ring.StrokeThickness = 1.4 if (is_h or is_a) else 0.8
                    try:
                        if not (is_h or is_a):
                            ring.StrokeDashArray = DoubleCollection([3.0, 3.0])
                    except Exception:
                        pass
                    ring.Fill = SolidColorBrush(Color.FromArgb(0, 0, 0, 0))
                    ring.IsHitTestVisible = False
                    cnv.Children.Add(ring)

                r = r0 + (1.5 if (is_a or is_h) else 0.0)
                f = fill
                if is_a:
                    f = u"#a5f3fc" if pt is first_sup[i] else u"#fda4af"
                elif is_h:
                    f = u"#67e8f9" if pt is first_sup[i] else u"#fb7185"
                _add_dot(cnv, pt["x"], pt["y"], r, f, stroke, 255)
                self._add_bar_hit(cnv, pt["x"], pt["y"], i, r)

            lab = TextBlock()
            lab.Text = unicode(i)
            lab.FontSize = 9
            lab.FontWeight = FontWeights.Bold
            lab.Foreground = brush_hex(u"#e0f2fe" if is_a or is_h else u"#67e8f9")
            Canvas.SetLeft(lab, first_sup[i]["x"] - 3)
            Canvas.SetTop(lab, first_sup[i]["y"] - 18)
            lab.IsHitTestVisible = False
            cnv.Children.Add(lab)
            lab2 = TextBlock()
            lab2.Text = unicode(i)
            lab2.FontSize = 9
            lab2.FontWeight = FontWeights.Bold
            lab2.Foreground = brush_hex(u"#fecaca")
            Canvas.SetLeft(lab2, first_inf[i]["x"] - 3)
            Canvas.SetTop(lab2, first_inf[i]["y"] + 8)
            lab2.IsHitTestVisible = False
            cnv.Children.Add(lab2)

        # dim
        dim = TextBlock()
        dim.Text = u"{0:.0f} × {1:.0f} cm  ·  clic + doble clic: estribo/traba".format(
            _SECTION_B_CM, _SECTION_H_CM
        )
        dim.FontSize = 10
        dim.Foreground = brush_hex(u"#64748b")
        dim.Width = cw
        dim.TextAlignment = TextAlignment.Center
        dim.IsHitTestVisible = False
        Canvas.SetLeft(dim, 0)
        Canvas.SetTop(dim, ch - 20)
        cnv.Children.Add(dim)

        self._refresh_snap_status()
        self._update_side_panel()

    def _draw_rubber_band(self, cnv, first_sup, first_inf, n):
        """Preview: estribo/traba 135° (estilo Muros) según hover tras clic 1."""
        if self.mode != u"draw":
            return
        p = self._pending_pair
        h = self._hover_bar
        if p is None:
            return
        m_stir = stirrup_margin_px()

        # Misma columna → traba 135°
        if h is not None and h == p and 0 <= p < n:
            draw_traba_vertical_column(
                cnv,
                first_sup[p]["x"], first_sup[p]["y"],
                first_inf[p]["x"], first_inf[p]["y"],
                margin=m_stir,
                brush=brush_hex(COLOR_TRABA, 200),
                thick=TIE_THICK,
            )
            tag = TextBlock()
            tag.Text = u"TRABA [{0}] · preview".format(p)
            tag.FontSize = 10
            tag.FontWeight = FontWeights.SemiBold
            tag.Foreground = brush_hex(u"#fdba74")
            tag.IsHitTestVisible = False
            x = (first_sup[p]["x"] + first_inf[p]["x"]) * 0.5
            Canvas.SetLeft(tag, x + 6)
            Canvas.SetTop(tag, (first_sup[p]["y"] + first_inf[p]["y"]) * 0.5 - 6)
            cnv.Children.Add(tag)
            return

        # Distinta columna → estribo 135°
        if h is not None and h != p and 0 <= h < n and 0 <= p < n:
            a, b = min(p, h), max(p, h)
            pts = [first_sup[a], first_sup[b], first_inf[a], first_inf[b]]
            hit = draw_estribo_bbox(
                cnv, pts, margin=m_stir,
                brush=brush_hex(COLOR_ESTRIBO, 200), thick=STIRRUP_THICK,
            )
            tag = TextBlock()
            tag.Text = u"ESTRIBO E({0}–{1}) · preview".format(a, b)
            tag.FontSize = 10
            tag.FontWeight = FontWeights.SemiBold
            tag.Foreground = brush_hex(u"#a5f3fc")
            tag.IsHitTestVisible = False
            if hit is not None:
                Canvas.SetLeft(tag, hit[0] + 4)
                Canvas.SetTop(tag, max(0.0, hit[1] - 16))
            cnv.Children.Add(tag)
            return

        # Sin snap B: sombra + pista de traba en A
        if self._mouse is not None and 0 <= p < n:
            try:
                mx = float(self._mouse[0])
            except Exception:
                return
            x_a = float(first_sup[p]["x"])
            rx = min(x_a, mx) - _PAIR_PAD
            ry = min(float(first_sup[p]["y"]), float(first_inf[p]["y"])) - _PAIR_PAD
            rw = abs(mx - x_a) + _PAIR_PAD * 2.0
            rh = abs(float(first_inf[p]["y"]) - float(first_sup[p]["y"])) + _PAIR_PAD * 2.0
            rect = Rectangle()
            rect.Width = max(8.0, rw)
            rect.Height = rh
            Canvas.SetLeft(rect, rx)
            Canvas.SetTop(rect, ry)
            rect.Stroke = brush_hex(u"#94a3b8", 160)
            rect.StrokeThickness = 1.2
            try:
                rect.StrokeDashArray = DoubleCollection([4.0, 4.0])
            except Exception:
                pass
            rect.Fill = SolidColorBrush(Color.FromArgb(20, 148, 163, 184))
            rect.IsHitTestVisible = False
            cnv.Children.Add(rect)
            x = (first_sup[p]["x"] + first_inf[p]["x"]) * 0.5
            ln = Line()
            ln.X1 = x
            ln.Y1 = first_sup[p]["y"]
            ln.X2 = x
            ln.Y2 = first_inf[p]["y"]
            ln.Stroke = brush_hex(u"#fb923c", 90)
            ln.StrokeThickness = 1.5
            try:
                ln.StrokeDashArray = DoubleCollection([3.0, 3.0])
            except Exception:
                pass
            ln.IsHitTestVisible = False
            cnv.Children.Add(ln)

    def _add_bar_hit(self, cnv, x, y, i, r):
        hit = Ellipse()
        hit_r = max(_SNAP_R, r + _HIT_PAD)
        hit.Width = hit_r * 2.0
        hit.Height = hit_r * 2.0
        Canvas.SetLeft(hit, x - hit_r)
        Canvas.SetTop(hit, y - hit_r)
        hit.Fill = SolidColorBrush(Color.FromArgb(1, 0, 0, 0))
        hit.StrokeThickness = 0
        hit.Cursor = Cursors.Hand
        hit.Tag = (u"bar", i)
        try:
            hit.MouseLeftButtonUp += MouseButtonEventHandler(
                lambda s, e, idx=i: self._on_bar_click(idx, e)
            )
        except Exception:
            pass
        cnv.Children.Add(hit)

    def _toggle_peri(self):
        self.draft[u"perimetral"] = not bool(self.draft.get(u"perimetral"))
        self._paint()

    def _on_bar_click(self, i, args=None):
        i = int(i)
        if self.mode == u"draw":
            try:
                import time as _time

                if _time.time() < float(getattr(self, u"_click_guard", 0) or 0):
                    return
            except Exception:
                pass
            if self._pending_pair is None:
                self._pending_pair = i
                self._paint()
                return
            a = int(self._pending_pair)
            b = i
            self._pending_pair = None
            if int(a) == int(b):
                ties = [int(t) for t in (self.draft.get(u"ties") or [])]
                if a in ties:
                    ties = [t for t in ties if t != a]
                else:
                    ties.append(a)
                    ties = sorted(set(ties))
                self.draft[u"ties"] = ties
            else:
                pair = _norm_pair(a, b)
                pairs = []
                for p in (self.draft.get(u"pairs") or []):
                    np = _norm_pair(p[0], p[1])
                    if np and np not in pairs:
                        pairs.append(np)
                if pair in pairs:
                    pairs = [p for p in pairs if p != pair]
                else:
                    pairs.append(pair)
                self.draft[u"pairs"] = pairs
            try:
                import time as _time

                self._click_guard = _time.time() + 0.35
            except Exception:
                self._click_guard = 0.0
            self._paint()
            return
        if self.mode == u"erase":
            # borrar by index
            ties = [int(t) for t in (self.draft.get(u"ties") or [])]
            if i in ties:
                self.draft[u"ties"] = [t for t in ties if t != i]
            else:
                pairs = []
                for p in (self.draft.get(u"pairs") or []):
                    np = _norm_pair(p[0], p[1])
                    if np and i not in np:
                        pairs.append(np)
                self.draft[u"pairs"] = pairs
            self._paint()
            return
        if self.mode == u"peri":
            return
        return

    def _erase_pair(self, i0, i1):
        pair = _norm_pair(i0, i1)
        if not pair:
            return
        self.draft[u"pairs"] = [
            p for p in (self.draft.get(u"pairs") or [])
            if _norm_pair(p[0], p[1]) != pair
        ]
        self._paint()

    def _erase_tie(self, k):
        self.draft[u"ties"] = [
            t for t in (self.draft.get(u"ties") or []) if int(t) != int(k)
        ]
        self._paint()


def show_lab(uiapp=None):
    """Abre el laboratorio (modeless)."""
    lab = ConfinDesignerLab()
    owner = None
    try:
        if uiapp is not None:
            # pyRevit: sin owner nativo fácil; None OK
            owner = None
    except Exception:
        owner = None
    return lab.show(owner=owner)


def run_pyrevit(uiapp):
    """Entry pyRevit sin selección de modelo."""
    show_lab(uiapp)
