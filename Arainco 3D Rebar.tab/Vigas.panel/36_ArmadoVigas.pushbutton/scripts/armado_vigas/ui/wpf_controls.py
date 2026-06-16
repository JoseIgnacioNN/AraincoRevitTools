# -*- coding: utf-8 -*-
"""Controles WPF reutilizables (stepper, combo, pinceles)."""

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import HorizontalAlignment, TextAlignment, Thickness, VerticalAlignment, FontWeights, GridLength, GridUnitType
from System.Windows.Controls import (
    Border,
    Button,
    ComboBox,
    ComboBoxItem,
    Orientation,
    StackPanel,
    TextBlock,
)
from System.Windows.Media import SolidColorBrush, Color

from armado_vigas.domain.constants import LONG_DIAM_OPTS, ESTRIBO_SPACING_MIN, ESTRIBO_SPACING_OPTS
from armado_vigas.ui import typography as typo
from armado_vigas.ui import theme as th


def brush_hex(hx, alpha=255):
    h = (hx or u"#64748b").strip().lstrip(u"#")
    if len(h) < 6:
        h = u"64748b"
    rr = int(h[0:2], 16)
    gg = int(h[2:4], 16)
    bb = int(h[4:6], 16)
    aa = max(0, min(255, int(alpha)))
    return SolidColorBrush(Color.FromArgb(aa, rr, gg, bb))


def label_small(text):
    tb = TextBlock()
    tb.Text = text or u""
    tb.Foreground = th.brush_fg_mid()
    tb.FontSize = typo.LABEL_FONT_PX
    tb.FontWeight = FontWeights.SemiBold
    tb.Margin = Thickness(0, 0, 0, 2)
    return tb


def _combo_font_size(compact):
    return typo.CTRL_FONT_PX if compact else 11.0


def _combo_height(compact):
    return typo.CTRL_HEIGHT_PX if compact else 24.0


# Opacidades tramo Tn (alpha 0–255) — mockup Opción D accentSoft
TRAMO_SOFT_ALPHA = {
    "fill": 66,
    "fillSel": 92,
    "stroke": 179,
    "strokeSel": 217,
    "border": 122,
    "text": 209,
    "chipBg": 20,
    "chipBgEdit": 33,
    "chipBorder": 107,
    "stripBg": 15,
    "stripBgSel": 28,
    "swatch": 128,
    "legendDot": 148,
    "halo": 31,
}


def accent_soft_brush(hex_color, key):
    return brush_hex(hex_color, TRAMO_SOFT_ALPHA.get(key, 128))


def _add_combo_item(cb, content, tag=None):
    item = ComboBoxItem()
    item.Content = content
    item.FontSize = typo.CTRL_FONT_PX
    if tag is not None:
        item.Tag = tag
    cb.Items.Add(item)
    return item


def make_stepper(win, value, min_v, max_v, step, on_change, compact=False, enabled=True):
    """Stepper compacto: valor a la izquierda · flechas ▲/▼ a la derecha."""
    from System.Windows.Controls import Grid, ColumnDefinition, RowDefinition

    shell = Border()
    shell.Background = th.brush_input()
    shell.BorderBrush = th.brush_border_input()
    shell.BorderThickness = Thickness(1)
    try:
        from System.Windows import CornerRadius
        shell.CornerRadius = CornerRadius(4.0)
    except Exception:
        pass
    shell.Padding = Thickness(0)
    shell.Height = _combo_height(compact)
    arrow_w = 16.0 if compact else 18.0
    val_min_w = 24.0 if compact else 28.0
    ctrl_fs = _combo_font_size(compact)
    btn_h = _combo_height(compact)

    panel = Grid()
    panel.SnapsToDevicePixels = True
    shell.Child = panel

    cd_val = ColumnDefinition()
    cd_val.Width = GridLength(1.0, GridUnitType.Star)
    cd_arr = ColumnDefinition()
    cd_arr.Width = GridLength(arrow_w)
    panel.ColumnDefinitions.Add(cd_val)
    panel.ColumnDefinitions.Add(cd_arr)

    val_tb = TextBlock()
    val_tb.Text = unicode(int(round(value)))
    val_tb.MinWidth = val_min_w
    val_tb.TextAlignment = TextAlignment.Center
    val_tb.VerticalAlignment = VerticalAlignment.Center
    val_tb.HorizontalAlignment = HorizontalAlignment.Center
    val_tb.Foreground = th.brush_fg_hi()
    val_tb.FontWeight = FontWeights.Bold
    val_tb.FontSize = ctrl_fs
    Grid.SetColumn(val_tb, 0)
    panel.Children.Add(val_tb)

    arrow_wrap = Border()
    arrow_wrap.BorderBrush = th.brush_border_input()
    arrow_wrap.BorderThickness = Thickness(1, 0, 0, 0)
    arrow_wrap.Background = th.brush_border_muted(180)
    Grid.SetColumn(arrow_wrap, 1)

    arrow_panel = Grid()
    arrow_wrap.Child = arrow_panel

    rd_up = RowDefinition()
    rd_up.Height = GridLength(1.0, GridUnitType.Star)
    rd_dn = RowDefinition()
    rd_dn.Height = GridLength(1.0, GridUnitType.Star)
    arrow_panel.RowDefinitions.Add(rd_up)
    arrow_panel.RowDefinitions.Add(rd_dn)

    sep_h = Border()
    sep_h.Height = 1.0
    sep_h.Background = th.brush_border_input()
    sep_h.VerticalAlignment = VerticalAlignment.Center
    Grid.SetRow(sep_h, 0)
    Grid.SetRowSpan(sep_h, 2)
    arrow_panel.Children.Add(sep_h)

    def _apply_style(btn):
        try:
            if win is not None:
                st = win.TryFindResource(u"BimToolsStepperZoneBtn")
                if st is not None:
                    btn.Style = st
        except Exception:
            pass
        btn.Padding = Thickness(0)
        btn.Margin = Thickness(0)
        btn.FontSize = 7.0 if compact else 8.0
        btn.HorizontalAlignment = HorizontalAlignment.Stretch
        btn.VerticalAlignment = VerticalAlignment.Stretch

    def _set_val(n):
        n = max(int(min_v), min(int(max_v), int(round(n))))
        val_tb.Text = unicode(n)
        if on_change:
            on_change(n)

    btn_up = Button()
    btn_up.Content = u"▲"
    _apply_style(btn_up)
    btn_dn = Button()
    btn_dn.Content = u"▼"
    _apply_style(btn_dn)
    Grid.SetRow(btn_up, 0)
    Grid.SetRow(btn_dn, 1)
    arrow_panel.Children.Add(btn_up)
    arrow_panel.Children.Add(btn_dn)
    panel.Children.Add(arrow_wrap)

    def _up(sender, args):
        try:
            cur = int(val_tb.Text)
        except Exception:
            cur = int(value)
        _set_val(cur + int(step))

    def _dn(sender, args):
        try:
            cur = int(val_tb.Text)
        except Exception:
            cur = int(value)
        _set_val(cur - int(step))

    try:
        from System.Windows import RoutedEventHandler as _REH
        btn_up.Click += _REH(_up)
        btn_dn.Click += _REH(_dn)
    except Exception:
        pass

    shell.IsEnabled = bool(enabled)
    if not enabled:
        shell.Opacity = 0.65
    return shell


def make_diam_combo(win, value, diam_opts=None, on_change=None, compact=False, enabled=True):
    opts = diam_opts or LONG_DIAM_OPTS
    cb = ComboBox()
    try:
        if win is not None:
            st = win.TryFindResource(u"Combo")
            if st is not None:
                cb.Style = st
    except Exception:
        pass
    diam_w = 52.0 if compact else 72.0
    cb.Width = diam_w
    cb.MinWidth = diam_w
    cb.MaxWidth = diam_w
    cb.Height = _combo_height(compact)
    cb.FontSize = _combo_font_size(compact)
    cb.Margin = Thickness(1, 0, 0, 0) if compact else Thickness(0)
    cur = int(value or opts[0])
    loading = [True]
    for d in opts:
        item = _add_combo_item(cb, u"ø{0}".format(int(d)), int(d))
        if int(d) == cur:
            cb.SelectedItem = item

    if on_change:
        def _changed(sender, args):
            if loading[0]:
                return
            try:
                sel = cb.SelectedItem
                if sel is not None and sel.Tag is not None:
                    on_change(int(sel.Tag))
            except Exception:
                pass
        try:
            from System.Windows.Controls import SelectionChangedEventHandler
            cb.SelectionChanged += SelectionChangedEventHandler(_changed)
        except Exception:
            pass
    loading[0] = False
    cb.IsEnabled = bool(enabled)
    if not enabled:
        cb.Opacity = 0.65
    return cb


def make_spacing_combo(win, value, on_change=None, compact=False, enabled=True, width=None):
    """Combo @ espaciado estribos (mm), valores cada 25."""
    opts = ESTRIBO_SPACING_OPTS
    try:
        cur = int(round(value))
    except Exception:
        cur = int(opts[0])
    cur = max(int(ESTRIBO_SPACING_MIN), cur)
    if cur not in opts:
        cur = min(opts, key=lambda o: abs(int(o) - cur))

    cb = ComboBox()
    try:
        if win is not None:
            st = win.TryFindResource(u"Combo")
            if st is not None:
                cb.Style = st
    except Exception:
        pass
    sp_w = width if width is not None else (52.0 if compact else 52.0)
    cb.Width = sp_w
    cb.MinWidth = sp_w
    cb.MaxWidth = sp_w
    cb.Height = _combo_height(compact)
    cb.FontSize = _combo_font_size(compact)
    cb.Margin = Thickness(1, 0, 0, 0) if compact else Thickness(0)

    loading = [True]
    for sp in opts:
        item = _add_combo_item(cb, unicode(int(sp)), int(sp))
        if int(sp) == cur:
            cb.SelectedItem = item

    if on_change:
        def _changed(sender, args):
            if loading[0]:
                return
            try:
                sel = cb.SelectedItem
                if sel is not None and sel.Tag is not None:
                    on_change(int(sel.Tag))
            except Exception:
                pass
        try:
            from System.Windows.Controls import SelectionChangedEventHandler
            cb.SelectionChanged += SelectionChangedEventHandler(_changed)
        except Exception:
            pass
    loading[0] = False
    cb.IsEnabled = bool(enabled)
    if not enabled:
        cb.Opacity = 0.65
    return cb


def make_spacing_stepper(win, value, on_change, compact=False):
    return make_stepper(
        win,
        value,
        ESTRIBO_SPACING_MIN,
        400,
        25,
        on_change,
        compact=compact,
    )


def make_string_combo(win, options, value, on_change, compact=False):
    cb = ComboBox()
    try:
        if win is not None:
            st = win.TryFindResource(u"ComboStretch")
            if st is None:
                st = win.TryFindResource(u"Combo")
            if st is not None:
                cb.Style = st
    except Exception:
        pass
    if compact:
        cb.MinWidth = 0.0
        cb.Height = _combo_height(True)
        cb.FontSize = _combo_font_size(True)
        cb.HorizontalAlignment = HorizontalAlignment.Stretch
    else:
        cb.MinWidth = 120.0
        cb.Height = 24.0
        cb.FontSize = 11.0
    cur = value
    loading = [True]
    for opt in options or []:
        item = _add_combo_item(cb, unicode(opt), opt)
        if opt == cur:
            cb.SelectedItem = item
    if on_change:
        def _changed(sender, args):
            if loading[0]:
                return
            try:
                sel = cb.SelectedItem
                if sel is not None:
                    on_change(sel.Tag)
            except Exception:
                pass
        try:
            from System.Windows.Controls import SelectionChangedEventHandler
            cb.SelectionChanged += SelectionChangedEventHandler(_changed)
        except Exception:
            pass
    loading[0] = False
    return cb


def make_capas_stepper(win, value, on_change, compact=False):
    from armado_vigas.domain.constants import CAPAS_MIN, CAPAS_MAX
    return make_stepper(win, value, CAPAS_MIN, CAPAS_MAX, 1, on_change, compact=compact)


def make_yesno_toggle(win, value, on_change, compact=False, enabled=True):
    """Toggle compacto Sí / No (estilo mockup suple inferior)."""
    shell = Border()
    shell.Background = th.brush_input()
    shell.BorderBrush = th.brush_border_input()
    shell.BorderThickness = Thickness(1)
    try:
        from System.Windows import CornerRadius
        shell.CornerRadius = CornerRadius(3.0)
    except Exception:
        pass
    shell.Padding = Thickness(0)
    shell.Height = _combo_height(compact)
    if compact:
        shell.HorizontalAlignment = HorizontalAlignment.Right

    panel = StackPanel()
    panel.Orientation = Orientation.Horizontal
    shell.Child = panel

    btn_h = typo.CTRL_HEIGHT_PX - 2.0 if compact else 22.0
    btn_fs = _combo_font_size(compact)
    btn_pad = 11.0 if compact else 12.0

    btn_yes = Button()
    btn_yes.Content = u"Sí"
    btn_yes.Padding = Thickness(btn_pad, 0, btn_pad, 0)
    btn_yes.Margin = Thickness(0)
    btn_yes.Height = btn_h
    btn_yes.FontSize = btn_fs
    btn_yes.FontWeight = FontWeights.Bold
    btn_yes.BorderThickness = Thickness(0)

    btn_no = Button()
    btn_no.Content = u"No"
    btn_no.Padding = Thickness(btn_pad, 0, btn_pad, 0)
    btn_no.Margin = Thickness(0)
    btn_no.Height = btn_h
    btn_no.FontSize = btn_fs
    btn_no.FontWeight = FontWeights.Bold
    btn_no.BorderThickness = Thickness(0)

    state = [bool(value)]

    def _apply_style():
        if state[0]:
            btn_yes.Background = brush_hex(u"#4ade80", 51)
            btn_yes.Foreground = brush_hex(u"#4ade80")
            btn_no.Background = th.brush_input()
            btn_no.Foreground = th.brush_fg_lo()
        else:
            btn_yes.Background = th.brush_input()
            btn_yes.Foreground = brush_hex(u"#64748b")
            btn_no.Background = brush_hex(u"#64748b", 64)
            btn_no.Foreground = brush_hex(u"#95b8cc")

    def _set(val):
        state[0] = bool(val)
        _apply_style()
        if on_change:
            on_change(state[0])

    def _yes(sender, args):
        _set(True)

    def _no(sender, args):
        _set(False)

    try:
        from System.Windows import RoutedEventHandler as _REH
        btn_yes.Click += _REH(_yes)
        btn_no.Click += _REH(_no)
    except Exception:
        pass

    _apply_style()
    panel.Children.Add(btn_yes)
    panel.Children.Add(btn_no)
    shell.IsEnabled = bool(enabled)
    if not enabled:
        shell.Opacity = 0.65
    return shell
