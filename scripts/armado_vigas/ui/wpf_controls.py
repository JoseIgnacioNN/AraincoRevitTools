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

from armado_vigas.domain.constants import (
    BAR_COUNT_MIN,
    BAR_COUNT_MAX,
    LONG_DIAM_OPTS,
    ESTRIBO_SPACING_MIN,
    ESTRIBO_SPACING_MAX,
)
from armado_vigas.ui import typography as typo
from armado_vigas.ui import theme as th
from armado_vigas.ui.net_ui import brush_hex as _brush_hex_cached
from armado_vigas.ui.net_ui import to_net_int_list

# Cache de Style WPF por (id(win), resource_key) — TryFindResource es caro en pyRevit.
_COMBO_STYLE_CACHE = {}


def clear_combo_style_cache():
    _COMBO_STYLE_CACHE.clear()


def _combo_style(win, stretch=False):
    """Resuelve Style Combo / ComboStretch una sola vez por ventana."""
    if win is None:
        return None
    key = (id(win), u"ComboStretch" if stretch else u"Combo")
    if key in _COMBO_STYLE_CACHE:
        return _COMBO_STYLE_CACHE[key]
    st = None
    try:
        st = win.TryFindResource(u"ComboStretch" if stretch else u"Combo")
        if st is None and stretch:
            st = win.TryFindResource(u"Combo")
    except Exception:
        st = None
    _COMBO_STYLE_CACHE[key] = st
    return st


def brush_hex(hx, alpha=255):
    """Brush cacheado + Freezado (compat API pública)."""
    return _brush_hex_cached(hx, alpha)


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
    return typo.CTRL_HEIGHT_PX if compact else 26.0


def _apply_fixed_control_height(ctrl, compact=True, stretch=False):
    """Altura legible sin recortar el texto del template Combo/TextBox.

    Con stretch: MinHeight (no Height fijo) para respetar ComboStretch Auto.
    Con tamaño fijo: Height = CTRL_HEIGHT_PX y padding vertical 0 (el theme
    suma 3+3 y con Height 22 recortaba ø8 / 200).
    """
    if ctrl is None:
        return
    h = float(_combo_height(compact))
    try:
        ctrl.VerticalContentAlignment = VerticalAlignment.Center
    except Exception:
        pass
    if stretch:
        try:
            from System.Windows import FrameworkElement

            ctrl.ClearValue(FrameworkElement.HeightProperty)
        except Exception:
            pass
        try:
            ctrl.MinHeight = h
        except Exception:
            pass
        try:
            # Padding vertical 0: el borde del template ya da aire.
            ctrl.Padding = Thickness(6, 0, 6, 0)
        except Exception:
            pass
        return
    try:
        ctrl.Height = h
        ctrl.MinHeight = h
    except Exception:
        pass
    try:
        ctrl.Padding = Thickness(6, 0, 6, 0)
    except Exception:
        pass


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


def make_diam_combo(win, value, diam_opts=None, on_change=None, compact=False, enabled=True, stretch=False):
    opts = diam_opts or LONG_DIAM_OPTS
    # Asegura cur y list compacta (muchos RebarBarType hinchan el bridge).
    try:
        cur = int(value or opts[0])
    except Exception:
        cur = int(opts[0]) if opts else 16
    try:
        base = [int(d) for d in (opts or LONG_DIAM_OPTS)]
    except Exception:
        base = list(LONG_DIAM_OPTS)
    if cur not in base:
        base.append(cur)
    # Preferir tabla corta estándar + valor actual si el catálogo es muy grande.
    if len(base) > 14:
        prefer = set(int(x) for x in LONG_DIAM_OPTS)
        prefer.add(cur)
        kept = [d for d in base if d in prefer]
        if cur not in kept:
            kept.append(cur)
        base = sorted(set(kept))
    opts = base
    # Normaliza a List[int] .NET para no dejar tuplas Python en el puente.
    try:
        net_opts = to_net_int_list(opts)
    except Exception:
        net_opts = None
    cb = ComboBox()
    try:
        st = _combo_style(win, stretch=stretch)
        if st is not None:
            cb.Style = st
    except Exception:
        pass
    if stretch:
        try:
            from System.Windows import FrameworkElement

            cb.ClearValue(FrameworkElement.WidthProperty)
            cb.ClearValue(FrameworkElement.MinWidthProperty)
            cb.ClearValue(FrameworkElement.MaxWidthProperty)
        except Exception:
            pass
        cb.MinWidth = 0.0
        try:
            cb.HorizontalAlignment = HorizontalAlignment.Stretch
        except Exception:
            pass
    else:
        diam_w = 56.0 if compact else 72.0
        cb.Width = diam_w
        cb.MinWidth = diam_w
        cb.MaxWidth = diam_w
    _apply_fixed_control_height(cb, compact=compact, stretch=stretch)
    cb.FontSize = _combo_font_size(compact)
    cb.Margin = Thickness(1, 0, 0, 0) if compact and not stretch else Thickness(0)
    loading = [True]
    seq = net_opts if net_opts is not None else opts
    for d in seq:
        di = int(d)
        item = _add_combo_item(cb, u"ø{0}".format(di), di)
        if di == cur:
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


def _is_digits_only_text(text):
    if text is None:
        return False
    for ch in unicode(text):
        if not ch.isdigit():
            return False
    return True


def make_spacing_input(win, value, on_change=None, compact=False, enabled=True, width=None, stretch=False):
    """Campo @ espaciado estribos (mm): entrada manual solo dígitos."""
    from System.Windows.Controls import TextBox
    from System.Windows.Input import Key, KeyEventHandler, TextCompositionEventHandler

    tb = TextBox()
    try:
        if win is not None:
            st = win.TryFindResource(u"BimToolsTextBoxDark")
            if st is not None:
                tb.Style = st
    except Exception:
        pass

    if stretch:
        try:
            from System.Windows import FrameworkElement

            tb.ClearValue(FrameworkElement.WidthProperty)
            tb.ClearValue(FrameworkElement.MinWidthProperty)
            tb.ClearValue(FrameworkElement.MaxWidthProperty)
        except Exception:
            pass
        tb.MinWidth = 0.0
        try:
            tb.HorizontalAlignment = HorizontalAlignment.Stretch
        except Exception:
            pass
        tb.Margin = Thickness(0)
    else:
        sp_w = width if width is not None else 52.0
        tb.Width = sp_w
        tb.MinWidth = sp_w
        tb.MaxWidth = sp_w
        tb.Margin = Thickness(1, 0, 0, 0) if compact else Thickness(0)

    _apply_fixed_control_height(tb, compact=compact, stretch=stretch)
    tb.FontSize = _combo_font_size(compact)
    tb.FontWeight = FontWeights.Bold
    tb.TextAlignment = TextAlignment.Center
    try:
        tb.VerticalContentAlignment = VerticalAlignment.Center
    except Exception:
        pass

    try:
        cur = int(round(value))
    except Exception:
        cur = int(ESTRIBO_SPACING_MIN)
    cur = max(int(ESTRIBO_SPACING_MIN), min(int(ESTRIBO_SPACING_MAX), cur))
    loading = [True]
    last_val = [cur]
    tb.Text = unicode(cur)

    def _preview_text_input(sender, e):
        try:
            e.Handled = not _is_digits_only_text(e.Text)
        except Exception:
            pass

    def _on_pasting(sender, e):
        try:
            from System.Windows import DataObject
            txt = None
            if e.SourceDataObject is not None and e.SourceDataObject.GetDataPresent("Text"):
                txt = e.SourceDataObject.GetData("Text")
            if not _is_digits_only_text(txt):
                e.CancelCommand()
        except Exception:
            try:
                e.CancelCommand()
            except Exception:
                pass

    def _commit():
        if loading[0]:
            return
        try:
            raw = (tb.Text or u"").strip()
            if not raw:
                n = last_val[0]
            else:
                n = int(raw)
            n = max(int(ESTRIBO_SPACING_MIN), min(int(ESTRIBO_SPACING_MAX), n))
            tb.Text = unicode(n)
            last_val[0] = n
            if on_change:
                on_change(n)
        except Exception:
            tb.Text = unicode(last_val[0])

    def _lost_focus(sender, args):
        _commit()

    def _key_down(sender, e):
        try:
            if e.Key == Key.Enter:
                _commit()
        except Exception:
            pass

    try:
        tb.PreviewTextInput += TextCompositionEventHandler(_preview_text_input)
    except Exception:
        pass
    try:
        from System.Windows import DataObject, DataObjectPastingEventHandler
        DataObject.AddPastingHandler(tb, DataObjectPastingEventHandler(_on_pasting))
    except Exception:
        pass
    try:
        from System.Windows import RoutedEventHandler as _REH
        tb.LostFocus += _REH(_lost_focus)
        tb.KeyDown += KeyEventHandler(_key_down)
    except Exception:
        pass

    loading[0] = False
    tb.IsEnabled = bool(enabled)
    if not enabled:
        tb.Opacity = 0.65
    return tb


def make_spacing_stepper(win, value, on_change, compact=False):
    return make_stepper(
        win,
        value,
        ESTRIBO_SPACING_MIN,
        ESTRIBO_SPACING_MAX,
        25,
        on_change,
        compact=compact,
    )


def make_string_combo(win, options, value, on_change, compact=False, stretch=True):
    from armado_vigas.ui.net_ui import to_net_string_list

    cb = ComboBox()
    try:
        st = _combo_style(win, stretch=bool(stretch or compact))
        if st is not None:
            cb.Style = st
    except Exception:
        pass
    if compact or stretch:
        try:
            from System.Windows import FrameworkElement

            cb.ClearValue(FrameworkElement.WidthProperty)
            cb.ClearValue(FrameworkElement.MaxWidthProperty)
        except Exception:
            pass
        cb.MinWidth = 0.0
        _apply_fixed_control_height(cb, compact=True, stretch=True)
        cb.FontSize = _combo_font_size(True)
        try:
            cb.HorizontalAlignment = HorizontalAlignment.Stretch
        except Exception:
            pass
    else:
        cb.MinWidth = 120.0
        _apply_fixed_control_height(cb, compact=False, stretch=False)
        cb.FontSize = 11.0
    cur = value
    loading = [True]
    # ComboBoxItem.Tag = str .NET (evita tipos Python dinámicos en selección).
    try:
        net_opts = to_net_string_list(options or [])
    except Exception:
        net_opts = options or []
    for opt in net_opts:
        try:
            s = unicode(opt)
        except NameError:
            s = str(opt)
        item = _add_combo_item(cb, s, s)
        if s == cur or opt == cur:
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


def make_int_combo(
    win,
    value,
    min_v,
    max_v,
    on_change=None,
    compact=False,
    enabled=True,
    step=1,
    stretch=True,
    format_item=None,
):
    """ComboBox de enteros en [min_v, max_v] (step 1 por defecto).

    Sustituye steppers ▲/▼ para cantidades (capas, n barras). Tag = int .NET.
    """
    try:
        lo = int(min_v)
        hi = int(max_v)
        step_i = max(1, int(step or 1))
    except Exception:
        lo, hi, step_i = 1, 3, 1
    if hi < lo:
        lo, hi = hi, lo
    try:
        cur = int(value)
    except Exception:
        cur = lo
    cur = max(lo, min(hi, cur))

    cb = ComboBox()
    try:
        # No reutilizar el nombre de step: el Style no puede entrar en range().
        style = _combo_style(win, stretch=stretch)
        if style is not None:
            cb.Style = style
    except Exception:
        pass

    if compact:
        cb.FontSize = _combo_font_size(True)
    else:
        cb.FontSize = _combo_font_size(False)
    if stretch:
        try:
            cb.MinWidth = 0.0
            cb.HorizontalAlignment = HorizontalAlignment.Stretch
        except Exception:
            pass
        try:
            from System.Windows import FrameworkElement

            cb.ClearValue(FrameworkElement.WidthProperty)
            cb.ClearValue(FrameworkElement.MaxWidthProperty)
        except Exception:
            pass
    else:
        cb.MinWidth = 48.0 if compact else 64.0
        cb.Width = 48.0 if compact else 64.0
        try:
            cb.MaxWidth = cb.Width
            cb.HorizontalAlignment = HorizontalAlignment.Left
        except Exception:
            pass
    _apply_fixed_control_height(cb, compact=compact, stretch=stretch)

    loading = [True]
    opts = list(range(lo, hi + 1, step_i))
    if cur not in opts:
        opts.append(cur)
        opts = sorted(set(opts))
    try:
        from armado_vigas.ui.net_ui import to_net_int_list

        net_opts = to_net_int_list(opts)
    except Exception:
        net_opts = opts

    for d in net_opts:
        di = int(d)
        if format_item is not None:
            try:
                text = format_item(di)
            except Exception:
                text = unicode(di)
        else:
            try:
                text = unicode(di)
            except NameError:
                text = str(di)
        item = _add_combo_item(cb, text, di)
        if di == cur:
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


def make_capas_combo(win, value, on_change, compact=False, stretch=False):
    """Cantidad de capas longitudinales (ComboBox, no stepper)."""
    from armado_vigas.domain.constants import CAPAS_MIN, CAPAS_MAX

    return make_int_combo(
        win,
        value,
        CAPAS_MIN,
        CAPAS_MAX,
        on_change,
        compact=compact,
        stretch=stretch,
        format_item=lambda n: u"{0}".format(n),
    )


def make_capas_stepper(win, value, on_change, compact=False):
    """Alias legacy → combo de capas."""
    return make_capas_combo(win, value, on_change, compact=compact, stretch=False)


def make_bar_count_combo(win, value, on_change=None, compact=False, enabled=True, stretch=False):
    """Cantidad de barras por capa (ComboBox)."""
    return make_int_combo(
        win,
        value,
        BAR_COUNT_MIN,
        BAR_COUNT_MAX,
        on_change,
        compact=compact,
        enabled=enabled,
        stretch=stretch,
    )


def make_yesno_toggle(win, value, on_change, compact=False, enabled=True, label=None):
    """Interruptor booleano (track + thumb), sin texto Sí/No junto al switch.

    Clic en el control para alternar. ``label`` opcional a la derecha
    (p. ej. Ini, Traslapo) — no es estado on/off.
    """
    from System.Windows import CornerRadius
    from System.Windows.Input import Cursors, MouseButtonEventHandler
    from System.Windows.Media import TranslateTransform

    host = Border()
    host.Background = brush_hex(u"#000000", 1)
    host.BorderThickness = Thickness(0)
    host.Padding = Thickness(0)
    host.Cursor = Cursors.Hand
    host.VerticalAlignment = VerticalAlignment.Center
    host.HorizontalAlignment = (
        HorizontalAlignment.Right if compact else HorizontalAlignment.Left
    )
    host.IsEnabled = bool(enabled)
    if not enabled:
        host.Opacity = 0.55

    row = StackPanel()
    row.Orientation = Orientation.Horizontal
    row.VerticalAlignment = VerticalAlignment.Center
    host.Child = row

    track_w = 30.0 if compact else 34.0
    track_h = 16.0 if compact else 18.0
    thumb_sz = 12.0 if compact else 14.0
    thumb_margin = 2.0
    thumb_on_x = track_w - thumb_sz - thumb_margin * 2.0

    track = Border()
    track.Width = track_w
    track.Height = track_h
    try:
        track.CornerRadius = CornerRadius(track_h * 0.5)
    except Exception:
        pass
    track.BorderThickness = Thickness(1)
    track.VerticalAlignment = VerticalAlignment.Center
    track.ClipToBounds = True
    track.SnapsToDevicePixels = True
    track.IsHitTestVisible = False

    thumb_xform = TranslateTransform(0.0, 0.0)
    thumb = Border()
    thumb.Width = thumb_sz
    thumb.Height = thumb_sz
    try:
        thumb.CornerRadius = CornerRadius(thumb_sz * 0.5)
    except Exception:
        pass
    thumb.Background = brush_hex(u"#e8f4f8")
    thumb.HorizontalAlignment = HorizontalAlignment.Left
    thumb.Margin = Thickness(thumb_margin, 0, 0, 0)
    thumb.VerticalAlignment = VerticalAlignment.Center
    thumb.RenderTransform = thumb_xform
    thumb.SnapsToDevicePixels = True
    thumb.IsHitTestVisible = False
    track.Child = thumb

    extra_lbl = None
    if label is not None and unicode(label).strip():
        extra_lbl = TextBlock()
        extra_lbl.Text = unicode(label)
        extra_lbl.FontSize = typo.CTRL_FONT_PX
        extra_lbl.FontWeight = FontWeights.SemiBold
        extra_lbl.VerticalAlignment = VerticalAlignment.Center
        extra_lbl.Margin = Thickness(6, 0, 0, 0)
        extra_lbl.Foreground = th.brush_fg_mid()
        extra_lbl.IsHitTestVisible = False

    state = [bool(value)]

    def _apply_visual():
        on = state[0]
        if on:
            track.Background = brush_hex(u"#22d3ee", 72)
            track.BorderBrush = brush_hex(u"#22d3ee", 200)
            thumb.Background = brush_hex(u"#e8f4f8")
            try:
                thumb_xform.X = thumb_on_x
            except Exception:
                pass
        else:
            track.Background = brush_hex(u"#122636")
            track.BorderBrush = th.brush_border()
            thumb.Background = brush_hex(u"#95b8cc")
            try:
                thumb_xform.X = 0.0
            except Exception:
                pass
        if extra_lbl is not None:
            extra_lbl.Foreground = th.brush_fg_hi() if on else th.brush_fg_lo()

    def _set(val, notify=True):
        state[0] = bool(val)
        _apply_visual()
        if notify and on_change:
            on_change(state[0])

    def _toggle(sender, args):
        try:
            if not host.IsEnabled:
                return
        except Exception:
            pass
        _set(not state[0], notify=True)
        try:
            args.Handled = True
        except Exception:
            pass

    try:
        from System.Windows.Input import MouseButtonEventHandler

        host.MouseLeftButtonUp += MouseButtonEventHandler(_toggle)
    except Exception:
        pass

    _apply_visual()
    row.Children.Add(track)
    if extra_lbl is not None:
        row.Children.Add(extra_lbl)

    try:
        host.Tag = {u"set": lambda v: _set(v, notify=False), u"get": lambda: state[0]}
    except Exception:
        pass
    return host
