# -*- coding: utf-8 -*-
"""Rail derecho — Mallas en muros (patrones UI/UX de Armado vigas)."""

from __future__ import print_function

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import (
    CornerRadius,
    FontWeights,
    GridLength,
    GridUnitType,
    HorizontalAlignment,
    TextWrapping,
    Thickness,
    VerticalAlignment,
)
from System.Windows.Controls import (
    Border,
    Button,
    CheckBox,
    ColumnDefinition,
    Dock,
    DockPanel,
    Grid,
    Orientation,
    StackPanel,
    TextBlock,
)
from System.Windows.Input import Cursors
from System.Windows.Media import SolidColorBrush, Color

try:
    from bimtools_ui_tokens import (
        ACCENT_PRIMARY,
        BG_APP,
        BG_PANEL,
        BG_PANEL_ELEVATED,
        BORDER,
        FG_BODY,
        FG_MUTED,
        FG_TITLE,
    )
except Exception:
    ACCENT_PRIMARY = u"#5BC0DE"
    BG_APP = u"#071018"
    BG_PANEL = u"#0a1620"
    BG_PANEL_ELEVATED = u"#0E1B32"
    BORDER = u"#21465C"
    FG_BODY = u"#95B8CC"
    FG_MUTED = u"#64748b"
    FG_TITLE = u"#E8F4F8"

# Acentos por orientación (armado_vigas/ui/rail_cards.py RAIL_TABS)
ACCENT_VERTICAL = u"#22d3ee"
ACCENT_HORIZONTAL = u"#fb7185"
BORDER_MUTED = u"#2d4455"

# Tipografía rail (armado_vigas/ui/typography.py)
CTRL_FONT_PX = 10.0
LABEL_FONT_PX = 9.0
META_FONT_PX = 8.0
TITLE_FONT_PX = 12.0

MALLAS_RAIL_TABS = (
    (u"vert", u"VERT", ACCENT_VERTICAL),
    (u"hor", u"HOR", ACCENT_HORIZONTAL),
)

MALLAS_CONCRETE_GRADES = (u"G25", u"G35", u"G45")
MALLAS_CONCRETE_GRADE_DEFAULT = u"G25"


def normalize_mallas_concrete_grade(grade):
    """``G25`` / ``G35`` / ``G45``; inválido o vacío → G25."""
    try:
        s = unicode(grade).strip().upper()
    except Exception:
        try:
            s = str(grade or u"").strip().upper()
        except Exception:
            s = u""
    if s in MALLAS_CONCRETE_GRADES:
        return s
    return MALLAS_CONCRETE_GRADE_DEFAULT


def _dosif_active_colors(grade):
    """(bg, border, fg) hex para chip activo de dosificación."""
    g = normalize_mallas_concrete_grade(grade)
    if g == u"G35":
        return (u"#4c1d95", u"#a78bfa", u"#ede9fe")
    if g == u"G45":
        return (u"#9d174d", u"#f472b6", u"#fce7f3")
    return (u"#0c4a6e", u"#38bdf8", u"#e0f2fe")


def _brush_hex(hex_color, alpha=255):
    s = (hex_color or u"#000000").lstrip(u"#")
    if len(s) != 6:
        s = u"071018"
    r = int(s[0:2], 16)
    g = int(s[2:4], 16)
    b = int(s[4:6], 16)
    return SolidColorBrush(Color.FromArgb(int(alpha), r, g, b))


def _corner(border, radius=4.0):
    try:
        border.CornerRadius = CornerRadius(float(radius))
    except Exception:
        pass
    return border


def make_role_badge(label, accent_hex=ACCENT_PRIMARY, bg_alpha=31):
    badge = Border()
    badge.Padding = Thickness(4, 1, 4, 1)
    badge.Background = _brush_hex(accent_hex, bg_alpha)
    badge.BorderBrush = _brush_hex(accent_hex, 107)
    badge.BorderThickness = Thickness(1)
    _corner(badge, 3.0)
    tb = TextBlock()
    tb.Text = (label or u"").upper()
    tb.FontSize = 10.0
    tb.FontWeight = FontWeights.Bold
    tb.Foreground = _brush_hex(accent_hex)
    badge.Child = tb
    return badge


def field_stack(label, control):
    """Etiqueta arriba + control (rail estrecho, estilo vigas)."""
    sp = StackPanel()
    sp.Orientation = Orientation.Vertical
    sp.HorizontalAlignment = HorizontalAlignment.Stretch
    tb = TextBlock()
    tb.Text = label or u""
    tb.Foreground = _brush_hex(FG_BODY)
    tb.FontSize = LABEL_FONT_PX
    tb.FontWeight = FontWeights.SemiBold
    tb.Margin = Thickness(0, 0, 0, 3)
    sp.Children.Add(tb)
    try:
        control.HorizontalAlignment = HorizontalAlignment.Stretch
        control.VerticalAlignment = VerticalAlignment.Center
        control.Margin = Thickness(0)
    except Exception:
        pass
    sp.Children.Add(control)
    return sp


def selection_meta_chip(text):
    """Chip meta scaneable (estilo build_config_viga_header)."""
    chip = Border()
    chip.Background = _brush_hex(BG_PANEL_ELEVATED)
    chip.BorderBrush = _brush_hex(BORDER)
    chip.BorderThickness = Thickness(1)
    chip.Padding = Thickness(8, 4, 8, 4)
    chip.Margin = Thickness(0, 0, 0, 0)
    chip.HorizontalAlignment = HorizontalAlignment.Stretch
    _corner(chip, 3.0)
    tb = TextBlock()
    tb.Text = text or u"1 muro seleccionado"
    tb.Foreground = _brush_hex(FG_BODY)
    tb.FontSize = CTRL_FONT_PX
    tb.TextWrapping = TextWrapping.Wrap
    chip.Child = tb
    return chip, tb


def build_config_muro_header(meta_text, phase_enabled=True, on_phase_toggle=None):
    """
    Cabecera global del rail: título + badge MALLA + toggle fase + chip meta.
    """
    block = Border()
    block.Margin = Thickness(0, 0, 0, 10)
    block.Padding = Thickness(10)
    block.Background = _brush_hex(BG_PANEL)
    block.BorderBrush = _brush_hex(BORDER)
    block.BorderThickness = Thickness(1)
    _corner(block, 4.0)

    root = StackPanel()
    root.Orientation = Orientation.Vertical

    title_row = DockPanel()
    title_row.LastChildFill = True
    title_row.Margin = Thickness(0, 0, 0, 6)

    right = StackPanel()
    right.Orientation = Orientation.Horizontal
    right.VerticalAlignment = VerticalAlignment.Center

    chk = CheckBox()
    chk.IsChecked = bool(phase_enabled)
    chk.ToolTip = u"Activar o desactivar el armado de malla"
    chk.VerticalAlignment = VerticalAlignment.Center
    chk.Margin = Thickness(8, 0, 0, 0)
    right.Children.Add(chk)

    badge = make_role_badge(u"MALLA", ACCENT_PRIMARY)
    badge.Margin = Thickness(8, 0, 0, 0)
    right.Children.Add(badge)

    DockPanel.SetDock(right, Dock.Right)
    title_row.Children.Add(right)

    title = TextBlock()
    title.Text = u"Configuración muro"
    title.Foreground = _brush_hex(FG_TITLE)
    title.FontSize = TITLE_FONT_PX
    title.FontWeight = FontWeights.SemiBold
    title.VerticalAlignment = VerticalAlignment.Center
    title_row.Children.Add(title)
    root.Children.Add(title_row)

    chip, chip_tb = selection_meta_chip(meta_text)
    root.Children.Add(chip)

    block.Child = root

    if on_phase_toggle is not None:
        try:
            from System.Windows import RoutedEventHandler

            def _chk(s, e):
                try:
                    on_phase_toggle(bool(chk.IsChecked))
                except Exception:
                    pass

            chk.Checked += RoutedEventHandler(_chk)
            chk.Unchecked += RoutedEventHandler(_chk)
        except Exception:
            pass

    return block, chip_tb, chk


def build_rail_tabs(active_id, on_select, tabs=None, phase_enabled=True):
    """Pestañas VERT / HOR a ancho completo (estilo build_rail_tabs vigas)."""
    tabs = tabs or MALLAS_RAIL_TABS
    grid = Grid()
    grid.Margin = Thickness(0, 0, 0, 10)
    grid.HorizontalAlignment = HorizontalAlignment.Stretch
    n = len(tabs)
    for _ in range(n):
        cd = ColumnDefinition()
        cd.Width = GridLength(1.0, GridUnitType.Star)
        grid.ColumnDefinitions.Add(cd)

    for i, (tab_id, label, color) in enumerate(tabs):
        sel = active_id == tab_id
        btn = Button()
        btn.Content = label + (u"" if phase_enabled else u" · off")
        btn.Padding = Thickness(4, 6, 4, 6)
        btn.FontSize = 11.0
        btn.FontWeight = FontWeights.Bold if sel else FontWeights.Normal
        btn.Cursor = Cursors.Hand
        btn.HorizontalAlignment = HorizontalAlignment.Stretch
        btn.HorizontalContentAlignment = HorizontalAlignment.Center
        btn.Margin = Thickness(0, 0, 4 if i < n - 1 else 0, 0)
        try:
            btn.MinWidth = 0.0
        except Exception:
            pass
        if sel:
            btn.BorderBrush = _brush_hex(color)
            btn.BorderThickness = Thickness(1.5)
            btn.Background = _brush_hex(BG_PANEL_ELEVATED)
            btn.Foreground = _brush_hex(FG_TITLE)
            btn.Opacity = 1.0 if phase_enabled else 0.55
        else:
            btn.BorderBrush = _brush_hex(BORDER_MUTED)
            btn.BorderThickness = Thickness(1)
            btn.Background = _brush_hex(BG_APP)
            btn.Foreground = _brush_hex(FG_MUTED)
            btn.Opacity = 0.38 if phase_enabled else 0.22

        def _click(sender, args, tid=tab_id):
            if on_select is not None:
                try:
                    on_select(tid)
                except Exception:
                    pass

        try:
            from System.Windows import RoutedEventHandler

            btn.Click += RoutedEventHandler(_click)
        except Exception:
            pass
        Grid.SetColumn(btn, i)
        grid.Children.Add(btn)
    return grid


def build_segmented_control(options, current_id, on_select, accent_hex=ACCENT_PRIMARY):
    """
    Botones segmentados (estilo dosificación G25/G35/G45).

    ``options``: iterable de (id, label).
    Retorna (root, buttons_dict, hint_tb) donde buttons_dict[id] = Button.
    """
    root = StackPanel()
    root.Orientation = Orientation.Vertical

    grid = Grid()
    grid.HorizontalAlignment = HorizontalAlignment.Stretch
    n = len(options)
    for _ in range(n):
        cd = ColumnDefinition()
        cd.Width = GridLength(1.0, GridUnitType.Star)
        grid.ColumnDefinitions.Add(cd)

    buttons = {}
    for i, (opt_id, label) in enumerate(options):
        active = opt_id == current_id
        btn = Button()
        btn.Content = label
        btn.Tag = opt_id
        btn.Height = 28.0
        btn.Padding = Thickness(2, 4, 2, 4)
        btn.FontSize = CTRL_FONT_PX
        btn.FontWeight = FontWeights.Bold if active else FontWeights.Normal
        btn.Margin = Thickness(0 if i == 0 else 3, 0, 0, 0)
        btn.HorizontalAlignment = HorizontalAlignment.Stretch
        btn.Cursor = Cursors.Hand
        if active:
            btn.Background = _brush_hex(accent_hex, 40)
            btn.BorderBrush = _brush_hex(accent_hex)
            btn.Foreground = _brush_hex(accent_hex)
            btn.BorderThickness = Thickness(1.2)
        else:
            btn.Background = _brush_hex(u"#0a1520")
            btn.BorderBrush = _brush_hex(BORDER)
            btn.Foreground = _brush_hex(FG_BODY)
            btn.BorderThickness = Thickness(1)

        def _on(s, e, oid=opt_id):
            if on_select is not None:
                try:
                    on_select(oid)
                except Exception:
                    pass

        try:
            from System.Windows import RoutedEventHandler

            btn.Click += RoutedEventHandler(_on)
        except Exception:
            pass
        buttons[opt_id] = btn
        Grid.SetColumn(btn, i)
        grid.Children.Add(btn)

    root.Children.Add(grid)
    return root, buttons


def refresh_segmented_control(buttons, current_id, accent_hex=ACCENT_PRIMARY):
    """Actualiza estilo activo de botones segmentados."""
    for opt_id, btn in (buttons or {}).items():
        active = opt_id == current_id
        try:
            btn.FontWeight = FontWeights.Bold if active else FontWeights.Normal
            if active:
                btn.Background = _brush_hex(accent_hex, 40)
                btn.BorderBrush = _brush_hex(accent_hex)
                btn.Foreground = _brush_hex(accent_hex)
                btn.BorderThickness = Thickness(1.2)
            else:
                btn.Background = _brush_hex(u"#0a1520")
                btn.BorderBrush = _brush_hex(BORDER)
                btn.Foreground = _brush_hex(FG_BODY)
                btn.BorderThickness = Thickness(1)
        except Exception:
            pass


def _hdr_row(title, badge, subtitle=None):
    dock = DockPanel()
    dock.LastChildFill = True
    dock.Margin = Thickness(0, 0, 0, 8 if subtitle else 6)
    if badge is not None:
        badge.Margin = Thickness(8, 0, 0, 0)
        badge.VerticalAlignment = VerticalAlignment.Center
        DockPanel.SetDock(badge, Dock.Right)
        dock.Children.Add(badge)
    tb = TextBlock()
    tb.Text = title or u""
    tb.Foreground = _brush_hex(FG_TITLE)
    tb.FontSize = TITLE_FONT_PX
    tb.FontWeight = FontWeights.SemiBold
    tb.VerticalAlignment = VerticalAlignment.Center
    tb.TextWrapping = TextWrapping.Wrap
    dock.Children.Add(tb)
    if subtitle:
        wrap = StackPanel()
        wrap.Children.Add(dock)
        sub = TextBlock()
        sub.Text = subtitle
        sub.Foreground = _brush_hex(FG_MUTED)
        sub.FontSize = CTRL_FONT_PX
        sub.TextWrapping = TextWrapping.Wrap
        sub.Margin = Thickness(0, 0, 0, 8)
        wrap.Children.Add(sub)
        return wrap
    return dock


def build_config_card(title, badge_label, accent_hex, subtitle, content, enabled=True):
    """Card del rail: borde acento, cabecera + cuerpo."""
    shell = Border()
    shell.Margin = Thickness(0, 0, 0, 0)
    shell.Padding = Thickness(10)
    shell.Background = _brush_hex(BG_PANEL)
    shell.BorderBrush = _brush_hex(accent_hex if enabled else BORDER)
    shell.BorderThickness = Thickness(1.5 if enabled else 1)
    shell.Opacity = 1.0 if enabled else 0.55
    _corner(shell, 4.0)

    root = StackPanel()
    root.Orientation = Orientation.Vertical
    badge = make_role_badge(badge_label, accent_hex) if badge_label else None
    root.Children.Add(_hdr_row(title, badge, subtitle))

    body = StackPanel()
    body.Orientation = Orientation.Vertical
    if content is not None:
        try:
            content.Margin = Thickness(0)
        except Exception:
            pass
        body.Children.Add(content)
    root.Children.Add(body)
    shell.Child = root
    return shell


def build_terminacion_section(title, options, current_id, on_select, hint_text, accent_hex):
    """
    Bloque extremo superior/inferior: título + segmentado + hint.
    Retorna (block, buttons, hint_tb).
    """
    block = StackPanel()
    block.Orientation = Orientation.Vertical
    block.Margin = Thickness(0, 10, 0, 0)

    hdr = TextBlock()
    hdr.Text = title or u""
    hdr.Foreground = _brush_hex(FG_BODY)
    hdr.FontSize = LABEL_FONT_PX
    hdr.FontWeight = FontWeights.SemiBold
    hdr.Margin = Thickness(0, 0, 0, 5)
    block.Children.Add(hdr)

    seg, buttons = build_segmented_control(options, current_id, on_select, accent_hex)
    block.Children.Add(seg)

    hint = TextBlock()
    hint.Text = hint_text or u""
    hint.Foreground = _brush_hex(FG_MUTED)
    hint.FontSize = META_FONT_PX
    hint.TextWrapping = TextWrapping.Wrap
    hint.Margin = Thickness(0, 6, 0, 0)
    block.Children.Add(hint)

    return block, buttons, hint


def refresh_dosificacion_segment(buttons, current_grade):
    """Actualiza estilo G25/G35/G45 del bloque dosificación."""
    cur = normalize_mallas_concrete_grade(current_grade)
    for g, btn in (buttons or {}).items():
        active = normalize_mallas_concrete_grade(g) == cur
        try:
            btn.FontWeight = FontWeights.Bold if active else FontWeights.Normal
            if active:
                bg, br, fg = _dosif_active_colors(cur)
                btn.Background = _brush_hex(bg)
                btn.BorderBrush = _brush_hex(br)
                btn.Foreground = _brush_hex(fg)
                btn.BorderThickness = Thickness(1.2)
            else:
                btn.Background = _brush_hex(u"#0a1520")
                btn.BorderBrush = _brush_hex(BORDER)
                btn.Foreground = _brush_hex(FG_BODY)
                btn.BorderThickness = Thickness(1)
        except Exception:
            pass


def build_dosificacion_segment(current_grade, on_select):
    """
    Segmentado G25 | G35 | G45 (tablas traslape/empotramiento/pata L).
    Retorna (block, buttons_dict).
    """
    current = normalize_mallas_concrete_grade(current_grade)
    block = StackPanel()
    block.Orientation = Orientation.Vertical
    block.Margin = Thickness(0, 0, 0, 10)

    lbl = TextBlock()
    lbl.Text = u"Dosificación del hormigón · lote"
    lbl.Foreground = _brush_hex(FG_BODY)
    lbl.FontSize = LABEL_FONT_PX
    lbl.FontWeight = FontWeights.SemiBold
    lbl.Margin = Thickness(0, 0, 0, 5)
    block.Children.Add(lbl)

    grid = Grid()
    grid.HorizontalAlignment = HorizontalAlignment.Stretch
    for _ in range(len(MALLAS_CONCRETE_GRADES)):
        cd = ColumnDefinition()
        cd.Width = GridLength(1.0, GridUnitType.Star)
        grid.ColumnDefinitions.Add(cd)

    buttons = {}
    for i, lab in enumerate(MALLAS_CONCRETE_GRADES):
        active = lab == current
        btn = Button()
        btn.Content = lab
        btn.Tag = lab
        btn.Height = 28.0
        btn.Padding = Thickness(2, 4, 2, 4)
        btn.FontSize = CTRL_FONT_PX
        btn.FontWeight = FontWeights.Bold if active else FontWeights.Normal
        btn.Margin = Thickness(0 if i == 0 else 3, 0, 0, 0)
        btn.HorizontalAlignment = HorizontalAlignment.Stretch
        btn.Cursor = Cursors.Hand
        btn.ToolTip = (
            u"Dosificación {0} · tablas de traslape, empotramiento y pata L "
            u"para todo el lote"
        ).format(lab)
        if active:
            bg, br, fg = _dosif_active_colors(lab)
            btn.Background = _brush_hex(bg)
            btn.BorderBrush = _brush_hex(br)
            btn.Foreground = _brush_hex(fg)
            btn.BorderThickness = Thickness(1.2)
        else:
            btn.Background = _brush_hex(u"#0a1520")
            btn.BorderBrush = _brush_hex(BORDER)
            btn.Foreground = _brush_hex(FG_BODY)
            btn.BorderThickness = Thickness(1)

        def _on(s, e, gid=lab):
            if on_select is not None:
                try:
                    on_select(gid)
                except Exception:
                    pass

        try:
            from System.Windows import RoutedEventHandler

            btn.Click += RoutedEventHandler(_on)
        except Exception:
            pass
        buttons[lab] = btn
        Grid.SetColumn(btn, i)
        grid.Children.Add(btn)

    block.Children.Add(grid)

    hint = TextBlock()
    hint.Text = (
        u"Afecta empotramientos verticales (L por Ø), patas L y referencias "
        u"del alzado según tablas G25/G35/G45."
    )
    hint.Foreground = _brush_hex(FG_MUTED)
    hint.FontSize = META_FONT_PX
    hint.TextWrapping = TextWrapping.Wrap
    hint.Margin = Thickness(0, 6, 0, 0)
    block.Children.Add(hint)

    return block, buttons
