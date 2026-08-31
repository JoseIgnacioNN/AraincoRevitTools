# -*- coding: utf-8 -*-
"""
Elevación Eje — UI WPF con listado de ejes (Grid) del proyecto.

Revit 2024+ | pyRevit | IronPython 2.7 / 3.4

Selecciona ejes y escala; crea elevaciones con contorno e etiquetas.
"""

from __future__ import print_function

import os
import re

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from Autodesk.Revit.DB import FilteredElementCollector, Grid
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler, TaskDialog

from System import AppDomain, EventHandler
from System.Windows import (
    CornerRadius,
    FontWeights,
    HorizontalAlignment,
    RoutedEventHandler,
    Thickness,
    VerticalAlignment,
    WindowState,
)
from System.Windows.Controls import (
    Border,
    Button,
    CheckBox,
    Orientation,
    StackPanel,
    TextBlock,
    TextChangedEventHandler,
)
from System.Windows.Input import Cursors, Key, KeyEventHandler, MouseButtonEventHandler
from System.Windows.Markup import XamlReader
from System.Windows.Media import Brushes, Color, SolidColorBrush

from bimtools_ui_tokens import BTN_MANUAL
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_wpf_shell import build_simple_tool_xaml
from revit_wpf_window_position import (
    bind_center_wpf_on_revit_monitor,
    position_wpf_window_center_on_monitor,
    revit_main_hwnd,
)

_APPDOMAIN_WINDOW_KEY = u"Arainco_ElevacionEje_UI"
_APPDOMAIN_EVENT_KEY = u"Arainco_ElevacionEje_ExtEvent"
_APPDOMAIN_HANDLER_KEY = u"Arainco_ElevacionEje_Handler"
_TOOL_DIALOG_TITLE = u"Arainco: Elevación Eje"

# Denominadores View.Scale (mismo set que Vistas por usuario / categoría).
VIEW_SCALE_RATIOS = (50, 75, 100, 125, 150, 175, 200)
DEFAULT_VIEW_SCALE = 50

_BRUSH_SEG_ON_BG = SolidColorBrush(Color.FromArgb(0x24, 0x5B, 0xB8, 0xD4))
_BRUSH_SEG_ON_BD = SolidColorBrush(Color.FromRgb(0x5B, 0xB8, 0xD4))
_BRUSH_SEG_OFF_BG = SolidColorBrush(Color.FromRgb(0x07, 0x10, 0x18))
_BRUSH_SEG_OFF_BD = SolidColorBrush(Color.FromRgb(0x1E, 0x33, 0x44))
_BRUSH_FG_HI = SolidColorBrush(Color.FromRgb(0xE8, 0xF4, 0xF8))
_BRUSH_FG_MID = SolidColorBrush(Color.FromRgb(0x95, 0xB8, 0xCC))
_BRUSH_ROW_HOVER = SolidColorBrush(Color.FromArgb(0x28, 0x5B, 0xC0, 0xDE))

_FILTER_PLACEHOLDER = u"Filtrar por nombre…"

_BODY_XAML = u"""
<Grid>
  <Grid.RowDefinitions>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="*"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
  </Grid.RowDefinitions>
  <Grid Grid.Row="0" Margin="0,0,0,8">
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="Auto"/>
    </Grid.ColumnDefinitions>
    <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
      <TextBlock Text="Ejes" Foreground="#95B8CC" FontSize="11" FontWeight="SemiBold"
                 VerticalAlignment="Center"/>
      <TextBlock x:Name="TxtCount" Margin="10,0,0,0" VerticalAlignment="Center"
                 Foreground="#64748b" FontSize="11" Text="0 de 0 seleccionados"/>
    </StackPanel>
    <StackPanel Grid.Column="1" Orientation="Horizontal">
      <Button x:Name="BtnSelectAll" Content="Seleccionar todo" Margin="0,0,6,0"
              Style="{StaticResource BtnSelectOutline}" MinWidth="120" Padding="8,2"
              ToolTip="Marcar los ejes visibles del filtro"/>
      <Button x:Name="BtnSelectNone" Content="Ninguno"
              Style="{StaticResource BtnSelectOutline}" MinWidth="72" Padding="8,2"
              ToolTip="Desmarcar los ejes visibles del filtro"/>
    </StackPanel>
  </Grid>
  <Grid Grid.Row="1" Margin="0,0,0,8" MinHeight="30">
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="Auto"/>
    </Grid.ColumnDefinitions>
    <Grid Grid.Column="0">
      <TextBox x:Name="TxtFilter" MinHeight="30"
               Style="{StaticResource BimToolsTextBoxDark}"
               ToolTip="Filtrar ejes por nombre"
               VerticalContentAlignment="Center" Padding="10,4,10,4"/>
      <TextBlock x:Name="TxtFilterPh" Text="Filtrar por nombre…"
                 IsHitTestVisible="False" Foreground="#64748b" FontSize="11"
                 VerticalAlignment="Center" Margin="12,0,0,0"/>
    </Grid>
    <Button x:Name="BtnRefresh" Grid.Column="1" Content="Actualizar"
            Margin="8,0,0,0" Padding="8,2" MinWidth="88" FontSize="11"
            Style="{StaticResource BtnSelectOutline}"
            ToolTip="Volver a leer los ejes del documento"
            VerticalAlignment="Stretch"/>
  </Grid>
  <Border Grid.Row="2" Background="#050E18" BorderBrush="#21465C" BorderThickness="1"
          CornerRadius="4" MinHeight="200">
    <ScrollViewer x:Name="ScrEjes" VerticalScrollBarVisibility="Auto"
                  HorizontalScrollBarVisibility="Disabled"
                  Padding="2,2">
      <StackPanel x:Name="PanelEjes"/>
    </ScrollViewer>
  </Border>
  <StackPanel Grid.Row="3">
    <TextBlock x:Name="TxtFilterEmpty" Margin="0,6,0,0" Visibility="Collapsed"
               Foreground="#64748b" FontSize="10"
               Text="Ningún eje coincide con el filtro."/>
    <TextBlock x:Name="TxtOcultos" Margin="0,8,0,0" Visibility="Collapsed"
               Foreground="#95B8CC" FontSize="11" TextWrapping="Wrap"
               Text=""/>
  </StackPanel>
  <TextBlock Grid.Row="4" Margin="0,14,0,6" Text="Escala de vista"
             Foreground="#95B8CC" FontSize="11" FontWeight="SemiBold"/>
  <WrapPanel Grid.Row="5" x:Name="PanelEscala" Orientation="Horizontal"/>
</Grid>
"""

_FOOTER_LEADING_XAML = (
    u'<Button x:Name="BtnManual" Content="Manual" '
    u'Style="{{StaticResource BtnSelectOutline}}" '
    u'Background="{bg}" MinWidth="96" Padding="8,2" '
    u'ToolTip="Abrir manual de usuario" VerticalAlignment="Center"/>'
).format(bg=BTN_MANUAL)

_FOOTER_ACTIONS_XAML = u"""
<Button x:Name="BtnClose" Content="Cerrar" Margin="0,0,8,0"
        Style="{StaticResource BtnSelectOutline}" MinWidth="100"/>
<Button x:Name="BtnCrear" Content="Crear elevaciones"
        Style="{StaticResource BtnPrimary}" MinWidth="168"
        ToolTip="Crear elevaciones para los ejes marcados"/>
"""


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _natural_sort_key(text):
    """Clave de orden natural: 1, 2, 10… y letras/símbolos intercalados."""
    s = _as_unicode(text).strip().lower()
    if not s:
        return ((1, u""),)
    parts = []
    for chunk in re.split(r"(\d+)", s):
        if not chunk:
            continue
        if chunk.isdigit():
            try:
                parts.append((0, int(chunk)))
            except Exception:
                parts.append((0, 0))
        else:
            parts.append((1, chunk))
    return tuple(parts) if parts else ((1, s),)


def _mostrar_aviso(uiapp, instruction, content=u"", ok_text=u"Entendido"):
    hwnd = None
    try:
        if uiapp is not None:
            hwnd = revit_main_hwnd(uiapp)
    except Exception:
        pass
    try:
        from bimtools_instruction_dialog import show_message_dialog

        show_message_dialog(
            _TOOL_DIALOG_TITLE,
            instruction,
            content=content,
            ok_text=ok_text,
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass
    try:
        body = instruction
        if content:
            body = instruction + u"\n\n" + content
        TaskDialog.Show(_TOOL_DIALOG_TITLE, body)
    except Exception:
        pass


def _resolve_manual_path():
    """Ruta a ``manual_usuario.html`` en la carpeta del pushbutton."""
    candidates = []
    try:
        import bimtools_paths

        pb = bimtools_paths.get_pushbutton_dir()
        if pb:
            candidates.append(os.path.join(pb, u"manual_usuario.html"))
    except Exception:
        pass
    try:
        ext_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        for tab_name in os.listdir(ext_dir):
            if not tab_name.endswith(u".tab"):
                continue
            panel = os.path.join(ext_dir, tab_name, u"Creación de Vistas.panel")
            if not os.path.isdir(panel):
                panel = os.path.join(
                    ext_dir, tab_name, u"Creacion de Vistas.panel",
                )
            if not os.path.isdir(panel):
                continue
            for pb_name in os.listdir(panel):
                if u"ElevacionEje" not in pb_name:
                    continue
                candidates.append(
                    os.path.join(panel, pb_name, u"manual_usuario.html")
                )
    except Exception:
        pass
    seen = set()
    for path in candidates:
        try:
            ap = os.path.normpath(os.path.abspath(path))
        except Exception:
            continue
        if ap in seen:
            continue
        seen.add(ap)
        if os.path.isfile(ap):
            return ap
    return None


def _open_manual(uiapp):
    path = _resolve_manual_path()
    if not path:
        _mostrar_aviso(
            uiapp,
            u"No se encontró manual_usuario.html.",
            content=u"Debe estar en la carpeta del pushbutton de la herramienta.",
        )
        return
    try:
        os.startfile(path)
    except Exception as ex:
        _mostrar_aviso(
            uiapp,
            u"No se pudo abrir el manual.",
            content=_as_unicode(ex),
        )


def listar_ejes_modelo(document, excluir_claves=None):
    """
    ``Grid`` del documento, ordenados con orden natural por nombre.

    Args:
        document: Document
        excluir_claves: set opcional de nombres de eje en minúsculas
            (p. ej. ya elevados con el Building Section actual)

    Returns:
        lista de ``(nombre, Grid)``
    """
    ejes = []
    if document is None:
        return ejes
    exclude = set(excluir_claves or [])
    try:
        for g in FilteredElementCollector(document).OfClass(Grid):
            if g is None or not isinstance(g, Grid):
                continue
            try:
                nombre = _as_unicode(g.Name).strip()
            except Exception:
                nombre = u""
            if not nombre:
                try:
                    nombre = u"Id {0}".format(g.Id.IntegerValue)
                except Exception:
                    nombre = u"(sin nombre)"
            if exclude:
                try:
                    from elevacion_eje import _clave_nombre_eje

                    if _clave_nombre_eje(nombre) in exclude:
                        continue
                except Exception:
                    if nombre.lower() in exclude:
                        continue
            ejes.append((nombre, g))
    except Exception:
        pass
    try:
        ejes.sort(key=lambda t: _natural_sort_key(t[0]))
    except Exception:
        try:
            ejes.sort(key=lambda t: t[0].lower())
        except Exception:
            pass
    return ejes


def _claves_ejes_ya_elevados(document, vft):
    """Set de claves de ejes ya elevados con el ViewFamilyType actual."""
    if document is None or vft is None:
        return set()
    try:
        from elevacion_eje import nombres_ejes_ya_elevados

        return set(nombres_ejes_ya_elevados(document, vft) or [])
    except Exception:
        return set()


def _attach_revit_owner(win, uiapp):
    if win is None or uiapp is None:
        return
    try:
        from System.Windows.Interop import WindowInteropHelper

        hwnd = revit_main_hwnd(uiapp)
        if hwnd is not None:
            WindowInteropHelper(win).Owner = hwnd
    except Exception:
        pass


def _prepare_window(win, uiapp):
    if win is None:
        return
    try:
        hwnd = revit_main_hwnd(uiapp)
        bind_center_wpf_on_revit_monitor(win, hwnd)
        position_wpf_window_center_on_monitor(win, hwnd)
    except Exception:
        pass
    _attach_revit_owner(win, uiapp)


def _build_xaml():
    return build_simple_tool_xaml(
        title=_TOOL_DIALOG_TITLE,
        styles_xml=BIMTOOLS_DARK_STYLES_XML,
        body_xaml=_BODY_XAML,
        footer_leading_xaml=_FOOTER_LEADING_XAML,
        footer_actions_xaml=_FOOTER_ACTIONS_XAML,
        width=520,
        min_width=520,
        height=700,
        min_height=700,
        resize_mode=u"NoResize",
        size_to_content_height=False,
    )


def _get_active_window():
    try:
        win = AppDomain.CurrentDomain.GetData(_APPDOMAIN_WINDOW_KEY)
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
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, win)
    except Exception:
        pass


def _clear_active_window():
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, None)
    except Exception:
        pass


def _pin_external_event(ext_event, handler):
    """Mantener vivos ExternalEvent/handler tras cerrar la UI WPF."""
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_EVENT_KEY, ext_event)
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_HANDLER_KEY, handler)
    except Exception:
        pass


def _unpin_external_event():
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_EVENT_KEY, None)
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_HANDLER_KEY, None)
    except Exception:
        pass


def _visibility_collapsed():
    from System.Windows import Visibility

    return Visibility.Collapsed


def _visibility_visible():
    from System.Windows import Visibility

    return Visibility.Visible


def _sync_filter_placeholder(txt, txt_ph):
    """Muestra el watermark solo cuando el TextBox está vacío."""
    if txt_ph is None:
        return
    try:
        empty = not _as_unicode(txt.Text if txt is not None else u"").strip()
    except Exception:
        empty = True
    try:
        txt_ph.Visibility = (
            _visibility_visible() if empty else _visibility_collapsed()
        )
    except Exception:
        pass


def _filter_query(txt):
    if txt is None:
        return u""
    try:
        return _as_unicode(txt.Text).strip().lower()
    except Exception:
        return u""


def _style_crear_button(btn, enabled, n_sel=0):
    """CTA primario visible; deshabilitado se ve claramente apagado."""
    if btn is None:
        return
    try:
        btn.IsEnabled = bool(enabled)
    except Exception:
        pass
    try:
        btn.Content = _texto_btn_crear(n_sel)
    except Exception:
        pass
    try:
        if enabled:
            btn.Opacity = 1.0
            btn.Background = SolidColorBrush(Color.FromRgb(0x5B, 0xC0, 0xDE))
            btn.Foreground = SolidColorBrush(Color.FromRgb(0x0A, 0x1A, 0x2F))
            btn.BorderBrush = SolidColorBrush(Color.FromRgb(0x87, 0xD9, 0xEE))
        else:
            btn.Opacity = 0.38
            btn.Background = SolidColorBrush(Color.FromRgb(0x1E, 0x33, 0x44))
            btn.Foreground = SolidColorBrush(Color.FromRgb(0x64, 0x74, 0x8B))
            btn.BorderBrush = SolidColorBrush(Color.FromRgb(0x21, 0x46, 0x5C))
    except Exception:
        try:
            btn.Opacity = 1.0 if enabled else 0.4
        except Exception:
            pass


def escala_seleccionada(state_or_ratio):
    """
    Denominador de escala (``View.Scale``), p. ej. 50 → 1:50.

    Acepta un ``dict`` con clave ``scale_ratio`` o un ``int`` directo.
    """
    if isinstance(state_or_ratio, dict):
        try:
            return int(state_or_ratio.get(u"scale_ratio", DEFAULT_VIEW_SCALE))
        except Exception:
            return int(DEFAULT_VIEW_SCALE)
    try:
        return int(state_or_ratio)
    except Exception:
        return int(DEFAULT_VIEW_SCALE)


def _ejes_seleccionados(rows, ejes):
    """Lista de ``(nombre, Grid)`` según filas con CheckBox marcado."""
    out = []
    if not rows or not ejes:
        return out
    by_id = {}
    for nombre, grid in ejes:
        try:
            by_id[int(grid.Id.IntegerValue)] = (nombre, grid)
        except Exception:
            continue
    for row in rows:
        cb = row.get(u"cb")
        if cb is None:
            continue
        try:
            if not cb.IsChecked:
                continue
        except Exception:
            continue
        try:
            eid = int(row.get(u"eid"))
        except Exception:
            continue
        pair = by_id.get(eid)
        if pair is not None:
            out.append(pair)
    return out


def _fill_escalas(panel, state, on_changed):
    """Botones segmentados 1:N en ``PanelEscala``."""
    if panel is None:
        return
    try:
        panel.Children.Clear()
    except Exception:
        pass
    buttons = []
    state[u"scale_buttons"] = buttons
    if u"scale_ratio" not in state:
        state[u"scale_ratio"] = int(DEFAULT_VIEW_SCALE)

    def _on_click(sender, _args):
        try:
            state[u"scale_ratio"] = int(sender.Tag)
        except Exception:
            state[u"scale_ratio"] = int(DEFAULT_VIEW_SCALE)
        _apply_scale_styles(state)
        if on_changed is not None:
            on_changed()

    for ratio in VIEW_SCALE_RATIOS:
        btn = Button()
        btn.Content = u"1:{0}".format(ratio)
        btn.Tag = int(ratio)
        btn.Margin = Thickness(0, 0, 4, 4)
        btn.Padding = Thickness(8, 6, 8, 6)
        btn.MinWidth = 52
        btn.FontSize = 12
        btn.Cursor = Cursors.Hand
        try:
            btn.Click += RoutedEventHandler(_on_click)
        except Exception:
            pass
        panel.Children.Add(btn)
        buttons.append(btn)
    _apply_scale_styles(state)


def _apply_scale_styles(state):
    buttons = state.get(u"scale_buttons") or []
    try:
        current = int(state.get(u"scale_ratio", DEFAULT_VIEW_SCALE))
    except Exception:
        current = int(DEFAULT_VIEW_SCALE)
    for btn in buttons:
        on = False
        try:
            on = int(btn.Tag) == current
        except Exception:
            on = False
        try:
            if on:
                btn.Background = _BRUSH_SEG_ON_BG
                btn.BorderBrush = _BRUSH_SEG_ON_BD
                btn.Foreground = _BRUSH_FG_HI
                btn.FontWeight = FontWeights.SemiBold
            else:
                btn.Background = _BRUSH_SEG_OFF_BG
                btn.BorderBrush = _BRUSH_SEG_OFF_BD
                btn.Foreground = _BRUSH_FG_MID
                btn.FontWeight = FontWeights.Normal
            btn.BorderThickness = Thickness(1)
        except Exception:
            pass


def _make_group_header(text, extra_top=False):
    tb = TextBlock()
    tb.Text = _as_unicode(text)
    tb.FontSize = 10
    tb.FontWeight = FontWeights.SemiBold
    try:
        tb.Foreground = _BRUSH_FG_MID
    except Exception:
        pass
    top = 10 if extra_top else 4
    tb.Margin = Thickness(6, top, 6, 4)
    return tb


def _append_eje_row(panel, rows, nombre, grid, keep, on_check_changed, grupo):
    try:
        eid = int(grid.Id.IntegerValue)
    except Exception:
        eid = None

    cb = CheckBox()
    try:
        cb.Content = u""
    except Exception:
        cb.Content = None
    cb.IsChecked = eid in keep if eid is not None else False
    cb.Margin = Thickness(0, 0, 0, 0)
    cb.Padding = Thickness(0, 0, 0, 0)
    cb.Cursor = Cursors.Hand
    cb.VerticalAlignment = VerticalAlignment.Center
    cb.VerticalContentAlignment = VerticalAlignment.Center
    cb.HorizontalAlignment = HorizontalAlignment.Left
    cb.Tag = eid

    label = TextBlock()
    label.Text = nombre
    label.FontSize = 12
    label.Margin = Thickness(8, 0, 0, 0)
    label.Padding = Thickness(0, 0, 0, 0)
    label.VerticalAlignment = VerticalAlignment.Center
    label.Cursor = Cursors.Hand
    try:
        label.Foreground = _BRUSH_FG_HI
    except Exception:
        pass

    row_panel = StackPanel()
    row_panel.Orientation = Orientation.Horizontal
    row_panel.VerticalAlignment = VerticalAlignment.Center
    try:
        row_panel.Children.Add(cb)
        row_panel.Children.Add(label)
    except Exception:
        pass

    border = Border()
    border.Background = Brushes.Transparent
    border.CornerRadius = CornerRadius(2)
    border.Padding = Thickness(6, 3, 6, 3)
    border.Margin = Thickness(0, 0, 0, 0)
    border.Cursor = Cursors.Hand
    border.Child = row_panel
    border.Tag = eid

    row = {
        u"cb": cb,
        u"label": label,
        u"border": border,
        u"nombre": nombre,
        u"eid": eid,
        u"visible": True,
        u"hover": False,
        u"grupo": grupo,
    }

    def _make_toggle(r):
        def _on_row_down(sender, args):
            try:
                src = args.OriginalSource
            except Exception:
                src = None
            if isinstance(src, CheckBox):
                return
            try:
                r[u"cb"].IsChecked = not bool(r[u"cb"].IsChecked)
                if args is not None:
                    args.Handled = True
            except Exception:
                pass

        return _on_row_down

    try:
        border.MouseLeftButtonDown += MouseButtonEventHandler(_make_toggle(row))
    except Exception:
        pass

    def _make_hover(r):
        def _enter(sender, args):
            r[u"hover"] = True
            _style_row(r)

        def _leave(sender, args):
            r[u"hover"] = False
            _style_row(r)

        return _enter, _leave

    enter, leave = _make_hover(row)
    try:
        from System.Windows.Input import MouseEventHandler

        border.MouseEnter += MouseEventHandler(enter)
        border.MouseLeave += MouseEventHandler(leave)
    except Exception:
        pass

    if on_check_changed is not None:
        try:
            cb.Checked += RoutedEventHandler(on_check_changed)
            cb.Unchecked += RoutedEventHandler(on_check_changed)
        except Exception:
            pass

    _style_row(row)
    panel.Children.Add(border)
    rows.append(row)


def _cargar_lista(panel, document, state, on_check_changed, checked_ids=None):
    """
    Rellena filas clicables agrupadas (numéricos / letras).

    Omite ejes ya elevados con el Building Section de ``state['vft']``.

    ``state['rows']`` = lista de dicts ``{cb, label, border, nombre, eid, visible, grupo}``.
    """
    if panel is None:
        state[u"ejes"] = []
        state[u"rows"] = []
        state[u"group_headers"] = []
        state[u"ocultos_count"] = 0
        return []

    try:
        panel.Children.Clear()
    except Exception:
        pass

    keep = set(checked_ids or [])
    ya = _claves_ejes_ya_elevados(document, state.get(u"vft"))
    todos = listar_ejes_modelo(document)
    if ya:
        try:
            from elevacion_eje import _clave_nombre_eje

            ejes = [
                (n, g) for n, g in todos
                if _clave_nombre_eje(n) not in ya
            ]
            state[u"ocultos_count"] = len(todos) - len(ejes)
        except Exception:
            ejes = listar_ejes_modelo(document, excluir_claves=ya)
            state[u"ocultos_count"] = max(0, len(todos) - len(ejes))
    else:
        ejes = todos
        state[u"ocultos_count"] = 0
    rows = []
    headers = []
    state[u"ejes"] = ejes
    state[u"rows"] = rows
    state[u"group_headers"] = headers

    nums = [(n, g) for n, g in ejes if _eje_es_numerico(n)]
    lets = [(n, g) for n, g in ejes if not _eje_es_numerico(n)]
    grupos = []
    if nums:
        grupos.append((u"num", u"Numéricos", nums))
    if lets:
        grupos.append((u"let", u"Letras", lets))

    for i_g, (gkey, glabel, pares) in enumerate(grupos):
        hdr = _make_group_header(glabel, extra_top=(i_g > 0))
        panel.Children.Add(hdr)
        headers.append({u"block": hdr, u"grupo": gkey})
        for nombre, grid in pares:
            _append_eje_row(
                panel, rows, nombre, grid, keep, on_check_changed, gkey,
            )

    return ejes


def _apply_filter(state, query, txt_empty=None):
    q = _as_unicode(query or u"").strip().lower()
    visible_n = 0
    visible_by_grupo = {}
    for row in state.get(u"rows") or []:
        nombre = _as_unicode(row.get(u"nombre")).lower()
        show = (not q) or (q in nombre)
        row[u"visible"] = show
        border = row.get(u"border")
        if border is not None:
            try:
                border.Visibility = (
                    _visibility_visible() if show else _visibility_collapsed()
                )
            except Exception:
                pass
        if show:
            visible_n += 1
            g = row.get(u"grupo")
            if g:
                visible_by_grupo[g] = int(visible_by_grupo.get(g, 0) or 0) + 1
    for hdr in state.get(u"group_headers") or []:
        block = hdr.get(u"block")
        if block is None:
            continue
        n_g = int(visible_by_grupo.get(hdr.get(u"grupo"), 0) or 0)
        try:
            block.Visibility = (
                _visibility_visible() if n_g > 0 else _visibility_collapsed()
            )
        except Exception:
            pass
    if txt_empty is not None:
        try:
            total = len(state.get(u"rows") or [])
            if total > 0 and visible_n == 0 and q:
                txt_empty.Visibility = _visibility_visible()
            else:
                txt_empty.Visibility = _visibility_collapsed()
        except Exception:
            pass
    return visible_n


def _set_checks_visible(state, checked):
    valor = True if checked else False
    for row in state.get(u"rows") or []:
        if not row.get(u"visible", True):
            continue
        cb = row.get(u"cb")
        if cb is None:
            continue
        try:
            cb.IsChecked = valor
        except Exception:
            pass


def _checked_ids(state):
    ids = set()
    for row in state.get(u"rows") or []:
        cb = row.get(u"cb")
        if cb is None:
            continue
        try:
            if not cb.IsChecked:
                continue
        except Exception:
            continue
        try:
            ids.add(int(row.get(u"eid")))
        except Exception:
            pass
    return ids


def _texto_contador(seleccionados, total, visible_n=None, filter_on=False):
    n = len(seleccionados)
    try:
        t = int(total)
    except Exception:
        t = 0
    if filter_on and visible_n is not None:
        return u"{0} visibles · {1} seleccionados".format(int(visible_n), n)
    if t <= 0:
        return u"0 de 0 seleccionados"
    return u"{0} de {1} seleccionados".format(n, t)


def _texto_estado(scale_ratio):
    try:
        return u"Escala 1:{0}".format(int(scale_ratio))
    except Exception:
        return u"Escala 1:{0}".format(DEFAULT_VIEW_SCALE)


def _texto_btn_crear(n):
    try:
        k = int(n)
    except Exception:
        k = 0
    if k <= 0:
        return u"Crear elevaciones"
    if k == 1:
        return u"Crear 1 elevación"
    return u"Crear {0} elevaciones".format(k)


def _texto_subtitulo(tipo):
    t = _as_unicode(tipo).strip()
    if not t:
        return u""
    return u"Building Section: {0}".format(t)


def _eje_es_numerico(nombre):
    s = _as_unicode(nombre).strip()
    if not s:
        return False
    try:
        return s[0].isdigit()
    except Exception:
        return False


def _row_is_checked(row):
    if not row:
        return False
    cb = row.get(u"cb")
    if cb is None:
        return False
    try:
        return bool(cb.IsChecked)
    except Exception:
        return False


def _style_row(row):
    if not row:
        return
    border = row.get(u"border")
    if border is None:
        return
    try:
        if _row_is_checked(row):
            border.Background = _BRUSH_SEG_ON_BG
        elif row.get(u"hover"):
            border.Background = _BRUSH_ROW_HOVER
        else:
            border.Background = Brushes.Transparent
    except Exception:
        pass


def _style_all_rows(state):
    for row in state.get(u"rows") or []:
        _style_row(row)


class _CrearElevacionesHandler(IExternalEventHandler):
    """Ejecuta la creación en el hilo de Revit (UI ya cerrada + ProgressBar)."""

    def __init__(self):
        self.ejes = []
        self.scale_ratio = DEFAULT_VIEW_SCALE
        self.uiapp_for_dialog = None

    def Execute(self, uiapp):
        from elevacion_eje import ejecutar_crear_elevaciones

        uidoc = None
        try:
            uidoc = uiapp.ActiveUIDocument
        except Exception:
            uidoc = None
        ok, msg, _vistas = (False, u"", [])
        try:
            ok, msg, _vistas = ejecutar_crear_elevaciones(
                uidoc,
                self.ejes,
                self.scale_ratio,
            )
        except Exception as ex:
            ok = False
            msg = _as_unicode(ex)
        finally:
            _unpin_external_event()

        host = self.uiapp_for_dialog or uiapp
        if ok:
            return
        try:
            _mostrar_aviso(
                host,
                u"No se completó la creación.",
                content=_as_unicode(msg),
            )
        except Exception:
            pass

    def GetName(self):
        return u"Arainco: Elevación Eje"


def run(revit):
    """Punto de entrada pyRevit: muestra la ventana con el listado de ejes."""
    existing = _get_active_window()
    if existing is not None:
        try:
            if existing.WindowState == WindowState.Minimized:
                existing.WindowState = WindowState.Normal
        except Exception:
            pass
        try:
            existing.Activate()
            existing.Focus()
        except Exception:
            pass
        _mostrar_aviso(revit, u"La herramienta ya esta en ejecucion.")
        return

    uidoc = revit.ActiveUIDocument
    if uidoc is None:
        _mostrar_aviso(revit, u"No hay documento activo.")
        return

    try:
        from elevacion_eje import (
            resolver_tipo_building_section,
            vista_permitida,
        )
    except Exception:
        vista_permitida = None
        resolver_tipo_building_section = None

    if vista_permitida is not None:
        ok_vista, msg_vista = vista_permitida(uidoc.ActiveView)
        if not ok_vista:
            _mostrar_aviso(revit, msg_vista)
            return

    doc = uidoc.Document
    tipo_section_label = u""
    vft0 = None
    if resolver_tipo_building_section is not None:
        vft0, sf0, err0 = resolver_tipo_building_section(doc, uidoc.ActiveView)
        if vft0 is None:
            _mostrar_aviso(
                revit,
                u"No se pudo determinar el tipo Building Section.",
                content=_as_unicode(err0),
            )
            return
        try:
            tipo_section_label = _as_unicode(vft0.Name).strip()
        except Exception:
            tipo_section_label = _as_unicode(sf0).strip()

    win = XamlReader.Parse(_build_xaml())
    txt_subtitle = win.FindName(u"TxtSubtitle")
    txt_status = win.FindName(u"TxtStatus")
    txt_count = win.FindName(u"TxtCount")
    txt_filter = win.FindName(u"TxtFilter")
    txt_filter_ph = win.FindName(u"TxtFilterPh")
    txt_filter_empty = win.FindName(u"TxtFilterEmpty")
    txt_ocultos = win.FindName(u"TxtOcultos")
    panel_ejes = win.FindName(u"PanelEjes")
    panel_escala = win.FindName(u"PanelEscala")
    btn_refresh = win.FindName(u"BtnRefresh")
    btn_close = win.FindName(u"BtnClose")
    btn_crear = win.FindName(u"BtnCrear")
    btn_all = win.FindName(u"BtnSelectAll")
    btn_none = win.FindName(u"BtnSelectNone")
    btn_manual = win.FindName(u"BtnManual")

    state = {
        u"ejes": [],
        u"rows": [],
        u"scale_ratio": int(DEFAULT_VIEW_SCALE),
        u"scale_buttons": [],
        u"busy": False,
        u"tipo_section": tipo_section_label,
        u"vft": vft0,
        u"ocultos_count": 0,
        u"group_headers": [],
    }

    crear_handler = _CrearElevacionesHandler()
    crear_event = ExternalEvent.Create(crear_handler)

    # Subtítulo: tipo Building Section resuelto desde Section Filter.
    if txt_subtitle is not None:
        if tipo_section_label:
            try:
                txt_subtitle.Text = _texto_subtitulo(tipo_section_label)
                txt_subtitle.Visibility = _visibility_visible()
            except Exception:
                pass
        else:
            try:
                txt_subtitle.Visibility = _visibility_collapsed()
            except Exception:
                try:
                    txt_subtitle.Text = u""
                except Exception:
                    pass

    if txt_filter_ph is not None:
        try:
            txt_filter_ph.Text = _FILTER_PLACEHOLDER
        except Exception:
            pass
    _sync_filter_placeholder(txt_filter, txt_filter_ph)
    _style_crear_button(btn_crear, False)

    def _set_status(text):
        if txt_status is not None:
            try:
                txt_status.Text = _as_unicode(text)
            except Exception:
                pass

    def _actualizar_ocultos_hint():
        if txt_ocultos is None:
            return
        n_ocultos = int(state.get(u"ocultos_count") or 0)
        total_disp = len(state.get(u"ejes") or [])
        try:
            if n_ocultos <= 0:
                txt_ocultos.Visibility = _visibility_collapsed()
                txt_ocultos.Text = u""
                return
            if total_disp <= 0:
                txt_ocultos.Text = (
                    u"Todos los ejes ya tienen elevación con este tipo."
                )
            elif n_ocultos == 1:
                txt_ocultos.Text = u"1 eje ya tiene elevación con este tipo."
            else:
                txt_ocultos.Text = (
                    u"{0} ejes ya tienen elevación con este tipo."
                ).format(n_ocultos)
            txt_ocultos.Visibility = _visibility_visible()
        except Exception:
            pass

    def _actualizar_subtitulo():
        if txt_subtitle is None:
            return
        tipo = _as_unicode(state.get(u"tipo_section") or u"").strip()
        try:
            if tipo:
                txt_subtitle.Text = _texto_subtitulo(tipo)
                txt_subtitle.Visibility = _visibility_visible()
            else:
                txt_subtitle.Visibility = _visibility_collapsed()
        except Exception:
            pass

    def _actualizar_estado(_sender=None, _args=None):
        seleccionados = _ejes_seleccionados(state.get(u"rows"), state.get(u"ejes"))
        total = len(state.get(u"ejes") or [])
        scale = escala_seleccionada(state)
        q = _filter_query(txt_filter)
        visible_n = 0
        for row in state.get(u"rows") or []:
            if row.get(u"visible", True):
                visible_n += 1
        if txt_count is not None:
            try:
                txt_count.Text = _texto_contador(
                    seleccionados,
                    total,
                    visible_n=visible_n,
                    filter_on=bool(q),
                )
            except Exception:
                pass
        _actualizar_ocultos_hint()
        _style_all_rows(state)
        if not state.get(u"busy"):
            if txt_status is not None:
                try:
                    txt_status.Text = _texto_estado(scale)
                except Exception:
                    pass
            _style_crear_button(
                btn_crear,
                len(seleccionados) > 0,
                n_sel=len(seleccionados),
            )

    def _reload(checked_ids=None):
        uidoc_now = revit.ActiveUIDocument
        doc_now = uidoc_now.Document if uidoc_now is not None else None
        active_now = uidoc_now.ActiveView if uidoc_now is not None else None
        if resolver_tipo_building_section is not None and doc_now is not None:
            vft_now, sf_now, err_now = resolver_tipo_building_section(
                doc_now, active_now,
            )
            if vft_now is not None:
                state[u"vft"] = vft_now
                try:
                    state[u"tipo_section"] = _as_unicode(vft_now.Name).strip()
                except Exception:
                    state[u"tipo_section"] = _as_unicode(sf_now).strip()
                _actualizar_subtitulo()
            elif err_now and not state.get(u"vft"):
                _mostrar_aviso(
                    revit,
                    u"No se pudo determinar el tipo Building Section.",
                    content=_as_unicode(err_now),
                )
        _cargar_lista(
            panel_ejes,
            doc_now,
            state,
            on_check_changed=_actualizar_estado,
            checked_ids=checked_ids,
        )
        _apply_filter(state, _filter_query(txt_filter), txt_filter_empty)
        _actualizar_estado()

    def _refresh(_sender=None, _args=None):
        _reload(checked_ids=_checked_ids(state))

    def _on_filter_changed(_sender=None, _args=None):
        _sync_filter_placeholder(txt_filter, txt_filter_ph)
        q = _filter_query(txt_filter)
        _apply_filter(state, q, txt_filter_empty)
        _actualizar_estado()

    def _select_all(_sender=None, _args=None):
        _set_checks_visible(state, True)
        _actualizar_estado()

    def _select_none(_sender=None, _args=None):
        _set_checks_visible(state, False)
        _actualizar_estado()

    def _on_crear(_sender, _args):
        if state.get(u"busy"):
            return
        seleccionados = _ejes_seleccionados(state.get(u"rows"), state.get(u"ejes"))
        if not seleccionados:
            _mostrar_aviso(revit, u"Marca al menos un eje para continuar.")
            return
        scale = escala_seleccionada(state)
        state[u"busy"] = True
        _style_crear_button(btn_crear, False)

        crear_handler.ejes = list(seleccionados)
        crear_handler.scale_ratio = scale
        crear_handler.uiapp_for_dialog = revit
        _pin_external_event(crear_event, crear_handler)
        try:
            crear_event.Raise()
        except Exception as ex:
            state[u"busy"] = False
            _unpin_external_event()
            _actualizar_estado()
            _mostrar_aviso(
                revit,
                u"No se pudo iniciar la creación.",
                content=_as_unicode(ex),
            )
            return
        try:
            win.Close()
        except Exception:
            pass

    def _on_manual(_sender, _args):
        _open_manual(revit)

    def _on_close(_sender, _args):
        try:
            win.Close()
        except Exception:
            pass

    def _on_key_down(_sender, args):
        if args.Key == Key.Escape:
            _on_close(None, None)

    def _on_closed(_sender, _args):
        _clear_active_window()

    _fill_escalas(panel_escala, state, on_changed=_actualizar_estado)
    _reload(checked_ids=set())

    if txt_filter is not None:
        try:
            txt_filter.TextChanged += TextChangedEventHandler(_on_filter_changed)
        except Exception:
            pass

    if btn_all is not None:
        btn_all.Click += RoutedEventHandler(_select_all)
    if btn_none is not None:
        btn_none.Click += RoutedEventHandler(_select_none)
    if btn_refresh is not None:
        btn_refresh.Click += RoutedEventHandler(_refresh)
    if btn_close is not None:
        btn_close.Click += RoutedEventHandler(_on_close)
    if btn_crear is not None:
        btn_crear.Click += RoutedEventHandler(_on_crear)
    if btn_manual is not None:
        btn_manual.Click += RoutedEventHandler(_on_manual)
    win.KeyDown += KeyEventHandler(_on_key_down)
    win.Closed += EventHandler(_on_closed)

    _set_active_window(win)
    _prepare_window(win, revit)
    try:
        win.Show()
    except Exception:
        _clear_active_window()
        raise
