# -*- coding: utf-8 -*-
"""UI WPF — Vistas por Usuario (shell visual alineado a Elevación Eje)."""

from __future__ import print_function

try:
    unicode
except NameError:
    unicode = str

import os

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from System import AppDomain, EventHandler
from System.Windows import (
    CornerRadius,
    FontWeights,
    GridLength,
    GridUnitType,
    HorizontalAlignment,
    RoutedEventHandler,
    TextTrimming,
    Thickness,
    VerticalAlignment,
)
from System.Windows.Controls import (
    Border,
    Button,
    CheckBox,
    ColumnDefinition,
    ComboBoxItem,
    Grid,
    Orientation,
    SelectionChangedEventHandler,
    StackPanel,
    TextBlock,
)
from System.Windows.Input import Cursors, MouseButtonEventHandler
from System.Windows.Markup import XamlReader
from System.Windows.Media import Brushes, Color, SolidColorBrush
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent, TaskDialog
from Autodesk.Revit.DB import FilteredElementCollector, Level

from bimtools_ui_tokens import BTN_MANUAL
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_wpf_shell import build_simple_tool_xaml
from revit_wpf_window_position import (
    bind_center_wpf_on_revit_monitor,
    position_wpf_window_center_on_monitor,
    revit_main_hwnd,
)

from vistas_por_usuario.constants import TRANSACTION_TITLE, VIEW_SCALE_RATIOS
from vistas_por_usuario import singleton
from vistas_por_usuario.people import (
    display_to_code,
    invalidate_modeladores_cache,
    load_modeladores,
    personas_paths,
)

_DIALOG_TITLE = TRANSACTION_TITLE
_WINDOW_TITLE = u"Arainco: Vistas por usuario"
_DEFAULT_VIEW_SCALE = 50
_APPDOMAIN_EVENT_KEY = u"Arainco_VistasPorUsuario_ExtEvent"
_APPDOMAIN_HANDLER_KEY = u"Arainco_VistasPorUsuario_Handler"

_BRUSH_SEG_ON_BG = SolidColorBrush(Color.FromArgb(0x24, 0x5B, 0xB8, 0xD4))
_BRUSH_SEG_ON_BD = SolidColorBrush(Color.FromRgb(0x5B, 0xB8, 0xD4))
_BRUSH_SEG_OFF_BG = SolidColorBrush(Color.FromRgb(0x07, 0x10, 0x18))
_BRUSH_SEG_OFF_BD = SolidColorBrush(Color.FromRgb(0x1E, 0x33, 0x44))
_BRUSH_FG_HI = SolidColorBrush(Color.FromRgb(0xE8, 0xF4, 0xF8))
_BRUSH_FG_MID = SolidColorBrush(Color.FromRgb(0x95, 0xB8, 0xCC))
_BRUSH_ROW_HOVER = SolidColorBrush(Color.FromArgb(0x28, 0x5B, 0xC0, 0xDE))

_BODY_XAML = u"""
<Grid>
  <Grid.RowDefinitions>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="*"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
  </Grid.RowDefinitions>

  <StackPanel Grid.Row="0" Margin="0,0,0,12">
    <TextBlock Text="Modelador" Foreground="#95B8CC"
               FontSize="11" FontWeight="SemiBold" Margin="0,0,0,4"/>
    <Grid>
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <ComboBox x:Name="CmbUsuario" Grid.Column="0"
                Style="{StaticResource ComboStretch}" IsEditable="False"
                ToolTip="Modelador: el código (p. ej. A.A.U) va a Subclasificacion y a los nombres de vista"/>
      <Border x:Name="ChipClasificacion" Grid.Column="1" Margin="6,0,0,0"
              Padding="8,0" MinWidth="56"
              Background="#245BB8D4" BorderBrush="#5BB8D4" BorderThickness="1"
              CornerRadius="4" VerticalAlignment="Stretch"
              ToolTip="Abreviación: una sola generación por este código en el documento">
        <TextBlock x:Name="TxtCodigo" Text="—" Foreground="#E8F4F8" FontSize="11"
                   FontFamily="Consolas" FontWeight="SemiBold"
                   VerticalAlignment="Center" HorizontalAlignment="Center"/>
      </Border>
      <Button x:Name="BtnGestionarPersonas" Grid.Column="2"
              Content="Personas" Margin="6,0,0,0" Padding="8,2" MinWidth="88"
              FontSize="11" Style="{StaticResource BtnSelectOutline}"
              VerticalAlignment="Stretch"
              ToolTip="Gestionar personas.json en el servidor de incidencias"/>
    </Grid>
  </StackPanel>

  <Grid Grid.Row="1" Margin="0,0,0,8">
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="Auto"/>
    </Grid.ColumnDefinitions>
    <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
      <TextBlock Text="Niveles" Foreground="#95B8CC" FontSize="11" FontWeight="SemiBold"
                 VerticalAlignment="Center"/>
      <TextBlock x:Name="TxtNivelCount" Margin="10,0,0,0" VerticalAlignment="Center"
                 Foreground="#64748b" FontSize="11" Text="0 de 0"/>
    </StackPanel>
    <StackPanel Grid.Column="1" Orientation="Horizontal">
      <Button x:Name="BtnSelAll" Content="Seleccionar todo" Margin="0,0,6,0"
              Style="{StaticResource BtnSelectOutline}" MinWidth="120" Padding="8,2"
              ToolTip="Marcar todos los niveles"/>
      <Button x:Name="BtnSelNone" Content="Ninguno"
              Style="{StaticResource BtnSelectOutline}" MinWidth="72" Padding="8,2"
              ToolTip="Desmarcar todos los niveles"/>
    </StackPanel>
  </Grid>

  <Border Grid.Row="2" Background="#050E18" BorderBrush="#21465C" BorderThickness="1"
          CornerRadius="4" MinHeight="200">
    <ScrollViewer VerticalScrollBarVisibility="Auto"
                  HorizontalScrollBarVisibility="Disabled"
                  Padding="2,2">
      <StackPanel x:Name="PanelNiveles"/>
    </ScrollViewer>
  </Border>

  <TextBlock Grid.Row="3" Margin="0,14,0,6" Text="Escala de vista"
             Foreground="#95B8CC" FontSize="11" FontWeight="SemiBold"/>
  <WrapPanel Grid.Row="4" x:Name="PanelEscala" Orientation="Horizontal"/>
</Grid>
"""

_FOOTER_HINT = (
    u"También crea plantillas y tipos Detail/Sección del modelador."
)

_FOOTER_LEADING_XAML = (
    u'<Button x:Name="BtnManual" Content="Manual" '
    u'Style="{{StaticResource BtnSelectOutline}}" '
    u'Background="{bg}" MinWidth="96" Padding="8,2" '
    u'ToolTip="Abrir manual de usuario" VerticalAlignment="Center"/>'
).format(bg=BTN_MANUAL)

_FOOTER_ACTIONS_XAML = u"""
<Button x:Name="BtnCancelar" Content="Cerrar" Margin="0,0,8,0"
        Style="{StaticResource BtnSelectOutline}" MinWidth="100"/>
<Button x:Name="BtnIniciar" Content="Crear Vistas"
        Style="{StaticResource BtnPrimary}" MinWidth="150"
        ToolTip="Cielo + Piso por nivel marcado. Si este modelador ya tiene vistas 02_TRABAJO, no se vuelve a generar"/>
"""


def _collect_levels_sorted(doc):
    """Niveles del documento ordenados por elevación."""
    levels = list(FilteredElementCollector(doc).OfClass(Level))
    levels.sort(key=lambda lv: lv.Elevation)
    return levels


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _format_level_elev(doc, elevation):
    """Cota del nivel para la columna derecha de la lista."""
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId

        meters = UnitUtils.ConvertFromInternalUnits(elevation, UnitTypeId.Meters)
        return u"{0:+.2f} m".format(float(meters))
    except Exception:
        try:
            return u"{0:+.2f}".format(float(elevation))
        except Exception:
            return u""


def _texto_contador(n_sel, total, n_plantas):
    try:
        n = int(n_sel)
        t = int(total)
        p = int(n_plantas)
    except Exception:
        return u"0 de 0"
    if t <= 0:
        return u"0 de 0"
    if n <= 0:
        return u"{0} de {1} seleccionados".format(n, t)
    return u"{0} de {1} · {2} plantas".format(n, t, p)


def mostrar_aviso(uiapp, instruction, content=u"", ok_text=u"Entendido"):
    """Diálogo informativo WPF (estilo BIMTools). Respaldo: TaskDialog nativo."""
    hwnd = None
    try:
        if uiapp is not None:
            hwnd = revit_main_hwnd(uiapp)
    except Exception:
        pass
    try:
        from bimtools_instruction_dialog import show_message_dialog

        show_message_dialog(
            _DIALOG_TITLE,
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
        TaskDialog.Show(_DIALOG_TITLE, body)
    except Exception:
        pass


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
        d = os.path.dirname(os.path.abspath(__file__))
        for _ in range(16):
            for tab_name in os.listdir(d):
                if not tab_name.endswith(u".tab"):
                    continue
                panel = os.path.join(d, tab_name, u"Creación de Vistas.panel")
                if not os.path.isdir(panel):
                    panel = os.path.join(
                        d, tab_name, u"Creacion de Vistas.panel",
                    )
                if not os.path.isdir(panel):
                    continue
                for pb_name in os.listdir(panel):
                    if u"VistasUsuario" not in pb_name:
                        continue
                    candidates.append(
                        os.path.join(panel, pb_name, u"manual_usuario.html")
                    )
            parent = os.path.dirname(d)
            if parent == d:
                break
            d = parent
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
        mostrar_aviso(
            uiapp,
            u"No se encontró manual_usuario.html.",
            content=u"Debe estar en la carpeta del pushbutton de la herramienta.",
        )
        return
    try:
        os.startfile(path)
    except Exception as ex:
        mostrar_aviso(
            uiapp,
            u"No se pudo abrir el manual.",
            content=_as_unicode(ex),
        )


def _build_xaml():
    return build_simple_tool_xaml(
        title=_WINDOW_TITLE,
        styles_xml=BIMTOOLS_DARK_STYLES_XML,
        body_xaml=_BODY_XAML,
        footer_leading_xaml=_FOOTER_LEADING_XAML,
        footer_actions_xaml=_FOOTER_ACTIONS_XAML,
        footer_hint_xaml=_FOOTER_HINT,
        width=520,
        min_width=520,
        height=700,
        min_height=700,
        resize_mode=u"NoResize",
        size_to_content_height=False,
    )


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


def _style_crear_button(btn, enabled):
    """CTA primario visible; deshabilitado se ve claramente apagado."""
    if btn is None:
        return
    try:
        btn.IsEnabled = bool(enabled)
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


class _LevelCheck(object):
    def __init__(self, level, checkbox, border):
        self.level = level
        self.checkbox = checkbox
        self.border = border
        self.hover = False


def _row_is_checked(item):
    if item is None or item.checkbox is None:
        return False
    try:
        return bool(item.checkbox.IsChecked)
    except Exception:
        return False


def _style_row(item):
    if item is None or item.border is None:
        return
    try:
        if _row_is_checked(item):
            item.border.Background = _BRUSH_SEG_ON_BG
        elif item.hover:
            item.border.Background = _BRUSH_ROW_HOVER
        else:
            item.border.Background = Brushes.Transparent
    except Exception:
        pass


class _CreateUserViewsHandler(IExternalEventHandler):
    """Ejecuta la creación en el hilo de Revit (UI ya cerrada)."""

    def __init__(self):
        self.request = None
        self.uiapp_for_dialog = None

    def Execute(self, uiapp):
        req = self.request
        self.request = None
        host = self.uiapp_for_dialog or uiapp
        if req is None:
            _unpin_external_event()
            return

        from vistas_por_usuario.service import (
            VistasPorUsuarioError,
            create_user_views,
            format_success_dialog,
            validate_user_views_not_exist,
        )

        try:
            uidoc = uiapp.ActiveUIDocument
            if uidoc is None:
                mostrar_aviso(host, u"No hay documento activo.")
                return

            doc = uidoc.Document
            ok, msg = validate_user_views_not_exist(doc, req.usuario_code)
            if not ok:
                mostrar_aviso(host, msg)
                return

            result = create_user_views(doc, req)
            instruction, content = format_success_dialog(
                result, req.usuario_display, req.usuario_code
            )
            mostrar_aviso(host, instruction, content, ok_text=u"Entendido")
        except VistasPorUsuarioError as ex:
            mostrar_aviso(host, str(ex))
        except Exception as ex:
            mostrar_aviso(
                host,
                u"Error al crear vistas.",
                content=u"{}".format(ex),
            )
        finally:
            _unpin_external_event()

    def GetName(self):
        return TRANSACTION_TITLE


class VistasPorUsuarioWindow(object):
    def __init__(self, doc, uidoc, revit_app):
        self._doc = doc
        self._uidoc = uidoc
        self._revit = revit_app
        self._level_checks = []
        self._scale_buttons = []
        self._scale_ratio = int(_DEFAULT_VIEW_SCALE)
        self._busy = False
        self._usuario_map = {}

        self._create_handler = _CreateUserViewsHandler()
        self._create_event = ExternalEvent.Create(self._create_handler)

        self._win = XamlReader.Parse(_build_xaml())
        self._cmb_usuario = self._win.FindName("CmbUsuario")
        self._panel_escala = self._win.FindName("PanelEscala")
        self._txt_codigo = self._win.FindName("TxtCodigo")
        self._panel_niveles = self._win.FindName("PanelNiveles")
        self._txt_nivel_count = self._win.FindName("TxtNivelCount")
        self._txt_subtitle = self._win.FindName("TxtSubtitle")
        self._txt_status = self._win.FindName("TxtStatus")
        self._btn_iniciar = self._win.FindName("BtnIniciar")
        self._btn_cancelar = self._win.FindName("BtnCancelar")
        btn_all = self._win.FindName("BtnSelAll")
        btn_none = self._win.FindName("BtnSelNone")

        if self._txt_subtitle is not None:
            try:
                self._txt_subtitle.Text = u"02_TRABAJO · Cielo + Piso por nivel"
            except Exception:
                pass

        self._fill_usuarios()
        self._fill_escalas()
        self._fill_niveles()
        self._refresh_form_state()

        self._cmb_usuario.SelectionChanged += SelectionChangedEventHandler(
            self._on_usuario_changed
        )
        self._btn_iniciar.Click += RoutedEventHandler(self._on_iniciar)
        self._btn_cancelar.Click += RoutedEventHandler(lambda s, e: self._win.Close())
        btn_manual = self._win.FindName("BtnManual")
        if btn_manual is not None:
            btn_manual.Click += RoutedEventHandler(
                lambda s, e: _open_manual(self._revit)
            )
        btn_gest = self._win.FindName("BtnGestionarPersonas")
        if btn_gest is not None:
            btn_gest.Click += RoutedEventHandler(self._on_gestionar_personas)
        if btn_all is not None:
            btn_all.Click += RoutedEventHandler(
                lambda s, e: self._set_all_levels(True)
            )
        if btn_none is not None:
            btn_none.Click += RoutedEventHandler(
                lambda s, e: self._set_all_levels(False)
            )

        self._win.Closed += EventHandler(lambda s, e: singleton.clear())

    def _combo_item_display(self, item):
        if item is None:
            return u""
        try:
            content = item.Content
        except Exception:
            content = None
        if isinstance(content, TextBlock):
            try:
                return _as_unicode(content.Text).strip()
            except Exception:
                return u""
        try:
            return _as_unicode(content).strip()
        except Exception:
            return u""

    def _selected_usuario_display(self):
        return self._combo_item_display(self._cmb_usuario.SelectedItem)

    def _on_usuario_changed(self, _sender, _e):
        self._refresh_form_state()

    def _fill_usuarios(self, prefer_display=None):
        """Combo = nombre; Tag/código = abreviación con puntos (clasificación)."""
        prefer = unicode(prefer_display or u"").strip()
        self._cmb_usuario.Items.Clear()
        items, mapping = load_modeladores()
        self._usuario_map = mapping or {}
        sel_idx = 0
        for i, display in enumerate(items):
            code = display_to_code(display, self._usuario_map)
            it = ComboBoxItem()
            tb = TextBlock()
            tb.Text = display
            try:
                tb.TextTrimming = TextTrimming.CharacterEllipsis
            except Exception:
                pass
            it.Content = tb
            it.Tag = code
            it.ToolTip = display
            self._cmb_usuario.Items.Add(it)
            if prefer and unicode(display).strip() == prefer:
                sel_idx = i
        if self._cmb_usuario.Items.Count > 0:
            self._cmb_usuario.SelectedIndex = sel_idx
        else:
            self._set_status(
                u"No hay modeladores en personas.json ni lista de respaldo."
            )

    def _on_gestionar_personas(self, _sender, _e):
        """Alta/edición en personas.json (mismo diálogo que Siguiente Revisión)."""
        prev = self._selected_usuario_display()
        try:
            from gestionar_personas_wpf import GestionarPersonasDialog, load_personas_list
            from System.Collections.ObjectModel import ObservableCollection
            from System.IO import Directory
        except Exception:
            mostrar_aviso(
                self._revit,
                u"No se pudo cargar el módulo de gestión de personas.",
            )
            return

        issues_dir, personas_file = personas_paths()
        oc = ObservableCollection[object]()
        for p in load_personas_list(personas_file):
            oc.Add(p)
        try:
            Directory.CreateDirectory(issues_dir)
        except Exception:
            pass

        prev_top = None
        try:
            prev_top = self._win.Topmost
            self._win.Topmost = False
        except Exception:
            pass
        try:
            GestionarPersonasDialog(
                oc,
                issues_dir,
                personas_file,
                uidoc=self._uidoc,
                revit_app=self._revit,
                owner=self._win,
            )
        except Exception as ex:
            mostrar_aviso(
                self._revit,
                u"No se pudo abrir el directorio de personas.",
                content=u"{}".format(ex),
            )
        finally:
            if prev_top is not None:
                try:
                    self._win.Topmost = prev_top
                except Exception:
                    pass

        invalidate_modeladores_cache()
        self._fill_usuarios(prefer_display=prev)
        self._refresh_form_state()

    def _fill_escalas(self):
        self._panel_escala.Children.Clear()
        self._scale_buttons = []
        self._scale_ratio = int(_DEFAULT_VIEW_SCALE)
        for ratio in VIEW_SCALE_RATIOS:
            btn = Button()
            btn.Content = u"1:{}".format(ratio)
            btn.Tag = ratio
            btn.Margin = Thickness(0, 0, 4, 4)
            btn.Padding = Thickness(8, 6, 8, 6)
            btn.MinWidth = 52
            btn.FontSize = 12
            btn.Cursor = Cursors.Hand
            btn.Click += RoutedEventHandler(self._on_scale_click)
            self._panel_escala.Children.Add(btn)
            self._scale_buttons.append(btn)
        self._apply_scale_button_styles()

    def _on_scale_click(self, sender, _e):
        try:
            self._scale_ratio = int(sender.Tag)
        except Exception:
            self._scale_ratio = int(_DEFAULT_VIEW_SCALE)
        self._apply_scale_button_styles()
        self._refresh_form_state()

    def _apply_scale_button_styles(self):
        for btn in self._scale_buttons:
            on = False
            try:
                on = int(btn.Tag) == int(self._scale_ratio)
            except Exception:
                on = False
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

    def _fill_niveles(self):
        self._panel_niveles.Children.Clear()
        self._level_checks = []
        levels = _collect_levels_sorted(self._doc)
        for lv in levels:
            try:
                name = _as_unicode(lv.Name or u"")
            except Exception:
                name = u"?"
            try:
                elev_txt = _format_level_elev(self._doc, lv.Elevation)
            except Exception:
                elev_txt = u""

            cb = CheckBox()
            try:
                cb.Content = u""
            except Exception:
                cb.Content = None
            cb.IsChecked = True
            cb.Margin = Thickness(0, 0, 0, 0)
            cb.Padding = Thickness(0, 0, 0, 0)
            cb.Cursor = Cursors.Hand
            cb.VerticalAlignment = VerticalAlignment.Center
            cb.VerticalContentAlignment = VerticalAlignment.Center
            cb.HorizontalAlignment = HorizontalAlignment.Left
            cb.Tag = lv

            label = TextBlock()
            label.Text = name
            label.FontSize = 12
            label.Margin = Thickness(8, 0, 8, 0)
            label.Padding = Thickness(0, 0, 0, 0)
            label.VerticalAlignment = VerticalAlignment.Center
            label.Cursor = Cursors.Hand
            try:
                label.Foreground = _BRUSH_FG_HI
                label.TextTrimming = TextTrimming.CharacterEllipsis
            except Exception:
                pass

            meta = TextBlock()
            meta.Text = elev_txt
            meta.FontSize = 10
            meta.VerticalAlignment = VerticalAlignment.Center
            meta.Cursor = Cursors.Hand
            try:
                from System.Windows.Media import FontFamily as WpfFontFamily

                meta.FontFamily = WpfFontFamily(u"Consolas")
            except Exception:
                pass
            try:
                meta.Foreground = _BRUSH_FG_MID
            except Exception:
                pass

            row_grid = Grid()
            row_grid.VerticalAlignment = VerticalAlignment.Center
            col_cb = ColumnDefinition()
            col_cb.Width = GridLength(0, GridUnitType.Auto)
            col_name = ColumnDefinition()
            col_name.Width = GridLength(1, GridUnitType.Star)
            col_meta = ColumnDefinition()
            col_meta.Width = GridLength(0, GridUnitType.Auto)
            row_grid.ColumnDefinitions.Add(col_cb)
            row_grid.ColumnDefinitions.Add(col_name)
            row_grid.ColumnDefinitions.Add(col_meta)
            Grid.SetColumn(cb, 0)
            Grid.SetColumn(label, 1)
            Grid.SetColumn(meta, 2)
            row_grid.Children.Add(cb)
            row_grid.Children.Add(label)
            row_grid.Children.Add(meta)

            border = Border()
            border.Background = Brushes.Transparent
            border.CornerRadius = CornerRadius(2)
            border.Padding = Thickness(6, 3, 6, 3)
            border.Margin = Thickness(0, 0, 0, 0)
            border.Cursor = Cursors.Hand
            border.Child = row_grid
            border.Tag = lv
            try:
                border.ToolTip = name if not elev_txt else u"{0}  {1}".format(name, elev_txt)
            except Exception:
                pass

            item = _LevelCheck(lv, cb, border)

            def _make_toggle(checkbox):
                def _on_row_down(sender, args):
                    try:
                        src = args.OriginalSource
                    except Exception:
                        src = None
                    if isinstance(src, CheckBox):
                        return
                    try:
                        checkbox.IsChecked = not bool(checkbox.IsChecked)
                        if args is not None:
                            args.Handled = True
                    except Exception:
                        pass

                return _on_row_down

            try:
                border.MouseLeftButtonDown += MouseButtonEventHandler(
                    _make_toggle(cb)
                )
            except Exception:
                pass

            def _make_hover(row):
                def _enter(sender, args):
                    row.hover = True
                    _style_row(row)

                def _leave(sender, args):
                    row.hover = False
                    _style_row(row)

                return _enter, _leave

            enter, leave = _make_hover(item)
            try:
                from System.Windows.Input import MouseEventHandler

                border.MouseEnter += MouseEventHandler(enter)
                border.MouseLeave += MouseEventHandler(leave)
            except Exception:
                pass

            cb.Checked += RoutedEventHandler(self._on_level_changed)
            cb.Unchecked += RoutedEventHandler(self._on_level_changed)
            _style_row(item)
            self._panel_niveles.Children.Add(border)
            self._level_checks.append(item)

    def _on_level_changed(self, _sender, _e):
        for row in self._level_checks:
            _style_row(row)
        self._refresh_form_state()

    def _set_all_levels(self, checked):
        for item in self._level_checks:
            item.checkbox.IsChecked = checked
        self._refresh_form_state()

    def _get_selected_usuario(self):
        sel = self._cmb_usuario.SelectedItem
        if sel is None:
            return None, None
        code = getattr(sel, "Tag", None) or u""
        display = self._combo_item_display(sel) or _as_unicode(code)
        return code, display

    def _get_selected_scale(self):
        try:
            return int(self._scale_ratio)
        except Exception:
            return int(_DEFAULT_VIEW_SCALE)

    def _get_selected_levels(self):
        out = []
        for item in self._level_checks:
            try:
                if item.checkbox.IsChecked:
                    out.append(item.level)
            except Exception:
                continue
        return out

    def _set_status(self, text):
        if self._txt_status is None:
            return
        try:
            self._txt_status.Text = _as_unicode(text or u"")
        except Exception:
            pass

    def _refresh_form_state(self):
        if self._busy:
            return

        code, display = self._get_selected_usuario()
        if self._txt_codigo is not None:
            self._txt_codigo.Text = unicode(code or u"—")
        try:
            name = _as_unicode(display).strip()
            c = _as_unicode(code).strip()
            if name and c:
                self._cmb_usuario.ToolTip = (
                    u"{0} ({1}). El código va a Subclasificacion y a los "
                    u"nombres de vista."
                ).format(name, c)
            else:
                self._cmb_usuario.ToolTip = (
                    u"Modelador: el código (p. ej. A.A.U) va a "
                    u"Subclasificacion y a los nombres de vista"
                )
        except Exception:
            pass
        try:
            chip = self._win.FindName("ChipClasificacion")
            if chip is not None:
                c = _as_unicode(code).strip() or u"—"
                chip.ToolTip = (
                    u"Abreviación {0}: una sola generación por este código "
                    u"en el documento."
                ).format(c)
        except Exception:
            pass

        levels = self._get_selected_levels()
        total = len(self._level_checks)
        n = len(levels)
        plantas = n * 2
        if self._txt_nivel_count is not None:
            self._txt_nivel_count.Text = _texto_contador(n, total, plantas)

        scale = self._get_selected_scale()
        can_run = bool(code) and n > 0
        _style_crear_button(self._btn_iniciar, can_run)

        if not code:
            self._set_status(u"Seleccione un modelador.")
        elif n == 0:
            self._set_status(u"Seleccione al menos un nivel.")
        else:
            self._set_status(u"{0} · 1:{1}".format(code, scale))

    def _on_iniciar(self, sender, args):
        if self._busy:
            return

        code, display = self._get_selected_usuario()
        if not code:
            self._set_status(u"Seleccione un modelador.")
            mostrar_aviso(self._revit, u"Seleccione un modelador.")
            return

        levels = self._get_selected_levels()
        if not levels:
            self._set_status(u"Seleccione al menos un nivel.")
            mostrar_aviso(self._revit, u"Seleccione al menos un nivel.")
            return

        from vistas_por_usuario.service import VistasPorUsuarioRequest

        scale = self._get_selected_scale()
        req = VistasPorUsuarioRequest(code, scale, list(levels), display)

        self._busy = True
        _style_crear_button(self._btn_iniciar, False)

        self._create_handler.request = req
        self._create_handler.uiapp_for_dialog = self._revit
        _pin_external_event(self._create_event, self._create_handler)
        try:
            self._create_event.Raise()
        except Exception as ex:
            self._busy = False
            _unpin_external_event()
            self._refresh_form_state()
            mostrar_aviso(
                self._revit,
                u"No se pudo iniciar la creación.",
                content=u"{}".format(ex),
            )
            return

        # Cerrar UI de inmediato; Execute usa el snapshot (sin controles WPF).
        try:
            self._win.Close()
        except Exception:
            pass

    def show(self):
        _prepare_window(self._win, self._revit)
        singleton.register(self._win)
        self._win.Show()


def show_vistas_por_usuario_ui(revit_app):
    if singleton.try_activate_existing():
        mostrar_aviso(revit_app, u"La herramienta ya está en ejecución.")
        return
    try:
        uidoc = revit_app.ActiveUIDocument
        doc = uidoc.Document
    except Exception:
        mostrar_aviso(revit_app, u"No hay documento activo.")
        return
    w = VistasPorUsuarioWindow(doc, uidoc, revit_app)
    w.show()
