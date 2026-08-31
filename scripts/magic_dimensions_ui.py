# -*- coding: utf-8 -*-
"""
UI WPF — Arainco: Magic Dimensions.

Formulario de escenario + flujo de selección en la vista activa.
Revit 2024+ | IronPython (pyRevit).
Entrada: ``73_MagicDimensions.pushbutton/script.py``.
"""

from __future__ import print_function

import os
import weakref

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from System import AppDomain, EventHandler
from System.Windows import RoutedEventHandler, WindowState
from System.Windows.Controls import ComboBoxItem, SelectionChangedEventHandler
from System.Windows.Input import Key, KeyEventHandler
from System.Windows.Markup import XamlReader
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler, TaskDialog

from bimtools_ui_tokens import BTN_MANUAL, FG_BODY, FONT_SIZE_BODY
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_wpf_shell import build_simple_tool_xaml
from revit_wpf_window_position import (
    position_wpf_window_top_left_at_active_view,
    revit_main_hwnd,
)
from magic_dimensions import (
    SCENARIO_FLOORS_SECTION_ELEVATION,
    SCENARIO_GRIDS_BUILDING_SECTION,
    SCENARIOS,
    _TOOL_TITLE,
    is_building_section_view,
    preselected_grids,
    resolve_active_view,
    run_scenario,
    view_display_name,
)
from cota_spot_elevacion_losas import (
    is_section_or_elevation_view,
    preselected_floors,
)

_SINGLETON_KEY = u"Arainco_MagicDimensions_UI"
_BUSY_KEY = u"Arainco_MagicDimensions_Busy"
_SETTINGS_KEY = u"Arainco_MagicDimensions_Settings"

_SUBTITLE = (
    u"Cotas automáticas según el escenario de la vista: "
    u"ejes en Building Section, o cota de elevación en losas, "
    u"fundaciones, muros y Structural Framing."
)

_BODY_XAML = u"""
<StackPanel>
  <TextBlock Style="{{StaticResource Label}}" Text="Escenario"/>
  <ComboBox x:Name="CboScenario" Style="{{StaticResource ComboStretch}}"
            MinHeight="32" Margin="0,0,0,10"/>
  <TextBlock x:Name="TxtHelp" TextWrapping="Wrap"
             Foreground="{fg}" FontSize="{fs}" LineHeight="17"/>
  <Border Margin="0,12,0,0" Background="#071018" BorderBrush="#21465C"
          BorderThickness="1" CornerRadius="4" Padding="12,10">
    <TextBlock x:Name="TxtView" TextWrapping="Wrap"
               Foreground="#E8F4F8" FontSize="12" LineHeight="18"/>
  </Border>
</StackPanel>
""".format(fg=FG_BODY, fs=FONT_SIZE_BODY)

_FOOTER_LEADING_XAML = (
    u'<Button x:Name="BtnManual" Content="Manual" '
    u'Style="{{StaticResource BtnSelectOutline}}" '
    u'Background="{bg}" MinWidth="96" MinHeight="32" MaxHeight="36" '
    u'Padding="8,2" VerticalAlignment="Center" '
    u'ToolTip="Abrir manual de usuario"/>'
).format(bg=BTN_MANUAL)

_FOOTER_ACTIONS_XAML = u"""
<Button x:Name="BtnRun" Content="Acotar"
        Style="{StaticResource BtnPrimary}" MinWidth="120"
        Margin="0,0,10,0" IsDefault="True"/>
<Button x:Name="BtnClose" Content="Cerrar"
        Style="{StaticResource BtnSelectOutline}" MinWidth="110"/>
"""


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


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
            _TOOL_TITLE,
            instruction=_as_unicode(instruction),
            content=_as_unicode(content) if content else None,
            ok_text=ok_text,
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass
    try:
        body = _as_unicode(instruction)
        if content:
            body = body + u"\n\n" + _as_unicode(content)
        TaskDialog.Show(_TOOL_TITLE, body)
    except Exception:
        pass


def _resolve_manual_path():
    try:
        import bimtools_paths

        pb = bimtools_paths.get_pushbutton_dir()
        if pb:
            path = os.path.join(pb, u"manual_usuario.html")
            if os.path.isfile(path):
                return path
    except Exception:
        pass
    return None


def _open_manual(uiapp):
    path = _resolve_manual_path()
    if not path:
        _mostrar_aviso(
            uiapp,
            u"No se encontró manual_usuario.html en la carpeta del botón.",
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


def _place_on_active_view(win, uiapp):
    if win is None:
        return
    uidoc = None
    try:
        uidoc = uiapp.ActiveUIDocument if uiapp is not None else None
    except Exception:
        uidoc = None
    hwnd = None
    try:
        hwnd = revit_main_hwnd(uiapp)
    except Exception:
        hwnd = None
    try:
        position_wpf_window_top_left_at_active_view(win, uidoc, hwnd)
    except Exception:
        pass


def _prepare_window(win, uiapp):
    if win is None:
        return
    _attach_revit_owner(win, uiapp)
    _place_on_active_view(win, uiapp)

    def _apply(sender, args):
        _place_on_active_view(win, uiapp)

    try:
        from System.Windows import RoutedEventHandler

        h = RoutedEventHandler(_apply)
        win.Loaded += h
        try:
            win.ContentRendered += h
        except Exception:
            pass
    except Exception:
        pass


def _build_xaml():
    return build_simple_tool_xaml(
        title=_TOOL_TITLE,
        styles_xml=BIMTOOLS_DARK_STYLES_XML,
        body_xaml=_BODY_XAML,
        footer_leading_xaml=_FOOTER_LEADING_XAML,
        footer_actions_xaml=_FOOTER_ACTIONS_XAML,
        width=520,
        min_width=520,
        resize_mode=u"NoResize",
        size_to_content_height=True,
    )


def _load_settings():
    try:
        data = AppDomain.CurrentDomain.GetData(_SETTINGS_KEY)
        if isinstance(data, dict):
            return data
    except Exception:
        pass
    return {}


def _save_settings(data):
    try:
        AppDomain.CurrentDomain.SetData(_SETTINGS_KEY, data)
    except Exception:
        pass


def _set_busy(flag):
    try:
        AppDomain.CurrentDomain.SetData(_BUSY_KEY, bool(flag))
    except Exception:
        pass


def _is_busy():
    try:
        return bool(AppDomain.CurrentDomain.GetData(_BUSY_KEY))
    except Exception:
        return False


class _AcotarHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref

    def GetName(self):
        return _TOOL_TITLE

    def Execute(self, uiapp):
        win = self._window_ref() if self._window_ref else None
        if win is None:
            _set_busy(False)
            return
        win.execute_acotar(uiapp)


class MagicDimensionsWindow(object):
    def __init__(self, uiapp):
        self._uiapp = uiapp
        self._handler = _AcotarHandler(weakref.ref(self))
        self._ext_event = ExternalEvent.Create(self._handler)
        self._win = XamlReader.Parse(_build_xaml())
        self._txt_subtitle = self._win.FindName(u"TxtSubtitle")
        self._txt_status = self._win.FindName(u"TxtStatus")
        self._txt_help = self._win.FindName(u"TxtHelp")
        self._txt_view = self._win.FindName(u"TxtView")
        self._cbo = self._win.FindName(u"CboScenario")
        self._wire_events()
        _prepare_window(self._win, uiapp)
        if self._txt_subtitle is not None:
            self._txt_subtitle.Text = _SUBTITLE
        self._fill_scenarios()
        self._refresh_context()

    def _wire_events(self):
        self._win.FindName(u"BtnRun").Click += RoutedEventHandler(self._on_run)
        self._win.FindName(u"BtnClose").Click += RoutedEventHandler(self._on_close)
        manual = self._win.FindName(u"BtnManual")
        if manual is not None:
            manual.Click += RoutedEventHandler(self._on_manual)
        self._win.KeyDown += KeyEventHandler(self._on_key_down)
        self._win.Closed += EventHandler(self._on_closed)

    def _fill_scenarios(self):
        if self._cbo is None:
            return
        self._cbo.Items.Clear()
        saved = _as_unicode(_load_settings().get(u"scenario") or u"")
        selected = None
        for key, label, help_text in SCENARIOS:
            item = ComboBoxItem()
            item.Content = label
            item.Tag = key
            item.ToolTip = help_text
            self._cbo.Items.Add(item)
            if key == saved:
                selected = item
        if selected is not None:
            self._cbo.SelectedItem = selected
        elif self._cbo.Items.Count > 0:
            self._cbo.SelectedIndex = 0
        self._cbo.SelectionChanged += SelectionChangedEventHandler(
            self._on_scenario_changed
        )
        self._sync_help()

    def _on_scenario_changed(self, sender, args):
        self._sync_help()
        key = self._selected_scenario()
        if key:
            _save_settings({u"scenario": key})
        self._refresh_context()

    def _selected_scenario(self):
        if self._cbo is None or self._cbo.SelectedItem is None:
            return SCENARIO_GRIDS_BUILDING_SECTION
        try:
            return _as_unicode(self._cbo.SelectedItem.Tag)
        except Exception:
            return SCENARIO_GRIDS_BUILDING_SECTION

    def _scenario_help(self):
        key = self._selected_scenario()
        for skey, _label, help_text in SCENARIOS:
            if skey == key:
                return help_text
        return u""

    def _sync_help(self):
        if self._txt_help is not None:
            self._txt_help.Text = self._scenario_help()

    def _uidoc(self):
        try:
            return self._uiapp.ActiveUIDocument
        except Exception:
            return None

    def _refresh_context(self):
        uidoc = self._uidoc()
        view = resolve_active_view(uidoc)
        vname = view_display_name(view)
        key = self._selected_scenario()
        if view is None:
            text = u"No hay vista gráfica activa."
            status = u"Abra una vista válida para el escenario."
        elif key == SCENARIO_FLOORS_SECTION_ELEVATION:
            if is_section_or_elevation_view(view):
                n_sel = len(preselected_floors(uidoc))
                if n_sel >= 2:
                    text = (
                        u"Vista: {0}\n"
                        u"Tipo: Sección / Alzado\n"
                        u"{1} elemento(s) ya seleccionados: al pulsar Acotar "
                        u"solo se pedirá el punto de la cota."
                    ).format(vname, n_sel)
                else:
                    text = (
                        u"Vista: {0}\n"
                        u"Tipo: Sección / Alzado\n"
                        u"Al pulsar Acotar, seleccione losas, fundaciones, "
                        u"muros o Structural Framing y luego un punto."
                    ).format(vname)
                status = u"Lista para acotar losas, fundaciones, muros o vigas."
            else:
                text = (
                    u"Vista: {0}\n"
                    u"Este escenario solo opera en Sección o Alzado."
                ).format(vname)
                status = u"Cambie a Sección o Alzado."
        elif is_building_section_view(view):
            n_sel = len(preselected_grids(uidoc))
            if n_sel >= 2:
                text = (
                    u"Vista: {0}\n"
                    u"Tipo: Building Section\n"
                    u"{1} eje(s) ya seleccionados: al pulsar Acotar "
                    u"solo se pedirá el punto de la cota."
                ).format(vname, n_sel)
            else:
                text = (
                    u"Vista: {0}\n"
                    u"Tipo: Building Section\n"
                    u"Al pulsar Acotar, seleccione ejes y luego un punto."
                ).format(vname)
            status = u"Lista para acotar."
        else:
            text = (
                u"Vista: {0}\n"
                u"El escenario de ejes solo opera en Building Section."
            ).format(vname)
            status = u"Cambie a una Building Section."
        if self._txt_view is not None:
            self._txt_view.Text = text
        self._set_status(status)

    def _set_status(self, text):
        if self._txt_status is not None:
            self._txt_status.Text = _as_unicode(text)

    def _on_manual(self, sender, args):
        _open_manual(self._uiapp)

    def _on_key_down(self, sender, args):
        if args.Key == Key.Escape:
            self._win.Close()

    def _on_close(self, sender, args):
        self._win.Close()

    def _on_closed(self, sender, args):
        _set_busy(False)
        try:
            AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, None)
        except Exception:
            pass

    def _aviso(self, instruction, content=u""):
        _mostrar_aviso(self._uiapp, instruction, content=content)

    def _restore_window(self):
        _set_busy(False)
        try:
            if self._win is not None and self._win.IsLoaded:
                self._refresh_context()
                self._win.Show()
                _place_on_active_view(self._win, self._uiapp)
                self._win.Activate()
        except Exception:
            pass

    def execute_acotar(self, uiapp):
        """Corre en contexto API (ExternalEvent): picks + transacción."""
        try:
            uidoc = None
            try:
                uidoc = uiapp.ActiveUIDocument if uiapp is not None else None
            except Exception:
                uidoc = None
            if uidoc is None:
                self._aviso(u"No hay documento activo.")
                return
            key = self._selected_scenario()
            _ok, status = run_scenario(
                key, uidoc, self._aviso, use_preselection=True
            )
            self._set_status(status)
        except Exception as ex:
            self._aviso(
                u"No se pudo crear la cota.",
                content=_as_unicode(ex),
            )
            self._set_status(_as_unicode(ex))
        finally:
            self._restore_window()

    def _on_run(self, sender, args):
        uidoc = self._uidoc()
        if uidoc is None:
            self._aviso(u"No hay documento activo.")
            return
        if self._ext_event is None:
            self._aviso(u"No se pudo iniciar el comando en Revit.")
            return

        try:
            self._win.Hide()
        except Exception:
            pass
        _set_busy(True)
        self._set_status(u"Seleccione elementos y el punto de la cota…")
        try:
            self._ext_event.Raise()
        except Exception as ex:
            self._aviso(
                u"No se pudo iniciar el comando en Revit.",
                content=_as_unicode(ex),
            )
            self._restore_window()

    def show(self):
        self._win.Show()


def _existing_window():
    try:
        w = AppDomain.CurrentDomain.GetData(_SINGLETON_KEY)
    except Exception:
        w = None
    if w is None:
        return None
    try:
        if hasattr(w, "_win") and w._win is not None and w._win.IsLoaded:
            return w
    except Exception:
        pass
    return None


def _focus_existing(win, uiapp):
    if not _is_busy():
        try:
            if win._win.WindowState == WindowState.Minimized:
                win._win.WindowState = WindowState.Normal
            if not win._win.IsVisible:
                win._win.Show()
            win._win.Activate()
        except Exception:
            pass
    _mostrar_aviso(uiapp, u"La herramienta ya esta en ejecucion.")


def show_magic_dimensions_window(revit):
    uiapp = revit
    existing = _existing_window()
    if existing is not None or _is_busy():
        if existing is not None:
            _focus_existing(existing, uiapp)
        else:
            _mostrar_aviso(uiapp, u"La herramienta ya esta en ejecucion.")
        return
    win = MagicDimensionsWindow(uiapp)
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, win)
    except Exception:
        pass
    win.show()


def run(revit):
    show_magic_dimensions_window(revit)
