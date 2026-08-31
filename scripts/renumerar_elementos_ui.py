# -*- coding: utf-8 -*-
"""
UI WPF — Arainco: Renumerar elementos.

Configura categoría, número inicial y tratamiento de duplicados;
después numera en el orden de selección (ESC termina).

Revit 2024+ | IronPython (pyRevit).
Entrada: ``31_RenumerarElementos.pushbutton/script.py``.
"""

from __future__ import print_function

try:
    unicode
except NameError:
    unicode = str

import os
import traceback

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from System import AppDomain, EventHandler
from System.Windows import RoutedEventHandler, Visibility, WindowState
from System.Windows.Controls import ComboBoxItem, SelectionChangedEventHandler
from System.Windows.Input import Key, KeyEventHandler
from System.Windows.Markup import XamlReader
from Autodesk.Revit.UI import TaskDialog

from bimtools_ui_tokens import BTN_MANUAL
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_wpf_shell import build_simple_tool_xaml
from revit_wpf_window_position import (
    bind_center_wpf_on_revit_monitor,
    position_wpf_window_center_on_monitor,
    revit_main_hwnd,
)
from renumerar_elementos import (
    DUPE_ALERT,
    DUPE_LABELS,
    DUPE_SKIP,
    DUPE_SWEEP,
    KEY_DOORS_BY_ROOM,
    RenumberStats,
    allowed_view_names,
    door_by_room_renumber,
    format_summary,
    is_allowed_view,
    list_options_for_view,
    pick_and_renumber,
)

_TOOL_TITLE = u"Arainco: Renumerar elementos"
_SINGLETON_KEY = u"Arainco_RenumerarElementos_UI"
_BUSY_KEY = u"Arainco_RenumerarElementos_Busy"
_SETTINGS_KEY = u"Arainco_RenumerarElementos_Settings"
_ALREADY_RUNNING = u"La herramienta ya esta en ejecucion."

_SUBTITLE = (
    u"Numera los elementos en el orden en que los selecciona. "
    u"Escape termina la selección."
)

_BODY_XAML = u"""
<StackPanel>
  <TextBlock Style="{StaticResource Label}" Text="Categoría"/>
  <ComboBox x:Name="CboCategory" Style="{StaticResource ComboStretch}"
            MinHeight="32" Margin="0,0,0,10"/>

  <StackPanel x:Name="PanelStart">
    <TextBlock Style="{StaticResource Label}" Text="Número inicial"/>
    <TextBox x:Name="TxtStart" Style="{StaticResource BimToolsTextBoxDark}"
             MinHeight="32" VerticalContentAlignment="Center"
             Text="1" ToolTip="Ejemplos: 1, 01, A1, 101A"/>
    <TextBlock Style="{StaticResource LabelSmall}" Margin="0,4,0,10"
               TextWrapping="Wrap" FontWeight="Normal"
               Text="Se incrementa en cada clic (1→2, A1→A2, 001→002)."/>
  </StackPanel>

  <TextBlock x:Name="TxtDoorHint" Style="{StaticResource LabelSmall}"
             Margin="0,0,0,10" TextWrapping="Wrap" FontWeight="Normal"
             Visibility="Collapsed"
             Text="La marca de la puerta se toma del número de la habitación (101 o 101A si hay varias)."/>

  <TextBlock Style="{StaticResource Label}" Text="Si el número ya está ocupado"/>
  <ComboBox x:Name="CboDupe" Style="{StaticResource ComboStretch}"
            MinHeight="32" Margin="0,0,0,10"/>

  <CheckBox x:Name="ChkPadding" Foreground="#95B8CC" FontSize="12"
            Margin="0,2,0,8" IsChecked="True"
            Content="Conservar ceros a la izquierda (001 → 002)"/>
  <CheckBox x:Name="ChkSkipSession" Foreground="#95B8CC" FontSize="12"
            Margin="0,0,0,0" IsChecked="True"
            Content="No volver a numerar el mismo elemento en esta sesión"/>
</StackPanel>
"""

_FOOTER_LEADING_XAML = (
    u'<Button x:Name="BtnManual" Content="Manual" '
    u'Style="{{StaticResource BtnSelectOutline}}" '
    u'Background="{bg}" MinWidth="96" MinHeight="32" MaxHeight="36" '
    u'Padding="8,2" VerticalAlignment="Center" '
    u'ToolTip="Abrir manual de usuario"/>'
).format(bg=BTN_MANUAL)

_FOOTER_ACTIONS_XAML = u"""
<Button x:Name="BtnStart" Content="Comenzar"
        Style="{StaticResource BtnPrimary}" MinWidth="130"
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
        _mostrar_aviso(uiapp, u"No se pudo abrir el manual.", content=_as_unicode(ex))


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
        title=_TOOL_TITLE,
        styles_xml=BIMTOOLS_DARK_STYLES_XML,
        body_xaml=_BODY_XAML,
        footer_leading_xaml=_FOOTER_LEADING_XAML,
        footer_actions_xaml=_FOOTER_ACTIONS_XAML,
        width=520,
        min_width=480,
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


def _combo_select_by_tag(combo, tag):
    if combo is None or tag is None:
        return
    wanted = _as_unicode(tag)
    for item in combo.Items:
        try:
            if _as_unicode(item.Tag) == wanted:
                combo.SelectedItem = item
                return
        except Exception:
            continue
    if combo.Items.Count > 0:
        combo.SelectedIndex = 0


def _combo_tag(combo):
    if combo is None:
        return None
    item = combo.SelectedItem
    if item is None:
        return None
    try:
        return item.Tag
    except Exception:
        return None


class RenumerarElementosWindow(object):
    def __init__(self, uiapp, view):
        self._uiapp = uiapp
        self._view = view
        self.accepted = None
        self._options = list_options_for_view(view)
        self._win = XamlReader.Parse(_build_xaml())
        self._txt_subtitle = self._win.FindName(u"TxtSubtitle")
        self._txt_status = self._win.FindName(u"TxtStatus")
        self._cbo_cat = self._win.FindName(u"CboCategory")
        self._cbo_dupe = self._win.FindName(u"CboDupe")
        self._txt_start = self._win.FindName(u"TxtStart")
        self._panel_start = self._win.FindName(u"PanelStart")
        self._txt_door_hint = self._win.FindName(u"TxtDoorHint")
        self._chk_padding = self._win.FindName(u"ChkPadding")
        self._chk_skip = self._win.FindName(u"ChkSkipSession")

        self._wire_events()
        _prepare_window(self._win, uiapp)
        if self._txt_subtitle is not None:
            self._txt_subtitle.Text = _SUBTITLE
        self._fill_combos()
        self._apply_settings(_load_settings())
        self._sync_category_ui()
        self._set_status(u"Configure y pulse Comenzar. Luego seleccione en orden.")

    def _wire_events(self):
        self._win.FindName(u"BtnStart").Click += RoutedEventHandler(self._on_start)
        self._win.FindName(u"BtnClose").Click += RoutedEventHandler(self._on_close)
        manual = self._win.FindName(u"BtnManual")
        if manual is not None:
            manual.Click += RoutedEventHandler(self._on_manual)
        if self._cbo_cat is not None:
            self._cbo_cat.SelectionChanged += SelectionChangedEventHandler(
                self._on_category_changed
            )
        self._win.KeyDown += KeyEventHandler(self._on_key_down)
        self._win.Closed += EventHandler(self._on_closed)

    def _fill_combos(self):
        if self._cbo_cat is not None:
            self._cbo_cat.Items.Clear()
            for opt in self._options:
                item = ComboBoxItem()
                item.Content = opt.label
                item.Tag = opt.key
                self._cbo_cat.Items.Add(item)
            if self._cbo_cat.Items.Count > 0:
                self._cbo_cat.SelectedIndex = 0
        if self._cbo_dupe is not None:
            self._cbo_dupe.Items.Clear()
            for key, label in DUPE_LABELS:
                item = ComboBoxItem()
                item.Content = label
                item.Tag = key
                self._cbo_dupe.Items.Add(item)
            self._cbo_dupe.SelectedIndex = 0

    def _apply_settings(self, data):
        if not data:
            return
        _combo_select_by_tag(self._cbo_cat, data.get(u"category"))
        _combo_select_by_tag(self._cbo_dupe, data.get(u"dupe_mode"))
        start = _as_unicode(data.get(u"start") or u"").strip()
        if start and self._txt_start is not None:
            self._txt_start.Text = start
        if self._chk_padding is not None and u"preserve_padding" in data:
            self._chk_padding.IsChecked = bool(data.get(u"preserve_padding"))
        if self._chk_skip is not None and u"skip_session" in data:
            self._chk_skip.IsChecked = bool(data.get(u"skip_session"))

    def _selected_option(self):
        key = _combo_tag(self._cbo_cat)
        for opt in self._options:
            if opt.key == key:
                return opt
        if self._options:
            return self._options[0]
        return None

    def _sync_category_ui(self):
        opt = self._selected_option()
        by_room = opt is not None and opt.key == KEY_DOORS_BY_ROOM
        if self._panel_start is not None:
            self._panel_start.Visibility = (
                Visibility.Collapsed if by_room else Visibility.Visible
            )
        if self._txt_door_hint is not None:
            self._txt_door_hint.Visibility = (
                Visibility.Visible if by_room else Visibility.Collapsed
            )
        if by_room:
            self._set_status(u"Seleccione pares puerta + habitación tras Comenzar.")
        else:
            self._set_status(u"Configure y pulse Comenzar. Luego seleccione en orden.")

    def _on_category_changed(self, sender, args):
        self._sync_category_ui()

    def _on_manual(self, sender, args):
        _open_manual(self._uiapp)

    def _on_key_down(self, sender, args):
        if args.Key == Key.Escape:
            self._win.Close()

    def _on_close(self, sender, args):
        self._win.Close()

    def _on_closed(self, sender, args):
        try:
            AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, None)
        except Exception:
            pass

    def _set_status(self, text):
        if self._txt_status is not None:
            self._txt_status.Text = _as_unicode(text)

    def _on_start(self, sender, args):
        opt = self._selected_option()
        if opt is None:
            _mostrar_aviso(self._uiapp, u"Seleccione una categoría.")
            return
        start = u"1"
        if self._txt_start is not None:
            start = _as_unicode(self._txt_start.Text).strip()
        if opt.key != KEY_DOORS_BY_ROOM and not start:
            _mostrar_aviso(self._uiapp, u"Indique el número inicial.")
            return
        dupe = _combo_tag(self._cbo_dupe) or DUPE_SWEEP
        if dupe not in (DUPE_SWEEP, DUPE_ALERT, DUPE_SKIP):
            dupe = DUPE_SWEEP
        preserve = True
        if self._chk_padding is not None:
            preserve = bool(self._chk_padding.IsChecked)
        skip_session = True
        if self._chk_skip is not None:
            skip_session = bool(self._chk_skip.IsChecked)
        self.accepted = {
            u"option": opt,
            u"start": start,
            u"dupe_mode": dupe,
            u"preserve_padding": preserve,
            u"skip_session": skip_session,
        }
        _save_settings(
            {
                u"category": opt.key,
                u"start": start,
                u"dupe_mode": dupe,
                u"preserve_padding": preserve,
                u"skip_session": skip_session,
            }
        )
        self._win.Close()

    def show_dialog(self):
        self._win.ShowDialog()


def _existing_controller():
    try:
        ctrl = AppDomain.CurrentDomain.GetData(_SINGLETON_KEY)
    except Exception:
        ctrl = None
    if ctrl is None:
        return None
    try:
        if hasattr(ctrl, "_win") and ctrl._win is not None and ctrl._win.IsVisible:
            return ctrl
    except Exception:
        pass
    return None


def _is_busy():
    try:
        return bool(AppDomain.CurrentDomain.GetData(_BUSY_KEY))
    except Exception:
        return False


def _set_busy(value):
    try:
        AppDomain.CurrentDomain.SetData(_BUSY_KEY, bool(value))
    except Exception:
        pass


def _focus_existing(ctrl, uiapp):
    try:
        if ctrl._win.WindowState == WindowState.Minimized:
            ctrl._win.WindowState = WindowState.Normal
        ctrl._win.Activate()
    except Exception:
        pass
    _mostrar_aviso(uiapp, _ALREADY_RUNNING)


def _show_config(uiapp, view):
    existing = _existing_controller()
    if existing is not None:
        _focus_existing(existing, uiapp)
        return None
    try:
        ctrl = RenumerarElementosWindow(uiapp, view)
    except Exception as ex:
        tb = traceback.format_exc()
        print(tb)
        _mostrar_aviso(
            uiapp,
            u"No se pudo abrir la ventana de la herramienta.",
            content=_as_unicode(ex) + u"\n\n" + _as_unicode(tb),
        )
        return None
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, ctrl)
    except Exception:
        pass
    try:
        ctrl.show_dialog()
    finally:
        try:
            AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, None)
        except Exception:
            pass
    return ctrl.accepted


def run(revit):
    uiapp = revit if hasattr(revit, "ActiveUIDocument") else None
    if uiapp is None:
        try:
            uiapp = revit.uiapp
        except Exception:
            uiapp = None
    if uiapp is None:
        return

    if _is_busy() or _existing_controller() is not None:
        existing = _existing_controller()
        if existing is not None:
            _focus_existing(existing, uiapp)
        else:
            _mostrar_aviso(uiapp, _ALREADY_RUNNING)
        return

    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        _mostrar_aviso(uiapp, u"No hay un documento activo.")
        return
    doc = uidoc.Document
    view = doc.ActiveView
    if view is None or not is_allowed_view(view):
        _mostrar_aviso(
            uiapp,
            u"Abra una vista de {0} e inténtelo de nuevo.".format(
                allowed_view_names()
            ),
        )
        return

    cfg = _show_config(uiapp, view)
    if not cfg:
        return

    option = cfg.get(u"option")
    if option is None:
        return

    stats = RenumberStats()
    _set_busy(True)
    try:
        if option.is_doors_by_room:
            door_by_room_renumber(
                uidoc,
                doc,
                option,
                cfg.get(u"dupe_mode") or DUPE_SWEEP,
                bool(cfg.get(u"preserve_padding", True)),
                bool(cfg.get(u"skip_session", True)),
                stats,
            )
        else:
            pick_and_renumber(
                uidoc,
                doc,
                option,
                cfg.get(u"start") or u"1",
                cfg.get(u"dupe_mode") or DUPE_SWEEP,
                bool(cfg.get(u"preserve_padding", True)),
                bool(cfg.get(u"skip_session", True)),
                stats,
            )
    except Exception as ex:
        tb = traceback.format_exc()
        print(tb)
        _mostrar_aviso(
            uiapp,
            u"Error durante la renumeración.",
            content=_as_unicode(ex) + u"\n\n" + _as_unicode(tb),
        )
        return
    finally:
        _set_busy(False)

    instruction, detail = format_summary(stats, option.label)
    _mostrar_aviso(uiapp, instruction, content=detail, ok_text=u"Cerrar")
