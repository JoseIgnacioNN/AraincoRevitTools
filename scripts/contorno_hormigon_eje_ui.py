# -*- coding: utf-8 -*-
"""UI WPF — contorno de hormigón por eje (Grid). Respaldo en scripts/ de la extensión."""

from __future__ import print_function

import weakref

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from System import AppDomain, EventHandler
from System.Windows import (
    RoutedEventHandler,
    TextWrapping,
    WindowState,
)
from System.Windows.Controls import ComboBoxItem
from System.Windows.Input import Key, KeyEventHandler
from System.Windows.Markup import XamlReader
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler, TaskDialog

from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from contorno_hormigon_eje import (
    _DIALOG_TITLE,
    _as_unicode,
    ejecutar_contorno,
    listar_ejes_modelo,
    recoger_hormigon_en_vista,
    vista_permitida,
)
from revit_wpf_window_position import (
    bind_center_wpf_on_revit_monitor,
    position_wpf_window_center_on_monitor,
    revit_main_hwnd,
)

_WINDOW_TITLE = _DIALOG_TITLE
_SINGLETON_KEY = u"Arainco_ContornoHormigonEje_UI"
_TX_GENERAR = u"Arainco: Contorno hormigón por eje"

_HELP_TEXT = (
    u"Genera el contorno del hormigón visible en la vista activa cortando "
    u"la unión booleana de todos los elementos con Material for Model Behavior "
    u"= Concrete con el plano del eje seleccionado.\n\n"
    u"Las detail lines resultantes se agrupan como CONTORNO + nombre del eje."
)

XAML = u"""
<Window
    x:Name="ContornoWin"
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="__WINDOW_TITLE__"
    Width="520"
    WindowStartupLocation="Manual"
    Background="Transparent"
    AllowsTransparency="True"
    FontFamily="Segoe UI"
    WindowStyle="None"
    ResizeMode="NoResize"
    SizeToContent="Height"
    ShowInTaskbar="False">
  <Window.Resources>
""" + BIMTOOLS_DARK_STYLES_XML + u"""
  </Window.Resources>
  <Border CornerRadius="8" Background="#071018" BorderBrush="#21465C"
          BorderThickness="1" Padding="22,20">
    <StackPanel>
      <TextBlock Text="__WINDOW_TITLE__" Foreground="#E8F4F8"
                 FontSize="16" FontWeight="Bold"/>
      <TextBlock x:Name="TxtSubtitle" Margin="0,6,0,0"
                 Foreground="#95B8CC" FontSize="11" TextWrapping="Wrap"/>
      <Border Margin="0,14,0,0" Background="#0a1620" BorderBrush="#21465C"
              BorderThickness="1" CornerRadius="6" Padding="12,10">
        <TextBlock Text="__HELP__" TextWrapping="Wrap"
                   Foreground="#95B8CC" FontSize="11" LineHeight="16"/>
      </Border>
      <TextBlock Margin="0,14,0,6" Text="Eje (Grid) para el plano de corte"
                 Foreground="#95B8CC" FontSize="11" FontWeight="SemiBold"/>
      <ComboBox x:Name="CmbEje" Style="{StaticResource ComboStretch}"
                MinHeight="32" MaxDropDownHeight="280"/>
      <Button x:Name="BtnGenerar" Margin="0,16,0,0"
              Content="Generar contorno"
              Style="{StaticResource BtnPrimary}"
              HorizontalAlignment="Stretch" MinHeight="36"/>
      <TextBlock x:Name="TxtStatus" Margin="0,12,0,0"
                 Foreground="#64748b" FontSize="10" TextWrapping="Wrap"/>
      <StackPanel Margin="0,14,0,0" Orientation="Horizontal"
                  HorizontalAlignment="Right">
        <Button x:Name="BtnClose" Content="Cerrar"
                Style="{StaticResource BtnSelectOutline}" MinWidth="108"/>
      </StackPanel>
    </StackPanel>
  </Border>
</Window>
"""


def _escape_xaml(text):
    s = _as_unicode(text)
    return (
        s.replace(u"&", u"&amp;")
        .replace(u"<", u"&lt;")
        .replace(u">", u"&gt;")
        .replace(u'"', u"&quot;")
    )


def _build_xaml():
    xaml = XAML
    xaml = xaml.replace(u"__WINDOW_TITLE__", _escape_xaml(_WINDOW_TITLE))
    xaml = xaml.replace(u"__HELP__", _escape_xaml(_HELP_TEXT))
    return xaml


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


class _GenerarContornoHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        win = self._window_ref()
        if win is None:
            return
        grid = win._grid_seleccionado()
        if grid is None:
            win._set_status(u"Selecciona un eje (Grid) en la lista.")
            return
        uidoc = uiapp.ActiveUIDocument
        ok, msg = ejecutar_contorno(uidoc, grid)
        if ok:
            win._set_status(msg)
            win._refresh_summary()
        else:
            win._set_status(u"Error: {0}".format(msg))
            try:
                TaskDialog.Show(_DIALOG_TITLE, msg)
            except Exception:
                pass

    def GetName(self):
        return _TX_GENERAR


class ContornoHormigonWindow(object):
    def __init__(self, uiapp):
        self._uiapp = uiapp
        self._win = None
        self._ejes = []
        self._win = XamlReader.Parse(_build_xaml())
        self._cmb = self._win.FindName(u"CmbEje")
        self._txt_subtitle = self._win.FindName(u"TxtSubtitle")
        self._txt_status = self._win.FindName(u"TxtStatus")

        self._handler = _GenerarContornoHandler(weakref.ref(self))
        self._ext_event = ExternalEvent.Create(self._handler)

        self._wire_events()
        _prepare_window(self._win, uiapp)
        self._refresh_summary()
        self._load_ejes()

    def _wire_events(self):
        self._win.FindName(u"BtnGenerar").Click += RoutedEventHandler(
            self._on_generar,
        )
        self._win.FindName(u"BtnClose").Click += RoutedEventHandler(
            self._on_close,
        )
        self._win.KeyDown += KeyEventHandler(self._on_key_down)
        self._win.Closed += EventHandler(self._on_closed)

    def _on_key_down(self, sender, args):
        if args.Key == Key.Escape:
            self._win.Close()

    def _on_generar(self, sender, args):
        self._set_status(u"Procesando…")
        self._ext_event.Raise()

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

    def _refresh_summary(self):
        uidoc = self._uiapp.ActiveUIDocument
        if uidoc is None:
            self._txt_subtitle.Text = u"No hay documento activo."
            return
        view = uidoc.ActiveView
        ok, msg = vista_permitida(view)
        if not ok:
            self._txt_subtitle.Text = msg
            return
        n = len(recoger_hormigon_en_vista(uidoc.Document, view))
        try:
            vname = _as_unicode(view.Name)
        except Exception:
            vname = u"Vista"
        self._txt_subtitle.Text = (
            u"Vista: {0} · {1} elemento(s) de hormigón (Concrete) visibles.".format(
                vname, n
            )
        )

    def _load_ejes(self):
        self._cmb.Items.Clear()
        self._ejes = []
        uidoc = self._uiapp.ActiveUIDocument
        if uidoc is None:
            return
        for nombre, grid in listar_ejes_modelo(uidoc.Document):
            item = ComboBoxItem()
            item.Content = nombre
            self._cmb.Items.Add(item)
            self._ejes.append(grid)
        if self._cmb.Items.Count > 0:
            self._cmb.SelectedIndex = 0
        else:
            self._set_status(u"No hay ejes (Grids) en el modelo.")

    def _grid_seleccionado(self):
        idx = -1
        try:
            idx = int(self._cmb.SelectedIndex)
        except Exception:
            idx = -1
        if idx < 0 or idx >= len(self._ejes):
            return None
        return self._ejes[idx]

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
        if hasattr(w, "_win") and w._win is not None and w._win.IsVisible:
            return w
    except Exception:
        pass
    return None


def _focus_existing(win):
    try:
        if win._win.WindowState == WindowState.Minimized:
            win._win.WindowState = WindowState.Normal
        win._win.Activate()
    except Exception:
        pass
    try:
        TaskDialog.Show(_DIALOG_TITLE, u"La herramienta ya está en ejecución.")
    except Exception:
        pass


def show_contorno_window(revit):
    uiapp = revit
    existing = _existing_window()
    if existing is not None:
        _focus_existing(existing)
        return

    win = ContornoHormigonWindow(uiapp)
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, win)
    except Exception:
        pass
    win.show()
