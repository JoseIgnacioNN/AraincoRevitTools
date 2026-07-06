# -*- coding: utf-8 -*-
"""UI WPF — selector de categorías para cotas de anchura (muro / fundación)."""

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
from System.Windows.Input import Key, KeyEventHandler
from System.Windows.Markup import XamlReader
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler, TaskDialog

from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from cota_anchura_selector import _TITULO, _TX_COTAR, ejecutar_cotas, resumen_vista
from revit_wpf_window_position import (
    bind_center_wpf_on_revit_monitor,
    position_wpf_window_center_on_monitor,
    revit_main_hwnd,
)

_WINDOW_TITLE = _TITULO
_SINGLETON_KEY = u"Arainco_CotaAnchuraSelector_UI"

_HELP_TEXT = (
    u"Crea cotas de anchura en planta entre las caras laterales de cada elemento, "
    u"posicionadas a mitad del largo según la LocationCurve (muro o muro host de la zapata).\n\n"
    u"Use la preselección en el modelo o el selector múltiple al pulsar «Cotar anchura»."
)

XAML = u"""
<Window
    x:Name="CotaAnchuraWin"
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="__WINDOW_TITLE__"
    Width="500"
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
    <Style x:Key="BimToolsCheckCategory" TargetType="CheckBox">
      <Setter Property="Foreground" Value="#E8F4F8"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Margin" Value="2,6,2,6"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CheckBox">
            <StackPanel Orientation="Horizontal">
              <Border x:Name="Box" Width="16" Height="16"
                      Background="#050E18" BorderBrush="#1A3A4D" BorderThickness="1"
                      CornerRadius="3" Margin="0,0,10,0" VerticalAlignment="Center">
                <Path x:Name="CheckMark" Visibility="Collapsed"
                      Data="M 2,8 L 6,12 L 14,3" Stroke="#5BC0DE" StrokeThickness="2"
                      StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
              </Border>
              <ContentPresenter VerticalAlignment="Center" RecognizesAccessKey="True"/>
            </StackPanel>
            <ControlTemplate.Triggers>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="CheckMark" Property="Visibility" Value="Visible"/>
                <Setter TargetName="Box" Property="BorderBrush" Value="#5BC0DE"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Box" Property="BorderBrush" Value="#5BC0DE"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
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
      <TextBlock Margin="0,14,0,6" Text="Categorías a cotar"
                 Foreground="#95B8CC" FontSize="11" FontWeight="SemiBold"/>
      <Border Background="#0a1620" BorderBrush="#21465C"
              BorderThickness="1" CornerRadius="6" Padding="12,8">
        <StackPanel>
          <CheckBox x:Name="ChkFundacion" Style="{StaticResource BimToolsCheckCategory}"
                    Content="Fundación estructural (zapata de muro)" IsChecked="True"/>
          <CheckBox x:Name="ChkMuro" Style="{StaticResource BimToolsCheckCategory}"
                    Content="Muro" IsChecked="True"/>
        </StackPanel>
      </Border>
      <Button x:Name="BtnCotar" Margin="0,16,0,0"
              Content="Cotar anchura"
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
    s = text if text is not None else u""
    try:
        s = unicode(s)  # noqa: F821 — IronPython 2
    except NameError:
        s = str(s)
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


class _CotarAnchuraHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        win_ctrl = self._window_ref()
        if win_ctrl is None:
            return

        incluir_fundacion = win_ctrl._chk_fundacion()
        incluir_muro = win_ctrl._chk_muro()

        if not incluir_fundacion and not incluir_muro:
            win_ctrl._set_status(u"Marque al menos una categoría.")
            try:
                TaskDialog.Show(_TITULO, u"Marque al menos una categoría para cotar.")
            except Exception:
                pass
            return

        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            win_ctrl._set_status(u"No hay documento activo.")
            return

        ok, msg = ejecutar_cotas(uidoc, incluir_fundacion, incluir_muro)
        if not msg:
            if not ok:
                win_ctrl._set_status(u"Operación cancelada.")
            return

        if ok:
            win_ctrl._set_status(msg)
        else:
            win_ctrl._set_status(u"Error: {0}".format(msg))
            try:
                TaskDialog.Show(_TITULO, msg)
            except Exception:
                pass

    def GetName(self):
        return _TX_COTAR


class CotaAnchuraSelectorWindow(object):
    def __init__(self, uiapp):
        self._uiapp = uiapp
        self._win = XamlReader.Parse(_build_xaml())
        self._txt_subtitle = self._win.FindName(u"TxtSubtitle")
        self._txt_status = self._win.FindName(u"TxtStatus")
        self._chk_fund = self._win.FindName(u"ChkFundacion")
        self._chk_muro_el = self._win.FindName(u"ChkMuro")

        self._handler = _CotarAnchuraHandler(weakref.ref(self))
        self._ext_event = ExternalEvent.Create(self._handler)

        self._wire_events()
        _prepare_window(self._win, uiapp)
        self._refresh_subtitle()

    def _wire_events(self):
        self._win.FindName(u"BtnCotar").Click += RoutedEventHandler(self._on_cotar)
        self._win.FindName(u"BtnClose").Click += RoutedEventHandler(self._on_close)
        self._win.KeyDown += KeyEventHandler(self._on_key_down)
        self._win.Closed += EventHandler(self._on_closed)

    def _on_key_down(self, sender, args):
        if args.Key == Key.Escape:
            self._win.Close()

    def _on_cotar(self, sender, args):
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
            self._txt_status.Text = text if text is not None else u""

    def _refresh_subtitle(self):
        uidoc = self._uiapp.ActiveUIDocument
        if self._txt_subtitle is not None:
            self._txt_subtitle.Text = resumen_vista(uidoc)

    def _chk_fundacion(self):
        try:
            return bool(self._chk_fund.IsChecked)
        except Exception:
            return False

    def _chk_muro(self):
        try:
            return bool(self._chk_muro_el.IsChecked)
        except Exception:
            return False

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
        TaskDialog.Show(_TITULO, u"La herramienta ya está en ejecución.")
    except Exception:
        pass


def show_cota_anchura_window(revit):
    existing = _existing_window()
    if existing is not None:
        _focus_existing(existing)
        return

    win = CotaAnchuraSelectorWindow(revit)
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, win)
    except Exception:
        pass
    win.show()
