# -*- coding: utf-8 -*-
"""Diálogos WPF (tema oscuro BIMTools) para View Unobscured en la vista activa."""

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")

from System.Windows import RoutedEventHandler
from System.Windows.Input import Key, KeyEventHandler
from System.Windows.Markup import XamlReader

try:
    from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
except Exception:
    BIMTOOLS_DARK_STYLES_XML = u""

_ACTION_DIALOG_XAML = u"""<Window
  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  Title="__TITLE__"
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
__BIMTOOLS_DARK_STYLES__
  </Window.Resources>
  <Border CornerRadius="8" Background="#071018" BorderBrush="#21465C"
          BorderThickness="1" Padding="22,20">
    <StackPanel>
      <TextBlock Text="__TITLE__" Foreground="#E8F4F8" FontSize="16" FontWeight="Bold"/>
      <TextBlock Margin="0,6,0,0" Text="__SUBTITLE__"
                 Foreground="#95B8CC" FontSize="11" TextWrapping="Wrap"/>
      <Border Margin="0,14,0,0" Background="#0a1620" BorderBrush="#21465C"
              BorderThickness="1" CornerRadius="6" Padding="12,10">
        <TextBlock Text="__SUMMARY__" TextWrapping="Wrap"
                   Foreground="#E8F4F8" FontSize="12" LineHeight="18"/>
      </Border>
      <TextBlock Margin="0,16,0,10" Text="Elige una acción:"
                 Foreground="#95B8CC" FontSize="11" FontWeight="SemiBold"/>
      <Button x:Name="BtnApply" Content="Aplicar View Unobscured"
              Style="{StaticResource BtnPrimary}" HorizontalAlignment="Stretch"
              MinHeight="34" Margin="0,0,0,8"/>
      <Button x:Name="BtnRemove" Content="Quitar View Unobscured"
              Style="{StaticResource BtnSelectOutline}" HorizontalAlignment="Stretch"
              MinHeight="34" Margin="0,0,0,8"/>
      <Button x:Name="BtnStatus" Content="Solo consultar estado"
              Style="{StaticResource BtnSelectOutline}" HorizontalAlignment="Stretch"
              MinHeight="34" Margin="0,0,0,8"/>
      <StackPanel Margin="0,8,0,0" Orientation="Horizontal" HorizontalAlignment="Right">
        <Button x:Name="BtnCancel" Content="Cancelar"
                Style="{StaticResource BtnSelectOutline}" MinWidth="108"/>
      </StackPanel>
    </StackPanel>
  </Border>
</Window>"""

_INFO_DIALOG_XAML = u"""<Window
  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  Title="__TITLE__"
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
__BIMTOOLS_DARK_STYLES__
  </Window.Resources>
  <Border CornerRadius="8" Background="#071018" BorderBrush="#21465C"
          BorderThickness="1" Padding="22,20">
    <StackPanel>
      <TextBlock Text="__TITLE__" Foreground="#E8F4F8" FontSize="16" FontWeight="Bold"/>
      <Border Margin="0,14,0,0" Background="#0a1620" BorderBrush="#21465C"
              BorderThickness="1" CornerRadius="6" Padding="12,10">
        <TextBlock Text="__CONTENT__" TextWrapping="Wrap"
                   Foreground="#E8F4F8" FontSize="12" LineHeight="18"/>
      </Border>
      <StackPanel Margin="0,22,0,0" Orientation="Horizontal" HorizontalAlignment="Right">
        <Button x:Name="BtnOk" Content="Aceptar" IsDefault="True"
                Style="{StaticResource BtnPrimary}" MinWidth="108"/>
      </StackPanel>
    </StackPanel>
  </Border>
</Window>"""


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _escape_xaml(text):
    s = _as_unicode(text)
    return (
        s.replace(u"&", u"&amp;")
        .replace(u"<", u"&lt;")
        .replace(u">", u"&gt;")
        .replace(u'"', u"&quot;")
    )


def _build_action_xaml(title, subtitle, summary):
    xaml = _ACTION_DIALOG_XAML.replace(
        u"__BIMTOOLS_DARK_STYLES__", BIMTOOLS_DARK_STYLES_XML
    )
    xaml = xaml.replace(u"__TITLE__", _escape_xaml(title))
    xaml = xaml.replace(u"__SUBTITLE__", _escape_xaml(subtitle))
    xaml = xaml.replace(u"__SUMMARY__", _escape_xaml(summary))
    return xaml


def _build_info_xaml(title, content):
    xaml = _INFO_DIALOG_XAML.replace(
        u"__BIMTOOLS_DARK_STYLES__", BIMTOOLS_DARK_STYLES_XML
    )
    xaml = xaml.replace(u"__TITLE__", _escape_xaml(title))
    xaml = xaml.replace(u"__CONTENT__", _escape_xaml(content))
    return xaml


def _attach_revit_owner(win, uiapp):
    if win is None or uiapp is None:
        return
    try:
        from System.Windows.Interop import WindowInteropHelper
        from revit_wpf_window_position import revit_main_hwnd

        hwnd = revit_main_hwnd(uiapp)
        if hwnd is not None:
            WindowInteropHelper(win).Owner = hwnd
    except Exception:
        pass


def _prepare_window(win, uiapp):
    if win is None:
        return
    try:
        from revit_wpf_window_position import (
            bind_center_wpf_on_revit_monitor,
            position_wpf_window_center_on_monitor,
            revit_main_hwnd,
        )

        hwnd = revit_main_hwnd(uiapp)
        bind_center_wpf_on_revit_monitor(win, hwnd)
        position_wpf_window_center_on_monitor(win, hwnd)
    except Exception:
        pass
    _attach_revit_owner(win, uiapp)


def _bind_escape_close(win, on_cancel):
    def _on_key(sender, args):
        if args.Key == Key.Escape:
            on_cancel(sender, args)
            args.Handled = True

    win.PreviewKeyDown += KeyEventHandler(_on_key)


def show_rebar_unobscured_action_dialog(title, subtitle, summary, uiapp=None):
    """
    Diálogo de acción estilo Armado Muros / BIMTools.

    Devuelve ``"apply"``, ``"remove"``, ``"status"`` o ``None``.
    """
    try:
        win = XamlReader.Parse(_build_action_xaml(title, subtitle, summary))
    except Exception:
        return None

    _prepare_window(win, uiapp)
    result = [None]

    def _choose(value):
        def _handler(sender, args):
            result[0] = value
            try:
                win.Close()
            except Exception:
                pass

        return _handler

    def _cancel(sender, args):
        result[0] = None
        try:
            win.Close()
        except Exception:
            pass

    try:
        mapping = (
            (u"BtnApply", u"apply"),
            (u"BtnRemove", u"remove"),
            (u"BtnStatus", u"status"),
        )
        for btn_name, value in mapping:
            btn = win.FindName(btn_name)
            if btn is not None:
                btn.Click += RoutedEventHandler(_choose(value))
        btn_cancel = win.FindName(u"BtnCancel")
        if btn_cancel is not None:
            btn_cancel.Click += RoutedEventHandler(_cancel)
        _bind_escape_close(win, _cancel)
        win.ShowDialog()
    except Exception:
        return None

    return result[0]


def show_info_dialog(title, content, uiapp=None):
    """Diálogo informativo con un solo botón Aceptar."""
    try:
        win = XamlReader.Parse(_build_info_xaml(title, content))
    except Exception:
        return

    _prepare_window(win, uiapp)

    def _close(sender, args):
        try:
            win.Close()
        except Exception:
            pass

    try:
        btn_ok = win.FindName(u"BtnOk")
        if btn_ok is not None:
            btn_ok.Click += RoutedEventHandler(_close)
        _bind_escape_close(win, _close)
        win.ShowDialog()
    except Exception:
        pass
