# -*- coding: utf-8 -*-
"""Diálogo modal WPF (tema oscuro BIMTools) — confirmación OK / Cancelar."""

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

_DIALOG = u"Arainco: Armado Muros"

_INSTRUCTION_DIALOG_XAML = u"""<Window
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
      <TextBlock x:Name="TxtInstruction" Margin="0,14,0,0" Text="__INSTRUCTION__"
                 TextWrapping="Wrap" Foreground="#E8F4F8" FontSize="12" LineHeight="18"/>
      <TextBlock x:Name="TxtContent" Margin="0,10,0,0" Text="__CONTENT__"
                 TextWrapping="Wrap" Foreground="#95B8CC" FontSize="11" LineHeight="16"/>
      <StackPanel Margin="0,22,0,0" Orientation="Horizontal" HorizontalAlignment="Right">
        <Button x:Name="BtnCancel" Content="__CANCEL__"
                Style="{StaticResource BtnSelectOutline}" MinWidth="108" Margin="0,0,10,0"/>
        <Button x:Name="BtnOk" Content="__OK__" IsDefault="True"
                Style="{StaticResource BtnPrimary}" MinWidth="108"/>
      </StackPanel>
    </StackPanel>
  </Border>
</Window>"""

_MESSAGE_DIALOG_XAML = u"""<Window
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
      <TextBlock x:Name="TxtInstruction" Margin="0,14,0,0" Text="__INSTRUCTION__"
                 TextWrapping="Wrap" Foreground="#E8F4F8" FontSize="12" LineHeight="18"/>
      <TextBlock x:Name="TxtContent" Margin="0,10,0,0" Text="__CONTENT__"
                 TextWrapping="Wrap" Foreground="#95B8CC" FontSize="11" LineHeight="16"/>
      <StackPanel Margin="0,22,0,0" Orientation="Horizontal" HorizontalAlignment="Right">
        <Button x:Name="BtnOk" Content="__OK__" IsDefault="True"
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


def _build_xaml(title, instruction, content, ok_text, cancel_text):
    xaml = _INSTRUCTION_DIALOG_XAML.replace(
        u"__BIMTOOLS_DARK_STYLES__", BIMTOOLS_DARK_STYLES_XML
    )
    xaml = xaml.replace(u"__TITLE__", _escape_xaml(title))
    xaml = xaml.replace(u"__INSTRUCTION__", _escape_xaml(instruction))
    xaml = xaml.replace(u"__CONTENT__", _escape_xaml(content))
    xaml = xaml.replace(u"__OK__", _escape_xaml(ok_text))
    xaml = xaml.replace(u"__CANCEL__", _escape_xaml(cancel_text))
    return xaml


def _build_message_xaml(title, instruction, content, ok_text):
    xaml = _MESSAGE_DIALOG_XAML.replace(
        u"__BIMTOOLS_DARK_STYLES__", BIMTOOLS_DARK_STYLES_XML
    )
    xaml = xaml.replace(u"__TITLE__", _escape_xaml(title))
    xaml = xaml.replace(u"__INSTRUCTION__", _escape_xaml(instruction))
    xaml = xaml.replace(u"__CONTENT__", _escape_xaml(content))
    xaml = xaml.replace(u"__OK__", _escape_xaml(ok_text))
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


def show_ok_cancel_dialog(
    title,
    instruction,
    content=u"",
    ok_text=u"Aceptar",
    cancel_text=u"Cancelar",
    hwnd_revit=None,
    uiapp=None,
):
    """
    Muestra un diálogo modal estilo BIMTools.
    Devuelve ``True`` si el usuario confirma; ``False`` si cancela o cierra.
    """
    try:
        win = XamlReader.Parse(
            _build_xaml(title, instruction, content, ok_text, cancel_text)
        )
    except Exception:
        return False

    try:
        from revit_wpf_window_position import (
            bind_center_wpf_on_revit_monitor,
            position_wpf_window_center_on_monitor,
        )

        bind_center_wpf_on_revit_monitor(win, hwnd_revit)
        position_wpf_window_center_on_monitor(win, hwnd_revit)
    except Exception:
        pass
    _attach_revit_owner(win, uiapp)

    accepted = [False]

    def _accept(sender, args):
        accepted[0] = True
        try:
            win.Close()
        except Exception:
            pass

    def _cancel(sender, args):
        accepted[0] = False
        try:
            win.Close()
        except Exception:
            pass

    def _on_key(sender, args):
        if args.Key == Key.Escape:
            _cancel(sender, args)
            args.Handled = True

    try:
        btn_ok = win.FindName(u"BtnOk")
        btn_cancel = win.FindName(u"BtnCancel")
        if btn_ok is not None:
            btn_ok.Click += RoutedEventHandler(_accept)
        if btn_cancel is not None:
            btn_cancel.Click += RoutedEventHandler(_cancel)
        win.PreviewKeyDown += KeyEventHandler(_on_key)
        win.ShowDialog()
    except Exception:
        return False

    return bool(accepted[0])


def show_message_dialog(
    title,
    instruction,
    content=u"",
    ok_text=u"Entendido",
    hwnd_revit=None,
    uiapp=None,
):
    """Diálogo modal informativo (solo Aceptar), mismo estilo WPF oscuro BIMTools."""
    try:
        win = XamlReader.Parse(
            _build_message_xaml(title, instruction, content, ok_text)
        )
    except Exception:
        return False

    try:
        from revit_wpf_window_position import (
            bind_center_wpf_on_revit_monitor,
            position_wpf_window_center_on_monitor,
        )

        bind_center_wpf_on_revit_monitor(win, hwnd_revit)
        position_wpf_window_center_on_monitor(win, hwnd_revit)
    except Exception:
        pass
    _attach_revit_owner(win, uiapp)

    def _accept(sender, args):
        try:
            win.Close()
        except Exception:
            pass

    def _on_key(sender, args):
        if args.Key == Key.Escape or args.Key == Key.Enter:
            _accept(sender, args)
            args.Handled = True

    try:
        btn_ok = win.FindName(u"BtnOk")
        if btn_ok is not None:
            btn_ok.Click += RoutedEventHandler(_accept)
        win.PreviewKeyDown += KeyEventHandler(_on_key)
        win.ShowDialog()
    except Exception:
        return False

    return True


def show_building_section_view_required(view, uiapp=None):
    """Aviso al inicio si la vista activa no es Building Section."""
    hwnd = None
    try:
        from revit_wpf_window_position import revit_main_hwnd

        if uiapp is not None:
            hwnd = revit_main_hwnd(uiapp)
    except Exception:
        pass
    try:
        from armado_muros_etiqueta_malla import texto_aviso_vista_building_section

        instruction, content = texto_aviso_vista_building_section(view)
    except Exception:
        instruction = (
            u"Esta herramienta solo puede ejecutarse en secciones "
            u"tipo Building Section."
        )
        content = u"Abra una sección de edificio antes de continuar."
    return show_message_dialog(
        _DIALOG,
        instruction,
        content,
        ok_text=u"Entendido",
        hwnd_revit=hwnd,
        uiapp=uiapp,
    )


def show_selection_instructions(uiapp=None):
    """
    Instrucciones previas a la selección en modelo (al ejecutar el botón).
    Devuelve ``True`` si el usuario pulsa Aceptar; ``False`` si cancela.
    """
    hwnd = None
    try:
        from revit_wpf_window_position import revit_main_hwnd

        if uiapp is not None:
            hwnd = revit_main_hwnd(uiapp)
    except Exception:
        pass
    return show_ok_cancel_dialog(
        _DIALOG,
        u"Seleccione uno o más muros a armar.",
        u"Pulse Aceptar para iniciar la selección en el modelo. "
        u"Finalice con la cinta (Finalizar) o cancela con Esc.",
        ok_text=u"Aceptar",
        cancel_text=u"Cancelar",
        hwnd_revit=hwnd,
        uiapp=uiapp,
    )
