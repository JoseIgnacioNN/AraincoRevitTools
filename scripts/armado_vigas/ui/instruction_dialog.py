# -*- coding: utf-8 -*-
"""
Diálogos modales de Armado vigas — shell Arainco / BIMTools.

Prioridad:
1. ``bimtools_instruction_dialog`` (cinta blanca + cuerpo oscuro, shell estándar)
2. Fallback local WPF (tema oscuro BIMTools)
3. ``TaskDialog`` de Revit (último recurso)
"""

from __future__ import print_function

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")

from System.Windows import RoutedEventHandler
from System.Windows.Input import Key, KeyEventHandler
from System.Windows.Markup import XamlReader

DIALOG_TITLE = u"Arainco: Armado vigas"

try:
    from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
except Exception:
    BIMTOOLS_DARK_STYLES_XML = u""

_OK_ONLY_XAML = u"""<Window
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

_OK_CANCEL_XAML = u"""<Window
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


def _resolve_hwnd(hwnd_revit=None, uiapp=None):
    if hwnd_revit is not None:
        return hwnd_revit
    try:
        from revit_wpf_window_position import revit_main_hwnd

        if uiapp is not None:
            return revit_main_hwnd(uiapp)
    except Exception:
        pass
    return None


def _position_win(win, hwnd_revit=None, uiapp=None):
    try:
        from revit_wpf_window_position import (
            bind_center_wpf_on_revit_monitor,
            position_wpf_window_center_on_monitor,
        )

        hwnd = _resolve_hwnd(hwnd_revit, uiapp)
        bind_center_wpf_on_revit_monitor(win, hwnd)
        position_wpf_window_center_on_monitor(win, hwnd)
    except Exception:
        pass
    if uiapp is None:
        return
    try:
        from System.Windows.Interop import WindowInteropHelper
        from revit_wpf_window_position import revit_main_hwnd

        hwnd = revit_main_hwnd(uiapp)
        if hwnd is not None:
            WindowInteropHelper(win).Owner = hwnd
    except Exception:
        pass


def _fill_template(template, title, instruction, content, ok_text, cancel_text=None):
    xaml = template.replace(u"__BIMTOOLS_DARK_STYLES__", BIMTOOLS_DARK_STYLES_XML)
    xaml = xaml.replace(u"__TITLE__", _escape_xaml(title))
    xaml = xaml.replace(u"__INSTRUCTION__", _escape_xaml(instruction))
    # Ocultar bloque de content si vacío: dejar texto vacío (TextBlock colapsa en altura baja).
    xaml = xaml.replace(u"__CONTENT__", _escape_xaml(content or u""))
    xaml = xaml.replace(u"__OK__", _escape_xaml(ok_text))
    if cancel_text is not None:
        xaml = xaml.replace(u"__CANCEL__", _escape_xaml(cancel_text))
    return xaml


def _taskdialog_fallback(title, instruction, content=u""):
    try:
        from Autodesk.Revit.UI import TaskDialog

        body = _as_unicode(instruction or u"")
        c = _as_unicode(content or u"").strip()
        if c:
            body = body + u"\n\n" + c if body else c
        TaskDialog.Show(_as_unicode(title) or DIALOG_TITLE, body or u"")
        return True
    except Exception:
        return False


def _show_local_ok_only(
    title, instruction, content, ok_text, hwnd_revit=None, uiapp=None,
):
    try:
        win = XamlReader.Parse(
            _fill_template(
                _OK_ONLY_XAML, title, instruction, content, ok_text,
            )
        )
    except Exception:
        return _taskdialog_fallback(title, instruction, content)

    _position_win(win, hwnd_revit, uiapp)

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
        btn = win.FindName(u"BtnOk")
        if btn is not None:
            btn.Click += RoutedEventHandler(_accept)
        win.PreviewKeyDown += KeyEventHandler(_on_key)
        # Content vacío: ocultar
        try:
            tb = win.FindName(u"TxtContent")
            if tb is not None and not _as_unicode(content).strip():
                from System.Windows import Visibility

                tb.Visibility = Visibility.Collapsed
        except Exception:
            pass
        win.ShowDialog()
        return True
    except Exception:
        return _taskdialog_fallback(title, instruction, content)


def _show_local_ok_cancel(
    title,
    instruction,
    content,
    ok_text,
    cancel_text,
    hwnd_revit=None,
    uiapp=None,
):
    try:
        win = XamlReader.Parse(
            _fill_template(
                _OK_CANCEL_XAML,
                title,
                instruction,
                content,
                ok_text,
                cancel_text=cancel_text,
            )
        )
    except Exception:
        return False

    _position_win(win, hwnd_revit, uiapp)
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
        try:
            tb = win.FindName(u"TxtContent")
            if tb is not None and not _as_unicode(content).strip():
                from System.Windows import Visibility

                tb.Visibility = Visibility.Collapsed
        except Exception:
            pass
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
    """Informativo (solo botón principal). Estilo extensión / BIMTools."""
    title = _as_unicode(title) or DIALOG_TITLE
    instruction = _as_unicode(instruction)
    content = _as_unicode(content)
    try:
        from bimtools_instruction_dialog import show_message_dialog as _shell

        return bool(
            _shell(
                title,
                instruction,
                content=content,
                ok_text=ok_text,
                hwnd_revit=hwnd_revit,
                uiapp=uiapp,
            )
        )
    except Exception:
        pass
    return _show_local_ok_only(
        title, instruction, content, ok_text, hwnd_revit=hwnd_revit, uiapp=uiapp,
    )


def show_ok_cancel_dialog(
    title,
    instruction,
    content=u"",
    ok_text=u"Aceptar",
    cancel_text=u"Cancelar",
    hwnd_revit=None,
    uiapp=None,
):
    """Confirmar / cancelar. Estilo extensión / BIMTools. ``True`` = aceptar."""
    title = _as_unicode(title) or DIALOG_TITLE
    instruction = _as_unicode(instruction)
    content = _as_unicode(content)
    try:
        from bimtools_instruction_dialog import show_ok_cancel_dialog as _shell

        return bool(
            _shell(
                title,
                instruction,
                content=content,
                ok_text=ok_text,
                cancel_text=cancel_text,
                hwnd_revit=hwnd_revit,
                uiapp=uiapp,
            )
        )
    except Exception:
        pass
    return _show_local_ok_cancel(
        title,
        instruction,
        content,
        ok_text,
        cancel_text,
        hwnd_revit=hwnd_revit,
        uiapp=uiapp,
    )


def show_info(
    instruction,
    content=u"",
    title=None,
    ok_text=u"Entendido",
    uiapp=None,
    hwnd_revit=None,
):
    """Aviso / resultado (cabecera Arainco)."""
    return show_message_dialog(
        title or DIALOG_TITLE,
        instruction,
        content=content,
        ok_text=ok_text,
        hwnd_revit=hwnd_revit,
        uiapp=uiapp,
    )


def show_yes_no(
    instruction,
    content=u"",
    title=None,
    yes_text=u"Sí",
    no_text=u"No",
    uiapp=None,
    hwnd_revit=None,
):
    """Confirmación Sí/No; ``True`` si elige Sí."""
    return show_ok_cancel_dialog(
        title or DIALOG_TITLE,
        instruction,
        content=content,
        ok_text=yes_text,
        cancel_text=no_text,
        hwnd_revit=hwnd_revit,
        uiapp=uiapp,
    )
