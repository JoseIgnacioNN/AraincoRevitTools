# -*- coding: utf-8 -*-
"""Diálogo modal WPF — shell estándar BIMTools (cinta blanca + cuerpo oscuro)."""

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")

from System.Windows import RoutedEventHandler
from System.Windows.Input import Key, KeyEventHandler
from System.Windows.Markup import XamlReader

from bimtools_ui_tokens import FG_BODY, FONT_SIZE_BODY
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_wpf_shell import build_simple_tool_xaml


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


def _build_message_xaml(title, instruction, content, ok_text):
    instruction = _as_unicode(instruction).strip()
    content = _as_unicode(content).strip()
    subtitle = u""
    if content:
        subtitle = instruction
        body_text = content
    else:
        body_text = instruction

    body_xaml = u"""
<StackPanel>
  <TextBlock TextWrapping="Wrap" Foreground="{fg}" FontSize="{fs}" LineHeight="17"
             Text="{text}"/>
</StackPanel>
""".format(
        fg=FG_BODY,
        fs=FONT_SIZE_BODY,
        text=_escape_xaml(body_text),
    )

    footer_xaml = u"""
<Button x:Name="BtnOk" Content="{ok}" IsDefault="True"
        Style="{{StaticResource BtnPrimary}}" MinWidth="108"/>
""".format(ok=_escape_xaml(ok_text))

    xaml = build_simple_tool_xaml(
        title=title,
        styles_xml=BIMTOOLS_DARK_STYLES_XML,
        body_xaml=body_xaml,
        footer_actions_xaml=footer_xaml,
        width=520,
        resize_mode=u"NoResize",
        size_to_content_height=True,
    )
    return xaml, subtitle


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


def show_message_dialog(
    title,
    instruction,
    content=u"",
    ok_text=u"Entendido",
    hwnd_revit=None,
    uiapp=None,
):
    """Diálogo modal informativo (solo Aceptar), shell estándar BIMTools."""
    try:
        xaml, subtitle = _build_message_xaml(title, instruction, content, ok_text)
        win = XamlReader.Parse(xaml)
    except Exception:
        return False

    if subtitle:
        try:
            txt_sub = win.FindName(u"TxtSubtitle")
            if txt_sub is not None:
                txt_sub.Text = subtitle
        except Exception:
            pass

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
