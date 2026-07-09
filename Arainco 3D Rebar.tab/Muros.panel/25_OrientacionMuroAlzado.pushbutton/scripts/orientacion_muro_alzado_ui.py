# -*- coding: utf-8 -*-
"""UI WPF — helpers de orientación de muros. Paquete portable."""

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
from System.ComponentModel import CancelEventHandler
from System.Windows import (
    RoutedEventHandler,
    WindowState,
)
from System.Windows.Input import Key, KeyEventHandler
from System.Windows.Markup import XamlReader
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler

from bimtools_ui_tokens import FG_BODY, FG_MUTED, FONT_SIZE_BODY, FONT_SIZE_HINT
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_wpf_shell import build_simple_tool_xaml
from orientacion_muro_alzado import (
    _TITULO,
    _as_unicode,
    ejecutar_actualizar_helpers,
    ejecutar_dibujar,
    ejecutar_eliminar_helpers,
    mostrar_aviso,
    resumen_vista,
)
from revit_wpf_window_position import (
    bind_center_wpf_on_revit_monitor,
    position_wpf_window_center_on_monitor,
    revit_main_hwnd,
)

_WINDOW_TITLE = _TITULO
_SINGLETON_KEY = u"Arainco_OrientacionMuroAlzado_UI"
_TX_DIBUJAR = u"Arainco: Helpers orientación muro en alzado"
_TX_ACTUALIZAR = u"Arainco: Actualizar helpers orientación muro en alzado"
_TX_CERRAR_LIMPIAR = u"Arainco: Limpiar helpers orientación muro al cerrar"

_TEXT_INTRO = (
    u"Dibuja flechas verdes centradas a mitad de altura de cada muro visible "
    u"en una vista Building Section (sección de edificio)."
)
_TEXT_ACTUALIZAR = (
    u"Tras ajustar la orientación de un muro en Revit, use \u00abActualizar "
    u"helpers\u00bb para redibujar las flechas seg\u00fan la Location Line actual."
)
_TEXT_FLECHA = (
    u"\u2192  Flecha \u2014 direcci\u00f3n Location Line (0 \u2192 1)"
)
_TEXT_ELIMINAR = (
    u"Al cerrar la ventana se eliminan los helpers dibujados en esta sesi\u00f3n."
)

_BODY_XAML_TEMPLATE = u"""
<StackPanel>
  <TextBlock TextWrapping="Wrap" Foreground="{fg}" FontSize="{fs}" LineHeight="17"
             Text="__TEXT_INTRO__"/>
  <TextBlock Margin="0,12,0,6" Style="{{StaticResource LabelSmall}}" Text="Leyenda"/>
  <StackPanel Margin="10,0,0,0">
    <TextBlock TextWrapping="Wrap" Foreground="{fg}" FontSize="{fs}" LineHeight="17"
               Text="__TEXT_FLECHA__"/>
  </StackPanel>
  <TextBlock Margin="0,12,0,0" TextWrapping="Wrap" Foreground="{fg}" FontSize="{fs}" LineHeight="17"
             Text="__TEXT_ACTUALIZAR__"/>
  <TextBlock Margin="0,8,0,0" TextWrapping="Wrap" Foreground="{fg_lo}" FontSize="{fs_lo}" LineHeight="15"
             Text="__TEXT_ELIMINAR__"/>
  <Button x:Name="BtnDibujar" Margin="0,14,0,0"
          Content="Dibujar helpers"
          Style="{{StaticResource BtnPrimary}}"
          HorizontalAlignment="Stretch" MinHeight="36"/>
  <Button x:Name="BtnActualizar" Margin="0,8,0,0"
          Content="Actualizar helpers"
          Style="{{StaticResource BtnSelectOutline}}"
          HorizontalAlignment="Stretch" MinHeight="32"
          IsEnabled="False"/>
</StackPanel>
""".format(fg=FG_BODY, fs=FONT_SIZE_BODY, fg_lo=FG_MUTED, fs_lo=FONT_SIZE_HINT)

_FOOTER_ACTIONS_XAML = u"""
<Button x:Name="BtnClose" Content="Cerrar"
        Style="{StaticResource BtnSelectOutline}" MinWidth="108"/>
"""


def _escape_xaml(text):
    s = _as_unicode(text)
    return (
        s.replace(u"&", u"&amp;")
        .replace(u"<", u"&lt;")
        .replace(u">", u"&gt;")
        .replace(u'"', u"&quot;")
    )


def _build_body_xaml():
    xaml = _BODY_XAML_TEMPLATE
    xaml = xaml.replace(u"__TEXT_INTRO__", _escape_xaml(_TEXT_INTRO))
    xaml = xaml.replace(u"__TEXT_FLECHA__", _escape_xaml(_TEXT_FLECHA))
    xaml = xaml.replace(u"__TEXT_ACTUALIZAR__", _escape_xaml(_TEXT_ACTUALIZAR))
    xaml = xaml.replace(u"__TEXT_ELIMINAR__", _escape_xaml(_TEXT_ELIMINAR))
    return xaml


def _build_xaml():
    return build_simple_tool_xaml(
        title=_WINDOW_TITLE,
        styles_xml=BIMTOOLS_DARK_STYLES_XML,
        body_xaml=_build_body_xaml(),
        footer_actions_xaml=_FOOTER_ACTIONS_XAML,
        width=520,
        resize_mode=u"NoResize",
        size_to_content_height=True,
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


class _DibujarHelpersHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        win = self._window_ref()
        if win is None:
            return
        uidoc = uiapp.ActiveUIDocument
        ok, msg, nuevos_ids = ejecutar_dibujar(uidoc, win._helper_ids)
        if ok:
            win._helper_ids = list(nuevos_ids or [])
            win._set_status(msg)
            win._refresh_summary()
            win._update_actualizar_enabled()
        else:
            win._set_status(u"Error: {0}".format(msg))
            try:
                mostrar_aviso(uiapp, msg)
            except Exception:
                pass

    def GetName(self):
        return _TX_DIBUJAR


class _ActualizarHelpersHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        win = self._window_ref()
        if win is None:
            return
        uidoc = uiapp.ActiveUIDocument
        ok, msg, nuevos_ids = ejecutar_actualizar_helpers(uidoc, win._helper_ids)
        if ok:
            win._helper_ids = list(nuevos_ids or [])
            win._set_status(msg)
            win._refresh_summary()
            win._update_actualizar_enabled()
        else:
            win._set_status(u"Error: {0}".format(msg))
            try:
                mostrar_aviso(uiapp, msg)
            except Exception:
                pass

    def GetName(self):
        return _TX_ACTUALIZAR


class _CerrarLimpiarHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref
        self._line_ids = []

    def set_line_ids(self, line_ids):
        self._line_ids = list(line_ids or [])

    def Execute(self, uiapp):
        ids = list(self._line_ids)
        self._line_ids = []
        try:
            uidoc = uiapp.ActiveUIDocument
            if uidoc is not None and ids:
                ejecutar_eliminar_helpers(
                    uidoc, ids, refrescar=False, incluir_legado=False
                )
        except Exception:
            pass
        win = self._window_ref()
        if win is None:
            return
        try:
            win._closing_after_cleanup = True
            if win._win is not None:
                win._win.Close()
        except Exception:
            pass

    def GetName(self):
        return _TX_CERRAR_LIMPIAR


class OrientacionMuroAlzadoWindow(object):
    def __init__(self, uiapp):
        self._uiapp = uiapp
        self._win = None
        self._helper_ids = []
        self._closing_after_cleanup = False
        self._win = XamlReader.Parse(_build_xaml())
        self._txt_subtitle = self._win.FindName(u"TxtSubtitle")
        self._txt_status = self._win.FindName(u"TxtStatus")
        self._btn_actualizar = self._win.FindName(u"BtnActualizar")

        self._handler_dibujar = _DibujarHelpersHandler(weakref.ref(self))
        self._ext_dibujar = ExternalEvent.Create(self._handler_dibujar)
        self._handler_actualizar = _ActualizarHelpersHandler(weakref.ref(self))
        self._ext_actualizar = ExternalEvent.Create(self._handler_actualizar)
        self._handler_cerrar_limpiar = _CerrarLimpiarHandler(weakref.ref(self))
        self._ext_cerrar_limpiar = ExternalEvent.Create(self._handler_cerrar_limpiar)

        self._wire_events()
        _prepare_window(self._win, uiapp)
        self._refresh_summary()

    def _wire_events(self):
        self._win.FindName(u"BtnDibujar").Click += RoutedEventHandler(
            self._on_dibujar,
        )
        self._win.FindName(u"BtnActualizar").Click += RoutedEventHandler(
            self._on_actualizar,
        )
        self._win.FindName(u"BtnClose").Click += RoutedEventHandler(
            self._on_close,
        )
        self._win.KeyDown += KeyEventHandler(self._on_key_down)
        self._win.Closing += CancelEventHandler(self._on_closing)
        self._win.Closed += EventHandler(self._on_closed)

    def _on_key_down(self, sender, args):
        if args.Key == Key.Escape:
            self._request_close()

    def _request_close(self):
        if self._win is not None:
            self._win.Close()

    def _on_closing(self, sender, args):
        if self._closing_after_cleanup:
            return
        if not self._helper_ids:
            return
        args.Cancel = True
        self._handler_cerrar_limpiar.set_line_ids(list(self._helper_ids))
        self._helper_ids = []
        self._ext_cerrar_limpiar.Raise()

    def _on_close(self, sender, args):
        self._request_close()

    def _on_dibujar(self, sender, args):
        self._set_status(u"Dibujando helpers…")
        self._ext_dibujar.Raise()

    def _on_actualizar(self, sender, args):
        self._set_status(u"Actualizando helpers…")
        self._ext_actualizar.Raise()

    def _update_actualizar_enabled(self):
        if self._btn_actualizar is not None:
            self._btn_actualizar.IsEnabled = bool(self._helper_ids)

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
        if self._txt_subtitle is not None:
            self._txt_subtitle.Text = resumen_vista(uidoc)

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
        mostrar_aviso(win._uiapp, u"La herramienta ya está en ejecución.")
    except Exception:
        pass


def show_orientacion_window(revit):
    uiapp = revit
    existing = _existing_window()
    if existing is not None:
        _focus_existing(existing)
        return

    win = OrientacionMuroAlzadoWindow(uiapp)
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, win)
    except Exception:
        pass
    win.show()
