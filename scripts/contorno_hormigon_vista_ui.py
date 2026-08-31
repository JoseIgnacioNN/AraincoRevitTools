# -*- coding: utf-8 -*-
"""UI WPF — Arainco: Redibujar contorno."""

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
from System.Windows.Input import Key, KeyEventHandler
from System.Windows.Markup import XamlReader
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler, TaskDialog

from bimtools_ui_tokens import BTN_MANUAL, FG_BODY, FONT_SIZE_BODY
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_wpf_shell import build_simple_tool_xaml
from contorno_hormigon_vista import (
    _DIALOG_TITLE,
    _as_unicode,
    ejecutar_contorno,
    leer_armadura_eje_desde_vista,
    limpiar_seleccion,
    recoger_hormigon_en_vista,
    vista_permitida,
    vista_tiene_parametro_armadura_eje,
)
from revit_wpf_window_position import (
    bind_center_wpf_on_revit_monitor,
    position_wpf_window_center_on_monitor,
    revit_main_hwnd,
)

_TOOL_DIALOG_TITLE = _DIALOG_TITLE
_SINGLETON_KEY = u"Arainco_RedibujarContorno_UI"
_TX_GENERAR = u"Arainco: Redibujar contorno"

_TEXT_INTRO = (
    u"Redibuja el contorno del hormigón visible en la vista activa: une los "
    u"elementos con Material for Model Behavior = Concrete, corta con el plano "
    u"vertical del eje indicado en Armadura_Eje y crea detail lines agrupadas "
    u"con el nombre de la vista (patrón Elevación Eje)."
)
_TEXT_DETALLE = (
    u"Las líneas se dibujan con estilo Medium Lines. Si la vista ya tenía un "
    u"grupo de contorno con el mismo nombre que la vista, se sustituye."
)

_BODY_XAML = u"""
<StackPanel>
  <TextBlock TextWrapping="Wrap" Foreground="{fg}" FontSize="{fs}" LineHeight="17"
             Text="{intro}"/>
  <TextBlock Margin="0,10,0,0" TextWrapping="Wrap" Foreground="{fg}" FontSize="{fs}"
             LineHeight="16" Text="{detalle}"/>
</StackPanel>
""".format(
    fg=FG_BODY,
    fs=FONT_SIZE_BODY,
    intro=_TEXT_INTRO.replace(u'"', u"&quot;"),
    detalle=_TEXT_DETALLE.replace(u'"', u"&quot;"),
)

_FOOTER_LEADING_XAML = (
    u'<Button x:Name="BtnManual" Content="Manual" '
    u'Style="{{StaticResource BtnSelectOutline}}" '
    u'Background="{bg}" MinWidth="96" Padding="8,2" '
    u'ToolTip="Abrir manual de usuario" VerticalAlignment="Center"/>'
).format(bg=BTN_MANUAL)

_FOOTER_ACTIONS_XAML = u"""
<Button x:Name="BtnClose" Content="Cerrar" Margin="0,0,8,0"
        Style="{StaticResource BtnSelectOutline}" MinWidth="108"/>
<Button x:Name="BtnGenerar" Content="Redibujar contorno"
        Style="{StaticResource BtnPrimary}" MinWidth="168"/>
"""


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
            _TOOL_DIALOG_TITLE,
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
        TaskDialog.Show(_TOOL_DIALOG_TITLE, body)
    except Exception:
        pass


def _resolve_manual_path():
    candidates = []
    try:
        import bimtools_paths

        pb = bimtools_paths.get_pushbutton_dir()
        if pb:
            candidates.append(os.path.join(pb, u"manual_usuario.html"))
    except Exception:
        pass
    for path in candidates:
        if path and os.path.isfile(path):
            return path
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
        title=_TOOL_DIALOG_TITLE,
        styles_xml=BIMTOOLS_DARK_STYLES_XML,
        body_xaml=_BODY_XAML,
        footer_leading_xaml=_FOOTER_LEADING_XAML,
        footer_actions_xaml=_FOOTER_ACTIONS_XAML,
        width=520,
        resize_mode=u"NoResize",
        size_to_content_height=True,
    )


class _GenerarContornoHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        win = self._window_ref()
        if win is None:
            return
        uidoc = uiapp.ActiveUIDocument
        ok, msg = ejecutar_contorno(uidoc)
        limpiar_seleccion(uidoc)
        if ok:
            try:
                win._win.Close()
            except Exception:
                pass
            return
        win._set_status(u"Error: {0}".format(msg))
        _mostrar_aviso(uiapp, msg)

    def GetName(self):
        return _TX_GENERAR


class RedibujarContornoWindow(object):
    def __init__(self, uiapp):
        self._uiapp = uiapp
        self._win = XamlReader.Parse(_build_xaml())
        self._txt_subtitle = self._win.FindName(u"TxtSubtitle")
        self._txt_status = self._win.FindName(u"TxtStatus")

        self._handler = _GenerarContornoHandler(weakref.ref(self))
        self._ext_event = ExternalEvent.Create(self._handler)

        self._wire_events()
        _prepare_window(self._win, uiapp)
        self._refresh_summary()

    def _wire_events(self):
        self._win.FindName(u"BtnGenerar").Click += RoutedEventHandler(
            self._on_generar,
        )
        self._win.FindName(u"BtnClose").Click += RoutedEventHandler(
            self._on_close,
        )
        manual = self._win.FindName(u"BtnManual")
        if manual is not None:
            manual.Click += RoutedEventHandler(self._on_manual)
        self._win.KeyDown += KeyEventHandler(self._on_key_down)
        self._win.Closed += EventHandler(self._on_closed)

    def _on_manual(self, sender, args):
        _open_manual(self._uiapp)

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
        if not vista_tiene_parametro_armadura_eje(view):
            self._txt_subtitle.Text = (
                u"Vista: {0} · {1} elemento(s) de hormigón (Concrete) visibles.\n"
                u"Armadura_Eje: la vista no tiene este parámetro.".format(vname, n)
            )
            return
        eje = leer_armadura_eje_desde_vista(view) or u"(vacío)"
        self._txt_subtitle.Text = (
            u"Vista: {0} · {1} elemento(s) de hormigón (Concrete) visibles.\n"
            u"Plano de corte: eje {2} (Armadura_Eje).".format(vname, n, eje)
        )

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


def _focus_existing(win, uiapp):
    try:
        if win._win.WindowState == WindowState.Minimized:
            win._win.WindowState = WindowState.Normal
        win._win.Activate()
    except Exception:
        pass
    _mostrar_aviso(uiapp, u"La herramienta ya está en ejecución.")


def show_contorno_window(revit):
    uiapp = revit
    existing = _existing_window()
    if existing is not None:
        _focus_existing(existing, uiapp)
        return

    win = RedibujarContornoWindow(uiapp)
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, win)
    except Exception:
        pass
    win.show()


def run(revit):
    show_contorno_window(revit)
