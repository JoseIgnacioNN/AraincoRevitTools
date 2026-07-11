# -*- coding: utf-8 -*-
"""UI WPF — selección de niveles para cuantificación de losa."""

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
    Thickness,
    VerticalAlignment,
    WindowState,
)
from System.Windows.Controls import CheckBox
from System.Windows.Input import Key, KeyEventHandler
from System.Windows.Markup import XamlReader
from System.Windows.Media import BrushConverter
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler

from bimtools_instruction_dialog import show_message_dialog
from bimtools_ui_tokens import FG_BODY, FG_MUTED, FONT_SIZE_BODY, FONT_SIZE_HINT
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_wpf_shell import build_simple_tool_xaml
from lib.rebar_schedule_cuantificacion_losa import (
    crear_o_actualizar_cuadros_por_nivel,
    listar_niveles,
)
from revit_wpf_window_position import (
    bind_center_wpf_on_revit_monitor,
    position_wpf_window_center_on_monitor,
    revit_main_hwnd,
)

_WINDOW_TITLE = u"Arainco: Cuantificación losa por nivel"
_SINGLETON_KEY = u"Arainco_CuantificacionLosa_UI"
_TX_GENERAR = u"Arainco: Cuantificación losa por nivel"

_BODY_XAML = u"""
<StackPanel>
  <TextBlock TextWrapping="Wrap" Foreground="{fg}" FontSize="{fs}" LineHeight="17"
             Text="Seleccione los niveles para los que desea generar tablas de cuantificación (4 tablas por nivel: Malla/Sin Malla × Superior/Inferior)."/>
  <DockPanel Margin="0,12,0,8" LastChildFill="False">
    <TextBlock DockPanel.Dock="Left" Style="{{StaticResource LabelSmall}}"
               VerticalAlignment="Center" Text="Niveles del proyecto"/>
    <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
      <Button x:Name="BtnAll" Content="Todos" Margin="0,0,6,0"
              Style="{{StaticResource BtnSelectOutline}}" MinWidth="72" MinHeight="28"
              Padding="10,4"/>
      <Button x:Name="BtnNone" Content="Ninguno"
              Style="{{StaticResource BtnSelectOutline}}" MinWidth="72" MinHeight="28"
              Padding="10,4"/>
    </StackPanel>
  </DockPanel>
  <Border Background="#0a1620" BorderBrush="#21465C" BorderThickness="1"
          CornerRadius="4" Padding="8" Height="300">
    <ScrollViewer VerticalScrollBarVisibility="Auto"
                  HorizontalScrollBarVisibility="Disabled">
      <StackPanel x:Name="PanelLevels"/>
    </ScrollViewer>
  </Border>
  <TextBlock x:Name="TxtSelectionHint" Margin="0,10,0,0"
             Foreground="{fg_lo}" FontSize="{fs_lo}" TextWrapping="Wrap"
             Text=""/>
</StackPanel>
""".format(fg=FG_BODY, fs=FONT_SIZE_BODY, fg_lo=FG_MUTED, fs_lo=FONT_SIZE_HINT)

_FOOTER_ACTIONS_XAML = u"""
<Button x:Name="BtnClose" Content="Cerrar" Margin="0,0,8,0"
        Style="{StaticResource BtnSelectOutline}" MinWidth="108" MinHeight="36"
        Padding="14,8"/>
<Button x:Name="BtnGenerar" Content="Generar tablas"
        Style="{StaticResource BtnPrimary}" MinWidth="140" MinHeight="36"
        Padding="14,8"/>
"""


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _fg_brush():
    try:
        return BrushConverter().ConvertFromString(FG_BODY)
    except Exception:
        return None


def mostrar_aviso(uiapp, instruction, content=u""):
    try:
        hwnd = revit_main_hwnd(uiapp)
        show_message_dialog(
            _WINDOW_TITLE,
            instruction=instruction,
            content=content or u"",
            ok_text=u"Entendido",
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
    except Exception:
        try:
            from Autodesk.Revit.UI import TaskDialog

            TaskDialog.Show(
                _WINDOW_TITLE,
                u"{0}\n\n{1}".format(instruction, content).strip(),
            )
        except Exception:
            print(u"{0}\n{1}".format(instruction, content))


def _build_xaml():
    # Ancho fijo; alto según contenido (sin hueco vacío ni botones recortados).
    return build_simple_tool_xaml(
        title=_WINDOW_TITLE,
        styles_xml=BIMTOOLS_DARK_STYLES_XML,
        body_xaml=_BODY_XAML,
        footer_actions_xaml=_FOOTER_ACTIONS_XAML,
        width=520,
        min_width=520,
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


class _GenerarHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        win = self._window_ref()
        if win is None:
            return
        try:
            win._ejecutar_generacion(uiapp)
        except Exception as ex:
            try:
                win._set_busy(False)
                win._set_status(u"Error: {0}".format(ex))
                mostrar_aviso(uiapp, u"Error al generar tablas.", _as_unicode(ex))
            except Exception:
                pass

    def GetName(self):
        return _TX_GENERAR


class CuantificacionLosaWindow(object):
    def __init__(self, uiapp):
        self._uiapp = uiapp
        self._level_checks = []  # list of (CheckBox, Level)
        self._busy = False

        xaml = _build_xaml()
        self._win = XamlReader.Parse(xaml)
        _prepare_window(self._win, uiapp)

        self._txt_title = self._win.FindName("TxtTitle")
        self._txt_subtitle = self._win.FindName("TxtSubtitle")
        self._txt_status = self._win.FindName("TxtStatus")
        self._txt_hint = self._win.FindName("TxtSelectionHint")
        self._panel_levels = self._win.FindName("PanelLevels")
        self._btn_all = self._win.FindName("BtnAll")
        self._btn_none = self._win.FindName("BtnNone")
        self._btn_generar = self._win.FindName("BtnGenerar")
        self._btn_close = self._win.FindName("BtnClose")

        if self._txt_subtitle is not None:
            self._txt_subtitle.Text = (
                u"Host Category = Floor · Malla / Sin Malla · Superior (F') / Inferior (F)"
            )

        self._handler_generar = _GenerarHandler(weakref.ref(self))
        self._ext_generar = ExternalEvent.Create(self._handler_generar)

        self._populate_levels()
        self._wire_events()
        self._update_selection_hint()
        self._set_status(u"Seleccione niveles y pulse Generar tablas.")

    def _wire_events(self):
        if self._btn_all is not None:
            self._btn_all.Click += self._on_all
        if self._btn_none is not None:
            self._btn_none.Click += self._on_none
        if self._btn_generar is not None:
            self._btn_generar.Click += self._on_generar
        if self._btn_close is not None:
            self._btn_close.Click += self._on_close
        self._win.Closed += EventHandler(self._on_closed)
        self._win.KeyDown += KeyEventHandler(self._on_key_down)

    def _on_key_down(self, sender, args):
        try:
            if args.Key == Key.Escape:
                self._win.Close()
        except Exception:
            pass

    def _populate_levels(self):
        self._level_checks = []
        if self._panel_levels is None:
            return
        try:
            self._panel_levels.Children.Clear()
        except Exception:
            pass

        uidoc = self._uiapp.ActiveUIDocument
        if uidoc is None:
            return
        doc = uidoc.Document
        levels = listar_niveles(doc)
        brush = _fg_brush()

        if not levels:
            empty = CheckBox()
            empty.Content = u"(No hay niveles en el proyecto)"
            empty.IsEnabled = False
            if brush is not None:
                empty.Foreground = brush
            self._panel_levels.Children.Add(empty)
            return

        for lv in levels:
            name = _as_unicode(getattr(lv, "Name", u"")) or u"(sin nombre)"
            chk = CheckBox()
            chk.Content = name
            chk.IsChecked = True
            chk.Margin = Thickness(2, 4, 2, 4)
            chk.VerticalContentAlignment = VerticalAlignment.Center
            if brush is not None:
                chk.Foreground = brush
            chk.Checked += self._on_check_changed
            chk.Unchecked += self._on_check_changed
            self._panel_levels.Children.Add(chk)
            self._level_checks.append((chk, lv))

    def _on_check_changed(self, sender, args):
        self._update_selection_hint()

    def _selected_levels(self):
        out = []
        for chk, lv in self._level_checks:
            try:
                if chk.IsChecked:
                    out.append(lv)
            except Exception:
                continue
        return out

    def _set_all_checked(self, value):
        for chk, _lv in self._level_checks:
            try:
                chk.IsChecked = bool(value)
            except Exception:
                pass
        self._update_selection_hint()

    def _on_all(self, sender, args):
        self._set_all_checked(True)

    def _on_none(self, sender, args):
        self._set_all_checked(False)

    def _update_selection_hint(self):
        n = len(self._selected_levels())
        total = len(self._level_checks)
        tables = n * 4
        text = u"{0} de {1} nivel(es) · {2} tabla(s) a generar/actualizar.".format(
            n, total, tables
        )
        if self._txt_hint is not None:
            self._txt_hint.Text = text

    def _set_status(self, text):
        if self._txt_status is not None:
            self._txt_status.Text = _as_unicode(text)

    def _set_busy(self, busy):
        self._busy = bool(busy)
        enabled = not self._busy
        for btn in (self._btn_generar, self._btn_all, self._btn_none, self._btn_close):
            if btn is not None:
                try:
                    btn.IsEnabled = enabled
                except Exception:
                    pass
        for chk, _lv in self._level_checks:
            try:
                chk.IsEnabled = enabled
            except Exception:
                pass

    def _on_generar(self, sender, args):
        if self._busy:
            return
        selected = self._selected_levels()
        if not selected:
            mostrar_aviso(
                self._uiapp,
                u"Seleccione al menos un nivel.",
                u"Marque uno o más niveles en la lista antes de generar.",
            )
            return
        self._set_busy(True)
        self._set_status(
            u"Generando tablas para {0} nivel(es)…".format(len(selected))
        )
        self._ext_generar.Raise()

    def _ejecutar_generacion(self, uiapp):
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            self._set_busy(False)
            self._set_status(u"No hay documento activo.")
            mostrar_aviso(uiapp, u"No hay documento activo.")
            return

        doc = uidoc.Document
        selected = self._selected_levels()
        if not selected:
            self._set_busy(False)
            self._set_status(u"Sin niveles seleccionados.")
            return

        t = Transaction(doc, _TX_GENERAR)
        t.Start()
        try:
            result = crear_o_actualizar_cuadros_por_nivel(doc, levels=selected)
            if (
                not result.get(u"ok")
                and not result.get(u"created")
                and not result.get(u"updated")
                and not result.get(u"regenerated")
            ):
                t.RollBack()
                self._set_busy(False)
                self._set_status(u"Sin cambios.")
                mostrar_aviso(
                    uiapp,
                    u"No se generaron tablas.",
                    u"Revise la plantilla y los parámetros de filtrado.",
                )
                return
            t.Commit()
        except Exception as ex:
            t.RollBack()
            self._set_busy(False)
            self._set_status(u"Error.")
            mostrar_aviso(uiapp, u"Error al generar tablas.", _as_unicode(ex))
            raise

        template = result.get(u"template")
        n_ok = (
            len(result.get(u"created") or [])
            + len(result.get(u"updated") or [])
            + len(result.get(u"regenerated") or [])
        )
        self._set_busy(False)
        self._set_status(u"Listo: {0} tabla(s) procesada(s).".format(n_ok))
        try:
            self._win.Close()
        except Exception:
            pass

    def _on_close(self, sender, args):
        try:
            self._win.Close()
        except Exception:
            pass

    def _on_closed(self, sender, args):
        try:
            AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, None)
        except Exception:
            pass

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


def show_cuantificacion_losa_window(revit):
    uiapp = revit
    existing = _existing_window()
    if existing is not None:
        _focus_existing(existing)
        return

    win = CuantificacionLosaWindow(uiapp)
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, win)
    except Exception:
        pass
    win.show()
