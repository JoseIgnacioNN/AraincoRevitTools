# -*- coding: utf-8 -*-
"""UI WPF — super selector de elementos por categoría de modelado."""

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
    Thickness,
    WindowState,
)
from System.Windows.Controls import CheckBox, TextBlock
from System.Windows.Input import Key, KeyEventHandler
from System.Windows.Markup import XamlReader
from Autodesk.Revit.DB import ViewSchedule, ViewSheet
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler, TaskDialog

from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from lib.elements_by_category import (
    collect_model_elements_in_view,
    group_elements_by_model_category,
    select_elements_in_model,
    summarize_view_elements,
)
from revit_wpf_window_position import (
    bind_center_wpf_on_revit_monitor,
    position_wpf_window_center_on_monitor,
    revit_main_hwnd,
)

_DIALOG_TITLE = u"Arainco: Super selector por categoría"
_WINDOW_TITLE = _DIALOG_TITLE
_SINGLETON_KEY = u"Arainco_SuperSelectorCategoria_UI"
_TX_SELECT = u"Arainco: Super selector por categoría"

_HELP_TEXT = (
    u"Lista las categorías de modelado con instancias en la vista activa.\n\n"
    u"Marca una o más categorías para seleccionar sus elementos en el modelo. "
    u"Al marcar o desmarcar un checkbox, la selección se actualiza con la "
    u"unión de todas las categorías marcadas."
)

XAML = u"""
<Window
    x:Name="SuperSelectorWin"
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="__WINDOW_TITLE__"
    Width="560"
    MaxHeight="720"
    WindowStartupLocation="Manual"
    Background="Transparent"
    AllowsTransparency="True"
    FontFamily="Segoe UI"
    WindowStyle="None"
    ResizeMode="CanResizeWithGrip"
    SizeToContent="Height"
    ShowInTaskbar="False">
  <Window.Resources>
""" + BIMTOOLS_DARK_STYLES_XML + u"""
    <Style x:Key="BimToolsCheckCategory" TargetType="CheckBox">
      <Setter Property="Foreground" Value="#E8F4F8"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Margin" Value="2,6,2,6"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CheckBox">
            <StackPanel Orientation="Horizontal">
              <Border x:Name="Box" Width="16" Height="16"
                      Background="#0a1620" BorderBrush="#21465C" BorderThickness="1"
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
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.45"/>
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
      <Border Margin="0,14,0,0" Background="#0a1620" BorderBrush="#21465C"
              BorderThickness="1" CornerRadius="6" Padding="12,10">
        <TextBlock x:Name="TxtSummary" TextWrapping="Wrap"
                   Foreground="#E8F4F8" FontSize="12" LineHeight="18"/>
      </Border>
      <TextBlock Margin="0,14,0,6" Text="Categorías de modelado en la vista"
                 Foreground="#95B8CC" FontSize="11" FontWeight="SemiBold"/>
      <Border Background="#0a1620" BorderBrush="#21465C"
              BorderThickness="1" CornerRadius="6" Padding="8,6"
              MaxHeight="320">
        <ScrollViewer VerticalScrollBarVisibility="Auto">
          <StackPanel x:Name="PanelCategories"/>
        </ScrollViewer>
      </Border>
      <StackPanel Margin="0,12,0,0" Orientation="Horizontal">
        <Button x:Name="BtnMarkAll" Content="Marcar todas"
                Style="{StaticResource BtnSelectOutline}" MinHeight="32"
                MinWidth="110" Margin="0,0,8,0"/>
        <Button x:Name="BtnMarkNone" Content="Ninguna"
                Style="{StaticResource BtnSelectOutline}" MinHeight="32"
                MinWidth="110" Margin="0,0,8,0"/>
        <Button x:Name="BtnRefresh" Content="Actualizar"
                Style="{StaticResource BtnSelectOutline}" MinHeight="32"
                MinWidth="110"/>
      </StackPanel>
      <Button x:Name="BtnSelectMarked" Margin="0,10,0,0"
              Content="Seleccionar categorías marcadas"
              Style="{StaticResource BtnPrimary}"
              HorizontalAlignment="Stretch" MinHeight="34"/>
      <TextBlock x:Name="TxtStatus" Margin="0,10,0,0"
                 Foreground="#64748b" FontSize="10" TextWrapping="Wrap"/>
      <StackPanel Margin="0,12,0,0" Orientation="Horizontal"
                  HorizontalAlignment="Right">
        <Button x:Name="BtnClose" Content="Cerrar"
                Style="{StaticResource BtnSelectOutline}" MinWidth="108"/>
      </StackPanel>
    </StackPanel>
  </Border>
</Window>
"""


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


class _SelectElementsHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        win = self._window_ref()
        if win is None:
            return
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            win._set_status(u"No hay documento activo.")
            return
        ids = win._collect_marked_element_ids()
        n = select_elements_in_model(uidoc, ids)
        if n < 1:
            win._set_status(u"Ninguna categoría marcada — selección vaciada.")
        else:
            labels = win._labels_marked()
            win._set_status(
                u"Seleccionados {0} elemento(s) de {1} categoría(s): {2}.".format(
                    n,
                    len(labels),
                    u", ".join(labels),
                ),
            )

    def GetName(self):
        return _TX_SELECT


class SuperSelectorWindow(object):
    def __init__(self, uiapp):
        self._uiapp = uiapp
        self._win = None
        self._view = None
        self._groups = []
        self._checkbox_by_label = {}
        self._updating_checks = False

        self._win = XamlReader.Parse(_build_xaml())
        self._panel = self._win.FindName(u"PanelCategories")
        self._txt_subtitle = self._win.FindName(u"TxtSubtitle")
        self._txt_summary = self._win.FindName(u"TxtSummary")
        self._txt_status = self._win.FindName(u"TxtStatus")

        self._handler_select = _SelectElementsHandler(weakref.ref(self))
        self._ext_select = ExternalEvent.Create(self._handler_select)

        self._wire_events()
        _prepare_window(self._win, uiapp)
        self._refresh_from_active_view()

    def _wire_events(self):
        self._win.FindName(u"BtnMarkAll").Click += RoutedEventHandler(
            self._on_mark_all,
        )
        self._win.FindName(u"BtnMarkNone").Click += RoutedEventHandler(
            self._on_mark_none,
        )
        self._win.FindName(u"BtnRefresh").Click += RoutedEventHandler(
            self._on_refresh,
        )
        self._win.FindName(u"BtnSelectMarked").Click += RoutedEventHandler(
            self._on_select_marked,
        )
        self._win.FindName(u"BtnClose").Click += RoutedEventHandler(
            self._on_close,
        )
        self._win.Closed += EventHandler(self._on_closed)
        self._win.PreviewKeyDown += KeyEventHandler(self._on_preview_key)

    def show(self):
        self._win.Show()

    def _on_preview_key(self, sender, args):
        if args.Key == Key.Escape:
            self._win.Close()

    def _on_closed(self, sender, args):
        _unregister_singleton()

    def _on_close(self, sender, args):
        self._win.Close()

    def _on_refresh(self, sender, args):
        self._refresh_from_active_view()

    def _on_mark_all(self, sender, args):
        self._set_mark_all(True)
        self._request_selection()

    def _on_mark_none(self, sender, args):
        self._set_mark_all(False)
        self._request_selection()

    def _on_select_marked(self, sender, args):
        self._request_selection()

    def _on_category_check_changed(self, sender, args):
        if self._updating_checks:
            return
        self._request_selection()

    def _request_selection(self):
        self._ext_select.Raise()

    def _set_status(self, text):
        if self._txt_status is not None:
            self._txt_status.Text = _as_unicode(text)

    def _refresh_from_active_view(self):
        uidoc = self._uiapp.ActiveUIDocument
        if uidoc is None:
            self._set_status(u"No hay documento activo.")
            return
        view = uidoc.ActiveView
        if isinstance(view, (ViewSheet, ViewSchedule)):
            self._set_status(
                u"Abre una vista de modelo (planta, alzado, sección, 3D…).",
            )
            return
        try:
            if getattr(view, u"IsTemplate", False):
                self._set_status(u"No se admite plantilla de vista.")
                return
        except Exception:
            pass

        doc = uidoc.Document
        self._view = view
        elements = collect_model_elements_in_view(doc, view)
        self._groups = group_elements_by_model_category(elements)
        summary = summarize_view_elements(elements, self._groups)
        view_name = _as_unicode(getattr(view, u"Name", None) or u"(vista)")
        if self._txt_subtitle is not None:
            self._txt_subtitle.Text = u"Vista activa: «{0}»".format(view_name)

        n_cat = summary.get(u"categories", 0)
        n_elem = summary.get(u"total", 0)
        if self._txt_summary is not None:
            if n_elem < 1:
                self._txt_summary.Text = (
                    u"No hay elementos de categorías de modelado en esta vista."
                )
            else:
                self._txt_summary.Text = (
                    u"Elementos de modelado en vista: {0}\n"
                    u"Categorías distintas: {1}".format(n_elem, n_cat)
                )

        self._refresh_categories()
        if n_elem < 1:
            self._set_status(u"Sin elementos que seleccionar.")
        else:
            self._set_status(
                u"Marca categorías para seleccionar sus elementos en el modelo.",
            )

    def _refresh_categories(self):
        self._checkbox_by_label = {}
        if self._panel is not None:
            self._panel.Children.Clear()

        if not self._groups:
            empty = TextBlock()
            empty.Text = u"(Sin categorías — no hay elementos de modelado en la vista)"
            empty.Foreground = self._win.FindName(u"TxtSummary").Foreground
            empty.FontSize = 12
            empty.TextWrapping = TextWrapping.Wrap
            empty.Margin = Thickness(4, 6, 4, 6)
            self._panel.Children.Add(empty)
            return

        self._updating_checks = True
        try:
            for grp in self._groups:
                label = grp[u"label"]
                count = grp[u"count"]

                cb = CheckBox()
                cb.Style = self._win.Resources[u"BimToolsCheckCategory"]
                cb.Content = u"{0}  ({1})".format(label, count)
                cb.IsChecked = False
                cb.Tag = label
                cb.ToolTip = u"Marcar para seleccionar «{0}» en el modelo".format(
                    label,
                )
                cb.Checked += RoutedEventHandler(self._on_category_check_changed)
                cb.Unchecked += RoutedEventHandler(self._on_category_check_changed)

                self._panel.Children.Add(cb)
                self._checkbox_by_label[label] = cb
        finally:
            self._updating_checks = False

    def _set_mark_all(self, checked):
        self._updating_checks = True
        try:
            for cb in self._checkbox_by_label.values():
                cb.IsChecked = bool(checked)
        finally:
            self._updating_checks = False

    def _labels_marked(self):
        out = []
        for label, cb in self._checkbox_by_label.items():
            if cb.IsChecked == True:
                out.append(label)
        out.sort()
        return out

    def _collect_marked_element_ids(self):
        labels_marked = set(self._labels_marked())
        out = []
        for grp in self._groups:
            if grp[u"label"] not in labels_marked:
                continue
            out.extend(grp[u"element_ids"])
        return out


def _unregister_singleton():
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, None)
    except Exception:
        pass


def _try_activate_existing():
    try:
        existing = AppDomain.CurrentDomain.GetData(_SINGLETON_KEY)
    except Exception:
        existing = None
    if existing is None:
        return False
    try:
        w = existing._win
        if w is None or not bool(getattr(w, u"IsLoaded", False)):
            raise RuntimeError(u"ventana cerrada")
        w.Activate()
        if w.WindowState == WindowState.Minimized:
            w.WindowState = WindowState.Normal
        TaskDialog.Show(_WINDOW_TITLE, u"La herramienta ya está en ejecución.")
        return True
    except Exception:
        _unregister_singleton()
        return False


def show_super_selector_ui(uiapp):
    if uiapp is None:
        TaskDialog.Show(_DIALOG_TITLE, u"No hay aplicación Revit activa.")
        return
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        TaskDialog.Show(_DIALOG_TITLE, u"No hay documento activo.")
        return
    if _try_activate_existing():
        return
    win = SuperSelectorWindow(uiapp)
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, win)
    except Exception:
        pass
    win.show()
