# -*- coding: utf-8 -*-
"""
Numerar columnas — interfaz WPF (módulo portable del pushbutton).
"""

import os
import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.UI import TaskDialog

import System
from System.Windows import RoutedEventHandler, WindowState, Visibility
from System.Windows.Markup import XamlReader
from System.Windows.Media import SolidColorBrush, Color

from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from numerar_columnas import (
    _TOOL_DIALOG_TITLE,
    analyze_column_stacks,
    apply_numeracion,
    format_resultado_numeracion,
)
from numerar_columnas_esquema import (
    measure_gallery_card_height,
    poblar_galeria_horizontal,
)

_APPDOMAIN_WINDOW_KEY = "BIMTools.NumerarColumnas.ActiveWindow"


def _get_active_window():
    try:
        win = System.AppDomain.CurrentDomain.GetData(_APPDOMAIN_WINDOW_KEY)
    except Exception:
        return None
    if win is None:
        return None
    try:
        _ = win.Title
    except Exception:
        _clear_active_window()
        return None
    try:
        if hasattr(win, "IsLoaded") and (not win.IsLoaded):
            _clear_active_window()
            return None
    except Exception:
        pass
    return win


def _set_active_window(win):
    try:
        System.AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, win)
    except Exception:
        pass


def _clear_active_window():
    try:
        System.AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, None)
    except Exception:
        pass


def _task_dialog_show(message, wpf_window=None):
    if wpf_window is not None:
        try:
            wpf_window.Topmost = False
        except Exception:
            pass
    try:
        TaskDialog.Show(_TOOL_DIALOG_TITLE, message)
    finally:
        if wpf_window is not None:
            try:
                wpf_window.Topmost = True
            except Exception:
                pass


def _load_xaml():
    return _NUMERAR_COL_XAML.replace(
        "__BIMTOOLS_DARK_STYLES__", BIMTOOLS_DARK_STYLES_XML
    )


_NUMERAR_COL_XAML = u"""
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="NumerarColumnasWin"
    Title="Arainco: Numerar columnas"
    Height="580" Width="980"
    MinHeight="480" MinWidth="720"
    WindowStartupLocation="CenterScreen"
    WindowState="Maximized"
    ResizeMode="CanResize"
    Background="#071018"
    FontFamily="Segoe UI"
    FontSize="12">
  <Window.Resources>
__BIMTOOLS_DARK_STYLES__
  </Window.Resources>
  <Border Background="#071018" BorderBrush="#21465C" BorderThickness="1" Padding="18">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*" MinHeight="200"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <StackPanel Grid.Row="0" Margin="0,0,0,12">
        <TextBlock Text="Arainco: Numerar columnas" Foreground="#E8F4F8"
                   FontSize="18" FontWeight="Bold"/>
        <TextBlock Foreground="#95B8CC" Margin="0,6,0,0" TextWrapping="Wrap"
                   Text="Un esquema por lote, ordenado por numeración (izquierda a derecha)."/>
      </StackPanel>

      <GroupBox Grid.Row="1" Style="{StaticResource GbParams}" Margin="0,0,0,12">
        <GroupBox.Header>
          <TextBlock Text="Criterio" FontWeight="SemiBold" Foreground="#E8F4F8"/>
        </GroupBox.Header>
        <TextBlock Foreground="#D5EAF2" TextWrapping="Wrap"
                   Text="Agrupa torres por tipos de columna y fundación. Cada configuración única es un lote; el esquema muestra un ejemplar y la cantidad de torres iguales."/>
      </GroupBox>

      <Border Grid.Row="2" CornerRadius="4" Padding="12" Margin="0,0,0,12" Background="#0b1924">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <TextBlock Grid.Row="0" Foreground="#94a3b8" FontSize="11" Margin="0,0,0,8" TextWrapping="Wrap"
                     Text="Esquemas por lote (vista previa). La numeración en el modelo se aplica solo con «Numerar columnas»."/>
          <Border Grid.Row="1" Background="#050a10" CornerRadius="4" Padding="10">
            <ScrollViewer x:Name="GalleryScroll" VerticalScrollBarVisibility="Disabled"
                          HorizontalScrollBarVisibility="Auto" VerticalAlignment="Stretch">
              <StackPanel x:Name="GalleryHost" Orientation="Horizontal"
                          VerticalAlignment="Stretch"/>
            </ScrollViewer>
          </Border>
        </Grid>
      </Border>

      <TextBlock x:Name="TxtResumen" Grid.Row="3" Foreground="#95B8CC" FontSize="11"
                 TextWrapping="Wrap" Margin="0,0,0,12"
                 Text="1) Analizar proyecto — vista previa.  2) Numerar columnas — escribe en el modelo."/>

      <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right">
        <Button x:Name="BtnCerrar" Content="Cerrar" Style="{StaticResource BtnSelectOutline}"
                MinWidth="100" Margin="0,0,10,0"/>
        <Button x:Name="BtnAnalizar" Content="Analizar proyecto"
                Style="{StaticResource BtnSelectOutline}" MinWidth="140" Margin="0,0,10,0"/>
        <Button x:Name="BtnNumerar" Content="Numerar columnas"
                Style="{StaticResource BtnPrimary}" MinWidth="150"/>
      </StackPanel>
    </Grid>
  </Border>
</Window>
"""


def _brush_ok():
    try:
        return SolidColorBrush(Color.FromRgb(213, 234, 242))
    except Exception:
        return None


def _brush_warn():
    try:
        return SolidColorBrush(Color.FromRgb(255, 200, 120))
    except Exception:
        return None


def _format_resumen(analysis):
    if not analysis or not analysis.get(u"ok"):
        return analysis.get(u"message") if analysis else u"Sin datos."
    n_l = analysis.get(u"n_configuraciones") or 0
    n_t = analysis.get(u"n_torres") or 0
    return (
        u"{0} torres en el proyecto \u00b7 {1} lotes (configuraciones). "
        u"Esquemas ordenados por número de lote."
    ).format(n_t, n_l)


def run(revit):
    existing = _get_active_window()
    if existing is not None:
        try:
            existing.WindowState = WindowState.Maximized
        except Exception:
            pass
        try:
            existing.Show()
            existing.Activate()
            existing.Focus()
        except Exception:
            pass
        _task_dialog_show(u"La herramienta ya esta en ejecucion.", existing)
        return

    uidoc = revit.ActiveUIDocument
    if uidoc is None:
        TaskDialog.Show(_TOOL_DIALOG_TITLE, u"No hay documento activo.")
        return
    doc = uidoc.Document

    win = XamlReader.Parse(_load_xaml())
    try:
        win.WindowState = WindowState.Maximized
    except Exception:
        pass
    state = {
        u"analysis": None,
        u"numerar_busy": False,
        u"numeracion_aplicada": False,
    }

    gallery_host = [None]
    gallery_scroll = [None]
    txt_resumen = win.FindName("TxtResumen")
    _last_gallery_card_h = [None]
    _idling_numerar_handler = [None]
    uiapp = revit

    def _resolve_gallery_hosts():
        if gallery_host[0] is None:
            try:
                gallery_host[0] = win.FindName("GalleryHost")
            except Exception:
                pass
        if gallery_scroll[0] is None:
            try:
                gallery_scroll[0] = win.FindName("GalleryScroll")
            except Exception:
                pass

    def _gallery_card_height():
        return measure_gallery_card_height(gallery_scroll[0], fallback=360.0)

    def _refresh_gallery():
        _resolve_gallery_hosts()
        host = gallery_host[0]
        card_h = _gallery_card_height()
        _last_gallery_card_h[0] = card_h
        analysis = state.get(u"analysis")
        if not analysis or not analysis.get(u"ok"):
            try:
                poblar_galeria_horizontal(host, [], card_height_px=card_h)
            except Exception as ex:
                _task_dialog_show(u"Galería: {0}".format(ex), win)
            if txt_resumen is not None:
                txt_resumen.Text = (
                    analysis.get(u"message") if analysis else u"Sin análisis."
                )
                txt_resumen.Foreground = SolidColorBrush(Color.FromRgb(255, 120, 120))
            return
        try:
            poblar_galeria_horizontal(
                host,
                analysis.get(u"lotes") or [],
                card_height_px=card_h,
            )
        except Exception as ex:
            _task_dialog_show(u"Error al dibujar esquemas:\n{0}".format(ex), win)
            return
        if txt_resumen is not None:
            txt_resumen.Text = _format_resumen(analysis)
            br = _brush_ok()
            if br is not None:
                txt_resumen.Foreground = br
        try:
            if host is not None:
                host.UpdateLayout()
            if gallery_scroll[0] is not None:
                gallery_scroll[0].UpdateLayout()
            win.UpdateLayout()
        except Exception:
            pass

    def _refit_gallery_on_resize(sender, args):
        _resolve_gallery_hosts()
        analysis = state.get(u"analysis")
        if not analysis or not analysis.get(u"ok"):
            return
        card_h = _gallery_card_height()
        if _last_gallery_card_h[0] is not None:
            if abs(float(card_h) - float(_last_gallery_card_h[0])) < 8.0:
                return
        _last_gallery_card_h[0] = card_h
        try:
            poblar_galeria_horizontal(
                gallery_host[0],
                analysis.get(u"lotes") or [],
                card_height_px=card_h,
            )
        except Exception:
            pass

    def _do_analyze():
        try:
            analysis = analyze_column_stacks(doc)
        except Exception as ex:
            _task_dialog_show(u"Error al analizar:\n{0}".format(ex), win)
            return
        state[u"analysis"] = analysis
        state[u"numeracion_aplicada"] = False
        _refresh_gallery()
        if not analysis.get(u"ok"):
            _task_dialog_show(analysis.get(u"message") or u"Error.", win)
        elif txt_resumen is not None:
            txt_resumen.Text = (
                _format_resumen(analysis)
                + u" Pulse «Numerar columnas» para escribir en el modelo."
            )
        if btn_num is not None and analysis.get(u"ok"):
            try:
                btn_num.IsEnabled = True
            except Exception:
                pass

    def _on_analizar(sender, args):
        _do_analyze()

    def _restore_window_after_revit_op():
        try:
            win.Show()
            win.WindowState = WindowState.Maximized
            win.Activate()
        except Exception:
            pass

    def _complete_numerar(res, err):
        state[u"numerar_busy"] = False
        if err:
            _restore_window_after_revit_op()
            if btn_num is not None:
                try:
                    btn_num.IsEnabled = True
                except Exception:
                    pass
            if txt_resumen is not None:
                txt_resumen.Text = u"Error al numerar."
                try:
                    txt_resumen.Foreground = SolidColorBrush(
                        Color.FromRgb(255, 120, 120)
                    )
                except Exception:
                    pass
            _task_dialog_show(u"Error al numerar:\n{0}".format(err), win)
            return

        state[u"numeracion_aplicada"] = True
        try:
            win.Close()
        except Exception:
            try:
                _restore_window_after_revit_op()
            except Exception:
                pass

    def _unregister_idling_numerar():
        h = _idling_numerar_handler[0]
        if h is None:
            return
        try:
            uiapp.Idling -= h
        except Exception:
            pass
        _idling_numerar_handler[0] = None

    def _on_numerar(sender, args):
        """Único punto que escribe «Numeracion Columna» en el modelo (transacción Revit)."""
        if state.get(u"numerar_busy"):
            return
        analysis = state.get(u"analysis")
        if not analysis or not analysis.get(u"ok"):
            _task_dialog_show(u"Analice el proyecto antes de numerar.", win)
            return

        ordered_groups = analysis[u"ordered_groups"]
        state[u"numerar_busy"] = True
        if btn_num is not None:
            try:
                btn_num.IsEnabled = False
            except Exception:
                pass
        if txt_resumen is not None:
            txt_resumen.Text = u"Numerando columnas en el modelo\u2026"
            try:
                br = _brush_ok()
                if br is not None:
                    txt_resumen.Foreground = br
            except Exception:
                pass

        _unregister_idling_numerar()

        def _on_idling_numerar(sender_idling, args_idling):
            _unregister_idling_numerar()
            res = None
            err = None
            try:
                res = apply_numeracion(doc, ordered_groups)
            except Exception as ex:
                err = unicode(ex)

            def _finish_on_ui():
                _complete_numerar(res, err)

            try:
                from System import Action
                from System.Windows.Threading import DispatcherPriority

                win.Dispatcher.BeginInvoke(
                    DispatcherPriority.Normal,
                    Action(_finish_on_ui),
                )
            except Exception:
                _complete_numerar(res, err)

        _idling_numerar_handler[0] = _on_idling_numerar
        try:
            win.Hide()
        except Exception:
            pass
        try:
            uiapp.Idling += _on_idling_numerar
        except Exception as ex:
            state[u"numerar_busy"] = False
            _complete_numerar(None, unicode(ex))

    def _on_close(sender, args):
        _unregister_idling_numerar()
        _clear_active_window()

    win.Closed += _on_close

    btn_cerrar = win.FindName("BtnCerrar")
    if btn_cerrar is not None:
        btn_cerrar.Click += RoutedEventHandler(lambda s, e: win.Close())

    btn_an = win.FindName("BtnAnalizar")
    if btn_an is not None:
        btn_an.Click += RoutedEventHandler(_on_analizar)

    btn_num = win.FindName("BtnNumerar")
    if btn_num is not None:
        try:
            btn_num.IsEnabled = False
        except Exception:
            pass
        btn_num.Click += RoutedEventHandler(_on_numerar)

    def _prepare_empty_gallery():
        _resolve_gallery_hosts()
        card_h = _gallery_card_height()
        try:
            poblar_galeria_horizontal(gallery_host[0], [], card_height_px=card_h)
        except Exception:
            pass

    def _on_loaded(sender, args):
        try:
            win.WindowState = WindowState.Maximized
        except Exception:
            pass
        _resolve_gallery_hosts()
        if gallery_scroll[0] is not None:
            try:
                gallery_scroll[0].SizeChanged += _refit_gallery_on_resize
            except Exception:
                pass

    def _on_content_rendered(sender, args):
        try:
            win.ContentRendered -= _on_content_rendered
        except Exception:
            pass
        try:
            from System import Action
            from System.Windows.Threading import DispatcherPriority

            def _idle():
                _prepare_empty_gallery()

            win.Dispatcher.BeginInvoke(
                DispatcherPriority.ApplicationIdle,
                Action(_idle),
            )
        except Exception:
            _prepare_empty_gallery()

    def _on_win_size_changed(sender, args):
        _refit_gallery_on_resize(None, None)

    try:
        win.SizeChanged += _on_win_size_changed
    except Exception:
        pass

    win.Loaded += _on_loaded
    win.ContentRendered += _on_content_rendered

    _set_active_window(win)
    try:
        win.ShowDialog()
    finally:
        _clear_active_window()
