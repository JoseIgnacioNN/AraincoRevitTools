# -*- coding: utf-8 -*-
"""
Numerar fundaciones — interfaz WPF (misma línea visual que armadura BIMTools).

Tema: ``bimtools_wpf_dark_theme.BIMTOOLS_DARK_STYLES_XML`` (Fundación Aislada / Malla en Losa).
"""

import os
import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    Transaction,
    FilteredElementCollector,
    FamilyInstance,
    WallFoundation,
    Floor,
)
from Autodesk.Revit.UI import TaskDialog

import System
from System.Windows import WindowState, Visibility, SizeToContent, Size
from System.Windows.Markup import XamlReader
from System.Windows.Media import SolidColorBrush, Color

from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from revit_wpf_window_position import position_wpf_window_top_left_at_active_view, revit_main_hwnd

from bimtools_paths import get_logo_paths

_APPDOMAIN_WINDOW_KEY = "BIMTools.NumerarFundaciones.ActiveWindow"
_TOOL_DIALOG_TITLE = u"BIMTools — Numerar Fundaciones"


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


# ── Lógica de numeración (antes en script del pushbutton) ────────────────────


def _get_dimensiones_y_volumen(element, doc):
    try:
        d = doc
        elem_type = d.GetElement(element.GetTypeId())
        if not elem_type:
            return ((0, 0, 0), 0)
        length_val, width_val, height_val = None, None, 1.0
        for name in ("Length", "Largo", "Longitud"):
            p = elem_type.LookupParameter(name)
            if p and p.HasValue:
                length_val = p.AsDouble() * 304.8
                break
        for name in ("Width", "Ancho"):
            p = elem_type.LookupParameter(name)
            if p and p.HasValue:
                width_val = p.AsDouble() * 304.8
                break
        for name in ("Height", "Depth", "Thickness", "Altura", "Profundidad", "Espesor"):
            p = elem_type.LookupParameter(name)
            if p and p.HasValue:
                height_val = p.AsDouble() * 304.8
                break
        if length_val is not None and width_val is not None:
            vol = length_val * width_val * height_val
            return ((length_val, width_val, height_val), vol)
    except Exception:
        pass
    return ((0, 0, 0), 0)


def _set_numeracion_fundacion(element, value):
    try:
        for param_name in ("Numeracion Fundacion", "Numeracion fundacion", "Numeracion", "Foundation Numbering"):
            p = element.LookupParameter(param_name)
            if p is not None and not p.IsReadOnly:
                p.Set(str(value))
                return True
    except Exception:
        pass
    return False


def _leer_numeracion_fundacion(element):
    if not element:
        return None
    try:
        for param_name in ("Numeracion Fundacion", "Numeracion fundacion", "Numeracion", "Foundation Numbering"):
            p = element.LookupParameter(param_name)
            if p is not None and p.HasValue:
                s = p.AsString()
                vs = p.AsValueString()
                val = s if s is not None else vs
                if val is not None and str(val).strip() and str(val).strip() != "0":
                    return str(val).strip()
                try:
                    d = p.AsDouble()
                    if d != 0.0:
                        return str(int(d)) if d == int(d) else str(d)
                except Exception:
                    pass
    except Exception:
        pass
    return None


def execute_numerar_fundaciones(doc):
    """
    Ejecuta la numeración en ``doc``.
    Retorna dict: ok (bool), text (unicode), is_error (bool).
    """
    from collections import defaultdict

    collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_StructuralFoundation
    ).WhereElementIsNotElementType()

    fundaciones = []
    total_cat = 0
    wall_found = 0
    floor_found = 0
    family_instance = 0
    otros = 0

    for elem in collector:
        total_cat += 1
        if isinstance(elem, WallFoundation):
            wall_found += 1
            continue
        if isinstance(elem, Floor):
            floor_found += 1
            continue
        if isinstance(elem, FamilyInstance):
            family_instance += 1
            fundaciones.append(elem)
        else:
            otros += 1

    if not fundaciones:
        msg = (
            u"No hay fundaciones aisladas.\n\n"
            u"Wall foundations y slab foundations se excluyen.\n\n"
            u"--- Diagnóstico ---\n"
            u"Total en categoría Structural Foundation: {}\n"
            u"  - Wall foundations: {}\n"
            u"  - Slab foundations (Floor): {}\n"
            u"  - FamilyInstance (aisladas): {}\n"
            u"  - Otros tipos: {}"
        ).format(total_cat, wall_found, floor_found, family_instance, otros)
        return {"ok": True, "text": msg, "is_error": False}

    grupos = defaultdict(list)
    for elem in fundaciones:
        dims, _ = _get_dimensiones_y_volumen(elem, doc)
        grupos[dims].append(elem)

    numeros_existentes = set()
    for elem in fundaciones:
        num = _leer_numeracion_fundacion(elem)
        if num is not None:
            try:
                numeros_existentes.add(int(num))
            except (ValueError, TypeError):
                pass

    grupos_sin_numero = [
        (dims, elems) for dims, elems in grupos.items()
        if not any(_leer_numeracion_fundacion(e) for e in elems)
    ]
    grupos_sin_numero.sort(
        key=lambda x: (
            x[0][0] * x[0][1] * x[0][2],
            x[0][0], x[0][1], x[0][2]
        )
    )

    siguiente = (max(numeros_existentes) if numeros_existentes else 0) + 1
    numero_por_grupo = {}
    for dims, elems in grupos_sin_numero:
        numero_por_grupo[dims] = str(siguiente)
        siguiente += 1

    for dims, elems in grupos.items():
        if dims not in numero_por_grupo:
            for e in elems:
                num = _leer_numeracion_fundacion(e)
                if num is not None:
                    numero_por_grupo[dims] = num
                    break

    trans = Transaction(doc, "Numerar fundaciones aisladas")
    try:
        trans.Start()
        numerados = 0
        sin_parametro = 0
        for dims, elementos in grupos.items():
            numero = numero_por_grupo.get(dims)
            if numero is None:
                continue
            for elem in elementos:
                if _leer_numeracion_fundacion(elem) is not None:
                    continue
                if _set_numeracion_fundacion(elem, numero):
                    numerados += 1
                else:
                    sin_parametro += 1

        trans.Commit()

        lineas_reporte = []
        for dims, elementos in grupos.items():
            numero = numero_por_grupo.get(dims)
            if numero is None:
                continue
            l_mm, w_mm, h_mm = dims
            dims_str = u"{:.2f} x {:.2f} x {:.2f} m".format(
                l_mm / 1000.0, w_mm / 1000.0, h_mm / 1000.0
            )
            cant = len(elementos)
            cant_str = u"{} unidad".format(cant) if cant == 1 else u"{} unidades".format(cant)
            lineas_reporte.append((numero, dims_str, cant_str))

        def _orden_numero(item):
            try:
                return int(item[0])
            except (ValueError, TypeError):
                return 9999
        lineas_reporte.sort(key=_orden_numero)

        if numerados == 0 and sin_parametro == 0:
            msg = u"No hay fundaciones sin numerar. Todas tienen valor en 'Numeracion Fundacion'."
        else:
            msg = u"Se numeraron {} fundación(es) sin numerar.".format(numerados)
            if numerados > 0:
                msg += u"\n\nSe agruparon con las ya numeradas que comparten las mismas dimensiones."
        if sin_parametro:
            msg += u"\n\n{} elemento(s) no tienen parámetro 'Numeracion Fundacion' editable.".format(sin_parametro)

        if lineas_reporte:
            msg += u"\n\n--- Listado de fundaciones numeradas ---\n"
            for numero, dims_str, cant_str in lineas_reporte:
                msg += u"\nF{}: {} - {}".format(numero, dims_str, cant_str)

        return {"ok": True, "text": msg, "is_error": False}

    except Exception as ex:
        if trans.HasStarted():
            trans.RollBack()
        return {"ok": False, "text": u"Error:\n{0}".format(ex), "is_error": True}


# ── XAML (tema oscuro compartido) ─────────────────────────────────────────────

_NUMERAR_FUND_XAML = u"""
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Arainco - Numerar Fundaciones"
    MaxHeight="720"
    WindowStartupLocation="Manual"
    Background="Transparent"
    AllowsTransparency="True"
    MinHeight="0"
    FontFamily="Segoe UI"
    WindowStyle="None"
    ResizeMode="NoResize"
    Topmost="True"
    UseLayoutRounding="True"
    SizeToContent="Height"
    Width="440"
    >
  <Window.Resources>
""" + BIMTOOLS_DARK_STYLES_XML + u"""
    <Style x:Key="TbInforme" TargetType="TextBox" BasedOn="{StaticResource CantSpinnerText}">
      <Setter Property="HorizontalAlignment" Value="Stretch"/>
      <Setter Property="FontWeight" Value="Normal"/>
      <Setter Property="Padding" Value="8,8"/>
      <Setter Property="Background" Value="#050E18"/>
      <Setter Property="Foreground" Value="#C8E4EF"/>
    </Style>
  </Window.Resources>
  <!-- Misma idea que Fundación Aislada: solo filas Auto (nada de *), SizeToContent Height en el Window. -->
  <Border x:Name="NumerarRootChrome" CornerRadius="10" Background="#0A1A2F" Padding="12"
          BorderBrush="#1A3A4D" BorderThickness="1"
          HorizontalAlignment="Stretch" VerticalAlignment="Top" ClipToBounds="True">
    <Border.Effect>
      <DropShadowEffect Color="#000000" BlurRadius="16" ShadowDepth="0" Opacity="0.35"/>
    </Border.Effect>
    <Grid HorizontalAlignment="Stretch">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <Border x:Name="TitleBar" Grid.Row="0" Background="#0E1B32" CornerRadius="6" Padding="10,8" Margin="0,0,0,10"
              BorderBrush="#21465C" BorderThickness="1" HorizontalAlignment="Stretch">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <Image x:Name="ImgLogo" Width="40" Height="40" Grid.Column="0"
                 Stretch="Uniform" Margin="0,0,10,0" VerticalAlignment="Center"/>
          <StackPanel Grid.Column="1" VerticalAlignment="Center">
            <TextBlock Text="Numerar fundaciones" FontSize="15" FontWeight="SemiBold"
                       Foreground="#E8F4F8"/>
            <TextBlock Text="Agrupa fundaciones aisladas por dimensiones y escribe Numeracion Fundacion"
                       FontSize="11" Foreground="#95B8CC" Margin="0,6,0,0" TextWrapping="Wrap"/>
          </StackPanel>
          <Button x:Name="BtnClose" Grid.Column="2" Style="{StaticResource BtnCloseX_MinimalNoBg}"
                  VerticalAlignment="Center" ToolTip="Cerrar"/>
        </Grid>
      </Border>

      <GroupBox Grid.Row="1" Style="{StaticResource GbParams}" Margin="0,0,0,10" HorizontalAlignment="Stretch">
        <GroupBox.Header>
          <TextBlock Text="Criterio" FontWeight="SemiBold" Foreground="#E8F4F8" FontSize="11"/>
        </GroupBox.Header>
        <StackPanel>
          <TextBlock Style="{StaticResource LabelSmall}" TextWrapping="Wrap" Margin="0,0,0,4"
                     Text="Solo fundaciones aisladas (FamilyInstance). Se excluyen wall foundations y slab foundations (Floor)."/>
          <TextBlock Style="{StaticResource LabelSmall}" TextWrapping="Wrap"
                     Text="Se agrupan por dimensiones de tipo (L × W × H). Solo se rellena el parámetro «Numeracion Fundacion» cuando está vacío; no se sobrescribe lo existente."/>
        </StackPanel>
      </GroupBox>

      <Grid x:Name="GridResultadoHost" Grid.Row="2" Margin="0,0,0,10" Visibility="Collapsed">
        <GroupBox x:Name="GbResultado" Style="{StaticResource GbParams}" Margin="0"
                  HorizontalAlignment="Stretch">
          <GroupBox.Header>
            <TextBlock Text="Resultado" FontWeight="SemiBold" Foreground="#E8F4F8" FontSize="11"/>
          </GroupBox.Header>
          <Border Background="#050E18" BorderBrush="#1A3A4D" BorderThickness="1" CornerRadius="4" Padding="0">
            <TextBox x:Name="TxtResult" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True"
                     MinHeight="0" MaxHeight="320" VerticalScrollBarVisibility="Auto"
                     Style="{StaticResource TbInforme}" Text=""/>
          </Border>
        </GroupBox>
      </Grid>

      <Button x:Name="BtnNumerar" Grid.Row="3" Content="Numerar fundaciones"
              Style="{StaticResource BtnPrimary}"
              HorizontalAlignment="Stretch" Margin="0"/>
    </Grid>
  </Border>
</Window>
"""


def _try_load_logo(image_control):
    if image_control is None:
        return
    from System.IO import FileAccess, FileMode, FileStream
    from System.Windows.Media.Imaging import BitmapCacheOption, BitmapImage

    for path in get_logo_paths():
        if not path or not os.path.isfile(path):
            continue
        stream = None
        try:
            stream = FileStream(path, FileMode.Open, FileAccess.Read)
            bmp = BitmapImage()
            bmp.BeginInit()
            bmp.StreamSource = stream
            bmp.CacheOption = BitmapCacheOption.OnLoad
            bmp.EndInit()
            bmp.Freeze()
            image_control.Source = bmp
            return
        except Exception:
            continue
        finally:
            if stream is not None:
                try:
                    stream.Dispose()
                except Exception:
                    pass


def _brush_error():
    try:
        return SolidColorBrush(Color.FromRgb(255, 120, 120))
    except Exception:
        return None


def _brush_ok():
    try:
        return SolidColorBrush(Color.FromRgb(200, 228, 239))
    except Exception:
        return None


def _snap_window_height_to_chrome(win):
    """
    Iguala el alto del Window al chrome sin hueco bajo el botón.

    ActualHeight del Border puede ser mayor que el contenido si el padre le
    asigna altura extra; Measure(Width, Infinity) + DesiredSize.Height evita eso.
    """
    try:
        chrome = win.FindName("NumerarRootChrome")
        if chrome is None:
            return
        try:
            win.UpdateLayout()
        except Exception:
            pass
        try:
            aw = float(win.Width)
            if aw < 1.0:
                aw = float(chrome.ActualWidth)
            if aw < 1.0:
                aw = 440.0
        except Exception:
            aw = 440.0
        inf = System.Double.PositiveInfinity
        try:
            chrome.InvalidateMeasure()
        except Exception:
            pass
        try:
            chrome.Measure(Size(aw, inf))
        except Exception:
            pass
        try:
            ch_desired = float(chrome.DesiredSize.Height)
            ch_actual = float(chrome.ActualHeight)
        except Exception:
            return
        ch = ch_desired if ch_desired > 1.0 else ch_actual
        if ch < 1.0:
            return
        win.SizeToContent = SizeToContent.Manual
        try:
            win.MinHeight = 0
            win.MaxHeight = 720
        except Exception:
            pass
        try:
            win.Height = ch
        except Exception:
            pass
    except Exception:
        pass


def run(revit):
    """Punto de entrada pyRevit: muestra la ventana."""
    existing = _get_active_window()
    if existing is not None:
        try:
            if existing.WindowState == WindowState.Minimized:
                existing.WindowState = WindowState.Normal
        except Exception:
            pass
        try:
            existing.Show()
        except Exception:
            pass
        try:
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

    win = XamlReader.Parse(_NUMERAR_FUND_XAML)

    _host_res = win.FindName("GridResultadoHost")
    if _host_res is not None:
        try:
            _host_res.Visibility = Visibility.Collapsed
        except Exception:
            pass

    def _on_close(sender, args):
        _clear_active_window()

    win.Closed += _on_close

    btn_close = win.FindName("BtnClose")
    if btn_close is not None:
        btn_close.Click += lambda s, e: win.Close()

    title_bar = win.FindName("TitleBar")
    if title_bar is not None:
        def _drag(s, e):
            try:
                win.DragMove()
            except Exception:
                pass
        title_bar.MouseLeftButtonDown += _drag

    txt_result = win.FindName("TxtResult")
    grid_resultado_host = win.FindName("GridResultadoHost")
    br_err = _brush_error()
    br_ok = _brush_ok()

    def _on_numerar(sender, args):
        res = execute_numerar_fundaciones(doc)
        if txt_result is not None:
            try:
                txt_result.Text = res.get("text") or u""
            except Exception:
                pass
            try:
                if res.get("is_error") and br_err is not None:
                    txt_result.Foreground = br_err
                elif br_ok is not None:
                    txt_result.Foreground = br_ok
            except Exception:
                pass
        if txt_result is not None:
            try:
                txt_result.MinHeight = 180
            except Exception:
                pass
        if grid_resultado_host is not None:
            try:
                grid_resultado_host.Visibility = Visibility.Visible
            except Exception:
                pass
        try:
            win.UpdateLayout()
        except Exception:
            pass
        _snap_window_height_to_chrome(win)

    btn_num = win.FindName("BtnNumerar")
    if btn_num is not None:
        btn_num.Click += _on_numerar

    img = win.FindName("ImgLogo")
    _try_load_logo(img)

    def _on_loaded(sender, args):
        try:
            hwnd = revit_main_hwnd(revit)
            position_wpf_window_top_left_at_active_view(win, uidoc, hwnd)
        except Exception:
            pass
        h = win.FindName("GridResultadoHost")
        if h is not None:
            try:
                h.Visibility = Visibility.Collapsed
            except Exception:
                pass

    win.Loaded += _on_loaded

    def _on_content_rendered(sender, args):
        try:
            win.ContentRendered -= _on_content_rendered
        except Exception:
            pass

        def _snap_idle():
            _snap_window_height_to_chrome(win)

        try:
            from System import Action
            from System.Windows.Threading import DispatcherPriority

            win.Dispatcher.BeginInvoke(
                DispatcherPriority.ApplicationIdle,
                Action(_snap_idle),
            )
        except Exception:
            _snap_idle()

    win.ContentRendered += _on_content_rendered

    _set_active_window(win)
    try:
        win.ShowDialog()
    finally:
        _clear_active_window()
