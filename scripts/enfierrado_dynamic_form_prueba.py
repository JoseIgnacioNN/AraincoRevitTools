# -*- coding: utf-8 -*-
"""
Formulario dinámico de prueba para enfierrado (WPF).

Objetivo:
- Mismo "form" base (Window WPF).
- Con `ComboBox` se cambia entre Tipo 1 y Tipo 2.
- Cada tipo llama a un módulo Python distinto.

Se utiliza para pruebas de integración UI dentro de Revit.
"""

import os
import sys
import clr
import weakref

clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")

from System.Collections.Generic import List  # noqa: F401 (compatibilidad futura)
from System.Windows.Markup import XamlReader
from System.Windows import Window, RoutedEventHandler
from System.Windows.Input import Key, KeyBinding, ModifierKeys, ApplicationCommands, CommandBinding
from System.Windows.Controls import SelectionChangedEventHandler
from System.Windows import Visibility
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption  # noqa: F401 (por si añadimos logo)

from Autodesk.Revit.UI import TaskDialog, ExternalEvent, IExternalEventHandler  # noqa: F401


# --- Constantes: extensiones / rutas ---
_script_dir = os.path.dirname(os.path.abspath(__file__))
_ext_root = os.path.dirname(_script_dir)
_scripts_dir = _script_dir
if _scripts_dir not in sys.path:
    sys.path.insert(0, _scripts_dir)

from revit_wpf_window_position import (
    position_wpf_window_top_left_at_active_view,
    revit_main_hwnd,
)
from System.Windows.Interop import WindowInteropHelper


XAML = r"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Enfierrado - Formulario Prueba"
        WindowStartupLocation="Manual"
        Width="520" Height="360"
        Background="#0A1C26" Foreground="#C8E4EF">
  <Window.Resources>
    <Style x:Key="BIMComboBoxStyle" TargetType="ComboBox">
      <Setter Property="Foreground" Value="#C8E4EF"/>
      <Setter Property="Background" Value="#1A3D52"/>
    </Style>
    <Style x:Key="BIMComboBoxItemStyle" TargetType="ComboBoxItem">
      <Setter Property="Foreground" Value="#C8E4EF"/>
      <Setter Property="Background" Value="#0F2535"/>
      <Setter Property="Padding" Value="8,5"/>
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="#1A3D52"/>
          <Setter Property="Foreground" Value="#C8E4EF"/>
        </Trigger>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#143246"/>
          <Setter Property="Foreground" Value="#C8E4EF"/>
        </Trigger>
      </Style.Triggers>
    </Style>
  </Window.Resources>
  <Grid Margin="16">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <TextBlock Grid.Row="0"
               Text="Selecciona el tipo"
               FontSize="12"
               FontWeight="SemiBold"
               Margin="0,0,0,8"/>

    <ComboBox Grid.Row="1" x:Name="CmbTipo"
              Style="{StaticResource BIMComboBoxStyle}"
              ItemContainerStyle="{StaticResource BIMComboBoxItemStyle}"
              Height="30" Margin="0,0,0,12"
              Padding="8,4"/>

    <Border Grid.Row="2"
            Background="#0F2535" BorderBrush="#1A3D52" BorderThickness="1"
            CornerRadius="6" Padding="12,10">
      <Grid>
        <StackPanel x:Name="PanelTipo1" Visibility="Collapsed">
          <Button x:Name="BtnHolaMario"
                  Content="Hola Mario"
                  Width="240" Height="34"
                  Background="#5BB8D4" Foreground="#0A1C26"
                  Margin="0,10,0,0"
                  HorizontalAlignment="Left"/>
        </StackPanel>

        <StackPanel x:Name="PanelTipo2" Visibility="Visible">
          <Button x:Name="BtnHolaJose"
                  Content="Hola Jose"
                  Width="240" Height="34"
                  Background="#5BB8D4" Foreground="#0A1C26"
                  Margin="0,10,0,0"
                  HorizontalAlignment="Left"/>
        </StackPanel>

        <StackPanel x:Name="PanelTipo3" Visibility="Collapsed">
          <TextBlock Text="Tipo 3" FontWeight="SemiBold" Margin="0,0,0,10"/>

          <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
            <TextBlock Text="Nombre:" Width="90" VerticalAlignment="Center"/>
            <TextBox x:Name="TxtTipo3Nombre" Width="240" Height="26" Background="#09161F" Foreground="#C8E4EF"/>
          </StackPanel>

          <StackPanel Orientation="Horizontal">
            <TextBlock Text="Apellido:" Width="90" VerticalAlignment="Center"/>
            <TextBox x:Name="TxtTipo3Apellido" Width="240" Height="26" Background="#09161F" Foreground="#C8E4EF"/>
          </StackPanel>

          <Button x:Name="BtnGenerarNombre"
                  Content="Generar Nombre"
                  Width="240" Height="34"
                  Background="#5BB8D4" Foreground="#0A1C26"
                  Margin="0,14,0,0"
                  HorizontalAlignment="Left"/>
        </StackPanel>

      </Grid>
    </Border>

    <TextBlock Grid.Row="3" x:Name="TxtEstado"
               Margin="0,10,0,10"
               Foreground="#5BB8D4"
               FontSize="11"
               TextWrapping="Wrap"/>

  </Grid>
</Window>
"""


class EnfierradoDynamicFormWindow(object):
    def __init__(self, revit, close_on_finish=False):
        self._revit = revit
        self._close_on_finish = bool(close_on_finish)

        self._win = XamlReader.Parse(XAML)

        self._cmb_tipo = self._win.FindName("CmbTipo")
        self._panel_tipo1 = self._win.FindName("PanelTipo1")
        self._panel_tipo2 = self._win.FindName("PanelTipo2")
        self._panel_tipo3 = self._win.FindName("PanelTipo3")
        self._txt_estado = self._win.FindName("TxtEstado")

        # Botones por tipo
        self._btn_hola_mario = self._win.FindName("BtnHolaMario")
        self._btn_hola_jose = self._win.FindName("BtnHolaJose")
        self._btn_generar_nombre = self._win.FindName("BtnGenerarNombre")

        # Tipo 3
        self._t3_nombre = self._win.FindName("TxtTipo3Nombre")
        self._t3_apellido = self._win.FindName("TxtTipo3Apellido")

        self._setup_ui()

    @staticmethod
    def _get_combo_text(combo):
        try:
            item = combo.SelectedItem
            if item is None:
                return ""
            return str(getattr(item, "Content", item))
        except Exception:
            return ""

    def _setup_ui(self):
        if self._cmb_tipo:
            self._cmb_tipo.Items.Add("Barra tipo 1")
            self._cmb_tipo.Items.Add("Barra tipo 2")
            self._cmb_tipo.Items.Add("Barra tipo 3")
            self._cmb_tipo.SelectedIndex = 1  # default: tipo 2
            self._cmb_tipo.SelectionChanged += SelectionChangedEventHandler(self._on_tipo_changed)

        if self._btn_hola_mario:
            self._btn_hola_mario.Click += RoutedEventHandler(self._on_hola_mario)
        if self._btn_hola_jose:
            self._btn_hola_jose.Click += RoutedEventHandler(self._on_hola_jose)
        if self._btn_generar_nombre:
            self._btn_generar_nombre.Click += RoutedEventHandler(self._on_generar_nombre)

        # default state
        self._on_tipo_changed(None, None)

        # Cerrar con ESC / comando cerrar
        def _on_close_cmd(sender, e):
            try:
                self._win.Close()
            except Exception:
                pass

        self._win.CommandBindings.Add(CommandBinding(ApplicationCommands.Close, _on_close_cmd))
        self._win.InputBindings.Add(KeyBinding(ApplicationCommands.Close, Key.Escape, ModifierKeys.None))

    def _set_estado(self, msg):
        try:
            if self._txt_estado:
                self._txt_estado.Text = msg
        except Exception:
            pass

    def _on_tipo_changed(self, sender, args):
        tipo = self._get_combo_text(self._cmb_tipo)

        try:
            if tipo == "Barra tipo 1":
                self._panel_tipo1.Visibility = Visibility.Visible
                self._panel_tipo2.Visibility = Visibility.Collapsed
                self._panel_tipo3.Visibility = Visibility.Collapsed
                self._set_estado("Tipo 1 seleccionado.")
            elif tipo == "Barra tipo 3":
                self._panel_tipo1.Visibility = Visibility.Collapsed
                self._panel_tipo2.Visibility = Visibility.Collapsed
                self._panel_tipo3.Visibility = Visibility.Visible
                self._set_estado("Tipo 3 seleccionado.")
            else:
                self._panel_tipo1.Visibility = Visibility.Collapsed
                self._panel_tipo2.Visibility = Visibility.Visible
                self._panel_tipo3.Visibility = Visibility.Collapsed
                self._set_estado("Tipo 2 seleccionado.")
        except Exception:
            pass

    def _on_hola_mario(self, sender, args):
        try:
            from enfierrado_hola_mario import run as run_mario

            run_mario(self._revit)
        except Exception as ex:
            try:
                TaskDialog.Show("Enfierrado - Error", u"Error en Hola Mario:\n{}".format(str(ex)))
            except Exception:
                pass

    def _on_hola_jose(self, sender, args):
        try:
            from enfierrado_hola_jose import run as run_jose

            run_jose(self._revit)
        except Exception as ex:
            try:
                TaskDialog.Show("Enfierrado - Error", u"Error en Hola Jose:\n{}".format(str(ex)))
            except Exception:
                pass

    def _on_generar_nombre(self, sender, args):
        try:
            nombre = (self._t3_nombre.Text or "").strip() if self._t3_nombre else ""
            apellido = (self._t3_apellido.Text or "").strip() if self._t3_apellido else ""

            from enfierrado_tipo3_generar_nombre import run as run_tipo3

            run_tipo3(self._revit, {"nombre": nombre, "apellido": apellido})
        except Exception as ex:
            try:
                TaskDialog.Show("Enfierrado - Error", u"Error al generar nombre:\n{}".format(str(ex)))
            except Exception:
                pass

    def show(self):
        """Muestra la ventana con Owner para integrarse con Revit."""
        uidoc = self._revit.ActiveUIDocument
        hwnd = None
        try:
            hwnd = revit_main_hwnd(self._revit.Application)
            if hwnd:
                helper = WindowInteropHelper(self._win)
                helper.Owner = hwnd
        except Exception:
            pass
        position_wpf_window_top_left_at_active_view(self._win, uidoc, hwnd)

        self._win.Show()
        self._win.Activate()


def run(revit, close_on_finish=False):
    """Punto de entrada del botón de prueba."""
    w = EnfierradoDynamicFormWindow(revit, close_on_finish=close_on_finish)
    w.show()
    return w

