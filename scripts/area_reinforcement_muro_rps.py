# -*- coding: utf-8 -*-
"""
Módulo compartido: Area Reinforcement Muro RPS.
Crea AreaReinforcement en muros de hormigón armado.
Misma lógica que Crear Area Reinf. RPS (losas): interfaz gráfica con malla exterior/interior,
tipo de barra, espaciado y ganchos configurables.
Usado por el botón 09_CrearAreaReinforcementMuroRPS.
"""

import os
import weakref
import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")

from System.Windows.Markup import XamlReader
from System.Windows import RoutedEventHandler
from System.Windows.Input import Key, KeyBinding, ModifierKeys, ApplicationCommands, CommandBinding
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption
from System import Uri, UriKind

from revit_wpf_window_position import (
    position_wpf_window_top_left_at_active_view,
    revit_main_hwnd,
)

from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    SubTransaction,
    Transaction,
    UnitUtils,
    UnitTypeId,
    Wall,
    XYZ,
)
from Autodesk.Revit.DB.Structure import AreaReinforcement
from Autodesk.Revit.UI import TaskDialog, ExternalEvent, IExternalEventHandler
from Autodesk.Revit.UI.Selection import ObjectType

# Importar funciones compartidas del módulo de losas
from area_reinforcement_losa import (
    _get_default_area_reinforcement_type_id,
    _get_rebar_bar_types,
    _aplicar_parametros_malla,
    _asignar_hook_a_area_reinforcement,
    _buscar_hook_por_largo,
    _crear_gancho_por_defecto,
    _crear_hook_desde_largo,
    _get_first_rebar_hook_type_id,
)

from bimtools_paths import get_logo_paths

# ── Constantes ─────────────────────────────────────────────────────────────
# Resta al espesor de muro para el largo del gancho (independiente de losas: −60 mm allí).
_HOOK_RESTA_MURO_MM = 40


def _obtener_espesor_muro_mm(wall):
    """
    Obtiene el espesor del muro en mm.
    Prioriza WALL_ATTR_WIDTH_PARAM (instancia) y LookupParameter('Default Thickness')
    en instancia y tipo como fallback.
    """
    if wall is None:
        return None
    # 1) BuiltInParameter WALL_ATTR_WIDTH_PARAM (espesor del muro)
    try:
        param = wall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)
        if param and param.HasValue:
            return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters)
    except Exception:
        pass
    # 2) LookupParameter "Default Thickness" en la instancia
    try:
        param = wall.LookupParameter("Default Thickness")
        if param and param.HasValue:
            return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters)
    except Exception:
        pass
    # 3) LookupParameter en el tipo (Espesor, Thickness)
    try:
        type_id = wall.GetTypeId()
        if type_id and type_id != ElementId.InvalidElementId:
            wall_type = wall.Document.GetElement(type_id)
            if wall_type:
                for pname in ("Default Thickness", "Thickness", "Espesor", "Width"):
                    param = wall_type.LookupParameter(pname)
                    if param and param.HasValue:
                        return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters)
    except Exception:
        pass
    return None


def _obtener_o_crear_hook_desde_espesor_muro(document, wall, en_transaccion=True):
    """
    Obtiene o crea un RebarHookType con Hook Length = espesor_muro - 40 mm.
    Debe ejecutarse ANTES de crear el Area Reinforcement.
    Retorna ElementId del gancho o InvalidElementId si no se pudo obtener.
    """
    espesor_mm = _obtener_espesor_muro_mm(wall)
    if espesor_mm is None:
        return ElementId.InvalidElementId
    largo_target = espesor_mm - _HOOK_RESTA_MURO_MM
    if largo_target <= 0:
        return ElementId.InvalidElementId
    hook = _buscar_hook_por_largo(document, largo_target)
    if hook:
        return hook.Id
    nuevo = _crear_hook_desde_largo(document, largo_target, en_transaccion=en_transaccion)
    return nuevo.Id if nuevo else ElementId.InvalidElementId


def _obtener_direccion_principal_muro(wall):
    """Obtiene la dirección principal del muro desde su curva de ubicación."""
    try:
        loc = wall.Location
        if loc is None:
            return XYZ(1, 0, 0)
        curve = getattr(loc, "Curve", None)
        if curve is None:
            return XYZ(1, 0, 0)
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        dx = p1.X - p0.X
        dy = p1.Y - p0.Y
        dz = p1.Z - p0.Z
        length = (dx * dx + dy * dy + dz * dz) ** 0.5
        if length > 1e-6:
            return XYZ(dx / length, dy / length, dz / length)
    except Exception:
        pass
    return XYZ(1, 0, 0)


# ── ExternalEvent Handlers ───────────────────────────────────────────────────
class ColocarAreaReinforcementMuroHandler(IExternalEventHandler):
    """Ejecuta la creación de Area Reinforcement en muros."""

    def __init__(self, window_ref, get_area_type_fn, aplicar_parametros_fn,
                 get_hook_type_fn, crear_gancho_fn, asignar_hook_fn):
        self._window_ref = window_ref
        self._get_area_type = get_area_type_fn
        self._aplicar_parametros = aplicar_parametros_fn
        self._get_hook_type = get_hook_type_fn
        self._crear_gancho = crear_gancho_fn
        self._asignar_hook = asignar_hook_fn
        self.wall_ids = []
        self.params_dict = {}
        self.layer_active_dict = {}
        self.area_reinforcement_type_id = None
        self.asignar_ganchos = True

    def _actualizar_estado(self, msg):
        try:
            win = self._window_ref() if self._window_ref else None
            if win and hasattr(win, "_set_estado"):
                win._set_estado(msg)
        except Exception:
            pass

    def Execute(self, uiapp):
        from Autodesk.Revit.DB import ElementId, Transaction, Wall, XYZ
        from Autodesk.Revit.DB.Structure import AreaReinforcement
        from Autodesk.Revit.UI import TaskDialog

        self._actualizar_estado(u"Creando Area Reinforcement...")
        try:
            doc = uiapp.ActiveUIDocument.Document
            uidoc = uiapp.ActiveUIDocument
            if not self.wall_ids:
                TaskDialog.Show("Area Reinforcement Muro RPS - Error", u"No hay muros seleccionados.")
                return
            if not self.params_dict or not any(
                pid and pid != ElementId.InvalidElementId
                for pid, _ in self.params_dict.values()
            ):
                TaskDialog.Show("Area Reinforcement Muro RPS - Error",
                    u"Selecciona al menos un tipo de barra válido.")
                return
            area_type_id = self.area_reinforcement_type_id or (
                self._get_area_type(doc) if self._get_area_type else None
            )
            if not area_type_id or area_type_id == ElementId.InvalidElementId:
                TaskDialog.Show("Area Reinforcement Muro RPS - Error",
                    u"No hay AreaReinforcementType en el proyecto. Crea uno manualmente.")
                return
            first_bar_id = None
            for pid, _ in self.params_dict.values():
                if pid and pid != ElementId.InvalidElementId:
                    first_bar_id = pid
                    break
            if not first_bar_id and self.params_dict:
                first_bar_id = list(self.params_dict.values())[0][0]

            creados = 0
            errores = []
            trans = Transaction(doc, "Area Reinforcement en muros")
            try:
                trans.Start()
                for wall_id in self.wall_ids:
                    wall = doc.GetElement(wall_id)
                    if not wall or not isinstance(wall, Wall):
                        errores.append(u"ID {}: no es muro válido".format(wall_id.IntegerValue))
                        continue
                    rebar_hook_type_id = ElementId.InvalidElementId
                    if self.asignar_ganchos:
                        sub = SubTransaction(doc)
                        try:
                            sub.Start()
                            rebar_hook_type_id = _obtener_o_crear_hook_desde_espesor_muro(doc, wall, en_transaccion=False)
                            if not rebar_hook_type_id or rebar_hook_type_id == ElementId.InvalidElementId:
                                rebar_hook_type_id = self._get_hook_type(doc) if self._get_hook_type else ElementId.InvalidElementId
                            if not rebar_hook_type_id or rebar_hook_type_id == ElementId.InvalidElementId:
                                rebar_hook_type_id = self._crear_gancho(doc) if self._crear_gancho else ElementId.InvalidElementId
                            sub.Commit()
                        except Exception as ex_sub:
                            if sub.HasStarted():
                                try:
                                    sub.RollBack()
                                except Exception:
                                    pass
                            rebar_hook_type_id = ElementId.InvalidElementId
                            errores.append(u"ID {}: gancho - {}".format(wall_id.IntegerValue, str(ex_sub)))
                            continue
                    major_direction = _obtener_direccion_principal_muro(wall)
                    try:
                        area_rein = AreaReinforcement.Create(
                            doc,
                            wall,
                            major_direction,
                            area_type_id,
                            first_bar_id,
                            rebar_hook_type_id,
                        )
                        if self.asignar_ganchos and self._asignar_hook and rebar_hook_type_id and rebar_hook_type_id != ElementId.InvalidElementId:
                            self._asignar_hook(area_rein, rebar_hook_type_id)
                        if area_rein and self._aplicar_parametros:
                            self._aplicar_parametros(
                                area_rein,
                                self.params_dict,
                                self.layer_active_dict,
                            )
                        creados += 1
                    except Exception as ex:
                        errores.append(u"ID {}: {}".format(wall_id.IntegerValue, str(ex)))
                trans.Commit()
                msg = u"Area Reinforcement creado en {} muro(s).".format(creados)
                if errores:
                    msg += u" Errores: " + "; ".join(errores[:3])
                    if len(errores) > 3:
                        msg += u"..."
                self._actualizar_estado(msg)
                TaskDialog.Show("Area Reinforcement Muro RPS", msg)
            except Exception as ex:
                if trans.HasStarted():
                    try:
                        trans.RollBack()
                    except Exception:
                        pass
                self._actualizar_estado(u"Error: {}".format(str(ex)))
                TaskDialog.Show("Area Reinforcement Muro RPS - Error", u"Error:\n\n{}".format(str(ex)))
        except Exception as ex:
            self._actualizar_estado(u"Error: {}".format(str(ex)))
            TaskDialog.Show("Area Reinforcement Muro RPS - Error", u"Error:\n\n{}".format(str(ex)))
        finally:
            try:
                win = self._window_ref() if self._window_ref else None
                if win and hasattr(win, "_win"):
                    win._win.Activate()
            except Exception:
                pass

    def GetName(self):
        return "ColocarAreaReinforcementMuro"


class SeleccionarMuroHandler(IExternalEventHandler):
    """Ejecuta la selección de muros en contexto API de Revit."""

    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        from Autodesk.Revit.UI.Selection import ObjectType
        from Autodesk.Revit.UI import TaskDialog
        from Autodesk.Revit.DB import BuiltInCategory

        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        win = self._window_ref()
        if not win:
            return
        try:
            win._document = doc
            refs = list(uidoc.Selection.PickObjects(
                ObjectType.Element,
                u"Selecciona uno o más muros. Finaliza con Finish o Cancel.",
            ))
            if not refs:
                win._set_estado(u"Selección cancelada.")
            else:
                wall_ids = []
                for ref in refs:
                    elem = doc.GetElement(ref.ElementId)
                    if elem and elem.Category and int(elem.Category.Id.IntegerValue) == int(BuiltInCategory.OST_Walls):
                        wall_ids.append(ref.ElementId)
                if wall_ids:
                    win._wall_ids = wall_ids
                    win._actualizar_info_muros(doc, wall_ids)
                    win._set_estado(u"{} muro(s) seleccionado(s).".format(len(wall_ids)))
                    if len(refs) > len(wall_ids):
                        TaskDialog.Show("Area Reinforcement Muro RPS",
                            u"{} muro(s) seleccionado(s). {} elemento(s) no son muros y se ignoraron.".format(
                                len(wall_ids), len(refs) - len(wall_ids)))
                else:
                    win._set_estado(u"Ningún muro en la selección.")
                    TaskDialog.Show("Area Reinforcement Muro RPS", u"Los elementos seleccionados no son muros.")
        except Exception as ex:
            err = str(ex).lower()
            if "cancel" not in err and "operation" not in err:
                win._set_estado(u"Error: {}.".format(str(ex)))
                TaskDialog.Show("Area Reinforcement Muro RPS - Error", str(ex))
            else:
                win._set_estado(u"Selección cancelada.")
        finally:
            try:
                win._win.Show()
                win._win.Activate()
            except Exception:
                pass

    def GetName(self):
        return "SeleccionarMuro"


# ── XAML — Estilo Arainco (igual que Area Reinforcement Losa) ───────────────
XAML = """
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Arainco - Area Reinforcement Muro RPS"
    Height="580" Width="520"
    MinHeight="540" MinWidth="480"
    WindowStartupLocation="Manual"
    Background="#0A1C26"
    FontFamily="Segoe UI"
    ResizeMode="CanResize">

  <Window.Resources>
    <Style x:Key="Label" TargetType="TextBlock">
      <Setter Property="Foreground"  Value="#4A8BA6"/>
      <Setter Property="FontSize"    Value="11"/>
      <Setter Property="FontWeight"  Value="SemiBold"/>
      <Setter Property="Margin"      Value="0,10,0,4"/>
    </Style>
    <Style x:Key="LabelSmall" TargetType="TextBlock" BasedOn="{StaticResource Label}">
      <Setter Property="FontSize"    Value="10"/>
      <Setter Property="Margin"      Value="0,4,0,2"/>
    </Style>
    <Style x:Key="Combo" TargetType="ComboBox">
      <Setter Property="Background"      Value="#0D2234"/>
      <Setter Property="Foreground"      Value="#FFFFFF"/>
      <Setter Property="BorderBrush"     Value="#1A3D52"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="FontSize"        Value="13"/>
      <Setter Property="Height"          Value="32"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBox">
            <Grid TextElement.Foreground="{TemplateBinding Foreground}">
              <Border x:Name="Border"
                      Background="{TemplateBinding Background}"
                      BorderBrush="{TemplateBinding BorderBrush}"
                      BorderThickness="{TemplateBinding BorderThickness}"
                      CornerRadius="6"/>
              <ToggleButton IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                            Focusable="False" Background="Transparent" BorderThickness="0"
                            HorizontalAlignment="Stretch" VerticalAlignment="Stretch"/>
              <ContentPresenter x:Name="ContentSite"
                                Content="{TemplateBinding SelectionBoxItem}"
                                ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                Margin="12,0,32,0" VerticalAlignment="Center" IsHitTestVisible="False"/>
              <TextBox x:Name="PART_EditableTextBox"
                       Visibility="Collapsed"
                       Background="Transparent" Foreground="{TemplateBinding Foreground}"
                       BorderThickness="0" Margin="12,0,32,0" VerticalAlignment="Center"
                       FontSize="{TemplateBinding FontSize}" FontFamily="{TemplateBinding FontFamily}"
                       CaretBrush="#5BB8D4" Padding="0" VerticalContentAlignment="Center"/>
              <TextBlock Text="&#9660;" FontSize="9" Foreground="#5BB8D4"
                         HorizontalAlignment="Right" VerticalAlignment="Center"
                         Margin="0,0,12,0" IsHitTestVisible="False"/>
              <Popup x:Name="PART_Popup"
                     IsOpen="{TemplateBinding IsDropDownOpen}"
                     AllowsTransparency="True" Focusable="False"
                     PopupAnimation="Fade" Placement="Bottom">
                <Border Background="#0D2234" BorderBrush="#1A3D52" BorderThickness="1"
                        MinWidth="{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}">
                  <ScrollViewer MaxHeight="200" VerticalScrollBarVisibility="Auto">
                    <ItemsPresenter/>
                  </ScrollViewer>
                </Border>
              </Popup>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsEditable" Value="True">
                <Setter TargetName="ContentSite" Property="Visibility" Value="Collapsed"/>
                <Setter TargetName="PART_EditableTextBox" Property="Visibility" Value="Visible"/>
              </Trigger>
              <Trigger Property="IsEditable" Value="False">
                <Setter TargetName="ContentSite" Property="Visibility" Value="Visible"/>
                <Setter TargetName="PART_EditableTextBox" Property="Visibility" Value="Collapsed"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Border" Property="BorderBrush" Value="#5BB8D4"/>
              </Trigger>
              <Trigger Property="IsDropDownOpen" Value="True">
                <Setter TargetName="Border" Property="BorderBrush" Value="#5BB8D4"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="ComboItem" TargetType="ComboBoxItem">
      <Setter Property="Background"  Value="#0D2234"/>
      <Setter Property="Foreground"  Value="#FFFFFF"/>
      <Setter Property="Padding"     Value="10,8"/>
      <Style.Triggers>
        <Trigger Property="IsHighlighted" Value="True">
          <Setter Property="Background" Value="#1A4F6A"/>
        </Trigger>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="#1A4F6A"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="BtnPrimary" TargetType="Button">
      <Setter Property="Background"      Value="#5BB8D4"/>
      <Setter Property="Foreground"      Value="#0A1C26"/>
      <Setter Property="FontWeight"      Value="Bold"/>
      <Setter Property="FontSize"        Value="13"/>
      <Setter Property="Padding"         Value="20,9"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}" CornerRadius="7"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="#7CCDE2"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter Property="Background" Value="#4AA5C0"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="BtnGhost" TargetType="Button" BasedOn="{StaticResource BtnPrimary}">
      <Setter Property="Background" Value="#0F2535"/>
      <Setter Property="Foreground" Value="#C8E4EF"/>
    </Style>
  </Window.Resources>

  <Grid Margin="16,12">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Header: LOGO + ENCABEZADO -->
    <Border Grid.Row="0" Background="#0F2535" CornerRadius="6" Padding="12,10" Margin="0,0,0,12">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Image x:Name="ImgLogo" Width="48" Height="48" Grid.Column="0"
               Stretch="Uniform" Margin="0,0,12,0" VerticalAlignment="Center"/>
        <StackPanel Grid.Column="1" VerticalAlignment="Center">
          <TextBlock Text="AREA REINFORCEMENT MURO" FontSize="16" FontWeight="Bold"
                     Foreground="#C8E4EF"/>
          <TextBlock Text="Selecciona muros y configura malla exterior e interior" FontSize="11"
                     Foreground="#4A8BA6" Margin="0,4,0,0"/>
        </StackPanel>
      </Grid>
    </Border>

    <!-- SELECCIONAR MUROS EN MODELO -->
    <StackPanel Grid.Row="1" Margin="0,0,0,10">
      <Button x:Name="BtnSeleccionar" Content="SELECCIONAR MUROS EN MODELO"
              Style="{StaticResource BtnGhost}"
              HorizontalAlignment="Stretch" Padding="16,12"/>
      <TextBlock x:Name="TxtMuroInfo" Text="Ningún muro seleccionado. Permite selección múltiple."
                 Foreground="#4A8BA6" FontSize="11" Margin="0,6,0,0" TextWrapping="Wrap"/>
    </StackPanel>

    <!-- MALLA EXTERIOR -->
    <Border Grid.Row="2" Background="#0F2535" CornerRadius="6" Padding="12,10" Margin="0,0,0,10"
            BorderBrush="#1A3D52" BorderThickness="1">
      <StackPanel>
        <CheckBox x:Name="ChkMallaExterior" IsChecked="True" Content="MALLA EXTERIOR"
                  Foreground="#C8E4EF" Style="{StaticResource Label}" VerticalAlignment="Center" Margin="0,0,0,8"/>
        <StackPanel x:Name="PanelExteriorContent" IsEnabled="True">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="12"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Border Grid.Column="0" Background="#0A1C26" CornerRadius="4" Padding="8,6" BorderBrush="#1A3D52" BorderThickness="1">
              <StackPanel>
                <TextBlock Text="Major" Style="{StaticResource LabelSmall}" Margin="0,0,0,4"/>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="6"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <StackPanel Grid.Column="0">
                    <TextBlock Text="Diam." Style="{StaticResource LabelSmall}"/>
                    <ComboBox x:Name="CmbExteriorMajorDiametro" Style="{StaticResource Combo}" IsEditable="False" IsReadOnly="True">
                      <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                    </ComboBox>
                  </StackPanel>
                  <StackPanel Grid.Column="2">
                    <TextBlock Text="Esp. (mm)" Style="{StaticResource LabelSmall}"/>
                    <ComboBox x:Name="CmbExteriorMajorEspaciamiento" Style="{StaticResource Combo}" IsEditable="True" IsReadOnly="False">
                      <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                    </ComboBox>
                  </StackPanel>
                </Grid>
              </StackPanel>
            </Border>
            <Border Grid.Column="2" Background="#0A1C26" CornerRadius="4" Padding="8,6" BorderBrush="#1A3D52" BorderThickness="1">
              <StackPanel>
                <TextBlock Text="Minor" Style="{StaticResource LabelSmall}" Margin="0,0,0,4"/>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="6"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <StackPanel Grid.Column="0">
                    <TextBlock Text="Diam." Style="{StaticResource LabelSmall}"/>
                    <ComboBox x:Name="CmbExteriorMinorDiametro" Style="{StaticResource Combo}" IsEditable="False" IsReadOnly="True">
                      <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                    </ComboBox>
                  </StackPanel>
                  <StackPanel Grid.Column="2">
                    <TextBlock Text="Esp. (mm)" Style="{StaticResource LabelSmall}"/>
                    <ComboBox x:Name="CmbExteriorMinorEspaciamiento" Style="{StaticResource Combo}" IsEditable="True" IsReadOnly="False">
                      <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                    </ComboBox>
                  </StackPanel>
                </Grid>
              </StackPanel>
            </Border>
          </Grid>
        </StackPanel>
      </StackPanel>
    </Border>

    <!-- MALLA INTERIOR -->
    <Border Grid.Row="3" Background="#0F2535" CornerRadius="6" Padding="12,10" Margin="0,0,0,10"
            BorderBrush="#1A3D52" BorderThickness="1">
      <StackPanel>
        <CheckBox x:Name="ChkMallaInterior" IsChecked="True" Content="MALLA INTERIOR"
                  Foreground="#C8E4EF" Style="{StaticResource Label}" VerticalAlignment="Center" Margin="0,0,0,8"/>
        <StackPanel x:Name="PanelInteriorContent" IsEnabled="True">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="12"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Border Grid.Column="0" Background="#0A1C26" CornerRadius="4" Padding="8,6" BorderBrush="#1A3D52" BorderThickness="1">
              <StackPanel>
                <TextBlock Text="Major" Style="{StaticResource LabelSmall}" Margin="0,0,0,4"/>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="6"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <StackPanel Grid.Column="0">
                    <TextBlock Text="Diam." Style="{StaticResource LabelSmall}"/>
                    <ComboBox x:Name="CmbInteriorMajorDiametro" Style="{StaticResource Combo}" IsEditable="False" IsReadOnly="True">
                      <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                    </ComboBox>
                  </StackPanel>
                  <StackPanel Grid.Column="2">
                    <TextBlock Text="Esp. (mm)" Style="{StaticResource LabelSmall}"/>
                    <ComboBox x:Name="CmbInteriorMajorEspaciamiento" Style="{StaticResource Combo}" IsEditable="True" IsReadOnly="False">
                      <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                    </ComboBox>
                  </StackPanel>
                </Grid>
              </StackPanel>
            </Border>
            <Border Grid.Column="2" Background="#0A1C26" CornerRadius="4" Padding="8,6" BorderBrush="#1A3D52" BorderThickness="1">
              <StackPanel>
                <TextBlock Text="Minor" Style="{StaticResource LabelSmall}" Margin="0,0,0,4"/>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="6"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <StackPanel Grid.Column="0">
                    <TextBlock Text="Diam." Style="{StaticResource LabelSmall}"/>
                    <ComboBox x:Name="CmbInteriorMinorDiametro" Style="{StaticResource Combo}" IsEditable="False" IsReadOnly="True">
                      <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                    </ComboBox>
                  </StackPanel>
                  <StackPanel Grid.Column="2">
                    <TextBlock Text="Esp. (mm)" Style="{StaticResource LabelSmall}"/>
                    <ComboBox x:Name="CmbInteriorMinorEspaciamiento" Style="{StaticResource Combo}" IsEditable="True" IsReadOnly="False">
                      <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                    </ComboBox>
                  </StackPanel>
                </Grid>
              </StackPanel>
            </Border>
          </Grid>
        </StackPanel>
      </StackPanel>
    </Border>

    <!-- OPCIONES -->
    <StackPanel Grid.Row="4" Margin="0,0,0,10">
      <CheckBox x:Name="ChkAsignarGanchos" IsChecked="True" Foreground="#C8E4EF"
                Content="Asignar ganchos (según espesor de muro - 40 mm)"
                VerticalAlignment="Center" Margin="0,4,0,4" FontSize="11"/>
    </StackPanel>

    <!-- COLOCAR ARMADURAS -->
    <StackPanel Grid.Row="5">
      <TextBlock x:Name="TxtEstado" Text="" Foreground="#5BB8D4" FontSize="11"
                 Margin="0,0,0,8" TextWrapping="Wrap"/>
      <Button x:Name="BtnColocar" Content="COLOCAR ARMADURAS"
              Style="{StaticResource BtnPrimary}"
              HorizontalAlignment="Stretch" Padding="20,12"/>
    </StackPanel>
  </Grid>
</Window>
"""


# ── Ventana principal ───────────────────────────────────────────────────────
class AreaReinforcementMuroRPSWindow(object):
    def __init__(self, revit):
        self._document = None
        self._wall_ids = []
        self._rebar_type_ids = {}
        self._area_reinforcement_type_id = None
        self._revit = revit
        self._win = XamlReader.Parse(XAML)
        self._colocar_handler = ColocarAreaReinforcementMuroHandler(
            weakref.ref(self),
            _get_default_area_reinforcement_type_id,
            _aplicar_parametros_malla,
            _get_first_rebar_hook_type_id,
            _crear_gancho_por_defecto,
            _asignar_hook_a_area_reinforcement,
        )
        self._colocar_event = ExternalEvent.Create(self._colocar_handler)
        self._seleccion_handler = SeleccionarMuroHandler(weakref.ref(self))
        self._seleccion_event = ExternalEvent.Create(self._seleccion_handler)
        self._setup_ui()

    def _set_estado(self, msg):
        try:
            txt = self._win.FindName("TxtEstado")
            if txt:
                txt.Text = msg
        except Exception:
            pass

    def _actualizar_info_muro(self, wall, document):
        try:
            txt = self._win.FindName("TxtMuroInfo")
            if txt and wall:
                name = wall.Name or "Muro"
                tipo = ""
                try:
                    tid = wall.GetTypeId()
                    if tid:
                        ft = document.GetElement(tid)
                        if ft:
                            tipo = ft.Name or ""
                except Exception:
                    pass
                info = u"{} | Tipo: {} | ID: {}".format(name, tipo, wall.Id.IntegerValue)
                txt.Text = info
        except Exception:
            pass

    def _actualizar_info_muros(self, document, wall_ids):
        try:
            txt = self._win.FindName("TxtMuroInfo")
            if txt and wall_ids:
                if len(wall_ids) == 1:
                    wall = document.GetElement(wall_ids[0])
                    if wall:
                        self._actualizar_info_muro(wall, document)
                    else:
                        txt.Text = u"1 muro seleccionado (ID: {})".format(wall_ids[0].IntegerValue)
                else:
                    txt.Text = u"{} muros seleccionados (IDs: {})".format(
                        len(wall_ids),
                        ", ".join(str(wid.IntegerValue) for wid in wall_ids[:5])
                        + ("..." if len(wall_ids) > 5 else ""),
                    )
        except Exception:
            pass

    def _setup_ui(self):
        btn_sel = self._win.FindName("BtnSeleccionar")
        btn_col = self._win.FindName("BtnColocar")
        if btn_col:
            btn_col.Click += RoutedEventHandler(self._on_colocar)
        if btn_sel:
            btn_sel.Click += RoutedEventHandler(self._on_seleccionar)
        chk_ext = self._win.FindName("ChkMallaExterior")
        chk_int = self._win.FindName("ChkMallaInterior")
        if chk_ext:
            chk_ext.Checked += RoutedEventHandler(self._on_chk_malla_exterior_changed)
            chk_ext.Unchecked += RoutedEventHandler(self._on_chk_malla_exterior_changed)
        if chk_int:
            chk_int.Checked += RoutedEventHandler(self._on_chk_malla_interior_changed)
            chk_int.Unchecked += RoutedEventHandler(self._on_chk_malla_interior_changed)
        def _on_close_cmd(sender, e):
            try:
                self._win.Close()
            except Exception:
                pass
        self._win.CommandBindings.Add(CommandBinding(ApplicationCommands.Close, _on_close_cmd))
        self._win.InputBindings.Add(KeyBinding(ApplicationCommands.Close, Key.Escape, ModifierKeys.None))
        self._win.Loaded += RoutedEventHandler(self._on_window_loaded)

    def _on_chk_malla_exterior_changed(self, sender, args):
        try:
            chk = self._win.FindName("ChkMallaExterior")
            enabled = chk.IsChecked == True if chk else True
            panel = self._win.FindName("PanelExteriorContent")
            if panel:
                panel.IsEnabled = enabled
                panel.Opacity = 1.0 if enabled else 0.35
        except Exception:
            pass

    def _on_chk_malla_interior_changed(self, sender, args):
        try:
            chk = self._win.FindName("ChkMallaInterior")
            enabled = chk.IsChecked == True if chk else True
            panel = self._win.FindName("PanelInteriorContent")
            if panel:
                panel.IsEnabled = enabled
                panel.Opacity = 1.0 if enabled else 0.35
        except Exception:
            pass

    def _sync_malla_checkboxes(self):
        try:
            self._on_chk_malla_exterior_changed(None, None)
            self._on_chk_malla_interior_changed(None, None)
        except Exception:
            pass

    def _on_window_loaded(self, sender, args):
        self._load_logo()
        self._cargar_combos()
        self._sync_malla_checkboxes()

    def _load_logo(self):
        try:
            img_ctrl = self._win.FindName("ImgLogo")
            if not img_ctrl:
                return
            for logo_path in get_logo_paths():
                if os.path.exists(logo_path):
                    bmp = BitmapImage()
                    bmp.BeginInit()
                    bmp.UriSource = Uri(logo_path, UriKind.Absolute)
                    bmp.CacheOption = BitmapCacheOption.OnLoad
                    bmp.EndInit()
                    bmp.Freeze()
                    img_ctrl.Source = bmp
                    break
        except Exception:
            pass

    def _cargar_combos(self):
        try:
            d = self._document or self._revit.ActiveUIDocument.Document
        except Exception:
            d = self._document
        self._area_reinforcement_type_id = _get_default_area_reinforcement_type_id(d)
        if not self._area_reinforcement_type_id:
            self._set_estado(u"No hay tipo de Area Reinforcement. Usa plantilla estructural.")
        bar_types = _get_rebar_bar_types(d)
        self._rebar_type_ids = {}
        for disp, bar_type in bar_types:
            self._rebar_type_ids[str(disp)] = bar_type.Id
        bar_disps = [disp for disp, _ in bar_types]
        espaciamientos = ["100", "150", "200", "250", "300"]
        diam_names = (
            "CmbExteriorMajorDiametro", "CmbExteriorMinorDiametro",
            "CmbInteriorMajorDiametro", "CmbInteriorMinorDiametro",
        )
        esp_names = (
            "CmbExteriorMajorEspaciamiento", "CmbExteriorMinorEspaciamiento",
            "CmbInteriorMajorEspaciamiento", "CmbInteriorMinorEspaciamiento",
        )
        for name in diam_names:
            cmb = self._win.FindName(name)
            if cmb:
                cmb.ItemsSource = bar_disps
                if bar_types:
                    cmb.SelectedIndex = min(1, len(bar_types) - 1) if len(bar_types) > 1 else 0
        for name in esp_names:
            cmb = self._win.FindName(name)
            if cmb:
                cmb.ItemsSource = espaciamientos
                cmb.SelectedIndex = 1  # 150 mm por defecto

    def _on_seleccionar(self, sender, args):
        self._win.Hide()
        self._seleccion_event.Raise()

    def _on_colocar(self, sender, args):
        from Autodesk.Revit.DB import ElementId
        from Autodesk.Revit.UI import TaskDialog

        try:
            if not self._wall_ids:
                TaskDialog.Show("Area Reinforcement Muro RPS", u"Primero selecciona uno o más muros.")
                self._set_estado(u"Selecciona muro(s) antes de crear.")
                return
            chk_ext = self._win.FindName("ChkMallaExterior")
            chk_int = self._win.FindName("ChkMallaInterior")
            if chk_ext and chk_int and chk_ext.IsChecked != True and chk_int.IsChecked != True:
                TaskDialog.Show("Area Reinforcement Muro RPS", u"Por lo menos una malla debe estar activada.")
                self._set_estado(u"Por lo menos una malla debe estar activada.")
                return
            area_type_id = self._area_reinforcement_type_id
            if not area_type_id:
                try:
                    d = self._document or self._revit.ActiveUIDocument.Document
                    area_type_id = _get_default_area_reinforcement_type_id(d)
                except Exception:
                    pass
            if not area_type_id:
                TaskDialog.Show("Area Reinforcement Muro RPS", u"No hay tipo de Area Reinforcement en el proyecto.")
                return
            rebar_ids = getattr(self, "_rebar_type_ids", {})
            chk_ext = self._win.FindName("ChkMallaExterior")
            chk_int = self._win.FindName("ChkMallaInterior")
            malla_exterior_activa = chk_ext.IsChecked == True if chk_ext else True
            malla_interior_activa = chk_int.IsChecked == True if chk_int else True
            layer_config = [
                ("exterior_major", "CmbExteriorMajorDiametro", "CmbExteriorMajorEspaciamiento", malla_exterior_activa),
                ("exterior_minor", "CmbExteriorMinorDiametro", "CmbExteriorMinorEspaciamiento", malla_exterior_activa),
                ("interior_major", "CmbInteriorMajorDiametro", "CmbInteriorMajorEspaciamiento", malla_interior_activa),
                ("interior_minor", "CmbInteriorMinorDiametro", "CmbInteriorMinorEspaciamiento", malla_interior_activa),
            ]
            params_dict = {}
            layer_active_dict = {}
            for layer_key, diam_name, esp_name, is_active in layer_config:
                cmb_diam = self._win.FindName(diam_name)
                cmb_esp = self._win.FindName(esp_name)
                bar_id = rebar_ids.get(str(cmb_diam.SelectedItem if cmb_diam else None), ElementId.InvalidElementId)
                esp = (cmb_esp.SelectedItem if cmb_esp and cmb_esp.SelectedItem else None) or (cmb_esp.Text if cmb_esp else None) or "150"
                params_dict[layer_key] = (bar_id, str(esp))
                layer_active_dict[layer_key] = bool(is_active)
            if not any(pid and pid != ElementId.InvalidElementId for pid, _ in params_dict.values()):
                TaskDialog.Show("Area Reinforcement Muro RPS", u"Selecciona al menos un diámetro válido.")
                return
            chk_ganchos = self._win.FindName("ChkAsignarGanchos")
            self._colocar_handler.wall_ids = list(self._wall_ids)
            self._colocar_handler.params_dict = params_dict
            self._colocar_handler.layer_active_dict = layer_active_dict
            self._colocar_handler.area_reinforcement_type_id = area_type_id
            self._colocar_handler.asignar_ganchos = chk_ganchos.IsChecked == True if chk_ganchos else True
            self._colocar_event.Raise()
            self._set_estado(u"Creando Area Reinforcement...")
        except Exception as ex:
            self._set_estado(u"Error: {}".format(str(ex)))
            TaskDialog.Show("Area Reinforcement Muro RPS - Error", u"Error:\n\n{}".format(str(ex)))

    def show(self):
        uidoc = self._revit.ActiveUIDocument
        hwnd = None
        try:
            from System.Windows.Interop import WindowInteropHelper

            hwnd = revit_main_hwnd(self._revit.Application)
            if hwnd:
                helper = WindowInteropHelper(self._win)
                helper.Owner = hwnd
        except Exception:
            pass
        position_wpf_window_top_left_at_active_view(self._win, uidoc, hwnd)
        self._document = self._revit.ActiveUIDocument.Document
        self._cargar_combos()
        self._win.Show()
        self._win.Activate()


def run(revit):
    """Punto de entrada: lanza la ventana de Area Reinforcement Muro RPS."""
    w = AreaReinforcementMuroRPSWindow(revit)
    w.show()
