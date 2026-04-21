# -*- coding: utf-8 -*-
"""
Recursos WPF compartidos (tema oscuro BIMTools) para ventanas embebidas en scripts IronPython.

Incluye: Label, LabelSmall, GbParams, Combo, ComboItem, BtnPrimary, BtnSelectOutline,
SpinRepeatBtn, CantSpinnerText, BtnCloseX_MinimalNoBg (linea visual Fundacion Aislada).

Uso: dentro de Window.Resources, tras Storyboard si aplica, concatenar BIMTOOLS_DARK_STYLES_XML.
"""

BIMTOOLS_DARK_STYLES_XML = u"""
    <Style x:Key="Label" TargetType="TextBlock">
      <Setter Property="Foreground"  Value="#95B8CC"/>
      <Setter Property="FontSize"    Value="11"/>
      <Setter Property="FontWeight"  Value="SemiBold"/>
      <Setter Property="Margin"      Value="0,2,0,1"/>
    </Style>
    <Style x:Key="LabelSmall" TargetType="TextBlock" BasedOn="{StaticResource Label}">
      <Setter Property="FontSize"    Value="10"/>
      <Setter Property="FontWeight"  Value="SemiBold"/>
      <Setter Property="Foreground"  Value="#95B8CC"/>
      <Setter Property="Margin"      Value="0,0,0,2"/>
    </Style>
    <Style x:Key="GbParams" TargetType="GroupBox">
      <Setter Property="BorderBrush" Value="#21465C"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Background" Value="#0E1B32"/>
      <Setter Property="Padding" Value="0,3,0,0"/>
      <Setter Property="Margin" Value="0,0,0,10"/>
      <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="GroupBox">
            <Grid SnapsToDevicePixels="True" HorizontalAlignment="Stretch">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <Border Grid.Row="0" Grid.Column="0" Background="#11253D" BorderBrush="{TemplateBinding BorderBrush}"
                      BorderThickness="1,1,1,0" CornerRadius="6,6,0,0" Padding="10,6,10,5" HorizontalAlignment="Stretch">
                <ContentPresenter ContentSource="Header" RecognizesAccessKey="True"
                                  HorizontalAlignment="Stretch" VerticalAlignment="Center"/>
              </Border>
              <Border Grid.Row="1" Grid.Column="0" Background="#0E1B32" BorderBrush="{TemplateBinding BorderBrush}"
                      BorderThickness="1,0,1,1" CornerRadius="0,0,6,6" Padding="10,7,10,9" HorizontalAlignment="Stretch">
                <ContentPresenter HorizontalAlignment="Stretch"/>
              </Border>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Combo" TargetType="ComboBox">
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Background"      Value="#050E18"/>
      <Setter Property="Foreground"      Value="#FFFFFF"/>
      <Setter Property="FontWeight"      Value="Bold"/>
      <Setter Property="BorderBrush"     Value="#1A3A4D"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="FontSize"        Value="11"/>
      <Setter Property="Height"          Value="24"/>
      <Setter Property="Width"           Value="110"/>
      <Setter Property="MinWidth"       Value="110"/>
      <Setter Property="MaxWidth"       Value="110"/>
      <Setter Property="HorizontalAlignment" Value="Left"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBox">
            <Grid SnapsToDevicePixels="True">
              <Border x:Name="Border" CornerRadius="5" Background="#050E18"
                      BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True">
                <Grid TextElement.Foreground="{TemplateBinding Foreground}"
                      TextElement.FontWeight="{TemplateBinding FontWeight}">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="18"/>
                  </Grid.ColumnDefinitions>
                  <ContentPresenter x:Name="ContentSite" Grid.Column="0"
                                    Content="{TemplateBinding SelectionBoxItem}"
                                    ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                    Margin="6,0,4,0" VerticalAlignment="Center"
                                    HorizontalAlignment="Left" IsHitTestVisible="False"/>
                  <TextBox x:Name="PART_EditableTextBox"
                           Grid.Column="0" Visibility="Collapsed"
                           Background="#050E18" Foreground="{TemplateBinding Foreground}"
                           BorderThickness="0" Margin="6,0,4,0" VerticalAlignment="Center"
                           FontSize="{TemplateBinding FontSize}" FontFamily="{TemplateBinding FontFamily}"
                           CaretBrush="#7AA3B8" Padding="0" VerticalContentAlignment="Center"
                           SelectionBrush="#2A4A5C" SelectionOpacity="0.95"
                           FocusVisualStyle="{x:Null}"/>
                  <Border Grid.Column="1" Background="#11253D" BorderBrush="#1A3A4D"
                          BorderThickness="1,0,0,0" CornerRadius="0,5,5,0" ClipToBounds="True">
                    <TextBlock Text="&#9660;" FontSize="8" Foreground="#7AA3B8"
                               HorizontalAlignment="Center" VerticalAlignment="Center"
                               IsHitTestVisible="False"/>
                  </Border>
                  <ToggleButton Grid.Column="0" Grid.ColumnSpan="2"
                                IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                Focusable="False" FocusVisualStyle="{x:Null}"
                                HorizontalAlignment="Stretch" VerticalAlignment="Stretch"
                                Background="Transparent" BorderThickness="0">
                    <ToggleButton.Template>
                      <ControlTemplate TargetType="ToggleButton">
                        <Border Background="Transparent" BorderThickness="0"/>
                      </ControlTemplate>
                    </ToggleButton.Template>
                  </ToggleButton>
                </Grid>
              </Border>
              <Popup x:Name="PART_Popup"
                     IsOpen="{TemplateBinding IsDropDownOpen}"
                     AllowsTransparency="True" Focusable="False"
                     PopupAnimation="Fade" Placement="Bottom"
                     PlacementTarget="{Binding ElementName=Border}">
                <Border Background="#050E18" BorderBrush="#1A3A4D" BorderThickness="1"
                        CornerRadius="5"
                        MinWidth="{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}">
                  <ScrollViewer MaxHeight="220" VerticalScrollBarVisibility="Auto">
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
                <Setter TargetName="Border" Property="Background" Value="#0B1728"/>
                <Setter TargetName="Border" Property="BorderBrush" Value="#4C7383"/>
              </Trigger>
              <Trigger Property="IsKeyboardFocusWithin" Value="True">
                <Setter TargetName="Border" Property="Background" Value="#0B1728"/>
                <Setter TargetName="Border" Property="BorderBrush" Value="#4C7383"/>
                <Setter TargetName="Border" Property="BorderThickness" Value="2"/>
              </Trigger>
              <Trigger Property="IsDropDownOpen" Value="True">
                <Setter TargetName="Border" Property="Background" Value="#0B1728"/>
                <Setter TargetName="Border" Property="BorderBrush" Value="#4C7383"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="ComboItem" TargetType="ComboBoxItem">
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Background"  Value="#050E18"/>
      <Setter Property="Foreground"  Value="#FFFFFF"/>
      <Setter Property="FontWeight"  Value="Bold"/>
      <Setter Property="Padding"     Value="8,2"/>
      <Style.Triggers>
        <Trigger Property="IsHighlighted" Value="True">
          <Setter Property="Background" Value="#1A3A4D"/>
        </Trigger>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="#1A3A4D"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="BtnPrimary" TargetType="Button">
      <Setter Property="Background"      Value="#5BC0DE"/>
      <Setter Property="Foreground"      Value="#0A1A2F"/>
      <Setter Property="FontWeight"      Value="Bold"/>
      <Setter Property="FontSize"        Value="12"/>
      <Setter Property="Padding"         Value="18,9"/>
      <Setter Property="BorderBrush"     Value="#87D9EE"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Root" Background="{TemplateBinding Background}" CornerRadius="5"
                    BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Root" Property="Background" Value="#74CEE8"/>
                <Setter TargetName="Root" Property="BorderBrush" Value="#A5E6F5"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="Root" Property="Background" Value="#49B1D0"/>
                <Setter TargetName="Root" Property="BorderBrush" Value="#79CDE4"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="BtnSelectOutline" TargetType="Button">
      <Setter Property="Background"      Value="#0A1627"/>
      <Setter Property="Foreground"      Value="#E8F4F8"/>
      <Setter Property="FontWeight"      Value="SemiBold"/>
      <Setter Property="FontSize"        Value="11"/>
      <Setter Property="Padding"         Value="10,7"/>
      <Setter Property="BorderBrush"     Value="#5BC0DE"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Root" Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="5"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Root" Property="Background" Value="#0E1B32"/>
                <Setter TargetName="Root" Property="BorderBrush" Value="#7ED4ED"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="Root" Property="Background" Value="#050E18"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="Root" Property="Opacity" Value="0.45"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <!-- Mismo patrón que enfierrado_vigas (spinner cantidad / separación estribos). -->
    <Style x:Key="SpinRepeatBtn" TargetType="RepeatButton">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="#7AA3B8"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="FontSize" Value="8"/>
      <Setter Property="Padding" Value="0"/>
      <Setter Property="Width" Value="18"/>
      <Setter Property="Focusable" Value="False"/>
      <Setter Property="Delay" Value="400"/>
      <Setter Property="Interval" Value="90"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RepeatButton">
            <Border x:Name="Bd" Background="{TemplateBinding Background}" Padding="2,0">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#1A3A4D"/>
                <Setter Property="Foreground" Value="#E8F4F8"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#0B1728"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="CantSpinnerText" TargetType="TextBox">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="#FFFFFF"/>
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="HorizontalAlignment" Value="Left"/>
      <Setter Property="Padding" Value="4,0,1,0"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="CaretBrush" Value="#7AA3B8"/>
    </Style>
    <Style x:Key="BtnCloseX_MinimalNoBg" TargetType="Button">
      <Setter Property="Width" Value="32"/>
      <Setter Property="Height" Value="32"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderBrush" Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Root" Background="Transparent">
              <Path x:Name="XIcon"
                    Data="M 9,9 L 23,23 M 23,9 L 9,23"
                    Stroke="#E8F4F8"
                    StrokeThickness="2.0"
                    StrokeStartLineCap="Round"
                    StrokeEndLineCap="Round"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="XIcon" Property="Stroke" Value="#5BC0DE"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
"""
