# -*- coding: utf-8 -*-
"""
Ventana WPF previsualización + parámetros malla Armado Muros Lineales (pushbutton autocontenido).

- Canvas: ``System.Windows.Controls.Canvas`` (Revit/pyRevit).
- Sección dibujada: solo geometría del muro; mallas sólo en parámetros.
- Instancia única por AppDomain (norma BIMTools).
- Ventana **modal** (`ShowDialog` + owner Revit): bloquea selección y demás acciones en la UI de Revit mientras está abierta.
"""

from __future__ import print_function

import os
import sys
import weakref
import clr

_sd_boot = os.path.dirname(os.path.abspath(__file__))
if _sd_boot not in sys.path:
    sys.path.insert(0, _sd_boot)
import bootstrap_paths

bootstrap_paths.pin_local_scripts_first()

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

import System.AppDomain as _AppDom

from System.Windows.Markup import XamlReader

from Autodesk.Revit.DB import (
    ElementId,
    FilteredElementCollector,
    StorageType,
    Transaction,
    UnitUtils,
    UnitTypeId,
    Wall,
)
from Autodesk.Revit.DB.Structure import RebarBarType, AreaReinforcementLayerType
from Autodesk.Revit.UI import TaskDialog, ExternalEvent, IExternalEventHandler

import armado_muros_lineales as geo

try:
    from bimtools_element_id import wall_id_int as _wall_id_int
except Exception:
    def _wall_id_int(wall):
        try:
            return geo._wall_id_int(wall)
        except Exception:
            try:
                return geo._element_id_int(wall.Id)
            except Exception:
                return None

try:
    import armado_muros_cabezal as cabezal
    _cab_import_error = None
except Exception as _e_cab:
    cabezal = None
    _cab_import_error = _e_cab

# Constantes mínimas si el import de cabezal falla (evita crash al abrir UI).
CABEZAL_EXTREMO_INICIO = u"inicio"
CABEZAL_EXTREMO_FIN = u"fin"
if cabezal is not None:
    CABEZAL_EXTREMO_INICIO = cabezal.CABEZAL_EXTREMO_INICIO
    CABEZAL_EXTREMO_FIN = cabezal.CABEZAL_EXTREMO_FIN


def _require_cabezal_mod():
    """True si ``armado_muros_cabezal`` cargó; si no, muestra el error de import."""
    if cabezal is not None:
        return True
    detail = u""
    if _cab_import_error is not None:
        try:
            detail = unicode(_cab_import_error)
        except Exception:
            detail = str(_cab_import_error)
    TaskDialog.Show(
        u"Arainco: Armado Muros — Error",
        u"No se pudo cargar el módulo de cabezal (armado_muros_cabezal).\n\n{0}".format(
            detail or u"Error desconocido.",
        ),
    )
    return False

try:
    from bimtools_runtime import skip_area_rein_ordered_parameters_scan
except Exception:
    def skip_area_rein_ordered_parameters_scan(doc=None):
        return True


def _area_rein_parameters_iter(area_rein, doc=None):
    """Iterador de parámetros AR; omitido fuera de Revit 2024 legacy."""
    if area_rein is None or skip_area_rein_ordered_parameters_scan(doc):
        return None
    if hasattr(area_rein, u"GetOrderedParameters"):
        try:
            return area_rein.GetOrderedParameters()
        except Exception:
            return None
    if hasattr(area_rein, u"Parameters"):
        try:
            return area_rein.Parameters
        except Exception:
            return None
    return None

try:
    import armado_muros_vecinos_extremos as _vec_ext
except Exception:
    _vec_ext = None

try:
    import armado_muros_cabezal_encuentro_l as _cab_enc_l
except Exception:
    _cab_enc_l = None
    try:
        import traceback as _tb_cab
        print(u"[armado_muros_preview_ui] cabezal import failed: {}".format(_tb_cab.format_exc()))
    except Exception:
        pass

try:
    from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
except Exception:
    BIMTOOLS_DARK_STYLES_XML = u""


def _copy_cabezal_layers(layers):
    """Copia capas sin deepcopy (ElementId de Revit no soporta ElementId())."""
    out = []
    for ly in layers or []:
        if isinstance(ly, dict):
            out.append(dict(ly))
        else:
            out.append(ly)
    return out


def _copy_cabezal_segment_bar_type_ids(raw):
    if not isinstance(raw, dict):
        return {}
    out = {}
    for k, v in raw.items():
        out[k] = list(v) if isinstance(v, list) else v
    return out


def _copy_cabezal_extremo_field(key, val):
    if key == u"layers":
        return _copy_cabezal_layers(val)
    if key == u"segment_bar_type_ids":
        return _copy_cabezal_segment_bar_type_ids(val)
    if key == u"confinement" and isinstance(val, dict):
        return dict(val)
    return val


def _copy_cabezal_extremo_config(ex_cfg):
    if not ex_cfg or not isinstance(ex_cfg, dict):
        if cabezal is not None:
            return cabezal.default_cabezal_extremo_config()
        return {}
    return {
        k: _copy_cabezal_extremo_field(k, v) for k, v in ex_cfg.items()
    }


def _build_preview_xaml():
    return XAML_PREVIEW.replace(u"__BIMTOOLS_DARK_STYLES__", BIMTOOLS_DARK_STYLES_XML)


UI_MODE_CABEZAL = u"cabezal"
UI_MODE_MALLAS = u"mallas"
UI_MODE_UNIFICADO = u"unificado"

_PREVIEW_APPDOMAIN_KEYS = {
    UI_MODE_CABEZAL: u"BIMTools.ArmadoMurosCabezal.Window",
    UI_MODE_MALLAS: u"BIMTools.ArmadoMurosMallas.Window",
    UI_MODE_UNIFICADO: u"BIMTools.ArmadoMuros.Window",
}


def _preview_singleton_key(mode):
    return _PREVIEW_APPDOMAIN_KEYS.get(mode, _PREVIEW_APPDOMAIN_KEYS[UI_MODE_MALLAS])


def _unregister_preview_singleton(mode):
    try:
        _AppDom.CurrentDomain.SetData(_preview_singleton_key(mode), None)
    except Exception:
        pass


def _register_preview_singleton(win, mode):
    try:
        _AppDom.CurrentDomain.SetData(_preview_singleton_key(mode), win)
    except Exception:
        pass


XAML_PREVIEW = u"""<Window
  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  Title="Arainco: Armado Muros"
  Height="960" Width="968"
  MinHeight="720" MinWidth="780" MaxWidth="1400"
  ResizeMode="CanResize"
  WindowStartupLocation="Manual"
  Background="#071018"
  FontFamily="Segoe UI"
  FontSize="12"
  ShowInTaskbar="False">
  <Window.Resources>
__BIMTOOLS_DARK_STYLES__
  </Window.Resources>
  <Border Background="#071018" BorderBrush="#21465C" BorderThickness="1" Padding="18">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <StackPanel Grid.Row="0" Margin="0,0,0,10">
        <TextBlock x:Name="TxtTitle" Text="Arainco: Armado Muros" Foreground="#E8F4F8" FontSize="18" FontWeight="Bold"/>
        <TextBlock x:Name="TxtSubtitle" Margin="0,6,0,0" Foreground="#95B8CC" TextWrapping="Wrap"
                   Text="Asistente: previsualización de mallas por sección y tramo."/>
      </StackPanel>

      <StackPanel x:Name="PnlModoMuro" Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,8">
        <CheckBox x:Name="ChkMuroTradicional" Content="Muro Tradicional" IsChecked="True"
                  Foreground="#E8F4F8" FontSize="11" Margin="0,0,24,0" VerticalAlignment="Center"/>
        <CheckBox x:Name="ChkMuroContencion" Content="Muro de Contención"
                  Foreground="#E8F4F8" FontSize="11" VerticalAlignment="Center"/>
      </StackPanel>

      <TextBlock x:Name="TxtInfoMuros" Grid.Row="2" Foreground="#95B8CC" FontSize="11"
                 Margin="0,0,0,12" TextWrapping="Wrap"/>

      <Grid Grid.Row="3">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Border x:Name="BdrColumnHeaders" Grid.Row="0" Background="#0a1620" BorderBrush="#21465C" BorderThickness="1,1,1,0"
                CornerRadius="4,4,0,0" Padding="8,6,8,4">
          <Grid x:Name="GrdColumnHeaders" Background="Transparent"
                SnapsToDevicePixels="True" HorizontalAlignment="Center"/>
        </Border>
        <ScrollViewer x:Name="ScrMuros" Grid.Row="1" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
          <Border Background="#0a1620" BorderBrush="#21465C" BorderThickness="1,0,1,0"
                  Padding="8,4,8,12">
            <Grid x:Name="GrdListaMuros" Background="Transparent"
                  ClipToBounds="False" SnapsToDevicePixels="True" HorizontalAlignment="Center"/>
          </Border>
        </ScrollViewer>
        <Border x:Name="BdrCabezalBulkActions" Grid.Row="2" Visibility="Collapsed"
                Background="#0a1620" BorderBrush="#21465C" BorderThickness="1,0,1,1"
                CornerRadius="0,0,4,4" Padding="8,8,8,12">
          <Grid x:Name="GrdCabezalBulkActions" Background="Transparent"
                SnapsToDevicePixels="True" HorizontalAlignment="Center"/>
        </Border>
      </Grid>

      <TextBlock x:Name="TxtFooterHint" Grid.Row="4" Foreground="#64748b" FontSize="10" TextWrapping="Wrap" Margin="0,8,0,0"
                 Text="Creación: cabezal (verticales, ini/fin) + malla AR → Remove System. Orden inferior→superior."/>

      <Grid Grid.Row="5" Margin="0,14,0,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="TxtEstado" Grid.Column="0" VerticalAlignment="Center"
                   Foreground="#64748b" FontSize="10" TextWrapping="Wrap" Margin="0,0,12,0"/>
        <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
          <Button x:Name="BtnAplicarCabezal" Content="Cabezal → todos"
                  Style="{StaticResource BtnSelectOutline}" MinWidth="118" Margin="0,0,8,0"/>
          <Button x:Name="BtnAplicarMallas" Content="Mallas → todos"
                  Style="{StaticResource BtnSelectOutline}" MinWidth="118" Margin="0,0,10,0"/>
          <Button x:Name="BtnCancelar" Content="Cancelar"
                  Style="{StaticResource BtnSelectOutline}" MinWidth="110" Margin="0,0,10,0"/>
          <Button x:Name="BtnCrear" Content="Crear Area Reinf."
                  Style="{StaticResource BtnPrimary}" MinWidth="180"/>
        </StackPanel>
      </Grid>
    </Grid>
  </Border>
</Window>"""


def _snap_cabezal_window_to_screen_left(win, uiapp=None):
    """Cabezal: snap real Win+← (SetWindowPos en px de pantalla)."""
    if win is None:
        return
    try:
        from revit_wpf_window_position import (
            revit_main_hwnd,
            snap_wpf_window_left_half,
        )

        hwnd = revit_main_hwnd(uiapp)
        if not snap_wpf_window_left_half(win, hwnd, half_width=True):
            _posicion_centro_horizontal(win, align_left=True)
    except Exception:
        _posicion_centro_horizontal(win, align_left=True)


def _resolve_uiapp_for_position(uidoc=None, revit=None):
    uiapp = None
    try:
        uiapp = uidoc.Application if uidoc else None
    except Exception:
        uiapp = None
    if uiapp is None and revit is not None:
        try:
            uiapp = getattr(revit, "Application", None) or revit
        except Exception:
            uiapp = revit
    return uiapp


def _position_preview_window(win, uidoc, revit, uses_cabezal_panels, before_snap=None):
    u"""Monitor secundario → maximizado; si no, snap izquierda (cabezal) o centro (mallas)."""
    if win is None:
        return False
    uiapp = _resolve_uiapp_for_position(uidoc, revit)
    hwnd = None
    try:
        from revit_wpf_window_position import (
            bind_maximize_wpf_on_secondary_monitor,
            bind_snap_wpf_window_left_half,
            revit_main_hwnd,
        )

        hwnd = revit_main_hwnd(uiapp)
        if bind_maximize_wpf_on_secondary_monitor(win, hwnd):
            if before_snap is not None:
                try:
                    before_snap()
                except Exception:
                    pass
            return True
        if uses_cabezal_panels:
            bind_snap_wpf_window_left_half(
                win,
                hwnd,
                half_width=True,
                before_snap=before_snap,
            )
        else:
            _posicion_centro_horizontal(win, align_left=False)
            if before_snap is not None:
                try:
                    before_snap()
                except Exception:
                    pass
    except Exception:
        try:
            if uses_cabezal_panels:
                if before_snap is not None:
                    try:
                        before_snap()
                    except Exception:
                        pass
                _snap_cabezal_window_to_screen_left(win, uiapp=uiapp)
            else:
                _posicion_centro_horizontal(win, align_left=False)
                if before_snap is not None:
                    try:
                        before_snap()
                    except Exception:
                        pass
        except Exception:
            pass
    return False


def _posicion_centro_horizontal(win, align_left=False):
    try:
        from System.Windows import SystemParameters, WindowStartupLocation

        wa = SystemParameters.WorkArea
        win.WindowStartupLocation = WindowStartupLocation.Manual
        wa_left = float(wa.Left)
        wa_top = float(wa.Top)
        wa_width = float(wa.Width)
        wa_height = float(wa.Height)
        if align_left:
            win.Left = wa_left
            win.Top = wa_top
            win.Height = wa_height
            snap_w = max(320.0, wa_width * 0.5)
            win.Width = snap_w
            try:
                if float(win.MinWidth) > snap_w:
                    win.MinWidth = snap_w
            except Exception:
                win.MinWidth = snap_w
        else:
            w = float(win.Width)
            if w <= 1.0:
                try:
                    if float(win.ActualWidth) > 1.0:
                        w = float(win.ActualWidth)
                except Exception:
                    pass
            left = wa_left + max(0.0, (wa_width - w) / 2.0)
            win.Left = left
            win.Top = wa_top + 48.0
    except Exception:
        pass


def _get_bar_types_sorted_display(document):
    out = []

    try:
        rts = list(FilteredElementCollector(document).OfClass(RebarBarType))
    except Exception:
        return []
    keyed = []
    for bt in rts:
        try:
            keyed.append((float(bt.BarNominalDiameter), bt))
        except Exception:
            keyed.append((0.0, bt))
    keyed.sort()
    for _d, bt in keyed:
        try:
            eid = geo._element_id_int(bt.Id)
            diam_ft = getattr(bt, "BarNominalDiameter", None)
            disp = u"ø {} mm".format(eid)
            try:
                if diam_ft is not None:
                    dmm = int(round(float(diam_ft) * 304.8))
                    if dmm > 0:
                        disp = u"\u00f8{} mm".format(dmm)
            except Exception:
                pass
            out.append((disp, bt))
        except Exception:
            continue
    return out


def _cabezal_compact_diam_label(label):
    """Etiqueta mínima para combo ø en cabezal (cabecera ya muestra «ø»)."""
    mm = _parse_diam_label_mm(label)
    if mm is not None:
        return str(mm)
    try:
        return unicode(label).strip()
    except Exception:
        return str(label or u"").strip()


def _parse_diam_label_mm(label):
    if not label:
        return None
    try:
        import re
        m = re.search(r"(\d+)", unicode(label))
        if m:
            return int(m.group(1))
    except Exception:
        pass
    return None


def _cabezal_diam_label_for_mm(diam_strings, mm):
    """Etiqueta de combo (p. ej. ``ø12 mm``) más cercana a ``mm``."""
    if not diam_strings:
        return None
    try:
        target = int(round(float(mm)))
    except Exception:
        target = 12
    for lab in diam_strings:
        if _parse_diam_label_mm(lab) == target:
            return lab
    best_lab = None
    best_diff = None
    for lab in diam_strings:
        d = _parse_diam_label_mm(lab)
        if d is None:
            continue
        diff = abs(d - target)
        if best_diff is None or diff < best_diff:
            best_diff = diff
            best_lab = lab
    return best_lab or diam_strings[0]


def _cabezal_default_diam_index(diam_strings, mm=None):
    if mm is None:
        mm = (
            cabezal.CABEZAL_DEFAULT_BAR_DIAM_MM
            if cabezal is not None
            else 12.0
        )
    lab = _cabezal_diam_label_for_mm(diam_strings, mm)
    if lab is None:
        return 0
    try:
        return diam_strings.index(lab)
    except ValueError:
        return 0


def _spacing_internal_mm(val_text):
    try:
        v = float(str(val_text).strip().replace(",", "."))
        return UnitUtils.ConvertToInternalUnits(v, UnitTypeId.Millimeters)
    except Exception:
        return UnitUtils.ConvertToInternalUnits(150.0, UnitTypeId.Millimeters)


def _capas_verticales_muro_keys(muro_contencion=False):
    u"""Capas con barras verticales: minor en muro tradicional, major en contención."""
    if muro_contencion:
        return (u"exterior_major", u"interior_major")
    return (u"exterior_minor", u"interior_minor")


def _capas_horizontales_muro_keys(muro_contencion=False):
    u"""Capas con barras horizontales: major en muro tradicional, minor en contención."""
    if muro_contencion:
        return (u"exterior_minor", u"interior_minor")
    return (u"exterior_major", u"interior_major")


def _param_coincide_capa_area(def_name, layer_key):
    dn = (def_name or u"").lower()
    ext_int = u"exterior" if u"exterior" in layer_key else u"interior"
    maj_min = u"major" if u"major" in layer_key else u"minor"
    top_bot = u"top" if u"exterior" in layer_key else u"bottom"
    tiene_cara = ext_int in dn or top_bot in dn
    tiene_capa = maj_min in dn or (u"dir 1" in dn and maj_min == u"major") or (
        u"dir 2" in dn and maj_min == u"minor"
    )
    return tiene_cara and tiene_capa


_LAYER_KEY_TO_AR_TYPE = {
    u"exterior_major": AreaReinforcementLayerType.TopOrFrontMajor,
    u"exterior_minor": AreaReinforcementLayerType.TopOrFrontMinor,
    u"interior_major": AreaReinforcementLayerType.BottomOrBackMajor,
    u"interior_minor": AreaReinforcementLayerType.BottomOrBackMinor,
}


def _aplicar_remove_verticales_por_cabezal(
    area_rein,
    layer_active_dict,
    ex_cfg_inicio=None,
    ex_cfg_fin=None,
    muro_contencion=False,
    doc=None,
):
    u"""
    Respaldo AR por parámetro Remove First/Last (mínimo 1+1).

    La correlación cabezal ↔ ``n_capas`` (capa k → barra k) se aplica tras
    Remove System en rebars ``exterior_minor`` / ``interior_minor``.
    """
    if area_rein is None:
        return

    _aplicar_remove_verticales_por_cabezal_por_parametro(
        area_rein,
        layer_active_dict,
        ex_cfg_inicio,
        ex_cfg_fin,
        muro_contencion,
        doc=doc,
    )


def _aplicar_remove_verticales_por_cabezal_por_parametro(
    area_rein,
    layer_active_dict,
    ex_cfg_inicio=None,
    ex_cfg_fin=None,
    muro_contencion=False,
    doc=None,
):
    u"""Respaldo AR: parámetros Remove First/Last según ``n_capas`` por extremo."""
    if area_rein is None:
        return
    n_remove_ini = 1
    n_remove_fin = 1
    if cabezal is not None:
        try:
            n_remove_ini = cabezal.malla_n_remove_por_extremo(ex_cfg_inicio)
            n_remove_fin = cabezal.malla_n_remove_por_extremo(ex_cfg_fin)
        except Exception:
            pass
    capas_v = _capas_verticales_muro_keys(muro_contencion)
    params_iter = _area_rein_parameters_iter(area_rein, doc)
    if not params_iter:
        return

    for layer_key in capas_v:
        if not layer_active_dict.get(layer_key, True):
            continue
        for param in params_iter:
            if param is None or param.IsReadOnly:
                continue
            try:
                def_name = param.Definition.Name or u""
                dn = def_name.lower()
                if not _param_coincide_capa_area(def_name, layer_key):
                    continue
                es_first = (
                    (u"remove" in dn and u"first" in dn)
                    or (u"eliminar" in dn and u"primera" in dn)
                    or (u"eliminar" in dn and u"first" in dn)
                )
                es_last = (
                    (u"remove" in dn and u"last" in dn)
                    or (u"eliminar" in dn and (u"última" in dn or u"ultima" in dn))
                    or (u"eliminar" in dn and u"last" in dn)
                )
                if not es_first and not es_last:
                    continue
                n_remove = n_remove_fin if es_last else n_remove_ini
                try:
                    n_max = cabezal.CABEZAL_MAX_CAPAS if cabezal is not None else 6
                    n_remove = max(1, min(int(n_remove), int(n_max)))
                except Exception:
                    n_remove = 1
                if param.StorageType == StorageType.Integer:
                    param.Set(int(n_remove))
                elif param.StorageType == StorageType.Double:
                    param.Set(float(n_remove))
                else:
                    try:
                        param.Set(int(n_remove))
                    except Exception:
                        pass
            except Exception:
                continue


def _aplicar_remove_horizontales_ultima_barra(
    area_rein,
    layer_active_dict,
    muro_contencion=False,
    doc=None,
):
    u"""Respaldo AR: Remove Last = 1 en capas horizontales exterior e interior."""
    if area_rein is None:
        return
    capas_h = _capas_horizontales_muro_keys(muro_contencion)
    params_iter = _area_rein_parameters_iter(area_rein, doc)
    if not params_iter:
        return

    for layer_key in capas_h:
        if not layer_active_dict.get(layer_key, True):
            continue
        for param in params_iter:
            if param is None or param.IsReadOnly:
                continue
            try:
                def_name = param.Definition.Name or u""
                dn = def_name.lower()
                if not _param_coincide_capa_area(def_name, layer_key):
                    continue
                es_last = (
                    (u"remove" in dn and u"last" in dn)
                    or (u"eliminar" in dn and (u"última" in dn or u"ultima" in dn))
                    or (u"eliminar" in dn and u"last" in dn)
                )
                if not es_last:
                    continue
                if param.StorageType == StorageType.Integer:
                    param.Set(1)
                elif param.StorageType == StorageType.Double:
                    param.Set(1.0)
                else:
                    try:
                        param.Set(1)
                    except Exception:
                        pass
            except Exception:
                continue


def _aplicar_parametros_malla(
    area_rein,
    params_dict,
    layer_active_dict,
    muro_contencion=False,
    ex_cfg_inicio=None,
    ex_cfg_fin=None,
    doc=None,
):

    if not area_rein:
        return
    layer_config = [
        ("exterior_major", [u"Exterior Major Spacing"],
         [u"Exterior Major Bar Type", u"Exterior Major Rebar Type"],
         AreaReinforcementLayerType.TopOrFrontMajor),
        ("exterior_minor", [u"Exterior Minor Spacing"],
         [u"Exterior Minor Bar Type", u"Exterior Minor Rebar Type"],
         AreaReinforcementLayerType.TopOrFrontMinor),
        ("interior_major", [u"Interior Major Spacing"],
         [u"Interior Major Bar Type", u"Interior Major Rebar Type"],
         AreaReinforcementLayerType.BottomOrBackMajor),
        ("interior_minor", [u"Interior Minor Spacing"],
         [u"Interior Minor Bar Type", u"Interior Minor Rebar Type"],
         AreaReinforcementLayerType.BottomOrBackMinor),
    ]
    dir_param_names = {
        "exterior_major": [
            u"Exterior Major Direction", u"Top Major Direction", u"Top Mayor Direction",
        ],
        "exterior_minor": [u"Exterior Minor Direction", u"Top Minor Direction"],
        "interior_major": [
            u"Interior Major Direction", u"Bottom Major Direction", u"Bottom Mayor Direction",
        ],
        "interior_minor": [
            u"Interior Minor Direction", u"Bottom Minor Direction",
        ],
    }
    for layer_key, spacing_names, bar_names, layer_type in layer_config:
        bar_type_id, spacing_mm = params_dict.get(layer_key, (None, "150"))
        is_active = layer_active_dict.get(layer_key, True)
        try:
            area_rein.SetLayerActive(layer_type, bool(is_active))
        except Exception:
            pass
        for pname in dir_param_names.get(layer_key, []):
            try:
                p = area_rein.LookupParameter(pname)
                if p and not p.IsReadOnly:
                    p.Set(1 if is_active else 0)
                    break
            except Exception:
                continue
        sp_int = _spacing_internal_mm(spacing_mm)
        for name in spacing_names:
            try:
                p = area_rein.LookupParameter(name)
                if p and not p.IsReadOnly:
                    p.Set(sp_int)
            except Exception:
                pass
        for name in bar_names:
            try:
                p = area_rein.LookupParameter(name)
                if p and not p.IsReadOnly and bar_type_id and bar_type_id != ElementId.InvalidElementId:
                    p.Set(bar_type_id)
            except Exception:
                pass
    params_iter = _area_rein_parameters_iter(area_rein, doc)
    if params_iter:
        for param in params_iter:
            if param is None or param.IsReadOnly:
                continue
            try:
                def_name = (param.Definition.Name or "").lower()
                for layer_key, _, _, _ in layer_config:
                    bar_type_id, spacing_mm = params_dict.get(layer_key, (None, "150"))
                    is_active = layer_active_dict.get(layer_key, True)
                    ext_int = "exterior" if "exterior" in layer_key else "interior"
                    top_bot = "top" if "exterior" in layer_key else "bottom"
                    maj_min = "major" if "major" in layer_key else "minor"
                    matches_layer = (ext_int in def_name or top_bot in def_name) and maj_min in def_name
                    if matches_layer:
                        try:
                            if "spacing" in def_name:
                                param.Set(_spacing_internal_mm(spacing_mm))
                            elif (
                                    ("bar" in def_name or "rebar" in def_name)
                                    and "type" in def_name
                                    and bar_type_id and bar_type_id != ElementId.InvalidElementId):
                                param.Set(bar_type_id)
                            elif "direction" in def_name:
                                param.Set(1 if is_active else 0)
                        except Exception:
                            pass
                        break
            except Exception:
                continue

    _aplicar_remove_horizontales_ultima_barra(
        area_rein,
        layer_active_dict,
        muro_contencion=bool(muro_contencion),
        doc=doc,
    )
    _aplicar_remove_verticales_por_cabezal(
        area_rein,
        layer_active_dict,
        ex_cfg_inicio=ex_cfg_inicio,
        ex_cfg_fin=ex_cfg_fin,
        muro_contencion=bool(muro_contencion),
        doc=doc,
    )


def _pop_legacy_extremo_marker_ids():
    """Recupera IDs de marcadores 3D huérfanos (sesiones anteriores) y limpia AppDomain."""
    key = u"__arainco_pending_marker_ids__"
    try:
        ids = _AppDom.CurrentDomain.GetData(key) or []
        _AppDom.CurrentDomain.SetData(key, [])
        return list(ids)
    except Exception:
        return []


def _mostrar_resumen_fin_armado_muros(msg_ok, titulo=u"Arainco: Armado Muros — resumen"):
    """Resumen de ejecución (sin TaskDialog; el estado queda en la ventana/handler)."""
    return


def _append_linea_etiquetas_malla_resumen(msg_ok, embed_res, n_muros=0):
    embed_res = embed_res or {}
    n_ok = int(embed_res.get(u"n_tags_rebar_malla", 0) or 0)
    n_skip = int(embed_res.get(u"n_tags_rebar_malla_skip", 0) or 0)
    n_fail = int(embed_res.get(u"n_tags_rebar_malla_fail", 0) or 0)
    msg_ok += (
        u"\nEtiquetas malla multihost (interior + exterior): "
        u"{0} ok, {1} omitidas, {2} fallo.".format(n_ok, n_skip, n_fail)
    )
    n_m = max(0, int(n_muros or 0))
    if n_m > 0:
        esperadas = n_m * 2
        if n_ok + n_fail + n_skip == 0:
            msg_ok += (
                u"\n  Ninguna etiqueta creada: use vista planta, alzado o sección "
                u"(no plantilla ni 3D)."
            )
        elif n_ok < esperadas:
            msg_ok += (
                u"\n  Con doble malla se esperan hasta {0} etiquetas "
                u"({1} muro(s) × vertical + horizontal).".format(esperadas, n_m)
            )
    return msg_ok


class _CrearCabezalEjecutarHandler(IExternalEventHandler):

    def __init__(self, window_ref):
        self._window_ref = window_ref
        self.walls = []
        self.cabezal_por_muro_id = {}

    def Execute(self, uiapp):
        from Autodesk.Revit.UI import TaskDialog
        wrap = None
        try:
            wrap = self._window_ref()
        except Exception:
            wrap = None
        if wrap is None or cabezal is None:
            return

        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            TaskDialog.Show(u"Armado muros", u"No hay documento activo.")
            return

        doc = uidoc.Document

        fallback_bt = cabezal.cabezal_resolve_bar_type_fallback(
            doc, self.cabezal_por_muro_id, self.walls,
        )

        try:
            import armado_muros_coronamiento as _cor_mod

            cor_res = _cor_mod.aplicar_coronamiento_muros(
                doc, self.walls, bar_type_fallback=fallback_bt,
            )
            cor_res = _cor_mod.aplicar_etiquetado_coronamiento(
                doc, cor_res, uidoc=uidoc,
            )
            if int(cor_res.get(u"n_fail", 0)):
                pass  # mensaje al final vía cab_res / TaskDialog extendido
            self._coronamiento_res = cor_res
        except Exception as ex_cor:
            self._coronamiento_res = {
                u"n_fail": 1,
                u"messages": [unicode(ex_cor)],
            }

        ref_walls_troceo = getattr(self, u"ref_walls_troceo", None)

        cab_res = cabezal.aplicar_cabezales_muros(
            doc,
            self.walls,
            self.cabezal_por_muro_id,
            bar_type_fallback=fallback_bt,
            ref_walls_troceo=ref_walls_troceo,
            uidoc=uidoc,
        )
        try:
            geo.aplicar_unobscured_armado_muros_en_vista(
                doc,
                uidoc,
                cab_res=cab_res,
                cor_res=cor_res,
                errores=cab_res.get(u"messages"),
            )
        except Exception:
            pass
        msg_ok = u"Cabezal — capas creadas: {0}, barras (total): {1}, error: {2}.".format(
            int(cab_res.get(u"n_created", 0)),
            int(cab_res.get(u"n_bars_total", 0)),
            int(cab_res.get(u"n_fail", 0)),
        )
        cor_res = getattr(self, u"_coronamiento_res", None) or {}
        if int(cor_res.get(u"n_created", 0)) or int(cor_res.get(u"n_fail", 0)):
            cor_line = u"Coronamiento (tope stack): "
            if int(cor_res.get(u"n_created", 0)):
                cor_line += u"{0}Ø{1} mm, {2} barra(s)".format(
                    int(cor_res.get(u"n_bars_spec", 0) or 0),
                    int(cor_res.get(u"diam_mm", 0) or 0),
                    int(cor_res.get(u"n_bars", 0) or 0),
                )
            else:
                cor_line += u"error"
            msg_ok = cor_line + u"\n" + msg_ok
        if int(cor_res.get(u"n_inferior_created", 0)) or int(cor_res.get(u"n_inferior_fail", 0)):
            cor_inf = u"Coronamiento inf. (fundación): sets={0}, barras={1}, error={2}.".format(
                int(cor_res.get(u"n_inferior_created", 0)),
                int(cor_res.get(u"n_inferior_bars", 0)),
                int(cor_res.get(u"n_inferior_fail", 0)),
            )
            msg_ok = cor_inf + u"\n" + msg_ok
        if int(cor_res.get(u"n_inferior_pie_created", 0)) or int(cor_res.get(u"n_inferior_pie_fail", 0)):
            cor_pie = u"Coronamiento pie (sin apil./fund.): sets={0}, barras={1}, error={2}.".format(
                int(cor_res.get(u"n_inferior_pie_created", 0)),
                int(cor_res.get(u"n_inferior_pie_bars", 0)),
                int(cor_res.get(u"n_inferior_pie_fail", 0)),
            )
            msg_ok = cor_pie + u"\n" + msg_ok
        if int(cor_res.get(u"n_voladizo_created", 0)) or int(cor_res.get(u"n_voladizo_fail", 0)):
            cor_vol = u"Coronamiento voladizo (reentrada): sets={0}, barras={1}, error={2}.".format(
                int(cor_res.get(u"n_voladizo_created", 0)),
                int(cor_res.get(u"n_voladizo_bars", 0)),
                int(cor_res.get(u"n_voladizo_fail", 0)),
            )
            msg_ok = cor_vol + u"\n" + msg_ok
        n_cor_tags = int(cor_res.get(u"n_cor_tags_created", 0))
        n_cor_tags_fail = int(cor_res.get(u"n_cor_tags_fail", 0))
        if n_cor_tags or n_cor_tags_fail:
            msg_ok = u"Etiquetas coronamiento: {0} ok, {1} fallo.\n".format(
                n_cor_tags, n_cor_tags_fail,
            ) + msg_ok
        n_embed_top = int(cab_res.get(u"n_kept_embed_top", 0))
        n_revert_top = int(cab_res.get(u"n_reverted_embed_top", 0))
        n_fund = int(cab_res.get(u"n_foundation_stretch", 0))
        n_troceo = int(cab_res.get(u"n_troceo_segments", 0))
        n_pata_t = int(cab_res.get(u"n_pata_top", 0))
        n_pata_b = int(cab_res.get(u"n_pata_bot", 0))
        n_conf = int(cab_res.get(u"n_confinement_created", 0))
        n_conf_tags = int(cab_res.get(u"n_conf_tags_created", 0))
        n_conf_tags_fail = int(cab_res.get(u"n_conf_tags_fail", 0))
        n_tags = int(cab_res.get(u"n_tags_created", 0))
        n_tags_fail = int(cab_res.get(u"n_tags_fail", 0))
        n_emp_ok = int(cab_res.get(u"n_empalme_markers_ok", 0))
        n_emp_fail = int(cab_res.get(u"n_empalme_markers_fail", 0))
        n_emp_dim_ok = int(cab_res.get(u"n_empalme_dims_ok", 0))
        n_emp_dim_fail = int(cab_res.get(u"n_empalme_dims_fail", 0))
        detail_parts = []
        if n_embed_top or n_revert_top:
            detail_parts.append(
                u"Embed +Z: mantenidos={0}, revertidos={1}".format(n_embed_top, n_revert_top),
            )
        if n_fund:
            detail_parts.append(u"Fundación -Z: {0}".format(n_fund))
        if n_troceo:
            detail_parts.append(u"Segmentos troceo: {0}".format(n_troceo))
        if n_pata_t or n_pata_b:
            detail_parts.append(
                u"Pata-L: sup={0}, inf={1}".format(n_pata_t, n_pata_b),
            )
        if n_conf:
            detail_parts.append(u"Confinamiento (estribos/trabas): {0}".format(n_conf))
        if n_conf_tags or n_conf_tags_fail:
            detail_parts.append(
                u"Etiquetas estribo: {0} ok, {1} fallo".format(
                    n_conf_tags, n_conf_tags_fail,
                ),
            )
        if n_tags or n_tags_fail:
            detail_parts.append(
                u"Etiquetas longitudinales: {0} ok, {1} omitidas/fallo".format(
                    n_tags, n_tags_fail,
                ),
            )
        if n_emp_ok or n_emp_fail:
            detail_parts.append(
                u"Marcadores empalme (detail): {0} ok, {1} fallo".format(
                    n_emp_ok, n_emp_fail,
                ),
            )
        if n_emp_dim_ok or n_emp_dim_fail:
            detail_parts.append(
                u"Cotas empalme: {0} ok, {1} fallo".format(
                    n_emp_dim_ok, n_emp_dim_fail,
                ),
            )
        if detail_parts:
            msg_ok += u"\n" + u"; ".join(detail_parts) + u"."
        msgs = list(cab_res.get(u"messages") or [])
        if msgs:
            msg_ok += u"\n\n" + u"\n".join(unicode(m) for m in msgs[:10])
            if len(msgs) > 10:
                msg_ok += u"\n…"
        _mostrar_resumen_fin_armado_muros(
            msg_ok, titulo=u"Arainco: Armado Muros — cabezal",
        )
        if wrap is not None:
            try:
                wrap._set_estado(msg_ok.replace(u"\n", u" "))
            except Exception:
                pass
            try:
                wrap._close_after_create()
            except Exception:
                pass

    def GetName(self):
        return u"AraincoCreacionMurosLinealesCabezal"


class _CrearMallasEjecutarHandler(IExternalEventHandler):

    def __init__(self, window_ref):
        self._window_ref = window_ref
        self.walls = []
        self.params_por_muro_id = {}
        self.cabezal_por_muro_id = {}
        self.area_reinforcement_type_id = ElementId.InvalidElementId
        self.muro_contencion = False

    def Execute(self, uiapp):
        from Autodesk.Revit.UI import TaskDialog

        wrap = None
        try:
            wrap = self._window_ref()
        except Exception:
            wrap = None

        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            TaskDialog.Show("Armado muros", u"No hay documento activo.")
            return

        doc = uidoc.Document
        n_muros = len(getattr(self, "walls", None) or [])

        ok, err, cover_n, embed_res = geo.crear_areas_malla_parametrizada(
            doc,
            self.walls,
            self.params_por_muro_id,
            self.area_reinforcement_type_id,
            _aplicar_parametros_malla,
            muro_contencion=bool(getattr(self, "muro_contencion", False)),
            uidoc=uidoc,
            cabezal_por_muro_id=getattr(self, "cabezal_por_muro_id", None),
            malla_activo_por_muro_id=getattr(
                self, u"malla_activo_por_muro_id", None,
            ),
        )
        try:
            geo.aplicar_unobscured_armado_muros_en_vista(
                doc,
                uidoc,
                embed_resumen=embed_res,
                rebars_malla_por_muro_id=(
                    (embed_res or {}).get(u"rebars_por_muro_id")
                ),
                errores=err,
            )
        except Exception:
            pass
        msg_ok = u"Rebar Cover ext/int {:.0f} mm, otras {:.0f} mm: {} muro(s).\nRebars creados: {} (IDs: {}).".format(
            geo.REBAR_COVER_MM_CARAS_EXT_INT,
            geo.REBAR_COVER_MM_OTRAS_CARAS,
            int(cover_n),
            len(ok),
            u",".join(str(x) for x in ok[:12]) if ok else u"—",
        )
        msg_ok = _append_linea_etiquetas_malla_resumen(msg_ok, embed_res, n_muros)
        if embed_res:
            msg_ok += u"\n\nVerticales ext/int (cabeza L=tabla; fund. pie: colisión→estira, no colisión→25+Ø/2):\n"
            msg_ok += u"  Estiradas={0}, retraídas={1}, pata L ext.={2}, fund. pie={3}, fund. retraídas={4}, pata L fund.={5}, omitidas={6}, error={7}.".format(
                int(embed_res.get(u"n_extended", 0)),
                int(embed_res.get(u"n_retracted", 0)),
                int(embed_res.get(u"n_pata_l", 0)),
                int(embed_res.get(u"n_fundacion_pie", 0)),
                int(embed_res.get(u"n_fundacion_retract", 0)),
                int(embed_res.get(u"n_pata_l_fund_pie", 0)),
                int(embed_res.get(u"n_skip", 0)),
                int(embed_res.get(u"n_fail", 0)),
            )
            msg_ok += u"\nHorizontales ext/int: retraída inicio 25+Ø/2, fin 25 mm → total={0} (ext.={1}, int.={2}); forma 06 → total={3} (ext.={4}, int.={5}).".format(
                int(embed_res.get(u"n_horiz_retract", 0)),
                int(embed_res.get(u"n_horiz_retract_ext", 0)),
                int(embed_res.get(u"n_horiz_retract_int", 0)),
                int(embed_res.get(u"n_horiz_pata_l", 0)),
                int(embed_res.get(u"n_horiz_pata_l_ext", 0)),
                int(embed_res.get(u"n_horiz_pata_l_int", 0)),
            )
            if int(embed_res.get(u"n_cabezal", 0)) or int(embed_res.get(u"n_cabezal_fail", 0)):
                msg_ok += u"\nCabezal (verticales, altura muro): creadas={0}, error={1}.".format(
                    int(embed_res.get(u"n_cabezal", 0)),
                    int(embed_res.get(u"n_cabezal_fail", 0)),
                )
            if int(embed_res.get(u"n_coronamiento", 0)) or int(embed_res.get(u"n_coronamiento_fail", 0)):
                msg_ok += u"\nCoronamiento sup. (tope stack): sets={0}, barras={1}, error={2}.".format(
                    int(embed_res.get(u"n_coronamiento", 0)),
                    int(embed_res.get(u"n_coronamiento_bars", 0)),
                    int(embed_res.get(u"n_coronamiento_fail", 0)),
                )
            if int(embed_res.get(u"n_coronamiento_inferior", 0)) or int(embed_res.get(u"n_coronamiento_inferior_fail", 0)):
                msg_ok += u"\nCoronamiento inf. (fundación): sets={0}, barras={1}, error={2}.".format(
                    int(embed_res.get(u"n_coronamiento_inferior", 0)),
                    int(embed_res.get(u"n_coronamiento_inferior_bars", 0)),
                    int(embed_res.get(u"n_coronamiento_inferior_fail", 0)),
                )
            if int(embed_res.get(u"n_coronamiento_inferior_pie", 0)) or int(embed_res.get(u"n_coronamiento_inferior_pie_fail", 0)):
                msg_ok += u"\nCoronamiento pie (sin apil./fund.): sets={0}, barras={1}, error={2}.".format(
                    int(embed_res.get(u"n_coronamiento_inferior_pie", 0)),
                    int(embed_res.get(u"n_coronamiento_inferior_pie_bars", 0)),
                    int(embed_res.get(u"n_coronamiento_inferior_pie_fail", 0)),
                )
            if int(embed_res.get(u"n_coronamiento_voladizo", 0)) or int(embed_res.get(u"n_coronamiento_voladizo_fail", 0)):
                msg_ok += u"\nCoronamiento voladizo (reentrada): sets={0}, barras={1}, error={2}.".format(
                    int(embed_res.get(u"n_coronamiento_voladizo", 0)),
                    int(embed_res.get(u"n_coronamiento_voladizo_bars", 0)),
                    int(embed_res.get(u"n_coronamiento_voladizo_fail", 0)),
                )
            if int(embed_res.get(u"n_coronamiento_tags", 0)) or int(embed_res.get(u"n_coronamiento_tags_fail", 0)):
                msg_ok += u"\nEtiquetas coronamiento: {0} ok, {1} fallo.".format(
                    int(embed_res.get(u"n_coronamiento_tags", 0)),
                    int(embed_res.get(u"n_coronamiento_tags_fail", 0)),
                )
            embed_msgs = embed_res.get(u"messages") or []
            if embed_msgs:
                msg_ok += u"\n" + u"\n".join(unicode(m) for m in embed_msgs[:5])
                if len(embed_msgs) > 5:
                    msg_ok += u"\n…"
        if ok and len(ok) > 12:
            msg_ok += u"…"
        if err:
            msg_ok += u"\n\n" + u"\n".join(err[:12])
            if len(err) > 12:
                msg_ok += u"\n…"
        _mostrar_resumen_fin_armado_muros(msg_ok)
        if wrap is not None:
            try:
                wrap._set_estado(msg_ok.replace(u"\n", u" ")[:500])
            except Exception:
                pass
            try:
                wrap._close_after_create()
            except Exception:
                pass

    def GetName(self):
        return "AraincoCreacionMurosLinealesAR"


class _CrearUnificadoEjecutarHandler(IExternalEventHandler):

    def __init__(self, window_ref):
        self._window_ref = window_ref
        self.walls = []
        self.params_por_muro_id = {}
        self.cabezal_por_muro_id = {}
        self.area_reinforcement_type_id = ElementId.InvalidElementId

    def Execute(self, uiapp):
        from Autodesk.Revit.UI import TaskDialog

        wrap = None
        try:
            wrap = self._window_ref()
        except Exception:
            wrap = None

        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            TaskDialog.Show(u"Arainco: Armado Muros", u"No hay documento activo.")
            return

        doc = uidoc.Document
        n_muros = len(getattr(self, "walls", None) or [])

        ok, err, cover_n, embed_res = geo.crear_armado_muros_unificado(
            doc,
            self.walls,
            self.params_por_muro_id,
            self.area_reinforcement_type_id,
            _aplicar_parametros_malla,
            self.cabezal_por_muro_id,
            uidoc=uidoc,
            malla_activo_por_muro_id=getattr(
                self, u"malla_activo_por_muro_id", None,
            ),
        )
        msg_ok = (
            u"Armado Muros — orden: longitudinales, confinamiento, mallas; "
            u"luego etiquetas (long., conf., malla).\n"
            u"Rebar Cover: {0} muro(s). Rebars AR: {1}.".format(
                int(cover_n),
                len(ok),
            )
        )
        msg_ok = _append_linea_etiquetas_malla_resumen(msg_ok, embed_res, n_muros)
        if embed_res:
            if int(embed_res.get(u"n_coronamiento", 0)) or int(embed_res.get(u"n_coronamiento_fail", 0)):
                msg_ok += u"\nCoronamiento sup. (tope stack): sets={0}, barras={1}, error={2}.".format(
                    int(embed_res.get(u"n_coronamiento", 0)),
                    int(embed_res.get(u"n_coronamiento_bars", 0)),
                    int(embed_res.get(u"n_coronamiento_fail", 0)),
                )
            if int(embed_res.get(u"n_coronamiento_inferior", 0)) or int(embed_res.get(u"n_coronamiento_inferior_fail", 0)):
                msg_ok += u"\nCoronamiento inf. (fundación): sets={0}, barras={1}, error={2}.".format(
                    int(embed_res.get(u"n_coronamiento_inferior", 0)),
                    int(embed_res.get(u"n_coronamiento_inferior_bars", 0)),
                    int(embed_res.get(u"n_coronamiento_inferior_fail", 0)),
                )
            if int(embed_res.get(u"n_coronamiento_inferior_pie", 0)) or int(embed_res.get(u"n_coronamiento_inferior_pie_fail", 0)):
                msg_ok += u"\nCoronamiento pie (sin apil./fund.): sets={0}, barras={1}, error={2}.".format(
                    int(embed_res.get(u"n_coronamiento_inferior_pie", 0)),
                    int(embed_res.get(u"n_coronamiento_inferior_pie_bars", 0)),
                    int(embed_res.get(u"n_coronamiento_inferior_pie_fail", 0)),
                )
            if int(embed_res.get(u"n_coronamiento_voladizo", 0)) or int(embed_res.get(u"n_coronamiento_voladizo_fail", 0)):
                msg_ok += u"\nCoronamiento voladizo (reentrada): sets={0}, barras={1}, error={2}.".format(
                    int(embed_res.get(u"n_coronamiento_voladizo", 0)),
                    int(embed_res.get(u"n_coronamiento_voladizo_bars", 0)),
                    int(embed_res.get(u"n_coronamiento_voladizo_fail", 0)),
                )
            if int(embed_res.get(u"n_coronamiento_tags", 0)) or int(embed_res.get(u"n_coronamiento_tags_fail", 0)):
                msg_ok += u"\nEtiquetas coronamiento: {0} ok, {1} fallo.".format(
                    int(embed_res.get(u"n_coronamiento_tags", 0)),
                    int(embed_res.get(u"n_coronamiento_tags_fail", 0)),
                )
            msg_ok += u"\n\nVerticales ext/int: estiradas={0}, retraídas={1}.".format(
                int(embed_res.get(u"n_extended", 0)),
                int(embed_res.get(u"n_retracted", 0)),
            )
            msg_ok += u"\nHorizontales ext/int: retraída={0} (ext.={1}, int.={2}); forma 06={3} (ext.={4}, int.={5}).".format(
                int(embed_res.get(u"n_horiz_retract", 0)),
                int(embed_res.get(u"n_horiz_retract_ext", 0)),
                int(embed_res.get(u"n_horiz_retract_int", 0)),
                int(embed_res.get(u"n_horiz_pata_l", 0)),
                int(embed_res.get(u"n_horiz_pata_l_ext", 0)),
                int(embed_res.get(u"n_horiz_pata_l_int", 0)),
            )
            if int(embed_res.get(u"n_cabezal", 0)) or int(embed_res.get(u"n_cabezal_fail", 0)):
                msg_ok += u"\nCabezal: capas={0}, fallos={1}.".format(
                    int(embed_res.get(u"n_cabezal", 0)),
                    int(embed_res.get(u"n_cabezal_fail", 0)),
                )
            n_tags = int(embed_res.get(u"n_tags_created", 0))
            if n_tags:
                msg_ok += u"\nEtiquetas cabezal longitudinales: {0}.".format(n_tags)
            embed_msgs = embed_res.get(u"messages") or []
            if embed_msgs:
                msg_ok += u"\n" + u"\n".join(unicode(m) for m in embed_msgs[:8])
                if len(embed_msgs) > 8:
                    msg_ok += u"\n…"
        if err:
            msg_ok += u"\n\n" + u"\n".join(err[:12])
            if len(err) > 12:
                msg_ok += u"\n…"
        _mostrar_resumen_fin_armado_muros(msg_ok)
        if wrap is not None:
            try:
                wrap._set_estado(msg_ok.replace(u"\n", u" ")[:500])
            except Exception:
                pass
            try:
                wrap._close_after_create()
            except Exception:
                pass

    def GetName(self):
        return u"AraincoCreacionArmadoMurosUnificado"


class ArmadoMurosPreviewWindow(object):

    # Paleta apagada UI — se asigna por espesor (mm), no por índice de fila.
    _THICKNESS_UI_PALETTE = (
        u"#4a7a88", u"#5a8268", u"#7a7348", u"#5a6690", u"#886070", u"#6a5888",
    )
    _COL_CYCLE = _THICKNESS_UI_PALETTE

    @staticmethod
    def _finite_px(value, default):
        """Devuelve ``default`` si ``value`` no es un número finito (p. ej. NaN en WPF sin medida)."""
        try:
            v = float(value)
        except Exception:
            return float(default)
        if v != v or v in (float("inf"), float("-inf")):
            return float(default)
        return v
    _CABEZAL_CAP_COL_PX = 420.0
    _CABEZAL_CAP_COL_MIN_PX = 340.0
    _CABEZAL_SECCION_COL_MIN_PX = 164.0
    _CABEZAL_CTRL_LBL_W_PX = 26.0
    _CABEZAL_CTRL_BTN_W_PX = 20.0
    _CABEZAL_CTRL_BTN_H_PX = 24.0
    _CABEZAL_CTRL_VAL_W_PX = 32.0
    _CABEZAL_CTRL_STEPPER_W_PX = 52.0
    _CABEZAL_CTRL_DIAM_W_PX = 56.0
    _CABEZAL_CTRL_STRIP_PX = 136.0
    _CABEZAL_CTRL_GAP_PX = 2.0
    _PREVIEW_ELEV_COL_PX = 500.0
    _PREVIEW_COL_PX = _CABEZAL_CAP_COL_PX * 2.0 + _PREVIEW_ELEV_COL_PX
    _CABEZAL_PREVIEW_CANVAS_H_PX = 60.0
    _CABEZAL_PREVIEW_CANVAS_W_PX = 175.0
    _CABEZAL_PREVIEW_BAR_SPAN_PX = 25.0
    _CABEZAL_PREVIEW_LAYER_PITCH_PX = 20.0
    _CABEZAL_ROW_FIXED_PX = 214.0
    _CABEZAL_ROW_MAX_PX = 214.0
    _CABEZAL_ROW_MIN_PX = 192.0
    _CABEZAL_LAYER_ROW_PX = 28.0
    _CABEZAL_LAYER_SCROLL_MAX_ROWS = 3
    _CABEZAL_EXTREMO_SPLIT_GAP_PX = 8.0
    _CABEZAL_EXTREMO_ARMADO_FRAC = 0.52
    _CABEZAL_CTRL_SCROLL_GUTTER_PX = 12.0
    _CABEZAL_SHELL_PAD_H_PX = 16.0
    _CABEZAL_SHELL_BORDER_H_PX = 2.0
    _CABEZAL_EXTREMO_BLOCK_GAP_PX = 10.0
    _CABEZAL_EXTREMO_INNER_GAP_PX = 8.0
    _CABEZAL_EXTREMO_TOOLBAR_ROW_PX = 24.0
    _CABEZAL_EXTREMO_TOOLBAR_GAP_PX = 6.0
    _TOGGLE_MINI_ANIM_MS = 420
    _TOGGLE_MINI_ANIM_INTERVAL_MS = 6
    _TOGGLE_MINI_FOLLOWUP_MS = 460
    _CABEZAL_PANEL_BODY_PAD_PX = 8.0
    _CABEZAL_UNIT_WRAP_PAD_TOP_PX = 4.0
    _CABEZAL_UNIT_WRAP_PAD_SIDE_PX = 8.0
    _CABEZAL_CAP_BASE_PX = 68.0
    _MESH_COL_PX = 300.0
    _ELEVATION_ROW_MIN_PX = 192.0
    _CABEZAL_PIE_SELECTOR_RESERVE_PX = 22.0
    _CABEZAL_SINGLE_TRAMO_ROW_EXTRA_PX = 18.0
    _MESH_ELEV_ROW_BODY_MIN_PX = 86.0
    _WINDOW_FRAME_PAD_PX = 96.0
    _PREVIEW_LEVEL_GUTTER_PX = 96
    _PREVIEW_WALL_LEVEL_GAP_PX = 12.0
    _LEVEL_RULER_COL_PX = 84.0
    _CABEZAL_ELEV_GAP_PX = 40.0
    _CABEZAL_ELEV_CANVAS_MARGIN_PX = 4.0
    _MESH_OVERLAY_COMPACT_MAX_W_PX = 212.0
    _BULK_CARD_H_PX = 148.0
    _EMPALME_TICK_W_PX = 10.0
    _EMPALME_TICK_STROKE_PX = 2.0
    _ENC_VEC_SEG_OFFSET_PX = 14.0
    _ENC_VEC_SEG_STROKE_PX = 1.5
    _ENC_VEC_SEG_Z = 8
    _EXTREMO_CHEVRON_Z = 30
    _TRAMO_CONN_STROKE_PX = 3.0
    _TRAMO_CONN_Z_GRID = 25
    _TRAMO_CONN_BADGE_W_PX = 30.0
    _TRAMO_CONN_BADGE_H_PX = 18.0

    def __init__(self, revit, uidoc, walls_list, mode=UI_MODE_MALLAS):
        self._revit = revit
        self._uidoc = uidoc
        self._ui_mode = mode if mode in (
            UI_MODE_CABEZAL, UI_MODE_MALLAS, UI_MODE_UNIFICADO,
        ) else UI_MODE_MALLAS
        if self._is_unificado_mode():
            self._modo_tradicional = True
        self.walls_list = [] if walls_list is None else list(walls_list)
        self.doc = uidoc.Document
        self._bar_labels_to_id = {}
        self._bar_id_to_label = {}
        self._diam_strings = []
        self._cabezal_diam_strings = []
        self._spacing_strings = ["100", "150", "200", "250", "300"]
        self._controls_by_wall_id = {}
        self._model_cache = {}
        self._foundation_cache = {}
        self._canvas_by_wall_id = {}
        self._right_by_wall_id = {}
        self._row_definitions = []
        self._preview_col_px = float(self._effective_preview_col_px())
        self._mesh_col_px = float(self._effective_mesh_col_px())
        self._preview_level_gutter_px = (
            0.0 if self._uses_cabezal_panels() else float(self._PREVIEW_LEVEL_GUTTER_PX)
        )
        self._ruler_canvas = None
        self._row_height_px = (
            float(getattr(self, "_CABEZAL_ROW_FIXED_PX", 280.0))
            if self._uses_cabezal_panels()
            else 200.0
        )
        self._prev_wrap_by_wall_id = {}
        self._row_grid_by_wall_id = {}
        self._length_scale_key = None
        self._preview_layout = None
        self._wall_extent_u_list = []
        self._max_extent_u_feet = 1.0
        self._modo_tradicional = True
        self._cabezal_by_wall_id = {}
        self._cabezal_ui_by_wall_id = {}
        self._cabezal_ui_by_segment = {}
        self._suppress_cabezal_stepper = False
        self._suppress_cabezal_confinement_cb = False
        self._suppress_cabezal_bulk_conf_cb = False
        self._suppress_cabezal_empalme_chk = False
        self._suppress_cabezal_armado_chk = False
        self._cabezal_bulk_ui = {}
        self._bulk_armado_toggle_hosts = {}
        self._malla_activo_by_wall_id = {}
        self._malla_ui_by_wall_id = {}
        self._suppress_malla_activo_chk = False
        self._bulk_malla_activo_chk = None
        self._bulk_mesh_params_body = None
        self._header_malla_activo_chk = None
        self._wall_thickness_color_map = {}
        self._defer_crear_raise = False
        self._cabezal_mounted_connectors = []
        self._ui_init_complete = False
        self._deferred_ui_init_started = False
        self._cabezal_segments_cache = {}
        self._single_tramo_row_cache = {}
        self._debounce_timers = {}
        self._stacked_layout = None

        self.walls_ordered = geo.ordenar_muros_por_base_asc(self.walls_list)
        self._walls_display_order = list(reversed(self.walls_ordered))

        self._win = None
        try:
            self._win = XamlReader.Parse(_build_preview_xaml())
        except Exception as ex_parse:
            TaskDialog.Show("Armado muros", u"No se cargó la ventana WPF:\n{}".format(str(ex_parse)))
            return

        if self._is_cabezal_mode():
            self._crear_handler = _CrearCabezalEjecutarHandler(weakref.ref(self))
        elif self._is_unificado_mode():
            self._crear_handler = _CrearUnificadoEjecutarHandler(weakref.ref(self))
        else:
            self._crear_handler = _CrearMallasEjecutarHandler(weakref.ref(self))
        self._crear_event = ExternalEvent.Create(self._crear_handler)

        _mode_ref = self._ui_mode
        try:
            from System import EventHandler as _EH_clr

            def _clear_singleton_slot(sender, evt):
                _unregister_preview_singleton(_mode_ref)

            self._win.Closed += _EH_clr(_clear_singleton_slot)
        except Exception:
            pass

        try:
            from System.Windows import RoutedEventHandler as _Ru_rd

            def _on_win_loaded(sender, evt):
                if getattr(self, "_deferred_ui_init_started", False):
                    return
                self._deferred_ui_init_started = True
                self._run_deferred_ui_init()

            def _on_win_size_changed(sender, evt):
                if not getattr(self, "_ui_init_complete", False):
                    return
                self._schedule_full_redraw()

            self._win.Loaded += _Ru_rd(_on_win_loaded)
            self._win.SizeChanged += _Ru_rd(_on_win_size_changed)
        except Exception:
            pass

        self._apply_mode_chrome()
        try:
            self._set_estado(u"Cargando interfaz…")
        except Exception:
            pass

    def _schedule_ui_debounce(self, key, callback, delay_ms=200):
        """Ejecuta ``callback`` una sola vez tras ``delay_ms`` sin nuevos eventos."""
        timers = getattr(self, "_debounce_timers", None) or {}
        self._debounce_timers = timers
        old = timers.get(key)
        if old is not None:
            try:
                old.Stop()
            except Exception:
                pass
        try:
            from System.Windows.Threading import DispatcherTimer
            from System import TimeSpan

            t = DispatcherTimer()
            t.Interval = TimeSpan(0, 0, 0, 0, int(delay_ms))

            def _tick(sender, args):
                try:
                    t.Stop()
                except Exception:
                    pass
                timers.pop(key, None)
                try:
                    callback()
                except Exception:
                    pass

            t.Tick += _tick
            timers[key] = t
            t.Start()
        except Exception:
            try:
                callback()
            except Exception:
                pass

    def _schedule_full_redraw(self, delay_ms=250):
        self._schedule_ui_debounce(
            u"full_redraw",
            self._redistribute_row_heights_and_redraw,
            delay_ms,
        )

    def _request_cabezal_preview_refresh(self, wid, extremo, debounce=True, delay_ms=180):
        if debounce:
            try:
                key = u"cprev_{0}_{1}".format(int(wid), extremo)
            except Exception:
                key = u"cprev_{0}_{1}".format(wid, extremo)
            self._schedule_ui_debounce(
                key,
                lambda w=wid, e=extremo: self._refresh_cabezal_preview(w, e),
                delay_ms,
            )
        else:
            self._refresh_cabezal_preview(wid, extremo)

    def _invalidate_cabezal_segments_cache(self, extremo=None):
        cache = getattr(self, "_cabezal_segments_cache", None)
        if cache is None:
            self._cabezal_segments_cache = {}
            cache = self._cabezal_segments_cache
        self._single_tramo_row_cache = {}
        if extremo is None:
            cache.clear()
            return
        try:
            del cache[extremo]
        except Exception:
            pass

    def _ensure_wall_preview_model(self, wid, wall):
        try:
            wid_i = int(wid)
        except Exception:
            wid_i = wid
        if wid_i not in self._model_cache:
            try:
                self._model_cache[wid_i] = geo.compute_section_preview_model(self.doc, wall)
            except Exception:
                self._model_cache[wid_i] = None
        self._ensure_foundation_preview_info(wid_i, wall)
        return self._model_cache.get(wid_i)

    def _ensure_foundation_preview_info(self, wid, wall):
        if wid in self._foundation_cache:
            return self._foundation_cache[wid]
        try:
            self._foundation_cache[wid] = self._load_foundation_preview_info(wall)
        except Exception:
            self._foundation_cache[wid] = None
        return self._foundation_cache[wid]

    def _run_deferred_ui_init(self):
        """Construcción pesada de UI tras mostrar la ventana (Loaded)."""
        if getattr(self, "_ui_init_complete", False):
            return
        try:
            self._prepare_ui_templates()
            if self._uses_cabezal_panels():
                view_right_xy = None
                try:
                    rd = self.doc.ActiveView.RightDirection
                    vr_x, vr_y = float(rd.X), float(rd.Y)
                    if (vr_x * vr_x + vr_y * vr_y) > 1e-9:
                        view_right_xy = (vr_x, vr_y)
                except Exception:
                    pass
                try:
                    self._stacked_layout = geo.compute_stacked_wall_layout(
                        self._walls_display_order,
                        view_right_xy=view_right_xy,
                    )
                except Exception:
                    self._stacked_layout = None
                if cabezal is not None:
                    legacy_ids = _pop_legacy_extremo_marker_ids()
                    if legacy_ids:
                        try:
                            t = Transaction(self.doc, u"Arainco: Limpiar marcadores extremos")
                            t.Start()
                            cabezal._delete_wall_extremo_markers(self.doc, legacy_ids)
                            t.Commit()
                        except Exception:
                            try:
                                t.RollBack()
                            except Exception:
                                pass
                self._init_cabezal_configs()
                self._sync_all_cabezal_troceo_auto()
                self._cabezal_cap_col_px_cached = self._compute_cabezal_cap_col_width_px()
                self._preview_col_px = float(self._effective_preview_col_px())
                self._refresh_wall_thickness_color_map()
            if self._is_mallas_mode() and not self._is_unificado_mode():
                self._sync_modo_from_checkboxes()
            self._build_wall_parameter_panels()
            self._wire_controls()
            self._refresh_info_txt()
            self._fit_window_to_content()
            self._redistribute_row_heights_and_redraw()
            self._ui_init_complete = True
            if self._is_unificado_mode():
                estado = u"Configura cabezal y malla; pulsa Crear armado completo."
            elif self._is_cabezal_mode():
                estado = u"Configura cabezales y pulsa Crear cabezales."
            else:
                estado = u"Configura mallas y pulsa Crear Area Reinf."
            self._set_estado(estado)
            if self._uses_cabezal_panels():
                try:
                    scr = self._win.FindName(u"ScrMuros")
                    if scr is not None:
                        scr.ScrollToEnd()
                except Exception:
                    pass
        except Exception as ex_init:
            self._ui_init_complete = True
            try:
                self._set_estado(u"Error al cargar UI: {0}".format(ex_init))
            except Exception:
                pass

    def _ensure_ui_ready_for_crear(self):
        if not getattr(self, "_ui_init_complete", False):
            try:
                self._run_deferred_ui_init()
            except Exception:
                pass
        return bool(getattr(self, "_ui_init_complete", False))

    def _is_cabezal_mode(self):
        return getattr(self, "_ui_mode", UI_MODE_MALLAS) == UI_MODE_CABEZAL

    def _is_unificado_mode(self):
        return getattr(self, "_ui_mode", UI_MODE_MALLAS) == UI_MODE_UNIFICADO

    def _uses_cabezal_panels(self):
        return self._is_cabezal_mode() or self._is_unificado_mode()

    def _is_mallas_mode(self):
        return getattr(self, "_ui_mode", UI_MODE_MALLAS) == UI_MODE_MALLAS

    def _effective_preview_col_px(self):
        if self._uses_cabezal_panels():
            gap = float(getattr(self, "_CABEZAL_ELEV_GAP_PX", 8.0))
            return float(self._cabezal_cap_col_px()) * 2.0 + gap * 2.0 + float(self._PREVIEW_ELEV_COL_PX)
        return float(self._PREVIEW_ELEV_COL_PX)

    def _cabezal_cap_col_px(self):
        cached = getattr(self, "_cabezal_cap_col_px_cached", None)
        if cached is not None:
            return float(cached)
        return float(self._compute_cabezal_cap_col_width_px())

    def _bulk_card_height_px(self):
        return float(getattr(self, u"_BULK_CARD_H_PX", 148.0))

    def _compute_cabezal_cap_col_width_px(self):
        """Ancho columna Final/Inicio + Configurador Global derivado del layout real."""
        _, _, _, strip_w, _, _ = self._cabezal_ctrl_metrics()
        scroll_gutter = float(getattr(self, "_CABEZAL_CTRL_SCROLL_GUTTER_PX", 12.0))
        split_gap = float(getattr(self, "_CABEZAL_EXTREMO_SPLIT_GAP_PX", 8.0))
        min_right = float(getattr(self, "_CABEZAL_SECCION_COL_MIN_PX", 164.0))
        min_arm = float(strip_w) + scroll_gutter
        content_w = min_arm + split_gap + min_right
        border_h = float(getattr(self, "_CABEZAL_SHELL_BORDER_H_PX", 2.0))
        pad_h = float(getattr(self, "_CABEZAL_SHELL_PAD_H_PX", 16.0))
        cap = content_w + border_h + pad_h
        floor = float(getattr(self, "_CABEZAL_CAP_COL_MIN_PX", 340.0))
        return max(floor, cap)

    @staticmethod
    def _ui_brush_hex(hx, alpha=238):
        from System.Windows.Media import SolidColorBrush, Color

        h = (hx or u"#5a6690").strip().lstrip(u"#")
        if len(h) < 6:
            h = u"5a6690"
        rr = int(h[0:2], 16)
        gg = int(h[2:4], 16)
        bb = int(h[4:6], 16)
        aa = max(0, min(255, int(alpha)))
        return SolidColorBrush(Color.FromArgb(aa, rr, gg, bb))

    @staticmethod
    def _ui_lighten_hex_brush(hx, mix=0.55, alpha=255):
        from System.Windows.Media import SolidColorBrush, Color

        h = (hx or u"#5a6690").strip().lstrip(u"#")
        if len(h) < 6:
            h = u"5a6690"
        rr = int(h[0:2], 16)
        gg = int(h[2:4], 16)
        bb = int(h[4:6], 16)
        m = max(0.0, min(1.0, float(mix)))
        rr = int(rr + (255 - rr) * m)
        gg = int(gg + (255 - gg) * m)
        bb = int(bb + (255 - bb) * m)
        aa = max(0, min(255, int(alpha)))
        return SolidColorBrush(Color.FromArgb(aa, rr, gg, bb))

    def _wall_elevation_color_hex(self, wall, row_index=0):
        if self._uses_cabezal_panels() and wall is not None:
            return self._wall_thickness_color_hex(wall)
        palette = list(ArmadoMurosPreviewWindow._COL_CYCLE)
        return palette[row_index % len(palette)]

    def _wall_thickness_mm_ui(self, wall):
        if wall is None:
            return 200.0
        try:
            th = geo.obtener_espesor_muro_mm_approx(wall)
            if th is not None and float(th) > 0.0:
                return float(th)
        except Exception:
            pass
        return 200.0

    def _wall_thickness_mm_key(self, wall):
        try:
            return int(round(float(self._wall_thickness_mm_ui(wall))))
        except Exception:
            return 200

    def _refresh_wall_thickness_color_map(self):
        """Espesor mm (redondeado) -> color UI; estable en toda la sesión."""
        palette = list(ArmadoMurosPreviewWindow._THICKNESS_UI_PALETTE)
        thicknesses = set()
        for w in getattr(self, u"walls_ordered", []) or []:
            thicknesses.add(self._wall_thickness_mm_key(w))
        if not thicknesses:
            thicknesses.add(200)
        mp = {}
        for i, th in enumerate(sorted(thicknesses)):
            mp[int(th)] = palette[i % len(palette)]
        self._wall_thickness_color_map = mp
        return mp

    def _wall_thickness_color_hex(self, wall):
        if not getattr(self, u"_wall_thickness_color_map", None):
            self._refresh_wall_thickness_color_map()
        key = self._wall_thickness_mm_key(wall)
        mp = getattr(self, u"_wall_thickness_color_map", {}) or {}
        if key in mp:
            return mp[key]
        palette = list(ArmadoMurosPreviewWindow._THICKNESS_UI_PALETTE)
        return palette[abs(int(key)) % len(palette)]

    def _build_wall_thickness_legend_panel(self):
        """Leyenda compacta e=XXX mm (solo si hay más de un espesor)."""
        from System.Windows.Controls import StackPanel, TextBlock, Border, Orientation
        from System.Windows import (
            Thickness,
            FontWeights,
            VerticalAlignment,
            CornerRadius,
        )
        from System.Windows.Media import SolidColorBrush, Color

        mp = getattr(self, u"_wall_thickness_color_map", None) or {}
        if len(mp) < 2:
            return None

        host = StackPanel()
        host.Orientation = Orientation.Horizontal
        host.Margin = Thickness(0, 0, 0, 0)
        host.VerticalAlignment = VerticalAlignment.Center

        hint = TextBlock()
        hint.Text = u"Espesor:"
        hint.Foreground = SolidColorBrush(Color.FromRgb(100, 116, 139))
        hint.FontSize = 9.0
        hint.FontWeight = FontWeights.SemiBold
        hint.Margin = Thickness(0, 0, 8, 0)
        hint.VerticalAlignment = VerticalAlignment.Center
        host.Children.Add(hint)

        for th in sorted(mp.keys()):
            hx = mp[th]
            row = StackPanel()
            row.Orientation = Orientation.Horizontal
            row.Margin = Thickness(0, 0, 10, 0)
            row.VerticalAlignment = VerticalAlignment.Center
            sw = Border()
            sw.Width = 10.0
            sw.Height = 10.0
            sw.CornerRadius = CornerRadius(2.0)
            sw.Background = self._ui_brush_hex(hx, alpha=220)
            sw.BorderBrush = self._ui_brush_hex(hx, alpha=255)
            sw.BorderThickness = Thickness(1)
            sw.Margin = Thickness(0, 0, 4, 0)
            sw.VerticalAlignment = VerticalAlignment.Center
            lbl = TextBlock()
            lbl.Text = u"e={0}".format(int(th))
            lbl.Foreground = SolidColorBrush(Color.FromRgb(148, 163, 184))
            lbl.FontSize = 9.0
            lbl.VerticalAlignment = VerticalAlignment.Center
            row.Children.Add(sw)
            row.Children.Add(lbl)
            host.Children.Add(row)
        return host

    def _effective_mesh_col_px(self):
        if self._uses_cabezal_panels():
            return 0.0
        return float(self._MESH_COL_PX)

    def _apply_mode_chrome(self):
        if self._win is None:
            return
        try:
            from System.Windows import Visibility

            uni = self._is_unificado_mode()
            cab = self._is_cabezal_mode()
            if uni:
                self._win.Title = u"Arainco: Armado Muros"
            elif cab:
                self._win.Title = u"Arainco: Cabezal muros"
            else:
                self._win.Title = u"Arainco: Mallas muros"
            title_tb = self._win.FindName(u"TxtTitle")
            if title_tb is not None:
                title_tb.Text = self._win.Title
            sub_tb = self._win.FindName(u"TxtSubtitle")
            if sub_tb is not None:
                if uni:
                    sub_tb.Visibility = Visibility.Visible
                    sub_tb.Text = (
                        u"Solo muro tradicional: cabezal ini/fin + malla ext.+int. "
                        u"No incluye muro de contención."
                    )
                elif cab:
                    sub_tb.Text = u""
                    sub_tb.Visibility = Visibility.Collapsed
                else:
                    sub_tb.Visibility = Visibility.Visible
                    sub_tb.Text = (
                        u"Area Reinforcement por tramo: diámetro y espaciado ext./int."
                    )
            foot_tb = self._win.FindName(u"TxtFooterHint")
            if foot_tb is not None:
                if uni:
                    foot_tb.Visibility = Visibility.Visible
                    foot_tb.Text = (
                        u"Creación: longitudinales → confinamiento → mallas; "
                        u"etiquetas long. → conf. → malla. "
                        u"Confinamiento: ø >= L0/3 (catálogo); @ fijo 100 mm. "
                        u"Orden inferior→superior."
                    )
                elif cab:
                    foot_tb.Text = u""
                    foot_tb.Visibility = Visibility.Collapsed
                else:
                    foot_tb.Visibility = Visibility.Visible
                    foot_tb.Text = (
                        u"Creación: AR + Remove System + post-proceso vert./horiz. "
                        u"(sin cabezal). Orden inferior→superior."
                    )
            pnl_modo = self._win.FindName(u"PnlModoMuro")
            if pnl_modo is not None:
                pnl_modo.Visibility = (
                    Visibility.Visible
                    if self._is_mallas_mode() and not uni
                    else Visibility.Collapsed
                )
            self._ensure_header_malla_activo_toggle()
            btn_cab = self._win.FindName(u"BtnAplicarCabezal")
            btn_mesh = self._win.FindName(u"BtnAplicarMallas")
            btn_crear = self._win.FindName(u"BtnCrear")
            if btn_cab is not None:
                btn_cab.Visibility = (
                    Visibility.Visible
                    if (cab or uni) and cabezal is not None
                    else Visibility.Collapsed
                )
            if btn_mesh is not None:
                btn_mesh.Visibility = (
                    Visibility.Visible
                    if self._is_mallas_mode() or uni
                    else Visibility.Collapsed
                )
            if btn_crear is not None:
                if uni:
                    btn_crear.Content = u"Crear armado completo"
                elif cab:
                    btn_crear.Content = u"Crear cabezales"
                else:
                    btn_crear.Content = u"Crear Area Reinf."
            bdr_bulk = self._win.FindName(u"BdrCabezalBulkActions")
            if bdr_bulk is not None:
                bdr_bulk.Visibility = (
                    Visibility.Visible
                    if (cab or uni) and cabezal is not None
                    else Visibility.Collapsed
                )
        except Exception:
            pass

    def _set_estado(self, msg):
        try:
            t = self._win.FindName("TxtEstado")
            if t is not None:
                t.Text = msg or ""
        except Exception:
            pass

    def _close_after_create(self):
        """Cierra la ventana y libera el singleton tras crear Area Reinforcement."""
        try:
            if self._win is not None:
                self._win.Close()
        except Exception:
            pass
        try:
            _unregister_preview_singleton(self._ui_mode)
        except Exception:
            pass

    def _mesh_modo_tradicional(self):
        return bool(getattr(self, "_modo_tradicional", True))

    def _sync_modo_from_checkboxes(self):
        if self._win is None:
            return
        try:
            chk_cont = self._win.FindName("ChkMuroContencion")
            if chk_cont is not None and chk_cont.IsChecked == True:
                self._modo_tradicional = False
                return
        except Exception:
            pass
        self._modo_tradicional = True
        try:
            chk_trad = self._win.FindName("ChkMuroTradicional")
            if chk_trad is not None:
                chk_trad.IsChecked = True
        except Exception:
            pass

    def _rebuild_mesh_ui(self):
        if self._win is None:
            return
        self._build_wall_parameter_panels()
        self._fit_window_to_content()
        self._redistribute_row_heights_and_redraw()

    def _apply_flat_combo(self, cb, narrow=False, stretch=False):
        if cb is None or self._win is None:
            return
        try:
            from System.Windows.Controls import ComboBox as _Cb

            if stretch:
                st = self._win.TryFindResource(u"ComboStretch")
            elif narrow:
                st = (
                    self._win.TryFindResource(u"ComboDiam")
                    or self._win.TryFindResource(u"Combo")
                )
            else:
                st = self._win.TryFindResource(u"Combo")
            it = self._win.TryFindResource(u"ComboItem")
            if st is not None:
                cb.Style = st
            if it is not None:
                cb.ItemContainerStyle = it
            if narrow:
                w = float(getattr(self, "_CABEZAL_CTRL_DIAM_W_PX", 56.0))
                cb.Width = w
                cb.MinWidth = w
                cb.MaxWidth = w
                try:
                    cb.FontSize = 10.0
                except Exception:
                    pass
            elif stretch:
                try:
                    cb.ClearValue(_Cb.WidthProperty)
                    cb.ClearValue(_Cb.MinWidthProperty)
                    cb.ClearValue(_Cb.MaxWidthProperty)
                    cb.MinWidth = 0.0
                except Exception:
                    pass
            else:
                try:
                    cb.ClearValue(_Cb.WidthProperty)
                    cb.ClearValue(_Cb.MaxWidthProperty)
                except Exception:
                    pass
        except Exception:
            pass

    def _apply_bimtools_textbox(self, tb):
        """TextBox tema oscuro (Empalme masivo, etc.)."""
        if tb is None or self._win is None:
            return
        try:
            from System.Windows.Controls import TextBox as _Tb
            from System.Windows.Media import SolidColorBrush, Color

            st = self._win.TryFindResource(u"BimToolsTextBoxDark")
            if st is not None:
                tb.Style = st
                return
        except Exception:
            pass
        try:
            from System.Windows import Thickness
            from System.Windows.Media import SolidColorBrush, Color

            tb.Background = SolidColorBrush(Color.FromRgb(5, 14, 24))
            tb.Foreground = SolidColorBrush(Color.FromRgb(232, 244, 248))
            tb.BorderBrush = SolidColorBrush(Color.FromRgb(33, 70, 92))
            tb.BorderThickness = Thickness(1)
            tb.CaretBrush = SolidColorBrush(Color.FromRgb(122, 163, 184))
        except Exception:
            pass

    def _apply_bimtools_toggle_mini(self, chk):
        """Estilo CheckBox sin caja nativa (solo ContentPresenter)."""
        if chk is None or self._win is None:
            return
        applied = False
        try:
            st = self._win.TryFindResource(u"BimToolsToggleMini")
            if st is not None:
                chk.Style = st
                applied = True
        except Exception:
            pass
        if applied:
            return
        try:
            from System.Windows.Controls import CheckBox as _Cb, ContentPresenter, ControlTemplate
            from System.Windows import VerticalAlignment
            from System.Windows.Markup import XamlReader

            tpl = XamlReader.Parse(
                u'<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" '
                u'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" '
                u'TargetType="{x:Type CheckBox}">'
                u'<ContentPresenter VerticalAlignment="Center"/>'
                u'</ControlTemplate>'
            )
            chk.Template = tpl
        except Exception:
            pass

    def _toggle_mini_parts_key(self, ui):
        return (ui or {}).get(u"toggle_mini_parts_key") or u"troceo_toggle_parts"

    def _build_toggle_mini_content(self, chk, ui):
        """Track + thumb en Content; refs en ui (IronPython no persiste attrs en CheckBox WPF)."""
        if chk is None or ui is None:
            return
        from System.Windows.Controls import StackPanel

        parts_key = self._toggle_mini_parts_key(ui)
        parts = ui.get(parts_key)
        if parts and isinstance(chk.Content, StackPanel):
            return
        ui.pop(parts_key, None)
        from System.Windows.Controls import Border, Orientation, StackPanel, TextBlock
        from System.Windows import HorizontalAlignment, VerticalAlignment, Thickness, CornerRadius, FontWeights
        from System.Windows.Media import Color, SolidColorBrush, TranslateTransform

        host = StackPanel()
        host.Orientation = Orientation.Horizontal
        host.VerticalAlignment = VerticalAlignment.Center

        track_fill = SolidColorBrush(Color.FromRgb(18, 38, 54))
        track_border = SolidColorBrush(Color.FromRgb(33, 70, 92))
        track = Border()
        track.Width = 32.0
        track.Height = 16.0
        track.CornerRadius = CornerRadius(8.0)
        track.Background = track_fill
        track.BorderBrush = track_border
        track.BorderThickness = Thickness(1)
        track.Margin = Thickness(0, 0, 6, 0)
        track.VerticalAlignment = VerticalAlignment.Center
        track.HorizontalAlignment = HorizontalAlignment.Left
        track.ClipToBounds = True
        track.SnapsToDevicePixels = True

        thumb_xform = TranslateTransform(0.0, 0.0)
        thumb = Border()
        thumb.Width = 12.0
        thumb.Height = 12.0
        thumb.CornerRadius = CornerRadius(6.0)
        thumb.Background = SolidColorBrush(Color.FromRgb(232, 244, 248))
        thumb.HorizontalAlignment = HorizontalAlignment.Left
        thumb.Margin = Thickness(2, 0, 0, 0)
        thumb.VerticalAlignment = VerticalAlignment.Center
        thumb.RenderTransform = thumb_xform
        thumb.SnapsToDevicePixels = True
        track.Child = thumb

        host.Children.Add(track)
        lbl_text = (ui or {}).get(u"toggle_mini_label")
        if lbl_text is None:
            lbl_text = u"Empalmes"
        if unicode(lbl_text).strip():
            lbl = TextBlock()
            lbl.Text = lbl_text
            lbl.FontSize = 10.0
            lbl.FontWeight = FontWeights.SemiBold
            lbl.VerticalAlignment = VerticalAlignment.Center
            lbl_fg = (ui or {}).get(u"toggle_mini_label_fg")
            if lbl_fg is not None:
                lbl.Foreground = lbl_fg
            else:
                lbl.Foreground = SolidColorBrush(Color.FromRgb(226, 232, 240))
            host.Children.Add(lbl)
        else:
            track.Margin = Thickness(0)
        chk.Content = host
        ui[parts_key] = {
            u"thumb_xform": thumb_xform,
            u"track_fill": track_fill,
            u"track_border": track_border,
            u"timer": None,
        }

    @staticmethod
    def _toggle_mini_color_on():
        from System.Windows.Media import Color
        return Color.FromRgb(34, 211, 238)

    @staticmethod
    def _toggle_mini_color_off_fill():
        from System.Windows.Media import Color
        return Color.FromRgb(18, 38, 54)

    @staticmethod
    def _toggle_mini_color_off_border():
        from System.Windows.Media import Color
        return Color.FromRgb(33, 70, 92)

    @staticmethod
    def _lerp_color_channel(a, b, t):
        return int(round(float(a) + (float(b) - float(a)) * float(t)))

    @staticmethod
    def _ease_in_out_sine(t):
        import math
        t = max(0.0, min(1.0, float(t)))
        return 0.5 * (1.0 - math.cos(math.pi * t))

    def _lerp_media_color(self, c0, c1, t):
        from System.Windows.Media import Color
        return Color.FromRgb(
            self._lerp_color_channel(c0.R, c1.R, t),
            self._lerp_color_channel(c0.G, c1.G, t),
            self._lerp_color_channel(c0.B, c1.B, t),
        )

    def _toggle_mini_parts(self, ui, parts_key=None):
        pk = parts_key or self._toggle_mini_parts_key(ui)
        parts = (ui or {}).get(pk) or {}
        return (
            parts.get(u"thumb_xform"),
            parts.get(u"track_fill"),
            parts.get(u"track_border"),
            parts,
        )

    def _stop_toggle_mini_anim(self, ui, parts_key=None):
        pk = parts_key or self._toggle_mini_parts_key(ui)
        parts = (ui or {}).get(pk)
        if not parts:
            return
        timer = parts.get(u"timer")
        if timer is not None:
            try:
                timer.Stop()
            except Exception:
                pass
            parts[u"timer"] = None

    def _apply_toggle_mini_visual(self, ui, checked, animate=False, parts_key=None):
        """Actualiza thumb + colores del track (snap o animado)."""
        if ui is None:
            return
        pk = parts_key or self._toggle_mini_parts_key(ui)
        if pk == u"armado_toggle_parts":
            chk = ui.get(u"armado_activo_chk")
        elif pk == u"malla_activo_toggle_parts":
            chk = ui.get(u"malla_activo_chk")
        else:
            chk = ui.get(u"troceo_por_muro_chk")
        if chk is None:
            chk = (
                ui.get(u"malla_activo_chk")
                or ui.get(u"troceo_por_muro_chk")
                or ui.get(u"armado_activo_chk")
            )
        if chk is not None and not ui.get(pk):
            try:
                self._build_toggle_mini_content(chk, ui)
            except Exception:
                pass
        thumb_xform, track_fill, track_border, parts = self._toggle_mini_parts(ui, pk)
        if thumb_xform is None or parts is None:
            return
        on = bool(checked)
        if not animate:
            self._stop_toggle_mini_anim(ui, pk)
            try:
                thumb_xform.X = 15.0 if on else 0.0
                on_c = self._toggle_on_color_for_ui(ui)
                off_fill = self._toggle_mini_color_off_fill()
                off_border = self._toggle_mini_color_off_border()
                if track_fill is not None:
                    track_fill.Color = on_c if on else off_fill
                if track_border is not None:
                    track_border.Color = on_c if on else off_border
            except Exception:
                pass
            return
        try:
            from System import TimeSpan
            from System.Windows import RoutedEventHandler as _REH
            from System.Windows.Threading import DispatcherTimer

            self._stop_toggle_mini_anim(ui, pk)

            target_x = 15.0 if on else 0.0
            start_x = float(thumb_xform.X)
            on_c = self._toggle_on_color_for_ui(ui)
            off_fill = self._toggle_mini_color_off_fill()
            off_border = self._toggle_mini_color_off_border()
            end_fill = on_c if on else off_fill
            end_border = on_c if on else off_border
            start_fill = track_fill.Color if track_fill is not None else off_fill
            start_border = track_border.Color if track_border is not None else off_border

            duration_ms = int(getattr(self, u"_TOGGLE_MINI_ANIM_MS", 420))
            interval_ms = int(getattr(self, u"_TOGGLE_MINI_ANIM_INTERVAL_MS", 6))
            steps = max(1, int(round(float(duration_ms) / float(interval_ms))))
            state = {u"i": 0}

            timer = DispatcherTimer()
            timer.Interval = TimeSpan.FromMilliseconds(interval_ms)

            def _tick(sender, args, u=ui, tm=timer, p=parts):
                try:
                    state[u"i"] += 1
                    t_raw = min(1.0, float(state[u"i"]) / float(steps))
                    t = self._ease_in_out_sine(t_raw)
                    thumb_xform.X = start_x + (target_x - start_x) * t
                    if track_fill is not None:
                        track_fill.Color = self._lerp_media_color(
                            start_fill, end_fill, t,
                        )
                    if track_border is not None:
                        track_border.Color = self._lerp_media_color(
                            start_border, end_border, t,
                        )
                    if state[u"i"] >= steps:
                        tm.Stop()
                        p[u"timer"] = None
                        thumb_xform.X = target_x
                        if track_fill is not None:
                            track_fill.Color = end_fill
                        if track_border is not None:
                            track_border.Color = end_border
                except Exception:
                    try:
                        tm.Stop()
                    except Exception:
                        pass
                    try:
                        self._apply_toggle_mini_visual(u, on, animate=False)
                    except Exception:
                        pass

            timer.Tick += _REH(_tick)
            parts[u"timer"] = timer
            timer.Start()
        except Exception:
            try:
                self._apply_toggle_mini_visual(ui, on, animate=False)
            except Exception:
                pass

    def _schedule_cabezal_empalme_followup(self, wid, extremo, delay_ms=None):
        """Aplaza redraw + rebuild para no bloquear la animación del toggle."""
        if delay_ms is None:
            delay_ms = int(getattr(self, u"_TOGGLE_MINI_FOLLOWUP_MS", 460))
        try:
            from System import TimeSpan
            from System.Windows import RoutedEventHandler as _REH
            from System.Windows.Threading import DispatcherTimer

            key = u"_cabezal_emp_follow_{0}".format(extremo)
            bag = getattr(self, key, None)
            if not isinstance(bag, dict):
                bag = {u"walls": set(), u"timer": None}
                setattr(self, key, bag)
            try:
                bag[u"walls"].add(int(wid))
            except Exception:
                pass
            old_timer = bag.get(u"timer")
            if old_timer is not None:
                try:
                    old_timer.Stop()
                except Exception:
                    pass
            timer = DispatcherTimer()
            timer.Interval = TimeSpan.FromMilliseconds(int(delay_ms))

            def _tick(sender, args, ex=extremo, wbag=bag, tm=timer, k=key):
                try:
                    tm.Stop()
                except Exception:
                    pass
                walls = list(wbag.get(u"walls") or [])
                wbag[u"walls"] = set()
                wbag[u"timer"] = None
                for w in walls:
                    try:
                        self._redraw_wall_elevation_canvas(w)
                    except Exception:
                        pass
                try:
                    self._rebuild_cabezal_all_walls_for_extremo(ex)
                except Exception:
                    pass

            timer.Tick += _REH(_tick)
            bag[u"timer"] = timer
            timer.Start()
        except Exception:
            try:
                self._redraw_wall_elevation_canvas(wid)
            except Exception:
                pass
            try:
                self._rebuild_cabezal_all_walls_for_extremo(extremo)
            except Exception:
                pass

    def _apply_cabezal_slider(self, sl):
        if sl is None or self._win is None:
            return
        try:
            st = self._win.TryFindResource(u"BimToolsSliderCompact")
            if st is not None:
                sl.Style = st
        except Exception:
            pass

    def _load_foundation_preview_info(self, wall):
        u"""Fundación estructural unida al muro (Join Geometry), misma regla que post-proceso."""
        if wall is None or self.doc is None:
            return None
        try:
            import armado_muros_verticales_embed_colision as _emb_mod

            funds = _emb_mod._fundaciones_estructurales_unidas_muro(self.doc, wall)
            if not funds:
                return None
            h_max = None
            for fund in funds:
                h_mm = _emb_mod._altura_bbox_elemento_mm(fund)
                if h_mm is not None:
                    h_max = max(h_max or 0.0, float(h_mm))
            return {
                u"count": len(funds),
                u"height_mm": h_max,
            }
        except Exception:
            return None

    def _foundation_footing_height_px(self, box_h, fund_info):
        if not fund_info:
            return 0.0
        bh = max(10.0, float(box_h))
        return min(28.0, max(16.0, bh * 0.22))

    def _draw_foundation_schematic(self, canv, x_off, draw_w, box_h, foot_h, fund_info, color_hex=None):
        u"""Zapata esquemática bajo el muro cuando hay fundación unida."""
        from System.Windows.Controls import Canvas as _Cn
        from System.Windows.Shapes import Rectangle as _Wr, Line as _Ln
        from System.Windows import FontWeights

        if not fund_info or foot_h <= 0.0:
            return

        hx = color_hex or u"#5a6690"
        fill_br = self._ui_brush_hex(hx, alpha=100)
        stroke_br = self._ui_brush_hex(hx, alpha=230)
        text_br = self._ui_lighten_hex_brush(hx, mix=0.52)

        foot_w = min(max(float(draw_w) * 1.16, float(draw_w) + 14.0), float(draw_w) + 28.0)
        foot_x = float(x_off) - (foot_w - float(draw_w)) * 0.5
        foot_y = max(0.0, float(box_h) - float(foot_h))

        foot = _Wr()
        foot.Width = foot_w
        foot.Height = float(foot_h)
        foot.Fill = fill_br
        foot.Stroke = stroke_br
        foot.StrokeThickness = 1.0
        _Cn.SetLeft(foot, foot_x)
        _Cn.SetTop(foot, foot_y)
        self._canvas_set_zindex(foot, 12)
        canv.Children.Add(foot)

        joint = _Ln()
        joint.X1 = foot_x + 1.0
        joint.X2 = foot_x + foot_w - 1.0
        joint.Y1 = foot_y
        joint.Y2 = foot_y
        joint.Stroke = stroke_br
        joint.StrokeThickness = 1.2
        self._canvas_set_zindex(joint, 13)
        canv.Children.Add(joint)

        h_mm = fund_info.get(u"height_mm")
        if h_mm is not None and float(h_mm) > 0.1:
            cap = u"Fund. H≈{:.0f}".format(float(h_mm))
        else:
            n_f = int(fund_info.get(u"count", 1))
            cap = u"Fund." if n_f <= 1 else u"Fund. ×{}".format(n_f)

        self._add_canvas_text_centered(
            canv,
            cap,
            foot_x + foot_w * 0.5,
            foot_y + float(foot_h) * 0.5,
            text_br,
            8.0,
            FontWeights.SemiBold,
            14,
        )

    def _prepare_ui_templates(self):
        bar_pairs = _get_bar_types_sorted_display(self.doc)
        self._diam_strings = [d for d, _ in bar_pairs]
        self._bar_labels_to_id = {}
        self._bar_id_to_label = {}
        for disp, bt in bar_pairs:
            self._bar_labels_to_id[disp] = bt.Id
            try:
                self._bar_labels_to_id[str(disp)] = bt.Id
            except Exception:
                pass
            compact = _cabezal_compact_diam_label(disp)
            if compact and compact != disp:
                self._bar_labels_to_id[compact] = bt.Id
                try:
                    self._bar_labels_to_id[str(compact)] = bt.Id
                except Exception:
                    pass
            try:
                self._bar_id_to_label[geo._element_id_int(bt.Id)] = disp
            except Exception:
                pass
        self._cabezal_diam_strings = [
            _cabezal_compact_diam_label(d) for d in self._diam_strings
        ]

    def _fill_cabezal_diam_combo(self, cb):
        strings = getattr(self, "_cabezal_diam_strings", None) or self._diam_strings
        if cb is not None and strings:
            cb.ItemsSource = list(strings)
            try:
                cb.SelectedIndex = _cabezal_default_diam_index(strings)
            except Exception:
                pass

    def _fill_combo_diam_esp(self, cb_diam, cb_esp):
        if cb_diam is not None and self._diam_strings:
            cb_diam.ItemsSource = list(self._diam_strings)
            try:
                cb_diam.SelectedIndex = _cabezal_default_diam_index(self._diam_strings)
            except Exception:
                pass
        if cb_esp is not None:
            cb_esp.ItemsSource = list(self._spacing_strings)
            try:
                cb_esp.SelectedIndex = 1
            except Exception:
                pass

    def _spacing_string_for_mm(self, spacing_mm):
        """Etiqueta de combo de espaciamiento más cercana a ``spacing_mm``."""
        try:
            target = int(round(float(spacing_mm)))
        except Exception:
            target = 200
        strings = list(getattr(self, u"_spacing_strings", None) or [u"200"])
        best = strings[0]
        best_diff = None
        for s in strings:
            try:
                v = int(round(float(str(s).strip().replace(u",", u"."))))
            except Exception:
                continue
            diff = abs(v - target)
            if best_diff is None or diff < best_diff:
                best_diff = diff
                best = s
        try:
            return unicode(best).strip()
        except Exception:
            return str(best).strip()

    def _set_spacing_combo_value(self, cb, spacing_mm):
        if cb is None:
            return
        esp_txt = self._spacing_string_for_mm(spacing_mm)
        try:
            items = list(cb.ItemsSource or getattr(self, u"_spacing_strings", []) or [])
            if esp_txt in items:
                cb.SelectedItem = esp_txt
                return
        except Exception:
            pass
        try:
            cb.Text = esp_txt
        except Exception:
            pass

    def _set_diam_combo_by_mm(self, cb, diam_mm):
        if cb is None or not self._diam_strings:
            return
        lab = _cabezal_diam_label_for_mm(self._diam_strings, diam_mm)
        if not lab:
            return
        ids_map = getattr(self, u"_bar_labels_to_id", {}) or {}
        bid = ids_map.get(lab) or ids_map.get(str(lab))
        if bid is not None:
            try:
                from Autodesk.Revit.DB import ElementId as _EI

                if bid != _EI.InvalidElementId:
                    self._set_diam_combo_selected_id(cb, bid)
                    return
            except Exception:
                pass
        try:
            items = list(cb.ItemsSource or self._diam_strings)
            if lab in items:
                cb.SelectedItem = lab
        except Exception:
            pass

    def _apply_malla_sic_defaults_for_wall(self, wall, ctr):
        """
        Pre-rellena combos de malla (doble S.I.C.) según espesor del muro.
        Editable después por el usuario.
        """
        if wall is None or not ctr:
            return
        mesh_keys = set(self._mesh_control_keys())
        if not any(k in ctr for k in mesh_keys):
            return
        try:
            diam_mm, spacing_mm = geo.malla_sic_defaults_para_muro(wall)
        except Exception:
            return
        for k in mesh_keys:
            cb = ctr.get(k)
            if cb is None:
                continue
            if k.endswith(u"_md") or k.endswith(u"_id"):
                self._set_diam_combo_by_mm(cb, diam_mm)
            elif k.endswith(u"_ms") or k.endswith(u"_is"):
                self._set_spacing_combo_value(cb, spacing_mm)

    def _cabezal_clamp_int(self, value, min_v, max_v, default=None):
        if default is None:
            default = min_v
        try:
            n = int(round(float(value)))
        except Exception:
            n = int(default)
        return max(int(min_v), min(int(max_v), n))

    def _cabezal_read_value_tb(self, tb, min_v, max_v, default):
        if tb is None:
            return self._cabezal_clamp_int(default, min_v, max_v, default)
        try:
            return self._cabezal_clamp_int(tb.Text, min_v, max_v, default)
        except Exception:
            return self._cabezal_clamp_int(default, min_v, max_v, default)

    def _apply_cabezal_stepper_btn(self, btn):
        if btn is None:
            return
        try:
            from System.Windows import Thickness, HorizontalAlignment, VerticalAlignment

            if self._win is not None:
                st = self._win.TryFindResource(u"BimToolsStepperZoneBtn")
                if st is not None:
                    btn.Style = st
            btn.Padding = Thickness(0, 0, 0, 0)
            btn.Margin = Thickness(0, 0, 0, 0)
            btn.HorizontalAlignment = HorizontalAlignment.Stretch
            btn.VerticalAlignment = VerticalAlignment.Stretch
            btn.FontSize = 11.0
        except Exception:
            pass

    def _wire_cabezal_stepper_shell_hover(self, shell):
        if shell is None:
            return
        try:
            from System.Windows import Thickness
            from System.Windows.Input import MouseEventHandler
            from System.Windows.Media import SolidColorBrush, Color

            bg_idle = SolidColorBrush(Color.FromRgb(5, 14, 24))
            bg_hot = SolidColorBrush(Color.FromRgb(11, 23, 40))
            br_idle = SolidColorBrush(Color.FromRgb(26, 58, 77))
            br_hot = SolidColorBrush(Color.FromRgb(76, 115, 131))

            def _idle(sender, e):
                try:
                    shell.Background = bg_idle
                    shell.BorderBrush = br_idle
                    shell.BorderThickness = Thickness(1)
                except Exception:
                    pass

            def _hot(sender, e):
                try:
                    shell.Background = bg_hot
                    shell.BorderBrush = br_hot
                    shell.BorderThickness = Thickness(1)
                except Exception:
                    pass

            shell.MouseEnter += MouseEventHandler(_hot)
            shell.MouseLeave += MouseEventHandler(_idle)
        except Exception:
            pass

    def _create_cabezal_stepper(self, min_v, max_v, initial, on_change=None, palette=None):
        from System.Windows.Controls import (
            Grid, Button, TextBlock, Border,
            ColumnDefinition, RowDefinition,
        )
        from System.Windows import (
            Thickness,
            FontWeights,
            VerticalAlignment,
            HorizontalAlignment,
            TextAlignment,
            CornerRadius,
            GridLength,
            GridUnitType,
        )
        from System.Windows.Media import SolidColorBrush, Color
        from System.Windows.Shapes import Path as WpfPath
        from System.Windows.Media import PathGeometry, PathFigure, LineSegment
        from System.Windows import Point as WpfPoint

        val_w = float(getattr(self, "_CABEZAL_CTRL_VAL_W_PX", 32.0))
        arrow_col_w = 20.0
        btn_h = float(getattr(self, "_CABEZAL_CTRL_BTN_H_PX", 24.0))
        step_w = val_w + arrow_col_w
        sep_br = SolidColorBrush(Color.FromRgb(26, 58, 77))
        zone_bg = SolidColorBrush(Color.FromRgb(17, 37, 61))

        shell = Border()
        shell.Background = SolidColorBrush(Color.FromRgb(5, 14, 24))
        shell.BorderBrush = sep_br
        shell.BorderThickness = Thickness(1)
        shell.CornerRadius = CornerRadius(4.0)
        shell.Padding = Thickness(0)
        shell.Width = step_w
        shell.MinWidth = step_w
        shell.MaxWidth = step_w
        shell.Height = btn_h
        shell.MinHeight = btn_h
        shell.MaxHeight = btn_h
        shell.VerticalAlignment = VerticalAlignment.Center
        shell.HorizontalAlignment = HorizontalAlignment.Center
        shell.SnapsToDevicePixels = True
        shell.ClipToBounds = True

        panel = Grid()
        panel.VerticalAlignment = VerticalAlignment.Stretch
        panel.HorizontalAlignment = HorizontalAlignment.Stretch
        panel.SnapsToDevicePixels = True
        cd_val = ColumnDefinition()
        cd_val.Width = GridLength(val_w, GridUnitType.Pixel)
        cd_arr = ColumnDefinition()
        cd_arr.Width = GridLength(arrow_col_w, GridUnitType.Pixel)
        panel.ColumnDefinitions.Add(cd_val)
        panel.ColumnDefinitions.Add(cd_arr)

        val_tb = TextBlock()
        val_tb.TextAlignment = TextAlignment.Center
        val_tb.VerticalAlignment = VerticalAlignment.Center
        val_tb.HorizontalAlignment = HorizontalAlignment.Center
        val_tb.FontSize = 11.0
        Grid.SetColumn(val_tb, 0)

        arrow_panel = Grid()
        arrow_panel.VerticalAlignment = VerticalAlignment.Stretch
        arrow_panel.HorizontalAlignment = HorizontalAlignment.Stretch
        arrow_panel.Background = zone_bg
        arrow_panel.SnapsToDevicePixels = True
        Grid.SetColumn(arrow_panel, 1)

        rd_up = RowDefinition()
        rd_up.Height = GridLength(1.0, GridUnitType.Star)
        rd_sep = RowDefinition()
        rd_sep.Height = GridLength(1.0, GridUnitType.Pixel)
        rd_dn = RowDefinition()
        rd_dn.Height = GridLength(1.0, GridUnitType.Star)
        arrow_panel.RowDefinitions.Add(rd_up)
        arrow_panel.RowDefinitions.Add(rd_sep)
        arrow_panel.RowDefinitions.Add(rd_dn)

        sep_line = Border()
        sep_line.Background = sep_br
        sep_line.Height = 1.0
        sep_line.HorizontalAlignment = HorizontalAlignment.Stretch
        Grid.SetRow(sep_line, 1)
        arrow_panel.Children.Add(sep_line)

        left_sep = Border()
        left_sep.Width = 1.0
        left_sep.Background = sep_br
        left_sep.HorizontalAlignment = HorizontalAlignment.Left
        left_sep.VerticalAlignment = VerticalAlignment.Stretch
        Grid.SetRowSpan(left_sep, 3)
        arrow_panel.Children.Add(left_sep)

        def _make_chevron_path(is_up):
            fg_dim = SolidColorBrush(Color.FromRgb(149, 184, 204))
            p = WpfPath()
            p.Stroke = fg_dim
            p.StrokeThickness = 1.4
            p.HorizontalAlignment = HorizontalAlignment.Center
            p.VerticalAlignment = VerticalAlignment.Center
            p.Width = 10.0
            p.Height = 6.0
            p.StrokeLineJoin = p.StrokeLineJoin.Round
            p.StrokeStartLineCap = p.StrokeStartLineCap.Round
            p.StrokeEndLineCap = p.StrokeEndLineCap.Round
            geo = PathGeometry()
            fig = PathFigure()
            fig.IsClosed = False
            fig.IsFilled = False
            if is_up:
                fig.StartPoint = WpfPoint(1.0, 5.0)
                fig.Segments.Add(LineSegment(WpfPoint(5.0, 1.0), True))
                fig.Segments.Add(LineSegment(WpfPoint(9.0, 5.0), True))
            else:
                fig.StartPoint = WpfPoint(1.0, 1.0)
                fig.Segments.Add(LineSegment(WpfPoint(5.0, 5.0), True))
                fig.Segments.Add(LineSegment(WpfPoint(9.0, 1.0), True))
            geo.Figures.Add(fig)
            p.Data = geo
            return p

        btn_up = Button()
        btn_up.Content = _make_chevron_path(True)
        self._apply_cabezal_stepper_btn(btn_up)
        Grid.SetRow(btn_up, 0)
        arrow_panel.Children.Add(btn_up)

        btn_dn = Button()
        btn_dn.Content = _make_chevron_path(False)
        self._apply_cabezal_stepper_btn(btn_dn)
        Grid.SetRow(btn_dn, 2)
        arrow_panel.Children.Add(btn_dn)

        state = {u"value": self._cabezal_clamp_int(initial, min_v, max_v, min_v)}

        def _apply_display(v):
            state[u"value"] = self._cabezal_clamp_int(v, min_v, max_v, min_v)
            val_tb.Text = str(state[u"value"])
            try:
                btn_dn.IsEnabled = state[u"value"] > int(min_v)
                btn_up.IsEnabled = state[u"value"] < int(max_v)
            except Exception:
                pass

        def _change(delta):
            if getattr(self, "_suppress_cabezal_stepper", False):
                return
            _apply_display(state[u"value"] + int(delta))
            if on_change is not None:
                try:
                    on_change(state[u"value"])
                except Exception:
                    pass

        try:
            from System.Windows import RoutedEventHandler as _REH

            btn_up.Click += _REH(lambda s, e: _change(1))
            btn_dn.Click += _REH(lambda s, e: _change(-1))
        except Exception:
            pass

        panel.Children.Add(val_tb)
        panel.Children.Add(arrow_panel)
        shell.Child = panel
        self._wire_cabezal_stepper_shell_hover(shell)
        _apply_display(state[u"value"])
        if palette is not None:
            self._cabezal_apply_stepper_value_style(val_tb, palette)
        else:
            val_tb.Foreground = SolidColorBrush(Color.FromRgb(238, 246, 250))
            val_tb.FontWeight = FontWeights.SemiBold

        return {
            u"panel": shell,
            u"value_tb": val_tb,
            u"btn_minus": btn_dn,
            u"btn_plus": btn_up,
            u"get_value": lambda: int(state[u"value"]),
            u"set_value": _apply_display,
            u"min_v": int(min_v),
            u"max_v": int(max_v),
        }

    def _cabezal_n_capas_from_ui(self, wid, extremo):
        if cabezal is None:
            return 2
        ui = self._cabezal_ui_ext(wid, extremo)
        try:
            cap_step = ui.get(u"capas_stepper")
            if cap_step is not None:
                try:
                    n_step = int(cap_step[u"get_value"]())
                    min_cap = self._cabezal_min_capas_for_wall_extremo(wid, extremo)
                    return max(
                        min_cap,
                        min(cabezal.CABEZAL_MAX_CAPAS, n_step),
                    )
                except Exception:
                    pass
        except Exception:
            pass
        try:
            tb = ui.get(u"capas_value_tb")
            if tb is not None:
                min_cap = self._cabezal_min_capas_for_wall_extremo(wid, extremo)
                return self._cabezal_read_value_tb(
                    tb,
                    min_cap,
                    cabezal.CABEZAL_MAX_CAPAS,
                    max(min_cap, 2),
                )
        except Exception:
            pass
        cfg = self._cabezal_by_wall_id.get(wid) or {}
        ly = list((cfg.get(extremo) or {}).get(u"layers") or [])
        try:
            n_cfg = int((cfg.get(extremo) or {}).get(u"n_capas", len(ly) or 2))
            return max(
                cabezal.CABEZAL_MIN_CAPAS,
                min(cabezal.CABEZAL_MAX_CAPAS, n_cfg),
            )
        except Exception:
            pass
        n = len(ly) if ly else 2
        return max(
            cabezal.CABEZAL_MIN_CAPAS,
            min(cabezal.CABEZAL_MAX_CAPAS, n),
        )

    def _max_capas_cabezal_wall(self, wid):
        if cabezal is None:
            return 2
        n_max = 2
        cfg = self._cabezal_by_wall_id.get(wid) or {}
        for ex in cabezal.CABEZAL_EXTREMOS:
            try:
                ly = (cfg.get(ex) or {}).get(u"layers") or []
                n_max = max(n_max, len(ly))
            except Exception:
                pass
        for ex in cabezal.CABEZAL_EXTREMOS:
            try:
                ui_ex = self._cabezal_ui_ext(wid, ex)
                tb = ui_ex.get(u"capas_value_tb")
                if tb is not None:
                    n_max = max(
                        n_max,
                        self._cabezal_read_value_tb(
                            tb,
                            cabezal.CABEZAL_MIN_CAPAS,
                            cabezal.CABEZAL_MAX_CAPAS,
                            2,
                        ),
                    )
            except Exception:
                pass
        return max(2, min(cabezal.CABEZAL_MAX_CAPAS, int(n_max)))

    def _cabezal_preview_canvas_size_px(self, wid=None, extremo=None):
        """Tamaño del canvas vista sección (menor en split extremo @ cap col)."""
        ch = self._cabezal_preview_height_px(wid, extremo)
        default_w = float(getattr(self, "_CABEZAL_PREVIEW_CANVAS_W_PX", 175.0))
        if wid is not None and extremo is not None:
            try:
                ui = self._cabezal_ui_ext(wid, extremo)
                pw = ui.get(u"preview_canvas_w_px")
                if pw is not None:
                    return float(pw), ch
            except Exception:
                pass
        return default_w, ch

    def _cabezal_shell_inner_width_px(self, cap_w=None):
        """Ancho útil dentro del Border cabezal (borde + padding horizontales)."""
        cap = float(
            cap_w if cap_w is not None
            else self._cabezal_cap_col_px(),
        )
        border_h = float(getattr(self, "_CABEZAL_SHELL_BORDER_H_PX", 2.0))
        pad_h = float(getattr(self, "_CABEZAL_SHELL_PAD_H_PX", 16.0))
        return max(280.0, cap - border_h - pad_h)

    def _cabezal_split_vline(self, sep_br, gap_px, min_h=0.0):
        """Separador vertical de altura completa en columna central del split."""
        from System.Windows.Controls import Grid, Border, RowDefinition
        from System.Windows import (
            GridLength,
            GridUnitType,
            HorizontalAlignment,
            VerticalAlignment,
        )

        gap = max(1.0, float(gap_px))
        host = Grid()
        host.Width = gap
        host.MinWidth = gap
        host.MaxWidth = gap
        host.HorizontalAlignment = HorizontalAlignment.Center
        host.VerticalAlignment = VerticalAlignment.Stretch
        if float(min_h) > 0.0:
            host.MinHeight = float(min_h)
        rd = RowDefinition()
        rd.Height = GridLength(1.0, GridUnitType.Star)
        host.RowDefinitions.Add(rd)
        line = Border()
        line.Width = 1.0
        line.Background = sep_br
        line.HorizontalAlignment = HorizontalAlignment.Center
        line.VerticalAlignment = VerticalAlignment.Stretch
        line.SnapsToDevicePixels = True
        host.Children.Add(line)
        return host

    def _cabezal_extremo_split_layout_px(self, cap_w=None):
        """Columnas Prop. C: armado izq. | sección der. dentro de cap_w."""
        cap = float(
            cap_w if cap_w is not None
            else self._cabezal_cap_col_px(),
        )
        gap = float(getattr(self, "_CABEZAL_EXTREMO_SPLIT_GAP_PX", 8.0))
        frac = float(getattr(self, "_CABEZAL_EXTREMO_ARMADO_FRAC", 0.52))
        _, _, _, strip_w, _, _ = self._cabezal_ctrl_metrics()
        scroll_gutter = float(getattr(self, "_CABEZAL_CTRL_SCROLL_GUTTER_PX", 12.0))
        min_arm = float(strip_w) + scroll_gutter
        min_right = float(getattr(self, "_CABEZAL_SECCION_COL_MIN_PX", 164.0))
        content_w = self._cabezal_shell_inner_width_px(cap)
        left_w = max(min_arm, int(content_w * frac))
        right_w = content_w - left_w - gap
        if right_w < min_right:
            right_w = min_right
            left_w = max(min_arm, content_w - gap - right_w)
        total = left_w + gap + right_w
        if total > content_w + 0.5:
            right_w = max(min_right, right_w - (total - content_w))
            left_w = max(min_arm, content_w - gap - right_w)
        preview_w = max(104.0, min(right_w, content_w - left_w - gap))
        return left_w, right_w, preview_w, content_w, gap

    def _cabezal_seccion_group_pad_px(self):
        return float(getattr(self, "_CABEZAL_SECCION_GROUP_PAD_PX", 4.0))

    def _cabezal_seccion_canvas_width_px(self, right_w):
        """Ancho del canvas dentro del Border Sección (resta padding del grupo)."""
        pad = self._cabezal_seccion_group_pad_px()
        return max(104.0, float(right_w) - 2.0 * pad)

    def _cabezal_preview_height_px(self, wid=None, extremo=None, n_capas=None):
        """Altura del canvas sección; mayor en encuentro L."""
        base = float(getattr(self, "_CABEZAL_PREVIEW_CANVAS_H_PX", 60.0))
        enc_h = float(getattr(self, "_CABEZAL_PREVIEW_CANVAS_ENC_L_H_PX", 88.0))
        if (
            wid is not None
            and extremo is not None
            and self._cabezal_extremo_es_encuentro_preview(wid, extremo)
        ):
            return enc_h
        return base

    def _cabezal_preview_draw_rect_px(self, cw=None, ch=None):
        """Rectángulo muro (outer) en px — siempre el mismo tamaño lógico."""
        if cw is None or ch is None:
            cw, ch = self._cabezal_preview_canvas_size_px()
        margin_x = float(getattr(self, "_CABEZAL_PREVIEW_CANVAS_MARGIN_X_PX", 8.0))
        margin_y = float(getattr(self, "_CABEZAL_PREVIEW_CANVAS_MARGIN_Y_PX", 5.0))
        draw_w = max(40.0, float(cw) - 2.0 * margin_x)
        draw_h = max(28.0, float(ch) - 2.0 * margin_y)
        x0 = margin_x
        y0 = margin_y + max(0.0, (float(ch) - draw_h - 2.0 * margin_y) * 0.5)
        return x0, y0, draw_w, draw_h

    def _cabezal_layers_scroll_max_height_px(self, n_capas=None, header_px=None):
        _, _, _, _, row_h, _ = self._cabezal_ctrl_metrics()
        max_rows = int(getattr(self, "_CABEZAL_LAYER_SCROLL_MAX_ROWS", 3))
        hdr = float(header_px if header_px is not None else row_h)
        return hdr + float(row_h) * float(max_rows)

    def _cabezal_wall_max_capas(self, wid):
        """Mayor n_capas entre Inicio/Final para estimar altura de fila."""
        if cabezal is None:
            return 2
        n = int(cabezal.CABEZAL_MIN_CAPAS)
        for ex in cabezal.CABEZAL_EXTREMOS:
            try:
                n = max(n, int(self._cabezal_n_capas_from_ui(wid, ex)))
            except Exception:
                pass
        return max(
            cabezal.CABEZAL_MIN_CAPAS,
            min(cabezal.CABEZAL_MAX_CAPAS, n),
        )

    def _cabezal_content_height_px(self, n_capas=None, wid=None, extremo=None):
        """Altura estimada del panel extremo (Prop. C split @ cap col)."""
        block_gap = float(getattr(self, "_CABEZAL_EXTREMO_BLOCK_GAP_PX", 10.0))
        inner_gap = float(getattr(self, "_CABEZAL_EXTREMO_INNER_GAP_PX", 8.0))
        wrap_pad = float(getattr(self, u"_CABEZAL_UNIT_WRAP_PAD_TOP_PX", 4.0))
        wrap_pad += float(getattr(self, u"_CABEZAL_UNIT_WRAP_PAD_SIDE_PX", 8.0))
        slack = float(getattr(self, "_CABEZAL_PANEL_BODY_PAD_PX", 8.0))
        if n_capas is None:
            n_capas = cabezal.CABEZAL_MIN_CAPAS if cabezal else 2
        try:
            n_capas = max(2, min(cabezal.CABEZAL_MAX_CAPAS, int(n_capas)))
        except Exception:
            n_capas = 2
        toolbar_h = float(getattr(self, "_CABEZAL_EXTREMO_TOOLBAR_ROW_PX", 24.0))
        toolbar_h += 4.0
        split_top = inner_gap * 0.35
        layers_block = self._cabezal_layers_scroll_max_height_px(n_capas)
        preview_h = self._cabezal_preview_height_px(wid, extremo)
        hdr = 9.0 + inner_gap
        armado_block = hdr + layers_block
        is_enc = (
            wid is not None
            and extremo is not None
            and self._cabezal_extremo_es_encuentro_preview(wid, extremo)
        )
        conf_footer = 0.0 if is_enc else 28.0
        preview_block = hdr + preview_h + block_gap + conf_footer
        body_h = max(armado_block, preview_block)
        return int(wrap_pad + toolbar_h + split_top + body_h + slack)

    def _cabezal_display_row_for_stack_index(self, stack_idx):
        walls = getattr(self, u"walls_ordered", []) or []
        od = getattr(self, u"_walls_display_order", []) or []
        if not (0 <= int(stack_idx) < len(walls)):
            return 0
        try:
            target = _wall_id_int(walls[int(stack_idx)])
        except Exception:
            return 0
        for ri, w in enumerate(od):
            try:
                if _wall_id_int(w) == target:
                    return ri
            except Exception:
                pass
        return 0

    def _cabezal_owner_wid_for(self, wid, extremo):
        if cabezal is None:
            return int(wid)
        seg = self._cabezal_segment_for_wall(wid, extremo)
        try:
            owner_idx = int(seg.get(u"owner_index", 0))
        except Exception:
            owner_idx = 0
        walls = getattr(self, u"walls_ordered", []) or []
        if walls and 0 <= owner_idx < len(walls):
            try:
                return _wall_id_int(walls[owner_idx])
            except Exception:
                pass
        return int(wid)

    def _cabezal_is_segment_owner(self, wid, extremo):
        return int(wid) == int(self._cabezal_owner_wid_for(wid, extremo))

    def _cabezal_segment_panel_height_px(self, seg, extremo):
        if cabezal is None or not seg:
            return float(getattr(self, u"_CABEZAL_ROW_MIN_PX", 232.0))
        try:
            owner_idx = int(seg.get(u"owner_index", 0))
        except Exception:
            owner_idx = 0
        walls = getattr(self, u"walls_ordered", []) or []
        owner_wid = 0
        if walls and 0 <= owner_idx < len(walls):
            try:
                owner_wid = _wall_id_int(walls[owner_idx])
            except Exception:
                owner_wid = 0
        try:
            n_capas = self._cabezal_wall_max_capas(owner_wid)
        except Exception:
            n_capas = cabezal.CABEZAL_MIN_CAPAS
        return float(self._cabezal_content_height_px(n_capas, owner_wid, extremo))

    def _cabezal_display_rows_for_segment(self, seg):
        rows = set()
        for wi in seg.get(u"wall_indices") or []:
            try:
                rows.add(self._cabezal_display_row_for_stack_index(int(wi)))
            except Exception:
                pass
        return sorted(rows)

    def _cabezal_segment_align_stack_index(self, seg):
        """Índice en ``walls_ordered`` del muro de referencia visual del tramo."""
        indices = list(seg.get(u"wall_indices") or [])
        if not indices:
            return 0
        if len(indices) < 2:
            return int(indices[0])
        return int(indices[1])

    def _cabezal_segment_align_display_row(self, seg):
        return self._cabezal_display_row_for_stack_index(
            self._cabezal_segment_align_stack_index(seg),
        )

    def _cabezal_segment_span_bounds(self, seg):
        rows = self._cabezal_display_rows_for_segment(seg)
        if not rows:
            return 0, 1
        return int(rows[0]), len(rows)

    def _cabezal_segment_combined_row_height(self, seg):
        row_heights = getattr(self, u"_last_row_heights", None)
        if not row_heights:
            row_heights = self._compute_row_heights()
        total = 0.0
        for ri in self._cabezal_display_rows_for_segment(seg):
            if 0 <= ri < len(row_heights):
                total += float(row_heights[ri])
        elev_min = float(getattr(self, u"_ELEVATION_ROW_MIN_PX", 188.0))
        return max(total, elev_min)

    def _cabezal_segment_align_top_margin_px(self, seg, cap_h):
        """Margen superior para centrar el panel en la fila ``wall_indices[1]``."""
        start_ri, span = self._cabezal_segment_span_bounds(seg)
        align_ri = self._cabezal_segment_align_display_row(seg)
        row_heights = getattr(self, u"_last_row_heights", None)
        if not row_heights:
            row_heights = self._compute_row_heights()
        offset_top = 0.0
        for ri in range(start_ri, align_ri):
            if 0 <= ri < len(row_heights):
                offset_top += float(row_heights[ri])
        align_row_h = float(row_heights[align_ri]) if 0 <= align_ri < len(row_heights) else 0.0
        if align_row_h <= 0.0:
            align_row_h = float(getattr(self, u"_ELEVATION_ROW_MIN_PX", 188.0))
        align_center = offset_top + align_row_h * 0.5
        top_margin = align_center - float(cap_h) * 0.5
        combined_h = self._cabezal_segment_combined_row_height(seg)
        top_margin = max(0.0, min(top_margin, max(0.0, combined_h - float(cap_h))))
        return top_margin

    def _mount_cabezal_segment_caps(self, cab_stack):
        """Un panel cabezal por tramo, centrado en ``wall_indices[1]`` (o [0] si hay uno)."""
        from System.Windows import VerticalAlignment, Thickness

        pending = getattr(self, u"_cabezal_pending_caps", []) or []
        self._cabezal_mounted_caps = []
        for extremo, owner_wid, wall, col, mirror in pending:
            seg = self._cabezal_segment_for_wall(owner_wid, extremo)
            start_ri, span = self._cabezal_segment_span_bounds(seg)
            cap_h = self._cabezal_segment_panel_height_px(seg, extremo)
            top_margin = self._cabezal_segment_align_top_margin_px(seg, cap_h)
            cap = self._build_cabezal_extremo_cap(
                owner_wid, wall, extremo, cap_h, mirror_preview=mirror,
            )
            cap.VerticalAlignment = VerticalAlignment.Top
            try:
                cap.Margin = Thickness(0, top_margin, 0, 0)
                cap.MinHeight = cap_h
            except Exception:
                pass
            try:
                from System.Windows.Controls import Grid
                Grid.SetRow(cap, start_ri)
                Grid.SetRowSpan(cap, span)
                Grid.SetColumn(cap, col)
            except Exception:
                pass
            cab_stack.Children.Add(cap)
            self._cabezal_mounted_caps.append((cap, seg, extremo))

    @staticmethod
    def _parse_hex_color_rgb(hex_str, default=(100, 116, 139)):
        try:
            s = str(hex_str or u"").strip().lstrip(u"#")
            if len(s) >= 6:
                return (
                    int(s[0:2], 16),
                    int(s[2:4], 16),
                    int(s[4:6], 16),
                )
        except Exception:
            pass
        return default

    def _cabezal_tramo_color_rgb(self, seg):
        try:
            idx = int((seg or {}).get(u"id", 0))
        except Exception:
            idx = 0
        palette = list(self._THICKNESS_UI_PALETTE)
        if not palette:
            return (100, 116, 139)
        return self._parse_hex_color_rgb(palette[idx % len(palette)])

    @staticmethod
    def _cabezal_tramo_label_id(seg):
        try:
            return int((seg or {}).get(u"id", 0)) + 1
        except Exception:
            return 1

    def _cabezal_tramo_connector_gap_col(self, owner_wid, extremo):
        if cabezal is None:
            return None
        ri = self._cabezal_display_row_for_stack_index(
            self._cabezal_wall_stack_index(int(owner_wid)),
        )
        ex_izq, ex_der = self._cabezal_extremos_lados_wall(int(owner_wid), ri)
        if extremo == ex_izq:
            return 1
        if extremo == ex_der:
            return 3
        return None

    def _layout_cabezal_tramo_connector(self, host, ui):
        from System.Windows.Controls import Canvas as _Cn
        from System.Windows import TextAlignment

        try:
            w = float(host.ActualWidth)
            h = float(host.ActualHeight)
        except Exception:
            return
        if w < 4.0 or h < 4.0:
            return
        cx = w * 0.5
        ln = ui.get(u"line")
        if ln is not None:
            ln.X1 = cx
            ln.X2 = cx
            ln.Y1 = 0.0
            ln.Y2 = h
        bw = float(getattr(self, u"_TRAMO_CONN_BADGE_W_PX", 30.0))
        bh = float(getattr(self, u"_TRAMO_CONN_BADGE_H_PX", 18.0))
        cy = h * 0.5
        bg = ui.get(u"badge_bg")
        if bg is not None:
            _Cn.SetLeft(bg, cx - bw * 0.5)
            _Cn.SetTop(bg, cy - bh * 0.5)
        tb = ui.get(u"badge_tb")
        if tb is not None:
            try:
                tb.Width = bw
                tb.TextAlignment = TextAlignment.Center
            except Exception:
                pass
            _Cn.SetLeft(tb, cx - bw * 0.5)
            _Cn.SetTop(tb, cy - 8.0)

    def _layout_cabezal_tramo_connectors_all(self):
        for host, seg, _extremo in (
            getattr(self, u"_cabezal_mounted_connectors", []) or []
        ):
            try:
                ui = host.Tag
            except Exception:
                ui = None
            if isinstance(ui, dict):
                self._layout_cabezal_tramo_connector(host, ui)

    def _build_cabezal_tramo_connector(self, seg, extremo):
        from System.Windows.Controls import Canvas as _Cn, TextBlock, Panel
        from System.Windows.Shapes import Line, Rectangle
        from System.Windows import (
            VerticalAlignment,
            HorizontalAlignment,
            FontWeights,
            TextAlignment,
        )
        from System.Windows.Media import SolidColorBrush, Color, Brushes

        r, g, b = self._cabezal_tramo_color_rgb(seg)
        seg_lbl = self._cabezal_tramo_label_id(seg)
        brush = SolidColorBrush(Color.FromRgb(int(r), int(g), int(b)))

        host = _Cn()
        host.Background = Brushes.Transparent
        host.ClipToBounds = False
        host.VerticalAlignment = VerticalAlignment.Stretch
        host.HorizontalAlignment = HorizontalAlignment.Stretch
        try:
            host.IsHitTestVisible = False
        except Exception:
            pass

        ln = Line()
        ln.Stroke = brush
        ln.StrokeThickness = float(getattr(self, u"_TRAMO_CONN_STROKE_PX", 3.0))
        try:
            from System.Windows.Media import PenLineCap
            ln.StrokeStartLineCap = PenLineCap.Round
            ln.StrokeEndLineCap = PenLineCap.Round
        except Exception:
            pass
        ln.Opacity = 0.92

        badge_bg = Rectangle()
        badge_bg.Width = float(getattr(self, u"_TRAMO_CONN_BADGE_W_PX", 30.0))
        badge_bg.Height = float(getattr(self, u"_TRAMO_CONN_BADGE_H_PX", 18.0))
        badge_bg.RadiusX = 4.0
        badge_bg.RadiusY = 4.0
        badge_bg.Fill = brush
        badge_bg.Opacity = 0.92
        badge_bg.IsHitTestVisible = False

        badge_tb = TextBlock()
        badge_tb.Text = u"T{0}".format(seg_lbl)
        badge_tb.FontSize = 11.0
        badge_tb.FontWeight = FontWeights.Bold
        badge_tb.Foreground = SolidColorBrush(Color.FromRgb(7, 16, 24))
        badge_tb.Width = float(getattr(self, u"_TRAMO_CONN_BADGE_W_PX", 30.0))
        badge_tb.TextAlignment = TextAlignment.Center
        badge_tb.IsHitTestVisible = False

        host.Children.Add(ln)
        host.Children.Add(badge_bg)
        host.Children.Add(badge_tb)
        self._canvas_set_zindex(ln, 1)
        self._canvas_set_zindex(badge_bg, 2)
        self._canvas_set_zindex(badge_tb, 3)

        ui = {u"line": ln, u"badge_bg": badge_bg, u"badge_tb": badge_tb}
        try:
            host.Tag = ui
        except Exception:
            pass

        def _on_size(sender, args, h=host, u=ui):
            self._layout_cabezal_tramo_connector(h, u)

        try:
            from System.Windows import SizeChangedEventHandler as _SCEH
            host.SizeChanged += _SCEH(_on_size)
        except Exception:
            try:
                from System.Windows import RoutedEventHandler as _REH
                host.SizeChanged += _REH(_on_size)
            except Exception:
                pass
        return host

    def _mount_cabezal_tramo_connectors(self, cab_stack):
        if cabezal is None or cab_stack is None:
            return
        from System.Windows.Controls import Grid, Panel
        from System.Windows import VerticalAlignment

        self._cabezal_mounted_connectors = []
        walls = getattr(self, u"walls_ordered", []) or []
        for extremo in cabezal.CABEZAL_EXTREMOS:
            for seg in self._cabezal_segments_for_extremo(extremo):
                try:
                    owner_idx = int(seg.get(u"owner_index", 0))
                except Exception:
                    owner_idx = 0
                if not (0 <= owner_idx < len(walls)):
                    continue
                try:
                    owner_wid = _wall_id_int(walls[owner_idx])
                except Exception:
                    continue
                gap_col = self._cabezal_tramo_connector_gap_col(owner_wid, extremo)
                if gap_col is None:
                    continue
                conn = self._build_cabezal_tramo_connector(seg, extremo)
                conn.VerticalAlignment = VerticalAlignment.Stretch
                start_ri, span = self._cabezal_segment_span_bounds(seg)
                try:
                    Grid.SetRow(conn, start_ri)
                    Grid.SetRowSpan(conn, span)
                    Grid.SetColumn(conn, gap_col)
                    Panel.SetZIndex(conn, int(getattr(self, u"_TRAMO_CONN_Z_GRID", 25)))
                except Exception:
                    pass
                cab_stack.Children.Add(conn)
                self._cabezal_mounted_connectors.append((conn, seg, extremo))
        self._layout_cabezal_tramo_connectors_all()

    def _cabezal_layer_steppers_are_live(self, ui):
        """True si los steppers n/ø pertenecen al panel activo (no UI destruida)."""
        if not ui or not isinstance(ui, dict):
            return False
        steppers = ui.get(u"layer_steppers") or []
        if not steppers:
            return False
        ctrl_grid = ui.get(u"controls_grid")
        if ctrl_grid is None:
            return False
        try:
            grid_children = list(ctrl_grid.Children)
        except Exception:
            return False
        for st in steppers:
            if not st or not isinstance(st, dict):
                continue
            panel = st.get(u"panel")
            if panel is None:
                return False
            try:
                if panel not in grid_children:
                    return False
            except Exception:
                return False
        return True

    def _cabezal_seed_segment_owners_from_predecessor(self, extremo):
        """
        Tras un empalme, copia armado del tramo anterior al propietario del
        tramo nuevo si aún tiene n/ø por defecto (p. ej. n=2).
        """
        if cabezal is None:
            return
        segs = self._cabezal_segments_for_extremo(extremo)
        if len(segs) <= 1:
            return
        copy_keys = (
            u"layers",
            u"n_capas",
            u"confinement",
            u"bar_type_id",
            u"conf_bar_type_id",
            u"segment_bar_type_ids",
            u"armado_activo",
            u"layer_spacing_mm",
            u"encuentro_tipo",
            u"vecino_wall_id",
            u"espesor_detectado_mm",
            u"espesor_seleccionado_mm",
            u"pitch_equitativo_mm",
            u"sic_encuentro_key",
            u"sic_encuentro_total",
            u"sic_encuentro_diam_mm",
            u"sic_basis",
        )
        walls = getattr(self, u"walls_ordered", []) or []
        prev_ex_cfg = None
        for seg in segs:
            try:
                owner_idx = int(seg.get(u"owner_index", 0))
            except Exception:
                owner_idx = 0
            if not (0 <= owner_idx < len(walls)):
                continue
            try:
                owner_wid = _wall_id_int(walls[owner_idx])
            except Exception:
                continue
            cfg = self._cabezal_by_wall_id.setdefault(
                owner_wid, cabezal.default_cabezal_muro_config(),
            )
            ex_cfg = cfg.setdefault(extremo, cabezal.default_cabezal_extremo_config())
            if prev_ex_cfg is not None and self._cabezal_extremo_needs_armado_seed(
                ex_cfg, prev_ex_cfg,
            ):
                for k in copy_keys:
                    if k in prev_ex_cfg:
                        ex_cfg[k] = _copy_cabezal_extremo_field(k, prev_ex_cfg[k])
                cabezal._normalize_cabezal_extremo_layers(ex_cfg)
                self._cabezal_propagate_segment_armado_from_owner(extremo, owner_wid)
            prev_ex_cfg = ex_cfg

    def _cabezal_extremo_needs_armado_seed(self, dst_ex, src_ex):
        """True si ``dst`` sigue con n por capa en default y ``src`` ya fue editado."""
        if cabezal is None or not src_ex or not dst_ex:
            return False
        try:
            src_layers = cabezal.cabezal_active_layers(src_ex)
            dst_layers = cabezal.cabezal_active_layers(dst_ex)
        except Exception:
            return False
        if not src_layers:
            return False
        if not dst_layers:
            return True
        src_has_custom_n = False
        dst_all_default_n = True
        for sl in src_layers:
            try:
                if int(sl.get(u"n_bars", cabezal.CABEZAL_MIN_BARRAS_POR_CAPA)) > cabezal.CABEZAL_MIN_BARRAS_POR_CAPA:
                    src_has_custom_n = True
                    break
            except Exception:
                pass
        if not src_has_custom_n:
            return False
        for dl in dst_layers:
            try:
                if int(dl.get(u"n_bars", cabezal.CABEZAL_MIN_BARRAS_POR_CAPA)) != cabezal.CABEZAL_MIN_BARRAS_POR_CAPA:
                    dst_all_default_n = False
                    break
            except Exception:
                pass
        return dst_all_default_n

    def _cabezal_propagate_segment_armado_from_owner(self, extremo, owner_wid):
        if cabezal is None:
            return
        seg = self._cabezal_segment_for_wall(owner_wid, extremo)
        owner_cfg = (
            (self._cabezal_by_wall_id.get(int(owner_wid)) or {}).get(extremo) or {}
        )
        if not owner_cfg:
            return
        copy_keys = (
            u"layers",
            u"n_capas",
            u"confinement",
            u"bar_type_id",
            u"conf_bar_type_id",
            u"segment_bar_type_ids",
            u"armado_activo",
            u"layer_spacing_mm",
            u"encuentro_tipo",
            u"vecino_wall_id",
            u"espesor_detectado_mm",
            u"espesor_seleccionado_mm",
            u"pitch_equitativo_mm",
            u"sic_encuentro_key",
            u"sic_encuentro_total",
            u"sic_encuentro_diam_mm",
            u"sic_basis",
        )
        walls = getattr(self, u"walls_ordered", []) or []
        for wi in seg.get(u"wall_indices") or []:
            try:
                wi = int(wi)
            except Exception:
                continue
            if not (0 <= wi < len(walls)):
                continue
            try:
                member_wid = _wall_id_int(walls[wi])
            except Exception:
                continue
            if int(member_wid) == int(owner_wid):
                continue
            cfg = self._cabezal_by_wall_id.setdefault(
                member_wid, cabezal.default_cabezal_muro_config(),
            )
            ex = cfg.setdefault(extremo, cabezal.default_cabezal_extremo_config())
            for k in copy_keys:
                if k in owner_cfg:
                    ex[k] = _copy_cabezal_extremo_field(k, owner_cfg[k])
            cabezal._normalize_cabezal_extremo_layers(ex)

    def _cabezal_stack_index_for_wid(self, wid):
        for i, w in enumerate(getattr(self, u"walls_ordered", []) or []):
            try:
                if _wall_id_int(w) == int(wid):
                    return i
            except Exception:
                pass
        return -1

    def _cabezal_row_is_single_wall_tramo(self, wid):
        """True si algún tramo (Inicio o Final) contiene solo este muro."""
        if cabezal is None:
            return False
        cache = getattr(self, "_single_tramo_row_cache", None)
        if cache is None:
            self._single_tramo_row_cache = {}
            cache = self._single_tramo_row_cache
        try:
            wid_i = int(wid)
        except Exception:
            wid_i = wid
        if wid_i in cache:
            return cache[wid_i]
        si = self._cabezal_stack_index_for_wid(wid)
        if si < 0:
            cache[wid_i] = False
            return False
        result = False
        for ex in cabezal.CABEZAL_EXTREMOS:
            for seg in self._cabezal_segments_for_extremo(ex):
                wis = seg.get(u"wall_indices") or []
                if len(wis) == 1:
                    try:
                        if int(wis[0]) == si:
                            result = True
                            break
                    except Exception:
                        pass
            if result:
                break
        cache[wid_i] = result
        return result

    def _row_height_for_wall_px(self, wid):
        elev_min = float(getattr(self, "_ELEVATION_ROW_MIN_PX", 192.0))
        if self._is_mallas_mode() or cabezal is None:
            return elev_min
        pie = float(getattr(self, "_CABEZAL_PIE_SELECTOR_RESERVE_PX", 22.0))
        mesh_body = float(getattr(self, "_MESH_ELEV_ROW_BODY_MIN_PX", 86.0))
        h = max(elev_min + pie, mesh_body + pie)
        if self._cabezal_row_is_single_wall_tramo(wid):
            h += float(getattr(self, "_CABEZAL_SINGLE_TRAMO_ROW_EXTRA_PX", 18.0))
        return h

    def _compute_row_heights(self):
        od = getattr(self, "_walls_display_order", []) or []
        if not od:
            return []
        return [
            self._row_height_for_wall_px(_wall_id_int(w)) for w in od
        ]

    def _ruler_col_px(self):
        if self._uses_cabezal_panels():
            return float(self._LEVEL_RULER_COL_PX)
        return 0.0

    def _layout_content_width_px(self):
        return (
            self._ruler_col_px()
            + float(self._preview_col_px)
            + (12.0 if self._mesh_col_px > 0.0 else 0.0)
            + float(self._mesh_col_px)
            + float(getattr(self, "_WINDOW_FRAME_PAD_PX", 40.0))
        )

    def _fit_window_to_content(self):
        if self._win is None:
            return
        try:
            from System import Double
            from System.Windows import ResizeMode

            w = float(self._layout_content_width_px())
            self._win.MinWidth = w
            if self._uses_cabezal_panels():
                self._win.ResizeMode = ResizeMode.CanResize
                self._win.MaxWidth = Double.PositiveInfinity
            else:
                self._win.Width = w
                self._win.MaxWidth = w
        except Exception:
            pass

    def _format_wall_row_label(self, wall, row_index, n_total):
        try:
            wid = _wall_id_int(wall)
        except Exception:
            wid = 0
        try:
            n_total = max(1, int(n_total))
            idx_inf = n_total - int(row_index)
        except Exception:
            idx_inf = int(row_index) + 1
        esp_txt = u"?"
        try:
            esp_mm = geo.obtener_espesor_muro_mm_approx(wall)
            if esp_mm is not None:
                esp_txt = u"{0:.0f}".format(float(esp_mm))
        except Exception:
            pass
        return u"#{0} · Id {1} · e={2} mm".format(idx_inf, wid, esp_txt)

    def _grid_content_width_px(self):
        return (
            self._ruler_col_px()
            + float(self._preview_col_px)
            + (12.0 if self._mesh_col_px > 0.0 else 0.0)
            + float(self._mesh_col_px)
        )

    def _apply_grid_content_width(self, grid):
        if grid is None:
            return
        try:
            from System.Windows import HorizontalAlignment

            w = float(self._grid_content_width_px())
            grid.Width = w
            grid.MinWidth = w
            grid.MaxWidth = w
            grid.HorizontalAlignment = HorizontalAlignment.Center
        except Exception:
            pass

    def _build_column_headers_panel(self, fg_lo, sep_br):
        from System.Windows.Controls import Border, Grid, TextBlock, ColumnDefinition, StackPanel
        from System.Windows import (
            GridLength,
            GridUnitType,
            Thickness,
            FontWeights,
            HorizontalAlignment,
            VerticalAlignment,
            TextWrapping,
        )
        from System.Windows.Media import SolidColorBrush, Color

        root = self._win.FindName("GrdColumnHeaders")
        if root is None:
            return
        root.Children.Clear()
        root.ColumnDefinitions.Clear()
        self._apply_grid_content_width(root)

        try:
            hdr_border = self._win.FindName("BdrColumnHeaders")
            if hdr_border is not None and self._uses_cabezal_panels():
                hdr_border.Margin = Thickness(0, 0, 18, 0)
            elif hdr_border is not None:
                hdr_border.Margin = Thickness(0)
        except Exception:
            pass

        ruler_px = self._ruler_col_px()
        col_offset = 0
        self._ruler_hdr_text = None
        if ruler_px > 0:
            cd_ruler = ColumnDefinition()
            cd_ruler.Width = GridLength(ruler_px, GridUnitType.Pixel)
            root.ColumnDefinitions.Add(cd_ruler)
            col_offset = 1

        cd0 = ColumnDefinition()
        cd0.Width = GridLength(float(self._preview_col_px), GridUnitType.Pixel)
        cd1 = ColumnDefinition()
        cd1.Width = GridLength(12.0, GridUnitType.Pixel)
        cd2 = ColumnDefinition()
        cd2.Width = GridLength(float(self._mesh_col_px), GridUnitType.Pixel)
        root.ColumnDefinitions.Add(cd0)
        if self._mesh_col_px > 0.0:
            root.ColumnDefinitions.Add(cd1)
            root.ColumnDefinitions.Add(cd2)

        hdr_wrap = Border()
        hdr_wrap.Padding = Thickness(0, 0, 0, 0)
        hdr_wrap.BorderBrush = sep_br
        hdr_wrap.BorderThickness = Thickness(0)
        Grid.SetColumn(hdr_wrap, 0)
        total_cols = col_offset + (3 if self._mesh_col_px > 0.0 else 1)
        Grid.SetColumnSpan(hdr_wrap, total_cols)

        outer = Grid()
        if ruler_px > 0:
            cd_ruler_hdr = ColumnDefinition()
            cd_ruler_hdr.Width = GridLength(ruler_px, GridUnitType.Pixel)
            outer.ColumnDefinitions.Add(cd_ruler_hdr)
        cd_a = ColumnDefinition()
        cd_a.Width = GridLength(float(self._preview_col_px), GridUnitType.Pixel)
        outer.ColumnDefinitions.Add(cd_a)
        if self._mesh_col_px > 0.0:
            cd_b = ColumnDefinition()
            cd_b.Width = GridLength(12.0, GridUnitType.Pixel)
            cd_c = ColumnDefinition()
            cd_c.Width = GridLength(float(self._mesh_col_px), GridUnitType.Pixel)
            outer.ColumnDefinitions.Add(cd_b)
            outer.ColumnDefinitions.Add(cd_c)

        acc = SolidColorBrush(Color.FromRgb(126, 184, 208))

        def _hdr_tb(text, col, size=10.5, parent=None):
            tb = TextBlock()
            tb.Text = text.upper()
            tb.Foreground = acc
            tb.FontSize = size
            tb.FontWeight = FontWeights.SemiBold
            tb.HorizontalAlignment = HorizontalAlignment.Center
            tb.VerticalAlignment = VerticalAlignment.Center
            tb.Margin = Thickness(0, 6, 0, 8)
            if parent is not None:
                Grid.SetColumn(tb, col)
                parent.Children.Add(tb)
            return tb

        _outer_col_prev = col_offset

        if self._uses_cabezal_panels():
            elev_gap = float(getattr(self, "_CABEZAL_ELEV_GAP_PX", 8.0))
            prev_hdr = Grid()
            if ruler_px > 0:
                cd_nivel = ColumnDefinition()
                cd_nivel.Width = GridLength(ruler_px, GridUnitType.Pixel)
                prev_hdr.ColumnDefinitions.Add(cd_nivel)
            cd_l = ColumnDefinition()
            cd_l.Width = GridLength(float(self._cabezal_cap_col_px()), GridUnitType.Pixel)
            cd_gl = ColumnDefinition()
            cd_gl.Width = GridLength(elev_gap, GridUnitType.Pixel)
            cd_m = ColumnDefinition()
            cd_m.Width = GridLength(1.0, GridUnitType.Star)
            cd_gr = ColumnDefinition()
            cd_gr.Width = GridLength(elev_gap, GridUnitType.Pixel)
            cd_r = ColumnDefinition()
            cd_r.Width = GridLength(float(self._cabezal_cap_col_px()), GridUnitType.Pixel)
            prev_hdr.ColumnDefinitions.Add(cd_l)
            prev_hdr.ColumnDefinitions.Add(cd_gl)
            prev_hdr.ColumnDefinitions.Add(cd_m)
            prev_hdr.ColumnDefinitions.Add(cd_gr)
            prev_hdr.ColumnDefinitions.Add(cd_r)
            Grid.SetColumn(prev_hdr, 0)
            Grid.SetColumnSpan(prev_hdr, total_cols)

            if ruler_px > 0:
                _hdr_tb(u"Nivel", 0, parent=prev_hdr)

            _hdr_lbl_izq = u"Final"
            _hdr_lbl_der = u"Inicio"
            if self._walls_display_order:
                _w0 = self._walls_display_order[0]
                _ex0_izq, _ex0_der = self._cabezal_extremos_lados_wall(
                    _wall_id_int(_w0), 0,
                )
                _hdr_lbl_izq = (
                    u"Inicio"
                    if _ex0_izq == CABEZAL_EXTREMO_INICIO
                    else u"Final"
                )
                _hdr_lbl_der = (
                    u"Inicio"
                    if _ex0_der == CABEZAL_EXTREMO_INICIO
                    else u"Final"
                )
            _hdr_tb(_hdr_lbl_izq, col_offset + 0, parent=prev_hdr)
            _elev_hdr_txt = (
                u"Elevación muros"
                if not self._is_unificado_mode()
                else u"Elevación · malla ext.+int."
            )
            _hdr_tb(_elev_hdr_txt, col_offset + 2, parent=prev_hdr)
            _hdr_tb(_hdr_lbl_der, col_offset + 4, parent=prev_hdr)
            _hdr_div_br = SolidColorBrush(Color.FromRgb(33, 70, 92))
            _div_cols = [col_offset + 1, col_offset + 3]
            if ruler_px > 0:
                _div_cols.insert(0, 0)
            for _hd_col in _div_cols:
                _hd = Border()
                _hd.Width = 1.0
                _hd.Background = _hdr_div_br
                if _hd_col == 0:
                    _hd.HorizontalAlignment = HorizontalAlignment.Right
                else:
                    _hd.HorizontalAlignment = HorizontalAlignment.Center
                _hd.VerticalAlignment = VerticalAlignment.Stretch
                Grid.SetColumn(_hd, _hd_col)
                prev_hdr.Children.Add(_hd)
            _hdr_line = Border()
            _hdr_line.Height = 1.0
            _hdr_line.Background = _hdr_div_br
            _hdr_line.VerticalAlignment = VerticalAlignment.Bottom
            _hdr_line.HorizontalAlignment = HorizontalAlignment.Stretch
            Grid.SetColumn(_hdr_line, 0)
            _hdr_line_span = (col_offset + 5) if ruler_px > 0 else 5
            Grid.SetColumnSpan(_hdr_line, _hdr_line_span)
            prev_hdr.Children.Add(_hdr_line)
            outer.Children.Add(prev_hdr)
            hint_txt = u""
        else:
            prev_hdr = TextBlock()
            prev_hdr.Text = u"Sección / elevación"
            prev_hdr.Foreground = acc
            prev_hdr.FontSize = 11.0
            prev_hdr.FontWeight = FontWeights.SemiBold
            prev_hdr.VerticalAlignment = VerticalAlignment.Center
            Grid.SetColumn(prev_hdr, _outer_col_prev)
            outer.Children.Add(prev_hdr)
            mesh_hdr = TextBlock()
            mesh_hdr.Text = u"Malla (ext. + int.)"
            mesh_hdr.Foreground = acc
            mesh_hdr.FontSize = 11.0
            mesh_hdr.FontWeight = FontWeights.SemiBold
            mesh_hdr.VerticalAlignment = VerticalAlignment.Center
            mesh_hdr.Margin = Thickness(4, 0, 0, 0)
            Grid.SetColumn(mesh_hdr, _outer_col_prev + 2)
            outer.Children.Add(mesh_hdr)
            hint_txt = (
                u"Configura mallas por tramo. Mallas → todos copia el tramo superior."
            )

        hint = TextBlock()
        hint.Text = hint_txt
        hint.Foreground = fg_lo
        hint.FontSize = 9.0
        hint.TextWrapping = TextWrapping.Wrap
        hint.Margin = Thickness(0, 6, 0, 0)

        shell = StackPanel()
        shell.Children.Add(outer)
        shell.Children.Add(hint)
        hdr_wrap.Child = shell
        root.Children.Add(hdr_wrap)

    def _mesh_control_keys(self):
        if self._mesh_modo_tradicional():
            return (
                u"ct_md", u"ct_ms", u"ct_id", u"ct_is",
            )
        return (
            u"cex_md", u"cex_ms", u"cex_id", u"cex_is",
            u"cix_md", u"cix_ms", u"cix_id", u"cix_is",
        )

    def _copy_combo_like(self, src_cb, dst_cb):
        if src_cb is None or dst_cb is None:
            return
        try:
            si = getattr(src_cb, "SelectedItem", None)
            if si is not None:
                dst_cb.SelectedItem = si
                return
        except Exception:
            pass
        try:
            txt = getattr(src_cb, "Text", None)
            if txt is not None:
                dst_cb.Text = txt
        except Exception:
            pass

    def _push_cabezal_config_to_ui(self, wid, redistribute=True):
        if cabezal is None:
            return
        cfg = self._cabezal_by_wall_id.get(wid) or {}
        for ex in cabezal.CABEZAL_EXTREMOS:
            ui = self._cabezal_ui_ext(wid, ex)
            if ui.get(u"capas_value_tb") is None:
                continue
            ex_cfg = cfg.get(ex) or cabezal.default_cabezal_extremo_config()
            cabezal._normalize_cabezal_extremo_layers(ex_cfg)
            try:
                n = int(ex_cfg.get(u"n_capas", len(ex_cfg.get(u"layers") or [])))
            except Exception:
                n = len(ex_cfg.get(u"layers") or []) or 2
            n = max(1, min(cabezal.CABEZAL_MAX_CAPAS, n))
            ui[u"layer_steppers"] = []
            self._suppress_cabezal_stepper = True
            try:
                cap_step = ui.get(u"capas_stepper")
                if cap_step is not None:
                    cap_step[u"set_value"](n)
                else:
                    tb_cap = ui.get(u"capas_value_tb")
                    if tb_cap is not None:
                        tb_cap.Text = str(n)
            finally:
                self._suppress_cabezal_stepper = False
            self._rebuild_cabezal_layer_sliders(wid, ex, redistribute=redistribute)
            self._refresh_cabezal_confinement_combo(wid, ex)
            self._apply_cabezal_armado_ui_state(wid, ex)
            self._request_cabezal_preview_refresh(wid, ex)

    def _cabezal_confinement_options_for_ui(self, wid, extremo):
        if cabezal is None:
            return []
        n = self._cabezal_n_capas_from_ui(wid, extremo)
        return cabezal.cabezal_confinement_options(n)

    def _on_cabezal_confinement_changed(self, wid, extremo):
        """Sync + redibujar preview tras cambio real en el combo de confinamiento."""
        if cabezal is None:
            return
        try:
            self._sync_cabezal_extremo_from_ui(wid, extremo)
        except Exception:
            pass
        self._request_cabezal_preview_refresh(wid, extremo)

    def _refresh_cabezal_confinement_combo(self, wid, extremo):
        if cabezal is None:
            return
        ui = self._cabezal_ui_ext(wid, extremo)
        cb = ui.get(u"confinement_cb")
        if cb is None:
            return
        cfg = self._cabezal_by_wall_id.setdefault(wid, {})
        ex_cfg = cfg.setdefault(extremo, cabezal.default_cabezal_extremo_config())
        n_capas = self._cabezal_n_capas_from_ui(wid, extremo)
        conf = cabezal.normalize_cabezal_confinement(
            ex_cfg.get(u"confinement"),
            n_capas,
        )
        ex_cfg[u"confinement"] = conf
        opts = self._cabezal_confinement_options_for_ui(wid, extremo)
        ui[u"confinement_options"] = opts
        self._suppress_cabezal_confinement_cb = True
        try:
            try:
                cb.Items.Clear()
                for _val, label in opts:
                    cb.Items.Add(label)
            except Exception:
                pass
            self._set_cabezal_confinement_combo_value(
                cb, wid, extremo, conf.get(u"type"),
            )
            try:
                if cabezal.cabezal_confinement_scenario_applies(n_capas):
                    cb.ToolTip = None
                else:
                    cb.ToolTip = (
                        u"Tipo 1 y Tipo 2 solo aplican con 2 a 6 capas "
                        u"(escenario actual)."
                    )
            except Exception:
                pass
        finally:
            self._suppress_cabezal_confinement_cb = False
        try:
            self._request_cabezal_preview_refresh(wid, extremo)
        except Exception:
            pass

    def _match_cabezal_confinement_label(self, txt, fresh_opts):
        """Resuelve etiqueta visible del combo → valor ``confinement.type``."""
        if cabezal is None:
            return None
        try:
            txt = unicode(txt or u"").strip()
        except Exception:
            txt = u""
        if not txt:
            return None
        if txt == u"Tipo 1" or u"Traba capa" in txt or u"traba capa" in txt.lower():
            return cabezal.CABEZAL_CONFINEMENT_TIE_LAYER_1
        if txt == u"Tipo 2" or u"\u00edndice 0" in txt or u"0 y 1" in txt:
            return cabezal.CABEZAL_CONFINEMENT_PERIMETER_0_1
        if txt == u"Sin confinamiento":
            return cabezal.CABEZAL_CONFINEMENT_NONE
        for val, lbl in fresh_opts or []:
            try:
                if lbl == txt or txt in lbl or lbl.startswith(txt):
                    return val
            except Exception:
                pass
        return None

    def _read_cabezal_confinement_combo(self, wid, extremo):
        if cabezal is None:
            return u"none"
        owner_wid = self._cabezal_owner_wid_for(wid, extremo)
        ui = self._cabezal_ui_ext(owner_wid, extremo)
        cb = ui.get(u"confinement_cb")
        fresh_opts = self._cabezal_confinement_options_for_ui(owner_wid, extremo)
        if cb is not None:
            for txt_src in (
                getattr(cb, u"SelectedItem", None),
                getattr(cb, u"Text", None),
            ):
                matched = self._match_cabezal_confinement_label(txt_src, fresh_opts)
                if matched:
                    return matched
            try:
                idx = int(cb.SelectedIndex)
            except Exception:
                idx = -1
            try:
                n_items = int(cb.Items.Count)
            except Exception:
                n_items = 0
            if n_items > 0 and 0 <= idx < n_items:
                try:
                    label = unicode(cb.Items[idx])
                except Exception:
                    label = u""
                matched = self._match_cabezal_confinement_label(label, fresh_opts)
                if matched:
                    return matched
                if idx < len(fresh_opts):
                    return fresh_opts[idx][0]
        if cb is None:
            cfg = self._cabezal_by_wall_id.get(owner_wid) or {}
            ex_cfg = cfg.get(extremo) or {}
            return cabezal.normalize_cabezal_confinement(
                ex_cfg.get(u"confinement"),
                self._cabezal_n_capas_from_ui(owner_wid, extremo),
            ).get(u"type")
        opts = ui.get(u"confinement_options") or fresh_opts
        try:
            idx = int(cb.SelectedIndex)
            if 0 <= idx < len(opts):
                return opts[idx][0]
        except Exception:
            pass
        return cabezal.CABEZAL_CONFINEMENT_NONE

    def _cabezal_combo_shows_perimeter_confinement(self, wid, extremo):
        """True si el combo UI indica estribo perimetral capas 0–1."""
        if cabezal is None:
            return False
        try:
            conf = self._read_cabezal_confinement_combo(wid, extremo)
            return cabezal.cabezal_confinement_is_perimeter(conf)
        except Exception:
            return False

    def _cabezal_combo_shows_tie_layer_1(self, wid, extremo):
        """True si el combo UI indica traba capa [1]."""
        if cabezal is None:
            return False
        try:
            conf = self._read_cabezal_confinement_combo(wid, extremo)
            return cabezal.cabezal_confinement_is_tie_layer_1(conf)
        except Exception:
            return False

    def _cabezal_confinement_is_active_preview(self, wid, extremo):
        """True si el preview debe dibujar confinamiento (estribo o traba)."""
        if cabezal is None:
            return False
        n_capas = max(
            self._cabezal_n_capas_from_ui(wid, extremo),
            len(self._cabezal_layers_live_from_ui(wid, extremo) or []),
        )
        if n_capas < 2:
            return False
        if not cabezal.cabezal_confinement_scenario_applies(n_capas):
            return False
        conf_type = self._cabezal_effective_confinement_type_for_preview(wid, extremo)
        try:
            return cabezal.cabezal_confinement_has_lote_z(conf_type)
        except Exception:
            return (
                conf_type == cabezal.CABEZAL_CONFINEMENT_PERIMETER_0_1
                or conf_type == cabezal.CABEZAL_CONFINEMENT_TIE_LAYER_1
            )

    def _cabezal_stirrup_segments_for_preview(self, wid, extremo, layout):
        conf_type = self._cabezal_effective_confinement_type_for_preview(wid, extremo)
        try:
            if not cabezal.cabezal_confinement_is_perimeter(conf_type):
                return None
        except Exception:
            if conf_type != cabezal.CABEZAL_CONFINEMENT_PERIMETER_0_1:
                return None
        segs = layout.get(u"stirrup_segments")
        if segs:
            return segs
        dots = layout.get(u"dots") or []
        if not dots:
            return None
        stirrup_idx = layout.get(u"stirrup_layer_indices") or [0, 1]
        try:
            rect = cabezal.cabezal_stirrup_preview_rect(
                dots, stirrup_idx, pad_frac=0.055,
            )
            if not rect:
                return None
            return cabezal.cabezal_stirrup_preview_segments(rect)
        except Exception:
            return None

    def _cabezal_effective_confinement_type_for_preview(self, wid, extremo):
        """Tipo de confinamiento para layout y canvas (combo tiene prioridad)."""
        if cabezal is None:
            return u"none"
        n_capas = max(
            self._cabezal_n_capas_from_ui(wid, extremo),
            len(self._cabezal_layers_live_from_ui(wid, extremo) or []),
        )
        conf_type = self._read_cabezal_confinement_combo(wid, extremo)
        try:
            if cabezal.cabezal_confinement_is_perimeter(conf_type):
                return conf_type
            if cabezal.cabezal_confinement_is_tie_layer_1(conf_type):
                return conf_type
        except Exception:
            if conf_type in (
                cabezal.CABEZAL_CONFINEMENT_PERIMETER_0_1,
                cabezal.CABEZAL_CONFINEMENT_TIE_LAYER_1,
            ):
                return conf_type
        cfg = self._cabezal_by_wall_id.get(wid) or {}
        ex_cfg = cfg.get(extremo) or {}
        norm = cabezal.normalize_cabezal_confinement(
            ex_cfg.get(u"confinement"),
            n_capas,
        )
        return norm.get(u"type") or cabezal.CABEZAL_CONFINEMENT_NONE

    def _cabezal_confinement_type_for_preview(self, wid, extremo):
        """Alias: tipo de confinamiento activo para el canvas."""
        return self._cabezal_effective_confinement_type_for_preview(wid, extremo)

    def _draw_cabezal_stirrup_overlay_preview(
        self, canv, layout, wid, extremo, _px, conf_type=None,
        bulk_preview=False,
    ):
        """
        Estribo perimetral capas 0–1 — ``Rectangle`` verde (estilo Armado Columnas).
        """
        from System.Windows.Controls import Canvas
        from System.Windows.Shapes import Rectangle
        from System.Windows.Media import SolidColorBrush, Color

        if canv is None or cabezal is None:
            return
        if conf_type is None:
            conf_type = self._cabezal_effective_confinement_type_for_preview(wid, extremo)
        show_stirrup = False
        try:
            show_stirrup = cabezal.cabezal_confinement_is_perimeter(conf_type)
        except Exception:
            show_stirrup = conf_type == cabezal.CABEZAL_CONFINEMENT_PERIMETER_0_1
        if not show_stirrup and not bulk_preview:
            show_stirrup = self._cabezal_combo_shows_perimeter_confinement(wid, extremo)
        if not show_stirrup:
            return

        dots = layout.get(u"dots") or []
        stirrup_idx = layout.get(u"stirrup_layer_indices")
        if stirrup_idx:
            idx_set = set(int(i) for i in stirrup_idx)
        elif bulk_preview:
            return
        else:
            idx_set = set((0, 1))
        subset = [
            d for d in dots
            if int(d.get(u"layer_index", int(d.get(u"layer", 1)) - 1)) in idx_set
        ]
        if len(subset) < 2:
            subset = list(dots)
        if len(subset) < 2:
            return

        bar_r = 3.5
        stir_thick = 2.2
        # Envolvente en px: tangente exterior de las barras + mitad del trazo del estribo.
        margin_px = bar_r + stir_thick * 0.5 + 1.5

        cxs = []
        cys = []
        for d in subset:
            try:
                cx, cy = _px(float(d[u"fx"]), float(d[u"fy"]))
            except Exception:
                continue
            cxs.append(cx)
            cys.append(cy)
        if len(cxs) < 2:
            return

        left = min(cxs) - margin_px
        top = min(cys) - margin_px
        rw = max(6.0, max(cxs) - min(cxs) + 2.0 * margin_px)
        rh = max(6.0, max(cys) - min(cys) + 2.0 * margin_px)

        br_green = SolidColorBrush(Color.FromRgb(0, 180, 80))

        r_el = Rectangle()
        r_el.Width = rw
        r_el.Height = rh
        r_el.Stroke = br_green
        r_el.StrokeThickness = stir_thick
        try:
            r_el.Fill = None
        except Exception:
            pass
        Canvas.SetLeft(r_el, left)
        Canvas.SetTop(r_el, top)
        self._canvas_set_zindex(r_el, 21)
        canv.Children.Add(r_el)

    def _draw_cabezal_tie_overlay_preview(
        self, canv, layout, wid, extremo, _px, conf_type=None,
        bulk_preview=False,
    ):
        """Trabas — líneas naranjas (tang. interior + empalmes por barra)."""
        from System.Windows.Controls import Canvas
        from System.Windows.Shapes import Line
        from System.Windows.Media import SolidColorBrush, Color

        if canv is None or cabezal is None:
            return

        ties = list(layout.get(u"tie_previews") or [])
        if not ties:
            one = layout.get(u"tie_preview")
            if one:
                ties = [one]
        if not ties:
            if conf_type is None:
                conf_type = self._cabezal_effective_confinement_type_for_preview(
                    wid, extremo,
                )
            show_tie = False
            try:
                show_tie = cabezal.cabezal_confinement_is_tie_layer_1(conf_type)
            except Exception:
                show_tie = conf_type == cabezal.CABEZAL_CONFINEMENT_TIE_LAYER_1
            if not show_tie and not bulk_preview:
                show_tie = self._cabezal_combo_shows_tie_layer_1(wid, extremo)
            if not show_tie:
                return
            dots = layout.get(u"dots") or []
            if not dots:
                return
            tie_li = cabezal.CABEZAL_TIE_LAYER_INDEX
            if bulk_preview:
                tie_layers = layout.get(u"tie_layer_indices") or []
                if tie_layers:
                    tie_li = int(tie_layers[0])
                else:
                    return
            try:
                tie = cabezal.cabezal_tie_preview_geometry(
                    dots,
                    layer_index=tie_li,
                    inner_y0=layout.get(u"inner_y0"),
                    inner_h=layout.get(u"inner_h"),
                )
            except Exception:
                tie = None
            if tie:
                ties = [tie]
        if not ties:
            return

        br_orange = SolidColorBrush(Color.FromRgb(240, 120, 40))
        tie_thick = 1.9

        def _add_seg(seg):
            if not seg or len(seg) < 2:
                return
            (fx0, fy0), (fx1, fy1) = seg[0], seg[1]
            x0, y0 = _px(float(fx0), float(fy0))
            x1, y1 = _px(float(fx1), float(fy1))
            ln = Line()
            ln.X1 = x0
            ln.Y1 = y0
            ln.X2 = x1
            ln.Y2 = y1
            ln.Stroke = br_orange
            ln.StrokeThickness = tie_thick
            self._canvas_set_zindex(ln, 22)
            canv.Children.Add(ln)

        for tie in ties:
            for seg in tie.get(u"segments") or []:
                _add_seg(seg)
            if not tie.get(u"segments"):
                for key in (u"leg", u"top_hook", u"bottom_hook"):
                    _add_seg(tie.get(key))
                for seg in tie.get(u"bar_grips") or []:
                    _add_seg(seg)

    def _set_cabezal_confinement_combo_value(self, cb, wid, extremo, conf_type):
        if cb is None or cabezal is None:
            return
        ui = self._cabezal_ui_ext(wid, extremo)
        opts = ui.get(u"confinement_options") or self._cabezal_confinement_options_for_ui(wid, extremo)
        target = conf_type or cabezal.CABEZAL_CONFINEMENT_NONE
        for i, (val, _lbl) in enumerate(opts):
            if val == target:
                try:
                    cb.SelectedIndex = i
                except Exception:
                    pass
                return
        try:
            cb.SelectedIndex = 0
        except Exception:
            pass

    def _apply_cabezal_to_all_from_first(self):
        if cabezal is None:
            return
        od = getattr(self, "_walls_display_order", []) or []
        if len(od) < 2:
            self._set_estado(u"Solo hay un tramo; nada que copiar.")
            return
        src_wid = _wall_id_int(od[0])
        self._sync_cabezal_from_ui(src_wid)
        src_cfg = self._cabezal_by_wall_id.get(src_wid) or {}
        armado_keys = (
            u"layers",
            u"n_capas",
            u"confinement",
            u"bar_type_id",
            u"conf_bar_type_id",
            u"segment_bar_type_ids",
            u"armado_activo",
            u"layer_spacing_mm",
            u"encuentro_tipo",
            u"vecino_wall_id",
            u"espesor_detectado_mm",
            u"espesor_seleccionado_mm",
            u"pitch_equitativo_mm",
            u"sic_encuentro_key",
            u"sic_encuentro_total",
            u"sic_encuentro_diam_mm",
            u"sic_basis",
        )
        n_tramos = 0
        for ex in cabezal.CABEZAL_EXTREMOS:
            src_ex = _copy_cabezal_extremo_config(src_cfg.get(ex) or {})
            owners = self._cabezal_walls_for_extremo(ex)
            for wid, _ri, _w in owners[1:]:
                cfg = self._cabezal_by_wall_id.setdefault(
                    wid, cabezal.default_cabezal_muro_config(),
                )
                dst_ex = cfg.setdefault(
                    ex, cabezal.default_cabezal_extremo_config(),
                )
                for k in armado_keys:
                    if k in src_ex:
                        dst_ex[k] = _copy_cabezal_extremo_field(k, src_ex[k])
                cabezal._normalize_cabezal_extremo_layers(dst_ex)
                self._cabezal_propagate_segment_armado_from_owner(ex, wid)
                self._push_cabezal_config_to_ui(wid, redistribute=False)
                n_tramos += 1
        self._redistribute_row_heights_and_redraw()
        lbl = self._format_wall_row_label(od[0], 0, len(od))
        self._set_estado(
            u"Armado cabezal copiado desde {0} a {1} tramo(s) restante(s).".format(
                lbl, n_tramos,
            ),
        )

    def _apply_mallas_to_all_from_first(self):
        od = getattr(self, "_walls_display_order", []) or []
        if len(od) < 2:
            self._set_estado(u"Solo hay un tramo; nada que copiar.")
            return
        src_wid = _wall_id_int(od[0])
        src_ctr = self._controls_by_wall_id.get(src_wid)
        if not src_ctr:
            return
        keys = self._mesh_control_keys()
        n = 0
        for w in od[1:]:
            wid = _wall_id_int(w)
            dst_ctr = self._controls_by_wall_id.get(wid)
            if not dst_ctr:
                continue
            for k in keys:
                self._copy_combo_like(src_ctr.get(k), dst_ctr.get(k))
            n += 1
        lbl = self._format_wall_row_label(od[0], 0, len(od))
        self._set_estado(
            u"Mallas copiadas desde {0} a {1} tramo(s) inferior(es).".format(lbl, n),
        )
        if self._is_unificado_mode():
            for w in od[1:]:
                try:
                    self._sync_cabezal_confinement_from_malla_wall(
                        _wall_id_int(w), refresh_preview=True,
                    )
                except Exception:
                    pass
            try:
                self._sync_cabezal_confinement_from_malla_wall(
                    _wall_id_int(od[0]), refresh_preview=True,
                )
            except Exception:
                pass

    def _init_cabezal_configs(self):
        if cabezal is None:
            return
        fb = ElementId.InvalidElementId
        fb_conf = ElementId.InvalidElementId
        if self._bar_labels_to_id and self._diam_strings:
            try:
                default_mm = (
                    cabezal.CABEZAL_DEFAULT_BAR_DIAM_MM
                    if cabezal is not None
                    else 12.0
                )
                lab = _cabezal_diam_label_for_mm(self._diam_strings, default_mm)
                fb = self._bar_labels_to_id.get(lab, fb) if lab else fb
                lab_conf = self._diam_strings[0]
                fb_conf = self._bar_labels_to_id.get(lab_conf, fb)
            except Exception:
                pass
        doc = getattr(self, u"doc", None)
        fb_arg = fb if fb != ElementId.InvalidElementId else None
        fb_conf_arg = fb_conf if fb_conf != ElementId.InvalidElementId else None
        for wall in self.walls_ordered:
            wid = _wall_id_int(wall)
            cfg = {}
            for ex in cabezal.CABEZAL_EXTREMOS:
                ex_cfg = None
                if doc is not None and _vec_ext is not None and _cab_enc_l is not None:
                    try:
                        vec = _vec_ext.vecino_principal_encuentro_l(doc, wall, ex)
                    except Exception:
                        vec = None
                    if vec is not None:
                        try:
                            ex_cfg = _cab_enc_l.cabezal_extremo_config_encuentro_l(
                                doc,
                                wall,
                                vec,
                                ex,
                                fallback_bar_type_id=fb_arg,
                                fallback_conf_bar_type_id=fb_conf_arg,
                            )
                        except Exception:
                            ex_cfg = None
                if ex_cfg is None:
                    ex_cfg = cabezal.cabezal_extremo_config_con_sic_longitudinal(
                        doc,
                        wall,
                        fallback_bar_type_id=fb_arg,
                        fallback_conf_bar_type_id=fb_conf_arg,
                    )
                cfg[ex] = ex_cfg
            self._cabezal_by_wall_id[wid] = cfg

    def _cabezal_ui_ext(self, wid, extremo):
        owner_wid = self._cabezal_owner_wid_for(wid, extremo)
        seg = self._cabezal_segment_for_wall(owner_wid, extremo)
        try:
            seg_id = int(seg.get(u"id", 0))
        except Exception:
            seg_id = 0
        by_ex = self._cabezal_ui_by_segment.setdefault(extremo, {})
        if seg_id in by_ex:
            return by_ex[seg_id]
        wall_ui = self._cabezal_ui_by_wall_id.get(owner_wid) or {}
        if extremo in wall_ui and isinstance(wall_ui.get(extremo), dict):
            return wall_ui[extremo]
        wall_ui = self._cabezal_ui_by_wall_id.setdefault(wid, {})
        if extremo not in wall_ui or not isinstance(wall_ui.get(extremo), dict):
            wall_ui[extremo] = {}
        return wall_ui[extremo]

    def _cabezal_extremo_ui_label(self, extremo):
        if cabezal is None:
            return unicode(extremo or u"")
        if extremo == cabezal.CABEZAL_EXTREMO_INICIO:
            return u"Inicio"
        return u"Final"

    def _cabezal_armado_activo_cfg(self, wid, extremo):
        if cabezal is None:
            return True
        cfg = self._cabezal_by_wall_id.get(wid) or {}
        ex_cfg = cfg.get(extremo) or {}
        return cabezal.cabezal_extremo_armado_activo(ex_cfg)

    def _cabezal_extremo_es_encuentro_l(self, wid, extremo):
        """True si el tramo/extremo usa geometría de encuentro L."""
        if cabezal is None:
            return False
        cfg = self._cabezal_by_wall_id.get(wid) or {}
        ex_cfg = cfg.get(extremo) or {}
        return cabezal.cabezal_extremo_es_encuentro_l(ex_cfg)

    def _cabezal_min_capas_for_wall_extremo(self, wid, extremo):
        if self._cabezal_extremo_es_encuentro_preview(wid, extremo):
            return 2
        return cabezal.CABEZAL_MIN_CAPAS if cabezal else 1

    def _cabezal_min_bars_for_wall_extremo(self, wid, extremo):
        if self._cabezal_extremo_es_encuentro_preview(wid, extremo):
            return 2
        return cabezal.CABEZAL_MIN_BARRAS_POR_CAPA if cabezal else 2

    def _refresh_cabezal_encuentro_ui_state(self, wid, extremo):
        """Encuentro L: oculta confinamiento, badge en toolbar y canvas ampliado."""
        from System.Windows import Visibility

        ui = self._cabezal_ui_ext(wid, extremo)
        wall = self._wall_for_integer_id(wid)
        enc_ctx = self._cabezal_preview_encuentro_ctx(wid, wall, extremo)
        is_enc = enc_ctx is not None
        vis_on = Visibility.Visible
        vis_off = Visibility.Collapsed

        for key in (u"confinement_lbl", u"confinement_cb"):
            el = ui.get(key)
            if el is not None:
                try:
                    el.Visibility = vis_off if is_enc else vis_on
                except Exception:
                    pass
        hint = ui.get(u"encuentro_hint_lbl")
        if hint is not None:
            try:
                hint.Visibility = vis_off
            except Exception:
                pass

        if is_enc and cabezal is not None:
            cfg = self._cabezal_by_wall_id.setdefault(wid, {})
            ex_cfg = cfg.setdefault(extremo, cabezal.default_cabezal_extremo_config())
            try:
                n_capas = int(self._cabezal_n_capas_from_ui(wid, extremo))
            except Exception:
                n_capas = 2
            ex_cfg[u"confinement"] = cabezal.normalize_cabezal_confinement(
                {u"type": cabezal.CABEZAL_CONFINEMENT_NONE}, n_capas,
            )

        base = ui.get(u"tramo_toolbar_text_base")
        tramo_text = base
        if is_enc and base:
            enc_tipo = (enc_ctx or {}).get(u"tipo") or u"L"
            tramo_text = u"{0} · Encuentro {1}".format(base, enc_tipo)
        cap_step = ui.get(u"capas_stepper")
        toolbar = ui.get(u"toolbar_grid")
        if toolbar is not None and cap_step is not None:
            try:
                pal = self._cabezal_ui_palette(extremo, u"unit")
                self._fill_cabezal_toolbar_row(
                    toolbar, cap_step.get(u"panel"), pal, tramo_text,
                )
            except Exception:
                pass

        try:
            prev_w, prev_h = self._cabezal_preview_canvas_size_px(wid, extremo)
            cv = ui.get(u"preview_canvas")
            if cv is not None:
                cv.Width = prev_w
                cv.Height = prev_h
                cv.MinWidth = prev_w
                cv.MaxWidth = prev_w
                cv.MinHeight = prev_h
                cv.MaxHeight = prev_h
                ui[u"preview_canvas_w_px"] = prev_w
        except Exception:
            pass

    def _cabezal_bulk_armado_activo(self, extremo):
        bulk = (getattr(self, "_cabezal_bulk_ui", None) or {}).get(extremo) or {}
        chk = bulk.get(u"armado_activo_chk")
        if chk is None:
            return True
        try:
            return bool(chk.IsChecked)
        except Exception:
            return True

    def _set_cabezal_extremo_panel_enabled(self, wid, extremo, enabled):
        ui = self._cabezal_ui_ext(wid, extremo)
        pal = self._cabezal_ui_palette(extremo, u"unit")
        opacity = 1.0 if enabled else float(pal.get(u"disabled_opacity", 0.42))
        targets = [
            ui.get(u"controls_scroll"),
            ui.get(u"split_grid"),
            ui.get(u"confinement_cb"),
            ui.get(u"confinement_lbl"),
            ui.get(u"encuentro_hint_lbl"),
        ]
        cap_step = ui.get(u"capas_stepper") or {}
        targets.append(cap_step.get(u"panel"))
        targets.append(cap_step.get(u"minus_btn"))
        targets.append(cap_step.get(u"plus_btn"))
        targets.append(cap_step.get(u"value_tb"))
        for st in ui.get(u"layer_steppers") or []:
            if isinstance(st, dict):
                targets.append(st.get(u"panel"))
                targets.append(st.get(u"minus_btn"))
                targets.append(st.get(u"plus_btn"))
                targets.append(st.get(u"value_tb"))
        for cb in ui.get(u"layer_diam_cbs") or []:
            targets.append(cb)
        for el in targets:
            if el is None:
                continue
            try:
                el.IsEnabled = bool(enabled)
                el.Opacity = opacity
            except Exception:
                pass

    def _set_cabezal_bulk_panel_enabled(self, extremo, enabled):
        bulk = (getattr(self, "_cabezal_bulk_ui", None) or {}).get(extremo) or {}
        pal = self._cabezal_ui_palette(extremo, u"bulk")
        opacity = 1.0 if enabled else float(pal.get(u"disabled_opacity", 0.42))
        body = bulk.get(u"config_body")
        if body is not None:
            try:
                body.IsEnabled = bool(enabled)
                body.Opacity = opacity
            except Exception:
                pass

    def _apply_cabezal_armado_ui_state(self, wid, extremo):
        """Atenúa el panel por muro según ``armado_activo`` (toggle solo en bulk)."""
        if cabezal is None:
            return
        on = self._cabezal_armado_activo_cfg(wid, extremo)
        self._set_cabezal_extremo_panel_enabled(wid, extremo, on)

    def _cabezal_extremos_lados_wall(self, wid, row_index):
        if cabezal is None:
            return u"inicio", u"fin"
        wall = None
        for w in self._walls_display_order:
            if _wall_id_int(w) == wid:
                wall = w
                break
        self._ensure_layout_cache()
        stacked = getattr(self, "_stacked_layout", None)
        if stacked is not None:
            return geo.cabezal_extremos_en_lados_stacked(wall, row_index, stacked)
        layout = getattr(self, "_preview_layout", None)
        return geo.cabezal_extremos_en_lados_preview(wall, row_index, layout)

    def _sync_cabezal_extremo_from_ui(self, wid, extremo, sync_confinement=True):
        if cabezal is None:
            return
        owner_wid = self._cabezal_owner_wid_for(wid, extremo)
        ui = self._cabezal_ui_ext(owner_wid, extremo)
        cfg = self._cabezal_by_wall_id.get(owner_wid)
        if cfg is None:
            return
        ex_cfg = cfg.setdefault(extremo, cabezal.default_cabezal_extremo_config())
        layers = []
        steppers = ui.get(u"layer_steppers") or []
        diam_cbs = ui.get(u"layer_diam_cbs") or []
        steppers_live = self._cabezal_layer_steppers_are_live(ui)
        n_layers = max(len(steppers), len(diam_cbs)) if steppers_live else 0
        n_active = self._cabezal_n_capas_from_ui(owner_wid, extremo)
        for i in range(n_layers):
            nb = cabezal.CABEZAL_MIN_BARRAS_POR_CAPA
            if i < len(steppers) and steppers[i] is not None:
                nb = self._cabezal_read_value_tb(
                    steppers[i][u"value_tb"],
                    cabezal.CABEZAL_MIN_BARRAS_POR_CAPA,
                    cabezal.CABEZAL_MAX_BARRAS_POR_CAPA,
                    cabezal.CABEZAL_MIN_BARRAS_POR_CAPA,
                )
            bid = self._cabezal_layer_diam_id(owner_wid, extremo, i)
            if i < n_active and i < len(diam_cbs) and diam_cbs[i] is not None:
                try:
                    bid = self._read_diam_combo_id(diam_cbs[i])
                except Exception:
                    pass
            layers.append({u"n_bars": nb, u"bar_type_id": bid})
        if layers:
            ex_cfg[u"layers"] = [
                cabezal._normalize_cabezal_layer_dict(ly, ex_cfg.get(u"bar_type_id"))
                for ly in layers
            ]
            cabezal._normalize_cabezal_extremo_layers(ex_cfg)
        chk = ui.get(u"troceo_por_muro_chk")
        if chk is not None:
            try:
                on = bool(chk.IsChecked)
                auto = self._cabezal_auto_troceo_for_wall(owner_wid, extremo)
                ex_cfg[u"troceo_por_muro_override"] = None if on == auto else on
                ex_cfg[u"troceo_auto_geom"] = auto
                ex_cfg[u"troceo_por_muro"] = on
            except Exception:
                pass
        try:
            ex_cfg[u"n_capas"] = self._cabezal_n_capas_from_ui(owner_wid, extremo)
        except Exception:
            pass
        if sync_confinement:
            conf_cb = ui.get(u"confinement_cb")
            if conf_cb is not None and not self._cabezal_extremo_es_encuentro_preview(owner_wid, extremo):
                try:
                    n_capas = int(ex_cfg.get(u"n_capas", cabezal.CABEZAL_MIN_CAPAS))
                except Exception:
                    n_capas = cabezal.CABEZAL_MIN_CAPAS
                skip_conf = False
                try:
                    if int(conf_cb.SelectedIndex) < 0 and int(conf_cb.Items.Count) == 0:
                        skip_conf = True
                except Exception:
                    pass
                if not skip_conf:
                    conf_val = self._read_cabezal_confinement_combo(owner_wid, extremo)
                    prev_conf = ex_cfg.get(u"confinement") or {}
                    if not isinstance(prev_conf, dict):
                        prev_conf = {}
                    merged = dict(prev_conf)
                    merged[u"type"] = conf_val
                    ex_cfg[u"confinement"] = cabezal.normalize_cabezal_confinement(
                        merged, n_capas,
                    )
        segs = self._cabezal_segments_for_extremo(extremo)
        cabezal._migrate_tramo_to_segment_bar_type_ids(
            ex_cfg, segs, ex_cfg.get(u"bar_type_id"),
        )
        self._cabezal_refresh_encuentro_pitch(owner_wid, extremo)
        self._cabezal_propagate_segment_armado_from_owner(extremo, owner_wid)

    def _migrate_cabezal_cfg_legacy(self, cfg):
        if not cfg or cabezal is None:
            return
        try:
            old_bid = cfg.pop(u"bar_type_id", None)
        except Exception:
            old_bid = None
        for ex in cabezal.CABEZAL_EXTREMOS:
            ex_cfg = cfg.setdefault(ex, cabezal.default_cabezal_extremo_config())
            if old_bid is not None and old_bid != ElementId.InvalidElementId:
                if ex_cfg.get(u"bar_type_id") in (None, ElementId.InvalidElementId):
                    ex_cfg[u"bar_type_id"] = old_bid
            layers = ex_cfg.get(u"layers") or []
            fb = ex_cfg.get(u"bar_type_id")
            if u"segment_bar_type_ids" not in ex_cfg:
                ex_cfg[u"segment_bar_type_ids"] = {}
            segs = self._cabezal_segments_for_extremo(ex)
            cabezal._migrate_tramo_to_segment_bar_type_ids(ex_cfg, segs, fb)
            normalized = []
            for ly in layers:
                normalized.append(
                    cabezal._normalize_cabezal_layer_dict(ly, fb),
                )
            if not normalized:
                normalized = [
                    cabezal.default_cabezal_layer_config(2, fb),
                    cabezal.default_cabezal_layer_config(2, fb),
                ]
            ex_cfg[u"layers"] = normalized
            if u"n_capas" not in ex_cfg:
                try:
                    ex_cfg[u"n_capas"] = max(
                        cabezal.CABEZAL_MIN_CAPAS,
                        min(cabezal.CABEZAL_MAX_CAPAS, len(normalized)),
                    )
                except Exception:
                    ex_cfg[u"n_capas"] = cabezal.CABEZAL_MIN_CAPAS
            cabezal._normalize_cabezal_extremo_layers(ex_cfg)
            if ex_cfg.get(u"conf_bar_type_id") in (None, ElementId.InvalidElementId):
                ex_cfg[u"conf_bar_type_id"] = fb
            if u"confinement" not in ex_cfg:
                n_capas = ex_cfg.get(u"n_capas", cabezal.CABEZAL_MIN_CAPAS)
                ex_cfg[u"confinement"] = cabezal.default_cabezal_confinement_config(n_capas)
            cabezal._normalize_cabezal_extremo_layers(ex_cfg)
        for ex in cabezal.CABEZAL_EXTREMOS:
            ex_cfg = cfg.setdefault(ex, cabezal.default_cabezal_extremo_config())
            if u"troceo_por_muro" not in ex_cfg:
                ex_cfg[u"troceo_por_muro"] = False

    def _sync_cabezal_confinement_from_malla_wall(self, wid, refresh_preview=True):
        """
        Sincroniza solo Ø/@ del confinamiento desde capas longitudinales.

        No relee el combo Tipo 1/2: un cambio de malla no debe alterar el tipo
        de confinamiento elegido en cabezal.
        """
        if cabezal is None:
            return
        if not (self._is_unificado_mode() or self._is_cabezal_mode()):
            return
        synced = set()
        for ex in cabezal.CABEZAL_EXTREMOS:
            owner_wid = self._cabezal_owner_wid_for(wid, ex)
            key = (ex, int(owner_wid))
            if key in synced:
                continue
            synced.add(key)
            cfg = self._cabezal_by_wall_id.get(owner_wid)
            if cfg is None:
                continue
            ex_cfg = cfg.get(ex)
            if ex_cfg:
                fb = cabezal.cabezal_sync_fallback_bar_type_id(
                    ex_cfg, cfg.get(u"bar_type_id"),
                )
                cabezal.cabezal_sync_confinement_from_extremo(
                    ex_cfg, self.doc, fb,
                )
        if refresh_preview:
            for ex in cabezal.CABEZAL_EXTREMOS:
                try:
                    self._request_cabezal_preview_refresh(wid, ex)
                except Exception:
                    pass

    def _sync_cabezal_from_ui(self, wid):
        if cabezal is None:
            return
        cfg = self._cabezal_by_wall_id.get(wid)
        if cfg is not None:
            self._migrate_cabezal_cfg_legacy(cfg)
        for ex in cabezal.CABEZAL_EXTREMOS:
            self._sync_cabezal_extremo_from_ui(wid, ex)
        if self._is_unificado_mode() or self._is_cabezal_mode():
            self._sync_cabezal_confinement_from_malla_wall(wid, refresh_preview=False)

    def _sync_cabezal_from_segment_owners(self, sync_confinement=True):
        """Lee UI de un controlador por tramo y propaga armado a muros del tramo."""
        if cabezal is None:
            return
        seen = set()
        for ex in cabezal.CABEZAL_EXTREMOS:
            for wid, _ri, _w in self._cabezal_walls_for_extremo(ex):
                key = (ex, int(wid))
                if key in seen:
                    continue
                seen.add(key)
                cfg = self._cabezal_by_wall_id.get(wid)
                if cfg is not None:
                    self._migrate_cabezal_cfg_legacy(cfg)
                self._sync_cabezal_extremo_from_ui(
                    wid, ex, sync_confinement=sync_confinement,
                )
        for wall in getattr(self, u"walls_ordered", []) or []:
            try:
                wid = _wall_id_int(wall)
            except Exception:
                continue
            if self._is_unificado_mode() or self._is_cabezal_mode():
                self._sync_cabezal_confinement_from_malla_wall(
                    wid, refresh_preview=False,
                )

    def _cabezal_layers_live_from_ui(self, wid, extremo):
        if cabezal is None:
            return []
        cfg = self._cabezal_by_wall_id.get(
            self._cabezal_owner_wid_for(wid, extremo),
        ) or {}
        ex_cfg = cfg.get(extremo) or {}
        ui = self._cabezal_ui_ext(wid, extremo)
        if not self._cabezal_layer_steppers_are_live(ui):
            return list(cabezal.cabezal_active_layers(ex_cfg))
        layers = []
        steppers = ui.get(u"layer_steppers") or []
        diam_cbs = ui.get(u"layer_diam_cbs") or []
        n_layers = max(len(steppers), len(diam_cbs))
        fb = ex_cfg.get(u"bar_type_id") or ElementId.InvalidElementId
        for i in range(n_layers):
            nb = cabezal.CABEZAL_MIN_BARRAS_POR_CAPA
            if i < len(steppers) and steppers[i] is not None:
                nb = self._cabezal_read_value_tb(
                    steppers[i][u"value_tb"],
                    cabezal.CABEZAL_MIN_BARRAS_POR_CAPA,
                    cabezal.CABEZAL_MAX_BARRAS_POR_CAPA,
                    cabezal.CABEZAL_MIN_BARRAS_POR_CAPA,
                )
            bid = self._cabezal_layer_diam_id(wid, extremo, i)
            if i < len(diam_cbs) and diam_cbs[i] is not None:
                try:
                    bid = self._read_diam_combo_id(diam_cbs[i])
                except Exception:
                    pass
            layers.append(
                cabezal._normalize_cabezal_layer_dict(
                    {u"n_bars": nb, u"bar_type_id": bid}, fb,
                ),
            )
        if layers:
            return layers[: self._cabezal_n_capas_from_ui(wid, extremo)]
        return list(cabezal.cabezal_active_layers(ex_cfg))

    def _refresh_cabezal_preview(self, wid, extremo):
        if cabezal is None:
            return
        ui = self._cabezal_ui_ext(wid, extremo)
        cv = ui.get(u"preview_canvas")
        if cv is None:
            return
        try:
            aw = float(cv.ActualWidth)
            ah = float(cv.ActualHeight)
            if aw < 8.0 or ah < 8.0:
                cw, ch = self._cabezal_preview_canvas_size_px(wid, extremo)
                try:
                    cv.Width = cw
                    cv.Height = ch
                except Exception:
                    pass
        except Exception:
            pass
        wall = None
        for w in self._walls_display_order:
            try:
                if _wall_id_int(w) == wid:
                    wall = w
                    break
            except Exception:
                pass
        if wall is None:
            return
        self._draw_cabezal_preview_canvas(cv, wall, wid, extremo)

    def _cabezal_resumen_texto(self, wid):
        if cabezal is None:
            return u""
        cfg = self._cabezal_by_wall_id.get(wid)
        if not cfg:
            return u""
        parts = []
        id_map = getattr(self, "_bar_id_to_label", {})
        for ex, lbl in (
            (cabezal.CABEZAL_EXTREMO_INICIO, u"Ini"),
            (cabezal.CABEZAL_EXTREMO_FIN, u"Fin"),
        ):
            ex_cfg = cfg.get(ex) or {}
            if not cabezal.cabezal_extremo_armado_activo(ex_cfg):
                parts.append(u"{0}:OFF".format(lbl))
                continue
            ly = cabezal.cabezal_active_layers(ex_cfg)
            if not ly:
                ly = list((ex_cfg or {}).get(u"layers") or [])
            if not ly:
                continue
            seg = self._cabezal_segment_for_wall(wid, ex)
            seg_id = int(seg.get(u"id", 0))
            try:
                n_capas_res = int(ex_cfg.get(u"n_capas", len(ly)))
            except Exception:
                n_capas_res = len(ly)
            nums = u"+".join(str(int(x.get(u"n_bars", 0))) for x in ly)
            diams = []
            for xi in range(n_capas_res):
                dl = u"?"
                try:
                    bid = self._cabezal_layer_diam_id(wid, ex, xi)
                    if bid and bid != ElementId.InvalidElementId:
                        dl = id_map.get(geo._element_id_int(bid), dl)
                except Exception:
                    pass
                try:
                    nb = int(ly[xi].get(u"n_bars", 2))
                except Exception:
                    nb = 2
                diams.append(u"L{0}\u00d8{1}\u00d7{2}".format(
                    xi + 1, dl.replace(u" mm", u""), nb,
                ))
            dtxt = u" ".join(diams) if diams else u""
            parts.append(u"{0}:{1}({2}|S{3})".format(
                lbl, len(ly), nums, seg_id + 1,
            ))
        return u"  ".join(parts)

    def _cabezal_empalme_stack_indices(self, extremo):
        if cabezal is None:
            return []
        walls = getattr(self, "walls_ordered", []) or []
        return cabezal._empalme_stack_indices(
            walls, self._cabezal_by_wall_id, extremo,
        )

    def _cabezal_segments_for_extremo(self, extremo):
        if cabezal is None:
            return []
        cache = getattr(self, "_cabezal_segments_cache", None)
        if cache is None:
            self._cabezal_segments_cache = {}
            cache = self._cabezal_segments_cache
        if extremo in cache:
            return cache[extremo]
        walls = getattr(self, "walls_ordered", []) or []
        segs = cabezal.build_cabezal_segments(
            len(walls), self._cabezal_empalme_stack_indices(extremo),
        )
        cache[extremo] = segs
        return segs

    def _cabezal_segment_for_wall(self, wid, extremo):
        stack_idx = self._cabezal_wall_stack_index(wid)
        segs = self._cabezal_segments_for_extremo(extremo)
        for seg in segs:
            if stack_idx in (seg.get(u"wall_indices") or []):
                return seg
        if segs:
            return segs[0]
        return {u"id": 0, u"owner_index": 0, u"wall_indices": []}

    def _cabezal_wall_stack_index(self, wid):
        walls = getattr(self, "walls_ordered", []) or []
        for i, w in enumerate(walls):
            try:
                if _wall_id_int(w) == wid:
                    return i
            except Exception:
                pass
        return 0

    def _cabezal_layer_diam_id(self, wid, extremo, layer_idx):
        if cabezal is None:
            return ElementId.InvalidElementId
        cfg = self._cabezal_by_wall_id.get(wid) or {}
        ex_cfg = cfg.get(extremo) or {}
        fb = ex_cfg.get(u"bar_type_id") or ElementId.InvalidElementId
        layers = ex_cfg.get(u"layers") or []
        li = int(layer_idx)
        ly = layers[li] if li < len(layers) else {}
        bid = (ly or {}).get(u"bar_type_id")
        if bid is None or bid == ElementId.InvalidElementId:
            bid = fb
        return bid

    def _cabezal_n_tramos_verticales(self, extremo):
        """Tramos verticales UI = 1 + empalmes activos en ese extremo (stack completo)."""
        if cabezal is None:
            return 1
        n_emp = 0
        for w in getattr(self, "_walls_display_order", []) or []:
            try:
                wid = _wall_id_int(w)
            except Exception:
                continue
            if self._cabezal_troceo_por_muro_activo(wid, extremo):
                n_emp += 1
        return cabezal.count_cabezal_tramos_verticales(n_emp)

    def _cabezal_ctrl_metrics(self):
        lbl_w = float(getattr(self, "_CABEZAL_CTRL_LBL_W_PX", 26.0))
        step_w = float(getattr(self, "_CABEZAL_CTRL_STEPPER_W_PX", 52.0))
        diam_w = float(getattr(self, "_CABEZAL_CTRL_DIAM_W_PX", 56.0))
        gap_w = float(getattr(self, "_CABEZAL_CTRL_GAP_PX", 2.0))
        strip_w = lbl_w + step_w + diam_w + gap_w
        row_h = float(getattr(self, "_CABEZAL_LAYER_ROW_PX", 26.0))
        return lbl_w, step_w, diam_w, strip_w, row_h, gap_w

    def _cabezal_ui_palette(self, extremo, context=u"bulk"):
        """Tokens tipográficos y acentos por extremo (bulk + unitario)."""
        from System.Windows.Media import SolidColorBrush, Color

        is_ini = (
            extremo == cabezal.CABEZAL_EXTREMO_INICIO
            if cabezal is not None else True
        )
        if is_ini:
            bar_rgb = (34, 211, 238)
            layer_rgb = (103, 232, 249)
        else:
            bar_rgb = (248, 113, 113)
            layer_rgb = (252, 165, 165)
        disabled_op = 0.42
        text_section = SolidColorBrush(Color.FromRgb(186, 198, 214))
        text_caption = SolidColorBrush(Color.FromRgb(156, 170, 188))
        if context == u"bulk":
            text_section = SolidColorBrush(Color.FromRgb(196, 208, 222))
            text_caption = SolidColorBrush(Color.FromRgb(168, 182, 198))
        return {
            u"text_title": SolidColorBrush(Color.FromRgb(226, 232, 240)),
            u"text_section": text_section,
            u"text_caption": text_caption,
            u"text_control": SolidColorBrush(Color.FromRgb(184, 201, 212)),
            u"text_value_rgb": (238, 246, 250),
            u"layer_active": SolidColorBrush(Color.FromRgb(*layer_rgb)),
            u"layer_inactive": SolidColorBrush(Color.FromRgb(84, 102, 120)),
            u"bar_accent": SolidColorBrush(Color.FromRgb(*bar_rgb)),
            u"bar_accent_rgb": bar_rgb,
            u"sep": SolidColorBrush(Color.FromRgb(33, 70, 92)),
            u"panel_fill": (
                SolidColorBrush(Color.FromRgb(10, 22, 32))
                if context == u"bulk"
                else SolidColorBrush(Color.FromArgb(180, 10, 22, 32))
            ),
            u"canvas_group_bg": SolidColorBrush(Color.FromArgb(140, 7, 16, 24)),
            u"canvas_group_border": SolidColorBrush(Color.FromRgb(33, 70, 92)),
            u"disabled_opacity": disabled_op,
            u"accent_bar_height_px": 3.0 if context == u"bulk" else 2.0,
        }

    def _mesh_ui_palette(self, context=u"bulk"):
        """Tokens visuales mallas (misma familia que cabezal bulk, sin acento verde)."""
        from System.Windows.Media import SolidColorBrush, Color

        accent_rgb = (56, 189, 248)
        return {
            u"panel_fill": SolidColorBrush(Color.FromRgb(10, 22, 32)),
            u"panel_fill_overlay": SolidColorBrush(Color.FromArgb(228, 10, 22, 32)),
            u"sep": SolidColorBrush(Color.FromRgb(33, 70, 92)),
            u"accent": SolidColorBrush(Color.FromRgb(*accent_rgb)),
            u"text_title": SolidColorBrush(Color.FromRgb(226, 232, 240)),
            u"text_section": SolidColorBrush(Color.FromRgb(148, 163, 184)),
            u"text_caption": SolidColorBrush(Color.FromRgb(100, 116, 139)),
            u"text_control": SolidColorBrush(Color.FromRgb(184, 201, 212)),
            u"accent_bar_height_px": 2.0 if context == u"unit" else 3.0,
            u"disabled_opacity": 0.42,
        }

    def _malla_activo_cfg(self, wid):
        try:
            return bool(self._malla_activo_by_wall_id.get(int(wid), True))
        except Exception:
            return True

    def _malla_activo_por_muro_id_dict(self):
        out = {}
        for w in getattr(self, u"walls_ordered", []) or []:
            wid = _wall_id_int(w)
            if wid is None:
                continue
            out[wid] = self._malla_activo_cfg(wid)
        return out

    def _set_malla_activo_all(self, activo, animate=False):
        on = bool(activo)
        for w in getattr(self, "_walls_display_order", []) or []:
            try:
                wid = _wall_id_int(w)
            except Exception:
                continue
            self._malla_activo_by_wall_id[wid] = on
            self._apply_malla_activo_ui_state(wid)
        self._sync_bulk_malla_activo_toggle(on, animate=animate)
        self._set_bulk_mesh_panel_enabled(on)

    def _malla_activo_toggle_hosts(self):
        hosts = []
        for attr in (u"_bulk_malla_activo_chk", u"_header_malla_activo_chk"):
            chk = getattr(self, attr, None)
            if chk is None:
                continue
            hosts.append({
                u"chk": chk,
                u"ui_stub": {
                    u"malla_activo_chk": chk,
                    u"toggle_accent_rgb": (56, 189, 248),
                    u"toggle_mini_parts_key": u"malla_activo_toggle_parts",
                    u"toggle_mini_label": u"Armado malla",
                },
            })
        return hosts

    def _sync_bulk_malla_activo_toggle(self, on, animate=False):
        self._suppress_malla_activo_chk = True
        try:
            for entry in self._malla_activo_toggle_hosts():
                chk = entry.get(u"chk")
                ui_stub = entry.get(u"ui_stub")
                if chk is None:
                    continue
                try:
                    chk.IsChecked = bool(on)
                except Exception:
                    pass
                if ui_stub:
                    try:
                        self._apply_toggle_mini_visual(
                            ui_stub,
                            on,
                            animate=animate,
                            parts_key=u"malla_activo_toggle_parts",
                        )
                    except Exception:
                        pass
        finally:
            self._suppress_malla_activo_chk = False

    def _set_bulk_mesh_panel_enabled(self, enabled):
        pal = self._mesh_ui_palette(u"bulk")
        opacity = 1.0 if enabled else float(pal.get(u"disabled_opacity", 0.42))
        body = getattr(self, u"_bulk_mesh_params_body", None)
        if body is not None:
            try:
                body.IsEnabled = bool(enabled)
                body.Opacity = opacity
            except Exception:
                pass

    def _apply_malla_activo_ui_state(self, wid):
        on = self._malla_activo_cfg(wid)
        pal = self._mesh_ui_palette(u"unit")
        opacity = 1.0 if on else float(pal.get(u"disabled_opacity", 0.42))
        ui = self._malla_ui_by_wall_id.get(wid) or {}
        for el in ui.get(u"panels") or []:
            if el is None:
                continue
            try:
                el.IsEnabled = bool(on)
                el.Opacity = opacity
            except Exception:
                pass
        ctr = self._controls_by_wall_id.get(wid)
        if ctr:
            for k in self._mesh_control_keys():
                cb = ctr.get(k)
                if cb is None:
                    continue
                try:
                    cb.IsEnabled = bool(on)
                    cb.Opacity = opacity
                except Exception:
                    pass

    def _register_malla_ui_targets(self, wid, panels):
        if wid is None:
            return
        entry = self._malla_ui_by_wall_id.setdefault(wid, {})
        existing = entry.get(u"panels") or []
        for p in panels or []:
            if p is not None and p not in existing:
                existing.append(p)
        entry[u"panels"] = existing
        if wid not in self._malla_activo_by_wall_id:
            self._malla_activo_by_wall_id[wid] = True
        self._apply_malla_activo_ui_state(wid)

    def _build_bulk_malla_activo_toggle(self, parent=None, header_mode=False):
        """Toggle «Armado malla» en configurador global (unificado)."""
        from System.Windows.Controls import CheckBox
        from System.Windows import HorizontalAlignment, VerticalAlignment, Thickness

        pal = self._mesh_ui_palette(u"bulk")
        accent_rgb = (56, 189, 248)
        ui_stub = {
            u"toggle_accent_rgb": accent_rgb,
            u"toggle_mini_parts_key": u"malla_activo_toggle_parts",
            u"toggle_mini_label": u"" if header_mode else u"Armado malla",
        }
        chk = CheckBox()
        chk.Margin = Thickness(0, 0, 0, 0 if header_mode else 8)
        chk.Padding = Thickness(0)
        chk.HorizontalAlignment = (
            HorizontalAlignment.Right if header_mode else HorizontalAlignment.Left
        )
        chk.VerticalAlignment = VerticalAlignment.Center
        self._apply_bimtools_toggle_mini(chk)
        self._build_toggle_mini_content(chk, ui_stub)
        chk.ToolTip = (
            u"Activa o desactiva la creación de malla (Area Reinforcement) "
            u"para todos los muros."
        )
        ui_stub[u"malla_activo_chk"] = chk
        if parent is not None:
            parent.Children.Add(chk)
        self._bulk_malla_activo_chk = chk

        def _on_bulk_malla(sender, args, c=chk):
            if getattr(self, "_suppress_malla_activo_chk", False):
                return
            try:
                on = bool(c.IsChecked)
            except Exception:
                on = False
            self._set_malla_activo_all(on, animate=True)

        try:
            from System.Windows import RoutedEventHandler as _REH
            chk.Checked += _REH(_on_bulk_malla)
            chk.Unchecked += _REH(_on_bulk_malla)
        except Exception:
            pass
        return chk

    def _ensure_header_malla_activo_toggle(self):
        """Toggle global en modo solo mallas (sin panel bulk central)."""
        if not self._is_mallas_mode() or self._is_unificado_mode():
            return
        if getattr(self, u"_header_malla_activo_chk", None) is not None:
            return
        from System.Windows.Controls import CheckBox
        from System.Windows import Thickness

        pnl = self._win.FindName(u"PnlModoMuro") if self._win else None
        if pnl is None:
            return
        pal = self._mesh_ui_palette(u"bulk")
        ui_stub = {
            u"toggle_accent_rgb": (56, 189, 248),
            u"toggle_mini_parts_key": u"malla_activo_toggle_parts",
            u"toggle_mini_label": u"Armado malla",
        }
        chk = CheckBox()
        chk.Margin = Thickness(24, 0, 0, 0)
        chk.Padding = Thickness(0)
        self._apply_bimtools_toggle_mini(chk)
        self._build_toggle_mini_content(chk, ui_stub)
        chk.ToolTip = (
            u"Activa o desactiva la malla para todos los tramos."
        )
        ui_stub[u"malla_activo_chk"] = chk
        pnl.Children.Add(chk)
        self._header_malla_activo_chk = chk

        def _on_hdr(sender, args, c=chk):
            if getattr(self, "_suppress_malla_activo_chk", False):
                return
            try:
                on = bool(c.IsChecked)
            except Exception:
                on = False
            self._set_malla_activo_all(on, animate=True)

        try:
            from System.Windows import RoutedEventHandler as _REH
            chk.Checked += _REH(_on_hdr)
            chk.Unchecked += _REH(_on_hdr)
        except Exception:
            pass
        od = getattr(self, "_walls_display_order", []) or []
        on = True
        if od:
            on = self._malla_activo_cfg(_wall_id_int(od[0]))
        self._sync_bulk_malla_activo_toggle(on, animate=False)

    def _build_mesh_apply_button(self, label, click_handler=None):
        from System.Windows.Controls import Button
        from System.Windows import Thickness, HorizontalAlignment

        btn = Button()
        btn.Content = label
        btn.Padding = Thickness(8, 4, 8, 4)
        btn.FontSize = 10.0
        btn.HorizontalAlignment = HorizontalAlignment.Stretch
        try:
            btn.Style = self._win.FindResource(u"BtnSelectOutline")
        except Exception:
            pass
        if click_handler is not None:
            try:
                from System.Windows import RoutedEventHandler as _REH

                btn.Click += _REH(click_handler)
            except Exception:
                pass
        return btn

    def _build_mesh_params_stack(
        self,
        ctr,
        md_key,
        ms_key,
        id_key,
        is_key,
        major_lbl=u"Vertical",
        minor_lbl=u"Horizontal",
        pal=None,
        wire_malla_conf=False,
        wall_id=None,
    ):
        from System.Windows.Controls import StackPanel, ComboBox, TextBlock, Orientation
        from System.Windows import Thickness, FontWeights

        if pal is None:
            pal = self._mesh_ui_palette()

        def _combo_pair(md_k, ms_k):
            cedm = ComboBox()
            cedm.IsEditable = False
            cedm.Margin = Thickness(0, 2, 6, 0)
            cem = ComboBox()
            cem.Margin = Thickness(0, 2, 0, 0)
            cem.IsEditable = True
            self._apply_flat_combo(cedm, narrow=False)
            self._apply_flat_combo(cem, narrow=True)
            ctr[md_k] = cedm
            ctr[ms_k] = cem
            row = StackPanel()
            row.Orientation = Orientation.Horizontal
            row.Children.Add(cedm)
            row.Children.Add(cem)
            self._fill_combo_diam_esp(cedm, cem)
            if wire_malla_conf and wall_id is not None:
                try:
                    from System.Windows import RoutedEventHandler as _REH_mesh

                    def _on_mesh_conf_change(sender, evt, w_id=wall_id):
                        if not self._malla_activo_cfg(w_id):
                            return
                        self._sync_cabezal_confinement_from_malla_wall(w_id)

                    cedm.SelectionChanged += _REH_mesh(_on_mesh_conf_change)
                    cem.SelectionChanged += _REH_mesh(_on_mesh_conf_change)
                    cem.LostFocus += _REH_mesh(_on_mesh_conf_change)
                except Exception:
                    pass
            return row

        side = StackPanel()
        side.Margin = Thickness(0)
        lmaj = TextBlock()
        lmaj.Text = major_lbl
        lmaj.Foreground = pal[u"text_caption"]
        lmaj.FontSize = 9.0
        side.Children.Add(lmaj)
        side.Children.Add(_combo_pair(md_key, ms_key))
        lmin = TextBlock()
        lmin.Text = minor_lbl
        lmin.Margin = Thickness(0, 4, 0, 0)
        lmin.Foreground = pal[u"text_caption"]
        lmin.FontSize = 9.0
        side.Children.Add(lmin)
        side.Children.Add(_combo_pair(id_key, is_key))
        return side

    def _build_mesh_params_row(
        self,
        ctr,
        md_key,
        ms_key,
        id_key,
        is_key,
        major_lbl=u"Vertical",
        minor_lbl=u"Horizontal",
        pal=None,
        wire_malla_conf=False,
        wall_id=None,
    ):
        """Vertical y Horizontal en la misma fila (bulk ancho, columna elevación)."""
        from System.Windows.Controls import (
            Grid,
            ColumnDefinition,
            StackPanel,
            ComboBox,
            TextBlock,
            Orientation,
        )
        from System.Windows import (
            Thickness,
            GridLength,
            GridUnitType,
            HorizontalAlignment,
            VerticalAlignment,
        )

        if pal is None:
            pal = self._mesh_ui_palette()

        grid = Grid()
        grid.HorizontalAlignment = HorizontalAlignment.Stretch
        grid.VerticalAlignment = VerticalAlignment.Top
        cd_maj = ColumnDefinition()
        cd_maj.Width = GridLength(1.0, GridUnitType.Star)
        cd_gap = ColumnDefinition()
        cd_gap.Width = GridLength(10.0, GridUnitType.Pixel)
        cd_min = ColumnDefinition()
        cd_min.Width = GridLength(1.0, GridUnitType.Star)
        for cd in (cd_maj, cd_gap, cd_min):
            grid.ColumnDefinitions.Add(cd)

        def _combo_pair(md_k, ms_k):
            cedm = ComboBox()
            cedm.IsEditable = False
            cedm.Margin = Thickness(0, 2, 6, 0)
            cem = ComboBox()
            cem.Margin = Thickness(0, 2, 0, 0)
            cem.IsEditable = True
            self._apply_flat_combo(cedm, narrow=False)
            self._apply_flat_combo(cem, narrow=True)
            ctr[md_k] = cedm
            ctr[ms_k] = cem
            row = StackPanel()
            row.Orientation = Orientation.Horizontal
            row.Children.Add(cedm)
            row.Children.Add(cem)
            self._fill_combo_diam_esp(cedm, cem)
            if wire_malla_conf and wall_id is not None:
                try:
                    from System.Windows import RoutedEventHandler as _REH_mesh

                    def _on_mesh_conf_change(sender, evt, w_id=wall_id):
                        if not self._malla_activo_cfg(w_id):
                            return
                        self._sync_cabezal_confinement_from_malla_wall(w_id)

                    cedm.SelectionChanged += _REH_mesh(_on_mesh_conf_change)
                    cem.SelectionChanged += _REH_mesh(_on_mesh_conf_change)
                    cem.LostFocus += _REH_mesh(_on_mesh_conf_change)
                except Exception:
                    pass
            return row

        def _side(col, lbl, md_k, ms_k):
            side = StackPanel()
            side.Orientation = Orientation.Vertical
            side.HorizontalAlignment = HorizontalAlignment.Stretch
            side.Margin = Thickness(0)
            lt = TextBlock()
            lt.Text = lbl
            lt.Foreground = pal[u"text_caption"]
            lt.FontSize = 9.0
            lt.Margin = Thickness(0, 0, 0, 0)
            side.Children.Add(lt)
            side.Children.Add(_combo_pair(md_k, ms_k))
            Grid.SetColumn(side, col)
            grid.Children.Add(side)

        _side(0, major_lbl, md_key, ms_key)
        _side(2, minor_lbl, id_key, is_key)
        return grid

    def _wrap_mesh_settings_card(self, body, title=None, compact=False):
        from System.Windows.Controls import Border, StackPanel, TextBlock
        from System.Windows import (
            Thickness,
            FontWeights,
            HorizontalAlignment,
            VerticalAlignment,
            CornerRadius,
        )

        pal = self._mesh_ui_palette(u"unit" if compact else u"bulk")
        shell = Border()
        shell.Background = (
            pal[u"panel_fill_overlay"] if compact else pal[u"panel_fill"]
        )
        shell.BorderBrush = pal[u"sep"]
        shell.BorderThickness = Thickness(1)
        shell.Padding = (
            Thickness(6, 5, 6, 5) if compact else Thickness(8, 7, 8, 7)
        )
        shell.Margin = Thickness(0)
        shell.HorizontalAlignment = HorizontalAlignment.Stretch
        shell.VerticalAlignment = VerticalAlignment.Top
        try:
            shell.CornerRadius = CornerRadius(4.0)
        except Exception:
            pass
        if compact:
            try:
                shell.MaxWidth = float(
                    getattr(self, u"_MESH_OVERLAY_COMPACT_MAX_W_PX", 212.0),
                )
                shell.HorizontalAlignment = HorizontalAlignment.Left
            except Exception:
                pass

        outer = StackPanel()
        accent_bar = Border()
        accent_bar.Height = float(pal[u"accent_bar_height_px"])
        accent_bar.Background = pal[u"accent"]
        accent_bar.Margin = Thickness(0, 0, 0, 4)
        accent_bar.HorizontalAlignment = HorizontalAlignment.Stretch
        outer.Children.Add(accent_bar)

        if title:
            hdr = TextBlock()
            hdr.Text = title
            hdr.Foreground = pal[u"text_section"]
            hdr.FontSize = 9.0 if compact else 10.0
            hdr.FontWeight = FontWeights.SemiBold
            hdr.Margin = Thickness(0, 0, 0, 4 if compact else 6)
            outer.Children.Add(hdr)

        outer.Children.Add(body)
        shell.Child = outer
        return shell

    def _build_bulk_action_card(
        self,
        title,
        panel_fill,
        sep_br,
        accent_br,
        body,
        toggle_chk=None,
        apply_button=None,
        stretch=True,
    ):
        """Tarjeta homogénea para acciones masivas (cabezal / mallas)."""
        from System.Windows.Controls import (
            Border,
            TextBlock,
            Grid,
            ColumnDefinition,
            RowDefinition,
        )
        from System.Windows import (
            Thickness,
            FontWeights,
            HorizontalAlignment,
            VerticalAlignment,
            CornerRadius,
            GridLength,
            GridUnitType,
            TextWrapping,
        )

        accent_h = 3.0

        shell = Border()
        shell.Background = panel_fill
        shell.BorderBrush = sep_br
        shell.BorderThickness = Thickness(1)
        shell.Padding = Thickness(8, 7, 8, 7)
        shell.Margin = Thickness(0)
        shell.VerticalAlignment = VerticalAlignment.Top
        if stretch:
            shell.HorizontalAlignment = HorizontalAlignment.Stretch
        else:
            shell.HorizontalAlignment = HorizontalAlignment.Center
        try:
            shell.CornerRadius = CornerRadius(4.0)
        except Exception:
            pass

        outer = Grid()
        outer.HorizontalAlignment = HorizontalAlignment.Stretch
        outer.VerticalAlignment = VerticalAlignment.Top
        for hkind in (
            GridLength.Auto,
            GridLength.Auto,
            GridLength.Auto,
            GridLength.Auto,
        ):
            rd = RowDefinition()
            rd.Height = hkind
            outer.RowDefinitions.Add(rd)

        accent_bar = Border()
        accent_bar.Height = accent_h
        accent_bar.Background = accent_br
        accent_bar.Margin = Thickness(0, 0, 0, 6)
        accent_bar.HorizontalAlignment = HorizontalAlignment.Stretch
        Grid.SetRow(accent_bar, 0)
        outer.Children.Add(accent_bar)

        hdr = Grid()
        hdr.Margin = Thickness(0, 0, 0, 6)
        hdr.HorizontalAlignment = HorizontalAlignment.Stretch
        cd_title = ColumnDefinition()
        cd_title.Width = GridLength(1.0, GridUnitType.Star)
        cd_toggle = ColumnDefinition()
        cd_toggle.Width = GridLength.Auto
        hdr.ColumnDefinitions.Add(cd_title)
        hdr.ColumnDefinitions.Add(cd_toggle)

        title_tb = TextBlock()
        title_tb.Text = title or u""
        try:
            from System.Windows.Media import SolidColorBrush, Color
            title_tb.Foreground = SolidColorBrush(Color.FromRgb(226, 232, 240))
        except Exception:
            pass
        title_tb.FontSize = 10.0
        title_tb.FontWeight = FontWeights.SemiBold
        title_tb.VerticalAlignment = VerticalAlignment.Center
        title_tb.HorizontalAlignment = HorizontalAlignment.Left
        title_tb.TextWrapping = TextWrapping.Wrap
        Grid.SetColumn(title_tb, 0)
        hdr.Children.Add(title_tb)

        if toggle_chk is not None:
            try:
                toggle_chk.Margin = Thickness(0)
                toggle_chk.Padding = Thickness(0)
                toggle_chk.HorizontalAlignment = HorizontalAlignment.Right
                toggle_chk.VerticalAlignment = VerticalAlignment.Center
            except Exception:
                pass
            Grid.SetColumn(toggle_chk, 1)
            hdr.Children.Add(toggle_chk)

        Grid.SetRow(hdr, 1)
        outer.Children.Add(hdr)

        if body is not None:
            try:
                body.HorizontalAlignment = HorizontalAlignment.Stretch
                body.VerticalAlignment = VerticalAlignment.Top
            except Exception:
                pass
            Grid.SetRow(body, 2)
            outer.Children.Add(body)

        if apply_button is not None:
            try:
                apply_button.Margin = Thickness(0, 6, 0, 0)
                apply_button.HorizontalAlignment = HorizontalAlignment.Stretch
                apply_button.MinWidth = 0.0
                apply_button.VerticalAlignment = VerticalAlignment.Top
            except Exception:
                pass
            Grid.SetRow(apply_button, 3)
            outer.Children.Add(apply_button)

        shell.Child = outer
        return shell

    def _cabezal_apply_stepper_value_style(self, val_tb, palette):
        if val_tb is None or palette is None:
            return
        from System.Windows.Media import SolidColorBrush, Color
        from System.Windows import FontWeights

        r, g, b = palette[u"text_value_rgb"]
        val_tb.Foreground = SolidColorBrush(Color.FromRgb(r, g, b))
        val_tb.FontWeight = FontWeights.SemiBold

    def _toggle_on_color_for_ui(self, ui):
        from System.Windows.Media import Color

        rgb = (ui or {}).get(u"toggle_accent_rgb")
        if rgb and len(rgb) >= 3:
            return Color.FromRgb(int(rgb[0]), int(rgb[1]), int(rgb[2]))
        return self._toggle_mini_color_on()

    def _init_cabezal_ctrl_grid_columns(self, grid):
        from System.Windows.Controls import ColumnDefinition
        from System.Windows import GridLength, GridUnitType

        if grid is None:
            return
        try:
            grid.ColumnDefinitions.Clear()
        except Exception:
            pass
        lbl_w, step_w, diam_w, strip_w, _, gap_w = self._cabezal_ctrl_metrics()
        for w in (lbl_w, step_w, gap_w, diam_w):
            cd = ColumnDefinition()
            cd.Width = GridLength(float(w), GridUnitType.Pixel)
            grid.ColumnDefinitions.Add(cd)
        try:
            grid.Width = strip_w
            grid.MinWidth = strip_w
            grid.MaxWidth = strip_w
        except Exception:
            pass
        return strip_w

    def _add_cabezal_ctrl_header_row(self, grid, row_idx, fg_lo):
        from System.Windows.Controls import TextBlock, Grid, RowDefinition
        from System.Windows import (
            GridLength,
            GridUnitType,
            FontWeights,
            VerticalAlignment,
            HorizontalAlignment,
        )

        _, _, _, _, row_h, _ = self._cabezal_ctrl_metrics()
        while grid.RowDefinitions.Count <= row_idx:
            rd = RowDefinition()
            rd.Height = GridLength(row_h, GridUnitType.Pixel)
            grid.RowDefinitions.Add(rd)

        for col, txt in ((0, u"capa"), (1, u"n"), (3, u"\u00d8")):
            tb = TextBlock()
            tb.Text = txt
            tb.Foreground = fg_lo
            tb.FontSize = 9.0
            tb.FontWeight = FontWeights.SemiBold
            tb.VerticalAlignment = VerticalAlignment.Center
            tb.HorizontalAlignment = HorizontalAlignment.Center
            Grid.SetRow(tb, row_idx)
            Grid.SetColumn(tb, col)
            grid.Children.Add(tb)

    def _add_cabezal_ctrl_row(
        self,
        grid,
        row_idx,
        label_text,
        stepper_panel,
        diam_cb,
        fg_lo,
        fg_accent,
        accent_label=False,
    ):
        from System.Windows.Controls import TextBlock, Grid, RowDefinition
        from System.Windows import (
            GridLength,
            GridUnitType,
            FontWeights,
            VerticalAlignment,
            HorizontalAlignment,
        )

        lbl_w, _, _, _, row_h, _ = self._cabezal_ctrl_metrics()
        while grid.RowDefinitions.Count <= row_idx:
            rd = RowDefinition()
            rd.Height = GridLength(row_h, GridUnitType.Pixel)
            grid.RowDefinitions.Add(rd)

        lbl = TextBlock()
        lbl.Text = label_text
        lbl.Foreground = fg_accent if accent_label else fg_lo
        lbl.FontSize = 11.0 if accent_label else 10.0
        lbl.FontWeight = FontWeights.SemiBold if accent_label else FontWeights.Normal
        lbl.Width = lbl_w
        lbl.VerticalAlignment = VerticalAlignment.Center
        lbl.HorizontalAlignment = HorizontalAlignment.Left
        Grid.SetRow(lbl, row_idx)
        Grid.SetColumn(lbl, 0)
        grid.Children.Add(lbl)

        if stepper_panel is not None:
            stepper_panel.VerticalAlignment = VerticalAlignment.Center
            stepper_panel.HorizontalAlignment = HorizontalAlignment.Center
            Grid.SetRow(stepper_panel, row_idx)
            Grid.SetColumn(stepper_panel, 1)
            grid.Children.Add(stepper_panel)

        if diam_cb is not None:
            diam_cb.VerticalAlignment = VerticalAlignment.Center
            diam_cb.HorizontalAlignment = HorizontalAlignment.Stretch
            Grid.SetRow(diam_cb, row_idx)
            Grid.SetColumn(diam_cb, 3)
            grid.Children.Add(diam_cb)

    def _clear_cabezal_ctrl_grid_rows(self, grid, from_row, to_row=None):
        from System.Windows.Controls import Grid

        if grid is None:
            return
        try:
            fr = int(from_row)
            tr = int(to_row) if to_row is not None else None
            to_remove = []
            for ch in list(grid.Children):
                try:
                    r = Grid.GetRow(ch)
                    if r < fr:
                        continue
                    if tr is not None and r >= tr:
                        continue
                    to_remove.append(ch)
                except Exception:
                    pass
            for ch in to_remove:
                grid.Children.Remove(ch)
            if tr is None:
                while grid.RowDefinitions.Count > fr:
                    grid.RowDefinitions.RemoveAt(grid.RowDefinitions.Count - 1)
        except Exception:
            pass

    def _bulk_armado_toggle_hosts_for(self, extremo):
        hosts = getattr(self, "_bulk_armado_toggle_hosts", None)
        if hosts is None:
            self._bulk_armado_toggle_hosts = {}
            hosts = self._bulk_armado_toggle_hosts
        return hosts.setdefault(extremo, [])

    def _sync_bulk_armado_toggle_widgets(self, extremo, activo, animate=False):
        """Mantiene sincronizados los toggles de cabezal (columnas Inicio/Final)."""
        self._suppress_cabezal_armado_chk = True
        try:
            for entry in self._bulk_armado_toggle_hosts_for(extremo):
                chk = entry.get(u"chk")
                ui_stub = entry.get(u"ui_stub")
                if chk is None:
                    continue
                try:
                    chk.IsChecked = bool(activo)
                except Exception:
                    pass
                if ui_stub:
                    try:
                        self._apply_toggle_mini_visual(
                            ui_stub,
                            activo,
                            animate=animate,
                            parts_key=u"armado_toggle_parts",
                        )
                    except Exception:
                        pass
        finally:
            self._suppress_cabezal_armado_chk = False

    def _set_cabezal_armado_extremo(self, extremo, activo, animate=False):
        """Activa/desactiva armado cabezal en todos los muros de un extremo."""
        if cabezal is None:
            return
        on = bool(activo)
        for w in getattr(self, "_walls_display_order", []) or []:
            try:
                wid = _wall_id_int(w)
            except Exception:
                continue
            cfg = self._cabezal_by_wall_id.setdefault(
                wid, cabezal.default_cabezal_muro_config(),
            )
            ex_cfg = cfg.setdefault(extremo, cabezal.default_cabezal_extremo_config())
            ex_cfg[u"armado_activo"] = on
            self._apply_cabezal_armado_ui_state(wid, extremo)
        self._sync_bulk_armado_toggle_widgets(extremo, on, animate=animate)
        self._set_cabezal_bulk_panel_enabled(extremo, on)
        try:
            self._refresh_cabezal_bulk_preview(extremo)
        except Exception:
            pass
        try:
            self._schedule_full_redraw()
        except Exception:
            pass

    def _add_bulk_armado_toggle(self, extremo, parent=None, margin_bottom=6, header_mode=False):
        """Añade un toggle masivo de armado cabezal (columnas Inicio/Final)."""
        from System.Windows.Controls import CheckBox
        from System.Windows import HorizontalAlignment, VerticalAlignment, Thickness

        if cabezal is None:
            return None
        pal = self._cabezal_ui_palette(extremo, u"bulk")
        accent_rgb = pal[u"bar_accent_rgb"]
        ex_lbl = self._cabezal_extremo_ui_label(extremo)
        ui_stub = {
            u"toggle_accent_rgb": accent_rgb,
            u"toggle_mini_parts_key": u"armado_toggle_parts",
            u"toggle_mini_label": (
                u"" if header_mode else u"Armado {0}".format(ex_lbl)
            ),
            u"toggle_mini_label_fg": pal[u"text_title"],
        }

        chk = CheckBox()
        chk.Margin = Thickness(0, 0, 0, 0 if header_mode else margin_bottom)
        chk.Padding = Thickness(0)
        chk.HorizontalAlignment = (
            HorizontalAlignment.Right if header_mode else HorizontalAlignment.Left
        )
        chk.VerticalAlignment = VerticalAlignment.Center
        self._apply_bimtools_toggle_mini(chk)
        self._build_toggle_mini_content(chk, ui_stub)
        chk.ToolTip = (
            u"Activa o desactiva el armado completo del cabezal en el extremo "
            u"{0} para todos los muros.".format(ex_lbl)
        )
        if parent is not None:
            parent.Children.Add(chk)
        self._bulk_armado_toggle_hosts_for(extremo).append({
            u"chk": chk,
            u"ui_stub": ui_stub,
        })
        bulk = self._cabezal_bulk_ui.setdefault(extremo, {})
        bulk[u"armado_activo_chk"] = chk

        def _on_bulk_armado(sender, args, ex=extremo, c=chk):
            if getattr(self, "_suppress_cabezal_armado_chk", False):
                return
            try:
                on = bool(c.IsChecked)
            except Exception:
                on = False
            self._set_cabezal_armado_extremo(ex, on, animate=True)

        try:
            from System.Windows import RoutedEventHandler as _REH
            chk.Checked += _REH(_on_bulk_armado)
            chk.Unchecked += _REH(_on_bulk_armado)
        except Exception:
            pass
        return chk

    def _ensure_cabezal_bulk_armado_toggle(self, extremo, parent, margin_bottom=0):
        """Toggle masivo en panel cabezal (columna Inicio/Final)."""
        if self._add_bulk_armado_toggle(extremo, parent, margin_bottom=margin_bottom) is None:
            return
        on = True
        od = getattr(self, "_walls_display_order", []) or []
        if od:
            try:
                on = self._cabezal_armado_activo_cfg(_wall_id_int(od[0]), extremo)
            except Exception:
                on = True
        self._sync_bulk_armado_toggle_widgets(extremo, on, animate=False)
        self._set_cabezal_bulk_panel_enabled(extremo, on)

    def _apply_cabezal_armado_extremo_all_walls(self, extremo, activo):
        self._set_cabezal_armado_extremo(extremo, activo, animate=False)

    def _ensure_cabezal_empalmes_checkbox(self, wid, extremo, parent_grid, grid_column=0):
        """Obsoleto: troceo/empalme solo en elevación (selectores I/F en pie del muro)."""
        return

    def _rebuild_cabezal_all_walls_for_extremo(self, extremo):
        """Recalcula paneles cabezal por tramo tras cambio de empalmes."""
        if cabezal is None:
            return
        self._invalidate_cabezal_segments_cache(extremo)
        try:
            self._sync_cabezal_from_segment_owners()
        except Exception:
            pass
        try:
            self._cabezal_seed_segment_owners_from_predecessor(extremo)
        except Exception:
            pass
        try:
            self._build_wall_parameter_panels()
            self._wire_controls()
            self._redraw_preview_canvas()
        except Exception:
            seen = set()
            for seg in self._cabezal_segments_for_extremo(extremo):
                try:
                    owner_idx = int(seg.get(u"owner_index", 0))
                except Exception:
                    owner_idx = 0
                walls = getattr(self, u"walls_ordered", []) or []
                if not (0 <= owner_idx < len(walls)):
                    continue
                try:
                    owner_wid = _wall_id_int(walls[owner_idx])
                except Exception:
                    continue
                if owner_wid in seen:
                    continue
                seen.add(owner_wid)
                self._rebuild_cabezal_layer_sliders(
                    owner_wid, extremo, redistribute=False,
                )
            try:
                self._redistribute_row_heights_and_redraw()
            except Exception:
                pass
        try:
            self._refresh_cabezal_bar_length_warns_for_extremo(extremo)
        except Exception:
            pass

    def _rebuild_cabezal_layer_sliders(self, wid, extremo, redistribute=True, refresh_bar_length=True):
        if cabezal is None:
            return
        ui = self._cabezal_ui_ext(wid, extremo)
        ctrl_grid = ui.get(u"controls_grid")
        if ctrl_grid is None:
            return
        strip_w = self._init_cabezal_ctrl_grid_columns(ctrl_grid)
        ui[u"layers_row_start"] = 1
        if self._cabezal_layer_steppers_are_live(ui):
            self._sync_cabezal_extremo_from_ui(wid, extremo)
        cfg = self._cabezal_by_wall_id.get(wid) or {}
        ex_cfg = cfg.get(extremo) or cabezal.default_cabezal_extremo_config()
        segs = self._cabezal_segments_for_extremo(extremo)
        cabezal._migrate_tramo_to_segment_bar_type_ids(
            ex_cfg, segs, ex_cfg.get(u"bar_type_id"),
        )
        cabezal._normalize_cabezal_extremo_layers(ex_cfg)
        layers = list(ex_cfg.get(u"layers") or [])

        n_capas = self._cabezal_n_capas_from_ui(wid, extremo)
        min_capas = self._cabezal_min_capas_for_wall_extremo(wid, extremo)
        n_capas = max(
            min_capas,
            min(cabezal.CABEZAL_MAX_CAPAS, int(n_capas)),
        )
        ex_cfg[u"n_capas"] = n_capas

        layer_start = int(ui.get(u"layers_row_start", 1))
        self._clear_cabezal_ctrl_grid_rows(ctrl_grid, 0)

        from System.Windows.Controls import ScrollBarVisibility

        pal = self._cabezal_ui_palette(extremo, u"unit")
        self._add_cabezal_ctrl_header_row(ctrl_grid, 0, pal[u"text_caption"])
        layer_steppers = []
        layer_diam_cbs = []
        layer_cap_lbls = []
        min_bars = self._cabezal_min_bars_for_wall_extremo(wid, extremo)

        def _make_layer_change_handler(layer_idx):
            def _on_layer_change(_v):
                if getattr(self, "_suppress_cabezal_stepper", False):
                    return
                if layer_idx >= self._cabezal_n_capas_from_ui(wid, extremo):
                    return
                self._sync_cabezal_extremo_from_ui(wid, extremo)
                self._request_cabezal_preview_refresh(wid, extremo)
            return _on_layer_change

        def _make_diam_change_handler(layer_i):
            def _on_diam_change():
                self._sync_cabezal_extremo_from_ui(wid, extremo)
                if int(layer_i) == 0:
                    self._sync_cabezal_confinement_from_malla_wall(
                        wid, refresh_preview=False,
                    )
                self._request_cabezal_preview_refresh(wid, extremo)
            return _on_diam_change

        diam_tip = u"Diámetro (ø) por capa"
        inactive_tip = u"Capa fuera del conteo CAPAS — solo lectura"

        for i in range(cabezal.CABEZAL_MAX_CAPAS):
            ly = layers[i] if i < len(layers) else cabezal.default_cabezal_layer_config(2, ex_cfg.get(u"bar_type_id"))
            layer_active = i < n_capas
            try:
                nb = int(ly.get(u"n_bars", 2))
            except Exception:
                nb = 2
            step = self._create_cabezal_stepper(
                min_bars,
                cabezal.CABEZAL_MAX_BARRAS_POR_CAPA,
                nb,
                on_change=_make_layer_change_handler(i),
                palette=pal,
            )
            try:
                step[u"value_tb"].ToolTip = (
                    u"L{0} — cantidad (n)".format(i + 1)
                    if layer_active else inactive_tip + u" — L{0}".format(i + 1)
                )
            except Exception:
                pass
            bid = self._cabezal_layer_diam_id(wid, extremo, i)
            diam_cb = self._create_cabezal_diam_combo(
                bid,
                on_change=_make_diam_change_handler(i),
            )
            try:
                tip = diam_tip + u" — L{0}".format(i + 1)
                if not layer_active:
                    tip = inactive_tip
                diam_cb.ToolTip = tip
            except Exception:
                pass
            row_fg = (
                pal[u"layer_active"] if layer_active else pal[u"layer_inactive"]
            )
            self._apply_cabezal_diam_combo_state(
                diam_cb, layer_active=layer_active, palette=pal,
            )
            self._apply_cabezal_stepper_enabled(step, layer_active)
            self._add_cabezal_ctrl_row(
                ctrl_grid,
                layer_start + i,
                u"{0}\u00aaC.".format(i + 1),
                step[u"panel"],
                diam_cb,
                pal[u"layer_inactive"],
                row_fg,
                accent_label=layer_active,
            )
            row_cap_lbl = None
            try:
                from System.Windows.Controls import Grid as _Gr
                for ch in list(ctrl_grid.Children):
                    if _Gr.GetRow(ch) == layer_start + i and _Gr.GetColumn(ch) == 0:
                        row_cap_lbl = ch
                        break
            except Exception:
                pass
            layer_cap_lbls.append(row_cap_lbl)
            layer_steppers.append(step)
            layer_diam_cbs.append(diam_cb)

        ui[u"layer_steppers"] = layer_steppers
        ui[u"layer_diam_cbs"] = layer_diam_cbs
        ui[u"layer_cap_lbls"] = layer_cap_lbls
        ui[u"layer_value_tbs"] = [
            st[u"value_tb"] if st else None for st in layer_steppers
        ]
        try:
            cap_step = ui.get(u"capas_stepper")
            if cap_step is not None:
                self._suppress_cabezal_stepper = True
                try:
                    cap_step[u"set_value"](n_capas)
                finally:
                    self._suppress_cabezal_stepper = False
            else:
                tb_cap = ui.get(u"capas_value_tb")
                if tb_cap is not None:
                    tb_cap.Text = str(n_capas)
        except Exception:
            pass
        try:
            cap_grid = ui.get(u"capas_grid") or ui.get(u"toolbar_grid")
            cap_step = ui.get(u"capas_stepper")
            if cap_grid is not None and cap_step is not None:
                self._fill_cabezal_toolbar_row(
                    cap_grid,
                    cap_step.get(u"panel"),
                    pal,
                    ui.get(u"tramo_toolbar_text"),
                )
        except Exception:
            pass
        try:
            scroll = ui.get(u"controls_scroll")
            if scroll is not None:
                scroll_max = self._cabezal_layers_scroll_max_height_px(
                    cabezal.CABEZAL_MAX_CAPAS,
                )
                scroll.MaxHeight = scroll_max
                if cabezal.CABEZAL_MAX_CAPAS > int(
                    getattr(self, "_CABEZAL_LAYER_SCROLL_MAX_ROWS", 3),
                ):
                    scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
                else:
                    scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled
        except Exception:
            pass
        try:
            cv = ui.get(u"preview_canvas")
            if cv is not None:
                prev_w, prev_h = self._cabezal_preview_canvas_size_px(wid, extremo)
                cv.Width = prev_w
                cv.Height = prev_h
                cv.MinWidth = prev_w
                cv.MaxWidth = prev_w
                cv.MinHeight = prev_h
                cv.MaxHeight = prev_h
        except Exception:
            pass
        try:
            wrap = ui.get(u"cab_wrap")
            if wrap is not None and strip_w is not None:
                cap_w = max(float(self._cabezal_cap_col_px()), float(strip_w) + 16.0)
                wrap.Width = cap_w
                wrap.MinWidth = cap_w
                wrap.MaxWidth = cap_w
        except Exception:
            pass
        self._request_cabezal_preview_refresh(wid, extremo, debounce=False)
        if refresh_bar_length:
            self._refresh_cabezal_bar_length_warn(wid, extremo)
        self._refresh_cabezal_encuentro_ui_state(wid, extremo)
        if redistribute:
            self._schedule_full_redraw()

    def _draw_cabezal_break_line(self, canv, x_edge, y0, h, stroke_br, z_index=4, face_right=True):
        """
        Símbolo de rotura vertical (corte arquitectónico) en el borde opuesto a la punta.

        Trayectoria: vertical → jog horizontal → diagonal → vuelta al eje → vertical.
        """
        from System.Windows.Shapes import Polyline
        from System.Windows import Point as WpfPoint
        from System.Windows.Media import PointCollection

        x = float(x_edge)
        ym = y0 + h * 0.5
        amp = max(2.5, min(5.0, h * 0.13))
        half_z = max(4.0, min(h * 0.22, h * 0.5 - 2.0))

        if face_right:
            x_jog = x - amp
            x_diag = x + amp
        else:
            x_jog = x + amp
            x_diag = x - amp

        y_top_br = ym - half_z
        y_bot_br = ym + half_z
        y_mid_a = ym - half_z * 0.28
        y_mid_b = ym + half_z * 0.28

        pc = PointCollection()
        pc.Add(WpfPoint(x, y0))
        pc.Add(WpfPoint(x, y_top_br))
        pc.Add(WpfPoint(x_jog, y_mid_a))
        pc.Add(WpfPoint(x_diag, y_mid_b))
        pc.Add(WpfPoint(x, y_bot_br))
        pc.Add(WpfPoint(x, y0 + h))

        pl = Polyline()
        pl.Points = pc
        pl.Stroke = stroke_br
        pl.StrokeThickness = 1.35
        try:
            pl.StrokeLineJoin = pl.StrokeLineJoin.Miter
            pl.StrokeStartLineCap = pl.StrokeStartLineCap.Round
            pl.StrokeEndLineCap = pl.StrokeEndLineCap.Round
        except Exception:
            pass
        pl.Fill = None
        self._canvas_set_zindex(pl, z_index)
        canv.Children.Add(pl)

    def _init_cabezal_toolbar_compact_grid(self, grid, strip_w):
        """Toolbar extremo: CAPAS + stepper (der.), sin empalme (solo en elevación)."""
        from System.Windows.Controls import ColumnDefinition, RowDefinition
        from System.Windows import GridLength, GridUnitType, HorizontalAlignment

        if grid is None:
            return
        try:
            grid.ColumnDefinitions.Clear()
            grid.RowDefinitions.Clear()
        except Exception:
            pass
        grid.HorizontalAlignment = HorizontalAlignment.Stretch
        row_px = float(getattr(self, "_CABEZAL_EXTREMO_TOOLBAR_ROW_PX", 24.0))
        try:
            grid.Width = strip_w
            grid.MinWidth = strip_w
            grid.MaxWidth = strip_w
        except Exception:
            pass
        rd = RowDefinition()
        rd.Height = GridLength(row_px, GridUnitType.Pixel)
        grid.RowDefinitions.Add(rd)
        cd_spacer = ColumnDefinition()
        cd_spacer.Width = GridLength(1.0, GridUnitType.Star)
        cd_cap_lbl = ColumnDefinition()
        cd_cap_lbl.Width = GridLength.Auto
        cd_cap_step = ColumnDefinition()
        cd_cap_step.Width = GridLength.Auto
        grid.ColumnDefinitions.Add(cd_spacer)
        grid.ColumnDefinitions.Add(cd_cap_lbl)
        grid.ColumnDefinitions.Add(cd_cap_step)

    def _init_cabezal_bulk_toolbar_grid(self, grid, content_w):
        """Toolbar masivo: Configurador Global (izq.) | CAPAS + stepper + Aplicar (der.)."""
        from System.Windows.Controls import ColumnDefinition, RowDefinition
        from System.Windows import GridLength, GridUnitType, HorizontalAlignment

        if grid is None:
            return
        try:
            grid.ColumnDefinitions.Clear()
            grid.RowDefinitions.Clear()
        except Exception:
            pass
        grid.HorizontalAlignment = HorizontalAlignment.Stretch
        row_px = float(getattr(self, "_CABEZAL_EXTREMO_TOOLBAR_ROW_PX", 24.0))
        try:
            cw = float(content_w)
            grid.Width = cw
            grid.MinWidth = cw
            grid.MaxWidth = cw
        except Exception:
            pass
        rd = RowDefinition()
        rd.Height = GridLength(row_px, GridUnitType.Pixel)
        grid.RowDefinitions.Add(rd)
        for wkind in (
            GridLength.Auto,
            GridLength(1.0, GridUnitType.Star),
            GridLength.Auto,
            GridLength.Auto,
            GridLength.Auto,
        ):
            cd = ColumnDefinition()
            cd.Width = wkind
            grid.ColumnDefinitions.Add(cd)

    def _fill_cabezal_bulk_toolbar_row(self, grid, cap_step_panel, apply_btn, palette):
        from System.Windows.Controls import TextBlock, Grid, Button
        from System.Windows import (
            FontWeights,
            VerticalAlignment,
            HorizontalAlignment,
            Thickness,
        )

        if grid is None:
            return
        cap_row = 0
        try:
            to_remove = []
            for child in list(grid.Children):
                to_remove.append(child)
            for child in to_remove:
                grid.Children.Remove(child)
        except Exception:
            pass

        mas = TextBlock()
        mas.Text = u"Configurador Global"
        mas.Foreground = palette[u"text_title"]
        mas.FontSize = 10.0
        mas.FontWeight = FontWeights.SemiBold
        mas.Margin = Thickness(0, 0, 0, 0)
        mas.VerticalAlignment = VerticalAlignment.Center
        mas.HorizontalAlignment = HorizontalAlignment.Left
        Grid.SetRow(mas, cap_row)
        Grid.SetColumn(mas, 0)
        grid.Children.Add(mas)

        lbl = TextBlock()
        lbl.Text = u"CAPAS"
        lbl.Foreground = palette[u"text_control"]
        lbl.FontSize = 10.0
        lbl.FontWeight = FontWeights.SemiBold
        lbl.Margin = Thickness(0, 0, 4, 0)
        lbl.VerticalAlignment = VerticalAlignment.Center
        lbl.HorizontalAlignment = HorizontalAlignment.Right
        Grid.SetRow(lbl, cap_row)
        Grid.SetColumn(lbl, 2)
        grid.Children.Add(lbl)

        if cap_step_panel is not None:
            try:
                parent = cap_step_panel.Parent
                if parent is not None:
                    parent.Children.Remove(cap_step_panel)
            except Exception:
                pass
            cap_step_panel.VerticalAlignment = VerticalAlignment.Center
            cap_step_panel.HorizontalAlignment = HorizontalAlignment.Right
            cap_step_panel.Margin = Thickness(0, 0, 4, 0)
            Grid.SetRow(cap_step_panel, cap_row)
            Grid.SetColumn(cap_step_panel, 3)
            grid.Children.Add(cap_step_panel)

        if apply_btn is not None:
            try:
                parent = apply_btn.Parent
                if parent is not None:
                    parent.Children.Remove(apply_btn)
            except Exception:
                pass
            apply_btn.VerticalAlignment = VerticalAlignment.Center
            apply_btn.HorizontalAlignment = HorizontalAlignment.Right
            apply_btn.Margin = Thickness(0, 0, 0, 0)
            Grid.SetRow(apply_btn, cap_row)
            Grid.SetColumn(apply_btn, 4)
            grid.Children.Add(apply_btn)

    def _init_cabezal_toolbar_grid(self, grid, strip_w):
        """Columnas toolbar: Empalmes (star) | CAPAS | stepper."""
        from System.Windows.Controls import ColumnDefinition, RowDefinition
        from System.Windows import GridLength, GridUnitType, HorizontalAlignment

        if grid is None:
            return
        try:
            grid.ColumnDefinitions.Clear()
            grid.RowDefinitions.Clear()
        except Exception:
            pass
        grid.HorizontalAlignment = HorizontalAlignment.Center
        try:
            grid.Width = strip_w
            grid.MinWidth = strip_w
            grid.MaxWidth = strip_w
        except Exception:
            pass
        rd_emp = RowDefinition()
        rd_emp.Height = GridLength(22.0, GridUnitType.Pixel)
        rd_cap = RowDefinition()
        rd_cap.Height = GridLength(22.0, GridUnitType.Pixel)
        grid.RowDefinitions.Add(rd_emp)
        grid.RowDefinitions.Add(rd_cap)
        cd_left = ColumnDefinition()
        cd_left.Width = GridLength(1.0, GridUnitType.Star)
        cd_cap_lbl = ColumnDefinition()
        cd_cap_lbl.Width = GridLength.Auto
        cd_cap_step = ColumnDefinition()
        cd_cap_step.Width = GridLength.Auto
        grid.ColumnDefinitions.Add(cd_left)
        grid.ColumnDefinitions.Add(cd_cap_lbl)
        grid.ColumnDefinitions.Add(cd_cap_step)

    def _init_cabezal_empalmes_row_grid(self, grid, strip_w):
        """Fila única Empalmes (toolbar extremo apilado)."""
        from System.Windows.Controls import ColumnDefinition, RowDefinition
        from System.Windows import GridLength, GridUnitType, HorizontalAlignment

        if grid is None:
            return
        try:
            grid.ColumnDefinitions.Clear()
            grid.RowDefinitions.Clear()
        except Exception:
            pass
        grid.HorizontalAlignment = HorizontalAlignment.Stretch
        try:
            grid.Width = strip_w
            grid.MinWidth = strip_w
            grid.MaxWidth = strip_w
        except Exception:
            pass
        rd = RowDefinition()
        rd.Height = GridLength(28.0, GridUnitType.Pixel)
        grid.RowDefinitions.Add(rd)
        cd = ColumnDefinition()
        cd.Width = GridLength(1.0, GridUnitType.Star)
        grid.ColumnDefinitions.Add(cd)

    def _init_cabezal_capas_row_grid(self, grid, strip_w):
        """Fila única CAPAS + stepper (toolbar extremo apilado)."""
        from System.Windows.Controls import ColumnDefinition, RowDefinition
        from System.Windows import GridLength, GridUnitType, HorizontalAlignment

        if grid is None:
            return
        try:
            grid.ColumnDefinitions.Clear()
            grid.RowDefinitions.Clear()
        except Exception:
            pass
        grid.HorizontalAlignment = HorizontalAlignment.Stretch
        try:
            grid.Width = strip_w
            grid.MinWidth = strip_w
            grid.MaxWidth = strip_w
        except Exception:
            pass
        rd = RowDefinition()
        rd.Height = GridLength(28.0, GridUnitType.Pixel)
        grid.RowDefinitions.Add(rd)
        cd_left = ColumnDefinition()
        cd_left.Width = GridLength(1.0, GridUnitType.Star)
        cd_cap_lbl = ColumnDefinition()
        cd_cap_lbl.Width = GridLength.Auto
        cd_cap_step = ColumnDefinition()
        cd_cap_step.Width = GridLength.Auto
        grid.ColumnDefinitions.Add(cd_left)
        grid.ColumnDefinitions.Add(cd_cap_lbl)
        grid.ColumnDefinitions.Add(cd_cap_step)

    def _cabezal_toolbar_cap_row_index(self, grid):
        try:
            n_rows = int(getattr(grid, "RowDefinitions", None).Count)
        except Exception:
            n_rows = 1
        if n_rows > 1:
            return 1
        return 0

    def _fill_cabezal_toolbar_row(self, grid, cap_step_panel, palette, tramo_text=None):
        """Tramo (izq.) + CAPAS + stepper (der.) en una fila compacta."""
        from System.Windows.Controls import TextBlock, Grid
        from System.Windows import (
            FontWeights,
            VerticalAlignment,
            HorizontalAlignment,
            Thickness,
        )

        if grid is None:
            return
        _, _, _, strip_w, _, _ = self._cabezal_ctrl_metrics()
        if len(getattr(grid, "ColumnDefinitions", None) or []) == 0:
            self._init_cabezal_toolbar_grid(grid, strip_w)
        cap_row = self._cabezal_toolbar_cap_row_index(grid)
        try:
            to_remove = []
            for child in list(grid.Children):
                try:
                    col = int(Grid.GetColumn(child))
                    if col >= 1 or (col == 0 and tramo_text):
                        to_remove.append(child)
                except Exception:
                    to_remove.append(child)
            for child in to_remove:
                grid.Children.Remove(child)
        except Exception:
            pass

        if tramo_text:
            tramo_tb = TextBlock()
            tramo_tb.Text = tramo_text
            tramo_tb.Foreground = palette[u"text_caption"]
            tramo_tb.FontSize = 9.0
            tramo_tb.FontWeight = FontWeights.SemiBold
            tramo_tb.VerticalAlignment = VerticalAlignment.Center
            tramo_tb.HorizontalAlignment = HorizontalAlignment.Left
            tramo_tb.Margin = Thickness(0, 0, 4, 0)
            Grid.SetRow(tramo_tb, cap_row)
            Grid.SetColumn(tramo_tb, 0)
            grid.Children.Add(tramo_tb)

        lbl = TextBlock()
        lbl.Text = u"CAPAS"
        lbl.Foreground = palette[u"text_control"]
        lbl.FontSize = 10.0
        lbl.FontWeight = FontWeights.SemiBold
        lbl.Margin = Thickness(8, 0, 4, 0)
        lbl.VerticalAlignment = VerticalAlignment.Center
        lbl.HorizontalAlignment = HorizontalAlignment.Right
        Grid.SetRow(lbl, cap_row)
        Grid.SetColumn(lbl, 1)
        grid.Children.Add(lbl)

        if cap_step_panel is not None:
            try:
                parent = cap_step_panel.Parent
                if parent is not None:
                    parent.Children.Remove(cap_step_panel)
            except Exception:
                pass
            cap_step_panel.VerticalAlignment = VerticalAlignment.Center
            cap_step_panel.HorizontalAlignment = HorizontalAlignment.Right
            cap_step_panel.Margin = Thickness(0, 0, 0, 0)
            Grid.SetRow(cap_step_panel, cap_row)
            Grid.SetColumn(cap_step_panel, 2)
            grid.Children.Add(cap_step_panel)

    def _fill_cabezal_capas_grid(self, grid, cap_step_panel, palette):
        """Compat.: delega en toolbar unificada."""
        self._fill_cabezal_toolbar_row(grid, cap_step_panel, palette)

    def _wall_for_integer_id(self, wid):
        try:
            wid_i = int(wid)
        except Exception:
            return None
        for w in getattr(self, u"walls_ordered", []) or []:
            try:
                if _wall_id_int(w) == wid_i:
                    return w
            except Exception:
                continue
        return None

    def _cabezal_preview_encuentro_ctx(self, wid, wall, extremo):
        """
        Contexto de preview planta encuentro (L/T comparten geometría L).
        None = cabezal libre (rect + punta).
        """
        if cabezal is None or wall is None:
            return None
        cfg = self._cabezal_by_wall_id.get(wid) or {}
        ex_cfg = cfg.get(extremo) or {}
        try:
            e_sel = float(geo.obtener_espesor_muro_mm_approx(wall) or 200.0)
        except Exception:
            e_sel = 200.0
        if cabezal.cabezal_extremo_es_encuentro_l(ex_cfg):
            try:
                e_det = float(ex_cfg.get(u"espesor_detectado_mm") or e_sel)
            except Exception:
                e_det = e_sel
            return {
                u"tipo": u"L",
                u"e_det": e_det,
                u"e_sel": e_sel,
                u"ex_cfg": ex_cfg,
                u"layout_encuentro": True,
            }
        enc = self._elevation_encuentro_vecino_at_extremo(wall, extremo)
        if enc is None or _cab_enc_l is None:
            return None
        vec = enc.get(u"vecino")
        if vec is None:
            return None
        try:
            e_det = float(_cab_enc_l.espesor_mm_wall(vec))
        except Exception:
            e_det = e_sel
        return {
            u"tipo": enc.get(u"tipo") or u"T",
            u"e_det": e_det,
            u"e_sel": e_sel,
            u"ex_cfg": ex_cfg,
            u"layout_encuentro": True,
        }

    def _cabezal_extremo_es_encuentro_preview(self, wid, extremo, wall=None):
        if wall is None:
            wall = self._wall_for_integer_id(wid)
        if wall is None:
            return self._cabezal_extremo_es_encuentro_l(wid, extremo)
        return self._cabezal_preview_encuentro_ctx(wid, wall, extremo) is not None

    def _draw_cabezal_encuentro_planta_l(
        self,
        canv,
        x0,
        y0,
        draw_w,
        draw_h,
        enc_ctx,
        mirror,
        br_fill,
        br_stroke,
        enc_geom=None,
    ):
        """Contorno planta L + zona intersección + unión discontinua (sin texto ni P_join)."""
        from System.Windows.Controls import Canvas as _Cn
        from System.Windows.Shapes import Path, Line, Rectangle
        from System.Windows import Point as WpfPoint
        from System.Windows.Media import PathGeometry, PathFigure, PolyLineSegment, SolidColorBrush, Color

        if _cab_enc_l is None or not enc_ctx:
            return None
        if enc_geom is None:
            enc_geom = _cab_enc_l.cabezal_encuentro_plan_l_polygon_local_px(
                draw_w,
                draw_h,
                enc_ctx.get(u"e_det"),
                enc_ctx.get(u"e_sel"),
                mirror=bool(mirror),
            )
        geom = enc_geom
        points = geom.get(u"points") or []
        if len(points) < 3:
            return None

        def _abs_pt(lx, ly):
            return WpfPoint(float(x0) + float(lx), float(y0) + float(ly))

        fig = PathFigure()
        fig.IsClosed = True
        fig.StartPoint = _abs_pt(points[0][0], points[0][1])
        seg = PolyLineSegment()
        for px, py in points[1:]:
            seg.Points.Add(_abs_pt(px, py))
        fig.Segments.Add(seg)
        pg = PathGeometry()
        pg.Figures.Add(fig)
        path = Path()
        path.Data = pg
        path.Fill = br_fill
        path.Stroke = br_stroke
        path.StrokeThickness = 1.4
        self._canvas_set_zindex(path, 1)
        canv.Children.Add(path)

        zr = geom.get(u"zone_rect") or {}
        if zr:
            wing = Rectangle()
            wing.Width = float(zr.get(u"w", 1.0))
            wing.Height = float(zr.get(u"h", 1.0))
            wing.Fill = SolidColorBrush(Color.FromArgb(26, 56, 189, 248))
            wing.Stroke = SolidColorBrush(Color.FromRgb(56, 189, 248))
            wing.StrokeThickness = 0.9
            try:
                wing.StrokeDashArray = self._wpf_double_collection(4.0, 3.0)
            except Exception:
                pass
            _Cn.SetLeft(wing, float(x0) + float(zr.get(u"x", 0.0)))
            _Cn.SetTop(wing, float(y0) + float(zr.get(u"y", 0.0)))
            self._canvas_set_zindex(wing, 2)
            canv.Children.Add(wing)

        jl = geom.get(u"join_line")
        if jl and len(jl) == 4:
            ln = Line()
            ln.X1 = float(x0) + float(jl[0])
            ln.Y1 = float(y0) + float(jl[1])
            ln.X2 = float(x0) + float(jl[2])
            ln.Y2 = float(y0) + float(jl[3])
            ln.Stroke = SolidColorBrush(Color.FromRgb(203, 213, 225))
            ln.StrokeThickness = 1.2
            try:
                ln.StrokeDashArray = self._wpf_double_collection(5.0, 3.0)
            except Exception:
                pass
            self._canvas_set_zindex(ln, 3)
            canv.Children.Add(ln)

        return geom

    @staticmethod
    def _wpf_double_collection(*values):
        from System.Windows.Media import DoubleCollection
        dc = DoubleCollection()
        for v in values:
            dc.Add(float(v))
        return dc

    def _draw_cabezal_preview_canvas(
        self,
        canv,
        wall,
        wid,
        extremo,
        layers_override=None,
        conf_type_override=None,
        canvas_w_px=None,
        mirror_preview=None,
        bulk_preview=False,
    ):
        from System.Windows.Controls import Canvas as _Cn
        from System.Windows.Shapes import Rectangle, Ellipse
        from System.Windows.Media import SolidColorBrush, Color

        if cabezal is None or canv is None or wall is None:
            return
        canv.Children.Clear()
        armado_on = (
            self._cabezal_bulk_armado_activo(extremo)
            if bulk_preview
            else self._cabezal_armado_activo_cfg(wid, extremo)
        )
        if not armado_on:
            from System.Windows.Controls import TextBlock
            from System.Windows import TextWrapping
            from System.Windows.Media import SolidColorBrush, Color

            tb = TextBlock()
            tb.Text = u"Armado desactivado"
            tb.Foreground = SolidColorBrush(Color.FromRgb(100, 116, 139))
            tb.FontSize = 9.0
            tb.TextWrapping = TextWrapping.Wrap
            _Cn.SetLeft(tb, 8.0)
            _Cn.SetTop(tb, 8.0)
            canv.Children.Add(tb)
            return
        if layers_override is not None:
            preview_layers = list(layers_override)
        else:
            preview_layers = list(self._cabezal_layers_live_from_ui(wid, extremo))
        if canvas_w_px is not None:
            cw = float(canvas_w_px)
            ch = self._cabezal_preview_height_px(wid, extremo)
        else:
            cw, ch = self._cabezal_preview_canvas_size_px(wid, extremo)
        x0, y0, draw_w, draw_h = self._cabezal_preview_draw_rect_px(cw, ch)
        draw_w_est = max(40.0, draw_w)
        draw_h_est = max(28.0, draw_h)

        try:
            e_mm = float(geo.obtener_espesor_muro_mm_approx(wall) or 200.0)
        except Exception:
            e_mm = 200.0

        cov_px = 10.0
        if bulk_preview or layers_override is not None:
            n_capas_preview = max(
                len(preview_layers),
                self._read_cabezal_bulk_capas(extremo),
            )
        else:
            n_capas_preview = max(
                len(preview_layers),
                self._cabezal_n_capas_from_ui(wid, extremo),
            )
        if conf_type_override is not None:
            conf_type = conf_type_override
            conf_norm = cabezal.normalize_cabezal_confinement(
                {u"type": conf_type_override}, n_capas_preview,
            )
        else:
            conf_type = self._cabezal_effective_confinement_type_for_preview(
                wid, extremo,
            )
            cfg_ex = (self._cabezal_by_wall_id.get(wid) or {}).get(extremo) or {}
            conf_norm = cabezal.normalize_cabezal_confinement(
                cfg_ex.get(u"confinement"), n_capas_preview,
            )
        stirrup_diam_preview = conf_norm.get(u"stirrup_diam_mm")
        cfg_ex = (self._cabezal_by_wall_id.get(wid) or {}).get(extremo) or {}
        enc_ctx = self._cabezal_preview_encuentro_ctx(wid, wall, extremo)
        if mirror_preview is not None:
            mirror = bool(mirror_preview)
        else:
            ui = self._cabezal_ui_ext(wid, extremo)
            mirror = bool(ui.get(u"mirror_preview", False))

        enc_geom = None
        zone_layout_w = draw_w
        zone_layout_h = draw_h
        if enc_ctx and _cab_enc_l is not None:
            enc_geom = _cab_enc_l.cabezal_encuentro_plan_l_polygon_local_px(
                draw_w,
                draw_h,
                enc_ctx.get(u"e_det"),
                enc_ctx.get(u"e_sel"),
                mirror=bool(mirror),
            )
            zr_pre = (enc_geom or {}).get(u"zone_rect") or {}
            if zr_pre:
                zone_layout_w = max(1.0, float(zr_pre.get(u"w", draw_w)))
                zone_layout_h = max(1.0, float(zr_pre.get(u"h", draw_h)))

        use_enc_layout = bool(
            enc_ctx and enc_ctx.get(u"layout_encuentro") and _cab_enc_l is not None
        )
        if use_enc_layout:
            try:
                e_det = float(enc_ctx.get(u"e_det") or e_mm)
            except Exception:
                e_det = e_mm
            try:
                e_sel = float(enc_ctx.get(u"e_sel") or e_mm)
            except Exception:
                e_sel = e_mm
            layout = _cab_enc_l.cabezal_seccion_preview_layout_encuentro_l(
                e_det,
                e_sel,
                preview_layers,
                cover_mm=cabezal.CABEZAL_COVER_MM,
                draw_w_px=zone_layout_w,
                draw_h_px=zone_layout_h,
                confinement_type=conf_type,
                confinement_stirrup_diam_mm=stirrup_diam_preview,
            )
            try:
                pitch_mm = float(layout.get(u"pitch_equitativo_mm") or 0.0)
                cfg_ex[u"pitch_equitativo_mm"] = pitch_mm
            except Exception:
                pass
        else:
            layer_pitch_px = float(
                getattr(self, "_CABEZAL_PREVIEW_LAYER_PITCH_PX", 20.0),
            )
            if canvas_w_px is not None:
                n_ly = max(1, len(preview_layers))
                side_pad = 14.0
                if n_ly > 1:
                    layer_pitch_px = min(
                        layer_pitch_px,
                        max(8.0, (draw_w - side_pad) / float(n_ly - 1)),
                    )
            layout = cabezal.cabezal_seccion_preview_layout(
                e_mm,
                preview_layers,
                cover_mm=cabezal.CABEZAL_COVER_MM,
                row_inset_frac_x=cov_px / draw_w_est,
                row_inset_frac_y=cov_px / draw_h_est,
                draw_w_px=draw_w,
                layer_pitch_px=layer_pitch_px,
                draw_h_px=draw_h,
                bar_span_px=float(getattr(self, "_CABEZAL_PREVIEW_BAR_SPAN_PX", 25.0)),
                confinement_type=conf_type,
                confinement_stirrup_diam_mm=stirrup_diam_preview,
            )
        stk = self._wall_thickness_color_hex(wall)

        def _clamp_byte(iv):
            return max(0, min(255, int(iv)))

        def _brush_hex(hx, aa=238):
            return self._ui_brush_hex(hx, alpha=aa)

        br_fill = _brush_hex(stk, aa=200)
        br_stroke = _brush_hex(stk, aa=220)
        br_bar_cn = SolidColorBrush(Color.FromRgb(34, 211, 238))
        br_edge = SolidColorBrush(Color.FromRgb(20, 20, 20))

        x1 = x0 + draw_w

        def _norm_fx(fx):
            f = float(fx)
            return (1.0 - f) if mirror else f

        zr = (enc_geom or {}).get(u"zone_rect") if enc_ctx else None
        if zr:
            zone_x = float(x0) + float(zr.get(u"x", 0.0))
            zone_y = float(y0) + float(zr.get(u"y", 0.0))
            zone_w = max(1.0, float(zr.get(u"w", draw_w)))
            zone_h = max(1.0, float(zr.get(u"h", draw_h)))

            def _px(fx, fy):
                return (
                    zone_x + _norm_fx(fx) * zone_w,
                    zone_y + float(fy) * zone_h,
                )
        else:

            def _px(fx, fy):
                return x0 + _norm_fx(fx) * draw_w, y0 + float(fy) * draw_h

        if enc_ctx:
            self._draw_cabezal_encuentro_planta_l(
                canv,
                x0,
                y0,
                draw_w,
                draw_h,
                enc_ctx,
                mirror,
                br_fill,
                br_stroke,
                enc_geom=enc_geom,
            )
        else:
            outer = Rectangle()
            outer.Width = draw_w
            outer.Height = draw_h
            outer.Fill = br_fill
            outer.Stroke = br_stroke
            outer.StrokeThickness = 1.0
            _Cn.SetLeft(outer, x0)
            _Cn.SetTop(outer, y0)
            self._canvas_set_zindex(outer, 1)
            canv.Children.Add(outer)

        bar_r = 3.5
        for dot in layout.get(u"dots") or []:
            cx, cy = _px(dot[u"fx"], dot[u"fy"])
            el = Ellipse()
            el.Width = bar_r * 2.0
            el.Height = bar_r * 2.0
            el.Fill = br_bar_cn
            el.Stroke = br_edge
            el.StrokeThickness = 0.6
            _Cn.SetLeft(el, cx - bar_r)
            _Cn.SetTop(el, cy - bar_r)
            self._canvas_set_zindex(el, 10)
            canv.Children.Add(el)

        if not use_enc_layout:
            self._draw_cabezal_stirrup_overlay_preview(
                canv, layout, wid, extremo, _px, conf_type=conf_type,
                bulk_preview=bulk_preview,
            )
            self._draw_cabezal_tie_overlay_preview(
                canv, layout, wid, extremo, _px, conf_type=conf_type,
                bulk_preview=bulk_preview,
            )

        if not enc_ctx:
            if mirror:
                self._draw_cabezal_break_line(
                    canv, x0, y0, draw_h, br_edge, 12, face_right=False,
                )
            else:
                self._draw_cabezal_break_line(
                    canv, x1, y0, draw_h, br_edge, 12, face_right=True,
                )

    def _draw_cabezal_encuentro_preview_labels(
        self,
        canv,
        layout,
        preview_layers,
        n_capas,
        ex_cfg,
        _px,
        x0,
        y0,
        draw_w,
        draw_h,
        enc_tipo=u"L",
        zone_rect=None,
    ):
        """Etiquetas de capas (n×ø) y ejes en preview encuentro L/T."""
        from System.Windows.Controls import Canvas as _Cn, TextBlock
        from System.Windows.Media import SolidColorBrush, Color

        br_axis = SolidColorBrush(Color.FromRgb(167, 139, 250))
        br_cap = SolidColorBrush(Color.FromRgb(125, 211, 252))
        br_sub = SolidColorBrush(Color.FromRgb(148, 163, 184))

        if zone_rect:
            zx = float(x0) + float(zone_rect.get(u"x", 0.0))
            zy = float(y0) + float(zone_rect.get(u"y", 0.0))
            zw = float(zone_rect.get(u"w", draw_w))
            zh = float(zone_rect.get(u"h", draw_h))
            axis_specs = (
                (u"Det.", zx + 2.0, zy + 2.0),
                (u"Sel.", zx + 2.0, zy + zh - 10.0),
            )
        else:
            axis_specs = (
                (u"Det.", x0 + 2.0, None),
                (u"Sel.", x0 + 2.0, None),
            )

        for spec in axis_specs:
            text = spec[0]
            tb = TextBlock()
            tb.Text = text
            tb.Foreground = br_axis
            tb.FontSize = 7.0
            if spec[2] is None:
                lx, ty = _px(0.04, 0.06 if text == u"Det." else 0.94)
                _Cn.SetLeft(tb, spec[1])
                _Cn.SetTop(tb, ty - 5.0)
            else:
                _Cn.SetLeft(tb, spec[1])
                _Cn.SetTop(tb, spec[2])
            self._canvas_set_zindex(tb, 12)
            canv.Children.Add(tb)

        try:
            pitch_v = float((ex_cfg or {}).get(u"pitch_equitativo_mm") or 0.0)
        except Exception:
            pitch_v = 0.0
        sub_tb = TextBlock()
        enc_label = u"Encuentro {0}".format(enc_tipo or u"L")
        sub_tb.Text = (
            u"{0} \u00b7 paso {1:.0f} mm".format(enc_label, pitch_v)
            if pitch_v > 0.5
            else enc_label
        )
        sub_tb.Foreground = br_sub
        sub_tb.FontSize = 7.0
        _Cn.SetLeft(sub_tb, x0 + 2.0)
        _Cn.SetTop(sub_tb, y0 + draw_h + 1.0)
        self._canvas_set_zindex(sub_tb, 12)
        canv.Children.Add(sub_tb)

        dots = layout.get(u"dots") or []
        by_layer = {}
        for dot in dots:
            try:
                li = int(dot.get(u"layer_index", 0))
            except Exception:
                li = 0
            by_layer.setdefault(li, []).append(dot)

        layers = list(preview_layers or [])
        try:
            n_show = max(1, min(int(n_capas), len(layers)))
        except Exception:
            n_show = len(layers)

        for i in range(n_show):
            ly = layers[i] if i < len(layers) else {}
            try:
                nb = int(ly.get(u"n_bars", 2))
            except Exception:
                nb = 2
            diam = self._diam_label_for_combo(None, ly.get(u"bar_type_id")) or u"?"
            col_dots = by_layer.get(i) or []
            if col_dots:
                avg_fx = sum(float(d[u"fx"]) for d in col_dots) / float(len(col_dots))
                avg_fy = sum(float(d[u"fy"]) for d in col_dots) / float(len(col_dots))
                cx, cy = _px(avg_fx, avg_fy)
            else:
                cx = x0 + draw_w * (0.2 + 0.6 * float(i) / max(1, n_show - 1))
                cy = y0 + draw_h * 0.5
            cap_tb = TextBlock()
            cap_tb.Text = u"C{0} {1}\u00d7{2}".format(i + 1, nb, diam)
            cap_tb.Foreground = br_cap
            cap_tb.FontSize = 7.0
            _Cn.SetLeft(cap_tb, cx - 14.0)
            if zone_rect:
                cap_y = float(y0) + float(zone_rect.get(u"y", 0.0)) + float(zone_rect.get(u"h", draw_h)) + 2.0
            else:
                cap_y = y0 + draw_h - 11.0
            _Cn.SetTop(cap_tb, cap_y)
            self._canvas_set_zindex(cap_tb, 12)
            canv.Children.Add(cap_tb)

    def _build_cabezal_tramo_placeholder(self, wid, wall, extremo, row_h):
        from System.Windows.Controls import Border, TextBlock, StackPanel
        from System.Windows import (
            Thickness,
            FontWeights,
            HorizontalAlignment,
            VerticalAlignment,
            CornerRadius,
        )
        from System.Windows.Media import SolidColorBrush, Color, Brushes

        pal = self._cabezal_ui_palette(extremo, u"unit")
        seg = self._cabezal_segment_for_wall(wid, extremo)
        try:
            seg_id = int(seg.get(u"id", 0)) + 1
        except Exception:
            seg_id = 1
        owner_wid = self._cabezal_owner_wid_for(wid, extremo)
        wall_indices = seg.get(u"wall_indices") or []
        muro_lbl = u"M{0}".format(
            u",".join(str(int(i) + 1) for i in wall_indices),
        )

        cap_w = float(self._cabezal_cap_col_px())
        wrap = Border()
        wrap.Width = cap_w
        wrap.MinWidth = cap_w
        wrap.MaxWidth = cap_w
        wrap.MinHeight = max(40.0, float(row_h))
        wrap.Background = Brushes.Transparent
        wrap.BorderBrush = pal[u"sep"]
        wrap.BorderThickness = Thickness(1)
        wrap.Padding = Thickness(6, 6, 6, 6)
        wrap.VerticalAlignment = VerticalAlignment.Stretch
        wrap.HorizontalAlignment = HorizontalAlignment.Center
        try:
            wrap.CornerRadius = CornerRadius(4.0)
        except Exception:
            pass

        sp = StackPanel()
        sp.VerticalAlignment = VerticalAlignment.Center
        sp.HorizontalAlignment = HorizontalAlignment.Center

        tb_t = TextBlock()
        tb_t.Text = u"T{0}".format(seg_id)
        tb_t.Foreground = SolidColorBrush(Color.FromRgb(126, 184, 208))
        tb_t.FontSize = 14.0
        tb_t.FontWeight = FontWeights.Bold
        tb_t.HorizontalAlignment = HorizontalAlignment.Center
        sp.Children.Add(tb_t)

        tb_m = TextBlock()
        tb_m.Text = muro_lbl
        tb_m.Foreground = SolidColorBrush(Color.FromRgb(100, 116, 139))
        tb_m.FontSize = 8.0
        tb_m.HorizontalAlignment = HorizontalAlignment.Center
        tb_m.Margin = Thickness(0, 4, 0, 0)
        sp.Children.Add(tb_m)

        wrap.Child = sp
        try:
            wrap.ToolTip = (
                u"Tramo T{0} ({1}) — editar en fila del muro propietario (Id {2})".format(
                    seg_id, muro_lbl, owner_wid,
                )
            )
        except Exception:
            pass
        return wrap

    def _create_cabezal_bar_length_warn_footer(self):
        """Pie del controlador: aviso compacto variante B (>12 m)."""
        from System.Windows.Controls import Border, StackPanel, TextBlock, Orientation
        from System.Windows import (
            Thickness,
            FontWeights,
            HorizontalAlignment,
            VerticalAlignment,
            TextWrapping,
            Visibility,
            CornerRadius,
        )
        from System.Windows.Media import SolidColorBrush, Color

        footer = Border()
        footer.Margin = Thickness(0, 6, 0, 0)
        footer.Padding = Thickness(0, 6, 0, 0)
        footer.BorderBrush = SolidColorBrush(Color.FromArgb(40, 255, 255, 255))
        footer.BorderThickness = Thickness(0, 1, 0, 0)
        footer.Visibility = Visibility.Collapsed
        footer.HorizontalAlignment = HorizontalAlignment.Stretch

        inner = Border()
        inner.Padding = Thickness(5, 5, 5, 5)
        inner.Background = SolidColorBrush(Color.FromArgb(31, 251, 191, 36))
        inner.BorderBrush = SolidColorBrush(Color.FromRgb(217, 119, 6))
        inner.BorderThickness = Thickness(1)
        try:
            inner.CornerRadius = CornerRadius(4.0)
        except Exception:
            pass

        row = StackPanel()
        row.Orientation = Orientation.Horizontal
        row.HorizontalAlignment = HorizontalAlignment.Stretch

        badge = Border()
        badge.Background = SolidColorBrush(Color.FromRgb(251, 191, 36))
        badge.Padding = Thickness(5, 2, 5, 2)
        badge.Margin = Thickness(0, 0, 6, 0)
        badge.VerticalAlignment = VerticalAlignment.Center
        try:
            badge.CornerRadius = CornerRadius(3.0)
        except Exception:
            pass
        badge_tb = TextBlock()
        badge_tb.Text = u">12 m"
        badge_tb.FontSize = 8.0
        badge_tb.FontWeight = FontWeights.Bold
        badge_tb.Foreground = SolidColorBrush(Color.FromRgb(69, 26, 3))
        badge.Child = badge_tb

        msg_tb = TextBlock()
        msg_tb.TextWrapping = TextWrapping.Wrap
        msg_tb.FontSize = 9.0
        msg_tb.Foreground = SolidColorBrush(Color.FromRgb(252, 211, 77))
        msg_tb.VerticalAlignment = VerticalAlignment.Center
        msg_tb.HorizontalAlignment = HorizontalAlignment.Stretch

        row.Children.Add(badge)
        row.Children.Add(msg_tb)
        inner.Child = row
        footer.Child = inner
        return footer, msg_tb

    def _refresh_cabezal_bar_length_warn(self, wid, extremo):
        """Actualiza aviso de largo máximo comercial al pie del controlador de tramo."""
        if cabezal is None:
            return
        ui = self._cabezal_ui_ext(wid, extremo)
        footer = ui.get(u"bar_length_warn_footer")
        msg_tb = ui.get(u"bar_length_warn_msg_tb")
        wrap = ui.get(u"cab_wrap")
        if footer is None or msg_tb is None:
            return

        from System.Windows import Visibility
        from System.Windows.Media import SolidColorBrush, Color

        seg = self._cabezal_segment_for_wall(wid, extremo)
        walls = getattr(self, u"walls_ordered", []) or []
        exceeds, L_mm = cabezal.cabezal_tramo_bar_length_status(walls, seg)
        pal = self._cabezal_ui_palette(extremo, u"unit")
        warn_br = SolidColorBrush(Color.FromRgb(217, 119, 6))

        if exceeds:
            footer.Visibility = Visibility.Visible
            msg_tb.Text = cabezal.cabezal_tramo_bar_length_warn_compact_text(L_mm)
            footer.ToolTip = cabezal.cabezal_tramo_bar_length_warn_tooltip(L_mm)
            if wrap is not None:
                wrap.BorderBrush = warn_br
        else:
            footer.Visibility = Visibility.Collapsed
            try:
                footer.ToolTip = None
            except Exception:
                pass
            if wrap is not None:
                wrap.BorderBrush = pal.get(u"sep") or warn_br

    def _refresh_cabezal_bar_length_warns_for_extremo(self, extremo):
        """Recalcula avisos de largo en todos los controladores de tramo de un extremo."""
        seen = set()
        for seg in self._cabezal_segments_for_extremo(extremo):
            try:
                owner_idx = int(seg.get(u"owner_index", 0))
            except Exception:
                owner_idx = 0
            walls = getattr(self, u"walls_ordered", []) or []
            if not (0 <= owner_idx < len(walls)):
                continue
            try:
                owner_wid = _wall_id_int(walls[owner_idx])
            except Exception:
                continue
            if owner_wid in seen:
                continue
            seen.add(owner_wid)
            self._refresh_cabezal_bar_length_warn(owner_wid, extremo)

    def _build_cabezal_extremo_cap(self, wid, wall, extremo, row_h, mirror_preview=False):
        from System.Windows.Controls import (
            Border,
            TextBlock,
            Canvas,
            Grid,
            StackPanel,
            ScrollViewer,
            RowDefinition,
            ColumnDefinition,
            ScrollBarVisibility,
        )
        from System.Windows import (
            Thickness,
            FontWeights,
            GridLength,
            GridUnitType,
            HorizontalAlignment,
            VerticalAlignment,
            CornerRadius,
            TextWrapping,
        )
        from System.Windows.Media import SolidColorBrush, Color, Brushes

        pal = self._cabezal_ui_palette(extremo, u"unit")
        sep_br = pal[u"sep"]

        cfg = self._cabezal_by_wall_id.get(wid) or {}
        ex_cfg = cfg.get(extremo) or cabezal.default_cabezal_extremo_config()
        layers0 = ex_cfg.get(u"layers") or [{u"n_bars": 2}, {u"n_bars": 2}]
        try:
            n0 = int(ex_cfg.get(u"n_capas", len(layers0)))
        except Exception:
            n0 = len(layers0)
        min_cap0 = self._cabezal_min_capas_for_wall_extremo(wid, extremo)
        n0 = max(min_cap0, min(cabezal.CABEZAL_MAX_CAPAS, n0))

        cap_w = float(self._cabezal_cap_col_px())
        left_w, right_w, preview_w, content_w, split_gap = (
            self._cabezal_extremo_split_layout_px(cap_w)
        )
        _, prev_h = self._cabezal_preview_canvas_size_px(wid, extremo)
        row_h_px = max(40.0, float(row_h))
        canvas_w = self._cabezal_seccion_canvas_width_px(right_w)

        block_gap = float(getattr(self, "_CABEZAL_EXTREMO_BLOCK_GAP_PX", 10.0))
        inner_gap = float(getattr(self, "_CABEZAL_EXTREMO_INNER_GAP_PX", 8.0))
        toolbar_gap = float(getattr(self, "_CABEZAL_EXTREMO_TOOLBAR_GAP_PX", 6.0))
        if toolbar_gap > 4.0:
            toolbar_gap = 4.0
        section_hdr_h = 9.0 + inner_gap * 0.35
        wrap_pad_top = float(getattr(self, u"_CABEZAL_UNIT_WRAP_PAD_TOP_PX", 4.0))
        wrap_pad_side = float(getattr(self, u"_CABEZAL_UNIT_WRAP_PAD_SIDE_PX", 8.0))

        wrap = Border()
        wrap.Width = cap_w
        wrap.MinWidth = cap_w
        wrap.MaxWidth = cap_w
        wrap.MinHeight = row_h_px
        wrap.Background = pal[u"panel_fill"]
        wrap.BorderBrush = sep_br
        wrap.BorderThickness = Thickness(1)
        wrap.Padding = Thickness(wrap_pad_side, wrap_pad_top, wrap_pad_side, wrap_pad_side)
        wrap.Margin = Thickness(0, 0, 0, 0)
        wrap.VerticalAlignment = VerticalAlignment.Stretch
        wrap.HorizontalAlignment = HorizontalAlignment.Center
        try:
            wrap.CornerRadius = CornerRadius(4.0)
        except Exception:
            pass

        content = StackPanel()
        content.VerticalAlignment = VerticalAlignment.Top
        content.HorizontalAlignment = HorizontalAlignment.Stretch
        content.Margin = Thickness(0, 0, 0, 0)
        try:
            content.Width = content_w
            content.MinWidth = content_w
            content.MaxWidth = content_w
        except Exception:
            pass

        seg = self._cabezal_segment_for_wall(wid, extremo)
        try:
            seg_id = int(seg.get(u"id", 0)) + 1
        except Exception:
            seg_id = 1
        wall_indices = seg.get(u"wall_indices") or []
        tramo_toolbar_text = u"Tramo T{0} · {1}".format(
            seg_id,
            u",".join(u"M{0}".format(int(i) + 1) for i in wall_indices),
        )

        toolbar_grid = Grid()
        toolbar_grid.HorizontalAlignment = HorizontalAlignment.Stretch
        toolbar_grid.VerticalAlignment = VerticalAlignment.Top
        toolbar_grid.Margin = Thickness(0, 0, 0, toolbar_gap)
        self._init_cabezal_toolbar_compact_grid(toolbar_grid, content_w)

        controls_scroll = ScrollViewer()
        controls_scroll.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
        controls_scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        controls_scroll.HorizontalAlignment = HorizontalAlignment.Left
        controls_scroll.VerticalAlignment = VerticalAlignment.Top
        controls_scroll.Margin = Thickness(0, 0, 0, 0)
        try:
            controls_scroll.Width = left_w
            controls_scroll.MaxWidth = left_w
        except Exception:
            pass

        controls_grid = Grid()
        controls_grid.VerticalAlignment = VerticalAlignment.Top
        controls_grid.HorizontalAlignment = HorizontalAlignment.Left
        controls_grid.ClipToBounds = False
        self._init_cabezal_ctrl_grid_columns(controls_grid)
        controls_scroll.Content = controls_grid
        try:
            controls_scroll.MaxHeight = self._cabezal_layers_scroll_max_height_px()
        except Exception:
            pass

        cv = Canvas()
        cv.Height = prev_h
        cv.Width = canvas_w
        cv.MinWidth = canvas_w
        cv.MaxWidth = canvas_w
        cv.MinHeight = prev_h
        cv.MaxHeight = prev_h
        cv.Background = Brushes.Transparent
        cv.VerticalAlignment = VerticalAlignment.Top
        cv.HorizontalAlignment = HorizontalAlignment.Center
        cv.Margin = Thickness(0, 0, 0, block_gap)
        try:
            cv.ClipToBounds = True
        except Exception:
            pass

        conf_lbl = TextBlock()
        conf_lbl.Text = u"Confinamiento"
        conf_lbl.Foreground = pal[u"text_section"]
        conf_lbl.FontSize = 9.0
        conf_lbl.FontWeight = FontWeights.SemiBold
        conf_lbl.HorizontalAlignment = HorizontalAlignment.Stretch
        conf_lbl.Margin = Thickness(0, 0, 0, inner_gap * 0.75)

        conf_cb = self._create_cabezal_confinement_combo(wid, extremo)
        try:
            conf_cb.MaxWidth = right_w
            conf_cb.HorizontalAlignment = HorizontalAlignment.Stretch
            conf_cb.Margin = Thickness(0, 0, 0, 0)
        except Exception:
            pass

        enc_hint_lbl = TextBlock()
        enc_hint_lbl.Text = (
            u"Capas en esp. detectado \u00b7 barras en esp. seleccionado"
        )
        enc_hint_lbl.Foreground = pal[u"text_caption"]
        enc_hint_lbl.FontSize = 8.0
        enc_hint_lbl.TextWrapping = TextWrapping.Wrap
        enc_hint_lbl.HorizontalAlignment = HorizontalAlignment.Stretch
        enc_hint_lbl.Margin = Thickness(0, inner_gap * 0.5, 0, 0)
        try:
            from System.Windows import Visibility
            enc_hint_lbl.Visibility = Visibility.Collapsed
        except Exception:
            pass

        seccion_group = Border()
        seccion_group.Background = pal[u"canvas_group_bg"]
        seccion_group.BorderBrush = pal[u"canvas_group_border"]
        seccion_group.BorderThickness = Thickness(1)
        seccion_group.Padding = Thickness(self._cabezal_seccion_group_pad_px())
        try:
            seccion_group.CornerRadius = CornerRadius(3.0)
            seccion_group.ClipToBounds = True
        except Exception:
            pass
        seccion_inner = StackPanel()
        seccion_inner.HorizontalAlignment = HorizontalAlignment.Stretch
        seccion_inner.Children.Add(cv)
        seccion_inner.Children.Add(enc_hint_lbl)
        seccion_inner.Children.Add(conf_lbl)
        seccion_inner.Children.Add(conf_cb)
        seccion_group.Child = seccion_inner

        armado_hdr = TextBlock()
        armado_hdr.Text = u"Armado"
        armado_hdr.Foreground = pal[u"text_section"]
        armado_hdr.FontSize = 9.0
        armado_hdr.FontWeight = FontWeights.SemiBold
        armado_hdr.Margin = Thickness(0, 0, 0, 0)
        armado_hdr.VerticalAlignment = VerticalAlignment.Bottom

        seccion_hdr = TextBlock()
        seccion_hdr.Text = u"Secci\u00f3n"
        seccion_hdr.Foreground = pal[u"text_section"]
        seccion_hdr.FontSize = 9.0
        seccion_hdr.FontWeight = FontWeights.SemiBold
        seccion_hdr.Margin = Thickness(0, 0, 0, 0)
        seccion_hdr.VerticalAlignment = VerticalAlignment.Bottom

        armado_col = StackPanel()
        armado_col.HorizontalAlignment = HorizontalAlignment.Left
        armado_col.VerticalAlignment = VerticalAlignment.Top
        try:
            armado_col.Width = left_w
            armado_col.MaxWidth = left_w
        except Exception:
            pass
        armado_col.Children.Add(controls_scroll)

        seccion_col = StackPanel()
        seccion_col.HorizontalAlignment = HorizontalAlignment.Stretch
        seccion_col.VerticalAlignment = VerticalAlignment.Top
        try:
            seccion_col.Width = right_w
            seccion_col.MaxWidth = right_w
        except Exception:
            pass
        seccion_col.Children.Add(seccion_group)

        split_sep = self._cabezal_split_vline(sep_br, split_gap)

        split_hdr_grid = Grid()
        split_hdr_grid.HorizontalAlignment = HorizontalAlignment.Stretch
        split_hdr_grid.VerticalAlignment = VerticalAlignment.Top
        split_hdr_grid.Margin = Thickness(0, 0, 0, inner_gap * 0.35)
        try:
            split_hdr_grid.Width = content_w
            split_hdr_grid.MinWidth = content_w
            split_hdr_grid.MaxWidth = content_w
            split_hdr_grid.Height = section_hdr_h
            split_hdr_grid.MinHeight = section_hdr_h
        except Exception:
            pass
        cd_h_arm = ColumnDefinition()
        cd_h_arm.Width = GridLength(left_w, GridUnitType.Pixel)
        cd_h_sep = ColumnDefinition()
        cd_h_sep.Width = GridLength(split_gap, GridUnitType.Pixel)
        cd_h_sec = ColumnDefinition()
        cd_h_sec.Width = GridLength(right_w, GridUnitType.Pixel)
        split_hdr_grid.ColumnDefinitions.Add(cd_h_arm)
        split_hdr_grid.ColumnDefinitions.Add(cd_h_sep)
        split_hdr_grid.ColumnDefinitions.Add(cd_h_sec)
        Grid.SetColumn(armado_hdr, 0)
        Grid.SetColumn(seccion_hdr, 2)
        split_hdr_grid.Children.Add(armado_hdr)
        split_hdr_grid.Children.Add(seccion_hdr)

        split_grid = Grid()
        split_grid.HorizontalAlignment = HorizontalAlignment.Stretch
        split_grid.VerticalAlignment = VerticalAlignment.Top
        split_grid.Margin = Thickness(0, 0, 0, 0)
        try:
            split_grid.Width = content_w
            split_grid.MinWidth = content_w
            split_grid.MaxWidth = content_w
        except Exception:
            pass
        cd_arm = ColumnDefinition()
        cd_arm.Width = GridLength(left_w, GridUnitType.Pixel)
        cd_sep = ColumnDefinition()
        cd_sep.Width = GridLength(split_gap, GridUnitType.Pixel)
        cd_sec = ColumnDefinition()
        cd_sec.Width = GridLength(right_w, GridUnitType.Pixel)
        split_grid.ColumnDefinitions.Add(cd_arm)
        split_grid.ColumnDefinitions.Add(cd_sep)
        split_grid.ColumnDefinitions.Add(cd_sec)
        Grid.SetColumn(armado_col, 0)
        Grid.SetColumn(split_sep, 1)
        Grid.SetColumn(seccion_col, 2)
        split_grid.Children.Add(armado_col)
        split_grid.Children.Add(split_sep)
        split_grid.Children.Add(seccion_col)

        bar_warn_footer, bar_warn_msg_tb = self._create_cabezal_bar_length_warn_footer()

        tramo_accent = Border()
        tramo_accent.Height = 2.0
        tramo_accent.Margin = Thickness(0, 0, 0, 4.0)
        tramo_accent.HorizontalAlignment = HorizontalAlignment.Stretch
        tr_rgb = self._cabezal_tramo_color_rgb(seg)
        tramo_accent.Background = SolidColorBrush(
            Color.FromRgb(int(tr_rgb[0]), int(tr_rgb[1]), int(tr_rgb[2])),
        )

        content.Children.Add(tramo_accent)
        content.Children.Add(toolbar_grid)
        content.Children.Add(split_hdr_grid)
        content.Children.Add(split_grid)
        content.Children.Add(bar_warn_footer)
        wrap.Child = content

        def _on_capas_change(_v):
            try:
                self._apply_cabezal_capas_to_wall_extremo(wid, extremo, int(_v))
            except Exception:
                pass

        cap_step = self._create_cabezal_stepper(
            self._cabezal_min_capas_for_wall_extremo(wid, extremo),
            cabezal.CABEZAL_MAX_CAPAS,
            n0,
            on_change=_on_capas_change,
            palette=pal,
        )
        self._fill_cabezal_toolbar_row(
            toolbar_grid, cap_step[u"panel"], pal, tramo_toolbar_text,
        )

        ui = self._cabezal_ui_ext(wid, extremo)
        ui[u"tramo_toolbar_text"] = tramo_toolbar_text
        ui[u"tramo_toolbar_text_base"] = tramo_toolbar_text
        ui[u"preview_canvas"] = cv
        ui[u"controls_grid"] = controls_grid
        ui[u"controls_scroll"] = controls_scroll
        ui[u"toolbar_grid"] = toolbar_grid
        ui[u"capas_grid"] = toolbar_grid
        ui[u"layers_row_start"] = 1
        ui[u"capas_stepper"] = cap_step
        ui[u"capas_value_tb"] = cap_step[u"value_tb"]
        ui[u"layer_steppers"] = []
        ui[u"layer_value_tbs"] = []
        ui[u"cab_wrap"] = wrap
        ui[u"extremo"] = extremo
        ui[u"mirror_preview"] = bool(mirror_preview)
        ui[u"confinement_cb"] = conf_cb
        ui[u"confinement_lbl"] = conf_lbl
        ui[u"encuentro_hint_lbl"] = enc_hint_lbl
        ui[u"preview_canvas_w_px"] = canvas_w
        ui[u"split_grid"] = split_grid
        ui[u"bar_length_warn_footer"] = bar_warn_footer
        ui[u"bar_length_warn_msg_tb"] = bar_warn_msg_tb

        wall_ui = self._cabezal_ui_by_wall_id.setdefault(wid, {})
        wall_ui[extremo] = ui
        try:
            seg_key = int(seg.get(u"id", 0))
        except Exception:
            seg_key = 0
        self._cabezal_ui_by_segment.setdefault(extremo, {})[seg_key] = ui
        ui[u"owner_wid"] = int(wid)
        ui[u"segment_id"] = seg_key

        try:
            extremo_lbl = (
                u"Inicio"
                if extremo == cabezal.CABEZAL_EXTREMO_INICIO
                else u"Final"
            )
            wrap.ToolTip = u"Cabezal — {0} · tramo T{1}".format(
                extremo_lbl, seg_id,
            )
        except Exception:
            pass

        def _on_preview_size(sender, args, w=wid, ex=extremo):
            self._request_cabezal_preview_refresh(w, ex, delay_ms=120)

        try:
            from System.Windows import RoutedEventHandler as _REH
            cv.SizeChanged += _REH(_on_preview_size)
        except Exception:
            pass

        self._rebuild_cabezal_layer_sliders(wid, extremo, redistribute=False)
        self._refresh_cabezal_confinement_combo(wid, extremo)
        self._refresh_cabezal_encuentro_ui_state(wid, extremo)
        self._apply_cabezal_armado_ui_state(wid, extremo)
        self._request_cabezal_preview_refresh(wid, extremo, debounce=False)
        self._refresh_cabezal_bar_length_warn(wid, extremo)
        return wrap

    def _create_cabezal_confinement_combo(self, wid, extremo):
        from System.Windows.Controls import ComboBox
        from System.Windows import Thickness, HorizontalAlignment

        cb = ComboBox()
        cb.IsEditable = False
        cb.Margin = Thickness(0, 0, 0, 4)
        cb.HorizontalAlignment = HorizontalAlignment.Stretch
        self._apply_flat_combo(cb, stretch=True)
        ui = self._cabezal_ui_ext(wid, extremo)
        ui[u"confinement_cb"] = cb

        def _on_conf_change(sender, args, w=wid, ex=extremo):
            if getattr(self, "_suppress_cabezal_confinement_cb", False):
                return
            self._on_cabezal_confinement_changed(w, ex)

        try:
            from System.Windows.Controls import SelectionChangedEventHandler as _SCEH
            from System.Windows import RoutedEventHandler as _REH
            cb.SelectionChanged += _SCEH(_on_conf_change)
            cb.DropDownClosed += _REH(_on_conf_change)
        except Exception:
            pass
        return cb

    def _cabezal_bulk_conf_options(self, extremo=None):
        if cabezal is None:
            return []
        n_capas = 2
        if extremo is not None:
            try:
                n_capas = self._read_cabezal_bulk_capas(extremo)
            except Exception:
                pass
        return cabezal.cabezal_confinement_options(n_capas)

    def _cabezal_bulk_default_conf_type(self, n_capas):
        """Confinamiento inicial del panel masivo (2–6 capas → Tipo 1)."""
        if cabezal is None:
            return u"none"
        try:
            n = int(n_capas)
        except Exception:
            n = 0
        if cabezal.cabezal_confinement_scenario_applies(n):
            return cabezal.CABEZAL_CONFINEMENT_TIE_LAYER_1
        return cabezal.CABEZAL_CONFINEMENT_NONE

    def _cabezal_confinement_for_capas_change(self, n_capas, prev_conf=None):
        """Confinamiento por defecto al variar CAPAS (2–6 capas → Tipo 1)."""
        if cabezal is None:
            return {}
        conf = cabezal.default_cabezal_confinement_config(n_capas)
        if isinstance(prev_conf, dict):
            try:
                conf[u"stirrup_diam_mm"] = float(
                    prev_conf.get(u"stirrup_diam_mm", conf.get(u"stirrup_diam_mm")),
                )
            except Exception:
                pass
        return conf

    def _refresh_cabezal_bulk_confinement_combo(self, extremo):
        if cabezal is None:
            return
        ui = (getattr(self, "_cabezal_bulk_ui", None) or {}).get(extremo) or {}
        cb = ui.get(u"conf_cb")
        if cb is None:
            return
        n_capas = self._read_cabezal_bulk_capas(extremo)
        bulk = self._cabezal_bulk_ui.setdefault(extremo, {})
        try:
            prev_n = int(bulk.get(u"_conf_n_capas", -1))
        except Exception:
            prev_n = -1
        try:
            n_now = int(n_capas)
        except Exception:
            n_now = cabezal.CABEZAL_MIN_CAPAS
        if prev_n != n_now:
            bulk[u"conf_type"] = self._cabezal_bulk_default_conf_type(n_now)
            bulk[u"_conf_n_capas"] = n_now
        prev_type = bulk.get(u"conf_type")
        if prev_type is None:
            prev_type = self._cabezal_bulk_default_conf_type(n_capas)
        conf = cabezal.normalize_cabezal_confinement({u"type": prev_type}, n_capas)
        opts = self._cabezal_bulk_conf_options(extremo)
        target = conf.get(u"type") or self._cabezal_bulk_default_conf_type(n_capas)
        bulk[u"conf_type"] = target
        self._suppress_cabezal_bulk_conf_cb = True
        try:
            try:
                cb.Items.Clear()
                for _val, label in opts:
                    cb.Items.Add(label)
            except Exception:
                pass
            for i, (val, _lbl) in enumerate(opts):
                if val == target:
                    try:
                        cb.SelectedIndex = i
                    except Exception:
                        pass
                    break
            else:
                try:
                    cb.SelectedIndex = 0
                except Exception:
                    pass
        finally:
            self._suppress_cabezal_bulk_conf_cb = False
        try:
            if cabezal.cabezal_confinement_scenario_applies(n_capas):
                cb.ToolTip = None
            else:
                cb.ToolTip = (
                    u"Tipo 1 y Tipo 2 solo aplican con 2 a 6 capas "
                    u"(escenario actual)."
                )
        except Exception:
            pass

    def _read_cabezal_bulk_conf_type(self, extremo):
        ui = (getattr(self, "_cabezal_bulk_ui", None) or {}).get(extremo) or {}
        cb = ui.get(u"conf_cb")
        opts = self._cabezal_bulk_conf_options(extremo)
        n_capas = self._read_cabezal_bulk_capas(extremo)
        default = self._cabezal_bulk_default_conf_type(n_capas)
        bulk = self._cabezal_bulk_ui.setdefault(extremo, {})
        if cb is None or not opts:
            bulk[u"conf_type"] = default
            return default
        try:
            idx = int(cb.SelectedIndex)
            if 0 <= idx < len(opts):
                val = opts[idx][0]
                bulk[u"conf_type"] = val
                return val
        except Exception:
            pass
        try:
            txt = unicode(cb.SelectedItem or cb.Text or u"")
            for val, lbl in opts:
                if lbl == txt or txt in lbl or lbl.startswith(txt):
                    bulk[u"conf_type"] = val
                    return val
        except Exception:
            pass
        stored = bulk.get(u"conf_type")
        if stored is not None:
            for val, _lbl in opts:
                if val == stored:
                    return stored
        bulk[u"conf_type"] = default
        return default

    def _read_cabezal_bulk_empalme_params(self, extremo):
        ui = (getattr(self, "_cabezal_bulk_ui", None) or {}).get(extremo) or {}
        start_1 = 2
        step = 2
        tb_s = ui.get(u"alt_start_tb")
        tb_t = ui.get(u"alt_step_tb")
        if tb_s is not None:
            try:
                start_1 = int(float(unicode(tb_s.Text or u"").strip().replace(u",", u".")))
            except Exception:
                pass
        if tb_t is not None:
            try:
                step = int(float(unicode(tb_t.Text or u"").strip().replace(u",", u".")))
            except Exception:
                pass
        return max(1, int(start_1)), max(1, int(step))

    def _read_cabezal_bulk_capas(self, extremo):
        if cabezal is None:
            return 2
        ui = (getattr(self, "_cabezal_bulk_ui", None) or {}).get(extremo) or {}
        cap_step = ui.get(u"capas_stepper")
        if cap_step is not None:
            try:
                return max(
                    cabezal.CABEZAL_MIN_CAPAS,
                    min(cabezal.CABEZAL_MAX_CAPAS, int(cap_step[u"get_value"]())),
                )
            except Exception:
                pass
        return cabezal.CABEZAL_MIN_CAPAS

    def _cabezal_refresh_encuentro_pitch(self, wid, extremo):
        """Recalcula paso equitativo de capas en encuentro L."""
        if cabezal is None or _cab_enc_l is None or _vec_ext is None:
            return
        cfg = self._cabezal_by_wall_id.get(wid) or {}
        ex_cfg = cfg.get(extremo) or {}
        if not cabezal.cabezal_extremo_es_encuentro_l(ex_cfg):
            return
        doc = getattr(self, u"doc", None)
        if doc is None:
            return
        wall = None
        for w in getattr(self, u"walls_ordered", []) or []:
            try:
                if _wall_id_int(w) == int(wid):
                    wall = w
                    break
            except Exception:
                pass
        if wall is None:
            return
        neighbor = None
        try:
            vid = ex_cfg.get(u"vecino_wall_id")
            if vid is not None:
                el = doc.GetElement(ElementId(int(vid)))
                if isinstance(el, Wall):
                    neighbor = el
        except Exception:
            neighbor = None
        if neighbor is None:
            try:
                neighbor = _vec_ext.vecino_principal_encuentro_l(doc, wall, extremo)
            except Exception:
                neighbor = None
        if neighbor is None:
            return
        try:
            _cab_enc_l.cabezal_encuentro_l_refresh_pitch_in_cfg(
                ex_cfg, wall, neighbor,
            )
        except Exception:
            pass

    def _apply_cabezal_capas_to_wall_extremo(self, wid, extremo, n_capas):
        """Misma escritura que el stepper CAPAS por fila (sin bloquear edición puntual)."""
        if cabezal is None:
            return
        try:
            n_capas = int(n_capas)
        except Exception:
            n_capas = cabezal.CABEZAL_MIN_CAPAS
        cfg = self._cabezal_by_wall_id.setdefault(wid, {})
        ex_cfg = cfg.setdefault(extremo, cabezal.default_cabezal_extremo_config())
        min_capas = (
            2
            if cabezal.cabezal_extremo_es_encuentro_l(ex_cfg)
            else cabezal.CABEZAL_MIN_CAPAS
        )
        try:
            n_capas = max(
                min_capas,
                min(cabezal.CABEZAL_MAX_CAPAS, n_capas),
            )
        except Exception:
            n_capas = min_capas
        prev_conf = ex_cfg.get(u"confinement")
        try:
            self._sync_cabezal_extremo_from_ui(wid, extremo)
        except Exception:
            pass
        ex_cfg[u"n_capas"] = n_capas
        try:
            cabezal._normalize_cabezal_extremo_layers(ex_cfg)
        except Exception:
            pass
        try:
            if cabezal.cabezal_extremo_es_encuentro_l(ex_cfg):
                ex_cfg[u"confinement"] = cabezal.normalize_cabezal_confinement(
                    {u"type": cabezal.CABEZAL_CONFINEMENT_NONE}, n_capas,
                )
            else:
                ex_cfg[u"confinement"] = self._cabezal_confinement_for_capas_change(
                    n_capas, prev_conf,
                )
        except Exception:
            pass
        self._cabezal_refresh_encuentro_pitch(wid, extremo)
        ui = self._cabezal_ui_ext(wid, extremo)
        cap_step = ui.get(u"capas_stepper")
        if cap_step is not None:
            self._suppress_cabezal_stepper = True
            try:
                cap_step[u"set_value"](n_capas)
            finally:
                self._suppress_cabezal_stepper = False
        else:
            tb_cap = ui.get(u"capas_value_tb")
            if tb_cap is not None:
                try:
                    tb_cap.Text = str(n_capas)
                except Exception:
                    pass
        self._apply_cabezal_layer_rows_active_state(wid, extremo)
        self._refresh_cabezal_confinement_combo(wid, extremo)
        self._refresh_cabezal_encuentro_ui_state(wid, extremo)
        self._request_cabezal_preview_refresh(wid, extremo)

    def _apply_cabezal_bulk_capas(self, extremo):
        if cabezal is None:
            return
        n_capas = self._read_cabezal_bulk_capas(extremo)
        walls = self._cabezal_walls_for_extremo(extremo)
        if not walls:
            return
        for wid, _ri, _w in walls:
            self._apply_cabezal_capas_to_wall_extremo(wid, extremo, n_capas)
        try:
            self._refresh_cabezal_segment_diam_combos(extremo)
        except Exception:
            pass
        ex_lbl = u"Inicio" if extremo == cabezal.CABEZAL_EXTREMO_INICIO else u"Final"
        self._set_estado(
            u"Capas masivas ({0}): {1} capa(s) en {2} tramo(s). Ajuste puntual por tramo.".format(
                ex_lbl, n_capas, len(walls),
            ),
        )
        try:
            self._update_cabezal_bulk_active_layers(extremo)
            self._refresh_cabezal_bulk_confinement_combo(extremo)
            self._refresh_cabezal_bulk_preview(extremo)
        except Exception:
            pass

    def _cabezal_bulk_layer_seed(self, extremo, layer_idx):
        """Valores iniciales n/ø para fila masiva desde el primer muro del extremo."""
        if cabezal is None:
            return 2, None
        od = getattr(self, "_walls_display_order", []) or []
        if not od:
            return 2, None
        try:
            walls = self._cabezal_walls_for_extremo(extremo)
            wid = walls[0][0] if walls else _wall_id_int(od[0])
        except Exception:
            try:
                wid = _wall_id_int(od[0])
            except Exception:
                return 2, None
        cfg = self._cabezal_by_wall_id.get(wid) or {}
        ex_cfg = cfg.get(extremo) or cabezal.default_cabezal_extremo_config()
        layers = ex_cfg.get(u"layers") or []
        li = int(layer_idx)
        ly = layers[li] if li < len(layers) else {}
        try:
            nb = int(ly.get(u"n_bars", 2))
        except Exception:
            nb = 2
        nb = max(
            cabezal.CABEZAL_MIN_BARRAS_POR_CAPA,
            min(cabezal.CABEZAL_MAX_BARRAS_POR_CAPA, nb),
        )
        bid = ly.get(u"bar_type_id")
        if bid is None or bid == ElementId.InvalidElementId:
            try:
                bid = self._cabezal_layer_diam_id(wid, extremo, li)
            except Exception:
                bid = ex_cfg.get(u"bar_type_id")
        return nb, bid

    def _read_cabezal_bulk_layer_armado(self, extremo):
        """Lista (n_bars, bar_type_id) por índice de capa desde controles masivos."""
        if cabezal is None:
            return []
        ui = (getattr(self, "_cabezal_bulk_ui", None) or {}).get(extremo) or {}
        steppers = ui.get(u"layer_n_steppers") or []
        diam_cbs = ui.get(u"layer_diam_cbs") or []
        out = []
        n_rows = max(len(steppers), len(diam_cbs), cabezal.CABEZAL_MAX_CAPAS)
        for i in range(n_rows):
            nb = cabezal.CABEZAL_MIN_BARRAS_POR_CAPA
            if i < len(steppers) and steppers[i] is not None:
                try:
                    nb = int(steppers[i][u"get_value"]())
                except Exception:
                    pass
            nb = max(
                cabezal.CABEZAL_MIN_BARRAS_POR_CAPA,
                min(cabezal.CABEZAL_MAX_BARRAS_POR_CAPA, int(nb)),
            )
            bid = ElementId.InvalidElementId
            if i < len(diam_cbs) and diam_cbs[i] is not None:
                bid = self._read_diam_combo_id(diam_cbs[i])
            out.append((nb, bid))
        return out

    def _apply_cabezal_n_bars_to_wall_layer(self, wid, extremo, layer_idx, n_bars):
        if cabezal is None:
            return
        try:
            li = int(layer_idx)
            nb = max(
                cabezal.CABEZAL_MIN_BARRAS_POR_CAPA,
                min(cabezal.CABEZAL_MAX_BARRAS_POR_CAPA, int(n_bars)),
            )
        except Exception:
            return
        n_capas = self._cabezal_n_capas_from_ui(wid, extremo)
        if li >= n_capas:
            return
        cfg = self._cabezal_by_wall_id.setdefault(wid, {})
        ex_cfg = cfg.setdefault(extremo, cabezal.default_cabezal_extremo_config())
        layers = list(ex_cfg.get(u"layers") or [])
        while len(layers) <= li:
            layers.append(
                cabezal.default_cabezal_layer_config(2, ex_cfg.get(u"bar_type_id")),
            )
        layers[li] = cabezal._normalize_cabezal_layer_dict(
            dict(layers[li], **{u"n_bars": nb}),
            ex_cfg.get(u"bar_type_id"),
        )
        ex_cfg[u"layers"] = layers
        cabezal._normalize_cabezal_extremo_layers(ex_cfg)
        ui = self._cabezal_ui_ext(wid, extremo)
        steppers = ui.get(u"layer_steppers") or []
        if li < len(steppers) and steppers[li] is not None:
            self._suppress_cabezal_stepper = True
            try:
                steppers[li][u"set_value"](nb)
            finally:
                self._suppress_cabezal_stepper = False
        self._request_cabezal_preview_refresh(wid, extremo)

    def _apply_cabezal_diam_to_owner_layer(self, wid, extremo, layer_idx, bar_type_id):
        """Escribe ø en la capa del muro."""
        if cabezal is None:
            return False
        try:
            li = int(layer_idx)
        except Exception:
            return False
        n_capas = self._cabezal_n_capas_from_ui(wid, extremo)
        if li >= n_capas:
            return False
        if bar_type_id is None or bar_type_id == ElementId.InvalidElementId:
            return False
        try:
            self._sync_cabezal_extremo_from_ui(wid, extremo)
        except Exception:
            pass
        cfg = self._cabezal_by_wall_id.setdefault(wid, {})
        ex_cfg = cfg.setdefault(extremo, cabezal.default_cabezal_extremo_config())
        layers = list(ex_cfg.get(u"layers") or [])
        while len(layers) <= li:
            layers.append(
                cabezal.default_cabezal_layer_config(2, ex_cfg.get(u"bar_type_id")),
            )
        layers[li] = cabezal._normalize_cabezal_layer_dict(
            dict(layers[li], **{u"bar_type_id": bar_type_id}),
            ex_cfg.get(u"bar_type_id"),
        )
        ex_cfg[u"layers"] = layers
        cabezal._normalize_cabezal_extremo_layers(ex_cfg)
        ui = self._cabezal_ui_ext(wid, extremo)
        diam_cbs = ui.get(u"layer_diam_cbs") or []
        if li < len(diam_cbs) and diam_cbs[li] is not None:
            self._set_diam_combo_selected_id(diam_cbs[li], bar_type_id)
        return True

    def _apply_cabezal_bulk_layer_armado(self, extremo):
        if cabezal is None:
            return
        n_capas_bulk = self._read_cabezal_bulk_capas(extremo)
        armado = self._read_cabezal_bulk_layer_armado(extremo)
        walls = self._cabezal_walls_for_extremo(extremo)
        if not walls or not armado:
            return
        n_layers_apply = min(
            n_capas_bulk,
            len(armado),
            cabezal.CABEZAL_MAX_CAPAS,
        )
        n_muros = 0
        for wid, _ri, _w in walls:
            n_muros += 1
            try:
                n_capas_wall = self._cabezal_n_capas_from_ui(wid, extremo)
            except Exception:
                n_capas_wall = n_capas_bulk
            n_apply = min(n_layers_apply, int(n_capas_wall))
            for li in range(n_apply):
                nb, _bid = armado[li]
                self._apply_cabezal_n_bars_to_wall_layer(wid, extremo, li, nb)
            for li in range(n_apply):
                _nb, bid = armado[li]
                self._apply_cabezal_diam_to_owner_layer(wid, extremo, li, bid)
        try:
            self._refresh_cabezal_segment_diam_combos(extremo)
        except Exception:
            pass
        ex_lbl = u"Inicio" if extremo == cabezal.CABEZAL_EXTREMO_INICIO else u"Final"
        self._set_estado(
            u"Armado masivo ({0}): n/ø en {1} capa(s), {2} tramo(s). "
            u"Edición por tramo activa.".format(
                ex_lbl, n_layers_apply, n_muros,
            ),
        )
        try:
            self._refresh_cabezal_bulk_preview(extremo)
        except Exception:
            pass

    def _cabezal_empalme_pattern_row_indices(self, n_rows, start_tramo_1, step):
        """Filas del canvas (0=arriba) que coinciden con patrón tramo base 1 = muro inferior."""
        if n_rows < 1:
            return set()
        start_slot = int(start_tramo_1) - 1
        if start_slot >= n_rows:
            return set()
        out = set()
        s = start_slot
        st = max(1, int(step))
        while s < n_rows:
            ri = n_rows - 1 - s
            if 0 <= ri < n_rows:
                out.add(ri)
            s += st
        return out

    def _cabezal_walls_for_extremo(self, extremo):
        od = getattr(self, "_walls_display_order", []) or []
        out = []
        seen = set()
        if cabezal is None:
            return out
        for seg in self._cabezal_segments_for_extremo(extremo):
            try:
                align_idx = self._cabezal_segment_align_stack_index(seg)
            except Exception:
                align_idx = int(seg.get(u"owner_index", 0))
            walls = getattr(self, u"walls_ordered", []) or []
            if not (0 <= align_idx < len(walls)):
                continue
            if align_idx in seen:
                continue
            seen.add(align_idx)
            w = walls[align_idx]
            try:
                wid = _wall_id_int(w)
            except Exception:
                continue
            ri = self._cabezal_display_row_for_stack_index(align_idx)
            out.append((wid, ri, w))
        return out

    def _cabezal_ordered_index_for_wall(self, wall):
        if wall is None:
            return -1
        try:
            target = _wall_id_int(wall)
        except Exception:
            return -1
        for i, w in enumerate(getattr(self, u"walls_ordered", []) or []):
            if w is None:
                continue
            try:
                if _wall_id_int(w) == target:
                    return i
            except Exception:
                pass
        return -1

    def _cabezal_auto_troceo_for_wall(self, wid, extremo):
        if cabezal is None:
            return False
        idx = -1
        for i, w in enumerate(getattr(self, u"walls_ordered", []) or []):
            try:
                if _wall_id_int(w) == int(wid):
                    idx = i
                    break
            except Exception:
                pass
        if idx < 1:
            return False
        try:
            flags = cabezal.compute_auto_troceo_por_muro_flags(
                self.walls_ordered,
                getattr(self, u"_stacked_layout", None),
                extremo,
            )
            return bool(flags[idx]) if idx < len(flags) else False
        except Exception:
            return False

    def _cabezal_troceo_is_manual(self, wid, extremo):
        if cabezal is None:
            return False
        cfg = self._cabezal_by_wall_id.get(wid) or {}
        ex_cfg = cfg.get(extremo) or {}
        return cabezal.troceo_override_from_extremo_cfg(ex_cfg) is not None

    def _sync_all_cabezal_troceo_auto(self):
        if not self._uses_cabezal_panels() or cabezal is None:
            return
        self._invalidate_cabezal_segments_cache()
        for ex in cabezal.CABEZAL_EXTREMOS:
            try:
                cabezal.sync_troceo_effective_for_extremo(
                    self.walls_ordered,
                    self._cabezal_by_wall_id,
                    getattr(self, u"_stacked_layout", None),
                    ex,
                )
            except Exception:
                pass

    def _refresh_cabezal_troceo_checkbox(self, wid, extremo):
        ui = self._cabezal_ui_ext(wid, extremo)
        chk = ui.get(u"troceo_por_muro_chk")
        cfg = self._cabezal_by_wall_id.get(wid) or {}
        ex_cfg = cfg.get(extremo) or {}
        on = bool(ex_cfg.get(u"troceo_por_muro"))
        if chk is not None:
            try:
                self._suppress_cabezal_empalme_chk = True
                chk.IsChecked = on
            finally:
                self._suppress_cabezal_empalme_chk = False
            try:
                self._apply_toggle_mini_visual(ui, on, animate=False)
            except Exception:
                pass
        try:
            from System.Windows.Media import SolidColorBrush, Color
            sep_br = SolidColorBrush(Color.FromRgb(33, 70, 92))
            acc_br = SolidColorBrush(Color.FromRgb(34, 211, 238))
            wrap = ui.get(u"cab_wrap")
            if wrap is not None:
                wrap.BorderBrush = acc_br if on else sep_br
        except Exception:
            pass

    def _set_cabezal_troceo_por_muro(self, wid, extremo, value, update_border=True):
        if cabezal is None:
            return
        self._invalidate_cabezal_segments_cache()
        cfg = self._cabezal_by_wall_id.setdefault(wid, {})
        ex_cfg = cfg.setdefault(extremo, cabezal.default_cabezal_extremo_config())
        auto = self._cabezal_auto_troceo_for_wall(wid, extremo)
        v = bool(value)
        ex_cfg[u"troceo_por_muro_override"] = None if v == auto else v
        ex_cfg[u"troceo_auto_geom"] = auto
        ex_cfg[u"troceo_por_muro"] = v
        if update_border:
            self._refresh_cabezal_troceo_checkbox(wid, extremo)

    def _cabezal_pie_selector_caption(self, wid, extremo):
        if cabezal is None:
            return u"Auto"
        cfg = self._cabezal_by_wall_id.get(wid) or {}
        ex_cfg = cfg.get(extremo) or {}
        auto = self._cabezal_auto_troceo_for_wall(wid, extremo)
        ov = cabezal.troceo_override_from_extremo_cfg(ex_cfg)
        if ov is None:
            return u"Auto·" if auto else u"Auto"
        if ov:
            return u"Tramo"
        return u"Cont."

    def _cycle_cabezal_troceo_at_pie(self, wid, extremo):
        if cabezal is None:
            return
        cfg = self._cabezal_by_wall_id.setdefault(wid, {})
        ex_cfg = cfg.setdefault(extremo, cabezal.default_cabezal_extremo_config())
        auto = self._cabezal_auto_troceo_for_wall(wid, extremo)
        ov = cabezal.troceo_override_from_extremo_cfg(ex_cfg)
        if ov is None:
            new_ov = True
        elif ov is True:
            new_ov = False
        else:
            new_ov = None
        if new_ov is None:
            ex_cfg[u"troceo_por_muro_override"] = None
            ex_cfg[u"troceo_auto_geom"] = auto
            ex_cfg[u"troceo_por_muro"] = bool(auto)
        else:
            ex_cfg[u"troceo_por_muro_override"] = bool(new_ov)
            ex_cfg[u"troceo_auto_geom"] = auto
            ex_cfg[u"troceo_por_muro"] = bool(new_ov)
        self._refresh_cabezal_troceo_checkbox(wid, extremo)
        try:
            self._redraw_wall_elevation_canvas(wid)
        except Exception:
            pass
        try:
            self._schedule_cabezal_empalme_followup(wid, extremo, delay_ms=80)
        except Exception:
            pass

    def _toggle_cabezal_troceo_at_pie(self, wid, extremo):
        self._cycle_cabezal_troceo_at_pie(wid, extremo)

    def _restore_cabezal_troceo_auto(self, extremo):
        if cabezal is None:
            return
        self._invalidate_cabezal_segments_cache(extremo)
        for wall in getattr(self, u"walls_ordered", []) or []:
            try:
                wid = _wall_id_int(wall)
            except Exception:
                continue
            cfg = self._cabezal_by_wall_id.setdefault(wid, {})
            ex_cfg = cfg.setdefault(extremo, cabezal.default_cabezal_extremo_config())
            ex_cfg[u"troceo_por_muro_override"] = None
        cabezal.sync_troceo_effective_for_extremo(
            self.walls_ordered,
            self._cabezal_by_wall_id,
            getattr(self, u"_stacked_layout", None),
            extremo,
        )
        for wid, _ri, _w in self._cabezal_walls_for_extremo(extremo):
            self._refresh_cabezal_troceo_checkbox(wid, extremo)
        for w in getattr(self, u"_walls_display_order", []) or []:
            try:
                self._redraw_wall_elevation_canvas(_wall_id_int(w))
            except Exception:
                pass
        self._rebuild_cabezal_all_walls_for_extremo(extremo)
        ex_lbl = u"Inicio" if extremo == cabezal.CABEZAL_EXTREMO_INICIO else u"Final"
        self._set_estado(
            u"Tramos restaurados a geometría ({0}).".format(ex_lbl),
        )

    def _force_single_cabezal_tramo(self, extremo):
        if cabezal is None:
            return
        n = 0
        for wall in getattr(self, u"walls_ordered", []) or []:
            try:
                wid = _wall_id_int(wall)
            except Exception:
                continue
            auto = self._cabezal_auto_troceo_for_wall(wid, extremo)
            cfg = self._cabezal_by_wall_id.setdefault(wid, {})
            ex_cfg = cfg.setdefault(extremo, cabezal.default_cabezal_extremo_config())
            ex_cfg[u"troceo_por_muro_override"] = False if auto else None
            ex_cfg[u"troceo_auto_geom"] = auto
            ex_cfg[u"troceo_por_muro"] = False
            n += 1
        for wid, _ri, _w in self._cabezal_walls_for_extremo(extremo):
            self._refresh_cabezal_troceo_checkbox(wid, extremo)
        for w in getattr(self, u"_walls_display_order", []) or []:
            try:
                self._redraw_wall_elevation_canvas(_wall_id_int(w))
            except Exception:
                pass
        self._rebuild_cabezal_all_walls_for_extremo(extremo)
        ex_lbl = u"Inicio" if extremo == cabezal.CABEZAL_EXTREMO_INICIO else u"Final"
        self._set_estado(
            u"Un solo tramo ({0}): sin cortes en pie en todo el stack.".format(ex_lbl),
        )

    def _apply_cabezal_bulk_confinement(self, extremo):
        if cabezal is None:
            return
        conf_type = self._read_cabezal_bulk_conf_type(extremo)
        walls = self._cabezal_walls_for_extremo(extremo)
        if not walls:
            return
        n = 0
        for wid, _ri, _w in walls:
            cfg = self._cabezal_by_wall_id.setdefault(wid, {})
            ex_cfg = cfg.setdefault(extremo, cabezal.default_cabezal_extremo_config())
            try:
                n_capas = self._cabezal_n_capas_from_ui(wid, extremo)
            except Exception:
                n_capas = cabezal.CABEZAL_MIN_CAPAS
            prev = ex_cfg.get(u"confinement") or {}
            if not isinstance(prev, dict):
                prev = {}
            merged = dict(prev)
            merged[u"type"] = conf_type
            ex_cfg[u"confinement"] = cabezal.normalize_cabezal_confinement(
                merged, n_capas,
            )
            self._refresh_cabezal_confinement_combo(wid, extremo)
            n += 1
        ex_lbl = u"Inicio" if extremo == cabezal.CABEZAL_EXTREMO_INICIO else u"Final"
        self._set_estado(
            u"Confinamiento masivo ({0}): {1} tramo(s). Ajuste puntual por tramo.".format(
                ex_lbl, n,
            ),
        )
        try:
            self._refresh_cabezal_bulk_preview(extremo)
        except Exception:
            pass

    def _apply_cabezal_bulk_alternate_empalme(self, extremo):
        if cabezal is None:
            return
        od = getattr(self, "_walls_display_order", []) or []
        n_rows = len(od)
        start_1, step = self._read_cabezal_bulk_empalme_params(extremo)
        pattern = self._cabezal_empalme_pattern_row_indices(n_rows, start_1, step)
        if not pattern:
            try:
                TaskDialog.Show(
                    u"Arainco: Cabezal muros",
                    u"No hay tramos que coincidan con «desde» / «cada». "
                    u"Tramo 1 = muro inferior (abajo en la lista).",
                )
            except Exception:
                pass
            return
        toggled = 0
        # Patrón por muro (no por tramo): un tramo sin empalmes agrupa todos los muros
        # y _cabezal_walls_for_extremo solo devuelve el owner de cada segmento.
        for ri, w in enumerate(od):
            if ri not in pattern:
                continue
            try:
                wid = _wall_id_int(w)
            except Exception:
                continue
            cfg = self._cabezal_by_wall_id.get(wid) or {}
            ex_cfg = cfg.get(extremo) or {}
            cur = bool(ex_cfg.get(u"troceo_por_muro"))
            self._set_cabezal_troceo_por_muro(wid, extremo, not cur)
            toggled += 1
        for w in od:
            try:
                self._redraw_wall_elevation_canvas(_wall_id_int(w))
            except Exception:
                pass
        self._rebuild_cabezal_all_walls_for_extremo(extremo)
        ex_lbl = u"Inicio" if extremo == cabezal.CABEZAL_EXTREMO_INICIO else u"Final"
        self._set_estado(
            u"Empalme alternado ({0}): {1} tramo(s). Edición por fila sigue activa.".format(
                ex_lbl, toggled,
            ),
        )

    def _apply_cabezal_bulk_clear_empalme(self, extremo):
        self._restore_cabezal_troceo_auto(extremo)

    def _cabezal_bulk_ref_wall(self):
        od = getattr(self, "_walls_display_order", []) or []
        return od[0] if od else None

    def _cabezal_layers_from_bulk_ui(self, extremo):
        if cabezal is None:
            return []
        n_capas = self._read_cabezal_bulk_capas(extremo)
        armado = self._read_cabezal_bulk_layer_armado(extremo)
        fb = ElementId.InvalidElementId
        wall = self._cabezal_bulk_ref_wall()
        if wall is not None:
            try:
                wid = _wall_id_int(wall)
                cfg = self._cabezal_by_wall_id.get(wid) or {}
                ex_cfg = cfg.get(extremo) or cabezal.default_cabezal_extremo_config()
                fb = ex_cfg.get(u"bar_type_id")
            except Exception:
                pass
        layers = []
        for i in range(n_capas):
            if i < len(armado):
                nb, bid = armado[i]
            else:
                nb, bid = cabezal.CABEZAL_MIN_BARRAS_POR_CAPA, fb
            if bid is None or bid == ElementId.InvalidElementId:
                bid = fb
            layers.append(
                cabezal._normalize_cabezal_layer_dict(
                    {u"n_bars": nb, u"bar_type_id": bid}, fb,
                )
            )
        return layers

    def _refresh_cabezal_bulk_preview(self, extremo):
        if cabezal is None:
            return
        ui = (getattr(self, "_cabezal_bulk_ui", None) or {}).get(extremo) or {}
        cv = ui.get(u"preview_canvas")
        if cv is None:
            return
        wall = self._cabezal_bulk_ref_wall()
        if wall is None:
            try:
                cv.Children.Clear()
            except Exception:
                pass
            return
        wid = _wall_id_int(wall)
        pw = ui.get(u"preview_canvas_w_px")
        layers = self._cabezal_layers_from_bulk_ui(extremo)
        conf_type = self._read_cabezal_bulk_conf_type(extremo)
        mirror = False
        try:
            wui = self._cabezal_ui_ext(wid, extremo)
            mirror = bool(wui.get(u"mirror_preview", False))
        except Exception:
            pass
        self._draw_cabezal_preview_canvas(
            cv,
            wall,
            wid,
            extremo,
            layers_override=layers,
            conf_type_override=conf_type,
            canvas_w_px=pw,
            mirror_preview=mirror,
            bulk_preview=True,
        )

    def _update_cabezal_bulk_active_layers(self, extremo):
        """Atenúa filas n/ø fuera del rango CAPAS del stepper masivo (solo visual)."""
        if cabezal is None:
            return
        from System.Windows import FontWeights

        n = self._read_cabezal_bulk_capas(extremo)
        ui = (getattr(self, "_cabezal_bulk_ui", None) or {}).get(extremo) or {}
        rows = ui.get(u"layer_row_ui") or []
        pal = self._cabezal_ui_palette(extremo, u"bulk")
        for i, row in enumerate(rows):
            active = i < n
            op = 1.0 if active else pal[u"disabled_opacity"]
            cap_lbl = row.get(u"cap_lbl")
            if cap_lbl is not None:
                try:
                    cap_lbl.Foreground = (
                        pal[u"layer_active"] if active else pal[u"layer_inactive"]
                    )
                    cap_lbl.FontSize = 11.0 if active else 10.0
                    cap_lbl.FontWeight = (
                        FontWeights.SemiBold if active else FontWeights.Normal
                    )
                except Exception:
                    pass
            for key in (u"n_panel", u"diam_cb"):
                el = row.get(key)
                if el is not None:
                    try:
                        el.Opacity = op
                    except Exception:
                        pass
            diam_cb = row.get(u"diam_cb")
            if diam_cb is not None:
                try:
                    self._apply_cabezal_diam_combo_state(
                        diam_cb, layer_active=active, palette=pal,
                    )
                except Exception:
                    pass

    def _build_cabezal_bulk_side_panel(self, extremo, parent_grid, grid_column):
        from System.Windows.Controls import (
            ColumnDefinition,
            Grid,
            Orientation,
            StackPanel,
            TextBlock,
            TextBox,
        )
        from System.Windows import (
            GridLength,
            GridUnitType,
            HorizontalAlignment,
            Thickness,
            FontWeights,
            VerticalAlignment,
        )

        if cabezal is None or parent_grid is None:
            return

        pal = self._cabezal_ui_palette(extremo, u"bulk")
        ex_lbl = self._cabezal_extremo_ui_label(extremo)

        fields = Grid()
        fields.Margin = Thickness(0)
        fields.HorizontalAlignment = HorizontalAlignment.Stretch
        fields.VerticalAlignment = VerticalAlignment.Top
        cd_e0 = ColumnDefinition()
        cd_e0.Width = GridLength(1.0, GridUnitType.Star)
        cd_e1 = ColumnDefinition()
        cd_e1.Width = GridLength(8.0, GridUnitType.Pixel)
        cd_e2 = ColumnDefinition()
        cd_e2.Width = GridLength(1.0, GridUnitType.Star)
        for cd in (cd_e0, cd_e1, cd_e2):
            fields.ColumnDefinitions.Add(cd)

        def _emp_field(col, title, default_txt, tip):
            sp = StackPanel()
            sp.Orientation = Orientation.Vertical
            sp.HorizontalAlignment = HorizontalAlignment.Stretch
            sp.VerticalAlignment = VerticalAlignment.Top
            lt = TextBlock()
            lt.Text = title
            lt.Foreground = pal[u"text_control"]
            lt.FontSize = 9.0
            lt.FontWeight = FontWeights.Normal
            lt.Margin = Thickness(0, 0, 0, 3)
            sp.Children.Add(lt)
            tb = TextBox()
            tb.Text = default_txt
            tb.MaxLength = 4
            tb.FontSize = 10.0
            tb.Padding = Thickness(4, 3, 4, 3)
            tb.ToolTip = tip
            tb.HorizontalAlignment = HorizontalAlignment.Stretch
            self._apply_bimtools_textbox(tb)
            sp.Children.Add(tb)
            Grid.SetColumn(sp, col)
            fields.Children.Add(sp)
            return tb

        tb_start = _emp_field(
            0,
            u"Desde tramo",
            u"2",
            u"Tramo 1 = muro inferior (abajo en la lista).",
        )
        tb_step = _emp_field(
            2,
            u"Cada N",
            u"2",
            u"2 = alternar uno s\u00ed / uno no en el patr\u00f3n.",
        )

        btn_apply = self._build_mesh_apply_button(
            u"Aplicar a todos",
            click_handler=lambda s, e, ex=extremo: self._apply_cabezal_bulk_alternate_empalme(ex),
        )
        btn_apply.ToolTip = u"Aplica el patr\u00f3n alternado desde / cada N a todos los tramos."
        btn_clear = self._build_mesh_apply_button(
            u"Limpiar",
            click_handler=lambda s, e, ex=extremo: self._apply_cabezal_bulk_clear_empalme(ex),
        )
        btn_clear.ToolTip = (
            u"Restaura todos los tramos del extremo a troceo autom\u00e1tico "
            u"(quita empalmes alternados manuales)."
        )

        btn_row = Grid()
        btn_row.Margin = Thickness(0)
        btn_row.HorizontalAlignment = HorizontalAlignment.Stretch
        cd_apply = ColumnDefinition()
        cd_apply.Width = GridLength(1.0, GridUnitType.Star)
        cd_gap = ColumnDefinition()
        cd_gap.Width = GridLength(6.0, GridUnitType.Pixel)
        cd_clear = ColumnDefinition()
        cd_clear.Width = GridLength(1.0, GridUnitType.Star)
        for cd in (cd_apply, cd_gap, cd_clear):
            btn_row.ColumnDefinitions.Add(cd)
        Grid.SetColumn(btn_apply, 0)
        Grid.SetColumn(btn_clear, 2)
        btn_row.Children.Add(btn_apply)
        btn_row.Children.Add(btn_clear)

        toggle_chk = self._add_bulk_armado_toggle(
            extremo, parent=None, header_mode=True,
        )
        on = True
        od = getattr(self, "_walls_display_order", []) or []
        if od:
            try:
                on = self._cabezal_armado_activo_cfg(_wall_id_int(od[0]), extremo)
            except Exception:
                on = True
        self._sync_bulk_armado_toggle_widgets(extremo, on, animate=False)

        shell = self._build_bulk_action_card(
            title=u"Armado {0}".format(ex_lbl),
            panel_fill=pal[u"panel_fill"],
            sep_br=pal[u"sep"],
            accent_br=pal[u"bar_accent"],
            body=fields,
            toggle_chk=toggle_chk,
            apply_button=btn_row,
            stretch=True,
        )

        Grid.SetColumn(shell, int(grid_column))
        parent_grid.Children.Add(shell)

        bulk = self._cabezal_bulk_ui.setdefault(extremo, {})
        bulk[u"config_body"] = fields
        bulk[u"panel"] = shell
        bulk[u"alt_start_tb"] = tb_start
        bulk[u"alt_step_tb"] = tb_step
        bulk[u"btn_apply_pattern"] = btn_apply
        bulk[u"btn_clear"] = btn_clear
        self._set_cabezal_bulk_panel_enabled(extremo, on)

    def _apply_bulk_mesh_to_all_walls(self):
        """Copia el configurador global de mallas (centro bulk) a todos los tramos."""
        src_ctr = getattr(self, u"_bulk_mesh_ctr", None) or {}
        if not src_ctr:
            self._apply_mallas_to_all_from_first()
            return
        keys = self._mesh_control_keys()
        n = 0
        for w in getattr(self, u"_walls_display_order", []) or []:
            wid = _wall_id_int(w)
            dst_ctr = self._controls_by_wall_id.get(wid)
            if not dst_ctr:
                continue
            for k in keys:
                self._copy_combo_like(src_ctr.get(k), dst_ctr.get(k))
            n += 1
        self._set_estado(
            u"Mallas (bulk centro) aplicadas a {0} tramo(s).".format(n),
        )
        if self._is_unificado_mode():
            for w in getattr(self, u"_walls_display_order", []) or []:
                try:
                    self._sync_cabezal_confinement_from_malla_wall(
                        _wall_id_int(w), refresh_preview=True,
                    )
                except Exception:
                    pass

    def _build_unificado_bulk_mesh_panel(self, ctr, fg_hi, fg_lo):
        pal = self._mesh_ui_palette(u"bulk")
        body = self._build_mesh_params_row(
            ctr,
            u"ct_md",
            u"ct_ms",
            u"ct_id",
            u"ct_is",
            major_lbl=u"Vertical",
            minor_lbl=u"Horizontal",
            pal=pal,
        )
        self._bulk_mesh_params_body = body

        btn = self._build_mesh_apply_button(
            u"Aplicar a todos",
            click_handler=lambda s, e: self._apply_bulk_mesh_to_all_walls(),
        )

        toggle_chk = self._build_bulk_malla_activo_toggle(
            parent=None, header_mode=True,
        )

        return self._build_bulk_action_card(
            title=u"Configurador \u2014 Mallas",
            panel_fill=pal[u"panel_fill"],
            sep_br=pal[u"sep"],
            accent_br=pal[u"accent"],
            body=body,
            toggle_chk=toggle_chk,
            apply_button=btn,
            stretch=True,
        )

    def _build_cabezal_bulk_actions_panel(self):
        from System.Windows.Controls import Grid, ColumnDefinition, RowDefinition, TextBlock
        from System.Windows import (
            GridLength,
            GridUnitType,
            HorizontalAlignment,
            Thickness,
            FontWeights,
            Visibility,
        )
        from System.Windows.Media import SolidColorBrush, Color

        if not self._uses_cabezal_panels() or cabezal is None:
            return
        bdr = self._win.FindName(u"BdrCabezalBulkActions")
        root = self._win.FindName(u"GrdCabezalBulkActions")
        if root is None:
            return
        root.Children.Clear()
        root.ColumnDefinitions.Clear()
        root.RowDefinitions.Clear()

        try:
            if bdr is not None:
                bdr.Visibility = Visibility.Visible
                bdr.Padding = Thickness(8, 8, 26, 12)
        except Exception:
            pass

        self._cabezal_bulk_ui = {}
        self._bulk_armado_toggle_hosts = {}

        rd0 = RowDefinition()
        rd0.Height = GridLength.Auto
        root.RowDefinitions.Add(rd0)

        from System.Windows.Controls import StackPanel

        header = StackPanel()
        header.Margin = Thickness(0, 0, 0, 8)
        title = TextBlock()
        title.Text = u"Configuraciones Globales"
        title.Foreground = SolidColorBrush(Color.FromRgb(100, 116, 139))
        title.FontSize = 9.0
        title.FontWeight = FontWeights.Bold
        title.Margin = Thickness(0, 0, 0, 4)
        header.Children.Add(title)
        legend = self._build_wall_thickness_legend_panel()
        if legend is not None:
            header.Children.Add(legend)
        Grid.SetRow(header, 0)
        root.Children.Add(header)

        rd1 = RowDefinition()
        rd1.Height = GridLength.Auto
        root.RowDefinitions.Add(rd1)

        outer = Grid()
        outer.Margin = Thickness(0)
        Grid.SetRow(outer, 1)

        ruler_px = self._ruler_col_px()
        elev_gap = float(getattr(self, "_CABEZAL_ELEV_GAP_PX", 8.0))
        col_offset = 0
        if ruler_px > 0:
            cd_r = ColumnDefinition()
            cd_r.Width = GridLength(ruler_px, GridUnitType.Pixel)
            outer.ColumnDefinitions.Add(cd_r)
            col_offset = 1

        cd_ini = ColumnDefinition()
        cd_ini.Width = GridLength(float(self._cabezal_cap_col_px()), GridUnitType.Pixel)
        cd_g1 = ColumnDefinition()
        cd_g1.Width = GridLength(elev_gap, GridUnitType.Pixel)
        cd_el = ColumnDefinition()
        cd_el.Width = GridLength(float(self._PREVIEW_ELEV_COL_PX), GridUnitType.Pixel)
        cd_g2 = ColumnDefinition()
        cd_g2.Width = GridLength(elev_gap, GridUnitType.Pixel)
        cd_fin = ColumnDefinition()
        cd_fin.Width = GridLength(float(self._cabezal_cap_col_px()), GridUnitType.Pixel)
        outer.ColumnDefinitions.Add(cd_ini)
        outer.ColumnDefinitions.Add(cd_g1)
        outer.ColumnDefinitions.Add(cd_el)
        outer.ColumnDefinitions.Add(cd_g2)
        outer.ColumnDefinitions.Add(cd_fin)

        od = getattr(self, "_walls_display_order", []) or []
        if od and cabezal is not None:
            try:
                ex_left, ex_right = self._cabezal_extremos_lados_wall(
                    _wall_id_int(od[0]), 0,
                )
            except Exception:
                ex_left = cabezal.CABEZAL_EXTREMO_INICIO
                ex_right = cabezal.CABEZAL_EXTREMO_FIN
        else:
            ex_left = cabezal.CABEZAL_EXTREMO_INICIO if cabezal else u"inicio"
            ex_right = cabezal.CABEZAL_EXTREMO_FIN if cabezal else u"fin"

        self._build_cabezal_bulk_side_panel(ex_left, outer, col_offset + 0)

        if self._is_unificado_mode():
            fg_hi = SolidColorBrush(Color.FromRgb(232, 244, 248))
            fg_lo = SolidColorBrush(Color.FromRgb(149, 184, 204))
            self._bulk_mesh_ctr = {}
            mesh_bulk = self._build_unificado_bulk_mesh_panel(
                self._bulk_mesh_ctr, fg_hi, fg_lo,
            )
            if od:
                on_m = self._malla_activo_cfg(_wall_id_int(od[0]))
                self._sync_bulk_malla_activo_toggle(on_m, animate=False)
                self._set_bulk_mesh_panel_enabled(on_m)
            Grid.SetColumn(mesh_bulk, col_offset + 2)
            outer.Children.Add(mesh_bulk)

        self._build_cabezal_bulk_side_panel(ex_right, outer, col_offset + 4)

        root.Children.Add(outer)
        try:
            w = float(self._grid_content_width_px())
            root.Width = w
            root.MinWidth = w
            root.MaxWidth = w
            root.HorizontalAlignment = HorizontalAlignment.Center
        except Exception:
            pass

    def _build_wall_parameter_panels(self):
        from System.Windows.Controls import (
            Border,
            StackPanel,
            TextBlock,
            ComboBox,
            Orientation,
            Grid,
            Canvas,
            ColumnDefinition,
            RowDefinition,
        )
        from System.Windows import (
            CornerRadius,
            GridLength,
            GridUnitType,
            Thickness,
            FontWeights,
            HorizontalAlignment,
            TextAlignment,
            TextWrapping,
            VerticalAlignment,
        )
        from System.Windows.Media import SolidColorBrush, Color, Brushes

        root = self._win.FindName("GrdListaMuros")
        if root is None:
            return

        root.Children.Clear()
        root.RowDefinitions.Clear()
        root.ColumnDefinitions.Clear()
        self._controls_by_wall_id = {}
        self._malla_ui_by_wall_id = {}
        self._canvas_by_wall_id = {}
        self._right_by_wall_id = {}
        self._prev_wrap_by_wall_id = {}
        self._row_grid_by_wall_id = {}
        self._cabezal_ui_by_wall_id = {}
        self._cabezal_ui_by_segment = {}
        self._cabezal_stack_grid = None
        self._cabezal_pending_caps = []
        self._cabezal_mounted_caps = []
        self._cabezal_mounted_connectors = []
        self._cabezal_elev_host_by_wid = {}
        self._row_definitions = []
        self._master_grid = root

        if self._uses_cabezal_panels():
            try:
                border_parent = root.Parent
                if border_parent is not None and hasattr(border_parent, "Padding"):
                    border_parent.Padding = Thickness(8, 4, 8, 6)
            except Exception:
                pass
        self._length_scale_key = None
        self._preview_layout = None

        fg_hi = SolidColorBrush(Color.FromRgb(232, 244, 248))
        fg_lo = SolidColorBrush(Color.FromRgb(149, 184, 204))
        sep_br = SolidColorBrush(Color.FromRgb(33, 70, 92))

        self._build_column_headers_panel(fg_lo, sep_br)

        ruler_px = self._ruler_col_px()
        self._grid_col_offset = 0
        if ruler_px > 0:
            cd_ruler = ColumnDefinition()
            cd_ruler.Width = GridLength(ruler_px, GridUnitType.Pixel)
            root.ColumnDefinitions.Add(cd_ruler)
            self._grid_col_offset = 1

        cd0 = ColumnDefinition()
        cd0.Width = GridLength(float(self._preview_col_px), GridUnitType.Pixel)
        cd1 = ColumnDefinition()
        cd1.Width = GridLength(12.0, GridUnitType.Pixel)
        cd2 = ColumnDefinition()
        cd2.Width = GridLength(float(self._mesh_col_px), GridUnitType.Pixel)
        root.ColumnDefinitions.Add(cd0)
        if self._mesh_col_px > 0.0:
            root.ColumnDefinitions.Add(cd1)
            root.ColumnDefinitions.Add(cd2)
        self._apply_grid_content_width(root)

        try:
            root.UseLayoutRounding = True
            root.SnapsToDevicePixels = True
        except Exception:
            pass

        n_rows = len(self._walls_display_order)
        row_heights = self._compute_row_heights()
        self._last_row_heights = list(row_heights)
        if len(row_heights) < n_rows:
            row_heights.extend(
                [float(self._row_height_px)] * (n_rows - len(row_heights)),
            )

        cab_stack = None
        self._cabezal_pending_caps = []
        if self._uses_cabezal_panels() and cabezal is not None:
            elev_gap = float(getattr(self, "_CABEZAL_ELEV_GAP_PX", 8.0))
            cab_stack = Grid()
            cab_stack.Margin = Thickness(0)
            cab_stack.ClipToBounds = False
            cab_stack.VerticalAlignment = VerticalAlignment.Stretch
            cab_stack.HorizontalAlignment = HorizontalAlignment.Stretch
            for rh in row_heights:
                rd_cab = RowDefinition()
                rd_cab.Height = GridLength(float(rh), GridUnitType.Pixel)
                cab_stack.RowDefinitions.Add(rd_cab)
            cd_ini = ColumnDefinition()
            cd_ini.Width = GridLength(
                float(self._cabezal_cap_col_px()), GridUnitType.Pixel,
            )
            cd_gl = ColumnDefinition()
            cd_gl.Width = GridLength(elev_gap, GridUnitType.Pixel)
            cd_el = ColumnDefinition()
            cd_el.Width = GridLength(1.0, GridUnitType.Star)
            cd_gr = ColumnDefinition()
            cd_gr.Width = GridLength(elev_gap, GridUnitType.Pixel)
            cd_fin = ColumnDefinition()
            cd_fin.Width = GridLength(
                float(self._cabezal_cap_col_px()), GridUnitType.Pixel,
            )
            cab_stack.ColumnDefinitions.Add(cd_ini)
            cab_stack.ColumnDefinitions.Add(cd_gl)
            cab_stack.ColumnDefinitions.Add(cd_el)
            cab_stack.ColumnDefinitions.Add(cd_gr)
            cab_stack.ColumnDefinitions.Add(cd_fin)
            self._cabezal_stack_grid = cab_stack

        self._ruler_top_row = 0
        if self._uses_cabezal_panels() and ruler_px > 0:
            rd_top = RowDefinition()
            rd_top.Height = GridLength(20.0, GridUnitType.Pixel)
            root.RowDefinitions.Add(rd_top)
            self._ruler_top_row = 1

        for rh in row_heights:
            rd = RowDefinition()
            rd.Height = GridLength(float(rh), GridUnitType.Pixel)
            root.RowDefinitions.Add(rd)
            self._row_definitions.append(rd)

        mesh_pal = self._mesh_ui_palette()

        def _mesh_params_block(
            title, major_lbl, minor_lbl, md_k, ms_k, id_k, is_k, ctr, wall_id=None,
        ):
            inner = self._build_mesh_params_stack(
                ctr,
                md_k,
                ms_k,
                id_k,
                is_k,
                major_lbl=major_lbl,
                minor_lbl=minor_lbl,
                pal=mesh_pal,
                wire_malla_conf=self._is_unificado_mode(),
                wall_id=wall_id,
            )
            if not title:
                return inner
            from System.Windows.Controls import StackPanel, TextBlock
            from System.Windows import FontWeights, Thickness

            side = StackPanel()
            hdr = TextBlock()
            hdr.Text = title
            hdr.Foreground = mesh_pal[u"text_section"]
            hdr.FontSize = 10.0
            hdr.FontWeight = FontWeights.SemiBold
            hdr.Margin = Thickness(0, 0, 0, 4)
            side.Children.Add(hdr)
            side.Children.Add(inner)
            return side

        for ri, wall in enumerate(self._walls_display_order):
            wid = _wall_id_int(wall)
            ctr = {}
            grid_row = ri + self._ruler_top_row
            row_h = float(row_heights[ri]) if ri < len(row_heights) else float(
                getattr(self, "_CABEZAL_ROW_FIXED_PX", 280.0)
                if self._uses_cabezal_panels()
                else getattr(self, "_row_height_px", 200.0)
            )

            prev_wrap = Border()
            prev_wrap.Background = Brushes.Transparent
            if self._uses_cabezal_panels():
                prev_wrap.BorderThickness = Thickness(0)
                prev_wrap.Margin = Thickness(0)
                prev_wrap.Padding = Thickness(0)
            else:
                prev_wrap.BorderThickness = Thickness(0, 0, 0, 1)
                prev_wrap.BorderBrush = sep_br
                prev_wrap.Margin = Thickness(0, 0, 0, 2)
                prev_wrap.Padding = Thickness(0, 2, 0, 4)
            prev_wrap.VerticalAlignment = VerticalAlignment.Stretch
            prev_wrap.HorizontalAlignment = HorizontalAlignment.Stretch
            prev_wrap.MinHeight = row_h
            prev_wrap.ClipToBounds = False
            Grid.SetRow(prev_wrap, grid_row)
            Grid.SetColumn(prev_wrap, self._grid_col_offset)
            self._prev_wrap_by_wall_id[wid] = prev_wrap

            cv = Canvas()
            cv.Background = Brushes.Transparent
            cv.ClipToBounds = False
            cv.HorizontalAlignment = HorizontalAlignment.Stretch
            cv.VerticalAlignment = VerticalAlignment.Stretch
            cv.Margin = Thickness(4, 0, 4, 0)
            cv.Height = row_h
            if self._uses_cabezal_panels():
                elev_gap = float(getattr(self, "_CABEZAL_ELEV_GAP_PX", 8.0))
                elev_canvas_w = max(
                    40.0,
                    float(self._PREVIEW_ELEV_COL_PX) - 8.0,
                )
                cv.Width = elev_canvas_w
            self._canvas_by_wall_id[wid] = cv

            if self._uses_cabezal_panels() and cabezal is not None and cab_stack is not None:
                ex_izq, ex_der = self._cabezal_extremos_lados_wall(wid, ri)

                wall_ui = self._cabezal_ui_by_wall_id.setdefault(wid, {})
                wall_ui[u"lado_izq"] = ex_izq
                wall_ui[u"lado_der"] = ex_der

                _div_brush = SolidColorBrush(Color.FromRgb(33, 70, 92))
                for _div_col in (1, 3):
                    _div = Border()
                    _div.Width = 1.0
                    _div.Background = _div_brush
                    _div.HorizontalAlignment = HorizontalAlignment.Center
                    _div.VerticalAlignment = VerticalAlignment.Stretch
                    Grid.SetRow(_div, ri)
                    Grid.SetColumn(_div, _div_col)
                    cab_stack.Children.Add(_div)

                cv.ClipToBounds = True
                if self._is_unificado_mode():
                    elev_host = Grid()
                    elev_host.ClipToBounds = False
                    elev_host.MinHeight = row_h
                    elev_host.VerticalAlignment = VerticalAlignment.Stretch
                    elev_host.HorizontalAlignment = HorizontalAlignment.Stretch

                    mesh_overlay = self._wrap_mesh_settings_card(
                        _mesh_params_block(
                            u"", u"Vertical", u"Horizontal",
                            u"ct_md", u"ct_ms", u"ct_id", u"ct_is",
                            ctr, wall_id=wid,
                        ),
                        title=u"Malla ext. + int.",
                        compact=True,
                    )

                    elev_host.Children.Add(cv)
                    elev_host.Children.Add(mesh_overlay)
                    wall_ui[u"mesh_overlay"] = mesh_overlay
                    self._register_malla_ui_targets(wid, [mesh_overlay])
                    chev_cv = Canvas()
                    # Sin Background el canvas no bloquea clics en zonas vacías;
                    # chevrones y selectores pie (hijos con Fill) siguen interactivos.
                    try:
                        chev_cv.Background = None
                    except Exception:
                        pass
                    chev_cv.ClipToBounds = False
                    chev_cv.HorizontalAlignment = HorizontalAlignment.Stretch
                    chev_cv.VerticalAlignment = VerticalAlignment.Stretch
                    elev_host.Children.Add(chev_cv)
                    wall_ui[u"chevron_canvas"] = chev_cv
                    Grid.SetRow(elev_host, ri)
                    Grid.SetColumn(elev_host, 2)
                    cab_stack.Children.Add(elev_host)
                    self._cabezal_elev_host_by_wid[wid] = elev_host
                    self._controls_by_wall_id[wid] = ctr
                else:
                    Grid.SetRow(cv, ri)
                    Grid.SetColumn(cv, 2)
                    cab_stack.Children.Add(cv)
                    self._controls_by_wall_id[wid] = ctr

                if self._cabezal_is_segment_owner(wid, ex_izq):
                    self._cabezal_pending_caps.append(
                        (ex_izq, wid, wall, 0, False),
                    )
                if self._cabezal_is_segment_owner(wid, ex_der):
                    self._cabezal_pending_caps.append(
                        (ex_der, wid, wall, 4, True),
                    )
            else:
                cv.Width = max(40.0, float(self._preview_col_px))
                prev_wrap.Child = cv
                root.Children.Add(prev_wrap)

            if self._uses_cabezal_panels() and ruler_px > 0:
                _ruler_div = Border()
                _ruler_div.Width = 1.0
                _ruler_div.Background = SolidColorBrush(Color.FromRgb(33, 70, 92))
                _ruler_div.HorizontalAlignment = HorizontalAlignment.Right
                _ruler_div.VerticalAlignment = VerticalAlignment.Stretch
                Grid.SetRow(_ruler_div, grid_row)
                Grid.SetColumn(_ruler_div, 0)
                root.Children.Add(_ruler_div)

            if self._is_mallas_mode() and self._mesh_col_px > 0.0:
                spacer = Border()
                spacer.Background = Brushes.Transparent
                spacer.Width = 12.0
                Grid.SetRow(spacer, grid_row)
                Grid.SetColumn(spacer, 1)

                right = Border()
                right.Background = Brushes.Transparent
                right.Padding = Thickness(4, 0, 6, 0)
                right.Margin = Thickness(0)
                right.VerticalAlignment = VerticalAlignment.Stretch
                right.HorizontalAlignment = HorizontalAlignment.Left
                right.MinHeight = row_h
                right.MaxWidth = float(self._mesh_col_px)
                right.BorderThickness = Thickness(0, 0, 0, 1)
                right.BorderBrush = sep_br
                Grid.SetRow(right, grid_row)
                Grid.SetColumn(right, 2)
                self._right_by_wall_id[wid] = right

                body = StackPanel()
                body.Orientation = Orientation.Vertical
                body.VerticalAlignment = VerticalAlignment.Center

                if self._mesh_modo_tradicional():
                    mesh_card = self._wrap_mesh_settings_card(
                        _mesh_params_block(
                            u"", u"Vertical", u"Horizontal",
                            u"ct_md", u"ct_ms", u"ct_id", u"ct_is",
                            ctr, wall_id=wid,
                        ),
                        title=u"Malla (ext. + int.)",
                        compact=False,
                    )
                    body.Children.Add(mesh_card)
                else:
                    mesh_row = StackPanel()
                    mesh_row.Orientation = Orientation.Horizontal
                    mesh_row.HorizontalAlignment = HorizontalAlignment.Left
                    mesh_row.VerticalAlignment = VerticalAlignment.Center
                    ext_side = self._wrap_mesh_settings_card(
                        _mesh_params_block(
                            u"", u"Vertical ext.", u"Horizontal ext.",
                            u"cex_md", u"cex_ms", u"cex_id", u"cex_is",
                            ctr,
                        ),
                        title=u"Exterior",
                        compact=False,
                    )
                    mesh_row.Children.Add(ext_side)
                    vsep = Border()
                    vsep.Width = 1.0
                    vsep.MinHeight = 72.0
                    vsep.Background = sep_br
                    vsep.Margin = Thickness(8, 2, 8, 2)
                    vsep.VerticalAlignment = VerticalAlignment.Stretch
                    mesh_row.Children.Add(vsep)
                    int_side = self._wrap_mesh_settings_card(
                        _mesh_params_block(
                            u"", u"Vertical int.", u"Horizontal int.",
                            u"cix_md", u"cix_ms", u"cix_id", u"cix_is",
                            ctr,
                        ),
                        title=u"Interior",
                        compact=False,
                    )
                    mesh_row.Children.Add(int_side)
                    body.Children.Add(mesh_row)

                right.Child = body
                self._register_malla_ui_targets(wid, [right])

                root.Children.Add(spacer)
                root.Children.Add(right)
                self._controls_by_wall_id[wid] = ctr

            ctr_malla = self._controls_by_wall_id.get(wid)
            if ctr_malla:
                self._apply_malla_sic_defaults_for_wall(wall, ctr_malla)

            self._draw_wall_polygon_on_canvas(
                cv, self._ensure_wall_preview_model(wid, wall), wall, ri,
            )

        if cab_stack is not None:
            self._mount_cabezal_segment_caps(cab_stack)
            self._mount_cabezal_tramo_connectors(cab_stack)
            Grid.SetRow(cab_stack, self._ruler_top_row)
            Grid.SetRowSpan(cab_stack, n_rows)
            Grid.SetColumn(cab_stack, self._grid_col_offset)
            root.Children.Add(cab_stack)

        if self._is_unificado_mode():
            for wall in self._walls_display_order:
                try:
                    self._sync_cabezal_confinement_from_malla_wall(
                        _wall_id_int(wall), refresh_preview=True,
                    )
                except Exception:
                    pass

        if self._grid_col_offset > 0 and n_rows > 0:
            ruler_cv = Canvas()
            ruler_cv.Background = Brushes.Transparent
            ruler_cv.ClipToBounds = False
            ruler_cv.HorizontalAlignment = HorizontalAlignment.Stretch
            ruler_cv.VerticalAlignment = VerticalAlignment.Stretch
            Grid.SetColumn(ruler_cv, 0)
            Grid.SetRow(ruler_cv, 0)
            Grid.SetRowSpan(ruler_cv, n_rows + self._ruler_top_row)
            top_row_h = 20.0 if self._ruler_top_row > 0 else 0.0
            total_h = sum(float(h) for h in row_heights) + top_row_h
            ruler_cv.Height = total_h
            ruler_cv.Width = ruler_px
            root.Children.Add(ruler_cv)
            self._ruler_canvas = ruler_cv
            self._draw_level_ruler(ruler_cv, row_heights)

        if self._uses_cabezal_panels() and cabezal is not None:
            self._build_cabezal_bulk_actions_panel()
            if self._is_unificado_mode():
                od = getattr(self, u"_walls_display_order", []) or []
                if od:
                    src_wid = _wall_id_int(od[0])
                    src_ctr = self._controls_by_wall_id.get(src_wid)
                    bulk_ctr = getattr(self, u"_bulk_mesh_ctr", None)
                    if src_ctr and bulk_ctr:
                        for k in self._mesh_control_keys():
                            self._copy_combo_like(src_ctr.get(k), bulk_ctr.get(k))

    def _wire_controls(self):
        try:
            from System.Windows import RoutedEventHandler as _REH
        except Exception:
            _REH = None

        if _REH is None:
            return

        btn = self._win.FindName("BtnCrear")
        if btn is not None:
            btn.Click += _REH(lambda s, e: self._on_crear_clicked())

        btn_cancel = self._win.FindName("BtnCancelar")
        if btn_cancel is not None:
            btn_cancel.Click += _REH(lambda s, e: self._on_cancel_clicked())

        btn_cab = self._win.FindName("BtnAplicarCabezal")
        if btn_cab is not None:
            btn_cab.Click += _REH(lambda s, e: self._apply_cabezal_to_all_from_first())

        btn_mesh = self._win.FindName("BtnAplicarMallas")
        if btn_mesh is not None:
            btn_mesh.Click += _REH(lambda s, e: self._apply_mallas_to_all_from_first())

        chk_trad = self._win.FindName("ChkMuroTradicional")
        chk_cont = self._win.FindName("ChkMuroContencion")

        def _on_trad_checked(s, e):
            try:
                if chk_cont is not None:
                    chk_cont.IsChecked = False
            except Exception:
                pass
            self._modo_tradicional = True
            self._rebuild_mesh_ui()

        def _on_cont_checked(s, e):
            try:
                if chk_trad is not None:
                    chk_trad.IsChecked = False
            except Exception:
                pass
            self._modo_tradicional = False
            self._rebuild_mesh_ui()

        if chk_trad is not None:
            chk_trad.Checked += _REH(_on_trad_checked)
        if chk_cont is not None:
            chk_cont.Checked += _REH(_on_cont_checked)

    def _on_cancel_clicked(self):
        try:
            if self._win is not None:
                self._win.Close()
        except Exception:
            pass

    def _refresh_info_txt(self):
        tb = self._win.FindName("TxtInfoMuros")
        if tb is None:
            return
        try:
            from System.Windows import Visibility
        except Exception:
            Visibility = None
        if self._is_cabezal_mode() and not self._is_unificado_mode():
            tb.Text = u""
            if Visibility is not None:
                tb.Visibility = Visibility.Collapsed
            return
        if Visibility is not None:
            tb.Visibility = Visibility.Visible
        n = len(self.walls_list)
        if n == 0:
            tb.Text = u"Sin muros."
            return
        ids = u", ".join(str(_wall_id_int(w)) for w in self.walls_ordered[:8])
        suf = "…" if len(self.walls_ordered) > 8 else ""
        n_fund = 0
        try:
            n_fund = sum(
                1 for w in self.walls_ordered
                if self._foundation_cache.get(_wall_id_int(w))
            )
        except Exception:
            n_fund = 0
        fund_note = u" Fundación unida: {} muro(s).".format(n_fund) if n_fund else u""
        modo_creacion = (
            u"rápido (un lote, sin animación)"
            if getattr(geo, u"MODO_EJECUCION_RAPIDA", True)
            else u"inferior→superior (animación por lote)"
        )
        tb.Text = (
            u"{} muro(s). Elevación: arriba = mayor cota Z; abajo = menor cota. "
            u"Creación {}: {}.{}"
        ).format(n, modo_creacion, ids + suf, fund_note)

    def _apply_cabezal_stepper_enabled(self, step, enabled):
        if step is None:
            return
        try:
            shell = step.get(u"panel")
            if shell is not None:
                shell.IsEnabled = bool(enabled)
                shell.Opacity = 1.0 if enabled else 0.42
        except Exception:
            pass
        if not enabled:
            try:
                if step.get(u"btn_plus") is not None:
                    step[u"btn_plus"].IsEnabled = False
                if step.get(u"btn_minus") is not None:
                    step[u"btn_minus"].IsEnabled = False
            except Exception:
                pass
        else:
            try:
                step[u"set_value"](step[u"get_value"]())
            except Exception:
                pass

    def _apply_cabezal_layer_rows_active_state(self, wid, extremo):
        """CAPAS=N habilita filas 0…N−1; el resto visible pero bloqueado."""
        if cabezal is None:
            return
        from System.Windows import FontWeights

        ui = self._cabezal_ui_ext(wid, extremo)
        n_capas = self._cabezal_n_capas_from_ui(wid, extremo)
        n_capas = max(
            cabezal.CABEZAL_MIN_CAPAS,
            min(cabezal.CABEZAL_MAX_CAPAS, int(n_capas)),
        )
        pal = self._cabezal_ui_palette(extremo, u"unit")
        steppers = ui.get(u"layer_steppers") or []
        diam_cbs = ui.get(u"layer_diam_cbs") or []
        cap_lbls = ui.get(u"layer_cap_lbls") or []
        n_rows = max(len(steppers), len(diam_cbs), cabezal.CABEZAL_MAX_CAPAS)
        for i in range(n_rows):
            active = i < n_capas
            if i < len(steppers):
                self._apply_cabezal_stepper_enabled(steppers[i], active)
                if not active:
                    try:
                        steppers[i][u"panel"].Opacity = pal[u"disabled_opacity"]
                    except Exception:
                        pass
                else:
                    try:
                        steppers[i][u"panel"].Opacity = 1.0
                    except Exception:
                        pass
            if i < len(cap_lbls) and cap_lbls[i] is not None:
                try:
                    cap_lbls[i].Foreground = (
                        pal[u"layer_active"] if active else pal[u"layer_inactive"]
                    )
                    cap_lbls[i].FontSize = 11.0 if active else 10.0
                    cap_lbls[i].FontWeight = (
                        FontWeights.SemiBold if active else FontWeights.Normal
                    )
                except Exception:
                    pass
            if i < len(diam_cbs):
                self._apply_cabezal_diam_combo_state(
                    diam_cbs[i], layer_active=active, palette=pal,
                )
        try:
            cfg = self._cabezal_by_wall_id.get(wid) or {}
            ex_cfg = cfg.setdefault(extremo, cabezal.default_cabezal_extremo_config())
            ex_cfg[u"n_capas"] = n_capas
        except Exception:
            pass

    def _apply_cabezal_diam_combo_state(self, cb, layer_active=True, palette=None):
        """Capas activas: ø editable; fuera de CAPAS bloqueado."""
        if cb is None:
            return
        from System.Windows.Media import SolidColorBrush, Color

        if palette is not None:
            if layer_active:
                try:
                    cb.IsEnabled = True
                    cb.Opacity = 1.0
                    cb.Foreground = palette[u"text_control"]
                except Exception:
                    pass
            else:
                try:
                    cb.IsEnabled = False
                    cb.Opacity = palette[u"disabled_opacity"]
                    cb.Foreground = palette[u"layer_inactive"]
                except Exception:
                    pass
            return

        if layer_active:
            try:
                cb.IsEnabled = True
                cb.Opacity = 1.0
                cb.Foreground = SolidColorBrush(Color.FromRgb(149, 184, 204))
            except Exception:
                pass
        else:
            try:
                cb.IsEnabled = False
                cb.Opacity = 0.42
                cb.Foreground = SolidColorBrush(Color.FromRgb(88, 108, 122))
            except Exception:
                pass

    def _refresh_cabezal_segment_diam_combos(self, extremo):
        """Actualiza combos ø desde la config del tramo (muro propietario)."""
        if cabezal is None:
            return
        for wid, _ri, _w in self._cabezal_walls_for_extremo(extremo):
            ui = self._cabezal_ui_ext(wid, extremo)
            diam_cbs = ui.get(u"layer_diam_cbs") or []
            if not diam_cbs:
                continue
            n_capas = self._cabezal_n_capas_from_ui(wid, extremo)
            pal = self._cabezal_ui_palette(extremo, u"unit")
            for i, cb in enumerate(diam_cbs):
                if cb is None:
                    continue
                bid = self._cabezal_layer_diam_id(wid, extremo, i)
                self._set_diam_combo_selected_id(cb, bid)
                self._apply_cabezal_diam_combo_state(
                    cb, layer_active=(i < n_capas), palette=pal,
                )
            self._request_cabezal_preview_refresh(wid, extremo, debounce=False)

    def _diam_combo_items(self, cb):
        if cb is None:
            return list(getattr(self, "_cabezal_diam_strings", None) or self._diam_strings or [])
        try:
            src = cb.ItemsSource
            if src is not None:
                return list(src)
        except Exception:
            pass
        return list(getattr(self, "_cabezal_diam_strings", None) or self._diam_strings or [])

    def _diam_label_for_combo(self, cb, bar_type_id):
        from Autodesk.Revit.DB import ElementId as _EI

        if bar_type_id is None or bar_type_id == _EI.InvalidElementId:
            return None
        try:
            iv = geo._element_id_int(bar_type_id)
        except Exception:
            return None
        full = getattr(self, "_bar_id_to_label", {}).get(iv)
        if not full:
            return None
        items = self._diam_combo_items(cb)
        compact = _cabezal_compact_diam_label(full)
        if compact in items:
            return compact
        if full in items:
            return full
        return compact

    def _read_diam_combo_id(self, cb):
        from Autodesk.Revit.DB import ElementId as _EI

        if cb is None:
            return _EI.InvalidElementId
        ids_map = getattr(self, "_bar_labels_to_id", {})
        si = getattr(cb, "SelectedItem", None)
        try:
            lab = unicode(si).strip() if si else u""
        except Exception:
            lab = str(si or "").strip()
        bid = ids_map.get(lab, _EI.InvalidElementId)
        if bid == _EI.InvalidElementId:
            bid = ids_map.get(str(lab), _EI.InvalidElementId)
        if bid == _EI.InvalidElementId and lab:
            mm = _parse_diam_label_mm(lab)
            if mm is not None:
                full = _cabezal_diam_label_for_mm(self._diam_strings, mm)
                if full:
                    bid = ids_map.get(full, _EI.InvalidElementId)
        return bid

    def _set_diam_combo_selected_id(self, cb, bar_type_id):
        if cb is None:
            return
        from Autodesk.Revit.DB import ElementId as _EI

        items = self._diam_combo_items(cb)
        if bar_type_id is None or bar_type_id == _EI.InvalidElementId:
            if items:
                try:
                    cb.SelectedIndex = _cabezal_default_diam_index(items)
                except Exception:
                    pass
            return
        lab = self._diam_label_for_combo(cb, bar_type_id)
        if lab and lab in items:
            try:
                cb.SelectedItem = lab
                return
            except Exception:
                pass
        if items:
            try:
                cb.SelectedIndex = _cabezal_default_diam_index(items)
            except Exception:
                pass

    def _create_cabezal_diam_combo(self, bar_type_id, on_change=None):
        from System.Windows.Controls import ComboBox
        from System.Windows import Thickness

        cb = ComboBox()
        cb.IsEditable = False
        cb.Margin = Thickness(0)
        self._apply_flat_combo(cb, narrow=True)
        self._fill_cabezal_diam_combo(cb)
        self._set_diam_combo_selected_id(cb, bar_type_id)
        if on_change is not None:
            try:
                from System.Windows import RoutedEventHandler as _REH

                def _on_sel(sender, args, fn=on_change):
                    fn()

                cb.SelectionChanged += _REH(_on_sel)
            except Exception:
                pass
        return cb

    def _parse_spacing_mm_txt(self, cmb_obj):
        if cmb_obj is None:
            return 150.0
        try:
            t = getattr(cmb_obj, "SelectedItem", None)
            txt = getattr(cmb_obj, "Text", None)
            sel = ""
            try:
                sel = unicode(t).strip()
            except Exception:
                sel = str(t or "").strip()
            if not sel and txt:
                try:
                    sel = unicode(txt).strip()
                except Exception:
                    sel = str(txt).strip()
            return float(sel.replace(",", "."))
        except Exception:
            return 150.0

    def _read_bar_spacing_from_controls(self, ctr, md_key, ms_key):
        from Autodesk.Revit.DB import ElementId as _EI

        ids_map = getattr(self, "_bar_labels_to_id", {})
        dc = ctr.get(md_key)
        ec = ctr.get(ms_key)
        si = getattr(dc, "SelectedItem", None)
        try:
            lab = unicode(si).strip() if si else ""
        except Exception:
            lab = str(si or "").strip()
        bid = ids_map.get(lab, _EI.InvalidElementId)
        if bid == _EI.InvalidElementId:
            bid = ids_map.get(str(lab), _EI.InvalidElementId)
        sei = getattr(ec, "SelectedItem", None)
        if sei is not None:
            try:
                esp_txt = unicode(sei).strip()
            except Exception:
                esp_txt = str(sei).strip()
        else:
            raw_t = getattr(ec, "Text", None) or u"150"
            try:
                esp_txt = unicode(raw_t).strip()
            except Exception:
                esp_txt = str(raw_t).strip()
        return bid, str(esp_txt)

    def _gather_params_for_wall_id(self, wid):
        ctr = self._controls_by_wall_id.get(wid)
        if not ctr:
            return {}, {}

        params_dict = {}
        layer_active_dict = {}

        if self._mesh_modo_tradicional():
            combo_cfg = (
                ("exterior_major", u"ct_md", u"ct_ms"),
                ("exterior_minor", u"ct_id", u"ct_is"),
            )
            for lk, dnk, ensk in combo_cfg:
                bid, esp_txt = self._read_bar_spacing_from_controls(ctr, dnk, ensk)
                params_dict[lk] = (bid, esp_txt)
                layer_active_dict[lk] = True
            params_dict["interior_major"] = params_dict["exterior_major"]
            params_dict["interior_minor"] = params_dict["exterior_minor"]
            layer_active_dict["interior_major"] = True
            layer_active_dict["interior_minor"] = True
            return params_dict, layer_active_dict

        combo_cfg = (
            ("exterior_major", u"cex_md", u"cex_ms"),
            ("exterior_minor", u"cex_id", u"cex_is"),
            ("interior_major", u"cix_md", u"cix_ms"),
            ("interior_minor", u"cix_id", u"cix_is"),
        )
        for lk, dnk, ensk in combo_cfg:
            bid, esp_txt = self._read_bar_spacing_from_controls(ctr, dnk, ensk)
            params_dict[lk] = (bid, esp_txt)
            layer_active_dict[lk] = True

        return params_dict, layer_active_dict

    def _redistribute_row_heights_and_redraw(self):
        from System.Windows import GridLength, GridUnitType

        if self._win is None:
            return
        od = getattr(self, "_walls_display_order", []) or []
        if not od:
            return

        heights = self._compute_row_heights()
        row_h = float(getattr(self, "_row_height_px", 128.0))
        if len(heights) < len(od):
            heights = [row_h] * len(od)

        row_defs = getattr(self, "_row_definitions", []) or []
        for i, hh in enumerate(heights):
            if i >= len(row_defs):
                break
            try:
                row_defs[i].Height = GridLength(float(hh), GridUnitType.Pixel)
            except Exception:
                pass

        self._last_row_heights = list(heights)

        cab_stack = getattr(self, u"_cabezal_stack_grid", None)
        if cab_stack is not None:
            for i, hh in enumerate(heights):
                if i >= cab_stack.RowDefinitions.Count:
                    break
                try:
                    cab_stack.RowDefinitions[i].Height = GridLength(
                        float(hh), GridUnitType.Pixel,
                    )
                except Exception:
                    pass
            for cap, seg, extremo in (
                getattr(self, u"_cabezal_mounted_caps", []) or []
            ):
                try:
                    from System.Windows import Thickness
                    cap_h = self._cabezal_segment_panel_height_px(seg, extremo)
                    top_margin = self._cabezal_segment_align_top_margin_px(seg, cap_h)
                    cap.MinHeight = cap_h
                    cap.Margin = Thickness(0, top_margin, 0, 0)
                    start_ri, span = self._cabezal_segment_span_bounds(seg)
                    from System.Windows.Controls import Grid
                    Grid.SetRow(cap, start_ri)
                    Grid.SetRowSpan(cap, span)
                except Exception:
                    pass
            for conn, seg, _extremo in (
                getattr(self, u"_cabezal_mounted_connectors", []) or []
            ):
                try:
                    start_ri, span = self._cabezal_segment_span_bounds(seg)
                    from System.Windows.Controls import Grid
                    Grid.SetRow(conn, start_ri)
                    Grid.SetRowSpan(conn, span)
                except Exception:
                    pass
            self._layout_cabezal_tramo_connectors_all()

        for i, wall in enumerate(od):
            wid = _wall_id_int(wall)
            hh = float(heights[i])
            pw = self._prev_wrap_by_wall_id.get(wid)
            if pw is not None:
                try:
                    pw.MinHeight = hh
                except Exception:
                    pass
            rg = self._row_grid_by_wall_id.get(wid)
            if rg is not None:
                try:
                    rg.MinHeight = hh
                except Exception:
                    pass
            elev_host = (
                getattr(self, u"_cabezal_elev_host_by_wid", {}) or {}
            ).get(wid)
            if elev_host is not None:
                try:
                    elev_host.MinHeight = hh
                except Exception:
                    pass
            cv = self._canvas_by_wall_id.get(wid)
            if cv is not None:
                cv.Height = hh
                md = self._ensure_wall_preview_model(wid, wall)
                self._draw_wall_polygon_on_canvas(cv, md, wall, i)
            if cabezal is not None:
                for ex in cabezal.CABEZAL_EXTREMOS:
                    owner_wid = self._cabezal_owner_wid_for(wid, ex)
                    ui_ex = self._cabezal_ui_ext(owner_wid, ex)
                    if not isinstance(ui_ex, dict):
                        continue
                    cab_cv = ui_ex.get(u"preview_canvas")
                    if cab_cv is not None:
                        owner_wall = wall
                        if int(owner_wid) != int(wid):
                            for w in od:
                                try:
                                    if _wall_id_int(w) == int(owner_wid):
                                        owner_wall = w
                                        break
                                except Exception:
                                    pass
                        self._draw_cabezal_preview_canvas(
                            cab_cv, owner_wall, owner_wid, ex,
                        )
            right = self._right_by_wall_id.get(wid)
            if right is not None:
                try:
                    right.MinHeight = hh
                except Exception:
                    pass

        self._redraw_level_ruler()

    def _level_theme_brushes(self):
        from System.Windows.Media import SolidColorBrush, Color

        return {
            u"line": SolidColorBrush(Color.FromRgb(71, 85, 105)),
            u"text": SolidColorBrush(Color.FromRgb(148, 163, 184)),
            u"disk": SolidColorBrush(Color.FromRgb(203, 213, 225)),
            u"bubble": SolidColorBrush(Color.FromRgb(34, 211, 238)),
        }

    def _canvas_set_zindex(self, elem, z_index):
        from System.Windows.Controls import Canvas as _Cn

        try:
            _Cn.SetZIndex(elem, int(z_index))
        except Exception:
            pass

    def _cabezal_troceo_por_muro_activo(self, wid, extremo):
        if cabezal is None:
            return False
        cfg = self._cabezal_by_wall_id.get(wid) or {}
        ex_cfg = cfg.get(extremo) or {}
        try:
            return bool(ex_cfg.get(u"troceo_por_muro"))
        except Exception:
            return False

    def _cabezal_extremo_chevron_metrics(self, wall_draw_h_px):
        h = max(20.0, float(wall_draw_h_px))
        half_h = max(5.0, min(9.0, h * 0.08))
        depth = max(6.0, min(11.0, half_h * 1.28))
        return half_h, depth

    def _extremo_to_elevation_lado(self, extremo, ex_izq, ex_der):
        if extremo == ex_izq:
            return u"izq"
        if extremo == ex_der:
            return u"der"
        return None

    def _elevation_vecino_short_label(self, neighbor):
        if neighbor is None:
            return u"Muro"
        try:
            eid = geo._element_id_int(neighbor.Id)
        except Exception:
            return u"Muro"
        try:
            nm = neighbor.Name
            if nm:
                return u"{0} ({1})".format(nm, eid)
        except Exception:
            pass
        return u"M{0}".format(eid)

    def _elevation_encuentro_vecino_at_extremo(self, wall, extremo):
        """Vecino en extremo inicio/fin; retorna dict lado/tipo/label o None."""
        if _vec_ext is None or wall is None or extremo not in (u"inicio", u"fin"):
            return None
        doc = self.doc
        vecinos = _vec_ext.vecinos_en_extremo(doc, wall, extremo)
        if not vecinos:
            return None

        neighbor = None
        tipo = u"T"
        if _cab_enc_l is not None:
            v_l = _vec_ext.vecino_principal_encuentro_l(doc, wall, extremo)
            if v_l is not None:
                neighbor = v_l
                tipo = u"L"
            else:
                for w in vecinos:
                    kind = _cab_enc_l.clasificar_encuentro_en_extremo(
                        doc, wall, w, extremo,
                    )
                    if kind == _cab_enc_l.CABEZAL_ENC_TIPO_L:
                        neighbor = w
                        tipo = u"L"
                        break
                    if kind == u"otro":
                        continue
                    if neighbor is None:
                        neighbor = w
                        tipo = u"T"
        else:
            neighbor = vecinos[0]

        if neighbor is None:
            return None
        return {
            u"tipo": tipo,
            u"vecino": neighbor,
            u"label": self._elevation_vecino_short_label(neighbor),
        }

    def _elevation_vecino_lado_cara(self, wall, neighbor, ex_izq, ex_der):
        """Lado izq./der. del canvas según posición del vecino respecto al eje."""
        if _vec_ext is None or wall is None or neighbor is None:
            return None
        wns = _vec_ext._load_wall_node_section()
        if wns is None:
            return None
        try:
            wall_line, _co = wns._location_as_line(wall)
        except Exception:
            return None
        if wall_line is None:
            return None
        try:
            e0 = wall_line.GetEndPoint(0)
            e1 = wall_line.GetEndPoint(1)
            ol = neighbor.Location
            from Autodesk.Revit.DB import LocationCurve
            if not isinstance(ol, LocationCurve):
                return None
            oc = ol.Curve
            if oc is None:
                return None
            nm = wns._midpoint_curve(oc)
            if nm is None:
                return None
            sx = float(e1.X - e0.X)
            sy = float(e1.Y - e0.Y)
            wx = float(nm.X - e0.X)
            wy = float(nm.Y - e0.Y)
            cross = sx * wy - sy * wx
            extremo = u"inicio" if cross >= 0.0 else u"fin"
        except Exception:
            return None
        return self._extremo_to_elevation_lado(extremo, ex_izq, ex_der)

    def _collect_elevation_vecino_encuentros(self, wall, row_index):
        """
        Encuentros L/T detectados para el canvas de elevación de una fila.

        Cada item: ``lado`` (izq|der), ``tipo`` (L|T), ``label`` (texto tooltip).
        """
        if wall is None or _vec_ext is None:
            return []
        try:
            wid = _wall_id_int(wall)
            ex_izq, ex_der = self._cabezal_extremos_lados_wall(wid, row_index)
        except Exception:
            ex_izq, ex_der = u"inicio", u"fin"

        out = []
        seen_lados = set()
        doc = self.doc

        for extremo in (ex_izq, ex_der):
            enc = self._elevation_encuentro_vecino_at_extremo(wall, extremo)
            if enc is None:
                continue
            lado = self._extremo_to_elevation_lado(extremo, ex_izq, ex_der)
            if lado is None or lado in seen_lados:
                continue
            seen_lados.add(lado)
            enc[u"lado"] = lado
            out.append(enc)

        try:
            laterales = _vec_ext.vecinos_cara_lateral_o_t(doc, wall)
        except Exception:
            laterales = []
        for nb in laterales or []:
            lado = self._elevation_vecino_lado_cara(wall, nb, ex_izq, ex_der)
            if lado is None or lado in seen_lados:
                continue
            seen_lados.add(lado)
            out.append({
                u"lado": lado,
                u"tipo": u"T",
                u"vecino": nb,
                u"label": self._elevation_vecino_short_label(nb),
            })

        return out

    def _draw_elevation_vecino_encuentros(
        self, canv, wall, row_index, x_off, draw_w, y_top, wall_draw_h_px,
    ):
        """
        Segmento vertical discontinuo que representa el muro vecino detectado (L/T).
        Offset fijo en px — solo representación gráfica.
        """
        if canv is None or wall is None:
            return
        encuentros = self._collect_elevation_vecino_encuentros(wall, row_index)
        if not encuentros:
            return

        from System.Windows.Shapes import Line as _Ln
        from System.Windows.Media import SolidColorBrush, Color, DoubleCollection

        y0 = float(y_top) + 1.0
        y1 = float(y_top) + max(4.0, float(wall_draw_h_px)) - 1.0
        x_left = float(x_off)
        x_right = float(x_off) + max(14.0, float(draw_w))
        off = float(getattr(self, u"_ENC_VEC_SEG_OFFSET_PX", 14.0))
        br_edge = SolidColorBrush(Color.FromRgb(15, 23, 42))
        br_dash = SolidColorBrush(Color.FromRgb(71, 85, 105))
        dash = DoubleCollection()
        dash.Add(5.0)
        dash.Add(4.0)

        for enc in encuentros:
            lado = enc.get(u"lado")
            if lado == u"izq":
                x_edge = x_left
                x_seg = x_edge + off
            elif lado == u"der":
                x_edge = x_right
                x_seg = x_edge - off
            else:
                continue

            edge_ln = _Ln()
            edge_ln.X1 = x_edge
            edge_ln.X2 = x_edge
            edge_ln.Y1 = y0
            edge_ln.Y2 = y1
            edge_ln.Stroke = br_edge
            edge_ln.StrokeThickness = 1.0
            try:
                edge_ln.SnapsToDevicePixels = True
            except Exception:
                pass
            self._canvas_set_zindex(edge_ln, int(self._ENC_VEC_SEG_Z))
            canv.Children.Add(edge_ln)

            seg = _Ln()
            seg.X1 = x_seg
            seg.X2 = x_seg
            seg.Y1 = y0
            seg.Y2 = y1
            seg.Stroke = br_dash
            seg.StrokeThickness = float(
                getattr(self, u"_ENC_VEC_SEG_STROKE_PX", 1.5),
            )
            seg.StrokeDashArray = dash
            try:
                seg.SnapsToDevicePixels = True
            except Exception:
                pass
            tipo = enc.get(u"tipo") or u"?"
            lbl = enc.get(u"label") or u"Muro"
            try:
                seg.ToolTip = u"Encuentro {0} · {1}".format(tipo, lbl)
            except Exception:
                pass
            self._canvas_set_zindex(seg, int(self._ENC_VEC_SEG_Z) + 1)
            canv.Children.Add(seg)

    def _draw_cabezal_extremo_chevron_elevation(
        self,
        canv,
        x_edge,
        y_mid,
        face_right,
        color_rgb,
        half_h,
        depth,
        z_index=None,
        tooltip=None,
    ):
        u"""Triángulo en borde del muro (opción B): apunta hacia el interior."""
        from System.Windows.Shapes import Polygon as _Wp
        from System.Windows import Point
        from System.Windows.Media import SolidColorBrush, Color, PointCollection

        if z_index is None:
            z_index = int(getattr(self, u"_EXTREMO_CHEVRON_Z", 30))

        try:
            r, g, b = color_rgb[:3]
        except Exception:
            r, g, b = 34, 211, 238

        x = float(x_edge)
        ym = float(y_mid)
        hh = float(half_h)
        dep = float(depth)

        plc = PointCollection()
        if face_right:
            plc.Add(Point(x, ym - hh))
            plc.Add(Point(x + dep, ym))
            plc.Add(Point(x, ym + hh))
        else:
            plc.Add(Point(x, ym - hh))
            plc.Add(Point(x - dep, ym))
            plc.Add(Point(x, ym + hh))

        poly = _Wp()
        poly.Points = plc
        poly.Fill = SolidColorBrush(Color.FromRgb(int(r), int(g), int(b)))
        poly.Stroke = SolidColorBrush(Color.FromRgb(20, 20, 20))
        poly.StrokeThickness = 0.6
        try:
            poly.SnapsToDevicePixels = True
        except Exception:
            pass
        if tooltip:
            try:
                poly.ToolTip = tooltip
            except Exception:
                pass
        self._canvas_set_zindex(poly, int(z_index))
        canv.Children.Add(poly)

    def _draw_cabezal_extremo_marks_elevation(
        self, canv, wall, row_index, x_off, draw_w, y_top, wall_draw_h_px,
    ):
        u"""Chevrones Inicio/Final en elevación, alineados con los bordes de la banda."""
        if not self._uses_cabezal_panels() or cabezal is None or wall is None:
            return
        od = getattr(self, u"_walls_display_order", []) or []
        if row_index < 0 or row_index >= len(od):
            return

        draw_w_f = max(14.0, float(draw_w))
        if draw_w_f < 18.0:
            return

        wid = _wall_id_int(wall)
        ex_izq, ex_der = self._cabezal_extremos_lados_wall(wid, row_index)
        y_mid = float(y_top) + max(10.0, float(wall_draw_h_px)) * 0.5
        half_h, depth = self._cabezal_extremo_chevron_metrics(wall_draw_h_px)
        x_left = float(x_off)
        x_right = float(x_off) + draw_w_f

        pal_izq = self._cabezal_ui_palette(ex_izq, u"unit")
        pal_der = self._cabezal_ui_palette(ex_der, u"unit")
        lbl_izq = (
            u"Inicio"
            if ex_izq == cabezal.CABEZAL_EXTREMO_INICIO
            else u"Final"
        )
        lbl_der = (
            u"Inicio"
            if ex_der == cabezal.CABEZAL_EXTREMO_INICIO
            else u"Final"
        )

        self._draw_cabezal_extremo_chevron_elevation(
            canv,
            x_left,
            y_mid,
            True,
            pal_izq.get(u"bar_accent_rgb"),
            half_h,
            depth,
            tooltip=lbl_izq,
        )
        self._draw_cabezal_extremo_chevron_elevation(
            canv,
            x_right,
            y_mid,
            False,
            pal_der.get(u"bar_accent_rgb"),
            half_h,
            depth,
            tooltip=lbl_der,
        )

    def _draw_empalme_tick_on_elevation(self, canv, y, x_edge, on_left):
        u"""Segmento horizontal rojo en borde izq./der. de la banda del muro."""
        from System.Windows.Shapes import Line as _Ln
        from System.Windows.Media import SolidColorBrush, Color

        tick_w = float(self._EMPALME_TICK_W_PX)
        yv = float(y)
        xe = float(x_edge)
        if on_left:
            x1 = xe
            x2 = xe + tick_w
        else:
            x2 = xe
            x1 = xe - tick_w
        ln = _Ln()
        ln.X1 = x1
        ln.X2 = x2
        ln.Y1 = yv
        ln.Y2 = yv
        ln.Stroke = SolidColorBrush(Color.FromRgb(220, 50, 60))
        ln.StrokeThickness = float(self._EMPALME_TICK_STROKE_PX)
        try:
            ln.SnapsToDevicePixels = True
        except Exception:
            pass
        self._canvas_set_zindex(ln, 35)
        canv.Children.Add(ln)

    def _empalme_tick_y_wall_base_px(self, wall_draw_h_px):
        """Cota Y del pie del muro en el canvas de elevación (borde inferior de la banda)."""
        return max(2.0, float(wall_draw_h_px) - 1.0)

    def _draw_cabezal_empalme_marks_elevation(
        self, canv, wall, row_index, x_off, draw_w, wall_draw_h_px,
    ):
        """
        Marcas de empalme en elevación (cabezal), en el pie del muro de la fila.
        Solo si «Define Empalmes» (troceo_por_muro) está activo en ese extremo.
        """
        if not self._uses_cabezal_panels() or cabezal is None or wall is None:
            return
        od = getattr(self, "_walls_display_order", []) or []
        if row_index < 0 or row_index >= len(od):
            return

        y_joint = self._empalme_tick_y_wall_base_px(wall_draw_h_px)
        x_left = float(x_off)
        x_right = float(x_off) + max(14.0, float(draw_w))

        wid = _wall_id_int(wall)
        ex_izq, ex_der = self._cabezal_extremos_lados_wall(wid, row_index)
        if (
            self._cabezal_armado_activo_cfg(wid, ex_izq)
            and self._cabezal_troceo_por_muro_activo(wid, ex_izq)
        ):
            self._draw_empalme_tick_on_elevation(canv, y_joint, x_left, True)
        if (
            self._cabezal_armado_activo_cfg(wid, ex_der)
            and self._cabezal_troceo_por_muro_activo(wid, ex_der)
        ):
            self._draw_empalme_tick_on_elevation(canv, y_joint, x_right, False)

    def _redraw_wall_elevation_canvas(self, wid):
        u"""Redibuja solo el canvas de elevación de un muro (marcas de empalme incluidas)."""
        od = getattr(self, "_walls_display_order", []) or []
        for ri, w in enumerate(od):
            try:
                if _wall_id_int(w) != int(wid):
                    continue
            except Exception:
                continue
            cv = self._canvas_by_wall_id.get(int(wid))
            if cv is None:
                return
            self._draw_wall_polygon_on_canvas(
                cv, self._ensure_wall_preview_model(int(wid), w), w, ri,
            )
            return

    def _add_canvas_text(self, canv, text, left, top, fg, font_size, font_weight, z_index):
        from System.Windows.Controls import TextBlock, Canvas as _Cn
        from System.Windows import FontWeights

        tb = TextBlock()
        tb.Text = text or u""
        tb.Foreground = fg
        tb.FontSize = float(font_size)
        try:
            tb.FontWeight = font_weight
        except Exception:
            tb.FontWeight = FontWeights.Normal
        _Cn.SetLeft(tb, float(left))
        _Cn.SetTop(tb, float(top))
        self._canvas_set_zindex(tb, z_index)
        canv.Children.Add(tb)
        return tb

    def _add_canvas_text_centered(self, canv, text, cx, cy, fg, font_size, font_weight, z_index):
        from System import Double
        from System.Windows import Size, FontWeights
        from System.Windows.Controls import TextBlock

        tb = TextBlock()
        tb.Text = text or u""
        tb.Foreground = fg
        tb.FontSize = float(font_size)
        try:
            tb.FontWeight = font_weight
        except Exception:
            tb.FontWeight = FontWeights.SemiBold
        try:
            tb.Measure(Size(Double.PositiveInfinity, Double.PositiveInfinity))
            lw = float(tb.DesiredSize.Width)
            lh = float(tb.DesiredSize.Height)
        except Exception:
            lw, lh = 48.0, 14.0
        self._add_canvas_text(
            canv, text, cx - lw / 2.0, cy - lh / 2.0, fg, font_size, font_weight, z_index,
        )
        return tb

    def _wall_section_label_lines(self, wall):
        u"""Líneas centrales del canvas: tipo (ej. M.H.A.) y espesor (ej. e=300)."""
        tipo = u"Muro"
        esp_mm = None

        try:
            esp_mm = geo.obtener_espesor_muro_mm_approx(wall)
        except Exception:
            esp_mm = None

        def _tipo_from_name(raw):
            if not raw:
                return None
            try:
                s = unicode(raw).strip()
            except Exception:
                s = str(raw or u"").strip()
            if not s:
                return None
            low = s.lower()
            idx = low.find(u" e=")
            if idx >= 0:
                return s[:idx].strip()
            return s

        try:
            wt = wall.WallType
            if wt is not None:
                tpart = _tipo_from_name(getattr(wt, "Name", None))
                if tpart:
                    tipo = tpart
        except Exception:
            pass

        if tipo == u"Muro":
            try:
                tpart = _tipo_from_name(getattr(wall, "Name", None))
                if tpart:
                    tipo = tpart
            except Exception:
                pass

        if esp_mm is not None:
            esp_line = u"e={:.0f}".format(float(esp_mm))
        else:
            esp_line = u"e=—"

        return tipo, esp_line

    def _wall_length_feet_approx(self, wall):
        try:
            loc = wall.Location
            crv = getattr(loc, "Curve", None)
            if crv is not None:
                return max(float(crv.Length), 1e-6)
        except Exception:
            pass
        try:
            bb = wall.get_BoundingBox(None)
            if bb is not None:
                dx = abs(float(bb.Max.X) - float(bb.Min.X))
                dy = abs(float(bb.Max.Y) - float(bb.Min.Y))
                dz = abs(float(bb.Max.Z) - float(bb.Min.Z))
                return max(dx, dy, dz, 1e-6)
        except Exception:
            pass
        return 1.0

    def _ensure_layout_cache(self):
        od = getattr(self, "_walls_display_order", []) or []
        try:
            key = tuple(_wall_id_int(w) for w in od)
        except Exception:
            key = None
        if key and key == getattr(self, "_length_scale_key", None):
            return

        layout = geo.compute_preview_horizontal_layout(od, self._model_cache)
        self._length_scale_key = key
        self._preview_layout = layout

        if layout is None:
            self._wall_extent_u_list = []
            self._max_extent_u_feet = 1.0
            return

        self._wall_extent_u_list = [
            float(it.get("extent_u", 1.0)) for it in (layout.get("items") or [])
        ]
        self._max_extent_u_feet = float(layout.get("max_extent_u", 1.0) or 1.0)

    def _wall_draw_width_px(self, row_index, canvas_width_px):
        self._ensure_layout_cache()
        wpx = max(
            10.0,
            self._finite_px(canvas_width_px, float(self._PREVIEW_ELEV_COL_PX)),
        )
        gutter = float(getattr(self, "_preview_level_gutter_px", self._PREVIEW_LEVEL_GUTTER_PX))
        left_pad = 6.0

        stacked = getattr(self, "_stacked_layout", None)
        if stacked is not None:
            st_items = stacked.get(u"items") or []
            if 0 <= row_index < len(st_items):
                item = st_items[row_index]
                span = float(stacked.get(u"global_span", 1.0))
                g_min = float(stacked.get(u"global_min", 0.0))
                right_pad = 6.0
                usable = max(20.0, wpx - left_pad - right_pad)
                length_u = float(item.get(u"length_u", span))
                u_start = float(item.get(u"u_start", g_min))

                draw_w = max(14.0, usable * (length_u / span))
                x_off = left_pad + ((u_start - g_min) / span) * usable

                x_off = max(left_pad, x_off)
                if x_off + draw_w > wpx - right_pad:
                    x_off = max(left_pad, wpx - right_pad - draw_w)
                return x_off, draw_w

        layout = getattr(self, "_preview_layout", None)
        if layout is None:
            draw_w = max(14.0, wpx - gutter - left_pad)
            return left_pad, draw_w

        items = layout.get("items") or []
        if not (0 <= row_index < len(items)):
            draw_w = max(14.0, wpx - gutter - left_pad)
            return left_pad, draw_w

        item = items[row_index]
        eu = float(item.get("extent_u", 1.0))
        u_pos = float(item.get("u_pos", 0.0))

        usable = max(20.0, wpx - gutter - left_pad)
        margin = left_pad
        wall_zone_max = max(left_pad + 14.0, wpx - gutter - float(
            getattr(self, "_PREVIEW_WALL_LEVEL_GAP_PX", 12.0),
        ))

        u_span_min = float(layout.get("u_span_min", 0.0))
        u_span_max = float(layout.get("u_span_max", 1.0))
        span = max(u_span_max - u_span_min, 1e-6)
        half = eu * 0.5
        draw_w = max(14.0, min(usable, usable * (eu / span)))
        u_left = u_pos - half
        x_off = margin + ((u_left - u_span_min) / span) * usable
        if x_off < left_pad:
            x_off = left_pad
        if x_off + draw_w > wall_zone_max:
            x_off = max(left_pad, wall_zone_max - draw_w)
        return x_off, draw_w

    def _wall_elevation_draw_metrics(self, wall, row_index, canv):
        """Geometría del prisma en el canvas de elevación (px)."""
        wpx = self._finite_px(
            canv.Width,
            max(40.0, float(self._PREVIEW_ELEV_COL_PX)),
        )
        if wpx < 25.0:
            wpx = max(40.0, float(self._PREVIEW_ELEV_COL_PX))
        hpx_row = self._finite_px(
            canv.Height,
            float(getattr(self, "_row_height_px", 128.0)),
        )
        if hpx_row < 10.0:
            hpx_row = float(getattr(self, "_row_height_px", 128.0))
        box_w = max(10.0, wpx)
        box_h = max(10.0, hpx_row)
        x_off, draw_w = self._wall_draw_width_px(row_index, box_w)
        fund_info = None
        if wall is not None:
            try:
                fund_info = self._foundation_cache.get(_wall_id_int(wall))
            except Exception:
                fund_info = None
        foot_h = self._foundation_footing_height_px(box_h, fund_info)
        wall_draw_h = (
            max(10.0, float(box_h) - float(foot_h))
            if foot_h > 0.0
            else float(box_h)
        )
        return x_off, draw_w, wall_draw_h, box_h

    def _position_mesh_overlay_for_wall(self, wall, row_index, canv):
        """Centra la tarjeta malla sobre el prisma dibujado (modo unificado)."""
        if not self._is_unificado_mode() or wall is None or canv is None:
            return
        try:
            wid = _wall_id_int(wall)
        except Exception:
            return
        wall_ui = self._cabezal_ui_by_wall_id.get(wid) or {}
        mesh = wall_ui.get(u"mesh_overlay")
        if mesh is None:
            return
        x_off, draw_w, wall_draw_h, _box_h = self._wall_elevation_draw_metrics(
            wall, row_index, canv,
        )
        mleft = float(getattr(self, u"_CABEZAL_ELEV_CANVAS_MARGIN_PX", 4.0))
        max_w = float(getattr(self, u"_MESH_OVERLAY_COMPACT_MAX_W_PX", 212.0))
        try:
            fit_w = max(72.0, min(max_w, float(draw_w) - 8.0))
            mesh.MaxWidth = fit_w
        except Exception:
            fit_w = max_w
        mw = fit_w
        mh = 72.0
        try:
            from System import Double
            from System.Windows import Size, HorizontalAlignment, VerticalAlignment, Thickness
            mesh.UpdateLayout()
            mw = float(mesh.ActualWidth)
            mh = float(mesh.ActualHeight)
            if mw <= 1.0 or mh <= 1.0:
                mesh.Measure(Size(Double.PositiveInfinity, Double.PositiveInfinity))
                mw = max(1.0, float(mesh.DesiredSize.Width))
                mh = max(1.0, float(mesh.DesiredSize.Height))
        except Exception:
            try:
                from System.Windows import HorizontalAlignment, VerticalAlignment, Thickness
            except Exception:
                return
        left = mleft + float(x_off) + max(0.0, (float(draw_w) - mw) * 0.5)
        top = max(2.0, (float(wall_draw_h) - mh) * 0.5)
        try:
            mesh.HorizontalAlignment = HorizontalAlignment.Left
            mesh.VerticalAlignment = VerticalAlignment.Top
            mesh.Margin = Thickness(left, top, 0, 0)
        except Exception:
            pass

    def _level_gutter_left(self, wpx):
        gutter = float(getattr(self, "_preview_level_gutter_px", self._PREVIEW_LEVEL_GUTTER_PX))
        return max(0.0, float(wpx) - gutter)

    def _preview_level_layout(self, wpx, x_off, draw_w):
        gutter = float(getattr(self, "_preview_level_gutter_px", self._PREVIEW_LEVEL_GUTTER_PX))
        gap = float(getattr(self, "_PREVIEW_WALL_LEVEL_GAP_PX", 12.0))
        gutter_left = max(0.0, float(wpx) - gutter)
        wall_right = min(float(x_off) + float(draw_w), gutter_left - gap)
        bubble_r = max(5.0, min(8.0, gutter * 0.11))
        bubble_cx = gutter_left + gutter * 0.64
        bubble_cx = max(bubble_cx, gutter_left + bubble_r + 16.0)
        bubble_cx = min(bubble_cx, float(wpx) - bubble_r - 8.0)
        return {
            u"gutter_left": gutter_left,
            u"wall_right": wall_right,
            u"bubble_r": bubble_r,
            u"bubble_cx": bubble_cx,
            u"gap": gap,
            u"card_width": float(wpx),
        }

    def _add_level_zone_delimiter(self, canv, gutter_left, y_top, y_bottom, z_index=4):
        u"""Segmento vertical que separa la zona de muros de la zona de cotas/nivel."""
        from System.Windows.Shapes import Line
        from System.Windows.Media import SolidColorBrush, Color

        ln = Line()
        x = float(gutter_left)
        ln.X1 = x
        ln.X2 = x
        ln.Y1 = float(y_top)
        ln.Y2 = max(float(y_bottom), float(y_top) + 1.0)
        ln.Stroke = SolidColorBrush(Color.FromRgb(71, 85, 105))
        ln.StrokeThickness = 1.0
        ln.Opacity = 0.85
        self._canvas_set_zindex(ln, int(z_index))
        canv.Children.Add(ln)

    def _add_canvas_label_block_centered(self, canv, lines, cx, cy, fg, font_size, font_weight, z_index):
        from System import Double
        from System.Windows import Size, FontWeights
        from System.Windows.Controls import TextBlock

        if not lines:
            return

        gap = 1.0
        sizes = []
        total_h = 0.0
        max_w = 0.0
        for txt in lines:
            tb = TextBlock()
            tb.Text = txt or u""
            tb.Foreground = fg
            tb.FontSize = float(font_size)
            try:
                tb.FontWeight = font_weight
            except Exception:
                tb.FontWeight = FontWeights.SemiBold
            try:
                tb.Measure(Size(Double.PositiveInfinity, Double.PositiveInfinity))
                lw = float(tb.DesiredSize.Width)
                lh = float(tb.DesiredSize.Height)
            except Exception:
                lw, lh = 40.0, 12.0
            sizes.append((lw, lh))
            max_w = max(max_w, lw)
            total_h += lh

        if len(lines) > 1:
            total_h += gap * float(len(lines) - 1)

        y = float(cy) - total_h / 2.0
        for i, txt in enumerate(lines):
            lw, lh = sizes[i]
            left = float(cx) - lw / 2.0
            self._add_canvas_text(
                canv, txt, left, y, fg, font_size, font_weight, z_index + i,
            )
            y += lh + (gap if i < len(lines) - 1 else 0.0)

    def _add_level_head_symbol(self, canv, line_y, line_x1, line_x2, bubble_cx, bubble_r, z_index=40):
        from System.Windows.Controls import Canvas as _Cn
        from System.Windows.Shapes import Line, Ellipse, Path
        from System.Windows.Media import Brushes, DoubleCollection, Geometry

        br = self._level_theme_brushes()
        zi = int(z_index)
        cx = float(bubble_cx)
        cy = float(line_y)
        r = float(bubble_r)
        x_end = max(float(line_x2), cx + r + 0.5)

        ln = Line()
        ln.Stroke = br[u"line"]
        ln.StrokeThickness = 0.9
        try:
            dashes = DoubleCollection()
            dashes.Add(4.0)
            dashes.Add(3.0)
            ln.StrokeDashArray = dashes
        except Exception:
            pass
        ln.X1 = float(line_x1)
        ln.Y1 = cy
        ln.X2 = x_end
        ln.Y2 = cy
        self._canvas_set_zindex(ln, zi)
        canv.Children.Add(ln)

        bx = cx - r
        by = cy - r
        dia = 2.0 * r

        disk = Ellipse()
        disk.Width = dia
        disk.Height = dia
        disk.Fill = br[u"disk"]
        disk.Stroke = Brushes.Transparent
        _Cn.SetLeft(disk, bx)
        _Cn.SetTop(disk, by)
        self._canvas_set_zindex(disk, zi + 1)
        canv.Children.Add(disk)

        try:
            p_tl = Path()
            p_tl.Data = Geometry.Parse(
                u"M {0},{1} L {2},{1} A {3},{3} 0 0 1 {0},{4} Z".format(cx, cy, cx - r, r, cy - r)
            )
            p_tl.Fill = br[u"bubble"]
            self._canvas_set_zindex(p_tl, zi + 2)
            canv.Children.Add(p_tl)

            p_br = Path()
            p_br.Data = Geometry.Parse(
                u"M {0},{1} L {2},{1} A {3},{3} 0 0 1 {0},{4} Z".format(cx, cy, cx + r, r, cy + r)
            )
            p_br.Fill = br[u"bubble"]
            self._canvas_set_zindex(p_br, zi + 3)
            canv.Children.Add(p_br)

            rim = Ellipse()
            rim.Width = dia
            rim.Height = dia
            rim.Fill = Brushes.Transparent
            rim.Stroke = br[u"text"]
            rim.StrokeThickness = 0.95
            _Cn.SetLeft(rim, bx)
            _Cn.SetTop(rim, by)
            self._canvas_set_zindex(rim, zi + 4)
            canv.Children.Add(rim)
        except Exception:
            pass

        return bubble_cx, bubble_r

    def _add_level_head_at_wall_base(self, canv, line_y, layout, z_index=40):
        bubble_cx = float(layout[u"bubble_cx"])
        bubble_r = float(layout[u"bubble_r"])

        # Discontinua: borde izquierdo de la card → borde derecho del símbolo de nivel.
        line_x1 = 4.0
        line_x2 = bubble_cx + bubble_r + 1.0
        self._add_level_head_symbol(
            canv, line_y, line_x1, line_x2, bubble_cx, bubble_r, z_index,
        )
        return bubble_cx, bubble_r

    def _add_level_label_at_base(self, canv, line_y, level_text, layout, z_index=50):
        from System import Double
        from System.Windows import Size, FontWeights
        from System.Windows.Controls import TextBlock

        br = self._level_theme_brushes()
        bubble_cx = float(layout[u"bubble_cx"])
        bubble_r = float(layout[u"bubble_r"])
        wall_right = float(layout[u"wall_right"])
        gutter_left = float(layout[u"gutter_left"])
        gap = float(layout.get(u"gap", 12.0))

        tb = TextBlock()
        tb.Text = level_text or u""
        tb.Foreground = br[u"text"]
        tb.FontSize = 9.0
        tb.FontWeight = FontWeights.SemiBold
        try:
            tb.Measure(Size(Double.PositiveInfinity, Double.PositiveInfinity))
            lw = float(tb.DesiredSize.Width)
            lh = float(tb.DesiredSize.Height)
        except Exception:
            lw, lh = 36.0, 12.0

        bubble_left = bubble_cx - bubble_r
        zone_left = gutter_left + 4.0
        zone_right = bubble_left - 6.0
        min_left = wall_right + gap

        if zone_right - zone_left >= lw and zone_right > min_left:
            lab_left = zone_left + max(0.0, (zone_right - zone_left - lw) * 0.5)
        else:
            lab_left = max(min_left, bubble_left - lw - 8.0)

        lab_left = max(lab_left, min_left)
        lab_top = max(4.0, float(line_y) - lh - 4.0)
        if lab_top < 4.0:
            lab_top = max(4.0, float(line_y) + 4.0)

        self._add_canvas_text(
            canv, level_text, lab_left, lab_top, br[u"text"], 9.0, FontWeights.SemiBold, z_index,
        )

    def _draw_level_ruler(self, canv, row_heights):
        u"""Dibuja la regleta de niveles con símbolo Revit, alineada con los bordes de cada fila."""
        from System.Windows.Controls import Canvas as _Cn, TextBlock
        from System.Windows.Shapes import Line, Ellipse, Path
        from System.Windows.Media import (
            SolidColorBrush, Color, Brushes, Geometry, DoubleCollection,
        )
        from System.Windows import FontWeights

        canv.Children.Clear()
        od = getattr(self, "_walls_display_order", []) or []
        if not od:
            return

        ruler_w = self._ruler_col_px()
        _top_pad = 20.0 if getattr(self, "_ruler_top_row", 0) > 0 else 0.0
        total_h = sum(float(h) for h in row_heights) + _top_pad

        br_text = SolidColorBrush(Color.FromRgb(34, 211, 238))
        br_guide = SolidColorBrush(Color.FromRgb(47, 60, 78))
        br_disk = SolidColorBrush(Color.FromRgb(207, 222, 232))
        br_bubble = SolidColorBrush(Color.FromRgb(31, 122, 173))
        br_rim = SolidColorBrush(Color.FromRgb(168, 197, 217))
        br_lead = SolidColorBrush(Color.FromRgb(74, 107, 130))

        bubble_r = 9.0
        bubble_cx = ruler_w * 0.72

        boundaries = []
        y_acc = 0.0
        for ri, wall in enumerate(od):
            top_y = y_acc + _top_pad
            hh = float(row_heights[ri]) if ri < len(row_heights) else 0.0

            foot_h = 0.0
            try:
                fund_info = self._foundation_cache.get(_wall_id_int(wall))
                if fund_info:
                    foot_h = self._foundation_footing_height_px(hh, fund_info)
            except Exception:
                pass

            bot_y = y_acc + hh - foot_h + _top_pad

            try:
                z_top = float(geo.cota_superior_muro_metros_aprox(wall))
            except Exception:
                z_top = 0.0
            try:
                z_bot = float(geo.cota_inferior_muro_metros_aprox(wall))
            except Exception:
                z_bot = 0.0

            boundaries.append({
                u"top_y": top_y,
                u"bot_y": bot_y,
                u"z_top": z_top,
                u"z_bot": z_bot,
                u"fund_h": foot_h,
                u"color_hex": self._wall_elevation_color_hex(wall, ri),
            })
            y_acc += hh

        drawn_levels = set()

        def _draw_level_head(y, level_m, is_fund=False, accent_hex=None):
            level_key = round(level_m, 3)
            if level_key in drawn_levels:
                return
            drawn_levels.add(level_key)

            zi = 40
            cx = bubble_cx
            cy = y
            r = bubble_r
            dia = 2.0 * r
            bx = cx - r
            by = cy - r

            if is_fund and accent_hex:
                br_head_disk = self._ui_lighten_hex_brush(accent_hex, mix=0.28)
                br_head_bubble = self._ui_brush_hex(accent_hex, alpha=255)
                br_head_text = self._ui_lighten_hex_brush(accent_hex, mix=0.45)
            else:
                br_head_disk = br_disk
                br_head_bubble = br_bubble
                br_head_text = br_text

            disk = Ellipse()
            disk.Width = dia
            disk.Height = dia
            disk.Fill = br_head_disk
            disk.Stroke = Brushes.Transparent
            _Cn.SetLeft(disk, bx)
            _Cn.SetTop(disk, by)
            self._canvas_set_zindex(disk, zi)
            canv.Children.Add(disk)

            try:
                q_fill = br_head_bubble
                p_tl = Path()
                p_tl.Data = Geometry.Parse(
                    u"M {0},{1} L {2},{1} A {3},{3} 0 0 1 {0},{4} Z".format(
                        cx, cy, cx - r, r, cy - r,
                    )
                )
                p_tl.Fill = q_fill
                self._canvas_set_zindex(p_tl, zi + 1)
                canv.Children.Add(p_tl)

                p_br = Path()
                p_br.Data = Geometry.Parse(
                    u"M {0},{1} L {2},{1} A {3},{3} 0 0 1 {0},{4} Z".format(
                        cx, cy, cx + r, r, cy + r,
                    )
                )
                p_br.Fill = q_fill
                self._canvas_set_zindex(p_br, zi + 2)
                canv.Children.Add(p_br)
            except Exception:
                pass

            rim = Ellipse()
            rim.Width = dia
            rim.Height = dia
            rim.Fill = Brushes.Transparent
            rim.Stroke = br_rim
            rim.StrokeThickness = 0.9
            _Cn.SetLeft(rim, bx)
            _Cn.SetTop(rim, by)
            self._canvas_set_zindex(rim, zi + 3)
            canv.Children.Add(rim)

            lead = Line()
            lead.X1 = cx - r - max(3.0, 0.35 * r)
            lead.Y1 = cy
            lead.X2 = cx - r
            lead.Y2 = cy
            lead.Stroke = br_lead
            lead.StrokeThickness = 0.9
            self._canvas_set_zindex(lead, zi + 4)
            canv.Children.Add(lead)

            guide = Line()
            guide.X1 = cx + r + 1.0
            guide.Y1 = cy
            guide.X2 = ruler_w + 800.0
            guide.Y2 = cy
            guide.Stroke = br_guide
            guide.StrokeThickness = 0.7
            guide.Opacity = 0.45
            try:
                dashes = DoubleCollection()
                dashes.Add(3.0)
                dashes.Add(4.0)
                guide.StrokeDashArray = dashes
            except Exception:
                pass
            self._canvas_set_zindex(guide, 2)
            canv.Children.Add(guide)

            if is_fund:
                label = u"Fund."
            else:
                label = u"{:.3f}".format(level_m)

            tb = TextBlock()
            tb.Text = label
            tb.Foreground = br_head_text
            tb.FontSize = 10.5
            tb.FontWeight = FontWeights.SemiBold

            try:
                from System import Double
                from System.Windows import Size as _Sz

                tb.Measure(_Sz(Double.PositiveInfinity, Double.PositiveInfinity))
                tw = float(tb.DesiredSize.Width)
                th = float(tb.DesiredSize.Height)
            except Exception:
                tw, th = 46.0, 14.0

            lab_right = cx - r - max(3.0, 0.35 * r) - 3.0
            lab_x = max(1.0, lab_right - tw)
            lab_y = cy - th - 2.0
            lab_y = min(lab_y, total_h - th - 1.0)

            _Cn.SetLeft(tb, lab_x)
            _Cn.SetTop(tb, lab_y)
            self._canvas_set_zindex(tb, zi + 10)
            canv.Children.Add(tb)

        for ri, bnd in enumerate(boundaries):
            if ri == 0:
                _draw_level_head(bnd[u"top_y"], bnd[u"z_top"])
            _draw_level_head(bnd[u"bot_y"], bnd[u"z_bot"])
            if bnd[u"fund_h"] > 0.0:
                fund_y = bnd[u"bot_y"] + bnd[u"fund_h"]
                _draw_level_head(
                    fund_y,
                    bnd[u"z_bot"],
                    is_fund=True,
                    accent_hex=bnd.get(u"color_hex"),
                )

    def _redraw_level_ruler(self):
        u"""Redibuja la regleta de niveles si existe."""
        ruler_cv = getattr(self, "_ruler_canvas", None)
        if ruler_cv is None:
            return
        heights = self._compute_row_heights()
        od = getattr(self, "_walls_display_order", []) or []
        if len(heights) < len(od):
            heights = [float(getattr(self, "_row_height_px", 128.0))] * len(od)
        _top_pad = 20.0 if getattr(self, "_ruler_top_row", 0) > 0 else 0.0
        total_h = sum(float(h) for h in heights) + _top_pad
        ruler_cv.Height = total_h
        self._draw_level_ruler(ruler_cv, heights)

    def _draw_wall_polygon_on_canvas(self, canv, md, wall, row_index):
        from System.Windows.Controls import Canvas as _Cn
        from System.Windows.Shapes import Polygon as _Wp
        from System.Windows import Point, FontWeights, Thickness
        from System.Windows.Media import SolidColorBrush, Color, PointCollection
        from System.Windows.Controls import Border as _Bd

        canv.Children.Clear()

        def _clamp_byte(iv):
            return max(0, min(255, int(iv)))

        def _brush_hex(hx, aa=238):
            h = hx.strip().lstrip("#")
            rr = int(h[0:2], 16)
            gg = int(h[2:4], 16)
            bb = int(h[4:6], 16)
            return SolidColorBrush(Color.FromArgb(_clamp_byte(aa), _clamp_byte(rr), _clamp_byte(gg), _clamp_byte(bb)))

        stk = self._wall_elevation_color_hex(wall, row_index)

        wpx = self._finite_px(
            canv.Width,
            max(40.0, float(self._PREVIEW_ELEV_COL_PX)),
        )
        if wpx < 25.0:
            wpx = max(40.0, float(self._PREVIEW_ELEV_COL_PX))

        hpx_row = self._finite_px(
            canv.Height,
            float(getattr(self, "_row_height_px", 128.0)),
        )
        if hpx_row < 10.0:
            hpx_row = float(getattr(self, "_row_height_px", 128.0))

        box_w = max(10.0, wpx)
        box_h = max(10.0, hpx_row)
        od = getattr(self, "_walls_display_order", []) or []
        n_od = len(od)
        is_first_row = row_index <= 0
        is_last_row = row_index >= n_od - 1
        x0 = 0.0
        y0 = 0.0
        stripe_h = box_h
        x_off, draw_w = self._wall_draw_width_px(row_index, box_w)
        lvl_layout = self._preview_level_layout(box_w, x_off, draw_w)

        fund_info = None
        if wall is not None:
            try:
                wid_f = _wall_id_int(wall)
                fund_info = self._ensure_foundation_preview_info(wid_f, wall)
            except Exception:
                fund_info = None

        foot_h = self._foundation_footing_height_px(box_h, fund_info)
        wall_draw_h = max(10.0, float(box_h) - float(foot_h)) if foot_h > 0.0 else float(box_h)

        fill_br = _brush_hex(stk, aa=34)
        stroke_br = _brush_hex(stk, aa=170)
        br_lvl = self._level_theme_brushes()

        if self._uses_cabezal_panels():
            band_border = Thickness(1.0)
        elif is_first_row and is_last_row:
            band_border = Thickness(1.0)
        elif is_first_row:
            band_border = Thickness(1.0, 1.0, 1.0, 0.0)
        elif is_last_row:
            band_border = Thickness(1.0, 0.0, 1.0, 0.0 if foot_h > 0.0 else 1.0)
        else:
            band_border = Thickness(1.0, 0.0, 1.0, 0.0)

        if not self._uses_cabezal_panels():
            self._add_level_zone_delimiter(
                canv,
                lvl_layout[u"gutter_left"],
                0.0,
                box_h,
                3,
            )

        band = _Bd()
        band.Width = draw_w
        band.Height = wall_draw_h
        band.Background = fill_br
        band.BorderBrush = stroke_br
        band.BorderThickness = band_border
        _Cn.SetLeft(band, x_off)
        _Cn.SetTop(band, y0)
        self._canvas_set_zindex(band, 1)
        try:
            band.SnapsToDevicePixels = True
        except Exception:
            pass
        if self._uses_cabezal_panels() and wall is not None:
            try:
                th_mm = self._wall_thickness_mm_key(wall)
                band.ToolTip = u"Espesor muro: {0} mm".format(th_mm)
            except Exception:
                pass
        canv.Children.Add(band)

        if fund_info and foot_h > 0.0:
            self._draw_foundation_schematic(
                canv, x_off, draw_w, box_h, foot_h, fund_info, color_hex=stk,
            )

        if md is not None:
            try:
                poly_uv = md["poly_uv_feet"]
                umin = float(md["u_min_feet"])
                umax = float(md["u_max_feet"])
                vmin = float(md["v_min_feet"])
                vmax = float(md["v_max_feet"])
            except Exception:
                poly_uv = None

            if poly_uv and len(poly_uv) >= 3:
                du = max(umax - umin, 1e-9)
                dv = max(vmax - vmin, 1e-9)

                plc = PointCollection()
                for uf, vf in poly_uv:
                    u_norm = (float(uf) - umin) / du
                    v_norm = (vmax - float(vf)) / dv
                    x = x_off + u_norm * draw_w
                    y = y0 + v_norm * wall_draw_h
                    plc.Add(Point(x, y))

                polygon = _Wp()
                polygon.Points = plc
                polygon.Fill = fill_br
                polygon.Stroke = stroke_br
                polygon.StrokeThickness = 0.0
                try:
                    polygon.SnapsToDevicePixels = True
                except Exception:
                    pass
                self._canvas_set_zindex(polygon, 2)
                canv.Children.Add(polygon)

        if wall is not None:
            tipo_line, esp_line = self._wall_section_label_lines(wall)
            self._add_canvas_label_block_centered(
                canv,
                [tipo_line, esp_line],
                x_off + draw_w / 2.0,
                wall_draw_h / 2.0,
                br_lvl[u"text"],
                10.0,
                FontWeights.SemiBold,
                25,
            )

        if wall is not None:
            self._draw_elevation_vecino_encuentros(
                canv, wall, row_index, x_off, draw_w, y0, wall_draw_h,
            )

        if wall is not None and not self._uses_cabezal_panels():
            try:
                z_m = float(geo.cota_inferior_muro_metros_aprox(wall))
                level_txt = u"{:.3f}".format(z_m)
            except Exception:
                level_txt = u"—"
            if foot_h > 0.0:
                y_base = max(2.0, float(box_h) - float(foot_h) - 1.0)
            elif is_last_row:
                y_base = max(2.0, float(box_h) - 1.0)
            else:
                y_base = max(2.0, float(box_h) - 1.0)
            self._add_level_head_at_wall_base(canv, y_base, lvl_layout, 40)
            self._add_level_label_at_base(canv, y_base, level_txt, lvl_layout, 50)

        marks_canv = canv
        if wall is not None and self._is_unificado_mode():
            try:
                wid_m = _wall_id_int(wall)
                alt_cv = (self._cabezal_ui_by_wall_id.get(wid_m) or {}).get(
                    u"chevron_canvas",
                )
                if alt_cv is not None:
                    alt_cv.Children.Clear()
                    marks_canv = alt_cv
                    try:
                        alt_cv.Width = canv.Width
                        alt_cv.Height = canv.Height
                    except Exception:
                        pass
            except Exception:
                pass

        if self._uses_cabezal_panels() and wall is not None:
            self._draw_cabezal_extremo_marks_elevation(
                marks_canv, wall, row_index, x_off, draw_w, y0, wall_draw_h,
            )
            self._draw_cabezal_empalme_marks_elevation(
                marks_canv, wall, row_index, x_off, draw_w, wall_draw_h,
            )
            self._draw_cabezal_troceo_pie_controls_elevation(
                marks_canv, wall, row_index, x_off, draw_w, wall_draw_h, box_h,
            )
        if wall is not None and self._is_unificado_mode():
            try:
                self._position_mesh_overlay_for_wall(wall, row_index, canv)
            except Exception:
                pass

    def _add_cabezal_pie_hit_rect(
        self, canv, x, y, w, h, wid, extremo, caption, pinned, active, auto_geom,
    ):
        from System.Windows.Controls import Canvas as _Cn
        from System.Windows.Input import Cursors
        from System.Windows.Media import SolidColorBrush, Color
        from System.Windows.Shapes import Rectangle as _Wr

        rect = _Wr()
        rect.Width = float(w)
        rect.Height = float(h)
        rect.RadiusX = 3.0
        rect.RadiusY = 3.0
        if pinned:
            fill = Color.FromArgb(235, 34, 211, 238)
            stroke = Color.FromArgb(255, 14, 116, 144)
        elif active and auto_geom:
            fill = Color.FromArgb(225, 180, 60, 60)
            stroke = Color.FromArgb(255, 220, 80, 80)
        elif active:
            fill = Color.FromArgb(210, 34, 120, 160)
            stroke = Color.FromArgb(255, 20, 80, 110)
        else:
            fill = Color.FromArgb(200, 45, 58, 72)
            stroke = Color.FromArgb(230, 80, 100, 120)
        rect.Fill = SolidColorBrush(fill)
        rect.Stroke = SolidColorBrush(stroke)
        rect.StrokeThickness = 1.0
        try:
            rect.Cursor = Cursors.Hand
        except Exception:
            pass
        ex_lbl = self._cabezal_extremo_ui_label(extremo)
        tip = u"{0}: clic cicla Auto → Tramo → Continuar".format(ex_lbl)
        if pinned:
            cap = caption or u""
            tip += u" ({0})".format(cap)
        elif auto_geom and active:
            tip += u" (geom.)"
        try:
            rect.ToolTip = tip
        except Exception:
            pass
        _Cn.SetLeft(rect, float(x))
        _Cn.SetTop(rect, float(y))
        self._canvas_set_zindex(rect, 55)

        def _on_click(sender, args, w=wid, ex=extremo):
            try:
                self._cycle_cabezal_troceo_at_pie(w, ex)
            except Exception:
                pass

        try:
            from System.Windows.Input import MouseButtonEventHandler as _MBEH
            rect.MouseLeftButtonDown += _MBEH(_on_click)
        except Exception:
            pass
        canv.Children.Add(rect)

        if caption:
            try:
                from System.Windows import FontWeights
                self._add_canvas_text_centered(
                    canv,
                    caption,
                    float(x) + float(w) / 2.0,
                    float(y) + float(h) / 2.0,
                    SolidColorBrush(Color.FromRgb(240, 248, 255)),
                    6.5,
                    FontWeights.SemiBold,
                    56,
                )
            except Exception:
                pass

    def _draw_cabezal_troceo_pie_controls_elevation(
        self, canv, wall, row_index, x_off, draw_w, wall_draw_h_px, box_h_px=None,
    ):
        if not self._uses_cabezal_panels() or cabezal is None or wall is None:
            return
        if self._cabezal_ordered_index_for_wall(wall) < 1:
            return
        try:
            wid = _wall_id_int(wall)
        except Exception:
            return
        row_h = max(float(wall_draw_h_px), float(box_h_px or wall_draw_h_px))
        y_joint = self._empalme_tick_y_wall_base_px(row_h)
        btn_w = 54.0
        btn_h = 18.0
        y_btn = max(2.0, y_joint - btn_h - 6.0)
        x_left = float(x_off) + 4.0
        x_right = float(x_off) + max(14.0, float(draw_w)) - btn_w - 4.0

        try:
            from System.Windows.Shapes import Line as _Ln
            from System.Windows.Media import SolidColorBrush, Color, DoubleCollection
            ln = _Ln()
            ln.X1 = float(x_off) + 6.0
            ln.X2 = float(x_off) + max(14.0, float(draw_w)) - 6.0
            ln.Y1 = y_joint
            ln.Y2 = y_joint
            ln.Stroke = SolidColorBrush(Color.FromArgb(140, 120, 140, 160))
            ln.StrokeThickness = 1.0
            try:
                ln.StrokeDashArray = DoubleCollection([3.0, 2.0])
            except Exception:
                pass
            self._canvas_set_zindex(ln, 52)
            canv.Children.Add(ln)
        except Exception:
            pass

        ex_izq, ex_der = self._cabezal_extremos_lados_wall(wid, row_index)
        if self._cabezal_armado_activo_cfg(wid, ex_izq):
            prefix = (
                u"I"
                if ex_izq == cabezal.CABEZAL_EXTREMO_INICIO
                else u"F"
            )
            cap = self._cabezal_pie_selector_caption(wid, ex_izq)
            self._add_cabezal_pie_hit_rect(
                canv, x_left, y_btn, btn_w, btn_h, wid, ex_izq,
                u"{0}·{1}".format(prefix, cap),
                self._cabezal_troceo_is_manual(wid, ex_izq),
                self._cabezal_troceo_por_muro_activo(wid, ex_izq),
                self._cabezal_auto_troceo_for_wall(wid, ex_izq),
            )
        if self._cabezal_armado_activo_cfg(wid, ex_der):
            prefix = (
                u"I"
                if ex_der == cabezal.CABEZAL_EXTREMO_INICIO
                else u"F"
            )
            cap = self._cabezal_pie_selector_caption(wid, ex_der)
            self._add_cabezal_pie_hit_rect(
                canv, x_right, y_btn, btn_w, btn_h, wid, ex_der,
                u"{0}·{1}".format(prefix, cap),
                self._cabezal_troceo_is_manual(wid, ex_der),
                self._cabezal_troceo_por_muro_activo(wid, ex_der),
                self._cabezal_auto_troceo_for_wall(wid, ex_der),
            )

    def _redraw_preview_canvas(self):
        self._schedule_full_redraw()

    def _schedule_crear_and_close(self):
        u"""Cierra el modal y dispara ExternalEvent tras ShowDialog (Revit no procesa Raise con diálogo abierto)."""
        self._defer_crear_raise = True
        try:
            if self._win is not None:
                self._win.Close()
        except Exception:
            pass

    def _raise_deferred_crear_if_needed(self):
        if not getattr(self, "_defer_crear_raise", False):
            return
        self._defer_crear_raise = False
        try:
            self._crear_event.Raise()
        except Exception as ex_ra:
            try:
                if self._is_unificado_mode():
                    titulo = u"Arainco: Armado Muros — Error"
                elif self._is_cabezal_mode():
                    titulo = u"Armado muros cabezal — Error"
                else:
                    titulo = u"Armado muros mallas — Error"
                TaskDialog.Show(titulo, unicode(ex_ra))
            except Exception:
                try:
                    TaskDialog.Show(u"Armado muros — Error", str(ex_ra))
                except Exception:
                    pass

    def _on_crear_clicked(self):
        if not self._ensure_ui_ready_for_crear():
            try:
                TaskDialog.Show(
                    u"Armado muros",
                    u"La interfaz aún se está cargando. Espere un momento e intente de nuevo.",
                )
            except Exception:
                pass
            return
        if self._is_cabezal_mode():
            self._on_crear_cabezal_clicked()
            return
        if self._is_unificado_mode():
            self._on_crear_unificado_clicked()
            return
        self._on_crear_mallas_clicked()

    def _on_crear_unificado_clicked(self):
        if cabezal is None:
            TaskDialog.Show(u"Arainco: Armado Muros", u"Módulo cabezal no disponible.")
            return
        if not self.walls_ordered:
            TaskDialog.Show(u"Arainco: Armado Muros", u"No hay muros.")
            self._set_estado(u"No hay muros cargados.")
            return

        self._sync_all_cabezal_troceo_auto()
        self._sync_cabezal_from_segment_owners(sync_confinement=False)

        cabezal_por = {}
        for wall in self.walls_ordered:
            wid = _wall_id_int(wall)
            cfg_c = self._cabezal_by_wall_id.get(wid)
            if cfg_c:
                for ex in cabezal.CABEZAL_EXTREMOS:
                    ex_cfg = cfg_c.get(ex)
                    if ex_cfg:
                        cabezal._normalize_cabezal_extremo_layers(ex_cfg)
                        segs = cabezal.build_cabezal_segments(
                            len(self.walls_ordered),
                            cabezal._empalme_stack_indices(
                                self.walls_ordered, self._cabezal_by_wall_id, ex,
                            ),
                        )
                        cabezal._migrate_tramo_to_segment_bar_type_ids(
                            ex_cfg, segs, ex_cfg.get(u"bar_type_id"),
                        )
                        if cabezal.cabezal_extremo_armado_activo(ex_cfg):
                            conf_val = self._read_cabezal_confinement_combo(wid, ex)
                            cabezal.cabezal_stamp_confinement_type(
                                ex_cfg, conf_val, self.doc,
                            )
            ok_c, msg_c = cabezal.validar_cabezal_config(cfg_c)
            if not ok_c:
                TaskDialog.Show(
                    u"Arainco: Armado Muros",
                    u"Muro {0}: {1}".format(wid, msg_c),
                )
                self._set_estado(msg_c)
                return
            cabezal_por[wid] = cabezal.cabezal_copy_muro_config(cfg_c)

        params_por = {}
        for wall in self.walls_ordered:
            wid = _wall_id_int(wall)
            ctr = self._controls_by_wall_id.get(wid)
            if not ctr:
                TaskDialog.Show(
                    u"Arainco: Armado Muros",
                    u"Panel interno incompleto (muro {}).".format(wid),
                )
                return
            pd, ld = self._gather_params_for_wall_id(wid)
            if self._malla_activo_cfg(wid):
                for lk in (
                    u"exterior_major", u"exterior_minor",
                    u"interior_major", u"interior_minor",
                ):
                    bid = pd.get(lk, (ElementId.InvalidElementId, u""))[0]
                    if bid is None or bid == ElementId.InvalidElementId:
                        TaskDialog.Show(
                            u"Arainco: Armado Muros",
                            u"Muro {}: elige barra en malla ({}).".format(wid, lk),
                        )
                        self._set_estado(u"Falta diámetro malla en muro {}.".format(wid))
                        return
            params_por[wid] = (pd, ld)

        tid = geo._get_default_area_reinforcement_type_id(self.doc)
        if not tid or tid == ElementId.InvalidElementId:
            TaskDialog.Show(
                u"Arainco: Armado Muros",
                u"No hay Area Reinforcement Type en el proyecto.",
            )
            self._set_estado(u"Sin tipo AR.")
            return

        self._crear_handler.walls = list(self.walls_ordered)
        self._crear_handler.params_por_muro_id = params_por
        self._crear_handler.cabezal_por_muro_id = cabezal_por
        self._crear_handler.area_reinforcement_type_id = tid
        self._crear_handler.malla_activo_por_muro_id = (
            self._malla_activo_por_muro_id_dict()
        )
        self._schedule_crear_and_close()

    def _on_crear_cabezal_clicked(self):
        if cabezal is None:
            TaskDialog.Show(u"Armado muros", u"Módulo cabezal no disponible.")
            return
        if not self.walls_ordered:
            TaskDialog.Show(u"Armado muros", u"No hay muros.")
            self._set_estado(u"No hay muros cargados.")
            return

        self._sync_all_cabezal_troceo_auto()
        self._sync_cabezal_from_segment_owners(sync_confinement=False)

        cabezal_por = {}
        for wall in self.walls_ordered:
            wid = _wall_id_int(wall)
            cfg_c = self._cabezal_by_wall_id.get(wid)
            if cfg_c:
                for ex in cabezal.CABEZAL_EXTREMOS:
                    ex_cfg = cfg_c.get(ex)
                    if ex_cfg:
                        cabezal._normalize_cabezal_extremo_layers(ex_cfg)
                        segs = cabezal.build_cabezal_segments(
                            len(self.walls_ordered),
                            cabezal._empalme_stack_indices(
                                self.walls_ordered, self._cabezal_by_wall_id, ex,
                            ),
                        )
                        cabezal._migrate_tramo_to_segment_bar_type_ids(
                            ex_cfg, segs, ex_cfg.get(u"bar_type_id"),
                        )
                        if cabezal.cabezal_extremo_armado_activo(ex_cfg):
                            conf_val = self._read_cabezal_confinement_combo(wid, ex)
                            cabezal.cabezal_stamp_confinement_type(
                                ex_cfg, conf_val, self.doc,
                            )
            ok_c, msg_c = cabezal.validar_cabezal_config(cfg_c)
            if not ok_c:
                TaskDialog.Show(
                    u"Armado muros",
                    u"Muro {0}: {1}".format(wid, msg_c),
                )
                self._set_estado(msg_c)
                return
            cabezal_por[wid] = cabezal.cabezal_copy_muro_config(cfg_c)

        self._crear_handler.walls = list(self.walls_ordered)
        self._crear_handler.cabezal_por_muro_id = cabezal_por
        self._crear_handler.ref_walls_troceo = None
        self._schedule_crear_and_close()

    def _seleccionar_muros_referencia_troceo(self):
        """
        Pide al usuario seleccionar muros de referencia para los planos de
        troceo. Retorna lista de Wall o None si cancela/no selecciona.
        """
        from Autodesk.Revit.UI import TaskDialog as _TD
        from Autodesk.Revit.UI import TaskDialogCommonButtons, TaskDialogResult
        td = _TD(u"Arainco: Troceo cabezal")
        td.MainInstruction = u"¿Desea seleccionar muros de referencia para troceo?"
        td.MainContent = (
            u"Los muros de referencia definen planos de corte horizontales.\n"
            u"Las barras fusionadas se cortarán en esos planos con alternancia A/B.\n\n"
            u"Si no selecciona muros, las barras se crearán sin troceo."
        )
        td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
        result = td.Show()
        if result != TaskDialogResult.Yes:
            return None
        try:
            from Autodesk.Revit.UI.Selection import ObjectType
            uidoc = self._uidoc
            if uidoc is None:
                return None
            self._set_estado(u"Seleccione muros de referencia para troceo (Esc para terminar)…")
            try:
                self._win.Hide()
            except Exception:
                pass
            ref_ids = uidoc.Selection.PickObjects(
                ObjectType.Element,
                u"Seleccione muros de referencia para planos de corte (troceo).",
            )
            try:
                self._win.Show()
            except Exception:
                pass
            if not ref_ids:
                return None
            doc = uidoc.Document
            ref_walls = []
            for r in ref_ids:
                el = doc.GetElement(r.ElementId)
                if el is not None and isinstance(el, Wall):
                    ref_walls.append(el)
            if ref_walls:
                self._set_estado(
                    u"{0} muros de referencia para troceo.".format(len(ref_walls)),
                )
                return ref_walls
            return None
        except Exception:
            try:
                self._win.Show()
            except Exception:
                pass
            return None

    def _on_crear_mallas_clicked(self):
        if not self.walls_ordered:
            TaskDialog.Show("Armado muros", u"No hay muros.")
            self._set_estado(u"No hay muros cargados.")
            return

        params_por = {}

        for wall in self.walls_ordered:
            wid = _wall_id_int(wall)
            ctr = self._controls_by_wall_id.get(wid)
            if not ctr:
                TaskDialog.Show("Armado muros", u"Panel interno incompleto (muro {}).".format(wid))
                return

            pd, ld = self._gather_params_for_wall_id(wid)
            if self._malla_activo_cfg(wid):
                need = [
                    "exterior_major", "exterior_minor",
                    "interior_major", "interior_minor",
                ]
                bad = []
                for lk in need:
                    bid = pd.get(lk, (ElementId.InvalidElementId, ""))[0]
                    if bid is None or bid == ElementId.InvalidElementId:
                        bad.append(lk)
                if bad:
                    TaskDialog.Show(
                        "Armado muros",
                        u"Muro {}: elige barra en capas activas.".format(wid),
                    )
                    self._set_estado(u"Falta diámetro en muro {}.".format(wid))
                    return

            params_por[wid] = (pd, ld)

        cabezal_por = {}

        tid = geo._get_default_area_reinforcement_type_id(self.doc)
        self._sync_modo_from_checkboxes()
        self._crear_handler.walls = list(self.walls_ordered)
        self._crear_handler.params_por_muro_id = params_por
        self._crear_handler.cabezal_por_muro_id = cabezal_por
        self._crear_handler.area_reinforcement_type_id = tid if tid else ElementId.InvalidElementId
        self._crear_handler.muro_contencion = not self._mesh_modo_tradicional()
        self._crear_handler.malla_activo_por_muro_id = (
            self._malla_activo_por_muro_id_dict()
        )

        if not tid or tid == ElementId.InvalidElementId:
            TaskDialog.Show(
                "Armado muros",
                u"No hay Area Reinforcement Type en el proyecto.",
            )
            self._set_estado(u"Sin tipo AR.")
            return

        self._schedule_crear_and_close()

    def _attach_revit_owner(self):
        u"""Ventana hija de Revit → ``ShowDialog`` bloquea interacción con el modelo."""
        if self._win is None:
            return
        try:
            from System.Windows.Interop import WindowInteropHelper

            hwnd = None
            try:
                from revit_wpf_window_position import revit_main_hwnd

                uiapp = None
                try:
                    uiapp = self._uidoc.Application if self._uidoc is not None else None
                except Exception:
                    uiapp = None
                if uiapp is None:
                    uiapp = getattr(self._revit, "Application", None) or self._revit
                hwnd = revit_main_hwnd(uiapp)
            except Exception:
                hwnd = None
            if hwnd is not None:
                WindowInteropHelper(self._win).Owner = hwnd
        except Exception:
            pass

    def Show(self):
        u"""Muestra la ventana de forma modal (bloquea la UI de Revit hasta cerrar)."""
        if self._win is None:
            return

        try:
            _position_preview_window(
                self._win,
                self._uidoc,
                self._revit,
                self._uses_cabezal_panels(),
                before_snap=self._attach_revit_owner,
            )
        except Exception:
            pass

        try:
            _register_preview_singleton(self._win, self._ui_mode)
        except Exception:
            pass

        try:
            self._win.ShowDialog()
        except Exception as ex_ss:
            try:
                TaskDialog.Show(u"Armado muros preview", unicode(ex_ss))
            except Exception:
                TaskDialog.Show(u"Armado muros preview", str(ex_ss))
        finally:
            try:
                _unregister_preview_singleton(self._ui_mode)
            except Exception:
                pass

        self._raise_deferred_crear_if_needed()


def _show_armado_muros_preview_impl(revit, uidoc, walls_list, mode):
    if mode in (UI_MODE_CABEZAL, UI_MODE_UNIFICADO) and not _require_cabezal_mod():
        return
    key = _preview_singleton_key(mode)
    existing = None
    try:
        existing = _AppDom.CurrentDomain.GetData(key)
    except Exception:
        existing = None
    if existing is not None:
        try:
            w = existing
            if not bool(getattr(w, "IsLoaded", False)):
                raise RuntimeError(u"Ventana anterior cerrada; se libera hueco singleton.")
            w.Activate()
            try:
                from System.Windows import WindowState as _WS

                if w.WindowState == _WS.Minimized:
                    w.WindowState = _WS.Normal
            except Exception:
                pass
            try:
                _position_preview_window(
                    w,
                    uidoc,
                    revit,
                    mode in (UI_MODE_CABEZAL, UI_MODE_UNIFICADO),
                )
            except Exception:
                pass
            if mode == UI_MODE_CABEZAL:
                titulo = u"Armado muros cabezal"
            elif mode == UI_MODE_UNIFICADO:
                titulo = u"Arainco: Armado Muros"
            else:
                titulo = u"Armado muros mallas"
            TaskDialog.Show(
                titulo,
                u"La herramienta ya está en ejecución.",
            )
            return
        except Exception:
            _unregister_preview_singleton(mode)

    preview = ArmadoMurosPreviewWindow(revit, uidoc, walls_list, mode=mode)
    if getattr(preview, "_win", None) is None:
        _unregister_preview_singleton(mode)
        return

    preview.Show()


def show_armado_muros_cabezal(revit, uidoc, walls_list):
    _show_armado_muros_preview_impl(revit, uidoc, walls_list, UI_MODE_CABEZAL)


def show_armado_muros_mallas(revit, uidoc, walls_list):
    _show_armado_muros_preview_impl(revit, uidoc, walls_list, UI_MODE_MALLAS)


def show_armado_muros_unificado(revit, uidoc, walls_list):
    _show_armado_muros_preview_impl(revit, uidoc, walls_list, UI_MODE_UNIFICADO)


def show_armado_muros_preview(revit, uidoc, walls_list):
    """Compatibilidad: abre el asistente de mallas."""
    show_armado_muros_mallas(revit, uidoc, walls_list)
