# -*- coding: utf-8 -*-
"""
Eliminar nubes de revisión — herramienta unificada.

Modos:
  - Lámina actual: nubes de la ViewSheet activa y de GetAllPlacedViews().
  - Múltiples láminas: selección por Clasificacion (excluye Splash); requiere
    parámetro Validacion en Project Information.

Revit 2025+ | pyRevit (CPython 3 / IronPython 3)
"""

from __future__ import print_function

try:
    unicode
except NameError:
    unicode = str

import os

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    FilteredElementCollector,
    Transaction,
    ViewSheet,
)
from Autodesk.Revit.UI import TaskDialog

from System import AppDomain, EventHandler
from System.Collections.Generic import List as ClrList
from System.Windows import (
    FontWeights,
    HorizontalAlignment,
    TextWrapping,
    Thickness,
    VerticalAlignment,
    Visibility,
    WindowState,
)
from System.Windows.Controls import (
    CheckBox,
    GroupBox,
    StackPanel,
    TextBlock,
)
from System.Windows.Input import Cursors, Key, KeyEventHandler
from System.Windows.Markup import XamlReader
from System.Windows.Media import Color, SolidColorBrush

try:
    from Autodesk.Revit.DB import RevisionCloud
except Exception:
    RevisionCloud = None

from bimtools_ui_tokens import BTN_MANUAL, FG_BODY, FG_TITLE, FONT_SIZE_BODY
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_wpf_shell import build_simple_tool_xaml
from revit_wpf_window_position import (
    bind_center_wpf_on_revit_monitor,
    position_wpf_window_center_on_monitor,
    revit_main_hwnd,
)

_APPDOMAIN_WINDOW_KEY = u"BIMTools.EliminarNubesRevision.ActiveWindow"
_TOOL_TITLE = u"Arainco: Eliminar nubes de revisión"
_TRANSACTION_NAME = u"Arainco: Eliminar nubes de revisión"
_ALREADY_RUNNING = u"La herramienta ya esta en ejecucion."

_PARAM_VALIDACION = u"Validacion"
_PARAM_CLASIFICACION = u"Clasificacion"
_SPLASH_SUBSTR = u"Splash"
_GROUP_PREFIX = u"Láminas "

_MODE_ACTUAL = u"actual"
_MODE_MULTI = u"multi"


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except Exception:
        return str(text)


def _element_id_key(eid):
    if eid is None:
        return None
    try:
        return int(eid.Value)
    except Exception:
        try:
            return int(eid.IntegerValue)
        except Exception:
            return None


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
            _TOOL_TITLE,
            instruction,
            content=content,
            ok_text=ok_text,
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass
    try:
        body = instruction
        if content:
            body = instruction + u"\n\n" + content
        TaskDialog.Show(_TOOL_TITLE, body)
    except Exception:
        pass


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


def _brush(hex_color):
    h = _as_unicode(hex_color).lstrip(u"#")
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    return SolidColorBrush(Color.FromRgb(r, g, b))


def _window_is_alive(win):
    if win is None:
        return False
    try:
        _ = win.Title
    except Exception:
        return False
    try:
        if hasattr(win, "IsLoaded") and (not win.IsLoaded):
            return False
    except Exception:
        pass
    return True


def _find_window_by_tool_title():
    try:
        from System.Windows import Application

        app = Application.Current
        if app is None:
            return None
        for w in app.Windows:
            try:
                txt = w.FindName(u"TxtTitle")
                if txt is not None and _as_unicode(txt.Text) == _TOOL_TITLE:
                    if _window_is_alive(w):
                        return w
            except Exception:
                continue
    except Exception:
        return None
    return None


def _get_active_window():
    try:
        win = AppDomain.CurrentDomain.GetData(_APPDOMAIN_WINDOW_KEY)
    except Exception:
        win = None
    if _window_is_alive(win):
        return win
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, None)
    except Exception:
        pass
    return _find_window_by_tool_title()


def _set_active_window(win):
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, win)
    except Exception:
        pass


def _clear_active_window():
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, None)
    except Exception:
        pass


def _activate_existing(win, uiapp):
    try:
        if win.WindowState == WindowState.Minimized:
            win.WindowState = WindowState.Normal
    except Exception:
        pass
    try:
        win.Activate()
        win.Focus()
    except Exception:
        pass
    _mostrar_aviso(uiapp, _ALREADY_RUNNING)


def _sheet_number(sheet):
    try:
        return _as_unicode(sheet.SheetNumber or u"").strip()
    except Exception:
        return u""


def _sheet_name(sheet):
    try:
        return _as_unicode(sheet.Name or u"").strip()
    except Exception:
        return u""


def _sheet_label(sheet):
    number = _sheet_number(sheet)
    name = _sheet_name(sheet)
    if number and name:
        return u"{0} - {1}".format(number, name)
    return number or name or u"(sin nombre)"


def _iter_revision_clouds_in_view(doc, view_id):
    bic = getattr(BuiltInCategory, "OST_RevisionClouds", None)
    if bic is not None:
        try:
            for el in (
                FilteredElementCollector(doc, view_id)
                .OfCategory(bic)
                .WhereElementIsNotElementType()
            ):
                yield el
            return
        except Exception:
            pass
    if RevisionCloud is None:
        return
    try:
        for el in (
            FilteredElementCollector(doc, view_id)
            .OfClass(RevisionCloud)
            .WhereElementIsNotElementType()
        ):
            yield el
    except Exception:
        return


def collect_revision_cloud_ids(doc, sheet):
    """Nubes visibles en la lámina y en las vistas colocadas (sin duplicados)."""
    view_ids = []
    seen_views = set()

    def _add_view_id(vid):
        key = _element_id_key(vid)
        if key is None or key in seen_views:
            return
        seen_views.add(key)
        view_ids.append(vid)

    _add_view_id(sheet.Id)
    try:
        for vid in sheet.GetAllPlacedViews():
            _add_view_id(vid)
    except Exception:
        pass

    clouds = {}
    for vid in view_ids:
        for el in _iter_revision_clouds_in_view(doc, vid):
            if el is None:
                continue
            key = _element_id_key(el.Id)
            if key is None or key in clouds:
                continue
            clouds[key] = el.Id
    return list(clouds.values())


def _collect_cloud_ids_for_sheets(doc, sheets):
    seen = set()
    ids = []
    for sheet in sheets:
        if sheet is None:
            continue
        try:
            cloud_ids = collect_revision_cloud_ids(doc, sheet)
        except Exception:
            cloud_ids = []
        for eid in cloud_ids:
            key = _element_id_key(eid)
            if key is not None and key not in seen:
                seen.add(key)
                ids.append(eid)
    return ids


def _delete_clouds(doc, cloud_ids):
    if not cloud_ids:
        return 0
    id_list = ClrList[ElementId]()
    for eid in cloud_ids:
        id_list.Add(eid)

    t = Transaction(doc, _TRANSACTION_NAME)
    t.Start()
    try:
        deleted = doc.Delete(id_list)
        t.Commit()
    except Exception:
        try:
            t.RollBack()
        except Exception:
            pass
        raise

    deleted_keys = set()
    if deleted is not None:
        for eid in deleted:
            key = _element_id_key(eid)
            if key is not None:
                deleted_keys.add(key)

    count = 0
    for eid in cloud_ids:
        key = _element_id_key(eid)
        if key in deleted_keys:
            count += 1
            continue
        try:
            if doc.GetElement(eid) is None:
                count += 1
        except Exception:
            count += 1
    return count


def _project_has_validacion(doc):
    pi = None
    try:
        pi = doc.ProjectInformation
    except Exception:
        pi = None
    if pi is None:
        try:
            bic = BuiltInCategory.OST_ProjectInformation
            for el in (
                FilteredElementCollector(doc)
                .OfCategory(bic)
                .WhereElementIsNotElementType()
            ):
                pi = el
                break
        except Exception:
            pi = None
    if pi is None:
        return False
    try:
        return pi.LookupParameter(_PARAM_VALIDACION) is not None
    except Exception:
        return False


def _param_string(element, name):
    try:
        p = element.LookupParameter(name)
    except Exception:
        p = None
    if p is None:
        return u""
    try:
        s = p.AsString()
        if s:
            return _as_unicode(s)
    except Exception:
        pass
    try:
        s = p.AsValueString()
        if s:
            return _as_unicode(s)
    except Exception:
        pass
    return u""


def _collect_sheet_groups(doc):
    """Láminas no Splash, agrupadas por Clasificacion."""
    groups = {}
    order = []
    collector = (
        FilteredElementCollector(doc)
        .OfClass(ViewSheet)
        .WhereElementIsNotElementType()
    )
    for sheet in collector:
        if _SPLASH_SUBSTR in _sheet_name(sheet):
            continue
        clasif = _param_string(sheet, _PARAM_CLASIFICACION)
        if clasif not in groups:
            groups[clasif] = []
            order.append(clasif)
        groups[clasif].append(sheet)
    result = []
    for clasif in order:
        sheets = groups[clasif]
        sheets.sort(key=lambda s: _sheet_label(s).lower())
        result.append((clasif, sheets))
    return result


def _get_active_sheet(uidoc, doc):
    view = None
    try:
        view = uidoc.ActiveView
    except Exception:
        view = None
    if view is None:
        try:
            view = doc.ActiveView
        except Exception:
            view = None
    if view is not None and isinstance(view, ViewSheet):
        return view
    return None


_BODY_XAML = u"""
<Grid>
  <Grid.RowDefinitions>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="*"/>
  </Grid.RowDefinitions>

  <Border Grid.Row="0" Background="#071018" BorderBrush="#21465C"
          BorderThickness="0,0,0,1" Padding="0,0,0,12" Margin="0,0,0,12">
    <StackPanel>
      <TextBlock Text="Ámbito" Style="{{StaticResource LabelSmall}}" Margin="0,0,0,8"/>
      <StackPanel Orientation="Horizontal">
        <RadioButton x:Name="RadioActual" GroupName="ModoNubes" Content="Lámina actual"
                     Margin="0,0,20,0" Foreground="{fg_title}" FontSize="{fs}"
                     VerticalContentAlignment="Center" Cursor="Hand"/>
        <RadioButton x:Name="RadioMulti" GroupName="ModoNubes" Content="Múltiples láminas"
                     Foreground="{fg_title}" FontSize="{fs}"
                     VerticalContentAlignment="Center" Cursor="Hand"/>
      </StackPanel>
    </StackPanel>
  </Border>

  <StackPanel x:Name="PanelActual" Grid.Row="1" Visibility="Collapsed">
    <TextBlock Text="Lámina activa" Style="{{StaticResource LabelSmall}}" Margin="0,0,0,8"/>
    <TextBlock x:Name="TxtSheetActual" TextWrapping="Wrap" FontWeight="SemiBold"
               Foreground="{fg_title}" FontSize="14" Margin="0,0,0,12"/>
    <TextBlock x:Name="TxtCloudsActual" TextWrapping="Wrap"
               Foreground="{fg}" FontSize="{fs}" LineHeight="18"/>
  </StackPanel>

  <Grid x:Name="PanelMulti" Grid.Row="1" Visibility="Collapsed">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
    </Grid.RowDefinitions>
    <DockPanel Grid.Row="0" LastChildFill="True" Margin="0,0,0,10">
      <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
        <Button x:Name="BtnNinguna" Content="Ninguna"
                Style="{{StaticResource BtnSelectOutline}}"
                MinWidth="88" Height="28" Margin="0,0,8,0" Padding="10,0"/>
        <Button x:Name="BtnTodas" Content="Todas"
                Style="{{StaticResource BtnSelectOutline}}"
                MinWidth="88" Height="28" Padding="10,0"/>
      </StackPanel>
      <TextBlock Text="Seleccionar láminas" Style="{{StaticResource Label}}"
                 VerticalAlignment="Center"/>
    </DockPanel>
    <ScrollViewer x:Name="ScrGroups" Grid.Row="1"
                  VerticalScrollBarVisibility="Auto"
                  HorizontalScrollBarVisibility="Disabled"
                  Padding="0,0,4,0">
      <StackPanel x:Name="PanelGroups"/>
    </ScrollViewer>
  </Grid>
</Grid>
""".format(fg=FG_BODY, fg_title=FG_TITLE, fs=FONT_SIZE_BODY)

_FOOTER_LEADING_XAML = (
    u'<Button x:Name="BtnManual" Content="Manual" '
    u'Style="{{StaticResource BtnSelectOutline}}" '
    u'Background="{bg}" MinWidth="96" Padding="8,2" '
    u'ToolTip="Abrir manual de usuario" VerticalAlignment="Center"/>'
).format(bg=BTN_MANUAL)

_FOOTER_ACTIONS_XAML = u"""
<Button x:Name="BtnCancel" Content="Cancelar" Margin="0,0,10,0"
        Style="{StaticResource BtnSelectOutline}" MinWidth="108"/>
<Button x:Name="BtnEliminar" Content="Eliminar" IsDefault="True"
        Style="{StaticResource BtnPrimary}" MinWidth="108"/>
"""


def _resolve_manual_path():
    """Ruta a ``manual_usuario.html`` en la carpeta del pushbutton."""
    candidates = []
    try:
        import bimtools_paths

        pb = bimtools_paths.get_pushbutton_dir()
        if pb:
            candidates.append(os.path.join(pb, u"manual_usuario.html"))
    except Exception:
        pass
    try:
        ext_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        for tab_name in os.listdir(ext_dir):
            if not tab_name.endswith(u".tab"):
                continue
            panel = os.path.join(ext_dir, tab_name, u"Modelado.panel")
            if not os.path.isdir(panel):
                continue
            for pb_name in os.listdir(panel):
                if u"EliminarNubesRevision" not in pb_name:
                    continue
                candidates.append(
                    os.path.join(panel, pb_name, u"manual_usuario.html")
                )
    except Exception:
        pass
    seen = set()
    for path in candidates:
        try:
            ap = os.path.normpath(os.path.abspath(path))
        except Exception:
            continue
        if ap in seen:
            continue
        seen.add(ap)
        if os.path.isfile(ap):
            return ap
    return None


def _open_manual(uiapp):
    path = _resolve_manual_path()
    if not path:
        _mostrar_aviso(
            uiapp,
            u"No se encontró manual_usuario.html.",
            content=u"Debe estar en la carpeta del pushbutton de la herramienta.",
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


def _build_xaml():
    return build_simple_tool_xaml(
        title=_TOOL_TITLE,
        styles_xml=BIMTOOLS_DARK_STYLES_XML,
        body_xaml=_BODY_XAML,
        footer_leading_xaml=_FOOTER_LEADING_XAML,
        footer_actions_xaml=_FOOTER_ACTIONS_XAML,
        footer_hint_xaml=(
            u"Se eliminan las nubes de la lámina y de las vistas colocadas en ella."
        ),
        width=600,
        min_width=520,
        height=720,
        min_height=480,
        resize_mode=u"CanResize",
    )


def _set_element_visible(element, visible):
    """Muestra u oculta un elemento WPF de forma explícita."""
    if element is None:
        return
    try:
        element.Visibility = Visibility.Visible if visible else Visibility.Collapsed
    except Exception:
        pass
    try:
        element.IsHitTestVisible = bool(visible)
    except Exception:
        pass
    try:
        element.Opacity = 1.0 if visible else 0.0
    except Exception:
        pass


def _set_checkboxes(checks, value):
    for cb in checks:
        try:
            cb.IsChecked = value
        except Exception:
            pass


def _fill_groups(win, groups):
    panel = win.FindName(u"PanelGroups")
    checks = []
    if panel is None:
        return checks

    try:
        panel.Children.Clear()
    except Exception:
        pass

    fg_title = _brush(FG_TITLE)
    fg_body = _brush(FG_BODY)
    gb_style = None
    try:
        gb_style = win.TryFindResource(u"GbParams")
    except Exception:
        gb_style = None

    for clasif, sheets in groups:
        gb = GroupBox()
        gb.Margin = Thickness(0, 0, 0, 10)
        if gb_style is not None:
            try:
                gb.Style = gb_style
            except Exception:
                pass

        header = TextBlock()
        header.Text = _GROUP_PREFIX + clasif
        header.FontWeight = FontWeights.SemiBold
        header.FontSize = 12
        try:
            header.Foreground = fg_title
        except Exception:
            pass
        gb.Header = header

        inner = StackPanel()
        inner.Margin = Thickness(2, 2, 2, 2)
        for sheet in sheets:
            cb = CheckBox()
            cb.Content = _sheet_label(sheet)
            cb.IsChecked = False
            cb.Margin = Thickness(0, 3, 0, 3)
            cb.Padding = Thickness(6, 2, 0, 2)
            cb.Cursor = Cursors.Hand
            cb.VerticalAlignment = VerticalAlignment.Center
            cb.HorizontalAlignment = HorizontalAlignment.Stretch
            cb.Tag = sheet
            try:
                cb.Foreground = fg_title
            except Exception:
                pass
            inner.Children.Add(cb)
            checks.append(cb)

        gb.Content = inner
        panel.Children.Add(gb)

    if not groups:
        empty = TextBlock()
        empty.Text = u"No hay láminas para mostrar."
        empty.TextWrapping = TextWrapping.Wrap
        try:
            empty.Foreground = fg_body
        except Exception:
            pass
        empty.Margin = Thickness(4, 8, 4, 8)
        panel.Children.Add(empty)

    return checks


def _show_unified_form(revit, doc, active_sheet, has_validacion, default_mode):
    win = XamlReader.Parse(_build_xaml())
    txt_subtitle = win.FindName(u"TxtSubtitle")
    txt_status = win.FindName(u"TxtStatus")
    panel_actual = win.FindName(u"PanelActual")
    panel_multi = win.FindName(u"PanelMulti")
    radio_actual = win.FindName(u"RadioActual")
    radio_multi = win.FindName(u"RadioMulti")
    txt_sheet_actual = win.FindName(u"TxtSheetActual")
    txt_clouds_actual = win.FindName(u"TxtCloudsActual")
    btn_eliminar = win.FindName(u"BtnEliminar")

    groups = _collect_sheet_groups(doc) if has_validacion else []
    checks = _fill_groups(win, groups) if has_validacion else []

    state = {
        u"mode": default_mode,
        u"suppress": False,
        u"cloud_ids_actual": [],
    }

    if txt_subtitle is not None:
        try:
            txt_subtitle.Text = (
                u"Elija el ámbito: lámina activa o un conjunto de láminas."
            )
        except Exception:
            pass

    def _set_status(text):
        if txt_status is None:
            return
        try:
            txt_status.Text = _as_unicode(text)
        except Exception:
            pass

    def _set_eliminar_enabled(enabled):
        if btn_eliminar is None:
            return
        try:
            btn_eliminar.IsEnabled = bool(enabled)
        except Exception:
            pass

    def _refresh_actual_panel():
        if active_sheet is None:
            if txt_sheet_actual is not None:
                try:
                    txt_sheet_actual.Text = u"(ninguna lámina activa)"
                except Exception:
                    pass
            if txt_clouds_actual is not None:
                try:
                    txt_clouds_actual.Text = (
                        u"Abra una lámina (no un viewport activado)."
                    )
                except Exception:
                    pass
            state[u"cloud_ids_actual"] = []
            return
        cloud_ids = collect_revision_cloud_ids(doc, active_sheet)
        state[u"cloud_ids_actual"] = cloud_ids
        if txt_sheet_actual is not None:
            try:
                txt_sheet_actual.Text = _sheet_label(active_sheet)
            except Exception:
                pass
        if txt_clouds_actual is not None:
            try:
                txt_clouds_actual.Text = u"Nubes encontradas: {0}.".format(
                    len(cloud_ids)
                )
            except Exception:
                pass

    def _apply_mode(mode, from_user=False):
        if mode == _MODE_MULTI and (not has_validacion):
            if from_user:
                _mostrar_aviso(
                    revit,
                    u"Este proyecto no admite el modo múltiples láminas.",
                    content=u"Use una plantilla Arainco de emisión, o el modo lámina actual.",
                )
            mode = _MODE_ACTUAL

        state[u"mode"] = mode
        is_actual = mode == _MODE_ACTUAL

        state[u"suppress"] = True
        try:
            if radio_actual is not None:
                radio_actual.IsChecked = bool(is_actual)
            if radio_multi is not None:
                radio_multi.IsChecked = bool(not is_actual)
        except Exception:
            pass
        state[u"suppress"] = False

        _set_element_visible(panel_actual, is_actual)
        _set_element_visible(panel_multi, not is_actual)

        if is_actual:
            _refresh_actual_panel()
            if active_sheet is None:
                _set_status(u"Abra una lámina para usar este modo.")
                _set_eliminar_enabled(False)
            else:
                n = len(state[u"cloud_ids_actual"])
                _set_status(u"Lámina activa. Nubes: {0}.".format(n))
                _set_eliminar_enabled(True)
        else:
            n_sheets = 0
            for _clasif, sheets in groups:
                n_sheets += len(sheets)
            _set_status(
                u"{0} lámina(s) en {1} grupo(s).".format(n_sheets, len(groups))
            )
            _set_eliminar_enabled(True)

    def _on_radio_actual(sender, args):
        if state[u"suppress"]:
            return
        try:
            if radio_actual is not None and radio_actual.IsChecked:
                _apply_mode(_MODE_ACTUAL, from_user=True)
        except Exception:
            pass

    def _on_radio_multi(sender, args):
        if state[u"suppress"]:
            return
        try:
            if radio_multi is not None and radio_multi.IsChecked:
                _apply_mode(_MODE_MULTI, from_user=True)
        except Exception:
            pass

    def _selected_sheets():
        selected = []
        seen = set()
        for cb in checks:
            try:
                if cb.IsChecked != True:
                    continue
                sheet = cb.Tag
                key = _element_id_key(sheet.Id)
            except Exception:
                continue
            if key is not None and key not in seen:
                seen.add(key)
                selected.append(sheet)
        return selected

    def _on_close(sender, args):
        _clear_active_window()

    def _on_key_down(sender, args):
        if args.Key == Key.Escape:
            try:
                win.Close()
            except Exception:
                pass
            args.Handled = True

    def _on_cancel(sender, args):
        try:
            win.Close()
        except Exception:
            pass

    def _on_todas(sender, args):
        _set_checkboxes(checks, True)
        _set_status(u"{0} lámina(s) marcada(s).".format(len(checks)))

    def _on_ninguna(sender, args):
        _set_checkboxes(checks, False)
        _set_status(u"Ninguna lámina marcada.")

    def _on_eliminar(sender, args):
        mode = state[u"mode"]
        if mode == _MODE_ACTUAL:
            if active_sheet is None:
                _set_status(u"Abra una lámina para usar este modo.")
                return
            cloud_ids = list(state[u"cloud_ids_actual"])
            if not cloud_ids:
                cloud_ids = collect_revision_cloud_ids(doc, active_sheet)
            if not cloud_ids:
                _set_status(u"No hay nubes de revisión en la lámina activa.")
                return
        else:
            sheets = _selected_sheets()
            if not sheets:
                _set_status(u"Seleccione al menos una lámina.")
                return
            try:
                cloud_ids = _collect_cloud_ids_for_sheets(doc, sheets)
            except Exception as ex:
                _mostrar_aviso(
                    revit,
                    u"No se pudieron recoger las nubes de revisión.",
                    content=_as_unicode(ex),
                )
                return
            if not cloud_ids:
                _set_status(
                    u"No hay nubes de revisión en las láminas seleccionadas."
                )
                return

        try:
            deleted = _delete_clouds(doc, cloud_ids)
        except Exception as ex:
            _mostrar_aviso(
                revit,
                u"No se pudieron eliminar las nubes de revisión.",
                content=_as_unicode(ex),
            )
            _set_status(u"Error al eliminar.")
            return

        try:
            win.Close()
        except Exception:
            pass

        if deleted:
            _mostrar_aviso(
                revit,
                u"Se han eliminado {0} nubes de revisión.".format(deleted),
                ok_text=u"Aceptar",
            )

    from System.Windows import RoutedEventHandler

    win.Closed += EventHandler(_on_close)
    win.PreviewKeyDown += KeyEventHandler(_on_key_down)

    btn_cancel = win.FindName(u"BtnCancel")
    if btn_cancel is not None:
        btn_cancel.Click += RoutedEventHandler(_on_cancel)
    if btn_eliminar is not None:
        btn_eliminar.Click += RoutedEventHandler(_on_eliminar)
    btn_manual = win.FindName(u"BtnManual")
    if btn_manual is not None:
        btn_manual.Click += RoutedEventHandler(
            lambda s, e: _open_manual(revit)
        )
    btn_todas = win.FindName(u"BtnTodas")
    if btn_todas is not None:
        btn_todas.Click += RoutedEventHandler(_on_todas)
    btn_ninguna = win.FindName(u"BtnNinguna")
    if btn_ninguna is not None:
        btn_ninguna.Click += RoutedEventHandler(_on_ninguna)

    if radio_actual is not None:
        radio_actual.Checked += RoutedEventHandler(_on_radio_actual)
    if radio_multi is not None:
        radio_multi.Checked += RoutedEventHandler(_on_radio_multi)

    state[u"suppress"] = True
    try:
        if radio_actual is not None:
            radio_actual.IsChecked = False
        if radio_multi is not None:
            radio_multi.IsChecked = False
    except Exception:
        pass
    state[u"suppress"] = False
    _apply_mode(default_mode, from_user=False)

    _prepare_window(win, revit)
    _set_active_window(win)
    try:
        win.ShowDialog()
    finally:
        _clear_active_window()


def run(revit):
    """Punto de entrada pyRevit."""
    existing = _get_active_window()
    if existing is not None:
        _activate_existing(existing, revit)
        return

    uidoc = None
    try:
        uidoc = revit.ActiveUIDocument
    except Exception:
        uidoc = None
    if uidoc is None:
        _mostrar_aviso(revit, u"No hay documento activo.")
        return

    doc = uidoc.Document
    if doc is None:
        _mostrar_aviso(revit, u"No hay documento activo.")
        return

    try:
        if doc.IsFamilyDocument:
            _mostrar_aviso(
                revit,
                u"Abra un proyecto con láminas.",
                content=u"Esta herramienta no se ejecuta en un documento de familia.",
            )
            return
    except Exception:
        pass

    try:
        if doc.IsReadOnly:
            _mostrar_aviso(revit, u"El documento está en solo lectura.")
            return
    except Exception:
        pass

    active_sheet = _get_active_sheet(uidoc, doc)
    has_validacion = _project_has_validacion(doc)

    if active_sheet is None and (not has_validacion):
        _mostrar_aviso(
            revit,
            u"No se puede abrir la herramienta.",
            content=(
                u"Abra una lámina, o trabaje en un proyecto / plantilla Arainco "
                u"de emisión."
            ),
        )
        return

    if active_sheet is not None:
        default_mode = _MODE_ACTUAL
    else:
        default_mode = _MODE_MULTI

    try:
        _show_unified_form(
            revit, doc, active_sheet, has_validacion, default_mode
        )
    except Exception as ex:
        _mostrar_aviso(
            revit,
            u"No se pudo abrir el formulario.",
            content=_as_unicode(ex),
        )
