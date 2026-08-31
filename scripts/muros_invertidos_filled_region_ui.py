# -*- coding: utf-8 -*-
"""
UI WPF — Arainco: Muros invertidos.

Ceiling Plan (ViewDirection +Z). Elige tipo de Filled Region y genera
las regiones del desacople (sin corte − cortes coincidentes).

Revit 2024+ | IronPython (pyRevit).
Entrada: ``32_MurosInvertidosFilledRegion.pushbutton/script.py``.
"""

from __future__ import print_function

try:
    unicode
except NameError:
    unicode = str

import os
import traceback

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from System import AppDomain, EventHandler
from System.Windows import RoutedEventHandler, WindowState
from System.Windows.Controls import ComboBoxItem
from System.Windows.Input import Key, KeyEventHandler
from System.Windows.Markup import XamlReader
from Autodesk.Revit.UI import TaskDialog

from bimtools_ui_tokens import BTN_MANUAL
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_wpf_shell import build_simple_tool_xaml
from revit_wpf_window_position import (
    bind_center_wpf_on_revit_monitor,
    position_wpf_window_center_on_monitor,
    revit_main_hwnd,
)
from muros_invertidos_filled_region import (
    classify_walls,
    generate,
    is_ceiling_plan_looking_up,
    list_filled_region_types,
)

_TOOL_TITLE = u"Arainco: Muros invertidos"
_SINGLETON_KEY = u"Arainco_MurosInvertidosFilledRegion_UI"
_SETTINGS_KEY = u"Arainco_MurosInvertidosFilledRegion_Settings"
_ALREADY_RUNNING = u"La herramienta ya esta en ejecucion."

_SUBTITLE = (
    u"Planta de techo con mirada hacia arriba. "
    u"Incluye muros del recorte aunque estén fuera del View Range "
    u"(invertidos sobre la losa)."
)

_BODY_XAML = u"""
<StackPanel>
  <TextBlock x:Name="TxtView" Style="{StaticResource LabelSmall}"
             TextWrapping="Wrap" Margin="0,0,0,10"/>
  <TextBlock x:Name="TxtLots" Style="{StaticResource LabelSmall}"
             TextWrapping="Wrap" Margin="0,0,0,12"/>
  <TextBlock Style="{StaticResource Label}" Text="Tipo de región rellena"/>
  <ComboBox x:Name="CboType" Style="{StaticResource ComboStretch}"
            MinHeight="32"/>
</StackPanel>
"""

_FOOTER_LEADING_XAML = (
    u'<Button x:Name="BtnManual" Content="Manual" '
    u'Style="{{StaticResource BtnSelectOutline}}" '
    u'Background="{bg}" MinWidth="96" MinHeight="32" MaxHeight="36" '
    u'Padding="8,2" VerticalAlignment="Center" '
    u'ToolTip="Abrir manual de usuario"/>'
).format(bg=BTN_MANUAL)

_FOOTER_ACTIONS_XAML = u"""
<Button x:Name="BtnGenerate" Content="Generar"
        Style="{StaticResource BtnPrimary}" MinWidth="130"
        Margin="0,0,10,0" IsDefault="True"/>
<Button x:Name="BtnClose" Content="Cerrar"
        Style="{StaticResource BtnSelectOutline}" MinWidth="110"/>
"""


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _elem_name(el):
    if el is None:
        return u""
    try:
        n = el.Name
        if n:
            return _as_unicode(n)
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import Element

        return _as_unicode(Element.Name.__get__(el))
    except Exception:
        return u""


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
        TaskDialog.Show(_TOOL_TITLE, body)
    except Exception:
        pass


def _resolve_manual_path():
    try:
        import bimtools_paths

        pb = bimtools_paths.get_pushbutton_dir()
        if pb:
            path = os.path.join(pb, u"manual_usuario.html")
            if os.path.isfile(path):
                return path
    except Exception:
        pass
    here = os.path.dirname(os.path.abspath(__file__))
    ext = os.path.dirname(here)
    try:
        for tab_name in os.listdir(ext):
            if not tab_name.endswith(u".tab"):
                continue
            panel = os.path.join(ext, tab_name, u"Modelado.panel")
            if not os.path.isdir(panel):
                continue
            for pb_name in os.listdir(panel):
                if u"MurosInvertidos" not in pb_name:
                    continue
                path = os.path.join(panel, pb_name, u"manual_usuario.html")
                if os.path.isfile(path):
                    return path
    except Exception:
        pass
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
        _mostrar_aviso(uiapp, u"No se pudo abrir el manual.", content=_as_unicode(ex))


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
        title=_TOOL_TITLE,
        styles_xml=BIMTOOLS_DARK_STYLES_XML,
        body_xaml=_BODY_XAML,
        footer_leading_xaml=_FOOTER_LEADING_XAML,
        footer_actions_xaml=_FOOTER_ACTIONS_XAML,
        width=520,
        min_width=480,
        resize_mode=u"NoResize",
        size_to_content_height=True,
    )


def _load_settings():
    try:
        data = AppDomain.CurrentDomain.GetData(_SETTINGS_KEY)
        if isinstance(data, dict):
            return data
    except Exception:
        pass
    return {}


def _save_settings(data):
    try:
        AppDomain.CurrentDomain.SetData(_SETTINGS_KEY, data)
    except Exception:
        pass


def _window_is_alive(win):
    try:
        return win is not None and win.IsLoaded
    except Exception:
        return False


def _existing_window():
    try:
        w = AppDomain.CurrentDomain.GetData(_SINGLETON_KEY)
    except Exception:
        w = None
    if _window_is_alive(w):
        return w
    try:
        from System.Windows import Application

        app = Application.Current
        if app is None:
            return None
        for ww in app.Windows:
            try:
                txt = ww.FindName(u"TxtTitle")
                if txt is not None and _as_unicode(txt.Text) == _TOOL_TITLE:
                    if _window_is_alive(ww):
                        return ww
            except Exception:
                continue
    except Exception:
        pass
    return None


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


def _combo_select_by_id(combo, id_int):
    if combo is None or id_int is None:
        return
    for item in combo.Items:
        try:
            if int(item.Tag) == int(id_int):
                combo.SelectedItem = item
                return
        except Exception:
            continue
    if combo.Items.Count > 0:
        combo.SelectedIndex = 0


class MurosInvertidosWindow(object):
    def __init__(self, uiapp, uidoc, view, types, cut_n, nocut_n, skipped, zcut):
        self._uiapp = uiapp
        self._uidoc = uidoc
        self._view = view
        self._types = types
        self._win = XamlReader.Parse(_build_xaml())
        self._txt_subtitle = self._win.FindName(u"TxtSubtitle")
        self._txt_status = self._win.FindName(u"TxtStatus")
        self._txt_view = self._win.FindName(u"TxtView")
        self._txt_lots = self._win.FindName(u"TxtLots")
        self._cbo = self._win.FindName(u"CboType")

        if self._txt_subtitle is not None:
            self._txt_subtitle.Text = _SUBTITLE
        try:
            vname = _as_unicode(view.Name)
        except Exception:
            vname = u"(vista activa)"
        ztxt = u"—"
        if zcut is not None:
            ztxt = u"{0:.3f} m".format(float(zcut) * 0.3048)
        if self._txt_view is not None:
            self._txt_view.Text = u"Vista: {0}  ·  Cut Plane ≈ {1}".format(vname, ztxt)
        if self._txt_lots is not None:
            self._txt_lots.Text = (
                u"Muros en corte: {0}   ·   sin corte: {1}   ·   sin geometría: {2}"
            ).format(cut_n, nocut_n, skipped)

        self._fill_types()
        _combo_select_by_id(self._cbo, _load_settings().get(u"type_id"))
        self._win.FindName(u"BtnGenerate").Click += RoutedEventHandler(
            self._on_generate
        )
        self._win.FindName(u"BtnClose").Click += RoutedEventHandler(self._on_close)
        manual = self._win.FindName(u"BtnManual")
        if manual is not None:
            manual.Click += RoutedEventHandler(self._on_manual)
        self._win.KeyDown += KeyEventHandler(self._on_key_down)
        self._win.Closed += EventHandler(self._on_closed)
        _prepare_window(self._win, uiapp)
        self._set_status(u"Elija el tipo y pulse Generar.")

    def _fill_types(self):
        if self._cbo is None:
            return
        self._cbo.Items.Clear()
        for t in self._types:
            item = ComboBoxItem()
            item.Content = _elem_name(t) or u"(sin nombre)"
            try:
                item.Tag = int(t.Id.IntegerValue)
            except Exception:
                try:
                    item.Tag = int(t.Id.Value)
                except Exception:
                    item.Tag = -1
            self._cbo.Items.Add(item)
        if self._cbo.Items.Count > 0:
            self._cbo.SelectedIndex = 0

    def _selected_type_id(self):
        if self._cbo is None or self._cbo.SelectedItem is None:
            return None
        try:
            from Autodesk.Revit.DB import ElementId

            return ElementId(int(self._cbo.SelectedItem.Tag))
        except Exception:
            return None

    def _set_status(self, text):
        if self._txt_status is not None:
            self._txt_status.Text = _as_unicode(text)

    def _on_manual(self, sender, args):
        _open_manual(self._uiapp)

    def _on_key_down(self, sender, args):
        if args.Key == Key.Escape:
            self._win.Close()

    def _on_close(self, sender, args):
        self._win.Close()

    def _on_closed(self, sender, args):
        try:
            AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, None)
        except Exception:
            pass

    def _on_generate(self, sender, args):
        type_id = self._selected_type_id()
        if type_id is None:
            _mostrar_aviso(self._uiapp, u"Seleccione un tipo de región rellena.")
            return
        try:
            _save_settings({u"type_id": int(self._cbo.SelectedItem.Tag)})
        except Exception:
            pass
        uidoc = self._uidoc
        view = None
        doc = None
        try:
            view = uidoc.ActiveView
            doc = uidoc.Document
        except Exception:
            view = self._view
            try:
                doc = view.Document
            except Exception:
                doc = None
        if doc is None or view is None:
            _mostrar_aviso(self._uiapp, u"No hay documento o vista activa.")
            return
        if not is_ceiling_plan_looking_up(view):
            _mostrar_aviso(
                self._uiapp,
                u"La vista activa debe ser una planta de techo con mirada hacia arriba.",
            )
            return
        try:
            result = generate(doc, view, type_id)
        except Exception as ex:
            _mostrar_aviso(
                self._uiapp,
                u"No se pudieron crear las regiones.",
                content=_as_unicode(ex) + u"\n\n" + _as_unicode(traceback.format_exc()),
            )
            return
        created = int(result.get(u"created") or 0)
        failed = int(result.get(u"failed") or 0)
        groups = int(result.get(u"groups") or 0)
        if created == 0 and groups == 0:
            _mostrar_aviso(
                self._uiapp,
                u"No hay desacople que representar.",
                content=(
                    u"Muros en corte: {0}\nMuros sin corte: {1}\n"
                    u"Las huellas coincidentes no dejan resto, o no hay muros sin corte."
                ).format(result.get(u"cut"), result.get(u"nocut")),
            )
            self._set_status(u"Sin regiones: huellas coincidentes o lote vacío.")
            return
        msg = u"Creadas {0} región(es).".format(created)
        if failed:
            msg = msg + u" {0} contorno(s) no válidos.".format(failed)
        self._set_status(msg)
        _mostrar_aviso(self._uiapp, msg)

    def show(self):
        try:
            AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, self._win)
        except Exception:
            pass
        self._win.ShowDialog()


def run(revit):
    uiapp = revit
    uidoc = None
    try:
        uidoc = uiapp.ActiveUIDocument
    except Exception:
        uidoc = None
    if uidoc is None:
        _mostrar_aviso(uiapp, u"No hay documento activo.")
        return
    view = None
    try:
        view = uidoc.ActiveView
    except Exception:
        view = None
    if view is None:
        _mostrar_aviso(uiapp, u"No hay vista activa.")
        return
    if not is_ceiling_plan_looking_up(view):
        _mostrar_aviso(
            uiapp,
            u"Esta herramienta solo opera en plantas de techo con mirada hacia arriba (View Direction Up).",
            content=u"Cambie a un Ceiling Plan que mire hacia +Z e inténtelo de nuevo.",
        )
        return
    doc = uidoc.Document
    types = list_filled_region_types(doc)
    if not types:
        _mostrar_aviso(
            uiapp,
            u"El documento no tiene tipos de región rellena (Filled Region).",
        )
        return
    existing = _existing_window()
    if existing is not None:
        _activate_existing(existing, uiapp)
        return
    cut_items, nocut_items, skipped, zcut = classify_walls(doc, view)
    win = MurosInvertidosWindow(
        uiapp,
        uidoc,
        view,
        types,
        len(cut_items),
        len(nocut_items),
        skipped,
        zcut,
    )
    win.show()
