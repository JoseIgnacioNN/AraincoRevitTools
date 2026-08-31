# -*- coding: utf-8 -*-
"""
UI WPF — Arainco: Unobscured mallas (Armadura_Malla = Yes).

Aplica o quita View Unobscured (+ sólido en vista) a las barras de la vista
activa cuyo parámetro booleano ``Armadura_Malla`` está en Yes.

Revit 2024+ | pyRevit / IronPython.
Entrada: ``06_MallaUnobscuredVista.pushbutton/script.py``.
"""

from __future__ import print_function

import os
import weakref

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from System import AppDomain, EventHandler
from System.Collections.Generic import List
from System.Windows import RoutedEventHandler, WindowState
from System.Windows.Input import Key, KeyEventHandler
from System.Windows.Markup import XamlReader
from Autodesk.Revit.DB import (
    ElementId,
    FilteredElementCollector,
    StorageType,
    Transaction,
    View,
    ViewSchedule,
    ViewSheet,
)
from Autodesk.Revit.DB.Structure import Rebar, RebarInSystem
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler, TaskDialog

from bimtools_rebar_3d_visibility import (
    apply_reinforcement_unobscured_in_view,
    summarize_reinforcement_unobscured_in_view,
)
from bimtools_ui_tokens import BTN_MANUAL, FG_BODY, FONT_SIZE_BODY
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_wpf_shell import build_simple_tool_xaml
from revit_wpf_window_position import (
    bind_center_wpf_on_revit_monitor,
    position_wpf_window_center_on_monitor,
    revit_main_hwnd,
)

_TOOL_TITLE = u"Arainco: Unobscured mallas"
_SINGLETON_KEY = u"Arainco_MallaUnobscuredVista_UI"
_PARAM_MALLA = u"Armadura_Malla"
_TX_APPLY = u"Arainco: View Unobscured mallas (aplicar)"
_TX_REMOVE = u"Arainco: View Unobscured mallas (quitar)"

_TEXT_INTRO = (
    u"Gestiona View Unobscured de las barras de malla en la vista activa "
    u"(parámetro Armadura_Malla = Yes). «Aplicar» las dibuja por delante "
    u"del hormigón; «Quitar» restaura el obscurecimiento. Solo afecta a "
    u"esta vista."
)

_BODY_XAML = u"""
<StackPanel>
  <TextBlock TextWrapping="Wrap" Foreground="{fg}" FontSize="{fs}" LineHeight="17"
             Text="{intro}"/>
  <Border Margin="0,12,0,0" Background="#071018" BorderBrush="#21465C"
          BorderThickness="1" CornerRadius="4" Padding="12,10">
    <TextBlock x:Name="TxtSummary" TextWrapping="Wrap"
               Foreground="#E8F4F8" FontSize="12" LineHeight="18"/>
  </Border>
  <StackPanel Orientation="Horizontal" Margin="0,12,0,0"
              HorizontalAlignment="Right">
    <Button x:Name="BtnRefresh" Content="Actualizar" Margin="0,0,8,0"
            Style="{{StaticResource BtnSelectOutline}}"
            MinWidth="104" MinHeight="32" MaxHeight="36"
            VerticalAlignment="Center"
            ToolTip="Volver a contar barras de malla en la vista activa"/>
    <Button x:Name="BtnRemove" Content="Quitar Unobscured" Margin="0,0,8,0"
            Style="{{StaticResource BtnSelectOutline}}"
            MinWidth="148" MinHeight="32" MaxHeight="36"
            VerticalAlignment="Center"/>
    <Button x:Name="BtnApply" Content="Aplicar Unobscured"
            Style="{{StaticResource BtnPrimary}}"
            MinWidth="156" MinHeight="32" MaxHeight="36"
            VerticalAlignment="Center"/>
  </StackPanel>
</StackPanel>
""".format(
    fg=FG_BODY,
    fs=FONT_SIZE_BODY,
    intro=_TEXT_INTRO.replace(u'"', u"&quot;"),
)

_FOOTER_LEADING_XAML = (
    u'<Button x:Name="BtnManual" Content="Manual" '
    u'Style="{{StaticResource BtnSelectOutline}}" '
    u'Background="{bg}" MinWidth="96" MinHeight="32" MaxHeight="36" '
    u'Padding="8,2" VerticalAlignment="Center" '
    u'ToolTip="Abrir manual de usuario"/>'
).format(bg=BTN_MANUAL)

_FOOTER_ACTIONS_XAML = (
    u'<Button x:Name="BtnClose" Content="Cerrar" '
    u'Style="{StaticResource BtnSelectOutline}" '
    u'MinWidth="108" MinHeight="32" MaxHeight="36" '
    u'VerticalAlignment="Center"/>'
)


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


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
    candidates = []
    try:
        import bimtools_paths

        pb = bimtools_paths.get_pushbutton_dir()
        if pb:
            candidates.append(os.path.join(pb, u"manual_usuario.html"))
    except Exception:
        pass
    for path in candidates:
        if path and os.path.isfile(path):
            return path
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
        _mostrar_aviso(
            uiapp,
            u"No se pudo abrir el manual.",
            content=_as_unicode(ex),
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


def _build_xaml():
    return build_simple_tool_xaml(
        title=_TOOL_TITLE,
        styles_xml=BIMTOOLS_DARK_STYLES_XML,
        body_xaml=_BODY_XAML,
        footer_leading_xaml=_FOOTER_LEADING_XAML,
        footer_actions_xaml=_FOOTER_ACTIONS_XAML,
        width=560,
        min_width=560,
        resize_mode=u"NoResize",
        size_to_content_height=True,
    )

def resolve_active_model_view(uidoc):
    """Vista gráfica de modelo activa (no plantilla, lámina ni cuadro)."""
    if uidoc is None:
        return None
    doc = uidoc.Document
    if doc is None:
        return None

    view = None
    try:
        view = getattr(uidoc, "ActiveGraphicalView", None)
    except Exception:
        view = None
    if view is None:
        try:
            view = uidoc.ActiveView
        except Exception:
            view = None
    if view is None:
        return None

    try:
        resolved = doc.GetElement(view.Id)
    except Exception:
        resolved = None
    if isinstance(resolved, View):
        view = resolved

    if not isinstance(view, View):
        return None
    try:
        if view.IsTemplate:
            return None
    except Exception:
        pass
    if isinstance(view, (ViewSheet, ViewSchedule)):
        return None
    return view


def _lookup_param(el, name):
    if el is None:
        return None
    try:
        p = el.LookupParameter(name)
        if p is not None:
            return p
    except Exception:
        pass
    try:
        return el.GetParameters(name)[0]
    except Exception:
        return None


def is_armadura_malla_yes(el):
    """True si ``Armadura_Malla`` está en Yes (booleano 1 o cadena Yes/Sí)."""
    p = _lookup_param(el, _PARAM_MALLA)
    if p is None:
        return False
    try:
        if not p.HasValue:
            return False
    except Exception:
        pass
    try:
        st = p.StorageType
    except Exception:
        st = None
    try:
        if st == StorageType.Integer:
            return int(p.AsInteger()) != 0
    except Exception:
        pass
    try:
        s = (p.AsValueString() or p.AsString() or u"").strip().lower()
    except Exception:
        s = u""
    if not s:
        return False
    return s in (u"yes", u"sí", u"si", u"1", u"true", u"verdadero")


def collect_malla_rebars_in_view(doc, view):
    """
    Rebar y RebarInSystem visibles en ``view`` con ``Armadura_Malla`` = Yes.
    """
    if doc is None or view is None or not isinstance(view, View):
        return []
    try:
        if view.IsTemplate:
            return []
    except Exception:
        pass

    out = []
    seen = set()
    view_id = view.Id
    for cls in (Rebar, RebarInSystem):
        try:
            elems = (
                FilteredElementCollector(doc, view_id)
                .OfClass(cls)
                .WhereElementIsNotElementType()
                .ToElements()
            )
        except Exception:
            elems = []
        for el in elems or []:
            if el is None or not is_armadura_malla_yes(el):
                continue
            try:
                key = int(el.Id.IntegerValue)
            except AttributeError:
                try:
                    key = int(el.Id.Value)
                except Exception:
                    continue
            except Exception:
                continue
            if key in seen:
                continue
            seen.add(key)
            out.append(el)
    return out


def _unhide_elements_in_view(view, elements):
    """Unhide en vista (ignora fallos); no es obligatorio para Unobscured."""
    if view is None or not elements:
        return 0
    ids = List[ElementId]()
    for el in elements:
        try:
            ids.Add(el.Id)
        except Exception:
            pass
    if ids.Count < 1:
        return 0
    try:
        view.UnhideElements(ids)
        return int(ids.Count)
    except Exception:
        return 0


def apply_malla_unobscured(doc, view, elements, unobscured):
    """
    Aplica o quita View Unobscured (+ sólido) a ``elements`` en ``view``.
    Devuelve el número de elementos procesados con éxito.
    """
    if not elements:
        return 0
    if unobscured:
        _unhide_elements_in_view(view, elements)
    return int(
        apply_reinforcement_unobscured_in_view(
            doc, elements, view, unobscured=bool(unobscured)
        )
        or 0
    )


class _UnobscuredHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref
        self.mode = None  # u"apply" | u"remove"

    def Execute(self, uiapp):
        win = self._window_ref()
        if win is None:
            return
        mode = self.mode
        self.mode = None
        if mode not in (u"apply", u"remove"):
            return

        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            win._set_status(u"No hay documento activo.")
            _mostrar_aviso(uiapp, u"No hay documento activo.")
            return

        doc = uidoc.Document
        view = resolve_active_model_view(uidoc)
        if view is None:
            win._set_status(u"Activa una vista de modelo (planta, alzado, sección o 3D).")
            _mostrar_aviso(
                uiapp,
                u"No hay una vista de modelo activa.",
                content=u"Activa una planta, alzado, sección o 3D e inténtalo de nuevo.",
            )
            return

        rebars = collect_malla_rebars_in_view(doc, view)
        if not rebars:
            win._refresh_summary()
            win._set_status(u"No hay barras con Armadura_Malla = Yes en la vista.")
            _mostrar_aviso(
                uiapp,
                u"No hay barras de malla (Armadura_Malla = Yes) en la vista activa.",
                content=u"Vista: {0}".format(
                    _as_unicode(getattr(view, "Name", None) or u"(vista)")
                ),
            )
            return

        unobscured = mode == u"apply"
        tx_name = _TX_APPLY if unobscured else _TX_REMOVE
        n_ok = 0
        t = Transaction(doc, tx_name)
        try:
            t.Start()
            n_ok = apply_malla_unobscured(doc, view, rebars, unobscured)
            t.Commit()
        except Exception as ex:
            try:
                t.RollBack()
            except Exception:
                pass
            win._set_status(u"Error: {0}".format(_as_unicode(ex)))
            _mostrar_aviso(
                uiapp,
                u"No se pudo modificar View Unobscured.",
                content=_as_unicode(ex),
            )
            return

        win._refresh_summary()
        if unobscured:
            msg = u"Unobscured aplicado a {0} barra(s) de malla.".format(n_ok)
        else:
            msg = u"Unobscured quitado de {0} barra(s) de malla.".format(n_ok)
        win._set_status(msg)

    def GetName(self):
        if self.mode == u"remove":
            return _TX_REMOVE
        return _TX_APPLY


class MallaUnobscuredVistaWindow(object):
    def __init__(self, uiapp):
        self._uiapp = uiapp
        self._win = XamlReader.Parse(_build_xaml())
        self._txt_subtitle = self._win.FindName(u"TxtSubtitle")
        self._txt_summary = self._win.FindName(u"TxtSummary")
        self._txt_status = self._win.FindName(u"TxtStatus")

        self._handler = _UnobscuredHandler(weakref.ref(self))
        self._ext_event = ExternalEvent.Create(self._handler)

        self._wire_events()
        _prepare_window(self._win, uiapp)
        self._refresh_summary()

    def _wire_events(self):
        self._win.FindName(u"BtnApply").Click += RoutedEventHandler(self._on_apply)
        self._win.FindName(u"BtnRemove").Click += RoutedEventHandler(self._on_remove)
        self._win.FindName(u"BtnRefresh").Click += RoutedEventHandler(self._on_refresh)
        self._win.FindName(u"BtnClose").Click += RoutedEventHandler(self._on_close)
        manual = self._win.FindName(u"BtnManual")
        if manual is not None:
            manual.Click += RoutedEventHandler(self._on_manual)
        self._win.KeyDown += KeyEventHandler(self._on_key_down)
        self._win.Closed += EventHandler(self._on_closed)

    def _on_manual(self, sender, args):
        _open_manual(self._uiapp)

    def _on_key_down(self, sender, args):
        if args.Key == Key.Escape:
            self._win.Close()

    def _on_refresh(self, sender, args):
        self._refresh_summary()
        self._set_status(u"Resumen actualizado.")

    def _on_apply(self, sender, args):
        self._set_status(u"Aplicando Unobscured…")
        self._handler.mode = u"apply"
        self._ext_event.Raise()

    def _on_remove(self, sender, args):
        self._set_status(u"Quitando Unobscured…")
        self._handler.mode = u"remove"
        self._ext_event.Raise()
    def _on_close(self, sender, args):
        self._win.Close()

    def _on_closed(self, sender, args):
        try:
            AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, None)
        except Exception:
            pass

    def _set_status(self, text):
        if self._txt_status is not None:
            self._txt_status.Text = _as_unicode(text)

    def _refresh_summary(self):
        uidoc = self._uiapp.ActiveUIDocument
        if uidoc is None:
            if self._txt_subtitle is not None:
                self._txt_subtitle.Text = u"No hay documento activo."
            if self._txt_summary is not None:
                self._txt_summary.Text = u"Abre un proyecto y una vista de modelo."
            return

        view = resolve_active_model_view(uidoc)
        if view is None:
            if self._txt_subtitle is not None:
                self._txt_subtitle.Text = (
                    u"Activa una vista de modelo (planta, alzado, sección o 3D)."
                )
            if self._txt_summary is not None:
                self._txt_summary.Text = (
                    u"No se puede operar sobre láminas, cuadros ni plantillas de vista."
                )
            return

        doc = uidoc.Document
        rebars = collect_malla_rebars_in_view(doc, view)
        stats = summarize_reinforcement_unobscured_in_view(doc, rebars, view)
        try:
            vname = _as_unicode(view.Name)
        except Exception:
            vname = u"Vista"

        if self._txt_subtitle is not None:
            self._txt_subtitle.Text = (
                u"Vista: {0} · {1} barra(s) con Armadura_Malla = Yes.".format(
                    vname, len(rebars)
                )
            )
        if self._txt_summary is not None:
            self._txt_summary.Text = (
                u"En la vista activa:\n"
                u"• Barras de malla: {0}\n"
                u"• Con Unobscured activo: {1}\n"
                u"• Sin Unobscured: {2}".format(
                    stats.get(u"total", 0),
                    stats.get(u"unobscured", 0),
                    stats.get(u"obscured", 0),
                )
            )

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


def _focus_existing(win, uiapp):
    try:
        if win._win.WindowState == WindowState.Minimized:
            win._win.WindowState = WindowState.Normal
        win._win.Activate()
    except Exception:
        pass
    _mostrar_aviso(uiapp, u"La herramienta ya esta en ejecucion.")


def show_malla_unobscured_window(revit):
    uiapp = revit
    existing = _existing_window()
    if existing is not None:
        _focus_existing(existing, uiapp)
        return

    win = MallaUnobscuredVistaWindow(uiapp)
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, win)
    except Exception:
        pass
    win.show()


def run(revit):
    show_malla_unobscured_window(revit)
