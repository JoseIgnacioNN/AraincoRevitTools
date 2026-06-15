# -*- coding: utf-8 -*-
"""UI WPF — seleccionar / resaltar / eliminar corrida por Armadura_Conjunto_GUID."""

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
    BuiltInCategory,
    DetailCurve,
    ElementId,
    FamilyInstance,
    FilledRegion,
    TextNote,
    Transaction,
)
from Autodesk.Revit.DB.Structure import Rebar
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler, TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from lib.corrida_guid import (
    ARMADURA_CONJUNTO_GUID_PARAM,
    collect_corrida_por_conjunto_guid,
    get_armadura_conjunto_guid,
)
from revit_wpf_window_position import (
    bind_center_wpf_on_revit_monitor,
    position_wpf_window_center_on_monitor,
    revit_main_hwnd,
)
from ui.ok_cancel_dialog import show_ok_cancel_dialog

_DIALOG_TITLE = u"Arainco: Corrida GUID"
_SINGLETON_KEY = u"Arainco_CorridaGUID_UI"
_WINDOW_TITLE = u"Arainco: Corrida GUID"

_HELP_TEXT = (
    u"Cada ejecución de armado asigna un identificador único "
    u"(GUID) a las barras, empalmes y croquis de despiece en el "
    u"parámetro «Armadura_Conjunto_GUID».\n\n"
    u"1. Pulsa «Seleccionar referencia» y elige una barra, un detail de empalme "
    u"o un elemento del lienzo de despiece.\n"
    u"2. Se mostrará el GUID y cuántos elementos comparten ese valor.\n"
    u"3. Puedes resaltar la corrida en el modelo o eliminarla por completo.\n\n"
    u"Incluye barras estructurales (Rebar), detail items de empalme y croquis "
    u"(curvas de detalle, textos, lienzo masking). "
    u"Etiquetas y cotas asociadas se eliminan al borrar las barras."
)

_SUMMARY_IDLE = (
    u"Aún no hay corrida identificada.\n"
    u"Selecciona una barra, empalme o elemento del lienzo de despiece."
)

_PICK_PROMPT = (
    u"Selecciona una barra (Rebar), empalme o croquis de despiece con "
    u"«Armadura_Conjunto_GUID» (p. ej. creados por Armado Muros o Armado Columnas)."
)

XAML = u"""
<Window
    x:Name="CorridaGuidWin"
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="__WINDOW_TITLE__"
    Width="540"
    WindowStartupLocation="Manual"
    Background="Transparent"
    AllowsTransparency="True"
    FontFamily="Segoe UI"
    WindowStyle="None"
    ResizeMode="NoResize"
    SizeToContent="Height"
    ShowInTaskbar="False">
  <Window.Resources>
""" + BIMTOOLS_DARK_STYLES_XML + u"""
  </Window.Resources>
  <Border CornerRadius="8" Background="#071018" BorderBrush="#21465C"
          BorderThickness="1" Padding="22,20">
    <StackPanel>
      <TextBlock Text="__WINDOW_TITLE__" Foreground="#E8F4F8"
                 FontSize="16" FontWeight="Bold"/>
      <TextBlock Margin="0,6,0,0" Text="Gestión de corridas Armado Muros"
                 Foreground="#95B8CC" FontSize="11"/>
      <Border Margin="0,14,0,0" Background="#0a1620" BorderBrush="#21465C"
              BorderThickness="1" CornerRadius="6" Padding="12,10">
        <TextBlock Text="__HELP__" TextWrapping="Wrap"
                   Foreground="#95B8CC" FontSize="11" LineHeight="16"/>
      </Border>
      <Button x:Name="BtnPick" Margin="0,16,0,0"
              Content="Seleccionar referencia"
              Style="{StaticResource BtnPrimary}"
              HorizontalAlignment="Stretch" MinHeight="34"/>
      <Border Margin="0,14,0,0" Background="#0a1620" BorderBrush="#21465C"
              BorderThickness="1" CornerRadius="6" Padding="12,10">
        <TextBlock x:Name="TxtSummary" Text="__SUMMARY_IDLE__"
                   TextWrapping="Wrap" Foreground="#E8F4F8"
                   FontSize="12" LineHeight="18"/>
      </Border>
      <Button x:Name="BtnHighlight" Margin="0,14,0,0"
              Content="Resaltar en modelo"
              Style="{StaticResource BtnSelectOutline}"
              HorizontalAlignment="Stretch" MinHeight="34"
              IsEnabled="False"/>
      <Button x:Name="BtnDelete" Margin="0,8,0,0"
              Content="Eliminar Conjunto"
              Style="{StaticResource BtnSelectOutline}"
              HorizontalAlignment="Stretch" MinHeight="34"
              IsEnabled="False"/>
      <StackPanel Margin="0,16,0,0" Orientation="Horizontal"
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
    xaml = xaml.replace(u"__SUMMARY_IDLE__", _escape_xaml(_SUMMARY_IDLE))
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


class _FiltroCorridaReferencia(ISelectionFilter):
    def AllowElement(self, elem):
        if isinstance(elem, Rebar):
            return True
        if isinstance(elem, (DetailCurve, TextNote, FilledRegion)):
            return True
        if isinstance(elem, FamilyInstance):
            try:
                cat = elem.Category
                if cat is not None:
                    return int(cat.Id.IntegerValue) == int(
                        BuiltInCategory.OST_DetailComponents,
                    )
            except Exception:
                pass
        return False

    def AllowReference(self, reference, position):
        return False


def _guid_snippet(guid, max_len=40):
    if not guid:
        return u""
    s = _as_unicode(guid).strip()
    if len(s) <= max_len:
        return s
    return s[: max_len - 1] + u"…"


def _summarize_corrida(doc, conjunto_guid, corrida, ref_eid=None):
    rebar_ids = (corrida or {}).get(u"rebar_ids") or []
    empalme_ids = (corrida or {}).get(u"empalme_ids") or []
    lienzo_ids = (corrida or {}).get(u"lienzo_ids") or []
    gid = _guid_snippet(conjunto_guid, 72)
    lines = [
        u"GUID: {0}".format(gid),
        u"Barras estructurales: {0}".format(len(rebar_ids)),
        u"Representaciones de empalme: {0}".format(len(empalme_ids)),
        u"Croquis de despiece (lienzo): {0}".format(len(lienzo_ids)),
        u"Total elementos de corrida: {0}".format(
            len(rebar_ids) + len(empalme_ids) + len(lienzo_ids),
        ),
    ]
    if ref_eid is not None:
        try:
            lines.append(u"Referencia: Id {0}".format(int(ref_eid.IntegerValue)))
        except Exception:
            try:
                lines.append(u"Referencia: Id {0}".format(int(ref_eid)))
            except Exception:
                pass
    host_ids = set()
    if doc is not None:
        for eid in rebar_ids:
            try:
                rb = doc.GetElement(eid)
            except Exception:
                continue
            if rb is None:
                continue
            try:
                hid = rb.GetHostId()
                if hid is not None and hid != ElementId.InvalidElementId:
                    host_ids.add(int(hid.IntegerValue))
            except Exception:
                pass
    if host_ids:
        lines.append(u"Muros / hosts distintos (barras): {0}".format(len(host_ids)))
    lines.append(u"")
    lines.append(
        u"Se resaltarán o eliminarán barras, empalmes y croquis de despiece con el "
        u"mismo «{0}».".format(ARMADURA_CONJUNTO_GUID_PARAM),
    )
    return u"\n".join(lines)


class _PickRebarHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        win = self._window_ref()
        if win is None:
            return
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            win._set_status_error(u"No hay documento activo.")
            win._restore_window()
            return
        doc = uidoc.Document
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element,
                _FiltroCorridaReferencia(),
                _PICK_PROMPT,
            )
        except OperationCanceledException:
            win._restore_window()
            return
        except Exception as ex:
            win._set_status_error(u"Error al seleccionar:\n{0}".format(ex))
            win._restore_window()
            return

        if ref is None:
            win._restore_window()
            return

        elem = doc.GetElement(ref.ElementId)
        guid = get_armadura_conjunto_guid(elem)
        if not guid:
            win._clear_corrida()
            win._set_summary(
                u"El elemento elegido no tiene valor en «{0}».\n\n"
                u"Solo aplica a barras, empalmes o croquis de despiece con ese "
                u"parámetro vinculado (p. ej. creados por Armado Muros o Armado Columnas).".format(
                    ARMADURA_CONJUNTO_GUID_PARAM,
                ),
                enabled=False,
            )
            win._restore_window()
            return

        corrida = collect_corrida_por_conjunto_guid(doc, guid)
        all_ids = corrida.get(u"all_ids") or []
        if not all_ids:
            win._clear_corrida()
            win._set_summary(
                u"No se encontraron elementos con GUID «{0}».".format(
                    _guid_snippet(guid),
                ),
                enabled=False,
            )
            win._restore_window()
            return

        win._apply_corrida(doc, guid, corrida, ref.ElementId)
        win._restore_window()

    def GetName(self):
        return u"Arainco: Corrida GUID — seleccionar barra"


class _HighlightHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        win = self._window_ref()
        if win is None:
            return
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            return
        ids = win._corrida_ids
        if not ids:
            return
        sel = List[ElementId]()
        for eid in ids:
            sel.Add(eid)
        try:
            uidoc.Selection.SetElementIds(sel)
            uidoc.ShowElements(sel)
        except Exception:
            pass

    def GetName(self):
        return u"Arainco: Corrida GUID — resaltar"


class _DeleteHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        win = self._window_ref()
        if win is None:
            return
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            win._set_status_error(u"No hay documento activo.")
            return
        doc = uidoc.Document
        ids = list(win._corrida_ids or [])
        guid = win._conjunto_guid
        if not ids or not guid:
            return

        n_rebar = len(win._rebar_ids or [])
        n_emp = len(win._empalme_ids or [])
        n_lienzo = len(win._lienzo_ids or [])
        if not show_ok_cancel_dialog(
            _DIALOG_TITLE,
            u"¿Eliminar la corrida seleccionada?",
            u"Se borrarán {0} elemento(s) con GUID:\n{1}\n\n"
            u"  · Barras: {2}\n"
            u"  · Empalmes (detail): {3}\n"
            u"  · Croquis de despiece: {4}\n\n"
            u"Esta acción no se puede deshacer fuera de Revit "
            u"(usa Deshacer inmediatamente si te equivocas).".format(
                len(ids), _guid_snippet(guid, 72), n_rebar, n_emp, n_lienzo,
            ),
            ok_text=u"Eliminar",
            cancel_text=u"Cancelar",
            uiapp=uiapp,
        ):
            return

        deleted = 0
        t = Transaction(doc, u"Arainco: Eliminar corrida GUID")
        try:
            t.Start()
            for eid in ids:
                try:
                    doc.Delete(eid)
                    deleted += 1
                except Exception:
                    pass
            t.Commit()
        except Exception as ex:
            if t.HasStarted():
                try:
                    t.RollBack()
                except Exception:
                    pass
            win._set_status_error(u"Error al eliminar:\n{0}".format(ex))
            return

        win._clear_corrida()
        win._set_summary(
            u"Corrida eliminada: {0} elemento(s) borrados "
            u"({1} barras, {2} empalmes, {3} croquis).\n"
            u"GUID: {4}".format(
                deleted, n_rebar, n_emp, n_lienzo, _guid_snippet(guid, 72),
            ),
            enabled=False,
        )
        try:
            uidoc.Selection.SetElementIds(List[ElementId]())
        except Exception:
            pass

    def GetName(self):
        return u"Arainco: Corrida GUID — eliminar"


class CorridaGuidWindow(object):
    def __init__(self, uiapp):
        self._uiapp = uiapp
        self._conjunto_guid = None
        self._rebar_ids = []
        self._empalme_ids = []
        self._lienzo_ids = []
        self._corrida_ids = []
        self._ref_eid = None
        self._win = XamlReader.Parse(_build_xaml())
        _prepare_window(self._win, uiapp)

        self._pick_handler = _PickRebarHandler(weakref.ref(self))
        self._highlight_handler = _HighlightHandler(weakref.ref(self))
        self._delete_handler = _DeleteHandler(weakref.ref(self))
        self._pick_event = ExternalEvent.Create(self._pick_handler)
        self._highlight_event = ExternalEvent.Create(self._highlight_handler)
        self._delete_event = ExternalEvent.Create(self._delete_handler)

        self._wire_commands()
        self._wire_close()

    def _find(self, name):
        try:
            return self._win.FindName(name)
        except Exception:
            return None

    def _wire_commands(self):
        btn_pick = self._find(u"BtnPick")
        btn_highlight = self._find(u"BtnHighlight")
        btn_delete = self._find(u"BtnDelete")
        if btn_pick is not None:
            btn_pick.Click += RoutedEventHandler(self._on_pick)
        if btn_highlight is not None:
            btn_highlight.Click += RoutedEventHandler(self._on_highlight)
        if btn_delete is not None:
            btn_delete.Click += RoutedEventHandler(self._on_delete)

    def _wire_close(self):
        btn_close = self._find(u"BtnClose")
        if btn_close is not None:
            btn_close.Click += RoutedEventHandler(self._on_close)

        def _on_key(sender, args):
            if args.Key == Key.Escape:
                self._on_close(sender, args)
                args.Handled = True

        self._win.PreviewKeyDown += KeyEventHandler(_on_key)
        self._win.Closed += EventHandler(self._on_closed)

    def _set_action_buttons_enabled(self, enabled):
        for name in (u"BtnHighlight", u"BtnDelete"):
            btn = self._find(name)
            if btn is not None:
                btn.IsEnabled = bool(enabled)

    def _set_summary(self, text, enabled=None):
        txt = self._find(u"TxtSummary")
        if txt is not None:
            txt.Text = _as_unicode(text)
        if enabled is not None:
            self._set_action_buttons_enabled(enabled)

    def _set_status_error(self, message):
        TaskDialog.Show(_DIALOG_TITLE, _as_unicode(message))

    def _clear_corrida(self):
        self._conjunto_guid = None
        self._rebar_ids = []
        self._empalme_ids = []
        self._lienzo_ids = []
        self._corrida_ids = []
        self._ref_eid = None
        self._set_action_buttons_enabled(False)

    def _apply_corrida(self, doc, guid, corrida, ref_eid):
        self._conjunto_guid = guid
        self._rebar_ids = list((corrida or {}).get(u"rebar_ids") or [])
        self._empalme_ids = list((corrida or {}).get(u"empalme_ids") or [])
        self._lienzo_ids = list((corrida or {}).get(u"lienzo_ids") or [])
        self._corrida_ids = list((corrida or {}).get(u"all_ids") or [])
        self._ref_eid = ref_eid
        self._set_summary(
            _summarize_corrida(
                doc,
                guid,
                {
                    u"rebar_ids": self._rebar_ids,
                    u"empalme_ids": self._empalme_ids,
                    u"lienzo_ids": self._lienzo_ids,
                },
                ref_eid,
            ),
            enabled=True,
        )

    def _restore_window(self):
        try:
            self._win.Show()
            self._win.Activate()
            if self._win.WindowState == WindowState.Minimized:
                self._win.WindowState = WindowState.Normal
        except Exception:
            pass

    def _on_pick(self, sender, args):
        try:
            self._win.Hide()
        except Exception:
            pass
        self._pick_event.Raise()

    def _on_highlight(self, sender, args):
        if not self._corrida_ids:
            return
        self._highlight_event.Raise()

    def _on_delete(self, sender, args):
        if not self._corrida_ids:
            return
        self._delete_event.Raise()

    def _on_close(self, sender, args):
        try:
            self._win.Close()
        except Exception:
            pass

    def _on_closed(self, sender, args):
        _unregister_singleton()

    def show(self):
        self._win.Show()


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
        if w is None or not bool(getattr(w, "IsLoaded", False)):
            raise RuntimeError(u"ventana cerrada")
        w.Activate()
        if w.WindowState == WindowState.Minimized:
            w.WindowState = WindowState.Normal
        TaskDialog.Show(_WINDOW_TITLE, u"La herramienta ya está en ejecución.")
        return True
    except Exception:
        _unregister_singleton()
        return False


def show_corrida_guid_ui(uiapp):
    if uiapp is None:
        TaskDialog.Show(_DIALOG_TITLE, u"No hay aplicación Revit activa.")
        return
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        TaskDialog.Show(_DIALOG_TITLE, u"No hay documento activo.")
        return
    if _try_activate_existing():
        return
    win = CorridaGuidWindow(uiapp)
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, win)
    except Exception:
        pass
    win.show()
