# -*- coding: utf-8 -*-
"""
UI WPF — Arainco: Copiar desde vínculo.

Copia al proyecto host elementos del modelo vinculado seleccionado,
conservando la posición mediante la transformación del vínculo.

Revit 2024+ | pyRevit / IronPython.
Entrada: ``30_CopiarDesdeVinculo.pushbutton/script.py``.
"""

from __future__ import print_function

try:
    unicode
except NameError:
    unicode = str

import os
import traceback
import weakref

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from System import AppDomain, EventHandler
from System.Windows import (
    FontWeights,
    GridLength,
    GridUnitType,
    HorizontalAlignment,
    RoutedEventHandler,
    TextWrapping,
    Thickness,
    VerticalAlignment,
    Visibility,
    WindowState,
)
from System.Windows.Controls import (
    CheckBox,
    ColumnDefinition,
    ComboBoxItem,
    Grid,
    SelectionChangedEventHandler,
    TextBlock,
    TextChangedEventHandler,
)
from System.Windows.Input import Cursors, Key, KeyEventHandler
from System.Windows.Markup import XamlReader
from System.Windows.Media import Color, SolidColorBrush
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler, TaskDialog

from bimtools_ui_tokens import BTN_MANUAL, FG_BODY, FG_MUTED, FG_TITLE, FONT_SIZE_BODY
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_wpf_shell import build_simple_tool_xaml
from copiar_desde_vinculo_service import (
    collect_link_categories,
    list_loaded_revit_links,
    run_copy_in_transaction,
    summarize_category_groups,
)
from revit_wpf_window_position import (
    bind_center_wpf_on_revit_monitor,
    position_wpf_window_center_on_monitor,
    revit_main_hwnd,
)

_TOOL_TITLE = u"Arainco: Copiar desde vínculo"
_SINGLETON_KEY = u"Arainco_CopiarDesdeVinculo_UI"
_ALREADY_RUNNING = u"La herramienta ya esta en ejecucion."
_TX_COPY = u"Arainco: Copiar desde vínculo"

_SUBTITLE = (
    u"Copia columnas, framing, muros, losas, grids, niveles y fundaciones "
    u"del vínculo al host, conservando su posición."
)

_BODY_XAML = u"""
<StackPanel>
  <TextBlock Style="{StaticResource Label}" Text="Vínculo Revit" Margin="0,0,0,4"/>
  <ComboBox x:Name="CboLink" Style="{StaticResource Combo}" MinHeight="32"
            Margin="0,0,0,8"/>

  <TextBlock x:Name="TxtSummary" TextWrapping="Wrap" Margin="0,0,0,14"
             Foreground="#95B8CC" FontSize="11" LineHeight="16"/>

  <DockPanel LastChildFill="True" Margin="0,0,0,6">
    <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" Margin="8,0,0,0">
      <Button x:Name="BtnMarkAll" Content="Todas"
              Style="{StaticResource BtnSelectOutline}"
              MinWidth="72" Height="28" Margin="0,0,6,0" Padding="8,0"/>
      <Button x:Name="BtnMarkNone" Content="Ninguna"
              Style="{StaticResource BtnSelectOutline}"
              MinWidth="72" Height="28" Margin="0,0,6,0" Padding="8,0"/>
      <Button x:Name="BtnRefresh" Content="Actualizar"
              Style="{StaticResource BtnSelectOutline}"
              MinWidth="88" Height="28" Padding="8,0"/>
    </StackPanel>
    <TextBox x:Name="TxtFilter" Style="{StaticResource BimToolsTextBoxDark}"
             MinHeight="28" VerticalContentAlignment="Center"
             ToolTip="Filtrar por nombre de categoría"/>
  </DockPanel>

  <TextBlock Style="{StaticResource Label}" Text="Categorías" Margin="0,4,0,4"/>
  <Border Background="#071018" BorderBrush="#21465C" BorderThickness="1"
          CornerRadius="4" Padding="6,4" MaxHeight="300">
    <ScrollViewer VerticalScrollBarVisibility="Auto"
                  HorizontalScrollBarVisibility="Disabled">
      <StackPanel x:Name="PanelCategories"/>
    </ScrollViewer>
  </Border>

  <StackPanel x:Name="PanelProgress" Margin="0,12,0,0" Visibility="Collapsed">
    <DockPanel LastChildFill="False" Margin="0,0,0,4">
      <TextBlock x:Name="TxtProgress" DockPanel.Dock="Left"
                 Foreground="#95B8CC" FontSize="11" TextWrapping="Wrap"
                 Text="Preparando…"/>
      <TextBlock x:Name="TxtProgressPct" DockPanel.Dock="Right"
                 Foreground="#64748b" FontSize="11"
                 HorizontalAlignment="Right" Text="0%"/>
    </DockPanel>
    <ProgressBar x:Name="BarProgress" Height="10" Minimum="0" Maximum="100"
                 Value="0" Background="#050E18" BorderBrush="#21465C"
                 BorderThickness="1" Foreground="#5BC0DE"/>
  </StackPanel>
</StackPanel>
"""

_PROGRESS_ACCENT_RGB = (91, 192, 222)

_FOOTER_LEADING_XAML = (
    u'<Button x:Name="BtnManual" Content="Manual" '
    u'Style="{{StaticResource BtnSelectOutline}}" '
    u'Background="{bg}" MinWidth="96" MinHeight="32" MaxHeight="36" '
    u'Padding="8,2" VerticalAlignment="Center" '
    u'ToolTip="Abrir manual de usuario"/>'
).format(bg=BTN_MANUAL)

_FOOTER_ACTIONS_XAML = u"""
<Button x:Name="BtnCopy" Content="Copiar al host"
        Style="{StaticResource BtnPrimary}" MinWidth="150"
        Margin="0,0,10,0"/>
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


def _brush(hex_color):
    h = _as_unicode(hex_color).lstrip(u"#")
    return SolidColorBrush(
        Color.FromRgb(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    )


def _link_display_label(entry):
    """Etiqueta corta y legible para el ComboBox de vínculos."""
    if not entry:
        return u"Vínculo"
    title = _as_unicode(entry.get(u"link_doc_title") or u"").strip()
    name = _as_unicode(entry.get(u"name") or u"").strip()
    # Nombre de instancia Revit suele ser "archivo.rvt : ubicación : …"
    short_name = name
    if u" : " in name:
        short_name = name.split(u" : ")[0].strip()
    elif u":" in name:
        short_name = name.split(u":")[0].strip()
    if title and short_name and title.lower() not in short_name.lower():
        return u"{0} — {1}".format(short_name, title)
    return short_name or title or u"Vínculo"


class _CopyProgress(object):
    """ProgressBar pyRevit (acento BIMTools) + sincronización con la UI WPF."""

    def __init__(self, total, title_prefix=None, ui_callback=None):
        self._total = max(1, int(total or 1))
        self._pb = None
        self._open = False
        self._title_prefix = title_prefix or _TOOL_TITLE
        self._ui_callback = ui_callback

    def __enter__(self):
        try:
            from pyrevit import forms as _pyrevit_forms

            self._pb = _pyrevit_forms.ProgressBar(
                title=self._title(0, u""),
                cancellable=False,
            )
            try:
                r, g, b = _PROGRESS_ACCENT_RGB
                self._pb.Resources[u"pyRevitAccentBrush"] = SolidColorBrush(
                    Color.FromRgb(r, g, b)
                )
            except Exception:
                pass
            self._pb.__enter__()
            self._open = True
        except Exception:
            self._pb = None
            self._open = False
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._open and self._pb is not None:
            try:
                self._pb.__exit__(exc_type, exc_val, exc_tb)
            except Exception:
                pass
        self._open = False
        self._pb = None
        return False

    def _title(self, current, label):
        cur = max(0, int(current))
        base = u"{0} — {1}/{2}".format(self._title_prefix, cur, int(self._total))
        if label:
            base = u"{0} · {1}".format(base, _as_unicode(label))
        return base

    def update(self, current, total=None, label=u""):
        tot = max(1, int(total or self._total))
        self._total = tot
        cur = max(0, min(int(current), tot))
        if self._ui_callback is not None:
            try:
                self._ui_callback(cur, tot, label)
            except Exception:
                pass
        if self._pb is None:
            return
        try:
            if hasattr(self._pb, u"update_progress"):
                try:
                    self._pb.update_progress(max(1, cur), max_value=tot)
                except TypeError:
                    try:
                        self._pb.update_progress(max(1, cur), max=tot)
                    except Exception:
                        pass
        except Exception:
            pass
        try:
            self._pb.title = self._title(cur, label)
        except Exception:
            pass


def _pump_wpf():
    """Fuerza un ciclo de mensajes WPF para refrescar la barra en el hilo de Revit."""
    try:
        from System import Action
        from System.Windows.Threading import Dispatcher, DispatcherPriority

        Dispatcher.CurrentDispatcher.Invoke(
            DispatcherPriority.Background, Action(lambda: None)
        )
    except Exception:
        pass


def _format_copy_summary(outcome):
    """
    Texto de resumen final: totales + desglose por categoría + warnings.

    Returns:
        (instruction, content) en unicode.
    """
    if not outcome:
        return (
            u"Copia finalizada sin resultados.",
            u"",
        )
    copied = int(outcome.get(u"copied", 0) or 0)
    requested = int(outcome.get(u"requested", 0) or 0)
    by_cat = outcome.get(u"by_category") or []
    errors = outcome.get(u"errors") or []
    warnings = outcome.get(u"warnings") or []
    warn_total = int(outcome.get(u"warning_count", 0) or 0)
    if warn_total < 1 and by_cat:
        for row in by_cat:
            warn_total += int(row.get(u"warning_count", 0) or 0)
            if not row.get(u"warning_count"):
                warn_total += len(row.get(u"warnings") or [])
    warn_unique = int(outcome.get(u"warning_unique_count", 0) or len(warnings))
    deleted_invalid = int(outcome.get(u"deleted_invalid", 0) or 0)

    instruction = u"Copiados {0} de {1} elemento(s).".format(copied, requested)
    extras = []
    if warn_total > 0:
        if warn_unique and warn_unique != warn_total:
            extras.append(
                u"{0} warning(s) ({1} tipos)".format(warn_total, warn_unique)
            )
        else:
            extras.append(u"{0} warning(s)".format(warn_total))
    if deleted_invalid > 0:
        extras.append(
            u"{0} elemento(s) inválido(s) eliminado(s)".format(deleted_invalid)
        )
    if errors:
        extras.append(u"errores en algunas categorías")
    if extras:
        instruction += u" · " + u", ".join(extras) + u"."

    lines = []
    if by_cat:
        lines.append(u"Detalle por categoría:")
        for row in by_cat:
            label = _as_unicode(row.get(u"label") or u"(categoría)")
            n_copied = int(row.get(u"copied", 0) or 0)
            n_req = int(row.get(u"requested", 0) or 0)
            err = _as_unicode(row.get(u"error") or u"").strip()
            n_warn = int(row.get(u"warning_count", 0) or 0)
            if n_warn < 1:
                n_warn = len(row.get(u"warnings") or [])
            if err:
                lines.append(
                    u"• {0}: {1} / {2} (error)".format(label, n_copied, n_req)
                )
            elif n_warn:
                lines.append(
                    u"• {0}: {1} / {2} ({3} warning(s))".format(
                        label, n_copied, n_req, n_warn
                    )
                )
            else:
                lines.append(u"• {0}: {1} / {2}".format(label, n_copied, n_req))

    if warnings or warn_total:
        lines.append(u"")
        lines.append(
            u"Warnings capturados (no revirtieron la copia) — "
            u"{0} en total, {1} tipo(s):".format(warn_total, warn_unique)
        )
        for w in warnings[:20]:
            lines.append(u"• {0}".format(_as_unicode(w)))
        if len(warnings) > 20:
            lines.append(
                u"• … y {0} tipos más (ver consola pyRevit).".format(
                    len(warnings) - 20
                )
            )

    if errors:
        lines.append(u"")
        lines.append(u"Errores:")
        for err in errors[:12]:
            lines.append(u"• {0}".format(_as_unicode(err)))

    return instruction, u"\n".join(lines)


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
        min_height=420,
        resize_mode=u"CanResizeWithGrip",
        size_to_content_height=True,
    )


class _CopyHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref
        self.request = None

    def Execute(self, uiapp):
        win = self._window_ref()
        if win is None:
            return
        req = self.request
        self.request = None
        if req is None:
            return

        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            win._set_status(u"No hay documento activo.")
            _mostrar_aviso(uiapp, u"No hay documento activo.")
            return

        doc = uidoc.Document
        if doc.IsFamilyDocument:
            win._set_status(u"No aplica a documentos de familia.")
            _mostrar_aviso(uiapp, u"No aplica a documentos de familia.")
            return

        if doc.IsReadOnly:
            win._set_status(u"El documento no es editable.")
            _mostrar_aviso(uiapp, u"El documento activo es de solo lectura.")
            return

        link_inst = req.get(u"link_instance")
        groups = req.get(u"groups") or []
        if link_inst is None:
            win._set_status(u"Selecciona un vínculo cargado.")
            _mostrar_aviso(uiapp, u"Selecciona un vínculo Revit cargado.")
            return
        if not groups:
            win._set_status(u"Marca al menos una categoría con elementos.")
            _mostrar_aviso(
                uiapp,
                u"Marca al menos una categoría con elementos para copiar.",
            )
            return

        n_steps = 0
        for g in groups:
            if g.get(u"element_ids"):
                n_steps += 1
        n_steps = max(1, n_steps)

        win._set_busy(True)
        win._show_progress(0, n_steps, u"Preparando…")

        def _on_progress(current, total, label):
            win._show_progress(current, total, label)
            _pump_wpf()

        outcome = None
        try:
            with _CopyProgress(
                n_steps,
                title_prefix=_TOOL_TITLE,
                ui_callback=_on_progress,
            ) as pb:

                def _cb(current, total, label):
                    pb.update(current, total, label)

                outcome = run_copy_in_transaction(
                    doc,
                    link_inst,
                    groups,
                    _TX_COPY,
                    progress_callback=_cb,
                )
        except Exception as ex:
            win._set_busy(False)
            win._hide_progress()
            win._set_status(u"Error durante la copia.")
            _mostrar_aviso(
                uiapp,
                u"No se pudieron copiar los elementos seleccionados.",
                content=_as_unicode(ex),
            )
            return

        win._set_busy(False)
        win._hide_progress()

        if outcome is None:
            win._set_status(u"Copia cancelada o sin resultados.")
            return

        committed = outcome.get(u"committed", False)
        copied = outcome.get(u"copied", 0)
        errors = outcome.get(u"errors") or []
        instruction, detail = _format_copy_summary(outcome)

        if not committed and copied < 1:
            win._set_status(u"Copia cancelada o sin resultados.")
            _mostrar_aviso(
                uiapp,
                u"No se pudieron copiar los elementos seleccionados.",
                content=detail
                or (
                    u"\n".join(errors[:8])
                    if errors
                    else u"Revit no devolvió elementos copiados."
                ),
            )
            return

        win._set_status(instruction)
        win._apply_result_summary(outcome)
        _mostrar_aviso(uiapp, instruction, content=detail, ok_text=u"Cerrar")

    def GetName(self):
        return _TX_COPY


class CopiarDesdeVinculoWindow(object):
    def __init__(self, uiapp):
        self._uiapp = uiapp
        self._win = XamlReader.Parse(_build_xaml())
        self._txt_subtitle = self._win.FindName(u"TxtSubtitle")
        self._txt_summary = self._win.FindName(u"TxtSummary")
        self._txt_status = self._win.FindName(u"TxtStatus")
        self._txt_filter = self._win.FindName(u"TxtFilter")
        self._cbo_link = self._win.FindName(u"CboLink")
        self._panel_categories = self._win.FindName(u"PanelCategories")
        self._panel_progress = self._win.FindName(u"PanelProgress")
        self._bar_progress = self._win.FindName(u"BarProgress")
        self._txt_progress = self._win.FindName(u"TxtProgress")
        self._txt_progress_pct = self._win.FindName(u"TxtProgressPct")
        self._btn_copy = self._win.FindName(u"BtnCopy")
        self._busy = False

        self._links = []
        self._groups = []
        self._checkboxes = []
        self._category_rows = []

        self._handler = _CopyHandler(weakref.ref(self))
        self._ext_event = ExternalEvent.Create(self._handler)

        self._wire_events()
        _prepare_window(self._win, uiapp)
        if self._txt_subtitle is not None:
            self._txt_subtitle.Text = _SUBTITLE
        if self._txt_filter is not None:
            try:
                self._txt_filter.Tag = u"filter"
            except Exception:
                pass
        self._reload_links()
        self._reload_categories()
        self._set_status(u"Selecciona categorías y pulsa Copiar al host.")

    def _wire_events(self):
        self._win.FindName(u"BtnCopy").Click += RoutedEventHandler(self._on_copy)
        self._win.FindName(u"BtnClose").Click += RoutedEventHandler(self._on_close)
        self._win.FindName(u"BtnMarkAll").Click += RoutedEventHandler(self._on_mark_all)
        self._win.FindName(u"BtnMarkNone").Click += RoutedEventHandler(self._on_mark_none)
        self._win.FindName(u"BtnRefresh").Click += RoutedEventHandler(self._on_refresh)
        manual = self._win.FindName(u"BtnManual")
        if manual is not None:
            manual.Click += RoutedEventHandler(self._on_manual)
        if self._cbo_link is not None:
            self._cbo_link.SelectionChanged += SelectionChangedEventHandler(
                self._on_link_changed
            )
        if self._txt_filter is not None:
            self._txt_filter.TextChanged += TextChangedEventHandler(
                self._on_filter_changed
            )
        self._win.KeyDown += KeyEventHandler(self._on_key_down)
        self._win.Closed += EventHandler(self._on_closed)

    def _on_manual(self, sender, args):
        _open_manual(self._uiapp)

    def _on_key_down(self, sender, args):
        if args.Key == Key.Escape:
            if self._busy:
                return
            self._win.Close()

    def _on_close(self, sender, args):
        if self._busy:
            return
        self._win.Close()

    def _on_closed(self, sender, args):
        try:
            AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, None)
        except Exception:
            pass

    def _on_mark_all(self, sender, args):
        for cb in self._visible_checkboxes():
            if cb.IsEnabled:
                cb.IsChecked = True
        self._refresh_selection_status()

    def _on_mark_none(self, sender, args):
        for cb in self._visible_checkboxes():
            cb.IsChecked = False
        self._refresh_selection_status()

    def _on_refresh(self, sender, args):
        self._reload_links()
        self._reload_categories()
        self._set_status(u"Lista actualizada.")

    def _on_link_changed(self, sender, args):
        self._reload_categories()

    def _on_filter_changed(self, sender, args):
        self._apply_filter()

    def _set_status(self, text):
        if self._txt_status is not None:
            self._txt_status.Text = _as_unicode(text)

    def _set_busy(self, busy):
        self._busy = bool(busy)
        enabled = not self._busy
        for name in (
            u"BtnCopy",
            u"BtnMarkAll",
            u"BtnMarkNone",
            u"BtnRefresh",
            u"CboLink",
            u"TxtFilter",
        ):
            ctrl = self._win.FindName(name)
            if ctrl is None:
                continue
            try:
                ctrl.IsEnabled = enabled
            except Exception:
                pass
        for cb in self._checkboxes:
            try:
                if self._busy:
                    cb.IsEnabled = False
                else:
                    grp = cb.Tag
                    cb.IsEnabled = int(grp.get(u"count", 0) or 0) > 0
            except Exception:
                pass

    def _show_progress(self, current, total, label=u""):
        tot = max(1, int(total or 1))
        cur = max(0, min(int(current or 0), tot))
        pct = int(round(100.0 * float(cur) / float(tot))) if tot else 0
        if self._panel_progress is not None:
            self._panel_progress.Visibility = Visibility.Visible
        if self._bar_progress is not None:
            try:
                self._bar_progress.Minimum = 0
                self._bar_progress.Maximum = 100
                self._bar_progress.Value = pct
            except Exception:
                pass
        if self._txt_progress is not None:
            if label:
                self._txt_progress.Text = u"Copiando {0}/{1}: {2}".format(
                    cur, tot, _as_unicode(label)
                )
            else:
                self._txt_progress.Text = u"Copiando {0}/{1}…".format(cur, tot)
        if self._txt_progress_pct is not None:
            self._txt_progress_pct.Text = u"{0}%".format(pct)
        self._set_status(
            u"Copiando {0}/{1}…".format(cur, tot)
            if not label
            else u"Copiando {0}/{1}: {2}".format(cur, tot, _as_unicode(label))
        )

    def _hide_progress(self):
        if self._panel_progress is not None:
            self._panel_progress.Visibility = Visibility.Collapsed
        if self._bar_progress is not None:
            try:
                self._bar_progress.Value = 0
            except Exception:
                pass

    def _host_doc(self):
        uidoc = self._uiapp.ActiveUIDocument
        if uidoc is None:
            return None
        return uidoc.Document

    def _selected_link_entry(self):
        if self._cbo_link is None or self._cbo_link.SelectedItem is None:
            return None
        item = self._cbo_link.SelectedItem
        try:
            tag = item.Tag
            if isinstance(tag, dict):
                return tag
        except Exception:
            pass
        return None

    def _reload_links(self):
        doc = self._host_doc()
        prev_id = None
        prev = self._selected_link_entry()
        if prev:
            prev_id = prev.get(u"link_id_int")

        self._links = list_loaded_revit_links(doc) if doc else []
        if self._cbo_link is None:
            return

        self._cbo_link.Items.Clear()
        selected_index = 0
        for i, entry in enumerate(self._links):
            label = _link_display_label(entry)
            entry[u"display"] = label
            item = ComboBoxItem()
            item.Content = label
            item.Tag = entry
            self._cbo_link.Items.Add(item)
            if prev_id is not None and entry.get(u"link_id_int") == prev_id:
                selected_index = i

        if self._links:
            self._cbo_link.SelectedIndex = selected_index

    def _reload_categories(self):
        entry = self._selected_link_entry()
        link_inst = entry.get(u"instance") if entry else None
        link_doc = None
        if link_inst is not None:
            try:
                link_doc = link_inst.GetLinkDocument()
            except Exception:
                link_doc = None
        self._groups = collect_link_categories(link_doc)
        self._build_category_panel()
        self._update_summary()
        self._refresh_selection_status()

    def _update_summary(self):
        stats = summarize_category_groups(self._groups)
        if self._txt_summary is None:
            return
        entry = self._selected_link_entry()
        if entry is None:
            self._txt_summary.Text = u"No hay vínculos Revit cargados en el proyecto."
            return
        self._txt_summary.Text = (
            u"{0} · {1} categoría(s) con elementos · {2} elemento(s)".format(
                entry.get(u"display") or u"Vínculo",
                stats.get(u"categories_with_elements", 0),
                stats.get(u"total", 0),
            )
        )

    def _apply_result_summary(self, outcome):
        """Actualiza el panel de resumen de la ventana con el desglose final."""
        if self._txt_summary is None or not outcome:
            return
        instruction, _detail = _format_copy_summary(outcome)
        lines = [instruction]
        for row in outcome.get(u"by_category") or []:
            label = _as_unicode(row.get(u"label") or u"")
            n_copied = int(row.get(u"copied", 0) or 0)
            n_req = int(row.get(u"requested", 0) or 0)
            err = _as_unicode(row.get(u"error") or u"").strip()
            n_warn = int(row.get(u"warning_count", 0) or 0)
            if n_warn < 1:
                n_warn = len(row.get(u"warnings") or [])
            if err:
                lines.append(u"• {0}: {1}/{2} (error)".format(label, n_copied, n_req))
            elif n_warn:
                lines.append(
                    u"• {0}: {1}/{2} ({3} warn)".format(label, n_copied, n_req, n_warn)
                )
            else:
                lines.append(u"• {0}: {1}/{2}".format(label, n_copied, n_req))
        n_warn_total = int(outcome.get(u"warning_count", 0) or 0)
        n_warn_unique = int(outcome.get(u"warning_unique_count", 0) or 0)
        if n_warn_total:
            if n_warn_unique and n_warn_unique != n_warn_total:
                lines.append(
                    u"Warnings: {0} ({1} tipos) — detalle en el diálogo / consola".format(
                        n_warn_total, n_warn_unique
                    )
                )
            else:
                lines.append(
                    u"Warnings: {0} (detalle en el diálogo / consola)".format(n_warn_total)
                )
        self._txt_summary.Text = u"\n".join(lines)

    def _make_category_row(self, grp):
        """Fila: checkbox + nombre a la izquierda, conteo a la derecha."""
        count = int(grp.get(u"count", 0) or 0)
        enabled = count > 0

        row = Grid()
        row.Margin = Thickness(2, 1, 2, 1)
        col0 = ColumnDefinition()
        col0.Width = GridLength(1, GridUnitType.Star)
        col1 = ColumnDefinition()
        col1.Width = GridLength.Auto
        row.ColumnDefinitions.Add(col0)
        row.ColumnDefinitions.Add(col1)

        cb = CheckBox()
        cb.Content = grp.get(u"label") or u""
        cb.IsChecked = False
        cb.IsEnabled = enabled
        cb.Margin = Thickness(4, 5, 4, 5)
        cb.Padding = Thickness(6, 0, 0, 0)
        cb.Cursor = Cursors.Hand if enabled else Cursors.Arrow
        cb.Foreground = _brush(FG_TITLE if enabled else FG_MUTED)
        cb.VerticalAlignment = VerticalAlignment.Center
        cb.Tag = grp
        cb.Checked += RoutedEventHandler(self._on_category_toggled)
        cb.Unchecked += RoutedEventHandler(self._on_category_toggled)
        Grid.SetColumn(cb, 0)
        row.Children.Add(cb)

        count_tb = TextBlock()
        count_tb.Text = u"{0}".format(count)
        count_tb.Foreground = _brush(FG_MUTED if not enabled else FG_BODY)
        count_tb.FontSize = 11
        count_tb.Margin = Thickness(8, 0, 8, 0)
        count_tb.VerticalAlignment = VerticalAlignment.Center
        count_tb.HorizontalAlignment = HorizontalAlignment.Right
        Grid.SetColumn(count_tb, 1)
        row.Children.Add(count_tb)

        row.Tag = grp
        return row, cb

    def _build_category_panel(self):
        if self._panel_categories is None:
            return
        self._panel_categories.Children.Clear()
        self._checkboxes = []
        self._category_rows = []

        if not self._groups:
            empty = TextBlock()
            empty.Text = u"Selecciona un vínculo para listar categorías."
            empty.TextWrapping = TextWrapping.Wrap
            empty.Foreground = _brush(FG_BODY)
            empty.Margin = Thickness(8, 10, 8, 10)
            self._panel_categories.Children.Add(empty)
            return

        for grp in self._groups:
            row, cb = self._make_category_row(grp)
            self._panel_categories.Children.Add(row)
            self._checkboxes.append(cb)
            self._category_rows.append(row)

        self._apply_filter()

    def _on_category_toggled(self, sender, args):
        self._refresh_selection_status()

    def _refresh_selection_status(self):
        groups = self._selected_groups()
        n_cats = len(groups)
        n_elems = 0
        for g in groups:
            n_elems += int(g.get(u"count", 0) or 0)
        if n_cats < 1:
            self._set_status(u"Ninguna categoría marcada.")
        else:
            self._set_status(
                u"{0} categoría(s) · {1} elemento(s) a copiar.".format(n_cats, n_elems)
            )

    def _visible_checkboxes(self):
        out = []
        for cb in self._checkboxes:
            try:
                parent = cb.Parent
                if parent is not None and parent.Visibility != Visibility.Visible:
                    continue
                if cb.Visibility == Visibility.Visible:
                    out.append(cb)
            except Exception:
                out.append(cb)
        return out

    def _apply_filter(self):
        query = u""
        if self._txt_filter is not None:
            query = (_as_unicode(self._txt_filter.Text) or u"").strip().lower()
        for row in self._category_rows:
            grp = row.Tag
            label = _as_unicode(grp.get(u"label") if grp else u"").lower()
            visible = (not query) or (query in label)
            row.Visibility = Visibility.Visible if visible else Visibility.Collapsed

    def _selected_groups(self):
        selected = []
        for cb in self._checkboxes:
            try:
                if cb.IsChecked and cb.IsEnabled:
                    grp = cb.Tag
                    if grp is not None and int(grp.get(u"count", 0) or 0) > 0:
                        selected.append(grp)
            except Exception:
                continue
        return selected

    def _on_copy(self, sender, args):
        if self._busy:
            return
        entry = self._selected_link_entry()
        groups = self._selected_groups()
        if entry is None:
            _mostrar_aviso(self._uiapp, u"Selecciona un vínculo Revit cargado.")
            return
        if not groups:
            _mostrar_aviso(
                self._uiapp,
                u"Marca al menos una categoría con elementos para copiar.",
            )
            return
        self._set_status(u"Copiando elementos…")
        self._show_progress(0, max(1, len(groups)), u"Preparando…")
        self._handler.request = {
            u"link_instance": entry.get(u"instance"),
            u"groups": groups,
        }
        self._ext_event.Raise()

    def show(self):
        self._win.Show()


def _existing_controller():
    try:
        ctrl = AppDomain.CurrentDomain.GetData(_SINGLETON_KEY)
    except Exception:
        ctrl = None
    if ctrl is None:
        return None
    try:
        if hasattr(ctrl, "_win") and ctrl._win is not None and ctrl._win.IsVisible:
            return ctrl
    except Exception:
        pass
    return None


def _focus_existing(ctrl, uiapp):
    try:
        if ctrl._win.WindowState == WindowState.Minimized:
            ctrl._win.WindowState = WindowState.Normal
        ctrl._win.Activate()
    except Exception:
        pass
    _mostrar_aviso(uiapp, _ALREADY_RUNNING)


def run(revit):
    uiapp = revit if hasattr(revit, "ActiveUIDocument") else None
    if uiapp is None:
        try:
            uiapp = revit.uiapp
        except Exception:
            uiapp = None
    if uiapp is None:
        return

    existing = _existing_controller()
    if existing is not None:
        _focus_existing(existing, uiapp)
        return

    try:
        ctrl = CopiarDesdeVinculoWindow(uiapp)
    except Exception as ex:
        tb = traceback.format_exc()
        print(tb)
        _mostrar_aviso(
            uiapp,
            u"No se pudo abrir la ventana de la herramienta.",
            content=_as_unicode(ex) + u"\n\n" + _as_unicode(tb),
        )
        return

    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, ctrl)
    except Exception:
        pass
    ctrl.show()
