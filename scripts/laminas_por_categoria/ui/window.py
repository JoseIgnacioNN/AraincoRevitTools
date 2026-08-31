# -*- coding: utf-8 -*-
"""UI WPF — Láminas por categoría."""

from __future__ import print_function

try:
    unicode
except NameError:
    unicode = str

import os
from datetime import datetime

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from System import AppDomain, EventHandler
from System.Windows import (
    RoutedEventHandler,
)
from System.Windows.Controls import (
    ComboBoxItem,
    SelectionChangedEventHandler,
    TextChangedEventHandler,
)
from System.Windows.Markup import XamlReader
from System.Windows.Media import Color, SolidColorBrush
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent, TaskDialog

from bimtools_ui_tokens import BTN_MANUAL
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_wpf_shell import build_simple_tool_xaml
from revit_wpf_window_position import (
    bind_center_wpf_on_revit_monitor,
    position_wpf_window_center_on_monitor,
    revit_main_hwnd,
)

from laminas_por_categoria import singleton
from laminas_por_categoria.constants import (
    CATEGORIA_OPTIONS,
    MONTH_OPTIONS,
    TRANSACTION_TITLE,
)
from laminas_por_categoria.people import load_personas, personas_json_path

_DIALOG_TITLE = TRANSACTION_TITLE
_WINDOW_TITLE = u"Arainco: Láminas por categoría"
_SUBTITLE = u"Crea láminas vacías con correlativo por Clasificacion."
_APPDOMAIN_EVENT_KEY = u"Arainco_LaminasPorCategoria_ExtEvent"
_APPDOMAIN_HANDLER_KEY = u"Arainco_LaminasPorCategoria_Handler"

_BODY_XAML = u"""
<Grid>
  <Grid.RowDefinitions>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
  </Grid.RowDefinitions>
  <Grid.ColumnDefinitions>
    <ColumnDefinition Width="150"/>
    <ColumnDefinition Width="*"/>
  </Grid.ColumnDefinitions>

  <TextBlock Grid.Row="0" Grid.Column="0" Text="Tipo de formato"
             Foreground="#95B8CC" FontSize="11" FontWeight="SemiBold"
             VerticalAlignment="Center" Margin="0,0,10,8" TextWrapping="Wrap"/>
  <ComboBox x:Name="CmbFormato" Grid.Row="0" Grid.Column="1" Margin="0,0,0,8"
            Style="{StaticResource ComboStretch}" IsEditable="False"
            ToolTip="Cajetín (se excluye EST_A_SPLASH SCREEN)"/>

  <TextBlock Grid.Row="1" Grid.Column="0" Text="Categoría"
             Foreground="#95B8CC" FontSize="11" FontWeight="SemiBold"
             VerticalAlignment="Center" Margin="0,0,10,8" TextWrapping="Wrap"/>
  <ComboBox x:Name="CmbCategoria" Grid.Row="1" Grid.Column="1" Margin="0,0,0,8"
            Style="{StaticResource ComboStretch}" IsEditable="False"
            ToolTip="Código escrito en Clasificacion y en el número (PG-001)"/>

  <TextBlock Grid.Row="2" Grid.Column="0" Text="Cantidad de láminas"
             Foreground="#95B8CC" FontSize="11" FontWeight="SemiBold"
             VerticalAlignment="Center" Margin="0,0,10,8" TextWrapping="Wrap"/>
  <TextBox x:Name="TxtCantidad" Grid.Row="2" Grid.Column="1" Margin="0,0,0,8"
           Style="{StaticResource BimToolsTextBoxDark}" Text="1"
           MinHeight="30" VerticalContentAlignment="Center"
           ToolTip="Entero mayor que 0"/>

  <TextBlock Grid.Row="3" Grid.Column="0" Text="Calculó"
             Foreground="#95B8CC" FontSize="11" FontWeight="SemiBold"
             VerticalAlignment="Center" Margin="0,0,10,8"/>
  <ComboBox x:Name="CmbCalculo" Grid.Row="3" Grid.Column="1" Margin="0,0,0,8"
            Style="{StaticResource ComboStretch}" IsEditable="False"/>

  <TextBlock Grid.Row="4" Grid.Column="0" Text="Revisó"
             Foreground="#95B8CC" FontSize="11" FontWeight="SemiBold"
             VerticalAlignment="Center" Margin="0,0,10,8"/>
  <ComboBox x:Name="CmbReviso" Grid.Row="4" Grid.Column="1" Margin="0,0,0,8"
            Style="{StaticResource ComboStretch}" IsEditable="False"/>

  <TextBlock Grid.Row="5" Grid.Column="0" Text="Aprobó"
             Foreground="#95B8CC" FontSize="11" FontWeight="SemiBold"
             VerticalAlignment="Center" Margin="0,0,10,8"/>
  <ComboBox x:Name="CmbAprobo" Grid.Row="5" Grid.Column="1" Margin="0,0,0,8"
            Style="{StaticResource ComboStretch}" IsEditable="False"/>

  <TextBlock Grid.Row="6" Grid.Column="0" Text="Dibujó"
             Foreground="#95B8CC" FontSize="11" FontWeight="SemiBold"
             VerticalAlignment="Center" Margin="0,0,10,8"/>
  <ComboBox x:Name="CmbDibujo" Grid.Row="6" Grid.Column="1" Margin="0,0,0,8"
            Style="{StaticResource ComboStretch}" IsEditable="False"/>

  <TextBlock Grid.Row="7" Grid.Column="0" Text="Fecha de creación"
             Foreground="#95B8CC" FontSize="11" FontWeight="SemiBold"
             VerticalAlignment="Center" Margin="0,0,10,8" TextWrapping="Wrap"/>
  <ComboBox x:Name="CmbFecha" Grid.Row="7" Grid.Column="1" Margin="0,0,0,8"
            Style="{StaticResource ComboStretch}" IsEditable="False"
            ToolTip="Se escribe en Sheet Issue Date (ENE. AAAA)"/>

  <TextBlock x:Name="TxtPreview" Grid.Row="8" Grid.Column="0" Grid.ColumnSpan="2"
             Foreground="#64748b" FontSize="11" TextWrapping="Wrap" Margin="0,4,0,0"
             Text=""/>
</Grid>
"""

_FOOTER_LEADING_XAML = (
    u'<Button x:Name="BtnManual" Content="Manual" '
    u'Style="{{StaticResource BtnSelectOutline}}" '
    u'Background="{bg}" MinWidth="96" Padding="8,2" '
    u'ToolTip="Abrir manual de usuario" VerticalAlignment="Center"/>'
).format(bg=BTN_MANUAL)

_FOOTER_ACTIONS_XAML = u"""
<Button x:Name="BtnCancelar" Content="Cancelar" Margin="0,0,8,0"
        Style="{StaticResource BtnSelectOutline}" MinWidth="100"/>
<Button x:Name="BtnIniciar" Content="Crear láminas"
        Style="{StaticResource BtnPrimary}" MinWidth="140"
        ToolTip="Crear láminas CONTENIDO LAMINA con el correlativo de la categoría"/>
"""


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except Exception:
        return str(text)


def mostrar_aviso(uiapp, instruction, content=u"", ok_text=u"Entendido"):
    hwnd = None
    try:
        if uiapp is not None:
            hwnd = revit_main_hwnd(uiapp)
    except Exception:
        pass
    try:
        from bimtools_instruction_dialog import show_message_dialog

        show_message_dialog(
            _DIALOG_TITLE,
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
        TaskDialog.Show(_DIALOG_TITLE, body)
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
    try:
        cursor = os.path.dirname(os.path.abspath(__file__))
        for _ in range(16):
            try:
                names = os.listdir(cursor)
            except Exception:
                names = []
            if any(_as_unicode(n).endswith(u".tab") for n in names):
                ext_dir = cursor
                break
            parent = os.path.dirname(cursor)
            if parent == cursor:
                ext_dir = None
                break
            cursor = parent
        else:
            ext_dir = None
        if ext_dir:
            for tab_name in os.listdir(ext_dir):
                if not _as_unicode(tab_name).endswith(u".tab"):
                    continue
                panel = os.path.join(ext_dir, tab_name, u"Modelado.panel")
                if not os.path.isdir(panel):
                    continue
                for pb_name in os.listdir(panel):
                    if u"LaminasPorCategoria" not in _as_unicode(pb_name):
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
        mostrar_aviso(
            uiapp,
            u"No se encontró manual_usuario.html.",
            content=u"Debe estar en la carpeta del pushbutton de la herramienta.",
        )
        return
    try:
        os.startfile(path)
    except Exception as ex:
        mostrar_aviso(
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
        title=_WINDOW_TITLE,
        styles_xml=BIMTOOLS_DARK_STYLES_XML,
        body_xaml=_BODY_XAML,
        footer_leading_xaml=_FOOTER_LEADING_XAML,
        footer_actions_xaml=_FOOTER_ACTIONS_XAML,
        width=480,
        min_width=480,
        height=0,
        min_height=0,
        resize_mode=u"NoResize",
        size_to_content_height=True,
    )


def _pin_external_event(ext_event, handler):
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_EVENT_KEY, ext_event)
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_HANDLER_KEY, handler)
    except Exception:
        pass


def _unpin_external_event():
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_EVENT_KEY, None)
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_HANDLER_KEY, None)
    except Exception:
        pass


def _style_crear_button(btn, enabled):
    if btn is None:
        return
    try:
        btn.IsEnabled = bool(enabled)
    except Exception:
        pass
    try:
        if enabled:
            btn.Opacity = 1.0
            btn.Background = SolidColorBrush(Color.FromRgb(0x5B, 0xC0, 0xDE))
            btn.Foreground = SolidColorBrush(Color.FromRgb(0x0A, 0x1A, 0x2F))
            btn.BorderBrush = SolidColorBrush(Color.FromRgb(0x87, 0xD9, 0xEE))
        else:
            btn.Opacity = 0.38
            btn.Background = SolidColorBrush(Color.FromRgb(0x1E, 0x33, 0x44))
            btn.Foreground = SolidColorBrush(Color.FromRgb(0x64, 0x74, 0x8B))
            btn.BorderBrush = SolidColorBrush(Color.FromRgb(0x21, 0x46, 0x5C))
    except Exception:
        try:
            btn.Opacity = 1.0 if enabled else 0.4
        except Exception:
            pass


def _fill_combo(combo, pairs, selected_index=0):
    """pairs: lista de (display, tag)."""
    if combo is None:
        return
    combo.Items.Clear()
    for display, tag in pairs:
        it = ComboBoxItem()
        it.Content = _as_unicode(display)
        it.Tag = tag
        it.ToolTip = _as_unicode(display)
        combo.Items.Add(it)
    if combo.Items.Count > 0:
        idx = int(selected_index or 0)
        if idx < 0:
            idx = 0
        if idx >= combo.Items.Count:
            idx = 0
        combo.SelectedIndex = idx


def _combo_tag(combo):
    if combo is None:
        return None
    try:
        item = combo.SelectedItem
    except Exception:
        item = None
    if item is None:
        return None
    try:
        return item.Tag
    except Exception:
        return None


def _combo_text(combo):
    if combo is None:
        return u""
    try:
        item = combo.SelectedItem
    except Exception:
        item = None
    if item is None:
        return u""
    try:
        return _as_unicode(item.Content).strip()
    except Exception:
        return u""


class _CreateLaminasHandler(IExternalEventHandler):
    def __init__(self):
        self.request = None
        self.uiapp_for_dialog = None

    def Execute(self, uiapp):
        req = self.request
        self.request = None
        host = self.uiapp_for_dialog or uiapp
        if req is None:
            _unpin_external_event()
            return
        from laminas_por_categoria.service import (
            LaminasPorCategoriaError,
            create_laminas,
            format_success_dialog,
        )

        try:
            uidoc = uiapp.ActiveUIDocument
            if uidoc is None:
                mostrar_aviso(host, u"No hay documento activo.")
                return
            doc = uidoc.Document
            result = create_laminas(doc, req)
            instruction, content = format_success_dialog(
                result, req.categoria
            )
            mostrar_aviso(host, instruction, content, ok_text=u"Entendido")
        except LaminasPorCategoriaError as ex:
            mostrar_aviso(host, _as_unicode(ex))
        except Exception as ex:
            mostrar_aviso(
                host,
                u"Error al crear láminas.",
                content=_as_unicode(ex),
            )
        finally:
            _unpin_external_event()

    def GetName(self):
        return TRANSACTION_TITLE


class LaminasPorCategoriaWindow(object):
    def __init__(self, doc, uidoc, revit_app):
        self._doc = doc
        self._uidoc = uidoc
        self._revit = revit_app
        self._busy = False
        self._title_blocks = []

        self._create_handler = _CreateLaminasHandler()
        self._create_event = ExternalEvent.Create(self._create_handler)

        self._win = XamlReader.Parse(_build_xaml())
        self._cmb_formato = self._win.FindName("CmbFormato")
        self._cmb_categoria = self._win.FindName("CmbCategoria")
        self._txt_cantidad = self._win.FindName("TxtCantidad")
        self._cmb_calculo = self._win.FindName("CmbCalculo")
        self._cmb_reviso = self._win.FindName("CmbReviso")
        self._cmb_aprobo = self._win.FindName("CmbAprobo")
        self._cmb_dibujo = self._win.FindName("CmbDibujo")
        self._cmb_fecha = self._win.FindName("CmbFecha")
        self._txt_preview = self._win.FindName("TxtPreview")
        self._txt_subtitle = self._win.FindName("TxtSubtitle")
        self._txt_status = self._win.FindName("TxtStatus")
        self._btn_iniciar = self._win.FindName("BtnIniciar")
        self._btn_cancelar = self._win.FindName("BtnCancelar")
        btn_manual = self._win.FindName("BtnManual")

        if self._txt_subtitle is not None:
            try:
                self._txt_subtitle.Text = _SUBTITLE
            except Exception:
                pass

        self._fill_formatos()
        self._fill_categorias()
        self._fill_personas()
        self._fill_fecha()
        self._refresh_form_state()

        if self._cmb_formato is not None:
            self._cmb_formato.SelectionChanged += SelectionChangedEventHandler(
                self._on_form_changed
            )
        if self._cmb_categoria is not None:
            self._cmb_categoria.SelectionChanged += SelectionChangedEventHandler(
                self._on_form_changed
            )
        if self._txt_cantidad is not None:
            self._txt_cantidad.TextChanged += TextChangedEventHandler(
                self._on_form_changed
            )

        self._btn_iniciar.Click += RoutedEventHandler(self._on_iniciar)
        self._btn_cancelar.Click += RoutedEventHandler(lambda s, e: self._win.Close())
        if btn_manual is not None:
            btn_manual.Click += RoutedEventHandler(
                lambda s, e: _open_manual(self._revit)
            )
        self._win.Closed += EventHandler(lambda s, e: singleton.clear())

    def _fill_formatos(self):
        from laminas_por_categoria.service import collect_title_blocks

        self._title_blocks = collect_title_blocks(self._doc)
        pairs = [(tb.name, tb.symbol_id) for tb in self._title_blocks]
        _fill_combo(self._cmb_formato, pairs, 0)

    def _fill_categorias(self):
        pairs = [(label, code) for code, label in CATEGORIA_OPTIONS]
        _fill_combo(self._cmb_categoria, pairs, 0)

    def _fill_personas(self):
        data = load_personas()
        calc_items, calc_idx = data[u"calculo"]
        rev_items, rev_idx = data[u"revision"]
        apr_items, apr_idx = data[u"aprobacion"]
        dib_items, dib_idx = data[u"dibujo"]
        _fill_combo(self._cmb_calculo, [(x, x) for x in calc_items], calc_idx)
        _fill_combo(self._cmb_reviso, [(x, x) for x in rev_items], rev_idx)
        _fill_combo(self._cmb_aprobo, [(x, x) for x in apr_items], apr_idx)
        _fill_combo(self._cmb_dibujo, [(x, x) for x in dib_items], dib_idx)

    def _fill_fecha(self):
        year = datetime.now().year
        month = datetime.now().month
        pairs = []
        for label, abbr in MONTH_OPTIONS:
            value = u"{0} {1}".format(abbr, year)
            pairs.append((label, value))
        idx = month - 1
        if idx < 0 or idx >= len(pairs):
            idx = 0
        _fill_combo(self._cmb_fecha, pairs, idx)

    def _parse_cantidad(self):
        if self._txt_cantidad is None:
            return None
        raw = _as_unicode(self._txt_cantidad.Text).strip()
        if not raw:
            return None
        try:
            n = int(raw)
        except Exception:
            return None
        if n < 1:
            return None
        return n

    def _on_form_changed(self, _sender, _e):
        self._refresh_form_state()

    def _set_status(self, text):
        if self._txt_status is None:
            return
        try:
            self._txt_status.Text = _as_unicode(text or u"")
        except Exception:
            pass

    def _refresh_form_state(self):
        if self._busy:
            return
        tb_id = _combo_tag(self._cmb_formato)
        cat = _as_unicode(_combo_tag(self._cmb_categoria) or u"").strip()
        n = self._parse_cantidad()
        can_run = tb_id is not None and bool(cat) and n is not None
        _style_crear_button(self._btn_iniciar, can_run)

        preview = u""
        if tb_id is None:
            self._set_status(u"No hay cajetines en el modelo (salvo splash).")
        elif not cat:
            self._set_status(u"Seleccione una categoría.")
        elif n is None:
            self._set_status(u"Indique una cantidad entera mayor que 0.")
        else:
            from laminas_por_categoria.service import preview_numbers

            numbers, skipped = preview_numbers(self._doc, cat, n)
            if not numbers:
                self._set_status(u"No se pudo calcular el correlativo.")
            elif len(numbers) == 1:
                preview = u"Siguiente: {0}".format(numbers[0])
                self._set_status(preview)
            else:
                preview = u"Siguiente: {0} … {1} ({2} láminas)".format(
                    numbers[0], numbers[-1], len(numbers)
                )
                self._set_status(preview)
            if skipped:
                extra = u"Se omiten {0} número(s) no numérico(s) al correlativo.".format(
                    len(skipped)
                )
                if preview:
                    preview = preview + u" · " + extra
                else:
                    preview = extra
        if self._txt_preview is not None:
            try:
                self._txt_preview.Text = preview
            except Exception:
                pass

    def _on_iniciar(self, sender, args):
        if self._busy:
            return
        tb_id = _combo_tag(self._cmb_formato)
        cat = _as_unicode(_combo_tag(self._cmb_categoria) or u"").strip()
        n = self._parse_cantidad()
        if tb_id is None or not cat or n is None:
            self._refresh_form_state()
            return

        from laminas_por_categoria.service import LaminasPorCategoriaRequest

        req = LaminasPorCategoriaRequest(
            title_block_id=tb_id,
            categoria=cat,
            cantidad=n,
            aprobo=_combo_text(self._cmb_aprobo),
            calculo=_combo_text(self._cmb_calculo),
            reviso=_combo_text(self._cmb_reviso),
            dibujo=_combo_text(self._cmb_dibujo),
            fecha=_as_unicode(_combo_tag(self._cmb_fecha) or u"").strip(),
        )
        self._busy = True
        _style_crear_button(self._btn_iniciar, False)
        self._create_handler.request = req
        self._create_handler.uiapp_for_dialog = self._revit
        _pin_external_event(self._create_event, self._create_handler)
        try:
            self._create_event.Raise()
        except Exception as ex:
            self._busy = False
            _unpin_external_event()
            self._refresh_form_state()
            mostrar_aviso(
                self._revit,
                u"No se pudo iniciar la creación.",
                content=_as_unicode(ex),
            )
            return
        try:
            self._win.Close()
        except Exception:
            pass

    def show(self):
        _prepare_window(self._win, self._revit)
        singleton.register(self._win)
        self._win.Show()


def show_laminas_por_categoria_ui(revit_app):
    if singleton.try_activate_existing():
        mostrar_aviso(revit_app, u"La herramienta ya está en ejecución.")
        return
    try:
        uidoc = revit_app.ActiveUIDocument
        doc = uidoc.Document
    except Exception:
        mostrar_aviso(revit_app, u"No hay documento activo.")
        return

    from laminas_por_categoria.service import project_has_validacion

    if not project_has_validacion(doc):
        mostrar_aviso(
            revit_app,
            u"Este proyecto no tiene el parámetro Validacion en "
            u"Información de proyecto.",
            content=u"Use una plantilla Arainco. El JSON de personas está en:\n{0}".format(
                personas_json_path()
            ),
        )
        return

    from laminas_por_categoria.service import collect_title_blocks

    if not collect_title_blocks(doc):
        mostrar_aviso(
            revit_app,
            u"No hay tipos de cajetín utilizables.",
            content=u"Cargue familias de Title Blocks. Se excluye EST_A_SPLASH SCREEN.",
        )
        return

    w = LaminasPorCategoriaWindow(doc, uidoc, revit_app)
    w.show()
