# -*- coding: utf-8 -*-
"""
Numerar fundaciones — interfaz WPF con shell estándar BIMTools.

Cinta blanca Arainco + cuerpo oscuro (``bimtools_wpf_shell`` / tokens / dark theme).
"""

from __future__ import print_function

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    Transaction,
    FilteredElementCollector,
    FamilyInstance,
    WallFoundation,
    Floor,
)
from Autodesk.Revit.UI import TaskDialog

from System import AppDomain, EventHandler
from System.Windows import Visibility, WindowState
from System.Windows.Input import Key, KeyEventHandler
from System.Windows.Markup import XamlReader
from System.Windows.Media import SolidColorBrush, Color

from bimtools_ui_tokens import FG_BODY, FG_MUTED, FONT_SIZE_BODY, FONT_SIZE_HINT
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_wpf_shell import build_simple_tool_xaml
from revit_wpf_window_position import (
    bind_center_wpf_on_revit_monitor,
    position_wpf_window_center_on_monitor,
    revit_main_hwnd,
)

_APPDOMAIN_WINDOW_KEY = u"Arainco_NumerarFundaciones_UI"
_TOOL_DIALOG_TITLE = u"Arainco: Numerar Fundaciones"

_TEXT_INTRO = (
    u"Agrupa fundaciones aisladas por dimensiones de tipo (L \u00d7 W \u00d7 H) "
    u"y escribe el par\u00e1metro Numeracion Fundacion solo cuando est\u00e1 vac\u00edo."
)
_TEXT_CRITERIO = (
    u"Solo FamilyInstance (fundaciones aisladas). Se excluyen wall foundations "
    u"y slab foundations (Floor). No se sobrescribe lo ya numerado."
)

_BODY_XAML = u"""
<StackPanel>
  <TextBlock TextWrapping="Wrap" Foreground="{fg}" FontSize="{fs}" LineHeight="17"
             Text="{intro}"/>
  <TextBlock Margin="0,12,0,0" TextWrapping="Wrap" Foreground="{fg_lo}" FontSize="{fs_lo}"
             LineHeight="15" Text="{criterio}"/>
  <Border x:Name="PanelResultado" Margin="0,14,0,0" Visibility="Collapsed"
          Background="#050E18" BorderBrush="#21465C" BorderThickness="1"
          CornerRadius="4" Padding="0">
    <TextBox x:Name="TxtResult" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True"
             MinHeight="120" MaxHeight="320" VerticalScrollBarVisibility="Auto"
             Style="{{StaticResource BimToolsTextBoxDark}}" Padding="8,8" Text=""/>
  </Border>
</StackPanel>
""".format(
    fg=FG_BODY,
    fs=FONT_SIZE_BODY,
    fg_lo=FG_MUTED,
    fs_lo=FONT_SIZE_HINT,
    intro=_TEXT_INTRO.replace(u'"', u"&quot;"),
    criterio=_TEXT_CRITERIO.replace(u'"', u"&quot;"),
)

_FOOTER_ACTIONS_XAML = u"""
<Button x:Name="BtnClose" Content="Cerrar" Margin="0,0,8,0"
        Style="{StaticResource BtnSelectOutline}" MinWidth="108"/>
<Button x:Name="BtnNumerar" Content="Numerar fundaciones"
        Style="{StaticResource BtnPrimary}" MinWidth="160"/>
"""


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _mostrar_aviso(uiapp, instruction, content=u"", ok_text=u"Entendido"):
    """Diálogo informativo WPF (estilo BIMTools). Respaldo: TaskDialog."""
    hwnd = None
    try:
        if uiapp is not None:
            hwnd = revit_main_hwnd(uiapp)
    except Exception:
        pass
    try:
        from bimtools_instruction_dialog import show_message_dialog

        show_message_dialog(
            _TOOL_DIALOG_TITLE,
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
        TaskDialog.Show(_TOOL_DIALOG_TITLE, body)
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


def _build_xaml():
    return build_simple_tool_xaml(
        title=_TOOL_DIALOG_TITLE,
        styles_xml=BIMTOOLS_DARK_STYLES_XML,
        body_xaml=_BODY_XAML,
        footer_actions_xaml=_FOOTER_ACTIONS_XAML,
        width=520,
        resize_mode=u"NoResize",
        size_to_content_height=True,
    )


# ── Lógica de numeración ─────────────────────────────────────────────────────


def _get_dimensiones_y_volumen(element, doc):
    try:
        d = doc
        elem_type = d.GetElement(element.GetTypeId())
        if not elem_type:
            return ((0, 0, 0), 0)
        length_val, width_val, height_val = None, None, 1.0
        for name in ("Length", "Largo", "Longitud"):
            p = elem_type.LookupParameter(name)
            if p and p.HasValue:
                length_val = p.AsDouble() * 304.8
                break
        for name in ("Width", "Ancho"):
            p = elem_type.LookupParameter(name)
            if p and p.HasValue:
                width_val = p.AsDouble() * 304.8
                break
        for name in ("Height", "Depth", "Thickness", "Altura", "Profundidad", "Espesor"):
            p = elem_type.LookupParameter(name)
            if p and p.HasValue:
                height_val = p.AsDouble() * 304.8
                break
        if length_val is not None and width_val is not None:
            vol = length_val * width_val * height_val
            return ((length_val, width_val, height_val), vol)
    except Exception:
        pass
    return ((0, 0, 0), 0)


def _set_numeracion_fundacion(element, value):
    try:
        for param_name in (
            "Numeracion Fundacion",
            "Numeracion fundacion",
            "Numeracion",
            "Foundation Numbering",
        ):
            p = element.LookupParameter(param_name)
            if p is not None and not p.IsReadOnly:
                p.Set(str(value))
                return True
    except Exception:
        pass
    return False


def _leer_numeracion_fundacion(element):
    if not element:
        return None
    try:
        for param_name in (
            "Numeracion Fundacion",
            "Numeracion fundacion",
            "Numeracion",
            "Foundation Numbering",
        ):
            p = element.LookupParameter(param_name)
            if p is not None and p.HasValue:
                s = p.AsString()
                vs = p.AsValueString()
                val = s if s is not None else vs
                if val is not None and str(val).strip() and str(val).strip() != "0":
                    return str(val).strip()
                try:
                    d = p.AsDouble()
                    if d != 0.0:
                        return str(int(d)) if d == int(d) else str(d)
                except Exception:
                    pass
    except Exception:
        pass
    return None


def execute_numerar_fundaciones(doc):
    """
    Ejecuta la numeración en ``doc``.
    Retorna dict: ok (bool), text (unicode), is_error (bool).
    """
    from collections import defaultdict

    collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_StructuralFoundation
    ).WhereElementIsNotElementType()

    fundaciones = []
    total_cat = 0
    wall_found = 0
    floor_found = 0
    family_instance = 0
    otros = 0

    for elem in collector:
        total_cat += 1
        if isinstance(elem, WallFoundation):
            wall_found += 1
            continue
        if isinstance(elem, Floor):
            floor_found += 1
            continue
        if isinstance(elem, FamilyInstance):
            family_instance += 1
            fundaciones.append(elem)
        else:
            otros += 1

    if not fundaciones:
        msg = (
            u"No hay fundaciones aisladas.\n\n"
            u"Wall foundations y slab foundations se excluyen.\n\n"
            u"--- Diagnóstico ---\n"
            u"Total en categoría Structural Foundation: {}\n"
            u"  - Wall foundations: {}\n"
            u"  - Slab foundations (Floor): {}\n"
            u"  - FamilyInstance (aisladas): {}\n"
            u"  - Otros tipos: {}"
        ).format(total_cat, wall_found, floor_found, family_instance, otros)
        return {"ok": True, "text": msg, "is_error": False}

    grupos = defaultdict(list)
    for elem in fundaciones:
        dims, _ = _get_dimensiones_y_volumen(elem, doc)
        grupos[dims].append(elem)

    numeros_existentes = set()
    for elem in fundaciones:
        num = _leer_numeracion_fundacion(elem)
        if num is not None:
            try:
                numeros_existentes.add(int(num))
            except (ValueError, TypeError):
                pass

    grupos_sin_numero = [
        (dims, elems) for dims, elems in grupos.items()
        if not any(_leer_numeracion_fundacion(e) for e in elems)
    ]
    grupos_sin_numero.sort(
        key=lambda x: (
            x[0][0] * x[0][1] * x[0][2],
            x[0][0],
            x[0][1],
            x[0][2],
        )
    )

    siguiente = (max(numeros_existentes) if numeros_existentes else 0) + 1
    numero_por_grupo = {}
    for dims, elems in grupos_sin_numero:
        numero_por_grupo[dims] = str(siguiente)
        siguiente += 1

    for dims, elems in grupos.items():
        if dims not in numero_por_grupo:
            for e in elems:
                num = _leer_numeracion_fundacion(e)
                if num is not None:
                    numero_por_grupo[dims] = num
                    break

    trans = Transaction(doc, u"Arainco: Numerar fundaciones aisladas")
    try:
        trans.Start()
        numerados = 0
        sin_parametro = 0
        for dims, elementos in grupos.items():
            numero = numero_por_grupo.get(dims)
            if numero is None:
                continue
            for elem in elementos:
                if _leer_numeracion_fundacion(elem) is not None:
                    continue
                if _set_numeracion_fundacion(elem, numero):
                    numerados += 1
                else:
                    sin_parametro += 1

        trans.Commit()

        lineas_reporte = []
        for dims, elementos in grupos.items():
            numero = numero_por_grupo.get(dims)
            if numero is None:
                continue
            l_mm, w_mm, h_mm = dims
            dims_str = u"{:.2f} x {:.2f} x {:.2f} m".format(
                l_mm / 1000.0, w_mm / 1000.0, h_mm / 1000.0
            )
            cant = len(elementos)
            cant_str = (
                u"{} unidad".format(cant)
                if cant == 1
                else u"{} unidades".format(cant)
            )
            lineas_reporte.append((numero, dims_str, cant_str))

        def _orden_numero(item):
            try:
                return int(item[0])
            except (ValueError, TypeError):
                return 9999

        lineas_reporte.sort(key=_orden_numero)

        if numerados == 0 and sin_parametro == 0:
            msg = (
                u"No hay fundaciones sin numerar. "
                u"Todas tienen valor en 'Numeracion Fundacion'."
            )
        else:
            msg = u"Se numeraron {} fundación(es) sin numerar.".format(numerados)
            if numerados > 0:
                msg += (
                    u"\n\nSe agruparon con las ya numeradas que "
                    u"comparten las mismas dimensiones."
                )
        if sin_parametro:
            msg += (
                u"\n\n{} elemento(s) no tienen parámetro "
                u"'Numeracion Fundacion' editable."
            ).format(sin_parametro)

        if lineas_reporte:
            msg += u"\n\n--- Listado de fundaciones numeradas ---\n"
            for numero, dims_str, cant_str in lineas_reporte:
                msg += u"\nF{}: {} - {}".format(numero, dims_str, cant_str)

        return {"ok": True, "text": msg, "is_error": False}

    except Exception as ex:
        if trans.HasStarted():
            trans.RollBack()
        return {"ok": False, "text": u"Error:\n{0}".format(ex), "is_error": True}


# ── Ventana ──────────────────────────────────────────────────────────────────


def _get_active_window():
    try:
        win = AppDomain.CurrentDomain.GetData(_APPDOMAIN_WINDOW_KEY)
    except Exception:
        return None
    if win is None:
        return None
    try:
        _ = win.Title
    except Exception:
        _clear_active_window()
        return None
    try:
        if hasattr(win, "IsLoaded") and (not win.IsLoaded):
            _clear_active_window()
            return None
    except Exception:
        pass
    return win


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


def _brush_error():
    try:
        return SolidColorBrush(Color.FromRgb(255, 120, 120))
    except Exception:
        return None


def _brush_ok():
    try:
        return SolidColorBrush(Color.FromRgb(200, 228, 239))
    except Exception:
        return None


def run(revit):
    """Punto de entrada pyRevit: muestra la ventana."""
    existing = _get_active_window()
    if existing is not None:
        try:
            if existing.WindowState == WindowState.Minimized:
                existing.WindowState = WindowState.Normal
        except Exception:
            pass
        try:
            existing.Activate()
            existing.Focus()
        except Exception:
            pass
        _mostrar_aviso(revit, u"La herramienta ya esta en ejecucion.")
        return

    uidoc = revit.ActiveUIDocument
    if uidoc is None:
        _mostrar_aviso(revit, u"No hay documento activo.")
        return
    doc = uidoc.Document

    win = XamlReader.Parse(_build_xaml())
    txt_subtitle = win.FindName(u"TxtSubtitle")
    txt_status = win.FindName(u"TxtStatus")
    txt_result = win.FindName(u"TxtResult")
    panel_resultado = win.FindName(u"PanelResultado")
    br_err = _brush_error()
    br_ok = _brush_ok()

    if txt_subtitle is not None:
        try:
            txt_subtitle.Text = (
                u"Agrupa fundaciones aisladas por dimensiones y escribe "
                u"Numeracion Fundacion."
            )
        except Exception:
            pass

    if panel_resultado is not None:
        try:
            panel_resultado.Visibility = Visibility.Collapsed
        except Exception:
            pass

    def _set_status(text):
        if txt_status is not None:
            try:
                txt_status.Text = _as_unicode(text)
            except Exception:
                pass

    def _on_close(sender, args):
        _clear_active_window()

    def _on_key_down(sender, args):
        if args.Key == Key.Escape:
            win.Close()

    def _on_btn_close(sender, args):
        win.Close()

    def _on_numerar(sender, args):
        _set_status(u"Numerando…")
        res = execute_numerar_fundaciones(doc)
        text = res.get("text") or u""
        if txt_result is not None:
            try:
                txt_result.Text = text
            except Exception:
                pass
            try:
                if res.get("is_error") and br_err is not None:
                    txt_result.Foreground = br_err
                elif br_ok is not None:
                    txt_result.Foreground = br_ok
            except Exception:
                pass
        if panel_resultado is not None:
            try:
                panel_resultado.Visibility = Visibility.Visible
            except Exception:
                pass
        if res.get("is_error"):
            _set_status(u"Error al numerar.")
        else:
            first_line = text.split(u"\n")[0] if text else u"Listo."
            _set_status(first_line)

    from System.Windows import RoutedEventHandler

    win.Closed += EventHandler(_on_close)
    win.KeyDown += KeyEventHandler(_on_key_down)

    btn_close = win.FindName(u"BtnClose")
    if btn_close is not None:
        btn_close.Click += RoutedEventHandler(_on_btn_close)

    btn_num = win.FindName(u"BtnNumerar")
    if btn_num is not None:
        btn_num.Click += RoutedEventHandler(_on_numerar)

    _prepare_window(win, revit)
    _set_status(u"Listo.")
    _set_active_window(win)
    try:
        win.ShowDialog()
    finally:
        _clear_active_window()
