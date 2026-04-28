# -*- coding: utf-8 -*-
"""
Enfierrado columnas — misma UI que ``enfierrado_vigas``; «Colocar armadura» genera **Structural
Rebar** según el eje fusionado (curvas procesadas) y planos auxiliares; no crea ``ModelLine``. **Varias capas de eje**: con
``_COLUMNAS_MODELO_LINE_MULTICAPA_ACTIVA`` activo, el recuadro «Capas» del 1.er segmento (texto
``TxtNumCapasSuperiores``, 2 o 3) envía varios anillos a ``ejecutar_model_lines_eje_columnas``
(aunque la 1.ª capa tenga 4 u 8 barras; la lógica de vigas limita ese control a 1 solo si
``n_capa1 < 12``, por eso aquí se lee el valor del control directamente). Si el campo manual de
curvas de eje tiene texto, se usa ese total como anillo único (sin multicapa). Si el eje supera
12 m se muestra el bloque de empalmes (informativo).
"""

import os
import sys
import weakref

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import BuiltInCategory, Transaction
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from Autodesk.Revit.UI.Selection import ISelectionFilter

_scripts_dir = os.path.dirname(os.path.abspath(__file__))
if _scripts_dir not in sys.path:
    sys.path.insert(0, _scripts_dir)

import enfierrado_vigas as ev

from bimtools_paths import get_logo_paths
from barras_bordes_losa_gancho_empotramiento import (
    _rebar_nominal_diameter_mm,
    _task_dialog_show,
    element_id_to_int,
)
from geometria_columnas_eje import (
    ejecutar_model_lines_eje_columnas,
    estimar_largo_max_mm_eje_columnas_fallback_ubicacion,
    estimar_largo_max_mm_eje_columnas_fusionado,
)

_APPDOMAIN_WINDOW_KEY = "BIMTools.EnfierradoColumnas.ActiveWindow"
# ``True``: 2.ª/3.ª anillo de ModelLine cuando «Capas» superior = 2 o 3 (véase
# ``_capas_num_curvas_eje_desde_ventana``).
_COLUMNAS_MODELO_LINE_MULTICAPA_ACTIVA = True


def _offset_mm_recubrimiento_desde_ventana(win):
    """
    Distancia desde la cara proyectada hacia el interior: 25 mm + Ø estribo (extremos)
    + mitad del Ø de barra longitudinal (combos activos de la misma UI que vigas).
    """
    doc = getattr(win, "_document", None)
    if doc is None:
        try:
            doc = win._revit.ActiveUIDocument.Document
        except Exception:
            doc = None
    if doc is not None and not getattr(win, "_entries", None):
        try:
            win._document = doc
            win._cargar_combos_diametro()
        except Exception:
            pass

    d_est = 0.0
    bt_e = ev._rebar_bar_type_desde_combo_diam(win, "CmbEstriboExtDiam")
    if bt_e is not None:
        try:
            d_est = float(_rebar_nominal_diameter_mm(bt_e) or 0.0)
        except Exception:
            d_est = 0.0

    d_long = 0.0
    chk_sup = win._win.FindName("ChkSuperior")
    chk_inf = win._win.FindName("ChkInferior")
    sup_on = chk_sup is None or chk_sup.IsChecked == True
    inf_on = chk_inf is None or chk_inf.IsChecked == True
    bt_long = None
    if sup_on:
        bt_long = ev._rebar_bar_type_desde_combo_diam(win, "CmbSupDiam")
    if bt_long is None and inf_on:
        bt_long = ev._rebar_bar_type_desde_combo_diam(win, "CmbInfDiam")
    if bt_long is None:
        bt_long = ev._rebar_bar_type_desde_combo_diam(win, "CmbSupDiam")
    if bt_long is None:
        bt_long = ev._rebar_bar_type_desde_combo_diam(win, "CmbInfDiam")
    if bt_long is not None:
        try:
            d_long = float(_rebar_nominal_diameter_mm(bt_long) or 0.0)
        except Exception:
            d_long = 0.0

    return 25.0 + d_est + 0.5 * d_long


def _rebar_bar_type_longitudinal_desde_ventana(win):
    """
    ``RebarBarType`` longitudinal según combos activos (Sup/Inf). None si no hay tipo.
    """
    doc = getattr(win, "_document", None)
    if doc is None:
        try:
            doc = win._revit.ActiveUIDocument.Document
        except Exception:
            doc = None
    if doc is not None and not getattr(win, "_entries", None):
        try:
            win._document = doc
            win._cargar_combos_diametro()
        except Exception:
            pass
    chk_sup = win._win.FindName("ChkSuperior")
    chk_inf = win._win.FindName("ChkInferior")
    sup_on = chk_sup is None or chk_sup.IsChecked == True
    inf_on = chk_inf is None or chk_inf.IsChecked == True
    bt_long = None
    if sup_on:
        bt_long = ev._rebar_bar_type_desde_combo_diam(win, "CmbSupDiam")
    if bt_long is None and inf_on:
        bt_long = ev._rebar_bar_type_desde_combo_diam(win, "CmbInfDiam")
    if bt_long is None:
        bt_long = ev._rebar_bar_type_desde_combo_diam(win, "CmbSupDiam")
    if bt_long is None:
        bt_long = ev._rebar_bar_type_desde_combo_diam(win, "CmbInfDiam")
    return bt_long


def _diametro_longitudinal_mm_desde_ventana(win):
    """Ø nominal (mm) de la barra longitudinal según combos activos (misma lógica que el offset)."""
    d_long = 0.0
    bt_long = _rebar_bar_type_longitudinal_desde_ventana(win)
    if bt_long is not None:
        try:
            d_long = float(_rebar_nominal_diameter_mm(bt_long) or 0.0)
        except Exception:
            d_long = 0.0
    return d_long


def _diametro_estribo_mm_desde_ventana(win):
    """Ø nominal (mm) del estribo desde ``CmbEstriboExtDiam`` (mismo tipo que el recubrimiento UI)."""
    doc = getattr(win, "_document", None)
    if doc is None:
        try:
            doc = win._revit.ActiveUIDocument.Document
        except Exception:
            doc = None
    if doc is not None and not getattr(win, "_entries", None):
        try:
            win._document = doc
            win._cargar_combos_diametro()
        except Exception:
            pass
    d_est = 0.0
    bt_e = ev._rebar_bar_type_desde_combo_diam(win, "CmbEstriboExtDiam")
    if bt_e is not None:
        try:
            d_est = float(_rebar_nominal_diameter_mm(bt_e) or 0.0)
        except Exception:
            d_est = 0.0
    return d_est


def _xaml_columnas_desde_vigas():
    """Misma plantilla WPF que vigas, con títulos adaptados a columnas."""
    x = ev._ENFIERRADO_VIGAS_XAML
    x = x.replace(
        'Title="Arainco - Armadura vigas"',
        'Title="Arainco - Armadura columnas"',
    )
    x = x.replace(
        'Text="Armadura Vigas"',
        'Text="Armadura columnas"',
    )
    x = x.replace(
        '      <StackPanel x:Name="PnlPieForm" Margin="0,12,0,0" '
        'HorizontalAlignment="Stretch">\n'
        '        <Button x:Name="BtnColocar" Content="Colocar armadura"',
        '      <StackPanel x:Name="PnlPieForm" Margin="0,12,0,0" '
        'HorizontalAlignment="Stretch">\n'
        '        <TextBlock TextWrapping="Wrap" MaxWidth="420" Foreground="#A8C8D8" FontSize="11" '
        'Margin="0,0,0,6"\n'
        '          Text="Líneas eje/columna: si el campo siguiente está vacío, se usa la '
        'cantidad de barras de «1ª Capa» (armadura superior). Valor manual aquí anula ese enlace. '
        'Reparto equitativo en caras laterales. '
        'ModelLines de 2.ª/3.ª capa: desactivadas temporalmente (solo se crea la 1.ª capa)."/>'
        '\n        <TextBox x:Name="TxtNumCurvasEjeColumna" Text="" '
        'Style="{StaticResource CantSpinnerText}" Margin="0,0,0,8"/>\n'
        '        <Button x:Name="BtnColocar" Content="Colocar armadura"',
    )
    x = x.replace(
        'Content="Seleccionar vigas para realizar empalmes"',
        'Content="Seleccionar elementos para empalmes (vigas o columnas)"',
    )
    x = x.replace(
        '<TextBlock Text="Trazo inicial mayor a 12 metros" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,0,0,6"/>',
        '<TextBlock Text="Eje de columna fusionado mayor a 12 m" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,0,0,6"/>',
    )
    x = x.replace(
        'Text="Opcional: otras vigas para planos de empalme (centro de eje). Si no pulsa el botón, se usan todas las vigas de la selección inicial del modelo."',
        'Text="Opcional: otras vigas o columnas para planos de empalme. Si no pulsa el botón, se usan las vigas Structural Framing de la selección inicial."',
    )
    return x


def _num_curvas_eje_desde_ventana(win):
    """
    Total de curvas de eje por columna.

    - Si ``TxtNumCurvasEjeColumna`` tiene texto: ese entero (override).
    - Si está vacío: misma cantidad que **1ª Capa** armadura superior (``CmbSupCant``), para pruebas alineadas con el armado.
    - Si no hay valor válido en ninguno: ``None`` (una línea por cara planar).
    """
    try:
        tb_ov = win._win.FindName("TxtNumCurvasEjeColumna")
    except Exception:
        tb_ov = None
    if tb_ov is not None:
        try:
            s_ov = (tb_ov.Text or "").strip()
        except Exception:
            s_ov = ""
        if s_ov:
            try:
                n_ov = int(float(s_ov.replace(",", ".")))
                return n_ov if n_ov >= 1 else None
            except Exception:
                pass
    try:
        tb_sup = win._win.FindName("CmbSupCant")
    except Exception:
        tb_sup = None
    if tb_sup is None:
        return None
    try:
        s = (tb_sup.Text or "").strip()
    except Exception:
        s = ""
    if not s:
        return None
    try:
        n = int(float(s.replace(",", ".")))
    except Exception:
        return None
    lo = int(getattr(ev, "_CANTIDAD_BARRAS_MIN", 1))
    hi = int(getattr(ev, "_CANTIDAD_BARRAS_MAX", 99))
    n = max(lo, min(hi, n))
    return n if n >= 1 else None


def _capas_num_curvas_eje_desde_ventana(win):
    """
    Lista de longitud ``n_capas`` para ``ejecutar_model_lines_eje_columnas`` (capas múltiples).
    Cuando ``_COLUMNAS_MODELO_LINE_MULTICAPA_ACTIVA`` es ``False``, siempre ``None`` (solo 1.ª capa).
    ``None`` también con override manual, superior desactivado o una sola capa en UI.

    El número de anillos se lee de ``TxtNumCapasSuperiores`` (2 o 3) **sin** aplicar el tope de
    vigas «solo 1 capa si 1.ª < 12 barras» (:func:`enfierrado_vigas._parse_capas_superiores_ventana`),
    porque el eje de columnas puede tener 2.ª capa con 4, 8, 12… curvas.
    """
    if not _COLUMNAS_MODELO_LINE_MULTICAPA_ACTIVA:
        return None
    try:
        tb_ov = win._win.FindName("TxtNumCurvasEjeColumna")
    except Exception:
        tb_ov = None
    if tb_ov is not None:
        try:
            s_ov = (tb_ov.Text or "").strip()
        except Exception:
            s_ov = ""
        if s_ov:
            return None
    chk_sup = win._win.FindName("ChkSuperior")
    sup_on = chk_sup is None or chk_sup.IsChecked == True
    if not sup_on:
        return None
    try:
        win._preparar_lectura_capas_superiores()
    except Exception:
        pass
    try:
        tb_cap = win._win.FindName("TxtNumCapasSuperiores")
        if tb_cap is not None:
            try:
                s = (tb_cap.Text or u"").strip()
            except Exception:
                s = u""
            if s:
                n_cap = int(float(s.replace(u",", u".")))
            else:
                n_cap = 1
        else:
            n_cap = 1
    except Exception:
        n_cap = 1
    n_cap = max(1, min(3, n_cap))
    if n_cap <= 1:
        return None
    try:
        n0 = int(
            ev._parse_cantidad_capa_ventana(
                win, u"CmbSupCant", forzar_combo_habilitado=True
            )
        )
    except Exception:
        n0 = 1
    n0 = max(1, n0)
    return [n0] * int(n_cap)


class EmpalmeColumnasVigasOColumnasFilter(ISelectionFilter):
    """Structural Framing o Structural Columns (planos de empalme / troceo)."""

    def AllowElement(self, elem):
        try:
            if elem is None or elem.Category is None:
                return False
            c = element_id_to_int(elem.Category.Id)
            return c in (
                int(BuiltInCategory.OST_StructuralFraming),
                int(BuiltInCategory.OST_StructuralColumns),
            )
        except Exception:
            return False

    def AllowReference(self, ref, pt):
        return False


class PickEmpalmeElementosColumnasHandler(IExternalEventHandler):
    """Selección de vigas y/o columnas para situar empalmes cuando el eje supera 12 m."""

    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        from Autodesk.Revit.UI.Selection import ObjectType

        win = self._window_ref()
        if win is None:
            return
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            win._set_estado(u"No hay documento activo.")
            return
        doc = uidoc.Document
        flt = EmpalmeColumnasVigasOColumnasFilter()
        try:
            refs = list(
                uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    flt,
                    u"Seleccione vigas o columnas estructurales para empalmes. Finalice con Finalizar.",
                )
            )
        except Exception:
            refs = []
            win._set_estado(u"Selección de elementos de empalme cancelada.")
            try:
                win._show_with_fade()
            except Exception:
                pass
            return
        if not refs:
            win._set_estado(u"Sin elementos de empalme.")
            try:
                win._show_with_fade()
            except Exception:
                pass
            return
        ids = []
        for r in refs:
            try:
                ids.append(r.ElementId)
            except Exception:
                pass
        win._empalme_framing_ids = ids
        win._document = doc
        try:
            win._refresh_empalme_text()
        except Exception:
            pass
        win._set_estado(
            u"{0} elemento(s) para empalmes (eje > 12 m).".format(len(ids))
        )
        try:
            win._show_with_fade()
        except Exception:
            pass

    def GetName(self):
        return u"PickEmpalmeElementosColumnas"


class ColocarEjeColumnasModelLineHandler(IExternalEventHandler):
    """Fusiona ejes colineales de columnas seleccionadas y crea ``Structural Rebar``."""

    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        win = self._window_ref()
        if win is None:
            return
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            try:
                win._set_estado(u"No hay documento activo.")
            except Exception:
                pass
            return
        doc = uidoc.Document
        ids = getattr(win, "_selected_element_ids", None) or []
        off_mm = _offset_mm_recubrimiento_desde_ventana(win)
        capas_nc = _capas_num_curvas_eje_desde_ventana(win)
        n_curvas = (
            None
            if capas_nc is not None
            else _num_curvas_eje_desde_ventana(win)
        )
        d_emp = _diametro_longitudinal_mm_desde_ventana(win)
        d_est = _diametro_estribo_mm_desde_ventana(win)
        emp_ids = list(getattr(win, "_empalme_framing_ids", None) or [])
        bt_long = _rebar_bar_type_longitudinal_desde_ventana(win)
        if bt_long is None:
            try:
                win._set_estado(
                    u"Seleccione diámetro longitudinal (combos Sup/Inf) para crear barras."
                )
            except Exception:
                pass
            return
        _, msg, ids_nuevos, ids_planos, ids_marc_norm = ejecutar_model_lines_eje_columnas(
            doc,
            uidoc,
            ids,
            offset_mm=off_mm,
            num_curvas_eje=n_curvas,
            empotramiento_diam_nominal_mm=d_emp if d_emp > 1e-9 else None,
            diam_estribo_nominal_mm=d_est if d_est > 1e-9 else None,
            empalme_element_ids=emp_ids,
            capas_num_curvas=capas_nc,
            rebar_bar_type=bt_long,
        )
        try:
            prev_ml = list(getattr(win, "_model_line_ids", None) or [])
            extra_ml = list(ids_nuevos or []) + list(ids_marc_norm or [])
            if extra_ml:
                win._model_line_ids = prev_ml + extra_ml
            if ids_planos:
                prev_sp = list(getattr(win, "_sketch_plane_ids", None) or [])
                win._sketch_plane_ids = prev_sp + list(ids_planos)
        except Exception:
            pass
        try:
            win._set_estado(msg)
        except Exception:
            pass

    def GetName(self):
        return u"ColocarEjeColumnasModelLine"


class EliminarModelLinesColumnasHandler(IExternalEventHandler):
    """
    Borra las ``ModelCurve`` y ``SketchPlane`` creados por la herramienta; ejecutar vía
    ``ExternalEvent`` (no desde el cierre directo del WPF).
    """

    def __init__(self):
        self._pending_doc = None
        self._pending_ids = None
        self._pending_sketch_plane_ids = None

    def armar(self, document, ids, sketch_plane_ids=None):
        self._pending_doc = document
        self._pending_ids = list(ids) if ids else []
        self._pending_sketch_plane_ids = (
            list(sketch_plane_ids) if sketch_plane_ids else []
        )

    def Execute(self, uiapp):
        doc = self._pending_doc
        ids = self._pending_ids
        ids_sp = self._pending_sketch_plane_ids
        self._pending_doc = None
        self._pending_ids = None
        self._pending_sketch_plane_ids = None
        if doc is None:
            try:
                uidoc = uiapp.ActiveUIDocument
                if uidoc is not None:
                    doc = uidoc.Document
            except Exception:
                pass
        if doc is None:
            return
        if not ids and not ids_sp:
            return
        try:
            with Transaction(
                doc, u"BIMTools — Quitar planos y líneas eje columnas"
            ) as t:
                t.Start()
                try:
                    for eid in ids_sp or []:
                        try:
                            if doc.GetElement(eid) is not None:
                                doc.Delete(eid)
                        except Exception:
                            pass
                    for eid in ids or []:
                        try:
                            if doc.GetElement(eid) is not None:
                                doc.Delete(eid)
                        except Exception:
                            pass
                except Exception:
                    try:
                        t.RollBack()
                    except Exception:
                        pass
                    return
                t.Commit()
        except Exception:
            pass

    def GetName(self):
        return u"BIMTools — Eliminar ModelLine columnas"


class EnfierradoColumnasWindow(ev.EnfierradoVigasWindow):
    """Misma ventana que vigas; reemplaza el evento «Colocar» por eje + ModelLine."""

    def __init__(self, revit):
        self._model_line_ids = []
        self._sketch_plane_ids = []
        self._eliminar_model_lines_handler = EliminarModelLinesColumnasHandler()
        self._eliminar_model_lines_event = ExternalEvent.Create(
            self._eliminar_model_lines_handler
        )
        ev.EnfierradoVigasWindow.__init__(
            self,
            revit,
            xaml_string=_xaml_columnas_desde_vigas(),
            logo_paths=get_logo_paths(),
            appdomain_window_key=_APPDOMAIN_WINDOW_KEY,
            tool_title_short=u"Enfierrado columnas",
        )
        self._empalme_pick_handler = PickEmpalmeElementosColumnasHandler(
            weakref.ref(self)
        )
        self._empalme_pick_event = ExternalEvent.Create(self._empalme_pick_handler)
        self._colocar_handler = ColocarEjeColumnasModelLineHandler(weakref.ref(self))
        self._colocar_event = ExternalEvent.Create(self._colocar_handler)
        try:
            from System.Windows.Controls import TextChangedEventHandler

            def _on_refresh_empalmes_por_texto(s, a):
                try:
                    self._refresh_empalmes_panel_from_selection()
                except Exception:
                    pass

            for _nm in (
                "TxtNumCurvasEjeColumna",
                "CmbSupCant",
                "CmbSup2Cant",
                "CmbSup3Cant",
                "TxtNumCapasSuperiores",
            ):
                _tb = self._win.FindName(_nm)
                if _tb is not None:
                    _tb.TextChanged += TextChangedEventHandler(_on_refresh_empalmes_por_texto)
        except Exception:
            pass
        try:
            from System.Windows import RoutedEventHandler

            def _on_closed_eliminar_model_lines(sender, args):
                try:
                    self._enqueue_eliminar_model_lines()
                except Exception:
                    pass

            self._win.Closed += RoutedEventHandler(_on_closed_eliminar_model_lines)
        except Exception:
            pass

    def _enqueue_eliminar_model_lines(self):
        ids = list(getattr(self, "_model_line_ids", None) or [])
        ids_sp = list(getattr(self, "_sketch_plane_ids", None) or [])
        self._model_line_ids = []
        self._sketch_plane_ids = []
        if not ids and not ids_sp:
            return
        doc = getattr(self, "_document", None)
        if doc is None:
            try:
                doc = self._revit.ActiveUIDocument.Document
            except Exception:
                doc = None
        if doc is None:
            return
        self._eliminar_model_lines_handler.armar(doc, ids, ids_sp)
        self._eliminar_model_lines_event.Raise()

    def _refresh_empalmes_panel_from_selection(self):
        """Asegura documento activo antes del umbral 12 m (la selección no siempre lo fija)."""
        try:
            if getattr(self, "_document", None) is None:
                self._document = self._revit.ActiveUIDocument.Document
        except Exception:
            pass
        ev.EnfierradoVigasWindow._refresh_empalmes_panel_from_selection(self)

    def _estimar_largo_max_trazo_mm_para_empalmes(self):
        """Mismo umbral 12 m que vigas: largo del eje fusionado de columnas (mm)."""
        doc = getattr(self, "_document", None)
        if doc is None:
            try:
                doc = self._revit.ActiveUIDocument.Document
            except Exception:
                doc = None
        ids = getattr(self, "_selected_element_ids", None) or []
        if doc is None or not ids:
            return None
        L = None
        try:
            off = _offset_mm_recubrimiento_desde_ventana(self)
            n_curvas = _num_curvas_eje_desde_ventana(self)
            L = estimar_largo_max_mm_eje_columnas_fusionado(
                doc, ids, off, n_curvas
            )
        except Exception:
            L = None
        if L is not None:
            return L
        try:
            return estimar_largo_max_mm_eje_columnas_fallback_ubicacion(doc, ids)
        except Exception:
            return None

    def _on_colocar(self, sender, args):
        if not self._selected_element_ids:
            self._set_estado(
                u"Seleccione al menos una columna estructural en el modelo."
            )
            return
        try:
            if getattr(self, "_document", None) is None:
                self._document = self._revit.ActiveUIDocument.Document
        except Exception:
            pass
        try:
            self._refresh_empalmes_panel_from_selection()
        except Exception:
            pass
        self._colocar_event.Raise()
        self._set_estado(u"En cola: eje fusionado y línea de modelo…")


def run_pyrevit(revit):
    if _scripts_dir not in sys.path:
        sys.path.insert(0, _scripts_dir)

    existing = ev._get_active_window(_APPDOMAIN_WINDOW_KEY)
    if existing is not None:
        ok = False
        try:
            from System.Windows import WindowState

            if existing.WindowState == WindowState.Minimized:
                existing.WindowState = WindowState.Normal
            existing.Show()
            existing.Activate()
            existing.Focus()
            ok = True
        except Exception:
            ev._clear_appdomain_window_key(_APPDOMAIN_WINDOW_KEY)
            existing = None
        if ok and existing is not None:
            _task_dialog_show(
                u"BIMTools — Enfierrado columnas",
                u"La herramienta ya está en ejecución.",
                existing,
            )
            return

    w = EnfierradoColumnasWindow(revit)
    try:
        w.show()
    except Exception:
        ev._clear_appdomain_window_key(_APPDOMAIN_WINDOW_KEY)
        raise
