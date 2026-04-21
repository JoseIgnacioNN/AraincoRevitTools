# -*- coding: utf-8 -*-
"""
Armadura Columnas V2 — misma base de UI que enfierrado columnas (``enfierrado_vigas``); lógica
independiente en este módulo para evolucionar sin alterar ``enfierrado_columnas``.

**Troceo de ``ModelLine`` (planos de empalme):** en esta herramienta debe existir una sola
operación lógica, centralizada en ``_ejecutar_troceo_empalme_columnas_v2_unico``. No invocar
``trocear_model_lines_con_planos_sketch_v2`` desde otros sitios del flujo Columnas V2.
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

from barras_bordes_losa_gancho_empotramiento import (
    _rebar_nominal_diameter_mm,
    _task_dialog_show,
    element_id_to_int,
)
from geometria_columnas_eje import (
    crear_sketch_planes_empalme_desde_location_curve,
    estimar_largo_max_mm_eje_columnas_fallback_ubicacion,
    estimar_largo_max_mm_eje_columnas_fusionado,
)
from geometria_columnas_v2_caras import ejecutar_v2_model_lines_cara_ancho
from troceo_model_curves_planos_empalme_v2 import (
    trocear_model_lines_con_planos_sketch_v2,
)

_EXT_ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), os.pardir))
_PUSHBUTTON_DIR = os.path.join(
    _EXT_ROOT,
    "BIMTools.tab",
    "Armadura.panel",
    "25_ArmaduraColumnasV2.pushbutton",
)
_LOGO_PATHS_COLUMNAS_V2 = [
    os.path.join(_PUSHBUTTON_DIR, "empresa_logo.png"),
    os.path.join(_PUSHBUTTON_DIR, "logo_empresa.png"),
    os.path.join(_PUSHBUTTON_DIR, "logo.png"),
]

_APPDOMAIN_WINDOW_KEY = "BIMTools.ArmaduraColumnasV2.ActiveWindow"


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


def _diametros_estribo_y_long_nominales_mm_desde_ventana(win):
    """Ø nominal (mm) del estribo y del longitudinal activo — para el paso entre líneas Cara A."""
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
    return d_est, d_long


def _xaml_columnas_desde_vigas():
    """Misma plantilla WPF que vigas, con títulos adaptados a columnas."""
    x = ev._ENFIERRADO_VIGAS_XAML
    x = x.replace(
        'Title="Arainco - Armadura vigas"',
        'Title="Arainco - Armadura columnas V2"',
    )
    x = x.replace(
        'Text="Armadura Vigas"',
        'Text="Armadura columnas V2"',
    )
    x = x.replace(
        '      <StackPanel x:Name="PnlPieForm" Margin="0,12,0,0" '
        'HorizontalAlignment="Stretch">\n'
        '        <Button x:Name="BtnColocar" Content="Colocar armadura"',
        '      <StackPanel x:Name="PnlPieForm" Margin="0,12,0,0" '
        'HorizontalAlignment="Stretch">\n'
        '        <TextBlock TextWrapping="Wrap" MaxWidth="420" Foreground="#A8C8D8" FontSize="11" '
        'Margin="0,0,0,6"\n'
        '          Text="A = par de mayor separación; B = menor (canto). '
        'Lados iguales en planta: cantidad nb en todos los lados. Paso según luz ortogonal. Ambas opuestas. '
        'Recubrimiento interior: '
        '25 mm + estribo + mitad Ø longitudinal (combos)."/>'
        '\n        <Grid Margin="0,0,0,8">\n'
        '          <Grid.ColumnDefinitions>\n'
        '            <ColumnDefinition Width="Auto"/>\n'
        '            <ColumnDefinition Width="76"/>\n'
        '            <ColumnDefinition Width="16"/>\n'
        '            <ColumnDefinition Width="Auto"/>\n'
        '            <ColumnDefinition Width="76"/>\n'
        '          </Grid.ColumnDefinitions>\n'
        '          <TextBlock Grid.Column="0" Text="Cara A (ancho)" Style="{StaticResource Label}" VerticalAlignment="Center" Margin="0,0,8,0"/>\n'
        '          <TextBox x:Name="TxtBarrasCaraAncho" Grid.Column="1" Text="2" Style="{StaticResource CantSpinnerText}"/>\n'
        '          <TextBlock Grid.Column="3" Text="Cara B (alto)" Style="{StaticResource Label}" VerticalAlignment="Center" Margin="0,0,8,0"/>\n'
        '          <TextBox x:Name="TxtBarrasCaraAlto" Grid.Column="4" Text="2" Style="{StaticResource CantSpinnerText}"/>\n'
        '        </Grid>\n'
        '        <CheckBox x:Name="ChkSegundaCapaColumnasV2" '
        'Content="Segunda capa: offset interior + paso entre líneas de la cara ortogonal (A↔B)" '
        'IsChecked="False" Foreground="#C8E4EF" Margin="0,0,0,4"/>\n'
        '        <CheckBox x:Name="ChkTerceraCapaColumnasV2" '
        'Content="Tercera capa: +2× ese paso (incluye 2.ª aunque no esté marcada)" '
        'IsChecked="False" Foreground="#C8E4EF" Margin="0,0,0,8"/>\n'
        '        <Button x:Name="BtnColocar" Content="Colocar armadura"',
    )
    x = x.replace(
        'Content="Seleccionar vigas para realizar empalmes"',
        'Content="Seleccionar columnas para empalmes"',
    )
    x = x.replace(
        '<TextBlock Text="Trazo inicial mayor a 12 metros" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,0,0,6"/>',
        '<TextBlock Text="Eje de columna fusionado mayor a 12 m" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,0,0,6"/>',
    )
    x = x.replace(
        'Text="Opcional: otras vigas para planos de empalme (centro de eje). Si no pulsa el botón, se usan todas las vigas de la selección inicial del modelo."',
        'Text="Opcional: otras columnas estructurales para empalme (eje > 12 m). El botón solo permite columnas; la selección inicial del comando también son columnas."',
    )
    return x


def _barras_cara_ancho_alto_desde_ventana(win):
    """Enteros (ancho, alto) desde ``TxtBarrasCaraAncho`` / ``TxtBarrasCaraAlto``."""
    lo = int(getattr(ev, "_CANTIDAD_BARRAS_MIN", 1))
    hi = int(getattr(ev, "_CANTIDAD_BARRAS_MAX", 99))
    na, nb = 2, 2
    try:
        tba = win._win.FindName("TxtBarrasCaraAncho")
        if tba is not None:
            s = (tba.Text or "").strip()
            if s:
                na = int(float(s.replace(",", ".")))
    except Exception:
        pass
    try:
        tbb = win._win.FindName("TxtBarrasCaraAlto")
        if tbb is not None:
            s = (tbb.Text or "").strip()
            if s:
                nb = int(float(s.replace(",", ".")))
    except Exception:
        pass
    na = max(lo, min(hi, na))
    nb = max(lo, min(hi, nb))
    return na, nb


class SeleccionarSoloColumnasV2Filter(ISelectionFilter):
    """Solo columnas estructurales (Structural Columns)."""

    def AllowElement(self, elem):
        try:
            if elem is None or elem.Category is None:
                return False
            return (
                element_id_to_int(elem.Category.Id)
                == int(BuiltInCategory.OST_StructuralColumns)
            )
        except Exception:
            return False

    def AllowReference(self, ref, pt):
        return False


class SeleccionarSoloColumnasV2Handler(IExternalEventHandler):
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
        flt = SeleccionarSoloColumnasV2Filter()
        try:
            refs = list(
                uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    flt,
                    u"Seleccione columnas estructurales. Finalice con Finalizar.",
                )
            )
        except Exception:
            refs = []
            win._set_estado(u"Selección cancelada.")
            try:
                win._show_with_fade()
            except Exception:
                pass
            return
        if not refs:
            win._set_estado(u"Sin elementos.")
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
        win._document = doc
        win._selected_element_ids = ids
        win._refresh_selection_text()
        try:
            win._refresh_empalmes_panel_from_selection()
        except Exception:
            pass
        win._set_estado(u"{0} columna(s) seleccionada(s).".format(len(ids)))
        try:
            win._show_with_fade()
        except Exception:
            pass

    def GetName(self):
        return u"SeleccionarColumnasV2"


class EmpalmeSoloColumnasEstructuralesV2Filter(ISelectionFilter):
    """Solo Structural Columns (empalmes en herramienta Columnas V2)."""

    def AllowElement(self, elem):
        try:
            if elem is None or elem.Category is None:
                return False
            c = element_id_to_int(elem.Category.Id)
            return c == int(BuiltInCategory.OST_StructuralColumns)
        except Exception:
            return False

    def AllowReference(self, ref, pt):
        return False


def _integer_values_from_element_ids(id_list):
    out = set()
    for eid in id_list or []:
        try:
            out.add(int(eid.IntegerValue))
        except Exception:
            pass
    return out


def _model_line_ids_para_troceo_empalme_v2(win):
    """Líneas de modelo de la herramienta menos marcadores de normal del pick de empalme."""
    ml = list(getattr(win, "_model_line_ids", None) or [])
    excl = _integer_values_from_element_ids(
        getattr(win, "_empalme_pick_model_line_ids", None) or []
    )
    out = []
    for eid in ml:
        try:
            k = int(eid.IntegerValue)
        except Exception:
            continue
        if k not in excl:
            out.append(eid)
    return out


def _actualizar_model_line_ids_tras_troceo_v2(win, ids_eliminados, ids_nuevos):
    dead = _integer_values_from_element_ids(ids_eliminados)
    ml = []
    for eid in getattr(win, "_model_line_ids", None) or []:
        try:
            if int(eid.IntegerValue) not in dead:
                ml.append(eid)
        except Exception:
            ml.append(eid)
    for eid in ids_nuevos or []:
        ml.append(eid)
    win._model_line_ids = ml


def _append_unique_element_ids(base_list, extra_ids):
    existing = _integer_values_from_element_ids(base_list)
    for e in extra_ids or []:
        try:
            k = int(e.IntegerValue)
        except Exception:
            continue
        if k not in existing:
            existing.add(k)
            base_list.append(e)


def _purge_win_list_attr_by_element_ids(win, attr_name, dead_ids):
    if not dead_ids:
        return
    dead = _integer_values_from_element_ids(dead_ids)
    lst = list(getattr(win, attr_name, None) or [])
    kept = []
    for x in lst:
        try:
            if int(x.IntegerValue) in dead:
                continue
        except Exception:
            pass
        kept.append(x)
    setattr(win, attr_name, kept)


def _ejecutar_troceo_empalme_columnas_v2_unico(
    document, win, ids_sketch_planes_empalme, empalme_column_ids
):
    """
    Único lugar que llama a ``trocear_model_lines_con_planos_sketch_v2`` para Armadura Columnas V2.

    Tras el pick de empalme (misma transacción que la creación de esos ``SketchPlane``): regenera,
    resuelve curvas candidatas y actualiza ``_model_line_ids``. **Colocar** no ejecuta troceo; solo
    este camino.
    """
    try:
        document.Regenerate()
    except Exception:
        pass
    curve_ids = _model_line_ids_para_troceo_empalme_v2(win)
    # #region agent log
    try:
        import json
        import time

        _lp = os.path.join(_EXT_ROOT, "debug-c561be.log")
        _n_excl = len(
            list(getattr(win, "_empalme_pick_model_line_ids", None) or [])
        )
        _prev_c = []
        for _x in (curve_ids or [])[:6]:
            try:
                _prev_c.append(int(_x.IntegerValue))
            except Exception:
                pass
        with open(_lp, "a") as _lf:
            _lf.write(
                json.dumps(
                    {
                        u"sessionId": u"c561be",
                        u"hypothesisId": u"H2",
                        u"location": u"empalme_v2:curve_ids_para_troceo",
                        u"message": u"troceo curve_ids from window",
                        u"data": {
                            u"n_curve_ids": len(curve_ids or []),
                            u"n_empalme_pick_ml_excl": _n_excl,
                            u"first_curve_ids": _prev_c,
                        },
                        u"timestamp": int(time.time() * 1000),
                    },
                    ensure_ascii=False,
                )
                + u"\n"
            )
    except Exception:
        pass
    # #endregion
    cols_fb = []
    _seen_col = set()
    for eid in list(empalme_column_ids or []) + list(
        getattr(win, "_selected_element_ids", None) or []
    ):
        k = element_id_to_int(eid)
        if k is None or k in _seen_col:
            continue
        _seen_col.add(k)
        cols_fb.append(eid)
    troceo_msg, _nv, _vj, troceo_diag = trocear_model_lines_con_planos_sketch_v2(
        document,
        curve_ids,
        ids_sketch_planes_empalme,
        column_ids_for_fallback=cols_fb,
    )
    try:
        troceo_diag = (
            u"— Contexto transacción —\n"
            u"SketchPlane empalme (nuevos en esta operación): {0}\n"
            u"IDs en _model_line_ids para troceo (excl. marcadores N): {1}\n"
            u"Columnas para respaldo bbox (selección ∪ empalme): {2}\n\n{3}"
        ).format(
            len(ids_sketch_planes_empalme or []),
            len(curve_ids),
            len(cols_fb),
            troceo_diag,
        )
    except Exception:
        pass
    try:
        _actualizar_model_line_ids_tras_troceo_v2(win, _vj, _nv)
    except Exception:
        pass
    return troceo_msg, troceo_diag


def _reemplazar_sketch_planes_pick_empalme_columnas_v2(document, win, empalme_column_ids):
    """
    Quita planos (y marcadores) del pick de empalme anterior y crea un ``SketchPlane`` por columna
    con origen en ``GetEndPoint(0)`` de la ``LocationCurve``.
    """
    prev_sp = list(getattr(win, "_empalme_pick_sketch_plane_ids", None) or [])
    prev_ml = list(getattr(win, "_empalme_pick_model_line_ids", None) or [])
    dead = list(prev_sp) + list(prev_ml)

    ids_sp = []
    ids_mk = []
    t = Transaction(document, u"BIMTools sketch planes empalme columnas")
    t.Start()
    try:
        _n_ml_pre_purge = len(list(getattr(win, "_model_line_ids", None) or []))
        for eid in dead:
            try:
                el = document.GetElement(eid)
            except Exception:
                el = None
            if el is not None:
                try:
                    document.Delete(eid)
                except Exception:
                    pass
        _purge_win_list_attr_by_element_ids(win, "_sketch_plane_ids", dead)
        _purge_win_list_attr_by_element_ids(win, "_model_line_ids", dead)
        # #region agent log
        try:
            import json
            import time

            _lp = os.path.join(_EXT_ROOT, "debug-c561be.log")
            _n_ml_post_purge = len(
                list(getattr(win, "_model_line_ids", None) or [])
            )
            with open(_lp, "a") as _lf:
                _lf.write(
                    json.dumps(
                        {
                            u"sessionId": u"c561be",
                            u"hypothesisId": u"H1",
                            u"location": u"empalme_v2:after_purge",
                            u"message": u"model_line_ids around empalme purge",
                            u"data": {
                                u"n_ml_pre_purge": _n_ml_pre_purge,
                                u"n_ml_post_purge": _n_ml_post_purge,
                                u"n_dead": len(dead),
                            },
                            u"timestamp": int(time.time() * 1000),
                        },
                        ensure_ascii=False,
                    )
                    + u"\n"
                )
        except Exception:
            pass
        # #endregion
        ids_sp, _planes, ids_mk = crear_sketch_planes_empalme_desde_location_curve(
            document,
            empalme_column_ids,
            crear_marcador_normal_primer_plano=False,
            crear_marcador_normal_cada_plano=True,
        )
        troceo_msg = u""
        troceo_diag = u""
        if not ids_sp:
            troceo_diag = (
                u"[0] No se creó ningún SketchPlane de empalme "
                u"(revisar ``LocationCurve`` / columnas elegidas).\n"
            )
        if ids_sp:
            troceo_msg, troceo_diag = _ejecutar_troceo_empalme_columnas_v2_unico(
                document, win, ids_sp, empalme_column_ids
            )
        t.Commit()
    except Exception as ex:
        try:
            t.RollBack()
        except Exception:
            pass
        try:
            err = ex.Message
        except Exception:
            err = str(ex)
        return 0, 0, err, u"", u""

    win._empalme_pick_sketch_plane_ids = list(ids_sp)
    win._empalme_pick_model_line_ids = list(ids_mk)
    all_sp = list(getattr(win, "_sketch_plane_ids", None) or [])
    all_ml = list(getattr(win, "_model_line_ids", None) or [])
    _append_unique_element_ids(all_sp, ids_sp)
    _append_unique_element_ids(all_ml, ids_mk)
    win._sketch_plane_ids = all_sp
    win._model_line_ids = all_ml
    return len(ids_sp), len(ids_mk), None, troceo_msg, troceo_diag


class PickEmpalmeElementosColumnasHandler(IExternalEventHandler):
    """Selección solo de columnas estructurales para empalme cuando el eje supera 12 m."""

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
        flt = EmpalmeSoloColumnasEstructuralesV2Filter()
        try:
            refs = list(
                uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    flt,
                    u"Seleccione columnas estructurales para empalmes. Finalice con Finalizar.",
                )
            )
        except Exception:
            refs = []
            win._set_estado(u"Selección de columnas de empalme cancelada.")
            try:
                win._show_with_fade()
            except Exception:
                pass
            return
        if not refs:
            win._set_estado(u"Sin columnas de empalme.")
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
        n_planos, n_normales, err_pl, troceo_msg, troceo_diag = (
            _reemplazar_sketch_planes_pick_empalme_columnas_v2(doc, win, ids)
        )
        if err_pl:
            try:
                win._set_estado(
                    u"{0} columna(s) empalme; planos: error — {1}".format(len(ids), err_pl)
                )
            except Exception:
                pass
        else:
            try:
                txt = (
                    u"{0} columna(s) empalme; {1} plano(s), {2} ModelLine(s) de normal.".format(
                        len(ids), n_planos, n_normales
                    )
                )
                if troceo_msg:
                    txt = txt + u" " + troceo_msg
                txt = txt + u" (Revise el cuadro «Diagnóstico troceo V2» si no hubo cortes.)"
                win._set_estado(txt)
            except Exception:
                pass
        if not err_pl and troceo_diag:
            try:
                _task_dialog_show(
                    u"BIMTools — Diagnóstico troceo V2",
                    troceo_diag,
                    win._win,
                )
            except Exception:
                pass
        try:
            win._show_with_fade()
        except Exception:
            pass

    def GetName(self):
        return u"PickEmpalmeElementosColumnas"


class ColocarV2CaraAnchoModelLineHandler(IExternalEventHandler):
    """Crea ``ModelCurve`` en cara ancho (fusión + marcador); sin ``Rebar`` en esta fase."""

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
        d_est_mm, d_long_mm = _diametros_estribo_y_long_nominales_mm_desde_ventana(win)
        na, nb = _barras_cara_ancho_alto_desde_ventana(win)
        chk_2 = win._win.FindName("ChkSegundaCapaColumnasV2")
        chk_3 = win._win.FindName("ChkTerceraCapaColumnasV2")
        segunda_capa = chk_2 is not None and chk_2.IsChecked == True
        tercera_capa = chk_3 is not None and chk_3.IsChecked == True
        msg, ids_eje, ids_marc_norm, ids_planos = ejecutar_v2_model_lines_cara_ancho(
            doc,
            uidoc,
            ids,
            na,
            nb,
            off_mm,
            diam_estribo_mm=d_est_mm,
            diam_longitudinal_mm=d_long_mm,
            segunda_capa=segunda_capa,
            tercera_capa=tercera_capa,
        )
        # #region agent log
        try:
            import json
            import time

            _lp = os.path.join(_EXT_ROOT, "debug-c561be.log")
            with open(_lp, "a") as _lf:
                _lf.write(
                    json.dumps(
                        {
                            u"sessionId": u"c561be",
                            u"hypothesisId": u"H1b",
                            u"location": u"colocar_v2:after_ejecutar",
                            u"message": u"Colocar V2 axis ids from geometry",
                            u"data": {
                                u"n_ids_eje": len(list(ids_eje or [])),
                                u"n_selected_cols": len(list(ids or [])),
                            },
                            u"timestamp": int(time.time() * 1000),
                        },
                        ensure_ascii=False,
                    )
                    + u"\n"
                )
        except Exception:
            pass
        # #endregion
        try:
            prev_ml = list(getattr(win, "_model_line_ids", None) or [])
            prev_mk = list(getattr(win, "_model_line_marker_ids", None) or [])
            ex_eje = list(ids_eje or [])
            ex_mk = list(ids_marc_norm or [])
            if ex_eje:
                win._model_line_ids = prev_ml + ex_eje
            if ex_mk:
                win._model_line_marker_ids = prev_mk + ex_mk
            if ids_planos:
                prev_sp = list(getattr(win, "_sketch_plane_ids", None) or [])
                for sid in ids_planos:
                    if sid not in prev_sp:
                        prev_sp.append(sid)
                win._sketch_plane_ids = prev_sp
        except Exception:
            pass
        try:
            win._set_estado(msg)
        except Exception:
            pass

    def GetName(self):
        return u"ColocarV2CaraAnchoModelLine"


class ArmaduraColumnasV2Window(ev.EnfierradoVigasWindow):
    """Misma ventana que vigas; reemplaza el evento «Colocar» por eje + ModelLine (V2)."""

    def __init__(self, revit):
        self._model_line_ids = []
        self._model_line_marker_ids = []
        self._sketch_plane_ids = []
        self._empalme_pick_sketch_plane_ids = []
        self._empalme_pick_model_line_ids = []
        ev.EnfierradoVigasWindow.__init__(
            self,
            revit,
            xaml_string=_xaml_columnas_desde_vigas(),
            logo_paths=_LOGO_PATHS_COLUMNAS_V2,
            appdomain_window_key=_APPDOMAIN_WINDOW_KEY,
            tool_title_short=u"Armadura Columnas V2",
        )
        self._seleccion_handler = SeleccionarSoloColumnasV2Handler(weakref.ref(self))
        self._seleccion_event = ExternalEvent.Create(self._seleccion_handler)
        self._empalme_pick_handler = PickEmpalmeElementosColumnasHandler(
            weakref.ref(self)
        )
        self._empalme_pick_event = ExternalEvent.Create(self._empalme_pick_handler)
        self._colocar_handler = ColocarV2CaraAnchoModelLineHandler(weakref.ref(self))
        self._colocar_event = ExternalEvent.Create(self._colocar_handler)
        try:
            from System.Windows.Controls import TextChangedEventHandler

            def _on_refresh_empalmes_por_texto(s, a):
                try:
                    self._refresh_empalmes_panel_from_selection()
                except Exception:
                    pass

            for _nm in (
                "TxtBarrasCaraAncho",
                "TxtBarrasCaraAlto",
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
            L = estimar_largo_max_mm_eje_columnas_fusionado(
                doc, ids, off, None
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
        self._set_estado(u"En cola: líneas de modelo (cara ancho)…")


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
                u"BIMTools — Armadura Columnas V2",
                u"La herramienta ya está en ejecución.",
                existing,
            )
            return

    w = ArmaduraColumnasV2Window(revit)
    try:
        w.show()
    except Exception:
        ev._clear_appdomain_window_key(_APPDOMAIN_WINDOW_KEY)
        raise
