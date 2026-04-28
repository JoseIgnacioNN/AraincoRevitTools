# -*- coding: utf-8 -*-
"""
Enfierrado fundaciones aisladas — formulario (misma línea visual que Crear Area Reinf RPS).

- Estilos WPF, combos y botones alineados con ``area_reinforcement_losa.py`` (Malla en Losa).
- Selección de una sola fundación (categoría Structural Foundation, no zapatas de muro).
- Tres grupos: armadura inferior / superior / lateral; inferior/superior: separación con spinner (100–300 mm, paso 10); lateral: cantidad 1–10 (misma lógica ▲▼ que vigas), default según ``ceil(h/200)−1`` (altura fundación).

Curvas de la cara inferior, detalle en vista y Rebar longitudinal inferior (polilínea U como con
inf.+sup.; solo inferior: largo de pata según tabla ø). Respaldo: barra recta con ganchos estándar.
"""

import math
import os
import re
import sys
import weakref
import clr
import System

_scripts_dir = os.path.dirname(os.path.abspath(__file__))
if _scripts_dir not in sys.path:
    sys.path.insert(0, _scripts_dir)

# Forzar recarga de módulos de geometría para evitar caché obsoleto en pyRevit.
for _m in list(sys.modules.keys()):
    if _m in ("geometria_fundacion_cara_inferior", "vista_seccion_enfierrado_vigas", "geometria_estribos_viga", "enfierrado_shaft_hashtag", "rebar_fundacion_cara_inferior"):
        sys.modules.pop(_m, None)

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from Autodesk.Revit.UI import TaskDialog, ExternalEvent, IExternalEventHandler
from Autodesk.Revit.DB import BuiltInCategory, Line, SubTransaction, Transaction, WallFoundation, XYZ
from Autodesk.Revit.UI.Selection import ISelectionFilter

from barras_bordes_losa_gancho_empotramiento import (
    _build_bar_type_entries,
    _rebar_nominal_diameter_mm,
    element_id_to_int,
    _task_dialog_show,
)
from revit_wpf_window_position import (
    position_wpf_window_top_left_at_active_view,
    revit_main_hwnd,
)

from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_paths import get_logo_paths

from geometria_fundacion_cara_inferior import (
    RECUBRIMIENTO_EXTREMOS_MM,
    aplicar_recubrimiento_inferior_completo_mm,
    evaluar_caras_paralelas_curva_mas_cercana,
    extraer_curva_lado_mayor_cara_inferior,
    extraer_curva_lado_mayor_cara_superior,
    extraer_curva_lado_menor_cara_inferior,
    extraer_curva_lado_menor_cara_superior,
    linea_horizontal_cara_lateral_a_cota_z,
    lineas_horizontales_perimetro_inferior_exterior,
    longitud_distribucion_perpendicular_barra_inferior_ft,
    construir_polilinea_u_fundacion_desde_eje_horizontal,
    longitud_array_lateral_altura_fundacion_menos_mm_ft,
    largo_gancho_u_tabla_mm,
    longitud_pata_u_fundacion_inf_sup_ft,
    normal_saliente_horizontal_paramento_para_barra_horizontal,
    obtener_marco_coordenadas_cara_inferior,
    obtener_marco_coordenadas_cara_superior,
    offset_linea_adicional_hacia_interior_mm,
    offset_linea_eje_barra_desde_cara_inferior_mm,
    offset_linea_hacia_interior_desde_cara_inferior_mm,
    primera_cota_z_armadura_lateral_ft,
    rango_z_caras_laterales_o_bbox,
    vector_reverso_cara_paralela_mas_cercana_a_barra,
)
from rebar_fundacion_cara_inferior import (
    HOOK_GANCHO_90_STANDARD_NAME,
    REBAR_SHAPE_NOMBRE_DEFECTO,
    aplicar_layout_fixed_number_rebar,
    aplicar_layout_maximum_spacing_rebar,
    crear_rebar_desde_curva_linea_con_ganchos,
    crear_rebar_polilinea_recta_sin_ganchos,
    crear_rebar_polilinea_u_malla_inf_sup_curve_loop,
    crear_rebar_u_shape_desde_eje_rebar_shape_nombrado,
)

_APPDOMAIN_WINDOW_KEY = "BIMTools.EnfierradoFundacionAislada.ActiveWindow"

def _parse_diameter_mm_from_bar_combo_label(lbl):
    """Obtiene el diámetro nominal (mm) desde el texto del combo (ø8 mm, ø12 mm  [Id …])."""
    if lbl is None:
        return None
    try:
        s = unicode(lbl)
    except Exception:
        return None
    s = s.replace(u"\u00f8", u" ").replace(u"ø", u" ")
    m = re.search(r"(\d+(?:\.\d+)?)", s)
    if not m:
        return None
    try:
        return float(m.group(1))
    except Exception:
        return None


_SEP_MM_MIN = 100
_SEP_MM_MAX = 300
_SEP_MM_STEP = 10
_SEP_MM_DEFAULT_VAL = 150

_LAT_CANT_MIN = 1
_LAT_CANT_MAX = 10
_LAT_CANT_STEP = 1
_LAT_CANT_DEFAULT_TXT = u"2"

_CAPTION_BTN_COLOCAR_ARMADURAS = u"Colocar armaduras"
_CAPTION_BTN_PROPAGAR_ARMADURAS = u"Propagar Armaduras"


def _numeracion_etiqueta_propagacion(num_val):
    """Etiqueta de numeración en UI (p. ej. parámetro «5» → «F5»)."""
    if num_val is None:
        return u""
    s = unicode(num_val).strip()
    if not s:
        return u""
    if s[0].upper() == u"F":
        return u"F" + s[1:]
    return u"F{0}".format(s)


def _normalize_sep_textbox(tb):
    """Ajusta el ``TextBox`` del spinner al rango 100–300 mm y al paso 10 (como enfierrado_vigas)."""
    if tb is None:
        return
    try:
        s = unicode(tb.Text).replace(u"mm", u"").strip()
        if not s:
            tb.Text = unicode(int(_SEP_MM_DEFAULT_VAL))
            return
        n = int(round(float(s)))
    except Exception:
        tb.Text = unicode(int(_SEP_MM_DEFAULT_VAL))
        return
    n = max(_SEP_MM_MIN, min(_SEP_MM_MAX, n))
    nmax = int((_SEP_MM_MAX - _SEP_MM_MIN) // _SEP_MM_STEP)
    steps = int(round((n - _SEP_MM_MIN) / float(_SEP_MM_STEP)))
    steps = max(0, min(nmax, steps))
    n = _SEP_MM_MIN + steps * _SEP_MM_STEP
    tb.Text = unicode(int(n))


def _normalize_lat_cant_textbox(tb):
    """Acota cantidad de barras laterales (1–10)."""
    if tb is None:
        return
    try:
        s = unicode(tb.Text).strip()
        if not s:
            tb.Text = _LAT_CANT_DEFAULT_TXT
            return
        n = int(float(s.replace(u",", u".")))
    except Exception:
        tb.Text = _LAT_CANT_DEFAULT_TXT
        return
    n = max(_LAT_CANT_MIN, min(_LAT_CANT_MAX, n))
    tb.Text = unicode(n)


def _leer_lat_cant_desde_textbox(tb, default_n):
    if tb is None:
        return int(default_n)
    try:
        s = unicode(tb.Text).strip()
        if not s:
            return int(default_n)
        n = int(round(float(s.replace(u",", u"."))))
    except Exception:
        return int(default_n)
    return max(_LAT_CANT_MIN, min(_LAT_CANT_MAX, n))


def _cantidad_laterales_fundacion_desde_altura_mm(h_mm):
    """
    Misma regla que vigas: ``ceil(h_mm/200) - 1``, mínimo 1; acotada 1–10 en este formulario.
    (In-line para no depender de importar ``enfierrado_vigas`` en pyRevit.)
    """
    if h_mm is None or float(h_mm) <= 0:
        n = int(_LAT_CANT_DEFAULT_TXT)
    else:
        n = int(math.ceil(float(h_mm) / 200.0)) - 1
        n = max(_LAT_CANT_MIN, n)
    return max(_LAT_CANT_MIN, min(_LAT_CANT_MAX, n))


def _altura_fundacion_mm_para_cantidad_lateral(elem):
    """
    Altura útil (mm) para sugerir cantidad de laterales: ``rango_z_caras_laterales_o_bbox`` y
    también la envolvente del elemento — se usa el **mayor** para no subestimar si sólo parte de
    las caras laterales tiene geometría en el familia.
    """
    if elem is None:
        return None
    h_ft = None
    try:
        z0, z1 = rango_z_caras_laterales_o_bbox(elem)
        if z0 is not None and z1 is not None:
            d = float(z1) - float(z0)
            if d > 1e-9:
                h_ft = d
    except Exception:
        pass
    try:
        bb = elem.get_BoundingBox(None)
        if bb is not None:
            dbb = float(bb.Max.Z) - float(bb.Min.Z)
            if dbb > 1e-9:
                if h_ft is None:
                    h_ft = dbb
                else:
                    h_ft = max(h_ft, dbb)
    except Exception:
        pass
    if h_ft is None or h_ft <= 1e-9:
        return None
    try:
        return float(h_ft) * 304.8
    except Exception:
        return None


def _z_range_fundacion_para_patas(el, eid, fz_range_by_eid):
    if fz_range_by_eid is not None:
        t = fz_range_by_eid.get(eid)
        if t is not None:
            z0, z1 = t
            if z0 is not None and z1 is not None:
                return z0, z1
    try:
        return rango_z_caras_laterales_o_bbox(el)
    except Exception:
        return None, None


def _leer_sep_mm_from_textblock(tb, default_mm):
    """Lee separación en mm desde el ``TextBox`` del spinner (100–300, paso 10)."""
    if tb is None:
        return float(default_mm)
    try:
        s = unicode(tb.Text).replace(u"mm", u"").strip()
        if not s:
            return float(default_mm)
        raw = int(round(float(s)))
        raw = max(_SEP_MM_MIN, min(_SEP_MM_MAX, raw))
        nmax = int((_SEP_MM_MAX - _SEP_MM_MIN) // _SEP_MM_STEP)
        steps = int(round((raw - _SEP_MM_MIN) / float(_SEP_MM_STEP)))
        steps = max(0, min(nmax, steps))
        return float(_SEP_MM_MIN + steps * _SEP_MM_STEP)
    except Exception:
        return float(default_mm)


def _resolver_rebar_bar_type_inferior_desde_combo(win, document):
    """
    ``RebarBarType`` asociado a ``CmbInfDiam`` para ``Rebar.CreateFromCurves*``.

    Usa el **índice seleccionado** y ``win._entries`` (mismo orden que los ítems del combo).
    Si la fila es la automática (sin elemento, solo etiqueta ``ø12 mm``), resuelve el tipo
    en el documento por diámetro nominal (misma lógica que ``_build_bar_type_entries``).
    """
    entries = getattr(win, "_entries", None) or []
    try:
        cmb = win._win.FindName("CmbInfDiam")
    except Exception:
        cmb = None
    if cmb is None:
        return None, u"No se encontró el combo Diámetro (inferior)."

    def _bt_por_mm(mm):
        if document is None or mm is None:
            return None
        try:
            from enfierrado_shaft_hashtag import resolver_bar_type_por_diametro_mm

            bt0, _exact, _delta = resolver_bar_type_por_diametro_mm(document, float(mm))
            return bt0
        except Exception:
            return None

    try:
        idx = int(cmb.SelectedIndex)
        if 0 <= idx < len(entries):
            bt, lbl = entries[idx]
            if bt is not None:
                return bt, None
            mm = _parse_diameter_mm_from_bar_combo_label(lbl)
            bt2 = _bt_por_mm(mm)
            if bt2 is not None:
                return bt2, None
    except Exception:
        pass

    try:
        sel = cmb.SelectedItem
        lab = unicode(sel) if sel is not None else u""
    except Exception:
        lab = u""
    for bt, lbl in entries:
        if unicode(lbl) != lab:
            continue
        if bt is not None:
            return bt, None
        mm = _parse_diameter_mm_from_bar_combo_label(lbl)
        bt2 = _bt_por_mm(mm)
        if bt2 is not None:
            return bt2, None
        break

    mm = _parse_diameter_mm_from_bar_combo_label(lab)
    bt3 = _bt_por_mm(mm)
    if bt3 is not None:
        return bt3, None

    return (
        None,
        u"No se pudo obtener RebarBarType desde «Diámetro (barra)» inferior. "
        u"Compruebe que el proyecto tiene tipos de barra cargados.",
    )


def _resolver_rebar_bar_type_superior_desde_combo(win, document):
    """``RebarBarType`` asociado a ``CmbSupDiam`` (misma convención que inferior)."""
    entries = getattr(win, "_entries", None) or []
    try:
        cmb = win._win.FindName("CmbSupDiam")
    except Exception:
        cmb = None
    if cmb is None:
        return None, u"No se encontró el combo Diámetro (superior)."

    def _bt_por_mm(mm):
        if document is None or mm is None:
            return None
        try:
            from enfierrado_shaft_hashtag import resolver_bar_type_por_diametro_mm

            bt0, _exact, _delta = resolver_bar_type_por_diametro_mm(document, float(mm))
            return bt0
        except Exception:
            return None

    try:
        idx = int(cmb.SelectedIndex)
        if 0 <= idx < len(entries):
            bt, lbl = entries[idx]
            if bt is not None:
                return bt, None
            mm = _parse_diameter_mm_from_bar_combo_label(lbl)
            bt2 = _bt_por_mm(mm)
            if bt2 is not None:
                return bt2, None
    except Exception:
        pass

    try:
        sel = cmb.SelectedItem
        lab = unicode(sel) if sel is not None else u""
    except Exception:
        lab = u""
    for bt, lbl in entries:
        if unicode(lbl) != lab:
            continue
        if bt is not None:
            return bt, None
        mm = _parse_diameter_mm_from_bar_combo_label(lbl)
        bt2 = _bt_por_mm(mm)
        if bt2 is not None:
            return bt2, None
        break

    mm = _parse_diameter_mm_from_bar_combo_label(lab)
    bt3 = _bt_por_mm(mm)
    if bt3 is not None:
        return bt3, None

    return (
        None,
        u"No se pudo obtener RebarBarType desde «Diámetro (barra)» superior. "
        u"Compruebe que el proyecto tiene tipos de barra cargados.",
    )


def _resolver_rebar_bar_type_lateral_desde_combo(win, document):
    """``RebarBarType`` asociado a ``CmbLatDiam`` (misma convención que inferior/superior)."""
    entries = getattr(win, "_entries", None) or []
    try:
        cmb = win._win.FindName("CmbLatDiam")
    except Exception:
        cmb = None
    if cmb is None:
        return None, u"No se encontró el combo Diámetro (lateral)."

    def _bt_por_mm(mm):
        if document is None or mm is None:
            return None
        try:
            from enfierrado_shaft_hashtag import resolver_bar_type_por_diametro_mm

            bt0, _exact, _delta = resolver_bar_type_por_diametro_mm(document, float(mm))
            return bt0
        except Exception:
            return None

    try:
        idx = int(cmb.SelectedIndex)
        if 0 <= idx < len(entries):
            bt, lbl = entries[idx]
            if bt is not None:
                return bt, None
            mm = _parse_diameter_mm_from_bar_combo_label(lbl)
            bt2 = _bt_por_mm(mm)
            if bt2 is not None:
                return bt2, None
    except Exception:
        pass

    try:
        sel = cmb.SelectedItem
        lab = unicode(sel) if sel is not None else u""
    except Exception:
        lab = u""
    for bt, lbl in entries:
        if unicode(lbl) != lab:
            continue
        if bt is not None:
            return bt, None
        mm = _parse_diameter_mm_from_bar_combo_label(lbl)
        bt2 = _bt_por_mm(mm)
        if bt2 is not None:
            return bt2, None
        break

    mm = _parse_diameter_mm_from_bar_combo_label(lab)
    bt3 = _bt_por_mm(mm)
    if bt3 is not None:
        return bt3, None

    return (
        None,
        u"No se pudo obtener RebarBarType desde «Diámetro (barra)» lateral. "
        u"Compruebe que el proyecto tiene tipos de barra cargados.",
    )


_FOUNDATION_CAT_ID = int(BuiltInCategory.OST_StructuralFoundation)
# Offset en planta desde caras laterales del perímetro (cota al eje): reduce solape con la malla ortogonal.
_RECUBRIMIENTO_PLANTA_CARAS_LATERALES_MM = 100.0
# Recubrimiento a caras horizontales (tapas): distancia hormigón–eje de barra (sin ø/2), fijo.
_RECUBRIMIENTO_CARA_HORIZONTAL_MM = 50.0
# Base (mm) para offset en planta y recorte en extremos de la curva perimetral inferior, junto con ø malla inferior.
_RECUBRIMIENTO_LATERAL_CARA_MM = 50.0
# Primera barra lateral: eje de la curva generatriz a esta distancia (mm) de la cara inferior, hacia el interior.
_OFFSET_EJE_PRIMERA_BARRA_LATERAL_DESDE_CARA_INFERIOR_MM = 100.0
# Largo del array (SetLayout): altura de la fundación menos este valor (mm).
_DESCUENTO_LARGO_ARRAY_LATERAL_MM = 200.0
# Patas modeladas en U (malla inf. + sup. activas): altura de fundación − este valor (mm).
# Misma lógica geométrica que ``_DESCUENTO_LONGITUD_PATA_U_FUNDACION_MM`` en geometría.
_DESCUENTO_PATA_U_INF_SUP_MM = 150.0

# Parámetro de proyecto sobre Rebar: ubicación de la armadura (malla inferior / superior / lateral).
_ARMA_UBICACION_PARAM = u"Armadura_Ubicacion"
_ARMA_UBICACION_INFERIOR = u"F"
_ARMA_UBICACION_SUPERIOR = u"F'"
_ARMA_UBICACION_LATERAL = u"L"


def _aplicar_armadura_ubicacion(rebar_element, valor_texto):
    """Rellena ``Armadura_Ubicacion`` en el ``Rebar`` si el parámetro existe y es escribible."""
    if rebar_element is None or valor_texto is None:
        return
    try:
        p = rebar_element.LookupParameter(_ARMA_UBICACION_PARAM)
        if p is None or p.IsReadOnly:
            return
        p.Set(valor_texto)
    except Exception:
        pass


def _altura_nominal_fundacion_ft(elem):
    """
    Obtiene la altura nominal (pies internos) de la fundación aislada.

    Estrategias en orden de prioridad:
    1. Nombre del tipo: último token numérico tras "x" (p. ej. ``3500x3500x800`` → 800 mm).
    2. Parámetros de instancia/tipo con nombres comunes (``Height``, ``Altura``, …).
    """
    if elem is None:
        return None

    # --- 1. Extraer altura del nombre del tipo (formato WxLxH) ---
    try:
        tipo = elem.Document.GetElement(elem.GetTypeId())
        type_name = u""
        if tipo is not None:
            try:
                type_name = unicode(tipo.Name) if tipo.Name else u""
            except Exception:
                pass
        if not type_name:
            try:
                from Autodesk.Revit.DB import BuiltInParameter
                p = elem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
                if p is not None:
                    type_name = unicode(p.AsString() or u"")
            except Exception:
                pass
        if type_name:
            # Busca el último bloque "NxNxN" en el nombre
            import re as _re
            m = _re.search(r'(\d+)[xX](\d+)[xX](\d+)', type_name)
            if m:
                h_mm = float(m.group(3))
                if h_mm > 1.0:
                    try:
                        from Autodesk.Revit.DB import UnitUtils, UnitTypeId
                        return UnitUtils.ConvertToInternalUnits(h_mm, UnitTypeId.Millimeters)
                    except Exception:
                        return h_mm / 304.8
    except Exception:
        pass

    # --- 2. Parámetros de instancia/tipo ---
    try:
        from Autodesk.Revit.DB import StorageType
    except Exception:
        return None
    _PARAM_NAMES = (
        u"Height", u"Altura", u"Espesor", u"Thickness",
        u"h", u"Alto", u"Foundation Height", u"Depth",
    )
    for src in (elem, tipo if 'tipo' in dir() else None):
        if src is None:
            continue
        for name in _PARAM_NAMES:
            try:
                p = src.LookupParameter(name)
                if p is not None and p.HasValue and p.StorageType == StorageType.Double:
                    val = float(p.AsDouble())
                    if val > 0.01:
                        return val
            except Exception:
                continue
    return None


def _leg_ft_pata_u_malla_inferior(z0p, z1p, d_mm_bar, sup_on, elem=None):
    """
    Longitud de cada pata (pies) para la polilínea U de la malla inferior.

    - Con **superior activo**: altura nominal − ``_DESCUENTO_PATA_U_INF_SUP_MM``.
      Prioriza el parámetro ``Height`` del elemento sobre la geometría de caras
      para evitar discrepancias por geometría de familia (p. ej. 804 mm vs 800 mm).
    - **Solo inferior**: largo según tabla BIMTools (:func:`largo_gancho_u_tabla_mm`),
      acotado a lo que permite la geometría.
    """
    # Descuento total: 150 mm + ø/2 (media caña del eje de la barra)
    try:
        descuento_mm = float(_DESCUENTO_PATA_U_INF_SUP_MM) + float(d_mm_bar) / 2.0
    except Exception:
        descuento_mm = float(_DESCUENTO_PATA_U_INF_SUP_MM)

    # Intentar obtener la altura desde el parámetro de la familia
    h_param_ft = _altura_nominal_fundacion_ft(elem) if elem is not None else None

    if h_param_ft is not None:
        leg_max = longitud_pata_u_fundacion_inf_sup_ft(
            0.0, h_param_ft, descuento_mm
        )
    else:
        leg_max = longitud_pata_u_fundacion_inf_sup_ft(
            z0p, z1p, descuento_mm
        )

    if sup_on:
        return leg_max
    # Solo inferior: largo de tabla ø, con el mismo descuento ø/2 que se aplica
    # en la polilínea U (eje de barra queda ø/2 más corto que el valor de tabla).
    hook_mm = largo_gancho_u_tabla_mm(d_mm_bar)
    leg_ft = None
    if hook_mm is not None:
        try:
            from bimtools_rebar_hook_lengths import pata_eje_curve_loop_mm_desde_tabla_mm
            eje_mm = pata_eje_curve_loop_mm_desde_tabla_mm(hook_mm, d_mm_bar)
        except Exception:
            eje_mm = None
        if eje_mm is None:
            try:
                eje_mm = float(hook_mm) - float(d_mm_bar) / 2.0
            except Exception:
                eje_mm = float(hook_mm)
        try:
            from Autodesk.Revit.DB import UnitUtils, UnitTypeId

            leg_ft = UnitUtils.ConvertToInternalUnits(
                float(eje_mm), UnitTypeId.Millimeters
            )
        except Exception:
            leg_ft = float(eje_mm) / 304.8
    if leg_ft is not None and leg_max is not None:
        return min(leg_ft, leg_max)
    if leg_ft is not None:
        return leg_ft
    return leg_max


# Duración del efecto de cierre (ms). La apertura usa el mismo valor (Storyboard + XAML).
_WINDOW_CLOSE_MS = 180
# ScrollViewer principal: al mostrar el pie de propagación, reduce un poco la zona central
# para que quepa el bloque extra sin chocar con MaxHeight de la ventana.
_SV_CONTENIDO_MAX_H_NORMAL = 600.0
_SV_CONTENIDO_MAX_H_PROPAGACION = 500.0

# Ancho del diálogo (misma lógica que ``area_reinforcement_losa``: 2×110 + @ + pads).
# Con el contenedor ``Informacion Armadura`` hay un borde/padding extra horizontal.
_FUND_INPUT_COLS_PER_ROW = 2
_FUND_COMBO_WIDTH_PX = 110
_FUND_DIAM_ESP_AT_COL_PX = 28
_FUND_BLOCK_PAD_H_PX = 16
_FUND_GROUPBOX_PAD_H_PX = 16
_FUND_OUTER_PAD_H_PX = 28
_FUND_ARM_INFO_GROUPBOX_EXTRA_H_PX = 12
_FUND_WIDTH_TITLE_MIN_PX = 288


def _fundacion_aislada_form_width_px(
    input_cols_per_row=None,
    combo_width_px=None,
):
    cols = int(input_cols_per_row or _FUND_INPUT_COLS_PER_ROW)
    cols = max(1, cols)
    c = int(combo_width_px or _FUND_COMBO_WIDTH_PX)
    row_inner = cols * c + _FUND_DIAM_ESP_AT_COL_PX + _FUND_BLOCK_PAD_H_PX
    w = (
        row_inner
        + _FUND_GROUPBOX_PAD_H_PX
        + _FUND_OUTER_PAD_H_PX
        + _FUND_ARM_INFO_GROUPBOX_EXTRA_H_PX
    )
    w = max(w, _FUND_WIDTH_TITLE_MIN_PX)
    return int((int(w) + 3) // 4 * 4)


# Recursos y chrome alineados con ``area_reinforcement_losa.XAML`` (Crear Area Reinf RPS).
_WPF_STORYBOARD_DUR_STR = "0:0:{0:.2f}".format(_WINDOW_CLOSE_MS / 1000.0)
_ENFIERRADO_FUND_XAML = (
    u"""
<Window
    x:Name="FundacionWin"
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Arainco - Armadura Fundacion Aislada"
    SizeToContent="Height"
    MaxHeight="920"
    WindowStartupLocation="Manual"
    Background="Transparent"
    AllowsTransparency="True"
    FontFamily="Segoe UI"
    WindowStyle="None"
    ResizeMode="NoResize"
    Topmost="True"
    UseLayoutRounding="True"
    >
  <!-- Apertura/cierre: ScaleTransform en FundRootScale (0,0). Igual que Wall Foundation Reinforcement. -->
  <Window.Resources>
    <Storyboard x:Key="FundOpenGrowStoryboard">
      <DoubleAnimation Storyboard.TargetName="FundRootScale" Storyboard.TargetProperty="ScaleX"
                       From="0" To="1" Duration="__WPF_STORYBOARD_DUR__" FillBehavior="HoldEnd">
        <DoubleAnimation.EasingFunction>
          <QuadraticEase EasingMode="EaseOut"/>
        </DoubleAnimation.EasingFunction>
      </DoubleAnimation>
      <DoubleAnimation Storyboard.TargetName="FundRootScale" Storyboard.TargetProperty="ScaleY"
                       From="0" To="1" Duration="__WPF_STORYBOARD_DUR__" FillBehavior="HoldEnd">
        <DoubleAnimation.EasingFunction>
          <QuadraticEase EasingMode="EaseOut"/>
        </DoubleAnimation.EasingFunction>
      </DoubleAnimation>
      <DoubleAnimation Storyboard.TargetName="FundacionWin" Storyboard.TargetProperty="Opacity"
                       From="0" To="1" Duration="__WPF_STORYBOARD_DUR__" FillBehavior="HoldEnd">
        <DoubleAnimation.EasingFunction>
          <QuadraticEase EasingMode="EaseOut"/>
        </DoubleAnimation.EasingFunction>
      </DoubleAnimation>
    </Storyboard>
""" + BIMTOOLS_DARK_STYLES_XML + u"""
  </Window.Resources>
  <Border x:Name="FundacionRootChrome" CornerRadius="10" Background="#0A1A2F" Padding="12"
          BorderBrush="#1A3A4D" BorderThickness="1" ClipToBounds="True" RenderTransformOrigin="0,0">
    <Border.Effect>
      <DropShadowEffect Color="#000000" BlurRadius="16" ShadowDepth="0" Opacity="0.35"/>
    </Border.Effect>
    <Border.RenderTransform>
      <ScaleTransform x:Name="FundRootScale" ScaleX="0" ScaleY="0"/>
    </Border.RenderTransform>
    <Grid HorizontalAlignment="Stretch">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <Border x:Name="TitleBar" Grid.Row="0" Background="#0E1B32" CornerRadius="6" Padding="10,8" Margin="0,0,0,10"
              BorderBrush="#21465C" BorderThickness="1" HorizontalAlignment="Stretch">
        <Grid HorizontalAlignment="Stretch">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <Image x:Name="ImgLogo" Width="38" Height="38" Grid.Column="0"
                 Stretch="Uniform" Margin="0,0,8,0" VerticalAlignment="Center"/>
          <StackPanel Grid.Column="1" VerticalAlignment="Center" Margin="0,0,4,0">
            <TextBlock Text="Armadura Fundacion Aislada" FontSize="13" FontWeight="SemiBold"
                       Foreground="#E8F4F8" TextWrapping="NoWrap"/>
          </StackPanel>
          <Button x:Name="BtnClose" Grid.Column="2"
                  Style="{StaticResource BtnCloseX_MinimalNoBg}"
                  VerticalAlignment="Center" HorizontalAlignment="Right" ToolTip="Cerrar"/>
        </Grid>
      </Border>

      <ScrollViewer x:Name="SvContenido" Grid.Row="1" VerticalScrollBarVisibility="Auto" MaxHeight="600" Margin="0,0,0,2">
        <StackPanel HorizontalAlignment="Stretch">
          <StackPanel Margin="0,0,0,8" HorizontalAlignment="Stretch">
            <Button x:Name="BtnSeleccionar" Content="Seleccionar fundación en modelo"
                    Style="{StaticResource BtnSelectOutline}"
                    HorizontalAlignment="Stretch"/>
          </StackPanel>

          <GroupBox Style="{StaticResource GbParams}" Margin="0,0,0,10" HorizontalAlignment="Stretch">
            <GroupBox.Header>
              <TextBlock Text="Informacion Armadura" FontWeight="SemiBold" Foreground="#E8F4F8" FontSize="11"/>
            </GroupBox.Header>
            <StackPanel>
          <GroupBox Style="{StaticResource GbParams}" Margin="0,0,0,10" HorizontalAlignment="Stretch">
            <GroupBox.Header>
              <CheckBox x:Name="ChkInferior" IsChecked="True" Content="Armadura Inferior"
                        Foreground="#E8F4F8" FontWeight="SemiBold" FontSize="11" VerticalAlignment="Center"/>
            </GroupBox.Header>
            <StackPanel x:Name="PanelInferior">
              <Grid HorizontalAlignment="Center">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="110"/>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="110"/>
                </Grid.ColumnDefinitions>
                <ComboBox Grid.Column="0" x:Name="CmbInfDiam" Style="{StaticResource Combo}" IsEditable="False" IsReadOnly="True">
                  <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                </ComboBox>
                <TextBlock Grid.Column="1" Text="@" FontSize="12" FontWeight="Bold"
                           Foreground="#95B8CC" VerticalAlignment="Center" HorizontalAlignment="Center" Margin="6,0,6,0"/>
                <Border Grid.Column="2" Width="110" Height="24" CornerRadius="4" Background="#050E18"
                        BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True">
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="18"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="TxtInfSepMm" Grid.Column="0" Style="{StaticResource CantSpinnerText}"
                             Text="150" Padding="6,0,6,0" HorizontalAlignment="Stretch" VerticalContentAlignment="Center"
                             ToolTip="Separación entre barras (mm): 100 a 300, paso 10"/>
                    <Border Grid.Column="1" Background="#0E1B32" BorderBrush="#1A3A4D"
                            BorderThickness="1,0,0,0" CornerRadius="0,4,4,0" ClipToBounds="True">
                      <Grid>
                        <Grid.RowDefinitions>
                          <RowDefinition Height="*"/>
                          <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <RepeatButton x:Name="BtnInfSepUp" Grid.Row="0" Style="{StaticResource SpinRepeatBtn}" Content="▲"
                                      ToolTip="Más 10 mm (máx. 300 mm)"/>
                        <RepeatButton x:Name="BtnInfSepDown" Grid.Row="1" Style="{StaticResource SpinRepeatBtn}" Content="▼"
                                      ToolTip="Menos 10 mm (mín. 100 mm)"/>
                      </Grid>
                    </Border>
                  </Grid>
                </Border>
              </Grid>
            </StackPanel>
          </GroupBox>

          <GroupBox Style="{StaticResource GbParams}" Margin="0,0,0,10" HorizontalAlignment="Stretch">
            <GroupBox.Header>
              <CheckBox x:Name="ChkSuperior" IsChecked="True" Content="Armadura Superior"
                        Foreground="#E8F4F8" FontWeight="SemiBold" FontSize="11" VerticalAlignment="Center"/>
            </GroupBox.Header>
            <StackPanel x:Name="PanelSuperior">
              <Grid HorizontalAlignment="Center">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="110"/>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="110"/>
                </Grid.ColumnDefinitions>
                <ComboBox Grid.Column="0" x:Name="CmbSupDiam" Style="{StaticResource Combo}" IsEditable="False" IsReadOnly="True">
                  <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                </ComboBox>
                <TextBlock Grid.Column="1" Text="@" FontSize="12" FontWeight="Bold"
                           Foreground="#95B8CC" VerticalAlignment="Center" HorizontalAlignment="Center" Margin="6,0,6,0"/>
                <Border Grid.Column="2" Width="110" Height="24" CornerRadius="4" Background="#050E18"
                        BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True">
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="18"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="TxtSupSepMm" Grid.Column="0" Style="{StaticResource CantSpinnerText}"
                             Text="150" Padding="6,0,6,0" HorizontalAlignment="Stretch" VerticalContentAlignment="Center"
                             ToolTip="Separación entre barras (mm): 100 a 300, paso 10"/>
                    <Border Grid.Column="1" Background="#0E1B32" BorderBrush="#1A3A4D"
                            BorderThickness="1,0,0,0" CornerRadius="0,4,4,0" ClipToBounds="True">
                      <Grid>
                        <Grid.RowDefinitions>
                          <RowDefinition Height="*"/>
                          <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <RepeatButton x:Name="BtnSupSepUp" Grid.Row="0" Style="{StaticResource SpinRepeatBtn}" Content="▲"
                                      ToolTip="Más 10 mm (máx. 300 mm)"/>
                        <RepeatButton x:Name="BtnSupSepDown" Grid.Row="1" Style="{StaticResource SpinRepeatBtn}" Content="▼"
                                      ToolTip="Menos 10 mm (mín. 100 mm)"/>
                      </Grid>
                    </Border>
                  </Grid>
                </Border>
              </Grid>
            </StackPanel>
          </GroupBox>

          <GroupBox Style="{StaticResource GbParams}" Margin="0,0,0,0" HorizontalAlignment="Stretch">
            <GroupBox.Header>
              <CheckBox x:Name="ChkLateral" IsChecked="True" Content="Armadura Lateral"
                        Foreground="#E8F4F8" FontWeight="SemiBold" FontSize="11" VerticalAlignment="Center"/>
            </GroupBox.Header>
            <StackPanel x:Name="PanelLateral">
              <Grid HorizontalAlignment="Center">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="110"/>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="110"/>
                </Grid.ColumnDefinitions>
                <Border Grid.Column="0" Width="110" Height="24" CornerRadius="4" Background="#050E18"
                        BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True">
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="18"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="TxtLatCant" Grid.Column="0" Style="{StaticResource CantSpinnerText}"
                             Text="2" Padding="6,0,6,0" HorizontalAlignment="Stretch" VerticalContentAlignment="Center"
                             ToolTip="Cantidad de barras laterales (1 a 10). Tras seleccionar la fundación, valor sugerido según altura."/>
                    <Border Grid.Column="1" Background="#0E1B32" BorderBrush="#1A3A4D"
                            BorderThickness="1,0,0,0" CornerRadius="0,4,4,0" ClipToBounds="True">
                      <Grid>
                        <Grid.RowDefinitions>
                          <RowDefinition Height="*"/>
                          <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <RepeatButton x:Name="BtnLatCantUp" Grid.Row="0" Style="{StaticResource SpinRepeatBtn}" Content="▲"
                                      ToolTip="Más 1 (máx. 10)"/>
                        <RepeatButton x:Name="BtnLatCantDown" Grid.Row="1" Style="{StaticResource SpinRepeatBtn}" Content="▼"
                                      ToolTip="Menos 1 (mín. 1)"/>
                      </Grid>
                    </Border>
                  </Grid>
                </Border>
                <Border Grid.Column="1" Background="Transparent" MinWidth="14" Margin="6,0,6,0"/>
                <ComboBox Grid.Column="2" x:Name="CmbLatDiam" Style="{StaticResource Combo}" IsEditable="False" IsReadOnly="True">
                  <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                </ComboBox>
              </Grid>
            </StackPanel>
          </GroupBox>
            </StackPanel>
          </GroupBox>
        </StackPanel>
      </ScrollViewer>

      <StackPanel Grid.Row="2" Margin="0" HorizontalAlignment="Stretch">
        <Border x:Name="BorderPropagacion" Visibility="Collapsed" Background="#0E1B32"
                BorderBrush="#1A3A4D" BorderThickness="1" CornerRadius="4" Padding="8,6" Margin="0,0,0,6">
          <TextBlock x:Name="TxtPropagacionTitulo" TextWrapping="Wrap" Foreground="#E8F4F8" FontSize="11"/>
        </Border>
        <Button x:Name="BtnColocar" Content="Colocar armaduras"
                Style="{StaticResource BtnPrimary}"
                HorizontalAlignment="Stretch"/>
      </StackPanel>
    </Grid>
  </Border>
</Window>
""").replace(u"__WPF_STORYBOARD_DUR__", _WPF_STORYBOARD_DUR_STR)


class FundacionAisladaSelectionFilter(ISelectionFilter):
    """Categoría Structural Foundation; excluye zapata de muro (WallFoundation)."""

    def AllowElement(self, elem):
        try:
            if elem is None:
                return False
            if isinstance(elem, WallFoundation):
                return False
            cat = elem.Category
            if cat is None:
                return False
            return element_id_to_int(cat.Id) == _FOUNDATION_CAT_ID
        except Exception:
            return False

    def AllowReference(self, ref, pt):
        return False


class SeleccionarFundacionesHandler(IExternalEventHandler):
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
        try:
            from geometria_fundacion_cara_inferior import clear_face_cache
            clear_face_cache()
        except Exception:
            pass
        flt = FundacionAisladaSelectionFilter()
        try:
            ref_pick = uidoc.Selection.PickObject(
                ObjectType.Element,
                flt,
                u"Seleccione una fundación aislada.",
            )
        except Exception:
            win._set_estado(u"Selección cancelada.")
            try:
                win._show_after_pick()
            except Exception:
                pass
            return

        if ref_pick is None:
            win._set_estado(u"Sin elemento.")
            try:
                win._show_after_pick()
            except Exception:
                pass
            return

        try:
            eid = ref_pick.ElementId
        except Exception:
            eid = None
        if eid is None:
            win._set_estado(u"Sin elemento.")
            try:
                win._show_after_pick()
            except Exception:
                pass
            return

        win._document = doc
        win._foundation_ids = [eid]
        try:
            win._hide_propagacion_ui()
        except Exception:
            pass
        try:
            win._refresh_laterales_cantidad_desde_fundaciones()
        except Exception:
            pass
        win._set_estado(u"Fundación seleccionada.")
        try:
            win._show_after_pick()
        except Exception:
            pass

    def GetName(self):
        return u"SeleccionarFundacionesAisladas"


class ColocarArmaduraFundacionStubHandler(IExternalEventHandler):
    """Reservado: aquí se enlazará la creación de curvas y barras."""

    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        win = self._window_ref()
        if win is None:
            return
        try:
            self._execute_colocar_stub(
                uiapp,
                aplicar_recubrimiento_inferior_completo_mm,
                extraer_curva_lado_menor_cara_inferior,
                extraer_curva_lado_mayor_cara_inferior,
                obtener_marco_coordenadas_cara_inferior,
            )
        except Exception as ex:
            try:
                msg = u"Error al colocar / evaluar:\n{0}".format(ex)
                _task_dialog_show(u"BIMTools — Armadura Fundacion Aislada", msg, win._win)
            except Exception:
                pass
            try:
                win._set_estado(u"Error: {0}".format(ex))
            except Exception:
                pass

    def _execute_colocar_stub(
        self,
        uiapp,
        aplicar_recubrimiento_inferior_completo_mm,
        extraer_curva_lado_menor_cara_inferior,
        extraer_curva_lado_mayor_cara_inferior,
        obtener_marco_coordenadas_cara_inferior,
    ):
        win = self._window_ref()
        if win is None:
            return
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            return
        doc = uidoc.Document
        # Limpiar caché de geometría de caras: cada Execute es una nueva operación sobre el modelo.
        try:
            from geometria_fundacion_cara_inferior import clear_face_cache
            clear_face_cache()
        except Exception:
            pass
        ids_run = list(win._foundation_ids or [])
        ovr = getattr(win, "_ids_run_override", None)
        if ovr:
            ids_run = list(ovr)
            try:
                win._ids_run_override = None
            except Exception:
                pass
        ok = 0
        err_ids = []
        z_mm_muestra = None
        lado_menor_mm_muestra = None
        lado_offset_mm_muestra = None
        rec_mm = _RECUBRIMIENTO_PLANTA_CARAS_LATERALES_MM
        try:
            sep_mm = _leer_sep_mm_from_textblock(
                win._win.FindName("TxtInfSepMm"), float(_SEP_MM_DEFAULT_VAL)
            )
        except Exception:
            sep_mm = float(_SEP_MM_DEFAULT_VAL)
        rec_sup_mm = _RECUBRIMIENTO_PLANTA_CARAS_LATERALES_MM
        try:
            sep_sup_mm = _leer_sep_mm_from_textblock(
                win._win.FindName("TxtSupSepMm"), float(_SEP_MM_DEFAULT_VAL)
            )
        except Exception:
            sep_sup_mm = float(_SEP_MM_DEFAULT_VAL)
        try:
            n_lat_cant = _leer_lat_cant_desde_textbox(
                win._win.FindName("TxtLatCant"),
                int(_LAT_CANT_DEFAULT_TXT),
            )
        except Exception:
            n_lat_cant = int(_LAT_CANT_DEFAULT_TXT)
        try:
            n_lat_cant = max(
                _LAT_CANT_MIN,
                min(_LAT_CANT_MAX, int(n_lat_cant)),
            )
        except Exception:
            n_lat_cant = int(_LAT_CANT_DEFAULT_TXT)

        try:
            _chk_inf = win._win.FindName("ChkInferior")
            inf_on = _chk_inf is None or _chk_inf.IsChecked == True
        except Exception:
            inf_on = True
        try:
            _chk_sup = win._win.FindName("ChkSuperior")
            sup_on = _chk_sup is None or _chk_sup.IsChecked == True
        except Exception:
            sup_on = True
        try:
            _chk_lat = win._win.FindName("ChkLateral")
            lat_on = _chk_lat is None or _chk_lat.IsChecked == True
        except Exception:
            lat_on = True

        # Pre-resolver bar types una sola vez; se reusan en el bucle de geometría y en la transacción.
        _bt_inf_pre, _err_bt_inf_pre = (
            _resolver_rebar_bar_type_inferior_desde_combo(win, doc) if inf_on else (None, None)
        )
        _bt_sup_pre, _err_bt_sup_pre = (
            _resolver_rebar_bar_type_superior_desde_combo(win, doc) if sup_on else (None, None)
        )
        _bt_lat_pre, _err_bt_lat_pre = (
            _resolver_rebar_bar_type_lateral_desde_combo(win, doc) if lat_on else (None, None)
        )

        # Recorte en extremos (hacia caras laterales / ganchos de la malla): el canto del hormigón
        # define c nominal; el eje de la barra queda a c + ø/2 (recubrimiento a fibra exterior ≈ c).
        ext_mesh_inf_mm = float(RECUBRIMIENTO_EXTREMOS_MM)
        ext_mesh_sup_mm = float(RECUBRIMIENTO_EXTREMOS_MM)
        if inf_on and _bt_inf_pre is not None:
            _dmi = _rebar_nominal_diameter_mm(_bt_inf_pre)
            if _dmi is not None:
                ext_mesh_inf_mm = float(RECUBRIMIENTO_EXTREMOS_MM) + 0.5 * float(_dmi)
        if sup_on and _bt_sup_pre is not None:
            _dms = _rebar_nominal_diameter_mm(_bt_sup_pre)
            if _dms is not None:
                ext_mesh_sup_mm = float(RECUBRIMIENTO_EXTREMOS_MM) + 0.5 * float(_dms)

        el_by_id = {}
        for _eid in ids_run:
            el_by_id[_eid] = doc.GetElement(_eid)

        marco_inf_by_eid = {}
        lh_pack_by_eid = {}
        fz_range_by_eid = {}

        elementos_ok = []
        elementos_ok_sup = []
        lado_mayor_mm_muestra = None
        lado_menor_mm_sup_muestra = None
        lado_mayor_mm_sup_muestra = None
        z_sup_mm_muestra = None
        err_ids_inf = []
        err_ids_sup = []
        err_ids_lat = []
        for eid in ids_run:
            el = el_by_id[eid]
            if el is None:
                err_ids.append(element_id_to_int(eid))
                continue
            marco = None
            tiene_alguna = False
            z_ft_elem = None
            res_menor = None
            curva = None

            if inf_on:
                marco = obtener_marco_coordenadas_cara_inferior(el)
                marco_inf_by_eid[eid] = marco
                res_menor = extraer_curva_lado_menor_cara_inferior(el)
                # Para zapatas cuadradas, todas las aristas tienen la misma longitud y
                # extraer_curva_lado_mayor devuelve la misma arista que la menor.
                # Pasamos excluir_curva para obtener la arista PERPENDICULAR adyacente.
                res_mayor = extraer_curva_lado_mayor_cara_inferior(
                    el,
                    excluir_curva=res_menor[0] if res_menor is not None else None,
                )
                # Si aún coinciden (geometría degenerada), descartar mayor.
                if res_menor is not None and res_mayor is not None:
                    try:
                        tol_s = 1.0 / 304.8
                        a0, a1 = res_menor[0].GetEndPoint(0), res_menor[0].GetEndPoint(1)
                        b0, b1 = res_mayor[0].GetEndPoint(0), res_mayor[0].GetEndPoint(1)
                        mismo = (
                            a0.DistanceTo(b0) < tol_s
                            and a1.DistanceTo(b1) < tol_s
                        ) or (
                            a0.DistanceTo(b1) < tol_s
                            and a1.DistanceTo(b0) < tol_s
                        )
                        if mismo:
                            res_mayor = None
                    except Exception:
                        pass
                if res_menor is not None:
                    curva, z_ft_elem = res_menor
                    tiene_alguna = True
                    curva_tratada, _co = aplicar_recubrimiento_inferior_completo_mm(
                        curva, el, rec_mm, ext_mesh_inf_mm
                    )
                    cara_pp = None
                    ev_par = None
                    try:
                        ev_par = evaluar_caras_paralelas_curva_mas_cercana(
                            el, curva_tratada
                        )
                        if ev_par and ev_par.get("mejor"):
                            cara_pp = ev_par["mejor"]
                    except Exception:
                        pass
                    # Para "menor" (barras en dirección corta), el array se distribuye
                    # a lo largo de la arista MAYOR (la perpendicular).
                    try:
                        perp_len_menor_mm = (
                            float(res_mayor[0].Length) * 304.8
                            if res_mayor is not None
                            else float(curva.Length) * 304.8
                        )
                    except Exception:
                        perp_len_menor_mm = None
                    elementos_ok.append(
                        (el, curva_tratada, marco, cara_pp, ev_par, u"menor", perp_len_menor_mm)
                    )
                    if lado_menor_mm_muestra is None:
                        try:
                            lado_menor_mm_muestra = float(curva.Length) * 304.8
                        except Exception:
                            pass
                    if lado_offset_mm_muestra is None and curva_tratada is not None:
                        try:
                            lado_offset_mm_muestra = float(curva_tratada.Length) * 304.8
                        except Exception:
                            pass

                if res_mayor is not None:
                    curva2, z2 = res_mayor
                    tiene_alguna = True
                    if z_ft_elem is None:
                        z_ft_elem = z2
                    curva_tratada2, _co2 = aplicar_recubrimiento_inferior_completo_mm(
                        curva2, el, rec_mm, ext_mesh_inf_mm
                    )
                    cara_pp2 = None
                    ev_par2 = None
                    try:
                        ev_par2 = evaluar_caras_paralelas_curva_mas_cercana(
                            el, curva_tratada2
                        )
                        if ev_par2 and ev_par2.get("mejor"):
                            cara_pp2 = ev_par2["mejor"]
                    except Exception:
                        pass
                    # Para "mayor" (barras en dirección larga), el array se distribuye
                    # a lo largo de la arista MENOR (la perpendicular).
                    try:
                        perp_len_mayor_mm = (
                            float(res_menor[0].Length) * 304.8
                            if res_menor is not None
                            else float(curva2.Length) * 304.8
                        )
                    except Exception:
                        perp_len_mayor_mm = None
                    elementos_ok.append(
                        (el, curva_tratada2, marco, cara_pp2, ev_par2, u"mayor", perp_len_mayor_mm)
                    )
                    if lado_mayor_mm_muestra is None:
                        try:
                            lado_mayor_mm_muestra = float(curva2.Length) * 304.8
                        except Exception:
                            pass

            if inf_on and not tiene_alguna:
                err_ids_inf.append(element_id_to_int(eid))

            tiene_sup = False
            z_ft_sup = None
            marco_sup = None
            res_menor_s = None
            curva_s = None

            if sup_on:
                marco_sup = obtener_marco_coordenadas_cara_superior(el)
                res_menor_s = extraer_curva_lado_menor_cara_superior(el)
                res_mayor_s = extraer_curva_lado_mayor_cara_superior(
                    el,
                    excluir_curva=res_menor_s[0] if res_menor_s is not None else None,
                )
                if res_menor_s is not None and res_mayor_s is not None:
                    try:
                        tol_s = 1.0 / 304.8
                        a0, a1 = res_menor_s[0].GetEndPoint(0), res_menor_s[0].GetEndPoint(1)
                        b0, b1 = res_mayor_s[0].GetEndPoint(0), res_mayor_s[0].GetEndPoint(1)
                        mismo = (
                            a0.DistanceTo(b0) < tol_s
                            and a1.DistanceTo(b1) < tol_s
                        ) or (
                            a0.DistanceTo(b1) < tol_s
                            and a1.DistanceTo(b0) < tol_s
                        )
                        if mismo:
                            res_mayor_s = None
                    except Exception:
                        pass
                if res_menor_s is not None:
                    curva_s, z_ft_sup = res_menor_s
                    tiene_sup = True
                    curva_tratada_s, _co_s = aplicar_recubrimiento_inferior_completo_mm(
                        curva_s, el, rec_sup_mm, ext_mesh_sup_mm
                    )
                    cara_pp_s = None
                    ev_par_s = None
                    try:
                        ev_par_s = evaluar_caras_paralelas_curva_mas_cercana(
                            el, curva_tratada_s
                        )
                        if ev_par_s and ev_par_s.get("mejor"):
                            cara_pp_s = ev_par_s["mejor"]
                    except Exception:
                        pass
                    try:
                        perp_len_menor_s_mm = (
                            float(res_mayor_s[0].Length) * 304.8
                            if res_mayor_s is not None
                            else float(curva_s.Length) * 304.8
                        )
                    except Exception:
                        perp_len_menor_s_mm = None
                    elementos_ok_sup.append(
                        (el, curva_tratada_s, marco_sup, cara_pp_s, ev_par_s, u"menor", perp_len_menor_s_mm)
                    )
                    if lado_menor_mm_sup_muestra is None:
                        try:
                            lado_menor_mm_sup_muestra = float(curva_s.Length) * 304.8
                        except Exception:
                            pass

                if res_mayor_s is not None:
                    curva2s, z2s = res_mayor_s
                    tiene_sup = True
                    if z_ft_sup is None:
                        z_ft_sup = z2s
                    curva_tratada2s, _co2s = (
                        aplicar_recubrimiento_inferior_completo_mm(
                            curva2s, el, rec_sup_mm, ext_mesh_sup_mm
                        )
                    )
                    cara_pp2s = None
                    ev_par2s = None
                    try:
                        ev_par2s = evaluar_caras_paralelas_curva_mas_cercana(
                            el, curva_tratada2s
                        )
                        if ev_par2s and ev_par2s.get("mejor"):
                            cara_pp2s = ev_par2s["mejor"]
                    except Exception:
                        pass
                    try:
                        perp_len_mayor_s_mm = (
                            float(res_menor_s[0].Length) * 304.8
                            if res_menor_s is not None
                            else float(curva2s.Length) * 304.8
                        )
                    except Exception:
                        perp_len_mayor_s_mm = None
                    elementos_ok_sup.append(
                        (
                            el,
                            curva_tratada2s,
                            marco_sup,
                            cara_pp2s,
                            ev_par2s,
                            u"mayor",
                            perp_len_mayor_s_mm,
                        )
                    )
                    if lado_mayor_mm_sup_muestra is None:
                        try:
                            lado_mayor_mm_sup_muestra = float(curva2s.Length) * 304.8
                        except Exception:
                            pass

            if sup_on and not tiene_sup:
                err_ids_sup.append(element_id_to_int(eid))

            tiene_lat = False
            if lat_on:
                try:
                    lh = lineas_horizontales_perimetro_inferior_exterior(el)
                    z0, z1 = rango_z_caras_laterales_o_bbox(el)
                    lh_pack_by_eid[eid] = lh
                    fz_range_by_eid[eid] = (z0, z1)
                    tiene_lat = (
                        lh is not None
                        and len(lh[0]) > 0
                        and z0 is not None
                        and z1 is not None
                        and float(z1) - float(z0) > 1e-6
                    )
                except Exception:
                    lh_pack_by_eid[eid] = None
                    fz_range_by_eid[eid] = (None, None)
                    tiene_lat = False
            if lat_on and not tiene_lat:
                err_ids_lat.append(element_id_to_int(eid))

            if (
                not inf_on
                and not sup_on
                and not lat_on
            ):
                err_ids.append(element_id_to_int(eid))
                continue

            if not (
                (inf_on and tiene_alguna)
                or (sup_on and tiene_sup)
                or (lat_on and tiene_lat)
            ):
                continue

            ok += 1
            if z_mm_muestra is None and z_ft_elem is not None:
                try:
                    z_mm_muestra = float(z_ft_elem) * 304.8
                except Exception:
                    pass
            if z_sup_mm_muestra is None and z_ft_sup is not None:
                try:
                    z_sup_mm_muestra = float(z_ft_sup) * 304.8
                except Exception:
                    pass
        if err_ids_inf:
            _task_dialog_show(
                u"BIMTools — Armadura Fundacion Aislada",
                u"No se pudo extraer la cara inferior ni curva de borde (lado corto/largo) en: {0}.".format(
                    u", ".join(unicode(i) for i in err_ids_inf)
                ),
                win._win,
            )
            try:
                win._set_estado(
                    u"Revisar geometría de fundación (Id sin cara inferior clara)."
                )
            except Exception:
                pass
            return
        if err_ids_sup:
            _task_dialog_show(
                u"BIMTools — Armadura Fundacion Aislada",
                u"No se pudo extraer la cara superior ni curva de borde (lado corto/largo) en: {0}.".format(
                    u", ".join(unicode(i) for i in err_ids_sup)
                ),
                win._win,
            )
            try:
                win._set_estado(
                    u"Revisar geometría de fundación (Id sin cara superior clara)."
                )
            except Exception:
                pass
            return
        if err_ids:
            _task_dialog_show(
                u"BIMTools — Armadura Fundacion Aislada",
                u"Active al menos un grupo de armadura (inferior, superior o lateral).",
                win._win,
            )
            try:
                win._set_estado(u"Marque inferior, superior y/o lateral.")
            except Exception:
                pass
            return
        if ok == 0:
            _task_dialog_show(
                u"BIMTools — Armadura Fundacion Aislada",
                u"No se pudo resolver geometría para ninguna fundación seleccionada "
                u"(según los grupos activados).",
                win._win,
            )
            try:
                win._set_estado(u"Sin geometría válida para los grupos activados.")
            except Exception:
                pass
            return
        msg = u""
        if inf_on:
            msg += u"Cara inferior resuelta en {0} elemento(s).".format(ok)
        if sup_on:
            if msg:
                msg += u"\n\n"
            msg += u"Cara superior resuelta en {0} elemento(s).".format(ok)
        if not msg:
            msg = u"Geometría resuelta en {0} elemento(s).".format(ok)
        if inf_on:
            msg += (
                u"\n\nPlanta: {0:.0f} mm a caras laterales (cota); extremos (eje, c nominal "
                u"{1:.0f} mm + ø/2): {2:.0f} mm a lo largo de la curva."
            ).format(
                float(_RECUBRIMIENTO_PLANTA_CARAS_LATERALES_MM),
                float(RECUBRIMIENTO_EXTREMOS_MM),
                float(ext_mesh_inf_mm),
            )
            if lado_menor_mm_muestra is not None:
                msg += u"\nLongitud borde lado corto (muestra): {0:.0f} mm.".format(
                    lado_menor_mm_muestra
                )
            if lado_mayor_mm_muestra is not None:
                msg += u"\nLongitud borde lado largo (muestra): {0:.0f} mm.".format(
                    lado_mayor_mm_muestra
                )
            if lado_offset_mm_muestra is not None:
                msg += u" Tras offset (muestra en lado corto): {0:.0f} mm.".format(
                    lado_offset_mm_muestra
                )
            if z_mm_muestra is not None:
                msg += u"\nElevación aprox. cara inferior (muestra): {0:.0f} mm.".format(
                    z_mm_muestra
                )
            msg += (
                u"\nEje de barra (Rebar inferior): cara horizontal = "
                u"{0:.0f} mm al hormigón + radio nominal (ø/2)."
            ).format(float(_RECUBRIMIENTO_CARA_HORIZONTAL_MM))
            msg += (
                u"\nLado largo inferior: segunda capa — el eje se desplaza +ø nominal "
                u"hacia el interior del hormigón respecto a la capa del lado corto."
            )
            msg += u"\nSeparación inferior (mm): {0:.0f} — regla de layout Maximum Spacing con ese paso.".format(
                float(sep_mm)
            )
        if sup_on:
            msg += (
                u"\n\nSuperior — planta {0:.0f} mm a caras laterales (cota); extremos (eje, c nominal "
                u"{1:.0f} mm + ø/2): {2:.0f} mm."
            ).format(
                float(_RECUBRIMIENTO_PLANTA_CARAS_LATERALES_MM),
                float(RECUBRIMIENTO_EXTREMOS_MM),
                float(ext_mesh_sup_mm),
            )
            if lado_menor_mm_sup_muestra is not None:
                msg += u"\nLongitud borde lado corto superior (muestra): {0:.0f} mm.".format(
                    lado_menor_mm_sup_muestra
                )
            if lado_mayor_mm_sup_muestra is not None:
                msg += u"\nLongitud borde lado largo superior (muestra): {0:.0f} mm.".format(
                    lado_mayor_mm_sup_muestra
                )
            if z_sup_mm_muestra is not None:
                msg += u"\nElevación aprox. cara superior (muestra): {0:.0f} mm.".format(
                    z_sup_mm_muestra
                )
            msg += (
                u"\nEje de barra (Rebar superior): cara horizontal = "
                u"{0:.0f} mm al hormigón + radio nominal (ø/2)."
            ).format(float(_RECUBRIMIENTO_CARA_HORIZONTAL_MM))
            msg += (
                u"\nLado largo superior: segunda capa — +ø nominal respecto al lado corto."
            )
            msg += u"\nSeparación superior (mm): {0:.0f} — Maximum Spacing con ese paso.".format(
                float(sep_sup_mm)
            )
        if inf_on and sup_on:
            msg += (
                u"\n\nInferior + superior: polilínea en U sin ganchos de tipo en API — patas en dirección "
                u"+N (normal saliente) de la **cara inferior** (malla inferior) o de la **cara superior** (malla superior); "
                u"longitud de cada pata ≈ altura de la fundación − {0:.0f} mm. Si la polilínea o "
                u"CreateFromCurves falla, **no** se coloca barra recta con ganchos en ese modo."
            ).format(float(_DESCUENTO_PATA_U_INF_SUP_MM))
        elif inf_on and not sup_on:
            msg += (
                u"\n\nSolo malla inferior: misma polilínea en U (sin ganchos de tipo en API); "
                u"longitud de cada pata según **tabla BIMTools por diámetro nominal** (ø), "
                u"acotada a la altura útil de la fundación. Si la polilínea o CreateFromCurves falla, "
                u"se intenta barra recta con ganchos estándar en el eje."
            )
        if lat_on:
            msg += (
                u"\n\nLateral — barras horizontales desde el perímetro de la cara inferior "
                u"(offsets en planta); primera barra: eje a {3:.0f} mm de la cara inferior (hacia "
                u"dentro); CreateFromCurves sin RebarHookType (misma línea que malla inf./sup.: "
                u"barra recta modelada sin ganchos de tipo); eje de referencia para norm según cara "
                u"vertical paralela; curva perimetral: offset en planta = {0:.0f} mm + "
                u"ø malla inferior + ø lateral/2; recorte en cada extremo (a lo largo de la curva) = "
                u"{0:.0f} mm + ø malla inferior; "
                u"largo del array = altura de la fundación − {4:.0f} mm; extremos (referencia) {1:.0f} mm; "
                u"cantidad de barras en altura (Fixed Number): {2:d}."
            ).format(
                float(_RECUBRIMIENTO_LATERAL_CARA_MM),
                float(RECUBRIMIENTO_EXTREMOS_MM),
                int(n_lat_cant),
                float(_OFFSET_EJE_PRIMERA_BARRA_LATERAL_DESDE_CARA_INFERIOR_MM),
                float(_DESCUENTO_LARGO_ARRAY_LATERAL_MM),
            )
        if err_ids_lat:
            msg += u"\n\nAviso: no se halló perímetro inferior útil o altura nula en Id: {0}.".format(
                u", ".join(unicode(i) for i in err_ids_lat)
            )
        lineas_eval_par = []
        for el_ev, curva_ev, _m, _cara_pp, ev, lado_etq, _perp in elementos_ok:
            if curva_ev is None:
                continue
            if ev is None:
                continue
            eid = element_id_to_int(el_ev.Id)
            lbl_ld = u"lado corto" if lado_etq == u"menor" else u"lado largo"
            cand = ev.get("candidatos") or []
            mejor = ev.get("mejor")
            dft = ev.get("distancia_ft")
            nd = ev.get("descartados_coplanar") or 0
            if mejor is not None and dft is not None:
                lineas_eval_par.append(
                    u"Id {0} ({1}): {2} cara(s) paralelas a la curva (excl. coplanar); "
                    u"la más cercana — dist. punto medio al plano ≈ {3:.0f} mm "
                    u"(norm/propagación: BasisZ de esa cara → plano barra).".format(
                        eid,
                        lbl_ld,
                        len(cand),
                        float(dft) * 304.8,
                    )
                )
            else:
                lineas_eval_par.append(
                    u"Id {0} ({1}): ninguna cara paralela no coplanar "
                    u"(coplanares descartadas: {2}).".format(eid, lbl_ld, nd)
                )
        for el_ev, curva_ev, _m, _cara_pp, ev, lado_etq, _perp in elementos_ok_sup:
            if curva_ev is None:
                continue
            if ev is None:
                continue
            eid = element_id_to_int(el_ev.Id)
            lbl_ld = (
                u"sup. lado corto" if lado_etq == u"menor" else u"sup. lado largo"
            )
            cand = ev.get("candidatos") or []
            mejor = ev.get("mejor")
            dft = ev.get("distancia_ft")
            nd = ev.get("descartados_coplanar") or 0
            if mejor is not None and dft is not None:
                lineas_eval_par.append(
                    u"Id {0} ({1}): {2} cara(s) paralelas a la curva (excl. coplanar); "
                    u"la más cercana — dist. punto medio al plano ≈ {3:.0f} mm "
                    u"(norm/propagación: BasisZ de esa cara → plano barra).".format(
                        eid,
                        lbl_ld,
                        len(cand),
                        float(dft) * 304.8,
                    )
                )
            else:
                lineas_eval_par.append(
                    u"Id {0} ({1}): ninguna cara paralela no coplanar "
                    u"(coplanares descartadas: {2}).".format(eid, lbl_ld, nd)
                )
        if lineas_eval_par:
            msg += (
                u"\n\nEvaluación: entre caras con plano paralelo a la tangente de la curva, "
                u"la más cercana al eje (dist. perpendicular):\n"
            )
            msg += u"\n".join(lineas_eval_par[:20])
            if len(lineas_eval_par) > 20:
                msg += u"\n…"

        n_rebar = 0
        n_rebar_sup = 0
        n_rebar_lat = 0
        lineas_norm_msg = []
        rebar_avisos = []
        rebar_ids_armadura_largo_total = []
        malla_fallback_ganchos = False

        # --- Pre-transacción: detectar si se necesita propagación (lectura pura) ---
        _a_prop_pre = []
        _num_val_pre = None
        if not getattr(win, "_saltar_propagacion_post", False):
            try:
                from numeracion_fundacion import (
                    buscar_fundaciones_por_numeracion,
                    leer_numeracion_fundacion,
                )
                for _eid_pre in ids_run:
                    _eln_pre = el_by_id.get(_eid_pre)
                    if _eln_pre is None:
                        continue
                    _num_val_pre = leer_numeracion_fundacion(_eln_pre)
                    if _num_val_pre is not None:
                        break
                if _num_val_pre is not None:
                    _todos_pre = buscar_fundaciones_por_numeracion(doc, _num_val_pre)
                    _ya_pre = set(element_id_to_int(x) for x in ids_run)
                    for _x_pre in _todos_pre:
                        if element_id_to_int(_x_pre) in _ya_pre:
                            continue
                        try:
                            _elp_pre = doc.GetElement(_x_pre)
                            if _elp_pre is None or isinstance(_elp_pre, WallFoundation):
                                continue
                            _cat_pre = _elp_pre.Category
                            if (
                                _cat_pre is None
                                or element_id_to_int(_cat_pre.Id) != _FOUNDATION_CAT_ID
                            ):
                                continue
                        except Exception:
                            continue
                        _a_prop_pre.append(_x_pre)
            except Exception:
                pass

        # Si hay fundaciones pendientes de propagación, las vistas se crean en el
        # segundo ciclo (propagación). De lo contrario, se crean en este mismo ciclo.
        _crear_vistas_en_este_ciclo = (
            getattr(win, "_saltar_propagacion_post", False) or not _a_prop_pre
        )

        _ultima_vista = None
        t = Transaction(doc, u"BIMTools: Armadura fundacion aislada")
        try:
            t.Start()
        except Exception as ex:
            _task_dialog_show(
                u"BIMTools — Armadura Fundacion Aislada",
                u"No se pudo iniciar la transacción:\n{0}".format(ex),
                win._win,
            )
            try:
                win._set_estado(u"Error de transacción.")
            except Exception:
                pass
            return
        try:
            _stx_rebar = SubTransaction(doc)
            _stx_rebar.Start()
            # Polilínea U en inferior: siempre que la malla inferior esté activa (igual criterio
            # que inf.+sup.; solo inferior usa largo de pata desde tabla ø vía _leg_ft_pata_u_malla_inferior).
            malla_inf_y_sup = inf_on and sup_on
            bt_inferior_cache = None
            if inf_on:
                bt, err_bt = _bt_inf_pre, _err_bt_inf_pre
                bt_inferior_cache = bt
                if bt is None:
                    rebar_avisos.append(
                        err_bt
                        or u"No hay RebarBarType inferior; no se creó barra."
                    )
                else:
                    for el, curva_tratada, marco_uvn, cara_pp, _ev_par, lado_etq, perp_len_mm in elementos_ok:
                        if curva_tratada is None:
                            rebar_avisos.append(
                                u"Id {0}: curva nula tras recubrimiento.".format(
                                    element_id_to_int(el.Id)
                                )
                            )
                            continue
                        lbl_ld = (
                            u"lado corto"
                            if lado_etq == u"menor"
                            else u"lado largo"
                        )
                        d_mm_bar = _rebar_nominal_diameter_mm(bt)
                        if d_mm_bar is None:
                            d_mm_bar = 0.0
                        n_cara = marco_uvn[3] if marco_uvn is not None and len(marco_uvn) > 3 else None
                        curva_rebar = offset_linea_eje_barra_desde_cara_inferior_mm(
                            curva_tratada,
                            n_cara,
                            _RECUBRIMIENTO_CARA_HORIZONTAL_MM,
                            d_mm_bar,
                        )
                        # Lado largo: segunda capa sobre la del lado corto; separación vertical = ø nominal.
                        if lado_etq == u"mayor" and d_mm_bar > 1e-9:
                            curva_rebar = offset_linea_adicional_hacia_interior_mm(
                                curva_rebar,
                                n_cara,
                                d_mm_bar,
                            )
                        z_hook_inf = vector_reverso_cara_paralela_mas_cercana_a_barra(
                            el, curva_rebar
                        )
                        r, err_rb, norm_rb = None, None, None
                        linea_marcador = curva_rebar
                        poli_u = None
                        try:
                            eid_el = el.Id
                        except Exception:
                            eid_el = None
                        z0p, z1p = _z_range_fundacion_para_patas(
                            el, eid_el, fz_range_by_eid
                        )
                        leg_ft = _leg_ft_pata_u_malla_inferior(
                            z0p, z1p, d_mm_bar, sup_on, elem=el
                        )
                        if leg_ft is not None and n_cara is not None:
                            # Cota al eje (c + ø/2) ya en curva_tratada/extremos; no acortar otra vez ø/2
                            # en el tramo central (misma regla que Wall Foundation / doc. geometría).
                            poli_u = construir_polilinea_u_fundacion_desde_eje_horizontal(
                                curva_rebar,
                                n_cara,
                                leg_ft,
                                d_mm_bar,
                                acortar_eje_central_para_cota_revit=False,
                            )
                        if poli_u is not None:
                            r, err_rb, norm_rb = (
                                crear_rebar_u_shape_desde_eje_rebar_shape_nombrado(
                                    doc,
                                    el,
                                    bt,
                                    poli_u,
                                    shape_nombre=REBAR_SHAPE_NOMBRE_DEFECTO,
                                    marco_cara_uvn=marco_uvn,
                                    cara_paralela=cara_pp,
                                    eje_referencia_z_ganchos=z_hook_inf,
                                )
                            )
                            if r is None:
                                r, err_rb, norm_rb = (
                                    crear_rebar_polilinea_u_malla_inf_sup_curve_loop(
                                        doc,
                                        el,
                                        bt,
                                        poli_u,
                                        poli_u[1],
                                        marco_cara_uvn=marco_uvn,
                                        cara_paralela=cara_pp,
                                        eje_referencia_z_ganchos=z_hook_inf,
                                    )
                                )
                            if r is None:
                                r, err_rb, norm_rb = (
                                    crear_rebar_polilinea_recta_sin_ganchos(
                                        doc,
                                        el,
                                        bt,
                                        poli_u,
                                        poli_u[1],
                                        marco_cara_uvn=marco_uvn,
                                        cara_paralela=cara_pp,
                                        eje_referencia_z_ganchos=z_hook_inf,
                                    )
                                )
                            if r is not None:
                                linea_marcador = poli_u[1]
                        if r is None and err_rb is None:
                            err_rb = (
                                u"No se pudo aplicar la polilínea U (revisar altura y "
                                u"normal cara inferior)."
                                if poli_u is None
                                else u"CreateFromCurves (polilínea) no generó la barra."
                            )
                        if r is None:
                            r, err_rb, norm_rb = crear_rebar_desde_curva_linea_con_ganchos(
                                doc,
                                el,
                                bt,
                                curva_rebar,
                                marco_cara_uvn=marco_uvn,
                                cara_paralela=cara_pp,
                                eje_referencia_z_ganchos=z_hook_inf,
                            )
                            if r is not None:
                                linea_marcador = curva_rebar
                                malla_fallback_ganchos = True
                                rebar_avisos.append(
                                    u"Id {0} ({1}): polilínea U rechazada por Revit; "
                                    u"se usó barra con ganchos en el eje.".format(
                                        element_id_to_int(el.Id),
                                        lbl_ld,
                                    )
                                )
                        if r is not None:
                            _aplicar_armadura_ubicacion(r, _ARMA_UBICACION_INFERIOR)
                            n_rebar += 1
                            if perp_len_mm is not None:
                                try:
                                    from Autodesk.Revit.DB import UnitUtils, UnitTypeId
                                    _span_mm = max(float(perp_len_mm) - 2.0 * float(rec_mm), 0.01)
                                    array_len_ft = float(
                                        UnitUtils.ConvertToInternalUnits(_span_mm, UnitTypeId.Millimeters)
                                    )
                                except Exception:
                                    array_len_ft = longitud_distribucion_perpendicular_barra_inferior_ft(
                                        el, curva_tratada, rec_mm, lado_etq
                                    )
                            else:
                                array_len_ft = longitud_distribucion_perpendicular_barra_inferior_ft(
                                    el, curva_tratada, rec_mm, lado_etq
                                )
                            _ok_lay, err_lay = aplicar_layout_maximum_spacing_rebar(
                                r, doc, sep_mm, array_len_ft
                            )
                            if not _ok_lay and err_lay:
                                rebar_avisos.append(
                                    u"Id {0} ({1}): layout Maximum Spacing: {2}".format(
                                        element_id_to_int(el.Id),
                                        lbl_ld,
                                        err_lay,
                                    )
                                )
                            try:
                                rebar_ids_armadura_largo_total.append(r.Id)
                            except Exception:
                                pass
                            if norm_rb is not None:
                                try:
                                    nu = norm_rb.Normalize()
                                    lineas_norm_msg.append(
                                        u"Id {0} ({1}): norm = ({2:.4f}, {3:.4f}, {4:.4f})".format(
                                            element_id_to_int(el.Id),
                                            lbl_ld,
                                            float(nu.X),
                                            float(nu.Y),
                                            float(nu.Z),
                                        )
                                    )
                                except Exception:
                                    lineas_norm_msg.append(
                                        u"Id {0} ({1}): norm (CreateFromCurves).".format(
                                            element_id_to_int(el.Id),
                                            lbl_ld,
                                        )
                                    )
                        elif err_rb:
                            rebar_avisos.append(
                                u"Id {0} ({1}): {2}".format(
                                    element_id_to_int(el.Id),
                                    lbl_ld,
                                    err_rb,
                                )
                            )
            if sup_on:
                bt_sup, err_bt_sup = _bt_sup_pre, _err_bt_sup_pre
                if bt_sup is None:
                    rebar_avisos.append(
                        err_bt_sup
                        or u"No hay RebarBarType superior; no se creó barra."
                    )
                else:
                    for (
                        el,
                        curva_tratada,
                        marco_uvn,
                        cara_pp,
                        _ev_par,
                        lado_etq,
                        perp_len_mm,
                    ) in elementos_ok_sup:
                        if curva_tratada is None:
                            rebar_avisos.append(
                                u"Id {0}: curva nula tras recubrimiento (superior).".format(
                                    element_id_to_int(el.Id)
                                )
                            )
                            continue
                        lbl_ld = (
                            u"sup. lado corto"
                            if lado_etq == u"menor"
                            else u"sup. lado largo"
                        )
                        d_mm_bar = _rebar_nominal_diameter_mm(bt_sup)
                        if d_mm_bar is None:
                            d_mm_bar = 0.0
                        n_cara = (
                            marco_uvn[3]
                            if marco_uvn is not None and len(marco_uvn) > 3
                            else None
                        )
                        curva_rebar = offset_linea_eje_barra_desde_cara_inferior_mm(
                            curva_tratada,
                            n_cara,
                            _RECUBRIMIENTO_CARA_HORIZONTAL_MM,
                            d_mm_bar,
                        )
                        if lado_etq == u"mayor" and d_mm_bar > 1e-9:
                            curva_rebar = offset_linea_adicional_hacia_interior_mm(
                                curva_rebar,
                                n_cara,
                                d_mm_bar,
                            )
                        z_hook_sup = vector_reverso_cara_paralela_mas_cercana_a_barra(
                            el, curva_rebar
                        )
                        if z_hook_sup is None:
                            z_hook_sup = XYZ(0.0, 0.0, -1.0)
                        try:
                            eid_sup = el.Id
                        except Exception:
                            eid_sup = None
                        r, err_rb, norm_rb = None, None, None
                        linea_marcador_sup = curva_rebar
                        poli_u_sup = None
                        if malla_inf_y_sup:
                            z0p, z1p = _z_range_fundacion_para_patas(
                                el, eid_sup, fz_range_by_eid
                            )
                            h_param_ft = _altura_nominal_fundacion_ft(el)
                            try:
                                _desc_sup_mm = float(_DESCUENTO_PATA_U_INF_SUP_MM) + float(d_mm_bar) / 2.0
                            except Exception:
                                _desc_sup_mm = float(_DESCUENTO_PATA_U_INF_SUP_MM)
                            if h_param_ft is not None:
                                leg_ft = longitud_pata_u_fundacion_inf_sup_ft(
                                    0.0, h_param_ft, _desc_sup_mm
                                )
                            else:
                                leg_ft = longitud_pata_u_fundacion_inf_sup_ft(
                                    z0p, z1p, _desc_sup_mm
                                )
                            if leg_ft is not None and n_cara is not None:
                                poli_u_sup = construir_polilinea_u_fundacion_desde_eje_horizontal(
                                    curva_rebar,
                                    n_cara,
                                    leg_ft,
                                    d_mm_bar,
                                    acortar_eje_central_para_cota_revit=False,
                                )
                            if poli_u_sup is not None:
                                r, err_rb, norm_rb = (
                                    crear_rebar_u_shape_desde_eje_rebar_shape_nombrado(
                                        doc,
                                        el,
                                        bt_sup,
                                        poli_u_sup,
                                        shape_nombre=REBAR_SHAPE_NOMBRE_DEFECTO,
                                        marco_cara_uvn=marco_uvn,
                                        cara_paralela=cara_pp,
                                        eje_referencia_z_ganchos=z_hook_sup,
                                    )
                                )
                                if r is None:
                                    r, err_rb, norm_rb = (
                                        crear_rebar_polilinea_u_malla_inf_sup_curve_loop(
                                            doc,
                                            el,
                                            bt_sup,
                                            poli_u_sup,
                                            poli_u_sup[1],
                                            marco_cara_uvn=marco_uvn,
                                            cara_paralela=cara_pp,
                                            eje_referencia_z_ganchos=z_hook_sup,
                                        )
                                    )
                                if r is None:
                                    r, err_rb, norm_rb = (
                                        crear_rebar_polilinea_recta_sin_ganchos(
                                            doc,
                                            el,
                                            bt_sup,
                                            poli_u_sup,
                                            poli_u_sup[1],
                                            marco_cara_uvn=marco_uvn,
                                            cara_paralela=cara_pp,
                                            eje_referencia_z_ganchos=z_hook_sup,
                                        )
                                    )
                                if r is not None:
                                    linea_marcador_sup = poli_u_sup[1]
                            if r is None and err_rb is None:
                                err_rb = (
                                    u"No se pudo aplicar la polilínea U (revisar altura y "
                                    u"normal cara superior)."
                                    if poli_u_sup is None
                                    else u"CreateFromCurves (polilínea) no generó la barra."
                                )
                            if r is None:
                                r, err_rb, norm_rb = crear_rebar_desde_curva_linea_con_ganchos(
                                    doc,
                                    el,
                                    bt_sup,
                                    curva_rebar,
                                    marco_cara_uvn=marco_uvn,
                                    cara_paralela=cara_pp,
                                    eje_referencia_z_ganchos=z_hook_sup,
                                )
                                if r is not None:
                                    linea_marcador_sup = curva_rebar
                                    malla_fallback_ganchos = True
                                    rebar_avisos.append(
                                        u"Id {0} ({1}): polilínea U rechazada por Revit; "
                                        u"se usó barra con ganchos en el eje.".format(
                                            element_id_to_int(el.Id),
                                            lbl_ld,
                                        )
                                    )
                        else:
                            r, err_rb, norm_rb = crear_rebar_desde_curva_linea_con_ganchos(
                                doc,
                                el,
                                bt_sup,
                                curva_rebar,
                                marco_cara_uvn=marco_uvn,
                                cara_paralela=cara_pp,
                                eje_referencia_z_ganchos=z_hook_sup,
                            )
                            linea_marcador_sup = curva_rebar
                        if r is not None:
                            _aplicar_armadura_ubicacion(r, _ARMA_UBICACION_SUPERIOR)
                            n_rebar_sup += 1
                            if perp_len_mm is not None:
                                try:
                                    from Autodesk.Revit.DB import UnitUtils, UnitTypeId
                                    _span_mm = max(float(perp_len_mm) - 2.0 * float(rec_sup_mm), 0.01)
                                    array_len_ft = float(
                                        UnitUtils.ConvertToInternalUnits(_span_mm, UnitTypeId.Millimeters)
                                    )
                                except Exception:
                                    array_len_ft = longitud_distribucion_perpendicular_barra_inferior_ft(
                                        el, curva_tratada, rec_sup_mm, lado_etq
                                    )
                            else:
                                array_len_ft = longitud_distribucion_perpendicular_barra_inferior_ft(
                                    el, curva_tratada, rec_sup_mm, lado_etq
                                )
                            _ok_lay, err_lay = aplicar_layout_maximum_spacing_rebar(
                                r, doc, sep_sup_mm, array_len_ft
                            )
                            if not _ok_lay and err_lay:
                                rebar_avisos.append(
                                    u"Id {0} ({1}): layout Maximum Spacing: {2}".format(
                                        element_id_to_int(el.Id),
                                        lbl_ld,
                                        err_lay,
                                    )
                                )
                            try:
                                rebar_ids_armadura_largo_total.append(r.Id)
                            except Exception:
                                pass
                            if norm_rb is not None:
                                try:
                                    nu = norm_rb.Normalize()
                                    lineas_norm_msg.append(
                                        u"Id {0} ({1}): norm = ({2:.4f}, {3:.4f}, {4:.4f})".format(
                                            element_id_to_int(el.Id),
                                            lbl_ld,
                                            float(nu.X),
                                            float(nu.Y),
                                            float(nu.Z),
                                        )
                                    )
                                except Exception:
                                    lineas_norm_msg.append(
                                        u"Id {0} ({1}): norm (CreateFromCurves).".format(
                                            element_id_to_int(el.Id),
                                            lbl_ld,
                                        )
                                    )
                        elif err_rb:
                            rebar_avisos.append(
                                u"Id {0} ({1}): {2}".format(
                                    element_id_to_int(el.Id),
                                    lbl_ld,
                                    err_rb,
                                )
                            )
            if lat_on:
                bt_lat, err_bt_lat = _bt_lat_pre, _err_bt_lat_pre
                if bt_lat is None:
                    rebar_avisos.append(
                        err_bt_lat
                        or u"No hay RebarBarType lateral; no se creó barra lateral."
                    )
                else:
                    d_mm_lat = _rebar_nominal_diameter_mm(bt_lat)
                    if d_mm_lat is None:
                        d_mm_lat = 0.0
                    # bt_inferior_cache ya fue resuelto antes del bucle de geometría.
                    bt_inf_malla = bt_inferior_cache if bt_inferior_cache is not None else _bt_inf_pre
                    d_mm_inf_malla = (
                        _rebar_nominal_diameter_mm(bt_inf_malla)
                        if bt_inf_malla is not None
                        else None
                    )
                    if d_mm_inf_malla is None:
                        d_mm_inf_malla = 0.0
                    off_planta_lat_mm = (
                        float(_RECUBRIMIENTO_LATERAL_CARA_MM)
                        + float(d_mm_inf_malla)
                        + 0.5 * float(d_mm_lat)
                    )
                    recorte_extremos_lat_mm = (
                        float(_RECUBRIMIENTO_LATERAL_CARA_MM)
                        + float(d_mm_inf_malla)
                    )
                    for eid in ids_run:
                        el = el_by_id.get(eid)
                        if el is None:
                            continue
                        if eid in marco_inf_by_eid:
                            marco_inf = marco_inf_by_eid[eid]
                        else:
                            try:
                                marco_inf = obtener_marco_coordenadas_cara_inferior(
                                    el
                                )
                            except Exception:
                                marco_inf = None
                            marco_inf_by_eid[eid] = marco_inf
                        n_inferior = None
                        if marco_inf is not None and len(marco_inf) > 3:
                            try:
                                ni = marco_inf[3]
                                if ni is not None and float(ni.GetLength()) > 1e-12:
                                    n_inferior = ni.Normalize()
                            except Exception:
                                n_inferior = None
                        lh_pack = lh_pack_by_eid.get(eid)
                        fz_min = None
                        fz_max = None
                        _fz = fz_range_by_eid.get(eid)
                        if _fz is not None:
                            fz_min, fz_max = _fz
                        if lh_pack is None or fz_min is None or fz_max is None:
                            continue
                        lineas_borde, _z_inf = lh_pack
                        if not lineas_borde:
                            continue
                        array_len_lat_ft = (
                            longitud_array_lateral_altura_fundacion_menos_mm_ft(
                                fz_min,
                                fz_max,
                                _DESCUENTO_LARGO_ARRAY_LATERAL_MM,
                            )
                        )
                        for line_borde in lineas_borde:
                            # Planta: 50 + ø inf. + ø lateral/2; recorte por extremo (longitud): 50 + ø inf.
                            curva_tratada, _co = aplicar_recubrimiento_inferior_completo_mm(
                                line_borde,
                                el,
                                off_planta_lat_mm,
                                recorte_extremos_lat_mm,
                            )
                            if curva_tratada is None:
                                continue
                            n_horiz_out = (
                                normal_saliente_horizontal_paramento_para_barra_horizontal(
                                    curva_tratada,
                                    el,
                                )
                            )
                            if n_horiz_out is None:
                                continue
                            curva_eje = curva_tratada
                            if n_inferior is not None:
                                curva_rebar = (
                                    offset_linea_hacia_interior_desde_cara_inferior_mm(
                                        curva_eje,
                                        n_inferior,
                                        _OFFSET_EJE_PRIMERA_BARRA_LATERAL_DESDE_CARA_INFERIOR_MM,
                                    )
                                )
                            else:
                                z_fb = primera_cota_z_armadura_lateral_ft(
                                    fz_min,
                                    fz_max,
                                    _RECUBRIMIENTO_LATERAL_CARA_MM,
                                    d_mm_lat,
                                )
                                curva_rebar = linea_horizontal_cara_lateral_a_cota_z(
                                    curva_eje,
                                    z_fb,
                                )
                            if curva_rebar is None:
                                continue
                            z_hook = vector_reverso_cara_paralela_mas_cercana_a_barra(
                                el,
                                curva_rebar,
                                excluir_caras_tapas_horizontales=True,
                            )
                            if z_hook is None:
                                try:
                                    zh = n_horiz_out.Negate()
                                    if (
                                        zh is not None
                                        and float(zh.GetLength()) > 1e-12
                                    ):
                                        z_hook = zh.Normalize()
                                except Exception:
                                    z_hook = None
                            # Dirección del conjunto (norm CreateFromCurves): normal **reversa**
                            # (opuesta a la saliente) de la cara inferior.
                            n_inf_rev = None
                            if n_inferior is not None:
                                try:
                                    nr = n_inferior.Negate()
                                    if nr is not None and float(nr.GetLength()) > 1e-12:
                                        n_inf_rev = nr.Normalize()
                                except Exception:
                                    n_inf_rev = None
                            norm_pri_lat = (
                                [n_inf_rev] if n_inf_rev is not None else None
                            )
                            # Gancho lateral via geometría (igual que malla inf/sup):
                            # polilínea U con patas horizontales hacia el interior.
                            hook_lat_mm = largo_gancho_u_tabla_mm(d_mm_lat)
                            leg_ft_lat = None
                            if hook_lat_mm is not None:
                                # Restar ø/2: el eje de la polilínea se traza hasta el
                                # centro de la barra, que queda ø/2 más corto que la
                                # cota de tabla (misma regla que malla inf/sup).
                                try:
                                    _d_round = float(int(round(float(d_mm_lat)))) if d_mm_lat else 0.0
                                    eje_lat_mm = float(hook_lat_mm) - 0.5 * _d_round
                                    eje_lat_mm = max(eje_lat_mm, 40.0)
                                except Exception:
                                    eje_lat_mm = float(hook_lat_mm)
                                try:
                                    from Autodesk.Revit.DB import UnitUtils, UnitTypeId
                                    leg_ft_lat = UnitUtils.ConvertToInternalUnits(
                                        eje_lat_mm, UnitTypeId.Millimeters
                                    )
                                except Exception:
                                    leg_ft_lat = eje_lat_mm / 304.8
                            r, err_rb, _norm_rb = None, None, None
                            if leg_ft_lat is not None:
                                # Acortar eje central ø/2 por cada extremo (total ø)
                                curva_rebar_lat = curva_rebar
                                try:
                                    from Autodesk.Revit.DB import UnitUtils, UnitTypeId as _UTI
                                    _half_d_ft = UnitUtils.ConvertToInternalUnits(
                                        float(d_mm_lat) / 2.0, _UTI.Millimeters
                                    )
                                except Exception:
                                    _half_d_ft = float(d_mm_lat) / 2.0 / 304.8
                                try:
                                    _p0 = curva_rebar.GetEndPoint(0)
                                    _p1 = curva_rebar.GetEndPoint(1)
                                    _tang = (_p1 - _p0)
                                    _tlen = float(_tang.GetLength())
                                    if _tlen > 2.0 * _half_d_ft + 1e-6:
                                        _tu = _tang.Multiply(1.0 / _tlen)
                                        _new_p0 = _p0 + _tu.Multiply(_half_d_ft)
                                        _new_p1 = _p1 - _tu.Multiply(_half_d_ft)
                                        curva_rebar_lat = Line.CreateBound(_new_p0, _new_p1)
                                except Exception:
                                    curva_rebar_lat = curva_rebar
                                poli_u_lat = construir_polilinea_u_fundacion_desde_eje_horizontal(
                                    curva_rebar_lat,
                                    n_horiz_out,
                                    leg_ft_lat,
                                    d_mm_lat,
                                    acortar_eje_central_para_cota_revit=False,
                                )
                                if poli_u_lat is not None:
                                    r, err_rb, _norm_rb = (
                                        crear_rebar_u_shape_desde_eje_rebar_shape_nombrado(
                                            doc,
                                            el,
                                            bt_lat,
                                            poli_u_lat,
                                            shape_nombre=REBAR_SHAPE_NOMBRE_DEFECTO,
                                            marco_cara_uvn=marco_inf,
                                            cara_paralela=None,
                                            eje_referencia_z_ganchos=z_hook,
                                        )
                                    )
                                    if r is None:
                                        r, err_rb, _norm_rb = (
                                            crear_rebar_polilinea_u_malla_inf_sup_curve_loop(
                                                doc,
                                                el,
                                                bt_lat,
                                                poli_u_lat,
                                                poli_u_lat[1],
                                                marco_cara_uvn=marco_inf,
                                                cara_paralela=None,
                                                eje_referencia_z_ganchos=z_hook,
                                            )
                                        )
                            # Fallback: barra recta sin ganchos
                            if r is None:
                                r, err_rb, _norm_rb = crear_rebar_polilinea_recta_sin_ganchos(
                                    doc,
                                    el,
                                    bt_lat,
                                    [curva_rebar],
                                    curva_rebar,
                                    marco_cara_uvn=marco_inf,
                                    cara_paralela=None,
                                    eje_referencia_z_ganchos=z_hook,
                                    normales_prioridad=norm_pri_lat,
                                )
                            if r is not None:
                                _aplicar_armadura_ubicacion(r, _ARMA_UBICACION_LATERAL)
                                n_rebar_lat += 1
                                _ok_lay_lat, err_lay_lat = (
                                    aplicar_layout_fixed_number_rebar(
                                        r,
                                        doc,
                                        n_lat_cant,
                                        array_len_lat_ft,
                                    )
                                )
                                if not _ok_lay_lat and err_lay_lat:
                                    rebar_avisos.append(
                                        u"Id {0} (lateral): layout Fixed Number (altura): {1}".format(
                                            element_id_to_int(el.Id),
                                            err_lay_lat,
                                        )
                                    )
                                try:
                                    rebar_ids_armadura_largo_total.append(r.Id)
                                except Exception:
                                    pass
                            elif err_rb:
                                rebar_avisos.append(
                                    u"Id {0} (lateral): {1}".format(
                                        element_id_to_int(el.Id),
                                        err_rb,
                                    )
                                )
            if rebar_ids_armadura_largo_total:
                try:
                    from enfierrado_shaft_hashtag import (
                        _apply_armadura_largo_total_to_rebars,
                    )

                    _apply_armadura_largo_total_to_rebars(
                        doc, rebar_ids_armadura_largo_total, rebar_avisos
                    )
                except Exception:
                    pass
            _stx_rebar.Commit()

            # --- Vistas y secciones dentro de la misma transacción principal ---
            _ultima_vista = None
            if _crear_vistas_en_este_ciclo:
                try:
                    from vista_seccion_enfierrado_vigas import (
                        crear_vista_planta_fundacion_aislada,
                        crear_secciones_fundacion_aislada,
                    )
                    _elem_vista = None
                    for _eid_v in (win._foundation_ids or []):
                        _elem_vista = doc.GetElement(_eid_v)
                        if _elem_vista is not None:
                            break
                    if _elem_vista is None:
                        for _eid_v in ids_run:
                            _elem_vista = el_by_id.get(_eid_v)
                            if _elem_vista is not None:
                                break
                    if _elem_vista is not None:
                        _vistas_a_crear = []
                        if inf_on and sup_on:
                            _vistas_a_crear = [
                                (_ARMA_UBICACION_INFERIOR, True),
                                (_ARMA_UBICACION_SUPERIOR, True),
                            ]
                        elif inf_on:
                            _vistas_a_crear = [(_ARMA_UBICACION_INFERIOR, True)]
                        elif sup_on:
                            _vistas_a_crear = [(_ARMA_UBICACION_SUPERIOR, True)]
                        else:
                            _vistas_a_crear = [(None, True)]
                        _ultima_vista = None
                        for _ub, _abrir in _vistas_a_crear:
                            _vista_planta, _av_planta = crear_vista_planta_fundacion_aislada(
                                doc,
                                _elem_vista,
                                uidocument=None,
                                gestionar_transaccion=False,
                                ubicacion_armadura=_ub,
                            )
                            if _av_planta:
                                rebar_avisos.append(u"Vista de planta: {0}".format(_av_planta))
                            if _vista_planta is not None:
                                _ultima_vista = _vista_planta
                        _vistas_sec, _av_sec = crear_secciones_fundacion_aislada(
                            doc,
                            _elem_vista,
                            uidocument=None,
                            gestionar_transaccion=False,
                        )
                        for _av in _av_sec:
                            rebar_avisos.append(u"Sección: {0}".format(_av))
                except Exception as _ex_vistas:
                    rebar_avisos.append(u"Vistas no creadas: {0}".format(_ex_vistas))

            t.Commit()
        except Exception as ex:
            try:
                _stx_rebar.RollBack()
            except Exception:
                pass
            try:
                t.RollBack()
            except Exception:
                pass
            _task_dialog_show(
                u"BIMTools — Armadura Fundacion Aislada",
                u"Error al crear armadura o detalle:\n{0}".format(ex),
                win._win,
            )
            try:
                win._set_estado(u"Error al colocar armadura.")
            except Exception:
                pass
            return

        # --- Post-commit: activar última vista + gestionar propagación ---
        _ultima_vista_post = _ultima_vista
        if _ultima_vista_post is not None and uidoc is not None:
            try:
                uidoc.ActiveView = _ultima_vista_post
            except Exception:
                pass

        if not getattr(win, "_saltar_propagacion_post", False):
            if _a_prop_pre:
                try:
                    from System.Windows import Visibility
                    win._propagacion_ids_a_confirmar = _a_prop_pre
                    win._propagacion_num_val = _num_val_pre
                    win._propagacion_pendiente_ui = True
                    br = win._win.FindName("BorderPropagacion")
                    _tx_lbl = win._win.FindName("TxtPropagacionTitulo")
                    if _tx_lbl is not None:
                        _tx_lbl.Text = (
                            u"{0} Fundaciones con Numeracion {1} encontradas."
                        ).format(
                            len(_a_prop_pre),
                            _numeracion_etiqueta_propagacion(_num_val_pre),
                        )
                    if br is not None:
                        br.Visibility = Visibility.Visible
                    btn_pr = win._win.FindName("BtnColocar")
                    if btn_pr is not None:
                        btn_pr.Content = _CAPTION_BTN_PROPAGAR_ARMADURAS
                    try:
                        seen = set()
                        merged_ids = []
                        for eid in list(ids_run) + list(_a_prop_pre):
                            k = element_id_to_int(eid)
                            if k in seen:
                                continue
                            seen.add(k)
                            merged_ids.append(eid)
                        win._refresh_laterales_cantidad_desde_ids(merged_ids)
                    except Exception:
                        pass
                    try:
                        win._refit_window_after_footer_change(True)
                    except Exception:
                        pass
                except Exception:
                    pass
        else:
            try:
                win._saltar_propagacion_post = False
            except Exception:
                pass

        if inf_on:
            if malla_inf_y_sup:
                if malla_fallback_ganchos:
                    msg += (
                        u"\n\nBarras inferiores creadas: {0} "
                        u"(inf.+sup.: la API no aceptó la polilínea U en este modelo; "
                        u"se usaron ganchos «{1}» en el eje — igual que modo solo inferior)."
                    ).format(n_rebar, HOOK_GANCHO_90_STANDARD_NAME)
                else:
                    msg += (
                        u"\n\nBarras inferiores creadas: {0} "
                        u"(inf.+sup.: polilínea U sin ganchos de tipo — patas según +N cara inferior)."
                    ).format(n_rebar)
            else:
                if malla_fallback_ganchos:
                    msg += (
                        u"\n\nBarras inferiores creadas: {0} "
                        u"(solo inferior: la API no aceptó la polilínea U; "
                        u"se usaron ganchos «{1}» en el eje)."
                    ).format(n_rebar, HOOK_GANCHO_90_STANDARD_NAME)
                else:
                    msg += (
                        u"\n\nBarras inferiores creadas: {0} "
                        u"(solo inferior: polilínea U sin ganchos de tipo — patas según tabla ø, "
                        u"acotadas a altura útil)."
                    ).format(n_rebar)
        else:
            msg += u"\n\nGrupo inferior desactivado: no se crearon barras inferiores."
        if sup_on:
            if malla_inf_y_sup:
                if malla_fallback_ganchos:
                    msg += (
                        u"\n\nBarras superiores creadas: {0} "
                        u"(inf.+sup.: la API no aceptó la polilínea U en este modelo; "
                        u"se usaron ganchos «{1}» en el eje.)."
                    ).format(n_rebar_sup, HOOK_GANCHO_90_STANDARD_NAME)
                else:
                    msg += (
                        u"\n\nBarras superiores creadas: {0} "
                        u"(inf.+sup.: polilínea U sin ganchos de tipo — patas según +N cara superior)."
                    ).format(n_rebar_sup)
            else:
                msg += u"\n\nBarras superiores creadas: {0} (ganchos «{1}» en ambos extremos).".format(
                    n_rebar_sup,
                    HOOK_GANCHO_90_STANDARD_NAME,
                )
        else:
            msg += u"\n\nGrupo superior desactivado: no se crearon barras superiores."
        if lat_on:
            msg += (
                u"\n\nBarras laterales (perímetro cara inferior) creadas: {0} "
                u"(CreateFromCurves sin RebarHookType — barra recta, sin ganchos de tipo en API)."
            ).format(n_rebar_lat)
        else:
            msg += u"\n\nGrupo lateral desactivado: no se crearon barras laterales."
        if lineas_norm_msg:
            msg += u"\n\nVector norm (CreateFromCurves), unitario:\n" + u"\n".join(
                lineas_norm_msg
            )
        if rebar_avisos:
            msg += u"\n" + u"\n".join(rebar_avisos[:10])

        try:
            est = u"Inferior: {0} — Superior: {1} — Lateral: {2}.".format(
                n_rebar,
                n_rebar_sup,
                n_rebar_lat,
            )
            if getattr(win, "_propagacion_pendiente_ui", False):
                est += u" — Propagación: pulse «{0}».".format(_CAPTION_BTN_PROPAGAR_ARMADURAS)
            win._set_estado(est)
        except Exception:
            pass

    def GetName(self):
        return u"ColocarArmaduraFundacionAislada"


def _clear_appdomain_window_key():
    try:
        System.AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, None)
    except Exception:
        pass


def _get_active_window():
    """Devuelve la ventana solo si sigue cargada y usable; si no, limpia la clave."""
    try:
        win = System.AppDomain.CurrentDomain.GetData(_APPDOMAIN_WINDOW_KEY)
    except Exception:
        return None
    if win is None:
        return None
    try:
        _ = win.Title
        if hasattr(win, "IsLoaded") and (not win.IsLoaded):
            _clear_appdomain_window_key()
            return None
    except Exception:
        _clear_appdomain_window_key()
        return None
    return win


class EnfierradoFundacionAisladaWindow(object):
    def __init__(self, revit):
        self._revit = revit
        self._document = None
        self._foundation_ids = []
        self._entries = []
        self._is_closing_with_fade = False
        self._propagacion_pendiente_ui = False
        self._propagacion_ids_a_confirmar = None
        self._propagacion_num_val = None
        self._ids_run_override = None
        self._saltar_propagacion_post = False

        from System.Windows import RoutedEventHandler
        from System.Windows.Input import ApplicationCommands, CommandBinding, Key, KeyBinding, ModifierKeys
        from System.Windows.Markup import XamlReader

        self._win = XamlReader.Parse(_ENFIERRADO_FUND_XAML)
        self._form_width_px = float(_fundacion_aislada_form_width_px())
        self._win.Width = self._form_width_px
        self._win.MinWidth = self._form_width_px
        self._win.MaxWidth = self._form_width_px
        self._open_grow_storyboard_started = False

        self._seleccion_handler = SeleccionarFundacionesHandler(weakref.ref(self))
        self._seleccion_event = ExternalEvent.Create(self._seleccion_handler)
        self._colocar_handler = ColocarArmaduraFundacionStubHandler(weakref.ref(self))
        self._colocar_event = ExternalEvent.Create(self._colocar_handler)

        self._setup_ui(RoutedEventHandler)
        self._wire_commands(RoutedEventHandler, ApplicationCommands, CommandBinding, KeyBinding, Key, ModifierKeys)
        self._wire_lifecycle_handlers()
        self._wire_open_grow_storyboard_completed()

    def _wire_open_grow_storyboard_completed(self):
        try:
            from System import EventHandler

            sb = self._win.TryFindResource("FundOpenGrowStoryboard")
            if sb is not None:
                sb.Completed += EventHandler(self._on_open_grow_storyboard_completed)
        except Exception:
            pass

    def _on_open_grow_storyboard_completed(self, sender, args):
        try:
            self._win.MinWidth = self._form_width_px
            self._win.MaxWidth = self._form_width_px
        except Exception:
            pass

    def _position_win_top_left_active_view(self):
        try:
            uidoc = self._revit.ActiveUIDocument if self._revit else None
            hwnd = None
            if self._revit is not None:
                try:
                    hwnd = revit_main_hwnd(self._revit.Application)
                except Exception:
                    hwnd = None
            position_wpf_window_top_left_at_active_view(self._win, uidoc, hwnd)
        except Exception:
            pass

    def _begin_open_grow_storyboard(self):
        """Misma lógica que Wall Foundation Reinforcement: Scale 0→1 + opacidad vía Storyboard."""
        if self._open_grow_storyboard_started:
            return
        self._open_grow_storyboard_started = True
        try:
            from System import TimeSpan
            from System.Windows import Duration, SizeToContent
            from System.Windows.Media import ScaleTransform

            sc = self._win.FindName("FundRootScale")
            if sc is not None:
                sc.ScaleX = 0.0
                sc.ScaleY = 0.0
            self._win.Width = float(self._form_width_px)
            try:
                self._win.SizeToContent = SizeToContent.Height
            except Exception:
                pass
            self._position_win_top_left_active_view()
            sb = self._win.TryFindResource("FundOpenGrowStoryboard")
            if sb is None:
                if sc is not None:
                    sc.ScaleX = sc.ScaleY = 1.0
                self._win.Opacity = 1.0
                return
            dur = Duration(TimeSpan.FromMilliseconds(float(_WINDOW_CLOSE_MS)))
            try:
                for i in range(int(sb.Children.Count)):
                    sb.Children[i].Duration = dur
            except Exception:
                pass
            sb.Begin(self._win, True)
        except Exception:
            try:
                self._win.Opacity = 1.0
            except Exception:
                pass

    def _wire_lifecycle_handlers(self):
        try:
            from System.Windows import RoutedEventHandler

            def _on_closed(sender, args):
                _clear_appdomain_window_key()

            self._win.Closed += RoutedEventHandler(_on_closed)
        except Exception:
            pass

    def _setup_ui(self, RoutedEventHandler):
        from System.IO import FileAccess, FileMode, FileStream
        from System.Windows.Media.Imaging import BitmapCacheOption, BitmapImage

        try:
            img = self._win.FindName("ImgLogo")
            if img is not None:
                for logo_path in get_logo_paths():
                    if os.path.isfile(logo_path):
                        stream = None
                        try:
                            stream = FileStream(logo_path, FileMode.Open, FileAccess.Read)
                            bmp = BitmapImage()
                            bmp.BeginInit()
                            bmp.StreamSource = stream
                            bmp.CacheOption = BitmapCacheOption.OnLoad
                            bmp.EndInit()
                            bmp.Freeze()
                            img.Source = bmp
                        finally:
                            if stream is not None:
                                try:
                                    stream.Dispose()
                                except Exception:
                                    pass
                        break
        except Exception:
            pass

        btn_sel = self._win.FindName("BtnSeleccionar")
        if btn_sel is not None:
            btn_sel.Click += RoutedEventHandler(self._on_seleccionar)
        btn_close = self._win.FindName("BtnClose")
        if btn_close is not None:
            btn_close.Click += RoutedEventHandler(lambda s, e: self._close_with_fade())
        btn_col = self._win.FindName("BtnColocar")
        if btn_col is not None:
            btn_col.Click += RoutedEventHandler(self._on_colocar)

        try:
            from System.Windows.Input import MouseButtonEventHandler

            title_bar = self._win.FindName("TitleBar")
            if title_bar is not None:
                title_bar.MouseLeftButtonDown += MouseButtonEventHandler(
                    lambda s, e: self._win.DragMove()
                )
            if btn_close is not None:
                btn_close.MouseLeftButtonDown += MouseButtonEventHandler(lambda s, e: setattr(e, "Handled", True))
        except Exception:
            pass

        for chk_name, panel_name in (
            ("ChkInferior", "PanelInferior"),
            ("ChkSuperior", "PanelSuperior"),
            ("ChkLateral", "PanelLateral"),
        ):
            chk = self._win.FindName(chk_name)
            pnl = self._win.FindName(panel_name)
            if chk is None or pnl is None:
                continue

            def _make_toggle(panel):
                def _toggle(s, a):
                    try:
                        en = s.IsChecked == True
                        panel.IsEnabled = en
                        panel.Opacity = 1.0 if en else 0.35
                    except Exception:
                        pass

                return _toggle

            chk.Checked += RoutedEventHandler(_make_toggle(pnl))
            chk.Unchecked += RoutedEventHandler(_make_toggle(pnl))

        chk_sup = self._win.FindName("ChkSuperior")
        chk_lat = self._win.FindName("ChkLateral")
        pnl_lat = self._win.FindName("PanelLateral")
        if chk_sup is not None and chk_lat is not None:

            def _on_superior_armadura_changed(sender, args):
                if sender.IsChecked != True:
                    chk_lat.IsChecked = False
                    chk_lat.IsEnabled = False
                    if pnl_lat is not None:
                        pnl_lat.IsEnabled = False
                        pnl_lat.Opacity = 0.35
                else:
                    chk_lat.IsChecked = True
                    chk_lat.IsEnabled = True
                    if pnl_lat is not None:
                        pnl_lat.IsEnabled = True
                        pnl_lat.Opacity = 1.0

            chk_sup.Checked += RoutedEventHandler(_on_superior_armadura_changed)
            chk_sup.Unchecked += RoutedEventHandler(_on_superior_armadura_changed)

        def _bind_sep_handlers(prefix):
            bu = self._win.FindName("Btn{}SepUp".format(prefix))
            bd = self._win.FindName("Btn{}SepDown".format(prefix))
            tb = self._win.FindName("Txt{}SepMm".format(prefix))

            def on_up(s, a):
                self._step_sep_spinner(prefix, _SEP_MM_STEP)

            def on_dn(s, a):
                self._step_sep_spinner(prefix, -_SEP_MM_STEP)

            if bu is not None:
                bu.Click += RoutedEventHandler(on_up)
            if bd is not None:
                bd.Click += RoutedEventHandler(on_dn)
            if tb is not None:
                def on_lost_focus(s, a, tbx=tb, pr=prefix):
                    _normalize_sep_textbox(tbx)
                    self._sync_sep_spinner_enabled(pr)

                tb.LostFocus += RoutedEventHandler(on_lost_focus)

        for pfx in ("Inf", "Sup"):
            _bind_sep_handlers(pfx)

        bu_lat = self._win.FindName("BtnLatCantUp")
        bd_lat = self._win.FindName("BtnLatCantDown")
        tb_lat = self._win.FindName("TxtLatCant")

        def on_lat_up(s, a):
            self._step_lat_cant_spinner(_LAT_CANT_STEP)

        def on_lat_dn(s, a):
            self._step_lat_cant_spinner(-_LAT_CANT_STEP)

        if bu_lat is not None:
            bu_lat.Click += RoutedEventHandler(on_lat_up)
        if bd_lat is not None:
            bd_lat.Click += RoutedEventHandler(on_lat_dn)
        if tb_lat is not None:

            def on_lat_lost_focus(s, a, tbx=tb_lat):
                _normalize_lat_cant_textbox(tbx)
                self._sync_lat_cant_spinner_enabled()

            tb_lat.LostFocus += RoutedEventHandler(on_lat_lost_focus)

        self._win.Loaded += RoutedEventHandler(self._on_window_loaded)

    def _on_window_loaded(self, sender, args):
        try:
            from System import Action
            from System.Windows.Threading import DispatcherPriority

            self._win.Dispatcher.BeginInvoke(
                Action(self._begin_open_grow_storyboard),
                DispatcherPriority.Loaded,
            )
        except Exception:
            try:
                self._begin_open_grow_storyboard()
            except Exception:
                pass

    def _wire_commands(self, RoutedEventHandler, ApplicationCommands, CommandBinding, KeyBinding, Key, ModifierKeys):
        try:
            from System.Windows.Input import ExecutedRoutedEventHandler

            self._win.CommandBindings.Add(
                CommandBinding(
                    ApplicationCommands.Close,
                    ExecutedRoutedEventHandler(lambda s, e: self._close_with_fade()),
                )
            )
            self._win.InputBindings.Add(
                KeyBinding(ApplicationCommands.Close, Key.Escape, ModifierKeys.None)
            )
        except Exception:
            pass

    def _close_with_fade(self):
        if getattr(self, "_is_closing_with_fade", False):
            return
        self._is_closing_with_fade = True
        try:
            from System import TimeSpan, EventHandler
            from System.Windows import Duration
            from System.Windows.Media import ScaleTransform
            from System.Windows.Media.Animation import DoubleAnimation, QuadraticEase, EasingMode

            sc = self._win.FindName("FundRootScale")
            dur = Duration(TimeSpan.FromMilliseconds(float(_WINDOW_CLOSE_MS)))
            ease_in = QuadraticEase()
            ease_in.EasingMode = EasingMode.EaseIn

            def _da(f0, f1):
                a = DoubleAnimation()
                a.From = float(f0)
                a.To = float(f1)
                a.Duration = dur
                a.EasingFunction = ease_in
                return a

            try:
                sx0 = float(sc.ScaleX) if sc is not None else 1.0
                sy0 = float(sc.ScaleY) if sc is not None else 1.0
            except Exception:
                sx0 = sy0 = 1.0
            try:
                op0 = float(self._win.Opacity)
            except Exception:
                op0 = 1.0

            op_anim = _da(op0, 0.0)
            ax = _da(sx0, 0.0)
            ay = _da(sy0, 0.0)

            def _done(sender, args):
                try:
                    self._win.Close()
                except Exception:
                    pass

            op_anim.Completed += EventHandler(_done)
            if sc is not None:
                sc.BeginAnimation(ScaleTransform.ScaleXProperty, ax)
                sc.BeginAnimation(ScaleTransform.ScaleYProperty, ay)
            self._win.BeginAnimation(self._win.OpacityProperty, op_anim)
        except Exception:
            try:
                self._win.Close()
            except Exception:
                pass
            self._is_closing_with_fade = False

    def _show_with_fade(self):
        """Igual que Wall Foundation Reinforcement: opacidad 0, Show, Activate; Storyboard en Loaded."""
        try:
            self._win.Opacity = 0.0
            if not self._win.IsVisible:
                self._win.Show()
            self._win.Activate()
        except Exception:
            pass
        self._is_closing_with_fade = False

    def _show_after_pick(self):
        """Tras PickObject la ventana ya está cargada: mostrar sin repetir Storyboard (como Malla en Losa)."""
        try:
            self._win.Show()
            self._win.Activate()
        except Exception:
            pass

    def _set_estado(self, msg):
        try:
            txt = self._win.FindName("TxtEstado")
            if txt is not None:
                txt.Text = msg or u""
        except Exception:
            pass

    def _sync_window_size_to_content(self):
        """Fuerza remeasure con SizeToContent=Height (Visibility desde ExternalEvent a veces no lo actualiza)."""
        try:
            from System.Windows import SizeToContent

            w = self._win
            w.SizeToContent = SizeToContent.Manual
            w.SizeToContent = SizeToContent.Height
            w.UpdateLayout()
        except Exception:
            pass

    def _apply_sv_max_for_propagacion_footer(self, propagacion_visible):
        try:
            sv = self._win.FindName("SvContenido")
            if sv is None:
                return
            sv.MaxHeight = (
                float(_SV_CONTENIDO_MAX_H_PROPAGACION)
                if propagacion_visible
                else float(_SV_CONTENIDO_MAX_H_NORMAL)
            )
        except Exception:
            pass

    def _refit_window_after_footer_change(self, propagacion_visible):
        self._apply_sv_max_for_propagacion_footer(propagacion_visible)
        self._sync_window_size_to_content()
        try:
            from System.Windows.Threading import (
                DispatcherOperationCallback,
                DispatcherPriority,
            )

            w = self._win
            wr = weakref.ref(self)

            def _deferred(_op):
                win = wr()
                if win is not None:
                    try:
                        win._sync_window_size_to_content()
                    except Exception:
                        pass
                return None

            w.Dispatcher.BeginInvoke(
                DispatcherPriority.ApplicationIdle,
                DispatcherOperationCallback(_deferred),
                None,
            )
        except Exception:
            pass

    def _hide_propagacion_ui(self):
        try:
            from System.Windows import Visibility

            br = self._win.FindName("BorderPropagacion")
            if br is not None:
                br.Visibility = Visibility.Collapsed
        except Exception:
            pass
        self._propagacion_pendiente_ui = False
        self._propagacion_ids_a_confirmar = None
        self._propagacion_num_val = None
        try:
            btn = self._win.FindName("BtnColocar")
            if btn is not None:
                btn.Content = _CAPTION_BTN_COLOCAR_ARMADURAS
        except Exception:
            pass
        self._refit_window_after_footer_change(False)

    def _step_sep_spinner(self, prefix, delta):
        tb = self._win.FindName("Txt{}SepMm".format(prefix))
        if tb is None:
            return
        try:
            v = int(float(unicode(tb.Text).strip()))
        except Exception:
            v = int(_SEP_MM_DEFAULT_VAL)
        v += int(delta)
        v = max(_SEP_MM_MIN, min(_SEP_MM_MAX, v))
        nmax = int((_SEP_MM_MAX - _SEP_MM_MIN) // _SEP_MM_STEP)
        steps = int(round((v - _SEP_MM_MIN) / float(_SEP_MM_STEP)))
        steps = max(0, min(nmax, steps))
        v = _SEP_MM_MIN + steps * _SEP_MM_STEP
        tb.Text = unicode(int(v))
        self._sync_sep_spinner_enabled(prefix)

    def _refresh_laterales_cantidad_desde_ids(self, id_list):
        """Actualiza ``TxtLatCant`` según la **mayor** altura entre los ``ElementId`` dados."""
        doc = getattr(self, "_document", None)
        tb = self._win.FindName("TxtLatCant")
        if tb is None or doc is None or not id_list:
            return
        h_max = None
        for eid in id_list:
            el = doc.GetElement(eid)
            if el is None:
                continue
            h = _altura_fundacion_mm_para_cantidad_lateral(el)
            if h is not None:
                h_max = h if h_max is None else max(h_max, h)
        if h_max is None:
            return
        n = _cantidad_laterales_fundacion_desde_altura_mm(h_max)
        try:
            tb.Text = unicode(n)
        except Exception:
            pass
        self._sync_lat_cant_spinner_enabled()

    def _refresh_laterales_cantidad_desde_fundaciones(self):
        ids = getattr(self, "_foundation_ids", None) or []
        self._refresh_laterales_cantidad_desde_ids(ids)

    def _step_lat_cant_spinner(self, delta):
        tb = self._win.FindName("TxtLatCant")
        if tb is None:
            return
        try:
            v = int(float(unicode(tb.Text).strip()))
        except Exception:
            v = int(_LAT_CANT_DEFAULT_TXT)
        v += int(delta)
        v = max(_LAT_CANT_MIN, min(_LAT_CANT_MAX, v))
        tb.Text = unicode(int(v))
        self._sync_lat_cant_spinner_enabled()

    def _sync_lat_cant_spinner_enabled(self):
        tb = self._win.FindName("TxtLatCant")
        bu = self._win.FindName("BtnLatCantUp")
        bd = self._win.FindName("BtnLatCantDown")
        if tb is None:
            return
        try:
            v = int(float(unicode(tb.Text).strip()))
        except Exception:
            v = int(_LAT_CANT_DEFAULT_TXT)
        v = max(_LAT_CANT_MIN, min(_LAT_CANT_MAX, v))
        if bu is not None:
            bu.IsEnabled = v < _LAT_CANT_MAX
        if bd is not None:
            bd.IsEnabled = v > _LAT_CANT_MIN

    def _sync_sep_spinner_enabled(self, prefix):
        tb = self._win.FindName("Txt{}SepMm".format(prefix))
        bu = self._win.FindName("Btn{}SepUp".format(prefix))
        bd = self._win.FindName("Btn{}SepDown".format(prefix))
        if tb is None:
            return
        try:
            v = int(float(unicode(tb.Text).strip()))
        except Exception:
            v = int(_SEP_MM_DEFAULT_VAL)
        v = max(_SEP_MM_MIN, min(_SEP_MM_MAX, v))
        if bu is not None:
            bu.IsEnabled = v < _SEP_MM_MAX
        if bd is not None:
            bd.IsEnabled = v > _SEP_MM_MIN

    def _init_sep_spinners(self):
        for pfx in ("Inf", "Sup"):
            tb = self._win.FindName("Txt{}SepMm".format(pfx))
            if tb is not None:
                try:
                    tb.Text = unicode(int(_SEP_MM_DEFAULT_VAL))
                except Exception:
                    pass
                _normalize_sep_textbox(tb)
            self._sync_sep_spinner_enabled(pfx)
        tb_lat = self._win.FindName("TxtLatCant")
        if tb_lat is not None:
            try:
                tb_lat.Text = unicode(int(_LAT_CANT_DEFAULT_TXT))
            except Exception:
                pass
            _normalize_lat_cant_textbox(tb_lat)
        self._sync_lat_cant_spinner_enabled()

    def _cargar_combos_diametro(self):
        doc = self._document
        if doc is None:
            return
        entries, err = _build_bar_type_entries(doc)
        self._entries = list(entries) if entries else []
        for name in ("CmbInfDiam", "CmbSupDiam", "CmbLatDiam"):
            cmb = self._win.FindName(name)
            if cmb is None:
                continue
            cmb.Items.Clear()
            cmb.IsEditable = False
            if err:
                self._set_estado(err)
                continue
            for _bt, lbl in self._entries:
                cmb.Items.Add(lbl)
            sel_idx = 0
            for i, (b, lbl) in enumerate(self._entries):
                dmm = None
                try:
                    if b is not None:
                        dmm = _rebar_nominal_diameter_mm(b)
                except Exception:
                    dmm = None
                if dmm == 8:
                    sel_idx = i
                    break
                if b is None and u"8" in unicode(lbl):
                    sel_idx = i
                    break
            try:
                cmb.SelectedIndex = min(sel_idx, max(0, cmb.Items.Count - 1))
            except Exception:
                cmb.SelectedIndex = 0

    def _on_seleccionar(self, sender, args):
        self._seleccion_event.Raise()

    def _on_colocar(self, sender, args):
        if getattr(self, "_propagacion_pendiente_ui", False):
            ids_extra = getattr(self, "_propagacion_ids_a_confirmar", None) or []
            self._hide_propagacion_ui()
            if not ids_extra:
                self._set_estado(u"No hay fundaciones para propagar.")
                return
            self._ids_run_override = list(ids_extra)
            self._saltar_propagacion_post = True
            self._colocar_event.Raise()
            self._set_estado(u"En cola: propagar armadura…")
            return
        if not self._foundation_ids:
            _task_dialog_show(
                u"BIMTools — Armadura Fundacion Aislada",
                u"Seleccione una fundación en el modelo.",
                self._win,
            )
            self._set_estado(u"")
            return
        chk_inf = self._win.FindName("ChkInferior")
        chk_sup = self._win.FindName("ChkSuperior")
        chk_lat = self._win.FindName("ChkLateral")
        if chk_inf and chk_sup and chk_lat:
            if (
                chk_inf.IsChecked != True
                and chk_sup.IsChecked != True
                and chk_lat.IsChecked != True
            ):
                _task_dialog_show(
                    u"BIMTools — Armadura Fundacion Aislada",
                    u"Active al menos un grupo de armadura (inferior, superior o lateral).",
                    self._win,
                )
                return
        self._colocar_event.Raise()
        self._set_estado(u"En cola: colocar armadura y detalle…")

    def show(self):
        uidoc = self._revit.ActiveUIDocument
        if uidoc is None:
            _task_dialog_show(
                u"Armadura Fundacion Aislada",
                u"No hay documento activo.",
                self._win,
            )
            return
        self._document = uidoc.Document
        hwnd = None
        try:
            hwnd = revit_main_hwnd(self._revit.Application)
        except Exception:
            pass
        try:
            from System.Windows.Interop import WindowInteropHelper

            if hwnd:
                helper = WindowInteropHelper(self._win)
                helper.Owner = hwnd
        except Exception:
            pass
        position_wpf_window_top_left_at_active_view(self._win, uidoc, hwnd)
        self._init_sep_spinners()
        if getattr(self, "_foundation_ids", None):
            try:
                self._refresh_laterales_cantidad_desde_fundaciones()
            except Exception:
                pass
        self._cargar_combos_diametro()
        self._set_estado(u"")
        self._show_with_fade()
        try:
            System.AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, self._win)
        except Exception:
            pass


def run_pyrevit(revit):
    if _scripts_dir not in sys.path:
        sys.path.insert(0, _scripts_dir)

    existing = _get_active_window()
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
            _clear_appdomain_window_key()
            existing = None
        if ok and existing is not None:
            _task_dialog_show(
                u"BIMTools — Armadura Fundacion Aislada",
                u"La herramienta ya está en ejecución.",
                existing,
            )
            return

    w = EnfierradoFundacionAisladaWindow(revit)
    try:
        w.show()
    except Exception:
        _clear_appdomain_window_key()
        raise
