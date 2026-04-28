# -*- coding: utf-8 -*-
"""
Wall Foundation Reinforcement — zapata de ``WallFoundation`` (Revit 2024+).

IMPORTANTE — Exclusión crítica (troceo):
    Las curvas a las que se aplicó un **estiramiento por empotramiento** deben
    **excluirse** de cualquier troceo automático: partir esos tramos en la API
    puede producir geometría inválida o armadura incoherente.
    Esta herramienta **solo** trocea el eje **recto** inferido del elemento
    (``LocationCurve``); no importa ``ModelLine`` ni curvas editadas manualmente
    con criterios de empotramiento.

- Conjuntos Rebar (``SetLayoutAsMaximumSpacing``) como en fundación aislada. Si la herramienta se
  ejecuta con una vista en planta activa, cada conjunto nuevo pasa a presentación **Middle** (solo
  barra central) en esa vista.
- Transversales: polilínea en U (``RebarShape`` «03» / respaldos en
  ``rebar_fundacion_cara_inferior``) adaptada al ancho y peralte de la zapata.
- Longitudinales: polilínea con patas (``CreateFromCurves*`` / forma «03» o lazo), igual criterio
  que la U; si el eje supera 12 m, troceo con traslape según tabla ø; cada tramo de eje respeta
  ``largo_máximo − pata`` en primera y última barra (pata = tabla por ø), intermedias hasta ``largo_máximo``.

**Geometría / uniones:** al colocar armadura se memorizan los elementos unidos
con ``JoinGeometryUtils.GetJoinedElements``, se ejecuta ``UnjoinGeometry`` sobre
cada par, se regenera, y se obtienen las curvas desde la **cara inferior** con
la misma cadena que fundación aislada (``extraer_curva_lado_mayor/menor``,
``aplicar_recubrimiento_inferior_completo_mm``,
``offset_linea_eje_barra_desde_cara_inferior_mm``, ``evaluar_caras_paralelas``…).
Tras crear las barras se restaura ``JoinGeometry`` con los mismos elementos
(dentro de la misma transacción).

Unidades internas: conversión con ``UnitUtils`` y ``UnitTypeId`` (``ForgeTypeId``).

Tras colocar: una ``IndependentTag`` por conjunto; el **tipo** dentro de la familía
``EST_A_STRUCTURAL REBAR TAG`` debe coincidir con el **nombre del RebarShape** modelado (p. ej. «03»).
En vista ortogonal se crea además una **Multi-Rebar Annotation** por conjunto con el tipo
«Recorrido Barras» (familia Multi-Rebar Annotations), si está cargado en el proyecto.
Con troceo longitudinal (> 12 m), en la **vista activa** se coloca el mismo **Detail Component**
de empalme que vigas / borde losa (``EST_D_DEATIL ITEM_EMPALME`` / tipo ``Empalme``), alineado al
eje de stock entre tramos consecutivos, y una **cota lineal** del traslape entre sus referencias
Left/Right (misma lógica que ``enfierrado_shaft_hashtag`` / vigas). Ese detalle y la cota se
vinculan con ``lap_detail_link_wall_foundation_schema`` para el **DMU**: solo depuración si
falta una barra; **no** se recoloca el tramo del símbolo (evita saltos a la primera barra del
layout).
Tras colocar armadura se genera una **vista en sección transversal** al ``LocationCurve`` (punto
medio), como en enfierrado de vigas (módulo ``vista_seccion_enfierrado_vigas``); al cerrar el
formulario se eliminan esas vistas de revisión.
"""

from __future__ import print_function

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

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FamilySymbol,
    FilteredElementCollector,
    IndependentTag,
    JoinGeometryUtils,
    Line,
    Options,
    Reference,
    StorageType,
    TagMode,
    TagOrientation,
    Transaction,
    UnitTypeId,
    UnitUtils,
    View3D,
    ViewDetailLevel,
    ViewPlan,
    WallFoundation,
    XYZ,
)
from Autodesk.Revit.DB.Structure import Rebar, RebarPresentationMode
from Autodesk.Revit.DB import LocationCurve
from Autodesk.Revit.UI import TaskDialog, ExternalEvent, IExternalEventHandler
from Autodesk.Revit.UI.Selection import ISelectionFilter

from barras_bordes_losa_gancho_empotramiento import (
    _build_bar_type_entries,
    _find_fixed_lap_detail_symbol_id,
    _rebar_nominal_diameter_mm,
    _task_dialog_show,
    element_id_to_int,
)
from geometria_viga_cara_superior_detalle import (
    _colocar_detail_item_traslape_en_vista,
    vista_permite_detail_curve,
)
from geometria_wall_foundation_cortes_muro import (
    geometria_inferior_wall_foundation_cortes_muro,
    vector_transversal_planta_desde_muro_host,
)
from geometria_fundacion_cara_inferior import (
    aplicar_recubrimiento_inferior_completo_mm,
    centro_xy_perimetro_inferior_doc,
    construir_polilinea_fundacion_ganchos_geometricos_desde_eje,
    construir_polilinea_u_fundacion_desde_eje_horizontal,
    evaluar_caras_paralelas_curva_mas_cercana,
    extraer_curva_lado_mayor_cara_inferior,
    extraer_curva_lado_menor_cara_inferior,
    largo_gancho_u_tabla_mm,
    longitud_pata_u_fundacion_inf_sup_ft,
    luz_proyeccion_perimetro_inferior_ft,
    obtener_marco_coordenadas_cara_inferior,
    offset_linea_eje_barra_desde_cara_inferior_mm,
    span_bruto_proyeccion_perimetro_inferior_ft,
    vector_reverso_cara_paralela_mas_cercana_a_barra,
)
from bimtools_rebar_hook_lengths import (
    pata_eje_curve_loop_mm_desde_tabla_mm,
    traslape_mm_from_nominal_diameter_mm,
)
from rebar_fundacion_cara_inferior import (
    REBAR_SHAPE_NOMBRE_DEFECTO,
    aplicar_layout_maximum_spacing_rebar,
    crear_rebar_polilinea_recta_sin_ganchos,
    crear_rebar_polilinea_u_malla_inf_sup_curve_loop,
    crear_rebar_u_shape_desde_eje_rebar_shape_nombrado,
    rebar_shape_display_name,
)
from revit_wpf_window_position import (
    position_wpf_window_top_left_at_active_view,
    revit_main_hwnd,
)

from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_paths import get_logo_paths

_APPDOMAIN_WINDOW_KEY = "BIMTools.WallFoundationReinforcement.ActiveWindow"

_FOUNDATION_CAT_ID = int(BuiltInCategory.OST_StructuralFoundation)

_SEP_MM_MIN = 100
_SEP_MM_MAX = 400
_SEP_MM_STEP = 10
_SEP_MM_DEFAULT = 100

_DOSIFICACION_HORMIGON_OPCIONES = (u"G25", u"G35", u"G45")
_DOSIFICACION_HORMIGON_DEFAULT = u"G25"

_RECO_HOR_MM = 50.0
_RECO_EXT_EJE_MM = 50.0
_DESC_PATA_U_MM = 150.0
# Offset en planta (mm) del perímetro inferior — misma base que fundación aislada / WF común.
_REC_OFF_PLANTA_INF_MM = 100.0
# Recorte en extremos del **lado mayor** (eje long.): cara de hormigón → **tangente** de la barra.
# Distancia a lo largo del eje hasta el **eje** de la barra = este valor + ø_long/2.
_REC_EXTREMOS_LONG_TANGENTE_MM = 50.0
# Distancia (mm) de la cara lateral de hormigón a la **fibra tangente** exterior del tramo/hook
# de la U (no al eje de la barra). Ajuste solo en ``_colocar_trans_u``.
_REC_LATERAL_CARA_U_MM = 50.0
_REC_EXTREMOS_INFERIOR_MM = 50.0

_MAX_STOCK_MM = 12000.0
_LAP_MM_MIN = 100.0
_LAP_MM_MAX = 4000.0
_LAP_DEFAULT_MM = 860.0

_MAX_BAR_USER_MIN_MM = 1000.0
_MAX_BAR_USER_MAX_MM = 12000.0

# Etiqueta de armadura por conjunto: familia de etiquetas; el **tipo** usado = nombre del
# ``RebarShape`` de la barra (p. ej. «03»), si existe en esa familia.
_WF_REBAR_TAG_FAMILY_NAME = u"EST_A_STRUCTURAL REBAR TAG"
# Tipo de Multi-Rebar Annotation (nombre en el selector de tipos de Revit).
_WF_MULTI_REBAR_ANNOTATION_TYPE_NAME = u"Recorrido Barras"


def _wf_norm_nombre_familia_etiqueta(s):
    if s is None:
        return u""
    try:
        t = unicode(s)
    except Exception:
        try:
            t = System.Convert.ToString(s)
        except Exception:
            t = u""
    return u" ".join(t.replace(u"\u00A0", u" ").split()).lower()


def _wf_primer_family_symbol_rebar_tag_por_nombre_familia(document, family_name):
    """
    Primer ``FamilySymbol`` de categoría etiqueta de armadura cuyo ``FamilyName`` coincide
    con ``family_name`` (sin distinguir mayúsculas / espacios).
    """
    if document is None or not family_name:
        return None
    tgt = _wf_norm_nombre_familia_etiqueta(family_name)
    if not tgt:
        return None
    try:
        col = (
            FilteredElementCollector(document)
            .OfClass(FamilySymbol)
            .OfCategory(BuiltInCategory.OST_RebarTags)
        )
        candidatos = []
        for sym in col:
            if sym is None:
                continue
            fn = u""
            try:
                fn = sym.FamilyName
            except Exception:
                pass
            if not fn:
                try:
                    fam = sym.Family
                    if fam is not None:
                        fn = fam.Name
                except Exception:
                    fn = u""
            if _wf_norm_nombre_familia_etiqueta(fn) != tgt:
                continue
            candidatos.append(sym)
    except Exception:
        return None
    if not candidatos:
        return None
    try:
        candidatos.sort(key=lambda x: (_wf_norm_nombre_familia_etiqueta(getattr(x, "Name", u""))))
    except Exception:
        pass
    sym0 = candidatos[0]
    try:
        if sym0 is not None and not sym0.IsActive:
            sym0.Activate()
    except Exception:
        pass
    return sym0


def _wf_nombres_tipo_family_symbol(sym):
    """Cadenas normalizadas comparables con el nombre de un ``RebarShape``."""
    out = []
    seen = set()
    if sym is None:
        return out
    try:
        n = getattr(sym, "Name", None)
        if n:
            c = _wf_norm_nombre_familia_etiqueta(n)
            if c and c not in seen:
                seen.add(c)
                out.append(c)
    except Exception:
        pass
    for bip_name in (u"SYMBOL_NAME_PARAM", u"ALL_MODEL_TYPE_NAME"):
        try:
            bip = getattr(BuiltInParameter, bip_name, None)
            if bip is None:
                continue
            p = sym.get_Parameter(bip)
            if p is None or not p.HasValue or p.StorageType != StorageType.String:
                continue
            c = _wf_norm_nombre_familia_etiqueta(p.AsString())
            if c and c not in seen:
                seen.add(c)
                out.append(c)
        except Exception:
            continue
    return out


def _wf_family_symbols_rebar_tag_en_familia(document, family_name):
    """Todos los ``FamilySymbol`` OST_RebarTags de la familia ``family_name``."""
    if document is None or not family_name:
        return []
    tgt = _wf_norm_nombre_familia_etiqueta(family_name)
    if not tgt:
        return []
    out = []
    try:
        col = (
            FilteredElementCollector(document)
            .OfClass(FamilySymbol)
            .OfCategory(BuiltInCategory.OST_RebarTags)
        )
        for sym in col:
            if sym is None:
                continue
            fn = u""
            try:
                fn = sym.FamilyName
            except Exception:
                pass
            if not fn:
                try:
                    fam = sym.Family
                    if fam is not None:
                        fn = fam.Name
                except Exception:
                    fn = u""
            if _wf_norm_nombre_familia_etiqueta(fn) != tgt:
                continue
            out.append(sym)
    except Exception:
        return []
    return out


def _wf_family_symbol_rebar_tag_por_nombre_shape(document, family_name, shape_name):
    """
    ``FamilySymbol`` en ``family_name`` cuyo nombre de tipo coincide con ``shape_name``
    (mismo criterio de normalización que el nombre visible del ``RebarShape``).
    """
    if document is None or not shape_name:
        return None
    key = _wf_norm_nombre_familia_etiqueta(shape_name)
    if not key:
        return None
    syms = _wf_family_symbols_rebar_tag_en_familia(document, family_name)
    exact = []
    for sym in syms:
        for cand in _wf_nombres_tipo_family_symbol(sym):
            if cand == key:
                exact.append(sym)
                break
    if len(exact) == 1:
        s0 = exact[0]
    elif len(exact) > 1:
        try:
            exact.sort(
                key=lambda x: _wf_norm_nombre_familia_etiqueta(getattr(x, "Name", u""))
            )
        except Exception:
            pass
        s0 = exact[0]
    else:
        s0 = None
    if s0 is not None:
        try:
            if not s0.IsActive:
                s0.Activate()
        except Exception:
            pass
    return s0


def _wf_rebar_shape_nombre_desde_barra(document, rebar):
    """Nombre visible del ``RebarShape`` asignado a la instancia ``rebar``."""
    if document is None or rebar is None:
        return u""
    sid = None
    try:
        sid = rebar.GetShapeId()
    except Exception:
        sid = None
    if sid is None or sid == ElementId.InvalidElementId:
        try:
            sid = rebar.RebarShapeId
        except Exception:
            sid = None
    if sid is None or sid == ElementId.InvalidElementId:
        return u""
    try:
        sh = document.GetElement(sid)
    except Exception:
        sh = None
    return rebar_shape_display_name(sh)


def _wf_punto_insercion_tag_rebar(rebar, view):
    if rebar is None:
        return None
    try:
        bb = rebar.get_BoundingBox(view)
        if bb is not None:
            return (bb.Min + bb.Max) * 0.5
    except Exception:
        pass
    try:
        bb0 = rebar.get_BoundingBox(None)
        if bb0 is not None:
            return (bb0.Min + bb0.Max) * 0.5
    except Exception:
        pass
    return None


def _wf_proyectar_punto_plano_vista(p, view):
    """Proyecta ``p`` al plano de la vista (corte en planta/alzado) para cabecera de etiqueta."""
    if p is None or view is None:
        return p
    try:
        vd = view.ViewDirection
        if vd is None or float(vd.GetLength()) < 1e-12:
            return p
        vd = vd.Normalize()
        vo = view.Origin
        if vo is None:
            return p
        d = float((p - vo).DotProduct(vd))
        return p - vd.Multiply(d)
    except Exception:
        return p


def _wf_referencias_tag_rebar(document, rebar, view):
    """
    Referencias para ``IndependentTag.Create`` (barra completa, posiciones del conjunto,
    subelementos y curvas de la geometría con ``ComputeReferences``).
    """
    refs = []
    seen = set()

    def _add_ref(r):
        if r is None:
            return
        try:
            k = r.ConvertToStableRepresentation(document)
        except Exception:
            try:
                k = unicode(r)
            except Exception:
                k = id(r)
        if k in seen:
            return
        seen.add(k)
        refs.append(r)

    try:
        subs = rebar.GetSubelements() if hasattr(rebar, "GetSubelements") else None
    except Exception:
        subs = None
    if subs:
        for sub in subs:
            if sub is None:
                continue
            try:
                if hasattr(sub, "GetReference"):
                    _add_ref(sub.GetReference())
            except Exception:
                continue

    try:
        npos = int(getattr(rebar, "NumberOfBarPositions", 0))
    except Exception:
        try:
            npos = (
                int(rebar.GetNumberOfBarPositions())
                if hasattr(rebar, "GetNumberOfBarPositions")
                else 0
            )
        except Exception:
            npos = 0
    if npos > 0:
        idxs = [0, int(npos / 2), max(0, npos - 1)]
        for idx in idxs:
            try:
                if hasattr(rebar, "GetReferenceToBarPosition"):
                    _add_ref(rebar.GetReferenceToBarPosition(idx))
                elif hasattr(rebar, "GetReferenceForBarPosition"):
                    _add_ref(rebar.GetReferenceForBarPosition(idx))
            except Exception:
                continue
    try:
        _add_ref(Reference(rebar))
    except Exception:
        pass

    def _collect_geom_refs(geom_elem):
        if geom_elem is None:
            return
        for go in geom_elem:
            if go is None:
                continue
            try:
                rgo = getattr(go, "Reference", None)
                if rgo is not None:
                    _add_ref(rgo)
            except Exception:
                pass
            try:
                gi = (
                    go.GetInstanceGeometry()
                    if hasattr(go, "GetInstanceGeometry")
                    else None
                )
                if gi is not None:
                    _collect_geom_refs(gi)
            except Exception:
                pass

    for use_view, incl_nv in ((True, False), (False, True), (False, False)):
        try:
            opts = Options()
            opts.ComputeReferences = True
            opts.IncludeNonVisibleObjects = incl_nv
            try:
                opts.DetailLevel = ViewDetailLevel.Fine
            except Exception:
                pass
            if use_view and view is not None:
                try:
                    opts.View = view
                except Exception:
                    pass
            geo = rebar.get_Geometry(opts)
            if geo is not None:
                _collect_geom_refs(geo)
        except Exception:
            continue
        if refs:
            break
    return refs


def _wf_vista_permite_independent_tag(view):
    if view is None:
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass
    try:
        if isinstance(view, View3D):
            return False
    except Exception:
        pass
    return True


def _wf_etiquetar_rebar_sets_independent_tag(document, view, rebars, avisos):
    """
    Una ``IndependentTag`` por cada ``Rebar`` (cada set): tipo de etiqueta = nombre del
    ``RebarShape`` de esa barra dentro de la familia ``_WF_REBAR_TAG_FAMILY_NAME``.
    """
    if document is None or view is None or not rebars or avisos is None:
        return 0
    if not _wf_vista_permite_independent_tag(view):
        avisos.append(
            u"Etiqueta «{0}»: use planta/alzado/sección (no plantilla ni 3D).".format(
                _WF_REBAR_TAG_FAMILY_NAME
            )
        )
        return 0
    if not _wf_family_symbols_rebar_tag_en_familia(document, _WF_REBAR_TAG_FAMILY_NAME):
        try:
            avisos.append(
                u"Etiqueta: no hay símbolos OST_RebarTags para familia «{0}».".format(
                    _WF_REBAR_TAG_FAMILY_NAME
                )
            )
        except Exception:
            pass
        return 0
    n_ok = 0
    for rb in rebars:
        if rb is None or not isinstance(rb, Rebar):
            continue
        shape_nm = _wf_rebar_shape_nombre_desde_barra(document, rb)
        tag_symbol = _wf_family_symbol_rebar_tag_por_nombre_shape(
            document, _WF_REBAR_TAG_FAMILY_NAME, shape_nm
        )
        if tag_symbol is None:
            try:
                rid = element_id_to_int(rb.Id)
            except Exception:
                rid = u"?"
            avisos.append(
                u"Etiqueta Id rebar {0}: no hay tipo «{1}» en «{2}» (nombre de RebarShape).".format(
                    rid,
                    shape_nm or u"?",
                    _WF_REBAR_TAG_FAMILY_NAME,
                )
            )
            continue
        try:
            tid = tag_symbol.Id
        except Exception:
            continue
        try:
            if tid is None or tid == ElementId.InvalidElementId:
                continue
        except Exception:
            pass
        p_raw = _wf_punto_insercion_tag_rebar(rb, view)
        p = _wf_proyectar_punto_plano_vista(p_raw, view)
        if p is None:
            try:
                rid = element_id_to_int(rb.Id)
            except Exception:
                rid = u"?"
            avisos.append(
                u"Etiqueta Id rebar {0}: sin punto de inserción.".format(rid)
            )
            continue
        refs = _wf_referencias_tag_rebar(document, rb, view)
        if not refs:
            try:
                rid = element_id_to_int(rb.Id)
            except Exception:
                rid = u"?"
            avisos.append(
                u"Etiqueta Id rebar {0}: sin referencia API.".format(rid)
            )
            continue
        created = None
        last_ex_msg = None
        for ref in refs:
            for orient in (TagOrientation.Horizontal, TagOrientation.Vertical):
                for add_leader in (False, True):
                    try:
                        created = IndependentTag.Create(
                            document,
                            tid,
                            view.Id,
                            ref,
                            add_leader,
                            orient,
                            p,
                        )
                    except Exception as _ex_tag:
                        created = None
                        try:
                            last_ex_msg = unicode(_ex_tag)
                        except Exception:
                            last_ex_msg = None
                    if created is not None:
                        break
                if created is not None:
                    break
            if created is not None:
                break
        if created is None:
            try:
                for ref in refs:
                    for orient in (TagOrientation.Horizontal, TagOrientation.Vertical):
                        for add_leader in (False, True):
                            try:
                                created = IndependentTag.Create(
                                    document,
                                    view.Id,
                                    ref,
                                    add_leader,
                                    TagMode.TM_ADDBY_CATEGORY,
                                    orient,
                                    p,
                                )
                                if created is not None:
                                    try:
                                        created.SetTypeId(tid)
                                    except Exception:
                                        pass
                            except Exception as _ex_tag2:
                                created = None
                                try:
                                    last_ex_msg = unicode(_ex_tag2)
                                except Exception:
                                    pass
                            if created is not None:
                                break
                        if created is not None:
                            break
                    if created is not None:
                        break
            except Exception:
                created = None
        if created is not None:
            try:
                if created.HasLeader:
                    created.HasLeader = False
            except Exception:
                pass
            n_ok += 1
        else:
            try:
                rid = element_id_to_int(rb.Id)
            except Exception:
                rid = u"?"
            _msg = u"Etiqueta Id rebar {0}: no se pudo crear con «{1}».".format(
                rid, _WF_REBAR_TAG_FAMILY_NAME
            )
            if last_ex_msg:
                try:
                    _msg += u" ({0})".format(last_ex_msg[:220])
                except Exception:
                    pass
            avisos.append(_msg)
    return int(n_ok)


def _wf_vista_es_planta(view):
    """``True`` si la vista activa es una planta (``ViewPlan``: planta de planta / estructura, etc.)."""
    if view is None:
        return False
    try:
        return isinstance(view, ViewPlan)
    except Exception:
        return False


def _wf_rebar_presentacion_solo_centro_en_vista(view, rebar_elem):
    """
    En la vista dada, presentación **Middle** del conjunto (equivalente a *Middle* en la UI de Revit).
    Solo aplica si ``CanApplyPresentationMode`` lo admite (p. ej. no barra única sin conjunto).
    """
    if view is None or rebar_elem is None:
        return
    try:
        if not isinstance(rebar_elem, Rebar):
            return
        if not rebar_elem.CanApplyPresentationMode(view):
            return
        rebar_elem.SetPresentationMode(view, RebarPresentationMode.Middle)
    except Exception:
        pass


def _wf_aplicar_presentacion_solo_barra_central_planta(view, rebars):
    if view is None or not rebars:
        return
    for rb in rebars:
        _wf_rebar_presentacion_solo_centro_en_vista(view, rb)

_WINDOW_OPEN_MS = 180
_WINDOW_CLOSE_MS = 180
_SV_MAX_H = 620.0

_FUND_INPUT_COLS_PER_ROW = 2
_FUND_COMBO_WIDTH_PX = 110
_FUND_DIAM_ESP_AT_COL_PX = 28
_FUND_BLOCK_PAD_H_PX = 16
_FUND_GROUPBOX_PAD_H_PX = 16
_FUND_OUTER_PAD_H_PX = 28
_FUND_WIDTH_TITLE_MIN_PX = 288


def _mm_to_ft(mm):
    try:
        return float(
            UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)
        )
    except Exception:
        return float(mm) / 304.8


def _ft_to_mm(ft):
    try:
        return float(
            UnitUtils.ConvertFromInternalUnits(float(ft), UnitTypeId.Millimeters)
        )
    except Exception:
        return float(ft) * 304.8


def _fund_form_width_px():
    cols = max(1, int(_FUND_INPUT_COLS_PER_ROW))
    c = int(_FUND_COMBO_WIDTH_PX)
    row_inner = cols * c + _FUND_DIAM_ESP_AT_COL_PX + _FUND_BLOCK_PAD_H_PX
    w = row_inner + _FUND_GROUPBOX_PAD_H_PX + _FUND_OUTER_PAD_H_PX
    w = max(w, _FUND_WIDTH_TITLE_MIN_PX)
    return int((int(w) + 3) // 4 * 4)


def _parse_diameter_mm_from_bar_combo_label(lbl):
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


def _snap_sep_mm(raw_n, default_val=_SEP_MM_DEFAULT):
    try:
        n = int(round(float(raw_n)))
    except Exception:
        return int(default_val)
    n = max(_SEP_MM_MIN, min(_SEP_MM_MAX, n))
    nmax = int((_SEP_MM_MAX - _SEP_MM_MIN) // _SEP_MM_STEP)
    steps = int(round((n - _SEP_MM_MIN) / float(_SEP_MM_STEP)))
    steps = max(0, min(nmax, steps))
    return _SEP_MM_MIN + steps * _SEP_MM_STEP


def _normalize_sep_tb(tb, default_val=_SEP_MM_DEFAULT):
    if tb is None:
        return
    try:
        s = unicode(tb.Text).replace(u"mm", u"").strip()
        if not s:
            tb.Text = unicode(int(default_val))
            return
        n = int(round(float(s.replace(u",", u"."))))
    except Exception:
        tb.Text = unicode(int(default_val))
        return
    v = _snap_sep_mm(n, default_val)
    tb.Text = unicode(int(v))


def _read_sep_tb(tb, default_val=_SEP_MM_DEFAULT):
    if tb is None:
        return int(default_val)
    try:
        s = unicode(tb.Text).replace(u"mm", u"").strip()
        if not s:
            return int(default_val)
        n = int(round(float(s.replace(u",", u"."))))
    except Exception:
        return int(default_val)
    return int(_snap_sep_mm(n, default_val))


def _read_dosificacion_hormigon(combo):
    if combo is None:
        return _DOSIFICACION_HORMIGON_DEFAULT
    try:
        si = combo.SelectedItem
        if si is not None:
            s = unicode(si).strip()
            if s in _DOSIFICACION_HORMIGON_OPCIONES:
                return s
    except Exception:
        pass
    try:
        s = unicode(combo.Text).strip().upper()
        for opt in _DOSIFICACION_HORMIGON_OPCIONES:
            if s == opt.upper():
                return opt
    except Exception:
        pass
    return _DOSIFICACION_HORMIGON_DEFAULT


def _normalize_lap_tb(tb):
    if tb is None:
        return
    try:
        s = unicode(tb.Text).strip()
        if not s:
            tb.Text = unicode(int(_LAP_DEFAULT_MM))
            return
        n = int(round(float(s.replace(u",", u"."))))
    except Exception:
        tb.Text = unicode(int(_LAP_DEFAULT_MM))
        return
    n = max(int(_LAP_MM_MIN), min(int(_LAP_MM_MAX), n))
    tb.Text = unicode(int(n))


def _read_lap_tb(tb):
    if tb is None:
        return float(_LAP_DEFAULT_MM)
    try:
        s = unicode(tb.Text).strip()
        if not s:
            return float(_LAP_DEFAULT_MM)
        n = float(s.replace(u",", u"."))
    except Exception:
        return float(_LAP_DEFAULT_MM)
    n = max(_LAP_MM_MIN, min(_LAP_MM_MAX, n))
    return n


def _wf_traslape_mm_longitudinal(d_long_mm, tlap, concrete_grade=None):
    """Traslape (mm) según tabla por ø longitudinal y dosificación; respaldo ``TxtLapMm``."""
    try:
        if d_long_mm is not None and float(d_long_mm) > 1e-6:
            v = traslape_mm_from_nominal_diameter_mm(
                float(d_long_mm), concrete_grade
            )
            if v is not None:
                return float(v)
    except Exception:
        pass
    return _read_lap_tb(tlap)


def _normalize_max_bar_tb(tb):
    if tb is None:
        return
    try:
        s = unicode(tb.Text).strip()
        if not s:
            tb.Text = unicode(int(_MAX_STOCK_MM))
            return
        n = int(round(float(s.replace(u",", u"."))))
    except Exception:
        tb.Text = unicode(int(_MAX_STOCK_MM))
        return
    n = max(int(_MAX_BAR_USER_MIN_MM), min(int(_MAX_BAR_USER_MAX_MM), n))
    tb.Text = unicode(int(n))


def _read_max_bar_tb(tb):
    if tb is None:
        return float(_MAX_STOCK_MM)
    try:
        s = unicode(tb.Text).strip()
        if not s:
            return float(_MAX_STOCK_MM)
        n = float(s.replace(u",", u"."))
    except Exception:
        return float(_MAX_STOCK_MM)
    n = max(_MAX_BAR_USER_MIN_MM, min(_MAX_BAR_USER_MAX_MM, n))
    return n


def _clear_appdomain_window_key():
    try:
        System.AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, None)
    except Exception:
        pass


def _get_active_window():
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


def _wf_width_ft(wf):
    if wf is None:
        return None
    for bip in (
        getattr(BuiltInParameter, "CONTINUOUS_FOOTING_WIDTH", None),
        getattr(BuiltInParameter, "STRUCTURAL_FOUNDATION_WIDTH", None),
    ):
        if bip is None:
            continue
        try:
            p = wf.get_Parameter(bip)
            if p is not None and p.HasValue:
                if p.StorageType == StorageType.String:
                    continue
                v = p.AsDouble()
                if v and v > 1e-9:
                    return float(v)
        except Exception:
            continue
    try:
        for nm in (u"Width", u"Ancho", u"Anchura"):
            lp = wf.LookupParameter(nm)
            if lp is not None and lp.HasValue:
                try:
                    v = lp.AsDouble()
                    if v and v > 1e-9:
                        return float(v)
                except Exception:
                    pass
    except Exception:
        pass
    return None


def _wf_z_range_ft(wf):
    bb = wf.get_BoundingBox(None)
    if bb is None:
        return None, None
    try:
        return float(bb.Min.Z), float(bb.Max.Z)
    except Exception:
        return None, None


def _wf_collect_joined_element_ids(document, wf):
    """``ElementId`` de elementos con *Join Geometry* respecto a la zapata."""
    if document is None or wf is None:
        return []
    out = []
    try:
        raw = JoinGeometryUtils.GetJoinedElements(document, wf)
    except Exception:
        return []
    if raw is None:
        return []
    try:
        for eid in raw:
            if eid is None:
                continue
            try:
                ii = element_id_to_int(eid)
            except Exception:
                continue
            if ii is not None:
                out.append(eid)
    except Exception:
        pass
    return out


def _wf_unjoin_all(document, wf, other_ids, avisos):
    """Desune la zapata de cada elemento en ``other_ids``. Añade avisos si falla algún par."""
    if not other_ids:
        return
    for oid in other_ids:
        try:
            oth = document.GetElement(oid)
        except Exception:
            oth = None
        if oth is None:
            continue
        try:
            JoinGeometryUtils.UnjoinGeometry(document, wf, oth)
        except Exception as ex:
            if avisos is not None:
                try:
                    avisos.append(
                        u"Unjoin Id {0}: {1}".format(element_id_to_int(oid), unicode(ex))
                    )
                except Exception:
                    avisos.append(u"Unjoin falló para un elemento unido.")


def _wf_rejoin_all(document, wf, other_ids, avisos):
    """Restaura *Join Geometry* con los mismos elementos."""
    if not other_ids:
        return
    for oid in other_ids:
        try:
            oth = document.GetElement(oid)
        except Exception:
            oth = None
        if oth is None:
            continue
        already_joined = False
        try:
            already_joined = bool(
                JoinGeometryUtils.AreElementsJoined(document, wf.Id, oid)
            )
        except Exception:
            pass
        if already_joined:
            continue
        try:
            JoinGeometryUtils.JoinGeometry(document, wf, oth)
        except Exception as ex:
            if avisos is not None:
                try:
                    avisos.append(
                        u"Join Id {0}: {1}".format(element_id_to_int(oid), unicode(ex))
                    )
                except Exception:
                    avisos.append(u"No se pudo re-unir un elemento.")


def _wf_perp_horizontal_xy(tu_xy):
    """Vector horizontal unitario ⟂ a ``tu_xy`` (proyección XY)."""
    if tu_xy is None:
        return XYZ.BasisX
    v = XYZ(float(tu_xy.X), float(tu_xy.Y), 0.0)
    if float(v.GetLength()) < 1e-12:
        return XYZ.BasisX
    u = v.Normalize()
    return XYZ(-float(u.Y), float(u.X), 0.0).Normalize()


def _wf_alinea_ancho_con_curva_ancho(wplan, width_line):
    """Mantiene el sentido de ``wplan`` alineado con la dirección de ``width_line`` en planta."""
    try:
        wref = width_line.GetEndPoint(1).Subtract(width_line.GetEndPoint(0))
        wxy = XYZ(float(wref.X), float(wref.Y), 0.0)
        if float(wxy.GetLength()) < 1e-12:
            return wplan
        wxy = wxy.Normalize()
        if float(wplan.DotProduct(wxy)) < 0.0:
            try:
                return wplan.Negate()
            except Exception:
                pass
    except Exception:
        pass
    return wplan


def _wf_punto_centro_u_en_franja(long_line, wdir, zmid, bbox_center_xy):
    """
    Centro del tramo horizontal de la U en planta: punto sobre la paralela al eje de la zapata
    (``long_line``) alineado en ancho con ``bbox_center_xy`` (idealmente centro del perímetro
    inferior o punto del ``LocationCurve``, no bbox global del proyecto).
    """
    try:
        lm = long_line.Evaluate(0.5, True)
        dv = XYZ(
            float(bbox_center_xy.X) - float(lm.X),
            float(bbox_center_xy.Y) - float(lm.Y),
            0.0,
        )
        d_w = float(dv.DotProduct(wdir))
        p = lm.Add(wdir.Multiply(d_w))
        return XYZ(float(p.X), float(p.Y), float(zmid))
    except Exception:
        return None


def _wf_traslada_linea_hacia_centro_bbox_planta(line, wf):
    """
    Traslada una ``Line`` en XY para pasar por el centro del **perímetro inferior** (centroide
    de muestreo del borde). Si no hay cara inferior, respaldo al centro del ``BoundingBox``.
    """
    if line is None or wf is None:
        return None
    p0 = line.GetEndPoint(0)
    p1 = line.GetEndPoint(1)
    tu = p1.Subtract(p0)
    if float(tu.GetLength()) < 1e-12:
        return None
    tu = tu.Normalize()
    zmid = 0.5 * (float(p0.Z) + float(p1.Z))
    cx, cy = None, None
    try:
        cxy = centro_xy_perimetro_inferior_doc(wf)
        if cxy is not None:
            cx, cy = float(cxy[0]), float(cxy[1])
    except Exception:
        cx, cy = None, None
    if cx is None:
        try:
            bb = wf.get_BoundingBox(None)
            if bb is None:
                return None
            cx = 0.5 * (float(bb.Min.X) + float(bb.Max.X))
            cy = 0.5 * (float(bb.Min.Y) + float(bb.Max.Y))
        except Exception:
            return None
    try:
        c = XYZ(cx, cy, zmid)
        v = c.Subtract(p0)
        tfoot = p0.Add(tu.Multiply(float(v.DotProduct(tu))))
        delta = c.Subtract(tfoot)
        return Line.CreateBound(p0.Add(delta), p1.Add(delta))
    except Exception:
        return None


def _wf_traslada_linea_hacia_interior_hormigon_mm(line, n_cara_saliente, mm):
    """
    Traslada los extremos de ``line`` hacia el interior del bloque (``-`` normal saliente),
    en mm. Misma convención que ``offset_linea_eje_barra_desde_cara_inferior_mm``.
    Usado para situar longitudinales por **encima** de la capa transversal (U).
    """
    if line is None:
        return None
    try:
        m = float(mm)
    except Exception:
        m = 0.0
    if m < 1e-6:
        return line
    try:
        if n_cara_saliente is not None and float(n_cara_saliente.GetLength()) > 1e-12:
            inward = n_cara_saliente.Normalize().Negate()
        else:
            inward = XYZ.BasisZ
        d_ft = _mm_to_ft(m)
        v = inward.Multiply(d_ft)
        return Line.CreateBound(
            line.GetEndPoint(0).Add(v),
            line.GetEndPoint(1).Add(v),
        )
    except Exception:
        return line


def _wf_span_luz_distribucion_bbox_ft(wf, line_ref, lado_malla, cap_geom_ft):
    """
    Luz del conjunto (pies) para *Wall Foundation*: proyección del **perímetro real** de la
    cara inferior sobre la dirección ⟂ a ``line_ref`` en planta (menos 2×rec), acotada por
    ``cap_geom_ft``. El parámetro ``lado_malla`` se ignora (histórico fundación aislada/bbox).

    No usa ``BoundingBox`` del proyecto — evita errores en zapatas giradas.
    """
    s = None
    try:
        s = luz_proyeccion_perimetro_inferior_ft(
            wf, line_ref, _REC_OFF_PLANTA_INF_MM, True
        )
    except Exception:
        s = None
    cap = max(0.0, float(cap_geom_ft))
    if s is None or float(s) <= 1e-9:
        out = cap if cap > 1e-9 else _mm_to_ft(10.0)
        return max(float(out), _mm_to_ft(10.0))
    s = float(s)
    if cap > 1e-9:
        s = min(s, cap)
    return max(s, _mm_to_ft(10.0))


def _wf_span_luz_along_eje_wall_desde_perimetro_ft(wf, line_ref, cap_ft):
    """
    Luz a lo largo del eje de la zapata (misma tangente en planta que ``line_ref``), desde
    el contorno inferior real — para reparto de transversales. Acota con ``cap_ft``.
    """
    s = None
    try:
        s = luz_proyeccion_perimetro_inferior_ft(
            wf, line_ref, _REC_OFF_PLANTA_INF_MM, False
        )
    except Exception:
        s = None
    cap = max(0.0, float(cap_ft))
    if s is None or float(s) <= 1e-9:
        out = cap if cap > 1e-9 else _mm_to_ft(10.0)
        return max(float(out), _mm_to_ft(10.0))
    s = float(s)
    if cap > 1e-9:
        s = min(s, cap)
    return max(s, _mm_to_ft(10.0))


def _wf_normales_prioridad_ancho_en_planta(width_line, long_line):
    """
    Normal de ``CreateFromCurves*`` para longitudinales: **dirección del ancho en planta**.
    Revit reparte el Rebar Set a lo largo de ±norm; con ``BasisZ`` el conjunto crecía en vertical.
    """
    try:
        q0 = width_line.GetEndPoint(0)
        q1 = width_line.GetEndPoint(1)
        w = XYZ(float(q1.X - q0.X), float(q1.Y - q0.Y), 0.0)
        if float(w.GetLength()) < 1e-12:
            return None
        w = w.Normalize()
        p0 = long_line.GetEndPoint(0)
        p1 = long_line.GetEndPoint(1)
        t = XYZ(float(p1.X - p0.X), float(p1.Y - p0.Y), 0.0)
        if float(t.GetLength()) > 1e-12:
            t = t.Normalize()
            if abs(float(t.DotProduct(w))) > 0.995:
                return None
        return [w]
    except Exception:
        return None


def _wf_norm_distribucion_longitudinal_en_planta(ln, width_line):
    """
    Normal para reparto **en planta** de longitudinales: **perpendicular al eje de la barra**
    (no la dirección de ``width_line``, que en trapecios puede no ser ⟂ al eje de la zapata).
    El sentido se alinea con ``width_line`` cuando hay ``dot(n, w) < 0``.
    """
    try:
        tu_ln = ln.GetEndPoint(1).Subtract(ln.GetEndPoint(0))
        tu_xy = XYZ(float(tu_ln.X), float(tu_ln.Y), 0.0)
        if float(tu_xy.GetLength()) < 1e-12:
            return None, None
        tu_xy = tu_xy.Normalize()
        n = _wf_perp_horizontal_xy(tu_xy)
        if width_line is not None:
            try:
                q0 = width_line.GetEndPoint(0)
                q1 = width_line.GetEndPoint(1)
                w = XYZ(float(q1.X - q0.X), float(q1.Y - q0.Y), 0.0)
                if float(w.GetLength()) > 1e-12:
                    w = w.Normalize()
                    if float(n.DotProduct(w)) < 0.0:
                        n = n.Negate()
            except Exception:
                pass
        return [n], n
    except Exception:
        return None, None


def _geometria_wf_cara_inferior_tol(
    wf, diam_long_mm, diam_trans_mm, tol, tnz, ultra_fallback=False
):
    """Un intento con tolerancias dadas (cara inferior fundación aislada)."""
    marco = obtener_marco_coordenadas_cara_inferior(
        wf, tol, tnz, ultra_fallback
    )
    if marco is None:
        return None
    n_cara = marco[3]
    r_men = extraer_curva_lado_menor_cara_inferior(wf, tol, tnz, ultra_fallback)
    r_may = extraer_curva_lado_mayor_cara_inferior(wf, tol, tnz, ultra_fallback)
    if r_men is None or r_may is None:
        return None
    c_men, _ = r_men
    c_may, _ = r_may
    try:
        d_l = float(diam_long_mm) if diam_long_mm else 0.0
        d_t = float(diam_trans_mm) if diam_trans_mm else 0.0
    except Exception:
        d_l = d_t = 0.0
    ext_long_mm = float(_REC_EXTREMOS_LONG_TANGENTE_MM)
    if d_l > 1e-6:
        ext_long_mm = float(_REC_EXTREMOS_LONG_TANGENTE_MM) + 0.5 * d_l
    ct_men, _ = aplicar_recubrimiento_inferior_completo_mm(
        c_men, wf, _REC_OFF_PLANTA_INF_MM, _REC_EXTREMOS_INFERIOR_MM
    )
    ct_may, _ = aplicar_recubrimiento_inferior_completo_mm(
        c_may, wf, _REC_OFF_PLANTA_INF_MM, ext_long_mm
    )
    if ct_men is None or ct_may is None:
        return None
    long_bar = offset_linea_eje_barra_desde_cara_inferior_mm(
        ct_may, n_cara, _RECO_HOR_MM, d_l
    )
    width_bar = offset_linea_eje_barra_desde_cara_inferior_mm(
        ct_men, n_cara, _RECO_HOR_MM, d_t
    )
    if long_bar is None or width_bar is None:
        return None
    ev = evaluar_caras_paralelas_curva_mas_cercana(wf, long_bar)
    cara_pp = None
    if isinstance(ev, dict):
        cara_pp = ev.get("mejor")
        if cara_pp is None:
            cara_pp = ev.get(u"mejor")
    z0, z1 = _wf_z_range_ft(wf)
    usable_w = float(width_bar.Length)
    return {
        "long_line": long_bar,
        "width_line": width_bar,
        "marco_uvn": marco,
        "cara_pp": cara_pp,
        "n_cara": n_cara,
        "z0": z0,
        "z1": z1,
        "usable_w_ft": usable_w,
    }


def _wf_geometria_fallback_bbox_location(wf, diam_long_mm, diam_trans_mm):
    """
    Respaldo para ``WallFoundation`` cuando no hay cara inferior reconocible
    (sólidos de símbolo, teselación, zapatas muy bajas, etc.): BoundingBox + eje
    de ``LocationCurve`` o lado mayor de la caja en planta.
    """
    bb = wf.get_BoundingBox(None)
    if bb is None:
        return None, u"Respaldo: sin BoundingBox."
    z0 = float(bb.Min.Z)
    z1 = float(bb.Max.Z)
    dx = float(bb.Max.X - bb.Min.X)
    dy = float(bb.Max.Y - bb.Min.Y)
    cx = 0.5 * (float(bb.Min.X) + float(bb.Max.X))
    cy = 0.5 * (float(bb.Min.Y) + float(bb.Max.Y))
    p0b = None
    p1b = None
    axis, _ = _axis_line_wall_foundation(wf)
    if axis is not None:
        p0 = axis.GetEndPoint(0)
        p1 = axis.GetEndPoint(1)
        p0b = XYZ(float(p0.X), float(p0.Y), z0)
        p1b = XYZ(float(p1.X), float(p1.Y), z0)
    if p0b is None or p1b is None:
        if dx >= dy:
            p0b = XYZ(float(bb.Min.X), cy, z0)
            p1b = XYZ(float(bb.Max.X), cy, z0)
        else:
            p0b = XYZ(cx, float(bb.Min.Y), z0)
            p1b = XYZ(cx, float(bb.Max.Y), z0)
    try:
        dxy = XYZ(
            float(p1b.X - p0b.X),
            float(p1b.Y - p0b.Y),
            0.0,
        )
        if float(dxy.GetLength()) < 1e-9:
            return None, u"Respaldo: longitud nula en planta."
        tu = dxy.Normalize()
    except Exception:
        return None, u"Respaldo: dirección longitudinal inválida."
    pm = p0b.Add(p1b.Subtract(p0b).Multiply(0.5))
    wdir = _wf_perp_horizontal_xy(tu)
    w_ft = _wf_width_ft(wf)
    if w_ft is None or w_ft < 1e-6:
        w_ft = min(dx, dy)
    if w_ft < 1e-6:
        return None, u"Respaldo: ancho nulo."
    half = 0.5 * float(w_ft)
    wm0 = pm.Subtract(wdir.Multiply(half))
    wm1 = pm.Add(wdir.Multiply(half))
    long_raw = Line.CreateBound(p0b, p1b)
    width_raw = Line.CreateBound(
        XYZ(float(wm0.X), float(wm0.Y), z0),
        XYZ(float(wm1.X), float(wm1.Y), z0),
    )
    n_out = XYZ.BasisZ.Negate()
    marco_syn = (pm, tu, wdir, n_out)
    try:
        d_l = float(diam_long_mm) if diam_long_mm else 0.0
        d_t = float(diam_trans_mm) if diam_trans_mm else 0.0
    except Exception:
        d_l = d_t = 0.0
    ext_long_mm = float(_REC_EXTREMOS_LONG_TANGENTE_MM)
    if d_l > 1e-6:
        ext_long_mm = float(_REC_EXTREMOS_LONG_TANGENTE_MM) + 0.5 * d_l
    ct_men, _ = aplicar_recubrimiento_inferior_completo_mm(
        width_raw, wf, _REC_OFF_PLANTA_INF_MM, _REC_EXTREMOS_INFERIOR_MM
    )
    ct_may, _ = aplicar_recubrimiento_inferior_completo_mm(
        long_raw, wf, _REC_OFF_PLANTA_INF_MM, ext_long_mm
    )
    if ct_men is None or ct_may is None:
        return None, u"Respaldo: recubrimiento dejó curva nula."
    long_bar = offset_linea_eje_barra_desde_cara_inferior_mm(
        ct_may, n_out, _RECO_HOR_MM, d_l
    )
    width_bar = offset_linea_eje_barra_desde_cara_inferior_mm(
        ct_men, n_out, _RECO_HOR_MM, d_t
    )
    if long_bar is None or width_bar is None:
        return None, u"Respaldo: offset eje de barra falló."
    ev = evaluar_caras_paralelas_curva_mas_cercana(wf, long_bar)
    cara_pp = None
    if isinstance(ev, dict):
        cara_pp = ev.get("mejor")
        if cara_pp is None:
            cara_pp = ev.get(u"mejor")
    usable_w = float(width_bar.Length)
    return {
        "long_line": long_bar,
        "width_line": width_bar,
        "marco_uvn": marco_syn,
        "cara_pp": cara_pp,
        "n_cara": n_out,
        "z0": z0,
        "z1": z1,
        "usable_w_ft": usable_w,
    }, None


def _geometria_wall_foundation_inferior(wf, diam_long_mm, diam_trans_mm):
    """
    Primero intenta **cortes planos** al sólido definidos por el ``LocationCurve``
    del muro host (``geometria_wall_foundation_cortes_muro``). Si no aplica,
    usa cara inferior + curvas mayor/menor (lógica fundación aislada) con varias
    tolerancias; luego un pase «ultra» (soleiras teseladas/inclinadas) vía
    ``ultra_fallback`` en geometría compartida; si todo falla, respaldo
    BoundingBox + eje.

    Debe llamarse con geometría fiable (p. ej. tras ``UnjoinGeometry`` y ``Regenerate``).

    Returns:
        tuple: ``(dict | None, mensaje_pista | None)`` — el segundo texto solo si hubo respaldo
        o fallo total (para avisos).
    """
    if wf is None:
        return None, u"Elemento nulo."
    try:
        g_cut = geometria_inferior_wall_foundation_cortes_muro(
            wf, diam_long_mm, diam_trans_mm
        )
    except Exception:
        g_cut = None
    if g_cut is not None:
        return g_cut, None
    tol_grid = (
        (0.05, 0.18),
        (0.12, 0.30),
        (0.25, 0.45),
        (0.50, 0.70),
    )
    for ultra in (False, True):
        for tol, tnz in tol_grid:
            try:
                g = _geometria_wf_cara_inferior_tol(
                    wf, diam_long_mm, diam_trans_mm, tol, tnz, ultra
                )
            except Exception:
                g = None
            if g is not None:
                return g, None
    fb, err_fb = _wf_geometria_fallback_bbox_location(wf, diam_long_mm, diam_trans_mm)
    if fb is not None:
        return fb, u"Geometría por respaldo (bbox/eje). Revise posición en modelo."
    return None, err_fb or u"No se extrajeron curvas del perímetro inferior."


def _longitud_eje_hint_mm(wf):
    """Largo característico del eje (mm) antes de transacción — sin desunir; puede ser aproximado."""
    if wf is None:
        return 0.0
    ax0, _ = _axis_line_wall_foundation(wf)
    if ax0 is not None:
        try:
            return _ft_to_mm(float(ax0.Length))
        except Exception:
            pass
    for tol, tnz in ((0.05, 0.18), (0.25, 0.45), (0.50, 0.70)):
        r = extraer_curva_lado_mayor_cara_inferior(wf, tol, tnz)
        if r is not None and r[0] is not None:
            try:
                return _ft_to_mm(float(r[0].Length))
            except Exception:
                pass
    bb = wf.get_BoundingBox(None)
    if bb is not None:
        try:
            dx = float(bb.Max.X - bb.Min.X)
            dy = float(bb.Max.Y - bb.Min.Y)
            return max(dx, dy) * 304.8
        except Exception:
            pass
    return 0.0


def _punto_centro_ancho_en_estacion(p_sta, wdir_unit, width_line):
    """
    Punto en el eje ancho (línea de la armadura transversal) más alineado con la estación
    ``p_sta`` sobre la zapata rectangular.
    """
    try:
        w0 = width_line.GetEndPoint(0)
        w1 = width_line.GetEndPoint(1)
        wmid = w0.Add(w1.Subtract(w0).Multiply(0.5))
        dv = wmid.Subtract(p_sta)
        dist = float(dv.DotProduct(wdir_unit))
        return p_sta.Add(wdir_unit.Multiply(dist))
    except Exception:
        return p_sta


def _axis_line_wall_foundation(wf):
    if wf is None:
        return None, u"Elemento nulo."
    loc = wf.Location
    if not isinstance(loc, LocationCurve):
        return None, u"La zapata no tiene LocationCurve."
    c = loc.Curve
    if c is None:
        return None, u"Curva de eje nula."
    if isinstance(c, Line):
        return c, None
    try:
        p0 = c.GetEndPoint(0)
        p1 = c.GetEndPoint(1)
        ln = Line.CreateBound(p0, p1)
        return ln, None
    except Exception as ex:
        return None, unicode(ex)


def _wf_tu_xy_desde_linea(line):
    """Unitario en XY de la dirección de una ``Line`` (proyección horizontal)."""
    if line is None:
        return None
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        v = XYZ(float(p1.X - p0.X), float(p1.Y - p0.Y), 0.0)
        if float(v.GetLength()) < 1e-12:
            return None
        return v.Normalize()
    except Exception:
        return None


def _wf_tu_xy_desde_axis_wall_foundation(ax):
    return _wf_tu_xy_desde_linea(_wf_location_curve_como_linea(ax))


def _wf_punto_referencia_planta_wall_foundation(wf, zmid):
    """Punto medio del ``LocationCurve`` de la zapata, con Z dada (host en planta)."""
    ax, _ = _axis_line_wall_foundation(wf)
    if ax is None:
        return None
    try:
        pm = ax.Evaluate(0.5, True)
        return XYZ(float(pm.X), float(pm.Y), float(zmid))
    except Exception:
        return None


def _wf_traslada_linea_hacia_punto_planta_xy(line, px, py):
    """
    Traslada ``line`` en XY (paralela a sí misma) para que pase por ``(px, py)`` en planta;
    mantiene la media de Z de los extremos.
    """
    if line is None:
        return None
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        tu = p1.Subtract(p0)
        if float(tu.GetLength()) < 1e-12:
            return None
        tu = tu.Normalize()
        zmid = 0.5 * (float(p0.Z) + float(p1.Z))
        c = XYZ(float(px), float(py), zmid)
        v = c.Subtract(p0)
        tfoot = p0.Add(tu.Multiply(float(v.DotProduct(tu))))
        delta = c.Subtract(tfoot)
        return Line.CreateBound(p0.Add(delta), p1.Add(delta))
    except Exception:
        return None


def _wf_location_curve_como_linea(ax):
    """``LocationCurve`` como ``Line`` (cuerda si la API devuelve arco u otra curva)."""
    if ax is None:
        return None
    if isinstance(ax, Line):
        return ax
    try:
        p0 = ax.GetEndPoint(0)
        p1 = ax.GetEndPoint(1)
        return Line.CreateBound(p0, p1)
    except Exception:
        return None


def _wf_line_eje_inferior_crudo_desde_location(wf, ax, zmid, diam_long_mm):
    """
    Eje longitudinal en planta (sin ``offset_linea_eje_barra``) según ``LocationCurve``,
    con el mismo criterio de extremos que ``ext_long_mm``. La longitud queda acotada por la
    **proyección del perímetro inferior** sobre la tangente (huella real), para no exceder
    el hormigón cuando el eje de Revit es más largo que la zapata en planta.
    """
    ax_ln = _wf_location_curve_como_linea(ax)
    if ax_ln is None:
        return None
    try:
        L = float(ax_ln.Length)
        if L < 1e-9:
            return None
        ext_long_mm = float(_REC_EXTREMOS_LONG_TANGENTE_MM)
        try:
            if diam_long_mm and float(diam_long_mm) > 1e-6:
                ext_long_mm += 0.5 * float(diam_long_mm)
        except Exception:
            pass
        ext_ft = _mm_to_ft(ext_long_mm)
        if L <= 2.0 * ext_ft + 1e-9:
            return None
        p0a = ax_ln.GetEndPoint(0)
        p1a = ax_ln.GetEndPoint(1)
        dvec = p1a.Subtract(p0a)
        tu_xy = XYZ(float(dvec.X), float(dvec.Y), 0.0)
        if float(tu_xy.GetLength()) < 1e-12:
            return None
        tu_xy = tu_xy.Normalize()
        pm = ax_ln.Evaluate(0.5, True)
        c = XYZ(float(pm.X), float(pm.Y), float(zmid))
        h = 0.5 * L
        half_len = h - ext_ft
        if wf is not None:
            try:
                raw_span = span_bruto_proyeccion_perimetro_inferior_ft(
                    wf,
                    XYZ(float(pm.X), float(pm.Y), float(zmid)),
                    tu_xy,
                    None,
                )
                if raw_span is not None and raw_span > 2.0 * ext_ft + 1e-9:
                    half_cap = 0.5 * float(raw_span) - ext_ft
                    if half_cap > 1e-9:
                        half_len = min(half_len, half_cap)
            except Exception:
                pass
        if half_len < 1e-9:
            return None
        pa = c.Subtract(tu_xy.Multiply(half_len))
        pb = c.Add(tu_xy.Multiply(half_len))
        return Line.CreateBound(pa, pb)
    except Exception:
        return None


def _wf_geo_alinear_strip_a_location_wall_foundation(
    wf, geo, diam_long_mm, diam_trans_mm
):
    """
    Fuerza ``long_line`` / ``width_line``: el eje longitudinal sigue el ``LocationCurve``
    de la zapata; la **dirección de transversales** (``width_line`` en planta) sale del
    ``LocationCurve`` del **muro host** (⟂ a la tangente del muro, vía
    ``vector_transversal_planta_desde_muro_host``), no de la perpendicular al eje de la
    soleira (que puede desalinearse del muro). Respaldo: perpendicular al eje de zapata.

    Si ``geo['use_cortes_lines_for_rebar']`` es verdadero (geometría por cortes
    plano∩sólido desde el muro host), **no modifica** ``long_line`` /
    ``width_line``: el armado usa exactamente esas ``Line``.
    """
    if wf is None or geo is None:
        return
    try:
        if geo.get("use_cortes_lines_for_rebar"):
            return
    except Exception:
        pass
    ax, _ = _axis_line_wall_foundation(wf)
    ax_ln = _wf_location_curve_como_linea(ax)
    tu_loc = _wf_tu_xy_desde_axis_wall_foundation(ax_ln)
    if tu_loc is None:
        return
    long_line = geo.get("long_line")
    width_line = geo.get("width_line")
    n_cara = geo.get("n_cara")
    if long_line is None or width_line is None or n_cara is None:
        return
    try:
        d_l = float(diam_long_mm) if diam_long_mm else 0.0
        d_t = float(diam_trans_mm) if diam_trans_mm else 0.0
    except Exception:
        d_l = d_t = 0.0
    try:
        pm = ax_ln.Evaluate(0.5, True)
        px, py = float(pm.X), float(pm.Y)
        zl = 0.5 * (
            float(long_line.GetEndPoint(0).Z) + float(long_line.GetEndPoint(1).Z)
        )
        zw = 0.5 * (
            float(width_line.GetEndPoint(0).Z) + float(width_line.GetEndPoint(1).Z)
        )
    except Exception:
        return
    long_raw = _wf_line_eje_inferior_crudo_desde_location(wf, ax_ln, zl, d_l)
    if long_raw is None:
        nl = _wf_traslada_linea_hacia_punto_planta_xy(long_line, px, py)
        nw = _wf_traslada_linea_hacia_punto_planta_xy(width_line, px, py)
        if nl is not None:
            geo["long_line"] = nl
        if nw is not None:
            geo["width_line"] = nw
        return
    try:
        new_long = offset_linea_eje_barra_desde_cara_inferior_mm(
            long_raw, n_cara, _RECO_HOR_MM, d_l
        )
    except Exception:
        new_long = None
    if new_long is None:
        nl = _wf_traslada_linea_hacia_punto_planta_xy(long_line, px, py)
        nw = _wf_traslada_linea_hacia_punto_planta_xy(width_line, px, py)
        if nl is not None:
            geo["long_line"] = nl
        if nw is not None:
            geo["width_line"] = nw
        return
    wdir = vector_transversal_planta_desde_muro_host(wf, pm)
    if wdir is None:
        wdir = _wf_perp_horizontal_xy(tu_loc)
    wdir = _wf_alinea_ancho_con_curva_ancho(wdir, width_line)
    try:
        wlen = float(width_line.Length)
    except Exception:
        wlen = 0.0
    if wlen < 1e-9:
        return
    c_w = XYZ(px, py, zw)
    w0 = c_w.Subtract(wdir.Multiply(0.5 * wlen))
    w1 = c_w.Add(wdir.Multiply(0.5 * wlen))
    try:
        width_raw = Line.CreateBound(w0, w1)
        new_width = offset_linea_eje_barra_desde_cara_inferior_mm(
            width_raw, n_cara, _RECO_HOR_MM, d_t
        )
    except Exception:
        new_width = None
    if new_width is None:
        return
    geo["long_line"] = new_long
    geo["width_line"] = new_width
    try:
        geo["usable_w_ft"] = float(new_width.Length)
    except Exception:
        pass
    try:
        ev = evaluar_caras_paralelas_curva_mas_cercana(wf, new_long)
        if isinstance(ev, dict):
            cp = ev.get("mejor")
            if cp is None:
                cp = ev.get(u"mejor")
            if cp is not None:
                geo["cara_pp"] = cp
    except Exception:
        pass


def _split_line_laps(line, max_len_ft, lap_ft):
    if line is None or not isinstance(line, Line):
        return []
    p0 = line.GetEndPoint(0)
    p1 = line.GetEndPoint(1)
    dvec = p1.Subtract(p0)
    L = float(dvec.GetLength())
    if L < 1e-9:
        return []
    edir = dvec.Normalize()
    if max_len_ft <= 1e-9:
        return [line]
    out = []
    start_d = 0.0
    guard = 0
    while start_d < L - 1e-9:
        guard += 1
        if guard > 10000:
            break
        end_d = min(start_d + max_len_ft, L)
        a = p0.Add(edir.Multiply(start_d))
        b = p0.Add(edir.Multiply(end_d))
        seg = Line.CreateBound(a, b)
        out.append(seg)
        if end_d >= L - 1e-9:
            break
        start_d = end_d - lap_ft
        if start_d < 0:
            start_d = 0.0
    return out


def _split_line_laps_longitudinal_eje_stock(line, max_bar_mm, pata_tab_mm, lap_mm):
    """
    Trocea el eje de longitudinales de forma que el **desarrollado** por barra (eje + patas
    según tabla) respete ``max_bar_mm`` stock:

    - **Primera** barra: tramo de eje ≤ ``max_bar_mm − pata_tab_mm`` (una pata al inicio).
    - **Intermedias**: eje ≤ ``max_bar_mm`` (sin patas en planta entre empalmes).
    - **Última**: eje ≤ ``max_bar_mm − pata_tab_mm`` (pata final); el largo de eje es el
      **remanente** tras empalmes intermedios al máximo de plantilla, no se iguala al del primer tramo.

    ``pata_tab_mm`` es el largo de pata de **tabla por ø**, no el tramo ya compensado para Revit.
    """
    if line is None or not isinstance(line, Line):
        return []
    try:
        mb = float(max_bar_mm)
        p_mm = float(pata_tab_mm)
        lap_mm = float(lap_mm)
    except Exception:
        return []
    if mb <= 1e-9 or p_mm < 0.0 or lap_mm < 0.0:
        return []
    lap_ft = _mm_to_ft(lap_mm)
    max_mid_ft = _mm_to_ft(mb)
    max_one_hook_axis_mm = max(mb - p_mm, 300.0)
    max_one_hook_ft = _mm_to_ft(max_one_hook_axis_mm)
    max_both_hooks_axis_mm = max(mb - 2.0 * p_mm, 300.0)
    max_both_hooks_ft = _mm_to_ft(max_both_hooks_axis_mm)

    p0 = line.GetEndPoint(0)
    p1 = line.GetEndPoint(1)
    dvec = p1.Subtract(p0)
    Ltot = float(dvec.GetLength())
    if Ltot < 1e-9:
        return []
    edir = dvec.Normalize()

    if Ltot <= max_both_hooks_ft + 1e-9:
        return [Line.CreateBound(p0, p1)]

    lengths_ft = []
    L1 = min(max_one_hook_ft, Ltot)
    rem = Ltot - L1 + lap_ft
    # `Ltot > max_both`: hacen falta al menos dos barras; si L1 = Ltot, `rem ≈ lap` y el troceo falla.
    if Ltot > max_both_hooks_ft + 1e-9 and rem <= lap_ft + _mm_to_ft(0.01):
        need_first = max(
            max_both_hooks_ft + _mm_to_ft(0.1),
            Ltot - max_one_hook_ft + lap_ft,
        )
        L1 = min(max_one_hook_ft, need_first)
        rem = Ltot - L1 + lap_ft
    lengths_ft.append(L1)
    LMIN_REM_FT = _mm_to_ft(50.0)
    while rem > max_one_hook_ft + 1e-9:
        ra = rem - max_mid_ft + lap_ft
        if ra > max_one_hook_ft + 1e-9:
            lengths_ft.append(max_mid_ft)
            rem = ra
        elif ra > 1e-9:
            lengths_ft.append(max_mid_ft)
            rem = ra
            break
        else:
            lf = rem + lap_ft - max_one_hook_ft
            L_mid = min(max_mid_ft, rem + lap_ft - LMIN_REM_FT)
            if L_mid + 1e-9 < lf:
                L_mid = min(max_mid_ft, max(lf, LMIN_REM_FT))
            if L_mid > 1e-9:
                lengths_ft.append(L_mid)
                rem = rem - L_mid + lap_ft
            break
    if rem > 1e-9:
        lengths_ft.append(rem)

    out = []
    pos = 0.0
    for i, ell in enumerate(lengths_ft):
        try:
            ell = float(ell)
        except Exception:
            continue
        if ell < 1e-9:
            continue
        if i > 0:
            pos -= lap_ft
        a = p0.Add(edir.Multiply(pos))
        b = p0.Add(edir.Multiply(pos + ell))
        out.append(Line.CreateBound(a, b))
        pos += ell
    return out


def _wf_puntos_simbologia_empalme_entre_segmentos_eje(segs):
    """
    Por cada junta entre tramos del eje de stock (troceo con traslape), devuelve
    ``(p0, p1)`` en 3D: inicio del tramo siguiente y fin del tramo anterior = zona de empalme.
    """
    out = []
    if not segs or len(segs) < 2:
        return out
    n = len(segs)
    for j in range(n - 1):
        try:
            s0 = segs[j]
            s1 = segs[j + 1]
            if s0 is None or s1 is None:
                continue
            pa = s1.GetEndPoint(0)
            pb = s0.GetEndPoint(1)
        except Exception:
            continue
        try:
            if pa.DistanceTo(pb) < _mm_to_ft(2.0):
                continue
        except Exception:
            continue
        out.append((pa, pb))
    return out


def _wf_inward_dirs_para_cota_empalme(wf, geo, pa, pb):
    """
    Direcciones para desplazar la línea de cota del traslape (mismo criterio que vigas):
    ``inward_3d`` = hacia el interior del hormigón desde la cara inferior; ``inward_xy`` =
    en planta, ⟂ al eje del empalme y apuntando hacia el eje / centro de referencia de la zapata.
    """
    inward_3d = None
    inward_xy = None
    if geo is not None:
        try:
            n_cara = geo.get("n_cara")
            if n_cara is not None and float(n_cara.GetLength()) > 1e-12:
                inward_3d = n_cara.Normalize().Negate()
        except Exception:
            inward_3d = None
    if wf is None or pa is None or pb is None:
        return inward_xy, inward_3d
    try:
        ax = pb.Subtract(pa)
        u = XYZ(float(ax.X), float(ax.Y), 0.0)
        if float(u.GetLength()) < 1e-12:
            return inward_xy, inward_3d
        u = u.Normalize()
        w = XYZ(-float(u.Y), float(u.X), 0.0)
        if float(w.GetLength()) < 1e-12:
            return inward_xy, inward_3d
        w = w.Normalize()
        pm = XYZ(
            0.5 * (float(pa.X) + float(pb.X)),
            0.5 * (float(pa.Y) + float(pb.Y)),
            0.5 * (float(pa.Z) + float(pb.Z)),
        )
        pref = _wf_punto_referencia_planta_wall_foundation(wf, float(pm.Z))
        if pref is None:
            try:
                gc = centro_xy_perimetro_inferior_doc(wf)
                if gc is not None:
                    pref = XYZ(float(gc[0]), float(gc[1]), float(pm.Z))
            except Exception:
                pref = None
        if pref is None:
            return inward_xy, inward_3d
        to_pref = XYZ(
            float(pref.X) - float(pm.X),
            float(pref.Y) - float(pm.Y),
            0.0,
        )
        if float(to_pref.GetLength()) < 1e-9:
            return inward_xy, inward_3d
        to_pref = to_pref.Normalize()
        if float(w.DotProduct(to_pref)) < 0.0:
            w = w.Negate()
        inward_xy = w
    except Exception:
        pass
    return inward_xy, inward_3d


def _wf_colocar_simbologia_empalme_eje(doc, view, segs, avisos, wf=None, geo=None):
    """
    Detail line-based de empalme (misma familia que vigas / borde losa) en la vista activa,
    con cota lineal del traslape entre referencias Left/Right del símbolo cuando la vista lo permite.

    Returns:
        ``(n_ok, lap_infos)`` donde ``lap_infos`` es una lista de dicts
        ``joint_idx, inst, dim_id`` para vincular al DMU tras crear los Rebar por tramo.
    """
    if doc is None or view is None or not segs or len(segs) < 2:
        return 0, []
    if not vista_permite_detail_curve(view):
        avisos.append(
            u"Simbología de empalme: la vista activa no admite detail components "
            u"(vistas 3D o plantilla); use planta o elevación/corte."
        )
        return 0, []
    try:
        from enfierrado_shaft_hashtag import (
            _create_overlap_dimension_from_detail_refs,
            _get_named_left_right_refs_from_detail_instance,
            _view_accepts_overlap_dimension,
        )
    except Exception:
        _create_overlap_dimension_from_detail_refs = None
        _get_named_left_right_refs_from_detail_instance = None
        _view_accepts_overlap_dimension = None
    sid, sym_err = _find_fixed_lap_detail_symbol_id(doc)
    if sid is None:
        if sym_err:
            avisos.append(sym_err)
        return 0, []
    lap_sym = doc.GetElement(sid)
    if lap_sym is None:
        if sym_err:
            avisos.append(sym_err)
        return 0, []
    if not isinstance(lap_sym, FamilySymbol):
        avisos.append(u"Simbología de empalme: el símbolo no es FamilySymbol.")
        return 0, []
    puntos = _wf_puntos_simbologia_empalme_entre_segmentos_eje(segs)
    n_ok = 0
    n_dim = 0
    aviso_refs_lap = None
    lap_infos = []
    do_dim = (
        _view_accepts_overlap_dimension is not None
        and _create_overlap_dimension_from_detail_refs is not None
        and _get_named_left_right_refs_from_detail_instance is not None
        and _view_accepts_overlap_dimension(view)
    )
    for joint_idx, (pa, pb) in enumerate(puntos):
        ok_d, err_d, lap_inst = _colocar_detail_item_traslape_en_vista(
            doc, view, lap_sym, pa, pb
        )
        if not ok_d:
            if err_d:
                avisos.append(err_d)
            continue
        n_ok += 1
        dim_eid = None
        if lap_inst is not None and do_dim:
            ref_l, ref_r, ref_err = _get_named_left_right_refs_from_detail_instance(
                lap_inst
            )
            if ref_err and aviso_refs_lap is None:
                aviso_refs_lap = ref_err
            if ref_l is not None and ref_r is not None:
                axis_u = None
                try:
                    dv = pb.Subtract(pa)
                    if dv.GetLength() > 1e-9:
                        axis_u = dv.Normalize()
                except Exception:
                    axis_u = None
                inward_xy, inward_3d = _wf_inward_dirs_para_cota_empalme(
                    wf, geo, pa, pb
                )
                ok_dm, msg_dm, dim_data = _create_overlap_dimension_from_detail_refs(
                    doc,
                    view,
                    ref_l,
                    ref_r,
                    pa,
                    pb,
                    axis_u,
                    lateral_hint=None,
                    line_offset_mm=450.0,
                    inward_dir_xy=inward_xy,
                    inward_dir_3d=inward_3d,
                    use_view_plane_dim_line=True,
                    flip_dimension_side=False,
                )
                if ok_dm and dim_data and dim_data.get("dim_id") is not None:
                    n_dim += 1
                    try:
                        dim_eid = ElementId(int(dim_data["dim_id"]))
                    except Exception:
                        dim_eid = None
                elif msg_dm:
                    avisos.append(
                        u"Cota empalme (longitudinales): {0}".format(msg_dm)
                    )
        lap_infos.append(
            {
                "joint_idx": int(joint_idx),
                "inst": lap_inst,
                "dim_id": dim_eid,
            }
        )
    if sym_err and n_ok > 0:
        avisos.append(sym_err)
    if aviso_refs_lap:
        avisos.append(aviso_refs_lap)
    if n_dim > 0:
        try:
            avisos.append(
                u"Cotas de traslape en símbolos de empalme: {0}.".format(int(n_dim))
            )
        except Exception:
            pass
    return n_ok, lap_infos


def _resolver_bar_type_from_combo(document, cmb, entries):
    if cmb is None:
        return None, u"Combo diámetro no encontrado."
    try:
        idx = int(cmb.SelectedIndex)
        if 0 <= idx < len(entries):
            bt, lbl = entries[idx]
            if bt is not None:
                return bt, None
            mm = _parse_diameter_mm_from_bar_combo_label(lbl)
            if mm is not None and document is not None:
                try:
                    from enfierrado_shaft_hashtag import resolver_bar_type_por_diametro_mm

                    bt2, _, _ = resolver_bar_type_por_diametro_mm(document, float(mm))
                    if bt2 is not None:
                        return bt2, None
                except Exception:
                    pass
    except Exception:
        pass
    try:
        sel = cmb.SelectedItem
        lab = unicode(sel) if sel is not None else u""
    except Exception:
        lab = u""
    for bt, lbl in entries:
        if unicode(lbl) == lab and bt is not None:
            return bt, None
    mm = _parse_diameter_mm_from_bar_combo_label(lab)
    if mm is not None and document is not None:
        try:
            from enfierrado_shaft_hashtag import resolver_bar_type_por_diametro_mm

            bt3, _, _ = resolver_bar_type_por_diametro_mm(document, float(mm))
            if bt3 is not None:
                return bt3, None
        except Exception:
            pass
    return None, u"No se pudo resolver RebarBarType."


def _ubicar_punto_eje_menos_recorte(p0, p1, tangent, recorte_ft):
    try:
        if float(tangent.GetLength()) < 1e-12:
            return None, None
        tu = tangent.Normalize()
    except Exception:
        return None, None
    return p0.Add(tu.Multiply(recorte_ft)), p1.Subtract(tu.Multiply(recorte_ft))


# --- XAML (mismo chrome oscuro que ``enfierrado_fundacion_aislada``) -----------------
_WF_XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        x:Name="WallFoundWin"
        Title="Armadura Fundacion Corrida"
        SizeToContent="Height" MaxHeight="920" WindowStartupLocation="Manual"
        Background="Transparent" AllowsTransparency="True" FontFamily="Segoe UI"
        WindowStyle="None" ResizeMode="NoResize" Topmost="True" UseLayoutRounding="True">
  <Window.Resources>
    <Storyboard x:Key="FundOpenGrowStoryboard">
      <DoubleAnimation Storyboard.TargetName="WallRootScale" Storyboard.TargetProperty="ScaleX"
                       From="0" To="1" Duration="0:0:0.18" FillBehavior="HoldEnd">
        <DoubleAnimation.EasingFunction><QuadraticEase EasingMode="EaseOut"/></DoubleAnimation.EasingFunction>
      </DoubleAnimation>
      <DoubleAnimation Storyboard.TargetName="WallRootScale" Storyboard.TargetProperty="ScaleY"
                       From="0" To="1" Duration="0:0:0.18" FillBehavior="HoldEnd">
        <DoubleAnimation.EasingFunction><QuadraticEase EasingMode="EaseOut"/></DoubleAnimation.EasingFunction>
      </DoubleAnimation>
      <DoubleAnimation Storyboard.TargetName="WallFoundWin" Storyboard.TargetProperty="Opacity"
                       From="0" To="1" Duration="0:0:0.18" FillBehavior="HoldEnd">
        <DoubleAnimation.EasingFunction><QuadraticEase EasingMode="EaseOut"/></DoubleAnimation.EasingFunction>
      </DoubleAnimation>
    </Storyboard>
""" + BIMTOOLS_DARK_STYLES_XML + u"""
  </Window.Resources>
  <Border x:Name="WallRootChrome" CornerRadius="10" Background="#0A1A2F" Padding="12"
          BorderBrush="#1A3A4D" BorderThickness="1" ClipToBounds="True" RenderTransformOrigin="0,0">
    <Border.Effect>
      <DropShadowEffect Color="#000000" BlurRadius="16" ShadowDepth="0" Opacity="0.35"/>
    </Border.Effect>
    <Border.RenderTransform>
      <ScaleTransform x:Name="WallRootScale" ScaleX="0" ScaleY="0"/>
    </Border.RenderTransform>
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>
      <Border x:Name="TitleBar" Grid.Row="0" Background="#0E1B32" CornerRadius="6" Padding="10,8" Margin="0,0,0,10"
              BorderBrush="#21465C" BorderThickness="1" HorizontalAlignment="Stretch">
        <Grid HorizontalAlignment="Stretch">
          <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
          <Image x:Name="ImgLogo" Grid.Column="0" Width="40" Height="40"
                 Stretch="Uniform" Margin="0,0,10,0" VerticalAlignment="Center"/>
          <StackPanel Grid.Column="1" VerticalAlignment="Center" Margin="0,0,8,0">
            <TextBlock Text="Armadura Fundacion Corrida" FontSize="15" FontWeight="SemiBold"
                       Foreground="#E8F4F8" TextWrapping="Wrap"/>
          </StackPanel>
          <Button x:Name="BtnClose" Grid.Column="2" Style="{StaticResource BtnCloseX_MinimalNoBg}" ToolTip="Cerrar"
                  VerticalAlignment="Center" HorizontalAlignment="Right"/>
        </Grid>
      </Border>
      <ScrollViewer x:Name="SvContenido" Grid.Row="1" VerticalScrollBarVisibility="Auto" MaxHeight="620">
        <StackPanel HorizontalAlignment="Stretch">
          <Button x:Name="BtnPick" Content="Seleccionar Fundaciones" Style="{StaticResource BtnSelectOutline}"
                  HorizontalAlignment="Stretch" Margin="0,0,0,8"/>
          <GroupBox Style="{StaticResource GbParams}" Margin="0" HorizontalAlignment="Left">
            <GroupBox.Header>
              <Grid VerticalAlignment="Center">
                <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="Informacion Armaduras" Foreground="#E8F4F8" FontWeight="SemiBold" FontSize="12"
                           VerticalAlignment="Center" HorizontalAlignment="Left"/>
                <ComboBox x:Name="CmbDosificacionHormigon" Grid.Column="1" Style="{StaticResource Combo}" Margin="16,0,0,0"
                          IsEditable="False" IsReadOnly="True" VerticalAlignment="Center" HorizontalAlignment="Right"
                          ToolTip="Dosificación del hormigón">
                  <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                </ComboBox>
              </Grid>
            </GroupBox.Header>
            <StackPanel>
              <!-- Tarjeta interior: título + fila de dos controles (boceto). -->
              <StackPanel x:Name="PanelTrans" Margin="0,0,0,0">
                <Border Background="#0E1B32" CornerRadius="4" Padding="8,8,8,8" BorderBrush="#1A3A4D" BorderThickness="1"
                        Margin="0,0,0,8" HorizontalAlignment="Left">
                  <StackPanel>
                    <TextBlock Text="Transversales" Foreground="#95B8CC" FontWeight="SemiBold" FontSize="11"
                               HorizontalAlignment="Left" Margin="0,0,0,8"/>
                    <Grid HorizontalAlignment="Left">
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="110"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="110"/>
                      </Grid.ColumnDefinitions>
                      <ComboBox Grid.Column="0" x:Name="CmbTransDiam" Style="{StaticResource Combo}" IsEditable="False" IsReadOnly="True">
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
                          <TextBox x:Name="TxtTransSep" Grid.Column="0" Style="{StaticResource CantSpinnerText}"
                                   Text="100" VerticalContentAlignment="Center"
                                   ToolTip="Separación transversal (mm): 100 a 400, paso 10"/>
                          <Border Grid.Column="1" Background="#0E1B32" BorderBrush="#1A3A4D"
                                  BorderThickness="1,0,0,0" CornerRadius="0,4,4,0" ClipToBounds="True">
                            <Grid>
                              <Grid.RowDefinitions>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="*"/>
                              </Grid.RowDefinitions>
                              <RepeatButton x:Name="BtnTransSepUp" Grid.Row="0" Style="{StaticResource SpinRepeatBtn}" Content="▲"
                                            ToolTip="Más 10 mm (máx. 400 mm)"/>
                              <RepeatButton x:Name="BtnTransSepDown" Grid.Row="1" Style="{StaticResource SpinRepeatBtn}" Content="▼"
                                            ToolTip="Menos 10 mm (mín. 100 mm)"/>
                            </Grid>
                          </Border>
                        </Grid>
                      </Border>
                    </Grid>
                  </StackPanel>
                </Border>
              </StackPanel>
              <StackPanel x:Name="PanelLong" Margin="0,0,0,0">
                <Border Background="#0E1B32" CornerRadius="4" Padding="8,8,8,8" BorderBrush="#1A3A4D" BorderThickness="1"
                        HorizontalAlignment="Left">
                  <StackPanel>
                    <TextBlock Text="Longitudinales" Foreground="#95B8CC" FontWeight="SemiBold" FontSize="11"
                               HorizontalAlignment="Left" Margin="0,0,0,8"/>
                    <Grid HorizontalAlignment="Left">
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="110"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="110"/>
                      </Grid.ColumnDefinitions>
                      <ComboBox Grid.Column="0" x:Name="CmbLongDiam" Style="{StaticResource Combo}" IsEditable="False" IsReadOnly="True">
                        <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                      </ComboBox>
                      <TextBlock Grid.Column="1" Text="@" FontSize="12" FontWeight="Bold"
                                 Foreground="#95B8CC" VerticalAlignment="Center" HorizontalAlignment="Center" Margin="6,0,6,0"/>
                      <Border Grid.Column="2" Width="110" Height="24" CornerRadius="4" Background="#050E18"
                              BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True" Margin="0">
                        <Grid>
                          <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="18"/>
                          </Grid.ColumnDefinitions>
                          <TextBox x:Name="TxtLongSep" Grid.Column="0" Style="{StaticResource CantSpinnerText}"
                                   Text="100" VerticalContentAlignment="Center"
                                   ToolTip="Separación longitudinal (mm): 100 a 400, paso 10"/>
                          <Border Grid.Column="1" Background="#0E1B32" BorderBrush="#1A3A4D"
                                  BorderThickness="1,0,0,0" CornerRadius="0,4,4,0" ClipToBounds="True">
                            <Grid>
                              <Grid.RowDefinitions>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="*"/>
                              </Grid.RowDefinitions>
                              <RepeatButton x:Name="BtnLongSepUp" Grid.Row="0" Style="{StaticResource SpinRepeatBtn}" Content="▲"
                                            ToolTip="Más 10 mm (máx. 400 mm)"/>
                              <RepeatButton x:Name="BtnLongSepDown" Grid.Row="1" Style="{StaticResource SpinRepeatBtn}" Content="▼"
                                            ToolTip="Menos 10 mm (mín. 100 mm)"/>
                            </Grid>
                          </Border>
                        </Grid>
                      </Border>
                    </Grid>
                    <Border x:Name="BorderTroceo" Visibility="Collapsed" Margin="0,10,0,0" HorizontalAlignment="Stretch"
                            Background="#050E18" BorderBrush="#1A3A4D" BorderThickness="1" CornerRadius="4" Padding="8,8,8,8">
                      <StackPanel>
                        <TextBlock Text="Troceo longitudinal (&gt; 12 m de eje)" Style="{StaticResource Label}" Margin="0,0,0,4"/>
                        <Grid Margin="0,0,0,6">
                          <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="12"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                          <StackPanel Grid.Column="0">
                            <TextBlock Text="Largo máx. tramo (mm)" Foreground="#95B8CC" FontSize="10"/>
                            <TextBox x:Name="TxtMaxBarMm" Style="{StaticResource CantSpinnerText}" Background="#050E18"
                                     BorderBrush="#1A3A4D" BorderThickness="1" Padding="4" Text="12000"/>
                          </StackPanel>
                          <StackPanel Grid.Column="2">
                            <TextBlock Text="Empalme / traslape (mm)" Foreground="#95B8CC" FontSize="10"/>
                            <TextBox x:Name="TxtLapMm" Style="{StaticResource CantSpinnerText}" Background="#050E18"
                                     BorderBrush="#1A3A4D" BorderThickness="1" Padding="4" Text="600"/>
                          </StackPanel>
                        </Grid>
                      </StackPanel>
                    </Border>
                  </StackPanel>
                </Border>
              </StackPanel>
            </StackPanel>
          </GroupBox>
        </StackPanel>
      </ScrollViewer>
      <StackPanel Grid.Row="2" Margin="0" HorizontalAlignment="Stretch">
        <Border Height="4" Background="Transparent"/>
        <Button x:Name="BtnColocar" Content="Colocar Armadura" Style="{StaticResource BtnPrimary}" HorizontalAlignment="Stretch"/>
      </StackPanel>
    </Grid>
  </Border>
</Window>
"""


def _wf_apply_selection_uidoc(uidoc, wf_id):
    """
    Deja la Wall Foundation como única selección de Revit (resaltado estándar)
    hasta que se limpie explícitamente (p. ej. tras Colocar).
    """
    if uidoc is None or wf_id is None:
        return
    try:
        from System.Collections.Generic import List

        doc = uidoc.Document
        if doc is None:
            return
        el = doc.GetElement(wf_id)
        if el is None or not isinstance(el, WallFoundation):
            return
        uidoc.Selection.SetElementIds(List[ElementId]([wf_id]))
    except Exception:
        pass


class WallFoundationOnlyFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            return elem is not None and isinstance(elem, WallFoundation)
        except Exception:
            return False

    def AllowReference(self, ref, pt):
        return False


class PickWallFoundationHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        from Autodesk.Revit.UI.Selection import ObjectType

        win = self._window_ref()
        if win is None:
            return
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            _task_dialog_show(
                u"BIMTools — Wall Foundation Reinforcement",
                u"No hay documento activo.",
                win._win,
            )
            return
        doc = uidoc.Document
        flt = WallFoundationOnlyFilter()
        try:
            r = uidoc.Selection.PickObject(
                ObjectType.Element,
                flt,
                u"Seleccione una zapata de muro (Wall Foundation).",
            )
        except Exception:
            win._show_after_pick()
            return
        if r is None:
            win._show_after_pick()
            return
        win._document = doc
        win._wf_id = r.ElementId
        _wf_apply_selection_uidoc(uidoc, r.ElementId)
        win._refresh_troceo_panel()
        win._show_after_pick()

    def GetName(self):
        return u"PickWallFoundation"


class ReselectWallFoundationHandler(IExternalEventHandler):
    """
    Restaura la selección de la zapata pendiente al volver al formulario modeless,
    si Revit la soltó al cambiar el foco.
    """

    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        win = self._window_ref()
        if win is None:
            return
        wf_id = getattr(win, "_wf_id", None)
        if wf_id is None:
            return
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            return
        try:
            cur_ids = list(uidoc.Selection.GetElementIds())
            if len(cur_ids) == 1:
                only = cur_ids[0]
                try:
                    if only == wf_id:
                        return
                except Exception:
                    try:
                        if element_id_to_int(only) == element_id_to_int(wf_id):
                            return
                    except Exception:
                        pass
        except Exception:
            pass
        _wf_apply_selection_uidoc(uidoc, wf_id)

    def GetName(self):
        return u"ReseleccionarWallFoundationPendiente"


class EliminarSeccionesRevisionWallFoundationHandler(IExternalEventHandler):
    """
    Borra vistas de sección de revisión tras cerrar el WPF modeless (Revit no permite
    transacción en el evento Closed).
    """

    def __init__(self):
        self._pending_doc = None
        self._pending_ids = None

    def armar(self, document, ids):
        self._pending_doc = document
        self._pending_ids = list(ids) if ids else []

    def Execute(self, uiapp):
        doc = self._pending_doc
        ids = self._pending_ids
        self._pending_doc = None
        self._pending_ids = None
        if doc is None:
            try:
                uidoc = uiapp.ActiveUIDocument
                if uidoc is not None:
                    doc = uidoc.Document
            except Exception:
                pass
        if doc is None:
            return
        try:
            from vista_seccion_enfierrado_vigas import (
                eliminar_vistas_seccion_revision_enfierrado,
            )

            eliminar_vistas_seccion_revision_enfierrado(
                doc,
                ids or [],
                uidocument=uiapp.ActiveUIDocument,
            )
        except Exception:
            pass

    def GetName(self):
        return u"EliminarSeccionesRevisionWallFoundation"


def _colocar_rebar_en_host(
    doc,
    wf,
    bt_long,
    long_sep_mm,
    long_line,
    width_line,
    usable_w_ft,
    needs_lap,
    max_bar_mm,
    lap_mm,
    avisos,
    geo=None,
    rebars_out=None,
    concrete_grade=None,
    active_view=None,
):
    """
    Longitudinales sobre ``long_line`` (3D) ya situada como eje de barra inferior.
    Ganchos como tramos de polilínea (``CreateFromCurves*``, misma familia que la U), no con
    sobrecarga de ``RebarHookType`` en línea recta.
    Troceo solo sobre el eje recto sintético (sin curvas con estiramiento empotramiento).
    """
    # Recorte en extremos ya aplicado en geometría (tangente 50 mm + ø/2 al eje).
    rec_e = 0.0
    p0 = long_line.GetEndPoint(0)
    p1 = long_line.GetEndPoint(1)
    try:
        tang = p1.Subtract(p0)
        L_ax = float(tang.GetLength())
        if L_ax < 1e-9:
            return 0
        tu = tang.Normalize()
    except Exception:
        return 0
    p0i, p1i = _ubicar_punto_eje_menos_recorte(p0, p1, tu, rec_e)
    if p0i is None or p1i is None:
        return 0
    z_al = 0.5 * (float(p0i.Z) + float(p1i.Z))
    pm_ref = _wf_punto_referencia_planta_wall_foundation(wf, z_al)
    if pm_ref is not None:
        _axis_align = _wf_traslada_linea_hacia_punto_planta_xy(
            Line.CreateBound(p0i, p1i), float(pm_ref.X), float(pm_ref.Y)
        )
    else:
        _axis_align = _wf_traslada_linea_hacia_centro_bbox_planta(
            Line.CreateBound(p0i, p1i), wf
        )
    if _axis_align is not None:
        try:
            p0i = _axis_align.GetEndPoint(0)
            p1i = _axis_align.GetEndPoint(1)
        except Exception:
            pass
    usable_w = max(0.0, float(usable_w_ft))
    if usable_w < _mm_to_ft(10.0):
        avisos.append(u"Ancho útil casi nulo; revise recubrimientos.")
        return 0
    n_cara = None
    marco_uvn = None
    cara_pp = None
    if geo is not None:
        n_cara = geo.get("n_cara")
        marco_uvn = geo.get("marco_uvn")
        cara_pp = geo.get("cara_pp")
        if n_cara is None and marco_uvn is not None and len(marco_uvn) > 3:
            n_cara = marco_uvn[3]
    if n_cara is None:
        n_cara = XYZ.BasisZ.Negate()
    z0 = z1 = None
    if geo is not None:
        z0 = geo.get("z0")
        z1 = geo.get("z1")
    if z0 is None or z1 is None:
        z0, z1 = _wf_z_range_ft(wf)
    cap_geom_ft = longitud_pata_u_fundacion_inf_sup_ft(z0, z1, _DESC_PATA_U_MM)
    if cap_geom_ft is None or cap_geom_ft < 1e-9:
        avisos.append(u"No se pudo resolver la altura de patas (longitudinales).")
        return 0
    gancho_tab_mm = None
    d_long_mm = _rebar_nominal_diameter_mm(bt_long)
    if d_long_mm is None:
        leg_ft = cap_geom_ft
    else:
        gancho_tab_mm = largo_gancho_u_tabla_mm(d_long_mm, concrete_grade)
        if gancho_tab_mm is None:
            leg_ft = cap_geom_ft
        else:
            eje_mm = pata_eje_curve_loop_mm_desde_tabla_mm(
                gancho_tab_mm, float(d_long_mm)
            )
            leg_ft = min(cap_geom_ft, _mm_to_ft(eje_mm))
    if leg_ft is None or leg_ft < 1e-9:
        avisos.append(u"No se pudo resolver la altura de patas (longitudinales).")
        return 0
    axis_seg = Line.CreateBound(p0i, p1i)
    if needs_lap:
        if gancho_tab_mm is not None:
            p_stock_mm = float(gancho_tab_mm)
        elif d_long_mm is not None:
            _gt = largo_gancho_u_tabla_mm(d_long_mm, concrete_grade)
            p_stock_mm = float(_gt) if _gt is not None else max(100.0, _ft_to_mm(leg_ft))
        else:
            p_stock_mm = max(100.0, _ft_to_mm(leg_ft))
        segs = _split_line_laps_longitudinal_eje_stock(
            axis_seg, float(max_bar_mm), p_stock_mm, float(lap_mm)
        )
    else:
        segs = [axis_seg]
    if not segs:
        return 0
    lap_infos = []
    if needs_lap and len(segs) > 1:
        n_sym, lap_infos = _wf_colocar_simbologia_empalme_eje(
            doc, active_view, segs, avisos, wf=wf, geo=geo
        )
        if n_sym > 0:
            try:
                avisos.append(
                    u"Simbología de empalme (longitudinales): {0} detalle(s).".format(
                        int(n_sym)
                    )
                )
            except Exception:
                pass
    n_tot = 0
    nseg = len(segs)
    rebar_id_per_seg = [None] * nseg
    for i, seg in enumerate(segs):
        try:
            ln = Line.CreateBound(seg.GetEndPoint(0), seg.GetEndPoint(1))
        except Exception:
            continue
        g0 = i == 0
        g1 = i == nseg - 1
        norm_pri, w_unit = _wf_norm_distribucion_longitudinal_en_planta(ln, width_line)
        if norm_pri is None:
            norm_pri = [XYZ.BasisZ]
            w_unit = None
        array_len_ft = _wf_span_luz_distribucion_bbox_ft(
            wf, ln, None, usable_w
        )
        # La primera barra del set queda fija; el conjunto crece ~arrayLen en planta. Si la curva
        # está en el centro del ancho, la mitad queda fuera del host → retroceder media luz.
        if w_unit is not None and float(array_len_ft) > _mm_to_ft(15.0):
            try:
                sh = w_unit.Multiply(float(array_len_ft) * 0.5)
                ln = Line.CreateBound(
                    ln.GetEndPoint(0).Subtract(sh),
                    ln.GetEndPoint(1).Subtract(sh),
                )
            except Exception:
                pass
        try:
            z_hook_ref = vector_reverso_cara_paralela_mas_cercana_a_barra(wf, ln)
        except Exception:
            z_hook_ref = XYZ.BasisZ
        if z_hook_ref is None:
            z_hook_ref = XYZ.BasisZ
        tramos, ln_ref = construir_polilinea_fundacion_ganchos_geometricos_desde_eje(
            ln,
            n_cara,
            leg_ft,
            g0,
            g1,
            d_long_mm,
            acortar_eje_central_para_cota_revit=False,
        )
        if tramos is None or ln_ref is None:
            avisos.append(
                u"Longitudinal tramo {0}: no se construyó la polilínea.".format(i + 1)
            )
            continue
        try:
            dev_mm = sum(_ft_to_mm(float(x.Length)) for x in tramos)
        except Exception:
            dev_mm = float(_MAX_STOCK_MM) + 1.0
        try:
            lim_stock_mm = float(max_bar_mm)
        except Exception:
            lim_stock_mm = float(_MAX_STOCK_MM)
        if dev_mm > lim_stock_mm + 5.0:
            avisos.append(
                u"Longitudinal tramo {0}: desarrollado ~{1:.0f} mm > límite ~{2:.0f} mm; no se crea.".format(
                    i + 1, dev_mm, lim_stock_mm
                )
            )
            continue
        r, err, _nv = None, None, None
        if len(tramos) == 3:
            poli = (tramos[0], tramos[1], tramos[2])
            r, err, _nv = crear_rebar_u_shape_desde_eje_rebar_shape_nombrado(
                doc,
                wf,
                bt_long,
                poli,
                shape_nombre=REBAR_SHAPE_NOMBRE_DEFECTO,
                marco_cara_uvn=marco_uvn,
                cara_paralela=cara_pp,
                eje_referencia_z_ganchos=z_hook_ref,
                normales_prioridad=norm_pri,
            )
            if r is None:
                r2, err2, _nv2 = crear_rebar_polilinea_u_malla_inf_sup_curve_loop(
                    doc,
                    wf,
                    bt_long,
                    poli,
                    poli[1],
                    marco_cara_uvn=marco_uvn,
                    cara_paralela=cara_pp,
                    eje_referencia_z_ganchos=z_hook_ref,
                    normales_prioridad=norm_pri,
                )
                if r2 is None:
                    r3, err3, _nv3 = crear_rebar_polilinea_recta_sin_ganchos(
                        doc,
                        wf,
                        bt_long,
                        poli,
                        poli[1],
                        marco_cara_uvn=marco_uvn,
                        cara_paralela=cara_pp,
                        eje_referencia_z_ganchos=z_hook_ref,
                        normales_prioridad=norm_pri,
                    )
                    r = r3
                    err = err3 or err2 or err
                    _nv = _nv3 or _nv2 or _nv
                else:
                    r = r2
                    err = err2
                    _nv = _nv2
        else:
            r, err, _nv = crear_rebar_polilinea_recta_sin_ganchos(
                doc,
                wf,
                bt_long,
                tramos,
                ln_ref,
                marco_cara_uvn=marco_uvn,
                cara_paralela=cara_pp,
                eje_referencia_z_ganchos=z_hook_ref,
                normales_prioridad=norm_pri,
            )
        if r is None:
            avisos.append(u"Longitudinal tramo {0}: {1}".format(i + 1, err or u"fallo"))
            continue
        try:
            rebar_id_per_seg[i] = r.Id
        except Exception:
            pass
        ok_l, wlay = aplicar_layout_maximum_spacing_rebar(
            r, doc, long_sep_mm, array_len_ft, flip_rebar_set=False
        )
        if not ok_l:
            avisos.append(
                u"Longitudinal tramo {0}: layout: {1}".format(i + 1, wlay or u"")
            )
        if rebars_out is not None:
            try:
                rebars_out.append(r)
            except Exception:
                pass
        try:
            n_tot += int(r.Quantity)
        except Exception:
            n_tot += 1
    if lap_infos:
        try:
            from lap_detail_link_wall_foundation_schema import (
                set_wall_foundation_lap_detail_rebar_link,
            )

            for info in lap_infos:
                try:
                    j = int(info.get("joint_idx", -1))
                except Exception:
                    continue
                inst = info.get("inst")
                dim_id = info.get("dim_id")
                if j < 0 or inst is None:
                    continue
                if j + 1 >= len(rebar_id_per_seg):
                    continue
                ra_id = rebar_id_per_seg[j]
                rb_id = rebar_id_per_seg[j + 1]
                if ra_id is None or rb_id is None:
                    continue
                set_wall_foundation_lap_detail_rebar_link(
                    inst, ra_id, rb_id, dim_id
                )
        except Exception:
            pass
    return n_tot


def _colocar_trans_u(
    doc, wf, bt_tr, trans_sep_mm, geo, avisos, rebars_out=None, concrete_grade=None
):
    long_line = geo.get("long_line")
    width_line = geo.get("width_line")
    marco_uvn = geo.get("marco_uvn")
    cara_pp = geo.get("cara_pp")
    z0 = geo.get("z0")
    z1 = geo.get("z1")
    if long_line is None or width_line is None:
        return 0
    rec_e = _mm_to_ft(_RECO_EXT_EJE_MM)
    p0 = long_line.GetEndPoint(0)
    p1 = long_line.GetEndPoint(1)
    try:
        tang = p1.Subtract(p0)
        L_ax = float(tang.GetLength())
        tu = tang.Normalize()
    except Exception:
        return 0
    if L_ax < 1e-9:
        return 0
    norm_u_create = None
    try:
        tu_xy = XYZ(float(tu.X), float(tu.Y), 0.0)
        if float(tu_xy.GetLength()) < 1e-12:
            return 0
        tu_xy = tu_xy.Normalize()
        wplan = _wf_perp_horizontal_xy(tu_xy)
        wplan = _wf_alinea_ancho_con_curva_ancho(wplan, width_line)
        wdir = wplan
        # ``norm`` de CreateFromCurvesAndShape: plano ⟂ al eje del muro (= ``tu_xy`` en planta);
        # sentido opuesto al que se usaba antes (``Negate``).
        try:
            norm_u_create = [tu_xy]
        except Exception:
            norm_u_create = None
    except Exception:
        return 0
    usable_w = max(0.0, float(width_line.Length) - 2.0 * rec_e)
    if usable_w < _mm_to_ft(10.0):
        avisos.append(u"Ancho insuficiente para U transversal.")
        return 0
    if z0 is None or z1 is None:
        z0, z1 = _wf_z_range_ft(wf)
    cap_geom_ft = longitud_pata_u_fundacion_inf_sup_ft(z0, z1, _DESC_PATA_U_MM)
    if cap_geom_ft is None or cap_geom_ft < 1e-9:
        avisos.append(u"No se pudo resolver la altura de patas (U).")
        return 0
    d_tr_mm = _rebar_nominal_diameter_mm(bt_tr)
    gancho_tab_mm = None
    if d_tr_mm is None:
        leg_ft = cap_geom_ft
    else:
        gancho_tab_mm = largo_gancho_u_tabla_mm(d_tr_mm, concrete_grade)
        if gancho_tab_mm is None:
            leg_ft = cap_geom_ft
        else:
            eje_mm = pata_eje_curve_loop_mm_desde_tabla_mm(
                gancho_tab_mm, float(d_tr_mm)
            )
            leg_ft = min(cap_geom_ft, _mm_to_ft(eje_mm))
    if leg_ft is None or leg_ft < 1e-9:
        avisos.append(u"No se pudo resolver la altura de patas (U).")
        return 0
    span_w_ft = _wf_span_luz_distribucion_bbox_ft(wf, long_line, None, usable_w)
    usable_l_ax = max(0.0, L_ax - 2.0 * rec_e)
    array_len = _wf_span_luz_along_eje_wall_desde_perimetro_ft(wf, long_line, usable_l_ax)
    # Plano vertical de la U = capa **transversal** (``width_line`` usa ø trans. en
    # ``offset_linea_eje_barra_desde_cara_inferior_mm``). ``long_line`` usa ø long.:
    # si mezclamos cotas se asumía implícitamente el mismo diámetro.
    try:
        zmid = 0.5 * (
            float(width_line.GetEndPoint(0).Z)
            + float(width_line.GetEndPoint(1).Z)
        )
    except Exception:
        zmid = 0.5 * (
            float(long_line.GetEndPoint(0).Z) + float(long_line.GetEndPoint(1).Z)
        )
    c_bb = None
    if wf is not None:
        c_bb = _wf_punto_referencia_planta_wall_foundation(wf, zmid)
    if c_bb is None:
        try:
            gc = centro_xy_perimetro_inferior_doc(wf) if wf is not None else None
            if gc is not None:
                c_bb = XYZ(float(gc[0]), float(gc[1]), zmid)
        except Exception:
            c_bb = None
    if c_bb is None:
        bbu = wf.get_BoundingBox(None) if wf is not None else None
        if bbu is not None:
            c_bb = XYZ(
                0.5 * (float(bbu.Min.X) + float(bbu.Max.X)),
                0.5 * (float(bbu.Min.Y) + float(bbu.Max.Y)),
                zmid,
            )
    if c_bb is None:
        try:
            pm = long_line.Evaluate(0.5, True)
            c_bb = XYZ(float(pm.X), float(pm.Y), zmid)
        except Exception:
            avisos.append(u"Transversal U: sin punto central.")
            return 0
    p_cen = _wf_punto_centro_u_en_franja(long_line, wdir, zmid, c_bb)
    if p_cen is None:
        p_cen = c_bb
    tu_h = XYZ(float(tu.X), float(tu.Y), 0.0)
    if float(tu_h.GetLength()) > 1e-12:
        tu_h = tu_h.Normalize()
        if float(array_len) > _mm_to_ft(15.0):
            try:
                p_cen = p_cen.Subtract(tu_h.Multiply(float(array_len) * 0.5))
            except Exception:
                pass
    half = 0.5 * float(span_w_ft)
    # Geometría común: ancho útil coherente con ~100 mm de offset de perímetro en planta.
    # Luego se acerca la U a la cara lateral hasta situar el **eje** a
    # ``_REC_LATERAL_CARA_U_MM + ø/2`` de la cara (recubrimiento medido a tangente de barra).
    try:
        _du_lat = max(
            0.0, float(_REC_OFF_PLANTA_INF_MM) - float(_REC_LATERAL_CARA_U_MM)
        )
        half = half + _mm_to_ft(_du_lat)
    except Exception:
        pass
    try:
        if d_tr_mm is not None:
            r_tr = 0.5 * float(d_tr_mm)
            if r_tr > 1e-6:
                half = half - _mm_to_ft(r_tr)
    except Exception:
        pass
    pa = p_cen.Subtract(wdir.Multiply(half))
    pb = p_cen.Add(wdir.Multiply(half))
    linea_eje = Line.CreateBound(pa, pb)
    try:
        leg_stock_mm = (
            float(gancho_tab_mm)
            if gancho_tab_mm is not None
            else float(_ft_to_mm(leg_ft))
        )
        u_len_mm = _ft_to_mm(float(linea_eje.Length)) + 2.0 * leg_stock_mm
    except Exception:
        u_len_mm = float(_MAX_STOCK_MM) + 1.0
    if u_len_mm > _MAX_STOCK_MM + 0.01:
        avisos.append(
            u"U transversal: desarrollado ~{0:.0f} mm > 12 m; no se crea.".format(
                u_len_mm
            )
        )
        return 0
    n_cara = geo.get("n_cara")
    if n_cara is None and marco_uvn is not None and len(marco_uvn) > 3:
        n_cara = marco_uvn[3]
    if n_cara is None:
        n_cara = XYZ.BasisZ.Negate()
    try:
        z_hook = vector_reverso_cara_paralela_mas_cercana_a_barra(wf, linea_eje)
    except Exception:
        z_hook = XYZ.BasisZ
    if z_hook is None:
        z_hook = XYZ.BasisZ
    poli = construir_polilinea_u_fundacion_desde_eje_horizontal(
        linea_eje,
        n_cara,
        leg_ft,
        d_tr_mm,
        acortar_eje_central_para_cota_revit=False,
    )
    if poli is None:
        avisos.append(u"No se construyó la polilínea U.")
        return 0
    r, err, _nv = crear_rebar_u_shape_desde_eje_rebar_shape_nombrado(
        doc,
        wf,
        bt_tr,
        poli,
        shape_nombre=REBAR_SHAPE_NOMBRE_DEFECTO,
        marco_cara_uvn=None,
        cara_paralela=None,
        eje_referencia_z_ganchos=z_hook,
        normales_prioridad=norm_u_create,
    )
    if r is None:
        r2, err2, _ = crear_rebar_polilinea_u_malla_inf_sup_curve_loop(
            doc,
            wf,
            bt_tr,
            poli,
            poli[1],
            marco_cara_uvn=None,
            cara_paralela=None,
            eje_referencia_z_ganchos=z_hook,
            normales_prioridad=norm_u_create,
        )
        if r2 is None:
            r3, err3, _ = crear_rebar_polilinea_recta_sin_ganchos(
                doc,
                wf,
                bt_tr,
                poli,
                poli[1],
                marco_cara_uvn=None,
                cara_paralela=None,
                eje_referencia_z_ganchos=z_hook,
                normales_prioridad=norm_u_create,
            )
            r = r3
            err = err3 or err2 or err
        else:
            r = r2
            err = err2
    if r is None:
        avisos.append(u"Transversal U: {0}".format(err or u"error"))
        return 0
    ok_l, wlay = aplicar_layout_maximum_spacing_rebar(
        r, doc, trans_sep_mm, array_len, flip_rebar_set=False
    )
    if not ok_l:
        avisos.append(u"Transversal: Maximum Spacing: {0}".format(wlay or u""))
    if rebars_out is not None:
        try:
            rebars_out.append(r)
        except Exception:
            pass
    try:
        return int(r.Quantity)
    except Exception:
        return 1


class ColocarWallFoundationHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        win = self._window_ref()
        if win is None:
            return
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            _task_dialog_show(
                u"BIMTools — Wall Foundation Reinforcement",
                u"No hay documento activo.",
                win._win,
            )
            return
        doc = uidoc.Document
        wf_id = getattr(win, "_wf_id", None)
        if wf_id is None:
            _task_dialog_show(
                u"BIMTools — Wall Foundation Reinforcement",
                u"Seleccione una Wall Foundation.",
                win._win,
            )
            return
        wf = doc.GetElement(wf_id)
        if wf is None or not isinstance(wf, WallFoundation):
            _task_dialog_show(
                u"BIMTools — Wall Foundation Reinforcement",
                u"Elemento inválido o no es Wall Foundation.",
                win._win,
            )
            return
        do_t = True
        do_l = True
        entries = getattr(win, "_entries", None) or []
        bt_tr, e1 = _resolver_bar_type_from_combo(
            doc, win._win.FindName("CmbTransDiam"), entries
        )
        bt_lo, e2 = _resolver_bar_type_from_combo(
            doc, win._win.FindName("CmbLongDiam"), entries
        )
        if do_t and bt_tr is None:
            _task_dialog_show(
                u"BIMTools — Wall Foundation Reinforcement",
                e1 or u"Tipo de barra transversal no válido.",
                win._win,
            )
            return
        if do_l and bt_lo is None:
            _task_dialog_show(
                u"BIMTools — Wall Foundation Reinforcement",
                e2 or u"Tipo de barra longitudinal no válido.",
                win._win,
            )
            return
        trans_sep = _read_sep_tb(win._win.FindName("TxtTransSep"))
        long_sep = _read_sep_tb(win._win.FindName("TxtLongSep"))
        try:
            win._dosificacion_hormigon = _read_dosificacion_hormigon(
                win._win.FindName("CmbDosificacionHormigon")
            )
        except Exception:
            win._dosificacion_hormigon = _DOSIFICACION_HORMIGON_DEFAULT
        joined_ids = _wf_collect_joined_element_ids(doc, wf)
        d_long_mm = _rebar_nominal_diameter_mm(bt_lo) if bt_lo else 0.0
        d_tr_mm = _rebar_nominal_diameter_mm(bt_tr) if bt_tr else 0.0
        tlap_ctrl = win._win.FindName("TxtLapMm")
        L_hint_mm = _longitud_eje_hint_mm(wf)
        if L_hint_mm < 1.0:
            _task_dialog_show(
                u"BIMTools — Wall Foundation Reinforcement",
                u"No se pudo estimar la longitud de la zapata (geometría o caja).",
                win._win,
            )
            return
        needs_lap_hint = float(L_hint_mm) > _MAX_STOCK_MM + 0.01
        if needs_lap_hint:
            max_pre = _read_max_bar_tb(win._win.FindName("TxtMaxBarMm"))
            lap_pre = (
                _wf_traslape_mm_longitudinal(
                    d_long_mm, tlap_ctrl, win._dosificacion_hormigon
                )
                if do_l
                else _read_lap_tb(tlap_ctrl)
            )
            if max_pre <= lap_pre + 1.0:
                _task_dialog_show(
                    u"BIMTools — Wall Foundation Reinforcement",
                    u"El largo máximo por tramo debe ser mayor que el empalme.",
                    win._win,
                )
                return
        avisos = []
        rebars_sets_creados = []
        active_view = uidoc.ActiveView
        n_t = n_l = 0
        linea_para_seccion = None
        t = Transaction(doc, u"BIMTools — Wall Foundation Reinforcement")
        t.Start()
        try:
            _wf_unjoin_all(doc, wf, joined_ids, avisos)
            try:
                doc.Regenerate()
            except Exception:
                pass
            wf = doc.GetElement(wf_id)
            if wf is None or not isinstance(wf, WallFoundation):
                raise Exception(u"La zapata dejó de ser válida tras desunir geometría.")
            geo, geo_hint = _geometria_wall_foundation_inferior(wf, d_long_mm, d_tr_mm)
            if geo is None:
                raise Exception(geo_hint or u"No se resolvió la geometría de la zapata.")
            if geo_hint:
                avisos.append(geo_hint)
            _wf_geo_alinear_strip_a_location_wall_foundation(
                wf, geo, d_long_mm, d_tr_mm
            )
            try:
                _ln_s = geo.get("long_line")
                if _ln_s is not None:
                    linea_para_seccion = Line.CreateBound(
                        _ln_s.GetEndPoint(0), _ln_s.GetEndPoint(1)
                    )
            except Exception:
                linea_para_seccion = None
            if linea_para_seccion is None and wf is not None:
                try:
                    loc = wf.Location
                    if isinstance(loc, LocationCurve) and loc.Curve is not None:
                        ax = _wf_location_curve_como_linea(loc.Curve)
                        if ax is not None:
                            linea_para_seccion = Line.CreateBound(
                                ax.GetEndPoint(0), ax.GetEndPoint(1)
                            )
                except Exception:
                    pass
            L_mm_act = _ft_to_mm(float(geo["long_line"].Length))
            needs_lap = float(L_mm_act) > _MAX_STOCK_MM + 0.01
            if needs_lap:
                max_mm = _read_max_bar_tb(win._win.FindName("TxtMaxBarMm"))
                lap_mm = (
                    _wf_traslape_mm_longitudinal(
                        d_long_mm, tlap_ctrl, win._dosificacion_hormigon
                    )
                    if do_l
                    else _read_lap_tb(tlap_ctrl)
                )
                if max_mm <= lap_mm + 1.0:
                    raise Exception(
                        u"El largo máximo por tramo debe ser mayor que el empalme."
                    )
            else:
                max_mm = float(_MAX_STOCK_MM)
                lap_mm = float(_LAP_DEFAULT_MM)
            if do_t:
                n_t = _colocar_trans_u(
                    doc,
                    wf,
                    bt_tr,
                    trans_sep,
                    geo,
                    avisos,
                    rebars_out=rebars_sets_creados,
                    concrete_grade=win._dosificacion_hormigon,
                )
            if do_l:
                long_axis = geo["long_line"]
                # Desde el eje ya desplazado con ø long. hasta el eje de la segunda capa
                # (long. sobre trans.): ``Δ = d_t`` hacia interior (p. ej. ø10/ø8 → 10 mm).
                if do_t and d_tr_mm:
                    try:
                        dt = float(d_tr_mm)
                    except Exception:
                        dt = 0.0
                    if dt > 1e-6:
                        long_axis = _wf_traslada_linea_hacia_interior_hormigon_mm(
                            long_axis,
                            geo.get("n_cara"),
                            dt,
                        )
                        if long_axis is None:
                            long_axis = geo["long_line"]
                n_l = _colocar_rebar_en_host(
                    doc,
                    wf,
                    bt_lo,
                    long_sep,
                    long_axis,
                    geo["width_line"],
                    geo["usable_w_ft"],
                    needs_lap,
                    max_mm,
                    lap_mm,
                    avisos,
                    geo=geo,
                    rebars_out=rebars_sets_creados,
                    concrete_grade=win._dosificacion_hormigon,
                    active_view=active_view,
                )
            _wf_rejoin_all(doc, wf, joined_ids, avisos)
            try:
                doc.Regenerate()
            except Exception:
                pass
            # Etiquetas **antes** de reducir el conjunto a «solo barra central» en planta (la API
            # de etiquetado suele requerir el conjunto completo visible).
            if rebars_sets_creados:
                try:
                    doc.Regenerate()
                except Exception:
                    pass
                _n_tags = _wf_etiquetar_rebar_sets_independent_tag(
                    doc,
                    active_view,
                    rebars_sets_creados,
                    avisos,
                )
                if _n_tags > 0:
                    try:
                        avisos.append(
                            u"Etiquetas «{0}» (tipo = RebarShape): {1} creada(s).".format(
                                _WF_REBAR_TAG_FAMILY_NAME, int(_n_tags)
                            )
                        )
                    except Exception:
                        pass
                try:
                    from geometria_estribos_viga import (
                        crear_multi_rebar_annotations_por_nombre_tipo,
                    )

                    _n_mra = crear_multi_rebar_annotations_por_nombre_tipo(
                        doc,
                        active_view,
                        rebars_sets_creados,
                        avisos,
                        _WF_MULTI_REBAR_ANNOTATION_TYPE_NAME,
                    )
                    if _n_mra > 0:
                        try:
                            avisos.append(
                                u"Multi-Rebar Annotation «{0}»: {1} creada(s).".format(
                                    _WF_MULTI_REBAR_ANNOTATION_TYPE_NAME,
                                    int(_n_mra),
                                )
                            )
                        except Exception:
                            pass
                except Exception:
                    pass
            if (
                _wf_vista_es_planta(active_view)
                and rebars_sets_creados
            ):
                _wf_aplicar_presentacion_solo_barra_central_planta(
                    active_view, rebars_sets_creados
                )
                try:
                    avisos.append(
                        u"Vista en planta: cada conjunto muestra solo la barra central."
                    )
                except Exception:
                    pass
            t.Commit()
        except Exception as ex:
            t.RollBack()
            _task_dialog_show(
                u"BIMTools — Wall Foundation Reinforcement",
                u"Error (se revirtió la transacción):\n{0}".format(ex),
                win._win,
            )
            return
        # Por defecto True: si el import falla (pyRevit / path), no bloquear silenciosamente.
        _crear_seccion_revision = True
        try:
            from vista_seccion_enfierrado_vigas import CREAR_SECCION_REVISION_WALL_FOUNDATION

            _crear_seccion_revision = bool(CREAR_SECCION_REVISION_WALL_FOUNDATION)
        except Exception:
            pass
        if _crear_seccion_revision and (n_t > 0 or n_l > 0):
            tsec = Transaction(doc, u"BIMTools — Sección revisión Wall Foundation")
            tsec.Start()
            try:
                from vista_seccion_enfierrado_vigas import (
                    crear_vistas_seccion_revision_wall_foundation,
                )

                wf_sec = doc.GetElement(wf_id)
                _vsec = []
                if wf_sec is None:
                    avisos.append(
                        u"Sección revisión: no se encontró la Wall Foundation en el documento."
                    )
                else:
                    _vsec, _av_sec = crear_vistas_seccion_revision_wall_foundation(
                        doc,
                        [wf_sec],
                        linea_eje=linea_para_seccion,
                        gestionar_transaccion=False,
                        uidocument=uidoc,
                    )
                    avisos.extend(_av_sec or [])
                    if _vsec:
                        try:
                            for _vs in _vsec:
                                if _vs is not None:
                                    win._secciones_revision_ids.append(_vs.Id)
                        except Exception:
                            pass
                        try:
                            nombres_vsec = []
                            for _vs in _vsec:
                                if _vs is not None:
                                    nombres_vsec.append(unicode(_vs.Name))
                            if nombres_vsec:
                                avisos.append(
                                    u"Vista(s) de sección (revisión armadura): "
                                    + u"; ".join(nombres_vsec)
                                    + u"."
                                )
                        except Exception:
                            pass
                    else:
                        if not (_av_sec or []):
                            avisos.append(
                                u"Sección revisión: no se creó ninguna vista. "
                                u"Compruebe que exista un tipo «Section» en el proyecto."
                            )
                tsec.Commit()
            except Exception as ex_sec:
                try:
                    tsec.RollBack()
                except Exception:
                    pass
                try:
                    avisos.append(
                        u"Sección revisión: {0}".format(unicode(ex_sec))
                    )
                except Exception:
                    avisos.append(u"Sección revisión: error inesperado.")
        elif (n_t > 0 or n_l > 0) and not _crear_seccion_revision:
            avisos.append(
                u"Sección revisión: desactivada (CREAR_SECCION_REVISION_WALL_FOUNDATION = False)."
            )
        try:
            from System.Collections.Generic import List

            uidoc.Selection.SetElementIds(List[ElementId]())
        except Exception:
            pass
        if avisos:
            try:
                txt = u"\n".join(avisos)
                if len(txt) > 5000:
                    txt = txt[:4900] + u"\n…"
                _task_dialog_show(
                    u"BIMTools — Wall Foundation — Resultado",
                    txt,
                    win._win,
                )
            except Exception:
                pass
        win._wf_id = None
        try:
            win._refresh_troceo_panel()
        except Exception:
            pass

    def GetName(self):
        return u"ColocarWallFoundationRebar"


class WallFoundationReinforcementWindow(object):
    def __init__(self, revit):
        self._revit = revit
        self._document = None
        self._wf_id = None
        self._entries = []
        self._is_closing_with_fade = False
        self._open_grow_storyboard_started = False
        self._base_top = None
        self._form_width_px = float(_fund_form_width_px())

        from System.Windows import RoutedEventHandler
        from System.Windows.Input import ApplicationCommands, CommandBinding, Key, KeyBinding, ModifierKeys
        from System.Windows.Markup import XamlReader

        self._win = XamlReader.Parse(_WF_XAML)
        self._win.Width = self._form_width_px
        self._win.MinWidth = self._form_width_px
        self._win.MaxWidth = self._form_width_px

        self._pick_handler = PickWallFoundationHandler(weakref.ref(self))
        self._pick_event = ExternalEvent.Create(self._pick_handler)
        self._col_handler = ColocarWallFoundationHandler(weakref.ref(self))
        self._col_event = ExternalEvent.Create(self._col_handler)
        self._reselect_handler = ReselectWallFoundationHandler(weakref.ref(self))
        self._reselect_event = ExternalEvent.Create(self._reselect_handler)
        self._secciones_revision_ids = []
        self._eliminar_secciones_handler = (
            EliminarSeccionesRevisionWallFoundationHandler()
        )
        self._eliminar_secciones_event = ExternalEvent.Create(
            self._eliminar_secciones_handler
        )

        self._wire_ui(RoutedEventHandler)
        self._wire_keys(ApplicationCommands, CommandBinding, KeyBinding, Key, ModifierKeys)
        self._wire_lifecycle()
        self._wire_storyboard_completed()
        self._wire_activate_resel()

    def _wire_activate_resel(self):
        try:
            from System import EventHandler

            self._win.Activated += EventHandler(self._on_win_activated_resel)
        except Exception:
            pass

    def _on_win_activated_resel(self, sender, args):
        try:
            if getattr(self, "_wf_id", None) is None:
                return
            self._reselect_event.Raise()
        except Exception:
            pass

    def _wire_lifecycle(self):
        from System import EventHandler
        from System.Windows import RoutedEventHandler

        self._win.Closed += EventHandler(self._on_win_closed)
        self._win.Loaded += RoutedEventHandler(self._on_loaded)

    def _on_win_closed(self, sender, args):
        _clear_appdomain_window_key()
        try:
            if getattr(self, "_secciones_revision_ids", None):
                self._enqueue_eliminar_secciones_revision()
        except Exception:
            pass

    def _enqueue_eliminar_secciones_revision(self):
        try:
            _ids = list(self._secciones_revision_ids or [])
            self._secciones_revision_ids = []
            _doc = getattr(self, "_document", None)
            if _doc is None:
                try:
                    _ud = self._revit.ActiveUIDocument
                    if _ud is not None:
                        _doc = _ud.Document
                except Exception:
                    _doc = None
            if _doc is None:
                return
            self._eliminar_secciones_handler.armar(_doc, _ids)
            self._eliminar_secciones_event.Raise()
        except Exception:
            pass

    def _wire_storyboard_completed(self):
        try:
            from System import EventHandler

            sb = self._win.TryFindResource("FundOpenGrowStoryboard")
            if sb is not None:
                sb.Completed += EventHandler(self._on_open_storyboard_completed)
        except Exception:
            pass

    def _on_open_storyboard_completed(self, sender, args):
        try:
            self._win.MinWidth = float(self._form_width_px)
            self._win.MaxWidth = float(self._form_width_px)
        except Exception:
            pass

    def _on_loaded(self, s, a):
        try:
            from System import Action
            from System.Windows.Threading import DispatcherPriority

            self._win.Dispatcher.BeginInvoke(
                Action(self._begin_open_sb),
                DispatcherPriority.Loaded,
            )
        except Exception:
            self._begin_open_sb()

    def _begin_open_sb(self):
        if self._open_grow_storyboard_started:
            return
        self._open_grow_storyboard_started = True
        try:
            from System import TimeSpan
            from System.Windows import Duration, SizeToContent
            from System.Windows.Media import ScaleTransform

            sc = self._win.FindName("WallRootScale")
            if sc is not None:
                sc.ScaleX = 0.0
                sc.ScaleY = 0.0
            self._win.Width = float(self._form_width_px)
            try:
                self._win.SizeToContent = SizeToContent.Height
            except Exception:
                pass
            self._position_win()
            sb = self._win.TryFindResource("FundOpenGrowStoryboard")
            if sb is None:
                if sc is not None:
                    sc.ScaleX = sc.ScaleY = 1.0
                self._win.Opacity = 1.0
                return
            dur = Duration(TimeSpan.FromMilliseconds(float(_WINDOW_OPEN_MS)))
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

    def _step_sep(self, tb, delta):
        if tb is None:
            return
        try:
            v = int(round(float(unicode(tb.Text).replace(u"mm", u"").strip())))
        except Exception:
            v = _SEP_MM_DEFAULT
        v += int(delta)
        v = _snap_sep_mm(v, _SEP_MM_DEFAULT)
        tb.Text = unicode(int(v))

    def _position_win(self):
        try:
            uidoc = self._revit.ActiveUIDocument if self._revit else None
            hwnd = None
            if self._revit is not None:
                try:
                    hwnd = revit_main_hwnd(self._revit.Application)
                except Exception:
                    pass
            position_wpf_window_top_left_at_active_view(self._win, uidoc, hwnd)
        except Exception:
            pass

    def _wire_ui(self, RoutedEventHandler):
        from System.IO import FileAccess, FileMode, FileStream
        from System.Windows.Media.Imaging import BitmapCacheOption, BitmapImage
        from System.Windows.Input import MouseButtonEventHandler

        img = self._win.FindName("ImgLogo")
        if img is not None:
            for pth in get_logo_paths():
                if os.path.isfile(pth):
                    stream = None
                    try:
                        stream = FileStream(pth, FileMode.Open, FileAccess.Read)
                        bmp = BitmapImage()
                        bmp.BeginInit()
                        bmp.StreamSource = stream
                        bmp.CacheOption = BitmapCacheOption.OnLoad
                        bmp.EndInit()
                        bmp.Freeze()
                        img.Source = bmp
                        try:
                            self._win.Icon = bmp
                        except Exception:
                            pass
                    finally:
                        if stream is not None:
                            try:
                                stream.Dispose()
                            except Exception:
                                pass
                    break
        bp = self._win.FindName("BtnPick")
        if bp is not None:
            bp.Click += RoutedEventHandler(lambda s, e: self._pick_event.Raise())
        bc = self._win.FindName("BtnClose")
        if bc is not None:
            bc.Click += RoutedEventHandler(lambda s, e: self._close())
        bcol = self._win.FindName("BtnColocar")
        if bcol is not None:
            bcol.Click += RoutedEventHandler(lambda s, e: self._col_event.Raise())
        tb = self._win.FindName("TitleBar")
        if tb is not None:
            tb.MouseLeftButtonDown += MouseButtonEventHandler(
                lambda s, e: self._win.DragMove()
            )
        def _bind_sep(tb_name, up_name, dn_name):
            tb = self._win.FindName(tb_name)
            bu = self._win.FindName(up_name)
            bd = self._win.FindName(dn_name)

            def up(s, a):
                self._step_sep(tb, _SEP_MM_STEP)

            def dn(s, a):
                self._step_sep(tb, -_SEP_MM_STEP)

            if bu is not None:
                bu.Click += RoutedEventHandler(up)
            if bd is not None:
                bd.Click += RoutedEventHandler(dn)
            if tb is not None:
                tb.LostFocus += RoutedEventHandler(
                    lambda s, a, tbx=tb: _normalize_sep_tb(tbx)
                )

        _bind_sep("TxtTransSep", "BtnTransSepUp", "BtnTransSepDown")
        _bind_sep("TxtLongSep", "BtnLongSepUp", "BtnLongSepDown")
        tmax = self._win.FindName("TxtMaxBarMm")
        tlap = self._win.FindName("TxtLapMm")
        if tmax is not None:
            tmax.LostFocus += RoutedEventHandler(
                lambda s, a: _normalize_max_bar_tb(tmax)
            )
        if tlap is not None:
            tlap.LostFocus += RoutedEventHandler(lambda s, a: _normalize_lap_tb(tlap))
        try:
            from System.Windows.Controls import SelectionChangedEventHandler

            cmb_long = self._win.FindName("CmbLongDiam")
            if cmb_long is not None:
                cmb_long.SelectionChanged += SelectionChangedEventHandler(
                    self._on_cmb_long_diam_selection_changed
                )
            cmb_dos = self._win.FindName("CmbDosificacionHormigon")
            if cmb_dos is not None:
                cmb_dos.SelectionChanged += SelectionChangedEventHandler(
                    lambda s, a: self._sync_lap_tb_from_long_diam()
                )
        except Exception:
            pass

    def _wire_keys(self, ApplicationCommands, CommandBinding, KeyBinding, Key, ModifierKeys):
        from System.Windows.Input import ExecutedRoutedEventHandler

        try:
            self._win.CommandBindings.Add(
                CommandBinding(
                    ApplicationCommands.Close,
                    ExecutedRoutedEventHandler(lambda s, e: self._close()),
                )
            )
            self._win.InputBindings.Add(
                KeyBinding(ApplicationCommands.Close, Key.Escape, ModifierKeys.None)
            )
        except Exception:
            pass

    def _show_after_pick(self):
        try:
            self._win.Show()
            self._win.Activate()
        except Exception:
            pass

    def _refresh_troceo_panel(self):
        try:
            from System.Windows import Visibility

            br = self._win.FindName("BorderTroceo")
            if br is None:
                return
            doc = self._document
            wf_id = self._wf_id
            if doc is None or wf_id is None:
                br.Visibility = Visibility.Collapsed
                return
            wf = doc.GetElement(wf_id)
            if wf is None or not isinstance(wf, WallFoundation):
                br.Visibility = Visibility.Collapsed
                return
            Lmm = _longitud_eje_hint_mm(wf)
            if Lmm < 1.0:
                br.Visibility = Visibility.Collapsed
                return
            vis = bool(float(Lmm) > _MAX_STOCK_MM + 0.01)
            br.Visibility = Visibility.Visible if vis else Visibility.Collapsed
            if vis:
                try:
                    self._sync_lap_tb_from_long_diam()
                except Exception:
                    pass
        except Exception:
            pass

    def _on_cmb_long_diam_selection_changed(self, sender, args):
        try:
            self._sync_lap_tb_from_long_diam()
        except Exception:
            pass

    def _sync_lap_tb_from_long_diam(self):
        doc = self._document
        cmb = self._win.FindName("CmbLongDiam")
        tlap = self._win.FindName("TxtLapMm")
        if doc is None or cmb is None or tlap is None:
            return
        entries = getattr(self, "_entries", None) or []
        bt, _ = _resolver_bar_type_from_combo(doc, cmb, entries)
        d_mm = _rebar_nominal_diameter_mm(bt) if bt else None
        if d_mm is None:
            return
        gr = _read_dosificacion_hormigon(
            self._win.FindName("CmbDosificacionHormigon")
        )
        v = traslape_mm_from_nominal_diameter_mm(float(d_mm), gr)
        if v is not None:
            tlap.Text = unicode(int(round(v)))

    def _close(self):
        if getattr(self, "_is_closing_with_fade", False):
            return
        try:
            self._enqueue_eliminar_secciones_revision()
        except Exception:
            pass
        self._is_closing_with_fade = True
        try:
            from System import TimeSpan
            from System.Windows import Duration
            from System.Windows.Media import ScaleTransform
            from System.Windows.Media.Animation import DoubleAnimation, QuadraticEase, EasingMode

            sc = self._win.FindName("WallRootScale")
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

            from System import EventHandler

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
                self._enqueue_eliminar_secciones_revision()
            except Exception:
                pass
            try:
                self._win.Close()
            except Exception:
                pass
            self._is_closing_with_fade = False

    def _show_with_fade(self):
        try:
            self._win.Opacity = 0.0
            if not self._win.IsVisible:
                self._win.Show()
            self._win.Activate()
        except Exception:
            pass
        self._is_closing_with_fade = False

    def _cargar_combos(self):
        doc = self._document
        if doc is None:
            return
        entries, err = _build_bar_type_entries(doc)
        if err:
            try:
                _task_dialog_show(
                    u"BIMTools — Wall Foundation Reinforcement",
                    err,
                    self._win,
                )
            except Exception:
                pass
            entries = entries or []
        self._entries = entries
        for name in ("CmbTransDiam", "CmbLongDiam"):
            cmb = self._win.FindName(name)
            if cmb is None:
                continue
            try:
                cmb.Items.Clear()
            except Exception:
                pass
            for bt, lbl in entries:
                try:
                    cmb.Items.Add(lbl)
                except Exception:
                    pass
            try:
                cmb.SelectedIndex = 0
            except Exception:
                pass
        cmb_dos = self._win.FindName("CmbDosificacionHormigon")
        if cmb_dos is not None:
            try:
                cmb_dos.Items.Clear()
                for lab in _DOSIFICACION_HORMIGON_OPCIONES:
                    cmb_dos.Items.Add(lab)
            except Exception:
                pass
            try:
                cmb_dos.SelectedIndex = 0
            except Exception:
                pass

    def show(self):
        uidoc = self._revit.ActiveUIDocument
        if uidoc is None:
            TaskDialog.Show(
                u"Wall Foundation Reinforcement",
                u"No hay documento activo.",
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
                WindowInteropHelper(self._win).Owner = hwnd
        except Exception:
            pass
        position_wpf_window_top_left_at_active_view(self._win, uidoc, hwnd)
        self._cargar_combos()
        _normalize_sep_tb(self._win.FindName("TxtTransSep"))
        _normalize_sep_tb(self._win.FindName("TxtLongSep"))
        self._refresh_troceo_panel()
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
                u"BIMTools — Wall Foundation Reinforcement",
                u"La herramienta ya está en ejecución.",
                existing,
            )
            return

    w = WallFoundationReinforcementWindow(revit)
    try:
        w.show()
    except Exception:
        _clear_appdomain_window_key()
        raise
