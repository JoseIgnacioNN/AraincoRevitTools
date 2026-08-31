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
Sobre los Rebar creados se activa **View Unobscured** (+ sólido) en la vista activa.
No se generan vistas de sección de revisión en el modelo.

**UI:** shell BIMTools con canvas de planta a la izquierda y canvas de sección (corte por el
ancho) en el rail derecho, junto a los parámetros ø/@.
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
    aplicar_recubrimiento_extremos_mm,
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

from bimtools_ui_tokens import WINDOW_CHROME_TITLE
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML

_APPDOMAIN_WINDOW_KEY = "BIMTools.WallFoundationReinforcement.ActiveWindow"
_FT_TO_MM = 304.8
_PLAN_PAD_FRAC = 0.08
_SECTION_PAD_PX = 28.0
_BRUSH_CACHE = {}
_COLOR_TRANS = u"#5BC0DE"
_COLOR_LONG = u"#4ade80"

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


def _wf_resolver_rebars_frescos(doc, rebars):
    """Relee ``Rebar`` desde el documento (ids válidos tras commit / regen)."""
    out = []
    if doc is None or not rebars:
        return out
    seen = set()
    for ref in rebars:
        if ref is None:
            continue
        try:
            rid = ref.Id if hasattr(ref, "Id") else ref
            el = doc.GetElement(rid)
        except Exception:
            el = None
        if el is None or not isinstance(el, Rebar):
            continue
        try:
            key = element_id_to_int(el.Id)
        except Exception:
            key = None
        if key is not None:
            if key in seen:
                continue
            seen.add(key)
        out.append(el)
    return out


def _wf_aplicar_unobscured_rebars(doc, view, rebars):
    """
    Activa View Unobscured (+ sólido en vista) en los Rebar indicados para ``view``.
    Devuelve cuántas barras recibieron Unobscured.
    """
    if doc is None or view is None or not rebars:
        return 0
    fresh = _wf_resolver_rebars_frescos(doc, rebars)
    if not fresh:
        return 0
    try:
        from bimtools_rebar_3d_visibility import apply_rebar_unobscured_in_view

        return int(apply_rebar_unobscured_in_view(doc, fresh, view) or 0)
    except Exception:
        n_ok = 0
        for rb in fresh:
            try:
                rb.SetUnobscuredInView(view, True)
                n_ok += 1
            except Exception:
                pass
            try:
                rb.SetSolidInView(view, True)
            except Exception:
                pass
        return n_ok


_FUND_INPUT_COLS_PER_ROW = 2
_FUND_COMBO_WIDTH_PX = 110
_FUND_DIAM_ESP_AT_COL_PX = 28
_FUND_BLOCK_PAD_H_PX = 16
_FUND_GROUPBOX_PAD_H_PX = 16
_FUND_OUTER_PAD_H_PX = 36
_FUND_ARM_INFO_GROUPBOX_EXTRA_H_PX = 12
_FUND_WIDTH_TITLE_MIN_PX = 420
_FUND_WIDTH_FOOTER_MIN_PX = 300


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
    w_content = (
        row_inner
        + _FUND_GROUPBOX_PAD_H_PX
        + _FUND_ARM_INFO_GROUPBOX_EXTRA_H_PX
    )
    w = max(
        w_content + _FUND_OUTER_PAD_H_PX,
        _FUND_WIDTH_TITLE_MIN_PX,
        _FUND_WIDTH_FOOTER_MIN_PX + _FUND_OUTER_PAD_H_PX,
    )
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


def _wf_as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _wf_brush(hex_color, alpha=255):
    from System.Windows.Media import Color, SolidColorBrush

    h = (_wf_as_unicode(hex_color) or u"#95B8CC").lstrip(u"#")
    if len(h) != 6:
        h = u"95B8CC"
    try:
        a = int(alpha)
    except Exception:
        a = 255
    key = (h, a)
    cached = _BRUSH_CACHE.get(key)
    if cached is not None:
        return cached
    brush = SolidColorBrush(
        Color.FromArgb(a, int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    )
    try:
        if brush.CanFreeze:
            brush.Freeze()
    except Exception:
        pass
    _BRUSH_CACHE[key] = brush
    return brush


def _wf_preview_width_ft_from_bbox(wf, tu_xy):
    """Ancho en planta proyectando el bbox sobre la normal al eje (ft)."""
    bb = wf.get_BoundingBox(None) if wf is not None else None
    if bb is None or tu_xy is None:
        return None
    try:
        nx = float(-tu_xy.Y)
        ny = float(tu_xy.X)
        nlen = (nx * nx + ny * ny) ** 0.5
        if nlen < 1e-12:
            return None
        nx /= nlen
        ny /= nlen
        corners = (
            (float(bb.Min.X), float(bb.Min.Y)),
            (float(bb.Max.X), float(bb.Min.Y)),
            (float(bb.Max.X), float(bb.Max.Y)),
            (float(bb.Min.X), float(bb.Max.Y)),
        )
        projs = [cx * nx + cy * ny for cx, cy in corners]
        return abs(max(projs) - min(projs))
    except Exception:
        return None


def _wf_preview_geo_mm(wf):
    """
    Geometría de preview (solo lectura, sin Unjoin).

    Returns dict:
      poly: lista [(x_mm, y_mm), ...] contorno en planta
      length_mm, width_mm, height_mm
      label: texto host
    o None.
    """
    if wf is None or not isinstance(wf, WallFoundation):
        return None
    ax, _ = _axis_line_wall_foundation(wf)
    tu = _wf_tu_xy_desde_linea(ax) if ax is not None else None
    length_ft = None
    p0 = p1 = None
    if ax is not None:
        try:
            p0 = ax.GetEndPoint(0)
            p1 = ax.GetEndPoint(1)
            length_ft = float(ax.Length)
        except Exception:
            length_ft = None
    if length_ft is None or length_ft < 1e-6 or p0 is None or tu is None:
        bb = wf.get_BoundingBox(None)
        if bb is None:
            return None
        try:
            dx = float(bb.Max.X - bb.Min.X)
            dy = float(bb.Max.Y - bb.Min.Y)
            if dx >= dy:
                length_ft = dx
                tu = XYZ(1.0, 0.0, 0.0)
                mid_y = 0.5 * (float(bb.Min.Y) + float(bb.Max.Y))
                mid_z = 0.5 * (float(bb.Min.Z) + float(bb.Max.Z))
                p0 = XYZ(float(bb.Min.X), mid_y, mid_z)
                p1 = XYZ(float(bb.Max.X), mid_y, mid_z)
            else:
                length_ft = dy
                tu = XYZ(0.0, 1.0, 0.0)
                mid_x = 0.5 * (float(bb.Min.X) + float(bb.Max.X))
                mid_z = 0.5 * (float(bb.Min.Z) + float(bb.Max.Z))
                p0 = XYZ(mid_x, float(bb.Min.Y), mid_z)
                p1 = XYZ(mid_x, float(bb.Max.Y), mid_z)
        except Exception:
            return None
    width_ft = _wf_width_ft(wf)
    if width_ft is None or width_ft < 1e-6:
        width_ft = _wf_preview_width_ft_from_bbox(wf, tu)
    if width_ft is None or width_ft < 1e-6:
        width_ft = _mm_to_ft(600.0)
    z0, z1 = _wf_z_range_ft(wf)
    if z0 is None or z1 is None:
        height_ft = _mm_to_ft(500.0)
    else:
        height_ft = max(_mm_to_ft(50.0), abs(float(z1) - float(z0)))
    try:
        tn = XYZ(float(-tu.Y), float(tu.X), 0.0)
        if float(tn.GetLength()) < 1e-12:
            tn = XYZ(0.0, 1.0, 0.0)
        else:
            tn = tn.Normalize()
        half = 0.5 * float(width_ft)
        c0 = p0.Add(tn.Multiply(half))
        c1 = p1.Add(tn.Multiply(half))
        c2 = p1.Add(tn.Multiply(-half))
        c3 = p0.Add(tn.Multiply(-half))
        poly = [
            (float(c0.X) * _FT_TO_MM, float(c0.Y) * _FT_TO_MM),
            (float(c1.X) * _FT_TO_MM, float(c1.Y) * _FT_TO_MM),
            (float(c2.X) * _FT_TO_MM, float(c2.Y) * _FT_TO_MM),
            (float(c3.X) * _FT_TO_MM, float(c3.Y) * _FT_TO_MM),
        ]
    except Exception:
        return None
    label = u"Wall Foundation"
    try:
        eid = element_id_to_int(wf.Id)
        if eid is not None:
            label = u"Wall Foundation Id {0}".format(eid)
    except Exception:
        pass
    try:
        nm = _wf_as_unicode(wf.Name)
        if nm:
            label = u"{0} · {1}".format(nm, label)
    except Exception:
        pass
    return {
        u"poly": poly,
        u"length_mm": float(length_ft) * _FT_TO_MM,
        u"width_mm": float(width_ft) * _FT_TO_MM,
        u"height_mm": float(height_ft) * _FT_TO_MM,
        u"label": label,
        u"p0_mm": (float(p0.X) * _FT_TO_MM, float(p0.Y) * _FT_TO_MM),
        u"p1_mm": (float(p1.X) * _FT_TO_MM, float(p1.Y) * _FT_TO_MM),
    }


def _wf_preview_positions_along(length_mm, spacing_mm, cover_mm):
    L = float(length_mm)
    e = max(1.0, float(spacing_mm))
    c = max(0.0, float(cover_mm))
    if L <= 2.0 * c + 1.0:
        return [L * 0.5]
    start = c
    end = L - c
    xs = []
    x = start
    guard = 0
    while x <= end + 0.5 and guard < 500:
        xs.append(x)
        x += e
        guard += 1
    if not xs:
        xs = [L * 0.5]
    if xs and abs(xs[-1] - end) > 0.5:
        if end - xs[-1] > 0.5 * e:
            xs.append(end)
        else:
            xs[-1] = end
    return xs


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


def _wf_apply_recubrimiento_ejes_franja(wf, geo, diam_long_mm, diam_trans_mm):
    """
    Recubrimiento en ejes long./ancho derivados de una **franja** dibujada.

    La franja ya delimita la zona en planta (sin offset de 100 mm del perímetro).
    Se aplica el mismo criterio que ``_geometria_wf_cara_inferior_tol`` en extremos
    y en vertical (50 mm + ø/2 al eje desde la cara inferior).
    """
    if geo is None:
        return None
    long_raw = geo.get("long_line")
    width_raw = geo.get("width_line")
    if long_raw is None or width_raw is None:
        return geo
    n_cara = geo.get("n_cara")
    if n_cara is None:
        marco = geo.get("marco_uvn")
        if marco is not None and len(marco) > 3:
            n_cara = marco[3]
    if n_cara is None:
        n_cara = XYZ.BasisZ.Negate()
    try:
        d_l = float(diam_long_mm) if diam_long_mm else 0.0
        d_t = float(diam_trans_mm) if diam_trans_mm else 0.0
    except Exception:
        d_l = d_t = 0.0
    ext_long_mm = float(_REC_EXTREMOS_LONG_TANGENTE_MM)
    if d_l > 1e-6:
        ext_long_mm = float(_REC_EXTREMOS_LONG_TANGENTE_MM) + 0.5 * d_l
    ct_long = aplicar_recubrimiento_extremos_mm(long_raw, ext_long_mm)
    ct_width = aplicar_recubrimiento_extremos_mm(
        width_raw, float(_REC_EXTREMOS_INFERIOR_MM)
    )
    if ct_long is None or ct_width is None:
        return geo
    long_bar = offset_linea_eje_barra_desde_cara_inferior_mm(
        ct_long, n_cara, _RECO_HOR_MM, d_l
    )
    width_bar = offset_linea_eje_barra_desde_cara_inferior_mm(
        ct_width, n_cara, _RECO_HOR_MM, d_t
    )
    if long_bar is None or width_bar is None:
        return geo
    geo["long_line"] = long_bar
    geo["width_line"] = width_bar
    try:
        geo["usable_w_ft"] = float(width_bar.Length)
    except Exception:
        pass
    geo["n_cara"] = n_cara
    return geo


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


# --- XAML shell: planta izquierda + rail seccion / params (como fundacion aislada) ----
_WF_XAML = (
    u"""
<Window
    x:Name="WallFoundWin"
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="__CHROME__"
    Height="820" Width="1040"
    MinHeight="760" MinWidth="980"
    WindowStartupLocation="Manual"
    Background="#071018"
    FontFamily="Segoe UI"
    FontSize="12"
    ShowInTaskbar="False"
    ResizeMode="CanResize"
    UseLayoutRounding="True">
  <Window.Resources>
"""
    + BIMTOOLS_DARK_STYLES_XML
    + u"""
  </Window.Resources>
  <Border Background="#071018" BorderBrush="#21465C" BorderThickness="1" Padding="18">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <StackPanel Grid.Row="0" Margin="0,0,0,8">
        <TextBlock x:Name="TxtTitle" Text="Arainco: Armadura Fundación Corrida"
                   Foreground="#E8F4F8" FontSize="18" FontWeight="Bold"/>
        <TextBlock x:Name="TxtSubtitle" Margin="0,6,0,0" Foreground="#95B8CC"
                   FontSize="11" TextWrapping="Wrap"
                   Text="Planta a la izquierda · sección por el ancho en el rail · transversales U y longitudinales."/>
      </StackPanel>

      <Grid Grid.Row="1">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="380"/>
        </Grid.ColumnDefinitions>

        <Border Grid.Column="0" Background="#0a1620" BorderBrush="#21465C"
                BorderThickness="1" CornerRadius="4,0,0,4" Padding="0">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Border Grid.Row="0" Background="#0a1620" BorderBrush="#21465C"
                    BorderThickness="0,0,0,1" Padding="8,6,8,4">
              <TextBlock x:Name="TxtCanvasHeader"
                         Foreground="#64748b" FontSize="10" FontWeight="SemiBold"
                         VerticalAlignment="Center"
                         Text="PLANTA · ZAPATA DE MURO"/>
            </Border>
            <Border Grid.Row="1" Background="#050E18" BorderBrush="Transparent"
                    BorderThickness="0" Padding="8,4,8,8">
              <Border Background="#050E18" BorderBrush="#21465C"
                      BorderThickness="1" CornerRadius="4">
                <Canvas x:Name="CvPlan" ClipToBounds="True" Background="#050E18"/>
              </Border>
            </Border>
          </Grid>
        </Border>

        <Border Grid.Column="1" Background="#0a1620" BorderBrush="#21465C"
                BorderThickness="1" CornerRadius="0,4,4,0" Padding="8,8">
          <ScrollViewer VerticalScrollBarVisibility="Auto"
                        HorizontalScrollBarVisibility="Disabled">
            <StackPanel x:Name="PnlSectionRail">

              <Border Background="#0a1620" BorderBrush="#21465C"
                      BorderThickness="1" CornerRadius="4" Padding="8" Margin="0,0,0,10">
                <StackPanel>
                  <TextBlock Text="SECCIÓN · ANCHO" Foreground="#64748b"
                             FontSize="10" FontWeight="SemiBold" Margin="0,0,0,6"/>
                  <Border Background="#050E18" BorderBrush="#21465C"
                          BorderThickness="1" CornerRadius="4" Height="220">
                    <Canvas x:Name="CvSection" ClipToBounds="True" Background="#050E18"/>
                  </Border>
                  <TextBlock x:Name="TxtSectionDims" Foreground="#64748b" FontSize="10"
                             Margin="0,6,0,0" TextWrapping="Wrap" Text=""/>
                </StackPanel>
              </Border>

              <Border Background="#0a1620" BorderBrush="#21465C"
                      BorderThickness="1" CornerRadius="4" Padding="10" Margin="0,0,0,10">
                <StackPanel>
                  <TextBlock Text="Fundación" Foreground="#E8F4F8"
                             FontSize="12" FontWeight="SemiBold" Margin="0,0,0,6"/>
                  <TextBlock x:Name="TxtHost" Foreground="#95B8CC" FontSize="11"
                             TextWrapping="Wrap" Text="— Sin selección —"/>
                  <Button x:Name="BtnPick" Content="Seleccionar fundación"
                          Style="{StaticResource BtnSelectOutline}"
                          HorizontalAlignment="Stretch" Margin="0,8,0,0" Padding="10,6"/>
                </StackPanel>
              </Border>

              <Border Background="#0a1620" BorderBrush="#21465C"
                      BorderThickness="1" CornerRadius="4" Padding="10" Margin="0,0,0,0">
                <StackPanel>
                  <Grid Margin="0,0,0,10">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Text="Armadura" Foreground="#E8F4F8"
                               FontSize="12" FontWeight="SemiBold" VerticalAlignment="Center"/>
                    <ComboBox x:Name="CmbDosificacionHormigon" Grid.Column="1"
                              Style="{StaticResource Combo}" Margin="8,0,0,0"
                              IsEditable="False" IsReadOnly="True" VerticalAlignment="Center"
                              MinWidth="72" ToolTip="Dosificación del hormigón">
                      <ComboBox.ItemContainerStyle>
                        <Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/>
                      </ComboBox.ItemContainerStyle>
                    </ComboBox>
                  </Grid>

                  <TextBlock Text="Transversales (U)" Foreground="#95B8CC" FontWeight="SemiBold"
                             FontSize="11" Margin="0,0,0,6"/>
                  <Grid Margin="0,0,0,12" HorizontalAlignment="Stretch">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="Auto"/>
                      <ColumnDefinition Width="110"/>
                    </Grid.ColumnDefinitions>
                    <ComboBox Grid.Column="0" x:Name="CmbTransDiam" Style="{StaticResource Combo}"
                              IsEditable="False" IsReadOnly="True">
                      <ComboBox.ItemContainerStyle>
                        <Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/>
                      </ComboBox.ItemContainerStyle>
                    </ComboBox>
                    <TextBlock Grid.Column="1" Text="@" FontSize="12" FontWeight="Bold"
                               Foreground="#95B8CC" VerticalAlignment="Center" Margin="6,0,6,0"/>
                    <Border Grid.Column="2" Height="24" CornerRadius="5" Background="#050E18"
                            BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True">
                      <Grid>
                        <Grid.ColumnDefinitions>
                          <ColumnDefinition Width="*"/>
                          <ColumnDefinition Width="18"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="TxtTransSep" Grid.Column="0"
                                 Style="{StaticResource CantSpinnerText}"
                                 Text="100" Padding="6,0,6,0" VerticalContentAlignment="Center"
                                 ToolTip="Separación transversal (mm): 100 a 400, pasos 10"/>
                        <Border Grid.Column="1" Background="#11253D" BorderBrush="#1A3A4D"
                                BorderThickness="1,0,0,0" CornerRadius="0,5,5,0" ClipToBounds="True">
                          <Grid>
                            <Grid.RowDefinitions>
                              <RowDefinition Height="*"/>
                              <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <RepeatButton x:Name="BtnTransSepUp" Grid.Row="0"
                                          Style="{StaticResource SpinRepeatBtn}" Content="▲"
                                          ToolTip="Más 10 mm (máx. 400 mm)"/>
                            <RepeatButton x:Name="BtnTransSepDown" Grid.Row="1"
                                          Style="{StaticResource SpinRepeatBtn}" Content="▼"
                                          ToolTip="Menos 10 mm (mín. 100 mm)"/>
                          </Grid>
                        </Border>
                      </Grid>
                    </Border>
                  </Grid>

                  <TextBlock Text="Longitudinales" Foreground="#95B8CC" FontWeight="SemiBold"
                             FontSize="11" Margin="0,0,0,6"/>
                  <Grid Margin="0,0,0,0" HorizontalAlignment="Stretch">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="Auto"/>
                      <ColumnDefinition Width="110"/>
                    </Grid.ColumnDefinitions>
                    <ComboBox Grid.Column="0" x:Name="CmbLongDiam" Style="{StaticResource Combo}"
                              IsEditable="False" IsReadOnly="True">
                      <ComboBox.ItemContainerStyle>
                        <Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/>
                      </ComboBox.ItemContainerStyle>
                    </ComboBox>
                    <TextBlock Grid.Column="1" Text="@" FontSize="12" FontWeight="Bold"
                               Foreground="#95B8CC" VerticalAlignment="Center" Margin="6,0,6,0"/>
                    <Border Grid.Column="2" Height="24" CornerRadius="5" Background="#050E18"
                            BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True">
                      <Grid>
                        <Grid.ColumnDefinitions>
                          <ColumnDefinition Width="*"/>
                          <ColumnDefinition Width="18"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="TxtLongSep" Grid.Column="0"
                                 Style="{StaticResource CantSpinnerText}"
                                 Text="100" Padding="6,0,6,0" VerticalContentAlignment="Center"
                                 ToolTip="Separación longitudinal (mm): 100 a 400, pasos 10"/>
                        <Border Grid.Column="1" Background="#11253D" BorderBrush="#1A3A4D"
                                BorderThickness="1,0,0,0" CornerRadius="0,5,5,0" ClipToBounds="True">
                          <Grid>
                            <Grid.RowDefinitions>
                              <RowDefinition Height="*"/>
                              <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <RepeatButton x:Name="BtnLongSepUp" Grid.Row="0"
                                          Style="{StaticResource SpinRepeatBtn}" Content="▲"
                                          ToolTip="Más 10 mm (máx. 400 mm)"/>
                            <RepeatButton x:Name="BtnLongSepDown" Grid.Row="1"
                                          Style="{StaticResource SpinRepeatBtn}" Content="▼"
                                          ToolTip="Menos 10 mm (mín. 100 mm)"/>
                          </Grid>
                        </Border>
                      </Grid>
                    </Border>
                  </Grid>

                  <Border x:Name="BorderTroceo" Visibility="Collapsed" Margin="0,10,0,0"
                          HorizontalAlignment="Stretch"
                          Background="#050E18" BorderBrush="#21465C" BorderThickness="1"
                          CornerRadius="4" Padding="8,8,8,8">
                    <StackPanel>
                      <TextBlock Text="Troceo longitudinal (&gt; 12 m de eje)"
                                 Style="{StaticResource Label}" Margin="0,0,0,4"/>
                      <Grid Margin="0,0,0,0">
                        <Grid.ColumnDefinitions>
                          <ColumnDefinition Width="*"/>
                          <ColumnDefinition Width="8"/>
                          <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Grid.Column="0">
                          <TextBlock Text="Largo máx. tramo (mm)" Foreground="#95B8CC" FontSize="10"/>
                          <TextBox x:Name="TxtMaxBarMm" Style="{StaticResource CantSpinnerText}"
                                   Background="#050E18" BorderBrush="#1A3A4D" BorderThickness="1"
                                   Padding="4" Text="12000"/>
                        </StackPanel>
                        <StackPanel Grid.Column="2">
                          <TextBlock Text="Empalme / traslape (mm)" Foreground="#95B8CC" FontSize="10"/>
                          <TextBox x:Name="TxtLapMm" Style="{StaticResource CantSpinnerText}"
                                   Background="#050E18" BorderBrush="#1A3A4D" BorderThickness="1"
                                   Padding="4" Text="600"/>
                        </StackPanel>
                      </Grid>
                    </StackPanel>
                  </Border>
                </StackPanel>
              </Border>

            </StackPanel>
          </ScrollViewer>
        </Border>
      </Grid>

      <TextBlock Grid.Row="2" x:Name="TxtHint" Foreground="#64748b" FontSize="10"
                 TextWrapping="Wrap" Margin="0,8,0,0"
                 Text="Rueda = zoom en planta · arrastre = pan · la sección muestra el corte por el ancho."/>

      <Grid Grid.Row="3" Margin="0,14,0,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <Button x:Name="BtnManual" Grid.Column="0" Content="Manual"
                Style="{StaticResource BtnSelectOutline}" MinWidth="96"
                Margin="0,0,12,0" Background="#2A5C3D"
                ToolTip="Abrir manual de usuario" VerticalAlignment="Center"/>
        <TextBlock x:Name="TxtEstado" Grid.Column="1" VerticalAlignment="Center"
                   Foreground="#64748b" FontSize="10" TextWrapping="Wrap" Margin="0,0,12,0"/>
        <StackPanel Grid.Column="2" Orientation="Horizontal" HorizontalAlignment="Right">
          <Button x:Name="BtnCancelar" Content="Cancelar"
                  Style="{StaticResource BtnSelectOutline}" MinWidth="110" Margin="0,0,10,0"/>
          <Button x:Name="BtnColocar" Content="Colocar armadura"
                  Style="{StaticResource BtnPrimary}" MinWidth="160"/>
        </StackPanel>
      </Grid>
    </Grid>
  </Border>
</Window>
"""
).replace(u"__CHROME__", WINDOW_CHROME_TITLE)


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
                u"Arainco: Wall Foundation Reinforcement",
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
        try:
            win._refresh_preview_from_selection()
        except Exception:
            pass
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
                u"Arainco: Wall Foundation Reinforcement",
                u"No hay documento activo.",
                win._win,
            )
            return
        doc = uidoc.Document
        wf_id = getattr(win, "_wf_id", None)
        if wf_id is None:
            _task_dialog_show(
                u"Arainco: Wall Foundation Reinforcement",
                u"Seleccione una Wall Foundation.",
                win._win,
            )
            return
        wf = doc.GetElement(wf_id)
        if wf is None or not isinstance(wf, WallFoundation):
            _task_dialog_show(
                u"Arainco: Wall Foundation Reinforcement",
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
                u"Arainco: Wall Foundation Reinforcement",
                e1 or u"Tipo de barra transversal no válido.",
                win._win,
            )
            return
        if do_l and bt_lo is None:
            _task_dialog_show(
                u"Arainco: Wall Foundation Reinforcement",
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
                u"Arainco: Wall Foundation Reinforcement",
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
                    u"Arainco: Wall Foundation Reinforcement",
                    u"El largo máximo por tramo debe ser mayor que el empalme.",
                    win._win,
                )
                return
        avisos = []
        rebars_sets_creados = []
        active_view = uidoc.ActiveView
        n_t = n_l = 0
        t = Transaction(doc, u"Arainco: Wall Foundation Reinforcement")
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
            if rebars_sets_creados:
                try:
                    _n_unob = _wf_aplicar_unobscured_rebars(
                        doc, active_view, rebars_sets_creados
                    )
                    if _n_unob > 0:
                        avisos.append(
                            u"View Unobscured (+ sólido): {0} barra(s) en la vista activa.".format(
                                int(_n_unob)
                            )
                        )
                except Exception:
                    pass
            t.Commit()
        except Exception as ex:
            t.RollBack()
            _task_dialog_show(
                u"Arainco: Wall Foundation Reinforcement",
                u"Error (se revirtió la transacción):\n{0}".format(ex),
                win._win,
            )
            return
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
                    u"Arainco: Wall Foundation — Resultado",
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
        try:
            win._refresh_preview_from_selection()
        except Exception:
            pass
        try:
            win._set_estado(u"Armadura colocada. Puede seleccionar otra fundación.")
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
        self._preview_geo = None
        self._view_zoom = 1.0
        self._view_pan_x = 0.0
        self._view_pan_y = 0.0
        self._scene_base = None
        self._panning = False
        self._pan_last = None
        self._ui_cv_plan = None
        self._ui_cv_section = None
        self._ui_txt_header = None
        self._ui_txt_section_dims = None
        self._ui_txt_host = None

        from System.Windows import RoutedEventHandler
        from System.Windows.Input import ApplicationCommands, CommandBinding, Key, KeyBinding, ModifierKeys
        from System.Windows.Markup import XamlReader

        self._win = XamlReader.Parse(_WF_XAML)

        self._pick_handler = PickWallFoundationHandler(weakref.ref(self))
        self._pick_event = ExternalEvent.Create(self._pick_handler)
        self._col_handler = ColocarWallFoundationHandler(weakref.ref(self))
        self._col_event = ExternalEvent.Create(self._col_handler)
        self._reselect_handler = ReselectWallFoundationHandler(weakref.ref(self))
        self._reselect_event = ExternalEvent.Create(self._reselect_handler)

        self._cache_ui_refs()
        self._wire_ui(RoutedEventHandler)
        self._wire_canvas_interaction()
        self._wire_keys(ApplicationCommands, CommandBinding, KeyBinding, Key, ModifierKeys)
        self._wire_lifecycle()
        self._wire_activate_resel()

    def _cache_ui_refs(self):
        win = self._win
        try:
            self._ui_cv_plan = win.FindName(u"CvPlan")
        except Exception:
            self._ui_cv_plan = None
        try:
            self._ui_cv_section = win.FindName(u"CvSection")
        except Exception:
            self._ui_cv_section = None
        try:
            self._ui_txt_header = win.FindName(u"TxtCanvasHeader")
        except Exception:
            self._ui_txt_header = None
        try:
            self._ui_txt_section_dims = win.FindName(u"TxtSectionDims")
        except Exception:
            self._ui_txt_section_dims = None
        try:
            self._ui_txt_host = win.FindName(u"TxtHost")
        except Exception:
            self._ui_txt_host = None

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

        self._win.Closed += EventHandler(self._on_win_closed)
        try:
            from System.Windows import SizeChangedEventHandler

            self._win.SizeChanged += SizeChangedEventHandler(self._on_size_changed)
        except Exception:
            try:
                self._win.SizeChanged += EventHandler(self._on_size_changed)
            except Exception:
                pass

    def _on_win_closed(self, sender, args):
        _clear_appdomain_window_key()

    def _on_size_changed(self, sender, args):
        try:
            self._redraw_all()
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
        try:
            self._redraw_all()
        except Exception:
            pass

    def _request_redraw(self, sender=None, args=None):
        try:
            self._redraw_all()
        except Exception:
            pass

    def _resolve_manual_path(self):
        candidates = []
        try:
            import bimtools_paths

            pb = bimtools_paths.get_pushbutton_dir()
            if pb:
                candidates.append(os.path.join(pb, u"manual_usuario.html"))
        except Exception:
            pass
        try:
            ext_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            for tab_name in os.listdir(ext_dir):
                if not tab_name.endswith(u".tab"):
                    continue
                panel = os.path.join(ext_dir, tab_name, u"Armadura.panel")
                if not os.path.isdir(panel):
                    continue
                for pb_name in os.listdir(panel):
                    if u"WallFoundationReinforcement" not in pb_name:
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

    def _open_manual(self, sender=None, args=None):
        path = self._resolve_manual_path()
        if not path:
            _task_dialog_show(
                u"Arainco: Armadura Fundación Corrida",
                u"No se encontró manual_usuario.html en la carpeta del botón.",
                self._win,
            )
            return
        try:
            os.startfile(path)
        except Exception as ex:
            _task_dialog_show(
                u"Arainco: Armadura Fundación Corrida",
                u"No se pudo abrir el manual:\n{0}".format(ex),
                self._win,
            )

    def _wire_ui(self, RoutedEventHandler):
        bp = self._win.FindName("BtnPick")
        if bp is not None:
            bp.Click += RoutedEventHandler(lambda s, e: self._pick_event.Raise())
        btn_cancel = self._win.FindName("BtnCancelar")
        if btn_cancel is not None:
            btn_cancel.Click += RoutedEventHandler(lambda s, e: self._close())
        bcol = self._win.FindName("BtnColocar")
        if bcol is not None:
            bcol.Click += RoutedEventHandler(lambda s, e: self._col_event.Raise())
        bman = self._win.FindName("BtnManual")
        if bman is not None:
            bman.Click += RoutedEventHandler(self._open_manual)

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
                    lambda s, a, tbx=tb: (
                        _normalize_sep_tb(tbx),
                        self._request_redraw(),
                    )
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
            cmb_trans = self._win.FindName("CmbTransDiam")
            if cmb_trans is not None:
                cmb_trans.SelectionChanged += SelectionChangedEventHandler(
                    self._request_redraw
                )
            cmb_dos = self._win.FindName("CmbDosificacionHormigon")
            if cmb_dos is not None:
                cmb_dos.SelectionChanged += SelectionChangedEventHandler(
                    lambda s, a: self._sync_lap_tb_from_long_diam()
                )
        except Exception:
            pass

    def _wire_canvas_interaction(self):
        from System.Windows.Input import (
            MouseButtonEventHandler,
            MouseEventHandler,
            MouseWheelEventHandler,
        )

        cv = self._ui_cv_plan
        if cv is None:
            return
        try:
            cv.MouseWheel += MouseWheelEventHandler(self._on_plan_wheel)
        except Exception:
            pass
        try:
            cv.MouseLeftButtonDown += MouseButtonEventHandler(self._on_plan_mouse_down)
            cv.MouseMove += MouseEventHandler(self._on_plan_mouse_move)
            cv.MouseLeftButtonUp += MouseButtonEventHandler(self._on_plan_mouse_up)
            cv.MouseLeave += MouseEventHandler(self._on_plan_mouse_up)
        except Exception:
            pass

    def _on_plan_wheel(self, sender, args):
        try:
            delta = int(args.Delta)
        except Exception:
            return
        factor = 1.1 if delta > 0 else (1.0 / 1.1)
        try:
            nz = float(self._view_zoom) * factor
            self._view_zoom = max(0.2, min(8.0, nz))
        except Exception:
            return
        self._redraw_plan()

    def _on_plan_mouse_down(self, sender, args):
        try:
            self._panning = True
            self._pan_last = args.GetPosition(self._ui_cv_plan)
            try:
                self._ui_cv_plan.CaptureMouse()
            except Exception:
                pass
        except Exception:
            self._panning = False
            self._pan_last = None

    def _on_plan_mouse_move(self, sender, args):
        if not self._panning or self._scene_base is None:
            return
        try:
            pos = args.GetPosition(self._ui_cv_plan)
            last = self._pan_last
            if last is None:
                self._pan_last = pos
                return
            dx_px = float(pos.X - last.X)
            dy_px = float(pos.Y - last.Y)
            scale = float(self._scene_base.get(u"scale") or 1.0)
            if scale < 1e-9:
                return
            self._view_pan_x -= dx_px / scale
            self._view_pan_y += dy_px / scale
            self._pan_last = pos
            self._redraw_plan()
        except Exception:
            pass

    def _on_plan_mouse_up(self, sender, args):
        if not self._panning:
            return
        self._panning = False
        self._pan_last = None
        try:
            if self._ui_cv_plan is not None:
                self._ui_cv_plan.ReleaseMouseCapture()
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
            self._redraw_all()
        except Exception:
            pass

    def _refresh_preview_from_selection(self):
        self._preview_geo = None
        self._view_zoom = 1.0
        self._view_pan_x = 0.0
        self._view_pan_y = 0.0
        doc = self._document
        wf_id = self._wf_id
        if doc is None or wf_id is None:
            if self._ui_txt_host is not None:
                self._ui_txt_host.Text = u"— Sin selección —"
            self._redraw_all()
            return
        wf = doc.GetElement(wf_id)
        geo = _wf_preview_geo_mm(wf)
        self._preview_geo = geo
        if self._ui_txt_host is not None:
            try:
                self._ui_txt_host.Text = (
                    geo.get(u"label") if geo else u"— Sin selección —"
                ) or u"— Sin selección —"
            except Exception:
                pass
        self._redraw_all()

    def _read_preview_bar_params(self):
        doc = self._document
        entries = getattr(self, "_entries", None) or []
        d_tr = 8.0
        d_lo = 8.0
        sep_tr = float(_SEP_MM_DEFAULT)
        sep_lo = float(_SEP_MM_DEFAULT)
        try:
            bt_tr, _ = _resolver_bar_type_from_combo(
                doc, self._win.FindName("CmbTransDiam"), entries
            )
            if bt_tr is not None:
                v = _rebar_nominal_diameter_mm(bt_tr)
                if v:
                    d_tr = float(v)
        except Exception:
            pass
        try:
            bt_lo, _ = _resolver_bar_type_from_combo(
                doc, self._win.FindName("CmbLongDiam"), entries
            )
            if bt_lo is not None:
                v = _rebar_nominal_diameter_mm(bt_lo)
                if v:
                    d_lo = float(v)
        except Exception:
            pass
        try:
            sep_tr = float(_read_sep_tb(self._win.FindName("TxtTransSep")))
        except Exception:
            pass
        try:
            sep_lo = float(_read_sep_tb(self._win.FindName("TxtLongSep")))
        except Exception:
            pass
        return d_tr, d_lo, sep_tr, sep_lo

    def _canvas_size(self, cv):
        w = h = 40.0
        try:
            w = float(cv.ActualWidth or 0)
        except Exception:
            pass
        try:
            h = float(cv.ActualHeight or 0)
        except Exception:
            pass
        if w < 40:
            try:
                w = float(cv.RenderSize.Width or 0)
            except Exception:
                pass
        if h < 40:
            try:
                h = float(cv.RenderSize.Height or 0)
            except Exception:
                pass
        return max(40.0, w), max(40.0, h)

    def _redraw_all(self):
        self._redraw_plan()
        self._redraw_section()

    def _redraw_plan(self):
        from System.Windows import Point as WpfPoint
        from System.Windows.Media import PointCollection
        from System.Windows.Shapes import Ellipse as WpfEllipse
        from System.Windows.Shapes import Line as WpfLine
        from System.Windows.Shapes import Polygon as WpfPolygon

        cv = self._ui_cv_plan
        if cv is None:
            return
        try:
            cv.Children.Clear()
        except Exception:
            return
        geo = self._preview_geo
        if not geo or not geo.get(u"poly"):
            try:
                if self._ui_txt_header is not None:
                    self._ui_txt_header.Text = u"PLANTA · ZAPATA DE MURO"
            except Exception:
                pass
            return

        poly = geo[u"poly"]
        xs = [float(p[0]) for p in poly]
        ys = [float(p[1]) for p in poly]
        min_x, max_x = min(xs), max(xs)
        min_y, max_y = min(ys), max(ys)
        span_x = max(1.0, max_x - min_x)
        span_y = max(1.0, max_y - min_y)
        cw, ch = self._canvas_size(cv)
        pad = _PLAN_PAD_FRAC
        fit = min(
            (cw * (1.0 - 2.0 * pad)) / span_x,
            (ch * (1.0 - 2.0 * pad)) / span_y,
        )
        fit = max(1e-6, fit)
        scale = fit * max(0.05, float(self._view_zoom))
        cx_mm = 0.5 * (min_x + max_x) + float(self._view_pan_x)
        cy_mm = 0.5 * (min_y + max_y) + float(self._view_pan_y)
        ox = cw / 2.0 - (cx_mm - min_x) * scale
        oy = ch / 2.0 - (max_y - cy_mm) * scale
        self._scene_base = {
            u"min_x": min_x,
            u"max_x": max_x,
            u"min_y": min_y,
            u"max_y": max_y,
            u"ox": ox,
            u"oy": oy,
            u"scale": scale,
        }

        def to_px(xmm, ymm):
            return (
                ox + (float(xmm) - min_x) * scale,
                oy + (max_y - float(ymm)) * scale,
            )

        wp = WpfPolygon()
        pc = PointCollection()
        for xmm, ymm in poly:
            px, py = to_px(xmm, ymm)
            pc.Add(WpfPoint(px, py))
        wp.Points = pc
        wp.Fill = _wf_brush(u"#1a3a4d", 160)
        wp.Stroke = _wf_brush(u"#5BC0DE")
        wp.StrokeThickness = 1.6
        try:
            cv.Children.Add(wp)
        except Exception:
            pass

        d_tr, d_lo, sep_tr, sep_lo = self._read_preview_bar_params()
        length_mm = float(geo.get(u"length_mm") or span_x)
        width_mm = float(geo.get(u"width_mm") or span_y)
        p0 = geo.get(u"p0_mm")
        p1 = geo.get(u"p1_mm")
        rec = float(_REC_OFF_PLANTA_INF_MM)
        if p0 is not None and p1 is not None:
            ax = float(p1[0] - p0[0])
            ay = float(p1[1] - p0[1])
            alen = (ax * ax + ay * ay) ** 0.5
            if alen > 1.0:
                ux, uy = ax / alen, ay / alen
                nx, ny = -uy, ux
                # Transversales: líneas a través del ancho
                for t in _wf_preview_positions_along(length_mm, sep_tr, rec):
                    cx = float(p0[0]) + ux * t
                    cy = float(p0[1]) + uy * t
                    half = 0.5 * max(1.0, width_mm - 2.0 * rec)
                    x0, y0 = cx + nx * half, cy + ny * half
                    x1, y1 = cx - nx * half, cy - ny * half
                    ln = WpfLine()
                    px0, py0 = to_px(x0, y0)
                    px1, py1 = to_px(x1, y1)
                    ln.X1, ln.Y1, ln.X2, ln.Y2 = px0, py0, px1, py1
                    ln.Stroke = _wf_brush(_COLOR_TRANS)
                    ln.StrokeThickness = max(1.2, min(2.8, d_tr * scale * 0.02))
                    try:
                        cv.Children.Add(ln)
                    except Exception:
                        pass
                # Longitudinales: líneas a lo largo del eje
                for s in _wf_preview_positions_along(width_mm, sep_lo, rec):
                    off = s - 0.5 * width_mm
                    x0 = float(p0[0]) + nx * off + ux * rec
                    y0 = float(p0[1]) + ny * off + uy * rec
                    x1 = float(p1[0]) + nx * off - ux * rec
                    y1 = float(p1[1]) + ny * off - uy * rec
                    ln = WpfLine()
                    px0, py0 = to_px(x0, y0)
                    px1, py1 = to_px(x1, y1)
                    ln.X1, ln.Y1, ln.X2, ln.Y2 = px0, py0, px1, py1
                    ln.Stroke = _wf_brush(_COLOR_LONG)
                    ln.StrokeThickness = max(1.0, min(2.4, d_lo * scale * 0.018))
                    try:
                        cv.Children.Add(ln)
                    except Exception:
                        pass

        try:
            if self._ui_txt_header is not None:
                self._ui_txt_header.Text = (
                    u"PLANTA · ZAPATA  ·  {0:.0f} × {1:.0f} mm"
                ).format(length_mm, width_mm)
        except Exception:
            pass

    def _redraw_section(self):
        from System.Windows.Controls import Canvas as WpfCanvas
        from System.Windows.Shapes import Ellipse as WpfEllipse
        from System.Windows.Shapes import Line as WpfLine
        from System.Windows.Shapes import Rectangle as WpfRectangle

        cv = self._ui_cv_section
        if cv is None:
            return
        try:
            cv.Children.Clear()
        except Exception:
            return
        geo = self._preview_geo
        if not geo:
            try:
                if self._ui_txt_section_dims is not None:
                    self._ui_txt_section_dims.Text = u""
            except Exception:
                pass
            return

        w_mm = max(1.0, float(geo.get(u"width_mm") or 600.0))
        h_mm = max(1.0, float(geo.get(u"height_mm") or 500.0))
        cw, ch = self._canvas_size(cv)
        pad = _SECTION_PAD_PX
        usable_w = max(20.0, cw - 2.0 * pad)
        usable_h = max(20.0, ch - 2.0 * pad - 8.0)
        scale = min(usable_w / w_mm, usable_h / h_mm)
        rw = w_mm * scale
        rh = h_mm * scale
        left = (cw - rw) * 0.5
        top = (ch - rh) * 0.5

        rect = WpfRectangle()
        rect.Width = rw
        rect.Height = rh
        rect.Fill = _wf_brush(u"#1a3a4d", 180)
        rect.Stroke = _wf_brush(u"#5BC0DE")
        rect.StrokeThickness = 1.5
        try:
            WpfCanvas.SetLeft(rect, left)
            WpfCanvas.SetTop(rect, top)
            cv.Children.Add(rect)
        except Exception:
            pass

        d_tr, d_lo, _sep_tr, sep_lo = self._read_preview_bar_params()
        rec_h = float(_RECO_HOR_MM)
        rec_lat = float(_REC_LATERAL_CARA_U_MM)
        y_tr = rec_h + 0.5 * d_tr
        x0 = rec_lat + 0.5 * d_tr
        x1 = w_mm - rec_lat - 0.5 * d_tr
        if x1 <= x0:
            x0, x1 = 0.0, w_mm

        def mm_to_px_x(xmm):
            return left + float(xmm) * scale

        def mm_to_px_y(ymm_from_bottom):
            return top + rh - float(ymm_from_bottom) * scale

        def add_line(xa, ya, xb, yb, color, thick):
            ln = WpfLine()
            ln.X1, ln.Y1 = mm_to_px_x(xa), mm_to_px_y(ya)
            ln.X2, ln.Y2 = mm_to_px_x(xb), mm_to_px_y(yb)
            ln.Stroke = _wf_brush(color)
            ln.StrokeThickness = thick
            try:
                cv.Children.Add(ln)
            except Exception:
                pass

        thick_u = max(1.4, min(3.0, d_tr * scale * 0.15))
        add_line(x0, y_tr, x1, y_tr, _COLOR_TRANS, thick_u)
        # Patas U hacia arriba (preview)
        leg = max(40.0, min(h_mm - rec_h - y_tr, float(_DESC_PATA_U_MM)))
        y_leg = min(h_mm - rec_h, y_tr + leg)
        add_line(x0, y_tr, x0, y_leg, _COLOR_TRANS, thick_u)
        add_line(x1, y_tr, x1, y_leg, _COLOR_TRANS, thick_u)

        # Longitudinales como círculos (corte por el ancho)
        y_lo = rec_h + d_tr + 0.5 * d_lo
        for xmm in _wf_preview_positions_along(w_mm, sep_lo, rec_lat + 0.5 * d_lo):
            rpx = max(2.0, min(6.0, 0.5 * d_lo * scale))
            el = WpfEllipse()
            el.Width = 2.0 * rpx
            el.Height = 2.0 * rpx
            el.Fill = _wf_brush(_COLOR_LONG)
            el.Stroke = _wf_brush(_COLOR_LONG)
            try:
                WpfCanvas.SetLeft(el, mm_to_px_x(xmm) - rpx)
                WpfCanvas.SetTop(el, mm_to_px_y(y_lo) - rpx)
                cv.Children.Add(el)
            except Exception:
                pass

        try:
            if self._ui_txt_section_dims is not None:
                self._ui_txt_section_dims.Text = (
                    u"Ancho {0:.0f} mm · Peralte {1:.0f} mm"
                ).format(w_mm, h_mm)
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
        try:
            self._redraw_all()
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
        try:
            self._win.Close()
        except Exception:
            pass

    def _set_estado(self, text):
        tb = self._win.FindName("TxtEstado")
        if tb is not None:
            try:
                tb.Text = text or u""
            except Exception:
                pass

    def _cargar_combos(self):
        doc = self._document
        if doc is None:
            return
        entries, err = _build_bar_type_entries(doc)
        if err:
            try:
                _task_dialog_show(
                    u"Arainco: Wall Foundation Reinforcement",
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
                u"Arainco: Armadura Fundación Corrida",
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
        self._set_estado(u"Seleccione una zapata de muro para ver el esquema.")
        try:
            if not self._win.IsVisible:
                self._win.Show()
            self._win.Activate()
        except Exception:
            pass
        try:
            self._win.UpdateLayout()
        except Exception:
            pass
        try:
            self._redraw_all()
        except Exception:
            pass
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
                u"Arainco: Armadura Fundación Corrida",
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

