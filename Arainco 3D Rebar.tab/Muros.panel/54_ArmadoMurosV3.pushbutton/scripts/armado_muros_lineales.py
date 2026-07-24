# -*- coding: utf-8 -*-
"""
Armado muros — Area Reinforcement rápido (pushbutton autocontenido).

Cada uso:
1. Seleccionar uno o más muros (filtro Wall + eje paralelo al plano de la vista activa).
2. Por muro:
   - ``LocationCurve``: punto medio ``Evaluate(0.5, True)``.
   - Vector tangente desde ``GetEndPoint(1) - GetEndPoint(0)`` (fallback ``ComputeDerivatives``).
   - Plano vertical que **contiene** la tangente y ``XYZ.BasisZ`` (normal ``T × Z``).
   - Intersección del plano con la **geometría sólida** del muro host y, si aplica, muros vecinos
     en extremos (criterio Armado muros nodo: ``armado_muros_vecinos_extremos``).
     Prohibido ``get_BoundingBox`` para el perímetro de creación.

3. ``AreaReinforcement.Create(...)`` sin gancho inicial (``InvalidElementId``).
4. Asigna Rebar Cover por cara: exterior/interior 25 mm, otras caras 0 mm (tipos existentes).
5. Tras crear cada Area Reinforcement, aplica ``RemoveAreaReinforcementSystem`` (rebars físicos).
6. En cada rebar set: horizontales excluyen la última barra (remove last); verticales según cabezal.
7. Post-proceso: verticales ext/int (cabeza, fundación); horizontales ext/int retraídas
   inicio ``25 mm + Ø/2``, fin ``25 mm`` + ``RebarShape`` «06» (int. inicio, ext. fin).
8. Varios muros: orden **de abajo hacia arriba**. Con ``MODO_EJECUCION_RAPIDA`` (defecto): un lote
   global, sin refresco entre lotes; un regenerar/refresco al cerrar el flujo. Si es False: lotes de
   ``MUROS_POR_LOTE_ANIMACION`` muros y aparición progresiva en vista.

Revit 2024+ · IronPython (pyRevit).
"""

from __future__ import print_function

import math
import os
import sys
import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BasePoint,
    BuiltInCategory,
    BuiltInParameter,
    Curve,
    ElementId,
    ElementTypeGroup,
    FilteredElementCollector,
    GeometryInstance,
    IntersectionResultArray,
    Line,
    LocationCurve,
    Options,
    Plane,
    PlanarFace,
    SetComparisonResult,
    Solid,
    Transaction,
    TransactionGroup,
    UnitUtils,
    UnitTypeId,
    ViewDetailLevel,
    Wall,
    XYZ,
)
from Autodesk.Revit.DB.Structure import (
    AreaReinforcement,
    AreaReinforcementType,
    MultiplanarOption,
    Rebar,
    RebarBarType,
    RebarCoverType,
)
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType


def _ensure_pushbutton_path():
    try:
        import bootstrap_paths
        return bootstrap_paths.pin_local_scripts_first()
    except Exception:
        d = os.path.dirname(os.path.abspath(__file__))
        if d and d not in sys.path:
            sys.path.insert(0, d)
        return d


_ensure_pushbutton_path()

try:
    from bimtools_runtime import (
        is_cpython3_runtime,
        is_legacy_revit2024_armado,
        pyrevit_progress_bar_enabled,
        use_transaction_group_armado_muros,
    )
except Exception:
    def is_cpython3_runtime():
        try:
            import sys
            return int(sys.version_info[0]) >= 3
        except Exception:
            return False

    def is_legacy_revit2024_armado(doc):
        return False

    def pyrevit_progress_bar_enabled(doc):
        return False

    def use_transaction_group_armado_muros(doc, within_parent_transaction_group=False):
        return False

try:
    from bimtools_element_id import (
        element_id_from_int,
        element_id_to_int,
        normalize_muro_id_dict,
        normalize_muro_id_key,
        revit_version_year,
        wall_id_int,
    )
except Exception:
    element_id_from_int = None
    element_id_to_int = None
    normalize_muro_id_dict = None
    normalize_muro_id_key = None
    revit_version_year = None
    wall_id_int = None

try:
    from armado_muros_nodo_shared import ajustar_inclusion_extremos_rebar_set_con_fallback
except Exception:
    ajustar_inclusion_extremos_rebar_set_con_fallback = None

try:
    from arearein_verticales_empotramiento_rps import _rebar_es_vertical_por_criterio
except Exception:
    _rebar_es_vertical_por_criterio = None

try:
    import armado_muros_vecinos_extremos as _vecinos_extremos_mod
except Exception:
    _vecinos_extremos_mod = None

try:
    import armado_muros_malla_rebar_tags as _malla_rebar_tags_mod
except Exception as _ex_malla_tags_imp:
    _malla_rebar_tags_mod = None
    _malla_rebar_tags_import_error = str(_ex_malla_tags_imp)
else:
    _malla_rebar_tags_import_error = None


def _reload_malla_rebar_tags_mod():
    """Asegura el módulo de tags cargado (sin importlib.reload — coste alto)."""
    global _malla_rebar_tags_mod, _malla_rebar_tags_import_error
    if _malla_rebar_tags_mod is not None:
        return True
    try:
        import armado_muros_malla_rebar_tags as _mt
        _malla_rebar_tags_mod = _mt
        _malla_rebar_tags_import_error = None
        return True
    except Exception as _ex_reload:
        _malla_rebar_tags_mod = None
        _malla_rebar_tags_import_error = str(_ex_reload)
        return False


def _stamp_malla_pre_etiquetar_lote(
    doc,
    lote,
    rebars_por_muro_id,
    params_por_muro_id,
    muro_contencion,
    rebars_horizontal_por_muro_id,
):
    """Sincroniza ``Armadura_Malla_Orientacion`` H./V. antes de etiquetar (en lineales)."""
    if doc is None or not lote or not rebars_por_muro_id:
        return
    rebars_lote = {}
    for wall in lote:
        try:
            wid = _wall_id_int(wall)
        except Exception:
            continue
        eids = rebars_por_muro_id.get(wid)
        if eids:
            rebars_lote[wid] = list(eids)
    if not rebars_lote:
        return
    t_pre = Transaction(
        doc, u"Arainco: Armado muros lineales — parámetros malla pre-etiqueta",
    )
    try:
        from armado_muros_txn import attach_rebar_outside_host_swallower
        attach_rebar_outside_host_swallower(t_pre)
    except Exception:
        pass
    try:
        t_pre.Start()
        _stamp_malla_params_rebars_por_muro(
            doc,
            rebars_lote,
            params_por_muro_id=params_por_muro_id,
            muro_contencion=muro_contencion,
            rebars_horizontal_por_muro_id=rebars_horizontal_por_muro_id,
        )
        _aplicar_spacing_malla_rebars_por_muro(
            doc,
            rebars_lote,
            params_por_muro_id=params_por_muro_id,
            muro_contencion=muro_contencion,
            rebars_horizontal_por_muro_id=rebars_horizontal_por_muro_id,
        )
        t_pre.Commit()
    except Exception:
        if t_pre.HasStarted():
            try:
                t_pre.RollBack()
            except Exception:
                pass


def _crear_tags_rebar_malla_tras_commit_lote(
    doc,
    view,
    lote,
    params_por_muro_id,
    rebars_por_muro_id,
    tags_rebar_malla,
    errores,
    muro_contencion=False,
    rebars_horizontal_por_muro_id=None,
    stamp_pre_etiqueta=True,
):
    _reload_malla_rebar_tags_mod()
    if _malla_rebar_tags_mod is None or view is None or tags_rebar_malla is None or not lote:
        return
    if stamp_pre_etiqueta:
        _stamp_malla_pre_etiquetar_lote(
            doc,
            lote,
            rebars_por_muro_id,
            params_por_muro_id,
            muro_contencion,
            rebars_horizontal_por_muro_id,
        )
    wall_by_id = {}
    for wall in lote:
        try:
            wall_by_id[_wall_id_int(wall)] = wall
        except Exception:
            pass
    for wall in lote:
        try:
            wid = _wall_id_int(wall)
        except Exception:
            continue
        eids = rebars_por_muro_id.get(wid)
        if not eids:
            continue
        try:
            tag_res = _malla_rebar_tags_mod.etiquetar_rebars_malla_en_vista(
                doc,
                view,
                {wid: list(eids)},
                params_por_muro_id=params_por_muro_id,
                muro_contencion=muro_contencion,
                walls_by_id=wall_by_id,
                rebars_horizontal_por_muro_id=rebars_horizontal_por_muro_id,
            )
            tags_rebar_malla[0] += int(tag_res.get(u"n_ok", 0))
            tags_rebar_malla[1] += int(tag_res.get(u"n_fail", 0))
            tags_rebar_malla[2] += int(tag_res.get(u"n_skip", 0))
            for msg in tag_res.get(u"messages") or []:
                if msg:
                    errores.append(u"Muro {}: {}.".format(wid, msg))
        except Exception as ex_tag:
            errores.append(
                u"Muro {}: etiquetas rebar malla — {}.".format(wid, str(ex_tag)),
            )


def _view_type_suffix_safe(view):
    try:
        from armado_muros_etiqueta_malla import _view_type_suffix
        return _view_type_suffix(view)
    except Exception:
        pass
    try:
        return str(view.ViewType)
    except Exception:
        return u"?"


def _resumen_etiquetas_malla_desde_contadores(tags_rebar_malla, messages=None):
    src = {
        u"n_tags_rebar_malla": int(tags_rebar_malla[0]),
        u"n_tags_rebar_malla_fail": int(tags_rebar_malla[1]),
        u"n_tags_rebar_malla_skip": int(tags_rebar_malla[2]),
    }
    if messages:
        src[u"messages"] = list(messages)
    return src


def _aplicar_etiquetas_malla_todos(
    doc, uidoc, walls_ord, params_por_muro_id, rebars_por_muro_id, errores, embed_resumen,
    cover_asignados=0, muro_contencion=False, stamp_pre_etiqueta=True,
):
    """Structural Rebar Tag de malla (``EST_A_STRUCTURAL REBAR TAG_MALLA``).

    ``stamp_pre_etiqueta=False`` en unificado: el post-lote ya stampó/spacing.
    """
    tags_rebar_malla = [0, 0, 0]
    tag_msgs = []
    _reload_malla_rebar_tags_mod()
    view_etiqueta = None
    vista_rebar_tag_ok = False
    if uidoc is not None:
        try:
            view_etiqueta = uidoc.ActiveView
            if _malla_rebar_tags_mod is not None:
                vista_rebar_tag_ok = _malla_rebar_tags_mod._vista_permite_tags_malla(
                    view_etiqueta,
                )
        except Exception:
            view_etiqueta = None
    if vista_rebar_tag_ok and view_etiqueta is not None and rebars_por_muro_id:
        horiz_map = {}
        if embed_resumen:
            horiz_map = embed_resumen.get(u"rebars_malla_horizontal_por_muro_id") or {}
        batch = _tamano_lote_ejecucion(len(walls_ord), MUROS_POR_LOTE_ANIMACION)
        for i0 in range(0, len(walls_ord), batch):
            lote = walls_ord[i0:i0 + batch]
            _crear_tags_rebar_malla_tras_commit_lote(
                doc,
                view_etiqueta,
                lote,
                params_por_muro_id,
                rebars_por_muro_id,
                tags_rebar_malla,
                errores,
                muro_contencion=muro_contencion,
                rebars_horizontal_por_muro_id=horiz_map,
                stamp_pre_etiqueta=stamp_pre_etiqueta,
            )
            _refrescar_vista_tras_lote(doc, uidoc)
    n_tag_rb = int(tags_rebar_malla[0])
    n_tag_fail = int(tags_rebar_malla[1])
    n_tag_skip = int(tags_rebar_malla[2])
    if rebars_por_muro_id and vista_rebar_tag_ok:
        return _merge_embed_resumen(
            embed_resumen,
            _resumen_etiquetas_malla_desde_contadores(tags_rebar_malla, tag_msgs),
        )
    if int(cover_asignados) > 0 and rebars_por_muro_id and not vista_rebar_tag_ok:
        vt_msg = _view_type_suffix_safe(view_etiqueta) if view_etiqueta else u"?"
        msg_v = (
            u"Etiquetas rebar malla: ninguna (vista activa: {}; use planta, alzado o sección).".format(
                vt_msg,
            )
        )
        if errores is not None:
            errores.append(msg_v)
        return _merge_embed_resumen(
            embed_resumen,
            {u"messages": [msg_v]},
        )
    if rebars_por_muro_id and not vista_rebar_tag_ok:
        msg_v = u"Etiquetas rebar malla: use planta, alzado o sección (no plantilla ni 3D)."
        if errores is not None:
            errores.append(msg_v)
        return _merge_embed_resumen(
            embed_resumen,
            {u"messages": [msg_v]},
        )
    if _malla_rebar_tags_mod is None and _malla_rebar_tags_import_error:
        errores.append(
            u"Etiquetas rebar malla: módulo no cargado — {}.".format(
                _malla_rebar_tags_import_error,
            ),
        )
    return embed_resumen


def _rebar_element_ids_lista_desde_mapa_por_muro(rebars_por_muro_id):
    """``ElementId`` únicos de rebars agrupados por muro."""
    from System.Collections.Generic import List as ClrList

    ids_out = ClrList[ElementId]()
    vistos = set()
    for _wid, id_list in (rebars_por_muro_id or {}).items():
        for eid in id_list or []:
            try:
                if eid is None:
                    continue
                if not isinstance(eid, ElementId):
                    if element_id_from_int is not None:
                        eid = element_id_from_int(eid)
                    else:
                        eid = ElementId(int(eid))
                if eid == ElementId.InvalidElementId:
                    continue
                k = _element_id_int(eid)
                if k in vistos:
                    continue
                vistos.add(k)
                ids_out.Add(eid)
            except Exception:
                continue
    return ids_out


def _rebars_desde_mapa_por_muro(doc, rebars_por_muro_id):
    """Elementos ``Rebar`` únicos a partir de ``rebars_por_muro_id``."""
    if doc is None or not rebars_por_muro_id:
        return []
    rebars = []
    vistos = set()
    for _wid, id_list in rebars_por_muro_id.items():
        for eid in id_list or []:
            k = _element_id_int(eid)
            if k is None or k in vistos:
                continue
            vistos.add(k)
            try:
                el = doc.GetElement(eid)
                if el is not None:
                    rebars.append(el)
            except Exception:
                pass
    return rebars


def _unhide_rebars_malla_en_vista(view, rebars_por_muro_id):
    """Revoca ``Hide in View → Elements`` sobre rebars de malla. Sin transacción."""
    if view is None or not rebars_por_muro_id:
        return 0
    ids_show = _rebar_element_ids_lista_desde_mapa_por_muro(rebars_por_muro_id)
    n_show = int(ids_show.Count)
    if n_show < 1:
        return 0
    view.UnhideElements(ids_show)
    return n_show


def _ids_rebars_malla_excluir_unobscured(rebars_por_muro_id):
    """Ids numéricos de rebars generados por malla (Remove Area Reinforcement System)."""
    ids = set()
    for _wid, id_list in (rebars_por_muro_id or {}).items():
        for eid in id_list or []:
            k = _element_id_int(eid)
            if k is not None:
                ids.add(k)
    return ids


def _agregar_rebar_id_si_aplica(out, seen, eid, ids_malla_excluir):
    k = _element_id_int(eid)
    if k is None or k in seen or k in ids_malla_excluir:
        return
    seen.add(k)
    out.append(eid)


def _recolectar_rebar_ids_unobscured_armado_muros(
    cab_res=None,
    embed_resumen=None,
    cor_res=None,
    ids_malla_excluir=None,
):
    """
    Rebars creados por Armado Muros (cabezal, confinamiento, coronamiento),
    excluyendo las asociadas a malla.
    """
    out = []
    seen = set()
    excl = ids_malla_excluir or set()

    if cab_res:
        for eid in cab_res.get(u"rebars_longitudinales_ids") or []:
            _agregar_rebar_id_si_aplica(out, seen, eid, excl)
        for _wid, lst in (cab_res.get(u"rebars_por_muro_id") or {}).items():
            for eid in lst or []:
                _agregar_rebar_id_si_aplica(out, seen, eid, excl)

    for src in (embed_resumen, cor_res):
        if not src:
            continue
        for eid in src.get(u"rebars_coronamiento_ids") or []:
            _agregar_rebar_id_si_aplica(out, seen, eid, excl)
        for iv in src.get(u"rebars_coronamiento_id_ints") or []:
            try:
                eid = ElementId(int(iv))
            except Exception:
                continue
            _agregar_rebar_id_si_aplica(out, seen, eid, excl)

    return out


def aplicar_unobscured_armado_muros_en_vista(
    doc,
    uidoc,
    cab_res=None,
    embed_resumen=None,
    cor_res=None,
    rebars_malla_por_muro_id=None,
    errores=None,
):
    """
    Visibilidad en la **vista activa** al ejecutar la herramienta:

    - Barras de **malla**: visibles (``UnhideElements``) y ``SetUnobscuredInView(False)``.
    - **Cabezal / confinamiento / coronamiento**: Unhide + ``SetUnobscuredInView(True)``
      (+ sólido). Crítico para coronamiento en pie/zapata oculto por el sólido.
    """
    vacio = {
        u"n_rebars_unobscured_on": 0,
        u"n_rebars_unobscured_off": 0,
        u"n_rebars_malla_visibles": 0,
    }
    if doc is None or uidoc is None:
        return vacio
    view = None
    try:
        view = uidoc.ActiveView
    except Exception:
        view = None
    if view is None:
        return vacio
    try:
        if getattr(view, "IsTemplate", False):
            return vacio
    except Exception:
        pass

    try:
        from bimtools_rebar_3d_visibility import (
            apply_rebar_unobscured_in_view,
            apply_rebar_unobscured_off_in_view,
        )
    except Exception as ex_imp:
        if errores is not None:
            errores.append(
                u"Visibilidad Armado Muros: módulo no disponible — {0}".format(ex_imp),
            )
        return vacio

    ids_malla = _ids_rebars_malla_excluir_unobscured(rebars_malla_por_muro_id)
    rebar_ids_on = _recolectar_rebar_ids_unobscured_armado_muros(
        cab_res=cab_res,
        embed_resumen=embed_resumen,
        cor_res=cor_res,
        ids_malla_excluir=ids_malla,
    )
    rebars_on = []
    ids_on_list = List[ElementId]()
    for eid in rebar_ids_on:
        try:
            el = doc.GetElement(eid)
            if el is not None:
                rebars_on.append(el)
                try:
                    ids_on_list.Add(el.Id)
                except Exception:
                    pass
        except Exception:
            pass
    rebars_malla = _rebars_desde_mapa_por_muro(doc, rebars_malla_por_muro_id)

    if not rebars_on and not rebars_malla:
        return vacio

    t = Transaction(doc, u"Arainco: Visibilidad barras Armado Muros en vista")
    n_on = 0
    n_off = 0
    n_vis = 0
    try:
        t.Start()
        try:
            from Autodesk.Revit.DB import BuiltInCategory, Category

            cat = Category.GetCategory(doc, BuiltInCategory.OST_Rebar)
            if cat is not None:
                view.SetCategoryHidden(cat.Id, False)
        except Exception:
            pass
        if rebars_malla:
            n_vis = _unhide_rebars_malla_en_vista(view, rebars_malla_por_muro_id)
            n_off = apply_rebar_unobscured_off_in_view(doc, rebars_malla, view)
        if rebars_on:
            if ids_on_list.Count > 0:
                try:
                    view.UnhideElements(ids_on_list)
                except Exception:
                    pass
            try:
                from Autodesk.Revit.DB.Structure import RebarPresentationMode

                for rb in rebars_on:
                    try:
                        rb.SetPresentationMode(view, RebarPresentationMode.All)
                    except Exception:
                        pass
            except Exception:
                pass
            n_on = apply_rebar_unobscured_in_view(doc, rebars_on, view)
        t.Commit()
    except Exception as ex:
        try:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
        except Exception:
            pass
        if errores is not None:
            try:
                msg = u"Visibilidad Armado Muros en vista: {0}".format(unicode(ex))
            except Exception:
                msg = u"Visibilidad Armado Muros en vista: {0}".format(str(ex))
            errores.append(msg)
        return vacio

    return {
        u"n_rebars_unobscured_on": int(n_on),
        u"n_rebars_unobscured_off": int(n_off),
        u"n_rebars_malla_visibles": int(n_vis),
    }


# ── Tolerancias (pies internos salvo uso explícito) ──────────────────────────
_MM_TOL = 2.5
try:
    _TOL_PT_MATCH_FT = UnitUtils.ConvertToInternalUnits(_MM_TOL, UnitTypeId.Millimeters)
except Exception:
    _TOL_PT_MATCH_FT = 0.01

_PLANE_LINE_DIST_TOL_FT = max(1e-5, UnitUtils.ConvertToInternalUnits(0.75, UnitTypeId.Millimeters))

# Rebar Cover por cara (mm) antes de crear Area Reinforcement.
REBAR_COVER_MM_CARAS_EXT_INT = 25.0
REBAR_COVER_MM_OTRAS_CARAS = 0.0
# True: un lote por fase, sin refresco de vista entre lotes (menor carga en Revit).
MODO_EJECUCION_RAPIDA = True
# Muros por lote si MODO_EJECUCION_RAPIDA es False (1 = aparición progresiva abajo→arriba).
MUROS_POR_LOTE_ANIMACION = 1

_ML_PBAR_BASE = u"Arainco: Armado muros lineales"


def _tamano_lote_ejecucion(n_items, lote_legacy):
    """Tamaño de lote: todo el conjunto en modo rápido; si no, ``max(1, lote_legacy)``."""
    try:
        n = int(n_items)
    except Exception:
        n = 0
    if n < 1:
        return 1
    if MODO_EJECUCION_RAPIDA:
        return n
    try:
        return max(1, int(lote_legacy))
    except Exception:
        return 1
# Un solo paso Deshacer para cabezal + mallas + etiquetas.
TXN_GROUP_ARMADO_MUROS_UNIFICADO = u"Arainco: Armado Muros v3"


def _ml_pbar_phase_title(base_title, total):
    """Título inicial 0/N (mismo patrón que Armado columnas / exportar láminas)."""
    try:
        t = max(int(total), 1)
    except Exception:
        t = 1
    return u"{} 0/{}".format(base_title, t)


def _ml_pbar_enabled(doc):
    try:
        return bool(pyrevit_progress_bar_enabled(doc))
    except Exception:
        return False


def _ml_pbar_start(title, count, doc=None):
    u"""
    ``forms.ProgressBar`` de pyRevit con acento BIMTools (91,192,222),
    igual que ``column_reinforcement_layout_rps._column_layout_pbar_start``.
    """
    if not _ml_pbar_enabled(doc):
        return None
    if count is None or int(count) < 1:
        return None
    try:
        from pyrevit import forms as _pyrevit_forms

        pb = _pyrevit_forms.ProgressBar(
            title=title,
            cancellable=False,
        )
        try:
            from System.Windows.Media import Color, SolidColorBrush

            r, g, b = (91, 192, 222)
            pb.Resources[u"pyRevitAccentBrush"] = SolidColorBrush(Color.FromRgb(r, g, b))
        except Exception:
            pass
        return pb
    except Exception:
        return None


def _ml_pbar_step(pb, current_index, count, base_title):
    u"""*current_index*: 0…count-1. Actualiza barra y título con contador."""
    if pb is None:
        return
    c = int(count) if count else 0
    if c < 1:
        c = 1
    i = int(current_index) + 1
    try:
        if hasattr(pb, u"update_progress"):
            try:
                pb.update_progress(i, max_value=c)
            except TypeError:
                try:
                    pb.update_progress(i, max=c)
                except Exception:
                    pass
    except Exception:
        pass
    try:
        pb.title = u"{} {}/{}".format(base_title, i, c)
    except Exception:
        pass


def _ml_pbar_enter(pb):
    if pb is None:
        return False
    try:
        pb.__enter__()
        return True
    except Exception:
        return False


def _ml_pbar_exit(pb, pbar_open):
    if not pbar_open or pb is None:
        return
    try:
        pb.__exit__(None, None, None)
    except Exception:
        pass


def _embed_resumen_vacio():
    return {
        u"n_extended": 0,
        u"n_retracted": 0,
        u"n_pata_l": 0,
        u"n_pata_l_fund_pie": 0,
        u"n_fundacion_pie": 0,
        u"n_fundacion_retract": 0,
        u"n_pie_muro_colision_revert": 0,
        u"n_pie_muro_retract": 0,
        u"n_pie_muro_pata_l": 0,
        u"n_pie_muro_pata_l_ext": 0,
        u"n_pie_muro_pata_l_int": 0,
        u"n_horiz_retract": 0,
        u"n_horiz_retract_ext": 0,
        u"n_horiz_retract_int": 0,
        u"n_horiz_pata_l": 0,
        u"n_horiz_pata_l_ext": 0,
        u"n_horiz_pata_l_int": 0,
        u"n_cabezal": 0,
        u"n_cabezal_fail": 0,
        u"n_tags_rebar_malla": 0,
        u"n_tags_rebar_malla_fail": 0,
        u"n_tags_rebar_malla_skip": 0,
        u"n_skip": 0,
        u"n_fail": 0,
        u"messages": [],
    }


def _merge_rebar_id_lists_por_muro(dest_map, src_map):
    """Une ``{wall_id: [rebar_id_int, ...]}`` sin duplicar."""
    out = dict(dest_map or {})
    for wid, lst in (src_map or {}).items():
        try:
            wkey = int(wid)
        except Exception:
            continue
        seen = set()
        merged = []
        for eid in list(out.get(wkey) or []) + list(lst or []):
            try:
                ei = int(eid)
            except Exception:
                continue
            if ei in seen:
                continue
            seen.add(ei)
            merged.append(ei)
        if merged:
            out[wkey] = merged
    return out


def _merge_embed_resumen(dest, src):
    if dest is None:
        dest = _embed_resumen_vacio()
    if not src:
        return dest
    for k in (
        u"n_extended",
        u"n_retracted",
        u"n_pata_l",
        u"n_pata_l_fund_pie",
        u"n_fundacion_pie",
        u"n_fundacion_retract",
        u"n_pie_muro_colision_revert",
        u"n_pie_muro_retract",
        u"n_pie_muro_pata_l",
        u"n_pie_muro_pata_l_ext",
        u"n_pie_muro_pata_l_int",
        u"n_horiz_retract",
        u"n_horiz_retract_ext",
        u"n_horiz_retract_int",
        u"n_horiz_pata_l",
        u"n_horiz_pata_l_ext",
        u"n_horiz_pata_l_int",
        u"n_cabezal",
        u"n_cabezal_fail",
        u"n_coronamiento",
        u"n_coronamiento_fail",
        u"n_coronamiento_bars",
        u"n_coronamiento_inferior",
        u"n_coronamiento_inferior_fail",
        u"n_coronamiento_inferior_bars",
        u"n_coronamiento_inferior_pie",
        u"n_coronamiento_inferior_pie_fail",
        u"n_coronamiento_inferior_pie_bars",
        u"n_coronamiento_voladizo",
        u"n_coronamiento_voladizo_fail",
        u"n_coronamiento_voladizo_bars",
        u"n_coronamiento_tags",
        u"n_coronamiento_tags_fail",
        u"n_tags_rebar_malla",
        u"n_tags_rebar_malla_fail",
        u"n_tags_rebar_malla_skip",
        u"n_skip",
        u"n_fail",
    ):
        dest[k] = int(dest.get(k, 0)) + int(src.get(k, 0))
    rebars_src = src.get(u"rebars_por_muro_id")
    if rebars_src:
        rebars_dest = dict(dest.get(u"rebars_por_muro_id") or {})
        rebars_dest.update(rebars_src)
        dest[u"rebars_por_muro_id"] = rebars_dest
    horiz_src = src.get(u"rebars_malla_horizontal_por_muro_id")
    if horiz_src:
        dest[u"rebars_malla_horizontal_por_muro_id"] = _merge_rebar_id_lists_por_muro(
            dest.get(u"rebars_malla_horizontal_por_muro_id"),
            horiz_src,
        )
    dest.setdefault(u"messages", [])
    dest[u"messages"].extend(src.get(u"messages") or [])
    dest[u"rebars_coronamiento_ids"] = _merge_rebar_element_id_lists(
        dest.get(u"rebars_coronamiento_ids"),
        src.get(u"rebars_coronamiento_ids"),
    )
    meta_dest = list(dest.get(u"rebars_coronamiento_tag_meta") or [])
    meta_src = list(src.get(u"rebars_coronamiento_tag_meta") or [])
    if meta_src:
        meta_dest.extend(meta_src)
        dest[u"rebars_coronamiento_tag_meta"] = meta_dest
    ints_dest = list(dest.get(u"rebars_coronamiento_id_ints") or [])
    ints_src = list(src.get(u"rebars_coronamiento_id_ints") or [])
    if ints_src:
        seen_i = set(ints_dest)
        for iv in ints_src:
            try:
                n = int(iv)
            except Exception:
                continue
            if n in seen_i:
                continue
            seen_i.add(n)
            ints_dest.append(n)
        dest[u"rebars_coronamiento_id_ints"] = ints_dest
    return dest


def _merge_rebar_element_id_lists(dest_list, src_list):
    """Une listas de ``ElementId`` de rebar sin duplicar."""
    out = list(dest_list or [])
    seen = set()
    for rid in out:
        try:
            seen.add(_element_id_int(rid))
        except Exception:
            pass
    for rid in src_list or []:
        if rid is None:
            continue
        try:
            iv = _element_id_int(rid)
        except Exception:
            continue
        if iv in seen:
            continue
        seen.add(iv)
        out.append(rid)
    return out


def _params_dict_for_wall_id(params_por_muro_id, wid):
    """Resuelve ``params_dict`` tolerando claves int/str."""
    if not params_por_muro_id or wid is None:
        return None
    keys = [wid]
    try:
        keys.append(int(wid))
    except Exception:
        pass
    try:
        keys.append(unicode(wid))
    except Exception:
        try:
            keys.append(str(wid))
        except Exception:
            pass
    seen = set()
    for key in keys:
        if key in seen:
            continue
        seen.add(key)
        try:
            tup = params_por_muro_id.get(key)
            if tup is not None:
                return tup[0]
        except Exception:
            continue
    return None


def _post_procesar_rebars_lote(
    doc,
    walls_ord,
    rebars_lote,
    params_por_muro_id=None,
    muro_contencion=False,
    cabezal_por_muro_id=None,
    fund_solids_cache=None,
):
    """
    Post-proceso de un lote de muros: verticales (empotramiento/patas L), horizontales
    (retraída/pata L) y exclusión final de extremos del set. Modifica ``rebars_lote`` in-place.

    Una Transaction de documento: colisión de malla usa caché de sólidos (sonda antes
    de mutar); mutaciones anidan SubTransaction. Exclusión + stamp al final.

    Regenerates: 1 tras mutaciones geométricas (antes de exclusión) + 1 al cierre
    (tras exclusión/spacing). Sin regenerates intermedios por exclusión/spacing.

    ``fund_solids_cache``: dict compartido entre lotes ``{fund_id: solids}``.
    """
    res = _embed_resumen_vacio()
    if doc is None or not rebars_lote:
        return res

    def _run_post():
        holder = {u"res": res}
        need_final_regen = False

        try:
            import armado_muros_verticales_embed_colision as _embed_mod

            holder[u"res"] = _merge_embed_resumen(
                holder[u"res"],
                _embed_mod.aplicar_empotramiento_verticales_cara_por_colision(
                    doc,
                    walls_ord,
                    rebars_lote,
                    params_por_muro_id=params_por_muro_id,
                    muro_contencion=muro_contencion,
                    cabezal_por_muro_id=cabezal_por_muro_id,
                    # Malla lote ya regeneró antes de Remove System.
                    regenerate_fund_geom=False,
                    fund_solids_cache=fund_solids_cache,
                ),
            )
        except Exception as ex_embed:
            holder[u"res"].setdefault(u"messages", [])
            holder[u"res"][u"messages"].append(
                u"Empotramiento vertical ext/int: {0}".format(str(ex_embed)),
            )

        try:
            import armado_muros_horizontales_retraida as _horiz_mod

            holder[u"res"] = _merge_embed_resumen(
                holder[u"res"],
                _horiz_mod.aplicar_retraida_horizontales_ext_int(
                    doc, walls_ord, rebars_lote,
                ),
            )
        except Exception as ex_horiz:
            holder[u"res"].setdefault(u"messages", [])
            holder[u"res"][u"messages"].append(
                u"Retraída horizontal ext/int: {0}".format(str(ex_horiz)),
            )

        ids_finales = []
        for eids in rebars_lote.values():
            ids_finales.extend(eids or [])
        if ids_finales:
            try:
                doc.Regenerate()
            except Exception:
                pass
            try:
                _desactivar_extremos_rebars_creados(
                    doc,
                    ids_finales,
                    cabezal_por_muro_id=cabezal_por_muro_id,
                    params_por_muro_id=params_por_muro_id,
                    muro_contencion=muro_contencion,
                    regenerate_each=False,
                    regenerate_after=False,
                )
                need_final_regen = True
            except Exception:
                pass

        try:
            from armado_muros_rebar_params import activar_armadura_arainco_por_ids

            for eids in (rebars_lote or {}).values():
                activar_armadura_arainco_por_ids(doc, eids)
            _stamp_malla_params_rebars_por_muro(
                doc,
                rebars_lote,
                params_por_muro_id=params_por_muro_id,
                muro_contencion=muro_contencion,
                rebars_horizontal_por_muro_id=holder[u"res"].get(
                    u"rebars_malla_horizontal_por_muro_id",
                ),
            )
            n_sp = _aplicar_spacing_malla_rebars_por_muro(
                doc,
                rebars_lote,
                params_por_muro_id=params_por_muro_id,
                muro_contencion=muro_contencion,
                rebars_horizontal_por_muro_id=holder[u"res"].get(
                    u"rebars_malla_horizontal_por_muro_id",
                ),
                regenerate_after=False,
            )
            if n_sp:
                need_final_regen = True
        except Exception:
            pass

        if need_final_regen and doc is not None:
            try:
                doc.Regenerate()
            except Exception:
                pass

        return holder[u"res"]

    try:
        from armado_muros_txn import run_in_transaction

        return run_in_transaction(
            doc,
            u"Arainco: Armado muros lineales — post lote",
            _run_post,
        )
    except Exception:
        return _run_post()


# ── Tipos ───────────────────────────────────────────────────────────────────
def _safe_volume(solid):
    try:
        return float(solid.Volume)
    except Exception:
        return None


# ── ElementId / tipo AR ───────────────────────────────────────────────────────
def _element_id_int(eid):
    if element_id_to_int is not None:
        return element_id_to_int(eid)
    if eid is None:
        return None
    try:
        return int(eid.Value)
    except Exception:
        pass
    try:
        return int(eid.IntegerValue)
    except Exception:
        pass
    try:
        return int(eid)
    except Exception:
        return None


def _wall_id_int(wall):
    if wall_id_int is not None:
        return wall_id_int(wall)
    return _element_id_int(getattr(wall, "Id", None))


def _normalize_muro_id_key(key):
    if normalize_muro_id_key is not None:
        return normalize_muro_id_key(key)
    try:
        return int(key)
    except Exception:
        return _element_id_int(key)


def _normalize_muro_id_dict(mapping):
    if normalize_muro_id_dict is not None:
        return normalize_muro_id_dict(mapping)
    return mapping or {}


def _revit_year(doc):
    if revit_version_year is not None:
        return revit_version_year(doc)
    return 0


def _unificado_usa_transaction_group_externo(doc):
    """
    Revit 2025+ / CPython 3: el ``TransactionGroup`` padre impide cabezal/mallas.
    Revit 2024 (IronPython 2) mantiene un solo Deshacer.
    """
    return use_transaction_group_armado_muros(doc, within_parent_transaction_group=False)


def _is_area_reinforcement_type(elem):
    if elem is None:
        return False
    try:
        if isinstance(elem, AreaReinforcementType):
            return True
    except Exception:
        pass
    try:
        tn = getattr(elem, "GetType", None)
        if tn and callable(tn):
            tinfo = tn()
            name = getattr(tinfo, "Name", "") or getattr(tinfo, "FullName", "") or str(tinfo)
            if "AreaReinforcementType" in str(name):
                return True
    except Exception:
        pass
    return False


def _get_default_area_reinforcement_type_id(document):
    def _valid_id(eid):
        if not eid or eid == ElementId.InvalidElementId:
            return False
        n = _element_id_int(eid)
        return n is not None and n >= 0

    try:
        for elem in FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_AreaRein):
            if elem is None:
                continue
            try:
                if hasattr(elem, "GetTypeId"):
                    tid = elem.GetTypeId()
                    if _valid_id(tid):
                        t = document.GetElement(tid)
                        if t and _is_area_reinforcement_type(t):
                            return tid
            except Exception:
                continue
    except Exception:
        pass

    try:
        default_id = document.GetDefaultElementTypeId(ElementTypeGroup.AreaReinforcementType)
        if _valid_id(default_id):
            t = document.GetElement(default_id)
            if t and _is_area_reinforcement_type(t):
                return default_id
            if t is not None:
                return default_id
    except Exception:
        pass

    try:
        for elem in (FilteredElementCollector(document)
                     .OfCategory(BuiltInCategory.OST_AreaRein)
                     .WhereElementIsElementType()):
            if elem and _is_area_reinforcement_type(elem):
                return elem.Id
    except Exception:
        pass

    try:
        for elem in FilteredElementCollector(document).OfClass(AreaReinforcementType):
            if elem and _is_area_reinforcement_type(elem):
                return elem.Id
    except Exception:
        try:
            rt = clr.GetClrType(AreaReinforcementType)
            for elem in FilteredElementCollector(document).OfClass(rt):
                if elem and _is_area_reinforcement_type(elem):
                    return elem.Id
        except Exception:
            pass

    return None


def _primer_bar_type_id(doc):
    try:
        for elem in FilteredElementCollector(doc).OfClass(RebarBarType):
            if elem:
                return elem.Id
    except Exception:
        pass
    return None


# ── Location + plano ─────────────────────────────────────────────────────────
def location_curve_wall(wall):
    if wall is None or not isinstance(wall, Wall):
        return None
    loc = wall.Location
    if not isinstance(loc, LocationCurve):
        return None
    return loc.Curve


def _tangente_ui_desde_location_curve(curve):
    """Vector unitario desde P1 − P0; respaldo ComputeDerivatives(0.5)."""
    if curve is None:
        return None
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        v = p1.Subtract(p0)
        if float(v.GetLength()) > 1e-12:
            return v.Normalize()
    except Exception:
        pass
    try:
        dv = curve.ComputeDerivatives(0.5, True)
        tx = getattr(dv, "BasisX", None)
        if tx is not None and float(tx.GetLength()) > 1e-12:
            return tx.Normalize()
    except Exception:
        pass
    return None


def _punto_centro_location_curve(curve):
    if curve is None:
        return None
    try:
        p = curve.Evaluate(0.5, True)
        if p is not None:
            return XYZ(float(p.X), float(p.Y), float(p.Z))
    except Exception:
        pass
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        return XYZ(
            0.5 * (float(p0.X) + float(p1.X)),
            0.5 * (float(p0.Y) + float(p1.Y)),
            0.5 * (float(p0.Z) + float(p1.Z)),
        )
    except Exception:
        return None


def plano_vertical_contiene_location_y_vertical_global(wall):
    """
    Plano definido por: punto medio de LocationCurve + vector tangente desde la Location.
    Interpretación física Revit (con el vector como **primer eje contenido en el plano**):
    contiene tangente ``T`` y la vertical ``BasisZ``. Normal ``N = T × Z``.
    """
    lc = location_curve_wall(wall)
    if lc is None:
        return None
    p_mid = _punto_centro_location_curve(lc)
    if p_mid is None:
        return None
    t = _tangente_ui_desde_location_curve(lc)
    if t is None:
        return None
    z_up = XYZ.BasisZ
    n = t.CrossProduct(z_up)
    ln = float(n.GetLength())
    if ln < 1e-12:
        n = t.CrossProduct(XYZ.BasisX)
        ln = float(n.GetLength())
    if ln < 1e-12:
        return None
    n = n.Multiply(1.0 / ln).Normalize()
    try:
        return Plane.CreateByNormalAndOrigin(n, p_mid), lc
    except Exception:
        return None


# ── Sólidos muro ─────────────────────────────────────────────────────────────
def _make_geometry_options_wall():
    opts = Options()
    try:
        opts.ComputeReferences = False
    except Exception:
        pass
    try:
        opts.DetailLevel = ViewDetailLevel.Fine
    except Exception:
        pass
    try:
        opts.IncludeNonVisibleObjects = True
    except Exception:
        pass
    return opts


def _iter_solids_wall(element):
    if element is None:
        return
    try:
        from bimtools_clr_collections import iterate_net_collection, safe_solid_volume
    except Exception:
        iterate_net_collection = None
        safe_solid_volume = None

    def _vol_ok(solid):
        if solid is None:
            return False
        if safe_solid_volume is not None:
            v = safe_solid_volume(solid)
            return v is not None and abs(v) > 1e-12
        v = _safe_volume(solid)
        return v is not None and abs(v) > 1e-12

    try:
        geom = element.get_Geometry(_make_geometry_options_wall())
        if geom is None:
            return
        if iterate_net_collection is not None:
            geom_items = iterate_net_collection(geom)
        else:
            geom_items = []
            try:
                for go in geom:
                    geom_items.append(go)
            except Exception:
                pass
        for go in geom_items:
            try:
                if isinstance(go, GeometryInstance):
                    solids_inst = []
                    try:
                        inst = go.GetInstanceGeometry()
                        if inst is not None:
                            if iterate_net_collection is not None:
                                inst_items = iterate_net_collection(inst)
                            else:
                                inst_items = list(inst)
                            for x in inst_items:
                                if isinstance(x, Solid) and _vol_ok(x):
                                    solids_inst.append(x)
                    except Exception:
                        solids_inst = []
                    if solids_inst:
                        for g in solids_inst:
                            yield g, None
                    else:
                        geoms = []
                        xf = None
                        try:
                            xf = go.Transform
                        except Exception:
                            xf = None
                        try:
                            sym = go.GetSymbolGeometry()
                            if sym is not None:
                                if iterate_net_collection is not None:
                                    sym_items = iterate_net_collection(sym)
                                else:
                                    sym_items = list(sym)
                                for x in sym_items:
                                    if isinstance(x, Solid) and _vol_ok(x):
                                        geoms.append(x)
                        except Exception:
                            geoms = []
                        for g in geoms:
                            yield g, xf
                elif isinstance(go, Solid):
                    if _vol_ok(go):
                        yield go, None
            except Exception:
                continue
    except Exception:
        return


# ── Intersección plano — geometría ────────────────────────────────────────────
def _signed_plane_dist(plane, p):
    try:
        return float(plane.Normal.DotProduct(p.Subtract(plane.Origin)))
    except Exception:
        return 0.0


def _plane_document_to_local(plane_doc, trf_to_doc):
    if plane_doc is None:
        return None
    if trf_to_doc is None:
        return plane_doc
    try:
        inv = trf_to_doc.Inverse
        o = inv.OfPoint(plane_doc.Origin)
        n = inv.OfVector(plane_doc.Normal)
        if float(n.GetLength()) < 1e-12:
            return None
        n = n.Normalize()
        return Plane.CreateByNormalAndOrigin(n, o)
    except Exception:
        return None


def _xyz_document_from_local(p, trf_to_doc):
    if p is None:
        return None
    if trf_to_doc is None:
        return XYZ(float(p.X), float(p.Y), float(p.Z))
    try:
        q = trf_to_doc.OfPoint(p)
        return XYZ(float(q.X), float(q.Y), float(q.Z))
    except Exception:
        return None


def _manual_intersect_segment_plane_points(p0, p1, plane):
    if plane is None or p0 is None or p1 is None:
        return []
    try:
        n = plane.Normal
        o = plane.Origin
        tol = max(float(_PLANE_LINE_DIST_TOL_FT), 1e-9)
        d0 = float(n.DotProduct(p0.Subtract(o)))
        d1 = float(n.DotProduct(p1.Subtract(o)))
        if abs(d0) <= tol and abs(d1) <= tol:
            return [
                XYZ(float(p0.X), float(p0.Y), float(p0.Z)),
                XYZ(float(p1.X), float(p1.Y), float(p1.Z)),
            ]
        if abs(d0) <= tol:
            return [XYZ(float(p0.X), float(p0.Y), float(p0.Z))]
        if abs(d1) <= tol:
            return [XYZ(float(p1.X), float(p1.Y), float(p1.Z))]
        if d0 * d1 < 0.0:
            denom = d0 - d1
            if abs(denom) < 1e-12:
                return []
            t = float(d0) / float(denom)
            if t < -tol or t > 1.0 + tol:
                return []
            t = max(0.0, min(1.0, t))
            v = p1.Subtract(p0)
            pt = p0.Add(v.Multiply(t))
            return [XYZ(float(pt.X), float(pt.Y), float(pt.Z))]
        return []
    except Exception:
        return []


def _curve_intersect_plane_points(curve, plane):
    if curve is None or not curve.IsBound or plane is None:
        return []
    pts = []
    try:
        arr = IntersectionResultArray()
        r = curve.Intersect(plane, arr)
        if arr is not None and int(arr.Size) > 0:
            for i in range(int(arr.Size)):
                try:
                    it = arr.get_Item(i)
                    if it is not None and it.XYZPoint is not None:
                        pts.append(it.XYZPoint)
                except Exception:
                    continue
        if (
                r in (
                    SetComparisonResult.Subset,
                    SetComparisonResult.Superset,
                    SetComparisonResult.Overlap,
                )
                and not pts):
            try:
                p0 = curve.GetEndPoint(0)
                p1 = curve.GetEndPoint(1)
                d0 = _signed_plane_dist(plane, p0)
                d1 = _signed_plane_dist(plane, p1)
                tol_c = max(_TOL_SEG_MERGE_FT * 0.15, 1e-5)
                if abs(d0) <= tol_c and abs(d1) <= tol_c:
                    return [p0, p1]
            except Exception:
                pass
    except Exception:
        pass
    return pts


def _segment_from_doc_points_longest_chord(pts_local_doc, trf_to_doc):
    if len(pts_local_doc) < 2:
        return None
    try:
        out = []
        for p in pts_local_doc:
            q = _xyz_document_from_local(p, trf_to_doc)
            if q is not None:
                out.append(q)
        if len(out) < 2:
            return None
        if len(out) == 2:
            return Line.CreateBound(out[0], out[1])
        best_i, best_j, best_d = 0, 1, -1.0
        n = len(out)
        for i in range(n):
            for j in range(i + 1, n):
                d = float(out[i].DistanceTo(out[j]))
                if d > best_d:
                    best_d = d
                    best_i, best_j = i, j
        if best_d < 1e-8:
            return None
        return Line.CreateBound(out[best_i], out[best_j])
    except Exception:
        return None


def _arista_segmento_si_coplanar(edge, plane_local, trf_to_doc):
    """Arista lineal contenida en el plano de corte (tramo completo del contorno)."""
    if edge is None or plane_local is None:
        return None
    try:
        crv = edge.AsCurve()
    except Exception:
        crv = None
    if crv is None or not crv.IsBound or not isinstance(crv, Line):
        return None
    try:
        p0 = crv.GetEndPoint(0)
        p1 = crv.GetEndPoint(1)
        tol = max(float(_PLANE_LINE_DIST_TOL_FT), float(_TOL_PT_MATCH_FT) * 0.5)
        if abs(_signed_plane_dist(plane_local, p0)) > tol:
            return None
        if abs(_signed_plane_dist(plane_local, p1)) > tol:
            return None
        q0 = _xyz_document_from_local(p0, trf_to_doc)
        q1 = _xyz_document_from_local(p1, trf_to_doc)
        if q0 is None or q1 is None or float(q0.DistanceTo(q1)) <= 1e-8:
            return None
        return Line.CreateBound(q0, q1)
    except Exception:
        return None


def _arista_plane_segment(edge, plane_local, trf_to_doc):
    if edge is None or plane_local is None:
        return None
    try:
        crv = edge.AsCurve()
    except Exception:
        crv = None
    if crv is None or not crv.IsBound:
        return None
    pts_loc = []
    try:
        pts_loc = _curve_intersect_plane_points(crv, plane_local)
        if len(pts_loc) < 2 and isinstance(crv, Line):
            p0 = crv.GetEndPoint(0)
            p1 = crv.GetEndPoint(1)
            pts_loc = _manual_intersect_segment_plane_points(p0, p1, plane_local)
    except Exception:
        pts_loc = []
    if len(pts_loc) >= 2:
        return _segment_from_doc_points_longest_chord(pts_loc, trf_to_doc)
    return None


def _collect_plane_segments_solid(solid, plane_doc, trf):
    segs = []
    pl = _plane_document_to_local(plane_doc, trf)
    if pl is None or solid is None:
        return segs
    try:
        ea = solid.Edges
        if ea is None:
            return segs
        ne = int(ea.Size)
    except Exception:
        return segs
    for i in range(ne):
        seg = None
        try:
            ed = ea.get_Item(i)
        except Exception:
            ed = None
        try:
            if ed is not None:
                seg = _arista_segmento_si_coplanar(ed, pl, trf)
                if seg is None:
                    seg = _arista_plane_segment(ed, pl, trf)
        except Exception:
            seg = None
        if seg is not None and isinstance(seg, Line) and float(seg.Length) > 1e-8:
            segs.append(seg)
    return segs


def _pts_close(a, b, tol):
    try:
        return float(a.DistanceTo(b)) <= tol
    except Exception:
        return False


def _orient_segment_to_continue(prev_end, cand, tol_match):
    a = cand.GetEndPoint(0)
    b = cand.GetEndPoint(1)
    if _pts_close(prev_end, a, tol_match):
        return Line.CreateBound(a, b)
    if _pts_close(prev_end, b, tol_match):
        return Line.CreateBound(b, a)
    return None


def _dedupe_near_segments(lines, tol_mid):
    if not lines:
        return []
    def mid_len(crv):
        p0 = crv.GetEndPoint(0)
        p1 = crv.GetEndPoint(1)
        m = p0.Add(p1).Multiply(0.5)
        return (
            round(float(m.X) / tol_mid) * tol_mid,
            round(float(m.Y) / tol_mid) * tol_mid,
            round(float(m.Z) / tol_mid) * tol_mid,
            round(float(crv.Length) / tol_mid) * tol_mid,
        )

    seen = set()
    out = []
    for ln in lines:
        key = mid_len(ln)
        if key in seen:
            continue
        seen.add(key)
        out.append(ln)
    return out


def _union_segments_into_ordered_curves(segments, tol_match):
    usable = []
    for s in segments:
        if isinstance(s, Line):
            ln = float(s.Length)
            if ln > 1e-8:
                usable.append(s)
    usable = _dedupe_near_segments(usable, max(tol_match * 5.0, 1e-5))
    if len(usable) < 3:
        return None

    keyed = [(-float(c.Length), i, c) for i, c in enumerate(usable)]
    keyed.sort()
    usable = [c for _neg_len, _i, c in keyed]
    n_seg = len(usable)

    for start_i in range(n_seg):
        for reverse_start in (False, True):
            unused = set(range(n_seg))
            unused.discard(start_i)
            chain = []
            s0 = usable[start_i]
            if not reverse_start:
                chain.append(Line.CreateBound(s0.GetEndPoint(0), s0.GetEndPoint(1)))
            else:
                chain.append(Line.CreateBound(s0.GetEndPoint(1), s0.GetEndPoint(0)))
            target = chain[-1].GetEndPoint(1)
            start_pt = chain[0].GetEndPoint(0)
            stall = 0

            guard = max(80, (n_seg + 5) * (n_seg + 5))

            while guard > 0:
                guard -= 1

                if not unused:
                    if _pts_close(start_pt, target, tol_match):
                        return chain
                    break

                if stall > n_seg + 3:
                    break

                next_k = None
                next_ori = None
                for k in unused:
                    o = _orient_segment_to_continue(target, usable[k], tol_match)
                    if o is not None:
                        next_k = k
                        next_ori = o
                        break

                if next_k is not None:
                    chain.append(next_ori)
                    unused.remove(next_k)
                    target = next_ori.GetEndPoint(1)
                    stall = 0
                    continue

                stall += 1

    return None


def _ejes_uv_seccion_muro(wall, plane_doc):
    """Ejes u (≈ tangente muro) y v (≈ vertical en el plano de corte) en coords. documento."""
    if plane_doc is None:
        return None
    try:
        n = plane_doc.Normal
        ln = float(n.GetLength())
        if ln < 1e-12:
            return None
        n_hat = XYZ(float(n.X) / ln, float(n.Y) / ln, float(n.Z) / ln)
        o_plane = XYZ(
            float(plane_doc.Origin.X),
            float(plane_doc.Origin.Y),
            float(plane_doc.Origin.Z),
        )
    except Exception:
        return None

    Maj = obtener_direccion_principal_muro(wall)
    try:
        t_par = Maj.DotProduct(n_hat)
        Tu_raw = XYZ(
            float(Maj.X) - float(t_par) * float(n_hat.X),
            float(Maj.Y) - float(t_par) * float(n_hat.Y),
            float(Maj.Z) - float(t_par) * float(n_hat.Z),
        )
    except Exception:
        Tu_raw = None
    if Tu_raw is None or float(Tu_raw.GetLength()) < 1e-12:
        try:
            g = XYZ.BasisZ
            gz = float(g.DotProduct(n_hat))
            Tu_raw = XYZ(
                float(g.X) - gz * float(n_hat.X),
                float(g.Y) - gz * float(n_hat.Y),
                float(g.Z) - gz * float(n_hat.Z),
            )
        except Exception:
            return None

    tul = float(Tu_raw.GetLength())
    if tul < 1e-12:
        return None
    u_hat = XYZ(float(Tu_raw.X) / tul, float(Tu_raw.Y) / tul, float(Tu_raw.Z) / tul)
    try:
        Vv_raw = n_hat.CrossProduct(u_hat)
        lv = float(Vv_raw.GetLength())
        if lv < 1e-12:
            return None
        v_hat = XYZ(float(Vv_raw.X) / lv, float(Vv_raw.Y) / lv, float(Vv_raw.Z) / lv)
    except Exception:
        return None
    return u_hat, v_hat, o_plane


def _punto_plano_uv(u_hat, v_hat, o_plane, u_feet, v_feet):
    return XYZ(
        float(o_plane.X) + float(u_hat.X) * u_feet + float(v_hat.X) * v_feet,
        float(o_plane.Y) + float(u_hat.Y) * u_feet + float(v_hat.Y) * v_feet,
        float(o_plane.Z) + float(u_hat.Z) * u_feet + float(v_hat.Z) * v_feet,
    )


def _curvas_rectangulo_uv(u_hat, v_hat, o_plane, umin, umax, vmin, vmax):
    if umax - umin < 1e-6 or vmax - vmin < 1e-6:
        return None
    p_bl = _punto_plano_uv(u_hat, v_hat, o_plane, umin, vmin)
    p_br = _punto_plano_uv(u_hat, v_hat, o_plane, umax, vmin)
    p_tr = _punto_plano_uv(u_hat, v_hat, o_plane, umax, vmax)
    p_tl = _punto_plano_uv(u_hat, v_hat, o_plane, umin, vmax)
    try:
        return [
            Line.CreateBound(p_bl, p_br),
            Line.CreateBound(p_br, p_tr),
            Line.CreateBound(p_tr, p_tl),
            Line.CreateBound(p_tl, p_bl),
        ]
    except Exception:
        return None


def _puntos_interseccion_plano_solido_doc(wall, plane_doc):
    """Puntos 3D (documento) donde aristas del sólido cortan el plano — sin BBAA."""
    out = []
    for sol, trf in _iter_solids_wall(wall):
        pl = _plane_document_to_local(plane_doc, trf)
        if pl is None:
            continue
        try:
            ea = sol.Edges
            ne = int(ea.Size)
        except Exception:
            continue
        for i in range(ne):
            try:
                ed = ea.get_Item(i)
            except Exception:
                ed = None
            if ed is None:
                continue
            try:
                crv = ed.AsCurve()
            except Exception:
                crv = None
            if crv is None or not crv.IsBound:
                continue
            pts_loc = []
            try:
                pts_loc = _curve_intersect_plane_points(crv, pl)
                if len(pts_loc) < 2 and isinstance(crv, Line):
                    p0 = crv.GetEndPoint(0)
                    p1 = crv.GetEndPoint(1)
                    pts_loc = _manual_intersect_segment_plane_points(p0, p1, pl)
            except Exception:
                pts_loc = []
            for p in pts_loc or []:
                q = _xyz_document_from_local(p, trf)
                if q is not None:
                    out.append(q)
    return out


def _curvas_perimetro_respaldo_uv_desde_geometria(wall, plane_doc, segments):
    """
    Respaldo sin BBAA: envolvente u/v en el plano de corte a partir de tramos o puntos
    reales de intersección plano ∩ sólido.
    """
    ejes = _ejes_uv_seccion_muro(wall, plane_doc)
    if ejes is None:
        return None
    u_hat, v_hat, o_plane = ejes

    us = []
    vs = []
    for seg in segments or []:
        try:
            for pt in (seg.GetEndPoint(0), seg.GetEndPoint(1)):
                d = pt.Subtract(o_plane)
                us.append(float(d.DotProduct(u_hat)))
                vs.append(float(d.DotProduct(v_hat)))
        except Exception:
            continue

    if len(us) < 2:
        for pt in _puntos_interseccion_plano_solido_doc(wall, plane_doc):
            try:
                d = pt.Subtract(o_plane)
                us.append(float(d.DotProduct(u_hat)))
                vs.append(float(d.DotProduct(v_hat)))
            except Exception:
                continue

    if len(us) < 2:
        return None

    return _curvas_rectangulo_uv(u_hat, v_hat, o_plane, min(us), max(us), min(vs), max(vs))


def _walls_para_corte_plano(host, muros_vecinos=None):
    """Host + vecinos únicos (sin duplicar el anfitrión)."""
    out = []
    seen = set()
    if host is not None and isinstance(host, Wall):
        try:
            seen.add(_element_id_int(host.Id))
        except Exception:
            pass
        out.append(host)
    for w in muros_vecinos or []:
        if w is None or not isinstance(w, Wall):
            continue
        try:
            wid = _wall_id_int(w)
        except Exception:
            wid = None
        if wid is not None and wid in seen:
            continue
        if wid is not None:
            seen.add(wid)
        out.append(w)
    return out


def _curvas_perimetro_respaldo_uv_desde_geometria_multi(host, walls_cut, plane_doc, segments):
    """
    Respaldo UV: envolvente en el plano de corte usando segmentos y puntos de todos los muros.
    Ejes u/v se derivan del ``host``.
    """
    ejes = _ejes_uv_seccion_muro(host, plane_doc)
    if ejes is None:
        return None
    u_hat, v_hat, o_plane = ejes

    us = []
    vs = []
    for seg in segments or []:
        try:
            for pt in (seg.GetEndPoint(0), seg.GetEndPoint(1)):
                d = pt.Subtract(o_plane)
                us.append(float(d.DotProduct(u_hat)))
                vs.append(float(d.DotProduct(v_hat)))
        except Exception:
            continue

    if len(us) < 2:
        for w in walls_cut or []:
            for pt in _puntos_interseccion_plano_solido_doc(w, plane_doc):
                try:
                    d = pt.Subtract(o_plane)
                    us.append(float(d.DotProduct(u_hat)))
                    vs.append(float(d.DotProduct(v_hat)))
                except Exception:
                    continue

    if len(us) < 2:
        return None

    return _curvas_rectangulo_uv(u_hat, v_hat, o_plane, min(us), max(us), min(vs), max(vs))


def _face_coplanar_con_plano(face, plane_doc, trf_to_doc, tol_ang=0.05, tol_dist_ft=None):
    if face is None or plane_doc is None or not isinstance(face, PlanarFace):
        return False
    if tol_dist_ft is None:
        tol_dist_ft = max(_TOL_PT_MATCH_FT * 2.0, 1e-4)
    try:
        n_loc = face.FaceNormal
        o_loc = face.Origin
        if trf_to_doc is not None:
            n_loc = trf_to_doc.OfVector(n_loc)
            o_loc = trf_to_doc.OfPoint(o_loc)
        ln = float(n_loc.GetLength())
        if ln < 1e-12:
            return False
        n_loc = n_loc.Multiply(1.0 / ln)
        pn = plane_doc.Normal
        pln = float(pn.GetLength())
        if pln < 1e-12:
            return False
        pn = pn.Multiply(1.0 / pln)
        if abs(float(n_loc.DotProduct(pn))) < (1.0 - tol_ang):
            return False
        return abs(_signed_plane_dist(plane_doc, o_loc)) <= tol_dist_ft
    except Exception:
        return False


def _curvas_perimetro_bbox_en_plano_seccion(wall, plane_doc):
    """Rectángulo en el plano de sección proyectando el BoundingBox del muro."""
    ejes = _ejes_uv_seccion_muro(wall, plane_doc)
    if ejes is None:
        return None
    u_hat, v_hat, o_plane = ejes
    try:
        bb = wall.get_BoundingBox(None)
        if bb is None:
            return None
        mn = bb.Min
        mx = bb.Max
        corners = (
            XYZ(mn.X, mn.Y, mn.Z),
            XYZ(mx.X, mn.Y, mn.Z),
            XYZ(mx.X, mx.Y, mn.Z),
            XYZ(mn.X, mx.Y, mn.Z),
            XYZ(mn.X, mn.Y, mx.Z),
            XYZ(mx.X, mn.Y, mx.Z),
            XYZ(mx.X, mx.Y, mx.Z),
            XYZ(mn.X, mx.Y, mx.Z),
        )
        us = []
        vs = []
        for pt in corners:
            d = pt.Subtract(o_plane)
            us.append(float(d.DotProduct(u_hat)))
            vs.append(float(d.DotProduct(v_hat)))
        if not us:
            return None
        return _curvas_rectangulo_uv(
            u_hat, v_hat, o_plane, min(us), max(us), min(vs), max(vs),
        )
    except Exception:
        return None


def _curvas_desde_faces_coplanares(wall, plane_doc):
    """Si alguna cara del sólido es coplanar al corte, usa sus CurveLoops."""
    try:
        from bimtools_clr_collections import (
            curves_from_curve_loop,
            iterate_net_collection,
            net_collection_count,
            safe_solid_volume,
        )
    except Exception:
        curves_from_curve_loop = None
        iterate_net_collection = None
        net_collection_count = None
        safe_solid_volume = None

    def _vol_ok(sol):
        if sol is None:
            return False
        if safe_solid_volume is not None:
            v = safe_solid_volume(sol)
            return v is not None and abs(v) > 1e-12
        return _safe_volume(sol) is not None and abs(_safe_volume(sol)) > 1e-12

    best = None
    best_len = -1.0
    for sol, trf in _iter_solids_wall(wall):
        if not _vol_ok(sol):
            continue
        try:
            faces = sol.Faces
            if faces is None:
                continue
            if net_collection_count is not None:
                nf = net_collection_count(faces)
            else:
                nf = int(faces.Size)
        except Exception:
            continue
        for i in range(nf):
            try:
                face = faces.get_Item(i)
            except Exception:
                try:
                    face = faces[i]
                except Exception:
                    face = None
            if not _face_coplanar_con_plano(face, plane_doc, trf):
                continue
            try:
                loops = face.GetEdgesAsCurveLoops()
            except Exception:
                loops = None
            if loops is None:
                continue
            try:
                if net_collection_count is not None:
                    nloops = net_collection_count(loops)
                elif bool(getattr(loops, u"IsEmpty", False)):
                    nloops = 0
                else:
                    nloops = int(loops.Size)
            except Exception:
                nloops = 0
            if nloops < 1:
                continue
            try:
                cl = loops.get_Item(0)
            except Exception:
                try:
                    cl = loops[0]
                except Exception:
                    cl = None
            if cl is None:
                continue
            if curves_from_curve_loop is not None:
                sub = curves_from_curve_loop(cl)
            else:
                sub = []
                try:
                    for j in range(int(cl.Count)):
                        c = cl.get_Item(j)
                        if c is not None and c.IsBound and float(c.Length) > 1e-8:
                            sub.append(c)
                except Exception:
                    sub = []
            if len(sub) >= 3:
                try:
                    perim = sum(float(c.Length) for c in sub)
                except Exception:
                    perim = 0.0
                if perim > best_len:
                    best_len = perim
                    best = sub
    return best


def _curvas_desde_faces_coplanares_multi(walls, plane_doc):
    """Cara coplanar de mayor perímetro entre varios muros."""
    best = None
    best_len = -1.0
    for wall in walls or []:
        sub = _curvas_desde_faces_coplanares(wall, plane_doc)
        if not sub:
            continue
        try:
            perim = sum(float(c.Length) for c in sub)
        except Exception:
            perim = 0.0
        if perim > best_len:
            best_len = perim
            best = sub
    return best


def _ordenar_curvas_cerradas(segments, tol_match):
    if not segments:
        return None
    ordered = _union_segments_into_ordered_curves(segments, tol_match)
    if ordered and len(ordered) >= 3:
        return ordered
    tol_relax = max(float(tol_match) * 4.0, 0.04)
    ordered = _union_segments_into_ordered_curves(segments, tol_relax)
    if ordered and len(ordered) >= 3:
        return ordered
    return None


def muros_vecinos_en_extremos_host(doc, host):
    """Vecinos en extremos (nodo); lista vacía si el módulo no está disponible."""
    if _vecinos_extremos_mod is None or doc is None or host is None:
        return []
    try:
        return _vecinos_extremos_mod.muros_vecinos_en_extremos(doc, host) or []
    except Exception:
        return []


def curvas_area_reinf_desde_muro_y_plano(_document, wall, plane_doc, muros_vecinos=None):
    """
    Curvas ordenadas cerradas del perímetro: intersección del plano con las aristas del sólido
    del muro host (y muros vecinos en extremos, si se indican). Geometría real, no BoundingBox.
    """
    walls_cut = _walls_para_corte_plano(wall, muros_vecinos)
    if not walls_cut:
        return None

    agg = []
    any_solid = False
    for w in walls_cut:
        pairs = []
        try:
            for s, trf in _iter_solids_wall(w):
                pairs.append((s, trf))
        except Exception:
            pairs = []
        if pairs:
            any_solid = True
        for sol, trf in pairs:
            try:
                agg.extend(_collect_plane_segments_solid(sol, plane_doc, trf))
            except Exception:
                pass

    if agg:
        try:
            agg = _dedupe_near_segments(agg, max(_TOL_PT_MATCH_FT * 3.0, 5e-4))
        except Exception:
            pass
        ordered = _ordenar_curvas_cerradas(agg, _TOL_PT_MATCH_FT)
        if ordered and len(ordered) >= 3:
            return ordered

    face_curves = _curvas_desde_faces_coplanares_multi(walls_cut, plane_doc)
    if face_curves and len(face_curves) >= 3:
        return face_curves

    if agg or any_solid:
        respaldo = _curvas_perimetro_respaldo_uv_desde_geometria_multi(
            wall, walls_cut, plane_doc, agg,
        )
        if respaldo and len(respaldo) >= 3:
            return respaldo

    bbox_curves = _curvas_perimetro_bbox_en_plano_seccion(wall, plane_doc)
    if bbox_curves and len(bbox_curves) >= 3:
        return bbox_curves

    return None


def obtener_direccion_principal_muro(wall):
    t = None
    lc = location_curve_wall(wall)
    if lc is not None:
        t = _tangente_ui_desde_location_curve(lc)
    if t is not None:
        return t
    try:
        loc = wall.Location
        if isinstance(loc, LocationCurve):
            c = getattr(loc, "Curve", None)
            if c is None:
                return XYZ(1, 0, 0)
            p0 = c.GetEndPoint(0)
            p1 = c.GetEndPoint(1)
            dx = p1.X - p0.X
            dy = p1.Y - p0.Y
            dz = p1.Z - p0.Z
            length = (dx * dx + dy * dy + dz * dz) ** 0.5
            if length > 1e-6:
                return XYZ(dx / length, dy / length, dz / length)
    except Exception:
        pass
    return XYZ(1, 0, 0)


def obtener_direccion_major_area_rein(wall, muro_contencion=False):
    """Major direction para ``AreaReinforcement.Create``: horizontal (tangente) o vertical (Z)."""
    if muro_contencion:
        try:
            z = XYZ.BasisZ
            ln = float(z.GetLength())
            if ln > 1e-12:
                return XYZ(float(z.X) / ln, float(z.Y) / ln, float(z.Z) / ln)
        except Exception:
            pass
        return XYZ(0.0, 0.0, 1.0)
    return obtener_direccion_principal_muro(wall)


def orden_vertices_xyz_desde_curvas_cerradas(curves):
    if not curves:
        return []
    verts = []
    for i, c in enumerate(curves):
        a = c.GetEndPoint(0)
        b = c.GetEndPoint(1)
        if i == 0:
            verts.append(XYZ(float(a.X), float(a.Y), float(a.Z)))
        verts.append(XYZ(float(b.X), float(b.Y), float(b.Z)))
    return verts


def _dedupe_vertices_xyz_sequence(vertices, tol_ft):
    out = []
    for p in vertices:
        if out and float(out[-1].DistanceTo(p)) <= tol_ft:
            continue
        out.append(p)
    if len(out) >= 2 and float(out[0].DistanceTo(out[-1])) <= tol_ft:
        out.pop()
    return out


def _primer_bar_desde_params(params_dict):
    for key in (
            "exterior_major", "exterior_minor", "interior_major", "interior_minor",
    ):
        pair = params_dict.get(key)
        if not pair:
            continue
        pid = pair[0]
        if pid and pid != ElementId.InvalidElementId:
            return pid
    return ElementId.InvalidElementId


def compute_section_preview_model(doc, wall):
    """
    Modelo para previsualización 2D: polígono sección del muro en coordenadas (u,v) pies internos.

    Tu = proyección al plano del eje mayor de armado (~tangente Location).
    Vv = N × Tu (en el plano de corte, altura característica).

    Devuelve None si no hay geometría útil.
    """
    res = plano_vertical_contiene_location_y_vertical_global(wall)
    if res is None:
        return None
    plane_doc, lc = res
    vecinos = muros_vecinos_en_extremos_host(doc, wall)
    curvas = curvas_area_reinf_desde_muro_y_plano(doc, wall, plane_doc, vecinos)
    if not curvas or len(curvas) < 3:
        return None

    try:
        n = plane_doc.Normal
        ln = float(n.GetLength())
        if ln < 1e-12:
            return None
        n_hat = XYZ(float(n.X) / ln, float(n.Y) / ln, float(n.Z) / ln)
        o_plane = XYZ(
            float(plane_doc.Origin.X),
            float(plane_doc.Origin.Y),
            float(plane_doc.Origin.Z),
        )
    except Exception:
        return None

    Maj = obtener_direccion_principal_muro(wall)
    try:
        t_par = Maj.DotProduct(n_hat)
        Tu_raw = XYZ(
            float(Maj.X) - float(t_par) * float(n_hat.X),
            float(Maj.Y) - float(t_par) * float(n_hat.Y),
            float(Maj.Z) - float(t_par) * float(n_hat.Z),
        )
    except Exception:
        Tu_raw = None
    if Tu_raw is None or float(Tu_raw.GetLength()) < 1e-12:
        try:
            g = XYZ.BasisZ
            gz = float(g.DotProduct(n_hat))
            Tu_raw = XYZ(
                float(g.X) - gz * float(n_hat.X),
                float(g.Y) - gz * float(n_hat.Y),
                float(g.Z) - gz * float(n_hat.Z),
            )
        except Exception:
            return None

    tul = float(Tu_raw.GetLength())
    if tul < 1e-12:
        return None
    u_hat = XYZ(float(Tu_raw.X) / tul, float(Tu_raw.Y) / tul, float(Tu_raw.Z) / tul)
    try:
        Vv_raw = n_hat.CrossProduct(u_hat)
        lv = float(Vv_raw.GetLength())
        if lv < 1e-12:
            return None
        v_hat = XYZ(float(Vv_raw.X) / lv, float(Vv_raw.Y) / lv, float(Vv_raw.Z) / lv)
    except Exception:
        return None

    verts_xyz = orden_vertices_xyz_desde_curvas_cerradas(curvas)
    verts_xyz = _dedupe_vertices_xyz_sequence(verts_xyz, _TOL_PT_MATCH_FT)
    if len(verts_xyz) < 3:
        return None

    uv_feet = []
    u_coords = []
    v_coords = []
    for p in verts_xyz:
        d = p.Subtract(o_plane)
        uf = float(d.DotProduct(u_hat))
        vf = float(d.DotProduct(v_hat))
        uv_feet.append((uf, vf))
        u_coords.append(uf)
        v_coords.append(vf)

    umin = min(u_coords)
    umax = max(u_coords)
    vmin = min(v_coords)
    vmax = max(v_coords)
    cuc = sum(u_coords) / float(len(u_coords))
    cvc = sum(v_coords) / float(len(v_coords))

    try:
        m_par = Maj.DotProduct(n_hat)
        Mproj = XYZ(
            float(Maj.X) - float(m_par) * float(n_hat.X),
            float(Maj.Y) - float(m_par) * float(n_hat.Y),
            float(Maj.Z) - float(m_par) * float(n_hat.Z),
        )
        mu = float(Mproj.DotProduct(u_hat))
        mv = float(Mproj.DotProduct(v_hat))
        major_angle_rad = math.atan2(mv, mu)
    except Exception:
        major_angle_rad = 0.0

    try:
        o_cen = XYZ(
            float(o_plane.X) + float(u_hat.X) * cuc + float(v_hat.X) * cvc,
            float(o_plane.Y) + float(u_hat.Y) * cuc + float(v_hat.Y) * cvc,
            float(o_plane.Z) + float(u_hat.Z) * cuc + float(v_hat.Z) * cvc,
        )
    except Exception:
        o_cen = o_plane

    try:
        th = float(UnitUtils.ConvertToInternalUnits(obtener_espesor_muro_mm_approx(wall) or 200.0, UnitTypeId.Millimeters))
    except Exception:
        th = UnitUtils.ConvertToInternalUnits(200.0, UnitTypeId.Millimeters)

    return {
        "poly_uv_feet": uv_feet,
        "u_min_feet": umin,
        "u_max_feet": umax,
        "v_min_feet": vmin,
        "v_max_feet": vmax,
        "centroid_u_feet": cuc,
        "centroid_v_feet": cvc,
        "major_angle_rad": major_angle_rad,
        "extent_u_feet": max(umax - umin, 1e-6),
        "extent_v_feet": max(vmax - vmin, 1e-6),
        "thickness_hint_feet": th,
        "plane_origin_xyz": o_cen,
    }


def wall_preview_anchor_xyz(wall, section_model=None):
    """Punto 3D de referencia para layout en canvas: centroide de sección o medio de LocationCurve."""
    if section_model is not None:
        try:
            p = section_model.get("plane_origin_xyz")
            if p is not None:
                return p
        except Exception:
            pass
    lc = location_curve_wall(wall)
    return _punto_centro_location_curve(lc)


def _wall_length_feet_approx(wall):
    try:
        lc = location_curve_wall(wall)
        if lc is not None:
            return max(float(lc.Length), 1e-6)
    except Exception:
        pass
    try:
        bb = wall.get_BoundingBox(None)
        if bb is not None:
            dx = abs(float(bb.Max.X) - float(bb.Min.X))
            dy = abs(float(bb.Max.Y) - float(bb.Min.Y))
            dz = abs(float(bb.Max.Z) - float(bb.Min.Z))
            return max(dx, dy, dz, 1e-6)
    except Exception:
        pass
    return 1.0


def _project_xy_on_axis(x, y, axis_xy):
    ax, ay = axis_xy
    return float(x) * ax + float(y) * ay


def _average_wall_tangent_xy(walls):
    sx, sy, n_axis = 0.0, 0.0, 0
    for w in walls or []:
        lc = location_curve_wall(w)
        t = _tangente_ui_desde_location_curve(lc)
        if t is None:
            continue
        tx, ty = float(t.X), float(t.Y)
        tl = (tx * tx + ty * ty) ** 0.5
        if tl < 1e-9:
            continue
        sx += tx / tl
        sy += ty / tl
        n_axis += 1
    if n_axis <= 0:
        return None
    al = (sx * sx + sy * sy) ** 0.5
    if al < 1e-9:
        return None
    return (sx / al, sy / al)


def _wall_bbox_extent_on_axis(wall, axis_xy):
    """Extensión del bounding box del muro proyectada sobre axis_xy (pies)."""
    try:
        bb = wall.get_BoundingBox(None)
        if bb is None:
            return None
        mn, mx = bb.Min, bb.Max
        corners = (
            (float(mn.X), float(mn.Y)),
            (float(mx.X), float(mn.Y)),
            (float(mx.X), float(mx.Y)),
            (float(mn.X), float(mx.Y)),
        )
        projs = [_project_xy_on_axis(x, y, axis_xy) for x, y in corners]
        umin = min(projs)
        umax = max(projs)
        ext = max(umax - umin, 1e-6)
        return ext, (umin + umax) * 0.5
    except Exception:
        return None


def _pick_plan_layout_axis(anchors, walls):
    """Eje en planta donde los centroides presentan mayor separación."""
    if not anchors:
        return (1.0, 0.0)

    cx = sum(float(p.X) for p in anchors) / float(len(anchors))
    cy = sum(float(p.Y) for p in anchors) / float(len(anchors))

    candidates = [(1.0, 0.0), (0.0, 1.0)]
    tan_xy = _average_wall_tangent_xy(walls)
    if tan_xy is not None:
        tx, ty = tan_xy
        candidates.append((tx, ty))
        candidates.append((-ty, tx))

    best_axis = (1.0, 0.0)
    best_spread = -1.0
    seen = set()
    for raw_ax, raw_ay in candidates:
        al = (raw_ax * raw_ax + raw_ay * raw_ay) ** 0.5
        if al < 1e-9:
            continue
        ax, ay = raw_ax / al, raw_ay / al
        key = (round(ax, 4), round(ay, 4))
        if key in seen:
            continue
        seen.add(key)
        vals = [_project_xy_on_axis(float(p.X) - cx, float(p.Y) - cy, (ax, ay)) for p in anchors]
        spread = max(vals) - min(vals) if vals else 0.0
        if spread > best_spread:
            best_spread = spread
            best_axis = (ax, ay)
    return best_axis


def compute_preview_horizontal_layout(walls, model_cache):
    """
    Posición horizontal relativa de muros para canvas.

    Proyecta el centroide de cada muro sobre el eje en planta con mayor separación
    entre centroides (X, Y, tangente media o su perpendicular). El largo en canvas
    usa la extensión del bounding box sobre ese mismo eje.
    """
    walls = list(walls or [])
    if not walls:
        return None

    anchors = []
    for w in walls:
        wid = _wall_id_int(w)
        md = (model_cache or {}).get(wid)
        p = wall_preview_anchor_xyz(w, md)
        if p is None:
            try:
                bb = w.get_BoundingBox(None)
                if bb is not None:
                    p = XYZ(
                        (float(bb.Min.X) + float(bb.Max.X)) * 0.5,
                        (float(bb.Min.Y) + float(bb.Max.Y)) * 0.5,
                        (float(bb.Min.Z) + float(bb.Max.Z)) * 0.5,
                    )
            except Exception:
                p = None
        if p is None:
            p = XYZ(0.0, 0.0, 0.0)
        anchors.append(p)

    axis_xy = _pick_plan_layout_axis(anchors, walls)
    ax, ay = axis_xy

    items = []
    u_mins = []
    u_maxs = []
    extents = []

    for i, w in enumerate(walls):
        p = anchors[i]
        u_pos = _project_xy_on_axis(float(p.X), float(p.Y), axis_xy)

        bb_ext = _wall_bbox_extent_on_axis(w, axis_xy)
        if bb_ext is not None:
            eu, _u_bb = bb_ext
        else:
            eu = None
            md = (model_cache or {}).get(_wall_id_int(w))
            if md is not None:
                try:
                    eu = float(md.get("extent_u_feet", 0.0))
                except Exception:
                    eu = None
            if eu is None or eu <= 1e-9:
                eu = _wall_length_feet_approx(w)
            eu = max(float(eu), 1e-6)

        extents.append(eu)
        half = eu * 0.5
        u_mins.append(u_pos - half)
        u_maxs.append(u_pos + half)
        items.append({"u_pos": u_pos, "extent_u": eu})

    u_span_min = min(u_mins) if u_mins else -0.5
    u_span_max = max(u_maxs) if u_maxs else 0.5
    if u_span_max - u_span_min < 1e-6:
        pad = max(0.5, max(extents) * 0.5 if extents else 0.5)
        u_span_min -= pad
        u_span_max += pad

    return {
        "axis_xy": axis_xy,
        "items": items,
        "u_span_min": u_span_min,
        "u_span_max": u_span_max,
        "max_extent_u": max(extents) if extents else 1.0,
    }


def _extract_wall_endpoints_xy(walls):
    """Extrae P0 y P1 (XY) de la LocationCurve de cada muro."""
    result = []
    for w in walls:
        crv = location_curve_wall(w)
        if crv is None:
            result.append(None)
            continue
        try:
            p0 = crv.GetEndPoint(0)
            p1 = crv.GetEndPoint(1)
            result.append((
                (float(p0.X), float(p0.Y)),
                (float(p1.X), float(p1.Y)),
            ))
        except Exception:
            result.append(None)
    return result


def compute_stacked_wall_layout(walls, view_right_xy=None):
    """Layout horizontal para muros apilados basado en LocationCurve.

    Proyecta P0 y P1 de cada muro sobre un eje compartido para obtener
    largo y posicion relativa reales.

    Si se proporciona *view_right_xy* (tupla (rx, ry) del RightDirection de la
    vista activa), se usa directamente como eje de proyeccion, garantizando que
    la orientacion izquierda/derecha del canvas coincida con la vista de Revit.
    """
    walls = list(walls or [])
    if not walls:
        return None

    endpoints = _extract_wall_endpoints_xy(walls)

    forced_axis = None
    if view_right_xy is not None:
        rx, ry = float(view_right_xy[0]), float(view_right_xy[1])
        rl = (rx * rx + ry * ry) ** 0.5
        if rl > 1e-9:
            forced_axis = (rx / rl, ry / rl)

    if forced_axis is not None:
        best_axis = forced_axis
    else:
        tan_xy = _average_wall_tangent_xy(walls)
        candidates = [(1.0, 0.0), (0.0, 1.0)]
        if tan_xy is not None:
            tx, ty = tan_xy
            candidates.insert(0, (tx, ty))

        best_axis = candidates[0]
        best_span = -1.0

        for ax, ay in candidates:
            al = (ax * ax + ay * ay) ** 0.5
            if al < 1e-9:
                continue
            ax, ay = ax / al, ay / al
            g_min, g_max = 1e18, -1e18
            for ep in endpoints:
                if ep is None:
                    continue
                (x0, y0), (x1, y1) = ep
                u0 = x0 * ax + y0 * ay
                u1 = x1 * ax + y1 * ay
                g_min = min(g_min, u0, u1)
                g_max = max(g_max, u0, u1)
            s = g_max - g_min if g_max > g_min else 0.0
            if s > best_span:
                best_span = s
                best_axis = (ax, ay)

    ax, ay = best_axis

    items = []
    all_u_min = []
    all_u_max = []

    for ep in endpoints:
        if ep is None:
            items.append({
                u"u_start": 0.0, u"u_end": 0.0, u"length_u": 1e-6,
                u"u0": 0.0, u"u1": 0.0,
            })
            all_u_min.append(0.0)
            all_u_max.append(0.0)
            continue

        (x0, y0), (x1, y1) = ep
        u0 = x0 * ax + y0 * ay
        u1 = x1 * ax + y1 * ay
        u_start = min(u0, u1)
        u_end = max(u0, u1)
        length_u = max(u_end - u_start, 1e-6)

        items.append({
            u"u_start": u_start,
            u"u_end": u_end,
            u"length_u": length_u,
            u"u0": u0,
            u"u1": u1,
        })
        all_u_min.append(u_start)
        all_u_max.append(u_end)

    if not all_u_min:
        return None

    global_min = min(all_u_min)
    global_max = max(all_u_max)
    global_span = max(global_max - global_min, 1e-6)

    return {
        u"axis_xy": best_axis,
        u"items": items,
        u"global_min": global_min,
        u"global_max": global_max,
        u"global_span": global_span,
    }


def cabezal_extremos_en_lados_stacked(wall, row_index, stacked_layout):
    """Asocia extremo de LocationCurve con izq/der usando stacked layout.

    Usa los valores u0/u1 pre-calculados del layout si estan disponibles,
    evitando re-proyectar los endpoints.
    P0 (inicio) con menor u → inicio a la izquierda.
    P0 (inicio) con mayor u → inicio a la derecha.
    """
    try:
        import armado_muros_cabezal as _cab
        ex_ini = _cab.CABEZAL_EXTREMO_INICIO
        ex_fin = _cab.CABEZAL_EXTREMO_FIN
    except Exception:
        ex_ini, ex_fin = u"inicio", u"fin"

    if stacked_layout is None or wall is None:
        return ex_ini, ex_fin

    st_items = stacked_layout.get(u"items") or []
    if not (0 <= int(row_index) < len(st_items)):
        return ex_ini, ex_fin

    item = st_items[int(row_index)]
    u_left = float(item.get(u"u_start", 0.0))

    if u"u0" in item and u"u1" in item:
        u0 = float(item[u"u0"])
        u1 = float(item[u"u1"])
    else:
        axis_xy = stacked_layout.get(u"axis_xy")
        if not axis_xy:
            return ex_ini, ex_fin
        lc = location_curve_wall(wall)
        if lc is None:
            return ex_ini, ex_fin
        try:
            p0 = lc.GetEndPoint(0)
            p1 = lc.GetEndPoint(1)
            u0 = _project_xy_on_axis(float(p0.X), float(p0.Y), axis_xy)
            u1 = _project_xy_on_axis(float(p1.X), float(p1.Y), axis_xy)
        except Exception:
            return ex_ini, ex_fin

    if abs(u0 - u_left) <= abs(u1 - u_left):
        return ex_ini, ex_fin
    return ex_fin, ex_ini


def cabezal_extremos_en_lados_preview(wall, row_index, layout):
    """
    Asocia extremo de LocationCurve con el lado izquierdo/derecho del tramo en elevación.

    Retorna (extremo_izquierda, extremo_derecha) como ``inicio`` / ``fin``.
    """
    try:
        import armado_muros_cabezal as _cab

        ex_ini = _cab.CABEZAL_EXTREMO_INICIO
        ex_fin = _cab.CABEZAL_EXTREMO_FIN
    except Exception:
        ex_ini, ex_fin = u"inicio", u"fin"

    if layout is None or wall is None:
        return ex_ini, ex_fin

    items = layout.get("items") or []
    if not (0 <= int(row_index) < len(items)):
        return ex_ini, ex_fin

    item = items[int(row_index)]
    try:
        u_left = float(item.get("u_pos", 0.0)) - float(item.get("extent_u", 1.0)) * 0.5
    except Exception:
        return ex_ini, ex_fin

    axis_xy = layout.get("axis_xy")
    if not axis_xy:
        return ex_ini, ex_fin

    lc = location_curve_wall(wall)
    if lc is None:
        return ex_ini, ex_fin

    try:
        p0 = lc.GetEndPoint(0)
        p1 = lc.GetEndPoint(1)
        u0 = _project_xy_on_axis(float(p0.X), float(p0.Y), axis_xy)
        u1 = _project_xy_on_axis(float(p1.X), float(p1.Y), axis_xy)
    except Exception:
        return ex_ini, ex_fin

    if abs(u0 - u_left) <= abs(u1 - u_left):
        return ex_ini, ex_fin
    return ex_fin, ex_ini


def obtener_espesor_muro_mm_approx(wall):
    """Espesor aproximado del muro (mm) sólo como referencia gráfica (no geometría de corte)."""
    try:

        param = wall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)
        if param and param.HasValue:
            return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters)
        param = wall.LookupParameter("Default Thickness")
        if param and param.HasValue:
            return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters)
        type_id = wall.GetTypeId()
        if type_id and type_id != ElementId.InvalidElementId:
            wt = wall.Document.GetElement(type_id)
            if wt:
                for pname in ("Default Thickness", "Thickness", "Espesor", "Width"):
                    p2 = wt.LookupParameter(pname)
                    if p2 and p2.HasValue:
                        return UnitUtils.ConvertFromInternalUnits(p2.AsDouble(), UnitTypeId.Millimeters)
    except Exception:
        pass
    return None


def malla_sic_default_por_espesor_mm(e_mm):
    """
    Doble malla M.H.A. (S.I.C.) — Ø y @ por espesor nominal del muro (mm).

    Valores por defecto para la UI; el usuario puede modificarlos antes de crear.
    Misma regla en exterior/interior y major/minor (doble malla).
    """
    try:
        e = int(round(float(e_mm)))
    except Exception:
        e = 200
    if e < 1:
        e = 200
    if e <= 200:
        return 8, 200
    if e <= 300:
        return 10, 200
    if e <= 399:
        return 12, 200
    if e <= 599:
        return 12, 150
    if e <= 799:
        return 16, 200
    return 16, 200


def malla_sic_defaults_para_muro(wall):
    """``(diam_mm, spacing_mm)`` según espesor del ``Wall`` (S.I.C.)."""
    e_mm = obtener_espesor_muro_mm_approx(wall)
    if e_mm is None:
        e_mm = 200.0
    return malla_sic_default_por_espesor_mm(e_mm)


def _wall_bottom_z_sort_key_ft(wall):
    try:
        bb = wall.get_BoundingBox(None)
        if bb is not None:
            return float(bb.Min.Z)
    except Exception:
        pass
    try:
        loc = wall.Location
        crv = getattr(loc, "Curve", None)
        if crv is not None:
            z0 = float(crv.GetEndPoint(0).Z)
            z1 = float(crv.GetEndPoint(1).Z)
            return min(z0, z1)
    except Exception:
        pass
    return 0.0


def _doc_from_element(elem):
    try:
        return elem.Document
    except Exception:
        return None


# Referencia de cota para etiquetas de elevación (UI). Geometría no cambia.
COTA_REF_SURVEY = u"survey"
COTA_REF_PROJECT_BASE = u"project_base"
COTA_REF_INTERNAL = u"internal"
COTA_REF_MODES = (
    COTA_REF_SURVEY,
    COTA_REF_PROJECT_BASE,
    COTA_REF_INTERNAL,
)


def normalize_cota_ref_mode(mode):
    """Normaliza mode → survey | project_base | internal (default survey)."""
    try:
        m = unicode(mode or u"").strip().lower()
    except Exception:
        m = u""
    if m in (u"pbp", u"base", u"project", u"projectbase"):
        return COTA_REF_PROJECT_BASE
    if m in (u"origen", u"origin", u"internal_origin", u"interno"):
        return COTA_REF_INTERNAL
    if m in COTA_REF_MODES:
        return m
    return COTA_REF_SURVEY


def _survey_point(doc):
    """``BasePoint`` del Survey Point (``IsShared``), o None."""
    if doc is None:
        return None
    try:
        sp = BasePoint.GetSurveyPoint(doc)
        if sp is not None:
            return sp
    except Exception:
        pass
    try:
        for bp in FilteredElementCollector(doc).OfClass(BasePoint):
            try:
                if bool(getattr(bp, u"IsShared", False)):
                    return bp
            except Exception:
                continue
    except Exception:
        pass
    return None


def _project_base_point(doc):
    """``BasePoint`` del Project Base Point (``IsShared == False``), o None."""
    if doc is None:
        return None
    try:
        pbp = BasePoint.GetProjectBasePoint(doc)
        if pbp is not None:
            return pbp
    except Exception:
        pass
    try:
        for bp in FilteredElementCollector(doc).OfClass(BasePoint):
            try:
                if not bool(getattr(bp, u"IsShared", True)):
                    return bp
            except Exception:
                continue
    except Exception:
        pass
    return None


def internal_z_ft_to_survey_meters(doc, z_ft):
    """Z interna (pies) → elevación Survey Point / shared (metros).

    Fórmula canónica (Revit API)::

        shared = ActiveProjectLocation.GetTotalTransform().Inverse.OfPoint(
            XYZ(0, 0, z_ft)
        )
        metros = shared.Z * 0.3048

    Respaldo con Survey Point::

        (z_ft - Position.Z) + SharedPosition.Z

    Si no hay transform ni Survey Point, usa Z interna × 0.3048.
    """
    return internal_z_ft_to_meters(doc, z_ft, COTA_REF_SURVEY)


def internal_z_ft_to_meters(doc, z_ft, mode=None):
    """Z interna (pies) → metros según referencia de cota.

    Modes:
    - ``survey``: Survey Point / shared via ``GetTotalTransform().Inverse``
    - ``project_base``: relativo a Project Base Point (``Position.Z``)
    - ``internal``: origen interno crudo ``z_ft * 0.3048``
    """
    z = float(z_ft or 0.0)
    ref = normalize_cota_ref_mode(mode)
    if ref == COTA_REF_INTERNAL:
        return z * 0.3048
    if ref == COTA_REF_PROJECT_BASE:
        if doc is not None:
            try:
                pbp = _project_base_point(doc)
                if pbp is not None:
                    return (z - float(pbp.Position.Z)) * 0.3048
            except Exception:
                pass
        return z * 0.3048
    # survey (default)
    if doc is not None:
        try:
            tr = doc.ActiveProjectLocation.GetTotalTransform()
            shared = tr.Inverse.OfPoint(XYZ(0.0, 0.0, z))
            return float(shared.Z) * 0.3048
        except Exception:
            pass
        try:
            sp = _survey_point(doc)
            if sp is not None:
                pos_z = float(sp.Position.Z)
                try:
                    shared_z = float(sp.SharedPosition.Z)
                except Exception:
                    shared_z = 0.0
                return (z - pos_z + shared_z) * 0.3048
        except Exception:
            pass
    return z * 0.3048


def cota_inferior_muro_metros_aprox(wall, mode=None):
    """Para UI: cota inferior del muro en metros (ref. survey / PBP / interno)."""
    z_ft = _wall_bottom_z_sort_key_ft(wall)
    return internal_z_ft_to_meters(
        _doc_from_element(wall), z_ft, mode=mode,
    )


def cota_superior_muro_metros_aprox(wall, mode=None):
    """Para UI: cota superior del muro en metros (ref. survey / PBP / interno)."""
    try:
        bb = wall.get_BoundingBox(None)
        if bb is not None:
            return internal_z_ft_to_meters(
                _doc_from_element(wall), float(bb.Max.Z), mode=mode,
            )
    except Exception:
        pass
    return cota_inferior_muro_metros_aprox(wall, mode=mode)


def ordenar_muros_por_base_asc(walls):
    """Del más bajo al más alto (menor Z de envolvente o curva). Lista nueva."""
    keyed = []
    for w in (walls or []):
        if w is None:
            continue
        keyed.append((_wall_bottom_z_sort_key_ft(w), w))
    keyed.sort()
    return [w for _z, w in keyed]


# Eje ⟂ normal de vista ⟺ eje paralelo al plano de la vista (≈ 2°).
_TOL_DOT_EJE_PARALELO_PLANO_VISTA = math.sin(math.radians(2.0))


def _normal_plano_vista_activa(view):
    """Normal unitaria del plano de la vista activa (``ViewDirection``)."""
    if view is None:
        return None
    try:
        vd = view.ViewDirection
        if vd is not None and float(vd.GetLength()) > 1e-12:
            return vd.Normalize()
    except Exception:
        pass
    return None


def _eje_muro_paralelo_plano_vista(wall, view):
    """
    True si el eje del muro (``LocationCurve``) es paralelo al plano de la vista.

    Criterio: ``|T · N_view| <= sin(2°)`` con ``T`` tangente unitaria y ``N_view`` normal de vista.
    """
    if wall is None or not isinstance(wall, Wall):
        return False
    n_view = _normal_plano_vista_activa(view)
    if n_view is None:
        return False
    lc = location_curve_wall(wall)
    if lc is None:
        return False
    t = _tangente_ui_desde_location_curve(lc)
    if t is None:
        return False
    try:
        return abs(float(t.DotProduct(n_view))) <= _TOL_DOT_EJE_PARALELO_PLANO_VISTA
    except Exception:
        return False


class _FiltroSoloMuros(ISelectionFilter):
    def __init__(self, view):
        self._view = view

    def AllowElement(self, elem):
        try:
            if not isinstance(elem, Wall):
                return False
            return _eje_muro_paralelo_plano_vista(elem, self._view)
        except Exception:
            return False

    def AllowReference(self, reference, position):
        return False


def _pick_muros(uidoc):
    doc = uidoc.Document
    view = uidoc.ActiveView
    _pick_muros._rechazo_paralelo = False
    try:
        refs = list(uidoc.Selection.PickObjects(
            ObjectType.Element,
            _FiltroSoloMuros(view),
            u"Seleccione muros cuyo eje sea visible en la vista activa "
            u"(paralelos al plano de corte) · Finalice en la cinta (Finalizar) o Esc para cancelar",
        ))
    except OperationCanceledException:
        return None
    walls = []
    for ref in refs:
        el = doc.GetElement(ref.ElementId)
        if el is not None and isinstance(el, Wall):
            if _eje_muro_paralelo_plano_vista(el, view):
                walls.append(el)
    if refs and not walls:
        _pick_muros._rechazo_paralelo = True
        TaskDialog.Show(
            u"Arainco: Armado Muros",
            u"Ningún muro seleccionado es paralelo al plano de la vista activa.\n\n"
            u"Solo se pueden armar muros cuyo eje (LocationCurve) se vea en la vista actual "
            u"(sección, alzado o planta). Los muros en punta o no alineados con el corte "
            u"quedan excluidos.",
        )
    return walls


# ── Rebar Cover ───────────────────────────────────────────────────────────────
def obtener_rebar_cover_type_por_mm(doc, mm_value):
    """Devuelve un ``RebarCoverType`` existente con ``CoverDistance`` ≈ ``mm_value``."""
    try:
        target_ft = float(UnitUtils.ConvertToInternalUnits(float(mm_value), UnitTypeId.Millimeters))
    except Exception:
        return None
    try:
        tol_ft = float(UnitUtils.ConvertToInternalUnits(0.01, UnitTypeId.Millimeters))
    except Exception:
        tol_ft = 1e-6

    try:
        collector = FilteredElementCollector(doc).OfClass(RebarCoverType)
    except Exception:
        return None

    for elem in collector:
        try:
            dist = float(elem.CoverDistance)
            if abs(dist - target_ft) <= tol_ft:
                return elem
        except Exception:
            continue
    return None


def _set_wall_cover_bip(wall, bip, cover_type):
    if wall is None or cover_type is None:
        return False
    try:
        p = wall.get_Parameter(bip)
        if p is None or p.IsReadOnly:
            return False
        p.Set(cover_type.Id)
        return True
    except Exception:
        return False


def asignar_rebar_cover_a_muro(wall, cover_ext_int, cover_other):
    """
    Asigna Rebar Cover por cara al muro:
    - Exterior / Interior → ``cover_ext_int`` (25 mm)
    - Other Faces → ``cover_other`` (0 mm)
    """
    if wall is None or cover_ext_int is None or cover_other is None:
        return False, u"tipos de cover inválidos"

    ok_ext = _set_wall_cover_bip(
        wall, BuiltInParameter.CLEAR_COVER_EXTERIOR, cover_ext_int,
    )
    ok_int = _set_wall_cover_bip(
        wall, BuiltInParameter.CLEAR_COVER_INTERIOR, cover_ext_int,
    )
    ok_other = _set_wall_cover_bip(
        wall, BuiltInParameter.CLEAR_COVER_OTHER, cover_other,
    )

    if ok_ext and ok_int and ok_other:
        return True, None

    faltan = []
    if not ok_ext:
        faltan.append(u"exterior")
    if not ok_int:
        faltan.append(u"interior")
    if not ok_other:
        faltan.append(u"otras caras")
    return False, u"no se pudo asignar: " + u", ".join(faltan)


def _resolver_tipos_rebar_cover(doc, errores=None):
    """Resuelve tipos de cover ext/int (25 mm) y otras caras (0 mm)."""
    cover_ext_int = obtener_rebar_cover_type_por_mm(doc, REBAR_COVER_MM_CARAS_EXT_INT)
    cover_other = obtener_rebar_cover_type_por_mm(doc, REBAR_COVER_MM_OTRAS_CARAS)
    if cover_ext_int is None:
        msg = u"No se encontró Rebar Cover {:.0f} mm en el proyecto.".format(
            REBAR_COVER_MM_CARAS_EXT_INT,
        )
        if errores is not None:
            errores.append(msg)
        return None, None, msg
    if cover_other is None:
        msg = u"No se encontró Rebar Cover {:.0f} mm en el proyecto.".format(
            REBAR_COVER_MM_OTRAS_CARAS,
        )
        if errores is not None:
            errores.append(msg)
        return None, None, msg
    return cover_ext_int, cover_other, None


def asignar_rebar_cover_a_muros(doc, walls, errores=None):
    """
    Asigna Rebar Cover a todos los muros antes de crear armadura.

    Evita que un cambio de cover al ejecutar la malla recalcule/estire rebars
    ya creados (p. ej. estribos de cabezal ligados al cover inicial del muro).
    """
    cover_ext_int, cover_other, err = _resolver_tipos_rebar_cover(doc, errores)
    if err:
        return 0
    n_ok = 0
    trans = Transaction(doc, u"Arainco: Armado muros — Rebar Cover")
    try:
        from armado_muros_txn import attach_rebar_outside_host_swallower
        attach_rebar_outside_host_swallower(trans)
    except Exception:
        pass
    try:
        trans.Start()
        for wall in walls or []:
            if wall is None:
                continue
            ok, err_wall = asignar_rebar_cover_a_muro(
                wall, cover_ext_int, cover_other,
            )
            if ok:
                n_ok += 1
            elif errores is not None and err_wall:
                try:
                    wid = _wall_id_int(wall) or 0
                except Exception:
                    wid = 0
                errores.append(
                    u"Muro {0}: Rebar Cover — {1}.".format(wid, err_wall),
                )
        try:
            doc.Regenerate()
        except Exception:
            pass
        trans.Commit()
    except Exception as ex:
        if trans.HasStarted():
            try:
                trans.RollBack()
            except Exception:
                pass
        if errores is not None:
            errores.append(u"Rebar Cover (lote): {0}".format(ex))
        return n_ok
    return n_ok


_MURO_VERT_MIN_ABS_TZ = 0.45


def _rebar_tangent_z_abs(rebar, pos_idx=0):
    if rebar is None:
        return None
    try:
        crvs = rebar.GetCenterlineCurves(
            False, False, False,
            MultiplanarOption.IncludeAllMultiplanarCurves,
            int(pos_idx),
        )
        if crvs is None or crvs.Count < 1:
            return None
        c0 = crvs[0]
        p0 = c0.GetEndPoint(0)
        p1 = c0.GetEndPoint(1)
        v = p1 - p0
        if v.GetLength() < 1e-12:
            return None
        t = v.Normalize()
        return abs(float(t.Z))
    except Exception:
        return None


def _rebar_max_tangent_z_abs(rebar, max_pos=None):
    """Máximo |t.Z| en todos los tramos del eje (patías L tras empotramiento)."""
    if rebar is None:
        return None
    max_tz = None
    try:
        from armado_muros_nodo_shared import _rebar_cantidad_posiciones
        n_pos = min(4, max(1, int(_rebar_cantidad_posiciones(rebar))))
    except Exception:
        n_pos = 1
    if max_pos is not None:
        try:
            n_pos = min(n_pos, max(1, int(max_pos)))
        except Exception:
            pass
    for pi in range(n_pos):
        try:
            crvs = rebar.GetCenterlineCurves(
                False, False, False,
                MultiplanarOption.IncludeAllMultiplanarCurves,
                int(pi),
            )
        except Exception:
            continue
        if crvs is None:
            continue
        try:
            n_crv = int(crvs.Count)
        except Exception:
            n_crv = 0
        for ci in range(n_crv):
            try:
                c = crvs[ci]
                p0 = c.GetEndPoint(0)
                p1 = c.GetEndPoint(1)
                v = p1 - p0
                if v.GetLength() < 1e-12:
                    continue
                tz = abs(float(v.Normalize().Z))
                if max_tz is None or tz > max_tz:
                    max_tz = tz
            except Exception:
                pass
    return max_tz


def _capas_verticales_muro_keys(muro_contencion=False):
    u"""Capas de malla con barras verticales (minor tradicional, major contención)."""
    if muro_contencion:
        return (u"exterior_major", u"interior_major")
    return (u"exterior_minor", u"interior_minor")


def _rebar_es_vertical_en_muro(rebar, host):
    if rebar is None or host is None:
        return False
    if _rebar_es_vertical_por_criterio is not None:
        try:
            from armado_muros_nodo_shared import _rebar_cantidad_posiciones
            n_try = min(4, max(1, int(_rebar_cantidad_posiciones(rebar))))
        except Exception:
            n_try = 1
        for pi in range(n_try):
            try:
                if bool(_rebar_es_vertical_por_criterio(rebar, host, pi)):
                    return True
            except Exception:
                pass
    if not isinstance(host, Wall):
        return False
    tz = _rebar_max_tangent_z_abs(rebar)
    if tz is None:
        tz = _rebar_tangent_z_abs(rebar, 0)
    if tz is None:
        return False
    return tz >= float(_MURO_VERT_MIN_ABS_TZ)


def _rebar_es_horizontal_dominante_en_muro(rebar, host):
    """True si el tramo dominante del eje es horizontal en muro (|t.Z| bajo)."""
    if rebar is None:
        return False
    try:
        from armado_muros_horizontales_retraida import (
            HORIZONTAL_MAX_ABS_TZ,
            PLAN_PREFER_X_FOR_HORIZONTAL,
            _rebar_horizontal_en_plano,
        )
        return bool(
            _rebar_horizontal_en_plano(
                rebar,
                PLAN_PREFER_X_FOR_HORIZONTAL,
                HORIZONTAL_MAX_ABS_TZ,
                host,
            ),
        )
    except Exception:
        pass
    try:
        from arearein_exterior_h_l135_rps import (
            HORIZONTAL_MAX_ABS_TZ,
            PLAN_PREFER_X_FOR_HORIZONTAL,
            _rebar_horizontal_en_plano,
        )
        return bool(
            _rebar_horizontal_en_plano(
                rebar,
                PLAN_PREFER_X_FOR_HORIZONTAL,
                HORIZONTAL_MAX_ABS_TZ,
                host,
            ),
        )
    except Exception:
        return False


def _rebar_es_malla_horizontal_para_stamp(rebar, host, params_dict=None, muro_contencion=False):
    """Clasificación horizontal para ``Armadura_Malla_Orientacion`` = H."""
    if rebar is None:
        return False
    try:
        import armado_muros_cabezal as _cab_malla

        if params_dict:
            es_h = bool(
                _cab_malla.rebar_coincide_tipo_capa_malla_horizontal(
                    rebar, params_dict, muro_contencion,
                ),
            )
            es_v = bool(
                _cab_malla.rebar_coincide_tipo_capa_malla_vertical(
                    rebar, params_dict, muro_contencion,
                ),
            )
            if es_h and not es_v:
                return True
            if _cab_malla.rebar_es_malla_vertical_por_tipo(
                rebar, params_dict, muro_contencion,
            ):
                return False
            if es_v and not es_h:
                return False
            if es_h and es_v:
                return _rebar_es_horizontal_dominante_en_muro(rebar, host)
    except Exception:
        pass
    return _rebar_es_horizontal_dominante_en_muro(rebar, host)


def _rebar_es_malla_vertical_para_stamp(rebar, host, params_dict=None, muro_contencion=False):
    """
    Clasificación vertical/horizontal para ``Armadura_Malla_*`` (interior y exterior).

    Menos conservadora que ``_rebar_es_malla_vertical_para_exclusion``: si el tipo
    coincide solo con capas verticales, o la geometría tiene tramo vertical (patas L),
    se considera vertical.
    """
    if rebar is None:
        return False
    if _rebar_es_horizontal_dominante_en_muro(rebar, host):
        return False
    try:
        import armado_muros_cabezal as _cab_malla

        if params_dict and _cab_malla.rebar_es_malla_vertical_por_tipo(
            rebar, params_dict, muro_contencion,
        ):
            return True
        es_v = bool(
            params_dict
            and _cab_malla.rebar_coincide_tipo_capa_malla_vertical(
                rebar, params_dict, muro_contencion,
            )
        )
        es_h = bool(
            params_dict
            and _cab_malla.rebar_coincide_tipo_capa_malla_horizontal(
                rebar, params_dict, muro_contencion,
            )
        )
        if es_h and not es_v:
            return False
        if es_h and es_v:
            if _rebar_es_horizontal_dominante_en_muro(rebar, host):
                return False
    except Exception:
        pass
    return _rebar_es_vertical_en_muro(rebar, host)


def _rebar_es_malla_vertical_para_exclusion(rebar, host, params_dict=None, muro_contencion=False):
    """
    Sets de malla vertical para correlación cabezal (``n_capas``).

    Tipo vertical del panel solo si no coincide también con capa horizontal
    (mismo ``RebarBarType`` en doble malla). Si hay ambigüedad, usa geometría
    (incluye patas L tras empotramiento).
    """
    if rebar is None:
        return False
    try:
        import armado_muros_cabezal as _cab_malla

        if params_dict and _cab_malla.rebar_es_malla_vertical_por_tipo(
            rebar, params_dict, muro_contencion,
        ):
            return True
    except Exception:
        pass
    return _rebar_es_vertical_en_muro(rebar, host)


def _ids_rebars_malla_horizontal_desde_param(doc, rebar_ids):
    """IDs con ``Armadura_Malla_Orientacion`` = H. (lectura tras stamp de creación)."""
    if doc is None or not rebar_ids:
        return []
    try:
        from armado_muros_rebar_params import get_armadura_malla_orientacion
    except Exception:
        return []
    out = []
    for eid in rebar_ids or []:
        try:
            rb = doc.GetElement(eid)
        except Exception:
            continue
        if rb is None:
            continue
        try:
            if get_armadura_malla_orientacion(rb) != u"horizontal":
                continue
        except Exception:
            continue
        rid = _element_id_int(getattr(rb, u"Id", None))
        if rid is not None:
            out.append(int(rid))
    return out


def _spacing_mm_desde_params_txt(esp_txt):
    try:
        return float(str(esp_txt).strip().replace(u",", u"."))
    except Exception:
        return 150.0


def _spacing_internal_desde_mm(spacing_mm):
    try:
        return UnitUtils.ConvertToInternalUnits(float(spacing_mm), UnitTypeId.Millimeters)
    except Exception:
        return UnitUtils.ConvertToInternalUnits(150.0, UnitTypeId.Millimeters)


def _layer_key_malla_orient_cara(orient, cara, muro_contencion=False):
    if orient == u"horizontal":
        if muro_contencion:
            return u"exterior_minor" if cara == u"exterior" else u"interior_minor"
        return u"exterior_major" if cara == u"exterior" else u"interior_major"
    if muro_contencion:
        return u"exterior_major" if cara == u"exterior" else u"interior_major"
    return u"exterior_minor" if cara == u"exterior" else u"interior_minor"


def _layer_key_malla_por_tipo_rebar(rebar, params_dict):
    if rebar is None or not params_dict:
        return None
    try:
        import armado_muros_cabezal as _cab

        for lk in (
            u"exterior_major", u"exterior_minor",
            u"interior_major", u"interior_minor",
        ):
            if _cab._rebar_coincide_tipo_capas_malla(rebar, params_dict, (lk,)):
                return lk
    except Exception:
        pass
    return None


def _orient_rebar_malla_para_spacing(
    rebar, host, params_dict, muro_contencion, horiz_ids=None,
):
    rid = _element_id_int(getattr(rebar, u"Id", None))
    if horiz_ids and rid is not None and rid in horiz_ids:
        return u"horizontal"
    try:
        from armado_muros_rebar_params import get_armadura_malla_orientacion

        o = get_armadura_malla_orientacion(rebar)
        if o == u"horizontal":
            return u"horizontal"
        if o == u"vertical":
            return u"vertical"
    except Exception:
        pass
    if _rebar_es_malla_horizontal_para_stamp(
        rebar, host, params_dict, muro_contencion,
    ):
        return u"horizontal"
    return u"vertical"


def _cara_rebar_malla_para_spacing(rebar, host):
    if _malla_rebar_tags_mod is not None:
        fn = getattr(_malla_rebar_tags_mod, u"_cara_rebar_en_muro", None)
        if fn is not None:
            try:
                return fn(rebar, host)
            except Exception:
                pass
    return None


def _rebar_layout_rule_nombre(rebar, acc):
    try:
        r = rebar.LayoutRule
        if r is not None:
            s = r.ToString() or u""
            if s:
                return s
    except Exception:
        pass
    if acc is not None:
        try:
            r = acc.GetLayoutRule()
            if r is not None:
                s = r.ToString() or u""
                if s:
                    return s
        except Exception:
            pass
    return u""


def _aplicar_spacing_ui_rebar(rebar, spacing_mm, doc=None, regenerate=True):
    """Fija el paso del set con el espaciamiento exacto del panel (mm)."""
    if rebar is None or spacing_mm is None:
        return False
    try:
        acc = rebar.GetShapeDrivenAccessor()
    except Exception:
        acc = None
    if acc is None:
        return False
    rule = _rebar_layout_rule_nombre(rebar, acc)
    if rule == u"Single" or u"Single" in rule:
        return False
    try:
        alen = float(acc.ArrayLength)
    except Exception:
        try:
            alen = float(acc.GetArrayLength())
        except Exception:
            alen = 0.0
    if alen < 1e-12:
        return False
    sp_ft = _spacing_internal_desde_mm(spacing_mm)
    if sp_ft >= alen - 1e-9:
        return False
    try:
        b_side = bool(acc.BarsOnNormalSide)
    except Exception:
        b_side = True
    try:
        inc0 = bool(rebar.IncludeFirstBar)
        inc1 = bool(rebar.IncludeLastBar)
    except Exception:
        inc0, inc1 = True, True

    def _try_set(b_side_):
        if rule in (u"MaximumSpacing", u""):
            acc.SetLayoutAsMaximumSpacing(sp_ft, alen, b_side_, inc0, inc1)
        elif rule == u"NumberWithSpacing":
            nbars = int(rebar.Quantity)
            acc.SetLayoutAsNumberWithSpacing(nbars, sp_ft, alen, b_side_, inc0, inc1)
        elif rule == u"MinimumClearSpacing":
            acc.SetLayoutAsMinimumClearSpacing(sp_ft, alen, b_side_, inc0, inc1)
        else:
            acc.SetLayoutAsMaximumSpacing(sp_ft, alen, b_side_, inc0, inc1)

    for b_try in (b_side, not b_side):
        try:
            _try_set(b_try)
            if doc is not None and regenerate:
                try:
                    doc.Regenerate()
                except Exception:
                    pass
            return True
        except Exception:
            continue
    return False


def _aplicar_spacing_malla_rebars_por_muro(
    doc,
    rebars_por_muro_id,
    params_por_muro_id=None,
    muro_contencion=False,
    rebars_horizontal_por_muro_id=None,
    regenerate_after=True,
):
    """
    Reaplica espaciamiento exacto del panel en sets de malla tras post-proceso.

    El copy-layout y la exclusión de extremos recalculan ``MaxSpacing``; las
    etiquetas rebar leen ese valor (p. ej. @149 en lugar de @150 del UI).

    ``regenerate_after=False``: el caller hace un Regenerate consolidado.
    """
    if doc is None or not rebars_por_muro_id:
        return 0
    horiz_map = rebars_horizontal_por_muro_id or {}
    n = 0
    for wid, eid_list in rebars_por_muro_id.items():
        host = None
        try:
            host = doc.GetElement(ElementId(int(wid)))
        except Exception:
            pass
        params_dict = _params_dict_for_wall_id(params_por_muro_id, wid)
        if not params_dict:
            continue
        horiz_set = set()
        for map_key in (wid,):
            for hid in horiz_map.get(map_key) or []:
                try:
                    horiz_set.add(int(hid))
                except Exception:
                    pass
        if not horiz_set:
            try:
                wint = int(wid)
                for hid in horiz_map.get(wint) or []:
                    try:
                        horiz_set.add(int(hid))
                    except Exception:
                        pass
            except Exception:
                pass
        for eid in eid_list or []:
            try:
                rb = doc.GetElement(eid)
            except Exception:
                continue
            if rb is None or not isinstance(rb, Rebar):
                continue
            if host is None:
                try:
                    host = doc.GetElement(rb.GetHostId())
                except Exception:
                    host = None
            orient = _orient_rebar_malla_para_spacing(
                rb, host, params_dict, muro_contencion, horiz_set,
            )
            cara = _cara_rebar_malla_para_spacing(rb, host)
            lk = None
            if cara in (u"exterior", u"interior"):
                lk = _layer_key_malla_orient_cara(orient, cara, muro_contencion)
            if lk is None:
                lk = _layer_key_malla_por_tipo_rebar(rb, params_dict)
            if lk is None:
                lk = _layer_key_malla_orient_cara(
                    orient, u"exterior", muro_contencion,
                )
            try:
                _, esp_txt = params_dict.get(lk, (None, u"150"))
            except Exception:
                esp_txt = u"150"
            spacing_mm = _spacing_mm_desde_params_txt(esp_txt)
            if _aplicar_spacing_ui_rebar(rb, spacing_mm, doc=doc, regenerate=False):
                n += 1
    if regenerate_after and n and doc is not None:
        try:
            doc.Regenerate()
        except Exception:
            pass
    return n


def _stamp_malla_params_rebars_por_muro(
    doc,
    rebars_por_muro_id,
    params_por_muro_id=None,
    muro_contencion=False,
    rebars_horizontal_por_muro_id=None,
):
    """
    Rellena ``Armadura_Malla`` = Yes, ``Armadura_Malla_Tipo`` / ``Armadura_Malla_Orientacion``.

    Verticales: Tipo = D.M., Orientación = V.
    Horizontales: Orientación = H.

    Fuente de verdad para etiquetas: el parámetro ``Armadura_Malla_Orientacion``.
    Si hay registro de post-proceso horizontal, ese set → H. y el resto → V.
    Si no, se respeta el parámetro ya estampado; si falta, tipo de capa / geometría.
    """
    if doc is None or not rebars_por_muro_id:
        return 0
    try:
        from armado_muros_rebar_params import (
            get_armadura_malla_orientacion,
            stamp_malla_horizontal_rebar,
            stamp_malla_vertical_rebar,
        )
    except Exception:
        return 0
    horiz_map = rebars_horizontal_por_muro_id or {}
    n = 0
    for wid, eid_list in rebars_por_muro_id.items():
        host = None
        try:
            host = doc.GetElement(ElementId(int(wid)))
        except Exception:
            pass
        params_dict = _params_dict_for_wall_id(params_por_muro_id, wid)
        horiz_set = set()
        for map_key in (wid,):
            for hid in horiz_map.get(map_key) or []:
                try:
                    horiz_set.add(int(hid))
                except Exception:
                    pass
        if not horiz_set:
            try:
                wint = int(wid)
                for hid in horiz_map.get(wint) or []:
                    try:
                        horiz_set.add(int(hid))
                    except Exception:
                        pass
            except Exception:
                pass
        for eid in eid_list or []:
            try:
                rb = doc.GetElement(eid)
            except Exception:
                continue
            if rb is None or not isinstance(rb, Rebar):
                continue
            if host is None:
                try:
                    host = doc.GetElement(rb.GetHostId())
                except Exception:
                    host = None
            rid = _element_id_int(getattr(rb, u"Id", None))
            if rid is not None and rid in horiz_set:
                stamp_malla_horizontal_rebar(rb)
                n += 1
                continue
            if horiz_set:
                stamp_malla_vertical_rebar(rb)
                n += 1
                continue
            try:
                orient_actual = get_armadura_malla_orientacion(rb)
            except Exception:
                orient_actual = None
            if orient_actual == u"horizontal":
                stamp_malla_horizontal_rebar(rb)
                n += 1
                continue
            if orient_actual == u"vertical":
                stamp_malla_vertical_rebar(rb)
                n += 1
                continue
            if _rebar_es_malla_horizontal_para_stamp(
                rb, host, params_dict, muro_contencion,
            ):
                stamp_malla_horizontal_rebar(rb)
            elif _rebar_es_malla_vertical_para_stamp(
                rb, host, params_dict, muro_contencion,
            ):
                stamp_malla_vertical_rebar(rb)
            elif _rebar_es_horizontal_dominante_en_muro(rb, host):
                stamp_malla_horizontal_rebar(rb)
            else:
                stamp_malla_vertical_rebar(rb)
            n += 1
    return n


def _excluir_extremos_rebar_por_orientacion(
    doc,
    rebar,
    host=None,
    ex_cfg_inicio=None,
    ex_cfg_fin=None,
    params_dict=None,
    muro_contencion=False,
    regenerate=True,
):
    """
    Malla vertical (minor): ``SetBarIncluded`` según ``n_capas`` cabezal por extremo.
    Horizontales (major ext/int): remove last bar (una barra en el extremo final del set).
    """
    if rebar is None:
        return False
    if host is None and doc is not None:
        try:
            host = doc.GetElement(rebar.GetHostId())
        except Exception:
            host = None
    if _rebar_es_malla_vertical_para_exclusion(
        rebar, host, params_dict, muro_contencion,
    ):
        try:
            import armado_muros_cabezal as _cab_malla

            return _cab_malla.aplicar_exclusion_verticales_malla_rebar(
                rebar, ex_cfg_inicio, ex_cfg_fin, doc=doc, host=host,
                regenerate=regenerate,
            )
        except Exception:
            if ajustar_inclusion_extremos_rebar_set_con_fallback is None:
                return False
            return ajustar_inclusion_extremos_rebar_set_con_fallback(
                rebar, doc, False, False, regenerate=regenerate,
            )
    try:
        import armado_muros_cabezal as _cab_malla

        return _cab_malla.aplicar_exclusion_horizontal_malla_ultima_barra(
            rebar, doc, regenerate=regenerate,
        )
    except Exception:
        if ajustar_inclusion_extremos_rebar_set_con_fallback is None:
            return False
        return ajustar_inclusion_extremos_rebar_set_con_fallback(
            rebar, doc, True, False, regenerate=regenerate,
        )


def _desactivar_extremos_rebars_creados(
    doc,
    rebar_ids,
    cabezal_por_muro_id=None,
    params_por_muro_id=None,
    muro_contencion=False,
    regenerate_each=True,
    regenerate_after=True,
):
    """
    Excluye barras de extremo según orientación (horizontal vs vertical).

    ``regenerate_after=False``: el caller hace un Regenerate consolidado
    (p. ej. post-lote unificado).
    """
    n_ok = 0
    n_skip = 0
    for eid in rebar_ids or []:
        try:
            rb = doc.GetElement(eid)
        except Exception:
            n_skip += 1
            continue
        if rb is None or not isinstance(rb, Rebar):
            n_skip += 1
            continue
        host = None
        ex_cfg_inicio = None
        ex_cfg_fin = None
        try:
            host = doc.GetElement(rb.GetHostId())
        except Exception:
            host = None
        params_dict = None
        if host is not None and params_por_muro_id:
            try:
                wid = _element_id_int(host.Id)
                tup = params_por_muro_id.get(wid)
                if tup is not None:
                    params_dict = tup[0]
            except Exception:
                params_dict = None
        if host is not None and cabezal_por_muro_id:
            try:
                import armado_muros_cabezal as _cab_malla

                wid = _element_id_int(host.Id)
                ex_cfg_inicio, ex_cfg_fin = _cab_malla.cabezal_extremos_config_for_muro(
                    cabezal_por_muro_id, wid,
                )
            except Exception:
                pass
        try:
            if _excluir_extremos_rebar_por_orientacion(
                doc,
                rb,
                host,
                ex_cfg_inicio,
                ex_cfg_fin,
                params_dict,
                muro_contencion,
                regenerate=regenerate_each,
            ):
                n_ok += 1
            else:
                n_skip += 1
        except Exception:
            n_skip += 1
    if (
        regenerate_after
        and (not regenerate_each)
        and n_ok
        and doc is not None
    ):
        try:
            doc.Regenerate()
        except Exception:
            pass
    return n_ok, n_skip


def _iter_element_ids_from_ilist(ilist):
    """Convierte ``IList<ElementId>`` de la API Revit a lista Python (IronPython-safe)."""
    if ilist is None:
        return []
    out = []
    try:
        n = int(ilist.Count)
    except Exception:
        n = 0
    if n > 0:
        for i in range(n):
            try:
                eid = ilist[i]
                if eid is not None and eid != ElementId.InvalidElementId:
                    out.append(eid)
            except Exception:
                pass
        if out:
            return out
    try:
        for eid in ilist:
            if eid is not None and eid != ElementId.InvalidElementId:
                out.append(eid)
    except Exception:
        pass
    return out


def remove_area_reinforcement_system(doc, area_rein):
    """
    Convierte un Area Reinforcement en rebars individuales.
    Equivalente a Remove Area Reinforcement System en la UI de Revit.
    """
    if doc is None or area_rein is None:
        raise Exception(u"Documento o AreaReinforcement inválido.")
    new_ids = AreaReinforcement.RemoveAreaReinforcementSystem(doc, area_rein)
    return _iter_element_ids_from_ilist(new_ids)


def _refrescar_vista_fin_flujo(doc, uidoc):
    """Un solo regenerar/refresco al terminar un flujo (modo rápido o cierre de animación)."""
    _refrescar_vista_tras_lote(doc, uidoc, forzar=True)


def _refrescar_vista_tras_lote(doc, uidoc, forzar=False):
    """Regenera y refresca la vista activa tras commit de un lote (efecto de aparición)."""
    if not forzar and MODO_EJECUCION_RAPIDA:
        return
    if doc is not None:
        try:
            doc.Regenerate()
        except Exception:
            pass
    if uidoc is None:
        return
    try:
        uidoc.RefreshActiveView()
    except Exception:
        pass
    try:
        uidoc.UpdateAllOpenViews()
    except Exception:
        pass


def _crear_armadura_un_muro_lineal(
    doc,
    wall,
    params_dict,
    layer_active_dict,
    area_type_id,
    aplicar_malla_cb,
    muro_contencion,
    cover_ext_int,
    cover_other,
    hook_invalido,
    errores,
    resumen_ok,
    rebars_por_muro_id,
    cabezal_por_muro_id=None,
    skip_cover_assign=False,
    skip_exclusion=False,
    defer_remove=False,
    pending_area_reins=None,
):
    """
    Crea AR (+ opcional Remove System / exclusión) para un muro (Transaction abierta).

    ``defer_remove=True``: solo Create AR + malla; agrega a ``pending_area_reins``
    ``(wid, area_rein_id, params_dict, layer_active_dict)`` para Remove System
    tras un Regenerate compartido del lote.

    ``skip_exclusion=True``: no excluye extremos aquí (lo hace el post-lote).
    Retorna 1 si el muro se procesó correctamente, 0 si no.
    """
    wid = _wall_id_int(wall)
    if not skip_cover_assign:
        ok_cover, err_cover = asignar_rebar_cover_a_muro(
            wall, cover_ext_int, cover_other,
        )
        if not ok_cover:
            errores.append(
                u"Muro {}: Rebar Cover — {}.".format(wid, err_cover or u"error"),
            )
            return 0

    bar_type_seed = _primer_bar_desde_params(params_dict)
    if not bar_type_seed or bar_type_seed == ElementId.InvalidElementId:
        errores.append(
            u"Muro {}: selecciona RebarBarType en capas activas.".format(wid),
        )
        return 0

    res = plano_vertical_contiene_location_y_vertical_global(wall)
    if res is None:
        errores.append(
            u"Muro {}: LocationCurve/Tangente o plano no resuelto.".format(wid)
        )
        return 0
    plane_doc, _lc_curve = res
    try:
        vecinos = muros_vecinos_en_extremos_host(doc, wall)
    except Exception as ex_v:
        errores.append(u"Muro {}: vecinos — {}.".format(wid, str(ex_v)))
        vecinos = []
    try:
        curvas = curvas_area_reinf_desde_muro_y_plano(doc, wall, plane_doc, vecinos)
    except Exception as ex_c:
        errores.append(u"Muro {}: curvas perímetro — {}.".format(wid, str(ex_c)))
        return 0
    if not curvas or len(curvas) < 3:
        errores.append(
            u"Muro {}: intersección plano-geometría sin perímetro suficiente.".format(wid)
        )
        return 0

    try:
        curve_list = List[Curve]()
        for c in curvas:
            curve_list.Add(c)
    except Exception as ex_list:
        errores.append(u"Muro {}: IList<Curve>: {}".format(wid, str(ex_list)))
        return 0

    major = obtener_direccion_major_area_rein(wall, muro_contencion)
    try:
        area_rein = AreaReinforcement.Create(
            doc,
            wall,
            curve_list,
            major,
            area_type_id,
            bar_type_seed,
            hook_invalido,
        )
    except Exception as ex_create:
        errores.append(u"Muro {}: Create AR — {}.".format(wid, str(ex_create)))
        return 0
    if area_rein is None:
        errores.append(u"Muro {}: Create no devolvió elemento.".format(wid))
        return 0
    area_rein_id = area_rein.Id
    ex_cfg_inicio = None
    ex_cfg_fin = None
    try:
        import armado_muros_cabezal as _cab_malla

        ex_cfg_inicio, ex_cfg_fin = _cab_malla.cabezal_extremos_config_for_muro(
            cabezal_por_muro_id, wid,
        )
    except Exception:
        pass
    if aplicar_malla_cb:
        try:
            aplicar_malla_cb(
                area_rein,
                params_dict,
                layer_active_dict,
                muro_contencion,
                ex_cfg_inicio,
                ex_cfg_fin,
                doc,
            )
        except Exception as ex_malla:
            errores.append(u"Muro {}: parámetros malla — {}.".format(wid, str(ex_malla)))
            return 0

    if defer_remove:
        if pending_area_reins is not None:
            pending_area_reins.append(
                (wid, area_rein_id, params_dict, layer_active_dict),
            )
        return 1

    return _finalizar_remove_system_muro_lineal(
        doc,
        wid,
        area_rein_id,
        params_dict,
        layer_active_dict,
        muro_contencion,
        errores,
        resumen_ok,
        rebars_por_muro_id,
        cabezal_por_muro_id=cabezal_por_muro_id,
        skip_exclusion=skip_exclusion,
        regenerate_before_remove=True,
    )


def _finalizar_remove_system_muro_lineal(
    doc,
    wid,
    area_rein_id,
    params_dict,
    layer_active_dict,
    muro_contencion,
    errores,
    resumen_ok,
    rebars_por_muro_id,
    cabezal_por_muro_id=None,
    skip_exclusion=False,
    regenerate_before_remove=True,
):
    """Remove Area System (+ exclusión opcional) tras Create AR."""
    if regenerate_before_remove:
        try:
            doc.Regenerate()
        except Exception:
            pass
    area_for_remove = doc.GetElement(area_rein_id)
    if area_for_remove is None:
        errores.append(
            u"Muro {}: Area Reinforcement no disponible tras regenerar.".format(wid),
        )
        return 0
    try:
        new_rebar_ids = remove_area_reinforcement_system(doc, area_for_remove)
    except Exception as ex_remove:
        errores.append(
            u"Muro {}: Remove Area System — {}.".format(wid, str(ex_remove)),
        )
        return 0
    if not new_rebar_ids:
        errores.append(
            u"Muro {}: Remove Area System no generó rebars.".format(wid),
        )
        return 0
    if not skip_exclusion:
        try:
            _desactivar_extremos_rebars_creados(
                doc,
                new_rebar_ids,
                cabezal_por_muro_id=cabezal_por_muro_id,
                params_por_muro_id={wid: (params_dict, layer_active_dict)},
                muro_contencion=muro_contencion,
                regenerate_each=False,
            )
        except Exception as ex_ext:
            errores.append(u"Muro {}: exclusión extremos — {}.".format(wid, str(ex_ext)))
            return 0
    rebars_por_muro_id[wid] = list(new_rebar_ids)
    for eid in new_rebar_ids:
        try:
            resumen_ok.append(_element_id_int(eid))
        except Exception:
            try:
                resumen_ok.append(int(eid))
            except Exception:
                pass
    return 1


# ── Ejecución principal ───────────────────────────────────────────────────────


def crear_areas_malla_parametrizada(doc, walls, params_por_muro_id,
                                    area_type_id, aplicar_malla_cb,
                                    muro_contencion=False, uidoc=None,
                                    cabezal_por_muro_id=None,
                                    defer_etiquetas_malla=False,
                                    malla_activo_por_muro_id=None,
                                    skip_coronamiento=False,
                                    skip_cover_assign=False,
                                    within_parent_transaction_group=False):
    """
    Crea Area Reinforcement por muro usando curvas desde intersección plano/sólido
    y aplica parámetros de malla (callback ``aplicar_malla_cb`` sobre cada nuevo elemento).

    ``params_por_muro_id``: ``{ wall_id_int: (params_dict, layer_active_dict), ... }``
    por cada ``wall.Id.IntegerValue``.
    ``muro_contencion``: si True, major direction vertical (``BasisZ``).

    Retorna ``(ids_rebar_creados, errores, muros_con_cover_asignado, embed_resumen)``.

    ``cabezal_por_muro_id``: solo correlación malla (exclusión verticales por ``n_capas``).
    La creación de cabezal longitudinales/confinamiento es responsabilidad del llamador
    (p. ej. ``crear_armado_muros_unificado``).

    Los muros se procesan de **abajo hacia arriba** (``ordenar_muros_por_base_asc``).
    Por lote: creación (AR + Remove System) → post-proceso (estiramientos, patas L, extremos)
    → refresco de vista. La animación muestra barras ya post-procesadas.
    Barra de progreso pyRevit (un paso por lote), igual que Armado columnas.
    """
    resumen_ok = []
    errores = []
    cover_asignados = 0
    rebars_por_muro_id = {}
    embed_resumen = None
    tags_rebar_malla = [0, 0, 0]
    _reload_malla_rebar_tags_mod()
    view_etiqueta = None
    vista_rebar_tag_ok = False
    if not defer_etiquetas_malla and uidoc is not None:
        try:
            view_etiqueta = uidoc.ActiveView
            if _malla_rebar_tags_mod is not None:
                vista_rebar_tag_ok = _malla_rebar_tags_mod._vista_permite_tags_malla(
                    view_etiqueta,
                )
        except Exception:
            view_etiqueta = None

    hook_invalido = ElementId.InvalidElementId

    cover_ext_int, cover_other, err_cover = _resolver_tipos_rebar_cover(doc)
    if err_cover:
        return [], [err_cover], 0, None

    walls_ord = ordenar_muros_por_base_asc(walls)
    fund_solids_cache = {}
    if not skip_coronamiento and walls_ord:
        try:
            cor_embed, cor_msgs = _aplicar_coronamiento_en_creacion(
                doc, walls_ord, params_por_muro_id,
            )
            embed_resumen = _merge_embed_resumen(embed_resumen, cor_embed)
            for m in cor_msgs:
                if m and int((cor_embed or {}).get(u"n_coronamiento_fail", 0)):
                    errores.append(m)
        except Exception as ex_cor:
            errores.append(u"Coronamiento: {0}".format(ex_cor))

    n_muros = len(walls_ord)
    batch = _tamano_lote_ejecucion(n_muros, MUROS_POR_LOTE_ANIMACION)
    n_lotes = int(math.ceil(float(n_muros) / float(batch))) if n_muros else 0

    _pb_crea = None
    _pbar_crea_open = False
    pb_lote_idx = 0

    use_own_tg = use_transaction_group_armado_muros(
        doc, within_parent_transaction_group=within_parent_transaction_group,
    )
    tg = None
    tg_started = False
    creacion_ok = False

    try:
        _pb_crea = _ml_pbar_start(
            _ml_pbar_phase_title(
                u"{} (armado + post)".format(_ML_PBAR_BASE),
                n_lotes,
            ),
            n_lotes,
            doc=doc,
        )
        _pbar_crea_open = _ml_pbar_enter(_pb_crea)

        if use_own_tg:
            tg = TransactionGroup(doc, u"Arainco: Armado muros lineales")
            tg.Start()
            tg_started = True

        for i0 in range(0, n_muros, batch):
            _ml_pbar_step(
                _pb_crea,
                pb_lote_idx,
                n_lotes,
                u"{} (armado + post)".format(_ML_PBAR_BASE),
            )
            pb_lote_idx += 1

            lote = walls_ord[i0:i0 + batch]
            i1 = min(i0 + batch, n_muros)
            if len(lote) == 1:
                w0 = lote[0]
                z_m = cota_inferior_muro_metros_aprox(w0)
                txn_name = u"Arainco: Armado muros lineales — muro {} (Z≈{:.2f} m)".format(
                    _wall_id_int(w0), z_m,
                )
            else:
                txn_name = u"Arainco: Armado muros lineales — lote {}–{} de {}".format(
                    i0 + 1, i1, n_muros,
                )

            t = Transaction(doc, txn_name)
            try:
                from armado_muros_txn import attach_rebar_outside_host_swallower
                attach_rebar_outside_host_swallower(t)
            except Exception:
                pass
            t.Start()
            lote_ok = False
            lote_wids_con_rebars = []
            pending_area_reins = []
            try:
                for wall in lote:
                    wid = _wall_id_int(wall)
                    if malla_activo_por_muro_id is not None:
                        if not malla_activo_por_muro_id.get(wid, True):
                            continue
                    try:
                        tup = params_por_muro_id.get(wid)
                        if tup is None:
                            errores.append(
                                u"Muro {}: falta configuración del panel.".format(wid)
                            )
                            continue
                        params_dict, layer_active_dict = tup
                    except Exception:
                        errores.append(u"Muro {}: parámetros inválidos.".format(wid))
                        continue

                    try:
                        n_creado = _crear_armadura_un_muro_lineal(
                            doc,
                            wall,
                            params_dict,
                            layer_active_dict,
                            area_type_id,
                            aplicar_malla_cb,
                            muro_contencion,
                            cover_ext_int,
                            cover_other,
                            hook_invalido,
                            errores,
                            resumen_ok,
                            rebars_por_muro_id,
                            cabezal_por_muro_id=cabezal_por_muro_id,
                            skip_cover_assign=skip_cover_assign,
                            skip_exclusion=True,
                            defer_remove=True,
                            pending_area_reins=pending_area_reins,
                        )
                    except Exception as ex_muro:
                        errores.append(
                            u"Muro {}: {}.".format(wid, str(ex_muro)),
                        )
                        continue

                if pending_area_reins:
                    try:
                        doc.Regenerate()
                    except Exception:
                        pass
                    for (
                        wid_ar,
                        area_rein_id,
                        params_dict_ar,
                        layer_active_ar,
                    ) in pending_area_reins:
                        n_fin = _finalizar_remove_system_muro_lineal(
                            doc,
                            wid_ar,
                            area_rein_id,
                            params_dict_ar,
                            layer_active_ar,
                            muro_contencion,
                            errores,
                            resumen_ok,
                            rebars_por_muro_id,
                            cabezal_por_muro_id=cabezal_por_muro_id,
                            skip_exclusion=True,
                            regenerate_before_remove=False,
                        )
                        if n_fin and not skip_cover_assign:
                            cover_asignados += n_fin
                        if rebars_por_muro_id.get(wid_ar):
                            lote_wids_con_rebars.append(wid_ar)
                            hids_crea = _ids_rebars_malla_horizontal_desde_param(
                                doc, rebars_por_muro_id.get(wid_ar),
                            )
                            if hids_crea:
                                embed_resumen = _merge_embed_resumen(
                                    embed_resumen,
                                    {
                                        u"rebars_malla_horizontal_por_muro_id": {
                                            int(wid_ar): hids_crea,
                                        },
                                    },
                                )

                t.Commit()
                lote_ok = True
            except Exception as ex_lote:
                try:
                    if t.HasStarted():
                        t.RollBack()
                except Exception:
                    pass
                errores.append(u"Lote {}–{}: {}".format(i0 + 1, i1, str(ex_lote)))

            if lote_ok and lote_wids_con_rebars:
                rebars_lote = {
                    wid: rebars_por_muro_id[wid]
                    for wid in lote_wids_con_rebars
                    if rebars_por_muro_id.get(wid)
                }
                if rebars_lote:
                    embed_resumen = _merge_embed_resumen(
                        embed_resumen,
                        _post_procesar_rebars_lote(
                            doc,
                            walls_ord,
                            rebars_lote,
                            params_por_muro_id,
                            muro_contencion,
                            cabezal_por_muro_id=cabezal_por_muro_id,
                            fund_solids_cache=fund_solids_cache,
                        ),
                    )

            if (
                lote_ok
                and not defer_etiquetas_malla
                and view_etiqueta is not None
                and vista_rebar_tag_ok
                and lote_wids_con_rebars
            ):
                horiz_lote = {}
                if embed_resumen:
                    horiz_lote = (
                        embed_resumen.get(u"rebars_malla_horizontal_por_muro_id") or {}
                    )
                if horiz_lote:
                    _stamp_malla_pre_etiquetar_lote(
                        doc,
                        lote,
                        rebars_por_muro_id,
                        params_por_muro_id,
                        muro_contencion,
                        horiz_lote,
                    )
                _crear_tags_rebar_malla_tras_commit_lote(
                    doc,
                    view_etiqueta,
                    lote,
                    params_por_muro_id,
                    rebars_por_muro_id,
                    tags_rebar_malla,
                    errores,
                    muro_contencion=muro_contencion,
                    rebars_horizontal_por_muro_id=horiz_lote,
                    stamp_pre_etiqueta=not bool(horiz_lote),
                )

            if lote_ok:
                _refrescar_vista_tras_lote(doc, uidoc)

        creacion_ok = True

    except Exception as ex:
        errores.append(str(ex))
    finally:
        _ml_pbar_exit(_pb_crea, _pbar_crea_open)
        if use_own_tg and tg_started and tg is not None:
            try:
                if creacion_ok:
                    tg.Assimilate()
                else:
                    tg.RollBack()
            except Exception:
                try:
                    if tg.HasStarted():
                        tg.RollBack()
                except Exception:
                    pass

    if creacion_ok and uidoc is not None and not defer_etiquetas_malla:
        _refrescar_vista_fin_flujo(doc, uidoc)

    if rebars_por_muro_id:
        embed_resumen = _merge_embed_resumen(
            embed_resumen,
            {u"rebars_por_muro_id": dict(rebars_por_muro_id)},
        )

    n_tag_rb = int(tags_rebar_malla[0])
    n_tag_fail = int(tags_rebar_malla[1])
    n_tag_skip = int(tags_rebar_malla[2])
    if (
        not defer_etiquetas_malla
        and rebars_por_muro_id
        and vista_rebar_tag_ok
    ):
        embed_resumen = _merge_embed_resumen(
            embed_resumen,
            _resumen_etiquetas_malla_desde_contadores(tags_rebar_malla),
        )
    elif (
        cover_asignados > 0
        and not defer_etiquetas_malla
        and rebars_por_muro_id
        and not vista_rebar_tag_ok
    ):
        vt_msg = _view_type_suffix_safe(view_etiqueta) if view_etiqueta else u"?"
        msg_v = (
            u"Etiquetas rebar malla: ninguna (vista activa: {}; use planta, alzado o sección).".format(
                vt_msg,
            )
        )
        errores.append(msg_v)
        embed_resumen = _merge_embed_resumen(
            embed_resumen,
            {u"messages": [msg_v]},
        )
    if _malla_rebar_tags_mod is None and _malla_rebar_tags_import_error:
        errores.append(
            u"Etiquetas rebar malla: módulo no cargado — {}.".format(
                _malla_rebar_tags_import_error,
            ),
        )

    embed_resumen = _aplicar_etiquetado_coronamiento_todos(
        doc, uidoc, embed_resumen, errores,
    )

    return resumen_ok, errores, cover_asignados, embed_resumen


def _fallback_bar_type_from_params(params_por_muro_id):
    try:
        for _wid, tup in (params_por_muro_id or {}).items():
            pd = tup[0] if tup else {}
            bid = pd.get(u"exterior_major", (None, u""))[0]
            if bid and bid != ElementId.InvalidElementId:
                return bid
    except Exception:
        pass
    return None


def _aplicar_coronamiento_en_creacion(
    doc, walls, params_por_muro_id=None,
    coronamiento_cfg=None, coronamiento_por_muro_id=None,
):
    """Coronamiento en tope/inf/voladizo; mapa por muro opcional (UI V3)."""
    embed = None
    msgs = []
    try:
        import armado_muros_coronamiento as _cor_mod

        fb_el = _fallback_bar_type_from_params(params_por_muro_id)
        fb = None
        if fb_el is not None and doc is not None:
            try:
                fb = doc.GetElement(fb_el)
            except Exception:
                fb = None
        cor_map = _normalize_muro_id_dict(coronamiento_por_muro_id)
        cor_res = _cor_mod.aplicar_coronamiento_muros(
            doc,
            walls,
            bar_type_fallback=fb,
            config=coronamiento_cfg,
            coronamiento_por_muro_id=cor_map if cor_map else None,
        )
        embed = _merge_embed_resumen(embed, {
            u"n_coronamiento": int(cor_res.get(u"n_created", 0)),
            u"n_coronamiento_fail": int(cor_res.get(u"n_fail", 0)),
            u"n_coronamiento_bars": int(cor_res.get(u"n_bars", 0)),
            u"n_coronamiento_inferior": int(cor_res.get(u"n_inferior_created", 0)),
            u"n_coronamiento_inferior_fail": int(cor_res.get(u"n_inferior_fail", 0)),
            u"n_coronamiento_inferior_bars": int(cor_res.get(u"n_inferior_bars", 0)),
            u"n_coronamiento_inferior_pie": int(cor_res.get(u"n_inferior_pie_created", 0)),
            u"n_coronamiento_inferior_pie_fail": int(cor_res.get(u"n_inferior_pie_fail", 0)),
            u"n_coronamiento_inferior_pie_bars": int(cor_res.get(u"n_inferior_pie_bars", 0)),
            u"n_coronamiento_voladizo": int(cor_res.get(u"n_voladizo_created", 0)),
            u"n_coronamiento_voladizo_fail": int(cor_res.get(u"n_voladizo_fail", 0)),
            u"n_coronamiento_voladizo_bars": int(cor_res.get(u"n_voladizo_bars", 0)),
            u"coronamiento_host_wall_id": cor_res.get(u"host_wall_id"),
            u"coronamiento_diam_mm": cor_res.get(u"diam_mm"),
            u"coronamiento_espesor_mm": cor_res.get(u"espesor_mm"),
            u"coronamiento_skipped": bool(cor_res.get(u"skipped")),
            u"rebars_coronamiento_ids": list(cor_res.get(u"rebars_coronamiento_ids") or []),
            u"rebars_coronamiento_id_ints": list(
                cor_res.get(u"rebars_coronamiento_id_ints") or [],
            ),
            u"rebars_coronamiento_tag_meta": list(
                cor_res.get(u"rebars_coronamiento_tag_meta") or [],
            ),
            u"messages": cor_res.get(u"messages") or [],
        })
        msgs = list(cor_res.get(u"messages") or [])
        if int(cor_res.get(u"n_fail", 0)):
            msgs.append(u"Coronamiento: fallo en creación.")
    except Exception as ex_cor:
        msgs.append(u"Coronamiento: {0}".format(ex_cor))
        embed = _merge_embed_resumen(embed, {u"n_coronamiento_fail": 1})
    return embed, msgs


def _aplicar_etiquetado_coronamiento_todos(
    doc, uidoc, embed_resumen, errores, aplicar_visibilidad=True,
):
    """Etiqueta coronamiento sup./inf. en vista activa (elevación/sección/planta).

    ``aplicar_visibilidad=False`` en unificado: Unobscured lo aplica el paso final.
    """
    if embed_resumen is None:
        return embed_resumen
    ids = embed_resumen.get(u"rebars_coronamiento_ids") or []
    id_ints = embed_resumen.get(u"rebars_coronamiento_id_ints") or []
    n_cor = (
        int(embed_resumen.get(u"n_coronamiento", 0))
        + int(embed_resumen.get(u"n_coronamiento_inferior", 0))
        + int(embed_resumen.get(u"n_coronamiento_inferior_pie", 0))
        + int(embed_resumen.get(u"n_coronamiento_voladizo", 0))
    )
    if not ids and not id_ints and n_cor < 1:
        return embed_resumen
    try:
        import armado_muros_coronamiento as _cor_mod

        tag_res = _cor_mod.aplicar_etiquetado_coronamiento(
            doc,
            {
                u"rebars_coronamiento_ids": list(ids),
                u"rebars_coronamiento_id_ints": list(id_ints),
                u"rebars_coronamiento_tag_meta": list(
                    embed_resumen.get(u"rebars_coronamiento_tag_meta") or [],
                ),
                u"n_created": int(embed_resumen.get(u"n_coronamiento", 0)),
                u"n_inferior_created": int(embed_resumen.get(u"n_coronamiento_inferior", 0)),
                u"n_inferior_pie_created": int(
                    embed_resumen.get(u"n_coronamiento_inferior_pie", 0),
                ),
                u"n_voladizo_created": int(embed_resumen.get(u"n_coronamiento_voladizo", 0)),
            },
            uidoc=uidoc,
            aplicar_visibilidad=aplicar_visibilidad,
        )
        embed_resumen = _merge_embed_resumen(embed_resumen, {
            u"n_coronamiento_tags": int(tag_res.get(u"n_cor_tags_created", 0)),
            u"n_coronamiento_tags_fail": int(tag_res.get(u"n_cor_tags_fail", 0)),
            u"messages": tag_res.get(u"messages") or [],
        })
        if int(tag_res.get(u"n_cor_tags_fail", 0)) and errores is not None:
            for m in tag_res.get(u"messages") or []:
                if m:
                    errores.append(m)
    except Exception as ex_tag:
        if errores is not None:
            errores.append(u"Etiquetas coronamiento: {0}".format(ex_tag))
    return embed_resumen


def _crear_armado_muros_unificado_impl(
    doc,
    walls,
    params_por_muro_id,
    area_type_id,
    aplicar_malla_cb,
    cabezal_por_muro_id,
    uidoc=None,
    malla_activo_por_muro_id=None,
    within_parent_transaction_group=False,
    coronamiento_cfg=None,
    coronamiento_por_muro_id=None,
):
    """
    Cuerpo del flujo unificado (cabezal, mallas, etiquetas).

    Si ``within_parent_transaction_group`` es True, el llamador ya abrió
    ``TransactionGroup`` ``Arainco: Armado Muros`` (Revit 2024).
    """
    errores = []
    embed_resumen = None
    walls_ord = ordenar_muros_por_base_asc(walls)
    cab_res = None
    params_por_muro_id = _normalize_muro_id_dict(params_por_muro_id)
    cabezal_por_muro_id = _normalize_muro_id_dict(cabezal_por_muro_id)
    malla_activo_por_muro_id = _normalize_muro_id_dict(malla_activo_por_muro_id)
    coronamiento_por_muro_id = _normalize_muro_id_dict(coronamiento_por_muro_id)

    n_cover_inicio = asignar_rebar_cover_a_muros(doc, walls_ord, errores)
    if n_cover_inicio <= 0 and walls_ord:
        errores.append(
            u"No se pudo asignar Rebar Cover al inicio (ext/int {:.0f} mm, otras {:.0f} mm).".format(
                REBAR_COVER_MM_CARAS_EXT_INT,
                REBAR_COVER_MM_OTRAS_CARAS,
            ),
        )
    embed_resumen = _merge_embed_resumen(embed_resumen, {
        u"n_rebar_cover_inicio": int(n_cover_inicio),
    })

    cor_embed, cor_msgs = _aplicar_coronamiento_en_creacion(
        doc,
        walls_ord,
        params_por_muro_id,
        coronamiento_cfg=coronamiento_cfg,
        coronamiento_por_muro_id=coronamiento_por_muro_id,
    )
    embed_resumen = _merge_embed_resumen(embed_resumen, cor_embed)
    for m in cor_msgs:
        if m and int((cor_embed or {}).get(u"n_coronamiento_fail", 0)):
            errores.append(m)

    if cabezal_por_muro_id:
        try:
            import armado_muros_cabezal as _cab_mod

            fallback_bt = _cab_mod.cabezal_resolve_bar_type_fallback(
                doc, cabezal_por_muro_id, walls_ord,
            )
            cab_res = _cab_mod.aplicar_cabezales_muros(
                doc,
                walls_ord,
                cabezal_por_muro_id,
                bar_type_fallback=fallback_bt,
                uidoc=uidoc,
                defer_etiquetado=True,
                within_parent_transaction_group=within_parent_transaction_group,
            )
            embed_resumen = _merge_embed_resumen(embed_resumen, {
                u"n_cabezal": int(cab_res.get(u"n_created", 0)),
                u"n_cabezal_fail": int(cab_res.get(u"n_fail", 0)),
                u"n_bars_total": int(cab_res.get(u"n_bars_total", 0)),
                u"n_confinement_created": int(cab_res.get(u"n_confinement_created", 0)),
                u"_cab_res": cab_res,
            })
            for m in cab_res.get(u"messages") or []:
                if m:
                    errores.append(m)
        except Exception as ex_cab:
            errores.append(u"Cabezal (barras): {0}".format(ex_cab))

    ok, err, cover_n, embed_malla = crear_areas_malla_parametrizada(
        doc,
        walls,
        params_por_muro_id,
        area_type_id,
        aplicar_malla_cb,
        muro_contencion=False,
        uidoc=uidoc,
        cabezal_por_muro_id=cabezal_por_muro_id,
        defer_etiquetas_malla=True,
        malla_activo_por_muro_id=malla_activo_por_muro_id,
        skip_coronamiento=True,
        skip_cover_assign=True,
        within_parent_transaction_group=within_parent_transaction_group,
    )
    errores.extend(err or [])
    embed_resumen = _merge_embed_resumen(embed_resumen, embed_malla)

    rebars_por_muro_id = {}
    if embed_resumen:
        rebars_por_muro_id = embed_resumen.get(u"rebars_por_muro_id") or {}

    if cab_res is not None:
        try:
            import armado_muros_cabezal as _cab_mod

            msgs_antes = len(cab_res.get(u"messages") or [])
            _cab_mod.cabezal_aplicar_etiquetado_longitudinal_pendiente(
                doc, cab_res, uidoc=uidoc,
            )
            _cab_mod.cabezal_aplicar_etiquetado_confinamiento_pendiente(
                doc, cab_res, uidoc=uidoc,
            )
            embed_resumen = _merge_embed_resumen(embed_resumen, {
                u"n_tags_created": int(cab_res.get(u"n_tags_created", 0)),
                u"n_tags_fail": int(cab_res.get(u"n_tags_fail", 0)),
                u"n_conf_tags_created": int(cab_res.get(u"n_conf_tags_created", 0)),
                u"n_conf_tags_fail": int(cab_res.get(u"n_conf_tags_fail", 0)),
                u"n_empalme_markers_ok": int(cab_res.get(u"n_empalme_markers_ok", 0)),
                u"n_empalme_markers_fail": int(cab_res.get(u"n_empalme_markers_fail", 0)),
            })
            # Mensajes de etiquetas se añaden tras el merge inicial de cab_res.
            for m in (cab_res.get(u"messages") or [])[msgs_antes:]:
                if m:
                    errores.append(m)
        except Exception as ex_tag:
            errores.append(u"Cabezal (etiquetas): {0}".format(ex_tag))

    embed_resumen = _aplicar_etiquetado_coronamiento_todos(
        doc, uidoc, embed_resumen, errores, aplicar_visibilidad=False,
    )

    embed_resumen = _aplicar_etiquetas_malla_todos(
        doc,
        uidoc,
        walls_ord,
        params_por_muro_id,
        rebars_por_muro_id,
        errores,
        embed_resumen,
        cover_asignados=int(cover_n),
        muro_contencion=False,
        stamp_pre_etiqueta=False,
    )

    vis_res = aplicar_unobscured_armado_muros_en_vista(
        doc,
        uidoc,
        cab_res=cab_res,
        embed_resumen=embed_resumen,
        rebars_malla_por_muro_id=rebars_por_muro_id,
        errores=errores,
    )
    if vis_res:
        extra_vis = {}
        if int(vis_res.get(u"n_rebars_unobscured_on", 0)):
            extra_vis[u"n_rebars_unobscured"] = int(
                vis_res.get(u"n_rebars_unobscured_on", 0),
            )
        if int(vis_res.get(u"n_rebars_unobscured_off", 0)):
            extra_vis[u"n_rebars_malla_unobscured_off"] = int(
                vis_res.get(u"n_rebars_unobscured_off", 0),
            )
        if int(vis_res.get(u"n_rebars_malla_visibles", 0)):
            extra_vis[u"n_rebars_malla_visibles"] = int(
                vis_res.get(u"n_rebars_malla_visibles", 0),
            )
        if extra_vis:
            embed_resumen = _merge_embed_resumen(embed_resumen, extra_vis)

    return ok, errores, cover_n, embed_resumen


def crear_armado_muros_unificado(
    doc,
    walls,
    params_por_muro_id,
    area_type_id,
    aplicar_malla_cb,
    cabezal_por_muro_id,
    uidoc=None,
    malla_activo_por_muro_id=None,
    coronamiento_cfg=None,
    coronamiento_por_muro_id=None,
):
    """
    Flujo unificado (solo muro tradicional), orden constructivo:
    1) Barras longitudinales (cabezal).
    2) Confinamiento (cabezal).
    3) Mallas (AR + Remove System + post-proceso).
    4) Etiquetas barras longitudinales (+ marcadores empalme).
    5) Etiquetas confinamiento.
    6) Etiquetas coronamiento (sup./inf.).
    7) Etiquetas mallas.
    8) Visibilidad en vista activa: malla visible + unobscured OFF;
       cabezal/coronamiento con unobscured ON.

    Envuelve todo en ``TransactionGroup`` ``Arainco: Armado Muros`` (un paso Deshacer).
    Los grupos y transacciones por lote internos (cabezal, lineales, etiquetas, etc.)
    se conservan sin cambios.
    """
    ok = []
    errores = []
    cover_n = 0
    embed_resumen = None
    conjunto_guid = None

    use_outer_tg = _unificado_usa_transaction_group_externo(doc)
    within_parent = bool(use_outer_tg)
    tg = None
    tg_started = False
    flujo_ok = False

    try:
        try:
            from armado_muros_rebar_params import iniciar_armadura_conjunto_guid_ejecucion
            conjunto_guid = iniciar_armadura_conjunto_guid_ejecucion()
        except Exception:
            conjunto_guid = None
        try:
            from armado_muros_rebar_params import iniciar_armadura_eje_ejecucion
            iniciar_armadura_eje_ejecucion(uidoc=uidoc)
        except Exception:
            pass
        if use_outer_tg:
            tg = TransactionGroup(doc, TXN_GROUP_ARMADO_MUROS_UNIFICADO)
            tg.Start()
            tg_started = True
        ok, errores, cover_n, embed_resumen = _crear_armado_muros_unificado_impl(
            doc,
            walls,
            params_por_muro_id,
            area_type_id,
            aplicar_malla_cb,
            cabezal_por_muro_id,
            uidoc=uidoc,
            malla_activo_por_muro_id=malla_activo_por_muro_id,
            within_parent_transaction_group=within_parent,
            coronamiento_cfg=coronamiento_cfg,
            coronamiento_por_muro_id=coronamiento_por_muro_id,
        )
        flujo_ok = True
    except Exception as ex:
        errores.append(u"Armado Muros (grupo transacciones): {0}".format(ex))
        flujo_ok = False
    finally:
        try:
            from armado_muros_rebar_params import finalizar_armadura_conjunto_guid_ejecucion
            finalizar_armadura_conjunto_guid_ejecucion()
        except Exception:
            pass
        try:
            from armado_muros_rebar_params import finalizar_armadura_eje_ejecucion
            finalizar_armadura_eje_ejecucion()
        except Exception:
            pass
        if tg_started:
            try:
                if flujo_ok:
                    tg.Assimilate()
                else:
                    tg.RollBack()
            except Exception:
                pass

    if flujo_ok and uidoc is not None:
        _refrescar_vista_fin_flujo(doc, uidoc)

    if conjunto_guid:
        embed_resumen = _merge_embed_resumen(
            embed_resumen or {},
            {u"armadura_conjunto_guid": conjunto_guid},
        )

    return ok, errores, cover_n, embed_resumen


def _open_preview_ui(revit, uidoc, walls, mode):
    try:
        import os as _os_ui
        import sys as _sys_ui

        _lib_dir = _os_ui.path.dirname(_os_ui.path.abspath(__file__))
        if _lib_dir and _lib_dir not in _sys_ui.path:
            _sys_ui.path.insert(0, _lib_dir)

        _preview_path = _os_ui.path.join(_lib_dir, u"armado_muros_preview_ui.py")
        if not _os_ui.path.isfile(_preview_path):
            TaskDialog.Show(
                u"Armado muros — Error",
                u"No se encuentra armado_muros_preview_ui.py en:\n{}".format(_lib_dir),
            )
            return

        import armado_muros_preview_ui as _ui_mod

        if mode == _ui_mod.UI_MODE_CABEZAL:
            _ui_mod.show_armado_muros_cabezal(revit, uidoc, walls)
        elif mode == _ui_mod.UI_MODE_UNIFICADO:
            _ui_mod.show_armado_muros_unificado(revit, uidoc, walls)
        else:
            _ui_mod.show_armado_muros_mallas(revit, uidoc, walls)
    except Exception as ex_ui:
        try:
            _detail = unicode(ex_ui)
        except Exception:
            _detail = str(ex_ui)
        TaskDialog.Show(
            u"Armado muros — Error",
            u"No se pudo cargar o abrir la interfaz.\n\n{}".format(_detail),
        )


def _guard_vista_armado_muros(uidoc, uiapp=None):
    """Bloquea el flujo si la vista activa no es Building Section."""
    view = None
    try:
        view = uidoc.ActiveView if uidoc is not None else None
    except Exception:
        view = None
    try:
        from armado_muros_etiqueta_malla import es_vista_building_section
        from armado_muros_instruction_dialog import show_building_section_view_required

        if es_vista_building_section(view):
            return True
        show_building_section_view_required(view, uiapp=uiapp)
    except Exception:
        try:
            from armado_muros_etiqueta_malla import mensaje_vista_requerida_armado_muros

            TaskDialog.Show(u"Arainco: Armado Muros v3", mensaje_vista_requerida_armado_muros(view))
        except Exception:
            TaskDialog.Show(
                u"Arainco: Armado Muros",
                u"Esta herramienta solo puede ejecutarse en secciones "
                u"tipo Building Section.",
            )
    return False


def run_cabezal(revit):
    uidoc = revit.ActiveUIDocument
    if uidoc is None:
        TaskDialog.Show(u"Armado muros cabezal — Error", u"No hay ventana Revit activa.")
        return
    if not _guard_vista_armado_muros(uidoc, revit):
        return
    walls = _pick_muros(uidoc)
    if walls is None:
        return
    if not walls and not getattr(_pick_muros, "_rechazo_paralelo", False):
        TaskDialog.Show(u"Armado muros cabezal", u"No seleccionaste muros válidos.")
        return
    if not walls:
        return
    _open_preview_ui(revit, uidoc, walls, u"cabezal")


def run_mallas(revit):
    uidoc = revit.ActiveUIDocument
    if uidoc is None:
        TaskDialog.Show(u"Armado muros mallas — Error", u"No hay ventana Revit activa.")
        return
    if not _guard_vista_armado_muros(uidoc, revit):
        return
    walls = _pick_muros(uidoc)
    if walls is None:
        return
    if not walls and not getattr(_pick_muros, "_rechazo_paralelo", False):
        TaskDialog.Show(u"Armado muros mallas", u"No seleccionaste muros válidos.")
        return
    if not walls:
        return
    _open_preview_ui(revit, uidoc, walls, u"mallas")


def run_unificado(revit):
    uidoc = revit.ActiveUIDocument
    if uidoc is None:
        TaskDialog.Show(u"Arainco: Armado Muros v3 — Error", u"No hay ventana Revit activa.")
        return
    if not _guard_vista_armado_muros(uidoc, revit):
        return
    from armado_muros_instruction_dialog import show_selection_instructions

    if not show_selection_instructions(revit):
        return
    walls = _pick_muros(uidoc)
    if walls is None:
        return
    if not walls and not getattr(_pick_muros, "_rechazo_paralelo", False):
        TaskDialog.Show(u"Arainco: Armado Muros v3", u"No seleccionaste muros válidos.")
        return
    if not walls:
        return
    _open_preview_ui(revit, uidoc, walls, u"unificado")


def run(revit):
    """Compatibilidad: abre el asistente unificado."""
    run_unificado(revit)
