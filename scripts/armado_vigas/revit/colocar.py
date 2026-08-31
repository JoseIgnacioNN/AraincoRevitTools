# -*- coding: utf-8 -*-
"""Colocación de armadura en vigas — longitudinales y estribos/confinamiento."""

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import (
    ExternalEvent,
    IExternalEventHandler,
)

from armado_vigas.ui.instruction_dialog import (
    DIALOG_TITLE,
    show_info,
    show_yes_no,
)

from bimtools_rebar_3d_visibility import (
    apply_rebar_unobscured_in_view,
    apply_reinforcement_unobscured_in_view,
)

from armado_vigas.revit.colocar_rebar import (
    colocar_armadura_longitudinal,
    find_longitudinal_guides_over_limit,
)
from armado_vigas.revit.colocar_lap_detail import colocar_marcadores_empalme_vigas
from armado_vigas.revit.colocar_estribos import colocar_estribos_confinamiento
from armado_vigas.revit.colocar_laterales import colocar_laterales
from armado_vigas.revit.colocar_progress import ColocarArmaduraProgress
from armado_vigas.revit.etiquetar_confinamiento import (
    etiquetar_confinamiento_en_vista,
    reset_inferior_lap_dim_host_registry,
)
from armado_vigas.revit.etiquetar_laterales import etiquetar_laterales_en_vista
from armado_vigas.revit.etiquetar_longitudinales import (
    etiquetar_longitudinales_en_vista,
    realinear_longitudinales_inf_tras_confinamiento,
)
from armado_vigas.domain.tramos import build_session_tramos, sort_beams
from armado_vigas.revit.session import SESSION
from armado_vigas.revit.txn import transaction_group_scope, transaction_scope

try:
    from armado_vigas.revit.armadura_conjunto_guid import (
        aplicar_conjunto_guid_elementos_creados,
        finalizar_corrida_conjunto_guid,
        iniciar_corrida_conjunto_guid,
    )
except Exception:
    aplicar_conjunto_guid_elementos_creados = None
    finalizar_corrida_conjunto_guid = None
    iniciar_corrida_conjunto_guid = None

try:
    from armado_vigas.revit.armadura_ubicacion import (
        aplicar_armadura_capa_longitudinales,
        aplicar_armadura_eje,
        aplicar_armadura_en_lamina,
        aplicar_armadura_ubicacion_laterales,
        aplicar_armadura_ubicacion_longitudinales,
        aplicar_marca_parametros_armado_vigas,
    )
except Exception:
    aplicar_armadura_capa_longitudinales = None
    aplicar_armadura_eje = None
    aplicar_armadura_en_lamina = None
    aplicar_armadura_ubicacion_laterales = None
    aplicar_armadura_ubicacion_longitudinales = None
    aplicar_marca_parametros_armado_vigas = None


def _restore_colocar_window(win):
    if win is not None and hasattr(win, u"restore_after_colocar"):
        try:
            win.restore_after_colocar()
        except Exception:
            pass


def _hide_colocar_window(win):
    if win is None:
        return
    if hasattr(win, u"hide_for_colocar_on_ui"):
        try:
            win.hide_for_colocar_on_ui()
            return
        except Exception:
            pass
    if hasattr(win, u"hide_for_colocar"):
        try:
            win.hide_for_colocar()
        except Exception:
            pass


def _format_exc(ex):
    try:
        return unicode(ex)
    except NameError:
        return str(ex)


def _collect_placed_rebars(*groups):
    """Lista única de Rebar creados en la corrida (sin duplicados por Id)."""
    out = []
    seen = set()
    for group in groups:
        for rb in group or []:
            if rb is None:
                continue
            try:
                eid = int(rb.Id.IntegerValue)
            except Exception:
                try:
                    eid = int(rb)
                except Exception:
                    eid = None
            if eid is not None:
                if eid in seen:
                    continue
                seen.add(eid)
            out.append(rb)
    return out


def _rebar_element_id_int(rb):
    if rb is None:
        return None
    try:
        return int(rb.Id.IntegerValue)
    except Exception:
        try:
            return int(rb)
        except Exception:
            return None


def _rebar_id_set(rebars):
    """Conjunto de ElementId.IntegerValue de rebars."""
    ids = set()
    for rb in rebars or []:
        eid = _rebar_element_id_int(rb)
        if eid is not None:
            ids.add(eid)
    return ids


def _collect_for_unobscured(rebars, rebars_lat=None):
    """Long. + estribos/conf.: excluye laterales (no llevan View Unobscured)."""
    exclude = _rebar_id_set(rebars_lat)
    if not exclude:
        return _collect_placed_rebars(rebars)
    out = []
    for rb in _collect_placed_rebars(rebars):
        eid = _rebar_element_id_int(rb)
        if eid is None or eid not in exclude:
            out.append(rb)
    return out


def _fresh_view(doc, view):
    if doc is None or view is None:
        return view
    try:
        v2 = doc.GetElement(view.Id)
        if v2 is not None:
            return v2
    except Exception:
        pass
    return view


def _apply_view_obscured_laterales(doc, view, rebars_lat, avisos=None):
    """
    Barras laterales del alma: Unobscured OFF y sólido OFF.

    Usa ids frescos del documento (no solo refs en memoria) para no fallar
    tras regenerate / pata L / re-host.
    """
    if doc is None or view is None or not rebars_lat:
        return 0
    from Autodesk.Revit.DB import ElementId

    try:
        from Autodesk.Revit.DB.Structure import Rebar
    except Exception:
        Rebar = None
    try:
        from bimtools_rebar_3d_visibility import ensure_rebar_obscured_in_view
    except Exception:
        ensure_rebar_obscured_in_view = None

    view = _fresh_view(doc, view)
    lat_ids = list(_rebar_id_set(rebars_lat))
    if not lat_ids:
        return 0

    fresh = []
    for eid in lat_ids:
        try:
            el = doc.GetElement(ElementId(int(eid)))
        except Exception:
            el = None
        if el is None:
            continue
        if Rebar is not None:
            try:
                if not isinstance(el, Rebar):
                    continue
            except Exception:
                pass
        fresh.append(el)

    if not fresh:
        return 0

    # Helper BIMTools: Unobscured OFF
    if ensure_rebar_obscured_in_view is not None:
        try:
            ensure_rebar_obscured_in_view(doc, fresh, view)
        except Exception:
            pass

    n_ok = 0
    for el in fresh:
        try:
            el = doc.GetElement(el.Id) or el
        except Exception:
            pass
        for attempt in range(2):
            try:
                el.SetUnobscuredInView(view, False)
                n_ok += 1
                break
            except Exception:
                if attempt == 0:
                    try:
                        doc.Regenerate()
                        el = doc.GetElement(el.Id) or el
                        view = _fresh_view(doc, view)
                    except Exception:
                        pass
        try:
            el.SetSolidInView(view, False)
        except Exception:
            try:
                fn = getattr(el, u"SetSolidInView", None)
                if fn is not None:
                    fn(view, False)
            except Exception:
                pass
        # Verificar; si quedó ON, reintentar
        try:
            if bool(el.IsUnobscuredInView(view)):
                el.SetUnobscuredInView(view, False)
        except Exception:
            pass

    try:
        doc.Regenerate()
    except Exception:
        pass
    return n_ok


def _apply_view_unobscured_active(
    doc,
    view,
    rebars,
    avisos=None,
    exclude_ids=None,
    presentation_first_last_ids=None,
):
    """
    Hace visibles en la vista activa Rebar de la corrida (long. / conf.).

    ``exclude_ids``: ids enteros de laterales (u otros) que NUNCA reciben Unobscured ON.
    """
    if doc is None or view is None or not rebars:
        return 0

    from Autodesk.Revit.DB import BuiltInCategory, Category, ElementId
    from System.Collections.Generic import List

    try:
        from Autodesk.Revit.DB.Structure import Rebar, RebarPresentationMode
    except Exception:
        Rebar = None
        RebarPresentationMode = None

    exclude_ids = set(exclude_ids or set())
    first_last_ids = set(presentation_first_last_ids or set())

    view = _fresh_view(doc, view)

    # Releer elementos frescos del documento; filtrar laterales por id.
    fresh = []
    id_list = List[ElementId]()
    for rb in rebars:
        el = None
        try:
            if hasattr(rb, u"Id"):
                el = doc.GetElement(rb.Id)
            else:
                el = doc.GetElement(ElementId(int(rb)))
        except Exception:
            el = rb if rb is not None else None
        if el is None:
            continue
        try:
            eid = int(el.Id.IntegerValue)
            if eid in exclude_ids:
                continue
        except Exception:
            pass
        if Rebar is not None:
            try:
                if not isinstance(el, Rebar):
                    el2 = doc.GetElement(el.Id)
                    if el2 is not None:
                        el = el2
            except Exception:
                pass
            # Doble check exclude tras re-get
            try:
                if int(el.Id.IntegerValue) in exclude_ids:
                    continue
            except Exception:
                pass
        fresh.append(el)
        try:
            id_list.Add(el.Id)
        except Exception:
            pass
    if not fresh:
        return 0

    # 0) Salir de Isolate/Hide temporal (si no, Unhide no pinta).
    try:
        from Autodesk.Revit.DB import TemporaryViewMode

        for mode_name in (
            u"TemporaryHideIsolate",
            u"RevealHiddenElements",
            u"TemporaryViewProperties",
        ):
            try:
                mode = getattr(TemporaryViewMode, mode_name, None)
                if mode is None:
                    continue
                if view.IsInTemporaryViewMode(mode):
                    view.DisableTemporaryViewMode(mode)
            except Exception:
                pass
    except Exception:
        pass

    # 1) Categoría Structural Rebar visible en la vista.
    try:
        cat = Category.GetCategory(doc, BuiltInCategory.OST_Rebar)
        if cat is None:
            try:
                cat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Rebar)
            except Exception:
                cat = None
        if cat is not None:
            try:
                view.SetCategoryHidden(cat.Id, False)
            except Exception:
                pass
    except Exception:
        pass

    # 2) Revocar Hide in View → Elements (solo los no-laterales).
    if id_list.Count > 0:
        try:
            view.UnhideElements(id_list)
        except Exception:
            pass

    try:
        doc.Regenerate()
    except Exception:
        pass

    # 3–4) Presentación + Unobscured + sólido (directo; helpers como refuerzo).
    n_ok = 0
    for el in fresh:
        if el is None:
            continue
        try:
            el = doc.GetElement(el.Id) or el
        except Exception:
            pass
        try:
            if int(el.Id.IntegerValue) in exclude_ids:
                continue
        except Exception:
            pass
        if RebarPresentationMode is not None:
            try:
                mode = RebarPresentationMode.All
                try:
                    if int(el.Id.IntegerValue) in first_last_ids:
                        mode = RebarPresentationMode.FirstLast
                except Exception:
                    pass
                if el.CanApplyPresentationMode(view):
                    el.SetPresentationMode(view, mode)
            except Exception:
                pass
        try:
            el.SetUnobscuredInView(view, True)
            n_ok += 1
        except Exception:
            pass
        try:
            el.SetSolidInView(view, True)
        except Exception:
            try:
                fn = getattr(el, u"SetSolidInView", None)
                if fn is not None:
                    fn(view, True)
            except Exception:
                pass

    # Fallback batch (p. ej. RebarInSystem / wrappers) — mismo filtro de exclusión.
    if n_ok <= 0:
        try:
            n_ok = int(
                apply_reinforcement_unobscured_in_view(
                    doc, fresh, view, unobscured=True, solid_in_view=True
                )
                or 0
            )
        except Exception:
            n_ok = 0
        if n_ok <= 0:
            try:
                apply_rebar_unobscured_in_view(doc, fresh, view)
                n_ok = len(fresh)
            except Exception:
                n_ok = 0

    try:
        doc.Regenerate()
    except Exception:
        pass

    # Tras regenerate, re-forzar OFF de laterales por si la API las activó al pintar.
    if exclude_ids:
        try:
            from Autodesk.Revit.DB.Structure import Rebar as _Rebar

            for eid in exclude_ids:
                try:
                    el = doc.GetElement(ElementId(int(eid)))
                except Exception:
                    el = None
                if el is None:
                    continue
                try:
                    if _Rebar is not None and not isinstance(el, _Rebar):
                        continue
                except Exception:
                    pass
                try:
                    el.SetUnobscuredInView(view, False)
                except Exception:
                    pass
                try:
                    el.SetSolidInView(view, False)
                except Exception:
                    pass
        except Exception:
            pass

    if avisos is not None:
        if n_ok > 0:
            avisos.append(
                u"Barras visibles en vista activa: {0}/{1}.".format(
                    int(n_ok), int(len(fresh))
                )
            )
        else:
            avisos.append(
                u"Aviso: no se pudo forzar visibilidad de rebar "
                u"en la vista activa (revisar VG / Unobscured)."
            )
    return n_ok


def _reapply_visibility_after_commit(
    doc,
    view,
    rebars,
    avisos=None,
    exclude_ids=None,
    presentation_first_last_ids=None,
):
    """
    Segunda pasada en transacción propia tras Commit (por si la grafía
    de la vista no quedó actualizada dentro de la txn larga).

    Debe invocarse **dentro** del ``TransactionGroup`` de armado para un solo Undo.
    """
    if doc is None or view is None or not rebars:
        return 0
    try:
        with transaction_scope(doc, u"Arainco: Visibilidad rebar vista"):
            return _apply_view_unobscured_active(
                doc,
                view,
                rebars,
                avisos=None,
                exclude_ids=exclude_ids,
                presentation_first_last_ids=presentation_first_last_ids,
            )
    except Exception as ex:
        if avisos is not None:
            try:
                avisos.append(
                    u"Aviso re-visibilidad post-commit: {0}".format(_format_exc(ex))
                )
            except Exception:
                pass
        return 0


def prompt_longitudinals_over_limit(over_limit):
    """
    Aviso de barras > 12 m con la ventana aún visible.

    Returns:
        True para continuar; False para cancelar.
    """
    if not over_limit:
        return True
    detail_lines = [
        u"· {0}: ≈ {1:.0f} mm".format(
            v[u"label"], float(v.get(u"length_mm") or 0.0)
        )
        for v in over_limit[:8]
    ]
    if len(over_limit) > 8:
        detail_lines.append(
            u"· … y {0} guía(s) más.".format(len(over_limit) - 8)
        )
    content = (
        u"Marque empalmes (Traslape sup/inf) en el canvas para trocear "
        u"las fibras y evitar barras mayores a 12 m.\n\n"
        + u"\n".join(detail_lines)
        + u"\n\n¿Desea colocar la armadura de todos modos?"
    )
    return show_yes_no(
        u"Hay barras longitudinales que superan 12 m",
        content=content,
        title=DIALOG_TITLE,
        yes_text=u"Sí, colocar",
        no_text=u"Cancelar",
    )


class ColocarArmaduraHandler(IExternalEventHandler):
    """
    Handler de proceso (AppDomain).

    No recrear ExternalEvent por cada apertura de ventana: en IronPython eso
    puede tumbar Revit en la 2.ª ejecución. El target se asigna antes de Raise.
    """

    def Execute(self, uiapp):
        win = _resolve_colocar_window()
        if win is None:
            try:
                show_info(
                    u"No se pudo acceder a la ventana de la herramienta.",
                    u"Cierre y vuelva a abrir Armado vigas.",
                    title=DIALOG_TITLE,
                    uiapp=uiapp,
                )
            except Exception:
                pass
            return
        try:
            self._execute_colocar(uiapp, win)
        except Exception as ex:
            msg = _format_exc(ex)
            try:
                win.set_status(u"Error: {0}".format(msg))
            except Exception:
                pass
            _restore_colocar_window(win)
            try:
                show_info(
                    u"No se pudo colocar la armadura.",
                    msg,
                    title=DIALOG_TITLE,
                    uiapp=uiapp,
                )
            except Exception:
                pass

    def _execute_colocar(self, uiapp, win):
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            win.set_status(u"Sin documento activo.")
            _restore_colocar_window(win)
            return
        doc = uidoc.Document
        if not SESSION.framing_elements:
            win.set_status(u"No hay vigas en el lote.")
            _restore_colocar_window(win)
            return

        try:
            win.set_status(u"Auditando longitudes…")
        except Exception:
            pass

        try:
            sorted_beams = sort_beams(list(SESSION.domain_beams or []))
            SESSION.tramos_sup, SESSION.tramos_inf = build_session_tramos(
                sorted_beams,
                empalme_beam_ids_sup=SESSION.empalme_beam_ids_sup,
                empalme_beam_ids_inf=SESSION.empalme_beam_ids_inf,
                split_empalme=SESSION.split_empalme,
            )
            SESSION.tramos = SESSION.tramos_sup
            over_limit = find_longitudinal_guides_over_limit(doc, SESSION)
        except Exception as ex:
            msg = _format_exc(ex)
            win.set_status(u"Error al auditar longitudes: {0}".format(msg))
            _restore_colocar_window(win)
            try:
                show_info(
                    u"No se pudo verificar la longitud de las barras.",
                    msg,
                    title=DIALOG_TITLE,
                    uiapp=uiapp,
                )
            except Exception:
                pass
            return

        if not prompt_longitudinals_over_limit(over_limit):
            win.set_status(u"Colocación cancelada: hay barras longitudinales > 12 m.")
            _restore_colocar_window(win)
            return

        # Por si Raise llegó con la ventana aún visible.
        _hide_colocar_window(win)

        view = uidoc.ActiveView
        n_lap_details = 0
        n_lap_dims = 0
        lap_res = {}
        rebars = []
        rebars_lat = []
        rebars_est = []
        conf_rebar_ids = set()
        conjunto_guid = None
        if iniciar_corrida_conjunto_guid is not None:
            conjunto_guid = iniciar_corrida_conjunto_guid()

        n_bars = 0
        n_tags = 0
        n_est = 0
        n_conf_tags = 0
        n_lat = 0
        n_lat_tags = 0
        avisos = []

        try:
            with ColocarArmaduraProgress(SESSION) as progress:
                # Un solo Undo: TransactionGroup.Assimilate (pyRevit 2025+).
                with transaction_group_scope(
                    doc, u"Arainco: Armado vigas", assimilate=True
                ):
                    with transaction_scope(doc, u"Arainco: Armado vigas"):
                        progress.step(u"longitudinales")
                        n_bars, avisos, rebars, long_by_side, lap_jobs = (
                            colocar_armadura_longitudinal(doc, SESSION)
                        )

                        progress.step(u"parámetros")
                        if aplicar_armadura_ubicacion_longitudinales is not None:
                            aplicar_armadura_ubicacion_longitudinales(long_by_side)
                        if aplicar_armadura_capa_longitudinales is not None:
                            aplicar_armadura_capa_longitudinales(long_by_side)

                        reset_inferior_lap_dim_host_registry()
                        progress.step(u"empalmes")
                        if lap_jobs and view is not None:
                            lap_res = colocar_marcadores_empalme_vigas(
                                doc, view, lap_jobs
                            )
                            n_lap_details = int(lap_res.get(u"n_ok") or 0)
                            n_lap_dims = int(lap_res.get(u"n_dims_ok") or 0)
                            for msg in lap_res.get(u"messages") or []:
                                if msg:
                                    avisos.append(msg)
                            if n_lap_details > 0:
                                avisos.append(
                                    u"Detail Items de traslape: {0}.".format(
                                        n_lap_details
                                    )
                                )
                            if n_lap_dims > 0:
                                avisos.append(
                                    u"Cotas de traslape: {0}.".format(n_lap_dims)
                                )

                        progress.step(u"etiquetas longitudinales")
                        if rebars and view is not None:
                            n_tags, avisos_tag, err_tag = (
                                etiquetar_longitudinales_en_vista(
                                    doc,
                                    view,
                                    rebars,
                                    use_transaction=False,
                                    rebars_by_side=long_by_side,
                                )
                            )
                            if avisos_tag:
                                avisos.extend(avisos_tag)
                            if err_tag:
                                avisos.append(err_tag)

                        progress.step(u"estribos y confinamiento")
                        if bool(getattr(SESSION, u"placeConf", True)):
                            n_est, avisos_est, rebars_est, conf_tag_jobs = (
                                colocar_estribos_confinamiento(
                                    doc, SESSION, view=view
                                )
                            )
                            avisos.extend(avisos_est or [])
                            rebars.extend(rebars_est or [])

                            conf_rebar_ids = _rebar_id_set(rebars_est)

                            progress.step(u"etiquetas confinamiento")
                            if conf_tag_jobs and view is not None:
                                n_conf_tags, avisos_conf, err_conf = (
                                    etiquetar_confinamiento_en_vista(
                                        doc,
                                        view,
                                        conf_tag_jobs,
                                        use_transaction=False,
                                    )
                                )
                                if avisos_conf:
                                    avisos.extend(avisos_conf)
                                if err_conf:
                                    avisos.append(err_conf)
                            if long_by_side and view is not None:
                                try:
                                    realinear_longitudinales_inf_tras_confinamiento(
                                        doc, view, long_by_side,
                                    )
                                except Exception:
                                    pass
                        else:
                            avisos.append(
                                u"CONF desactivado · sin estribos/confinamiento."
                            )
                            conf_tag_jobs = None
                            conf_rebar_ids = set()

                        if getattr(SESSION, "lateralesEnabled", False):
                            # Unobscured de long./conf. ANTES de crear laterales.
                            progress.step(u"visibilidad")
                            vis_pre = _collect_placed_rebars(rebars)
                            if vis_pre and view is not None:
                                _apply_view_unobscured_active(
                                    doc,
                                    view,
                                    vis_pre,
                                    avisos=avisos,
                                    exclude_ids=None,
                                    presentation_first_last_ids=conf_rebar_ids,
                                )

                            progress.step(u"laterales")
                            n_lat, avisos_lat, rebars_lat, err_lat = colocar_laterales(
                                doc, SESSION, view=view
                            )
                            if avisos_lat:
                                avisos.extend(avisos_lat)
                            if err_lat:
                                avisos.append(err_lat)
                            if (
                                rebars_lat
                                and aplicar_armadura_ubicacion_laterales is not None
                            ):
                                try:
                                    n_ub_lat = aplicar_armadura_ubicacion_laterales(
                                        rebars_lat
                                    )
                                    if n_ub_lat > 0:
                                        avisos.append(
                                            u"Armadura_Ubicacion=Lateral en {0} barra(s).".format(
                                                n_ub_lat
                                            )
                                        )
                                except Exception:
                                    pass
                            # Tras crear: ON por defecto de Revit → forzar OFF.
                            if rebars_lat and view is not None:
                                _apply_view_obscured_laterales(
                                    doc, view, rebars_lat, avisos=None
                                )
                            if rebars_lat and view is not None:
                                n_lat_tags, avisos_lat_tag, err_lat_tag = (
                                    etiquetar_laterales_en_vista(
                                        doc,
                                        view,
                                        rebars_lat,
                                        framing_elements=SESSION.framing_elements,
                                        use_transaction=False,
                                    )
                                )
                                if avisos_lat_tag:
                                    avisos.extend(avisos_lat_tag)
                                if err_lat_tag:
                                    avisos.append(err_lat_tag)
                                _apply_view_obscured_laterales(
                                    doc, view, rebars_lat, avisos=None
                                )

                        lat_ids = _rebar_id_set(rebars_lat)

                        progress.step(u"finalización")
                        # View Unobscured: long. + estribos/conf. — NUNCA laterales.
                        vis_rebars = _collect_for_unobscured(rebars, rebars_lat)
                        if vis_rebars and view is not None:
                            _apply_view_unobscured_active(
                                doc,
                                view,
                                vis_rebars,
                                avisos=avisos,
                                exclude_ids=lat_ids,
                                presentation_first_last_ids=conf_rebar_ids,
                            )
                        if rebars_lat and view is not None:
                            _apply_view_obscured_laterales(
                                doc, view, rebars_lat, avisos=None
                            )
                        all_rebars = _collect_placed_rebars(rebars, rebars_lat)
                        # Marca Arainco: Eje · Arainco=Yes · Malla=No · Nivel (host viga).
                        if (
                            aplicar_marca_parametros_armado_vigas is not None
                            and all_rebars
                        ):
                            try:
                                stamp_stats = aplicar_marca_parametros_armado_vigas(
                                    rebars,
                                    view=view,
                                    rebars_laterales=rebars_lat,
                                    document=doc,
                                )
                                n_st = int((stamp_stats or {}).get(u"arainco") or 0)
                                if n_st > 0:
                                    avisos.append(
                                        u"Parámetros Arainco estampados en {0} barra(s).".format(
                                            n_st
                                        )
                                    )
                            except Exception as ex_stamp:
                                try:
                                    avisos.append(
                                        u"Aviso estampa parámetros: {0}".format(
                                            _format_exc(ex_stamp)
                                        )
                                    )
                                except Exception:
                                    pass
                        # Compat: Lámina / Eje sueltos por si la marca unificada no corrió.
                        if aplicar_armadura_en_lamina is not None:
                            aplicar_armadura_en_lamina(
                                rebars, view, rebars_laterales=rebars_lat
                            )
                        if (
                            aplicar_armadura_eje is not None
                            and aplicar_marca_parametros_armado_vigas is None
                        ):
                            aplicar_armadura_eje(
                                rebars, view, rebars_laterales=rebars_lat
                            )
                        if aplicar_conjunto_guid_elementos_creados is not None:
                            aplicar_conjunto_guid_elementos_creados(
                                doc,
                                view,
                                rebars,
                                rebars_laterales=rebars_lat,
                                lap_result=lap_res,
                                conjunto_guid=conjunto_guid,
                            )
                        # Tras estampas / GUID: reasegurar laterales sin Unobscured.
                        if rebars_lat and view is not None:
                            _apply_view_obscured_laterales(
                                doc, view, rebars_lat, avisos=None
                            )

                    # Reaplicar visibilidad en txn hija del mismo grupo (mismo Undo).
                    try:
                        lat_ids = _rebar_id_set(rebars_lat)
                        vis_after = _collect_for_unobscured(rebars, rebars_lat)
                        view_vis = view
                        try:
                            av = uidoc.ActiveView
                            if av is not None:
                                view_vis = av
                        except Exception:
                            pass
                        if vis_after and view_vis is not None:
                            _reapply_visibility_after_commit(
                                doc,
                                view_vis,
                                vis_after,
                                avisos=avisos,
                                exclude_ids=lat_ids,
                                presentation_first_last_ids=conf_rebar_ids,
                            )
                        # Vista de trabajo (start) y activa: OFF en laterales.
                        views_off = []
                        seen_vids = set()
                        for vv in (view, view_vis):
                            if vv is None:
                                continue
                            try:
                                vid = int(vv.Id.IntegerValue)
                            except Exception:
                                continue
                            if vid in seen_vids:
                                continue
                            seen_vids.add(vid)
                            views_off.append(vv)
                        if rebars_lat and views_off:
                            try:
                                with transaction_scope(
                                    doc, u"Arainco: Laterales sin Unobscured"
                                ):
                                    for vv in views_off:
                                        _apply_view_obscured_laterales(
                                            doc, vv, rebars_lat, avisos=None
                                        )
                            except Exception:
                                pass
                    except Exception:
                        pass
        except Exception as ex:
            try:
                msg = unicode(ex)
            except NameError:
                msg = str(ex)
            win.set_status(u"Error: {0}".format(msg))
            _restore_colocar_window(win)
            try:
                show_info(
                    u"No se pudo colocar la armadura.",
                    msg,
                    title=DIALOG_TITLE,
                    uiapp=uiapp,
                )
            except Exception:
                pass
            return
        finally:
            if finalizar_corrida_conjunto_guid is not None:
                finalizar_corrida_conjunto_guid()

        msg = u"Rebar: {0} barra(s) longitudinales".format(n_bars)
        if n_lap_details > 0:
            msg += u", {0} empalme(s)".format(n_lap_details)
        if n_lap_dims > 0:
            msg += u", {0} cota(s) traslape".format(n_lap_dims)
        n_etiq_total = int(n_tags or 0) + int(n_conf_tags or 0) + int(n_lat_tags or 0)
        if n_etiq_total > 0:
            msg += u", {0} etiqueta(s)".format(n_etiq_total)
        if n_est > 0:
            msg += u", {0} pos. estribo/confin.".format(n_est)
        if n_lat > 0:
            msg += u", {0} pos. laterales".format(n_lat)
        msg += u"."
        tie_avisos = [
            a for a in (avisos or [])
            if u"traba" in (a or u"").lower() or u"Trabas" in (a or u"")
        ]
        if avisos:
            msg += u" · {0} aviso(s).".format(len(avisos))
            if len(avisos) <= 2:
                msg += u" " + u" · ".join(avisos[:2])
        win.set_status(msg)
        if tie_avisos:
            try:
                show_info(
                    u"Problemas con trabas de confinamiento",
                    u"Las barras longitudinales y/o estribos se colocaron, "
                    u"pero hubo problemas con trabas de confinamiento:\n\n"
                    + u"\n".join(tie_avisos[:8]),
                    title=DIALOG_TITLE,
                    uiapp=uiapp,
                )
            except Exception:
                pass
        elif n_bars <= 0 and n_est <= 0 and n_lat <= 0:
            detail = msg
            if avisos:
                detail += u"\n\n" + u"\n".join(avisos[:8])
            try:
                show_info(
                    u"No se colocó armadura.",
                    detail,
                    title=DIALOG_TITLE,
                    uiapp=uiapp,
                )
            except Exception:
                pass
            _restore_colocar_window(win)
            return
        win.request_close()

    def GetName(self):
        return u"ArmadoVigasColocarRebar"


_COLOCAR_TARGET = None
_COLOCAR_EVENT = None
_COLOCAR_HANDLER = None


def _resolve_colocar_window():
    global _COLOCAR_TARGET
    win = _COLOCAR_TARGET
    if win is not None:
        return win
    try:
        from armado_vigas.ui.window import get_existing_armado_vigas_window

        return get_existing_armado_vigas_window()
    except Exception:
        return None


def ensure_colocar_event():
    """Un solo ExternalEvent por AppDomain. Nunca Dispose."""
    global _COLOCAR_EVENT, _COLOCAR_HANDLER
    if _COLOCAR_EVENT is None:
        _COLOCAR_HANDLER = ColocarArmaduraHandler()
        _COLOCAR_EVENT = ExternalEvent.Create(_COLOCAR_HANDLER)
    return _COLOCAR_EVENT


def set_colocar_target(win):
    global _COLOCAR_TARGET
    _COLOCAR_TARGET = win


def clear_colocar_target(win=None):
    global _COLOCAR_TARGET
    if win is None or _COLOCAR_TARGET is win:
        _COLOCAR_TARGET = None


def raise_colocar_armadura(win):
    """
    Oculta la UI, asigna target y encola ExternalEvent.

    Revit no procesa Raise mientras el foco queda atrapado en WPF modeless
    maximizado; ocultar primero libera el idle del host.
    """
    set_colocar_target(win)
    try:
        if win is not None and hasattr(win, u"hide_for_colocar"):
            win.hide_for_colocar()
    except Exception:
        pass
    # Devolver foco a Revit para que procese el idle / ExternalEvent.
    try:
        uiapp = getattr(win, u"_uiapp", None) if win is not None else None
        if uiapp is not None:
            from revit_wpf_window_position import revit_main_hwnd
            from ctypes import windll

            hwnd = revit_main_hwnd(uiapp)
            if hwnd:
                windll.user32.SetForegroundWindow(int(hwnd))
    except Exception:
        pass
    evt = ensure_colocar_event()
    evt.Raise()
    return True


# Compatibilidad con imports previos del handler.
ColocarGuiasHandler = ColocarArmaduraHandler
