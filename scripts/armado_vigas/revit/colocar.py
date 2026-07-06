# -*- coding: utf-8 -*-
"""Colocación de armadura en vigas — longitudinales y estribos/confinamiento."""

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI import (
    IExternalEventHandler,
    TaskDialog,
    TaskDialogCommonButtons,
    TaskDialogResult,
)

from bimtools_rebar_3d_visibility import (
    apply_rebar_unobscured_in_view,
    ensure_rebar_obscured_in_view,
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
        aplicar_armadura_en_lamina,
        aplicar_armadura_ubicacion_longitudinales,
    )
except Exception:
    aplicar_armadura_capa_longitudinales = None
    aplicar_armadura_en_lamina = None
    aplicar_armadura_ubicacion_longitudinales = None


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


def prompt_longitudinals_over_limit(over_limit):
    """
    Aviso de barras > 12 m con la ventana aún visible.

    Returns:
        True para continuar; False para cancelar.
    """
    if not over_limit:
        return True
    td = TaskDialog(u"Arainco: Armado vigas")
    td.MainInstruction = u"Hay barras longitudinales que superan 12 m"
    detail_lines = [
        u"· {0}: ≈ {1:.0f} mm".format(v[u"label"], v[u"length_mm"])
        for v in over_limit[:8]
    ]
    if len(over_limit) > 8:
        detail_lines.append(
            u"· … y {0} guía(s) más.".format(len(over_limit) - 8)
        )
    td.MainContent = (
        u"Marque empalmes (Traslape sup/inf) en el canvas para trocear "
        u"las fibras y evitar barras mayores a 12 m.\n\n"
        + u"\n".join(detail_lines)
        + u"\n\n¿Desea colocar la armadura de todos modos?"
    )
    td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
    try:
        return int(td.Show()) == int(TaskDialogResult.Yes)
    except Exception:
        return td.Show() == TaskDialogResult.Yes


class ColocarArmaduraHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        win = self._window_ref()
        if win is None:
            try:
                TaskDialog.Show(
                    u"Arainco: Armado vigas",
                    u"No se pudo acceder a la ventana de la herramienta.\n"
                    u"Cierre y vuelva a abrir Armado vigas.",
                )
            except Exception:
                pass
            return
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            win.set_status(u"Sin documento activo.")
            return
        doc = uidoc.Document
        if not SESSION.framing_elements:
            win.set_status(u"No hay vigas en el lote.")
            return

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
            try:
                TaskDialog.Show(
                    u"Arainco: Armado vigas",
                    u"No se pudo verificar la longitud de las barras:\n\n{0}".format(msg),
                )
            except Exception:
                pass
            return

        if not prompt_longitudinals_over_limit(over_limit):
            win.set_status(u"Colocación cancelada: hay barras longitudinales > 12 m.")
            return

        _hide_colocar_window(win)

        view = uidoc.ActiveView
        n_lap_details = 0
        n_lap_dims = 0
        lap_res = {}
        rebars_lat = []
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

        with ColocarArmaduraProgress(SESSION) as progress:
            t = Transaction(doc, u"Arainco: Armado vigas")
            t.Start()
            try:
                progress.step(u"longitudinales")
                n_bars, avisos, rebars, long_by_side, lap_jobs = colocar_armadura_longitudinal(
                    doc, SESSION
                )

                progress.step(u"parámetros")
                if aplicar_armadura_ubicacion_longitudinales is not None:
                    aplicar_armadura_ubicacion_longitudinales(long_by_side)
                if aplicar_armadura_capa_longitudinales is not None:
                    aplicar_armadura_capa_longitudinales(long_by_side)

                reset_inferior_lap_dim_host_registry()
                progress.step(u"empalmes")
                if lap_jobs and view is not None:
                    lap_res = colocar_marcadores_empalme_vigas(doc, view, lap_jobs)
                    n_lap_details = int(lap_res.get(u"n_ok") or 0)
                    n_lap_dims = int(lap_res.get(u"n_dims_ok") or 0)
                    for msg in lap_res.get(u"messages") or []:
                        if msg:
                            avisos.append(msg)
                    if n_lap_details > 0:
                        avisos.append(
                            u"Detail Items de traslape: {0}.".format(n_lap_details)
                        )
                    if n_lap_dims > 0:
                        avisos.append(
                            u"Cotas de traslape: {0}.".format(n_lap_dims)
                        )

                progress.step(u"etiquetas longitudinales")
                if rebars and view is not None:
                    n_tags, avisos_tag, err_tag = etiquetar_longitudinales_en_vista(
                        doc,
                        view,
                        rebars,
                        use_transaction=False,
                        rebars_by_side=long_by_side,
                    )
                    if avisos_tag:
                        avisos.extend(avisos_tag)
                    if err_tag:
                        avisos.append(err_tag)

                progress.step(u"estribos y confinamiento")
                n_est, avisos_est, rebars_est, conf_tag_jobs = colocar_estribos_confinamiento(
                    doc, SESSION, view=view
                )
                avisos.extend(avisos_est or [])
                rebars.extend(rebars_est or [])

                progress.step(u"etiquetas confinamiento")
                if conf_tag_jobs and view is not None:
                    n_conf_tags, avisos_conf, err_conf = etiquetar_confinamiento_en_vista(
                        doc,
                        view,
                        conf_tag_jobs,
                        use_transaction=False,
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

                if getattr(SESSION, "lateralesEnabled", False):
                    progress.step(u"laterales")
                    n_lat, avisos_lat, rebars_lat, err_lat = colocar_laterales(doc, SESSION)
                    if avisos_lat:
                        avisos.extend(avisos_lat)
                    if err_lat:
                        avisos.append(err_lat)
                    if rebars_lat and view is not None:
                        n_lat_tags, avisos_lat_tag, err_lat_tag = etiquetar_laterales_en_vista(
                            doc,
                            view,
                            rebars_lat,
                            framing_elements=SESSION.framing_elements,
                            use_transaction=False,
                        )
                        if avisos_lat_tag:
                            avisos.extend(avisos_lat_tag)
                        if err_lat_tag:
                            avisos.append(err_lat_tag)
                    if rebars_lat and view is not None:
                        ensure_rebar_obscured_in_view(doc, rebars_lat, view)

                progress.step(u"finalización")
                if rebars and view is not None:
                    apply_rebar_unobscured_in_view(doc, rebars, view)
                if aplicar_armadura_en_lamina is not None:
                    aplicar_armadura_en_lamina(rebars, view, rebars_laterales=rebars_lat)
                if aplicar_conjunto_guid_elementos_creados is not None:
                    aplicar_conjunto_guid_elementos_creados(
                        doc,
                        view,
                        rebars,
                        rebars_laterales=rebars_lat,
                        lap_result=lap_res,
                        conjunto_guid=conjunto_guid,
                    )
                t.Commit()
            except Exception as ex:
                try:
                    t.RollBack()
                except Exception:
                    pass
                try:
                    msg = unicode(ex)
                except NameError:
                    msg = str(ex)
                win.set_status(u"Error: {0}".format(msg))
                _restore_colocar_window(win)
                try:
                    TaskDialog.Show(
                        u"Arainco: Armado vigas",
                        u"No se pudo colocar la armadura:\n\n{0}".format(msg),
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
                TaskDialog.Show(
                    u"Arainco: Armado vigas — trabas",
                    u"Las barras longitudinales y/o estribos se colocaron, "
                    u"pero hubo problemas con trabas de confinamiento:\n\n"
                    + u"\n".join(tie_avisos[:8]),
                )
            except Exception:
                pass
        elif n_bars <= 0 and n_est <= 0 and n_lat <= 0:
            detail = msg
            if avisos:
                detail += u"\n\n" + u"\n".join(avisos[:8])
            try:
                TaskDialog.Show(u"Arainco: Armado vigas", detail)
            except Exception:
                pass
            _restore_colocar_window(win)
            return
        win.request_close()

    def GetName(self):
        return u"ArmadoVigasColocarRebar"


# Compatibilidad con imports previos del handler.
ColocarGuiasHandler = ColocarArmaduraHandler
