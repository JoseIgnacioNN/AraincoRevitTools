# -*- coding: utf-8 -*-
"""Colocación de armadura en vigas — longitudinales y estribos/confinamiento."""

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI import IExternalEventHandler, TaskDialog

from bimtools_rebar_3d_visibility import (
    apply_rebar_unobscured_in_view,
    ensure_rebar_obscured_in_view,
)

from armado_vigas.revit.colocar_rebar import colocar_armadura_longitudinal
from armado_vigas.revit.colocar_lap_detail import colocar_marcadores_empalme_vigas
from armado_vigas.revit.colocar_estribos import colocar_estribos_confinamiento
from armado_vigas.revit.colocar_laterales import colocar_laterales
from armado_vigas.revit.etiquetar_confinamiento import etiquetar_confinamiento_en_vista
from armado_vigas.revit.etiquetar_laterales import etiquetar_laterales_en_vista
from armado_vigas.revit.etiquetar_longitudinales import (
    etiquetar_longitudinales_en_vista,
    realinear_longitudinales_inf_tras_confinamiento,
)
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


class ColocarArmaduraHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        win = self._window_ref()
        if win is None:
            return
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            win.set_status(u"Sin documento activo.")
            return
        doc = uidoc.Document
        if not SESSION.framing_elements:
            win.set_status(u"No hay vigas en el lote.")
            return

        view = uidoc.ActiveView
        n_lap_details = 0
        n_lap_dims = 0
        lap_res = {}
        rebars_lat = []
        conjunto_guid = None
        if iniciar_corrida_conjunto_guid is not None:
            conjunto_guid = iniciar_corrida_conjunto_guid()

        t = Transaction(doc, u"Arainco: Armado vigas")
        t.Start()
        avisos = []
        try:
            n_bars, avisos, rebars, long_by_side, lap_jobs = colocar_armadura_longitudinal(
                doc, SESSION
            )
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
            n_tags = 0
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
            n_est, avisos_est, rebars_est, conf_tag_jobs = colocar_estribos_confinamiento(
                doc, SESSION, view=view
            )
            avisos.extend(avisos_est or [])
            rebars.extend(rebars_est or [])
            n_conf_tags = 0
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
            n_lat = 0
            n_lat_tags = 0
            if getattr(SESSION, "lateralesEnabled", False):
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
                # Laterales: nunca View Unobscured (lista aparte de longitudinales/estribos).
                if rebars_lat and view is not None:
                    ensure_rebar_obscured_in_view(doc, rebars_lat, view)
            if rebars and view is not None:
                apply_rebar_unobscured_in_view(doc, rebars, view)
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
            return
        win.request_close()

    def GetName(self):
        return u"ArmadoVigasColocarRebar"


# Compatibilidad con imports previos del handler.
ColocarGuiasHandler = ColocarArmaduraHandler
