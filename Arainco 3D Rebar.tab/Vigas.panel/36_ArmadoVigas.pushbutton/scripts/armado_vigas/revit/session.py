# -*- coding: utf-8 -*-
"""Estado compartido entre handlers ExternalEvent (singleton de sesión)."""

from armado_vigas.domain.laterales import (
    LATERALES_DIAM_DEFAULT,
    suggest_n_laterales_from_beams,
)
from armado_vigas.domain.tramos import build_session_tramos, sort_beams


class ArmadoVigasSession(object):
    def __init__(self):
        self.reset()

    def reset(self):
        self.all_element_ids = []
        self.framing_elements = []
        self.domain_beams = []
        self.domain_beams_by_element_id = {}
        self.empalme_beam_ids_sup = set()
        self.empalme_beam_ids_inf = set()
        self.split_empalme = True
        self.apoyos_loaded = False
        self.apoyos = []
        self.tramos_sup = []
        self.tramos_inf = []
        self.tramos = []
        self.last_message = u""
        self.direction_overlay_ids = []
        self.direction_overlay_view_id = None
        self.lateralesEnabled = True
        self.nLaterales = 1
        self.diamLaterales = LATERALES_DIAM_DEFAULT

    def set_selection(self, document, refs_or_elements, view=None):
        from armado_vigas.revit.adapters import elements_from_refs, framing_from_elements
        from armado_vigas.revit.adapters import domain_beams_from_framing, apoyos_from_elements
        from armado_vigas.revit.view_order import assign_beam_view_order, assign_beam_col_endpoints
        from geometria_empotramiento_extremos import element_ids_desde_elementos

        elems = elements_from_refs(document, refs_or_elements)
        self.all_element_ids = element_ids_desde_elementos(elems)
        self.framing_elements = framing_from_elements(elems)
        apoyos = apoyos_from_elements(elems)
        self.apoyos = apoyos
        self.apoyos_loaded = bool(apoyos) and bool(self.framing_elements)
        self.domain_beams = domain_beams_from_framing(
            document, self.framing_elements, apoyos
        )
        assign_beam_view_order(self.domain_beams, view)
        assign_beam_col_endpoints(self.domain_beams, self.apoyos, view)
        self.domain_beams_by_element_id = {}
        for beam in self.domain_beams:
            eid = beam.get("elementIdInt")
            if eid is not None:
                self.domain_beams_by_element_id[eid] = beam
        sorted_beams = sort_beams(self.domain_beams)
        self.tramos_sup, self.tramos_inf = build_session_tramos(
            sorted_beams,
            empalme_beam_ids_sup=self.empalme_beam_ids_sup,
            empalme_beam_ids_inf=self.empalme_beam_ids_inf,
            split_empalme=self.split_empalme,
        )
        self.tramos = self.tramos_sup
        self.nLaterales = suggest_n_laterales_from_beams(self.domain_beams)
        self.last_message = u"{0} elem. · {1} viga(s) · sup {2} / inf {3} tramo(s)".format(
            len(self.all_element_ids),
            len(self.framing_elements),
            len(self.tramos_sup),
            len(self.tramos_inf),
        )


SESSION = ArmadoVigasSession()
