# -*- coding: utf-8 -*-
"""Estado compartido entre handlers ExternalEvent (singleton de sesión)."""

from armado_vigas.domain.constants import (
    CONCRETE_GRADE_DEFAULT,
    normalize_concrete_grade,
)
from armado_vigas.domain.bar_ends import (
    BAR_END_MODE_DEFAULT,
    ensure_session_bar_end_modes,
    normalize_bar_end_mode,
)
from armado_vigas.domain.laterales import (
    LATERALES_DIAM_DEFAULT,
    suggest_n_laterales_from_beams,
)
from armado_vigas.domain.tramos import build_session_tramos, sort_beams


class ArmadoVigasSession(object):
    def __init__(self):
        self.concreteGrade = CONCRETE_GRADE_DEFAULT
        self.barEndStartSup = BAR_END_MODE_DEFAULT
        self.barEndEndSup = BAR_END_MODE_DEFAULT
        self.barEndStartInf = BAR_END_MODE_DEFAULT
        self.barEndEndInf = BAR_END_MODE_DEFAULT
        self.reset()

    def reset(self):
        # Preferencias de herramienta: conservar entre lotes.
        prev_grade = getattr(self, "concreteGrade", CONCRETE_GRADE_DEFAULT)
        prev_ends = {
            u"barEndStartSup": getattr(self, u"barEndStartSup", BAR_END_MODE_DEFAULT),
            u"barEndEndSup": getattr(self, u"barEndEndSup", BAR_END_MODE_DEFAULT),
            u"barEndStartInf": getattr(self, u"barEndStartInf", BAR_END_MODE_DEFAULT),
            u"barEndEndInf": getattr(self, u"barEndEndInf", BAR_END_MODE_DEFAULT),
        }
        self.all_element_ids = []
        self.framing_elements = []
        self.domain_beams = []
        self.domain_beams_by_element_id = {}
        self.empalme_beam_ids_sup = set()
        self.empalme_beam_ids_inf = set()
        self.split_empalme = True
        self.apoyos_loaded = False
        self.apoyos = []
        # Suple SUP por apoyo (ids columna/muro; sin losas). Ver domain.suple_superior.
        self.suple_sup_apoyo_ids = set()
        self.suple_sup_cfg_by_apoyo = {}  # id → {n, diam}
        self.selected_suple_apoyo_id = None
        self.tramos_sup = []
        self.tramos_inf = []
        self.tramos = []
        # ø / n por tramo Tn (topología) · ver domain.tramo_armado
        self.tramo_armado = {u"sup": {}, u"inf": {}}
        self.last_message = u""
        self.direction_overlay_ids = []
        self.direction_overlay_view_id = None
        self.lateralesEnabled = True
        # Qué colocar al pulsar «Colocar» (se sincroniza desde toggles del rail).
        self.placeSup = True
        self.placeInf = True
        self.placeConf = True
        self.nLaterales = 1
        self.diamLaterales = LATERALES_DIAM_DEFAULT
        self.bar_diameters_mm = None
        self.concreteGrade = normalize_concrete_grade(prev_grade)
        for k, v in prev_ends.items():
            setattr(self, k, normalize_bar_end_mode(v))
        ensure_session_bar_end_modes(self)
        # Vigas hormigón unidas a la selección (paralelas / no paralelas a la vista).
        self.joined_framing = {
            "all": [],
            "parallel": [],
            "not_parallel": [],
            "by_element_id": {},
            "counts": {"all": 0, "parallel": 0, "not_parallel": 0},
        }

    def set_concrete_grade(self, grade):
        self.concreteGrade = normalize_concrete_grade(grade)
        return self.concreteGrade

    def set_selection(self, document, refs_or_elements, view=None):
        from armado_vigas.revit.adapters import elements_from_refs, framing_from_elements
        from armado_vigas.revit.adapters import domain_beams_from_framing, apoyos_from_elements
        from armado_vigas.revit.view_order import assign_beam_view_order, assign_beam_col_endpoints
        from geometria_empotramiento_extremos import element_ids_desde_elementos
        from armado_vigas.revit.rebar_resources import list_bar_diameters_mm

        elems = elements_from_refs(document, refs_or_elements)
        self.bar_diameters_mm = list_bar_diameters_mm(document)
        self.all_element_ids = element_ids_desde_elementos(elems)
        self.framing_elements = framing_from_elements(elems)
        apoyos = apoyos_from_elements(elems, document, view)
        self.apoyos = apoyos
        self.apoyos_loaded = bool(apoyos) and bool(self.framing_elements)
        self.domain_beams = domain_beams_from_framing(
            document, self.framing_elements, apoyos, view=view
        )
        assign_beam_view_order(self.domain_beams, view)
        # Re-enriquecer solid AABB + apoyos (coherente con u de LocationCurve)
        try:
            from armado_vigas.revit.elev_geometry import (
                assign_beam_supports_by_proximity,
                enrich_session_elev_geometry,
            )

            enrich_session_elev_geometry(
                self.domain_beams, self.apoyos, view, document=document
            )
            assign_beam_supports_by_proximity(self.domain_beams, self.apoyos, view)
        except Exception:
            assign_beam_col_endpoints(self.domain_beams, self.apoyos, view)
        # Flags de suple SUP derivados de apoyos activos (tras colStart/colEnd).
        try:
            from armado_vigas.domain.suple_superior import (
                ensure_session_suple_sup,
                sync_beams_suple_from_apoyo_set,
            )

            ensure_session_suple_sup(self)
            # Quitar de la selección ids de losa o apoyos ya no presentes.
            valid = set()
            for ap in self.apoyos or []:
                try:
                    from armado_vigas.domain.suple_superior import apoyo_allows_suple_sup

                    if ap and apoyo_allows_suple_sup(ap) and ap.get("id"):
                        try:
                            valid.add(unicode(ap.get("id")))
                        except NameError:
                            valid.add(str(ap.get("id")))
                except Exception:
                    try:
                        if ap and ap.get("id"):
                            try:
                                valid.add(unicode(ap.get("id")))
                            except NameError:
                                valid.add(str(ap.get("id")))
                    except Exception:
                        pass
            try:
                self.suple_sup_apoyo_ids = set(
                    a for a in (self.suple_sup_apoyo_ids or set()) if a in valid
                )
            except Exception:
                self.suple_sup_apoyo_ids = set()
            sync_beams_suple_from_apoyo_set(self, self.domain_beams)
        except Exception:
            pass
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

        # Vigas unidas a la selección: clasificar // y no-// al plano de vista.
        try:
            from armado_vigas.revit.joined_framing import (
                detect_joined_concrete_framing,
                format_joined_summary,
                not_parallel_joined_labels,
            )

            self.joined_framing = detect_joined_concrete_framing(
                document, self.framing_elements, view,
            )
            try:
                from armado_vigas.revit.elev_geometry import enrich_joined_framing_geometry

                enrich_joined_framing_geometry(self.joined_framing, view, document)
            except Exception:
                pass
            join_txt = format_joined_summary(self.joined_framing)
            np_labels = not_parallel_joined_labels(self.joined_framing)
            if np_labels:
                join_txt += u" · no // vista: {0}".format(u", ".join(np_labels))
        except Exception:
            self.joined_framing = {
                "all": [],
                "parallel": [],
                "not_parallel": [],
                "by_element_id": {},
                "counts": {"all": 0, "parallel": 0, "not_parallel": 0},
            }
            join_txt = u""

        self.last_message = (
            u"{0} elem. · {1} viga(s) · sup {2} / inf {3} tramo(s){4}".format(
                len(self.all_element_ids),
                len(self.framing_elements),
                len(self.tramos_sup),
                len(self.tramos_inf),
                join_txt,
            )
        )


SESSION = ArmadoVigasSession()
