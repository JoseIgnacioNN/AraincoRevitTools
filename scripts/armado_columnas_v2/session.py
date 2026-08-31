# -*- coding: utf-8 -*-
"""Sesión de UI Armado columnas V2."""

from __future__ import division

_KIND_LABEL_ES = {
    u"column": u"columna",
    u"foundation": u"fundación",
    u"beam": u"viga",
    u"floor": u"losa",
}


class ColumnArmadoSession(object):
    def __init__(self):
        self.reset()

    def reset(self):
        self.domain_members = []
        self.domain_columns = []
        self.domain_foundations = []
        self.domain_beams = []
        self.domain_floors = []
        self.selected_ids = set()
        self.preview_id = None
        self.last_message = u""
        self.is_demo = False
        self.view_name = u""
        self.all_element_ids = []

    def set_selection(self, document, refs_or_elements, view=None):
        from armado_columnas_v2.revit.adapters import (
            domain_members_from_selection,
            elements_from_refs,
        )

        self.reset()
        elems = elements_from_refs(document, refs_or_elements)
        members = domain_members_from_selection(document, elems, view)
        self.domain_members = members
        self.domain_columns = [m for m in members if m.get("kind") == u"column"]
        self.domain_foundations = [m for m in members if m.get("kind") == u"foundation"]
        self.domain_beams = [m for m in members if m.get("kind") == u"beam"]
        self.domain_floors = [m for m in members if m.get("kind") == u"floor"]
        self.all_element_ids = [
            m.get("elementIdInt") for m in members if m.get("elementIdInt") is not None
        ]
        try:
            self.view_name = view.Name if view is not None else u""
        except Exception:
            self.view_name = u""

        if members:
            first = None
            for m in members:
                if m.get("kind") == u"column":
                    first = m
                    break
            if first is not None:
                self.selected_ids = set([first["id"]])
                self.preview_id = first["id"]
            else:
                # Sin columnas: contexto visible, sin selección en canvas
                self.selected_ids = set()
                self.preview_id = None

        n_c = len(self.domain_columns)
        n_f = len(self.domain_foundations)
        n_v = len(self.domain_beams)
        n_l = len(self.domain_floors)
        vista = self.view_name or u"vista activa"
        self.last_message = (
            u"{0} col · {1} fund · {2} viga(s) · {3} losa(s) · elevación «{4}»".format(
                n_c, n_f, n_v, n_l, vista
            )
        )
        self.is_demo = False
        return members

    def selected_members(self):
        ids = self.selected_ids or set()
        return [m for m in (self.domain_members or []) if m.get("id") in ids]

    def selected_columns(self):
        return [m for m in self.selected_members() if m.get("kind") == u"column"]

    def preview_member(self):
        pid = self.preview_id
        if pid:
            for m in self.domain_members or []:
                if m.get("id") == pid:
                    return m
        selected = self.selected_members()
        if selected:
            return selected[0]
        members = self.domain_members or []
        return members[0] if members else None

    def preview_column(self):
        return self.preview_member()

    def select_column(self, member_id, multi=False):
        """Selecciona solo columnas en el canvas (otros kinds no son seleccionables)."""
        if not member_id:
            return
        target = None
        for m in self.domain_members or []:
            if m.get("id") == member_id:
                target = m
                break
        if target is None or target.get("kind") != u"column":
            return
        if multi:
            if member_id in self.selected_ids:
                if len(self.selected_ids) > 1:
                    self.selected_ids.discard(member_id)
            else:
                self.selected_ids.add(member_id)
        else:
            self.selected_ids = set([member_id])
        self.preview_id = member_id

    @staticmethod
    def kind_label_es(kind):
        return _KIND_LABEL_ES.get(kind or u"", kind or u"elemento")


SESSION = ColumnArmadoSession()
