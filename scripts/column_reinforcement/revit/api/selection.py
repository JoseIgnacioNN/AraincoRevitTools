# -*- coding: utf-8 -*-
"""Adapter de selección Revit.

La lógica concreta permanece en el motor legado durante la fase 1; este archivo
marca el contrato para mover filtros y `PickObjects` sin contaminar servicios.
"""


class SelectionAdapter(object):
    def __init__(self, uidoc):
        self.uidoc = uidoc

    def pick_objects(self, object_type, selection_filter, prompt):
        return self.uidoc.Selection.PickObjects(object_type, selection_filter, prompt)
