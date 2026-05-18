# -*- coding: utf-8 -*-
"""Validación de cada paso del wizard antes de avanzar.

Pura: no importa Revit API.  Devuelve listas de mensajes de error.
"""


class ValidationEngine(object):

    # ------------------------------------------------------------------ #
    #  Step 1 — Selección de columnas                                     #
    # ------------------------------------------------------------------ #

    def validate_step1(self, column_groups):
        errors = []
        if not column_groups:
            errors.append(u"Selecciona al menos una columna.")
        return errors

    # ------------------------------------------------------------------ #
    #  Step 2 — Troceo / traslapes                                        #
    # ------------------------------------------------------------------ #

    def validate_step2(self, splice_segments):
        errors = []
        if not splice_segments:
            errors.append(u"No se generaron segmentos de troceo.")
            return errors
        for seg in splice_segments:
            if seg.height_mm < 100:
                errors.append(
                    u"Segmento {0} es demasiado corto ({1:.0f} mm).".format(
                        seg.segment_id, seg.height_mm
                    )
                )
        return errors

    # ------------------------------------------------------------------ #
    #  Step 3 — Repartición de barras                                     #
    # ------------------------------------------------------------------ #

    def validate_step3(self, column_groups, distributions):
        errors = []
        for group in column_groups:
            dist = next((d for d in distributions if d.group_id == group.group_id), None)
            if dist is None:
                errors.append(
                    u"Grupo {0} no tiene distribución de barras configurada.".format(
                        group.group_id
                    )
                )
                continue
            if group.section.is_square and dist.side_a_count != dist.side_b_count:
                errors.append(
                    u"Grupo {0}: sección cuadrada requiere mismo número de barras en A y B.".format(
                        group.group_id
                    )
                )
        return errors

    # ------------------------------------------------------------------ #
    #  Step 4 — Diámetros                                                 #
    # ------------------------------------------------------------------ #

    def validate_step4(self, splice_segments):
        errors = []
        valid_diameters = {8, 10, 12, 16, 20, 22, 25, 28, 32, 36}
        for seg in splice_segments:
            if seg.diameter_mm not in valid_diameters:
                errors.append(
                    u"Segmento {0}: diámetro {1} mm no reconocido.".format(
                        seg.segment_id, seg.diameter_mm
                    )
                )
        return errors
