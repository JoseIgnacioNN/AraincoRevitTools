# -*- coding: utf-8 -*-
"""Servicio de creación de barras longitudinales en Revit.

Recibe un WizardRequest con toda la configuración del wizard y crea las barras
usando la Revit API.  Usa TransactionRunner para garantizar el prefijo Arainco:.
"""

import os
import sys

_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
_SCRIPTS_DIR = os.path.dirname(os.path.dirname(_THIS_DIR))
if _SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, _SCRIPTS_DIR)

from column_reinforcement.revit.api.transactions import TransactionRunner
from column_reinforcement_v2.models.wizard_request import WizardResult

FEET_TO_MM = 304.8
MM_TO_FEET = 1.0 / FEET_TO_MM


class RebarCreationService(object):
    """Crea barras longitudinales en Revit a partir de un WizardRequest.

    Arquitectura de creación:
    - Por cada ColumnGroup se obtienen los elementos Revit.
    - Por cada elemento se aplica la distribución y los segmentos de troceo.
    - Se crean curvas Line verticales para cada barra y se llama
      Rebar.CreateFromCurves() (Revit 2024+).
    """

    def __init__(self, doc, version_adapter):
        self.doc = doc
        self.version_adapter = version_adapter

    def execute(self, wizard_request):
        from Autodesk.Revit.DB import Transaction

        runner = TransactionRunner(self.doc, Transaction)
        bars_created = [0]

        def _create():
            count = self._create_all_bars(wizard_request)
            bars_created[0] = count

        try:
            runner.run(u"Armado Columnas Wizard", _create)
            return WizardResult(
                success=True,
                message=u"Se crearon {0} barras correctamente.".format(bars_created[0]),
                bars_created=bars_created[0],
            )
        except Exception as ex:
            return WizardResult(
                success=False,
                message=u"Error al crear barras: {0}".format(ex),
            )

    def _create_all_bars(self, request):
        """Itera grupos → elementos → segmentos y crea barras longitudinales."""
        from Autodesk.Revit.DB import ElementId
        from Autodesk.Revit.DB import Curve as RevitCurve
        from Autodesk.Revit.DB.Structure import (
            Rebar, RebarStyle, RebarHookOrientation,
        )
        from System.Collections.Generic import List

        errors = []
        count  = 0

        for group in request.column_groups:
            dist = request.distribution_for_group(group.group_id)
            if dist is None:
                continue

            bar_positions = self._compute_bar_positions(group, dist, request.cover_mm)
            bar_type = self._find_bar_type_for_segment(request)

            if bar_type is None:
                errors.append(
                    u"No se encontró RebarBarType en el documento. "
                    u"Carga al menos un tipo de barra antes de ejecutar."
                )
                break

            for elem_id_int in group.element_ids:
                host = self.doc.GetElement(ElementId(elem_id_int))
                if host is None:
                    continue

                for segment in request.splice_segments:
                    for pos_x_mm, pos_y_mm in bar_positions:
                        curve = self._make_vertical_line(
                            host, pos_x_mm, pos_y_mm,
                            segment.z_start_mm, segment.z_end_mm,
                        )
                        if curve is None:
                            continue

                        curves_list = List[RevitCurve]()
                        curves_list.Add(curve)

                        try:
                            Rebar.CreateFromCurves(
                                self.doc,
                                RebarStyle.Standard,
                                bar_type,
                                None,                      # startHookType
                                None,                      # endHookType
                                host,
                                self._normal_vec(),        # XYZ(0,1,0) — perpendicular al eje Z
                                curves_list,
                                RebarHookOrientation.Left,
                                RebarHookOrientation.Left,
                                True,                      # useExistingShapeIfPossible
                                True,                      # createNewShape si no existe
                            )
                            count += 1
                        except Exception as ex:
                            errors.append(u"[seg {0}] {1}".format(segment.segment_id, ex))

        if errors:
            raise Exception(u"\n".join(errors[:5]))

        return count

    def _compute_bar_positions(self, group, dist, cover_mm):
        """Devuelve lista de (x_mm, y_mm) relativas al centroide de la sección."""
        w = group.section.side_b_mm
        h = group.section.side_a_mm
        c = cover_mm

        positions = []
        na = dist.side_a_count  # barras por cara A (lado h)
        nb = dist.side_b_count  # barras por cara B (lado w)

        # Esquinas
        corners = [(-w/2+c, -h/2+c), (w/2-c, -h/2+c),
                   (w/2-c,  h/2-c), (-w/2+c,  h/2-c)]
        positions.extend(corners)

        # Barras intermedias cara inferior y superior (y = ±h/2)
        for i in range(1, nb - 1):
            x = -w/2 + c + i * (w - 2*c) / (nb - 1)
            positions.append((x, -h/2 + c))
            positions.append((x,  h/2 - c))

        # Barras intermedias cara izquierda y derecha (x = ±w/2)
        for j in range(1, na - 1):
            y = -h/2 + c + j * (h - 2*c) / (na - 1)
            positions.append((-w/2 + c, y))
            positions.append(( w/2 - c, y))

        return positions

    def _make_vertical_line(self, host, x_mm, y_mm, z_start_mm, z_end_mm):
        try:
            from Autodesk.Revit.DB import Line, XYZ, LocationPoint, LocationCurve
            loc = host.Location
            if isinstance(loc, LocationPoint):
                loc_pt = loc.Point
            elif isinstance(loc, LocationCurve):
                loc_pt = loc.Curve.GetEndPoint(0)
            else:
                return None
            x = loc_pt.X + x_mm * MM_TO_FEET
            y = loc_pt.Y + y_mm * MM_TO_FEET
            p0 = XYZ(x, y, z_start_mm * MM_TO_FEET)
            p1 = XYZ(x, y, z_end_mm * MM_TO_FEET)
            return Line.CreateBound(p0, p1)
        except Exception:
            return None

    def _normal_vec(self):
        # Normal perpendicular a la dirección de la barra (eje Z).
        # XYZ(0,1,0) satisface la restricción de Revit API: normal ⊥ curvas.
        from Autodesk.Revit.DB import XYZ
        return XYZ(0, 1, 0)

    def _find_bar_type_for_segment(self, request, segment_idx=0):
        """Busca el RebarBarType cuyo diámetro coincide con el primer segmento.

        Estrategia de búsqueda:
        1. Coincidencia exacta (±2 mm) con el diámetro del segmento.
        2. Primer tipo disponible como fallback.
        """
        from Autodesk.Revit.DB import FilteredElementCollector
        from Autodesk.Revit.DB.Structure import RebarBarType

        target_mm = 20
        if request.splice_segments and segment_idx < len(request.splice_segments):
            target_mm = request.splice_segments[segment_idx].diameter_mm

        all_types = list(FilteredElementCollector(self.doc).OfClass(RebarBarType))
        if not all_types:
            return None

        for bt in all_types:
            try:
                d_mm = bt.BarDiameter * FEET_TO_MM
                if abs(d_mm - target_mm) < 2.0:
                    return bt
            except Exception:
                pass

        return all_types[0]
