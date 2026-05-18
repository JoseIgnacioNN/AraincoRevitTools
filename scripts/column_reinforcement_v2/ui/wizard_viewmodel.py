# -*- coding: utf-8 -*-
"""ViewModel central del wizard.

Encapsula el estado completo de los 4 pasos: datos cargados desde Revit,
configuraciones del usuario y el método to_request() que produce el DTO
final para RebarCreationService.

No tiene referencias a controles WPF; el controlador de ventana es quien
lee y escribe los controles usando FindName().
"""

from column_reinforcement_v2.models.wizard_request import WizardRequest
from column_reinforcement_v2.services.column_grouping_service import ColumnGroupingService
from column_reinforcement_v2.services.splice_logic_service import SpliceLogicService
from column_reinforcement_v2.services.distribution_service import DistributionService

# Diámetros disponibles en mm (dropdown Step 4)
AVAILABLE_DIAMETERS = [8, 10, 12, 16, 20, 22, 25, 28, 32, 36]
DEFAULT_DIAMETER_MM = 20


class WizardViewModel(object):
    """Estado mutable del wizard; actualizado por el WindowController."""

    def __init__(self):
        # Step 1 — resultado de la selección Revit
        self.column_groups = []       # [ColumnGroup]
        self.total_columns = 0
        self.total_height_m = 0.0
        self.z_bottom_mm = 0.0
        self.z_top_mm = 0.0

        # Step 2 — puntos de corte habilitados por el usuario
        # Clave: cota z en mm; Valor: bool (True = empalme activo)
        self.cut_points = {}          # {float: bool}
        self.splice_segments = []     # [SpliceSegment]

        # Step 3 — distribución de barras
        self.distributions = []       # [RebarDistribution]

        # Step 4 — diámetros (ajustados sobre splice_segments directamente)
        self.available_diameters = AVAILABLE_DIAMETERS

        # Servicios (sin dependencia Revit)
        self._splice_svc = SpliceLogicService()
        self._dist_svc = DistributionService()

    # ------------------------------------------------------------------ #
    #  Step 1: carga de datos desde Revit                                 #
    # ------------------------------------------------------------------ #

    def load_from_elements(self, elements):
        """Agrupa los elementos Revit seleccionados y pobla el Step 1."""
        svc = ColumnGroupingService()
        self.column_groups = svc.group(elements)
        self._refresh_step1_stats()
        self._init_step2_defaults()
        self._init_step3_defaults()

    def _refresh_step1_stats(self):
        if not self.column_groups:
            self.total_columns = 0
            self.total_height_m = 0.0
            self.z_bottom_mm = 0.0
            self.z_top_mm = 0.0
            return
        self.total_columns = sum(g.column_count for g in self.column_groups)
        self.z_bottom_mm = min(g.z_bottom_mm for g in self.column_groups)
        self.z_top_mm = max(g.z_top_mm for g in self.column_groups)
        self.total_height_m = (self.z_top_mm - self.z_bottom_mm) / 1000.0

    # ------------------------------------------------------------------ #
    #  Step 2: empalmes / troceo                                          #
    # ------------------------------------------------------------------ #

    def _init_step2_defaults(self):
        """Determina todos los puntos de corte posibles y activa los de grupo por defecto.

        Estrategia:
        - Puntos disponibles = límites entre grupos (sección cambia) + juntas
          internas de cada grupo (columna a columna dentro de misma sección).
        - Por defecto: solo se activan los límites entre grupos distintos.
        - Si solo hay un grupo, ninguno se activa (el usuario elige).
        """
        group_boundary_cuts = set(self._splice_svc.default_cut_points(self.column_groups))

        all_potential = set(group_boundary_cuts)
        for g in self.column_groups:
            for z in getattr(g, "column_z_joints", []):
                all_potential.add(z)

        # Activar por defecto solo los límites entre grupos distintos
        self.cut_points = {z: (z in group_boundary_cuts) for z in all_potential}
        self._rebuild_segments()

    def set_cut_active(self, z_mm, active):
        self.cut_points[z_mm] = bool(active)
        self._rebuild_segments()

    def _rebuild_segments(self):
        active_cuts = [z for z, active in self.cut_points.items() if active]
        # Preservar diámetros ya configurados antes de reconstruir
        old_diameters = {s.segment_id: s.diameter_mm for s in self.splice_segments}
        self.splice_segments = self._splice_svc.generate_segments(
            self.column_groups, active_cuts
        )
        for seg in self.splice_segments:
            seg.diameter_mm = old_diameters.get(seg.segment_id, DEFAULT_DIAMETER_MM)

    @property
    def all_cut_points_sorted(self):
        """Lista de (z_mm, label, is_checkbox, is_active) ordenada de arriba a abajo.

        - is_checkbox=True  → nivel con checkbox activable por el usuario.
        - is_checkbox=False → nivel extremo (solo visual, no activable).
        """
        if not self.column_groups:
            return []

        extreme_zs = {self.z_bottom_mm, self.z_top_mm}

        # Todos los niveles a mostrar: extremos + puntos de corte disponibles
        all_zs = set(extreme_zs) | set(self.cut_points.keys())

        result = []
        for z in sorted(all_zs, reverse=True):
            label = u"N+{0:.2f}".format(z / 1000.0)
            is_checkbox = (z in self.cut_points) and (z not in extreme_zs)
            is_active   = self.cut_points.get(z, False)
            result.append((z, label, is_checkbox, is_active))
        return result

    # ------------------------------------------------------------------ #
    #  Step 3: distribución de barras                                     #
    # ------------------------------------------------------------------ #

    def _init_step3_defaults(self):
        self.distributions = self._dist_svc.create_defaults(self.column_groups)

    def distribution_for(self, group_id):
        return self._dist_svc.get_for_group(self.distributions, group_id)

    def increment_bars(self, group_id, side):
        d = self.distribution_for(group_id)
        if d is None:
            return
        if side == "A":
            self._dist_svc.increment_a(d)
        else:
            self._dist_svc.increment_b(d)

    def decrement_bars(self, group_id, side):
        d = self.distribution_for(group_id)
        if d is None:
            return
        if side == "A":
            self._dist_svc.decrement_a(d)
        else:
            self._dist_svc.decrement_b(d)

    # ------------------------------------------------------------------ #
    #  Step 4: diámetros                                                  #
    # ------------------------------------------------------------------ #

    def set_diameter(self, segment_id, diameter_mm):
        for seg in self.splice_segments:
            if seg.segment_id == segment_id:
                seg.diameter_mm = int(diameter_mm)
                break

    # ------------------------------------------------------------------ #
    #  Resumen y DTO final                                                #
    # ------------------------------------------------------------------ #

    @property
    def segment_count(self):
        return len(self.splice_segments)

    @property
    def average_lap_height_label(self):
        if not self.splice_segments:
            return u"—"
        avg_lap = sum(s.lap_length_mm() for s in self.splice_segments) / len(self.splice_segments)
        first_diam = self.splice_segments[0].diameter_mm
        return u"50 Ø{0} = {1:.0f} mm".format(first_diam, avg_lap)

    def to_request(self, cover_mm=25.0):
        return WizardRequest(
            column_groups=self.column_groups,
            splice_segments=self.splice_segments,
            distributions=self.distributions,
            cover_mm=cover_mm,
        )
