# -*- coding: utf-8 -*-
u"""
Capa de presentación para "Armado Columnas".

Este módulo actúa como fachada portátil:

1. Define ``ColumnLayoutWizardOutcome`` (solo datos, sin WPF).
2. Expone ``show_column_layout_wizard_singleton`` que muestra el asistente
   y devuelve el outcome con todas las decisiones del usuario.

La lógica WPF detallada (rejilla, troceo, estribos) vive en los controladores
originales de scripts/column_reinforcement/ui/, importados dinámicamente.
Para que el pushbutton sea completamente auto-contenido en el futuro, copiar
el contenido de esos módulos aquí y reemplazar la carga XAML por
``load_xaml_from_ui_folder`` (ver ui/wpf_helpers.py).

REGLA DE CAPA:
- No crea elementos Revit ni abre transacciones.
- Sólo lee del documento (para poblar combos).
- Toda escritura en modelo ocurre en creators/.
"""
import os
import sys

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System import AppDomain

# ---------------------------------------------------------------------------
# Asegurar que scripts/ está en sys.path para los controladores WPF existentes
# ---------------------------------------------------------------------------

def _ensure_scripts_in_path():
    """
    Agrega scripts/ al sys.path si no está ya.
    Necesario mientras los controladores WPF vivan en scripts/.
    En una refactoring futura completa este bloque desaparece.
    """
    # __file__ es <pushbutton>/ui/controllers.py
    ui_dir  = os.path.dirname(os.path.abspath(__file__))  # pushbutton/ui
    pb_dir  = os.path.dirname(ui_dir)                     # pushbutton
    ext_dir = os.path.dirname(os.path.dirname(pb_dir))    # extension root
    # La carpeta scripts/ en el root de la extension
    scripts_dir = os.path.join(ext_dir, "scripts")
    for d in (scripts_dir, pb_dir):
        nd = os.path.normpath(d)
        if os.path.isdir(nd) and nd not in sys.path:
            sys.path.insert(0, nd)


_ensure_scripts_in_path()


# ---------------------------------------------------------------------------
# Outcome (data-object puro)
# ---------------------------------------------------------------------------

class ColumnLayoutWizardOutcome(object):
    """
    Resultado inmutable del asistente WPF.

    Atributos:
        cancelled             (bool): usuario canceló.
        already_running       (bool): ya había una instancia activa.
        section_grid_config   (dict): {(s_mm, L_mm): {bars_a, bars_b, cover, include_inner}}.
        troceo_outcome        (obj) : TroceoSchemeOutcome o None.
        stirrup_configs       (dict): {(s_mm, L_mm): StirrupSectionData}.
        stirrup_spacing_by_column_id (dict): {element_id_iv: float_mm}.
        stirrup_bar_type_by_column_id (dict): {element_id_iv: RebarBarType}.
        global_long_bar_diam_mm (float): Ø nominal longitudinal.
        concrete_grade        (str|None): 'G25'/'G35'/'G45' o None.
        has_stirrups          (bool): hay configuración de estribos.
        troceo_diams_by_label (dict): {label: float_mm} para layout troceo.
        troceo_cut_planes_a   (list): planos de corte tipo A.
    """

    def __init__(
        self,
        cancelled=False,
        already_running=False,
        section_grid_config=None,
        troceo_outcome=None,
        stirrup_configs=None,
        stirrup_spacing_by_column_id=None,
        stirrup_bar_type_by_column_id=None,
        global_long_bar_diam_mm=12.0,
        concrete_grade=None,
    ):
        self.cancelled               = bool(cancelled)
        self.already_running         = bool(already_running)
        self.section_grid_config     = section_grid_config or {}
        self.troceo_outcome          = troceo_outcome
        self.stirrup_configs         = stirrup_configs or {}
        self.stirrup_spacing_by_column_id  = stirrup_spacing_by_column_id  or {}
        self.stirrup_bar_type_by_column_id = stirrup_bar_type_by_column_id or {}
        self.global_long_bar_diam_mm = float(global_long_bar_diam_mm)
        self.concrete_grade          = concrete_grade

    # ----- Propiedades derivadas -----

    @property
    def has_stirrups(self):
        return bool(self.stirrup_configs)

    @property
    def troceo_diams_by_label(self):
        """
        ``{label: float_mm}`` desde ``troceo_outcome``.
        El outcome expone una lista de tramos; aquí se aplana por etiqueta.
        """
        out = {}
        try:
            to = self.troceo_outcome
            if to is None:
                return out
            for seg in getattr(to, "segments", []) or []:
                lbl = getattr(seg, "ubicacion", None) or getattr(seg, "label", None)
                dmm = getattr(seg, "bar_diam_mm", None) or getattr(seg, "diam_mm", None)
                if lbl is not None and dmm is not None:
                    out[u"{}".format(lbl)] = float(dmm)
        except Exception:
            pass
        return out

    @property
    def troceo_cut_planes_a(self):
        """Lista de Plane de troceo tipo A (puede estar vacía)."""
        try:
            to = self.troceo_outcome
            if to is None:
                return []
            return list(getattr(to, "cut_planes_a", []) or [])
        except Exception:
            return []


# ---------------------------------------------------------------------------
# Entry point principal
# ---------------------------------------------------------------------------

def show_column_layout_wizard_singleton(
    section_meta,
    troceo_rows,
    uiapp,
    uidoc,
    doc,
    default_bar_diam_mm=12.0,
):
    """
    Muestra el asistente WPF o enfoca la instancia ya abierta.

    Parámetros:
        section_meta  : list de (section_key_mm_tuple, title_str) por sección detectada.
        troceo_rows   : salida de core.geometry.build_troceo_scheme_rows.
        uiapp, uidoc, doc: handles del entorno Revit.
        default_bar_diam_mm: Ø nominal preseleccionado en la UI.

    Devuelve:
        ColumnLayoutWizardOutcome.
    """
    try:
        from column_reinforcement.ui.column_layout_wizard_window import (
            show_column_layout_wizard_singleton as _orig_show,
            ColumnLayoutWizardOutcome as _OrigOutcome,
        )
    except Exception as ex:
        raise ImportError(
            u"No se pudo importar el asistente WPF original.\n"
            u"Asegúrate de que scripts/column_reinforcement/ui/ existe.\n"
            u"Error: {}".format(ex)
        )

    orig_outcome = _orig_show(
        section_meta=section_meta,
        troceo_rows=troceo_rows,
        uiapp=uiapp,
        uidoc=uidoc,
        doc=doc,
        default_bar_diam_mm=float(default_bar_diam_mm),
    )

    if orig_outcome is None:
        return ColumnLayoutWizardOutcome(cancelled=True)

    try:
        already = bool(getattr(orig_outcome, "already_running", False))
        cancelled = bool(getattr(orig_outcome, "cancelled", False))
    except Exception:
        already   = False
        cancelled = True

    if already:
        return ColumnLayoutWizardOutcome(already_running=True)
    if cancelled:
        return ColumnLayoutWizardOutcome(cancelled=True)

    return ColumnLayoutWizardOutcome(
        cancelled=False,
        already_running=False,
        section_grid_config=dict(getattr(orig_outcome, "section_grid_config", {}) or {}),
        troceo_outcome=getattr(orig_outcome, "troceo_outcome", None),
        stirrup_configs=dict(getattr(orig_outcome, "stirrup_configs", {}) or {}),
        stirrup_spacing_by_column_id=dict(
            getattr(orig_outcome, "stirrup_spacing_by_column_id", {}) or {}
        ),
        stirrup_bar_type_by_column_id=dict(
            getattr(orig_outcome, "stirrup_bar_type_by_column_id", {}) or {}
        ),
        global_long_bar_diam_mm=float(
            getattr(orig_outcome, "global_long_bar_diam_mm", default_bar_diam_mm)
            if hasattr(orig_outcome, "global_long_bar_diam_mm")
            else default_bar_diam_mm
        ),
        concrete_grade=getattr(orig_outcome, "concrete_grade", None),
    )
