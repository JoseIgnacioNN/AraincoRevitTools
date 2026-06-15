# -*- coding: utf-8 -*-
"""Bootstrap de rutas e imports para pushbutton portable 34_ArmadoMuros."""

from __future__ import print_function

import os
import sys

_MODULES_TO_PURGE = (
    "armado_muros_instruction_dialog",
    "armado_muros_portable_path",
    "bimtools_element_id",
    "bimtools_runtime",
    "bimtools_clr_collections",
    "armado_muros_lineales",
    "armado_muros_vecinos_extremos",
    "armado_muros_preview_ui",
    "armado_muros_verticales_embed_colision",
    "armado_muros_horizontales_retraida",
    "armado_muros_cabezal",
    "armado_muros_cabezal_tags",
    "armado_muros_cabezal_empalme_markers",
    "armado_muros_cabezal_encuentro_l",
    "armado_muros_coronamiento",
    "armado_muros_etiqueta_malla",
    "armado_muros_malla_rebar_tags",
    "armado_muros_rebar_params",
    "armado_muros_run",
    "armado_muros_nodo_shared",
    "armado_muros_lap_detail_shared",
    "arearein_verticales_empotramiento_rps",
    "bimtools_wpf_dark_theme",
    "bimtools_rebar_hook_lengths",
    "bimtools_rebar_3d_visibility",
    "rebar_extender_l_ganchos_135_rps",
    "arearein_exterior_h_l135_rps",
    "arearein_interior_h_l135_rps",
    "revit_wpf_window_position",
    "wall_node_boolean_section_rps",
    "bimtools_joined_geometry",
    "geometria_fundacion_cara_inferior",
    "rebar_fundacion_cara_inferior",
    "enfierrado_shaft_hashtag",
    "embed_anchorage_link_schema",
    "lap_detail_link_schema",
    "bootstrap_paths",
)


def setup_armado_muros_paths():
    try:
        from armado_muros_portable_path import ensure_pushbutton_on_path

        return ensure_pushbutton_on_path()
    except Exception:
        try:
            root = os.path.dirname(os.path.abspath(__file__))
        except NameError:
            root = os.getcwd()
        scripts = os.path.join(root, u"scripts")
        target = scripts if os.path.isdir(scripts) else root
        if target and os.path.isdir(target) and target not in sys.path:
            sys.path.insert(0, target)
        if root and root != target and root not in sys.path:
            sys.path.insert(1, root)
        return target


def purge_armado_muros_modules():
    for name in _MODULES_TO_PURGE:
        try:
            if name in sys.modules:
                del sys.modules[name]
        except Exception:
            pass
