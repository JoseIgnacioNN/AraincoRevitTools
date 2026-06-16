# -*- coding: utf-8 -*-
"""Bootstrap portable 36_ArmadoVigas: rutas, purge e inicio de runtime."""

from __future__ import print_function

import os
import sys

_MODULES_TO_PURGE = (
    "bootstrap",
    "pin_local_scripts",
    "bootstrap_paths",
    "portable_layout",
    "geometria_empotramiento_extremos",
    "geometria_colision_vigas",
    "geometria_fundacion_cara_inferior",
    "geometria_viga_cara_superior_detalle",
    "geometria_estribos_viga",
    "evaluacion_curva_puntos_obstaculos",
    "armadura_vigas_capas",
    "rebar_fundacion_cara_inferior",
    "enfierrado_shaft_hashtag",
    "bimtools_rebar_hook_lengths",
    "bimtools_rebar_3d_visibility",
    "bimtools_wpf_dark_theme",
    "bimtools_paths",
    "revit_wpf_window_position",
    "lap_detail_link_vigas_schema",
    "embed_anchorage_link_schema",
    "barras_bordes_losa_gancho_empotramiento",
    "armado_muros_cabezal_tags",
    "armado_vigas",
)


def pushbutton_dir():
    try:
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    except NameError:
        return os.getcwd()


def local_scripts_dir():
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except NameError:
        return os.getcwd()


def _prepend_path(path):
    if not path or not os.path.isdir(path):
        return
    path = os.path.abspath(path)
    try:
        while path in sys.path:
            sys.path.remove(path)
    except ValueError:
        pass
    sys.path.insert(0, path)


def _portable_import_roots(scripts_root):
    pb = pushbutton_dir()
    if pb and pb not in sys.path:
        sys.path.insert(0, pb)
    try:
        from portable_layout import is_portable_layout, portable_import_roots

        if is_portable_layout(scripts_root):
            return portable_import_roots(scripts_root)
    except Exception:
        pass
    return []


def find_extension_scripts_dir(from_dir=None):
    """``.../extension/scripts/armado_vigas`` (desarrollo sin copia local)."""
    cursor = from_dir or pushbutton_dir()
    for _ in range(12):
        candidate = os.path.join(cursor, "scripts")
        if os.path.isdir(os.path.join(candidate, "armado_vigas")):
            return os.path.abspath(candidate)
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    return None


def resolve_scripts_dir():
    """Prioriza ``<pushbutton>/scripts/``; si falta el paquete, extensión canónica."""
    local = local_scripts_dir()
    if os.path.isdir(os.path.join(local, "armado_vigas")):
        return local
    return find_extension_scripts_dir()


def ensure_scripts_on_path(scripts_dir=None):
    """Antepone subcarpetas portable + ``scripts/`` + raíz del pushbutton."""
    sd = scripts_dir or resolve_scripts_dir()
    pb = pushbutton_dir()
    ordered = []
    if sd:
        ordered.extend(_portable_import_roots(sd))
        ordered.append(sd)
    if pb and pb not in ordered:
        ordered.append(pb)
    for d in reversed(ordered):
        _prepend_path(d)
    try:
        import pin_local_scripts

        pin_local_scripts.pin_local_scripts_first()
        return pin_local_scripts.local_scripts_dir()
    except Exception:
        return sd


def setup_armado_vigas_paths():
    return ensure_scripts_on_path()


def purge_armado_vigas_modules():
    for name in _MODULES_TO_PURGE:
        try:
            if name in sys.modules:
                del sys.modules[name]
        except Exception:
            pass
    for key in list(sys.modules.keys()):
        if key == "armado_vigas" or key.startswith("armado_vigas."):
            try:
                del sys.modules[key]
            except Exception:
                pass


def prepare_runtime(pushbutton_dir_path):
    """Registra pushbutton, purge compartido y orden final de ``sys.path``."""
    from armado_vigas.bootstrap_paths import (
        ensure_paths,
        purge_shared_modules,
        set_pushbutton_dir,
    )

    set_pushbutton_dir(pushbutton_dir_path)
    purge_shared_modules()
    ensure_paths(pushbutton_dir_path)
