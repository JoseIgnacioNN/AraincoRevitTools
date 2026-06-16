# -*- coding: utf-8 -*-
"""Resolución de rutas: pushbutton local primero, extensión BIMTools después."""

from __future__ import print_function

import os
import sys

_PB_DIR = None
_SCRIPTS_CANONICAL = None

_SHARED_MODULES_TO_PURGE = (
    "geometria_empotramiento_extremos",
    "geometria_colision_vigas",
    "geometria_fundacion_cara_inferior",
    "geometria_viga_cara_superior_detalle",
    "geometria_estribos_viga",
    "evaluacion_curva_puntos_obstaculos",
    "rebar_fundacion_cara_inferior",
    "armadura_vigas_capas",
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
)


def _prepend_path(path):
    if not path:
        return
    path = os.path.abspath(path)
    try:
        while path in sys.path:
            sys.path.remove(path)
    except ValueError:
        pass
    sys.path.insert(0, path)


def purge_shared_modules():
    """Evita imports obsoletos de ``sys.modules`` (p. ej. otra herramienta en la sesión)."""
    for name in _SHARED_MODULES_TO_PURGE:
        sys.modules.pop(name, None)
    for key in list(sys.modules.keys()):
        if key == "armado_vigas" or key.startswith("armado_vigas."):
            sys.modules.pop(key, None)


def pushbutton_dir():
    global _PB_DIR
    if _PB_DIR is None:
        try:
            _PB_DIR = os.environ.get("ARAINCO_ARMADO_VIGAS_PB_DIR")
        except Exception:
            _PB_DIR = None
        if not _PB_DIR:
            try:
                _PB_DIR = os.path.dirname(os.path.abspath(__file__))
                for _ in range(4):
                    if _PB_DIR and os.path.basename(_PB_DIR) == "scripts":
                        _PB_DIR = os.path.dirname(_PB_DIR)
                        break
                    _PB_DIR = os.path.dirname(_PB_DIR)
            except NameError:
                _PB_DIR = os.getcwd()
    return _PB_DIR


def set_pushbutton_dir(path):
    global _PB_DIR
    _PB_DIR = path
    try:
        os.environ["ARAINCO_ARMADO_VIGAS_PB_DIR"] = path or ""
    except Exception:
        pass


def find_scripts_dir(pushbutton_dir_path=None):
    """
    1) ``<pushbutton>/scripts/armado_vigas/__init__.py`` (portable).
    2) ``.../BIMTools.extension/scripts/`` subiendo directorios.
    """
    pb = pushbutton_dir_path or pushbutton_dir()
    if pb:
        local = os.path.join(pb, "scripts")
        if os.path.isdir(os.path.join(local, "armado_vigas")):
            return os.path.abspath(local)

    cursor = pb or os.getcwd()
    for _ in range(12):
        candidate = os.path.join(cursor, "scripts")
        if os.path.isdir(os.path.join(candidate, "armado_vigas")):
            return os.path.abspath(candidate)
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    return None


def find_armado_muros_scripts_dir(pushbutton_dir_path=None):
    """Scripts con ``armado_muros_cabezal_tags.py`` (portable, flat o 34_ArmadoMuros)."""
    pb = pushbutton_dir_path or pushbutton_dir()
    if pb:
        local_scripts = os.path.join(pb, u"scripts")
        for rel in (
            os.path.join(u"shared", u"tags", u"armado_muros_cabezal_tags.py"),
            u"armado_muros_cabezal_tags.py",
        ):
            if os.path.isfile(os.path.join(local_scripts, rel)):
                return os.path.abspath(os.path.dirname(os.path.join(local_scripts, rel)))

    ext = find_scripts_dir(pushbutton_dir_path)
    roots = []
    if ext:
        roots.append(os.path.dirname(ext))
    pb = pushbutton_dir_path or pushbutton_dir()
    if pb:
        roots.append(pb)
    cursor = pb or os.getcwd()
    for _ in range(12):
        roots.append(cursor)
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    seen = set()
    for root in roots:
        if not root:
            continue
        candidate = os.path.join(
            root,
            u"BIMTools.tab",
            u"Armadura.panel",
            u"34_ArmadoMuros.pushbutton",
            u"scripts",
        )
        candidate = os.path.abspath(candidate)
        if candidate in seen:
            continue
        seen.add(candidate)
        if os.path.isfile(os.path.join(candidate, u"armado_muros_cabezal_tags.py")):
            return candidate
    return None


def _portable_import_roots(scripts_root):
    """Subcarpetas ``infra/`` y ``shared/*`` del layout portable empaquetado."""
    if not scripts_root:
        return []
    marker = os.path.join(scripts_root, u"infra", u"bimtools_paths.py")
    if not os.path.isfile(marker):
        return []
    subdirs = (
        u"infra",
        os.path.join(u"shared", u"geometria"),
        os.path.join(u"shared", u"rebar"),
        os.path.join(u"shared", u"schemas"),
        os.path.join(u"shared", u"tags"),
    )
    roots = []
    for sub in subdirs:
        candidate = os.path.join(scripts_root, sub)
        if os.path.isdir(candidate):
            roots.append(os.path.abspath(candidate))
    return roots


def ensure_paths(pushbutton_dir_path=None):
    """Prioriza ``scripts/`` local, subcarpetas portable y canónico en ``sys.path``."""
    pb = pushbutton_dir_path or pushbutton_dir()
    local = None
    if pb:
        candidate = os.path.join(pb, "scripts")
        if os.path.isdir(candidate):
            local = os.path.abspath(candidate)
    ext = find_scripts_dir(pb)
    muros = find_armado_muros_scripts_dir(pb)
    ordered = []
    if local:
        for sub in _portable_import_roots(local):
            if sub not in ordered:
                ordered.append(sub)
        ordered.append(local)
    if ext and ext not in ordered:
        ordered.append(ext)
    if muros and muros not in ordered:
        ordered.append(muros)
    for d in reversed(ordered):
        _prepend_path(d)
    return ordered
