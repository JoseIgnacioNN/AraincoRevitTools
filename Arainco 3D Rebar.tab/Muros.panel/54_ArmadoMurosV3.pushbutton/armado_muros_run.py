# -*- coding: utf-8 -*-
"""Bootstrap de rutas e imports para pushbutton portable 54_ArmadoMurosV3."""

from __future__ import print_function

import os
import sys

_MODULES_TO_PURGE = (
    "armado_muros_instruction_dialog",
    "bimtools_instruction_dialog",
    "bimtools_ui_tokens",
    "bimtools_wpf_shell",
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
    "armado_muros_txn",
    "armado_muros_nodo_shared",
    "armado_muros_lap_detail_shared",
    "armado_muros_v3_elevation",
    "armado_muros_v3_troceo",
    "armado_muros_v3_segments",
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

# Fingerprint en memoria de proceso (sin AppDomain: evita fallo al importar sin clr System).
_LAST_FINGERPRINT = None
_FINGERPRINT_APPDOMAIN_KEY = u"BIMTools.ArmadoMurosV3.ScriptsFingerprint"


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
    global _LAST_FINGERPRINT
    for name in _MODULES_TO_PURGE:
        try:
            if name in sys.modules:
                del sys.modules[name]
        except Exception:
            pass
    _LAST_FINGERPRINT = None
    try:
        import clr

        clr.AddReference("System")
        import System

        System.AppDomain.CurrentDomain.SetData(_FINGERPRINT_APPDOMAIN_KEY, None)
    except Exception:
        pass


def _scripts_dir():
    try:
        root = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        root = os.getcwd()
    scripts = os.path.join(root, u"scripts")
    return scripts if os.path.isdir(scripts) else root


def _package_fingerprint():
    """Fingerprint barato: path + mtime del preview UI (módulo más grande)."""
    scripts = _scripts_dir()
    marker = os.path.join(scripts, u"armado_muros_preview_ui.py")
    try:
        # IronPython: getmtime → int; evitar format {:.3f} (ValueError).
        mtime_key = int(os.path.getmtime(marker))
    except Exception:
        mtime_key = 0
    return u"{0}|{1}".format(
        os.path.normcase(os.path.abspath(scripts)),
        mtime_key,
    )


def _module_loaded_from_this_package(mod_name):
    mod = sys.modules.get(mod_name)
    if mod is None:
        return False
    scripts = os.path.normcase(os.path.abspath(_scripts_dir()))
    try:
        fn = getattr(mod, u"__file__", None) or u""
        fn = os.path.normcase(os.path.abspath(fn))
    except Exception:
        return False
    try:
        return fn.startswith(scripts + os.sep) or fn.startswith(scripts + u"/")
    except Exception:
        return False


def _get_cached_fingerprint():
    global _LAST_FINGERPRINT
    if _LAST_FINGERPRINT is not None:
        return _LAST_FINGERPRINT
    try:
        import clr

        clr.AddReference("System")
        import System

        return System.AppDomain.CurrentDomain.GetData(_FINGERPRINT_APPDOMAIN_KEY)
    except Exception:
        return None


def _set_cached_fingerprint(fp):
    global _LAST_FINGERPRINT
    _LAST_FINGERPRINT = fp
    try:
        import clr

        clr.AddReference("System")
        import System

        System.AppDomain.CurrentDomain.SetData(_FINGERPRINT_APPDOMAIN_KEY, fp)
    except Exception:
        pass


def ensure_armado_muros_modules_fresh(force=False):
    """
    Purge solo si el paquete cambió, force=True, o los módulos cargados
    no pertenecen a este pushbutton (p. ej. tras abrir 34_ArmadoMuros).
    """
    fp = _package_fingerprint()
    if force:
        purge_armado_muros_modules()
        _set_cached_fingerprint(fp)
        return True

    cached = _get_cached_fingerprint()
    preview_ok = _module_loaded_from_this_package(u"armado_muros_preview_ui")
    lineales_ok = _module_loaded_from_this_package(u"armado_muros_lineales")
    any_loaded = (u"armado_muros_preview_ui" in sys.modules) or (
        u"armado_muros_lineales" in sys.modules
    )
    same_pkg = (not any_loaded) or (preview_ok and lineales_ok)

    if cached == fp and same_pkg:
        return False

    purge_armado_muros_modules()
    _set_cached_fingerprint(fp)
    return True
