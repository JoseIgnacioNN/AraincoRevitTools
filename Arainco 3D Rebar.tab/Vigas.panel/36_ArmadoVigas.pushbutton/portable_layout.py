# -*- coding: utf-8 -*-
"""
Layout de ``scripts/`` empaquetado — fuente única para sync y bootstrap.

Cada clave es un subdirectorio bajo ``<pushbutton>/scripts/`` que se añade a
``sys.path`` para conservar imports planos (``from geometria_* import …``).
"""

from __future__ import print_function

import os

# (subcarpeta bajo scripts/, tuplas de nombres de archivo)
SHARED_MODULE_BUCKETS = (
    (u"infra", (
        u"bimtools_paths.py",
        u"bimtools_rebar_hook_lengths.py",
        u"bimtools_rebar_3d_visibility.py",
        u"bimtools_wpf_dark_theme.py",
        u"revit_wpf_window_position.py",
    )),
    (os.path.join(u"shared", u"geometria"), (
        u"geometria_colision_vigas.py",
        u"geometria_empotramiento_extremos.py",
        u"geometria_fundacion_cara_inferior.py",
        u"geometria_viga_cara_superior_detalle.py",
        u"geometria_estribos_viga.py",
        u"evaluacion_curva_puntos_obstaculos.py",
    )),
    (os.path.join(u"shared", u"rebar"), (
        u"armadura_vigas_capas.py",
        u"rebar_fundacion_cara_inferior.py",
        u"enfierrado_shaft_hashtag.py",
    )),
    (os.path.join(u"shared", u"schemas"), (
        u"embed_anchorage_link_schema.py",
        u"lap_detail_link_vigas_schema.py",
        u"barras_bordes_losa_gancho_empotramiento.py",
    )),
    (os.path.join(u"shared", u"tags"), (
        u"armado_muros_cabezal_tags.py",
    )),
)

# Infra fija del botón (no se sobrescribe en sync)
STATIC_INFRA_FILES = (
    u"bootstrap.py",
    u"pin_local_scripts.py",
    u"bootstrap_paths.py",
)

_PORTABLE_MARKER = os.path.join(u"infra", u"bimtools_paths.py")


def all_shared_filenames():
    out = []
    for _bucket, files in SHARED_MODULE_BUCKETS:
        out.extend(files)
    return tuple(out)


def portable_import_subdirs():
    return tuple(bucket for bucket, _files in SHARED_MODULE_BUCKETS)


def is_portable_layout(scripts_root):
    if not scripts_root:
        return False
    return os.path.isfile(os.path.join(scripts_root, _PORTABLE_MARKER))


def portable_import_roots(scripts_root):
    """Rutas absolutas de subcarpetas importables (orden: infra → shared/*)."""
    if not scripts_root or not is_portable_layout(scripts_root):
        return []
    roots = []
    for sub in portable_import_subdirs():
        candidate = os.path.join(scripts_root, sub)
        if os.path.isdir(candidate):
            roots.append(os.path.abspath(candidate))
    return roots
