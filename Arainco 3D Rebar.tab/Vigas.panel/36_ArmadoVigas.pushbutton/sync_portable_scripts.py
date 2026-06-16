# -*- coding: utf-8 -*-
"""
Sincroniza ``scripts/`` empaquetado del pushbutton desde fuentes canónicas.

Uso:
    python sync_portable_scripts.py

Copia ``armado_vigas/`` y módulos compartidos en subcarpetas según ``portable_layout.py``.
No sobrescribe infra fija (``bootstrap.py``, ``pin_local_scripts.py``, ``bootstrap_paths.py``).
"""

from __future__ import print_function

import os
import shutil
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

from portable_layout import SHARED_MODULE_BUCKETS, STATIC_INFRA_FILES, all_shared_filenames

_REPO_ROOT = os.path.abspath(os.path.join(_HERE, os.pardir, os.pardir, os.pardir))
_CANONICAL_SCRIPTS = os.path.join(_REPO_ROOT, "scripts")
_DEST_SCRIPTS = os.path.join(_HERE, "scripts")

_TAGS_SRC = os.path.join(
    _REPO_ROOT,
    "BIMTools.tab",
    "Armadura.panel",
    "34_ArmadoMuros.pushbutton",
    "scripts",
    "armado_muros_cabezal_tags.py",
)


def _copy_file(src, dest):
    if not os.path.isfile(src):
        raise IOError(u"No existe: {0}".format(src))
    dest_dir = os.path.dirname(dest)
    if not os.path.isdir(dest_dir):
        os.makedirs(dest_dir)
    shutil.copy2(src, dest)
    print(u"  + {0}".format(os.path.relpath(dest, _HERE)))


def _copy_tree(src, dest):
    if not os.path.isdir(src):
        raise IOError(u"No existe carpeta: {0}".format(src))
    if os.path.isdir(dest):
        shutil.rmtree(dest)
    shutil.copytree(src, dest)
    n = sum(len(files) for _root, _dirs, files in os.walk(dest))
    print(u"  + {0}/ ({1} archivos)".format(os.path.relpath(dest, _HERE), n))


def _resolve_src(filename):
    if filename == u"armado_muros_cabezal_tags.py":
        return _TAGS_SRC
    return os.path.join(_CANONICAL_SCRIPTS, filename)


def _remove_stale_flat_copies():
    """Elimina copias planas obsoletas en la raíz de scripts/."""
    for name in all_shared_filenames():
        stale = os.path.join(_DEST_SCRIPTS, name)
        if os.path.isfile(stale):
            os.remove(stale)
            print(u"  - {0}".format(os.path.relpath(stale, _HERE)))


def sync():
    if not os.path.isdir(_CANONICAL_SCRIPTS):
        print(u"ERROR: no se encontró scripts/ en {0}".format(_CANONICAL_SCRIPTS))
        return 1

    print(u"Sincronizando Armado vigas portable -> {0}".format(_DEST_SCRIPTS))
    if not os.path.isdir(_DEST_SCRIPTS):
        os.makedirs(_DEST_SCRIPTS)

    pkg_src = os.path.join(_CANONICAL_SCRIPTS, "armado_vigas")
    pkg_dest = os.path.join(_DEST_SCRIPTS, "armado_vigas")
    _copy_tree(pkg_src, pkg_dest)

    for bucket, files in SHARED_MODULE_BUCKETS:
        for name in files:
            _copy_file(_resolve_src(name), os.path.join(_DEST_SCRIPTS, bucket, name))

    _remove_stale_flat_copies()
    print(u"Listo.")
    return 0


if __name__ == "__main__":
    sys.exit(sync())
