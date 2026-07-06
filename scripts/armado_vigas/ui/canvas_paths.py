# -*- coding: utf-8 -*-
"""Ruta al mockup HTML del canvas (portable + extensión)."""

import os

_MOCKUP_NAME = "armado_vigas_canvas_mockup.html"


def find_canvas_mockup_path(pushbutton_dir=None):
    """
    1) ``<pushbutton>/mockups/armado_vigas_canvas_mockup.html``
    2) ``BIMTools.tab/Armadura.panel/mockups/`` bajo la raíz de extensión
    3) Subir desde pushbutton hasta encontrar el archivo
    """
    candidates = []
    if pushbutton_dir:
        candidates.append(
            os.path.join(pushbutton_dir, "mockups", _MOCKUP_NAME)
        )
    try:
        from armado_vigas.bootstrap_paths import find_scripts_dir

        scripts_dir = find_scripts_dir(pushbutton_dir)
        if scripts_dir:
            ext_root = os.path.dirname(os.path.dirname(scripts_dir))
            candidates.append(
                os.path.join(
                    ext_root,
                    "BIMTools.tab",
                    "Armadura.panel",
                    "mockups",
                    _MOCKUP_NAME,
                )
            )
    except Exception:
        pass

    if pushbutton_dir:
        cursor = pushbutton_dir
        for _ in range(14):
            candidates.append(
                os.path.join(
                    cursor,
                    "BIMTools.tab",
                    "Armadura.panel",
                    "mockups",
                    _MOCKUP_NAME,
                )
            )
            candidates.append(os.path.join(cursor, "mockups", _MOCKUP_NAME))
            parent = os.path.dirname(cursor)
            if parent == cursor:
                break
            cursor = parent

    seen = set()
    for p in candidates:
        if not p:
            continue
        ap = os.path.normpath(os.path.abspath(p))
        if ap in seen:
            continue
        seen.add(ap)
        if os.path.isfile(ap):
            return ap
    return None


def mockup_file_path(pushbutton_dir=None):
    """Ruta absoluta al HTML del canvas."""
    path = find_canvas_mockup_path(pushbutton_dir)
    if not path:
        return None
    return os.path.normpath(os.path.abspath(path))


def mockup_uri(pushbutton_dir=None):
    """Compatibilidad: devuelve ruta local absoluta (no URI con query)."""
    return mockup_file_path(pushbutton_dir)
