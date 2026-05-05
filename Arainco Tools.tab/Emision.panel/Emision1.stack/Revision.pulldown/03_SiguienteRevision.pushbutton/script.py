# -*- coding: utf-8 -*-
"""
Revisiones — botón autocontenido.

Toda la lógica vive en esta carpeta (subcarpeta ``siguiente_revision/`` y módulos
auxiliares junto a ``script.py``). Puede copiarse a otra extensión sin depender
de ``scripts/`` de BIMTools.

Logos: coloque ``logo.png``, ``empresa_logo.png`` o ``logo_empresa.png`` aquí;
``icon.png`` sirve como icono de la cinta y como último recurso para la cabecera WPF.
"""

from __future__ import print_function

__title__ = u"Arainco: Revisiones"
__author__ = "BIMTools"
__doc__ = (
    "Crea una nueva entrada de revisión por lámina a partir del último correlativo "
    "en el índice de la lámina y actualiza datos de revisión/nubes configurados."
)

import os
import sys

_pb = os.path.dirname(os.path.abspath(__file__))

# Primero el pushbutton: resuelve siguient_revision y módulos empaquetados
# antes que cualquier otro ``scripts/`` de la extensión.
if _pb not in sys.path:
    sys.path.insert(0, _pb)

_pkg_init = os.path.join(_pb, "siguiente_revision", "__init__.py")
if not os.path.isfile(_pkg_init):
    try:
        from pyrevit import forms

        forms.alert(
            u"Falta la carpeta siguiente_revision junto a este script.",
            title=u"Revisiones",
        )
    except Exception:
        pass
    sys.exit(1)

try:
    import bimtools_paths  # noqa: E402

    bimtools_paths.set_pushbutton_dir(_pb)
except Exception:
    pass

import siguiente_revision  # noqa: E402

reload(siguiente_revision)  # noqa: E402

siguiente_revision.main(__revit__)
