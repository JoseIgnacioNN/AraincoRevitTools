# -*- coding: utf-8 -*-
"""
Rutas de logo corporativo para WPF, portables entre extensiones y carpetas de pushbutton.

Prioridad (cada sección añade candidatos sin duplicar):
  1) Carpeta del pushbutton registrada con :func:`set_pushbutton_dir` (desde script.py)
  2) ``<raíz_extensión>/assets/`` y ``<raíz_extensión>/branding/`` (mismos nombres de archivo)
  3) Si existe ``BIMTools.tab``, listas heredadas por botón (compatibilidad con el layout actual)

``raíz_extensión`` se infiere de la ruta de este módulo (hijo de ``scripts/``). Cualquier
herramienta que añada ``.../ExtensionName.extension/scripts`` al path puede
``import bimtools_paths`` y recibir la raíz de esa extensión.
"""

import os

_LOGO_NAMES = ("empresa_logo.png", "logo_empresa.png", "logo.png")
_pushbutton_dir = None


def set_pushbutton_dir(path):
    """
    Registrar el directorio del .pushbutton antes de abrir el formulario (o antes de
    ``import`` si el módulo resuelve logos al importar).

    Ejemplo en script.py:
        _d = os.path.dirname(os.path.abspath(__file__))
        bimtools_paths.set_pushbutton_dir(_d)
    """
    global _pushbutton_dir
    if path and os.path.isdir(os.path.normpath(path)):
        _pushbutton_dir = os.path.normpath(os.path.abspath(path))
    else:
        _pushbutton_dir = None


def get_pushbutton_dir():
    return _pushbutton_dir


def default_extension_root():
    """Raíz de la extensión asumiendo este archivo en ``.../extension/scripts/bimtools_paths.py``."""
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def get_logo_paths(extension_root=None):
    """
    Lista ordenada de rutas candidatas a probar (existencia se comprueba al cargar la imagen).
    """
    if extension_root is None:
        extension_root = default_extension_root()
    else:
        extension_root = os.path.normpath(os.path.abspath(extension_root))

    out = []
    seen = set()

    def _add(p):
        p = os.path.normpath(os.path.abspath(p))
        if p not in seen:
            seen.add(p)
            out.append(p)

    if _pushbutton_dir:
        for name in _LOGO_NAMES:
            _add(os.path.join(_pushbutton_dir, name))
        _add(os.path.join(_pushbutton_dir, "icon.png"))
    for sub in ("assets", "branding"):
        base = os.path.join(extension_root, sub)
        for name in _LOGO_NAMES:
            _add(os.path.join(base, name))
        _add(os.path.join(base, "icon.png"))

    bimtools_tab = os.path.join(extension_root, "BIMTools.tab")
    if not os.path.isdir(bimtools_tab):
        return out

    panel = os.path.join(bimtools_tab, "Armadura.panel")
    model = os.path.join(bimtools_tab, "Modelado.panel")
    inc = os.path.join(
        bimtools_tab,
        "Incidencias.panel",
        "Incidencias.stack",
        "01_BIMIssue.pushbutton",
    )
    pushbuttons = (
        os.path.join(panel, "08_CrearAreaReinforcementRPS.pushbutton"),
        os.path.join(panel, "22_EnfierradoFundacionAislada.pushbutton"),
        os.path.join(panel, "02_NumerarFundaciones.pushbutton"),
        os.path.join(panel, "20_BordeLosaGanchoEmpotramiento.pushbutton"),
        os.path.join(panel, "23_EnfierradoVigas.pushbutton"),
        os.path.join(panel, "24_EnfierradoColumnas.pushbutton"),
        os.path.join(panel, "25_ArmaduraColumnasV2.pushbutton"),
        os.path.join(panel, "26_WallFoundationReinforcement.pushbutton"),
        os.path.join(panel, "09_CrearAreaReinforcementMuroRPS.pushbutton"),
        os.path.join(model, "01_SpotElevVerticesLosa.pushbutton"),
    )
    for pb in pushbuttons:
        for name in _LOGO_NAMES:
            _add(os.path.join(pb, name))
    _add(os.path.join(inc, "logo.png"))
    return out
