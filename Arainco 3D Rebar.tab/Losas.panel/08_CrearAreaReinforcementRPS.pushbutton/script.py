# -*- coding: utf-8 -*-
"""
Crear Area Reinforcement RPS — Usa la interfaz gráfica de Area Reinforcement Losa.
"""

__title__ = "Crear Area\nReinf. RPS"
__author__ = "BIMTools"
__doc__ = "Abre la interfaz de Area Reinforcement Losa para crear mallas en losas."

import os
import sys

# Cargar el módulo compartido desde `extension/scripts` para evitar duplicación
# (esta carpeta del pushbutton contiene una copia local, pero el botón debe usar
# el módulo centralizado).
_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
#
# Estructura (desde este script):
#   <ext_root>/BIMTools.tab/Armadura.panel/08_CrearAreaReinforcementRPS.pushbutton/script.py
# Entonces `scripts/` queda en:
#   <ext_root>/scripts/
#
_ext_root = os.path.dirname(os.path.dirname(os.path.dirname(_pushbutton_dir)))
_shared_scripts_dir = os.path.join(_ext_root, "scripts")

# IronPython mantiene módulos en sys.modules: sin esto, cambios (p. ej. ganchos espesor−60) no se aplican hasta reiniciar Revit.
try:
    if "area_reinforcement_losa" in sys.modules:
        del sys.modules["area_reinforcement_losa"]
except Exception:
    pass

# Preferir módulo compartido, pero mantener fallback a la copia local
# (para poder migrar solo la carpeta del pushbutton).
_added_shared = False
if os.path.isdir(_shared_scripts_dir) and _shared_scripts_dir not in sys.path:
    sys.path.insert(0, _shared_scripts_dir)
    _added_shared = True

try:
    import bimtools_paths

    bimtools_paths.set_pushbutton_dir(_pushbutton_dir)
    from area_reinforcement_losa import run  # shared module
except ImportError:
    # Limpiar posible entrada parcial y hacer fallback a la copia local.
    try:
        if "area_reinforcement_losa" in sys.modules:
            del sys.modules["area_reinforcement_losa"]
    except Exception:
        pass

    if _pushbutton_dir not in sys.path:
        sys.path.insert(0, _pushbutton_dir)
    if os.path.isdir(_shared_scripts_dir) and _shared_scripts_dir not in sys.path:
        sys.path.insert(0, _shared_scripts_dir)
    import bimtools_paths

    bimtools_paths.set_pushbutton_dir(_pushbutton_dir)
    from area_reinforcement_losa import run  # local module
finally:
    # Si no logramos importar desde shared (o ya no es necesario), no dejamos
    # rutas “basura” en sys.path.
    # (No es destructivo: reinsertar no es necesario en PyRevit/IronPython.)
    if _added_shared:
        try:
            if _shared_scripts_dir in sys.path:
                sys.path.remove(_shared_scripts_dir)
        except Exception:
            pass

run(__revit__, close_on_finish=True)
