# -*- coding: utf-8 -*-
"""
Ejecución al cargar la extensión BIMTools (pyRevit).

Aquí se registran los DMU que reaccionan a cambios en Rebar. pyRevit busca
`startup.py` en la raíz del directorio `*.extension`, al mismo nivel que las
carpetas `*.tab`.

- DMU etiquetas / RebarShape: `ENABLE_REBAR_SHAPE_TAG_AUTO_SYNC` y
  `rebar_tag_shape_sync_core.REBAR_TAG_SYNC_DEFAULT_FAMILY_NAMES`.
- DMU ``Armadura_Largo Total`` (desactivado): el parámetro lo rellenan las
  herramientas al crear barras; ver `_apply_armadura_largo_total_to_rebars`.
- DMU anotaciones (empalme + cotas empotramiento): `ENABLE_LAP_DETAIL_LINK_DMU` y
  `scripts/lap_detail_updater_dmu.py`. Empalmes **vigas** usan schema aparte
  (`lap_detail_link_vigas_schema.py`) y geometría opcional `compute_lap_segment_endpoints_vigas`.
- DMU marcadores de cota confinamiento (columnas): `ENABLE_CONFINEMENT_DIM_LINK_DMU` y
  `scripts/confinement_dim_updater_dmu.py`.
"""

from __future__ import print_function

import os
import sys

# Interruptor global para habilitar/deshabilitar el DMU de sincronización
# automática de Rebar Tag por Shape.
ENABLE_REBAR_SHAPE_TAG_AUTO_SYNC = False

# DMU desactivado: ``Armadura_Largo Total`` solo se escribe desde herramientas (no al editar Rebar).
ENABLE_ARMADURA_LARGO_TOTAL_DMU = False

# Reposicionar / limpiar Detail Components de empalme ligados a pares de Rebar.
ENABLE_LAP_DETAIL_LINK_DMU = True

# Borrar DetailCurve marcadores al eliminar cotas de confinamiento (columnas).
ENABLE_CONFINEMENT_DIM_LINK_DMU = True


def _register():
    ext_root = os.path.abspath(os.path.dirname(__file__))
    scripts_dir = os.path.join(ext_root, "scripts")
    if scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)

    from pyrevit import HOST_APP

    addin_id = HOST_APP.addin_id

    if ENABLE_REBAR_SHAPE_TAG_AUTO_SYNC:
        from rebar_shape_tag_updater_dmu import register_rebar_shape_tag_updater

        register_rebar_shape_tag_updater(addin_id, doc=None)

    if ENABLE_ARMADURA_LARGO_TOTAL_DMU:
        from armadura_largo_total_updater_dmu import register_armadura_largo_total_updater

        register_armadura_largo_total_updater(addin_id, doc=None)
    else:
        from armadura_largo_total_updater_dmu import unregister_armadura_largo_total_updater

        try:
            unregister_armadura_largo_total_updater(addin_id)
        except Exception:
            pass

    if ENABLE_LAP_DETAIL_LINK_DMU:
        from lap_detail_updater_dmu import register_lap_detail_link_updater

        register_lap_detail_link_updater(addin_id, doc=None)

    if ENABLE_CONFINEMENT_DIM_LINK_DMU:
        from confinement_dim_updater_dmu import register_confinement_dim_link_updater

        register_confinement_dim_link_updater(addin_id, doc=None)


try:
    _register()
except Exception:
    import traceback

    traceback.print_exc()
