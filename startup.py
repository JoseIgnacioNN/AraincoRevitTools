# -*- coding: utf-8 -*-
"""
Ejecución al cargar la extensión BIMTools (pyRevit).

Aquí se registran los DMU que reaccionan a cambios en Rebar. pyRevit busca
`startup.py` en la raíz del directorio `*.extension`, al mismo nivel que las
carpetas `*.tab`.

- DMU etiquetas / RebarShape: `ENABLE_REBAR_SHAPE_TAG_AUTO_SYNC`
  (tipo homónimo en la misma familia de cada IndependentTag).
- DMU ``Armadura_Largo Total`` (desactivado): el parámetro lo rellenan las
  herramientas al crear barras; ver `_apply_armadura_largo_total_to_rebars`.
- DMU anotaciones (empalme + cotas empotramiento): `ENABLE_LAP_DETAIL_LINK_DMU` y
  `scripts/lap_detail_updater_dmu.py`. Empalmes **vigas** usan schema aparte
  (`lap_detail_link_vigas_schema.py`) y geometría opcional `compute_lap_segment_endpoints_vigas`.
- DMU marcadores de cota confinamiento (columnas): `ENABLE_CONFINEMENT_DIM_LINK_DMU` y
  `scripts/confinement_dim_updater_dmu.py`.
- DMU color rojo barras >12 m y reset al acortar (vista activa / todas las vistas):
  `ENABLE_REBAR_LARGO_EXCESO_COLOR_DMU` y
  `scripts/rebar_largo_exceso_color_updater_dmu.py`.
- DMU patas L por cambio de diámetro: `ENABLE_REBAR_PATA_L_DIAMETER_DMU` y
  `scripts/rebar_pata_l_diameter_updater_dmu.py`.
- DMU tipo Detail según «Section Filter»:
  `ENABLE_SECTION_TYPE_FROM_FILTER_DMU` y
  `scripts/section_type_from_filter_updater_dmu.py`
  (+ hooks `ID_SECTION` / `ID_OBJECTS_CALLOUT`).
- DMU nombre de vista Detail al insertar en lámina:
  `ENABLE_DETAIL_VIEW_SHEET_RENAME_DMU` y
  `scripts/detail_view_sheet_rename_updater_dmu.py`.
- DMU tipo de viewport Seccion al insertar Detail en lámina:
  `ENABLE_DETAIL_VIEW_VIEWPORT_TYPE_DMU` y
  `scripts/detail_view_viewport_type_updater_dmu.py`.
- Interceptar «Duplicate as a Dependent»: hook
  `hooks/command-before-exec[ID_CREATE_DEPENDENT_VIEW].py` +
  `scripts/dependent_view_duplicate_intercept.py`
  (`ENABLE_DEPENDENT_VIEW_DUPLICATE_INTERCEPT` solo registra binding manual de respaldo).
"""

from __future__ import print_function

import os
import sys

# Interruptor global para habilitar/deshabilitar el DMU de sincronización
# automática de Rebar Tag por Shape.
ENABLE_REBAR_SHAPE_TAG_AUTO_SYNC = True

# DMU desactivado: ``Armadura_Largo Total`` solo se escribe desde herramientas (no al editar Rebar).
ENABLE_ARMADURA_LARGO_TOTAL_DMU = False

# Reposicionar / limpiar Detail Components de empalme ligados a pares de Rebar.
ENABLE_LAP_DETAIL_LINK_DMU = True

# Borrar DetailCurve marcadores al eliminar cotas de confinamiento (columnas).
ENABLE_CONFINEMENT_DIM_LINK_DMU = True

# Color rojo automático en vista activa si Rebar > 12 m (excluye vistas 3D).
ENABLE_REBAR_LARGO_EXCESO_COLOR_DMU = True

# Ajustar largo de pata L al cambiar RebarBarType (tabla BIMTools por Ø).
ENABLE_REBAR_PATA_L_DIAMETER_DMU = True

# Al crear una vista Detail desde una Building Section, asignar el tipo Detail
# cuyo nombre contiene el «Section Filter» de la vista origen.
ENABLE_SECTION_TYPE_FROM_FILTER_DMU = True

# Al insertar una vista Detail en una lámina, View Name = Sheet Number_Detail Number.
ENABLE_DETAIL_VIEW_SHEET_RENAME_DMU = True

# Al insertar una vista Detail en una lámina, tipo de viewport = Seccion.
ENABLE_DETAIL_VIEW_VIEWPORT_TYPE_DMU = True

# Binding manual de respaldo (el camino principal es el hook en hooks/).
# Dejar en False si el hook está activo, para no duplicar suscripciones.
ENABLE_DEPENDENT_VIEW_DUPLICATE_INTERCEPT = False


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

    if ENABLE_REBAR_LARGO_EXCESO_COLOR_DMU:
        from rebar_largo_exceso_color_updater_dmu import (
            register_rebar_largo_exceso_color_updater,
        )

        register_rebar_largo_exceso_color_updater(addin_id, doc=None)
    else:
        from rebar_largo_exceso_color_updater_dmu import (
            unregister_rebar_largo_exceso_color_updater,
        )

        try:
            unregister_rebar_largo_exceso_color_updater(addin_id)
        except Exception:
            pass

    if ENABLE_REBAR_PATA_L_DIAMETER_DMU:
        from rebar_pata_l_diameter_updater_dmu import (
            register_rebar_pata_l_diameter_updater,
        )

        register_rebar_pata_l_diameter_updater(addin_id, doc=None)
    else:
        from rebar_pata_l_diameter_updater_dmu import (
            unregister_rebar_pata_l_diameter_updater,
        )

        try:
            unregister_rebar_pata_l_diameter_updater(addin_id)
        except Exception:
            pass

    if ENABLE_SECTION_TYPE_FROM_FILTER_DMU:
        from section_type_from_filter_updater_dmu import (
            register_section_type_from_filter_updater,
        )

        register_section_type_from_filter_updater(addin_id, doc=None)
    else:
        from section_type_from_filter_updater_dmu import (
            unregister_section_type_from_filter_updater,
        )

        try:
            unregister_section_type_from_filter_updater(addin_id)
        except Exception:
            pass

    if ENABLE_DETAIL_VIEW_SHEET_RENAME_DMU:
        from detail_view_sheet_rename_updater_dmu import (
            register_detail_view_sheet_rename_updater,
        )

        register_detail_view_sheet_rename_updater(addin_id, doc=None)
    else:
        from detail_view_sheet_rename_updater_dmu import (
            unregister_detail_view_sheet_rename_updater,
        )

        try:
            unregister_detail_view_sheet_rename_updater(addin_id)
        except Exception:
            pass

    if ENABLE_DETAIL_VIEW_VIEWPORT_TYPE_DMU:
        from detail_view_viewport_type_updater_dmu import (
            register_detail_view_viewport_type_updater,
        )

        register_detail_view_viewport_type_updater(addin_id, doc=None)
    else:
        from detail_view_viewport_type_updater_dmu import (
            unregister_detail_view_viewport_type_updater,
        )

        try:
            unregister_detail_view_viewport_type_updater(addin_id)
        except Exception:
            pass

    if ENABLE_DEPENDENT_VIEW_DUPLICATE_INTERCEPT:
        from dependent_view_duplicate_intercept import (
            register_dependent_view_duplicate_intercept,
        )

        uiapp = getattr(HOST_APP, "uiapp", None)
        if uiapp is not None:
            register_dependent_view_duplicate_intercept(uiapp)


try:
    _register()
except Exception:
    import traceback

    traceback.print_exc()
