# -*- coding: utf-8 -*-
from armado_vigas.geometry.extremos import (
    aplicar_extremos_a_linea_fusionada,
    mark_pata_l_keep_geometry,
)
from armado_vigas.geometry.longitudinales import (
    build_longitudinal_guides_for_chain,
    build_longitudinal_guides_for_run,
)
from armado_vigas.geometry.retract_muros_noparalelos import (
    aplicar_estiramiento_extremos_vigas_noparalelas,
    aplicar_retracto_extremos_muros_noparalelos,
)
