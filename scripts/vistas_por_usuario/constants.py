# -*- coding: utf-8 -*-
"""Constantes — Vistas por Usuario (migración desde Dynamo VistasPorUsuario_script.dyn)."""

from __future__ import print_function

CLASIFICACION = u"02_TRABAJO"
ZONA_DEFAULT = u"GENERAL"
DISCIPLINE_DETAIL_SECTION = 2

TRANSACTION_TITLE = u"Arainco: Vistas por usuario"

# ViewFamilyType por nombre (sustituye ElementId hardcodeados del .dyn)
VFT_NAME_CIELO = u"Structural Plan (Cielo)"
VFT_NAME_PISO = u"Structural Plan (Piso)"
VFT_NAME_FOUNDATION = u"Structural Foundation Plan"
VFT_NAME_DETAIL_FAMILY = u"Architectural Detail"
VFT_NAME_SECTION_FAMILY = u"Architectural Section"

# Plantillas semilla (nombre exacto de View Template en el proyecto, como en el .dyn).
# Se duplican a 02_TRABAJO_{abreviacion}_…
MASTER_TEMPLATE_SEEDS = {
    u"cielo": u"Architectural Reflected Ceiling Plan",
    u"piso": u"Structural Foundation Plan",
    u"detail": u"Architectural Detail",
    u"section": u"Architectural Section",
}

# Compat: alias antiguo (ya no se usa para búsqueda)
MASTER_TEMPLATE_MARKERS = MASTER_TEMPLATE_SEEDS

# Escalas del formulario Dynamo (denominador View.Scale)
VIEW_SCALE_RATIOS = (50, 75, 100, 125, 150, 175, 200)

# Modeladores: personas.json (misma fuente que Siguiente Revisión).
# Código de clasificación = abreviación con puntos (p. ej. J.N.N.).

PARAM_PLAN = (
    (u"Clasificacion", CLASIFICACION),
    (u"Subclasificacion", None),  # abreviación modelador (p. ej. J.N.N.)
    (u"Section Filter", None),  # idem
    (u"Zona", ZONA_DEFAULT),
)

PARAM_DETAIL_SECTION_TYPE = (
    (u"Clasificacion", CLASIFICACION),
    (u"Subclasificacion", None),
    (u"Zona", ZONA_DEFAULT),
    (u"Discipline", DISCIPLINE_DETAIL_SECTION),
    (u"Section Filter", None),
)

APP_DOMAIN_KEY = u"BIMTools.VistasPorUsuario.ActiveWindow"
