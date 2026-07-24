# -*- coding: utf-8 -*-
"""Constantes — Vistas por Categoría (migración desde Dynamo VistasPorCategoria_script.dyn)."""

from __future__ import print_function

CLASIFICACION = u"01_ENTREGABLE"
ZONA_DEFAULT = u"GENERAL"
DISCIPLINE_DETAIL_SECTION = 2

TRANSACTION_TITLE = u"Arainco: Vistas por categoría"

# ViewFamilyType por nombre (sustituye ElementId hardcodeados del .dyn)
VFT_NAME_CIELO = u"Structural Plan (Cielo)"
VFT_NAME_PISO = u"Structural Plan (Piso)"
VFT_NAME_FOUNDATION = u"Structural Foundation Plan"
VFT_NAME_DETAIL_FAMILY = u"Architectural Detail"
VFT_NAME_SECTION_FAMILY = u"Architectural Section"

# Plantillas semilla (nombre exacto de View Template en el proyecto, como en el .dyn).
# Se duplican a 01_ENTREGABLE_{categoria}_…_{zona}
MASTER_TEMPLATE_SEEDS = {
    u"cielo": u"Architectural Reflected Ceiling Plan",
    u"piso": u"Structural Foundation Plan",
    u"detail": u"Architectural Detail",
    u"section": u"Architectural Section",
}

# Escalas del formulario Dynamo (denominador View.Scale)
VIEW_SCALE_RATIOS = (50, 75, 100, 125, 150, 175, 200)

# Categorías de entregable (código, etiqueta). Valores = código Dynamo (split " - "[0]).
CATEGORIA_OPTIONS = (
    (u"00_PG", u"00_PG - PLANTAS GENERALES"),
    (u"01_LO", u"01_LO - PLANTAS LOSAS"),
    (u"02_MA", u"02_MA - ELEVACION POR EJE"),
    (u"03_VH", u"03_VH - ELEVACION VIGAS"),
    (u"04_CH", u"04_CH - ELEVACION COLUMNAS"),
    (u"05_EM", u"05_EM - ESTRUCTURA METALICA"),
    (u"06_ES", u"06_ES - DETALLE ESCALERA"),
    (u"07_DE", u"07_DE - DETALLE ESTANQUE"),
    (u"08_RP", u"08_RP - DETALLE RAMPA"),
    (u"09_MC", u"09_MC - DETALLE MONTACARGA"),
    (u"10_PC", u"10_PC - PLANTAS DE CARGA"),
    (u"11_CF", u"11_CF - CIELOS FALSOS"),
    (u"12_OE", u"12_OE - OBRAS EXTERIORES"),
    (u"13_SK", u"13_SK - ESQUEMAS"),
    (u"14_RF", u"14_RF - REFUERZOS"),
    (u"15_DM", u"15_DM - DEMOLICION"),
    (u"16_EX", u"16_EX - EXCAVACION"),
)

# Parámetros: None se sustituye en runtime (categoria / zona / section_filter).
PARAM_PLAN_KEYS = (
    u"Clasificacion",
    u"Subclasificacion",
    u"Zona",
    u"Section Filter",
)

PARAM_DETAIL_SECTION_KEYS = (
    u"Clasificacion",
    u"Subclasificacion",
    u"Zona",
    u"Discipline",
    u"Section Filter",
)

APP_DOMAIN_KEY = u"BIMTools.VistasPorCategoria.ActiveWindow"
