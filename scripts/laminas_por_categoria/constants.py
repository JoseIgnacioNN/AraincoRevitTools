# -*- coding: utf-8 -*-
"""Constantes — Láminas por categoría (migración de laminasPorCategoria_script.dyn)."""

from __future__ import print_function

TRANSACTION_TITLE = u"Arainco: Láminas por categoría"
APP_DOMAIN_KEY = u"BIMTools.LaminasPorCategoria.ActiveWindow"

SHEET_NAME = u"CONTENIDO LAMINA"
PARAM_CLASIFICACION = u"Clasificacion"
PARAM_VALIDACION = u"Validacion"
TITLE_BLOCK_SPLASH_NAME = u"EST_A_SPLASH SCREEN"
NUMBER_PAD_WIDTH = 3
PERSONAS_FILE_NAME = u"personas.json"

# Keys = texto del combo; value = código escrito en Clasificacion y en SheetNumber.
CATEGORIA_OPTIONS = (
    (u"PG", u"PG - PLANTA GENERAL"),
    (u"LO", u"LO - LOSAS"),
    (u"MA", u"MA - MUROS"),
    (u"VH", u"VH - VIGAS"),
    (u"CH", u"CH - COLUMNAS"),
    (u"EM", u"EM - ESTRUCTURA METALICA"),
    (u"ES", u"ES - ESCALERA"),
    (u"DE", u"DE - DETALLE ESTANQUE"),
    (u"RP", u"RP - RAMPAS"),
    (u"MC", u"MC - MONTACARGAS"),
    (u"PC", u"PC - PLANO DE CARGAS"),
    (u"SK", u"SK - FICHA"),
    (u"DM", u"DM - DEMOLICION"),
    (u"RF", u"RF - REFUERZO"),
    (u"CF", u"CF - CIELO FALSO"),
)

# Combo muestra el mes; en lámina se escribe «ENE. 2026» (espacio tras el punto).
MONTH_OPTIONS = (
    (u"ENERO", u"ENE."),
    (u"FEBRERO", u"FEB."),
    (u"MARZO", u"MAR."),
    (u"ABRIL", u"ABR."),
    (u"MAYO", u"MAY."),
    (u"JUNIO", u"JUN."),
    (u"JULIO", u"JUL."),
    (u"AGOSTO", u"AGO."),
    (u"SEPTIEMBRE", u"SEP."),
    (u"OCTUBRE", u"OCT."),
    (u"NOVIEMBRE", u"NOV."),
    (u"DICIEMBRE", u"DIC."),
)

# Respaldo si personas.json falta o está mal formado. Índices = lista ya ordenada.
PERSONAS_FALLBACK = {
    u"calculo": {
        u"default_index": 1,
        u"items": (
            u"N. VICUÑA",
            u"C. MANRIQUEZ",
            u"M. ECHEVERRIA",
            u"F. TORELLI",
            u"A. MARTINEZ",
            u"M. SAEZ",
            u"J. TORELLI",
            u"F. IBARRA",
        ),
    },
    u"revision": {
        u"default_index": 0,
        u"items": (
            u"N. VICUÑA",
            u"C. MANRIQUEZ",
            u"M. ECHEVERRIA",
            u"F. TORELLI",
            u"A. MARTINEZ",
            u"M. SAEZ",
            u"J. TORELLI",
            u"F. IBARRA",
        ),
    },
    u"aprobacion": {
        u"default_index": 2,
        u"items": (
            u"N. VICUÑA",
            u"C. MANRIQUEZ",
            u"P. ARANEDA",
        ),
    },
    u"dibujo": {
        u"default_index": 0,
        u"items": (
            u"J. NUÑEZ",
            u"P. CAROCA",
            u"A. SANCHEZ",
            u"S. JIMENEZ",
            u"H. CARRASCO",
            u"R. PEREZ",
            u"C. ORREGO",
            u"L. BARRIOS",
            u"R. AGUILA",
            u"J. NEIRA",
            u"C. CANCINO",
        ),
    },
}

PERSONA_ROLES = (u"calculo", u"revision", u"aprobacion", u"dibujo")
