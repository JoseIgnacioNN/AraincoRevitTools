# -*- coding: utf-8 -*-
"""
Constantes globales de la herramienta Revisiones.

Centraliza todas las magic strings y valores de configuración para eliminar
referencias dispersas a strings literales en el código.
"""

from __future__ import print_function

# --- Singleton / AppDomain ---
APP_DOMAIN_KEY = u"BIMTools.SiguienteRevision.ActiveWindow"

# --- Transacción Revit ---
TX_NAME = u"Arainco: Revisiones"

# --- ProgressBar ---
PBAR_TITLE_BASE = u"Arainco - Revisiones"
PBAR_ACCENT_RGB = (91, 192, 222)

# --- personas.json ---
ISSUES_DIR = u"Y:\\00_SERVIDOR DE INCIDENCIAS"
PERSONAS_FILE_NAME = u"personas.json"
PERSONA_ROL_MODELADOR = u"Modelador"
PERSONA_ROL_INGENIERO = u"Ingeniero"

# --- Sufijos de parámetros de cajetín (convención Rnn_XX_SUFIJO) ---
SUFFIX_NUM = u"01_NUM"
SUFFIX_DES = u"02_DES"
SUFFIX_DIR = u"03_DIR"
SUFFIX_DIB = u"03_DIB"   # familias que no usan DIR
SUFFIX_REV = u"04_REV"
SUFFIX_APR = u"05_APR"
SUFFIX_FCH = u"06_FCH"

MAX_REVISION_SLOTS = 20

# --- Parámetro en nubes de revisión (opcional) ---
PARAM_CANTIDAD_REVISIONES = u"CANTIDAD_REVISIONES"

# --- Láminas excluidas del listado ---
EXCLUDE_SHEET_NAME_SUBSTR = u"splash screen"

# --- Layouts de parámetros en cajetín ---
LAYOUT_RNN_FIELD = u"rnn_field"   # R02_01_NUM, R02_02_DES, …
LAYOUT_R01_ROW  = u"r01_row"      # R01_02_NUM, R01_02_DES, …

# --- Fecha ---
DATE_FORMAT = u"dd.MM.yy"
DATE_DAYS_BEFORE = 5
DATE_DAYS_AFTER  = 20

# --- Iniciales de respaldo (sin personas.json) ---
DIBUJO_INICIALES = (
    u"J.N.N.",
    u"P.C.C.",
    u"A.S.A.",
    u"C.M.H.",
    u"S.J.M.",
    u"R.P.C.",
    u"C.O.C.",
    u"H.C.P.",
    u"L.B.M.",
    u"J.N.O.",
    u"T.M.M.",
    u"B.L.A.",
    u"X.X.X.",
)

# --- Descripciones de revisión ---
DESCRIPCIONES = (
    u"PRELIMINAR",
    u"PARA APROBACION",
    u"EMITIDO PARA APROBACION MUNICIPAL",
    u"MODIFICACION GENERAL",
    u"APROBADO PARA LICITACION",
    u"ACTUALIZACION APL",
    u"MODIFICA LO INDICADO",
    u"APROBADO PARA CONSTRUCCION",
    u"PARA MUNICIPALIDAD",
    u"APTO PARA LICITACION",
    u"PARA LICITACION",
    u"INGRESA UNA DESCRIPCION",
    u"PARA INGRESO MUNICIPAL",
    u"MODIFICA Y AGREGA LO INDICADO",
    u"INFORMATIVO",
    u"PARA REVISION",
    u"ACTUALIZA LICITACION",
)
