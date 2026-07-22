# -*- coding: utf-8 -*-
"""Bootstrap de rutas e imports para pushbutton portable 04_ExportarLaminasPDFDWG."""

from __future__ import print_function

import os
import sys

# Solo módulos de esta herramienta (hot-reload). No usar startswith("lib")/("ui")
# sin punto: borraría lib2to3, uiautomation, etc. y puede dejar diálogos pyRevit en blanco.
_MODULES_TO_PURGE = (
    "run",
    "export_laminas_run_04",
    "bootstrap",
    "bootstrap_path",
    # Capas locales
    "lib",
    "mvvm",
    "ui",
    "infra",
    # Copias planas de acceso / tema (scripts/)
    "corporate_access",
    "bimtools_script_guard",
    "bimtools_instruction_dialog",
    "bimtools_ui_tokens",
    "bimtools_wpf_shell",
    "bimtools_wpf_dark_theme",
    "revit_wpf_window_position",
    # Nombres planos (versiones anteriores)
    "export_laminas_app",
    "export_laminas_run",
    "export_laminas_portable_path",
    "export_laminas_view",
    "export_laminas_vm",
    "export_laminas_services",
    "export_laminas_strategies",
    "export_laminas_commands",
    "export_laminas_instruction_dialog",
    "componer_nombre_lamina_ui",
    "exportar_laminas_pdf_dwg",
    "export_laminas_naming_schema",
    "sheet_export_manager",
    "listado_planos_excel_core",
    "bimtools_paths",
    # Nombres históricos (imp.load_source en versiones anteriores)
    "bimtools_exportar_laminas_pdf_dwg__04pushbutton",
    "bimtools_componer_nombre_lamina_ui__04pushbutton",
    "bimtools_listado_planos_excel_core__04export_laminas",
    "bimtools_exportlam_services__04",
    "bimtools_exportlam_strategies__04",
    "bimtools_exportlam_commands__04",
    "bimtools_exportlam_vm__04",
    "bimtools_exportlam_view__04",
    "bimtools_paths__ExportarLaminasPDFDWG",
    "bimtools_paths__ComponerNombreLamina",
)

# Paquetes locales bajo scripts/: solo raíz exacta o hijos con punto.
_PACKAGE_ROOTS = (
    "lib",
    "mvvm",
    "ui",
    "infra",
)


def setup_export_laminas_paths():
    """Inserta ``<pushbutton>/scripts/`` al frente de ``sys.path``."""
    try:
        scripts_dir = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        scripts_dir = os.getcwd()
    if scripts_dir and os.path.isdir(scripts_dir):
        try:
            while scripts_dir in sys.path:
                sys.path.remove(scripts_dir)
        except Exception:
            pass
        sys.path.insert(0, scripts_dir)
    return scripts_dir


def purge_export_laminas_modules():
    for name in _MODULES_TO_PURGE:
        try:
            if name in sys.modules:
                del sys.modules[name]
        except Exception:
            pass
    for key in list(sys.modules.keys()):
        for root in _PACKAGE_ROOTS:
            if key == root or key.startswith(root + "."):
                try:
                    del sys.modules[key]
                except Exception:
                    pass
                break
