# -*- coding: utf-8 -*-
"""
Arainco: Exportar Láminas — PDF, DWG y listado Excel.

Versión portable: dependencias en ``<pushbutton>/scripts/``.
Ver ESTRUCTURA_PORTABLE.txt para despliegue.
"""

__title__ = "Exportar\nLáminas"
__author__ = "BIMTools"
__doc__ = (
    "Selecciona y exporta láminas (PDF/DWG). Nombre Personalizado: encabezado «Nombre de archivo». "
    "Opcional: listado Excel de las seleccionadas (plantilla TemplateListado en la carpeta del botón). "
    "Ruta de entrega completa y editable; «Examinar…» la completa con YYYY.MM.DD_ENTREGA (hoy). "
    "Subcarpetas PDF y DWG."
)

import os
import sys
import imp

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_DIALOG_TITLE = u"Arainco: Exportar Láminas"
_MAIN_MODULE = u"run.py"
_MAIN_MODULE_ID = u"export_laminas_run_04"

# Módulos mínimos del paquete portable (arranque + acceso + UI).
_REQUIRED_MODULES = (
    _MAIN_MODULE,
    u"bootstrap.py",
    u"corporate_access.py",
    u"bimtools_script_guard.py",
    u"bimtools_instruction_dialog.py",
    u"bimtools_ui_tokens.py",
    u"bimtools_wpf_shell.py",
    u"bimtools_wpf_dark_theme.py",
    u"revit_wpf_window_position.py",
    u"lib/sheet_export_manager.py",
    u"lib/exportar_laminas_pdf_dwg.py",
    u"mvvm/export_laminas_vm.py",
    u"ui/export_laminas_view.py",
)


def _as_unicode(value):
    if value is None:
        return u""
    try:
        return unicode(value)
    except NameError:
        return str(value)
    except Exception:
        try:
            return str(value)
        except Exception:
            return u""


def _show_error(message):
    text = _as_unicode(message).strip() or u"Error desconocido al iniciar la herramienta."
    try:
        TaskDialog.Show(_DIALOG_TITLE, text)
    except Exception:
        try:
            from pyrevit import forms

            forms.alert(text, title=_DIALOG_TITLE)
        except Exception:
            print(text)


_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
_scripts_dir = os.path.abspath(os.path.join(_pushbutton_dir, u"scripts"))

_missing = []
for _name in _REQUIRED_MODULES:
    if not os.path.isfile(os.path.join(_scripts_dir, _name.replace(u"/", os.sep))):
        _missing.append(_name)

if _missing:
    _show_error(
        u"Paquete portable incompleto. Faltan en scripts/:\n\n- {0}".format(
            u"\n- ".join(_missing)
        )
    )
else:
    # Bootstrap de acceso vive junto a script.py (paquete portable).
    if _pushbutton_dir not in sys.path:
        sys.path.insert(0, _pushbutton_dir)

    if _scripts_dir not in sys.path:
        sys.path.insert(0, _scripts_dir)

    try:
        from bootstrap import purge_export_laminas_modules, setup_export_laminas_paths

        setup_export_laminas_paths()
        purge_export_laminas_modules()
    except Exception as ex:
        _show_error(u"Error en bootstrap:\n\n{0}".format(_as_unicode(ex)))
    else:
        try:
            import bimtools_access_bootstrap as _bimtools_access
        except Exception as ex:
            _show_error(
                u"No se pudo cargar bimtools_access_bootstrap.py:\n\n{0}".format(
                    _as_unicode(ex)
                )
            )
        else:
            if _bimtools_access.require_tool_access(__file__, __revit__, __title__):
                try:
                    setup_export_laminas_paths()
                    purge_export_laminas_modules()
                    _module_path = os.path.join(_scripts_dir, _MAIN_MODULE)
                    _mod = imp.load_source(_MAIN_MODULE_ID, _module_path)
                    _mod.main(__revit__)
                except Exception as ex:
                    _show_error(
                        u"Error al ejecutar la herramienta:\n\n{0}".format(
                            _as_unicode(ex)
                        )
                    )
