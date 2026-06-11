# -*- coding: utf-8 -*-
"""
Composition root — Exportar Láminas (MVVM).

Capas en ``scripts/``:
  run.py                         Punto de entrada lógico (singleton + grafo DI)
  ui/export_laminas_view.py      Vista WPF
  mvvm/export_laminas_vm.py      ViewModel
  mvvm/export_laminas_commands.py RelayCommand
  mvvm/export_laminas_services.py Servicios (Revit, carpetas, progreso)
  mvvm/export_laminas_strategies.py Estrategias PDF / DWG / Listado
  lib/exportar_laminas_pdf_dwg.py Dominio: DataTable, naming, FCH
  lib/sheet_export_manager.py    Motor PDF/DWG
  lib/listado_planos_excel_core.py Generación Excel
  ui/componer_nombre_lamina_ui.py Diálogo Nombre Personalizado
  ui/export_laminas_instruction_dialog.py Mensajes modales BIMTools
  lib/export_laminas_naming_schema.py Extensible Storage (receta de nombres)
"""

import os

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

import System  # noqa: E402

from infra.bimtools_paths import set_pushbutton_dir  # noqa: E402
from infra.revit_wpf_window_position import revit_main_hwnd  # noqa: E402
from lib.exportar_laminas_pdf_dwg import (  # noqa: E402
    build_sheets_datatable,
    datatable_row_matches_fecha_entrega_selection,
    evaluate_naming_recipe,
    get_persisted_naming_recipe_segments,
    list_fch_entrega_parameter_names_in_model,
    list_naming_source_options,
    persist_naming_recipe_segments,
    sanitize_file_base,
    unique_fecha_entrega_values_from_datatable,
)
from mvvm.export_laminas_commands import RelayCommand  # noqa: E402
from mvvm.export_laminas_services import (  # noqa: E402
    BloquearComandosRevit,
    FolderBrowserService,
    ProgressService,
    RevitWindowService,
)
from mvvm.export_laminas_strategies import (  # noqa: E402
    DwgExportStrategy,
    ListadoExportStrategy,
    PdfExportStrategy,
)
from mvvm.export_laminas_vm import ExportarLaminasViewModel  # noqa: E402
from ui.componer_nombre_lamina_ui import show_componer_nombre_dialog  # noqa: E402
from ui.export_laminas_instruction_dialog import show_message_dialog  # noqa: E402
from ui.export_laminas_view import ExportarLaminasView  # noqa: E402

try:
    from pyrevit import forms as _pyrevit_forms
except Exception:
    _pyrevit_forms = None

_SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
_PB = os.path.dirname(_SCRIPTS_DIR)
_TEMPLATE_LISTADO_XLSX = os.path.join(_PB, u"TemplateListado.xlsx")

_APPDOMAIN_WINDOW_KEY = u"BIMTools.ExportarLaminasPDFDWG.ActiveWindow"
_TASK_TITLE = u"Arainco: Exportar Láminas"


def _clear_appdomain_window():
    try:
        System.AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, None)
    except Exception:
        pass


def _get_active_tool_window():
    try:
        win = System.AppDomain.CurrentDomain.GetData(_APPDOMAIN_WINDOW_KEY)
    except Exception:
        return None
    if win is None:
        return None
    try:
        _ = win.Title
        if hasattr(win, "IsLoaded") and not win.IsLoaded:
            _clear_appdomain_window()
            return None
    except Exception:
        _clear_appdomain_window()
        return None
    return win


def _show_message(revit, msg, wpf_win=None):
    try:
        uiapp = None
        try:
            uiapp = revit.Application
        except Exception:
            uiapp = revit
        hwnd = revit_main_hwnd(uiapp)
        top = None
        if wpf_win is not None:
            try:
                top = wpf_win.Topmost
                wpf_win.Topmost = False
            except Exception:
                top = None
        try:
            show_message_dialog(
                _TASK_TITLE,
                msg,
                u"",
                ok_text=u"Entendido",
                hwnd_revit=hwnd,
                uiapp=uiapp,
            )
        finally:
            if wpf_win is not None and top is not None:
                try:
                    wpf_win.Topmost = top
                except Exception:
                    pass
    except Exception:
        pass


def _load_listado_core():
    try:
        from lib import listado_planos_excel_core as core

        return core
    except Exception:
        return None


def _show_componer_nombre_dialog(owner, doc, table, list_options_fn, evaluate_fn):
    show_componer_nombre_dialog(
        owner,
        doc,
        table,
        list_options_fn,
        evaluate_fn,
        get_persisted_recipe_fn=get_persisted_naming_recipe_segments,
        persist_recipe_fn=persist_naming_recipe_segments,
    )


def _build_and_show(revit):
    doc = revit.ActiveUIDocument.Document
    revit_svc = RevitWindowService(revit=revit, bloquear_cls=BloquearComandosRevit)
    folder_svc = FolderBrowserService(revit_application=revit.Application)
    progress_svc = ProgressService(pyrevit_forms=_pyrevit_forms)

    pdf_strategy = PdfExportStrategy(sanitize_fn=sanitize_file_base)
    dwg_strategy = DwgExportStrategy(sanitize_fn=sanitize_file_base)
    listado_strategy = ListadoExportStrategy(core_module=_load_listado_core())

    vm = ExportarLaminasViewModel(
        doc=doc,
        revit=revit,
        build_sheets_fn=build_sheets_datatable,
        list_fch_fn=list_fch_entrega_parameter_names_in_model,
        unique_fch_fn=unique_fecha_entrega_values_from_datatable,
        row_matches_fch_fn=datatable_row_matches_fecha_entrega_selection,
        sanitize_fn=sanitize_file_base,
        list_naming_opts_fn=list_naming_source_options,
        eval_naming_fn=evaluate_naming_recipe,
        revit_svc=revit_svc,
        progress_svc=progress_svc,
        pdf_strategy=pdf_strategy,
        dwg_strategy=dwg_strategy,
        listado_strategy=listado_strategy,
        template_listado_path=_TEMPLATE_LISTADO_XLSX,
        relay_command_cls=RelayCommand,
    )

    view = ExportarLaminasView(
        view_model=vm,
        folder_svc=folder_svc,
        revit_svc=revit_svc,
        show_componer_nombre_fn=_show_componer_nombre_dialog,
        appdomain_win_key=_APPDOMAIN_WINDOW_KEY,
    )
    view.show()


def main(revit):
    set_pushbutton_dir(_PB)
    existing = _get_active_tool_window()
    if existing is not None:
        ok = False
        try:
            from System.Windows import WindowState

            if existing.WindowState == WindowState.Minimized:
                existing.WindowState = WindowState.Normal
            existing.Show()
            existing.Activate()
            existing.Focus()
            ok = True
        except Exception:
            _clear_appdomain_window()
            existing = None
        if ok and existing is not None:
            _show_message(
                revit,
                u"La herramienta ya está en ejecución.\n\n"
                u"Si actualizó el script, cierre esta ventana y vuelva a abrir "
                u"la herramienta para cargar la nueva versión.",
                existing,
            )
            return

    try:
        _build_and_show(revit)
    except Exception as ex:
        _show_message(
            revit,
            u"No se pudo abrir el formulario.\n\n{0}".format(unicode(str(ex))),
        )
