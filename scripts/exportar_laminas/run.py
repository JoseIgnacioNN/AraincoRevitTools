# -*- coding: utf-8 -*-
"""
Composition root — Exportar Láminas (MVVM).

Paquete ``scripts/exportar_laminas/`` (botón ligero).
"""

import os

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

import System  # noqa: E402

try:
    unicode
except NameError:
    unicode = str


def _as_unicode(value):
    if value is None:
        return u""
    try:
        return unicode(value)
    except Exception:
        try:
            return str(value)
        except Exception:
            return u""

from bimtools_paths import get_pushbutton_dir  # noqa: E402
from revit_wpf_window_position import revit_main_hwnd  # noqa: E402
from exportar_laminas.lib.exportar_laminas_pdf_dwg import (  # noqa: E402
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
from exportar_laminas.mvvm.export_laminas_commands import RelayCommand  # noqa: E402
from exportar_laminas.mvvm.export_laminas_services import (  # noqa: E402
    BloquearComandosRevit,
    FolderBrowserService,
    ProgressService,
    RevitWindowService,
)
from exportar_laminas.mvvm.export_laminas_strategies import (  # noqa: E402
    DwgExportStrategy,
    ListadoExportStrategy,
    PdfExportStrategy,
)
from exportar_laminas.mvvm.export_laminas_vm import ExportarLaminasViewModel  # noqa: E402
from exportar_laminas.ui.componer_nombre_lamina_ui import show_componer_nombre_dialog  # noqa: E402
from exportar_laminas.ui.export_laminas_instruction_dialog import show_message_dialog  # noqa: E402
from exportar_laminas.ui.export_laminas_view import ExportarLaminasView  # noqa: E402

try:
    from pyrevit import forms as _pyrevit_forms
except Exception:
    _pyrevit_forms = None

def _template_listado_path():
    pb = get_pushbutton_dir()
    if not pb:
        return u""
    return os.path.join(pb, u"TemplateListado.xlsx")

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
    text = _as_unicode(msg).strip() or u"Error desconocido."
    shown = False
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
            shown = bool(
                show_message_dialog(
                    _TASK_TITLE,
                    text,
                    u"",
                    ok_text=u"Entendido",
                    hwnd_revit=hwnd,
                    uiapp=uiapp,
                )
            )
        finally:
            if wpf_win is not None and top is not None:
                try:
                    wpf_win.Topmost = top
                except Exception:
                    pass
    except Exception:
        shown = False
    if shown:
        return
    try:
        from Autodesk.Revit.UI import TaskDialog

        TaskDialog.Show(_TASK_TITLE, text)
    except Exception:
        try:
            if _pyrevit_forms is not None:
                _pyrevit_forms.alert(text, title=_TASK_TITLE)
        except Exception:
            pass


def _load_listado_core():
    try:
        from exportar_laminas.lib import listado_planos_excel_core as core

        return core
    except Exception:
        return None


def _show_componer_nombre_dialog(owner, doc, table, list_options_fn, evaluate_fn, catalog=None):
    show_componer_nombre_dialog(
        owner,
        doc,
        table,
        list_options_fn,
        evaluate_fn,
        get_persisted_recipe_fn=lambda d: get_persisted_naming_recipe_segments(
            d, catalog=catalog
        ),
        persist_recipe_fn=lambda d, segs: persist_naming_recipe_segments(
            d, segs, catalog=catalog
        ),
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
        template_listado_path=_template_listado_path(),
        relay_command_cls=RelayCommand,
    )

    def _show_componer(owner, doc, table, list_fn, eval_fn):
        _show_componer_nombre_dialog(
            owner, doc, table, list_fn, eval_fn, catalog=vm._catalog
        )

    view = ExportarLaminasView(
        view_model=vm,
        folder_svc=folder_svc,
        revit_svc=revit_svc,
        show_componer_nombre_fn=_show_componer,
        appdomain_win_key=_APPDOMAIN_WINDOW_KEY,
    )
    view.show()


def main(revit):
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
            u"No se pudo abrir el formulario.\n\n{0}".format(_as_unicode(ex)),
        )


def run(revit):
    """Punto de entrada del botón ligero."""
    main(revit)
