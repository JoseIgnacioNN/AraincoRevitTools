# -*- coding: utf-8 -*-
"""
Exportar láminas a PDF y/o DWG con nombre de archivo personalizado por fila.
Carpeta de entrega: ruta completa en el cuadro (tras «Examinar…» se propone
YYYY.MM.DD_ENTREGA bajo la carpeta elegida; el nombre es editable). Dentro: PDF, DWG y
opcionalmente listado Excel (láminas seleccionadas; plantilla TemplateListado en la
carpeta del botón). Progreso de exportación: barras ``pyrevit.forms.ProgressBar``
consecutivas — DWG, luego PDF, luego listado Excel. Acento cian #5BC0DE.

Arquitectura MVVM
-----------------
  View           export_laminas_view.py      WPF window + XAML + event wiring
  ViewModel      export_laminas_vm.py        Estado + lógica de negocio
  Commands       export_laminas_commands.py  RelayCommand
  Services       export_laminas_services.py  BloquearComandosRevit, RevitWindowService, FolderBrowserService, ProgressService
  Strategies     export_laminas_strategies.py PdfExportStrategy, DwgExportStrategy, ListadoExportStrategy

Copias locales en esta carpeta (exportable a otra extensión sin depender de scripts/):
  revit_wpf_window_position.py, bimtools_wpf_dark_theme.py, bimtools_paths.py
  El bloqueo de Revit durante exportación vive en export_laminas_services (BloquearComandosRevit).

Sigue siendo opcional ``scripts/export_laminas_naming_schema.py`` para persistir recetas de nombre.
"""

__title__ = "Exportar\nLáminas"
__author__ = "BIMTools"
__doc__ = (
    "Selecciona y exporta láminas (PDF/DWG). Nombre Personalizado: encabezado «Nombre de archivo». "
    "Opcional: listado Excel de las seleccionadas (plantilla TemplateListado en la carpeta del botón). "
    "Ruta de entrega completa y editable; «Examinar…» la completa con YYYY.MM.DD_ENTREGA (hoy). "
    "Subcarpetas PDF y DWG."
)

import imp
import os
import sys

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")
clr.AddReference("System.Data")
clr.AddReference("System.Windows.Forms")

# ---------------------------------------------------------------------------
# Bootstrapping de rutas
# ---------------------------------------------------------------------------

_pb = os.path.dirname(os.path.abspath(__file__))

# Añadir pushbutton dir a sys.path (permite importar módulos del botón directamente)
if _pb not in sys.path:
    sys.path.insert(0, _pb)

# Buscar scripts/ del repositorio (export_laminas_naming_schema u otros; el botón ya trae
# revit_wpf_window_position, bimtools_wpf_dark_theme y bimtools_paths en su carpeta).
_d = _pb
for _ in range(24):
    _sp = os.path.join(_d, "scripts")
    if os.path.isfile(os.path.join(_sp, "revit_wpf_window_position.py")):
        if _sp not in sys.path:
            sys.path.insert(0, _sp)
        break
    _p = os.path.dirname(_d)
    if _p == _d:
        break
    _d = _p
else:
    _sp = os.path.abspath(
        os.path.join(_pb, os.pardir, os.pardir, os.pardir, os.pardir, "scripts")
    )
    if os.path.isdir(_sp) and _sp not in sys.path:
        sys.path.insert(0, _sp)

# ---------------------------------------------------------------------------
# Carga de módulos externos al pushbutton
# ---------------------------------------------------------------------------

# bimtools_paths – logo (primero copia local del pushbutton; si no existe, scripts/)
bimtools_paths = None
try:
    _bimtools_paths_fp = os.path.join(_pb, "bimtools_paths.py")
    if not os.path.isfile(_bimtools_paths_fp):
        _bimtools_paths_fp = os.path.join(
            os.path.abspath(os.path.join(_pb, os.pardir, os.pardir, os.pardir)),
            "scripts",
            "bimtools_paths.py",
        )
    if os.path.isfile(_bimtools_paths_fp):
        bimtools_paths = imp.load_source(
            "bimtools_paths__ExportarLaminasPDFDWG", _bimtools_paths_fp
        )
    else:
        import bimtools_paths as _btp
        bimtools_paths = _btp
    bimtools_paths.set_pushbutton_dir(_pb)
except Exception:
    bimtools_paths = None

# exportar_laminas_pdf_dwg – DataTable, naming, helpers de exportación
_EXPORT_MOD_PATH = os.path.join(_pb, "exportar_laminas_pdf_dwg.py")
_EXPORT_MOD_NAME = u"bimtools_exportar_laminas_pdf_dwg__04pushbutton"
for _k in (u"exportar_laminas_pdf_dwg", _EXPORT_MOD_NAME):
    try:
        if _k in sys.modules:
            del sys.modules[_k]
    except Exception:
        pass
if not os.path.isfile(_EXPORT_MOD_PATH):
    raise IOError(
        u"Falta el archivo junto al botón: {0}".format(_EXPORT_MOD_PATH)
    )
_export_el = imp.load_source(_EXPORT_MOD_NAME, _EXPORT_MOD_PATH)

build_sheets_datatable = _export_el.build_sheets_datatable
list_fch_entrega_parameter_names_in_model = _export_el.list_fch_entrega_parameter_names_in_model
unique_fecha_entrega_values_from_datatable = _export_el.unique_fecha_entrega_values_from_datatable
datatable_row_matches_fecha_entrega_selection = _export_el.datatable_row_matches_fecha_entrega_selection
export_sheet_pdf = _export_el.export_sheet_pdf
export_sheet_dwg = _export_el.export_sheet_dwg
sanitize_file_base = _export_el.sanitize_file_base
list_naming_source_options = _export_el.list_naming_source_options
evaluate_naming_recipe = _export_el.evaluate_naming_recipe

# componer_nombre_lamina_ui – diálogo de nombre personalizado
_COMPOSER_PATH = os.path.join(_pb, "componer_nombre_lamina_ui.py")
_COMPOSER_MOD_NAME = u"bimtools_componer_nombre_lamina_ui__04pushbutton"
if not os.path.isfile(_COMPOSER_PATH):
    raise IOError(u"Falta el módulo de UI: {0}".format(_COMPOSER_PATH))
_composer_el = imp.load_source(_COMPOSER_MOD_NAME, _COMPOSER_PATH)
show_componer_nombre_dialog = _composer_el.show_componer_nombre_dialog

# listado_planos_excel_core – generación de listado Excel
_LISTADO_CORE_PATH = os.path.join(_pb, u"listado_planos_excel_core.py")
_LISTADO_CORE_NAME = u"bimtools_listado_planos_excel_core__04export_laminas"
_TEMPLATE_LISTADO_XLSX = os.path.join(_pb, u"TemplateListado.xlsx")
_listado_planos_core = None
try:
    if os.path.isfile(_LISTADO_CORE_PATH):
        try:
            if _LISTADO_CORE_NAME in sys.modules:
                del sys.modules[_LISTADO_CORE_NAME]
        except Exception:
            pass
        _listado_planos_core = imp.load_source(_LISTADO_CORE_NAME, _LISTADO_CORE_PATH)
except Exception:
    _listado_planos_core = None

# Módulos MVVM del botón – cargados con imp para evitar conflictos de caché
def _load_mvvm_module(name, filename):
    mod_name = u"bimtools_exportlam_{0}__04".format(name)
    path = os.path.join(_pb, filename)
    try:
        if mod_name in sys.modules:
            del sys.modules[mod_name]
    except Exception:
        pass
    return imp.load_source(mod_name, path)

_svc_mod  = _load_mvvm_module("services",   "export_laminas_services.py")
_strat_mod = _load_mvvm_module("strategies", "export_laminas_strategies.py")
_cmd_mod  = _load_mvvm_module("commands",   "export_laminas_commands.py")
_vm_mod   = _load_mvvm_module("vm",         "export_laminas_vm.py")
_view_mod = _load_mvvm_module("view",       "export_laminas_view.py")

# ---------------------------------------------------------------------------
# Importaciones Revit y dependencias opcionales
# ---------------------------------------------------------------------------

import System  # noqa: E402

try:
    _BloquearComandosRevit = _svc_mod.BloquearComandosRevit
except Exception:
    _BloquearComandosRevit = None

try:
    from pyrevit import forms as _pyrevit_forms
except Exception:
    _pyrevit_forms = None

doc = __revit__.ActiveUIDocument.Document  # noqa: F821

# ---------------------------------------------------------------------------
# Constantes de AppDomain (singleton de ventana)
# ---------------------------------------------------------------------------

_APPDOMAIN_WINDOW_KEY = u"BIMTools.ExportarLaminasPDFDWG.ActiveWindow"
_APPDOMAIN_CONTROLLER_KEY = u"BIMTools.ExportarLaminasPDFDWG.ActiveController"
_TASK_DLG_TITLE = u"Exportar Láminas"


def _clear_appdomain():
    try:
        System.AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, None)
    except Exception:
        pass
    try:
        System.AppDomain.CurrentDomain.SetData(_APPDOMAIN_CONTROLLER_KEY, None)
    except Exception:
        pass


def _get_active_tool_window():
    """Devuelve la ventana WPF activa si existe un formulario en ejecución."""
    try:
        win = System.AppDomain.CurrentDomain.GetData(_APPDOMAIN_WINDOW_KEY)
    except Exception:
        return None
    if win is None:
        return None
    try:
        _ = win.Title
        if hasattr(win, "IsLoaded") and not win.IsLoaded:
            _clear_appdomain()
            return None
    except Exception:
        _clear_appdomain()
        return None
    return win


def _task_dialog_safe(msg, wpf_win=None):
    from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult
    top = None
    if wpf_win is not None:
        try:
            top = wpf_win.Topmost
            wpf_win.Topmost = False
        except Exception:
            top = None
    try:
        td = TaskDialog(_TASK_DLG_TITLE)
        try:
            td.TitleAutoPrefix = False
        except Exception:
            pass
        td.MainInstruction = msg
        td.CommonButtons = TaskDialogCommonButtons.Ok
        td.DefaultButton = TaskDialogResult.Ok
        td.Show()
    finally:
        if wpf_win is not None and top is not None:
            try:
                wpf_win.Topmost = top
            except Exception:
                pass

# ---------------------------------------------------------------------------
# Composición del grafo de dependencias (Composition Root)
# ---------------------------------------------------------------------------

def _build_and_show():
    # Services
    revit_svc = _svc_mod.RevitWindowService(
        revit=__revit__,  # noqa: F821
        bloquear_cls=_BloquearComandosRevit,
    )
    folder_svc = _svc_mod.FolderBrowserService(
        revit_application=__revit__.Application,  # noqa: F821
    )
    progress_svc = _svc_mod.ProgressService(
        pyrevit_forms=_pyrevit_forms,
    )

    # Strategies
    pdf_strategy = _strat_mod.PdfExportStrategy(
        sanitize_fn=sanitize_file_base,
    )
    dwg_strategy = _strat_mod.DwgExportStrategy(
        sanitize_fn=sanitize_file_base,
    )
    listado_strategy = _strat_mod.ListadoExportStrategy(
        core_module=_listado_planos_core,
    )

    # ViewModel
    vm = _vm_mod.ExportarLaminasViewModel(
        doc=doc,
        revit=__revit__,  # noqa: F821
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
        relay_command_cls=_cmd_mod.RelayCommand,
    )

    # View
    view = _view_mod.ExportarLaminasView(
        view_model=vm,
        folder_svc=folder_svc,
        revit_svc=revit_svc,
        show_componer_nombre_fn=show_componer_nombre_dialog,
        bimtools_paths_mod=bimtools_paths,
        appdomain_win_key=_APPDOMAIN_WINDOW_KEY,
        appdomain_ctrl_key=_APPDOMAIN_CONTROLLER_KEY,
    )
    view.show()

# ---------------------------------------------------------------------------
# Punto de entrada (con control de instancia única)
# ---------------------------------------------------------------------------

def main():
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
            _clear_appdomain()
            existing = None
        if ok and existing is not None:
            _task_dialog_safe(
                u"La herramienta ya está en ejecución.\n\n"
                u"Si actualizó el script, cierre esta ventana y vuelva a abrir "
                u"la herramienta para cargar la nueva versión.",
                existing,
            )
            return

    try:
        _build_and_show()
    except Exception as ex:
        try:
            _task_dialog_safe(
                u"No se pudo abrir el formulario.\n\n{0}".format(unicode(str(ex)))
            )
        except Exception:
            pass


main()
