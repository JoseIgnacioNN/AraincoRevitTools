# -*- coding: utf-8 -*-
"""
siguiente_revision — Entry point del paquete MVVM de Revisiones.

Este módulo es el único punto de entrada desde el pushbutton pyRevit:

    import siguiente_revision
    reload(siguiente_revision)
    siguiente_revision.main(__revit__)

Los sub-módulos se importan explícitamente aquí para que un reload() en
pyRevit propague correctamente los cambios a las capas de infraestructura,
servicios y UI.
"""

from __future__ import print_function

try:
    unicode
except NameError:
    unicode = str

# ---------------------------------------------------------------------------
# Re-imports de sub-módulos con imports relativos (IronPython safe).
# Imports relativos evitan el problema chicken-and-egg donde el módulo
# padre no está totalmente inicializado al ejecutar __init__.py.
# ---------------------------------------------------------------------------

from . import constants
from .infrastructure import transaction as _tx
from .infrastructure import revit_version as _rv
from .infrastructure import singleton as _sg
from .services import parameter_service as _ps
from .services import people_service as _peo
from .services import sheet_service as _sht
from .services import revision_service as _rev
from .viewmodels import base_vm as _bvm
from .viewmodels import revision_vm as _rvm
from .ui import revision_window as _win


# ---------------------------------------------------------------------------
# main()
# ---------------------------------------------------------------------------

def main(revit_app):
    """
    Punto de entrada de la herramienta Revisiones desde el pushbutton pyRevit.

    Args:
        revit_app: objeto __revit__ (UIApplication) inyectado por pyRevit.
    """
    from Autodesk.Revit.UI import TaskDialog

    # --- Resolución de doc / uidoc ---
    try:
        uidoc = revit_app.ActiveUIDocument
        doc   = uidoc.Document
    except Exception:
        try:
            doc   = revit_app.ActiveUIDocument.Document
            uidoc = revit_app.ActiveUIDocument
        except Exception:
            TaskDialog.Show(u"Revisiones", u"No hay un documento activo.")
            return

    # --- Singleton: evitar doble apertura ---
    from siguiente_revision.infrastructure import singleton
    if singleton.try_activate_existing():
        TaskDialog.Show(u"Revisiones", u"La herramienta ya está en ejecución.")
        return

    # --- Verificación de revisiones del proyecto ---
    from siguiente_revision.services import revision_service, sheet_service
    from siguiente_revision.infrastructure.revit_version import RevitVersionAdapter

    ver = RevitVersionAdapter(revit_app.Application)
    ordered = revision_service.get_ordered_revision_ids(doc)
    if not ordered:
        TaskDialog.Show(
            u"Revisiones",
            u"El proyecto no tiene revisiones definidas.\n"
            u"Cree al menos una en Gestión de revisiones.",
        )
        return

    idx_rev0 = revision_service.index_of_revision_display_number(doc, ordered, u"0")
    sheets_all = sheet_service.collect_sheets(doc)

    # --- ViewModel ---
    from siguiente_revision.viewmodels.revision_vm import RevisionViewModel
    vm = RevisionViewModel(doc)

    # --- Ventana ---
    from siguiente_revision.ui import revision_window as win_mod

    try:
        win = win_mod.load_xaml()
    except Exception as ex:
        TaskDialog.Show(
            u"Revisiones",
            u"No se pudo cargar la interfaz de usuario:\n\n{0}".format(str(ex)),
        )
        return

    win_mod.build_and_wire(
        win,
        vm,
        doc,
        uidoc,
        revit_app.Application,
        idx_rev0,
        sheets_all,
    )

    # --- Posición de ventana ---
    try:
        from revit_wpf_window_position import center_on_revit
        center_on_revit(win, revit_app.Application)
    except Exception:
        pass

    # --- Registro singleton + ShowDialog ---
    singleton.register(win)
    try:
        win.ShowDialog()
    finally:
        singleton.clear()
        # Detach row change handler
        try:
            d = vm._row_delegate
            tbl = vm.sheet_table
            if tbl is not None and d is not None:
                tbl.RowChanged -= d
        except Exception:
            pass

    # --- Post-diálogo: aplicar revisiones ---
    if not vm.dialog_result:
        return

    sel = vm.selected_sheets
    if not sel:
        from pyrevit import forms
        forms.alert(u"Marque al menos una lámina.", title=u"Revisiones")
        return

    if not vm.description:
        from pyrevit import forms
        forms.alert(u"Seleccione una descripción.", title=u"Revisiones")
        return

    from System import DateTime
    from System.Globalization import CultureInfo

    try:
        DateTime.ParseExact(vm.fecha_str, u"dd.MM.yy", CultureInfo.InvariantCulture)
    except Exception:
        from pyrevit import forms
        forms.alert(
            u"No se interpretó la fecha. Use formato dd.MM.yy.",
            title=u"Revisiones",
        )
        return

    form_data = vm.build_form_data()
    result = revision_service.apply(
        doc,
        sel,
        form_data,
        revit_uiapp=revit_app,
    )

    msg = u"Revisiones aplicadas correctamente en {} lámina(s).".format(result.done)
    if result.errors:
        msg += u"\n\nAdvertencias:\n" + result.error_text

    from pyrevit import forms
    forms.alert(msg, title=u"Revisiones")
