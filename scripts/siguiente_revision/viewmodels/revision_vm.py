# -*- coding: utf-8 -*-
"""
RevisionViewModel — estado y lógica de la herramienta Revisiones.

Reemplaza el dict «state» del código monolítico original por un ViewModel real
con propiedades observables, mapas de personas y acceso a servicios.
El código-behind de la ventana lee y escribe este ViewModel en lugar
de manipular el dict directamente.
"""

from __future__ import print_function

try:
    unicode
except NameError:
    unicode = str

from siguiente_revision.viewmodels.base_vm import ObservableObject
from siguiente_revision.commands.relay_command import RelayCommand
from siguiente_revision.services import people_service
from siguiente_revision.services import revision_service
from siguiente_revision.constants import PERSONA_ROL_MODELADOR, PERSONA_ROL_INGENIERO


class RevisionViewModel(ObservableObject):
    """
    Estado completo del formulario de Revisiones.

    Propiedades observables:
        description    Descripción seleccionada.
        dibujo_display Nombre mostrado en el combo Dibujó.
        reviso_display Nombre mostrado en el combo Revisó.
        aprobo_display Nombre mostrado en el combo Aprobó.
        fecha_str      Fecha seleccionada (dd.MM.yy).
        emit_rev0      True = modo Revisión 0; False = siguiente automática.
        dialog_result  True si el usuario confirmó con OK.

    Mapas internos:
        map_dib   {display: sheet_value} para Dibujó.
        map_ing   {display: sheet_value} para Revisó/Aprobó.

    Propiedades derivadas de solo lectura:
        dibujo    Valor que va a los parámetros de lámina (abreviación).
        reviso    Idem para Revisó.
        aprobo    Idem para Aprobó.
    """

    def __init__(self, doc):
        super(RevisionViewModel, self).__init__()
        self._doc = doc

        # --- Estado formulario ---
        self._description    = u""
        self._dibujo_display = u""
        self._reviso_display = u""
        self._aprobo_display = u""
        self._fecha_str      = u""
        self._emit_rev0      = False
        self._dialog_result  = False

        # --- Mapas personas ---
        self.map_dib = {}
        self.map_ing = {}

        # --- Láminas seleccionadas (resueltas al aceptar) ---
        self.selected_sheets = []

        # --- Datos internos de UI (tabla y grid; controlados por la ventana) ---
        self.sheet_table       = None
        self._grid             = None
        self._row_delegate     = None
        self._syncing_sel_all  = False
        self._sel_anchor       = None
        self._hdr_chk          = None
        self._has_revision_zero = False

        # --- Comandos ---
        self.ok_command     = RelayCommand(self._cmd_ok)
        self.cancel_command = RelayCommand(self._cmd_cancel)

        # Referencia a la ventana (asignada por RevisionWindow al montar)
        self._win = None
        self._close_requested = False

    # -----------------------------------------------------------------------
    # Propiedades del formulario
    # -----------------------------------------------------------------------

    @property
    def description(self):
        return self._description

    @description.setter
    def description(self, v):
        self.set_property(u"_description", unicode(v or u"").strip(), u"description")

    @property
    def dibujo_display(self):
        return self._dibujo_display

    @dibujo_display.setter
    def dibujo_display(self, v):
        self.set_property(u"_dibujo_display", unicode(v or u""), u"dibujo_display")

    @property
    def reviso_display(self):
        return self._reviso_display

    @reviso_display.setter
    def reviso_display(self, v):
        self.set_property(u"_reviso_display", unicode(v or u""), u"reviso_display")

    @property
    def aprobo_display(self):
        return self._aprobo_display

    @aprobo_display.setter
    def aprobo_display(self, v):
        self.set_property(u"_aprobo_display", unicode(v or u""), u"aprobo_display")

    @property
    def fecha_str(self):
        return self._fecha_str

    @fecha_str.setter
    def fecha_str(self, v):
        self.set_property(u"_fecha_str", unicode(v or u"").strip(), u"fecha_str")

    @property
    def emit_rev0(self):
        return self._emit_rev0

    @emit_rev0.setter
    def emit_rev0(self, v):
        self.set_property(u"_emit_rev0", bool(v), u"emit_rev0")

    @property
    def dialog_result(self):
        return self._dialog_result

    # -----------------------------------------------------------------------
    # Valores para las láminas (traducción display → abreviación)
    # -----------------------------------------------------------------------

    @property
    def dibujo(self):
        return people_service.display_to_sheet_value(self._dibujo_display, self.map_dib)

    @property
    def reviso(self):
        return people_service.display_to_sheet_value(self._reviso_display, self.map_ing)

    @property
    def aprobo(self):
        return people_service.display_to_sheet_value(self._aprobo_display, self.map_ing)

    # -----------------------------------------------------------------------
    # FormData para el servicio de revisión
    # -----------------------------------------------------------------------

    def build_form_data(self):
        """Construye el FormData que recibe RevisionService.apply()."""
        return revision_service.FormData(
            description  = self._description,
            dibujo       = self.dibujo,
            reviso       = self.reviso,
            aprobo       = self.aprobo,
            fecha_str    = self._fecha_str,
            emit_rev0    = self._emit_rev0,
        )

    # -----------------------------------------------------------------------
    # Carga de personas
    # -----------------------------------------------------------------------

    def load_people(self):
        """
        Recarga los mapas de personas desde personas.json.
        Devuelve (dib_items, ing_items) para rellenar los combos en la ventana.
        """
        dib_items, dib_map = people_service.load_display_map(PERSONA_ROL_MODELADOR)
        if not dib_items:
            dib_items, dib_map = people_service.fallback_items()

        ing_items, ing_map = people_service.load_display_map(PERSONA_ROL_INGENIERO)
        if not ing_items:
            ing_items, ing_map = people_service.fallback_items()

        self.map_dib = dib_map
        self.map_ing = ing_map
        return dib_items, ing_items

    # -----------------------------------------------------------------------
    # Comandos internos (llamados por la ventana vía Click o RelayCommand)
    # -----------------------------------------------------------------------

    def _cmd_ok(self, _param=None):
        """Confirmar — la ventana debe extraer los valores del combo antes de llamar a accept()."""
        pass

    def _cmd_cancel(self, _param=None):
        """Cancelar — la ventana llama a cancel()."""
        pass

    def accept(self, win):
        """
        Llamado por la ventana cuando el usuario pulsa OK.
        Lee los valores actuales de los controles WPF y cierra.
        """
        from siguiente_revision.services.sheet_service import collect_checked_sheets
        from System.Globalization import CultureInfo
        from System import DateTime

        desc_cb = win.FindName("CbDescripcion")
        self._description = (
            unicode(desc_cb.SelectedItem) if desc_cb.SelectedItem is not None else u""
        ).strip()

        self._dibujo_display = unicode(win.FindName("CbDibujo").SelectedItem or u"")
        self._reviso_display = unicode(win.FindName("CbReviso").SelectedItem or u"")
        self._aprobo_display = unicode(win.FindName("CbAprobo").SelectedItem or u"")
        self._fecha_str      = unicode(win.FindName("CbFecha").SelectedItem or u"").strip()

        r_punct = win.FindName("RadRevisionPuntual")
        try:
            is_checked = r_punct.IsChecked if r_punct is not None else None
            if hasattr(is_checked, "HasValue"):
                self._emit_rev0 = bool(is_checked.HasValue and is_checked.Value)
            else:
                self._emit_rev0 = bool(is_checked)
        except Exception:
            self._emit_rev0 = False

        if self._emit_rev0 and not self._has_revision_zero:
            from Autodesk.Revit.UI import TaskDialog
            TaskDialog.Show(
                u"Revisiones",
                u"No hay revisión número 0 en Gestión de revisiones del proyecto.",
            )
            return

        try:
            gd = win.FindName("GridSheets")
            if gd is not None:
                gd.CommitEdit()
        except Exception:
            pass

        self.selected_sheets = collect_checked_sheets(self._doc, self.sheet_table)
        self._dialog_result = True
        self._close_win()

    def cancel(self, win):
        """Llamado por la ventana cuando el usuario cancela."""
        self._dialog_result = False
        self._close_win()

    def _close_win(self):
        if self._win is not None:
            try:
                self._win.Close()
            except Exception:
                pass
