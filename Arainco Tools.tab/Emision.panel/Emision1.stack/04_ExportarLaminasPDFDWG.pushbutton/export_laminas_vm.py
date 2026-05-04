# -*- coding: utf-8 -*-
"""
ViewModel de la herramienta Exportar Láminas.

ExportarLaminasViewModel encapsula todo el estado de la herramienta y la lógica
de negocio (selección de láminas, filtrado, configuración de exportación,
orquestación de las fases PDF/DWG/Listado) sin referencias directas a WPF.

La comunicación con la Vista se realiza mediante callbacks registrados en tiempo
de construcción de la Vista (patrón Observer ligero, equivalente a eventos .NET):

    vm.bind_on_show_ok(lambda msg: revit_svc.show_ok(msg, win))
    vm.bind_on_exporting_changed(lambda exp: view.update_controls(exp))

Dependencias inyectadas en el constructor:
    doc, revit               → Documento y __revit__ de Revit.
    build_sheets_fn          → build_sheets_datatable
    list_fch_fn              → list_fch_entrega_parameter_names_in_model
    unique_fch_fn            → unique_fecha_entrega_values_from_datatable
    row_matches_fch_fn       → datatable_row_matches_fecha_entrega_selection
    sanitize_fn              → sanitize_file_base
    list_naming_opts_fn      → list_naming_source_options
    eval_naming_fn           → evaluate_naming_recipe
    revit_svc                → RevitWindowService
    progress_svc             → ProgressService
    pdf_strategy             → PdfExportStrategy
    dwg_strategy             → DwgExportStrategy
    listado_strategy         → ListadoExportStrategy
    template_listado_path    → ruta a TemplateListado.xlsx
    relay_command_cls        → clase RelayCommand
"""

import os
import sys

_pb = os.path.dirname(os.path.abspath(__file__))
if _pb not in sys.path:
    sys.path.insert(0, _pb)

import clr  # noqa: E402

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import ElementId, ViewSheet  # noqa: E402

# Etiquetas de fase para la barra de progreso
_PBAR_DWG = u"Arainco - Exportando DWG"
_PBAR_PDF = u"Arainco - Exportando PDF"
_PBAR_LISTADO = u"Arainco - Exportando listado Excel"


# ---------------------------------------------------------------------------
# Helpers de rutas (sin dependencias de UI)
# ---------------------------------------------------------------------------

def _default_delivery_folder_name():
    """Nombre sugerido: YYYY.MM.DD_ENTREGA (fecha local de hoy)."""
    from datetime import date
    return date.today().strftime("%Y.%m.%d") + u"_ENTREGA"


def _is_standard_delivery_folder(basename):
    """True si el segmento de ruta cumple el patrón YYYY.MM.DD_ENTREGA."""
    try:
        s = unicode(basename)
    except Exception:
        return False
    if not s.endswith(u"_ENTREGA"):
        return False
    date_part = s[:-8]
    parts = date_part.split(u".")
    if len(parts) != 3:
        return False
    if len(parts[0]) != 4 or len(parts[1]) != 2 or len(parts[2]) != 2:
        return False
    return all(p.isdigit() for p in parts)


# ---------------------------------------------------------------------------
# ViewModel
# ---------------------------------------------------------------------------

class ExportarLaminasViewModel(object):
    """
    ViewModel principal de Exportar Láminas.

    No importa ningún tipo WPF; todo acceso a la UI se realiza a través de
    callbacks registrados por la Vista.
    """

    def __init__(
        self,
        doc,
        revit,
        build_sheets_fn,
        list_fch_fn,
        unique_fch_fn,
        row_matches_fch_fn,
        sanitize_fn,
        list_naming_opts_fn,
        eval_naming_fn,
        revit_svc,
        progress_svc,
        pdf_strategy,
        dwg_strategy,
        listado_strategy,
        template_listado_path,
        relay_command_cls,
    ):
        # Dependencias inyectadas
        self._doc = doc
        self._revit = revit
        self._build_sheets = build_sheets_fn
        self._list_fch = list_fch_fn
        self._unique_fch = unique_fch_fn
        self._row_matches_fch = row_matches_fch_fn
        self._sanitize = sanitize_fn
        self._list_naming_opts = list_naming_opts_fn
        self._eval_naming = eval_naming_fn
        self._revit_svc = revit_svc
        self._progress_svc = progress_svc
        self._pdf_strategy = pdf_strategy
        self._dwg_strategy = dwg_strategy
        self._listado_strategy = listado_strategy
        self._template_listado_path = template_listado_path

        # Estado interno
        self._carpeta = u""
        self._do_pdf = True
        self._do_dwg = True
        self._do_listado = True
        self._is_exporting = False
        self._fch_param_names = []

        # Callbacks registrados por la Vista (patrón Observer)
        self._on_show_ok = None
        self._on_show_errors = None
        self._on_ask_open_folder = None
        self._on_exporting_changed = None
        self._on_estado_changed = None

        # Carga inicial de datos
        self._table = build_sheets_fn(doc)
        self._fch_param_names = list_fch_fn(doc)

        # Comandos (RelayCommand)
        self.export_command = relay_command_cls(
            execute_fn=self._execute_export,
            can_execute_fn=lambda: not self._is_exporting,
        )
        self.refresh_command = relay_command_cls(
            execute_fn=self._refresh_sheets,
        )

    # -----------------------------------------------------------------------
    # Propiedades públicas (read / read-write)
    # -----------------------------------------------------------------------

    @property
    def doc(self):
        return self._doc

    @property
    def revit(self):
        return self._revit

    @property
    def table(self):
        """DataTable de láminas (para binding directo al DataGrid)."""
        return self._table

    @property
    def table_view(self):
        """DefaultView del DataTable (para ItemsSource del DataGrid)."""
        return self._table.DefaultView

    @property
    def fch_param_names(self):
        return self._fch_param_names

    @property
    def list_naming_source_options(self):
        return self._list_naming_opts

    @property
    def evaluate_naming_recipe(self):
        return self._eval_naming

    @property
    def carpeta(self):
        return self._carpeta

    @carpeta.setter
    def carpeta(self, value):
        try:
            self._carpeta = unicode(value).strip() if value else u""
        except Exception:
            self._carpeta = u""

    @property
    def do_pdf(self):
        return self._do_pdf

    @do_pdf.setter
    def do_pdf(self, value):
        self._do_pdf = bool(value)

    @property
    def do_dwg(self):
        return self._do_dwg

    @do_dwg.setter
    def do_dwg(self, value):
        self._do_dwg = bool(value)

    @property
    def do_listado(self):
        return self._do_listado

    @do_listado.setter
    def do_listado(self, value):
        self._do_listado = bool(value)

    @property
    def is_exporting(self):
        return self._is_exporting

    # -----------------------------------------------------------------------
    # Registro de callbacks (View → VM)
    # -----------------------------------------------------------------------

    def bind_on_show_ok(self, callback):
        """callback(message: str) → None"""
        self._on_show_ok = callback

    def bind_on_show_errors(self, callback):
        """callback(main_instruction: str, errors: list) → None"""
        self._on_show_errors = callback

    def bind_on_ask_open_folder(self, callback):
        """callback(folder_path: str) → bool (True = usuario quiere abrir la carpeta)"""
        self._on_ask_open_folder = callback

    def bind_on_exporting_changed(self, callback):
        """callback(is_exporting: bool) → None"""
        self._on_exporting_changed = callback

    def bind_on_estado_changed(self, callback):
        """callback(estado_text: str) → None"""
        self._on_estado_changed = callback

    # -----------------------------------------------------------------------
    # Notificaciones internas hacia la Vista
    # -----------------------------------------------------------------------

    def _notify_show_ok(self, message):
        if self._on_show_ok is not None:
            try:
                self._on_show_ok(message)
            except Exception:
                pass

    def _notify_show_errors(self, instruction, errors):
        if self._on_show_errors is not None:
            try:
                self._on_show_errors(instruction, errors)
            except Exception:
                pass

    def _notify_ask_open_folder(self, folder_path):
        if self._on_ask_open_folder is not None:
            try:
                return bool(self._on_ask_open_folder(folder_path))
            except Exception:
                return False
        return False

    def _notify_exporting(self, is_exporting):
        self._is_exporting = is_exporting
        if self._on_exporting_changed is not None:
            try:
                self._on_exporting_changed(is_exporting)
            except Exception:
                pass

    def _notify_estado_changed(self):
        if self._on_estado_changed is not None:
            try:
                self._on_estado_changed(self.get_estado_text())
            except Exception:
                pass

    # -----------------------------------------------------------------------
    # Lógica de estado de la tabla
    # -----------------------------------------------------------------------

    def get_estado_text(self):
        """Texto de estado: «N láminas | M seleccionadas»."""
        n = self._table.Rows.Count
        ns = 0
        for i in range(n):
            try:
                if self._row_is_selected_raw(self._table.Rows[i]):
                    ns += 1
            except Exception:
                pass
        return u"{0} láminas  |  {1} seleccionadas".format(n, ns)

    def get_selected_indices(self):
        """Lista de índices de filas seleccionadas en el DataTable."""
        return [
            i for i in range(self._table.Rows.Count)
            if self._row_is_selected_raw(self._table.Rows[i])
        ]

    def get_fch_unique_values(self):
        """Lista de valores únicos de fecha de entrega del DataTable."""
        try:
            return list(self._unique_fch(self._table))
        except Exception:
            return []

    def get_visible_selection_state(self):
        """(n_visible, n_selected_visible) del DefaultView activo (respeta el filtro Buscar)."""
        dv = self._table.DefaultView
        n = dv.Count
        n_sel = 0
        for i in range(n):
            try:
                if self._row_is_selected_raw(dv[i].Row):
                    n_sel += 1
            except Exception:
                pass
        return n, n_sel

    def on_row_changed(self):
        """Llamado por la Vista cuando cambia cualquier fila del DataTable."""
        self._notify_estado_changed()

    def on_cell_edit_ending(self):
        """Llamado por la Vista al terminar la edición de una celda."""
        self._notify_estado_changed()

    # -----------------------------------------------------------------------
    # Selección de filas
    # -----------------------------------------------------------------------

    def row_is_selected(self, data_row):
        """Devuelve True si la fila DataRow tiene Sel=True."""
        return self._row_is_selected_raw(data_row)

    def _row_is_selected_raw(self, row):
        try:
            sel = row[u"Sel"]
            if self._nullable_bool(sel):
                return True
            if unicode(str(sel)).lower() == u"true":
                return True
        except Exception:
            pass
        return False

    def toggle_all_visible(self):
        """
        Alterna la selección de todas las filas visibles en el DefaultView.
        Devuelve el nuevo valor bool (True = todas marcadas).
        """
        from System import Boolean
        dv = self._table.DefaultView
        n = dv.Count
        if n == 0:
            return False
        n_sel = sum(
            1 for i in range(n) if self._row_is_selected_raw(dv[i].Row)
        )
        new_val = n_sel != n
        for i in range(n):
            try:
                dv[i].Row[u"Sel"] = Boolean(new_val)
            except Exception:
                pass
        self._notify_estado_changed()
        return new_val

    def set_rows_selected(self, data_row_views, new_value):
        """Fija la selección de una lista de DataRowView al valor indicado."""
        from System import Boolean
        for rv in data_row_views:
            try:
                rv[u"Sel"] = Boolean(new_value)
            except Exception:
                pass
        self._notify_estado_changed()

    def toggle_row(self, data_row_view, new_value=None):
        """Alterna (o fija) la selección de una única DataRowView."""
        from System import Boolean
        try:
            if new_value is None:
                new_value = not self._row_is_selected_raw(data_row_view.Row)
            data_row_view[u"Sel"] = Boolean(new_value)
        except Exception:
            pass
        self._notify_estado_changed()

    # -----------------------------------------------------------------------
    # Filtros
    # -----------------------------------------------------------------------

    def apply_search_filter(self, text):
        """Aplica el filtro de búsqueda al DefaultView del DataTable."""
        dv = self._table.DefaultView
        try:
            t = unicode(text).strip() if text else u""
        except Exception:
            t = u""
        if not t:
            dv.RowFilter = u""
        else:
            esc = t.replace(u"'", u"''")
            dv.RowFilter = (
                u"[SheetNumber] LIKE '%{0}%'"
                u" OR [SheetName] LIKE '%{0}%'"
                u" OR [FechaEntrega] LIKE '%{0}%'"
            ).format(esc)
        self._notify_estado_changed()

    def apply_fch_selection(self, key):
        """Marca las filas cuyo parámetro FCH coincide con key."""
        from System import Boolean
        try:
            key = unicode(key).strip()
        except Exception:
            key = u""
        if not key:
            return
        for i in range(self._table.Rows.Count):
            self._table.Rows[i][u"Sel"] = Boolean(
                self._row_matches_fch(self._table.Rows[i], key)
            )
        self._notify_estado_changed()

    # -----------------------------------------------------------------------
    # Carpeta de salida
    # -----------------------------------------------------------------------

    def set_carpeta_from_browse(self, base_path):
        """
        Procesa la ruta elegida en FolderBrowserDialog:
        – Si el último segmento no es una carpeta de entrega estándar,
          añade YYYY.MM.DD_ENTREGA como subcarpeta.
        Actualiza self.carpeta; la Vista puede leerlo y sincronizar el control.
        """
        if not base_path:
            return
        try:
            bn = os.path.basename(base_path.rstrip(u"\\/"))
            if _is_standard_delivery_folder(bn):
                new_txt = os.path.normpath(base_path)
            else:
                suf = _default_delivery_folder_name()
                new_txt = os.path.normpath(os.path.join(base_path, suf))
            self._carpeta = new_txt
        except Exception:
            pass

    # -----------------------------------------------------------------------
    # Refresh de datos
    # -----------------------------------------------------------------------

    def _refresh_sheets(self):
        """Recarga las láminas desde el modelo de Revit."""
        try:
            self._table = self._build_sheets(self._doc)
            self._fch_param_names = self._list_fch(self._doc)
        except Exception:
            pass
        self._notify_estado_changed()

    def get_refreshed_table(self):
        """Refresca y devuelve (table, fch_param_names) para que la Vista actualice el binding."""
        self._refresh_sheets()
        return self._table, self._fch_param_names

    # -----------------------------------------------------------------------
    # Exportación
    # -----------------------------------------------------------------------

    def _execute_export(self):
        """
        Orquesta la exportación completa (DWG → PDF → Listado Excel).
        No recibe parámetros WPF; usa callbacks para notificar a la Vista.
        """
        do_pdf = self._do_pdf
        do_dwg = self._do_dwg
        do_listado = self._do_listado

        if not do_pdf and not do_dwg and not do_listado:
            self._notify_show_ok(
                u"Marque al menos un formato: PDF, DWG o listado de planos."
            )
            return

        carpeta = self._carpeta
        if not carpeta:
            self._notify_show_ok(
                u"Indique la carpeta de entrega (ruta completa en el cuadro; "
                u"use «Examinar…» si no la tiene)."
            )
            return
        try:
            carpeta = os.path.normpath(carpeta)
        except Exception:
            pass

        pdf_dir = os.path.join(carpeta, u"PDF")
        dwg_dir = os.path.join(carpeta, u"DWG")
        try:
            if do_pdf and not os.path.isdir(pdf_dir):
                os.makedirs(pdf_dir)
            if do_dwg and not os.path.isdir(dwg_dir):
                os.makedirs(dwg_dir)
            if do_listado and not do_pdf and not do_dwg:
                if not os.path.isdir(carpeta):
                    os.makedirs(carpeta)
        except Exception as ex:
            self._notify_show_ok(
                u"No se pudieron crear las carpetas:\n\n{}".format(
                    unicode(str(ex))
                )
            )
            return

        selected_indices = self.get_selected_indices()
        n_sel = len(selected_indices)

        if do_listado and n_sel == 0:
            self._notify_show_ok(
                u"Para el listado Excel debe seleccionar al menos una lámina en la tabla."
            )
            return

        if do_listado:
            if not os.path.isfile(self._template_listado_path):
                self._notify_show_ok(
                    u"No se encontró la plantilla de listado:\n{0}".format(
                        self._template_listado_path
                    )
                )
                return
            if not self._listado_strategy.available:
                self._notify_show_ok(
                    u"No se pudo cargar el módulo de listado Excel "
                    u"(listado_planos_excel_core)."
                )
                return

        # Recopilar láminas para el listado
        sheets_for_listado = []
        if do_listado:
            for i in selected_indices:
                row = self._table.Rows[i]
                try:
                    sid = int(row[u"IdInt"])
                except Exception:
                    continue
                try:
                    eid = ElementId(sid)
                except Exception:
                    try:
                        eid = ElementId(long(sid))
                    except Exception:
                        continue
                el = self._doc.GetElement(eid)
                if el is not None and isinstance(el, ViewSheet):
                    sheets_for_listado.append(el)

        errores = []
        n_dwg_ok = 0
        n_pdf_ok = 0

        ps = self._progress_svc

        self._notify_exporting(True)
        try:
            self._revit_svc.block_revit()
            try:
                # -- Fase DWG -------------------------------------------------
                if do_dwg and n_sel > 0:
                    ps.begin_phase(ps.phase_title(_PBAR_DWG, n_sel), n_sel)
                    sn = [0]
                    for i in selected_indices:
                        got = self._parse_row_job(i)
                        if got is None:
                            ps.step(sn[0], n_sel, _PBAR_DWG)
                            sn[0] += 1
                            continue
                        eid, custom = got
                        try:
                            ok = self._dwg_strategy.execute(
                                self._doc, dwg_dir, eid, custom
                            )
                            if ok:
                                n_dwg_ok += 1
                            else:
                                errores.append(
                                    u"DWG — {0}: la exportación no generó archivo.".format(
                                        custom
                                    )
                                )
                        except Exception as ex:
                            errores.append(
                                u"DWG — {0}: {1}".format(custom, unicode(str(ex)))
                            )
                        finally:
                            ps.step(sn[0], n_sel, _PBAR_DWG)
                            sn[0] += 1
                    ps.end_phase()

                # -- Fase PDF -------------------------------------------------
                if do_pdf and n_sel > 0:
                    ps.begin_phase(ps.phase_title(_PBAR_PDF, n_sel), n_sel)
                    sn = [0]
                    for i in selected_indices:
                        got = self._parse_row_job(i)
                        if got is None:
                            ps.step(sn[0], n_sel, _PBAR_PDF)
                            sn[0] += 1
                            continue
                        eid, custom = got
                        try:
                            ok = self._pdf_strategy.execute(
                                self._doc, pdf_dir, eid, custom
                            )
                            if ok:
                                n_pdf_ok += 1
                            else:
                                errores.append(
                                    u"PDF — {0}: la exportación devolvió False.".format(
                                        custom
                                    )
                                )
                        except Exception as ex:
                            errores.append(
                                u"PDF — {0}: {1}".format(custom, unicode(str(ex)))
                            )
                        finally:
                            ps.step(sn[0], n_sel, _PBAR_PDF)
                            sn[0] += 1
                    ps.end_phase()

                # -- Fase Listado Excel ----------------------------------------
                if do_listado and n_sel > 0:
                    ps.begin_phase(ps.phase_title(_PBAR_LISTADO, 1), 1)
                    if not sheets_for_listado:
                        errores.append(
                            u"Listado Excel: la selección no contiene láminas "
                            u"válidas para el listado."
                        )
                    else:
                        try:
                            listado_out = os.path.join(
                                carpeta,
                                self._listado_strategy.default_filename(self._doc),
                            )
                            self._listado_strategy.execute(
                                self._revit,
                                self._template_listado_path,
                                listado_out,
                                sheets_for_listado,
                                self._get_fecha_emision_override(),
                            )
                        except Exception as ex:
                            errores.append(
                                u"Listado Excel: {0}".format(unicode(str(ex)))
                            )
                    ps.step(0, 1, _PBAR_LISTADO)
                    ps.end_phase()
            finally:
                self._revit_svc.unblock_revit()
                ps.end_phase()
        finally:
            self._notify_exporting(False)

        # -- Resultado: errores + pregunta para abrir carpeta -----------------
        if errores:
            self._notify_show_errors(
                u"Se registraron errores durante la exportación.", errores
            )
        try:
            open_root = os.path.normpath(unicode(carpeta).strip())
            if open_root and os.path.isdir(open_root):
                if self._notify_ask_open_folder(open_root):
                    try:
                        os.startfile(open_root)
                    except Exception:
                        try:
                            import subprocess
                            subprocess.Popen([u"explorer", open_root])
                        except Exception:
                            pass
        except Exception:
            pass

        self._notify_estado_changed()

    def _parse_row_job(self, row_index):
        """
        Extrae (ElementId, custom_name) de la fila indicada del DataTable.
        Devuelve None si la fila no tiene datos válidos.
        """
        row = self._table.Rows[row_index]
        try:
            sid = int(row[u"IdInt"])
        except Exception:
            return None
        try:
            eid = ElementId(sid)
        except Exception:
            try:
                eid = ElementId(long(sid))
            except Exception:
                return None
        try:
            custom = unicode(row[u"CustomName"]).strip()
        except Exception:
            custom = u""
        if not custom:
            custom = u"export"
        custom = self._sanitize(custom)
        return eid, custom

    def _get_fecha_emision_override(self):
        """
        La Vista registra un callable que devuelve la fecha del combo.
        Si no está registrado, devuelve None (sin override).
        """
        cb = getattr(self, "_get_fecha_from_view", None)
        if cb is None:
            return None
        try:
            return cb()
        except Exception:
            return None

    def bind_get_fecha_emision(self, callback):
        """
        La Vista registra un callable() → str|None que obtiene la fecha
        seleccionada en CmbFechaEntrega (para el listado Excel).
        """
        self._get_fecha_from_view = callback

    # -----------------------------------------------------------------------
    # Utilidades estáticas
    # -----------------------------------------------------------------------

    @staticmethod
    def _nullable_bool(wpf_nullable):
        """Convierte Nullable<bool> de WPF a bool de Python sin lanzar excepción."""
        try:
            if wpf_nullable is None:
                return False
            if hasattr(wpf_nullable, u"HasValue"):
                return bool(wpf_nullable.HasValue and wpf_nullable.Value)
            return unicode(wpf_nullable).strip().lower() == u"true"
        except Exception:
            return False
