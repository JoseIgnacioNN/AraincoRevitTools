# -*- coding: utf-8 -*-
"""
Estrategias de exportación (patrón Strategy) para la herramienta Exportar Láminas.

PdfExportStrategy     – exporta una ViewSheet a un único PDF.
DwgExportStrategy     – exporta una ViewSheet a DWG (MergedViews).
ListadoExportStrategy – genera el listado Excel de láminas seleccionadas.

Cada estrategia encapsula la dependencia externa (SheetExportManager,
listado_planos_excel_core) y expone una interfaz uniforme: execute().
"""

from lib.sheet_export_manager import SheetExportManager  # noqa: E402

# Nombre del setup DWG del proyecto utilizado por defecto.
DWG_SETUP_DEFAULT = u"Default"


# ---------------------------------------------------------------------------
# PdfExportStrategy
# ---------------------------------------------------------------------------

class PdfExportStrategy(object):
    """Exporta una lámina a un PDF único usando SheetExportManager."""

    def __init__(self, sanitize_fn=None):
        """sanitize_fn – función de saneado de nombres (opcional; se delega a SheetExportManager)."""
        self._sanitize_fn = sanitize_fn

    def execute(self, doc, pdf_dir, element_id, file_base):
        """
        Exporta la lámina indicada a PDF en pdf_dir.

        :param doc:        Document de Revit.
        :param pdf_dir:    Carpeta de salida (debe existir).
        :param element_id: ElementId de la ViewSheet.
        :param file_base:  Nombre de archivo deseado (sin extensión).
        :returns: True si se generó el archivo.
        :raises:  Excepción ante error crítico del motor de exportación.
        """
        mgr = SheetExportManager(doc, sanitize_file_base_fn=self._sanitize_fn)
        return bool(mgr.export_pdf(pdf_dir, element_id, file_base))


# ---------------------------------------------------------------------------
# DwgExportStrategy
# ---------------------------------------------------------------------------

class DwgExportStrategy(object):
    """Exporta una lámina a DWG (MergedViews) usando SheetExportManager."""

    def __init__(self, setup_name=None, sanitize_fn=None):
        """
        setup_name   – nombre del ExportDWGSettings del proyecto (p. ej. «Default»).
        sanitize_fn  – función de saneado opcional.
        """
        self._setup = setup_name or DWG_SETUP_DEFAULT
        self._sanitize_fn = sanitize_fn

    def execute(self, doc, dwg_dir, element_id, file_base):
        """
        Exporta la lámina indicada a DWG en dwg_dir.

        :param doc:        Document de Revit.
        :param dwg_dir:    Carpeta de salida (debe existir).
        :param element_id: ElementId de la ViewSheet.
        :param file_base:  Nombre de archivo deseado (sin extensión).
        :returns: True si se generó el archivo.
        """
        mgr = SheetExportManager(doc, sanitize_file_base_fn=self._sanitize_fn)
        return bool(mgr.export_dwg(dwg_dir, element_id, file_base, self._setup))


# ---------------------------------------------------------------------------
# ListadoExportStrategy
# ---------------------------------------------------------------------------

class ListadoExportStrategy(object):
    """Genera el listado Excel de láminas usando listado_planos_excel_core."""

    def __init__(self, core_module):
        """core_module – módulo listado_planos_excel_core (puede ser None si no está disponible)."""
        self._core = core_module

    @property
    def available(self):
        """True si el módulo de listado está disponible."""
        return self._core is not None

    def default_filename(self, doc):
        """Nombre de archivo Excel sugerido para el documento dado."""
        try:
            return self._core.default_listado_workbook_filename(doc)
        except Exception:
            return u"Listado_Laminas.xlsx"

    def execute(self, revit, template_path, out_path, sheets, fecha_override=None):
        """
        Genera el archivo Excel del listado.

        :param revit:          Referencia __revit__.
        :param template_path:  Ruta al TemplateListado.xlsx.
        :param out_path:       Ruta de salida del Excel generado.
        :param sheets:         Lista de ViewSheet.
        :param fecha_override: Texto de fecha a usar en todas las filas (o None).
        :returns: (n_rows, truncated) – número de filas escritas y si se truncó.
        """
        return self._core.run_export_sheets(
            revit, template_path, out_path, sheets, fecha_override
        )
