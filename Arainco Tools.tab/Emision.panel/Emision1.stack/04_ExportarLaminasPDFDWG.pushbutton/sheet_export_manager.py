# -*- coding: utf-8 -*-
"""
Núcleo de exportación de láminas (PDF / DWG) para Revit 2024–2026.

Diseñado para entornos pyRevit sobre .NET Framework (2024) y .NET 8 (2025/2026):
listas genéricas ``System.Collections.Generic.List``, tipos DB actuales y
``UnitUtils`` con ``ForgeTypeId`` donde aplica.

La API ``Document.Export`` para PDF/DWG no requiere ``Transaction``; el ámbito
``RevitTransactionScope`` se expone para extensiones o pasos que modifiquen el modelo.
"""

from __future__ import absolute_import

import os
import re

try:
    unicode
except NameError:
    unicode = str

import clr  # noqa: F401 / pyRevit

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (  # noqa: E402
    DWGExportOptions,
    ElementId,
    ExportDWGSettings,
    PDFExportOptions,
)
from Autodesk.Revit.DB import UnitUtils  # noqa: E402
from System.Collections.Generic import List  # noqa: E402

_INVALID_WIN_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def _unicode_safe(value, default=u""):
    try:
        if value is None:
            return default
        return unicode(value).strip()
    except Exception:
        try:
            return str(value).strip()
        except Exception:
            return default


def sanitize_file_base(name):
    """
    Normaliza el nombre de archivo (sin ruta): caracteres válidos en Windows
    y manejo seguro de nombres vacíos o solo caracteres prohibidos.

    :param name: Texto de entrada (puede contener caracteres Unicode válidos en NTFS).
    :returns: Cadena segura; nunca cadena vacía.
    """
    if name is None:
        return u"export"
    s = _unicode_safe(name, u"export")
    if not s:
        return u"export"
    s = _INVALID_WIN_CHARS.sub(u"_", s)
    s = s.rstrip(u" .")
    return s if s else u"export"


def parse_revit_year(application):
    """
    Obtiene el año mayor de Revit desde ``Application.VersionNumber`` (p. ej. ``2025``).

    :param application: ``Autodesk.Revit.ApplicationServices.Application`` o compatible.
    :returns: Entero 0 si no se pudo interpretar.
    """
    if application is None:
        return 0
    try:
        raw = _unicode_safe(application.VersionNumber, u"")
    except Exception:
        raw = u""
    if not raw:
        return 0
    token = raw
    for sep in (u" ", u".", u"-"):
        if sep in token:
            token = token.split(sep)[0]
            break
    try:
        return int(token)
    except Exception:
        return 0


class RevitTransactionScope(object):
    """
    Context manager para ``Transaction`` con cierre explícito ante excepciones.

    No usar para ``Document.Export`` (PDF/DWG), que no debe ir dentro de una transacción.

    :param document: Documento activo.
    :param name: Nombre visible de la transacción en el historial de Undo.
    """

    __slots__ = ("_doc", "_name", "_tx")

    def __init__(self, document, name):
        self._doc = document
        self._name = name or u"BIMTools"
        self._tx = None

    def __enter__(self):
        from Autodesk.Revit.DB import Transaction

        self._tx = Transaction(self._doc, self._name)
        self._tx.Start()
        return self._tx

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._tx is None:
            return False
        try:
            if exc_type is not None:
                self._tx.RollBack()
            else:
                self._tx.Commit()
        except Exception:
            try:
                self._tx.RollBack()
            except Exception:
                pass
        finally:
            self._tx = None
        return False


def internal_length_to_mm(length_internal):
    """
    Convierte longitud interna de Revit a milímetros (API moderna con ``ForgeTypeId``).

    Útil para márgenes, escalas de hábito o futuros ajustes de tamaño de lámina.
    Falla en silencio y devuelve ``None`` si la conversión no está disponible.

    :param length_internal: Valor en unidades internas de Revit.
    :returns: ``float`` en mm o ``None``.
    """
    if length_internal is None:
        return None
    try:
        from Autodesk.Revit.DB import UnitTypeId
    except Exception:
        return None
    try:
        return float(UnitUtils.ConvertFromInternalUnits(float(length_internal), UnitTypeId.Millimeters))
    except Exception:
        return None


def mm_to_internal_length(mm):
    """
    Convierte milímetros a unidades internas de Revit.

    :param mm: Valor en mm.
    :returns: ``float`` interno o ``None`` si falla.
    """
    if mm is None:
        return None
    try:
        from Autodesk.Revit.DB import UnitTypeId
    except Exception:
        return None
    try:
        return float(UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters))
    except Exception:
        return None


def _pyrevit_logger():
    try:
        from pyrevit import script

        return script.get_logger()
    except Exception:
        return None


def _log_warning(message):
    log = _pyrevit_logger()
    if log is None:
        return
    try:
        log.warning(message)
    except Exception:
        pass


def _log_debug(message):
    log = _pyrevit_logger()
    if log is None:
        return
    try:
        log.debug(message)
    except Exception:
        pass


def _export_result_ok(result):
    if result is None:
        return False
    try:
        return int(result.Count) > 0
    except Exception:
        pass
    try:
        return len(result) > 0
    except Exception:
        pass
    try:
        return bool(result)
    except Exception:
        return False


def _make_unique_stem(folder, stem, extension):
    """
    Si ``stem + extension`` ya existe en ``folder``, añade ``_2``, ``_3``, …

    Evita sobrescritura silenciosa con nombres duplicados o caracteres
    colisionantes tras sanitizar.

    :param folder: Directorio de salida (existente).
    :param stem: Nombre sin extensión.
    :param extension: Extensión incluyendo punto (``.pdf`` o ``.dwg``).
    :returns: Nuevo stem único.
    """
    stem = sanitize_file_base(stem)
    if stem.lower().endswith(extension.lower()):
        stem = stem[: -len(extension)]
    candidate = stem
    try:
        n = 2
        while os.path.isfile(os.path.join(folder, candidate + extension)):
            candidate = u"{}_{}".format(stem, n)
            n += 1
    except Exception:
        return stem
    return candidate


class SheetExportManager(object):
    """
    Orquesta exportación de una ``ViewSheet`` a PDF o DWG con opciones compatibles
    2024–2026 y rutas únicas ante colisiones.

    :param doc: ``Document`` de Revit.
    :param sanitize_file_base_fn: Opcional; sustituye la función de saneado de nombres (tests / inyección).
    """

    __slots__ = ("_doc", "_app", "_year", "_sanitize")

    def __init__(self, doc, sanitize_file_base_fn=None):
        self._doc = doc
        try:
            self._app = doc.Application if doc is not None else None
        except Exception:
            self._app = None
        self._year = parse_revit_year(self._app)
        self._sanitize = sanitize_file_base_fn or sanitize_file_base

    @property
    def document(self):
        return self._doc

    @property
    def revit_year(self):
        return self._year

    def export_pdf(self, folder, sheet_id, custom_base):
        """
        Exporta una lámina a un único PDF (``PDFExportOptions.Combine = True``).

        Revit 2024+: ``FileName`` y ``Combine`` siguen siendo el patrón recomendado
        para un archivo por lámina con nombre controlado.

        :param folder: Carpeta de salida.
        :param sheet_id: ``ElementId`` de la ``ViewSheet``.
        :param custom_base: Nombre deseado (con o sin ``.pdf``).
        :returns: ``True`` si ``Document.Export`` devolvió éxito (colección de nombres no vacía).
        """
        if self._doc is None or sheet_id is None:
            return False
        base = self._sanitize(custom_base)
        if not base.lower().endswith(u".pdf"):
            base = base + u".pdf"
        stem = base[:-4]
        stem = _make_unique_stem(folder, stem, u".pdf")
        base = stem + u".pdf"

        opts = PDFExportOptions()
        try:
            opts.Combine = True
            opts.FileName = base
            ids = List[ElementId]()
            ids.Add(sheet_id)
            result = self._doc.Export(folder, ids, opts)
            ok = _export_result_ok(result)
            if not ok:
                _log_warning(u"[Exportar láminas] PDF sin salida para: {}".format(base))
            else:
                _log_debug(u"[Exportar láminas] PDF OK año={} archivo={}".format(self._year, base))
            return ok
        except Exception as ex:
            _log_warning(u"[Exportar láminas] PDF error: {}".format(_unicode_safe(ex, u"error")))
            return False
        finally:
            try:
                opts.Dispose()
            except Exception:
                pass

    def export_dwg(self, folder, sheet_id, custom_base, dwg_setup_name=None):
        """
        Exporta una lámina a DWG con ``MergedViews=True`` (una salida, sin xrefs por vistas).

        Opciones base: ``ExportDWGSettings.FindByName`` si existe el setup indicado;
        en otro caso ``DWGExportOptions()`` nuevo.

        :param folder: Carpeta de salida.
        :param sheet_id: ``ElementId`` de la ``ViewSheet``.
        :param custom_base: Nombre sin ruta; ``.dwg`` se elimina si viene incluido.
        :param dwg_setup_name: Nombre del setup de proyecto (p. ej. ``Default``).
        :returns: ``True`` si la exportación devolvió éxito.
        """
        if self._doc is None or sheet_id is None:
            return False
        base = self._sanitize(custom_base)
        if base.lower().endswith(u".dwg"):
            base = base[:-4]
        base = _make_unique_stem(folder, base, u".dwg")

        dwg_opts = None
        try:
            sn = u""
            if dwg_setup_name is not None:
                sn = _unicode_safe(dwg_setup_name, u"")
            if sn:
                try:
                    st = ExportDWGSettings.FindByName(self._doc, sn)
                except Exception:
                    st = None
                if st is not None:
                    try:
                        dwg_opts = st.GetDWGExportOptions()
                    except Exception:
                        dwg_opts = None
            if dwg_opts is None:
                dwg_opts = DWGExportOptions()
            try:
                dwg_opts.MergedViews = True
            except Exception:
                pass
            ids = List[ElementId]()
            ids.Add(sheet_id)
            result = self._doc.Export(folder, base, ids, dwg_opts)
            ok = _export_result_ok(result)
            if not ok:
                _log_warning(u"[Exportar láminas] DWG sin salida para: {}".format(base))
            else:
                _log_debug(u"[Exportar láminas] DWG OK año={} archivo={}".format(self._year, base))
            return ok
        except Exception as ex:
            _log_warning(u"[Exportar láminas] DWG error: {}".format(_unicode_safe(ex, u"error")))
            return False
        finally:
            if dwg_opts is not None:
                try:
                    dwg_opts.Dispose()
                except Exception:
                    pass


def transaction(document, name):
    """
    Alias usable como ``with transaction(doc, u\"Nombre\"):`` (decorador de contexto).

    :returns: :class:`RevitTransactionScope`
    """
    return RevitTransactionScope(document, name)
