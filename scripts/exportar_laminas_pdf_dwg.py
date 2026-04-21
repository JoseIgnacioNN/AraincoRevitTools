# -*- coding: utf-8 -*-
"""
Exportación de láminas (ViewSheet) a PDF y DWG para BIMTools.
Usado por 04_ExportarLaminasPDFDWG.pushbutton.
"""

import re

import clr

clr.AddReference("RevitAPI")
clr.AddReference("System.Data")

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInParameter,
    DWGExportOptions,
    ElementId,
    FilteredElementCollector,
    PDFExportOptions,
    ViewSheet,
)
from System import Boolean, Int32, String  # noqa: E402
from System.Collections.Generic import List  # noqa: E402
from System.Data import DataColumn, DataTable  # noqa: E402


_INVALID_WIN_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def element_id_int(eid):
    try:
        return int(eid.IntegerValue)
    except Exception:
        try:
            return int(eid.Value)
        except Exception:
            return 0


def sanitize_file_base(name):
    """
    Nombre de archivo sin ruta ni extensión obligatoria.
    Elimina caracteres no válidos en Windows y recorta espacios.
    """
    if name is None:
        return u"export"
    try:
        s = unicode(name).strip()
    except Exception:
        s = u"export"
    if not s:
        return u"export"
    s = _INVALID_WIN_CHARS.sub(u"_", s)
    s = s.rstrip(u" .")
    return s if s else u"export"


def _param_as_string(elem, bip):
    try:
        p = elem.get_Parameter(bip)
        if p is None:
            return u""
        return (p.AsString() or p.AsValueString() or u"").strip()
    except Exception:
        return u""


def _param_by_bip_name(elem, bip_name):
    """Resuelve BuiltInParameter por nombre en texto para no fallar si falta el miembro del enum."""
    try:
        name = bip_name
        try:
            if isinstance(bip_name, unicode):
                name = str(bip_name)
        except Exception:
            pass
        bip = getattr(BuiltInParameter, name, None)
        if bip is None:
            return u""
        return _param_as_string(elem, bip)
    except Exception:
        return u""


def _sheet_revision_display(sheet, doc):
    """
    Texto de revisión para la grilla: GetCurrentRevision primero, luego parámetro de lámina si existe.
    """
    try:
        rid = sheet.GetCurrentRevision()
        if rid is not None:
            try:
                if rid == ElementId.InvalidElementId:
                    rid = None
            except Exception:
                pass
        if rid is not None:
            rev = doc.GetElement(rid)
            if rev is not None:
                for attr in (u"RevisionNumber", u"Description", u"SequenceNumber"):
                    try:
                        val = getattr(rev, attr, None)
                        if val is None:
                            continue
                        s = unicode(val).strip()
                        if s:
                            return s
                    except Exception:
                        pass
    except Exception:
        pass
    t = _param_by_bip_name(sheet, u"SHEET_CURRENT_REVISION")
    if t:
        return t
    try:
        p = sheet.LookupParameter(u"Current Revision")
        if p is not None:
            t = (p.AsString() or p.AsValueString() or u"").strip()
            if t:
                return t
    except Exception:
        pass
    return u""


def _sheet_size_display(sheet):
    try:
        sw = _param_by_bip_name(sheet, u"SHEET_WIDTH")
        sh = _param_by_bip_name(sheet, u"SHEET_HEIGHT")
        if sw and sh:
            return u"{} x {}".format(sw, sh)
        if sw:
            return sw
        if sh:
            return sh
    except Exception:
        pass
    return u""


def default_custom_name(sheet):
    num = u""
    try:
        num = (sheet.SheetNumber or u"").strip()
    except Exception:
        pass
    name = u""
    try:
        name = (sheet.Name or u"").strip()
    except Exception:
        pass
    if num and name:
        return sanitize_file_base(u"{}_{}".format(num, name))
    if num:
        return sanitize_file_base(num)
    if name:
        return sanitize_file_base(name)
    return u"Sheet_{}".format(element_id_int(sheet.Id))


def build_sheets_datatable(doc):
    """
    DataTable columnas: Sel, SheetNumber, SheetName, Revision, Size, CustomName, IdInt
    """
    dt = DataTable()
    dt.Columns.Add(DataColumn(u"Sel", clr.GetClrType(Boolean)))
    dt.Columns.Add(DataColumn(u"SheetNumber", clr.GetClrType(String)))
    dt.Columns.Add(DataColumn(u"SheetName", clr.GetClrType(String)))
    dt.Columns.Add(DataColumn(u"Revision", clr.GetClrType(String)))
    dt.Columns.Add(DataColumn(u"Size", clr.GetClrType(String)))
    dt.Columns.Add(DataColumn(u"CustomName", clr.GetClrType(String)))
    dt.Columns.Add(DataColumn(u"IdInt", clr.GetClrType(Int32)))

    sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
    sheets = sorted(sheets, key=lambda s: (s.SheetNumber or u"").upper())

    try:
        dt.BeginLoadData()
    except Exception:
        pass
    for sheet in sheets:
        row = dt.NewRow()
        row[u"Sel"] = Boolean(False)
        row[u"SheetNumber"] = sheet.SheetNumber or u""
        row[u"SheetName"] = sheet.Name or u""
        row[u"Revision"] = _sheet_revision_display(sheet, doc)
        row[u"Size"] = _sheet_size_display(sheet)
        row[u"CustomName"] = default_custom_name(sheet)
        row[u"IdInt"] = element_id_int(sheet.Id)
        dt.Rows.Add(row)

    try:
        dt.EndLoadData()
    except Exception:
        pass

    return dt


def export_sheet_pdf(doc, folder, sheet_id, custom_base):
    """
    Un PDF por lámina. custom_base sin ruta; se añade .pdf si falta.
    PDFExportOptions.Combine=True con una sola vista respeta FileName.
    """
    base = sanitize_file_base(custom_base)
    if not base.lower().endswith(u".pdf"):
        base = base + u".pdf"

    opts = PDFExportOptions()
    try:
        opts.Combine = True
        opts.FileName = base
        ids = List[ElementId]()
        ids.Add(sheet_id)
        return doc.Export(folder, ids, opts)
    finally:
        try:
            opts.Dispose()
        except Exception:
            pass


def export_sheet_dwg(doc, folder, sheet_id, custom_base):
    """
    Un DWG por lámina. custom_base sin extensión .dwg (Revit la añade).
    """
    base = sanitize_file_base(custom_base)
    if base.lower().endswith(u".dwg"):
        base = base[:-4]

    dwg_opts = DWGExportOptions()
    ids = List[ElementId]()
    ids.Add(sheet_id)
    return doc.Export(folder, base, ids, dwg_opts)
