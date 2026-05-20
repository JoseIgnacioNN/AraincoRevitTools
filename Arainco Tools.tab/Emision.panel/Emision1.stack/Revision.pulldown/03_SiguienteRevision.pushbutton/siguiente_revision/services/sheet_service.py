# -*- coding: utf-8 -*-
"""
SheetService — colección y construcción de la tabla de láminas.

Provee las funciones de acceso a ViewSheet del documento Revit y la construcción
del DataTable que alimenta el DataGrid en la UI.
"""

from __future__ import print_function

try:
    unicode
except NameError:
    unicode = str

import clr
clr.AddReference("System.Data")
clr.AddReference("RevitAPI")

from System import Boolean, String
from System import Int64
from System.Data import DataColumn, DataTable

from Autodesk.Revit.DB import FilteredElementCollector, ViewSheet

from siguiente_revision.constants import EXCLUDE_SHEET_NAME_SUBSTR


def collect_sheets(doc):
    """
    Retorna todas las láminas no-placeholder del documento, ordenadas por número y nombre.
    """
    out = []
    for vs in FilteredElementCollector(doc).OfClass(ViewSheet):
        if getattr(vs, "IsPlaceholder", False):
            continue
        out.append(vs)
    return sorted(out, key=lambda s: (s.SheetNumber or u"", s.Name or u""))


def sheet_display(sheet):
    """Texto «número - nombre» para mensajes."""
    return u"{} - {}".format(sheet.SheetNumber or u"?", sheet.Name or u"?")


def sheet_revision_cell(sheet, doc):
    """Revisión actual formateada para la celda de la tabla."""
    try:
        from exportar_laminas_pdf_dwg import _sheet_revision_display
        if _sheet_revision_display is not None:
            return (unicode(_sheet_revision_display(sheet, doc)).strip())
    except Exception:
        pass
    return u""


def row_id_for_table(sheet):
    """
    Valor numérico estable de sheet.Id como Int64.

    Revit 2024+ usa identificadores de 64 bits; no usar Int32 en la tabla.
    """
    try:
        return int(sheet.Id.Value)
    except Exception:
        try:
            return int(sheet.Id.IntegerValue)
        except Exception:
            return 0


def build_selection_table(doc, sheets_all, revision_service=None):
    """
    Construye un DataTable con columnas:
        Sel, SelEnabled, SheetNumber, SheetName, Revision, NuevaRevision, IdInt.

    Omite láminas con «splash screen» en el nombre.

    Args:
        doc: Revit Document.
        sheets_all: lista de ViewSheet.
        revision_service: RevisionService opcional para calcular NuevaRevision/SelEnabled.
    """
    tbl = DataTable()
    tbl.Columns.Add(DataColumn(u"Sel",          clr.GetClrType(Boolean)))
    tbl.Columns.Add(DataColumn(u"SelEnabled",   clr.GetClrType(Boolean)))
    tbl.Columns.Add(DataColumn(u"SheetNumber",  clr.GetClrType(String)))
    tbl.Columns.Add(DataColumn(u"SheetName",    clr.GetClrType(String)))
    tbl.Columns.Add(DataColumn(u"Revision",     clr.GetClrType(String)))
    tbl.Columns.Add(DataColumn(u"NuevaRevision",clr.GetClrType(String)))
    tbl.Columns.Add(DataColumn(u"IdInt",        clr.GetClrType(Int64)))
    try:
        tbl.BeginLoadData()
    except Exception:
        pass
    for s in sheets_all:
        nm = unicode(s.Name or u"")
        if EXCLUDE_SHEET_NAME_SUBSTR in nm.lower():
            continue
        row = tbl.NewRow()
        row[u"Sel"]           = False
        row[u"SelEnabled"]    = True
        row[u"SheetNumber"]   = unicode(s.SheetNumber or u"")
        row[u"SheetName"]     = unicode(s.Name or u"")
        row[u"Revision"]      = sheet_revision_cell(s, doc)
        row[u"NuevaRevision"] = u""
        row[u"IdInt"]         = Int64(row_id_for_table(s))
        tbl.Rows.Add(row)
    try:
        tbl.EndLoadData()
    except Exception:
        pass
    try:
        tbl.AcceptChanges()
    except Exception:
        pass
    return tbl


def collect_checked_sheets(doc, tbl, version_adapter=None):
    """
    Retorna la lista de ViewSheet cuya columna «Sel» es True en el DataTable.
    """
    from siguiente_revision.infrastructure.revit_version import RevitVersionAdapter

    out = []
    if tbl is None:
        return out
    adapter = version_adapter
    n = int(tbl.Rows.Count)
    for i in range(n):
        row = tbl.Rows[i]
        try:
            if not bool(row[u"Sel"]):
                continue
            eid_val = row[u"IdInt"]
            if adapter is not None:
                eid = adapter.element_id_from_int(eid_val)
            else:
                from Autodesk.Revit.DB import ElementId
                from System import Int64 as _I64
                try:
                    eid = ElementId(_I64(int(eid_val)))
                except Exception:
                    eid = ElementId(int(eid_val))
            el = doc.GetElement(eid)
            if isinstance(el, ViewSheet):
                out.append(el)
        except Exception:
            continue
    return out
