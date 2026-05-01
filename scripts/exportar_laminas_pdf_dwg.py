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
    BuiltInCategory,
    BuiltInParameter,
    Category,
    DWGExportOptions,
    ElementId,
    ExportDWGSettings,
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


# --- Composición de nombre de archivo (parámetros + especiales) ---
_NAMING_KEY_SEP = u"\x1f"


def _iter_element_parameters(elem):
    try:
        lst = elem.GetOrderedParameters()
        if lst is not None:
            for p in lst:
                yield p
            return
    except Exception:
        pass
    try:
        for p in elem.Parameters:
            yield p
    except Exception:
        pass


def _lookup_sheet_parameter_display_string(sheet, definition_name):
    if not definition_name:
        return u""
    try:
        p = sheet.LookupParameter(definition_name)
        if p is None:
            return u""
        s = (p.AsString() or p.AsValueString() or u"").strip()
        return unicode(s) if s else u""
    except Exception:
        return u""


def _sheets_category_id(doc):
    try:
        cat = Category.GetCategory(doc, BuiltInCategory.OST_Sheets)
        if cat is None:
            return None
        return cat.Id
    except Exception:
        return None


def _binding_includes_sheets_category(binding, sheets_cat_id):
    if binding is None or sheets_cat_id is None:
        return False
    try:
        cats = binding.Categories
    except Exception:
        return False
    try:
        tid = sheets_cat_id.IntegerValue
    except Exception:
        try:
            tid = int(sheets_cat_id.Value)
        except Exception:
            return False
    try:
        for c in cats:
            try:
                if c is None or c.Id is None:
                    continue
                if c.Id.IntegerValue == tid:
                    return True
            except Exception:
                continue
    except Exception:
        pass
    return False


def _definition_names_from_parameter_bindings_for_sheets(doc, sheets_cat_id):
    names = []
    if sheets_cat_id is None:
        return names
    try:
        it = doc.ParameterBindings.ForwardIterator()
        it.Reset()
        while it.MoveNext():
            try:
                defn = it.Key
                binding = it.Current
            except Exception:
                continue
            if defn is None or binding is None:
                continue
            if not _binding_includes_sheets_category(binding, sheets_cat_id):
                continue
            try:
                nm = (defn.Name or u"").strip()
                if nm:
                    names.append(nm)
            except Exception:
                continue
    except Exception:
        pass
    return names


def _definition_names_on_sheet(sheet):
    out = []
    for p in _iter_element_parameters(sheet):
        try:
            if p is None:
                continue
            dfn = p.Definition
            if dfn is None:
                continue
            nm = (dfn.Name or u"").strip()
            if nm:
                out.append(nm)
        except Exception:
            continue
    return out


def list_naming_source_options(doc):
    opts = []
    opts.append(
        {u"Key": u"SPECIAL" + _NAMING_KEY_SEP + u"SheetNumber", u"Label": u"Número de lámina"}
    )
    opts.append({u"Key": u"SPECIAL" + _NAMING_KEY_SEP + u"SheetName", u"Label": u"Nombre de lámina"})
    opts.append({u"Key": u"SPECIAL" + _NAMING_KEY_SEP + u"Revision", u"Label": u"Revisión (actual)"})
    opts.append({u"Key": u"SPECIAL" + _NAMING_KEY_SEP + u"Size", u"Label": u"Tamaño de lámina"})

    seen = set()
    sheets_cat_id = _sheets_category_id(doc)

    for nm in _definition_names_from_parameter_bindings_for_sheets(doc, sheets_cat_id):
        if nm:
            seen.add(nm)

    try:
        sheets = list(FilteredElementCollector(doc).OfClass(ViewSheet).ToElements())
    except Exception:
        sheets = []

    for sh in sheets:
        for nm in _definition_names_on_sheet(sh):
            if nm:
                seen.add(nm)

    for nm in sorted(seen, key=lambda x: x.upper()):
        opts.append({u"Key": u"DEF" + _NAMING_KEY_SEP + nm, u"Label": nm})

    return opts


def resolve_naming_source_value(sheet, doc, key):
    if not key:
        return u""
    try:
        k = unicode(key)
    except Exception:
        return u""
    if _NAMING_KEY_SEP not in k:
        return u""
    kind, payload = k.split(_NAMING_KEY_SEP, 1)
    kind = (kind or u"").strip().upper()
    payload = payload or u""

    if kind == u"SPECIAL":
        if payload == u"SheetNumber":
            try:
                return unicode(sheet.SheetNumber or u"").strip()
            except Exception:
                return u""
        if payload == u"SheetName":
            try:
                return unicode(sheet.Name or u"").strip()
            except Exception:
                return u""
        if payload == u"Revision":
            return _sheet_revision_display(sheet, doc)
        if payload == u"Size":
            return _sheet_size_display(sheet)
        return u""

    if kind == u"DEF":
        return _lookup_sheet_parameter_display_string(sheet, payload)

    return u""


def evaluate_naming_recipe(sheet, doc, recipe_segments):
    if sheet is None or doc is None:
        return sanitize_file_base(u"")
    parts = []
    for seg in recipe_segments or []:
        try:
            src_key = seg.get(u"Key", u"")
        except Exception:
            src_key = u""
        val = resolve_naming_source_value(sheet, doc, src_key)
        try:
            pre = unicode(seg.get(u"Prefix") or u"")
            suf = unicode(seg.get(u"Suffix") or u"")
            sep = unicode(seg.get(u"Separator") or u"")
        except Exception:
            pre = suf = sep = u""
        parts.append(pre + val + suf + sep)
    raw = u"".join(parts)
    return sanitize_file_base(raw)


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


def export_sheet_dwg(doc, folder, sheet_id, custom_base, dwg_setup_name=None):
    """
    Un DWG por lámina. custom_base sin extensión .dwg (Revit la añade).
    dwg_setup_name: nombre de ExportDWGSettings del proyecto (p. ej. u"Default"); si falta o no existe,
    se usan opciones por defecto. Siempre se fuerza MergedViews=True (una sola salida,
    sin xrefs por vistas en lámina / vínculos como en UI desmarcado).
    """
    base = sanitize_file_base(custom_base)
    if base.lower().endswith(u".dwg"):
        base = base[:-4]

    dwg_opts = None
    try:
        sn = u""
        if dwg_setup_name is not None:
            try:
                sn = unicode(dwg_setup_name).strip()
            except Exception:
                sn = u""
        if sn:
            try:
                st = ExportDWGSettings.FindByName(doc, sn)
            except Exception:
                st = None
            if st is not None:
                try:
                    dwg_opts = st.GetDWGExportOptions()
                except Exception:
                    dwg_opts = None
        if dwg_opts is None:
            dwg_opts = DWGExportOptions()
        dwg_opts.MergedViews = True
        ids = List[ElementId]()
        ids.Add(sheet_id)
        return doc.Export(folder, base, ids, dwg_opts)
    finally:
        if dwg_opts is not None:
            try:
                dwg_opts.Dispose()
            except Exception:
                pass
