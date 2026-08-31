# -*- coding: utf-8 -*-
"""
Dominio Exportar Láminas: DataTable de láminas, filtros FCH, naming y persistencia.

Dominio del paquete ``exportar_laminas``. La receta «Nombre personalizado» se
persiste vía ``export_laminas_naming_schema``.
"""

import clr

clr.AddReference("RevitAPI")
clr.AddReference("System.Data")

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInCategory,
    BuiltInParameter,
    Category,
    ElementId,
    FilteredElementCollector,
    ViewSheet,
)
from System import Boolean, Int32, String  # noqa: E402
from System.Data import DataColumn, DataTable  # noqa: E402

from exportar_laminas.lib.sheet_export_manager import SheetExportManager, sanitize_file_base  # noqa: E402

try:
    import exportar_laminas.lib.export_laminas_naming_schema as _lam_nm_store  # noqa: E402
except Exception:
    _lam_nm_store = None


def element_id_int(eid):
    try:
        return int(eid.IntegerValue)
    except Exception:
        try:
            return int(eid.Value)
        except Exception:
            return 0


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
# Clave interna: TIPO + separador U+001F + carga (p. ej. nombre de parámetro)
_NAMING_KEY_SEP = u"\x1f"
_FECHA_ENTREGA_PARAM_SUBSTRING = u"FCH"
_FECHA_ENTREGA_TOKEN_SEP = u"\x1f"


def _iter_element_parameters(elem):
    """Itera parámetros de un elemento (API según versión de Revit)."""
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
    """Valor legible de un parámetro de lámina por nombre de definición."""
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
    """ElementId de la categoría Láminas (OST_Sheets)."""
    try:
        cat = Category.GetCategory(doc, BuiltInCategory.OST_Sheets)
        if cat is None:
            return None
        return cat.Id
    except Exception:
        return None


def _binding_includes_sheets_category(binding, sheets_cat_id):
    """True si el binding de proyecto incluye la categoría Láminas."""
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
    """Nombres de parámetros de proyecto enlazados a la categoría Láminas."""
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
    """Nombres de definición de todos los parámetros visibles en una lámina."""
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


class SheetCatalog(object):
    """
    Índice de láminas del documento: un collector y un walk de parámetros.

    Evita repetir FilteredElementCollector + GetOrderedParameters al abrir
    (DataTable, FCH, opciones de naming y receta persistida).
    """

    def __init__(self, doc):
        self.doc = doc
        self.sheets = []
        self.param_names = set()
        self.fch_param_names = []
        self._fch_pairs_by_id = {}
        self._naming_opts = None
        self._naming_keys = None

    def fch_pairs_for(self, sheet):
        if sheet is None:
            return []
        return self._fch_pairs_by_id.get(element_id_int(sheet.Id), [])

    def naming_options(self):
        if self._naming_opts is None:
            self._naming_opts = _build_naming_source_options(self.doc, self.param_names)
            self._naming_keys = set()
            for o in self._naming_opts:
                try:
                    k = o.get(u"Key", u"")
                    if k:
                        self._naming_keys.add(unicode(k))
                except Exception:
                    continue
        return self._naming_opts

    def naming_keys(self):
        self.naming_options()
        return self._naming_keys or set()


def build_sheet_catalog(doc):
    """Construye :class:`SheetCatalog` con un único pase sobre las láminas."""
    cat = SheetCatalog(doc)
    if doc is None:
        return cat
    try:
        sheets = list(FilteredElementCollector(doc).OfClass(ViewSheet).ToElements())
    except Exception:
        sheets = []
    try:
        cat.sheets = sorted(sheets, key=lambda s: (s.SheetNumber or u"").upper())
    except Exception:
        cat.sheets = list(sheets)
    all_names = set()
    all_fch = set()
    sub = _FECHA_ENTREGA_PARAM_SUBSTRING.upper()
    for sh in cat.sheets:
        sid = element_id_int(sh.Id)
        fch_pairs = []
        seen_fch = set()
        for p in _iter_element_parameters(sh):
            try:
                if p is None:
                    continue
                dfn = p.Definition
                if dfn is None:
                    continue
                nm = (dfn.Name or u"").strip()
                if not nm:
                    continue
            except Exception:
                continue
            all_names.add(nm)
            if sub in nm.upper() and nm not in seen_fch:
                seen_fch.add(nm)
                try:
                    val = (p.AsString() or p.AsValueString() or u"").strip()
                except Exception:
                    val = u""
                fch_pairs.append((nm, val))
                all_fch.add(nm)
        try:
            fch_pairs.sort(key=lambda x: x[0].upper())
        except Exception:
            pass
        cat._fch_pairs_by_id[sid] = fch_pairs
    cat.param_names = all_names
    cat.fch_param_names = sorted(all_fch, key=lambda x: x.upper())
    return cat


def _build_naming_source_options(doc, instance_param_names):
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
    if instance_param_names:
        for nm in instance_param_names:
            if nm:
                seen.add(nm)
    for nm in sorted(seen, key=lambda x: x.upper()):
        opts.append({u"Key": u"DEF" + _NAMING_KEY_SEP + nm, u"Label": nm})
    return opts


def list_naming_source_options(doc, catalog=None):
    """
    Opciones para el compositor: campos especiales de lámina + parámetros
    enlazados a categoría Láminas y unión de parámetros presentes en las láminas.
    """
    if catalog is not None:
        return catalog.naming_options()
    if doc is None:
        return _build_naming_source_options(None, None)
    cat = build_sheet_catalog(doc)
    return cat.naming_options()


def resolve_naming_source_value(sheet, doc, key):
    """Resuelve una clave de list_naming_source_options a texto (sin sanitizar)."""
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


def filter_recipe_segments_for_document(doc, segments, catalog=None):
    """
    Conserva solo segmentos cuya ``Key`` sigue existiendo en ``list_naming_source_options``.
    """
    if doc is None or not segments:
        return []
    if catalog is not None:
        allowed = catalog.naming_keys()
    else:
        allowed = set()
        for o in list_naming_source_options(doc):
            try:
                k = o.get(u"Key", u"")
                if k:
                    allowed.add(unicode(k))
            except Exception:
                continue
    out = []
    for s in segments:
        try:
            k = unicode(s.get(u"Key", u""))
        except Exception:
            k = u""
        if k not in allowed:
            continue
        try:
            out.append(
                {
                    u"Key": k,
                    u"Prefix": unicode(s.get(u"Prefix") or u""),
                    u"Suffix": unicode(s.get(u"Suffix") or u""),
                    u"Separator": unicode(s.get(u"Separator") or u""),
                }
            )
        except Exception:
            continue
    return out


def get_persisted_naming_recipe_segments(doc, catalog=None):
    """
    Lee la receta guardada en el documento (Extensible Storage) y la filtra por opciones vigentes.
    """
    if _lam_nm_store is None or doc is None:
        return []
    try:
        raw = _lam_nm_store.load_recipe_segments(doc)
    except Exception:
        raw = []
    return filter_recipe_segments_for_document(doc, raw, catalog=catalog)


def persist_naming_recipe_segments(doc, segments_dict_list, catalog=None):
    """Guarda la receta en ``ProjectInformation`` (una transacción)."""
    if _lam_nm_store is None or doc is None:
        return False
    try:
        filtered = filter_recipe_segments_for_document(
            doc, segments_dict_list, catalog=catalog
        )
        return _lam_nm_store.save_recipe_segments(doc, filtered)
    except Exception:
        return False


def evaluate_naming_recipe(sheet, doc, recipe_segments):
    """
    recipe_segments: lista de dicts con claves Key, Prefix, Suffix, Separator (unicode).
    Devuelve nombre base ya sanitizado para usar como CustomName.
    """
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
    """
    Nombre de archivo sugerido por defecto: «Número de lámina - Nombre de lámina».
    """
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
        return sanitize_file_base(u"{} - {}".format(num, name))
    if num:
        return sanitize_file_base(num)
    if name:
        return sanitize_file_base(name)
    return u"Sheet_{}".format(element_id_int(sheet.Id))


def _fch_param_names_on_sheet_sorted(sheet):
    """Nombres de definición en la lámina cuyo nombre contiene ``FCH`` (orden alfabético, sin duplicados)."""
    sub = _FECHA_ENTREGA_PARAM_SUBSTRING.upper()
    found = []
    for nm in _definition_names_on_sheet(sheet):
        try:
            n = (nm or u"").strip()
        except Exception:
            n = u""
        if n and sub in n.upper():
            found.append(n)
    seen = set()
    out = []
    for n in sorted(found, key=lambda x: x.upper()):
        if n not in seen:
            seen.add(n)
            out.append(n)
    return out


def _sheet_fch_entrega_display_and_tokens(sheet, catalog=None):
    """
    Recorre **todos** los parámetros FCH de la lámina, acumula valores de texto no vacíos
    (como Dynamo con varias claves). Devuelve (texto para grilla, cadena de tokens para filtrar).
    """
    values_ordered = []
    seen_val = set()
    if catalog is not None:
        pairs = catalog.fch_pairs_for(sheet)
        iterable = [val for _nm, val in pairs]
    else:
        iterable = []
        for nm in _fch_param_names_on_sheet_sorted(sheet):
            iterable.append(_lookup_sheet_parameter_display_string(sheet, nm))
    for disp in iterable:
        disp = (disp or u"").strip()
        if disp and disp not in seen_val:
            seen_val.add(disp)
            values_ordered.append(disp)
    if not values_ordered:
        return u"", u""
    display = u"; ".join(values_ordered)
    tokens = _FECHA_ENTREGA_TOKEN_SEP.join(values_ordered)
    return display, tokens


def list_fch_entrega_parameter_names_in_model(doc, catalog=None):
    """Todos los nombres de parámetro distintos que contienen ``FCH`` y existen en al menos una lámina."""
    if catalog is not None:
        return list(catalog.fch_param_names)
    if doc is None:
        return []
    return list(build_sheet_catalog(doc).fch_param_names)


def model_has_fch_entrega_params(doc):
    return len(list_fch_entrega_parameter_names_in_model(doc)) > 0


def discover_fecha_entrega_param_name(doc):
    """
    Compatibilidad: devuelve un solo nombre FCH (el primero alfabéticamente).
    Para lógica de selección use :func:`_sheet_fch_entrega_display_and_tokens` / tokens en DataTable.
    """
    names = list_fch_entrega_parameter_names_in_model(doc)
    return names[0] if names else None


def unique_fecha_entrega_values_from_datatable(dt):
    """Valores únicos no vacíos entre todas las filas (desde ``FechaEntregaTokens``), ordenados."""
    if dt is None:
        return []
    seen = set()
    out = []
    n = dt.Rows.Count
    for i in range(n):
        try:
            tok = dt.Rows[i][u"FechaEntregaTokens"]
            tok = unicode(tok) if tok is not None else u""
        except Exception:
            tok = u""
        for part in tok.split(_FECHA_ENTREGA_TOKEN_SEP):
            v = (part or u"").strip()
            if v and v not in seen:
                seen.add(v)
                out.append(v)
    out.sort(key=lambda x: x.upper())
    return out


def datatable_row_matches_fecha_entrega_selection(row, selected_key):
    """True si algún valor FCH de la fila coincide exactamente con ``selected_key`` (trim)."""
    if row is None or not selected_key:
        return False
    try:
        key = unicode(selected_key).strip()
    except Exception:
        key = u""
    if not key:
        return False
    try:
        tok = row[u"FechaEntregaTokens"]
        tok = unicode(tok) if tok is not None else u""
    except Exception:
        tok = u""
    for part in tok.split(_FECHA_ENTREGA_TOKEN_SEP):
        if (part or u"").strip() == key:
            return True
    return False


def build_sheets_datatable(doc, fecha_entrega_param_name=None, catalog=None):
    """
    DataTable columnas: Sel, SheetNumber, SheetName, Revision, Size, FechaEntrega,
    FechaEntregaTokens, CustomName, IdInt

    ``fecha_entrega_param_name`` se ignora: se consideran **todos** los parámetros cuyo nombre
    contiene ``FCH`` en cada lámina.
    """
    if catalog is None:
        catalog = build_sheet_catalog(doc)

    dt = DataTable()
    dt.Columns.Add(DataColumn(u"Sel", clr.GetClrType(Boolean)))
    dt.Columns.Add(DataColumn(u"SheetNumber", clr.GetClrType(String)))
    dt.Columns.Add(DataColumn(u"SheetName", clr.GetClrType(String)))
    dt.Columns.Add(DataColumn(u"Revision", clr.GetClrType(String)))
    dt.Columns.Add(DataColumn(u"Size", clr.GetClrType(String)))
    dt.Columns.Add(DataColumn(u"FechaEntrega", clr.GetClrType(String)))
    dt.Columns.Add(DataColumn(u"FechaEntregaTokens", clr.GetClrType(String)))
    dt.Columns.Add(DataColumn(u"CustomName", clr.GetClrType(String)))
    dt.Columns.Add(DataColumn(u"IdInt", clr.GetClrType(Int32)))

    sheets = catalog.sheets
    persisted_recipe = get_persisted_naming_recipe_segments(doc, catalog=catalog)

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
        disp, tok = _sheet_fch_entrega_display_and_tokens(sheet, catalog=catalog)
        row[u"FechaEntrega"] = disp
        row[u"FechaEntregaTokens"] = tok
        if persisted_recipe:
            row[u"CustomName"] = evaluate_naming_recipe(sheet, doc, persisted_recipe)
        else:
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
    Un PDF por lámina. Delega en :class:`SheetExportManager` (opciones vigentes 2024–2026).

    :param doc: Documento de Revit.
    :param folder: Carpeta de salida.
    :param sheet_id: ``ElementId`` de la lámina.
    :param custom_base: Nombre de archivo deseado sin extensión (Revit añade ``.pdf``).
    :returns: ``True`` si la exportación generó salida.
    """
    return SheetExportManager(doc, sanitize_file_base_fn=sanitize_file_base).export_pdf(
        folder, sheet_id, custom_base
    )


def export_sheet_dwg(doc, folder, sheet_id, custom_base, dwg_setup_name=None):
    """
    Un DWG por lámina con ``MergedViews`` (ver :class:`SheetExportManager`).

    :returns: ``True`` si la exportación generó salida.
    """
    return SheetExportManager(doc, sanitize_file_base_fn=sanitize_file_base).export_dwg(
        folder, sheet_id, custom_base, dwg_setup_name
    )
