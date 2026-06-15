# -*- coding: utf-8 -*-
"""
Numeración de columnas estructurales apiladas (módulo portable del pushbutton).

Agrupa torres de columnas (y fundación asociada) por «firma» de tipos y alturas,
luego asigna el mismo número de lote a cada torre de la misma configuración.
Escribe el parámetro de instancia «Numeracion Columna» en todas las columnas de cada torre.
"""

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import (
    BuiltInCategory,
    FilteredElementCollector,
    StorageType,
    Transaction,
)
from Autodesk.Revit.UI import TaskDialog

_TRANSACTION_NAME = u"Arainco: Numeración de columnas"
_TOOL_DIALOG_TITLE = u"Arainco: Numeración de columnas"

# Tolerancia en pies (0.05 pies ≈ 1.5 cm)
TOLERANCE = 0.05

_NUMERACION_COLUMNA_PARAM_CANDIDATES = (
    u"Numeracion Columna",
    u"Numeración Columna",
)


def _type_id_int(type_id):
    if type_id is None or type_id == type_id.InvalidElementId:
        return -1
    try:
        return int(type_id.Value)
    except Exception:
        try:
            return type_id.IntegerValue
        except Exception:
            return -1


def _lookup_numeracion_param(element):
    for name in _NUMERACION_COLUMNA_PARAM_CANDIDATES:
        param = element.LookupParameter(name)
        if param is not None and not param.IsReadOnly:
            return param
    return None


def _set_numeracion_columna(element, value):
    param = _lookup_numeracion_param(element)
    if param is None:
        return False
    if param.StorageType == StorageType.Integer:
        param.Set(int(value))
    else:
        param.Set(str(value))
    return True


def _type_display_name(doc, type_id):
    if doc is None or type_id is None:
        return u"?"
    try:
        if type_id == type_id.InvalidElementId:
            return u"?"
    except Exception:
        pass
    try:
        sym = doc.GetElement(type_id)
        if sym is None:
            return u"?"
        for attr in ("Name",):
            try:
                n = getattr(sym, attr, None)
                if n:
                    return unicode(n)
            except Exception:
                pass
        try:
            fn = sym.FamilyName
            if fn:
                return unicode(fn)
        except Exception:
            pass
    except Exception:
        pass
    return u"?"


def abbrev_type_label(name, max_len=22):
    """Nombre corto para la leyenda del esquema (el completo va en ToolTip)."""
    if not name:
        return u"?"
    s = unicode(name).strip().replace(u"_", u" ")
    up = s.upper()
    if u"FUNDACION" in up and u"AISLAD" in up:
        return u"Fund. aislada"
    if u"SIN FUND" in up:
        return u"Sin fund."
    for marker in (u"COLUMN", u"PILAR", u"FOUNDATION", u"FUNDACION"):
        if marker in up:
            pos = up.find(marker)
            tail = s[pos + len(marker) :].strip(u" _-")
            if tail:
                s = tail
                up = s.upper()
            break
    for token in (
        u"EST M ",
        u"EST ",
        u"STRUCTURAL ",
        u"RECTANGULAR",
        u"RECT ",
        u"COLUMN ",
        u"FOUNDATION ",
    ):
        if up.startswith(token):
            s = s[len(token) :].strip()
            up = s.upper()
    while u"  " in s:
        s = s.replace(u"  ", u" ")
    s = s.strip()
    if len(s) > max_len:
        return s[: max(1, max_len - 1)] + u"\u2026"
    return s if s else u"?"


class StackedElement(object):
    """Geometría de bounding box para columnas y fundaciones."""

    def __init__(self, element):
        self.element = element
        self.id = element.Id
        self.type_id = element.GetTypeId()

        bb = element.get_BoundingBox(None)
        if bb:
            self.min_x = bb.Min.X
            self.min_y = bb.Min.Y
            self.min_z = bb.Min.Z
            self.max_x = bb.Max.X
            self.max_y = bb.Max.Y
            self.max_z = bb.Max.Z
            self.height = abs(self.max_z - self.min_z)
        else:
            self.min_x = self.min_y = self.min_z = 0
            self.max_x = self.max_y = self.max_z = 0
            self.height = 0


class StackedColumn(StackedElement):
    def __init__(self, element):
        StackedElement.__init__(self, element)
        self.top_col = None
        self.bottom_col = None
        self.foundation = None

    def is_directly_under(self, other_col):
        if abs(self.max_z - other_col.min_z) > TOLERANCE:
            return False
        overlap_x = (self.max_x > other_col.min_x - TOLERANCE) and (
            self.min_x < other_col.max_x + TOLERANCE
        )
        overlap_y = (self.max_y > other_col.min_y - TOLERANCE) and (
            self.min_y < other_col.max_y + TOLERANCE
        )
        return overlap_x and overlap_y

    def is_resting_on(self, foundation):
        if abs(self.min_z - foundation.max_z) > TOLERANCE:
            return False
        overlap_x = (self.max_x > foundation.min_x - TOLERANCE) and (
            self.min_x < foundation.max_x + TOLERANCE
        )
        overlap_y = (self.max_y > foundation.min_y - TOLERANCE) and (
            self.min_y < foundation.max_y + TOLERANCE
        )
        return overlap_x and overlap_y


def _build_signature(root):
    signature_list = []
    if root.foundation:
        f_height = round(root.foundation.height, 2)
        signature_list.append(
            u"F_{0}_{1}".format(
                _type_id_int(root.foundation.type_id),
                f_height,
            )
        )
    else:
        signature_list.append(u"F_None")
    current = root
    while current is not None:
        signature_list.append(u"C_{0}".format(_type_id_int(current.type_id)))
        current = current.top_col
    return tuple(signature_list)


def _segments_from_root(doc, root):
    """Segmentos de abajo a arriba para el esquema simplificado."""
    segs = []
    if root.foundation:
        h = float(root.foundation.height)
        if h < 0.05:
            h = 0.5
        full = _type_display_name(doc, root.foundation.type_id)
        segs.append(
            {
                u"kind": u"fundacion",
                u"label": abbrev_type_label(full),
                u"label_full": full,
                u"h_ft": h,
            }
        )
    else:
        segs.append(
            {
                u"kind": u"sin_fundacion",
                u"label": u"Sin fund.",
                u"label_full": u"Sin fundación",
                u"h_ft": 0.35,
            }
        )
    tramo = 1
    current = root
    while current is not None:
        h = float(current.height)
        if h < 0.05:
            h = 1.0
        full = _type_display_name(doc, current.type_id)
        segs.append(
            {
                u"kind": u"columna",
                u"label": abbrev_type_label(full),
                u"label_full": full,
                u"h_ft": h,
                u"tramo": tramo,
            }
        )
        current = current.top_col
        tramo += 1
    return segs


def analyze_column_stacks(doc):
    """
    Analiza el proyecto y devuelve torres agrupadas por firma.

    Returns:
        dict con keys: ``ok``, ``message``, ``roots``, ``ordered_groups``,
        ``lotes`` (lista ordenada por número de lote para la UI).
    """
    if doc is None:
        return {u"ok": False, u"message": u"No hay documento."}

    col_collector = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_StructuralColumns)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    found_collector = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_StructuralFoundation)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    if not col_collector:
        return {
            u"ok": False,
            u"message": u"No se encontraron columnas estructurales en el proyecto.",
        }

    columns_data = [
        StackedColumn(c) for c in col_collector if c.get_BoundingBox(None) is not None
    ]
    foundations_data = [
        StackedElement(f)
        for f in found_collector
        if f.get_BoundingBox(None) is not None
    ]

    for c1 in columns_data:
        for c2 in columns_data:
            if c1.id == c2.id:
                continue
            if c1.is_directly_under(c2):
                c1.top_col = c2
                c2.bottom_col = c1

    roots = [c for c in columns_data if c.bottom_col is None]

    for root in roots:
        for f in foundations_data:
            if root.is_resting_on(f):
                root.foundation = f
                break

    stack_groups = {}
    for root in roots:
        sig = _build_signature(root)
        if sig not in stack_groups:
            stack_groups[sig] = []
        stack_groups[sig].append(root)

    ordered_groups = sorted(stack_groups.items(), key=lambda kv: kv[0])
    lotes = []
    for lote_no, (signature, matching_roots) in enumerate(ordered_groups, 1):
        rep = matching_roots[0]
        lotes.append(
            {
                u"lote": lote_no,
                u"torres_count": len(matching_roots),
                u"signature": signature,
                u"representante": rep,
                u"segmentos": _segments_from_root(doc, rep),
            }
        )

    lotes.sort(key=lambda x: int(x[u"lote"]))

    return {
        u"ok": True,
        u"message": u"",
        u"roots": roots,
        u"ordered_groups": ordered_groups,
        u"lotes": lotes,
        u"n_configuraciones": len(lotes),
        u"n_torres": len(roots),
    }


def apply_numeracion(doc, ordered_groups):
    """
    Escribe «Numeracion Columna» según ``ordered_groups``
    (lista de ``(signature, [StackedColumn root, ...])``).
    """
    parametros_no_encontrados = 0
    t = Transaction(doc, _TRANSACTION_NAME)
    t.Start()
    try:
        for lote_no, (_sig, matching_roots) in enumerate(ordered_groups, 1):
            for root in matching_roots:
                current = root
                while current is not None:
                    if not _set_numeracion_columna(current.element, lote_no):
                        parametros_no_encontrados += 1
                    current = current.top_col
        t.Commit()
    except Exception:
        t.RollBack()
        raise
    return {u"parametros_no_encontrados": parametros_no_encontrados}


def run(revit):
    """Entrada pyRevit: abre la interfaz con galería de esquemas."""
    from numerar_columnas_ui import run as run_ui

    run_ui(revit)


def run_headless(revit):
    """Numeración directa sin UI (compatibilidad / pruebas)."""
    uidoc = revit.ActiveUIDocument
    if uidoc is None:
        TaskDialog.Show(_TOOL_DIALOG_TITLE, u"No hay un documento activo.")
        return
    doc = uidoc.Document
    analysis = analyze_column_stacks(doc)
    if not analysis.get(u"ok"):
        TaskDialog.Show(_TOOL_DIALOG_TITLE, analysis.get(u"message") or u"Error.")
        return
    apply_numeracion(doc, analysis[u"ordered_groups"])


def format_resultado_numeracion(n_torres, n_config, parametros_no_encontrados):
    """Texto de resumen para la barra de la UI (sin TaskDialog)."""
    msg = (
        u"Numeración aplicada: {0} torres, {1} lotes."
    ).format(int(n_torres), int(n_config))
    n_miss = int(parametros_no_encontrados or 0)
    if n_miss > 0:
        msg += (
            u" Advertencia: no se pudo escribir «Numeracion Columna» en "
            u"{0} columna(s) (parámetro ausente o solo lectura)."
        ).format(n_miss)
    return msg
