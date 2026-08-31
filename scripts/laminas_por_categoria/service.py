# -*- coding: utf-8 -*-
"""
Servicio Revit — creación de láminas por categoría.

Equivalente Python de laminasPorCategoria_script.dyn:
- Exige parámetro Validacion en Project Information (existencia).
- Cajetines OST_TitleBlocks excepto EST_A_SPLASH SCREEN.
- Correlativo {cat}-{NNN} a 3 dígitos según Clasificacion.
- Nombre fijo CONTENIDO LAMINA.
- Parámetros de firma / fecha / Clasificacion.

Revit 2024–2026 · IronPython (pyRevit).
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FamilySymbol,
    FilteredElementCollector,
    StorageType,
    Transaction,
    TransactionStatus,
    ViewSheet,
)

from laminas_por_categoria.constants import (
    NUMBER_PAD_WIDTH,
    PARAM_CLASIFICACION,
    PARAM_VALIDACION,
    SHEET_NAME,
    TITLE_BLOCK_SPLASH_NAME,
    TRANSACTION_TITLE,
)

try:
    unicode
except NameError:
    unicode = str


class LaminasPorCategoriaError(Exception):
    """Error de validación o configuración previo / durante la transacción."""


class LaminasPorCategoriaRequest(object):
    def __init__(
        self,
        title_block_id,
        categoria,
        cantidad,
        aprobo,
        calculo,
        reviso,
        dibujo,
        fecha,
    ):
        self.title_block_id = title_block_id
        self.categoria = _as_unicode(categoria).strip()
        self.cantidad = int(cantidad)
        self.aprobo = _as_unicode(aprobo).strip()
        self.calculo = _as_unicode(calculo).strip()
        self.reviso = _as_unicode(reviso).strip()
        self.dibujo = _as_unicode(dibujo).strip()
        self.fecha = _as_unicode(fecha).strip()


class LaminasPorCategoriaResult(object):
    def __init__(self):
        self.created = []
        self.skipped_invalid_numbers = []
        self.warnings = []


class TitleBlockInfo(object):
    def __init__(self, symbol_id, name):
        self.symbol_id = symbol_id
        self.name = name


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except Exception:
        return str(text)


def _element_name(element):
    if element is None:
        return u""
    try:
        n = element.Name
        if n:
            return _as_unicode(n).strip()
    except Exception:
        pass
    try:
        p = element.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p is not None and p.HasValue:
            return _as_unicode(p.AsString() or u"").strip()
    except Exception:
        pass
    return u""


def _param_string(element, name):
    if element is None:
        return u""
    try:
        p = element.LookupParameter(name)
        if p is None:
            return u""
        if p.StorageType == StorageType.String:
            return _as_unicode(p.AsString() or u"").strip()
        try:
            return _as_unicode(p.AsValueString() or u"").strip()
        except Exception:
            return u""
    except Exception:
        return u""


def _has_parameter(element, name):
    if element is None:
        return False
    try:
        if element.LookupParameter(name) is not None:
            return True
    except Exception:
        pass
    try:
        target = _as_unicode(name).strip().lower()
        for p in element.Parameters:
            try:
                defn = p.Definition
                if defn is None:
                    continue
                if _as_unicode(defn.Name).strip().lower() == target:
                    return True
            except Exception:
                continue
    except Exception:
        pass
    return False


def _set_param_string(element, name, value, builtin=None):
    p = None
    if builtin is not None:
        try:
            p = element.get_Parameter(builtin)
        except Exception:
            p = None
    if p is None:
        try:
            p = element.LookupParameter(name)
        except Exception:
            p = None
    if p is None or p.IsReadOnly:
        return False
    sval = _as_unicode(value)
    try:
        if p.StorageType == StorageType.String:
            p.Set(sval)
            return True
    except Exception:
        pass
    try:
        p.SetValueString(sval)
        return True
    except Exception:
        return False


def project_has_validacion(doc):
    """True si Project Information tiene el parámetro Validacion (existencia)."""
    try:
        info = doc.ProjectInformation
    except Exception:
        info = None
    if info is not None and _has_parameter(info, PARAM_VALIDACION):
        return True
    try:
        for el in FilteredElementCollector(doc).OfCategory(
            BuiltInCategory.OST_ProjectInformation
        ).WhereElementIsNotElementType():
            if _has_parameter(el, PARAM_VALIDACION):
                return True
    except Exception:
        pass
    return False


def collect_title_blocks(doc):
    """FamilySymbol de cajetín, sin el tipo splash. Orden alfabético por nombre."""
    items = []
    splash = TITLE_BLOCK_SPLASH_NAME.strip().lower()
    try:
        symbols = (
            FilteredElementCollector(doc)
            .OfClass(FamilySymbol)
            .OfCategory(BuiltInCategory.OST_TitleBlocks)
        )
    except Exception:
        return items
    for sym in symbols:
        name = _element_name(sym)
        if not name:
            continue
        if name.strip().lower() == splash:
            continue
        try:
            sid = int(sym.Id.IntegerValue)
        except Exception:
            continue
        items.append(TitleBlockInfo(sid, name))
    try:
        items.sort(key=lambda t: t.name.lower())
    except Exception:
        items.sort(key=lambda t: t.name)
    return items


def collect_sheets(doc):
    sheets = []
    try:
        for sh in FilteredElementCollector(doc).OfClass(ViewSheet):
            try:
                if sh.IsPlaceholder:
                    continue
            except Exception:
                pass
            sheets.append(sh)
    except Exception:
        pass
    return sheets


def _parse_numeric_suffix(sheet_number):
    """
    Segundo token tras '-'. None si no hay correlativo numérico.

    PG-013 → 13; PG → None; PG-A01 → None.
    """
    raw = _as_unicode(sheet_number).strip()
    if not raw:
        return None
    parts = raw.split(u"-")
    if len(parts) < 2:
        return None
    token = _as_unicode(parts[1]).strip()
    if not token:
        return None
    try:
        return int(token)
    except Exception:
        return None


def max_suffix_for_categoria(doc, categoria):
    """
    Máximo correlativo de láminas con Clasificacion == categoria.

    Números ilegibles se omiten (no abortan). Sin válidos → 0.
    Devuelve (max_int, lista_numeros_omitidos).
    """
    cat = _as_unicode(categoria).strip()
    skipped = []
    maximum = 0
    for sh in collect_sheets(doc):
        clas = _param_string(sh, PARAM_CLASIFICACION)
        if clas != cat:
            continue
        number = u""
        try:
            number = _as_unicode(sh.SheetNumber or u"").strip()
        except Exception:
            number = u""
        suffix = _parse_numeric_suffix(number)
        if suffix is None:
            if number:
                skipped.append(number)
            continue
        if suffix > maximum:
            maximum = suffix
    return maximum, skipped


def existing_sheet_numbers(doc):
    numbers = set()
    for sh in collect_sheets(doc):
        try:
            n = _as_unicode(sh.SheetNumber or u"").strip()
        except Exception:
            n = u""
        if n:
            numbers.add(n)
    return numbers


def format_sheet_number(categoria, suffix):
    cat = _as_unicode(categoria).strip()
    num = int(suffix)
    return u"{0}-{1:0{2}d}".format(cat, num, NUMBER_PAD_WIDTH)


def preview_numbers(doc, categoria, cantidad):
    """Lista de números que se asignarían (evita colisión global de SheetNumber)."""
    cat = _as_unicode(categoria).strip()
    try:
        n = int(cantidad)
    except Exception:
        n = 0
    if not cat or n < 1:
        return [], []
    maximum, skipped = max_suffix_for_categoria(doc, cat)
    occupied = existing_sheet_numbers(doc)
    numbers = []
    candidate = maximum + 1
    while len(numbers) < n:
        sheet_no = format_sheet_number(cat, candidate)
        if sheet_no not in occupied:
            numbers.append(sheet_no)
            occupied.add(sheet_no)
        candidate += 1
        if candidate > maximum + n + 5000:
            break
    return numbers, skipped


def _set_clasificacion(doc, sheet, value):
    if _set_param_string(sheet, PARAM_CLASIFICACION, value):
        return True
    try:
        tbs = list(
            FilteredElementCollector(doc, sheet.Id)
            .OfCategory(BuiltInCategory.OST_TitleBlocks)
            .WhereElementIsNotElementType()
        )
    except Exception:
        tbs = []
    wrote = False
    for tb in tbs:
        if _set_param_string(tb, PARAM_CLASIFICACION, value):
            wrote = True
    return wrote


def _resolve_title_block(doc, symbol_id):
    try:
        eid = ElementId(int(symbol_id))
        el = doc.GetElement(eid)
    except Exception:
        el = None
    if el is None or not isinstance(el, FamilySymbol):
        raise LaminasPorCategoriaError(
            u"No se encontró el tipo de cajetín seleccionado."
        )
    try:
        if not el.IsActive:
            el.Activate()
            doc.Regenerate()
    except Exception:
        pass
    return el


def create_laminas(doc, request):
    if request is None:
        raise LaminasPorCategoriaError(u"No hay datos de creación.")
    if not project_has_validacion(doc):
        raise LaminasPorCategoriaError(
            u"Este proyecto no tiene el parámetro Validacion en "
            u"Información de proyecto. Use una plantilla Arainco."
        )
    cat = request.categoria
    if not cat:
        raise LaminasPorCategoriaError(u"Seleccione una categoría.")
    if request.cantidad < 1:
        raise LaminasPorCategoriaError(u"La cantidad debe ser un entero mayor que 0.")

    numbers, skipped = preview_numbers(doc, cat, request.cantidad)
    if len(numbers) != request.cantidad:
        raise LaminasPorCategoriaError(
            u"No se pudieron generar {0} números de lámina libres.".format(
                request.cantidad
            )
        )

    result = LaminasPorCategoriaResult()
    result.skipped_invalid_numbers = list(skipped)
    if skipped:
        result.warnings.append(
            u"Se omitieron {0} lámina(s) de {1} con número no numérico "
            u"al calcular el correlativo: {2}.".format(
                len(skipped), cat, u", ".join(skipped[:8])
            )
        )

    t = Transaction(doc, TRANSACTION_TITLE)
    t.Start()
    try:
        symbol = _resolve_title_block(doc, request.title_block_id)
        tb_id = symbol.Id
        for sheet_no in numbers:
            sheet = ViewSheet.Create(doc, tb_id)
            if sheet is None:
                raise LaminasPorCategoriaError(
                    u"Revit no pudo crear la lámina {0}.".format(sheet_no)
                )
            try:
                sheet.SheetNumber = sheet_no
            except Exception as ex:
                raise LaminasPorCategoriaError(
                    u"No se pudo asignar el número {0}: {1}.".format(sheet_no, ex)
                )
            try:
                sheet.Name = SHEET_NAME
            except Exception:
                _set_param_string(
                    sheet, u"Sheet Name", SHEET_NAME, BuiltInParameter.SHEET_NAME
                )

            ok_all = True
            ok_all = _set_param_string(
                sheet, u"Approved By", request.aprobo, BuiltInParameter.SHEET_APPROVED_BY
            ) and ok_all
            ok_all = _set_param_string(
                sheet, u"Designed By", request.calculo, BuiltInParameter.SHEET_DESIGNED_BY
            ) and ok_all
            ok_all = _set_param_string(
                sheet, u"Checked By", request.reviso, BuiltInParameter.SHEET_CHECKED_BY
            ) and ok_all
            ok_all = _set_param_string(
                sheet, u"Drawn By", request.dibujo, BuiltInParameter.SHEET_DRAWN_BY
            ) and ok_all
            ok_all = _set_param_string(
                sheet,
                u"Sheet Issue Date",
                request.fecha,
                BuiltInParameter.SHEET_ISSUE_DATE,
            ) and ok_all
            if not _set_clasificacion(doc, sheet, cat):
                result.warnings.append(
                    u"{0}: no se pudo escribir Clasificacion.".format(sheet_no)
                )
            elif not ok_all:
                result.warnings.append(
                    u"{0}: alguno de los parámetros de firma/fecha no se escribió.".format(
                        sheet_no
                    )
                )
            result.created.append(sheet_no)
        t.Commit()
    except Exception:
        try:
            if t.GetStatus() == TransactionStatus.Started:
                t.RollBack()
        except Exception:
            pass
        raise
    return result


def format_success_dialog(result, categoria):
    n = len(result.created) if result is not None else 0
    instruction = u"Se crearon {0} lámina(s) {1}.".format(n, categoria)
    lines = []
    if result is not None and result.created:
        if n <= 12:
            lines.append(u"Números: {0}.".format(u", ".join(result.created)))
        else:
            lines.append(
                u"Números: {0} … {1}.".format(result.created[0], result.created[-1])
            )
    lines.append(u"Nombre: {0}.".format(SHEET_NAME))
    if result is not None:
        for w in result.warnings:
            lines.append(w)
    content = u"\n".join(lines)
    return instruction, content
