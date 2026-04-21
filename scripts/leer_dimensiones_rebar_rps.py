# -*- coding: utf-8 -*-
"""
Lee y muestra parámetros de dimensiones de una barra de armadura (Rebar) seleccionada.

Uso: Revit 2024 + RevitPythonShell (RPS) o pyRevit. Selecciona un único Rebar y ejecuta el script.

Incluye:
- Tipo de barra (RebarBarType) y diámetro nominal si está disponible.
- Forma (RebarShape) y nombre.
- Valores de instancia con tipo de dato longitud (SpecTypeId.Length), p. ej. A, B, C de la forma.
- BuiltInParameter habituales de rebar relacionados con longitudes/recubrimientos.
- Si aplica, datos del ShapeDrivenAccessor (normal, regla de trazado, espaciado, longitud de conjunto).
- Longitud geométrica aproximada (suma de curvas de la línea media, posición 0 del conjunto).
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    SpecTypeId,
    StorageType,
    UnitUtils,
    UnitTypeId,
)
from Autodesk.Revit.DB.Structure import MultiplanarOption, Rebar, RebarBarType, RebarShape

try:
    from Autodesk.Revit.DB.Structure import RebarShapeDrivenLayoutRule
except Exception:
    RebarShapeDrivenLayoutRule = None

# ── RPS / pyRevit ────────────────────────────────────────────────────────────
try:
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
except NameError:
    doc = uidoc = None


def _sep(title):
    print(u"\n=== {} ===".format(title))


def _fmt_double_internal(val_internal):
    """Muestra valor en unidades internas y en mm (referencia)."""
    try:
        mm = UnitUtils.ConvertFromInternalUnits(float(val_internal), UnitTypeId.Millimeters)
        return u"{:.4f} ft (int.) | {:.3f} mm".format(float(val_internal), mm)
    except Exception:
        return u"{!r}".format(val_internal)


def _param_line(param):
    """Una línea de texto para un parámetro de instancia."""
    if param is None:
        return None
    try:
        name = param.Definition.Name
        st = param.StorageType
        if st == StorageType.Double:
            raw = param.AsDouble()
            avs = None
            try:
                avs = param.AsValueString()
            except Exception:
                pass
            if avs:
                return u"  {} = {}  (raw {})".format(name, avs, _fmt_double_internal(raw))
            return u"  {} = {}".format(name, _fmt_double_internal(raw))
        if st == StorageType.Integer:
            return u"  {} = {}".format(name, int(param.AsInteger()))
        if st == StorageType.String:
            return u"  {} = {}".format(name, param.AsString() or u"")
        if st == StorageType.ElementId:
            eid = param.AsElementId()
            if eid is None or eid == ElementId.InvalidElementId:
                return u"  {} = (ninguno)".format(name)
            el = doc.GetElement(eid)
            lab = el.Name if el is not None else u"?"
            return u"  {} = {} [id {}]".format(name, lab, eid.IntegerValue)
    except Exception as ex:
        return u"  (error leyendo parámetro: {})".format(ex)
    return None


def _is_length_datatype(param):
    try:
        dt = param.Definition.GetDataType()
        if dt is not None and SpecTypeId.Length is not None:
            return dt == SpecTypeId.Length
    except Exception:
        pass
    return False


def _bip_line(elem, bip_name):
    try:
        bip = getattr(BuiltInParameter, bip_name, None)
        if bip is None:
            return None
        p = elem.get_Parameter(bip)
        return _param_line(p)
    except Exception:
        return None


def _print_shape_driven(rebar):
    try:
        acc = rebar.GetShapeDrivenAccessor()
    except Exception as ex:
        print(u"  (sin acceso ShapeDriven: {})".format(ex))
        return
    if acc is None:
        print(u"  (ShapeDrivenAccessor no disponible para este Rebar)")
        return

    try:
        n = acc.Normal
        if n is not None:
            print(u"  Normal: ({:.6f}, {:.6f}, {:.6f})".format(n.X, n.Y, n.Z))
    except Exception:
        pass

    for label, call in (
        (u"Regla de trazado (LayoutRule)", lambda: acc.GetLayoutRule()),
        (u"Espaciado (Spacing)", lambda: acc.GetSpacing()),
        (u"Longitud de conjunto (ArrayLength)", lambda: acc.GetArrayLength()),
        (u"Altura (Height)", lambda: getattr(acc, "Height", None)),
    ):
        try:
            val = call()
            if val is not None:
                if label.startswith(u"Espaciado") or label.startswith(u"Longitud"):
                    print(u"  {}: {}".format(label, _fmt_double_internal(float(val))))
                elif RebarShapeDrivenLayoutRule is not None and label.startswith(u"Regla"):
                    try:
                        print(u"  {}: {} ({})".format(label, val, int(val)))
                    except Exception:
                        print(u"  {}: {}".format(label, val))
                else:
                    print(u"  {}: {}".format(label, val))
        except Exception:
            pass


def _geom_length_mm(rebar):
    try:
        crv = rebar.GetCenterlineCurves(
            False,
            False,
            False,
            MultiplanarOption.IncludeOnlyPlanarCurves,
            0,
        )
        if not crv:
            return None
        total = 0.0
        for c in crv:
            if c is not None:
                total += float(c.Length)
        mm = UnitUtils.ConvertFromInternalUnits(total, UnitTypeId.Millimeters)
        return total, mm
    except Exception:
        return None


def _run():
    global doc, uidoc
    if doc is None or uidoc is None:
        print(u"Error: ejecuta el script en RPS/pyRevit (__revit__ no disponible).")
        return

    ids = list(uidoc.Selection.GetElementIds())
    if len(ids) != 1:
        print(u"Selecciona exactamente un elemento (Rebar). Actualmente: {} seleccionados.".format(len(ids)))
        return

    elem = doc.GetElement(ids[0])
    if elem is None:
        print(u"Error: no se pudo obtener el elemento.")
        return
    if not isinstance(elem, Rebar):
        print(u"Error: el elemento no es un Rebar (es {}).".format(type(elem).__name__))
        return

    rebar = elem
    _sep(u"Rebar")
    print(u"  Id: {}".format(rebar.Id.IntegerValue))
    print(u"  Cantidad / posiciones: {} / {}".format(rebar.Quantity, rebar.NumberOfBarPositions))

    type_id = rebar.GetTypeId()
    if type_id and type_id != ElementId.InvalidElementId:
        bt = doc.GetElement(type_id)
        if isinstance(bt, RebarBarType):
            _sep(u"Tipo de barra (RebarBarType)")
            print(u"  Nombre: {}".format(bt.Name))
            try:
                bd = bt.BarDiameter
                print(u"  BarDiameter (API): {}".format(_fmt_double_internal(float(bd))))
            except Exception:
                pass
            for bip_name in ("REBAR_BAR_DIAMETER",):
                line = _bip_line(bt, bip_name)
                if line:
                    print(line)
            for nm in (u"Bar Diameter", u"Diámetro de barra", u"Diámetro"):
                p = bt.LookupParameter(nm)
                ln = _param_line(p)
                if ln:
                    print(ln)
                    break

    shape_id = rebar.GetShapeId()
    if shape_id and shape_id != ElementId.InvalidElementId:
        sh = doc.GetElement(shape_id)
        if isinstance(sh, RebarShape):
            _sep(u"Forma (RebarShape)")
            print(u"  Nombre: {}".format(sh.Name))

    _sep(u"Parámetros de instancia (tipo longitud / dimensiones de forma)")
    length_params = []
    other_lines = []
    try:
        it = rebar.Parameters
        for param in it:
            if param is None:
                continue
            if param.StorageType != StorageType.Double:
                continue
            if not _is_length_datatype(param):
                continue
            line = _param_line(param)
            if line:
                length_params.append(line)
    except Exception as ex:
        print(u"  (error iterando Parameters: {})".format(ex))

    if length_params:
        for line in sorted(set(length_params)):
            print(line)
    else:
        print(u"  (ningún parámetro Double+Length detectado; revisa idioma/API del proyecto)")

    _sep(u"BuiltInParameter de rebar (longitudes / recubrimientos, si existen)")
    for bip_name in (
        "REBAR_BAR_LENGTH",
        "REBAR_BAR_DIAMETER",
        "REBAR_BAR_MAXIMUM_BEND_RADIUS",
        "REBAR_HOOK_LENGTH_START",
        "REBAR_HOOK_LENGTH_END",
        "REBAR_COVER_TOP",
        "REBAR_COVER_BOTTOM",
        "REBAR_COVER_OTHER",
    ):
        line = _bip_line(rebar, bip_name)
        if line:
            print(line)

    _sep(u"Shape driven (layout)")
    _print_shape_driven(rebar)

    _sep(u"Geometría (línea media, posición 0)")
    gl = _geom_length_mm(rebar)
    if gl:
        total_int, mm = gl
        print(u"  Suma longitudes curvas: {} | {:.3f} mm".format(_fmt_double_internal(total_int), mm))
    else:
        print(u"  (no se pudo calcular GetCenterlineCurves para posición 0)")


_run()
