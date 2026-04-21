# -*- coding: utf-8 -*-
"""
Script RPS: estribos rectangulares en vigas (Structural Framing, tipo Beam).

Ejecutable en RevitPythonShell (RPS) o pyRevit. Requiere una o más vigas seleccionadas.
En RPS, ejecuta el .py desde archivo (p. ej. menú del shell / Execute Script) para que exista
ruta; si pegas el código en la consola, el script igual funciona (no depende de imports locales).

Enfoque:
- Un Rebar por viga: lazo rectangular en el plano ⟂ al eje (normal = eje).
- Tras crearlo, SetLayoutAsMaximumSpacing; por defecto primera/última barra activas (``include_end_bars=True``). El flujo BIMTools de vigas pasa ``False`` solo en el tramo de estribos **centrales**.
- Recubrimiento, paso máximo y respeto en extremos (END_INSET_MM) configurables abajo.

Limitaciones:
- Vigas con eje casi vertical (|eje·Z| > 0,92) se omiten (misma regla que armadura_vigas_capas).
- Sección y ancho/canto se toman de parámetros de tipo (Width/Height/…) o del bounding box.
"""

from __future__ import print_function

import os
import sys
import clr

# RPS / ejecución desde consola: a veces no existe __file__.
_scripts_dir = None
try:
    _scripts_dir = os.path.dirname(os.path.abspath(__file__))
except NameError:
    _argv = getattr(sys, "argv", None) or []
    if _argv and _argv[0] and not str(_argv[0]).startswith("-"):
        try:
            _scripts_dir = os.path.dirname(os.path.abspath(_argv[0]))
        except Exception:
            _scripts_dir = None
if _scripts_dir and os.path.isdir(_scripts_dir) and _scripts_dir not in sys.path:
    sys.path.insert(0, _scripts_dir)

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

import System
from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BuiltInCategory,
    Curve,
    ElementId,
    FamilyInstance,
    FilteredElementCollector,
    Line,
    LocationCurve,
    Transaction,
    UnitUtils,
    UnitTypeId,
    XYZ,
)
from Autodesk.Revit.DB.Structure import (
    Rebar,
    RebarBarType,
    RebarHookOrientation,
    RebarHookType,
    RebarShape,
    RebarStyle,
    StructuralMaterialType,
    StructuralType,
)

try:
    from Autodesk.Revit.DB.Structure import RebarShapeDrivenLayoutRule
except Exception:
    RebarShapeDrivenLayoutRule = None

# ── Parámetros de prueba (editar) ───────────────────────────────────────────
COVER_MM = 25.0
SPACING_MM = 150.0
END_INSET_MM = 50.0
# Si no es None, usa el RebarBarType cuyo nombre contenga este texto (insensible a mayúsculas).
BAR_TYPE_NAME_CONTAINS = None
# Si True, al fallar una viga se imprime el último mensaje de excepción de la API Revit.
PRINT_API_ERRORS = True
# Máximo de RebarShape con estilo StirrupTie a probar en el fallback.
MAX_STIRRUP_SHAPES_TO_TRY = 16
# ─────────────────────────────────────────────────────────────────────────────

try:
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
except NameError:
    doc = uidoc = None


def _mm_to_internal(mm):
    return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)


def _curve_array_from_lines(lines):
    """IList<Curve>: array CLR de Line (a veces falla en IronPython vs List[Curve])."""
    ct = clr.GetClrType(Line).BaseType
    arr = System.Array.CreateInstance(ct, len(lines))
    for i, ln in enumerate(lines):
        arr[i] = ln
    return arr


def _curve_list_ilist(lines):
    """IList<Curve> como List[Curve] (preferido para CreateFromCurves en RPS)."""
    lst = List[Curve]()
    for ln in lines:
        lst.Add(ln)
    return lst


def _curve_list_object(lines):
    """IList para CreateFromCurvesAndShape (mismo patrón que armadura_vigas_capas)."""
    lst = List[object]()
    for ln in lines:
        lst.Add(ln)
    return lst


def _stirrup_center_at_station(host, p0, axis, dist_along):
    """
    Punto en el plano de la sección: sobre la línea analítica + corrección al centro del bbox
    del host en el plano ⟂ eje (si la referencia no pasa por el hormigón, CreateFromCurves falla).
    """
    try:
        ax = axis.Normalize()
    except Exception:
        ax = axis
    pt = p0 + ax * float(dist_along)
    try:
        bb = host.get_BoundingBox(None)
        if bb is None:
            return pt
        c_bb = XYZ(
            0.5 * (float(bb.Min.X) + float(bb.Max.X)),
            0.5 * (float(bb.Min.Y) + float(bb.Max.Y)),
            0.5 * (float(bb.Min.Z) + float(bb.Max.Z)),
        )
        v = c_bb - pt
        v = v - ax * float(v.DotProduct(ax))
        return pt + v
    except Exception:
        return pt


def _host_material_warning(host):
    """Aviso si la viga no es hormigón (Revit no hospeda armadura en acero)."""
    try:
        sm = host.StructuralMaterialType
        if sm == StructuralMaterialType.Steel:
            return u"material estructural = Acero (cambia a hormigón o usa familia concreta para Rebar)."
        if sm == StructuralMaterialType.Wood:
            return u"material estructural = Madera (no válido como host de armadura)."
    except Exception:
        pass
    return None


def _pick_first_hook_type(document):
    for ht in FilteredElementCollector(document).OfClass(RebarHookType):
        return ht
    return None


def _read_width_depth_ft(document, elem):
    """Ancho y canto (pies internos) desde tipo o bbox."""
    et = document.GetElement(elem.GetTypeId()) if elem else None
    w, d = None, None
    if et:
        for n in ("Width", "Ancho", "Ancho nominal", "b", "B"):
            p = et.LookupParameter(n)
            if p and p.HasValue:
                w = float(p.AsDouble())
                break
        for n in ("Height", "Depth", "Altura", "Profundidad", "h", "H", "d"):
            p = et.LookupParameter(n)
            if p and p.HasValue:
                d = float(p.AsDouble())
                break
    bb = elem.get_BoundingBox(None)
    if bb is not None:
        dx = abs(bb.Max.X - bb.Min.X)
        dy = abs(bb.Max.Y - bb.Min.Y)
        dz = abs(bb.Max.Z - bb.Min.Z)
        dims = sorted([dx, dy, dz], reverse=True)
        small = sorted(dims[1:]) if len(dims) >= 3 else [dims[-1], dims[-1]]
        bbox_w = float(small[0])
        bbox_d = float(small[1])
        if not w or w <= 0:
            w = bbox_w
        if not d or d <= 0:
            d = bbox_d
    if not w or w <= 0:
        w = 1.0
    if not d or d <= 0:
        d = 1.0
    return w, d


def _beam_frame(curve):
    """Eje unitario, ancho (⟂ eje, horizontal típico), canto (profundidad de sección)."""
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    raw = p1 - p0
    ln = raw.GetLength()
    if ln < 1e-9:
        return None
    axis = raw.Normalize()
    if abs(axis.Z) > 0.92:
        return None
    z_up = XYZ.BasisZ
    width_dir = axis.CrossProduct(z_up)
    if width_dir.GetLength() < 1e-9:
        width_dir = axis.CrossProduct(XYZ.BasisX)
    width_dir = width_dir.Normalize()
    depth_dir = width_dir.CrossProduct(axis).Normalize()
    if depth_dir.Z < 0:
        depth_dir = depth_dir.Negate()
        width_dir = axis.CrossProduct(depth_dir).Normalize()
    return axis, width_dir, depth_dir, p0, p1, ln


def _pick_bar_type(document):
    col = list(FilteredElementCollector(document).OfClass(RebarBarType))
    if not col:
        return None
    if BAR_TYPE_NAME_CONTAINS:
        key = BAR_TYPE_NAME_CONTAINS.lower()
        for bt in col:
            try:
                if key in (bt.Name or "").lower():
                    return bt
            except Exception:
                continue
    return col[0]


def _first_stirrup_distance_ft(beam_len_ft, inset_ft):
    """Distancia desde p0 al primer estribo (al inicio de la zona de reparto)."""
    if beam_len_ft <= 2.0 * inset_ft + 1e-6:
        return 0.5 * beam_len_ft
    return float(inset_ft)


def _stirrup_array_length_ft(beam_len_ft, inset_ft):
    """Longitud del conjunto a lo largo del eje (entre recortes en apoyos)."""
    if beam_len_ft <= 2.0 * inset_ft + 1e-6:
        return 0.0
    return max(0.0, float(beam_len_ft) - 2.0 * float(inset_ft))


def _rebar_quantity(rebar):
    try:
        return int(rebar.Quantity)
    except Exception:
        try:
            return int(rebar.NumberOfBarPositions)
        except Exception:
            return 1


def _apply_maximum_spacing_layout(rebar, spacing_ft, array_len_ft, include_end_bars=True):
    """
    Rebar set con separación máxima (misma idea que armado por máx. spacing en fundación + varias combinaciones).
    La propagación del set con normal = eje de viga va a lo largo del vano.

    Args:
        include_end_bars: si True, ``includeFirstBar`` y ``includeLastBar`` en True; si False,
            ambos en False (estribos centrales BIMTools: sin barra en los extremos del tramo del set).
    """
    acc = rebar.GetShapeDrivenAccessor()
    if acc is None:
        return False
    doc = rebar.Document
    try:
        if array_len_ft < 1e-6:
            acc.SetLayoutAsSingle()
            return True
        if spacing_ft >= array_len_ft - 1e-9:
            acc.SetLayoutAsSingle()
            return True
    except Exception:
        pass

    # Si el vano es mayor que el paso, exigimos más de una posición (evita quedar en Single).
    expect_multi = array_len_ft > spacing_ft * 1.01 + 1e-6

    # API: SetLayoutAsMaximumSpacing(spacing, arrayLen, barsOnNormalSide, includeFirstBar, includeLastBar).
    if include_end_bars:
        combos = (
            (True, True, True),
            (False, True, True),
        )
    else:
        combos = (
            (True, False, False),
            (False, False, False),
        )
    for b_side, inc0, inc1 in combos:
        try:
            acc.SetLayoutAsMaximumSpacing(spacing_ft, array_len_ft, b_side, inc0, inc1)
            try:
                acc.FlipRebarSet()
                acc.SetLayoutAsMaximumSpacing(spacing_ft, array_len_ft, b_side, inc0, inc1)
            except Exception:
                pass
            try:
                doc.Regenerate()
            except Exception:
                pass
            if expect_multi:
                qty = _rebar_quantity(rebar)
                bad_rule = False
                if RebarShapeDrivenLayoutRule is not None:
                    try:
                        lr = acc.LayoutRule
                        bad_rule = lr != RebarShapeDrivenLayoutRule.MaximumSpacing
                    except Exception:
                        bad_rule = False
                if qty < 2 or bad_rule:
                    continue
            return True
        except Exception:
            continue
    # Si la validación qty/LayoutRule fue demasiado estricta (p. ej. Quantity antes de regenerar), reintentar sin filtrar.
    if expect_multi:
        for b_side, inc0, inc1 in combos:
            try:
                acc.SetLayoutAsMaximumSpacing(spacing_ft, array_len_ft, b_side, inc0, inc1)
                try:
                    acc.FlipRebarSet()
                    acc.SetLayoutAsMaximumSpacing(spacing_ft, array_len_ft, b_side, inc0, inc1)
                except Exception:
                    pass
                return True
            except Exception:
                continue
    return False


def _build_rect_corners(center, width_dir, depth_dir, half_w, half_d):
    w = width_dir
    d = depth_dir
    a = center - w * half_w - d * half_d
    b = center + w * half_w - d * half_d
    c = center + w * half_w + d * half_d
    e = center - w * half_w + d * half_d
    return a, b, c, e


def _attempt_create_from_curves(document, host, bar_type, axis_norm, curves_ilist):
    """CreateFromCurves StirrupTie; devuelve (rebar, último_error_texto)."""
    # Revit 2024+ / IronPython: la sobrecarga activa suele pedir RebarHookType o null, no ElementId.
    hook_none = (None, None)
    hook_pick = _pick_first_hook_type(document)
    hook_pairs = [hook_none]
    if hook_pick is not None:
        hook_pairs.append((hook_pick, hook_pick))

    last_err = None
    norms = [axis_norm]
    try:
        norms.append(axis_norm.Negate())
    except Exception:
        pass
    orient_pairs = (
        (RebarHookOrientation.Right, RebarHookOrientation.Right),
        (RebarHookOrientation.Left, RebarHookOrientation.Left),
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
    )
    for h0, h1 in hook_pairs:
        for nvec in norms:
            for so, eo in orient_pairs:
                for use_ex, create_new in ((True, True), (True, False), (False, True), (False, False)):
                    try:
                        r = Rebar.CreateFromCurves(
                            document,
                            RebarStyle.StirrupTie,
                            bar_type,
                            h0,
                            h1,
                            host,
                            nvec,
                            curves_ilist,
                            so,
                            eo,
                            use_ex,
                            create_new,
                        )
                        if r:
                            return r, None
                    except Exception as ex:
                        try:
                            last_err = str(ex)
                        except Exception:
                            last_err = u"(excepción sin mensaje)"
                        continue
    return None, last_err


def _try_create_stirrup_via_shapes(document, host, bar_type, axis_norm, lines):
    """Fallback: RebarShape del proyecto con RebarStyle.StirrupTie."""
    hook = _pick_first_hook_type(document)
    if hook is None:
        return None, u"No hay RebarHookType (necesario para CreateFromCurvesAndShape)."
    curves_clr = _curve_list_object(lines)
    last_err = None
    norms = [axis_norm]
    try:
        norms.append(axis_norm.Negate())
    except Exception:
        pass
    orient_pairs = (
        (RebarHookOrientation.Right, RebarHookOrientation.Right),
        (RebarHookOrientation.Left, RebarHookOrientation.Left),
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
    )
    n_try = 0
    for shape in FilteredElementCollector(document).OfClass(RebarShape):
        try:
            if shape.RebarStyle != RebarStyle.StirrupTie:
                continue
        except Exception:
            continue
        n_try += 1
        if n_try > MAX_STIRRUP_SHAPES_TO_TRY:
            break
        for nvec in norms:
            for so, eo in orient_pairs:
                try:
                    r = Rebar.CreateFromCurvesAndShape(
                        document,
                        shape,
                        bar_type,
                        hook,
                        hook,
                        host,
                        nvec,
                        curves_clr,
                        so,
                        eo,
                    )
                    if r:
                        return r, None
                except Exception as ex:
                    last_err = str(ex)
                    continue
    if n_try == 0:
        return None, (
            last_err
            or u"No hay RebarShape con estilo StirrupTie en el proyecto (o ninguno encajó)."
        )
    return None, last_err


def _try_create_stirrup_all(document, host, bar_type, axis_norm, lines):
    """List[Curve], array CLR y fallback por RebarShape."""
    last_err = None
    for curves in (_curve_list_ilist(lines), _curve_array_from_lines(lines)):
        r, err = _attempt_create_from_curves(document, host, bar_type, axis_norm, curves)
        if r:
            return r, None
        if err:
            last_err = err
    r2, err2 = _try_create_stirrup_via_shapes(document, host, bar_type, axis_norm, lines)
    if r2:
        return r2, None
    return None, err2 or last_err


def _create_one_stirrup_at(
    document, host, bar_type, p0, axis, width_dir, depth_dir, dist_along, half_w, half_d
):
    center = _stirrup_center_at_station(host, p0, axis, dist_along)
    a, b, c, e = _build_rect_corners(center, width_dir, depth_dir, half_w, half_d)
    lines_fwd = [
        Line.CreateBound(a, b),
        Line.CreateBound(b, c),
        Line.CreateBound(c, e),
        Line.CreateBound(e, a),
    ]
    r, err = _try_create_stirrup_all(document, host, bar_type, axis, lines_fwd)
    if r:
        return r, None
    lines_rev = [
        Line.CreateBound(a, e),
        Line.CreateBound(e, c),
        Line.CreateBound(c, b),
        Line.CreateBound(b, a),
    ]
    r2, err2 = _try_create_stirrup_all(document, host, bar_type, axis, lines_rev)
    if r2:
        return r2, None
    return None, err2 or err


def _is_beam(elem):
    if elem is None or not isinstance(elem, FamilyInstance):
        return False
    try:
        if elem.Category is None:
            return False
        if int(elem.Category.Id.IntegerValue) != int(BuiltInCategory.OST_StructuralFraming):
            return False
    except Exception:
        return False
    st = getattr(elem, "StructuralType", None)
    return st == StructuralType.Beam


def run(document, uidocument):
    docu = document
    uid = uidocument
    if docu is None or uid is None:
        print(u"Error: no hay documento activo (__revit__ no definido).")
        return

    ids = list(uid.Selection.GetElementIds())
    if not ids:
        print(u"Error: selecciona al menos una viga (Structural Framing, tipo Beam).")
        return

    bar_type = _pick_bar_type(docu)
    if bar_type is None:
        print(u"Error: no hay RebarBarType en el proyecto.")
        return

    try:
        bar_diam = float(bar_type.BarNominalDiameter)
        if bar_diam <= 0:
            bar_diam = float(getattr(bar_type, "BarModelDiameter", 0) or 0.04)
    except Exception:
        bar_diam = 0.04

    cover_ft = _mm_to_internal(COVER_MM)
    spacing_ft = _mm_to_internal(SPACING_MM)
    inset_ft = _mm_to_internal(END_INSET_MM)

    total = 0
    errores = []

    t = Transaction(docu, u"BIMTools RPS: estribos en vigas")
    t.Start()
    try:
        for eid in ids:
            host = docu.GetElement(eid)
            if not _is_beam(host):
                errores.append(u"ID {}: no es viga (Beam).".format(eid.IntegerValue))
                continue
            loc = host.Location
            if not isinstance(loc, LocationCurve):
                errores.append(u"ID {}: sin LocationCurve.".format(eid.IntegerValue))
                continue
            curve = loc.Curve
            frame = _beam_frame(curve)
            if frame is None:
                errores.append(
                    u"ID {}: geometría no válida o viga demasiado vertical.".format(
                        eid.IntegerValue
                    )
                )
                continue
            axis, width_dir, depth_dir, p0, p1, beam_len = frame
            w_ft, d_ft = _read_width_depth_ft(docu, host)
            half_w = 0.5 * w_ft - cover_ft - 0.5 * bar_diam
            half_d = 0.5 * d_ft - cover_ft - 0.5 * bar_diam
            if half_w <= 1e-4 or half_d <= 1e-4:
                errores.append(
                    u"ID {}: recubrimiento/diámetro dejan luz útil nula (revisa COVER_MM o sección).".format(
                        eid.IntegerValue
                    )
                )
                continue

            mat_w = _host_material_warning(host)
            if mat_w:
                print(u"Aviso ID {}: {}".format(eid.IntegerValue, mat_w))

            dist_first = _first_stirrup_distance_ft(beam_len, inset_ft)
            array_len = _stirrup_array_length_ft(beam_len, inset_ft)

            r, first_api_err = _create_one_stirrup_at(
                docu, host, bar_type, p0, axis, width_dir, depth_dir, dist_first, half_w, half_d
            )
            if not r:
                msg = u"ID {}: no se pudo crear el estribo base.".format(eid.IntegerValue)
                if PRINT_API_ERRORS and first_api_err:
                    msg += u" Detalle: {}".format(first_api_err)
                errores.append(msg)
                continue

            layout_ok = _apply_maximum_spacing_layout(r, spacing_ft, array_len)
            if not layout_ok:
                try:
                    docu.Delete(r.Id)
                except Exception:
                    pass
                errores.append(
                    u"ID {}: estribo creado pero falló SetLayoutAsMaximumSpacing.".format(
                        eid.IntegerValue
                    )
                )
                continue

            qty = _rebar_quantity(r)
            total += 1
            print(
                u"Viga ID {}: 1 conjunto (máx. separación), {} posición(es), s≤{} mm, L_conj≈{} mm, rec {} mm.".format(
                    eid.IntegerValue,
                    qty,
                    int(SPACING_MM),
                    int(round(array_len * 304.8)),
                    int(COVER_MM),
                )
            )
        t.Commit()
    except Exception as ex:
        if t.HasStarted():
            try:
                t.RollBack()
            except Exception:
                pass
        print(u"Error en transacción: {}".format(str(ex)))
        return

    print(u"Total conjuntos de estribos (Rebar): {}".format(total))
    for msg in errores:
        print(msg)


# Solo al ejecutar este .py como script (RPS / consola), no al importar desde
# geometria_estribos_viga u otras herramientas — si no, run() se disparaba al
# cargar el módulo y mostraba el error de «selecciona una viga» en otras tools.
if __name__ == "__main__":
    try:
        if doc is not None and uidoc is not None:
            run(doc, uidoc)
    except NameError:
        pass
