# -*- coding: utf-8 -*-
"""
Divide barras de armadura (Rebar) seleccionadas en tramos de como máximo un largo dado (mm).

Enfoque: lee la línea media por posición de barra, parte la polilínea en segmentos de longitud
objetivo y recrea cada tramo con Rebar.CreateFromCurves. Elimina el elemento original.

Limitaciones:
- Rebar con host válido y geometría basada en curvas soportadas (líneas y arcos acotados).
- Conjuntos con varias posiciones: se procesa cada posición y se borra un solo elemento al final.
- Ganchos: el primer tramo conserva el gancho de inicio del original; el último tramo de cada
  posición conserva el gancho de fin; los tramos intermedios van sin ganchos.
"""

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import Arc, ElementId, Line, Transaction, UnitUtils, UnitTypeId, XYZ
from Autodesk.Revit.DB.Structure import (
    MultiplanarOption,
    Rebar,
    RebarBarType,
    RebarHookOrientation,
    RebarHookType,
    RebarStyle,
)
from Autodesk.Revit.UI import TaskDialog

import System


def _curve_clr_type():
    """Tipo CLR abstracto Curve (Line.BaseType) para reflexión y genéricos."""
    return clr.GetClrType(Line).BaseType


def _line_clr_type():
    return clr.GetClrType(Line)


def _arc_clr_type():
    return clr.GetClrType(Arc)


def _trim_curve_segment(crv, pa, pb):
    """
    Subcurva entre parámetros raw pa y pb (misma convención que CreateTrimmedCurve).
    IronPython a menudo no expone Curve.CreateTrimmedCurve; usamos geometría + reflexión.
    """
    if crv is None:
        raise ValueError(u"Curva nula.")
    if abs(float(pa) - float(pb)) < 1e-12:
        raise ValueError(u"Parámetros de recorte coinciden.")

    tcrv = clr.GetClrType(crv)
    tname = tcrv.Name
    fn = tcrv.FullName or u""

    is_line = _line_clr_type().Equals(tcrv) or _line_clr_type().IsAssignableFrom(tcrv)
    if not is_line and (tname == "Line" or fn.endswith(".Line")):
        is_line = True

    if is_line:
        return Line.CreateBound(crv.Evaluate(pa, False), crv.Evaluate(pb, False))

    is_arc = _arc_clr_type().Equals(tcrv) or _arc_clr_type().IsAssignableFrom(tcrv)
    if not is_arc and (tname == "Arc" or fn.endswith(".Arc")):
        is_arc = True

    if is_arc:
        tmid = 0.5 * (float(pa) + float(pb))
        return Arc.Create(
            crv.Evaluate(pa, False),
            crv.Evaluate(tmid, False),
            crv.Evaluate(pb, False),
        )

    raise ValueError(
        u"Tipo de curva no soportado para recorte: {} (solo línea y arco).".format(tname)
    )


def _segment_length_between(crv, pa, pb):
    seg = _trim_curve_segment(crv, pa, pb)
    return float(seg.Length)


def _mm_to_internal(mm):
    return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)


def _npos(rebar):
    try:
        return int(rebar.NumberOfBarPositions)
    except Exception:
        return int(rebar.Quantity)


def _hook_type(doc, hook_id):
    if hook_id is None or hook_id == ElementId.InvalidElementId:
        return None
    e = doc.GetElement(hook_id)
    return e if isinstance(e, RebarHookType) else None


def _rebar_normal(rebar):
    try:
        acc = rebar.GetShapeDrivenAccessor()
        if acc is not None:
            n = acc.Normal
            if n is not None and n.GetLength() > 1e-12:
                return n.Normalize()
    except Exception:
        pass
    return XYZ.BasisZ


def _param_at_dist_from_start(curve, dist):
    """Distancia de arco desde GetEndParameter(0) hasta el parámetro devuelto."""
    p0 = curve.GetEndParameter(0)
    p1 = curve.GetEndParameter(1)
    if dist <= 1e-12:
        return p0
    if dist >= curve.Length - 1e-9:
        return p1
    prev_p = p0
    for k in range(1, 33):
        t = float(k) / 32.0
        pm = p0 + t * (p1 - p0)
        try:
            Lm = _segment_length_between(curve, p0, pm)
        except Exception:
            continue
        if Lm >= dist:
            lo_p, hi_p = prev_p, pm
            for _ in range(45):
                mid = 0.5 * (lo_p + hi_p)
                try:
                    Lmid = _segment_length_between(curve, p0, mid)
                except Exception:
                    hi_p = mid
                    continue
                if Lmid < dist:
                    lo_p = mid
                else:
                    hi_p = mid
            return 0.5 * (lo_p + hi_p)
        prev_p = pm
    return p1


def _split_curve_chain_into_chunks(curves, chunk_len_internal, max_chunks=800):
    """
    curves: IList o lista de Curve conectadas en orden (salida de GetCenterlineCurves).
    Devuelve lista de listas de Curve (cada una es un tramo continuo).
    """
    if not curves or chunk_len_internal <= 1e-12:
        return []
    clist = [c for c in curves if c is not None]
    if not clist:
        return []

    total = sum(float(c.Length) for c in clist)
    if total <= chunk_len_internal + 1e-9:
        return [clist]

    chunks = []
    current = []
    accum = 0.0
    idx = 0
    offset_on_curve = 0.0
    eps = 1e-9

    while idx < len(clist):
        if len(chunks) + (1 if current else 0) > max_chunks:
            raise ValueError(
                u"Demasiados tramos (> {}). Aumenta el largo objetivo o divide la selección.".format(max_chunks)
            )

        c = clist[idx]
        c_len = float(c.Length)
        rem = c_len - offset_on_curve
        need = chunk_len_internal - accum

        if rem <= eps:
            idx += 1
            offset_on_curve = 0.0
            continue

        p0 = c.GetEndParameter(0)
        p1 = c.GetEndParameter(1)

        if rem > need + eps:
            ps = _param_at_dist_from_start(c, offset_on_curve)
            pe = _param_at_dist_from_start(c, offset_on_curve + need)
            try:
                piece = _trim_curve_segment(c, ps, pe)
            except Exception as ex:
                raise ValueError(u"No se pudo trocear una curva del trazado: {}".format(ex))

            current.append(piece)
            chunks.append(current)
            current = []
            accum = 0.0
            offset_on_curve += need
            if offset_on_curve >= c_len - eps:
                idx += 1
                offset_on_curve = 0.0
        else:
            ps = _param_at_dist_from_start(c, offset_on_curve)
            try:
                piece = _trim_curve_segment(c, ps, p1)
            except Exception as ex:
                raise ValueError(u"No se pudo extraer el final de una curva: {}".format(ex))
            current.append(piece)
            accum += float(piece.Length)
            idx += 1
            offset_on_curve = 0.0
            if accum >= chunk_len_internal - eps:
                chunks.append(current)
                current = []
                accum = 0.0

    if current:
        chunks.append(current)

    return chunks


def _create_rebar_chunk(
    doc,
    curves_list,
    host,
    norm,
    bar_type,
    style,
    start_hook,
    end_hook,
    start_orient,
    end_orient,
):
    # IronPython: evitar List[T](python_list) y genéricos raros; usar Array CLR de Curve.
    ct = _curve_clr_type()
    n = len(curves_list)
    arr = System.Array.CreateInstance(ct, n)
    for i in range(n):
        arr[i] = curves_list[i]
    return Rebar.CreateFromCurves(
        doc,
        style,
        bar_type,
        start_hook,
        end_hook,
        host,
        norm,
        arr,
        start_orient,
        end_orient,
        True,
        True,
    )


def dividir_un_rebar(document, rebar, largo_segmento_mm, max_chunks=800):
    """
    Divide un Rebar en tramos de longitud máxima largo_segmento_mm.
    Returns:
        (ok: bool, creados: int, mensaje: unicode)
    """
    if not isinstance(rebar, Rebar):
        return False, 0, u"No es un elemento Rebar."

    if largo_segmento_mm <= 0:
        return False, 0, u"El largo debe ser mayor que cero."

    chunk_len = _mm_to_internal(largo_segmento_mm)
    host = document.GetElement(rebar.GetHostId())
    if host is None:
        return False, 0, u"La barra no tiene host válido."

    bar_type = document.GetElement(rebar.GetTypeId())
    if not isinstance(bar_type, RebarBarType):
        return False, 0, u"No se pudo obtener RebarBarType."

    try:
        style = rebar.Style
    except Exception:
        style = RebarStyle.Standard

    norm = _rebar_normal(rebar)
    h0_id = rebar.GetHookTypeId(0)
    h1_id = rebar.GetHookTypeId(1)
    hook_start_full = _hook_type(document, h0_id)
    hook_end_full = _hook_type(document, h1_id)

    try:
        so0 = rebar.GetHookOrientation(0)
        so1 = rebar.GetHookOrientation(1)
    except Exception:
        so0 = RebarHookOrientation.Right
        so1 = RebarHookOrientation.Left

    n_positions = _npos(rebar)
    created = 0

    t = Transaction(document, u"Dividir armadura por longitud")
    t.Start()
    try:
        for pi in range(n_positions):
            curves = rebar.GetCenterlineCurves(
                False, False, False, MultiplanarOption.IncludeAllMultiplanarCurves, pi
            )
            if curves is None or curves.Count == 0:
                continue

            chain = [curves[i] for i in range(curves.Count)]
            chunks = _split_curve_chain_into_chunks(chain, chunk_len, max_chunks=max_chunks)
            nch = len(chunks)
            for ci, chunk_curves in enumerate(chunks):
                if not chunk_curves:
                    continue
                sh = hook_start_full if (ci == 0) else None
                eh = hook_end_full if (ci == nch - 1) else None
                nr = _create_rebar_chunk(
                    document,
                    chunk_curves,
                    host,
                    norm,
                    bar_type,
                    style,
                    sh,
                    eh,
                    so0 if sh is not None else RebarHookOrientation.Right,
                    so1 if eh is not None else RebarHookOrientation.Left,
                )
                if nr is None:
                    raise RuntimeError(u"CreateFromCurves devolvió None (tramo {}/{})".format(ci + 1, nch))
                created += 1

        if created == 0:
            t.RollBack()
            return False, 0, u"No se generó ningún tramo (revisa la geometría de la barra)."

        document.Delete(rebar.Id)
        t.Commit()
    except Exception as ex:
        t.RollBack()
        try:
            msg = unicode(ex) if ex else u"Error en transacción."
        except NameError:
            msg = str(ex) if ex else u"Error en transacción."
        return False, 0, msg

    return True, created, u""


def dividir_seleccion(uidocument, largo_segmento_mm, max_chunks=800):
    """
    Procesa la selección actual del UIDocument.
    Returns:
        (total_ok, total_creados, lineas_log)
    """
    doc = uidocument.Document
    sel = uidocument.Selection.GetElementIds()
    if sel is None or sel.Count == 0:
        return 0, 0, [u"No hay elementos seleccionados."]

    logs = []
    total_creados = 0
    ok_count = 0

    for eid in sel:
        el = doc.GetElement(eid)
        if not isinstance(el, Rebar):
            logs.append(u"{}: omitido (no es Rebar).".format(eid.IntegerValue))
            continue
        ok, n, msg = dividir_un_rebar(doc, el, largo_segmento_mm, max_chunks=max_chunks)
        if ok:
            ok_count += 1
            total_creados += n
            logs.append(u"{}: {} tramo(s) creado(s).".format(eid.IntegerValue, n))
        else:
            logs.append(u"{}: {}".format(eid.IntegerValue, msg))

    return ok_count, total_creados, logs


def run_pyrevit(__revit__, largo_segmento_mm=None, max_chunks=800):
    """
    Entrada para botón pyRevit. Si largo_segmento_mm es None, pide valor por consola InputBox.
    """
    uidoc = __revit__.ActiveUIDocument
    if largo_segmento_mm is None:
        try:
            clr.AddReference("Microsoft.VisualBasic")
            from Microsoft.VisualBasic import Interaction

            s = Interaction.InputBox(
                u"Largo máximo de cada tramo (mm):",
                u"Dividir armadura por longitud",
                u"6000",
            )
        except Exception:
            s = None
        if not s or not str(s).strip():
            TaskDialog.Show(u"Dividir armadura", u"Operación cancelada.")
            return
        try:
            largo_segmento_mm = float(str(s).strip().replace(",", "."))
        except Exception:
            TaskDialog.Show(u"Dividir armadura", u"No se pudo interpretar el número.")
            return

    okc, total, logs = dividir_seleccion(uidoc, largo_segmento_mm, max_chunks=max_chunks)
    summary = u"Elementos divididos: {}. Tramos nuevos: {}.\n\n{}".format(
        okc, total, u"\n".join(logs[:40])
    )
    if len(logs) > 40:
        summary += u"\n..."
    TaskDialog.Show(u"Dividir armadura por longitud", summary)


def main_rps():
    """RevitPythonShell: ajusta LARGO_MM y ejecuta con selección actual."""
    LARGO_MM = 6000.0
    MAX_CHUNKS = 800
    try:
        uidoc = __revit__.ActiveUIDocument  # noqa: F821
    except Exception:
        TaskDialog.Show(u"Dividir armadura", u"Ejecuta en RPS con __revit__ disponible.")
        return
    okc, total, logs = dividir_seleccion(uidocument, LARGO_MM, max_chunks=MAX_CHUNKS)
    for line in logs:
        print(line)
    print(u"Listo: {} elemento(s), {} tramo(s).".format(okc, total))
