# -*- coding: utf-8 -*-
"""Aplicación de extremos empotrado/gancho sobre líneas fusionadas."""

from geometria_empotramiento_extremos import (
    DIAM_NOMINAL_RESPALDO_MM,
    MODO_EMPOTRAMIENTO,
    MODO_GANCHO,
    RECUBRIMIENTO_FIBRA_MM,
    SONDA_COLISION_MM,
    aplicar_extremos_linea,
    element_ids_desde_elementos,
    resolver_extremo_linea,
    _hook_mm_desde_diametro,
    _mm_to_ft,
    _MIN_LINE_LEN_FT,
)

from armado_vigas.domain.bar_ends import empotramiento_mm_for_diam
from armado_vigas.domain.constants import (
    BAR_END_MODE_AUTO,
    BAR_END_MODE_EMP,
    BAR_END_MODE_PATA_L,
    normalize_bar_end_mode,
)

try:
    from Autodesk.Revit.DB import Line
except Exception:
    Line = None

__all__ = [
    "MODO_EMPOTRAMIENTO",
    "MODO_GANCHO",
    "aplicar_extremos_a_linea_fusionada",
    "element_ids_desde_elementos",
    "resolver_extremo_linea",
    "force_end_mode_meta",
    "mark_pata_l_keep_geometry",
    "aplicar_empotramiento_extremos_marcados",
]


def mark_pata_l_keep_geometry(meta, diam_nominal_mm=None, punto=None, concrete_grade=None):
    """
    Marca pata L en meta **sin mover** el extremo (conserva estirón por viga).

    Usado tras colisión con viga transversal: la longitud ya incluye
    ``+(b/2 − 25 mm)``; solo se añade gancho L en colocación.
    """
    try:
        d_mm = float(diam_nominal_mm or DIAM_NOMINAL_RESPALDO_MM)
    except Exception:
        d_mm = DIAM_NOMINAL_RESPALDO_MM
    if d_mm <= 1e-9:
        d_mm = DIAM_NOMINAL_RESPALDO_MM
    g = concrete_grade
    try:
        from armado_vigas.domain.concrete_lengths import (
            hook_mm_for_diameter,
            session_concrete_grade,
        )

        g = concrete_grade if concrete_grade is not None else session_concrete_grade()
        hook_mm = float(hook_mm_for_diameter(d_mm, g) or 0.0)
    except Exception:
        try:
            hook_mm = float(_hook_mm_desde_diametro(d_mm, g) or 0.0)
        except Exception:
            hook_mm = 0.0
    out = {}
    if meta:
        try:
            out.update(meta)
        except Exception:
            pass
    if punto is not None:
        out[u"punto"] = punto
    elif meta and meta.get(u"punto") is not None:
        out[u"punto"] = meta.get(u"punto")
    out[u"modo"] = MODO_GANCHO
    out[u"pata_l"] = True
    out[u"hook_mm"] = float(hook_mm)
    out[u"forced"] = True
    out[u"beam_collision"] = True
    # Geometría longitudinal ya fijada por estirón; no retraer aquí.
    out[u"delta_mm"] = float(out.get(u"delta_mm") or 0.0)
    out[u"descripcion"] = (
        u"Colisión muro/viga: pata L {0:.0f} mm (estirón conservado)".format(
            float(hook_mm)
        )
    )
    return out


def aplicar_empotramiento_extremos_marcados(
    line, emp_meta, diam_nominal_mm=None, concrete_grade=None
):
    """
    Estira extremos marcados por colisión con muro // a la vista.

    Longitud = empotramiento de tabla según Ø y dosificación.
    Conserva pata L / otros extremos no marcados.

    Returns:
        ``(line_out, meta_start, meta_end)`` — metas solo en extremos aplicados;
        el otro extremo queda ``None`` (el caller fusiona con metas previas).
    """
    meta_i = None
    meta_f = None
    if line is None or not emp_meta or not emp_meta.get(u"applied"):
        return line, meta_i, meta_f
    try:
        d_mm = float(diam_nominal_mm or DIAM_NOMINAL_RESPALDO_MM)
    except Exception:
        d_mm = DIAM_NOMINAL_RESPALDO_MM
    if d_mm <= 1e-9:
        d_mm = DIAM_NOMINAL_RESPALDO_MM

    emp_mm, desc = empotramiento_mm_for_diam(d_mm, concrete_grade=concrete_grade)
    try:
        emp_mm = float(emp_mm or 0.0)
    except Exception:
        emp_mm = 0.0
    if emp_mm <= 1e-6:
        return line, meta_i, meta_f

    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        v = p1 - p0
        L = float(v.GetLength())
        u = v.Normalize()
    except Exception:
        return line, meta_i, meta_f
    if L < float(_MIN_LINE_LEN_FT):
        return line, meta_i, meta_f

    d_ft = _mm_to_ft(emp_mm)
    changed = False
    if emp_meta.get(u"start"):
        try:
            p0 = p0 - u.Multiply(d_ft)
            changed = True
        except Exception:
            pass
        wall_id = None
        try:
            wall_id = (emp_meta.get(u"start") or {}).get(u"wall_id")
        except Exception:
            wall_id = None
        meta_i = {
            u"punto": p0,
            u"modo": MODO_EMPOTRAMIENTO,
            u"delta_mm": float(emp_mm),
            u"hook_mm": 0.0,
            u"emp_mm": float(emp_mm),
            u"sonda_mm": 0.0,
            u"pata_l": False,
            u"forced": True,
            u"wall_parallel": True,
            u"wall_id": wall_id,
            u"descripcion": u"Muro // vista: {0}".format(
                desc or u"Empotramiento Ø{0} → {1:.0f} mm".format(
                    int(round(d_mm)), float(emp_mm)
                )
            ),
        }
    if emp_meta.get(u"end"):
        try:
            p1 = p1 + u.Multiply(d_ft)
            changed = True
        except Exception:
            pass
        wall_id = None
        try:
            wall_id = (emp_meta.get(u"end") or {}).get(u"wall_id")
        except Exception:
            wall_id = None
        meta_f = {
            u"punto": p1,
            u"modo": MODO_EMPOTRAMIENTO,
            u"delta_mm": float(emp_mm),
            u"hook_mm": 0.0,
            u"emp_mm": float(emp_mm),
            u"sonda_mm": 0.0,
            u"pata_l": False,
            u"forced": True,
            u"wall_parallel": True,
            u"wall_id": wall_id,
            u"descripcion": u"Muro // vista: {0}".format(
                desc or u"Empotramiento Ø{0} → {1:.0f} mm".format(
                    int(round(d_mm)), float(emp_mm)
                )
            ),
        }

    if not changed:
        return line, meta_i, meta_f
    try:
        if float((p1 - p0).GetLength()) < float(_MIN_LINE_LEN_FT):
            return line, meta_i, meta_f
        if Line is not None:
            return Line.CreateBound(p0, p1), meta_i, meta_f
    except Exception:
        pass
    return line, meta_i, meta_f


def force_end_mode_meta(
    meta, mode, p_ext, dir_saliente_unit, diam_nominal_mm=None, concrete_grade=None
):
    """
    Sobrescribe el resultado de la sonda con un modo de usuario.

    ``mode``: ``auto`` (sin cambio) | ``emp`` | ``pata_l``.
    """
    mode = normalize_bar_end_mode(mode)
    if mode == BAR_END_MODE_AUTO or mode is None:
        return meta

    try:
        d_mm = float(diam_nominal_mm or DIAM_NOMINAL_RESPALDO_MM)
    except Exception:
        d_mm = DIAM_NOMINAL_RESPALDO_MM
    if d_mm <= 1e-9:
        d_mm = DIAM_NOMINAL_RESPALDO_MM

    mm_s = float(SONDA_COLISION_MM)
    mm_rec = float(RECUBRIMIENTO_FIBRA_MM)
    try:
        du = dir_saliente_unit.Normalize() if dir_saliente_unit is not None else None
    except Exception:
        du = dir_saliente_unit

    if mode == BAR_END_MODE_EMP:
        # Estirón = sonda 50 mm + longitud de desarrollo/empotramiento por Ø + grade.
        emp_mm, desc = empotramiento_mm_for_diam(d_mm, concrete_grade=concrete_grade)
        delta = mm_s + float(emp_mm or 0.0)
        p_nuevo = p_ext
        if p_ext is not None and du is not None:
            try:
                p_nuevo = p_ext + du.Multiply(_mm_to_ft(delta))
            except Exception:
                p_nuevo = p_ext
        return {
            u"punto": p_nuevo,
            u"modo": MODO_EMPOTRAMIENTO,
            u"delta_mm": float(delta),
            u"hook_mm": 0.0,
            u"emp_mm": float(emp_mm or 0.0),
            u"sonda_mm": mm_s,
            u"pata_l": False,
            u"forced": True,
            u"descripcion": u"Forzado: {0}".format(
                desc or u"Empotramiento Ø{0} mm → {1:.0f} mm".format(
                    int(round(d_mm)), float(emp_mm or 0.0)
                )
            ),
        }

    # Pata L (gancho libre)
    retract_mm = mm_rec + 0.5 * d_mm
    delta = -retract_mm
    p_nuevo = p_ext
    if p_ext is not None and du is not None:
        try:
            p_nuevo = p_ext + du.Multiply(_mm_to_ft(delta))
        except Exception:
            p_nuevo = p_ext
    g = concrete_grade
    try:
        from armado_vigas.domain.concrete_lengths import (
            hook_mm_for_diameter,
            session_concrete_grade,
        )

        g = concrete_grade if concrete_grade is not None else session_concrete_grade()
        hook_mm = float(hook_mm_for_diameter(d_mm, g) or 0.0)
    except Exception:
        try:
            hook_mm = float(_hook_mm_desde_diametro(d_mm, g) or 0.0)
        except Exception:
            hook_mm = 0.0
    return {
        u"punto": p_nuevo,
        u"modo": MODO_GANCHO,
        u"delta_mm": float(delta),
        u"hook_mm": float(hook_mm or 0.0),
        u"emp_mm": 0.0,
        u"sonda_mm": mm_s,
        u"pata_l": True,
        u"forced": True,
        u"descripcion": u"Forzado: pata L {0:.0f} mm".format(float(hook_mm or 0.0)),
    }


def aplicar_extremos_a_linea_fusionada(
    document,
    line,
    ids_seleccion,
    host_chain_elements,
    diam_nominal_mm,
    resolver_inicio=True,
    resolver_fin=True,
    end_mode_start=None,
    end_mode_end=None,
    concrete_grade=None,
):
    """
    Wrapper de :func:`geometria_empotramiento_extremos.aplicar_extremos_linea`.

    ``host_chain_elements``: vigas de la cadena colineal (excluidas de colisión).

    ``end_mode_start`` / ``end_mode_end``: preferencia de usuario sobre extremos de
    **curva** (0 / 1). Usar ``auto`` o ``None`` para dejar la sonda.
    """
    try:
        from armado_vigas.domain.concrete_lengths import session_concrete_grade

        grade = (
            concrete_grade
            if concrete_grade is not None
            else session_concrete_grade()
        )
    except Exception:
        grade = concrete_grade

    ids_excluir = element_ids_desde_elementos(host_chain_elements)
    try:
        line_out, meta_i, meta_f = aplicar_extremos_linea(
            document,
            line,
            ids_seleccion,
            ids_excluir=ids_excluir,
            diam_nominal_mm=diam_nominal_mm,
            resolver_inicio=resolver_inicio,
            resolver_fin=resolver_fin,
            concrete_grade=grade,
        )
    except TypeError:
        # Módulo legacy sin ``concrete_grade``.
        line_out, meta_i, meta_f = aplicar_extremos_linea(
            document,
            line,
            ids_seleccion,
            ids_excluir=ids_excluir,
            diam_nominal_mm=diam_nominal_mm,
            resolver_inicio=resolver_inicio,
            resolver_fin=resolver_fin,
        )

    if line is None:
        return line_out, meta_i, meta_f

    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        t = (p1 - p0).Normalize()
    except Exception:
        return line_out, meta_i, meta_f

    mode_s = normalize_bar_end_mode(end_mode_start) if end_mode_start is not None else BAR_END_MODE_AUTO
    mode_e = normalize_bar_end_mode(end_mode_end) if end_mode_end is not None else BAR_END_MODE_AUTO

    pa = p0
    pb = p1
    if resolver_inicio and mode_s != BAR_END_MODE_AUTO:
        meta_i = force_end_mode_meta(
            meta_i, mode_s, p0, t.Negate(), diam_nominal_mm, concrete_grade=grade
        )
        if meta_i and meta_i.get(u"punto") is not None:
            pa = meta_i[u"punto"]
        elif line_out is not None:
            try:
                pa = line_out.GetEndPoint(0)
            except Exception:
                pa = p0
    elif line_out is not None:
        try:
            pa = line_out.GetEndPoint(0)
        except Exception:
            pass
        if meta_i and meta_i.get(u"punto") is not None:
            pa = meta_i[u"punto"]

    if resolver_fin and mode_e != BAR_END_MODE_AUTO:
        meta_f = force_end_mode_meta(
            meta_f, mode_e, p1, t, diam_nominal_mm, concrete_grade=grade
        )
        if meta_f and meta_f.get(u"punto") is not None:
            pb = meta_f[u"punto"]
        elif line_out is not None:
            try:
                pb = line_out.GetEndPoint(1)
            except Exception:
                pb = p1
    elif line_out is not None:
        try:
            pb = line_out.GetEndPoint(1)
        except Exception:
            pass
        if meta_f and meta_f.get(u"punto") is not None:
            pb = meta_f[u"punto"]

    # Recalcular línea si hubo forzado en extremos de curva.
    if (
        (resolver_inicio and mode_s != BAR_END_MODE_AUTO)
        or (resolver_fin and mode_e != BAR_END_MODE_AUTO)
    ):
        try:
            if pa.DistanceTo(pb) < _MIN_LINE_LEN_FT:
                return None, meta_i, meta_f
            if Line is not None:
                return Line.CreateBound(pa, pb), meta_i, meta_f
        except Exception:
            return line_out, meta_i, meta_f

    return line_out, meta_i, meta_f
