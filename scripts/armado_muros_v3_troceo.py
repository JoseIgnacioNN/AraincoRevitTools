# -*- coding: utf-8 -*-
"""
Troceo por muro/tramo — ciclo Auto → Tramo → Cont. (estilo Armado Muros V2).

- ``auto``: sigue geometría (cambio de espesor o de largo vs inferior).
- ``tramo``: fuerza empalme ON (inicia tramo de barra).
- ``cont``: fuerza continuidad OFF (no empalme).

La evaluación Auto **no** aplica al muro de Z más baja (base del stack);
empieza en el segundo muro apilado hacia arriba.
"""

from __future__ import print_function

TROCEO_AUTO = u"auto"
TROCEO_TRAMO = u"tramo"
TROCEO_CONT = u"cont"

# Instrumentación elev/troceo — poner False para silenciar en pyRevit output.
# Ver ``log_elev_troceo_debug`` / prints ``[ArmadoMachones:troceo]``.
DEBUG_TROCEO = False

_FT_TO_MM = 304.8
_THICK_TOL_FT = 2.0 / _FT_TO_MM
_LEN_TOL_FT = 5.0 / _FT_TO_MM  # ~5 mm en largo / extremo U


def cycle_troceo_mode(mode):
    m = mode or TROCEO_AUTO
    if m == TROCEO_AUTO:
        return TROCEO_TRAMO
    if m == TROCEO_TRAMO:
        return TROCEO_CONT
    return TROCEO_AUTO


def _elem_thickness_ft(elem):
    """Espesor real (Wall.Width), no lado corto de get_machon_dimensions."""
    if elem is None:
        return None
    try:
        w = float(elem.Width)
        if w > 1e-12:
            return w
    except Exception:
        pass
    try:
        from machon_geometry import get_machon_dimensions, is_wall_element

        if is_wall_element(elem):
            return float(get_machon_dimensions(elem)[0])
    except Exception:
        pass
    return None


def _elem_length_ft(elem):
    if elem is None:
        return None
    try:
        from machon_geometry import _wall_length_ft, is_wall_element

        if is_wall_element(elem):
            lu = _wall_length_ft(elem)
            if lu is not None and float(lu) > 1e-12:
                return float(lu)
    except Exception:
        pass
    try:
        loc = elem.Location
        crv = getattr(loc, u"Curve", None)
        if crv is not None:
            return float(crv.Length)
    except Exception:
        pass
    return None


def auto_geom_suggests_empalme(elem, elem_below):
    """True si el muro sugiere troceo vs el inferior (espesor o largo)."""
    if elem is None or elem_below is None:
        return False
    t0 = _elem_thickness_ft(elem_below)
    t1 = _elem_thickness_ft(elem)
    if t0 is not None and t1 is not None:
        if abs(float(t1) - float(t0)) > _THICK_TOL_FT:
            return True
    l0 = _elem_length_ft(elem_below)
    l1 = _elem_length_ft(elem)
    if l0 is not None and l1 is not None:
        if abs(float(l1) - float(l0)) > _LEN_TOL_FT:
            return True
    return False


def auto_meta_suggests_empalme(m, m_below):
    """Igual que ``auto_geom_suggests_empalme`` usando campos de elev meta."""
    if not m or not m_below:
        return False
    try:
        t0 = float(m_below.get(u"thick_mm") or 0.0)
        t1 = float(m.get(u"thick_mm") or 0.0)
    except Exception:
        t0 = t1 = 0.0
    if t0 > 1.0 and t1 > 1.0 and abs(t1 - t0) > 2.0:
        return True
    try:
        l0 = float(m_below.get(u"length_ft") or m_below.get(u"length_u") or 0.0)
        l1 = float(m.get(u"length_ft") or m.get(u"length_u") or 0.0)
    except Exception:
        l0 = l1 = 0.0
    if l0 > 1e-9 and l1 > 1e-9 and abs(l1 - l0) > _LEN_TOL_FT:
        return True
    # Respaldo: elementos Revit si meta incompleta.
    return auto_geom_suggests_empalme(m.get(u"elem"), m_below.get(u"elem"))


def lowest_z_meta_index(wall_meta):
    """Índice del muro con Z más baja (base del stack). ``None`` si vacío."""
    meta = list(wall_meta or [])
    if not meta:
        return None

    def _z_at(i):
        try:
            return float(meta[i].get(u"z_mm") or 0.0)
        except Exception:
            return 0.0

    return min(range(len(meta)), key=lambda i: (_z_at(i), i))


def compute_auto_troceo_flags(walls_ordered):
    """
    Flags Auto por muro apilado (orden base → cima = Z ascendente).

    El muro de Z más baja (índice 0) queda **fuera** de la evaluación;
    la geometría se compara desde el **segundo** muro hacia arriba
    (cada uno vs el inmediatamente inferior), como Muros V2.
    """
    n = len(walls_ordered or [])
    flags = [False] * n
    if n < 2:
        return flags
    for i in range(1, n):
        try:
            flags[i] = bool(
                auto_geom_suggests_empalme(
                    walls_ordered[i],
                    walls_ordered[i - 1],
                )
            )
        except Exception:
            flags[i] = False
    return flags


def compute_auto_troceo_flags_from_meta(wall_meta):
    """
    Igual que ``compute_auto_troceo_flags``, usando meta de elevación.

    Identifica el muro base por **mínimo ``z_mm``** (no solo índice 0) y lo
    excluye; evalúa el resto en orden de Z ascendente vs el muro inferior
    **solo si hay contacto de apilamiento en Z** (no encadenar muros
    desconectados solo por orden de cota — evitaba falsos tramos / fantasma).
    """
    meta = list(wall_meta or [])
    n = len(meta)
    flags = [False] * n
    if n < 2:
        return flags

    def _z_at(i):
        try:
            return float(meta[i].get(u"z_mm") or 0.0)
        except Exception:
            return 0.0

    def _h_at(i):
        try:
            return float(meta[i].get(u"height_mm") or 0.0)
        except Exception:
            return 0.0

    def _stacked_contact(i_upper, i_lower):
        """Tope del inferior ≈ base del superior (mm), tol ~15 mm."""
        z_bot_u = _z_at(i_upper)
        z_top_l = _z_at(i_lower) + _h_at(i_lower)
        return abs(z_top_l - z_bot_u) <= 15.0

    by_z = sorted(range(n), key=lambda i: (_z_at(i), i))
    # by_z[0] = muro de elevación Z más baja → sin evaluación Auto
    for rank in range(1, n):
        i = by_z[rank]
        below_i = by_z[rank - 1]
        if not _stacked_contact(i, below_i):
            flags[i] = False
            continue
        try:
            flags[i] = bool(auto_meta_suggests_empalme(meta[i], meta[below_i]))
        except Exception:
            flags[i] = False
        if DEBUG_TROCEO and flags[i]:
            try:
                m = meta[i]
                b = meta[below_i]
                print(
                    u"[ArmadoMachones:troceo] auto empalme W{0}←W{1} "
                    u"thick {2:.0f}→{3:.0f} mm  len {4:.0f}→{5:.0f} mm".format(
                        i,
                        below_i,
                        float(b.get(u"thick_mm") or 0.0),
                        float(m.get(u"thick_mm") or 0.0),
                        float(b.get(u"length_ft") or b.get(u"length_u") or 0.0)
                        * _FT_TO_MM,
                        float(m.get(u"length_ft") or m.get(u"length_u") or 0.0)
                        * _FT_TO_MM,
                    )
                )
            except Exception:
                pass
    return flags


def effective_empalme(mode, auto_suggest):
    """Bool efectivo de empalme/troceo según modo."""
    m = mode or TROCEO_AUTO
    if m == TROCEO_TRAMO:
        return True
    if m == TROCEO_CONT:
        return False
    return bool(auto_suggest)


def pie_caption(mode, auto_suggest):
    m = mode or TROCEO_AUTO
    if m == TROCEO_TRAMO:
        return u"Tramo"
    if m == TROCEO_CONT:
        return u"Cont."
    return u"Auto·" if auto_suggest else u"Auto"


def empalme_indices_from_modes(modes, auto_flags, base_index=None):
    """
    Índices (0-based) con empalme efectivo.

    El muro base (Z más baja / ``base_index``, por defecto 0) nunca entra.
    """
    out = []
    n = max(len(modes or []), len(auto_flags or []))
    try:
        base_i = 0 if base_index is None else int(base_index)
    except Exception:
        base_i = 0
    for i in range(n):
        if i == base_i:
            continue
        mode = TROCEO_AUTO
        if modes and i < len(modes):
            mode = modes[i]
        auto = False
        if auto_flags and i < len(auto_flags):
            auto = bool(auto_flags[i])
        if effective_empalme(mode, auto):
            out.append(i)
    return out


def log_elev_troceo_debug(wall_meta, autos=None, modes=None):
    """Una línea por muro + base min-Z. No-op si ``DEBUG_TROCEO`` es False."""
    if not DEBUG_TROCEO:
        return
    meta = list(wall_meta or [])
    n = len(meta)
    base_i = lowest_z_meta_index(meta)
    try:
        print(
            u"[ArmadoMachones:troceo] elev n={0} base_i={1} (min z_mm)".format(
                n, base_i
            )
        )
    except Exception:
        pass
    if autos is None:
        try:
            autos = compute_auto_troceo_flags_from_meta(meta)
        except Exception:
            autos = [False] * n
    for i, m in enumerate(meta):
        try:
            eid = m.get(u"eid")
            z_mm = float(m.get(u"z_mm") or 0.0)
            h_mm = float(m.get(u"height_mm") or 0.0)
            th = float(m.get(u"thick_mm") or 0.0)
            lu = float(m.get(u"length_ft") or m.get(u"length_u") or 0.0)
            af = bool(autos[i]) if autos and i < len(autos) else False
            md = TROCEO_AUTO
            if modes and i < len(modes):
                md = modes[i] or TROCEO_AUTO
            mark = u" BASE" if i == base_i else u""
            print(
                u"[ArmadoMachones:troceo] W{0} id={1} z={2:.1f} h={3:.1f} "
                u"e={4:.0f} L={5:.0f} auto={6} mode={7}{8}".format(
                    i,
                    eid,
                    z_mm,
                    h_mm,
                    th,
                    lu * _FT_TO_MM,
                    af,
                    md,
                    mark,
                )
            )
        except Exception as ex:
            try:
                print(u"[ArmadoMachones:troceo] W{0} log-err {1}".format(i, ex))
            except Exception:
                pass
