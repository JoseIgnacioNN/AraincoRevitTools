# -*- coding: utf-8 -*-
"""
Extremos de armadura longitudinal — sonda de colisión y resolución empotrado vs gancho.

Tras fusión colineal de la guía:

1. Desde cada nodo extremo, sonda **50 mm** hacia fuera (tangente saliente).
2. Colisión booleana contra **todo** el lote seleccionado (vigas, columnas, muros),
   excluyendo ids host indicados (p. ej. cadena de la viga).
3. **Con colisión (no columna):** +50 mm + longitud de empotramiento (tabla
   ``EXTENSION_MM_BY_BAR_DIAMETER_MM``).
4. **Con colisión en columna estructural:** igual estirón de empotramiento; sobre el tramo
   estirado se recorta al volumen/cara del pilar y se aplica **−(25 + Ø/2)** mm desde el
   límite interior del pilar (sin sonda adicional); flag ``pata_l`` + ``hook_mm`` para
   polilínea L al colocar.
5. **Sin colisión:** retractar sonda (volver al nodo fibra) + **−(25 + Ø/2)** mm hacia el
   interior; ``pata_l`` + ``hook_mm`` para polilínea L al colocar.

No abre transacciones.
"""

from __future__ import division

import math

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    Arc,
    BuiltInCategory,
    CurveLoop,
    Frame,
    GeometryCreationUtilities,
    GeometryInstance,
    Line,
    Options,
    PlanarFace,
    Plane,
    Solid,
    SolidCurveIntersectionMode,
    SolidCurveIntersectionOptions,
    UV,
    ViewDetailLevel,
    XYZ,
)

from System.Collections.Generic import List

from geometria_colision_vigas import obtener_solidos_elemento, solidos_intersectan_por_booleana

try:
    from geometria_viga_cara_superior_detalle import (
        _iter_planar_faces_elemento,
        _param_line_plane_intersection_distance,
        _punto_sobre_cara_planar,
    )
except Exception:
    _iter_planar_faces_elemento = None
    _param_line_plane_intersection_distance = None
    _punto_sobre_cara_planar = None

try:
    from evaluacion_curva_puntos_obstaculos import _punto_en_volumen_solido
except Exception:
    _punto_en_volumen_solido = None

try:
    from enfierrado_shaft_hashtag import (
        _nearest_tabulated_bar_diameter_mm,
        extension_mm_por_diametro_nominal_mm,
    )
except Exception:
    _nearest_tabulated_bar_diameter_mm = None
    extension_mm_por_diametro_nominal_mm = None

try:
    from bimtools_rebar_hook_lengths import hook_length_mm_from_nominal_diameter_mm
except Exception:
    hook_length_mm_from_nominal_diameter_mm = None

# Sonda diagnóstica (mm) — igual retrae si no hay colisión.
SONDA_COLISION_MM = 50.0
RECUBRIMIENTO_FIBRA_MM = 25.0
DIAM_NOMINAL_RESPALDO_MM = 16.0

_PROBE_SPHERE_RADIUS_MM = 2.0
_SONDA_MUESTRAS = 6
_TOL_VOLUMEN_PROBE_FT3 = 1e-12
_TOL_SPLIT_FT = 1.0 / 304.8
_MIN_LINE_LEN_FT = 1.0 / 304.8 * 5.0
_MIN_FACE_AREA_FT2 = 1e-8

MODO_EMPOTRAMIENTO = u"empotramiento"
MODO_GANCHO = u"gancho"


def _mm_to_ft(mm):
    return float(mm) / 304.8


def _geometry_options():
    opts = Options()
    try:
        opts.ComputeReferences = False
    except Exception:
        pass
    try:
        opts.DetailLevel = ViewDetailLevel.Fine
    except Exception:
        pass
    try:
        opts.IncludeNonVisibleObjects = True
    except Exception:
        pass
    return opts


def _iter_solids_elemento(elemento, opts):
    if elemento is None:
        return
    try:
        ge = elemento.get_Geometry(opts)
    except Exception:
        return
    if ge is None:
        return
    for obj in ge:
        if obj is None:
            continue
        if isinstance(obj, Solid):
            try:
                if float(obj.Volume) < 1e-9:
                    continue
            except Exception:
                continue
            yield obj
        elif isinstance(obj, GeometryInstance):
            try:
                sub = obj.GetInstanceGeometry()
            except Exception:
                continue
            if sub is None:
                continue
            for g2 in sub:
                if isinstance(g2, Solid):
                    try:
                        if float(g2.Volume) < 1e-9:
                            continue
                    except Exception:
                        continue
                    yield g2


def _solid_esfera_en_centro(centre, radius_ft):
    if centre is None:
        return None
    try:
        r = float(radius_ft)
    except Exception:
        return None
    if r <= 1e-12:
        return None
    try:
        frame = Frame(centre, XYZ.BasisX, XYZ.BasisY, XYZ.BasisZ)
        p1 = centre - XYZ.BasisZ.Multiply(r)
        p2 = centre + XYZ.BasisZ.Multiply(r)
        p_mid = centre + XYZ.BasisX.Multiply(r)
        arc = Arc.Create(p1, p2, p_mid)
        ln = Line.CreateBound(arc.GetEndPoint(1), arc.GetEndPoint(0))
        loop = CurveLoop()
        loop.Append(arc)
        loop.Append(ln)
        loops = List[CurveLoop]()
        loops.Add(loop)
        return GeometryCreationUtilities.CreateRevolvedGeometry(
            frame, loops, 0.0, 2.0 * math.pi
        )
    except Exception:
        return None


def _id_en_excluidos(eid, excluir_ids):
    for x in excluir_ids or []:
        try:
            if int(x.IntegerValue) == int(eid.IntegerValue):
                return True
        except Exception:
            try:
                if x == eid:
                    return True
            except Exception:
                pass
    return False


def _probe_esfera_toca_seleccion(document, ids_seleccion, excluir_ids, p_probe):
    el = _probe_esfera_primer_elemento(document, ids_seleccion, excluir_ids, p_probe)
    return el is not None


def _probe_esfera_primer_elemento(document, ids_seleccion, excluir_ids, p_probe):
    """Primer elemento de la selección cuyo sólido intersecta la esfera de sonda."""
    if document is None or p_probe is None:
        return None
    r_ft = _mm_to_ft(_PROBE_SPHERE_RADIUS_MM)
    sphere = _solid_esfera_en_centro(p_probe, r_ft)
    opts = _geometry_options()
    for eid in ids_seleccion or []:
        if _id_en_excluidos(eid, excluir_ids):
            continue
        try:
            el = document.GetElement(eid)
        except Exception:
            el = None
        if el is None:
            continue
        for s in _iter_solids_elemento(el, opts):
            hit = False
            if sphere is not None:
                try:
                    if solidos_intersectan_por_booleana(
                        sphere, s, tol_volumen=_TOL_VOLUMEN_PROBE_FT3
                    ):
                        hit = True
                except Exception:
                    pass
            elif _punto_en_volumen_solido is not None:
                try:
                    if _punto_en_volumen_solido(s, p_probe):
                        hit = True
                except Exception:
                    pass
            if hit:
                return el
    return None


def _elemento_es_columna_estructural(elem):
    if elem is None:
        return False
    try:
        cat = elem.Category
        if cat is None:
            return False
        return int(cat.Id.IntegerValue) == int(BuiltInCategory.OST_StructuralColumns)
    except Exception:
        return False


def _sonda_tramo_info_seleccion(
    document, ids_seleccion, excluir_ids, p_ext, du_unit, len_mm, n_muestras
):
    """
    Muestreo del tramo de sonda: colisión global y columnas estructurales detectadas.

    Returns:
        ``dict`` con ``choca``, ``columnas`` (lista sin duplicados), ``t_colision`` (0…1),
        ``elemento_colision``.
    """
    info = {
        u"choca": False,
        u"columnas": [],
        u"t_colision": None,
        u"elemento_colision": None,
    }
    try:
        L = _mm_to_ft(float(len_mm))
    except Exception:
        return info
    if p_ext is None or du_unit is None or L <= 1e-12:
        return info
    try:
        du = du_unit.Normalize()
    except Exception:
        return info
    n = max(2, int(n_muestras or _SONDA_MUESTRAS))
    cols_seen = set()
    for i in range(n + 1):
        t = float(i) / float(n)
        p = p_ext + du.Multiply(L * t)
        el = _probe_esfera_primer_elemento(document, ids_seleccion, excluir_ids, p)
        if el is None:
            continue
        info[u"choca"] = True
        if info[u"t_colision"] is None:
            info[u"t_colision"] = t
            info[u"elemento_colision"] = el
        if _elemento_es_columna_estructural(el):
            try:
                eid = int(el.Id.IntegerValue)
            except Exception:
                eid = None
            if eid is not None and eid not in cols_seen:
                cols_seen.add(eid)
                info[u"columnas"].append(el)
    return info


def _sonda_tramo_toca_seleccion(
    document, ids_seleccion, excluir_ids, p_ext, du_unit, len_mm, n_muestras
):
    info = _sonda_tramo_info_seleccion(
        document, ids_seleccion, excluir_ids, p_ext, du_unit, len_mm, n_muestras
    )
    return bool(info.get(u"choca"))


def _solid_line_inside_param_intervals(line, solid):
    """Tramos del segmento dentro del sólido (parámetros nativos de la curva)."""
    out = []
    if line is None or solid is None:
        return out
    try:
        scio = SolidCurveIntersectionOptions()
        try:
            scio.ResultType = SolidCurveIntersectionMode.CurveSegmentsInside
        except Exception:
            pass
        sci = solid.IntersectWithCurve(line, scio)
    except Exception:
        return out
    if sci is None:
        return out
    try:
        n = int(sci.SegmentCount)
    except Exception:
        return out
    if n < 1:
        return out
    for i in range(n):
        try:
            ext = sci.GetCurveSegmentExtents(i)
            s0 = float(ext.StartParameter)
            s1 = float(ext.EndParameter)
            if s1 < s0:
                s0, s1 = s1, s0
            out.append((s0, s1))
        except Exception:
            continue
    return out


def _merge_param_intervals(intervals):
    if not intervals:
        return []
    iv = sorted(intervals, key=lambda x: x[0])
    merged = []
    for a, b in iv:
        if b < a:
            a, b = b, a
        if not merged or a > merged[-1][1] + 1e-9:
            merged.append([a, b])
        else:
            merged[-1][1] = max(merged[-1][1], b)
    return [(float(a), float(b)) for a, b in merged]


def _intervalos_dentro_elemento(line, elemento):
    """Unión de intervalos paramétricos de ``line`` dentro de los sólidos del elemento."""
    if line is None or elemento is None:
        return []
    acc = []
    try:
        solids = obtener_solidos_elemento(elemento, _geometry_options())
    except Exception:
        solids = []
    for s in solids or []:
        acc.extend(_solid_line_inside_param_intervals(line, s))
    return _merge_param_intervals(acc)


def _plano_desde_face_local(face):
    if face is None or not isinstance(face, PlanarFace):
        return None
    try:
        n = face.FaceNormal
        if n is None or n.GetLength() < 1e-12:
            return None
        n = n.Normalize()
        bb_uv = face.GetBoundingBox()
        if bb_uv is None:
            return None
        u_mid = 0.5 * (float(bb_uv.Min.U) + float(bb_uv.Max.U))
        v_mid = 0.5 * (float(bb_uv.Min.V) + float(bb_uv.Max.V))
        o = face.Evaluate(UV(u_mid, v_mid))
        if o is None:
            return None
        return Plane.CreateByNormalAndOrigin(n, o)
    except Exception:
        return None


def _param_line_plane_distance_fallback(line, plane):
    if _param_line_plane_intersection_distance is not None:
        try:
            return _param_line_plane_intersection_distance(line, plane)
        except Exception:
            pass
    if line is None or plane is None:
        return None
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        d_raw = p1 - p0
        L = float(d_raw.GetLength())
        if L < _MIN_LINE_LEN_FT:
            return None
        du = d_raw.Normalize()
        n = plane.Normal
        if n is None or n.GetLength() < 1e-12:
            return None
        n = n.Normalize()
        o = plane.Origin
        denom = float(du.DotProduct(n))
        if abs(denom) < 1e-12:
            return None
        s = float((o - p0).DotProduct(n)) / denom
        if s < -_TOL_SPLIT_FT or s > L + _TOL_SPLIT_FT:
            return None
        return max(0.0, min(L, s))
    except Exception:
        return None


def _punto_sobre_cara_fallback(pt, face):
    if _punto_sobre_cara_planar is not None:
        try:
            return _punto_sobre_cara_planar(pt, face)
        except Exception:
            pass
    if pt is None or face is None:
        return False
    try:
        r = face.Project(pt)
        if r is None:
            return False
        d = float(r.Distance)
        return d <= _TOL_SPLIT_FT
    except Exception:
        return False


def _iter_caras_planas_elemento(elemento):
    if _iter_planar_faces_elemento is not None:
        try:
            for f in _iter_planar_faces_elemento(elemento):
                yield f
            return
        except Exception:
            pass
    opts = _geometry_options()
    for s in _iter_solids_elemento(elemento, opts):
        try:
            faces = s.Faces
        except Exception:
            continue
        for f in faces:
            if isinstance(f, PlanarFace):
                yield f


def _params_interseccion_linea_caras_elemento(line, elemento):
    """Parámetros (0…Length) donde ``line`` corta caras planas del elemento."""
    out = []
    if line is None or elemento is None:
        return out
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        L = float(line.Length)
        if L < _MIN_LINE_LEN_FT:
            return out
    except Exception:
        return out
    for face in _iter_caras_planas_elemento(elemento):
        if face is None or not isinstance(face, PlanarFace):
            continue
        try:
            if face.Area < _MIN_FACE_AREA_FT2:
                continue
        except Exception:
            pass
        pl = _plano_desde_face_local(face)
        if pl is None:
            continue
        t = _param_line_plane_distance_fallback(line, pl)
        if t is None:
            continue
        try:
            du = (p1 - p0).Normalize()
            q = p0 + du.Multiply(t)
        except Exception:
            continue
        if not _punto_sobre_cara_fallback(q, face):
            continue
        out.append(float(t))
    return out


def _sonda_tramo_toca_elemento(document, elemento, p_ext, du_unit, len_mm, n_muestras):
    """Sonda sobre un único elemento (sin extensión adicional)."""
    if elemento is None:
        return False
    try:
        eid = elemento.Id
    except Exception:
        return False
    return _sonda_tramo_toca_seleccion(
        document, [eid], [], p_ext, du_unit, len_mm, n_muestras
    )


def _refinar_punto_extremo_columna(
    p_ext,
    du,
    columna,
    stretch_mm,
    d_mm,
    mm_rec,
    document=None,
):
    """
    Tras empotramiento completo: recorta al volumen/cara del pilar y retrae
    ``−(mm_rec + Ø/2)`` desde el límite interior detectado.

    Returns:
        ``(p_final, meta_extra)`` — ``meta_extra`` vacío si no hubo refinamiento.
    """
    meta = {}
    if p_ext is None or du is None or columna is None:
        return None, meta
    try:
        L_mm = float(stretch_mm)
    except Exception:
        return None, meta
    if L_mm <= 1e-9:
        return None, meta
    L_ft = _mm_to_ft(L_mm)
    try:
        p_emp = p_ext + du.Multiply(L_ft)
        seg = Line.CreateBound(p_ext, p_emp)
    except Exception:
        return None, meta
    try:
        seg_len = float(seg.Length)
    except Exception:
        return None, meta
    if seg_len < _MIN_LINE_LEN_FT:
        return None, meta

    if document is not None:
        if not _sonda_tramo_toca_elemento(
            document, columna, p_ext, du, L_mm, _SONDA_MUESTRAS
        ):
            return None, meta

    inside = _intervalos_dentro_elemento(seg, columna)
    face_params = _params_interseccion_linea_caras_elemento(seg, columna)

    t_entry = None
    t_exit = None
    if inside:
        t_entry = float(inside[0][0])
        t_exit = float(inside[-1][1])
        for a, b in inside:
            if a < t_entry:
                t_entry = float(a)
            if b > t_exit:
                t_exit = float(b)
    elif face_params:
        face_params = sorted(face_params)
        t_entry = float(face_params[0])
        t_exit = float(face_params[-1])

    if t_exit is None:
        return None, meta

    t_entry = max(0.0, float(t_entry or 0.0))
    t_exit = min(seg_len, max(t_entry, float(t_exit)))

    retract_mm = float(mm_rec) + 0.5 * float(d_mm)
    retract_ft = _mm_to_ft(retract_mm)

    try:
        p_boundary = p_ext + du.Multiply(t_exit)
        p_final = p_boundary - du.Multiply(retract_ft)
    except Exception:
        return None, meta

    try:
        col_id = int(columna.Id.IntegerValue)
    except Exception:
        col_id = None

    t_face = t_entry
    if face_params:
        t_face = min(face_params, key=lambda t: abs(float(t) - t_entry))

    meta = {
        u"columna_id": col_id,
        u"columna_refinado": True,
        u"t_entry_ft": float(t_entry),
        u"t_exit_ft": float(t_exit),
        u"t_face_ft": float(t_face) if t_face is not None else None,
        u"retract_mm": float(retract_mm),
        u"punto_limite": p_boundary,
    }
    return p_final, meta


def _empotramiento_mm_desde_diametro(diam_nominal_mm):
    if extension_mm_por_diametro_nominal_mm is None:
        return 0.0, u""
    d_in = diam_nominal_mm
    try:
        d_in = float(diam_nominal_mm)
    except Exception:
        d_in = DIAM_NOMINAL_RESPALDO_MM
    if d_in <= 1e-9:
        d_in = DIAM_NOMINAL_RESPALDO_MM
    if _nearest_tabulated_bar_diameter_mm is not None:
        try:
            d_tab = _nearest_tabulated_bar_diameter_mm(d_in)
        except Exception:
            d_tab = d_in
    else:
        d_tab = d_in
    try:
        ext_mm, desc = extension_mm_por_diametro_nominal_mm(d_tab)
        return max(0.0, float(ext_mm)), desc
    except Exception:
        return 0.0, u""


def _hook_mm_desde_diametro(diam_nominal_mm):
    if hook_length_mm_from_nominal_diameter_mm is None:
        return 0.0
    try:
        d = float(diam_nominal_mm or DIAM_NOMINAL_RESPALDO_MM)
    except Exception:
        d = DIAM_NOMINAL_RESPALDO_MM
    if d <= 1e-9:
        d = DIAM_NOMINAL_RESPALDO_MM
    return float(hook_length_mm_from_nominal_diameter_mm(d))


def resolver_extremo_linea(
    document,
    p_ext,
    dir_saliente_unit,
    ids_seleccion,
    ids_excluir=None,
    diam_nominal_mm=None,
    mm_sonda=None,
    mm_recubrimiento=None,
):
    """
    Resuelve un extremo de fibra fusionada.

    Args:
        p_ext: nodo fibra (``XYZ``) antes de sonda.
        dir_saliente_unit: tangente **hacia fuera** del tramo.
        ids_seleccion: ``ElementId`` de todo el lote inicial.
        ids_excluir: hosts a ignorar (cadena colineal).
        diam_nominal_mm: Ø barra para tabla empotramiento / gancho.

    Returns:
        ``dict`` con ``punto``, ``modo`` (``MODO_*``), ``delta_mm`` (desde ``p_ext``,
        positivo = hacia fuera), ``hook_mm`` (solo ``MODO_GANCHO``), ``emp_mm``,
        ``sonda_mm``, ``descripcion``, ``pata_l`` (bool). Si la colisión es en columna
        estructural, puede incluir ``columna_id``, ``columna_refinado``,
        ``punto_limite_columna``, ``t_face_ft``, ``retract_mm``.
    """
    mm_s = float(mm_sonda if mm_sonda is not None else SONDA_COLISION_MM)
    mm_rec = float(
        mm_recubrimiento if mm_recubrimiento is not None else RECUBRIMIENTO_FIBRA_MM
    )
    try:
        d_mm = float(diam_nominal_mm or DIAM_NOMINAL_RESPALDO_MM)
    except Exception:
        d_mm = DIAM_NOMINAL_RESPALDO_MM
    if d_mm <= 1e-9:
        d_mm = DIAM_NOMINAL_RESPALDO_MM

    if p_ext is None or dir_saliente_unit is None:
        return {
            u"punto": p_ext,
            u"modo": MODO_GANCHO,
            u"delta_mm": 0.0,
            u"hook_mm": _hook_mm_desde_diametro(d_mm),
            u"emp_mm": 0.0,
            u"sonda_mm": mm_s,
            u"descripcion": u"Sin geometría de extremo.",
        }

    try:
        du = dir_saliente_unit.Normalize()
    except Exception:
        du = dir_saliente_unit

    sonda_info = _sonda_tramo_info_seleccion(
        document,
        ids_seleccion,
        ids_excluir,
        p_ext,
        du,
        mm_s,
        _SONDA_MUESTRAS,
    )
    choca = bool(sonda_info.get(u"choca"))
    columnas_sonda = list(sonda_info.get(u"columnas") or [])

    if choca:
        emp_mm, desc = _empotramiento_mm_desde_diametro(d_mm)
        delta = mm_s + emp_mm
        try:
            p_nuevo = p_ext + du.Multiply(_mm_to_ft(delta))
        except Exception:
            p_nuevo = p_ext

        meta = {
            u"punto": p_nuevo,
            u"modo": MODO_EMPOTRAMIENTO,
            u"delta_mm": float(delta),
            u"hook_mm": 0.0,
            u"emp_mm": float(emp_mm),
            u"sonda_mm": mm_s,
            u"descripcion": desc or u"Empotramiento Ø{0} mm".format(int(round(d_mm))),
            u"colision_columna": bool(columnas_sonda),
        }

        if columnas_sonda:
            columna = columnas_sonda[0]
            p_ref, ref_meta = _refinar_punto_extremo_columna(
                p_ext,
                du,
                columna,
                delta,
                d_mm,
                mm_rec,
                document=document,
            )
            if p_ref is not None and ref_meta:
                try:
                    hook_mm = _hook_mm_desde_diametro(d_mm)
                    delta_ref_ft = float((p_ref - p_ext).DotProduct(du))
                    delta_ref_mm = delta_ref_ft * 304.8
                    meta[u"punto"] = p_ref
                    meta[u"delta_mm"] = float(delta_ref_mm)
                    meta[u"columna_id"] = ref_meta.get(u"columna_id")
                    meta[u"columna_refinado"] = True
                    meta[u"t_face_ft"] = ref_meta.get(u"t_face_ft")
                    meta[u"punto_limite_columna"] = ref_meta.get(u"punto_limite")
                    meta[u"retract_mm"] = ref_meta.get(u"retract_mm")
                    meta[u"pata_l"] = True
                    meta[u"hook_mm"] = float(hook_mm)
                    meta[u"descripcion"] = (
                        u"Pilar {0} · recorte cara + −({1}+{2}/2) mm · pata L {3:.0f} mm".format(
                            ref_meta.get(u"columna_id") or u"?",
                            int(round(mm_rec)),
                            int(round(d_mm)),
                            hook_mm,
                        )
                    )
                except Exception:
                    meta[u"columna_refinado"] = False
                    meta[u"descripcion"] = (
                        (desc or u"Empotramiento Ø{0} mm".format(int(round(d_mm))))
                        + u" · refinamiento pilar falló"
                    )
            else:
                meta[u"columna_refinado"] = False
                if len(columnas_sonda) > 1:
                    meta[u"columnas_detectadas"] = len(columnas_sonda)
                meta[u"descripcion"] = (
                    (desc or u"Empotramiento Ø{0} mm".format(int(round(d_mm))))
                    + u" · pilar sin cara/volumen utilizable"
                )

        return meta

    retract_mm = mm_rec + 0.5 * d_mm
    delta = -retract_mm
    try:
        p_nuevo = p_ext + du.Multiply(_mm_to_ft(delta))
    except Exception:
        p_nuevo = p_ext
    hook_mm = _hook_mm_desde_diametro(d_mm)
    return {
        u"punto": p_nuevo,
        u"modo": MODO_GANCHO,
        u"delta_mm": float(delta),
        u"hook_mm": float(hook_mm),
        u"emp_mm": 0.0,
        u"sonda_mm": mm_s,
        u"pata_l": True,
        u"descripcion": u"Extremo libre · −({0}+{1}/2) mm · pata L {2:.0f} mm".format(
            int(round(mm_rec)), int(round(d_mm)), hook_mm
        ),
    }


def aplicar_extremos_linea(
    document,
    line,
    ids_seleccion,
    ids_excluir=None,
    diam_nominal_mm=None,
    resolver_inicio=True,
    resolver_fin=True,
):
    """
    Aplica :func:`resolver_extremo_linea` en ``GetEndPoint(0)`` y/o ``(1)``.

    Returns:
        ``(line_nueva, meta_inicio, meta_fin)`` — ``line_nueva`` puede ser ``None`` si inválida.
    """
    if line is None:
        return None, None, None
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        t = (p1 - p0).Normalize()
    except Exception:
        return None, None, None

    meta_i = meta_f = None
    pa, pb = p0, p1

    if resolver_inicio:
        meta_i = resolver_extremo_linea(
            document,
            p0,
            t.Negate(),
            ids_seleccion,
            ids_excluir,
            diam_nominal_mm=diam_nominal_mm,
        )
        if meta_i and meta_i.get(u"punto") is not None:
            pa = meta_i[u"punto"]

    if resolver_fin:
        meta_f = resolver_extremo_linea(
            document,
            p1,
            t,
            ids_seleccion,
            ids_excluir,
            diam_nominal_mm=diam_nominal_mm,
        )
        if meta_f and meta_f.get(u"punto") is not None:
            pb = meta_f[u"punto"]

    try:
        if pa.DistanceTo(pb) < _MIN_LINE_LEN_FT:
            return None, meta_i, meta_f
        return Line.CreateBound(pa, pb), meta_i, meta_f
    except Exception:
        return None, meta_i, meta_f


def element_ids_desde_elementos(elementos):
    """Lista de ``ElementId`` desde elementos Revit."""
    out = []
    for el in elementos or []:
        if el is None:
            continue
        try:
            out.append(el.Id)
        except Exception:
            pass
    return out
