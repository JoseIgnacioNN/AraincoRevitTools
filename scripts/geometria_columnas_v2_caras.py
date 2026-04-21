# -*- coding: utf-8 -*-
"""
Armadura columnas V2 — geometría en caras laterales por **separación real** entre opuestos.

**Cara A:** par con **mayor** separación entre opuestos; ``na`` barras; paso y recubrimiento en
cara usan la **luz ortogonal** (menor separación del otro par) y ``min(Width, Depth)``.

**Cara B:** par **menor**; ``nb``; luz ortogonal = mayor separación y ``max(Width, Depth)``.
**Lados iguales** (misma ``sep`` en todos los pares): en todas las caras se usa **``nb``** (Cara B).

Por cada tarea se generan líneas en **las dos caras opuestas** del par (orden de anclas).
Fusión, deduplicado por punto medio, sondeo/empotramiento, ``ModelCurve`` y marcadores como antes.
Primera capa: comentario ``BIMTools_ColV2_EjeCara``; capas 2+ ``BIMTools_ColV2_EjeCara·L2`` / ``·L3``…
(troceo con planos de empalme: todas las capas; ver ``troceo_model_curves_planos_empalme_v2``).
"""

from Autodesk.Revit.DB import BuiltInParameter, Line, Transaction, XYZ
from Autodesk.Revit.DB.Structure import StructuralType

from armadura_vigas_capas import _read_width_depth_ft
from geometria_columnas_eje import (
    _aplicar_offset_interior_cara,
    _curva_eje_para_proyeccion,
    _dedupe_lineas_misma_posicion,
    _dedupe_lineas_mismo_punto_medio,
    _EMPOTRAMIENTO_PRUEBA_MM,
    _face_basis_x_unit_tuple,
    _mm_to_ft,
    _project_point_to_plane,
    aplicar_empotramiento_lineas_unificadas,
    crear_marcador_normal_curva_eje,
    filtrar_solo_structural_columns,
    fundaciones_estructurales_unidas,
    sketch_plane_para_linea,
    _lateral_face_anchors_from_columna,
    _vector_ancho_en_cara,
)
from geometria_viga_cara_superior_detalle import _MIN_LINE_LEN_FT, _unificar_lineas_colineales

# Tolerancia paralelismo / distancia punto–recta (pies).
_TOL_LINE_GROUP_FT = 2.0 / 304.8
# Descuento fijo (mm) sobre el ancho de cara para el tramo útil entre líneas (misma convención
# que ``_OCHO_BARRAS_DESCUENTO_ANCHO_MM`` en eje columnas).
_ANCHO_CARA_DESCUENTO_MM = 50.0
# Estirón al extremo del eje en ``GetEndPoint(0)``: altura bbox fundación − este valor (mm).
_ESTIRON_FUNDACION_DESCUENTO_MM = 50.0
# Acortamiento en el extremo del eje en ``GetEndPoint(1)`` (mm), misma idea que ``_RECORTE_FIN_MM``.
_ACORTE_EXTREMO_FIN_MM = 25.0
# Tolerancia (pies): todos los pares laterales con la misma separación ⇒ sección cuadrada / regular.
_TOL_SECCION_CUADRADA_FT = 2.0 / 304.8


def _altura_envolvente_z_fundaciones_mm(elementos):
    """Rango Z (mm) de la envolvente de todos los ``BoundingBox`` de fundaciones unidas."""
    z_lo = None
    z_hi = None
    for el in elementos or []:
        if el is None:
            continue
        try:
            bb = el.get_BoundingBox(None)
            if bb is None:
                continue
            lo = float(bb.Min.Z)
            hi = float(bb.Max.Z)
            z_lo = lo if z_lo is None else min(z_lo, lo)
            z_hi = hi if z_hi is None else max(z_hi, hi)
        except Exception:
            continue
    if z_lo is None or z_hi is None:
        return 0.0
    return max(0.0, (z_hi - z_lo) * 304.8)


def _mm_estiron_desde_start_si_fundacion(document, columna):
    """
    Si hay fundación estructural unida (*Join Geometry*), retorna
    ``altura_fundación_mm − 50``; si no, ``0``.
    """
    if document is None or columna is None:
        return 0.0
    try:
        funds = fundaciones_estructurales_unidas(document, columna)
    except Exception:
        funds = []
    if not funds:
        return 0.0
    h_mm = _altura_envolvente_z_fundaciones_mm(funds)
    return max(0.0, h_mm - float(_ESTIRON_FUNDACION_DESCUENTO_MM))


def _extender_linea_desde_start_column_curve(ln, crv_columna, delta_mm):
    """
    Alarga el extremo del tramo alineado con ``GetEndPoint(0)`` de ``crv_columna`` en
    ``delta_mm`` siguiendo **−axis** (axis = ``(end1 − end0).Normalize()`` de la curva
    de la columna: hacia la cimentación bajo el arranque analítico).
    """
    if ln is None or crv_columna is None or delta_mm <= 1e-9:
        return ln
    if not crv_columna.IsBound:
        return ln
    try:
        p0 = crv_columna.GetEndPoint(0)
        p1 = crv_columna.GetEndPoint(1)
        dvec = p1 - p0
        lg = float(dvec.GetLength())
        if lg < 1e-12:
            return ln
        axis = dvec.Normalize()
    except Exception:
        return ln
    try:
        qa = ln.GetEndPoint(0)
        qb = ln.GetEndPoint(1)
        ta = float((qa - p0).DotProduct(axis))
        tb = float((qb - p0).DotProduct(axis))
    except Exception:
        return ln
    delta_ft = _mm_to_ft(float(delta_mm))
    try:
        if ta <= tb:
            q_lo = qa
            q_hi = qb
        else:
            q_lo = qb
            q_hi = qa
        q_lo_new = q_lo - axis.Multiply(delta_ft)
        if ta <= tb:
            return Line.CreateBound(q_lo_new, q_hi)
        return Line.CreateBound(q_hi, q_lo_new)
    except Exception:
        return ln


def _acortar_linea_extremo_fin_column_curve(ln, crv_columna, delta_mm):
    """
    **Estiramiento negativo** / recorte de ``delta_mm`` en el extremo del tramo correspondiente a
    ``GetEndPoint(1)`` de ``crv_columna`` (mayor proyección sobre el eje 0→1).
    """
    if ln is None or crv_columna is None or delta_mm <= 1e-9:
        return ln
    if not crv_columna.IsBound:
        return ln
    try:
        p0 = crv_columna.GetEndPoint(0)
        p1 = crv_columna.GetEndPoint(1)
        dvec = p1 - p0
        lg = float(dvec.GetLength())
        if lg < 1e-12:
            return ln
        axis = dvec.Normalize()
    except Exception:
        return ln
    try:
        qa = ln.GetEndPoint(0)
        qb = ln.GetEndPoint(1)
        ta = float((qa - p0).DotProduct(axis))
        tb = float((qb - p0).DotProduct(axis))
    except Exception:
        return ln
    delta_ft = _mm_to_ft(float(delta_mm))
    try:
        if ta <= tb:
            q_lo = qa
            q_hi = qb
        else:
            q_lo = qb
            q_hi = qa
        q_hi_new = q_hi - axis.Multiply(delta_ft)
        if q_lo.DistanceTo(q_hi_new) < _MIN_LINE_LEN_FT:
            return ln
        if ta <= tb:
            return Line.CreateBound(q_lo, q_hi_new)
        return Line.CreateBound(q_hi_new, q_lo)
    except Exception:
        return ln


def _basis_x_xyz_desde_cara(face):
    """``XYZ`` unitario = eje X local de la cara (``ComputeDerivatives`` en UV medio)."""
    t = _face_basis_x_unit_tuple(face)
    if t is None:
        return None
    try:
        v = XYZ(float(t[0]), float(t[1]), float(t[2]))
        if v.GetLength() < 1e-12:
            return None
        return v.Normalize()
    except Exception:
        return None


def _direccion_traslacion_bxis_en_cara_respecto_eje(face, axis_unit):
    """
    Dirección **en el plano de la cara** y **perpendicular** al eje proyectado del pilar, obtenida
    proyectando el ``BasisX`` de la cara: si el paramétrico U sigue la dirección del pilar (muy
    habitual), ``BasisX`` es casi ∥ al eje y una traslación pura en X **no desplaza** la línea
    lateralmente (misma recta). Por eso se elimina la componente paralela al eje; si queda
    degenerado se usa ``_vector_ancho_en_cara``. El sentido se alinea con ``_vector_ancho_en_cara``.
    """
    if face is None or axis_unit is None:
        return None
    try:
        ax = axis_unit.Normalize()
    except Exception:
        return None
    if ax is None or ax.GetLength() < 1e-12:
        return None
    u_ref = _vector_ancho_en_cara(face, ax)
    bx = _basis_x_xyz_desde_cara(face)
    if bx is None:
        return u_ref
    try:
        n = face.FaceNormal
        if n is None or n.GetLength() < 1e-12:
            return u_ref
        n = n.Normalize()
        bx_t = bx - n.Multiply(float(bx.DotProduct(n)))
        if bx_t.GetLength() < 1e-9:
            return u_ref
        bx_t = bx_t.Normalize()
        para = float(bx_t.DotProduct(ax))
        perp = bx_t - ax.Multiply(para)
        ln = perp.GetLength()
        if ln < 1e-6:
            return u_ref
        dir_tr = perp.Multiply(1.0 / ln)
        if u_ref is not None and float(dir_tr.DotProduct(u_ref)) < 0.0:
            dir_tr = dir_tr.Negate()
        return dir_tr
    except Exception:
        return u_ref


def _vector_distribucion_menos_eje_x_cara(face, axis_unit):
    """
    Dirección para **copias** en cara: **opuesta** al eje X paramétrico de la cara, en el plano
    tangente y ⟂ al eje columna (misma proyección que en
    :func:`_direccion_traslacion_bxis_en_cara_respecto_eje`, sin fijar sentido con ``u_ref``).
    Si degenera: ``−_vector_ancho_en_cara``.
    """
    if face is None or axis_unit is None:
        return None
    try:
        ax = axis_unit.Normalize()
    except Exception:
        return None
    if ax is None or ax.GetLength() < 1e-12:
        return None
    u_ref = _vector_ancho_en_cara(face, ax)
    bx = _basis_x_xyz_desde_cara(face)
    if bx is None:
        try:
            return u_ref.Negate() if u_ref is not None else None
        except Exception:
            return None
    try:
        n = face.FaceNormal
        if n is None or n.GetLength() < 1e-12:
            return u_ref.Negate() if u_ref is not None else None
        n = n.Normalize()
        bx_t = bx - n.Multiply(float(bx.DotProduct(n)))
        if bx_t.GetLength() < 1e-9:
            return u_ref.Negate() if u_ref is not None else None
        bx_t = bx_t.Normalize()
        para = float(bx_t.DotProduct(ax))
        perp = bx_t - ax.Multiply(para)
        ln = perp.GetLength()
        if ln < 1e-6:
            return u_ref.Negate() if u_ref is not None else None
        dir_pos = perp.Multiply(1.0 / ln)
        return dir_pos.Negate()
    except Exception:
        try:
            return u_ref.Negate() if u_ref is not None else None
        except Exception:
            return None


def _mm_traslacion_bxis_ancho_menos_recubrimiento(width_sep_ft, offset_mm, width_type_ft=None):
    """
    Distancia (mm) en la dirección anterior: ``ancho/2 − (25 + Ø estribo + Ø long./2)``.
    Se usa el mayor entre el ancho derivado del par de caras y el **ancho de tipo** (pies), para
    no quedar en 0 si la geometría BRep es más estrecha que el parámetro del tipo.
    """
    try:
        w_mm = float(width_sep_ft) * 304.8
        if width_type_ft is not None:
            try:
                wt = float(width_type_ft) * 304.8
                if wt > w_mm:
                    w_mm = wt
            except Exception:
                pass
        off = float(offset_mm or 0.0)
        return max(0.0, 0.5 * w_mm - off)
    except Exception:
        return 0.0


def _ancho_cara_proyeccion_mm(width_sep_ft, width_type_ft):
    """Milímetros: mayor entre separación BRep del par ancho y ancho de tipo."""
    try:
        w_mm = float(width_sep_ft) * 304.8
    except Exception:
        w_mm = 0.0
    if width_type_ft is not None:
        try:
            wt = float(width_type_ft) * 304.8
            if wt > w_mm:
                w_mm = wt
        except Exception:
            pass
    return w_mm


def _paso_propagacion_mm_cara_a(n_barras, w_mm, diam_estribo_mm, diam_long_mm):
    """
    Tramo útil ``L = w - 50 - 2·Ø est - Ø long``; paso entre líneas consecutivas = ``L / (n-1)``.
    Casos 2/3/4 barras coinciden con la especificación (divide en 1, 2 o 3 intervalos).
    """
    try:
        n = int(n_barras)
    except Exception:
        n = 1
    if n <= 1:
        return 0.0, 0.0
    try:
        w = float(w_mm)
        de = float(diam_estribo_mm or 0.0)
        dl = float(diam_long_mm or 0.0)
        L = w - float(_ANCHO_CARA_DESCUENTO_MM) - 2.0 * de - dl
    except Exception:
        L = 0.0
    L = max(0.0, L)
    step = L / float(n - 1)
    return L, step


def _paso_entre_lineas_distrib_mm(sep_ft, tipo_ft, n_barras, d_est_mm, d_long_mm):
    """Paso ``L/(n−1)`` (mm) entre líneas copiadas en cara, misma receta que primera capa."""
    w_mm = _ancho_cara_proyeccion_mm(sep_ft, tipo_ft)
    _, step = _paso_propagacion_mm_cara_a(
        n_barras, w_mm, d_est_mm, d_long_mm
    )
    return float(step)


def _pasos_cara_a_y_b_mm(tareas, w_ft, d_ft, na, nb, d_est_mm, d_long_mm):
    """
    Pasos entre líneas (mm) de la receta **A** (``na``, luz ortogonal ``wmin``) y **B** (``nb``, ``wmaj``).
    En sección con lados iguales devuelve el mismo valor dos veces. Si solo hay tarea A, ``step_b``
    se calcula con ``nb`` y ``min(w,d)``.
    """
    step_a = None
    step_b = None
    step_sq = None
    for t in tareas or []:
        try:
            _, sep, tipo, n, etq = t
        except Exception:
            continue
        st = _paso_entre_lineas_distrib_mm(sep, tipo, n, d_est_mm, d_long_mm)
        if u"lados iguales" in etq:
            step_sq = st
            break
        if u"Cara A (ancho)" in etq:
            step_a = st
        if u"Cara B (alto)" in etq:
            step_b = st
    if step_sq is not None:
        x = float(step_sq)
        return x, x
    try:
        wmin = min(float(w_ft), float(d_ft))
    except Exception:
        wmin = 1.0
    if step_b is None and tareas:
        try:
            sep0 = float(tareas[0][1])
        except Exception:
            sep0 = 1.0
        step_b = _paso_entre_lineas_distrib_mm(sep0, wmin, nb, d_est_mm, d_long_mm)
    if step_a is None:
        step_a = 0.0
    if step_b is None:
        step_b = 0.0
    return float(step_a), float(step_b)


def _offset_add_segunda_capa_mm(etiqueta_t, step_a_mm, step_b_mm):
    """Receta A segunda capa suma paso B; receta B suma paso A; lados iguales suman ese paso."""
    if u"lados iguales" in etiqueta_t:
        return float(step_a_mm)
    if u"Cara A (ancho)" in etiqueta_t:
        return float(step_b_mm)
    if u"Cara B (alto)" in etiqueta_t:
        return float(step_a_mm)
    return 0.0


def _lineas_copias_eje_x_cara(face, axis_unit, linea_trasladada, n_barras, w_mm, d_est_mm, d_long_mm):
    """
    Tras la traslación inicial, copia la línea en sentido **−eje X** de la cara de proyección
    (unitario ⟂ eje columna en el plano); desplazamientos ``0, step, 2·step, …`` con
    ``step = L/(n-1)``.
    """
    if linea_trasladada is None:
        return []
    try:
        n = max(1, min(99, int(n_barras)))
    except Exception:
        n = 1
    if axis_unit is None:
        return [linea_trasladada]
    try:
        ax = axis_unit.Normalize()
    except Exception:
        return [linea_trasladada]
    if ax is None or ax.GetLength() < 1e-12:
        return [linea_trasladada]
    dir_x = _vector_distribucion_menos_eje_x_cara(face, ax)
    if dir_x is None:
        return [linea_trasladada]
    _, step_mm = _paso_propagacion_mm_cara_a(n, w_mm, d_est_mm, d_long_mm)
    out = []
    for k in range(n):
        try:
            v = dir_x.Multiply(_mm_to_ft(k * step_mm))
            ln_k = _trasladar_linea_según_vector(linea_trasladada, v)
            if ln_k is not None and ln_k.Length >= _MIN_LINE_LEN_FT:
                out.append(ln_k)
        except Exception:
            continue
    return out if out else [linea_trasladada]


def _trasladar_linea_según_vector(linea, vec_xyz):
    """Copia de ``linea`` desplazada ``vec_xyz`` (pies)."""
    if linea is None or vec_xyz is None:
        return None
    try:
        p0 = linea.GetEndPoint(0) + vec_xyz
        p1 = linea.GetEndPoint(1) + vec_xyz
        if p0.DistanceTo(p1) < _MIN_LINE_LEN_FT:
            return None
        return Line.CreateBound(p0, p1)
    except Exception:
        return None


def _es_instancia_columna_estructural(el):
    if el is None:
        return False
    try:
        st = getattr(el, "StructuralType", None)
        return st == StructuralType.Column
    except Exception:
        return False


def _separacion_par_caras_ft(ancla_a, ancla_b):
    """Distancia entre centros de cara a lo largo de la normal de la primera."""
    try:
        fa = ancla_a.get(u"face")
        ca = ancla_a.get(u"center")
        cb = ancla_b.get(u"center")
        if fa is None or ca is None or cb is None:
            return None
        n = fa.FaceNormal
        if n is None or n.GetLength() < 1e-12:
            return None
        n = n.Normalize()
        return abs(float((cb - ca).DotProduct(n)))
    except Exception:
        return None


def _pares_caras_opuestas(anchors):
    """Agrupa anclas en pares de caras con normales opuestas (~rectángulo)."""
    if not anchors:
        return []
    pairs = []
    used = [False] * len(anchors)
    for i in range(len(anchors)):
        if used[i]:
            continue
        ni = None
        try:
            fi = anchors[i][u"face"]
            ni = fi.FaceNormal.Normalize()
        except Exception:
            continue
        if ni is None:
            continue
        for j in range(i + 1, len(anchors)):
            if used[j]:
                continue
            try:
                nj = anchors[j][u"face"].FaceNormal.Normalize()
            except Exception:
                continue
            if nj is None:
                continue
            if float(ni.DotProduct(nj)) < -0.92:
                pairs.append((anchors[i], anchors[j]))
                used[i] = True
                used[j] = True
                break
    return pairs


def _pares_laterales_misma_geometria(par_a, par_b):
    """``True`` si los dos pares son el mismo par de caras."""
    if par_a is None or par_b is None:
        return False
    try:
        a0, a1 = par_a
        b0, b1 = par_b
        fa0, fa1 = a0[u"face"], a1[u"face"]
        fb0, fb1 = b0[u"face"], b1[u"face"]
        return (fa0 is fb0 and fa1 is fb1) or (fa0 is fb1 and fa1 is fb0)
    except Exception:
        return False


def _tareas_v2_ancho_canto_desde_pares(pairs, w_ft, d_ft, na, nb):
    """
    Lista de ``(par, sep_en_cara_ft, tipo_en_cara_ft, n_barras, etiqueta_ui)``.

    El **par** es la cara lateral donde se dibuja (mayor separación = Cara A, menor = B).
    ``sep_en_cara_ft`` / ``tipo_en_cara_ft`` son la **luz ortogonal** (el otro par): en esa cara
    la distribución de barras y la traslación ``(luz/2 − recub.)`` van en la dirección ⟂ al eje,
    no en la separación del propio par.

    Si **todos** los pares tienen la misma ``sep`` (lados iguales en planta), se usa ``nb`` y
    la misma luz para cada par (sección cuadrada o regular).
    """
    items = []
    for pr in pairs or []:
        try:
            a, b = pr
        except Exception:
            continue
        s = _separacion_par_caras_ft(a, b)
        if s is None or s < 1e-9:
            continue
        items.append((float(s), pr))
    if not items:
        return []
    try:
        wmaj = max(float(w_ft), float(d_ft))
        wmin = min(float(w_ft), float(d_ft))
    except Exception:
        wmaj, wmin = 1.0, 1.0
    items.sort(key=lambda t: (-t[0], id(t[1][0])))

    if len(items) >= 2:
        s_ref = float(items[0][0])
        iguales = True
        for s, _ in items[1:]:
            if abs(float(s) - s_ref) > float(_TOL_SECCION_CUADRADA_FT):
                iguales = False
                break
        if iguales:
            out_sq = []
            for s, pr in items:
                out_sq.append(
                    (pr, float(s), wmaj, nb, u"Cara B (lados iguales)")
                )
            return out_sq

    if len(items) == 1:
        s0, pr0 = items[0]
        return [(pr0, float(s0), wmaj, na, u"Cara A (ancho)")]
    s_hi, pr_hi = items[0]
    s_lo, pr_lo = items[-1]
    out = [(pr_hi, float(s_lo), wmin, na, u"Cara A (ancho)")]
    if not _pares_laterales_misma_geometria(pr_hi, pr_lo):
        out.append((pr_lo, float(s_hi), wmaj, nb, u"Cara B (alto)"))
    return out


def _proyectar_punto_a_cara_planar(face, pt):
    if face is None or pt is None:
        return None
    try:
        r = face.Project(pt)
        if r is None:
            return None
        return r.XYZPoint
    except Exception:
        return None


def _dist_punto_a_linea(pt, linea):
    try:
        p0 = linea.GetEndPoint(0)
        d = linea.GetEndPoint(1) - p0
        ln = float(d.GetLength())
        if ln < 1e-12:
            return 1e30
        du = d.Multiply(1.0 / ln)
        v = pt - p0
        para = du.Multiply(float(v.DotProduct(du)))
        perp = v - para
        return float(perp.GetLength())
    except Exception:
        return 1e30


def _lineas_paralelas_misma_infinita(la, lb, tol_dist):
    try:
        da = la.GetEndPoint(1) - la.GetEndPoint(0)
        db = lb.GetEndPoint(1) - lb.GetEndPoint(0)
        if da.GetLength() < 1e-12 or db.GetLength() < 1e-12:
            return False
        da = da.Normalize()
        db = db.Normalize()
        if abs(abs(float(da.DotProduct(db))) - 1.0) > 0.02:
            return False
        return _dist_punto_a_linea(la.GetEndPoint(0), lb) <= tol_dist
    except Exception:
        return False


def _agrupar_lineas_colineales(lineas, tol_dist=_TOL_LINE_GROUP_FT):
    if not lineas:
        return []
    grupos = []
    for ln in lineas:
        colocado = False
        for g in grupos:
            if g and _lineas_paralelas_misma_infinita(ln, g[0], tol_dist):
                g.append(ln)
                colocado = True
                break
        if not colocado:
            grupos.append([ln])
    return grupos


def _fundir_grupos(grupos):
    salida = []
    for g in grupos:
        if not g:
            continue
        m = _unificar_lineas_colineales(g, g[0])
        if m is not None and m.Length >= _MIN_LINE_LEN_FT:
            salida.append(m)
    return salida


def _fundir_grupos_con_capa(grupos, lineas_all, metas_capa):
    """
    Como ``_fundir_grupos`` pero devuelve ``(lineas, capas)`` con el índice de capa
    (``mult_capa`` 0-based del bucle de capas) por tramo fusionado: ``min`` en el grupo.
    """
    salida_lineas = []
    salida_capas = []
    for g in grupos:
        if not g:
            continue
        capas_en_g = []
        for ln in g:
            found = False
            for i, l in enumerate(lineas_all or []):
                if l is ln:
                    try:
                        ki = int(metas_capa[i])
                    except Exception:
                        ki = 0
                    capas_en_g.append(ki)
                    found = True
                    break
            if not found:
                capas_en_g.append(0)
        try:
            capa_grp = min(capas_en_g)
        except Exception:
            capa_grp = 0
        m = _unificar_lineas_colineales(g, g[0])
        if m is not None and m.Length >= _MIN_LINE_LEN_FT:
            salida_lineas.append(m)
            salida_capas.append(capa_grp)
    return salida_lineas, salida_capas


# Comentario en ModelCurve: primera capa (troceo por planos empalme). Capas 2+ llevan sufijo ``·Ln``.
COMMENT_COL_V2_EJE_CARA_CAPA1 = u"BIMTools_ColV2_EjeCara"


def _comentario_capa_model_curve_col_v2(capa_mult):
    """``capa_mult`` 0 = primera capa; 1 = segunda (``·L2``), etc."""
    try:
        k = int(capa_mult)
    except Exception:
        k = 0
    if k <= 0:
        return COMMENT_COL_V2_EJE_CARA_CAPA1
    return u"{0}·L{1}".format(COMMENT_COL_V2_EJE_CARA_CAPA1, k + 1)


def _crear_model_curve_y_sketch(document, ln, capa_mult=0):
    """``(model_curve_id, sketch_plane_id)`` o ``(None, None)``."""
    if document is None or ln is None:
        return None, None
    try:
        sp = sketch_plane_para_linea(document, ln)
        if sp is None:
            return None, None
        ln_use = ln
        try:
            pl = sp.GetPlane()
            if pl is not None:
                p0 = ln.GetEndPoint(0)
                p1 = ln.GetEndPoint(1)
                q0 = _project_point_to_plane(p0, pl)
                q1 = _project_point_to_plane(p1, pl)
                if (
                    q0 is not None
                    and q1 is not None
                    and q0.DistanceTo(q1) >= _MIN_LINE_LEN_FT
                ):
                    ln_use = Line.CreateBound(q0, q1)
        except Exception:
            pass
        mc = document.Create.NewModelCurve(ln_use, sp)
        if mc is None:
            return None, None
        sp_id = None
        try:
            sp_id = sp.Id
        except Exception:
            pass
        try:
            p_c = mc.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
            if p_c is not None and not p_c.IsReadOnly:
                p_c.Set(_comentario_capa_model_curve_col_v2(capa_mult))
        except Exception:
            pass
        return mc.Id, sp_id
    except Exception:
        return None, None


def _generar_lineas_v2_un_par(
    document,
    col,
    crv,
    p_end0,
    p_end1,
    par,
    sep_ft,
    tipo_dim_ft,
    n_barras,
    offset_mm,
    diam_estribo_mm,
    diam_longitudinal_mm,
    detalle,
    etiqueta_cara=u"Cara",
    offset_interior_extra_mm=0.0,
):
    """
    Líneas de eje en la cara dada por ``par``.
    ``sep_ft`` y ``tipo_dim_ft`` deben ser la **luz en la dirección de reparto en planta** (par
    **ortogonal** al de esta cara): paso entre copias y ``(luz/2 − recub.)`` en ``dir_bx``.
    ``offset_interior_extra_mm``: sumado solo al offset **hacia el interior** (normal a cara).
    """
    out = []
    if par is None or sep_ft is None or sep_ft < 1e-6:
        return out
    try:
        n = int(n_barras)
    except Exception:
        n = 1
    n = max(1, min(99, n))

    ancla_a, ancla_b = par
    face_w = ancla_a[u"face"]
    q0 = _proyectar_punto_a_cara_planar(face_w, p_end0)
    q1 = _proyectar_punto_a_cara_planar(face_w, p_end1)
    if q0 is None or q1 is None or q0.DistanceTo(q1) < _MIN_LINE_LEN_FT:
        detalle.append(
            u"{0}: proyección al plano de cara degenerada.".format(etiqueta_cara)
        )
        return out

    base_ln = Line.CreateBound(q0, q1)
    if base_ln is None:
        return out
    axis_pre = q1 - q0
    if axis_pre.GetLength() < 1e-12:
        return out
    axis_pre_u = axis_pre.Normalize()
    dir_bx = _direccion_traslacion_bxis_en_cara_respecto_eje(face_w, axis_pre_u)
    d_bx_mm = _mm_traslacion_bxis_ancho_menos_recubrimiento(
        sep_ft, offset_mm, tipo_dim_ft
    )
    if dir_bx is not None and d_bx_mm > 1e-6:
        try:
            v_bx = dir_bx.Multiply(_mm_to_ft(d_bx_mm))
            ln_bx = _trasladar_linea_según_vector(base_ln, v_bx)
            if ln_bx is not None:
                base_ln = ln_bx
        except Exception:
            pass
    try:
        p_ax0 = base_ln.GetEndPoint(0)
        p_ax1 = base_ln.GetEndPoint(1)
    except Exception:
        return out
    axis_on_face = p_ax1 - p_ax0
    if axis_on_face.GetLength() < 1e-12:
        return out
    axis_on_face = axis_on_face.Normalize()
    w_mm_face = _ancho_cara_proyeccion_mm(sep_ft, tipo_dim_ft)
    try:
        de_use = float(diam_estribo_mm if diam_estribo_mm is not None else 0.0)
        dl_use = float(
            diam_longitudinal_mm if diam_longitudinal_mm is not None else 0.0
        )
    except Exception:
        de_use, dl_use = 0.0, 0.0
    barras = _lineas_copias_eje_x_cara(
        face_w, axis_on_face, base_ln, n, w_mm_face, de_use, dl_use
    )
    estiron_mm = _mm_estiron_desde_start_si_fundacion(document, col)
    try:
        off_in = float(offset_mm or 0.0) + float(offset_interior_extra_mm or 0.0)
    except Exception:
        off_in = float(offset_mm or 0.0)
    for bn in barras:
        bn_in = _aplicar_offset_interior_cara(bn, face_w, off_in)
        if bn_in is None or bn_in.Length < _MIN_LINE_LEN_FT:
            continue
        if estiron_mm > 1e-6:
            bn_ex = _extender_linea_desde_start_column_curve(bn_in, crv, estiron_mm)
            if bn_ex is not None and bn_ex.Length >= _MIN_LINE_LEN_FT:
                bn_in = bn_ex
        bn_in = _acortar_linea_extremo_fin_column_curve(
            bn_in, crv, float(_ACORTE_EXTREMO_FIN_MM)
        )
        if bn_in is None or bn_in.Length < _MIN_LINE_LEN_FT:
            continue
        out.append(bn_in)
    return out


def ejecutar_v2_model_lines_cara_ancho(
    document,
    _uidocument,
    element_ids,
    n_barras_cara_ancho,
    n_barras_cara_alto,
    offset_mm,
    diam_estribo_mm=None,
    diam_longitudinal_mm=None,
    segunda_capa=False,
    tercera_capa=False,
):
    """
    Par más separado → cara A (``na``); el reparto usa la **luz del otro** par y ``min(w,d)``.
    Par más estrecho → B (``nb``) con luz del ancho y ``max(w,d)``. Dos opuestas por par.
    Capas interiores extra: offset interior += ``k`` × paso cara ortogonal (``k`` = 1 o 2).
    ``tercera_capa`` incluye segunda aunque el check de segunda esté apagado.

    Returns:
        ``(mensaje, ids_model_curve_eje, ids_model_curve_marcador_normal, ids_sketch_planes)``
    """
    vac = []
    if document is None:
        return u"No hay documento.", vac, vac, vac
    try:
        na = int(n_barras_cara_ancho)
    except Exception:
        na = 0
    try:
        nb = int(n_barras_cara_alto)
    except Exception:
        nb = 0
    na = max(1, min(99, na))
    nb = max(1, min(99, nb))
    try:
        de_mm = float(diam_estribo_mm if diam_estribo_mm is not None else 0.0)
        dl_mm = float(
            diam_longitudinal_mm if diam_longitudinal_mm is not None else 0.0
        )
    except Exception:
        de_mm, dl_mm = 0.0, 0.0

    if tercera_capa:
        capas_mult = [0, 1, 2]
    elif segunda_capa:
        capas_mult = [0, 1]
    else:
        capas_mult = [0]

    cols = filtrar_solo_structural_columns(document, element_ids or [])
    if not cols:
        return (
            u"No hay columnas estructurales en la selección.",
            vac,
            vac,
            vac,
        )

    lineas_tras_offset = []
    capas_por_linea = []
    detalle = []
    hubo_columna_sin_par_b = False

    for col in cols:
        if not _es_instancia_columna_estructural(col):
            detalle.append(u"Omitido (no es columna estructural).")
            continue
        crv = _curva_eje_para_proyeccion(col)
        if crv is None or not crv.IsBound:
            detalle.append(u"Columna sin curva de ubicación válida.")
            continue
        try:
            p_end0 = crv.GetEndPoint(0)
            p_end1 = crv.GetEndPoint(1)
        except Exception:
            detalle.append(u"No se leyeron extremos de la columna.")
            continue

        w_ft, d_ft = _read_width_depth_ft(document, col, crv)
        anchors = _lateral_face_anchors_from_columna(col)
        pairs = _pares_caras_opuestas(anchors)
        if len(pairs) < 1:
            detalle.append(u"Sin par de caras laterales reconocible.")
            continue
        tareas = _tareas_v2_ancho_canto_desde_pares(pairs, w_ft, d_ft, na, nb)
        if not tareas:
            detalle.append(u"No se clasificaron pares laterales por separación.")
            continue
        if len(tareas) < 2:
            hubo_columna_sin_par_b = True

        step_a_mm, step_b_mm = _pasos_cara_a_y_b_mm(
            tareas, w_ft, d_ft, na, nb, de_mm, dl_mm
        )
        for mult_capa in capas_mult:
            try:
                k = int(mult_capa)
            except Exception:
                k = 0
            if k < 0:
                k = 0
            for par_t, sep_t, tipo_t, n_t, etiqueta_t in tareas:
                try:
                    ac0, ac1 = par_t
                except Exception:
                    continue
                base_extra = _offset_add_segunda_capa_mm(
                    etiqueta_t, step_a_mm, step_b_mm
                )
                extra_mm = float(base_extra) * float(k)
                for iori, par_ord in enumerate(((ac0, ac1), (ac1, ac0))):
                    suf = u"{0} · opuesta {1}".format(etiqueta_t, iori + 1)
                    if k >= 1:
                        suf += u" · capa {0}".format(k + 1)
                    for _ln in _generar_lineas_v2_un_par(
                        document,
                        col,
                        crv,
                        p_end0,
                        p_end1,
                        par_ord,
                        float(sep_t),
                        float(tipo_t),
                        int(n_t),
                        float(offset_mm),
                        diam_estribo_mm,
                        diam_longitudinal_mm,
                        detalle,
                        etiqueta_cara=suf,
                        offset_interior_extra_mm=float(extra_mm),
                    ):
                        lineas_tras_offset.append(_ln)
                        capas_por_linea.append(k)

    if hubo_columna_sin_par_b:
        detalle.append(
            u"Una o más columnas sin segundo par lateral: Cara B omitida en ellas."
        )

    if not lineas_tras_offset:
        return (
            u"No se generaron líneas en caras laterales. " + u" ".join(detalle[:3]),
            vac,
            vac,
            vac,
        )

    lineas_tras_offset, capas_por_linea, n_dup_punto = _dedupe_lineas_mismo_punto_medio(
        lineas_tras_offset, capas_por_linea, tol_ft=None
    )
    if n_dup_punto > 0:
        detalle.append(
            u"Tramos duplicados por punto medio: {0} descartado(s).".format(n_dup_punto)
        )
    lineas_tras_offset, capas_por_linea, n_dup_pos = _dedupe_lineas_misma_posicion(
        lineas_tras_offset, capas_por_linea, tol_ft=None
    )
    if n_dup_pos > 0:
        detalle.append(
            u"Tramos duplicados por posición: {0} descartado(s).".format(n_dup_pos)
        )
    if not lineas_tras_offset:
        return (
            u"No quedaron líneas tras deduplicar por punto medio. " + u" ".join(detalle[:3]),
            vac,
            vac,
            vac,
        )

    grupos = _agrupar_lineas_colineales(lineas_tras_offset)
    fusionadas, fusion_capas = _fundir_grupos_con_capa(
        grupos, lineas_tras_offset, capas_por_linea
    )
    if not fusionadas:
        return (
            u"No hubo tramos fusionados.",
            vac,
            vac,
            vac,
        )

    try:
        d_nom_sonda = float(diam_longitudinal_mm) if diam_longitudinal_mm is not None else 0.0
    except Exception:
        d_nom_sonda = 0.0
    try:
        fusionadas, fusion_capas = aplicar_empotramiento_lineas_unificadas(
            document,
            fusionadas,
            list(element_ids or []),
            cols,
            float(_EMPOTRAMIENTO_PRUEBA_MM),
            diam_nominal_mm=d_nom_sonda if d_nom_sonda > 1e-9 else None,
            line_metas=fusion_capas,
        )
    except Exception:
        fusionadas = []
        fusion_capas = []
    if not fusionadas:
        return (
            u"Tramos inválidos o demasiado cortos tras sondeo/empotramiento.",
            vac,
            vac,
            vac,
        )
    if len(fusion_capas) < len(fusionadas):
        fusion_capas.extend([0] * (len(fusionadas) - len(fusion_capas)))
    elif len(fusion_capas) > len(fusionadas):
        fusion_capas = fusion_capas[: len(fusionadas)]

    ids_eje = []
    ids_marc = []
    ids_sp = []
    try:
        with Transaction(
            document, u"BIMTools — Columnas V2 líneas por separación cara A/B"
        ) as t:
            t.Start()
            try:
                for ln, cap_k in zip(fusionadas, fusion_capas):
                    mid, sid = _crear_model_curve_y_sketch(document, ln, capa_mult=cap_k)
                    if mid is not None:
                        ids_eje.append(mid)
                    if sid is not None:
                        ids_sp.append(sid)
                    mnorm = crear_marcador_normal_curva_eje(document, ln)
                    if mnorm is not None:
                        ids_marc.append(mnorm)
                        try:
                            el_m = document.GetElement(mnorm)
                            sp_m = el_m.SketchPlane if el_m is not None else None
                            if sp_m is not None:
                                ids_sp.append(sp_m.Id)
                        except Exception:
                            pass
                t.Commit()
            except Exception as ex:
                try:
                    t.RollBack()
                except Exception:
                    pass
                return (u"Error en transacción: {0}".format(ex), vac, vac, vac)
    except Exception as ex:
        return (u"Error: {0}".format(ex), vac, vac, vac)

    n_eje = len(fusionadas)
    n_marc = n_eje
    n_total_mc = len(ids_eje) + len(ids_marc)
    capas_txt = u""
    if tercera_capa:
        capas_txt = u" Capas 1–3 (offset + k·paso ortogonal)."
    elif segunda_capa:
        capas_txt = u" Capas 1–2 (+ paso ortogonal)."
    msg = (
        u"A/B por geometría (sep. mayor/menor); UI {0}/{1} barra(s).{5} "
        u"Ejes fusionados: {2}; marcadores normal: {3}; elementos modelo totales: {4}."
    ).format(
        na,
        nb,
        n_eje,
        n_marc,
        n_total_mc,
        capas_txt,
    )
    if detalle:
        msg += u" Avisos: " + u" ".join(detalle[:2])
    return msg, ids_eje, ids_marc, ids_sp
