# -*- coding: utf-8 -*-
"""
Troceo Columnas V2 — **solo** para ``ModelLine`` / ``ModelCurve`` usando ``SketchPlane`` de empalme.

Lógica **aislada** del troceo automático de vigas (``geometria_viga_cara_superior_detalle``,
``_punto_troceo_segmento_plano_empalme``, etc.): intersección línea–plano, segmentos nuevos
``Line.CreateBound``, sin traslape, sin empalme ni recortes posteriores.

Requisitos: curvas de trabajo como **líneas** acotadas (el flujo V2 de caras solo crea ``Line``).
"""

from Autodesk.Revit.DB import (
    BuiltInParameter,
    CurveElement,
    FilteredElementCollector,
    IntersectionResultArray,
    Line,
    LocationCurve,
    LocationPoint,
    ModelCurve,
    Plane,
    SketchPlane,
    XYZ,
)

# ~1 mm (pies internos)
_TOL_PIE_FT = 1.0 / 304.8
# Tramo mínimo entre cortes o respecto a extremos (evita segmentos degenerados)
_MIN_TRAMO_FT = max(4.0 / 304.8, _TOL_PIE_FT * 4.0)
# Margen desde extremos del segmento para aceptar un plano de empalme (~2 mm; antes ~4 mm rechazaba cortes en juntas).
_MARGEN_EXTREMO_CORTE_FT = max(2.0 / 304.8, _TOL_PIE_FT * 3.0)
# Bbox ampliada por columna: primero 5 m; si no hay coincidencias, 20 m (líneas en cara muy desplazadas vs BoundingBox del pilar).
_BBOX_SALTO_COLUMNAS_MM = (5000.0, 20000.0)
# Ejes demasiado cortos no se trocean (marcador normal ~300 mm y artefactos).
_MIN_LONGITUD_EJE_TROCEO_MM = 800.0
# Límite de líneas detalladas en el informe (evita diálogos enormes).
_MAX_LINEAS_DETALLE_DIAG = 35


try:
    from barras_bordes_losa_gancho_empotramiento import element_id_to_int
except Exception:

    def element_id_to_int(eid):
        if eid is None:
            return None
        try:
            return int(eid.Value)
        except Exception:
            try:
                return int(eid.IntegerValue)
            except Exception:
                return None


def _diag_ln(diag_list, text):
    try:
        diag_list.append(unicode(text))
    except Exception:
        diag_list.append(str(text))


def _fmt_xyz_mm(p):
    try:
        return u"({0:.1f},{1:.1f},{2:.1f})".format(
            float(p.X) * 304.8, float(p.Y) * 304.8, float(p.Z) * 304.8
        )
    except Exception:
        return u"?"


def _fmt_vec3(v):
    try:
        return u"({0:.4f},{1:.4f},{2:.4f})".format(float(v.X), float(v.Y), float(v.Z))
    except Exception:
        return u"?"


def _dist_signed_pt_plano(pt, plane):
    try:
        n = plane.Normal
        if n.GetLength() < 1e-12:
            return None
        n = n.Normalize()
        return float(n.DotProduct(pt - plane.Origin))
    except Exception:
        return None


def _es_model_curve(elemento):
    if elemento is None:
        return False
    if isinstance(elemento, ModelCurve):
        return True
    try:
        return elemento.GetType().Name == "ModelCurve"
    except Exception:
        return False


def _iter_model_curves(document):
    """
    Revit (API reciente) no admite ``OfClass(ModelCurve)`` en el colector; se usa
    ``CurveElement`` y se conservan instancias ``ModelCurve``.
    """
    if document is None:
        return
    try:
        for ce in FilteredElementCollector(document).OfClass(CurveElement):
            if ce is None:
                continue
            if isinstance(ce, ModelCurve):
                yield ce
                continue
            try:
                if ce.GetType().Name == "ModelCurve":
                    yield ce
            except Exception:
                pass
    except Exception:
        return


def _enteros_ids(ids):
    s = set()
    for eid in ids or []:
        try:
            s.add(int(eid.IntegerValue))
        except Exception:
            pass
    return s


def _interseccion_segmento_plano_algebraico(p0, p1, plane):
    """
    Intersección segmento–plano por álgebra (respaldo si la API devuelve resultado vacío
    o un ``SetComparisonResult`` poco fiable).
    ``p(t) = p0 + t (p1-p0)``, ``t`` en ``[0,1]`` (incluye plano en extremo de junta).
    """
    if plane is None or p0 is None or p1 is None:
        return None
    try:
        n = plane.Normal
        if n.GetLength() < 1e-12:
            return None
        n = n.Normalize()
        o = plane.Origin
        d = p1 - p0
        L = float(d.GetLength())
        if L < 1e-12:
            return None
        dist0 = float(n.DotProduct(p0 - o))
        dist1 = float(n.DotProduct(p1 - o))
        denom = dist0 - dist1
        tol = max(1e-12, L * 1e-10)
        if abs(denom) <= tol:
            return None
        t = dist0 / denom
        if t < -1e-6 or t > 1.0 + 1e-6:
            return None
        t_cl = min(1.0, max(0.0, float(t)))
        return p0 + d.Multiply(t_cl)
    except Exception:
        return None


def _interseccion_segmento_plano(p0, p1, plane):
    """Punto de corte del segmento ``[p0,p1]`` con ``plane``, o ``None``."""
    if plane is None or p0 is None or p1 is None:
        return None
    try:
        seg = Line.CreateBound(p0, p1)
        arr = IntersectionResultArray()
        seg.Intersect(plane, arr)
        if arr is not None and arr.Size > 0:
            try:
                pt_api = arr.get_Item(0).XYZPoint
                s_chk = _parametro_distancia_desde_p0(p0, p1, pt_api)
                L_chk = float(p0.DistanceTo(p1))
                if s_chk is not None and L_chk > 1e-12:
                    if -_TOL_PIE_FT <= s_chk <= L_chk + _TOL_PIE_FT:
                        return pt_api
            except Exception:
                pass
        return _interseccion_segmento_plano_algebraico(p0, p1, plane)
    except Exception:
        return _interseccion_segmento_plano_algebraico(p0, p1, plane)


def _parametro_distancia_desde_p0(p0, p1, pt):
    """Distancia con signo desde ``p0`` hacia ``p1`` hasta la proyección de ``pt`` en la recta."""
    try:
        v = p1 - p0
        L = float(v.GetLength())
        if L < 1e-12:
            return None
        u = v.Normalize()
        return float((pt - p0).DotProduct(u))
    except Exception:
        return None


def _comentario_model_curve(mc):
    """Comentarios de instancia; varios accesos por idioma / caché de parámetro en Revit."""
    if mc is None:
        return u""
    try:
        p = mc.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if p is not None:
            s = p.AsString()
            if not s:
                s = p.AsValueString()
            if s:
                return s
    except Exception:
        pass
    for nm in (u"Comments", u"Comentarios", u"Commentaires"):
        try:
            p = mc.LookupParameter(nm)
            if p is None:
                continue
            s = p.AsString()
            if not s:
                s = p.AsValueString()
            if s:
                return s
        except Exception:
            continue
    return u""


def _es_marcador_bimtools_excluir_troceo(mc):
    """
    Marcadores BIMTools de normal (empalme, eje, etc.): **nunca** trocear.
    Las únicas curvas válidas son ejes en cara V2 (comentario ``BIMTools_ColV2_EjeCara``…).
    """
    try:
        s = _comentario_model_curve(mc)
        try:
            s_cmp = unicode(s)
        except Exception:
            s_cmp = str(s)
        if u"BIMTools_Normal" in s_cmp:
            return True
    except Exception:
        pass
    try:
        nm = mc.Name or u""
        try:
            nm_cmp = unicode(nm)
        except Exception:
            nm_cmp = str(nm)
        if u"BIMTools_Normal" in nm_cmp:
            return True
    except Exception:
        pass
    try:
        pm = mc.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
        if pm is not None:
            ms = pm.AsString()
            if not ms:
                ms = pm.AsValueString()
            if ms and ms.strip().lower() == u"n":
                return True
    except Exception:
        pass
    return False


def _es_model_curve_col_v2_eje_cara_para_troceo(mc):
    """
    Solo líneas de **proyección de eje a caras** (``geometria_columnas_v2_caras``): comentario
    ``BIMTools_ColV2_EjeCara`` o ``·L2`` / ``·L3``. No marcadores ni otras ``ModelLine``.
    """
    if mc is None:
        return False
    if _es_marcador_bimtools_excluir_troceo(mc):
        return False
    s = _comentario_model_curve(mc)
    try:
        s_cmp = unicode(s).strip()
    except Exception:
        s_cmp = str(s).strip()
    if not s_cmp or u"BIMTools_ColV2_EjeCara" not in s_cmp:
        return False
    return True


def _geometry_curve_es_linea_acotada(mc):
    """``True`` si ``GeometryCurve`` es ``Line`` acotada (ejes Col V2; no arcos ni nulos)."""
    try:
        crv = mc.GeometryCurve
        if crv is None or not crv.IsBound:
            return False
        if isinstance(crv, Line):
            return True
        try:
            return crv.GetType().Name == u"Line"
        except Exception:
            return False
    except Exception:
        return False


def _filtrar_element_ids_model_curve_col_v2_eje_cara(document, id_list):
    """
    Candidatos a troceo desde **``_model_line_ids``** (solo ``ids_eje`` de Colocar, sin marcadores V2).

    - Si el comentario Col V2 es legible: se usa.
    - Si no (API Revit), se confía en el origen: ``ModelCurve`` + ``Line`` acotada y no marcador BIMTools.
      No aplica a barridos por documento/bbox (ahí sigue siendo obligatorio el comentario).
    """
    out = []
    if document is None or not id_list:
        return out
    for eid in id_list:
        try:
            el = document.GetElement(eid)
        except Exception:
            el = None
        if el is None or not _es_model_curve(el):
            continue
        if _es_marcador_bimtools_excluir_troceo(el):
            continue
        if _es_model_curve_col_v2_eje_cara_para_troceo(el):
            out.append(eid)
            continue
        if _geometry_curve_es_linea_acotada(el):
            out.append(eid)
    return out


def _lista_cajas_ampliadas_por_columna(document, column_element_ids, margin_ft):
    """
    Una o más cajas AABB por pilar: primero ``BoundingBox(None)`` ampliado; si no hay,
    ``LocationCurve`` o ``LocationPoint`` del host con el mismo margen (pies).
    """
    boxes = []
    if document is None or not column_element_ids:
        return boxes
    try:
        m = float(margin_ft)
    except Exception:
        m = 5000.0 / 304.8
    for eid in column_element_ids:
        try:
            el = document.GetElement(eid)
        except Exception:
            el = None
        if el is None:
            continue
        bb = el.get_BoundingBox(None)
        if bb is not None:
            try:
                lo, hi = bb.Min, bb.Max
                boxes.append(
                    (
                        XYZ(float(lo.X) - m, float(lo.Y) - m, float(lo.Z) - m),
                        XYZ(float(hi.X) + m, float(hi.Y) + m, float(hi.Z) + m),
                    )
                )
                continue
            except Exception:
                pass
        loc = getattr(el, "Location", None)
        if isinstance(loc, LocationCurve):
            try:
                crv = loc.Curve
                if crv is not None and crv.IsBound:
                    p0 = crv.GetEndPoint(0)
                    p1 = crv.GetEndPoint(1)
                    boxes.append(
                        (
                            XYZ(
                                min(p0.X, p1.X) - m,
                                min(p0.Y, p1.Y) - m,
                                min(p0.Z, p1.Z) - m,
                            ),
                            XYZ(
                                max(p0.X, p1.X) + m,
                                max(p0.Y, p1.Y) + m,
                                max(p0.Z, p1.Z) + m,
                            ),
                        )
                    )
            except Exception:
                pass
            continue
        if isinstance(loc, LocationPoint):
            try:
                pt = loc.Point
                boxes.append(
                    (
                        XYZ(float(pt.X) - m, float(pt.Y) - m, float(pt.Z) - m),
                        XYZ(float(pt.X) + m, float(pt.Y) + m, float(pt.Z) + m),
                    )
                )
            except Exception:
                pass
    return boxes


def _aabb_segmento_solapa_caja(p0, p1, lo, hi):
    """``True`` si la caja AABB del segmento ``p0–p1`` interseca la caja ``[lo, hi]``."""
    try:
        sx0 = min(float(p0.X), float(p1.X))
        sy0 = min(float(p0.Y), float(p1.Y))
        sz0 = min(float(p0.Z), float(p1.Z))
        sx1 = max(float(p0.X), float(p1.X))
        sy1 = max(float(p0.Y), float(p1.Y))
        sz1 = max(float(p0.Z), float(p1.Z))
        if sx1 < float(lo.X) or sx0 > float(hi.X):
            return False
        if sy1 < float(lo.Y) or sy0 > float(hi.Y):
            return False
        if sz1 < float(lo.Z) or sz0 > float(hi.Z):
            return False
        return True
    except Exception:
        return False


def _collect_model_curves_near_columnas_v2(
    document, column_element_ids, margin_mm=None
):
    """
    ``ModelCurve`` eje cara ColV2 (**cualquier capa**) cuyo segmento solapa **alguna** caja ampliada por pilar.
    ``margin_mm``: ampliación de la AABB del pilar (por defecto 5 m).
    Excluye marcadores de normal BIMTools.
    """
    out = []
    n_intersect_bbox = 0
    n_fail_comment = 0
    if document is None or not column_element_ids:
        return out
    try:
        mm = float(margin_mm) if margin_mm is not None else 5000.0
    except Exception:
        mm = 5000.0
    margin_ft = mm / 304.8
    boxes = _lista_cajas_ampliadas_por_columna(document, column_element_ids, margin_ft)
    if not boxes:
        return out
    try:
        for mc in _iter_model_curves(document):
            if _es_marcador_bimtools_excluir_troceo(mc):
                continue
            try:
                crv = mc.GeometryCurve
            except Exception:
                crv = None
            if crv is None or not crv.IsBound:
                continue
            try:
                p0 = crv.GetEndPoint(0)
                p1 = crv.GetEndPoint(1)
            except Exception:
                continue
            if not any(
                _aabb_segmento_solapa_caja(p0, p1, lo, hi) for lo, hi in boxes
            ):
                continue
            n_intersect_bbox += 1
            if not _es_model_curve_col_v2_eje_cara_para_troceo(mc):
                n_fail_comment += 1
                continue
            try:
                out.append(mc.Id)
            except Exception:
                pass
    except Exception:
        pass
    # #region agent log
    if boxes and not out:
        try:
            import json
            import os
            import time

            _lp = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), os.pardir, "debug-c561be.log"
            )
            with open(_lp, "a") as _lf:
                _lf.write(
                    json.dumps(
                        {
                            u"sessionId": u"c561be",
                            u"hypothesisId": u"H3",
                            u"location": u"troceo_v2:bbox_collect_empty",
                            u"message": u"bbox collect returned 0 ColV2-comment curves",
                            u"data": {
                                u"n_intersect_bbox": n_intersect_bbox,
                                u"n_fail_comment_after_bbox": n_fail_comment,
                                u"margin_mm": mm,
                                u"n_boxes": len(boxes),
                            },
                            u"timestamp": int(time.time() * 1000),
                        },
                        ensure_ascii=False,
                    )
                    + u"\n"
                )
        except Exception:
            pass
    # #endregion
    return out


def _diag_recoleccion_bbox(document, column_element_ids, margin_mm=5000.0):
    """Texto breve si el respaldo bbox no devolvió curvas."""
    lines = []
    try:
        mm = float(margin_mm)
    except Exception:
        mm = 5000.0
    m_ft = mm / 304.8
    bx = _lista_cajas_ampliadas_por_columna(document, column_element_ids, m_ft)
    lines.append(
        u"    Cajas ampliadas (~{0:.0f} mm): {1} de {2} Id columnas.".format(
            mm, len(bx), len(column_element_ids or [])
        )
    )
    try:
        n_mc = sum(1 for _x in _iter_model_curves(document))
        lines.append(u"    ModelCurve en documento (aprox.): {0}.".format(n_mc))
    except Exception as ex:
        lines.append(u"    Conteo ModelCurve: error ({0}).".format(ex))
    return u"\n".join(lines)


def _collect_model_curve_ids_col_v2_desde_documento(document):
    """
    Respaldos: ``ModelCurve`` eje cara V2, **todas las capas** (``geometria_columnas_v2_caras``).
    Solo se usa si la ventana no aportó IDs en ``_model_line_ids``.
    """
    out = []
    if document is None:
        return out
    try:
        for mc in _iter_model_curves(document):
            try:
                if _es_model_curve_col_v2_eje_cara_para_troceo(mc):
                    out.append(mc.Id)
            except Exception:
                continue
    except Exception:
        pass
    return out


def _planos_desde_sketch_planes(document, sketch_plane_ids):
    """``list`` de ``(ElementId, Plane)`` válidos."""
    out = []
    for sid in sketch_plane_ids or []:
        try:
            el = document.GetElement(sid)
        except Exception:
            el = None
        if el is None or not isinstance(el, SketchPlane):
            continue
        try:
            pl = el.GetPlane()
        except Exception:
            pl = None
        if pl is None:
            continue
        out.append((sid, pl))
    return out


def _leer_marca_y_comentarios_texto(elemento):
    """Valores de Marca y Comentarios antes de borrar el elemento."""
    mark_s = None
    comm_s = None
    if elemento is None:
        return mark_s, comm_s
    try:
        pm = elemento.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
        if pm is not None:
            mark_s = pm.AsString()
            if not mark_s:
                mark_s = pm.AsValueString()
    except Exception:
        pass
    try:
        pc = elemento.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if pc is not None:
            comm_s = pc.AsString()
            if not comm_s:
                comm_s = pc.AsValueString()
    except Exception:
        pass
    return mark_s, comm_s


def _aplicar_marca_y_comentarios_texto(elemento, mark_s, comm_s):
    if elemento is None:
        return
    try:
        if mark_s:
            pd = elemento.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
            if pd is not None and not pd.IsReadOnly:
                pd.Set(mark_s)
    except Exception:
        pass
    try:
        if comm_s:
            pd = elemento.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
            if pd is not None and not pd.IsReadOnly:
                pd.Set(comm_s)
    except Exception:
        pass


def trocear_model_lines_con_planos_sketch_v2(
    document,
    model_curve_ids,
    sketch_plane_ids,
    column_ids_for_fallback=None,
):
    """
    Trocea cada ``ModelCurve`` cuyo ``GeometryCurve`` sea ``Line`` donde cruce alguno de los
    planos de los ``SketchPlane`` indicados.

    ``column_ids_for_fallback``: si no hay candidatos desde la ventana, se buscan en el documento
    únicamente ``ModelCurve`` con comentario **BIMTools_ColV2_EjeCara** (todas las capas), opcionalmente
    acotadas por bbox de pilares. No se trocean marcadores ni otras líneas de modelo.

    Returns:
        ``(mensaje_unicode, ids_nuevos_model_curve, ids_eliminados, diagnostico_multilinea)``
    """
    ids_nuevos = []
    ids_eliminados = []
    diag = []
    if document is None:
        _diag_ln(diag, u"[1] Documento nulo.")
        return u"No hay documento.", ids_nuevos, ids_eliminados, u"\n".join(diag)

    planos_info = _planos_desde_sketch_planes(document, sketch_plane_ids)
    _diag_ln(
        diag,
        u"[2] SketchPlane Id recibidos: {0}; válidos (GetPlane): {1}".format(
            len(sketch_plane_ids or []),
            len(planos_info),
        ),
    )
    for i, (sid, pl) in enumerate(planos_info):
        try:
            _diag_ln(
                diag,
                u"    Plano {0}: Id={1}, orig_mm={2}, n={3}".format(
                    i + 1,
                    element_id_to_int(sid),
                    _fmt_xyz_mm(pl.Origin),
                    _fmt_vec3(pl.Normal),
                ),
            )
        except Exception as ex:
            _diag_ln(diag, u"    Plano {0}: error al leer ({1})".format(i + 1, ex))
    if not planos_info:
        _diag_ln(diag, u"[3] Revise que los elementos sean SketchPlane y existan en el documento.")
        return (
            u"Troceo V2: no hay planos de empalme (SketchPlane) válidos.",
            ids_nuevos,
            ids_eliminados,
            u"\n".join(diag),
        )

    entrada = list(model_curve_ids or [])
    mc_ids = []
    prefijo_origen = u""
    if entrada:
        n_ventana_antes = len(entrada)
        mc_ids = _filtrar_element_ids_model_curve_col_v2_eje_cara(document, list(entrada))
        if n_ventana_antes and not mc_ids:
            prefijo_origen = (
                u"lista ventana: 0 eje cara V2 de {0} Id; ".format(n_ventana_antes)
            )
    if entrada and mc_ids:
        origen_curvas = u"lista ventana (_model_line_ids)"
    elif not mc_ids:
        mc_ids = _collect_model_curve_ids_col_v2_desde_documento(document)
        if mc_ids:
            origen_curvas = (
                prefijo_origen + u"respaldo: comentario BIMTools_ColV2_EjeCara (todas las capas)"
            )
        elif column_ids_for_fallback:
            _bbox_mm_usado = None
            for _mm_try in _BBOX_SALTO_COLUMNAS_MM:
                mc_try = _collect_model_curves_near_columnas_v2(
                    document, column_ids_for_fallback, margin_mm=_mm_try
                )
                if mc_try:
                    mc_ids = mc_try
                    _bbox_mm_usado = _mm_try
                    break
            if mc_ids:
                origen_curvas = prefijo_origen + u"respaldo: bbox columnas (Id pilares: {0}, margen {1:.0f} mm)".format(
                    len(column_ids_for_fallback),
                    float(_bbox_mm_usado or 5000.0),
                )
            else:
                origen_curvas = (
                    prefijo_origen
                    + u"respaldo bbox sin ej ColV2 (comentario BIMTools_ColV2_EjeCara en volumen de pilares)."
                )
                try:
                    for _mm_try in _BBOX_SALTO_COLUMNAS_MM:
                        _diag_ln(
                            diag,
                            _diag_recoleccion_bbox(
                                document, column_ids_for_fallback, margin_mm=_mm_try
                            ),
                        )
                except Exception:
                    pass
        else:
            origen_curvas = prefijo_origen + u"sin respaldo (faltan Id columnas para envolvente)"
    # #region agent log
    try:
        import json
        import os
        import time

        _lp = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), os.pardir, "debug-c561be.log"
        )
        with open(_lp, "a") as _lf:
            _lf.write(
                json.dumps(
                    {
                        u"sessionId": u"c561be",
                        u"hypothesisId": u"H0",
                        u"location": u"troceo_v2:candidatas_resueltas",
                        u"message": u"resolved troceo candidates",
                        u"n_entrada": len(entrada),
                        u"n_mc_ids": len(mc_ids),
                        u"origen_curvas": origen_curvas,
                        u"timestamp": int(time.time() * 1000),
                    },
                    ensure_ascii=False,
                )
                + u"\n"
            )
    except Exception:
        pass
    # #endregion
    _diag_ln(diag, u"[4] IDs en _model_line_ids (ventana): {0}".format(len(entrada)))
    _diag_ln(diag, u"[5] Curvas candidatas: {0}; origen: {1}.".format(len(mc_ids), origen_curvas))
    if entrada[:8]:
        _ids_preview = []
        for x in entrada[:8]:
            try:
                _ids_preview.append(unicode(element_id_to_int(x)))
            except Exception:
                _ids_preview.append(str(element_id_to_int(x)))
        _diag_ln(diag, u"    Primeros Id ventana: {0}".format(u", ".join(_ids_preview)))
    if not mc_ids:
        _diag_ln(
            diag,
            u"[6] Sin curvas eje cara V2. Revisar comentario BIMTools_ColV2_EjeCara, "
            u"colocar líneas en esta sesión o bbox cerca de columnas.",
        )
        return (
            u"Troceo V2: no hay líneas de modelo para cortar.",
            ids_nuevos,
            ids_eliminados,
            u"\n".join(diag),
        )

    n_procesadas = 0
    n_troceadas = 0
    n_tramos = 0
    n_logged = 0
    n_no_mc = 0
    n_no_line = 0

    for eid in mc_ids:
        eid_i = element_id_to_int(eid)
        try:
            el = document.GetElement(eid)
        except Exception:
            el = None
        if el is None:
            n_no_mc += 1
            if n_logged < _MAX_LINEAS_DETALLE_DIAG:
                _diag_ln(diag, u"  Id {0}: elemento nulo.".format(eid_i))
                n_logged += 1
            continue
        if not _es_model_curve(el):
            try:
                tn = el.GetType().Name
            except Exception:
                tn = u"?"
            n_no_mc += 1
            if n_logged < _MAX_LINEAS_DETALLE_DIAG:
                _diag_ln(diag, u"  Id {0}: tipo «{1}» (se espera ModelCurve).".format(eid_i, tn))
                n_logged += 1
            continue
        n_procesadas += 1
        try:
            crv = el.GeometryCurve
        except Exception as ex_gc:
            crv = None
            if n_logged < _MAX_LINEAS_DETALLE_DIAG:
                _diag_ln(diag, u"  Id {0}: GeometryCurve error: {1}".format(eid_i, ex_gc))
                n_logged += 1
            continue
        if crv is None or not crv.IsBound:
            if n_logged < _MAX_LINEAS_DETALLE_DIAG:
                _diag_ln(diag, u"  Id {0}: curva no acotada o nula.".format(eid_i))
                n_logged += 1
            continue
        es_linea = isinstance(crv, Line)
        if not es_linea:
            try:
                es_linea = crv.GetType().Name == "Line"
            except Exception:
                es_linea = False
        if not es_linea:
            n_no_line += 1
            try:
                tcn = crv.GetType().Name
            except Exception:
                tcn = u"?"
            if n_logged < _MAX_LINEAS_DETALLE_DIAG:
                _diag_ln(
                    diag,
                    u"  Id {0}: GeometryCurve tipo «{1}» (solo Line está soportada).".format(eid_i, tcn),
                )
                n_logged += 1
            continue
        try:
            p0 = crv.GetEndPoint(0)
            p1 = crv.GetEndPoint(1)
        except Exception:
            continue
        try:
            L = float(p0.DistanceTo(p1))
        except Exception:
            continue
        if L * 304.8 < _MIN_LONGITUD_EJE_TROCEO_MM:
            if n_logged < _MAX_LINEAS_DETALLE_DIAG:
                _diag_ln(
                    diag,
                    u"  Id {0}: longitud {1:.0f} mm < mínimo eje troceo ({2:.0f} mm).".format(
                        eid_i, L * 304.8, _MIN_LONGITUD_EJE_TROCEO_MM
                    ),
                )
                n_logged += 1
            continue
        if L < 2.0 * _MARGEN_EXTREMO_CORTE_FT:
            if n_logged < _MAX_LINEAS_DETALLE_DIAG:
                _diag_ln(
                    diag,
                    u"  Id {0}: longitud {1:.0f} mm < mínimo para trocear.".format(eid_i, L * 304.8),
                )
                n_logged += 1
            continue

        try:
            udir = (p1 - p0).Normalize()
        except Exception:
            continue

        me = _MARGEN_EXTREMO_CORTE_FT
        puntos_corte = []
        res_planes = []
        for ip, (_unused_sid, pl) in enumerate(planos_info):
            pt = _interseccion_segmento_plano(p0, p1, pl)
            d0 = _dist_signed_pt_plano(p0, pl)
            d1 = _dist_signed_pt_plano(p1, pl)
            d0m = d0 * 304.8 if d0 is not None else None
            d1m = d1 * 304.8 if d1 is not None else None
            if pt is None:
                res_planes.append(u"P{0}: sin intersección (d0={1} d1={2} mm)".format(ip + 1, d0m, d1m))
                continue
            s = _parametro_distancia_desde_p0(p0, p1, pt)
            if s is None:
                res_planes.append(u"P{0}: intersección sin parámetro s".format(ip + 1))
                continue
            if s < -_TOL_PIE_FT or s > L + _TOL_PIE_FT:
                res_planes.append(
                    u"P{0}: s fuera del tramo ({1:.1f} mm en L={2:.0f} mm)".format(
                        ip + 1, s * 304.8, L * 304.8
                    )
                )
                continue
            s_use = float(s)
            if s_use <= me:
                s_use = float(me)
            elif s_use >= L - me:
                s_use = float(L - me)
            if L <= 2.0 * me + _MIN_TRAMO_FT:
                res_planes.append(
                    u"P{0}: tramo {1:.0f} mm demasiado corto para margen de corte.".format(
                        ip + 1, L * 304.8
                    )
                )
                continue
            if s_use <= me + 1e-10 or s_use >= L - me - 1e-10:
                res_planes.append(
                    u"P{0}: no cabe corte interior (L={1:.0f} mm, marg {2:.1f} mm).".format(
                        ip + 1, L * 304.8, me * 304.8
                    )
                )
                continue
            smm = s_use * 304.8
            try:
                pt_use = p0 + udir.Multiply(s_use)
            except Exception:
                pt_use = pt
            puntos_corte.append((s_use, pt_use))
            res_planes.append(u"P{0}: corte válido s={1:.1f} mm".format(ip + 1, smm))

        if not puntos_corte:
            if n_logged < _MAX_LINEAS_DETALLE_DIAG:
                _diag_ln(
                    diag,
                    u"  Id {0}: sin punto interior de corte. {1}".format(eid_i, u" | ".join(res_planes)),
                )
                n_logged += 1
            continue

        puntos_corte.sort(key=lambda t: t[0])
        s_param = []
        last_s = None
        for s, _pt in puntos_corte:
            if last_s is not None and abs(s - last_s) < _TOL_PIE_FT:
                continue
            if last_s is not None and abs(s - last_s) < _MIN_TRAMO_FT:
                s = 0.5 * (s + last_s)
                s_param[-1] = s
                last_s = s
                continue
            s_param.append(s)
            last_s = s

        if not s_param:
            if n_logged < _MAX_LINEAS_DETALLE_DIAG:
                _diag_ln(diag, u"  Id {0}: s_param vacío tras deduplicar.".format(eid_i))
                n_logged += 1
            continue

        try:
            sp_orig = el.SketchPlane
        except Exception:
            sp_orig = None
        if sp_orig is None:
            if n_logged < _MAX_LINEAS_DETALLE_DIAG:
                _diag_ln(diag, u"  Id {0}: SketchPlane nulo en ModelCurve.".format(eid_i))
                n_logged += 1
            continue

        vertices = [p0]
        for s in s_param:
            vertices.append(p0 + udir.Multiply(s))
        vertices.append(p1)

        tramos = []
        for i in range(len(vertices) - 1):
            a = vertices[i]
            b = vertices[i + 1]
            if a.DistanceTo(b) < me:
                continue
            seg = Line.CreateBound(a, b)
            tramos.append(seg)

        if len(tramos) < 2:
            if n_logged < _MAX_LINEAS_DETALLE_DIAG:
                _diag_ln(
                    diag,
                    u"  Id {0}: solo {1} tramo(s) tras partir (se necesitan ≥2).".format(eid_i, len(tramos)),
                )
                n_logged += 1
            continue

        mark_s, comm_s = _leer_marca_y_comentarios_texto(el)

        try:
            document.Delete(el.Id)
        except Exception as ex_del:
            if n_logged < _MAX_LINEAS_DETALLE_DIAG:
                _diag_ln(diag, u"  Id {0}: Delete falló: {1}".format(eid_i, ex_del))
                n_logged += 1
            continue
        ids_eliminados.append(eid)
        n_troceadas += 1
        n_creados_mc = 0
        for seg in tramos:
            try:
                mc = document.Create.NewModelCurve(seg, sp_orig)
            except Exception as ex_mc:
                if n_logged < _MAX_LINEAS_DETALLE_DIAG:
                    _diag_ln(diag, u"  Id {0}: NewModelCurve error: {1}".format(eid_i, ex_mc))
                    n_logged += 1
                mc = None
            if mc is not None:
                try:
                    _aplicar_marca_y_comentarios_texto(mc, mark_s, comm_s)
                except Exception:
                    pass
                try:
                    ids_nuevos.append(mc.Id)
                except Exception:
                    pass
                n_tramos += 1
                n_creados_mc += 1
        if n_logged < _MAX_LINEAS_DETALLE_DIAG:
            _diag_ln(
                diag,
                u"  Id {0}: TROCEADO → {1} tramo(s) nuevos.".format(eid_i, n_creados_mc),
            )
            n_logged += 1

    _diag_ln(
        diag,
        u"[7] Resumen: ModelCurve consideradas={0}, troceadas={1}, tramos creados={2}, "
        u"no-ModelCurve/otras={3}, no-Line={4}.".format(
            n_procesadas,
            n_troceadas,
            n_tramos,
            n_no_mc,
            n_no_line,
        ),
    )

    msg = u"Troceo V2: {0} línea(s) revisada(s), {1} troceada(s), {2} tramo(s) nuevos, {3} plano(s).".format(
        n_procesadas,
        n_troceadas,
        n_tramos,
        len(planos_info),
    )
    if n_troceadas == 0:
        msg = (
            u"Troceo V2: ningún corte aplicado. Revise el cuadro «Diagnóstico troceo V2»."
        )
    diag_txt = u"\n".join(diag)
    return msg, ids_nuevos, ids_eliminados, diag_txt


def trocear_model_lines_con_planos_sketch_v2_en_transaccion(
    document, model_curve_ids, sketch_plane_ids, column_ids_for_fallback=None
):
    """
    Envuelve :func:`trocear_model_lines_con_planos_sketch_v2` en una única ``Transaction``.
    """
    from Autodesk.Revit.DB import Transaction

    if document is None:
        return u"No hay documento.", [], [], u""

    t = Transaction(document, u"BIMTools — Troceo V2 planos empalme")
    t.Start()
    try:
        msg, nuevos, viejos, diag = trocear_model_lines_con_planos_sketch_v2(
            document,
            model_curve_ids,
            sketch_plane_ids,
            column_ids_for_fallback,
        )
        t.Commit()
        return msg, nuevos, viejos, diag
    except Exception as ex:
        try:
            t.RollBack()
        except Exception:
            pass
        try:
            err = ex.Message
        except Exception:
            err = str(ex)
        try:
            _dex = unicode(ex)
        except Exception:
            _dex = str(ex)
        return u"Troceo V2: error — {0}".format(err), [], [], _dex
