# -*- coding: utf-8 -*-
"""
Rebar longitudinal a partir de una ``Line`` en la cara inferior de fundación.

``CreateFromCurvesAndShape`` usa por defecto el ``RebarShape`` existente nombrado **«03»**
(si existe en el proyecto); si no, la primera forma Standard/SimpleLine.
Para la U con forma nombrada (p. ej. «03»), ``CreateFromCurvesAndShape`` recibe los **tres**
tramos rectos de la U (curvas sin ganchos en la lista, según API). ``startHook``/``endHook``
en ``null`` no fuerzan un ``RebarHookType`` en la creación. Se prueba orientación unificada
con ``XYZ.BasisZ`` y la sobrecarga con ángulos 0 / sin end treatment.
``CreateFromCurves`` solo con ``useExistingShapeIfPossible=True`` y ``createNewShape=False``
para no crear formas nuevas. Si falla o los ganchos son asimétricos, ``CreateFromCurves`` con
``ElementId`` por extremo (incl. ``InvalidElementId`` sin gancho).
"""

import clr
import System

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    BuiltInParameter,
    Curve,
    CurveLoop,
    ElementId,
    FilteredElementCollector,
    Line,
    Plane,
    StorageType,
    XYZ,
)
from Autodesk.Revit.DB.Structure import (
    Rebar,
    RebarHookOrientation,
    RebarHookType,
    RebarShape,
    RebarStyle,
)

from System.Collections.Generic import List

# Nombre visible del gancho en el proyecto (coincidencia exacta con el tipo mostrado).
HOOK_GANCHO_90_STANDARD_NAME = u"Standard - 90 deg."
# Forma de armadura existente en el proyecto (no crear una nueva vía API).
REBAR_SHAPE_NOMBRE_DEFECTO = u"03"

# Longitud mínima de barra (pies) para no enviar geometría degenerada a la API.
_MIN_CURVE_LENGTH_FT = 1.0 / 304.8 * 5.0  # ~5 mm

# Caché de RebarShape y RebarHookType por documento (clave: doc.Title).
# Se invalida si el título del documento cambia (cambio de proyecto).
_SHAPE_CACHE = {}   # {(doc_title, shape_name): RebarShape | None}
_HOOK_CACHE = {}    # {(doc_title, hook_name): RebarHookType | None}
_SHAPE_LINEA_SIMPLE_CACHE = {}  # {doc_title: RebarShape | None}

# Caché de la combinación ganadora de parámetros para CreateFromCurvesAndShape y
# CreateFromCurves por documento. Una vez que una barra se crea con éxito se guarda
# (normal_key, so, eo, use_ex, create_new) y se prueba primero en la siguiente barra.
# Esto evita decenas de llamadas fallidas a la API en barras 2, 3, … de la misma ejecución.
_REBAR_AND_SHAPE_WIN = {}   # {doc_title: (nv_key_tuple, hook_o_val)} — para AndShape
_REBAR_MIN_WIN = {}         # {doc_title: (nv_key_tuple, so_val, eo_val)} — para minimo


def _doc_title_safe(document):
    try:
        return document.Title
    except Exception:
        return None


def _hook_compare_string(value):
    if value is None:
        return u""
    try:
        t = unicode(value)
    except Exception:
        try:
            t = System.Convert.ToString(value)
        except Exception:
            return u""
    try:
        return t.replace(u"\u00A0", u" ").strip()
    except Exception:
        return u""


def rebar_hook_type_display_name(ht):
    if ht is None:
        return u""
    try:
        n = unicode(getattr(ht, "Name", None) or u"").strip()
        if n:
            return n
    except Exception:
        pass
    for bip_name in (u"SYMBOL_NAME_PARAM", u"ALL_MODEL_TYPE_NAME"):
        try:
            bip = getattr(BuiltInParameter, bip_name, None)
            if bip is None:
                continue
            p = ht.get_Parameter(bip)
            if p is None or not p.HasValue:
                continue
            if p.StorageType == StorageType.String:
                s = unicode(p.AsString() or u"").strip()
                if s:
                    return s
        except Exception:
            continue
    return u""


def rebar_shape_display_name(sh):
    """Nombre visible del ``RebarShape`` (``.Name`` o parámetro de tipo)."""
    if sh is None:
        return u""
    try:
        n = unicode(getattr(sh, "Name", None) or u"").strip()
        if n:
            return n
    except Exception:
        pass
    for bip_name in (u"SYMBOL_NAME_PARAM", u"ALL_MODEL_TYPE_NAME"):
        try:
            bip = getattr(BuiltInParameter, bip_name, None)
            if bip is None:
                continue
            p = sh.get_Parameter(bip)
            if p is None or not p.HasValue:
                continue
            if p.StorageType == StorageType.String:
                s = unicode(p.AsString() or u"").strip()
                if s:
                    return s
        except Exception:
            continue
    return u""


def buscar_rebar_shape_por_nombre(document, shape_name):
    """Resuelve un ``RebarShape`` por nombre mostrado (coincidencia exacta, sin crear tipos)."""
    if document is None:
        return None
    target = _hook_compare_string(shape_name)
    if not target:
        return None
    try:
        doc_key = (document.Title, target)
    except Exception:
        doc_key = None
    if doc_key is not None and doc_key in _SHAPE_CACHE:
        return _SHAPE_CACHE[doc_key]
    try:
        shapes = list(FilteredElementCollector(document).OfClass(RebarShape))
    except Exception:
        if doc_key is not None:
            _SHAPE_CACHE[doc_key] = None
        return None
    result = None
    for sh in shapes:
        if _hook_compare_string(rebar_shape_display_name(sh)) == target:
            try:
                result = document.GetElement(sh.Id)
            except Exception:
                result = sh
            break
    if doc_key is not None:
        _SHAPE_CACHE[doc_key] = result
    return result


def buscar_rebar_hook_type_por_nombre(document, hook_name):
    if document is None:
        return None
    target = _hook_compare_string(hook_name)
    if not target:
        return None
    try:
        doc_key = (document.Title, target)
    except Exception:
        doc_key = None
    if doc_key is not None and doc_key in _HOOK_CACHE:
        return _HOOK_CACHE[doc_key]
    try:
        hook_types = list(FilteredElementCollector(document).OfClass(RebarHookType))
    except Exception:
        if doc_key is not None:
            _HOOK_CACHE[doc_key] = None
        return None
    result = None
    for ht in hook_types:
        if _hook_compare_string(rebar_hook_type_display_name(ht)) == target:
            try:
                result = document.GetElement(ht.Id)
            except Exception:
                result = ht
            break
    if doc_key is not None:
        _HOOK_CACHE[doc_key] = result
    return result


def _validar_linea_horizontal(line):
    if line is None or not isinstance(line, Line):
        return False
    try:
        L = float(line.Length)
    except Exception:
        return False
    if L < _MIN_CURVE_LENGTH_FT:
        return False
    for i in (0, 1):
        try:
            p = line.GetEndPoint(i)
            for a in (p.X, p.Y, p.Z):
                if abs(float(a)) > 1e15:
                    return False
        except Exception:
            return False
    return True


def _normal_planta_perpendicular_a_linea(line):
    """Normal al plano que contiene la barra horizontal (en XY): ⟂ a la tangente en planta."""
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        d = XYZ(p1.X - p0.X, p1.Y - p0.Y, 0.0)
        if d.GetLength() < 1e-12:
            return None
        t = d.Normalize()
        n = XYZ(-t.Y, t.X, 0.0)
        if n.GetLength() < 1e-12:
            return None
        return n.Normalize()
    except Exception:
        return None


def _tangente_unitaria_linea(line):
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        d = p1 - p0
        if d.GetLength() < 1e-12:
            return None
        return d.Normalize()
    except Exception:
        return None


def _orientacion_ganchos_unificada_segun_z_proyecto(line, norm_vec, eje_referencia_z=None):
    """
    Un solo ``RebarHookOrientation`` para **inicio y fin**, coherente con el plano de la
    barra (``norm``) y un eje de referencia en Z (por defecto **+Z** del proyecto): usa el
    sentido de ``T × N`` frente a ese eje (lateral del triedro de la sección) para elegir
    Left o Right.

    Para la **cara superior** de la fundación, pasar ``eje_referencia_z = XYZ(0,0,-1)``
    (Z reverso) para que los ganchos queden hacia el interior del host.

    La API interpreta Left/Right con la normal de la barra como «arriba»; aquí se fija
    una regla única ligada al eje de referencia para ambos ganchos.
    """
    if line is None or norm_vec is None:
        return RebarHookOrientation.Right
    t = _tangente_unitaria_linea(line)
    if t is None:
        return RebarHookOrientation.Right
    try:
        z_ref = eje_referencia_z if eje_referencia_z is not None else XYZ.BasisZ
        n = norm_vec.Normalize()
        lateral = t.CrossProduct(n)
        if lateral.GetLength() < 1e-12:
            return RebarHookOrientation.Right
        lateral = lateral.Normalize()
        if float(lateral.DotProduct(z_ref)) >= 0.0:
            return RebarHookOrientation.Right
        return RebarHookOrientation.Left
    except Exception:
        return RebarHookOrientation.Right


def _norm_desde_eje_v_cara_perpendicular_a_barra(line, marco):
    """
    Vector ``XYZ`` para la sobrecarga ``norm`` de ``CreateFromCurves*``, a partir del
    **eje V** (Y paramétrico de la cara, ``BasisY``), **invertido** (``-V``), proyectado
    perpendicular a la tangente de la barra; luego se devuelve el **vector reverso**
    (``-w``) para que el conjunto no se reparta hacia fuera del host.

    Si ``-V`` es paralelo a la barra, se usa la proyección de U; si sigue degenerado, ``N × T``.
    Cada resultado unitario se niega antes de devolverlo.
    """
    if line is None or marco is None or len(marco) < 4:
        return None
    t = _tangente_unitaria_linea(line)
    if t is None:
        return None
    _o, u_ax, v_ax, n_ax = marco[0], marco[1], marco[2], marco[3]

    def _proyecta_ortogonal(a_vec):
        if a_vec is None:
            return None
        try:
            a = a_vec.Normalize()
            w = a - t.Multiply(a.DotProduct(t))
        except Exception:
            return None
        if w.GetLength() < 1e-10:
            return None
        try:
            return w.Normalize()
        except Exception:
            return None

    # Base: eje Y de la cara en sentido contrario al paramétrico (-V), proyectado ⟂ barra.
    v_menos = None
    if v_ax is not None:
        try:
            v_menos = v_ax.Negate()
        except Exception:
            v_menos = None
    w = _proyecta_ortogonal(v_menos)
    if w is None and u_ax is not None:
        w = _proyecta_ortogonal(u_ax)
    if w is None and n_ax is not None:
        try:
            w = t.CrossProduct(n_ax.Normalize())
            if w.GetLength() < 1e-10:
                return None
            w = w.Normalize()
        except Exception:
            return None
    if w is None:
        return None
    # Reverso del norm “geométrico” anterior: interior del host en lugar de exterior.
    try:
        return w.Negate()
    except Exception:
        return None


def _lista_normales_create_from_curves(
    line,
    marco_cara_uvn,
    cara_paralela=None,
    host_element=None,
    normales_prioridad=None,
):
    """
    Orden: ``normales_prioridad`` (si se pasan); luego ``norm`` desde **BasisZ** de la cara
    paralela más cercana; luego marco inferior (``-V`` + reverso); luego planta / Z.
    """
    seen = []
    out = []

    def _add(v):
        if v is None:
            return
        try:
            vn = v.Normalize()
            key = (round(vn.X, 5), round(vn.Y, 5), round(vn.Z, 5))
            if key in seen:
                return
            seen.append(key)
            out.append(vn)
        except Exception:
            pass

    if normales_prioridad:
        for pv in normales_prioridad:
            if pv is None:
                continue
            _add(pv)
            try:
                _add(pv.Negate())
            except Exception:
                pass

    if cara_paralela is not None:
        try:
            from geometria_fundacion_cara_inferior import (
                obtener_norm_plano_barra_desde_basisz_cara_paralela,
            )

            fp = cara_paralela[0]
            tr = cara_paralela[1]
            npp = obtener_norm_plano_barra_desde_basisz_cara_paralela(
                line, fp, tr, elemento=host_element
            )
            if npp is not None:
                _add(npp)
                try:
                    _add(npp.Negate())
                except Exception:
                    pass
        except Exception:
            pass

    if marco_cara_uvn is not None:
        nv = _norm_desde_eje_v_cara_perpendicular_a_barra(line, marco_cara_uvn)
        if nv is not None:
            _add(nv)
            try:
                _add(nv.Negate())
            except Exception:
                pass

    npl = _normal_planta_perpendicular_a_linea(line)
    if npl is not None:
        _add(npl)
        try:
            _add(npl.Negate())
        except Exception:
            pass

    for extra in (XYZ.BasisZ,):
        try:
            _add(extra)
            _add(extra.Negate())
        except Exception:
            _add(XYZ.BasisZ)

    return out


def _rebar_shape_linea_simple(document):
    """Respaldo: primera forma Standard con SimpleLine (como FundacionAislada)."""
    if document is None:
        return None
    try:
        doc_key = document.Title
    except Exception:
        doc_key = None
    if doc_key is not None and doc_key in _SHAPE_LINEA_SIMPLE_CACHE:
        return _SHAPE_LINEA_SIMPLE_CACHE[doc_key]
    try:
        shapes = list(FilteredElementCollector(document).OfClass(RebarShape))
    except Exception:
        if doc_key is not None:
            _SHAPE_LINEA_SIMPLE_CACHE[doc_key] = None
        return None
    result = None
    for shape in shapes:
        try:
            if shape.RebarStyle != RebarStyle.Standard:
                continue
            if getattr(shape, "SimpleLine", False):
                result = shape
                break
        except Exception:
            continue
    if result is None:
        for shape in shapes:
            try:
                if shape.RebarStyle == RebarStyle.Standard:
                    result = shape
                    break
            except Exception:
                continue
    if doc_key is not None:
        _SHAPE_LINEA_SIMPLE_CACHE[doc_key] = result
    return result


def _rebar_shape_para_fundacion_inferior(document):
    """
    Por defecto el tipo existente **«03»**; si no está en el proyecto, respaldo SimpleLine/Standard.
    No crea ``RebarShape`` nuevos.
    """
    if document is None:
        return None
    sh = buscar_rebar_shape_por_nombre(document, REBAR_SHAPE_NOMBRE_DEFECTO)
    if sh is not None:
        return sh
    return _rebar_shape_linea_simple(document)


def _curves_list_una_linea(line):
    """``IList<Curve>`` para APIs que aceptan lista genérica."""
    lst = List[object]()
    lst.Add(line)
    return lst


def _curves_ilist_curve_typed(line):
    """Alternativa ``List[Curve]`` por si el binding exige el tipo base."""
    try:
        lst = List[Curve]()
        lst.Add(line)
        return lst
    except Exception:
        return _curves_list_una_linea(line)


def _create_from_curves_un_intento(
    document,
    host,
    bar_type,
    curves_ilist,
    nvec,
    so,
    eo,
    hook_type,
    use_existing,
    create_new,
):
    """
    Cubre las dos formas habituales de la misma sobrecarga de 11 parámetros:

    - Ganchos como ``RebarHookType`` (firma que muestra la documentación reciente).
    - Ganchos como ``ElementId`` (proyectos/API donde aún resuelve el otro método).
    """
    if document is None or host is None or bar_type is None or hook_type is None:
        return None
    # 1) RebarHookType — como en la firma pública de Revit (startHook / endHook).
    try:
        r = Rebar.CreateFromCurves(
            document,
            RebarStyle.Standard,
            bar_type,
            hook_type,
            hook_type,
            host,
            nvec,
            curves_ilist,
            so,
            eo,
            use_existing,
            create_new,
        )
        if r:
            return r
    except Exception:
        pass
    # 2) ElementId (mismo orden de argumentos; otra sobrecarga según versión).
    try:
        hid = hook_type.Id
        r = Rebar.CreateFromCurves(
            document,
            RebarStyle.Standard,
            bar_type,
            hid,
            hid,
            host,
            nvec,
            curves_ilist,
            so,
            eo,
            use_existing,
            create_new,
        )
        if r:
            return r
    except Exception:
        pass
    return None


def _hook_start_end_ids(hook_type, gancho_en_inicio, gancho_en_fin):
    """``ElementId`` de gancho por extremo; ``InvalidElementId`` si no aplica."""
    inv = ElementId.InvalidElementId
    if not gancho_en_inicio and not gancho_en_fin:
        return inv, inv
    if hook_type is None:
        return inv, inv
    try:
        hid = hook_type.Id
    except Exception:
        return inv, inv
    return (hid if gancho_en_inicio else inv), (hid if gancho_en_fin else inv)


def _set_rebar_hook_type_id_end(rebar, end_idx, type_id, document, max_attempts=1):
    """Asigna ``SetHookTypeId``. En caso de fallo de verificación, reintenta una sola vez
    con ``Regenerate`` (reducido de 2 a 1 para evitar regeneraciones innecesarias mid-loop)."""
    if rebar is None:
        return False
    e = int(end_idx)
    for attempt in range(int(max_attempts)):
        try:
            rebar.SetHookTypeId(e, type_id)
        except Exception:
            return False
        try:
            got = rebar.GetHookTypeId(e)
            if int(got.IntegerValue) == int(type_id.IntegerValue):
                return True
        except Exception:
            pass
        if document is not None and attempt + 1 < int(max_attempts):
            try:
                document.Regenerate()
            except Exception:
                pass
    return False


def _create_from_curves_un_intento_dos_ids(
    document,
    host,
    bar_type,
    curves_ilist,
    nvec,
    so,
    eo,
    id_start,
    id_end,
    use_existing,
    create_new,
):
    """
    ``CreateFromCurves`` con ganchos por ``ElementId`` (asimetría o sin ganchos).

    Sin ganchos: algunas versiones resuelven ``(None, None)`` y otras ``(Invalid, Invalid)``;
    se prueban ambas (mismo criterio que ``enfierrado_shaft_hashtag``).
    """
    if document is None or host is None or bar_type is None:
        return None
    inv = ElementId.InvalidElementId
    try:
        both_inv = id_start == inv and id_end == inv
    except Exception:
        both_inv = False
    if both_inv:
        try_pairs = ((None, None), (inv, inv))
    else:
        try_pairs = ((id_start, id_end),)
    for h0, h1 in try_pairs:
        try:
            r = Rebar.CreateFromCurves(
                document,
                RebarStyle.Standard,
                bar_type,
                h0,
                h1,
                host,
                nvec,
                curves_ilist,
                so,
                eo,
                use_existing,
                create_new,
            )
            if r:
                return r
        except Exception:
            continue
    return None


def _try_create_rebar_no_hooks_then_assign_hooks(
    document,
    host,
    bar_type,
    curves_ilist,
    nvec,
    so,
    eo,
    use_existing,
    create_new,
    hook_type,
    gancho_en_inicio,
    gancho_en_fin,
):
    """
    Si ``CreateFromCurves`` con pareja (gancho, inv) falla, crea sin ganchos y fija
    ``HookTypeId`` por extremo (Revit a veces no acepta la sobrecarga mixta vía IronPython).
    """
    if document is None or host is None or bar_type is None or hook_type is None:
        return None
    if not gancho_en_inicio and not gancho_en_fin:
        return None
    inv = ElementId.InvalidElementId
    try:
        hid = hook_type.Id
    except Exception:
        return None
    for h0, h1 in ((None, None), (inv, inv)):
        try:
            r = Rebar.CreateFromCurves(
                document,
                RebarStyle.Standard,
                bar_type,
                h0,
                h1,
                host,
                nvec,
                curves_ilist,
                so,
                eo,
                use_existing,
                create_new,
            )
            if not r:
                continue
            _set_rebar_hook_type_id_end(r, 0, inv, document)
            _set_rebar_hook_type_id_end(r, 1, inv, document)
            _set_rebar_hook_type_id_end(
                r, 0, hid if gancho_en_inicio else inv, document
            )
            _set_rebar_hook_type_id_end(
                r, 1, hid if gancho_en_fin else inv, document
            )
            return r
        except Exception:
            continue
    return None


def _intentar_create_from_curves_minimo_extremos(
    document,
    host,
    bar_type,
    line,
    hook_type,
    marco_cara_uvn=None,
    cara_paralela=None,
    eje_referencia_z_ganchos=None,
    normales_prioridad=None,
    gancho_en_inicio=True,
    gancho_en_fin=True,
):
    """
    Igual que :func:`_intentar_create_from_curves_minimo` pero con ganchos opcionales
    por extremo (``InvalidElementId`` donde no hay gancho).
    """
    if document is None or host is None or bar_type is None or line is None:
        return None, None
    if (gancho_en_inicio or gancho_en_fin) and hook_type is None:
        return None, None
    id_s, id_e = _hook_start_end_ids(hook_type, gancho_en_inicio, gancho_en_fin)
    ct = clr.GetClrType(Line).BaseType
    arr = System.Array.CreateInstance(ct, 1)
    arr[0] = line
    curves_variants = (arr, _curves_ilist_curve_typed(line))
    normals = _lista_normales_create_from_curves(
        line,
        marco_cara_uvn,
        cara_paralela=cara_paralela,
        host_element=host,
        normales_prioridad=normales_prioridad,
    )
    seen = []
    bool_pairs = ((True, False),)
    orient_pairs = (
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
        (RebarHookOrientation.Right, RebarHookOrientation.Right),
        (RebarHookOrientation.Left, RebarHookOrientation.Left),
    )

    doc_key = _doc_title_safe(document)
    win = _REBAR_MIN_WIN.get(doc_key) if doc_key else None

    def _attempt_single(nvec, so, eo):
        """Prueba ambas variantes de curves_ilist y retorna el Rebar o None."""
        for use_ex, create_new in bool_pairs:
            for curves_ilist in curves_variants:
                if curves_ilist is None:
                    continue
                if gancho_en_inicio and gancho_en_fin:
                    r = _create_from_curves_un_intento(
                        document, host, bar_type, curves_ilist,
                        nvec, so, eo, hook_type, use_ex, create_new,
                    )
                else:
                    r = _create_from_curves_un_intento_dos_ids(
                        document, host, bar_type, curves_ilist,
                        nvec, so, eo, id_s, id_e, use_ex, create_new,
                    )
                    if r is None and hook_type is not None and (gancho_en_inicio or gancho_en_fin):
                        r = _try_create_rebar_no_hooks_then_assign_hooks(
                            document, host, bar_type, curves_ilist,
                            nvec, so, eo, use_ex, create_new,
                            hook_type, gancho_en_inicio, gancho_en_fin,
                        )
                if r:
                    return r
        return None

    # Probar primero la combinación ganadora de ejecuciones previas.
    if win is not None:
        win_nv_key, win_so, win_eo = win
        for nvec in normals:
            if nvec is None:
                continue
            try:
                nv_key = (round(nvec.X, 6), round(nvec.Y, 6), round(nvec.Z, 6))
            except Exception:
                nv_key = None
            if nv_key == win_nv_key:
                r = _attempt_single(nvec, win_so, win_eo)
                if r:
                    return r, nvec
                break

    for nvec in normals:
        if nvec is None:
            continue
        key = (round(nvec.X, 6), round(nvec.Y, 6), round(nvec.Z, 6))
        if key in seen:
            continue
        seen.append(key)
        hook_primary = _orientacion_ganchos_unificada_segun_z_proyecto(
            line, nvec, eje_referencia_z=eje_referencia_z_ganchos
        )
        orient_tries = [
            (hook_primary, hook_primary),
            (RebarHookOrientation.Right, RebarHookOrientation.Left),
            (RebarHookOrientation.Left, RebarHookOrientation.Right),
        ]
        if gancho_en_inicio != gancho_en_fin:
            orient_tries.extend(orient_pairs)
        for so, eo in orient_tries:
            r = _attempt_single(nvec, so, eo)
            if r:
                if doc_key is not None:
                    try:
                        nv_key = (round(nvec.X, 6), round(nvec.Y, 6), round(nvec.Z, 6))
                        _REBAR_MIN_WIN[doc_key] = (nv_key, so, eo)
                    except Exception:
                        pass
                return r, nvec
    return None, None


def _validar_tramos_polilinea_minimo(lines):
    if not lines:
        return False
    for ln in lines:
        if ln is None or not isinstance(ln, Line):
            return False
        try:
            if float(ln.Length) < _MIN_CURVE_LENGTH_FT:
                return False
        except Exception:
            return False
    return True


def _ilist_curves_open_chain_reversed(la):
    """
    Misma polilínea abierta recorrida en sentido inverso (útil si ``CreateFromCurves`` rechaza
    el orden original de una L).
    """
    if la is None or len(la) < 2:
        return None
    try:
        out = List[Curve]()
        for ln in reversed(la):
            if ln is None:
                continue
            p0 = ln.GetEndPoint(0)
            p1 = ln.GetEndPoint(1)
            out.Add(Line.CreateBound(p1, p0))
    except Exception:
        return None
    try:
        if out is None or int(out.Count) < 2:
            return None
    except Exception:
        return None
    return out


def _intentar_create_from_curves_polilinea_sin_ganchos(
    document,
    host,
    bar_type,
    curves_ilist,
    line_ref_normales,
    marco_cara_uvn=None,
    cara_paralela=None,
    eje_referencia_z_ganchos=None,
    normales_prioridad=None,
):
    """
    Polilínea de varios ``Line`` conectados sin ``RebarHookType`` (ganchos modelados como tramos).
    """
    if (
        document is None
        or host is None
        or bar_type is None
        or curves_ilist is None
        or line_ref_normales is None
    ):
        return None, None
    id_s = ElementId.InvalidElementId
    id_e = ElementId.InvalidElementId
    la = []
    try:
        _cnt = int(curves_ilist.Count)
    except Exception:
        try:
            _cnt = len(curves_ilist)
        except Exception:
            _cnt = 0
    for _ii in range(_cnt):
        try:
            la.append(curves_ilist[_ii])
        except Exception:
            break
    normals_prio_geom = []
    if len(la) == 2:
        n_g = _normal_plano_polilinea_dos_tramos(la)
        if n_g is not None:
            normals_prio_geom.append(n_g)
            try:
                nn = XYZ(-float(n_g.X), -float(n_g.Y), -float(n_g.Z))
                if float(nn.GetLength()) > 1e-12:
                    normals_prio_geom.append(nn.Normalize())
            except Exception:
                pass
        n_t = _normal_plano_polilinea_L_desde_tangentes(la)
        if n_t is not None:
            normals_prio_geom.append(n_t)
            try:
                nn2 = XYZ(-float(n_t.X), -float(n_t.Y), -float(n_t.Z))
                if float(nn2.GetLength()) > 1e-12:
                    normals_prio_geom.append(nn2.Normalize())
            except Exception:
                pass
    elif len(la) >= 3:
        n_g = _normal_plano_polilinea_u_tres_tramos(la)
        if n_g is not None:
            normals_prio_geom.append(n_g)
            try:
                nn = XYZ(-float(n_g.X), -float(n_g.Y), -float(n_g.Z))
                if float(nn.GetLength()) > 1e-12:
                    normals_prio_geom.append(nn.Normalize())
            except Exception:
                pass
    seen_geom = []
    normals_head = []
    for n in normals_prio_geom:
        if n is None:
            continue
        try:
            u = n.Normalize()
        except Exception:
            u = n
        key = (round(float(u.X), 6), round(float(u.Y), 6), round(float(u.Z), 6))
        if key in seen_geom:
            continue
        seen_geom.append(key)
        normals_head.append(u)
    normals_marco = _lista_normales_create_from_curves(
        line_ref_normales,
        marco_cara_uvn,
        cara_paralela=cara_paralela,
        host_element=host,
        normales_prioridad=normales_prioridad,
    )
    normals = normals_head + [x for x in normals_marco if x is not None]
    bool_pairs = (
        (True, True),
        (True, False),
        (False, True),
    )
    orient_pairs = (
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
        (RebarHookOrientation.Right, RebarHookOrientation.Right),
        (RebarHookOrientation.Left, RebarHookOrientation.Left),
    )
    gancho_en_inicio = False
    gancho_en_fin = False
    hook_primary = RebarHookOrientation.Right
    multi_tramo = len(la) > 1
    curve_variants = [curves_ilist]
    rev_il = _ilist_curves_open_chain_reversed(la)
    if rev_il is not None:
        curve_variants.append(rev_il)
    for curves_pass in curve_variants:
        if curves_pass is None:
            continue
        seen = []
        for nvec in normals:
            if nvec is None:
                continue
            key = (round(nvec.X, 6), round(nvec.Y, 6), round(nvec.Z, 6))
            if key in seen:
                continue
            seen.append(key)
            try:
                hook_primary = _orientacion_ganchos_unificada_segun_z_proyecto(
                    line_ref_normales, nvec, eje_referencia_z=eje_referencia_z_ganchos
                )
            except Exception:
                hook_primary = RebarHookOrientation.Right
            orient_tries = [
                (hook_primary, hook_primary),
                (RebarHookOrientation.Right, RebarHookOrientation.Left),
                (RebarHookOrientation.Left, RebarHookOrientation.Right),
            ]
            if gancho_en_inicio != gancho_en_fin or multi_tramo:
                orient_tries.extend(orient_pairs)
            for so, eo in orient_tries:
                for use_ex, create_new in bool_pairs:
                    r = _create_from_curves_un_intento_dos_ids(
                        document,
                        host,
                        bar_type,
                        curves_pass,
                        nvec,
                        so,
                        eo,
                        id_s,
                        id_e,
                        use_ex,
                        create_new,
                    )
                    if r:
                        return r, nvec
    return None, None


def _curves_ilist_desde_curve_loop_o_misma_lista(curves_ilist):
    """
    Si ``CurveLoop.Create`` acepta las curvas, devuelve la lista recorrida del lazo y el lazo
    (normal vía ``GetPlane()``). Si la U es abierta y Revit rechaza el lazo, devuelve la misma
    lista de entrada y ``curve_loop=None``.
    """
    if curves_ilist is None:
        return curves_ilist, None
    try:
        cl = CurveLoop.Create(curves_ilist)
    except Exception:
        return curves_ilist, None
    if cl is None:
        return curves_ilist, None
    try:
        out = List[Curve]()
        for c in cl:
            if c is not None:
                out.Add(c)
    except Exception:
        return curves_ilist, None
    if out is None or out.Count < 1:
        return curves_ilist, None
    return out, cl


def _normal_plano_polilinea_u_tres_tramos(lineas):
    """
    Normal unitaria del plano de la U (tres ``Line`` consecutivas: pata, eje, pata).
    Usa puntos en los extremos de las patas y vértices del eje para ``Plane.CreateByThreePoints``.
    """
    if lineas is None or len(lineas) < 3:
        return None
    try:
        ln0, ln1, ln2 = lineas[0], lineas[1], lineas[2]
        p_q0 = ln0.GetEndPoint(0)
        p0 = ln0.GetEndPoint(1)
        p1 = ln1.GetEndPoint(1)
        p_q1 = ln2.GetEndPoint(1)
    except Exception:
        return None
    for trio in ((p_q0, p0, p_q1), (p_q0, p1, p_q1)):
        try:
            pl = Plane.CreateByThreePoints(trio[0], trio[1], trio[2])
            if pl is None:
                continue
            n = pl.Normal
            if n is None or float(n.GetLength()) < 1e-12:
                continue
            return n.Normalize()
        except Exception:
            continue
    return None


def _normal_plano_polilinea_L_desde_tangentes(lineas):
    """Normal ⟂ al plano de la L como ``t0 × t1`` en el vértice común."""
    if lineas is None or len(lineas) < 2:
        return None
    try:
        t0 = _tangente_unitaria_linea(lineas[0])
        t1 = _tangente_unitaria_linea(lineas[1])
        if t0 is None or t1 is None:
            return None
        n = t0.CrossProduct(t1)
        if n is None or float(n.GetLength()) < 1e-12:
            return None
        return n.Normalize()
    except Exception:
        return None


def _normal_plano_polilinea_dos_tramos(lineas):
    """
    Normal unitaria del plano de una L (dos ``Line`` consecutivas que comparten vértice).
    """
    if lineas is None or len(lineas) < 2:
        return None
    try:
        ln0, ln1 = lineas[0], lineas[1]
        p_a = ln0.GetEndPoint(0)
        p_m = ln0.GetEndPoint(1)
        p_b = ln1.GetEndPoint(1)
    except Exception:
        return None
    try:
        pl = Plane.CreateByThreePoints(p_a, p_m, p_b)
        if pl is None:
            return None
        n = pl.Normal
        if n is None or float(n.GetLength()) < 1e-12:
            return None
        return n.Normalize()
    except Exception:
        return None


def crear_rebar_polilinea_u_malla_inf_sup_curve_loop(
    document,
    host,
    bar_type,
    lineas,
    linea_central_para_normales,
    marco_cara_uvn=None,
    cara_paralela=None,
    eje_referencia_z_ganchos=None,
    normales_prioridad=None,
):
    """
    normales_prioridad: opcional, mismo criterio que en ``_lista_normales_create_from_curves``.

    Malla **inferior y superior** activas: ``CurveLoop`` + ``GetPlane().Normal`` cuando Revit
    forma un lazo válido; si no (U abierta típica), misma ``IList<Curve>`` y normal por plano
    de tres puntos. ``Rebar.CreateFromCurves`` sin ``RebarHookType``; prueba ``(True, True)`` y
    ``(True, False)`` y varias orientaciones de gancho.

    Returns:
        tuple: ``(rebar | None, mensaje_error | None, norm_createfromcurves | None)``.
    """
    if document is None or host is None or bar_type is None or not lineas:
        return None, None, None
    seq = list(lineas)
    if len(seq) != 3:
        return None, None, None
    if linea_central_para_normales is None:
        return None, None, None
    if not _validar_tramos_polilinea_minimo(seq):
        return None, None, None
    try:
        curves_ilist = List[Curve]()
        for ln in seq:
            curves_ilist.Add(ln)
    except Exception:
        return None, None, None
    rebar_curves, curve_loop = _curves_ilist_desde_curve_loop_o_misma_lista(curves_ilist)
    normals_prio = []
    if curve_loop is not None:
        try:
            pl = curve_loop.GetPlane()
            if pl is not None:
                n0 = pl.Normal
                if n0 is not None and float(n0.GetLength()) > 1e-12:
                    try:
                        normals_prio.append(n0.Normalize())
                    except Exception:
                        normals_prio.append(n0)
        except Exception:
            pass
    n3 = _normal_plano_polilinea_u_tres_tramos(seq)
    if n3 is not None:
        normals_prio.append(n3)
        try:
            nn = XYZ(-float(n3.X), -float(n3.Y), -float(n3.Z))
            if float(nn.GetLength()) > 1e-12:
                normals_prio.append(nn.Normalize())
        except Exception:
            pass
    seen_n = []
    normals_head = []
    for n in normals_prio:
        if n is None:
            continue
        try:
            u = n.Normalize()
        except Exception:
            u = n
        key = (round(float(u.X), 6), round(float(u.Y), 6), round(float(u.Z), 6))
        if key in seen_n:
            continue
        seen_n.append(key)
        normals_head.append(u)
    normals_marco = _lista_normales_create_from_curves(
        linea_central_para_normales,
        marco_cara_uvn,
        cara_paralela=cara_paralela,
        host_element=host,
        normales_prioridad=normales_prioridad,
    )
    normals_order = normals_head + [x for x in normals_marco if x is not None]
    inv = ElementId.InvalidElementId
    bool_pairs = ((True, True), (True, False))
    orient_pairs = (
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
        (RebarHookOrientation.Right, RebarHookOrientation.Right),
        (RebarHookOrientation.Left, RebarHookOrientation.Left),
    )
    try:
        hook_primary = _orientacion_ganchos_unificada_segun_z_proyecto(
            linea_central_para_normales,
            normals_order[0] if normals_order else None,
            eje_referencia_z=eje_referencia_z_ganchos,
        )
    except Exception:
        hook_primary = RebarHookOrientation.Right
    orient_tries = [(hook_primary, hook_primary)]
    orient_tries.extend(orient_pairs)
    for nvec in normals_order:
        if nvec is None:
            continue
        try:
            nu = nvec.Normalize()
        except Exception:
            nu = nvec
        for so, eo in orient_tries:
            for use_ex, create_new in bool_pairs:
                r = _create_from_curves_un_intento_dos_ids(
                    document,
                    host,
                    bar_type,
                    rebar_curves,
                    nu,
                    so,
                    eo,
                    inv,
                    inv,
                    use_ex,
                    create_new,
                )
                if r:
                    return r, None, nu
    return (
        None,
        u"No CreateFromCurves (CurveLoop/plano) para esta polilínea U.",
        None,
    )


def _agent_dbg_shape_u_03(msg, data=None, hypothesis_id=""):
    # #region agent log
    try:
        import json
        import os
        import time

        p = os.path.normpath(
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "debug-9aea0b.log")
        )
        line = {
            "sessionId": "9aea0b",
            "location": "rebar_fundacion_cara_inferior:crear_rebar_u_shape_desde_eje",
            "message": msg,
            "data": data or {},
            "timestamp": int(time.time() * 1000),
            "hypothesisId": hypothesis_id,
        }
        with open(p, "a") as f:
            f.write(json.dumps(line) + "\n")
    except Exception:
        pass
    # #endregion


def crear_rebar_u_shape_desde_eje_rebar_shape_nombrado(
    document,
    host,
    bar_type,
    polilinea_u_tres_tramos,
    shape_nombre=None,
    marco_cara_uvn=None,
    cara_paralela=None,
    eje_referencia_z_ganchos=None,
    normales_prioridad=None,
):
    """
    normales_prioridad: probadas primero como ``norm`` en ``CreateFromCurvesAndShape`` (p. ej.
    invertir el sentido de propagación del conjunto sin mover la polilínea).

    Barra en U vía ``Rebar.CreateFromCurvesAndShape`` y un ``RebarShape`` del proyecto
    (p. ej. **«03»**). Revit exige curvas que formen lazo válido y coincidan con la forma;
    se pasan los **tres** tramos rectos de la U (sin arcos de gancho en la lista). En la
    creación no se asigna ``RebarHookType``: ``None`` / null en API.

    Returns:
        tuple: ``(rebar | None, mensaje_error | None, norm_createfromcurves | None)``.
    """
    if shape_nombre is None:
        shape_nombre = REBAR_SHAPE_NOMBRE_DEFECTO
    if document is None or host is None or bar_type is None or not polilinea_u_tres_tramos:
        return None, u"Argumentos incompletos.", None
    lineas = list(polilinea_u_tres_tramos)
    if len(lineas) != 3:
        return None, u"La U requiere exactamente tres tramos rectos.", None
    linea_eje = lineas[1]
    if not _validar_tramos_polilinea_minimo(lineas):
        return None, u"Un tramo de la polilínea es demasiado corto o no válido.", None
    shape = buscar_rebar_shape_por_nombre(document, shape_nombre)
    if shape is None:
        return (
            None,
            u"No se encontró RebarShape «{0}» en el proyecto.".format(shape_nombre),
            None,
        )
    try:
        curves_ilist = List[Curve]()
        for ln in lineas:
            curves_ilist.Add(ln)
    except Exception:
        return None, u"No se pudo construir la lista de curvas.", None
    try:
        _ll = float(linea_eje.Length)
    except Exception:
        _ll = -1.0
    _defines_hooks = None
    try:
        from Autodesk.Revit.DB.Structure import ReinforcementSettings

        _rs = ReinforcementSettings.GetReinforcementSettings(document)
        if _rs is not None:
            _defines_hooks = bool(_rs.RebarShapeDefinesHooks)
    except Exception:
        pass
    _agent_dbg_shape_u_03(
        "shape_u_entry",
        {
            "shape_nombre": shape_nombre,
            "eje_len_ft": _ll,
            "n_curves": 3,
            "rebar_shape_defines_hooks": _defines_hooks,
        },
        "H5",
    )
    normals = _lista_normales_create_from_curves(
        linea_eje,
        marco_cara_uvn,
        cara_paralela=cara_paralela,
        host_element=host,
        normales_prioridad=normales_prioridad,
    )
    seen = []
    last_ex = None
    for nvec in normals:
        if nvec is None:
            continue
        key = (round(nvec.X, 6), round(nvec.Y, 6), round(nvec.Z, 6))
        if key in seen:
            continue
        seen.append(key)
        try:
            hook_primary = _orientacion_ganchos_unificada_segun_z_proyecto(
                linea_eje,
                nvec,
                eje_referencia_z=eje_referencia_z_ganchos,
            )
        except Exception:
            hook_primary = RebarHookOrientation.Right
        orient_tries = [
            (hook_primary, hook_primary),
            (RebarHookOrientation.Right, RebarHookOrientation.Left),
            (RebarHookOrientation.Left, RebarHookOrientation.Right),
            (RebarHookOrientation.Right, RebarHookOrientation.Right),
            (RebarHookOrientation.Left, RebarHookOrientation.Left),
        ]
        for so, eo in orient_tries:
            try:
                r = Rebar.CreateFromCurvesAndShape(
                    document,
                    shape,
                    bar_type,
                    None,
                    None,
                    host,
                    nvec,
                    curves_ilist,
                    so,
                    eo,
                    0.0,
                    0.0,
                    ElementId.InvalidElementId,
                    ElementId.InvalidElementId,
                )
                if r:
                    _agent_dbg_shape_u_03("shape_u_ok", {"shape_nombre": shape_nombre}, "H5")
                    return r, None, nvec
            except Exception as ex:
                try:
                    last_ex = unicode(ex)
                except Exception:
                    last_ex = str(ex)
            try:
                r = Rebar.CreateFromCurvesAndShape(
                    document,
                    shape,
                    bar_type,
                    None,
                    None,
                    host,
                    nvec,
                    curves_ilist,
                    so,
                    eo,
                )
                if r:
                    _agent_dbg_shape_u_03("shape_u_ok", {"shape_nombre": shape_nombre}, "H5")
                    return r, None, nvec
            except Exception as ex:
                try:
                    last_ex = unicode(ex)
                except Exception:
                    last_ex = str(ex)
                continue
    _agent_dbg_shape_u_03(
        "shape_u_failed",
        {"shape_nombre": shape_nombre, "last_ex": last_ex},
        "H2",
    )
    return (
        None,
        u"No CreateFromCurvesAndShape con RebarShape «{0}» y el eje dado.".format(
            shape_nombre
        ),
        None,
    )


def crear_rebar_polilinea_recta_sin_ganchos(
    document,
    host,
    bar_type,
    lineas,
    linea_central_para_normales,
    marco_cara_uvn=None,
    cara_paralela=None,
    eje_referencia_z_ganchos=None,
    normales_prioridad=None,
):
    """
    Crea ``Rebar`` a partir de varias ``Line`` consecutivas (sin tipos de gancho).

    Returns:
        tuple: ``(rebar | None, mensaje_error | None, norm_createfromcurves | None)``.
    """
    if document is None or host is None or bar_type is None or not lineas:
        return None, u"Argumentos incompletos.", None
    if linea_central_para_normales is None:
        return None, u"Falta línea de referencia para normales.", None
    if not _validar_tramos_polilinea_minimo(lineas):
        return None, u"Un tramo de la polilínea es demasiado corto o no válido.", None
    try:
        curves_ilist = List[Curve]()
        for ln in lineas:
            curves_ilist.Add(ln)
    except Exception:
        return None, u"No se pudo construir la lista de curvas.", None
    r, nv = _intentar_create_from_curves_polilinea_sin_ganchos(
        document,
        host,
        bar_type,
        curves_ilist,
        linea_central_para_normales,
        marco_cara_uvn=marco_cara_uvn,
        cara_paralela=cara_paralela,
        eje_referencia_z_ganchos=eje_referencia_z_ganchos,
        normales_prioridad=normales_prioridad,
    )
    if r is not None:
        return r, None, nv
    return (
        None,
        u"No se pudo crear el Rebar desde la polilínea (CreateFromCurves sin ganchos).",
        None,
    )


def _intentar_create_from_curves_and_shape(
    document,
    host,
    bar_type,
    line,
    hook_type,
    marco_cara_uvn=None,
    cara_paralela=None,
    eje_referencia_z_ganchos=None,
    normales_prioridad=None,
):
    """
    Camino principal: ``norm`` desde **BasisZ** de la cara paralela cercana (si hay), luego marco inferior.

    ``eje_referencia_z_ganchos``: vector unitario (~Z) para :func:`_orientacion_ganchos_unificada_segun_z_proyecto`
    (p. ej. ``XYZ(0,0,-1)`` en cara superior).
    """
    if document is None or host is None or bar_type is None or line is None or hook_type is None:
        return None, None
    shape = _rebar_shape_para_fundacion_inferior(document)
    if shape is None:
        return None, None
    curves = _curves_list_una_linea(line)
    deduped = _lista_normales_create_from_curves(
        line,
        marco_cara_uvn,
        cara_paralela=cara_paralela,
        host_element=host,
        normales_prioridad=normales_prioridad,
    )

    doc_key = _doc_title_safe(document)
    win = _REBAR_AND_SHAPE_WIN.get(doc_key) if doc_key else None

    def _try_nv_hook(nv, hook_o):
        try:
            r = Rebar.CreateFromCurvesAndShape(
                document, shape, bar_type, hook_type, hook_type,
                host, nv, curves, hook_o, hook_o,
                0.0, 0.0,
                ElementId.InvalidElementId, ElementId.InvalidElementId,
            )
            if r:
                return r
        except Exception:
            pass
        try:
            r = Rebar.CreateFromCurvesAndShape(
                document, shape, bar_type, hook_type, hook_type,
                host, nv, curves, hook_o, hook_o,
            )
            if r:
                return r
        except Exception:
            pass
        return None

    # Probar primero la combinación ganadora de llamadas previas.
    if win is not None:
        win_nv_key, win_hook_o_val = win
        for nv in deduped:
            if nv is None:
                continue
            try:
                nv_key = (round(nv.X, 5), round(nv.Y, 5), round(nv.Z, 5))
            except Exception:
                nv_key = None
            if nv_key == win_nv_key:
                try:
                    hook_o = type(win_hook_o_val)(int(win_hook_o_val))
                except Exception:
                    hook_o = win_hook_o_val
                r = _try_nv_hook(nv, hook_o)
                if r:
                    return r, nv
                break

    for nv in deduped:
        if nv is None:
            continue
        hook_o = _orientacion_ganchos_unificada_segun_z_proyecto(
            line, nv, eje_referencia_z=eje_referencia_z_ganchos
        )
        r = _try_nv_hook(nv, hook_o)
        if r:
            if doc_key is not None:
                try:
                    nv_key = (round(nv.X, 5), round(nv.Y, 5), round(nv.Z, 5))
                    _REBAR_AND_SHAPE_WIN[doc_key] = (nv_key, hook_o)
                except Exception:
                    pass
            return r, nv
    return None, None


def _intentar_create_from_curves_minimo(
    document,
    host,
    bar_type,
    line,
    hook_type,
    marco_cara_uvn=None,
    cara_paralela=None,
    eje_referencia_z_ganchos=None,
    normales_prioridad=None,
):
    """
    Pocos intentos; ``norm`` vía :func:`_lista_normales_create_from_curves` (BasisZ cara paralela, etc.).
    Equivale a :func:`_intentar_create_from_curves_minimo_extremos` con ganchos en ambos extremos.
    """
    return _intentar_create_from_curves_minimo_extremos(
        document,
        host,
        bar_type,
        line,
        hook_type,
        marco_cara_uvn=marco_cara_uvn,
        cara_paralela=cara_paralela,
        eje_referencia_z_ganchos=eje_referencia_z_ganchos,
        normales_prioridad=normales_prioridad,
        gancho_en_inicio=True,
        gancho_en_fin=True,
    )


def crear_rebar_desde_curva_linea_con_ganchos(
    document,
    host,
    bar_type,
    curve,
    hook_type_name=None,
    marco_cara_uvn=None,
    cara_paralela=None,
    eje_referencia_z_ganchos=None,
    normales_prioridad=None,
    gancho_en_inicio=True,
    gancho_en_fin=True,
):
    """
    Crea un ``Rebar`` con la curva dada (``Line``) y ganchos opcionales por extremo.

    Con ``gancho_en_inicio`` / ``gancho_en_fin`` en ``False`` se usa ``InvalidElementId`` en ese
    extremo (p. ej. tramos intermedios de una corrida con traslape).

    Args:
        marco_cara_uvn: resultado de :func:`geometria_fundacion_cara_inferior.obtener_marco_coordenadas_cara_inferior`
            ``(origin, u, v, n)``.
        cara_paralela: ``(face, transform)`` de la cara paralela más cercana a la curva
            (:func:`geometria_fundacion_cara_inferior.evaluar_caras_paralelas_curva_mas_cercana`).
            Si existe, el ``norm`` de ``CreateFromCurves*`` prioriza el **BasisZ** interno de esa cara
            (proyectado ⟂ barra, orientado al interior del host) para la propagación del conjunto.
        normales_prioridad: lista opcional de ``XYZ`` (no necesariamente unitarios) probados **primero**
            como ``norm`` en ``CreateFromCurves*`` (p. ej. normal de la cara inferior en armadura lateral).
        eje_referencia_z_ganchos: eje (~unitario) usado con la regla Left/Right de ganchos respecto al Z del proyecto.
            Por defecto ``None`` (= ``BasisZ``). Para armadura en la **cara superior**, usar ``XYZ(0,0,-1)``
            para orientar los ganchos hacia el interior del hormigón.
        gancho_en_inicio: si ``False``, el extremo inicial de la línea no tendrá gancho.
        gancho_en_fin: si ``False``, el extremo final no tendrá gancho.

    Returns:
        tuple: ``(rebar | None, mensaje_error | None, norm_createfromcurves | None)``.
        ``norm_createfromcurves`` es el ``XYZ`` (unitario) pasado a ``CreateFromCurves*`` en el
        intento que tuvo éxito; sirve para marcadores y depuración.
    """
    if hook_type_name is None:
        hook_type_name = HOOK_GANCHO_90_STANDARD_NAME
    if document is None or host is None or bar_type is None or curve is None:
        return None, u"Argumentos incompletos.", None
    if not isinstance(curve, Line):
        return None, u"La curva debe ser una línea.", None
    if not _validar_linea_horizontal(curve):
        return None, u"Curva demasiado corta o no válida.", None
    hook_type = None
    if gancho_en_inicio or gancho_en_fin:
        hook_type = buscar_rebar_hook_type_por_nombre(document, hook_type_name)
        if hook_type is None:
            return (
                None,
                u"No se encontró RebarHookType '{0}' en el proyecto.".format(hook_type_name),
                None,
            )

    r, nv = None, None
    if gancho_en_inicio and gancho_en_fin:
        r, nv = _intentar_create_from_curves_and_shape(
            document,
            host,
            bar_type,
            curve,
            hook_type,
            marco_cara_uvn=marco_cara_uvn,
            cara_paralela=cara_paralela,
            eje_referencia_z_ganchos=eje_referencia_z_ganchos,
            normales_prioridad=normales_prioridad,
        )
    if r is not None:
        return r, None, nv

    r2, nv2 = _intentar_create_from_curves_minimo_extremos(
        document,
        host,
        bar_type,
        curve,
        hook_type,
        marco_cara_uvn=marco_cara_uvn,
        cara_paralela=cara_paralela,
        eje_referencia_z_ganchos=eje_referencia_z_ganchos,
        normales_prioridad=normales_prioridad,
        gancho_en_inicio=gancho_en_inicio,
        gancho_en_fin=gancho_en_fin,
    )
    if r2 is not None:
        return r2, None, nv2

    return (
        None,
        u"No se pudo crear el Rebar (CreateFromCurvesAndShape ni CreateFromCurves). "
        u"Compruebe tipo de barra, gancho y host.",
        None,
    )


def aplicar_layout_maximum_spacing_rebar(
    rebar, document, separacion_mm, array_length_ft, flip_rebar_set=True
):
    """
    Aplica ``RebarShapeDrivenAccessor.SetLayoutAsMaximumSpacing`` (regla *Maximum Spacing*).

    Args:
        rebar: ``Rebar`` creado con ``CreateFromCurves*``.
        document: ``Document`` (regeneración opcional).
        separacion_mm: paso máximo entre barras (mm), desde el combo Separación inferior.
        array_length_ft: longitud del conjunto en la dirección de distribución (pies).
        flip_rebar_set: si es ``False``, no llama a ``FlipRebarSet``. Zapata de muro usa ``False``
            y fija el sentido con el ``norm`` de creación (p. ej. ``normales_prioridad`` en U).

    Returns:
        tuple: ``(ok: bool, aviso | None)`` — si ``ok`` es False, ``aviso`` describe el fallo.
    """
    if rebar is None:
        return False, u"Rebar nulo."
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId

        sep_ft = UnitUtils.ConvertToInternalUnits(float(separacion_mm), UnitTypeId.Millimeters)
    except Exception:
        sep_ft = float(separacion_mm) / 304.8
    acc = None
    try:
        acc = rebar.GetShapeDrivenAccessor()
    except Exception:
        acc = None
    except:
        acc = None
    if acc is None:
        return False, u"GetShapeDrivenAccessor no disponible (no es barra dirigida por forma)."
    try:
        al = float(array_length_ft)
    except Exception:
        al = 1.0
    try:
        if al < 1e-9:
            acc.SetLayoutAsSingle()
            return True, None
        if sep_ft >= al - 1e-9:
            acc.SetLayoutAsSingle()
            return True, None
    except Exception as ex:
        return False, unicode(ex)
    except:
        return False, u"SetLayoutAsSingle: error (CLR) al preparar conjunto."
    # API: SetLayoutAsMaximumSpacing(spacing, arrayLength, barsOnNormalSide, includeFirstBar, includeLastBar).
    # includeFirstBar / includeLastBar = True: barras visibles en ambos extremos del conjunto.
    combos = (
        (True, True, True),
        (False, True, True),
    )
    last_err = None
    try:
        for b_side, inc_first, inc_last in combos:
            try:
                acc.SetLayoutAsMaximumSpacing(
                    sep_ft, al, b_side, inc_first, inc_last
                )
                # Invierte el sentido de propagación del conjunto respecto al predeterminado del plano de barra.
                if flip_rebar_set:
                    try:
                        acc.FlipRebarSet()
                    except Exception:
                        pass
                return True, None
            except Exception as ex:
                try:
                    last_err = unicode(ex)
                except Exception:
                    last_err = u"SetLayoutAsMaximumSpacing"
                continue
            except:
                last_err = u"SetLayoutAsMaximumSpacing (CLR)"
                continue
    except:
        return (
            False,
            u"Maximum Spacing: error al aplicar conjunto (p. ej. barra polilinea sin shape-driven).",
        )
    return False, last_err or u"No se pudo fijar Maximum Spacing."


def _rebar_cantidad_posiciones(rebar):
    try:
        return int(rebar.Quantity)
    except Exception:
        try:
            return int(rebar.NumberOfBarPositions)
        except Exception:
            return 1


def aplicar_layout_fixed_number_rebar(rebar, document, n_barras, array_length_ft):
    """
    Aplica ``RebarShapeDrivenAccessor.SetLayoutAsFixedNumber`` (regla *Fixed Number*).

    Args:
        rebar: ``Rebar`` creado con ``CreateFromCurves*``.
        document: ``Document`` (``Regenerate`` opcional entre intentos).
        n_barras: número de barras del conjunto (p. ej. cantidad armadura superior).
        array_length_ft: longitud del conjunto en la dirección de distribución (pies).

    Returns:
        tuple: ``(ok: bool, aviso | None)``.
    """
    if rebar is None:
        return False, u"Rebar nulo."
    try:
        acc = rebar.GetShapeDrivenAccessor()
    except Exception:
        acc = None
    if acc is None:
        return False, u"GetShapeDrivenAccessor no disponible (no es barra dirigida por forma)."
    try:
        n = int(n_barras)
    except Exception:
        n = 1
    if n <= 1:
        try:
            acc.SetLayoutAsSingle()
            return True, None
        except Exception as ex:
            try:
                return False, unicode(ex)
            except Exception:
                return False, u"SetLayoutAsSingle"

    try:
        al = float(array_length_ft)
    except Exception:
        al = 0.0
    arr_len = max(al, 5.0 / 304.8)

    def _qty_ok():
        try:
            return _rebar_cantidad_posiciones(rebar) == int(n)
        except Exception:
            return False

    combos = (
        (True, True, True),
        (False, True, True),
        (True, False, False),
        (False, False, False),
    )
    last_err_holder = [None]

    def _try_all_combos():
        for b_side, inc0, inc1 in combos:
            try:
                acc.SetLayoutAsFixedNumber(
                    int(n), float(arr_len), b_side, inc0, inc1
                )
                if _qty_ok():
                    return True
            except Exception as ex:
                try:
                    last_err_holder[0] = unicode(ex)
                except Exception:
                    last_err_holder[0] = u"SetLayoutAsFixedNumber"
        return False

    if _try_all_combos():
        return True, None
    if document is not None:
        try:
            document.Regenerate()
        except Exception:
            pass
    try:
        acc.FlipRebarSet()
    except Exception:
        pass
    if _try_all_combos():
        return True, None
    return False, (last_err_holder[0] or u"No se pudo fijar Fixed Number.")
