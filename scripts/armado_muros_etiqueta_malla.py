# -*- coding: utf-8 -*-
"""
Utilidades de vista y geometría para mallas en Armado muros.

Incluye centroide del muro, validación de vista Building Section y helpers
usados por ``armado_muros_malla_rebar_tags``. La etiqueta TextNote D.M./H./V.
(``crear_etiqueta_texto_malla_muro``) quedó obsoleta: las mallas se etiquetan
con Structural Rebar Tag desde las barras.
"""

from __future__ import print_function

import clr
import unicodedata

try:
    unicode
except NameError:
    unicode = str

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    GeometryInstance,
    HorizontalTextAlignment,
    Options,
    Solid,
    TextNote,
    TextNoteOptions,
    TextNoteType,
    VerticalTextAlignment,
    XYZ,
)

# Tipo de texto obligatorio para la etiqueta (nombre en el proyecto Revit).
TEXT_NOTE_TYPE_NAME = u"2.5mm Arial_Arrow Filled 15 Degree"
ETIQUETA_PREFIJO = u"D.M."
_VISTAS_ETIQUETA = frozenset(("Elevation", "Section", "Detail"))


def _norm_name(s):
    if s is None:
        return u""
    try:
        t = unicodedata.normalize("NFC", unicode(s))
    except Exception:
        try:
            t = unicode(s)
        except Exception:
            t = str(s)
    for ch in (u"\xa0", u"\u200b", u"\u202f", u"\u2002", u"\u2003", u"\ufeff"):
        t = t.replace(ch, u" ")
    return u" ".join(t.split())


def _canon_name_key(s):
    """Clave de comparación insensible a espacios Unicode y mayúsculas."""
    return _norm_name(s).lower()


def _names_equal(a, b):
    if not a or not b:
        return False
    ka = _canon_name_key(a)
    kb = _canon_name_key(b)
    if not ka or not kb:
        return False
    if ka == kb:
        return True
    try:
        if unicode(ka) == unicode(kb):
            return True
    except Exception:
        pass
    return False


def _view_type_suffix(view):
    """Devuelve 'Section', 'Elevation', etc. (IronPython a veces devuelve el nombre completo del enum)."""
    if view is None:
        return u""
    try:
        vt = view.ViewType
        try:
            s = vt.ToString()
        except Exception:
            s = str(vt)
    except Exception:
        return u""
    s = (s or u"").strip()
    if u"." in s:
        s = s.split(u".")[-1]
    return s


def es_vista_elevacion_seccion(view):
    if view is None:
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import ViewType

        vt = view.ViewType
        if vt in (ViewType.Elevation, ViewType.Section, ViewType.Detail):
            return True
    except Exception:
        pass
    suf = _view_type_suffix(view)
    return suf in _VISTAS_ETIQUETA


_VISTA_DETALLE_MARKERS = (
    u"detail",
    u"detalle",
    u"callout",
    u"recuadro",
    u"detailed",
)
_BUILDING_SECTION_NAME_MARKERS = (
    u"building section",
    u"sección de edificio",
    u"seccion de edificio",
)


def _es_view_family_type(vft):
    if vft is None:
        return False
    try:
        from Autodesk.Revit.DB import ViewFamilyType

        if isinstance(vft, ViewFamilyType):
            return True
    except Exception:
        pass
    try:
        return type(vft).__name__ == u"ViewFamilyType"
    except Exception:
        pass
    return hasattr(vft, u"ViewFamily") and hasattr(vft, u"Name")


def _enum_igual(valor, enum_obj):
    if valor is None or enum_obj is None:
        return False
    try:
        if valor == enum_obj:
            return True
    except Exception:
        pass
    try:
        if int(valor) == int(enum_obj):
            return True
    except Exception:
        pass
    try:
        a = _canon_name_key(valor.ToString() if hasattr(valor, u"ToString") else valor)
        b = _canon_name_key(enum_obj.ToString() if hasattr(enum_obj, u"ToString") else enum_obj)
        if a and b and a.split(u".")[-1] == b.split(u".")[-1]:
            return True
    except Exception:
        pass
    return False


def _vista_es_tipo_section(view):
    if view is None:
        return False
    if _view_type_suffix(view) == u"Section":
        return True
    try:
        from Autodesk.Revit.DB import ViewType

        return _enum_igual(view.ViewType, ViewType.Section)
    except Exception:
        return False


def _vista_es_tipo_detail(view):
    if view is None:
        return False
    if _view_type_suffix(view) == u"Detail":
        return True
    try:
        from Autodesk.Revit.DB import ViewType

        return _enum_igual(view.ViewType, ViewType.Detail)
    except Exception:
        return False


def _parametro_texto(element, *builtins):
    if element is None:
        return u""
    for bip in builtins:
        try:
            p = element.get_Parameter(bip)
            if p is None:
                continue
            s = p.AsValueString()
            if s:
                return _norm_name(s)
        except Exception:
            pass
        try:
            p = element.get_Parameter(bip)
            if p is None:
                continue
            s = p.AsString()
            if s:
                return _norm_name(s)
        except Exception:
            pass
    return u""


def _extraer_tipo_desde_familia_y_tipo(texto):
    s = _norm_name(texto)
    if not s:
        return u""
    if u":" in s:
        return _norm_name(s.split(u":", 1)[1])
    return s


def _view_family_type_element(view):
    if view is None:
        return None
    doc = None
    try:
        doc = view.Document
    except Exception:
        doc = None
    if doc is None:
        return None
    try:
        tid = view.GetTypeId()
        if tid is not None and tid != ElementId.InvalidElementId:
            vft = doc.GetElement(tid)
            if _es_view_family_type(vft):
                return vft
    except Exception:
        pass
    return None


def _nombre_vft_es_building_section(name):
    n = _canon_name_key(name or u"")
    if not n:
        return False
    for bad in _VISTA_DETALLE_MARKERS:
        if bad in n:
            return False
    for ok in _BUILDING_SECTION_NAME_MARKERS:
        if ok in n:
            return True
    return False


def _view_family_type_name(view):
    vft = _view_family_type_element(view)
    if vft is not None:
        try:
            nm = vft.Name or u""
            if nm:
                return nm
        except Exception:
            pass
    try:
        raw = _parametro_texto(
            view,
            BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM,
            BuiltInParameter.ALL_MODEL_TYPE_NAME,
            BuiltInParameter.SYMBOL_NAME_PARAM,
        )
        tipo = _extraer_tipo_desde_familia_y_tipo(raw)
        if tipo:
            return tipo
        if raw:
            return raw
    except Exception:
        pass
    return u""


def _vft_es_familia_section(vft):
    if vft is None:
        return False
    try:
        from Autodesk.Revit.DB import ViewFamily

        return _enum_igual(vft.ViewFamily, ViewFamily.Section)
    except Exception:
        pass
    try:
        vf = vft.ViewFamily
        s = vf.ToString() if hasattr(vf, u"ToString") else str(vf)
        return u"Section" in (s or u"")
    except Exception:
        return False


def es_vista_building_section(view):
    """True si la vista es sección de edificio (``ViewType.Section`` + tipo Building Section)."""
    if view is None:
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass
    if _vista_es_tipo_detail(view):
        return False
    if not _vista_es_tipo_section(view):
        return False

    vft = _view_family_type_element(view)
    if vft is not None:
        if not _vft_es_familia_section(vft):
            return False
        try:
            if _nombre_vft_es_building_section(vft.Name):
                return True
        except Exception:
            pass

    nombre_tipo = _view_family_type_name(view)
    if _nombre_vft_es_building_section(nombre_tipo):
        return True

    # Respaldo: ViewType.Section + ViewFamily.Section sin marcadores de detalle en el nombre.
    if vft is not None and _vft_es_familia_section(vft):
        n = _canon_name_key(nombre_tipo or u"")
        if n:
            for bad in _VISTA_DETALLE_MARKERS:
                if bad in n:
                    return False
        return True

    return False


def es_vista_alzado_o_seccion_edificio(view):
    """Alias retrocompatible: solo secciones Building Section."""
    return es_vista_building_section(view)


def texto_aviso_vista_building_section(view):
    """Devuelve ``(instruction, content)`` para el diálogo WPF de la herramienta."""
    vn = u""
    vt_s = _view_type_suffix(view) or u"desconocido"
    try:
        vn = view.Name or u""
    except Exception:
        pass
    vft_name = _view_family_type_name(view)
    fam_tipo = u""
    try:
        fam_tipo = _parametro_texto(view, BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)
    except Exception:
        pass
    instruction = (
        u"Esta herramienta solo puede ejecutarse en secciones "
        u"tipo Building Section."
    )
    content = u"Vista activa: «{0}» ({1}).".format(vn, vt_s)
    if vft_name:
        content += u"\nTipo de vista: «{0}».".format(vft_name)
    elif fam_tipo:
        content += u"\nFamilia y tipo: «{0}».".format(fam_tipo)
    content += u"\n\nAbra una sección de edificio antes de continuar."
    return instruction, content


def mensaje_vista_requerida_armado_muros(view):
    """Texto plano (respaldo si no carga WPF)."""
    instruction, content = texto_aviso_vista_building_section(view)
    return instruction + u"\n\n" + content


def _proyectar_punto_plano_vista(p, view):
    if p is None or view is None:
        return p
    try:
        vd = view.ViewDirection
        if vd is None or float(vd.GetLength()) < 1e-12:
            return p
        vd = vd.Normalize()
        vo = view.Origin
        if vo is None:
            return p
        d = float((p - vo).DotProduct(vd))
        return p - vd.Multiply(d)
    except Exception:
        return p


def _solid_centroid(solid):
    try:
        return solid.ComputeCentroid()
    except Exception:
        return None


def centroide_geometria_muro(element, view=None):
    if element is None:
        return None
    opts = Options()
    opts.ComputeReferences = False
    try:
        if view is not None:
            opts.View = view
    except Exception:
        pass
    try:
        geom_elem = element.get_Geometry(opts)
    except Exception:
        geom_elem = None
    if geom_elem is not None:
        vol_sum = 0.0
        sx = sy = sz = 0.0
        for obj in geom_elem:
            if obj is None:
                continue
            if isinstance(obj, Solid) and obj.Volume > 1e-12:
                c = _solid_centroid(obj)
                if c is None:
                    continue
                v = obj.Volume
                sx += c.X * v
                sy += c.Y * v
                sz += c.Z * v
                vol_sum += v
            elif isinstance(obj, GeometryInstance):
                try:
                    inst_geom = obj.GetInstanceGeometry()
                    if inst_geom is None:
                        continue
                    for g in inst_geom:
                        if isinstance(g, Solid) and g.Volume > 1e-12:
                            c = _solid_centroid(g)
                            if c is None:
                                continue
                            v = g.Volume
                            sx += c.X * v
                            sy += c.Y * v
                            sz += c.Z * v
                            vol_sum += v
                except Exception:
                    pass
        if vol_sum > 1e-12:
            return XYZ(sx / vol_sum, sy / vol_sum, sz / vol_sum)
    try:
        bb = element.get_BoundingBox(view)
        if bb is None:
            bb = element.get_BoundingBox(None)
        if bb and bb.Min and bb.Max:
            mn, mx = bb.Min, bb.Max
            return XYZ((mn.X + mx.X) * 0.5, (mn.Y + mx.Y) * 0.5, (mn.Z + mx.Z) * 0.5)
    except Exception:
        pass
    return None


def _text_type_name(tnt):
    if tnt is None:
        return u""
    try:
        n = tnt.Name
        if n and str(n).strip():
            return _norm_name(n)
    except Exception:
        pass
    for bip in (
        BuiltInParameter.SYMBOL_NAME_PARAM,
        BuiltInParameter.ALL_MODEL_TYPE_NAME,
    ):
        try:
            p = tnt.get_Parameter(bip)
            if p and p.HasValue:
                s = p.AsString()
                if s and str(s).strip():
                    return _norm_name(s)
        except Exception:
            continue
    return u""


def _collect_text_note_types(document):
    out = []
    try:
        col = FilteredElementCollector(document).OfClass(TextNoteType)
        try:
            col = col.WhereElementIsElementType()
        except Exception:
            pass
        for tnt in col:
            if tnt:
                out.append(tnt)
    except Exception:
        pass
    return out


def _find_text_note_type_named(document, exact_name):
    target = _norm_name(exact_name)
    if not target:
        return None
    target_key = _canon_name_key(exact_name)
    types = _collect_text_note_types(document)

    def _type_name_keys(tnt):
        keys = []
        nn = _text_type_name(tnt)
        if nn:
            keys.append(_canon_name_key(nn))
        try:
            raw = tnt.Name
            if raw:
                keys.append(_canon_name_key(raw))
        except Exception:
            pass
        return keys

    for tnt in types:
        try:
            for nk in _type_name_keys(tnt):
                if nk and (nk == target_key or _names_equal(nk, target_key)):
                    return tnt
        except Exception:
            continue

    # Respaldo: único «2.5mm … Arrow Filled 15 Degree» (IronPython a veces falla == en unicode).
    candidates = []
    for tnt in types:
        try:
            nk = _canon_name_key(_text_type_name(tnt))
            if not nk:
                continue
            if (
                nk.startswith(u"2.5")
                and u"arial" in nk
                and u"arrow" in nk
                and u"filled" in nk
                and u"15" in nk
                and u"degree" in nk
                and u"open dot" not in nk
            ):
                candidates.append(tnt)
        except Exception:
            continue
    if len(candidates) == 1:
        return candidates[0]
    for tnt in candidates:
        if _canon_name_key(_text_type_name(tnt)) == target_key:
            return tnt
    if candidates:
        return candidates[0]
    return None


def _text_note_type_obligatorio(document):
    return _find_text_note_type_named(document, TEXT_NOTE_TYPE_NAME)


def _nombres_tipos_similares(document, max_items=5):
    target = _norm_name(TEXT_NOTE_TYPE_NAME).lower()
    hits = []
    for tnt in _collect_text_note_types(document):
        nn = _text_type_name(tnt).lower()
        if not nn:
            continue
        if u"arial" in nn and u"2.5" in nn:
            hits.append(_text_type_name(tnt))
        elif target and (target in nn or nn in target):
            hits.append(_text_type_name(tnt))
    uniq = []
    for h in hits:
        if h not in uniq:
            uniq.append(h)
        if len(uniq) >= max_items:
            break
    return uniq


def _eid_int(eid):
    try:
        return int(eid.Value)
    except Exception:
        try:
            return int(eid.IntegerValue)
        except Exception:
            return u"?"


def _append_error(errores, wall, msg):
    if errores is None:
        return
    errores.append(u"Muro {}: {}.".format(_eid_int(wall.Id) if wall else u"?", msg))


def _diametro_mm_bar_type(document, bar_type_id):
    if not bar_type_id or bar_type_id == ElementId.InvalidElementId:
        return None
    try:
        bt = document.GetElement(bar_type_id)
        if bt is None:
            return None
        diam_ft = bt.BarNominalDiameter
        return int(round(float(diam_ft) * 304.8))
    except Exception:
        return None


def _capas_etiqueta(params_dict, layer_active_dict):
    def _capa_activa(key):
        return layer_active_dict.get(key, True)

    def _datos(key):
        bar_id, esp = params_dict.get(key, (ElementId.InvalidElementId, u"150"))
        if not _capa_activa(key):
            return ElementId.InvalidElementId, None
        return bar_id, esp

    if _capa_activa("exterior_major") or _capa_activa("exterior_minor"):
        pref = "exterior"
    elif _capa_activa("interior_major") or _capa_activa("interior_minor"):
        pref = "interior"
    else:
        pref = "exterior"

    h_bar, h_esp = _datos(u"{}_major".format(pref))
    v_bar, v_esp = _datos(u"{}_minor".format(pref))
    if h_esp is None and v_esp is not None:
        h_esp = v_esp
    if v_esp is None and h_esp is not None:
        v_esp = h_esp
    if h_esp is None:
        h_esp = u"150"
    if v_esp is None:
        v_esp = u"150"
    return h_bar, h_esp, v_bar, v_esp


def _formatear_fila(prefijo, diam_mm, esp_mm):
    try:
        d = int(round(float(diam_mm)))
    except (TypeError, ValueError):
        d = 0
    try:
        e = int(round(float(str(esp_mm).strip())))
    except (TypeError, ValueError):
        e = 150
    return u"{0}=\u00f8{1}@{2}".format(prefijo, d, e)


def _texto_etiqueta(h_diam_mm, h_esp_mm, v_diam_mm, v_esp_mm):
    h_line = _formatear_fila(u"H.", h_diam_mm, h_esp_mm)
    v_line = _formatear_fila(u"V.", v_diam_mm, v_esp_mm)
    return u"{0}\t{1}\n\t{2}".format(ETIQUETA_PREFIJO, h_line, v_line)


def _crear_text_note(document, view, origin, texto, text_type, usar_alineacion=True):
    def _opts(con_alineacion):
        try:
            o = TextNoteOptions(text_type.Id)
        except Exception:
            o = TextNoteOptions()
            o.TypeId = text_type.Id
        if con_alineacion:
            try:
                o.HorizontalAlignment = HorizontalTextAlignment.Center
            except Exception:
                pass
            try:
                o.VerticalAlignment = VerticalTextAlignment.Middle
            except Exception:
                pass
        return o

    last_ex = None
    for con_alin in (True, False):
        try:
            TextNote.Create(document, view.Id, origin, texto, _opts(con_alin))
            return None
        except Exception as ex:
            last_ex = ex
    return last_ex


def crear_etiqueta_texto_malla_muro(
    document,
    view,
    wall,
    params_dict,
    layer_active_dict,
    errores=None,
):
    """
    Crea TextNote en el centroide del muro (proyectado al plano de la vista).
    Usa el tipo de texto ``TEXT_NOTE_TYPE_NAME`` (comparación normalizada del nombre).
    Retorna True si se creó la etiqueta.
    """
    if wall is None:
        return False

    if not es_vista_elevacion_seccion(view):
        _append_error(
            errores,
            wall,
            u"etiqueta solo en elevación/sección (vista activa: {})".format(
                _view_type_suffix(view) or u"?",
            ),
        )
        return False

    h_bar, h_esp, v_bar, v_esp = _capas_etiqueta(params_dict, layer_active_dict)
    h_diam = _diametro_mm_bar_type(document, h_bar)
    v_diam = _diametro_mm_bar_type(document, v_bar)
    if h_diam is None and v_diam is not None:
        h_diam = v_diam
    if v_diam is None and h_diam is not None:
        v_diam = h_diam
    if h_diam is None and v_diam is None:
        _append_error(errores, wall, u"sin diámetro de barra válido para la etiqueta")
        return False
    if h_diam is None:
        h_diam = 0
    if v_diam is None:
        v_diam = 0

    origin = centroide_geometria_muro(wall, view)
    if origin is None:
        _append_error(errores, wall, u"no se pudo calcular el centroide del muro")
        return False
    origin = _proyectar_punto_plano_vista(origin, view)

    text_type = _text_note_type_obligatorio(document)
    if text_type is None:
        sim = _nombres_tipos_similares(document)
        msg = u"no existe el tipo de texto «{}»".format(TEXT_NOTE_TYPE_NAME)
        if sim:
            msg += u" (similares: {})".format(u"; ".join(sim))
        _append_error(errores, wall, msg)
        return False

    texto = _texto_etiqueta(h_diam, h_esp, v_diam, v_esp)
    ex = _crear_text_note(document, view, origin, texto, text_type)
    if ex is None:
        return True
    _append_error(
        errores,
        wall,
        u"TextNote.Create falló ({0}): {1}".format(
            _text_type_name(text_type),
            str(ex),
        ),
    )
    return False
