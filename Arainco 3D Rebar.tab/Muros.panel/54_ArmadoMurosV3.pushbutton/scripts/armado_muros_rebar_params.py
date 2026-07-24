# -*- coding: utf-8 -*-
"""Parámetros compartidos en Rebar creados por Armado Muros."""

import clr

clr.AddReference("System")
clr.AddReference("RevitAPI")
from System import AppDomain

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    Level,
    StorageType,
    Wall,
)
from Autodesk.Revit.DB.Structure import MultiplanarOption, Rebar

ARMADURA_ARAINCO_PARAM = u"Armadura_Arainco"
ARMADURA_CAPA_PARAM = u"Armadura_Capa"
ARMADURA_MALLA_PARAM = u"Armadura_Malla"
ARMADURA_MALLA_TIPO_PARAM = u"Armadura_Malla_Tipo"
ARMADURA_MALLA_ORIENTACION_PARAM = u"Armadura_Malla_Orientacion"
ARMADURA_CONJUNTO_GUID_PARAM = u"Armadura_Conjunto_GUID"
ARMADURA_NIVEL_PARAM = u"Armadura_Nivel"
ARMADURA_EJE_PARAM = u"Armadura_Eje"
ARMADURA_MALLA_TIPO_DM = u"D.M."
ARMADURA_MALLA_ORIENT_V = u"V."
ARMADURA_MALLA_ORIENT_H = u"H."
_APPDOMAIN_CONJUNTO_GUID_KEY = u"Arainco_ArmadoMurosV3_Conjunto_GUID"
_APPDOMAIN_ARMADURA_EJE_KEY = u"Arainco_ArmadoMurosV3_Armadura_Eje"


def _norm_param_def_name(name):
    if name is None:
        return u""
    try:
        t = unicode(name).replace(u"\u00A0", u" ").strip()
    except Exception:
        try:
            t = str(name).strip()
        except Exception:
            return u""
    return t


def _iter_element_parameters(element):
    if element is None:
        return
    try:
        for p in element.Parameters:
            yield p
    except Exception:
        pass


def _iter_rebar_parameters(rebar):
    return _iter_element_parameters(rebar)


def _find_element_parameter(element, param_name):
    """
    ``LookupParameter`` y, si falla, barrido por ``Parameters``
    (algunos parámetros compartidos solo aparecen al iterar).
    """
    if element is None or not param_name:
        return None
    target = _norm_param_def_name(param_name).lower()
    if not target:
        return None
    try:
        p = element.LookupParameter(param_name)
        if p is not None:
            return p
    except Exception:
        pass
    try:
        for p in _iter_element_parameters(element):
            if p is None:
                continue
            try:
                dn = _norm_param_def_name(p.Definition.Name).lower()
            except Exception:
                continue
            if dn == target:
                return p
    except Exception:
        pass
    return None


def _find_rebar_parameter(rebar, param_name):
    return _find_element_parameter(rebar, param_name)


def generar_armadura_conjunto_guid():
    """UUID de corrida (único por ejecución unificada)."""
    try:
        import uuid
        return unicode(uuid.uuid4())
    except Exception:
        pass
    try:
        import System
        return unicode(System.Guid.NewGuid())
    except Exception:
        return None


def iniciar_armadura_conjunto_guid_ejecucion(conjunto_guid=None):
    """Activa el GUID de corrida en AppDomain para estampar rebars de la ejecución."""
    gid = conjunto_guid or generar_armadura_conjunto_guid()
    if not gid:
        return None
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_CONJUNTO_GUID_KEY, gid)
    except Exception:
        pass
    return gid


def obtener_armadura_conjunto_guid_actual():
    """GUID de corrida activo o ``None``."""
    try:
        gid = AppDomain.CurrentDomain.GetData(_APPDOMAIN_CONJUNTO_GUID_KEY)
    except Exception:
        gid = None
    if not gid:
        return None
    try:
        t = unicode(gid).strip()
    except Exception:
        try:
            t = str(gid or u"").strip()
        except Exception:
            return None
    return t or None


def finalizar_armadura_conjunto_guid_ejecucion():
    """Limpia el GUID de corrida al terminar la ejecución."""
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_CONJUNTO_GUID_KEY, None)
    except Exception:
        pass


def _param_value_as_text(param):
    """Lee texto de un ``Parameter`` (AsString / AsValueString)."""
    if param is None:
        return None
    val = None
    try:
        val = param.AsString()
    except Exception:
        pass
    if not val:
        try:
            val = param.AsValueString()
        except Exception:
            pass
    if not val:
        return None
    try:
        t = unicode(val).strip()
    except Exception:
        try:
            t = str(val or u"").strip()
        except Exception:
            return None
    return t or None


def leer_armadura_eje_desde_vista(view):
    """
    Lee ``Armadura_Eje`` de la vista activa (o la vista indicada).

    Returns:
        texto normalizado o ``None`` si no hay valor / parámetro.
    """
    if view is None:
        return None
    p = _find_element_parameter(view, ARMADURA_EJE_PARAM)
    return _param_value_as_text(p)


def iniciar_armadura_eje_ejecucion(uidoc=None, view=None, eje_valor=None):
    """
    Cachea ``Armadura_Eje`` de la vista activa para estampar rebars de la corrida.

    Prioridad: ``eje_valor`` explícito → ``view`` → ``uidoc.ActiveView``.
    """
    valor = None
    if eje_valor is not None:
        try:
            valor = unicode(eje_valor).strip()
        except Exception:
            try:
                valor = str(eje_valor or u"").strip()
            except Exception:
                valor = None
        if not valor:
            valor = None
    if valor is None:
        vista = view
        if vista is None and uidoc is not None:
            try:
                vista = uidoc.ActiveView
            except Exception:
                vista = None
        valor = leer_armadura_eje_desde_vista(vista)
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_ARMADURA_EJE_KEY, valor)
    except Exception:
        pass
    return valor


def obtener_armadura_eje_actual():
    """Valor ``Armadura_Eje`` cacheado de la corrida o ``None``."""
    try:
        val = AppDomain.CurrentDomain.GetData(_APPDOMAIN_ARMADURA_EJE_KEY)
    except Exception:
        val = None
    if not val:
        return None
    try:
        t = unicode(val).strip()
    except Exception:
        try:
            t = str(val or u"").strip()
        except Exception:
            return None
    return t or None


def finalizar_armadura_eje_ejecucion():
    """Limpia el valor ``Armadura_Eje`` de corrida al terminar la ejecución."""
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_ARMADURA_EJE_KEY, None)
    except Exception:
        pass


def stamp_armadura_eje(rebar, eje_valor=None):
    """Escribe ``Armadura_Eje`` en un ``Rebar`` (valor explícito o cache de corrida)."""
    if rebar is None or not isinstance(rebar, Rebar):
        return False
    valor = eje_valor if eje_valor is not None else obtener_armadura_eje_actual()
    if not valor:
        return False
    return _set_rebar_string_param(rebar, ARMADURA_EJE_PARAM, valor)


def _set_element_string_param(element, param_name, value):
    """Escribe un parámetro de instancia tipo texto en un elemento."""
    if element is None or not param_name:
        return False
    try:
        valor = unicode(value)
    except Exception:
        try:
            valor = str(value)
        except Exception:
            return False
    p = _find_element_parameter(element, param_name)
    if p is None or p.IsReadOnly:
        return False
    try:
        st = p.StorageType
        if st == StorageType.String:
            p.Set(valor)
            return True
    except Exception:
        pass
    try:
        p.SetValueString(valor)
        return True
    except Exception:
        pass
    try:
        p.Set(valor)
        return True
    except Exception:
        return False


def _set_rebar_string_param(rebar, param_name, value):
    """Escribe un parámetro de instancia tipo texto en un ``Rebar``."""
    if rebar is None or not isinstance(rebar, Rebar) or not param_name:
        return False
    return _set_element_string_param(rebar, param_name, value)


def stamp_armadura_conjunto_guid(element, conjunto_guid=None):
    """Escribe ``Armadura_Conjunto_GUID`` en un elemento (Rebar, Detail Item, etc.)."""
    if element is None:
        return False
    gid = conjunto_guid or obtener_armadura_conjunto_guid_actual()
    if not gid:
        return False
    return _set_element_string_param(element, ARMADURA_CONJUNTO_GUID_PARAM, gid)


def stamp_armadura_conjunto_guid_por_ids(doc, element_ids, conjunto_guid=None):
    """Escribe ``Armadura_Conjunto_GUID`` en una lista de ``ElementId``."""
    if doc is None or not element_ids:
        return 0
    gid = conjunto_guid or obtener_armadura_conjunto_guid_actual()
    if not gid:
        return 0
    n = 0
    for eid in element_ids:
        try:
            el = doc.GetElement(eid)
        except Exception:
            continue
        if stamp_armadura_conjunto_guid(el, conjunto_guid=gid):
            n += 1
    return n


def get_armadura_conjunto_guid(element):
    """Lee ``Armadura_Conjunto_GUID`` de un elemento o ``None``."""
    if element is None:
        return None
    p = _find_element_parameter(element, ARMADURA_CONJUNTO_GUID_PARAM)
    if p is None:
        return None
    val = None
    try:
        val = p.AsString()
    except Exception:
        pass
    if not val:
        try:
            val = p.AsValueString()
        except Exception:
            pass
    if not val:
        return None
    try:
        t = unicode(val).strip()
    except Exception:
        try:
            t = str(val or u"").strip()
        except Exception:
            return None
    return t or None


def _normalize_conjunto_guid_target(conjunto_guid):
    if not conjunto_guid:
        return None
    try:
        target = unicode(conjunto_guid).strip()
    except Exception:
        try:
            target = str(conjunto_guid or u"").strip()
        except Exception:
            return None
    return target or None


def collect_rebars_por_conjunto_guid(doc, conjunto_guid):
    """
    Devuelve ``ElementId`` de rebars con el mismo ``Armadura_Conjunto_GUID``.

    ``conjunto_guid`` debe coincidir exactamente con el valor leído por
    ``get_armadura_conjunto_guid`` (texto ya normalizado).
    """
    target = _normalize_conjunto_guid_target(conjunto_guid)
    if doc is None or not target:
        return []

    from Autodesk.Revit.DB import FilteredElementCollector

    ids = []
    try:
        rebars = (
            FilteredElementCollector(doc)
            .OfClass(Rebar)
            .WhereElementIsNotElementType()
        )
    except Exception:
        return []

    for rebar in rebars:
        try:
            gid = get_armadura_conjunto_guid(rebar)
        except Exception:
            continue
        if gid == target:
            try:
                ids.append(rebar.Id)
            except Exception:
                pass
    return ids


def collect_empalmes_por_conjunto_guid(doc, conjunto_guid):
    """
    Devuelve ``ElementId`` de representaciones de empalme (Detail Items)
    con el mismo ``Armadura_Conjunto_GUID``.
    """
    target = _normalize_conjunto_guid_target(conjunto_guid)
    if doc is None or not target:
        return []

    from Autodesk.Revit.DB import BuiltInCategory, FamilyInstance, FilteredElementCollector

    ids = []
    try:
        details = (
            FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_DetailComponents)
            .WhereElementIsNotElementType()
        )
    except Exception:
        return []

    for el in details:
        if not isinstance(el, FamilyInstance):
            continue
        try:
            gid = get_armadura_conjunto_guid(el)
        except Exception:
            continue
        if gid == target:
            try:
                ids.append(el.Id)
            except Exception:
                pass
    return ids


def collect_corrida_por_conjunto_guid(doc, conjunto_guid):
    """
    Barras + representaciones de empalme con el mismo GUID de corrida.

    Retorna ``dict`` con claves ``rebar_ids``, ``empalme_ids``, ``all_ids``.
    """
    rebar_ids = collect_rebars_por_conjunto_guid(doc, conjunto_guid)
    empalme_ids = collect_empalmes_por_conjunto_guid(doc, conjunto_guid)
    all_ids = list(rebar_ids)
    seen = set()
    for eid in rebar_ids:
        try:
            seen.add(int(eid.IntegerValue))
        except Exception:
            pass
    for eid in empalme_ids:
        try:
            key = int(eid.IntegerValue)
        except Exception:
            key = None
        if key is not None and key not in seen:
            all_ids.append(eid)
            seen.add(key)
    return {
        u"rebar_ids": rebar_ids,
        u"empalme_ids": empalme_ids,
        u"all_ids": all_ids,
    }


def armadura_capa_valor_desde_layer(layer_index):
    """Capa índice 0 → ``(1ºC.)``, índice 1 → ``(2ºC.)``, etc."""
    try:
        li = int(layer_index)
    except Exception:
        li = 0
    n = max(0, li) + 1
    return u"({0}ºC.)".format(n)


def set_armadura_capa_desde_layer(rebar, layer_index):
    """Escribe ``Armadura_Capa`` según índice de capa (0-based → 1ºC., 2ºC., …)."""
    if rebar is None or not isinstance(rebar, Rebar):
        return False
    valor_txt = armadura_capa_valor_desde_layer(layer_index)
    return _set_rebar_string_param(rebar, ARMADURA_CAPA_PARAM, valor_txt)


def stamp_armadura_nivel(rebar, level_name):
    """Escribe ``Armadura_Nivel`` (nombre de nivel) en un ``Rebar``."""
    if rebar is None or not level_name:
        return False
    return _set_rebar_string_param(rebar, ARMADURA_NIVEL_PARAM, level_name)


def _rebar_document(rebar):
    if rebar is None:
        return None
    try:
        return rebar.Document
    except Exception:
        return None


def _rebar_host_element(rebar, doc=None):
    """
    Elemento host del rebar (``GetHostId``).

    Si el host inmediato es ``AreaReinforcement``, resuelve el host estructural
    (muro/losa) para poder leer Base Constraint.
    """
    if rebar is None or not isinstance(rebar, Rebar):
        return None
    document = doc or _rebar_document(rebar)
    if document is None:
        return None
    try:
        hid = rebar.GetHostId()
    except Exception:
        return None
    if hid is None or hid == ElementId.InvalidElementId:
        return None
    try:
        host = document.GetElement(hid)
    except Exception:
        return None
    if host is None:
        return None
    try:
        from Autodesk.Revit.DB.Structure import AreaReinforcement
        if isinstance(host, AreaReinforcement):
            try:
                ar_hid = host.GetHostId()
            except Exception:
                ar_hid = None
            if ar_hid is not None and ar_hid != ElementId.InvalidElementId:
                try:
                    structural = document.GetElement(ar_hid)
                except Exception:
                    structural = None
                if structural is not None:
                    return structural
    except Exception:
        pass
    return host


def _nivel_nombre_desde_element_id(doc, level_id):
    if doc is None or level_id is None or level_id == ElementId.InvalidElementId:
        return None
    try:
        if int(level_id.IntegerValue) < 0:
            return None
    except Exception:
        pass
    try:
        level = doc.GetElement(level_id)
    except Exception:
        return None
    if level is None or not isinstance(level, Level):
        return None
    try:
        name = level.Name
    except Exception:
        return None
    if name is None:
        return None
    try:
        text = unicode(name).strip()
    except Exception:
        try:
            text = str(name).strip()
        except Exception:
            return None
    return text or None


def nivel_nombre_base_constraint_muro(doc, wall):
    """
    Nombre del nivel de ``WALL_BASE_CONSTRAINT`` del muro anfitrión.

    Respaldo: ``Wall.LevelId`` y parámetros de nivel de instancia.
    """
    if doc is None or wall is None:
        return None
    try:
        p = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)
        if p is not None and p.HasValue and p.StorageType == StorageType.ElementId:
            name = _nivel_nombre_desde_element_id(doc, p.AsElementId())
            if name:
                return name
    except Exception:
        pass
    try:
        lid = wall.LevelId
        name = _nivel_nombre_desde_element_id(doc, lid)
        if name:
            return name
    except Exception:
        pass
    for bip_name in (
        u"INSTANCE_REFERENCE_LEVEL_PARAM",
        u"LEVEL_PARAM",
        u"SCHEDULE_LEVEL_PARAM",
    ):
        try:
            bip = getattr(BuiltInParameter, bip_name, None)
            if bip is None:
                continue
            p = wall.get_Parameter(bip)
            if p is None or not p.HasValue or p.StorageType != StorageType.ElementId:
                continue
            name = _nivel_nombre_desde_element_id(doc, p.AsElementId())
            if name:
                return name
        except Exception:
            continue
    return None


def stamp_armadura_nivel_desde_host_muro(rebar, wall=None, doc=None):
    """
    ``Armadura_Nivel`` desde Base Constraint del muro host.

    Usado en malla y confinamiento.
    """
    if rebar is None or not isinstance(rebar, Rebar):
        return False
    document = doc or _rebar_document(rebar)
    host = wall
    if host is None:
        host = _rebar_host_element(rebar, document)
    if host is None or not isinstance(host, Wall):
        return False
    name = nivel_nombre_base_constraint_muro(document, host)
    if not name:
        return False
    return stamp_armadura_nivel(rebar, name)


def _curve_length_safe(curve):
    if curve is None:
        return 0.0
    try:
        return float(curve.Length)
    except Exception:
        pass
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        return float(p0.DistanceTo(p1))
    except Exception:
        return 0.0


def _rebar_centerline_curves(rebar, pos_idx=0):
    """Curvas de eje del rebar (posición ``pos_idx``); varios overload de API."""
    if rebar is None or not isinstance(rebar, Rebar):
        return None
    attempts = (
        (False, False, False, MultiplanarOption.IncludeAllMultiplanarCurves, int(pos_idx)),
        (False, False, False, MultiplanarOption.IncludeOnlyPlanarCurves, int(pos_idx)),
        (False, False, False),
    )
    for args in attempts:
        try:
            crvs = rebar.GetCenterlineCurves(*args)
        except Exception:
            continue
        if crvs is None:
            continue
        try:
            if int(crvs.Count) < 1:
                continue
        except Exception:
            continue
        return crvs
    return None


def rebar_main_segment_startpoint(rebar, pos_idx=0):
    """
    StartPoint del segmento principal (curva más larga) de la centerline.

    Returns:
        XYZ o None.
    """
    crvs = _rebar_centerline_curves(rebar, pos_idx=pos_idx)
    if crvs is None:
        return None
    best = None
    best_len = -1.0
    try:
        n = int(crvs.Count)
    except Exception:
        n = 0
    for i in range(n):
        try:
            c = crvs[i]
        except Exception:
            continue
        L = _curve_length_safe(c)
        if L > best_len:
            best_len = L
            best = c
    if best is None:
        return None
    try:
        return best.GetEndPoint(0)
    except Exception:
        return None


def listar_niveles_proyecto(doc):
    """Niveles del proyecto (no plantilla)."""
    if doc is None:
        return []
    try:
        levels = list(
            FilteredElementCollector(doc)
            .OfClass(Level)
            .WhereElementIsNotElementType()
        )
    except Exception:
        try:
            levels = list(
                FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Levels)
                .WhereElementIsNotElementType()
            )
        except Exception:
            return []
    out = []
    for lv in levels:
        if lv is None or not isinstance(lv, Level):
            continue
        out.append(lv)
    return out


def nivel_mas_cercano_a_z(doc, z_elevation):
    """``Level`` cuya elevación está más cerca de ``z_elevation`` (pies internos)."""
    levels = listar_niveles_proyecto(doc)
    if not levels:
        return None
    try:
        z = float(z_elevation)
    except Exception:
        return None
    best = None
    best_dist = None
    for lv in levels:
        try:
            elev = float(lv.Elevation)
        except Exception:
            continue
        dist = abs(elev - z)
        if best is None or dist < best_dist:
            best = lv
            best_dist = dist
    return best


def nivel_nombre_mas_cercano_a_z(doc, z_elevation):
    lv = nivel_mas_cercano_a_z(doc, z_elevation)
    if lv is None:
        return None
    try:
        name = lv.Name
    except Exception:
        return None
    if name is None:
        return None
    try:
        text = unicode(name).strip()
    except Exception:
        try:
            text = str(name).strip()
        except Exception:
            return None
    return text or None


def stamp_armadura_nivel_desde_centerline(rebar, doc=None, pos_idx=0):
    """
    ``Armadura_Nivel`` = nivel más cercano al StartPoint del segmento principal.

    Usado en longitudinales de cabezal y coronamiento.
    """
    if rebar is None or not isinstance(rebar, Rebar):
        return False
    document = doc or _rebar_document(rebar)
    if document is None:
        return False
    pt = rebar_main_segment_startpoint(rebar, pos_idx=pos_idx)
    if pt is None:
        return False
    try:
        z = float(pt.Z)
    except Exception:
        return False
    name = nivel_nombre_mas_cercano_a_z(document, z)
    if not name:
        return False
    return stamp_armadura_nivel(rebar, name)


def stamp_cabezal_longitudinal_rebar(rebar, layer_index=0):
    """``Armadura_Arainco`` + ``Armadura_Malla`` = No + ``Armadura_Capa`` + nivel centerline."""
    if rebar is None:
        return rebar
    activar_armadura_arainco(rebar)
    set_armadura_capa_desde_layer(rebar, layer_index)
    stamp_armadura_nivel_desde_centerline(rebar)
    return rebar


def stamp_confinamiento_rebar(rebar):
    """Confinamiento: ``Armadura_Arainco`` + ``Armadura_Nivel`` desde base del muro host."""
    if rebar is None:
        return rebar
    activar_armadura_arainco(rebar)
    stamp_armadura_nivel_desde_host_muro(rebar)
    return rebar


def stamp_coronamiento_rebar(rebar):
    """Coronamiento: ``Armadura_Arainco`` + ``Armadura_Nivel`` desde centerline."""
    if rebar is None:
        return rebar
    activar_armadura_arainco(rebar)
    stamp_armadura_nivel_desde_centerline(rebar)
    return rebar


def _set_rebar_yes_no_param(rebar, param_name, yes=True):
    """Escribe un parámetro Yes/No (Integer 0/1, bool o texto Yes/No)."""
    if rebar is None or not isinstance(rebar, Rebar) or not param_name:
        return False
    p = _find_rebar_parameter(rebar, param_name)
    if p is None or p.IsReadOnly:
        return False
    if yes:
        candidates = (1, True, u"1", u"Yes", u"yes", u"Sí", u"SI")
    else:
        candidates = (0, False, u"0", u"No", u"no")
    try:
        st = p.StorageType
        if st == StorageType.Integer:
            p.Set(1 if yes else 0)
            return True
    except Exception:
        pass
    for val in candidates:
        try:
            p.Set(val)
            return True
        except Exception:
            continue
    try:
        p.SetValueString(u"Yes" if yes else u"No")
        return True
    except Exception:
        return False


def activar_armadura_malla(rebar, yes=True):
    """Escribe ``Armadura_Malla`` Yes/No (cuantificación tablas)."""
    return _set_rebar_yes_no_param(rebar, ARMADURA_MALLA_PARAM, yes=yes)


def activar_armadura_arainco(rebar):
    """
    Activa ``Armadura_Arainco`` y marca ``Armadura_Malla`` = No.

    Las barras de malla pisan después con ``stamp_malla_*`` → ``Armadura_Malla`` = Yes.
    """
    if rebar is None or not isinstance(rebar, Rebar):
        return False
    ok = _set_rebar_yes_no_param(rebar, ARMADURA_ARAINCO_PARAM, yes=True)
    activar_armadura_malla(rebar, yes=False)
    stamp_armadura_conjunto_guid(rebar)
    stamp_armadura_eje(rebar)
    return ok


def activar_armadura_arainco_por_ids(doc, rebar_ids):
    """Activa ``Armadura_Arainco`` (+ ``Armadura_Malla`` = No) en una lista de ``ElementId``."""
    if doc is None or not rebar_ids:
        return 0
    n = 0
    for eid in rebar_ids:
        try:
            el = doc.GetElement(eid)
            if activar_armadura_arainco(el):
                n += 1
        except Exception:
            pass
    return n


def stamp_malla_vertical_rebar(rebar):
    """Barras verticales de malla: ``Armadura_Malla`` = Yes, Tipo D.M., orientación V + nivel host."""
    if rebar is None:
        return rebar
    activar_armadura_malla(rebar, yes=True)
    _set_rebar_string_param(rebar, ARMADURA_MALLA_TIPO_PARAM, ARMADURA_MALLA_TIPO_DM)
    _set_rebar_string_param(
        rebar, ARMADURA_MALLA_ORIENTACION_PARAM, ARMADURA_MALLA_ORIENT_V,
    )
    stamp_armadura_conjunto_guid(rebar)
    stamp_armadura_eje(rebar)
    stamp_armadura_nivel_desde_host_muro(rebar)
    return rebar


def stamp_malla_horizontal_rebar(rebar):
    """Barras horizontales de malla: ``Armadura_Malla`` = Yes, orientación H + nivel host."""
    if rebar is None:
        return rebar
    activar_armadura_malla(rebar, yes=True)
    _set_rebar_string_param(
        rebar, ARMADURA_MALLA_ORIENTACION_PARAM, ARMADURA_MALLA_ORIENT_H,
    )
    stamp_armadura_conjunto_guid(rebar)
    stamp_armadura_eje(rebar)
    stamp_armadura_nivel_desde_host_muro(rebar)
    return rebar


def get_armadura_malla_orientacion(rebar):
    """Lee ``Armadura_Malla_Orientacion`` (``V.`` / ``H.``) o ``None``."""
    if rebar is None:
        return None
    p = _find_rebar_parameter(rebar, ARMADURA_MALLA_ORIENTACION_PARAM)
    if p is None:
        return None
    val = None
    try:
        val = p.AsString()
    except Exception:
        pass
    if not val:
        try:
            val = p.AsValueString()
        except Exception:
            pass
    if not val:
        try:
            if p.HasValue and p.StorageType == StorageType.String:
                val = p.AsString()
        except Exception:
            pass
    if not val:
        return None
    try:
        t = unicode(val).strip()
    except Exception:
        try:
            t = str(val or u"").strip()
        except Exception:
            return None
    tl = t.lower().replace(u"\u00a0", u" ")
    if tl in (ARMADURA_MALLA_ORIENT_V.lower(), u"v", u"v."):
        return u"vertical"
    if tl in (ARMADURA_MALLA_ORIENT_H.lower(), u"h", u"h."):
        return u"horizontal"
    if tl.startswith(u"v"):
        return u"vertical"
    if tl.startswith(u"h"):
        return u"horizontal"
    return None
