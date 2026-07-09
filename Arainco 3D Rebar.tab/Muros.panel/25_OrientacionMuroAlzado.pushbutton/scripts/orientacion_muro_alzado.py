# -*- coding: utf-8 -*-
"""
Helpers gráficos de orientación de muros en alzado / sección.

Revit 2024+ | pyRevit | IronPython 3.4

En la vista activa dibuja una flecha centrada en cada muro visible:
  • Dirección de la Location Line (punto 0 → punto 1).

Paquete portable: ``25_OrientacionMuroAlzado.pushbutton/scripts/``.
"""

from __future__ import division

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    Category,
    Color,
    ElementId,
    FilteredElementCollector,
    GraphicsStyle,
    GraphicsStyleType,
    Group,
    Line,
    LocationCurve,
    OverrideGraphicSettings,
    Plane,
    Transaction,
    ViewFamily,
    ViewSchedule,
    ViewSheet,
    ViewType,
    Wall,
    XYZ,
)

from System import Int64

_TITULO = u"Arainco: Helpers orientación muro en alzado"
_TITULO_ACTUALIZAR = u"Arainco: Actualizar helpers orientación muro en alzado"
_PREFIJO_GRUPO = u"Arainco_ORIENT_MURO"
_WIDE_LINES_NAMES = (u"<Wide Lines>", u"Wide Lines")
_COLOR_VERDE = Color(0, 200, 0)
_MIN_LINE_LEN_FT = 1.0 / 304.8
_LEN_MIN_FT = 0.8
_LEN_MAX_FT = 8.0
_LEN_FRACCION_MURO = 0.35

_VISTA_DETALLE_MARKERS = (
    u"detail",
    u"detalle",
    u"callout",
    u"recuadro",
    u"detailed",
)
_BUILDING_SECTION_MARKERS = (
    u"building section",
    u"sección de edificio",
    u"seccion de edificio",
)


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _canon_key(text):
    return _as_unicode(text).strip().lower()


def _view_type_suffix(view):
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


def _enum_equals(valor, enum_obj):
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
        a = _canon_key(valor.ToString() if hasattr(valor, u"ToString") else valor)
        b = _canon_key(
            enum_obj.ToString() if hasattr(enum_obj, u"ToString") else enum_obj
        )
        if a and b and a.split(u".")[-1] == b.split(u".")[-1]:
            return True
    except Exception:
        pass
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
                return _as_unicode(s).strip()
        except Exception:
            pass
        try:
            p = element.get_Parameter(bip)
            if p is None:
                continue
            s = p.AsString()
            if s:
                return _as_unicode(s).strip()
        except Exception:
            pass
    return u""


def _view_family_type_element(view):
    if view is None:
        return None
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
            if vft is not None and hasattr(vft, u"ViewFamily"):
                return vft
    except Exception:
        pass
    return None


def _view_family_type_name(view):
    vft = _view_family_type_element(view)
    if vft is not None:
        try:
            nm = vft.Name or u""
            if nm:
                return _as_unicode(nm)
        except Exception:
            pass
    try:
        raw = _parametro_texto(
            view,
            BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM,
            BuiltInParameter.ALL_MODEL_TYPE_NAME,
            BuiltInParameter.SYMBOL_NAME_PARAM,
        )
        if u":" in raw:
            raw = raw.split(u":", 1)[1].strip()
        if raw:
            return raw
    except Exception:
        pass
    return u""


def _nombre_es_building_section(name):
    n = _canon_key(name or u"")
    if not n:
        return False
    for bad in _VISTA_DETALLE_MARKERS:
        if bad in n:
            return False
    for ok in _BUILDING_SECTION_MARKERS:
        if ok in n:
            return True
    return False


def _vft_es_familia_section(vft):
    if vft is None:
        return False
    try:
        return _enum_equals(vft.ViewFamily, ViewFamily.Section)
    except Exception:
        pass
    try:
        vf = vft.ViewFamily
        s = vf.ToString() if hasattr(vf, u"ToString") else str(vf)
        return u"Section" in (s or u"")
    except Exception:
        return False


def es_vista_building_section(view):
    """True si la vista activa es una sección de edificio (Building Section)."""
    if view is None:
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass
    if _view_type_suffix(view) == u"Detail":
        return False
    try:
        if _enum_equals(view.ViewType, ViewType.Detail):
            return False
    except Exception:
        pass
    if _view_type_suffix(view) != u"Section":
        try:
            if not _enum_equals(view.ViewType, ViewType.Section):
                return False
        except Exception:
            return False

    vft = _view_family_type_element(view)
    if vft is not None:
        if not _vft_es_familia_section(vft):
            return False
        try:
            if _nombre_es_building_section(vft.Name):
                return True
        except Exception:
            pass

    nombre_tipo = _view_family_type_name(view)
    if _nombre_es_building_section(nombre_tipo):
        return True

    if vft is not None and _vft_es_familia_section(vft):
        n = _canon_key(nombre_tipo or u"")
        if n:
            for bad in _VISTA_DETALLE_MARKERS:
                if bad in n:
                    return False
        return True

    return False


def texto_aviso_vista_building_section(view):
    """Devuelve ``(instruction, content)`` para el diálogo WPF de la herramienta."""
    vn = u""
    vt_s = _view_type_suffix(view) or u"desconocido"
    try:
        vn = _as_unicode(view.Name)
    except Exception:
        pass
    vft_name = _view_family_type_name(view)
    instruction = (
        u"Esta herramienta solo se ejecuta en vistas tipo Building Section "
        u"(sección de edificio)."
    )
    content = u"Vista activa: «{0}» ({1}).".format(vn, vt_s)
    if vft_name:
        content += u"\nTipo de vista: «{0}».".format(vft_name)
    content += u"\n\nAbra una sección de edificio antes de continuar."
    return instruction, content


def _mensaje_vista_no_building_section(view):
    instruction, content = texto_aviso_vista_building_section(view)
    return instruction + u"\n\n" + content


def mostrar_aviso(uiapp, instruction, content=u""):
    """Diálogo informativo WPF (estilo BIMTools). Respaldo: TaskDialog nativo."""
    hwnd = None
    try:
        from revit_wpf_window_position import revit_main_hwnd

        if uiapp is not None:
            hwnd = revit_main_hwnd(uiapp)
    except Exception:
        pass
    try:
        from bimtools_instruction_dialog import show_message_dialog

        show_message_dialog(
            _TITULO,
            instruction,
            content=content,
            ok_text=u"Entendido",
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass
    try:
        from Autodesk.Revit.UI import TaskDialog

        body = instruction
        if content:
            body = instruction + u"\n\n" + content
        TaskDialog.Show(_TITULO, body)
    except Exception:
        pass


def _validar_vista_building_section(view):
    if view is None:
        return False, u"No hay vista activa."
    ok_bs = es_vista_building_section(view)
    if not ok_bs:
        return False, _mensaje_vista_no_building_section(view)
    return True, None


def _unit_vector(vec):
    try:
        ln = float(vec.GetLength())
        if ln < 1e-12:
            return None
        return vec.Divide(ln)
    except Exception:
        return None


def _vista_permite_detail_lines(view):
    """Solo admite vistas Building Section."""
    return _validar_vista_building_section(view)


def _plano_vista(view):
    try:
        return Plane.CreateByNormalAndOrigin(view.ViewDirection, view.Origin)
    except Exception:
        return None


def _proyectar_punto(p, plane):
    if p is None or plane is None:
        return None
    try:
        n = _unit_vector(plane.Normal)
        if n is None:
            return None
        dist = float(p.Subtract(plane.Origin).DotProduct(n))
        return p.Subtract(n.Multiply(dist))
    except Exception:
        return None


def _proyectar_vector_en_plano(vec, plane):
    if vec is None or plane is None:
        return None
    try:
        n = _unit_vector(plane.Normal)
        if n is None:
            return None
        comp = float(vec.DotProduct(n))
        return vec.Subtract(n.Multiply(comp))
    except Exception:
        return None


def _perpendicular_en_plano(dir_unit, plane):
    if dir_unit is None or plane is None:
        return None
    try:
        n = _unit_vector(plane.Normal)
        if n is None:
            return None
        return _unit_vector(n.CrossProduct(dir_unit))
    except Exception:
        return None


def _muros_en_vista(doc, view):
    try:
        col = FilteredElementCollector(doc, view.Id)
        col = col.OfCategory(BuiltInCategory.OST_Walls)
        col = col.WhereElementIsNotElementType()
        return [w for w in col.ToElements() if isinstance(w, Wall)]
    except Exception:
        return []


def _direccion_en_plano_desde_location(wall, plane):
    """Dirección de trazado en el plano de la vista (Location Line 0 → 1)."""
    try:
        loc = wall.Location
        if not isinstance(loc, LocationCurve):
            return None
        crv = loc.Curve
        if crv is None or not crv.IsBound:
            return None

        p0 = _proyectar_punto(crv.GetEndPoint(0), plane)
        p1 = _proyectar_punto(crv.GetEndPoint(1), plane)
        if p0 is not None and p1 is not None:
            vec = p1.Subtract(p0)
            if float(vec.GetLength()) >= _MIN_LINE_LEN_FT:
                return _unit_vector(vec)

        if isinstance(crv, Line):
            dir_plano = _proyectar_vector_en_plano(crv.Direction, plane)
            if dir_plano is not None and float(dir_plano.GetLength()) >= _MIN_LINE_LEN_FT:
                return _unit_vector(dir_plano)

        return None
    except Exception:
        return None


def _centro_bbox_muro(wall, view):
    """Centro 3D del muro según BoundingBox en la vista (mitad de altura visible)."""
    bb = None
    try:
        bb = wall.get_BoundingBox(view)
    except Exception:
        pass
    if bb is None:
        try:
            bb = wall.get_BoundingBox(None)
        except Exception:
            pass
    if bb is None:
        return None
    try:
        return XYZ(
            (float(bb.Min.X) + float(bb.Max.X)) * 0.5,
            (float(bb.Min.Y) + float(bb.Max.Y)) * 0.5,
            (float(bb.Min.Z) + float(bb.Max.Z)) * 0.5,
        )
    except Exception:
        return None


def _punto_medio_en_plano(wall, view, plane):
    """
    Centro del helper en el plano de la vista.

    Antes del cambio a «mitad de altura» se usaba solo
    ``Evaluate(0.5)`` de la Location Line proyectado al plano (sin warning de grupo).
    Ahora: centro del BoundingBox del muro **en la vista activa**, proyectado al plano
    (mitad de altura visible, sin mezclar XY de location con Z global).
    """
    pt = _centro_bbox_muro(wall, view)
    if pt is not None:
        proj = _proyectar_punto(pt, plane)
        if proj is not None:
            return proj
    try:
        loc = wall.Location
        if not isinstance(loc, LocationCurve):
            return None
        crv = loc.Curve
        if crv is None:
            return None
        return _proyectar_punto(crv.Evaluate(0.5, True), plane)
    except Exception:
        return None


def _longitud_en_plano(wall, plane):
    try:
        loc = wall.Location
        if not isinstance(loc, LocationCurve):
            return _LEN_MIN_FT
        crv = loc.Curve
        p0 = _proyectar_punto(crv.GetEndPoint(0), plane)
        p1 = _proyectar_punto(crv.GetEndPoint(1), plane)
        if p0 is None or p1 is None:
            return _LEN_MIN_FT
        return float(p0.DistanceTo(p1))
    except Exception:
        return _LEN_MIN_FT


def _nombre_grupo_vista(nombre_vista):
    """Clave legada para grupos de versiones anteriores."""
    base = u"{0}_{1}".format(_PREFIJO_GRUPO, _as_unicode(nombre_vista).strip())
    if base == _PREFIJO_GRUPO + u"_":
        base = u"{0}_Vista".format(_PREFIJO_GRUPO)
    return base


def _get_lines_category(doc):
    try:
        lc = Category.GetCategory(doc, BuiltInCategory.OST_Lines)
        if lc is not None:
            return lc
    except Exception:
        pass
    try:
        return doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines)
    except Exception:
        return None


def _norm_upper(text):
    try:
        return _as_unicode(text).strip().upper()
    except Exception:
        return u""


def _element_id_entero(eid):
    if eid is None:
        return None
    try:
        return int(eid.Value)
    except Exception:
        try:
            return int(eid.IntegerValue)
        except Exception:
            return None


def _element_id_desde_entero(val):
    if val is None:
        return None
    try:
        return ElementId(Int64(int(val)))
    except Exception:
        try:
            return ElementId(int(val))
        except Exception:
            return None


def _element_id_equals(a, b):
    ai = _element_id_entero(a)
    bi = _element_id_entero(b)
    if ai is None or bi is None:
        return False
    return ai == bi


def _graphics_style_id_for_subcategory(sub):
    if sub is None:
        return None
    try:
        gs = sub.GetGraphicsStyle(GraphicsStyleType.Projection)
        if gs is not None:
            return gs.Id
    except Exception:
        pass
    return None


def _line_style_display_name(doc, style_element_id):
    if doc is None or style_element_id is None:
        return u""
    el = doc.GetElement(style_element_id)
    if el is None:
        return u""
    try:
        nm = el.Name
        if nm:
            return _as_unicode(nm).strip()
    except Exception:
        pass
    try:
        cg = getattr(el, u"GraphicsStyleCategory", None)
        if cg is not None:
            return _as_unicode(cg.Name).strip()
    except Exception:
        pass
    return u""


def _resolve_wide_lines_style_id(doc):
    """GraphicsStyle.Id del estilo «Wide Lines» (subcategoría Líneas)."""
    lines_cat = _get_lines_category(doc)
    targets = set()
    for name in _WIDE_LINES_NAMES:
        targets.add(_norm_upper(name))
        bare = name.strip(u"<>").strip()
        if bare:
            targets.add(_norm_upper(bare))

    fallback_id = None
    if lines_cat is not None:
        try:
            for sub in lines_cat.SubCategories:
                try:
                    nm_u = _norm_upper(sub.Name)
                except Exception:
                    continue
                style_id = _graphics_style_id_for_subcategory(sub)
                if style_id is None:
                    continue
                if nm_u in targets:
                    return style_id
                bare_targets = [t.strip(u"<>").strip() for t in targets]
                for bare in bare_targets:
                    if bare and (nm_u == bare or bare in nm_u):
                        fallback_id = style_id
                        break
        except Exception:
            pass

    if fallback_id is not None:
        return fallback_id

    if lines_cat is None:
        return None
    try:
        parent_iv = _element_id_entero(lines_cat.Id)
    except Exception:
        parent_iv = None
    try:
        for gs in FilteredElementCollector(doc).OfClass(GraphicsStyle):
            try:
                cg = getattr(gs, u"GraphicsStyleCategory", None)
                if cg is None:
                    cg = getattr(gs, u"Category", None)
                if cg is None:
                    continue
                if parent_iv is not None:
                    pc = getattr(cg, u"Parent", None)
                    if pc is None or _element_id_entero(pc.Id) != parent_iv:
                        continue
                nm_u = _norm_upper(cg.Name)
                if nm_u in targets:
                    return gs.Id
                for t in targets:
                    bare = t.strip(u"<>").strip()
                    if bare and (nm_u == bare or bare in nm_u):
                        return gs.Id
            except Exception:
                continue
    except Exception:
        pass
    return None


def _pick_applicable_line_style_id(doc, detail_curve, preferred_id):
    """Elige un LineStyleId válido para la DetailCurve (p. ej. Wide Lines)."""
    applicable = []
    try:
        applicable = list(detail_curve.GetLineStyleIds())
    except Exception:
        pass
    if not applicable:
        return preferred_id
    if preferred_id is not None:
        try:
            piv = _element_id_entero(preferred_id)
            for aid in applicable:
                if piv is not None and _element_id_entero(aid) == piv:
                    return aid
        except Exception:
            pass
    for aid in applicable:
        nm_u = _norm_upper(_line_style_display_name(doc, aid))
        for name in _WIDE_LINES_NAMES:
            target = _norm_upper(name)
            bare = target.strip(u"<>").strip()
            if nm_u == target or (bare and (nm_u == bare or bare in nm_u)):
                return aid
    return preferred_id


def _aplicar_line_style(doc, detail_curve, style_id):
    if detail_curve is None or style_id is None:
        return
    style_id = _pick_applicable_line_style_id(doc, detail_curve, style_id)
    try:
        if style_id == ElementId.InvalidElementId:
            return
    except Exception:
        pass
    try:
        detail_curve.LineStyleId = style_id
    except Exception:
        pass


def _aplicar_override_verde_en_vista(view, element_id):
    """Verde solo en la vista activa (Override Graphics in View by Element)."""
    if view is None or element_id is None:
        return
    try:
        ogs = OverrideGraphicSettings()
        ogs.SetProjectionLineColor(_COLOR_VERDE)
        try:
            ogs.SetCutLineColor(_COLOR_VERDE)
        except Exception:
            pass
        view.SetElementOverrides(element_id, ogs)
    except Exception:
        pass


def _crear_detail_line(doc, view, p0, p1, line_style_id=None):
    if p0 is None or p1 is None:
        return None
    try:
        if float(p0.DistanceTo(p1)) < _MIN_LINE_LEN_FT:
            return None
        dc = doc.Create.NewDetailCurve(view, Line.CreateBound(p0, p1))
        _aplicar_line_style(doc, dc, line_style_id)
        return dc
    except Exception:
        return None


def _lineas_helper(centro, dir_unit, perp, arrow_len):
    """Flecha de orientación: eje + punta (sin marca central ni referencias)."""
    tip = centro.Add(dir_unit.Multiply(arrow_len))
    wing = arrow_len * 0.18
    left = tip.Subtract(dir_unit.Multiply(wing)).Add(perp.Multiply(wing * 0.65))
    right = tip.Subtract(dir_unit.Multiply(wing)).Subtract(perp.Multiply(wing * 0.65))
    return [
        (centro, tip),
        (left, tip),
        (right, tip),
    ]


def _dibujar_helpers(doc, view, walls, plane):
    """Crea detail lines (Wide Lines si existe) sin agrupar."""
    detail_ids = []
    dibujados = 0
    wide_style_id = _resolve_wide_lines_style_id(doc)

    for wall in walls:
        centro = _punto_medio_en_plano(wall, view, plane)
        dir_unit = _direccion_en_plano_desde_location(wall, plane)
        if centro is None or dir_unit is None:
            continue

        perp = _perpendicular_en_plano(dir_unit, plane)
        if perp is None:
            continue

        long_muro = _longitud_en_plano(wall, plane)
        arrow_len = max(min(long_muro * _LEN_FRACCION_MURO, _LEN_MAX_FT), _LEN_MIN_FT)
        creadas_muro = 0

        for p0, p1 in _lineas_helper(centro, dir_unit, perp, arrow_len):
            dc = _crear_detail_line(doc, view, p0, p1, wide_style_id)
            if dc is not None:
                detail_ids.append(dc.Id)
                creadas_muro += 1

        if creadas_muro == 0:
            continue

        dibujados += 1

    for eid in detail_ids:
        _aplicar_override_verde_en_vista(view, eid)

    return dibujados, detail_ids


def _element_ids_desde_enteros(id_values):
    ids = []
    for v in id_values or []:
        try:
            if isinstance(v, ElementId):
                ids.append(v)
            else:
                eid = _element_id_desde_entero(v)
                if eid is not None:
                    ids.append(eid)
        except Exception:
            continue
    return ids


def _ids_enteros_desde_element_ids(element_ids):
    out = []
    for eid in element_ids or []:
        iv = _element_id_entero(eid)
        if iv is not None:
            out.append(iv)
    return out


def _eliminar_detail_lines(doc, line_ids):
    eids = _element_ids_desde_enteros(line_ids)
    valid = []
    for eid in eids:
        try:
            if doc.GetElement(eid) is not None:
                valid.append(eid)
        except Exception:
            pass
    if not valid:
        return 0
    try:
        from System.Collections.Generic import List

        col = List[ElementId]()
        for eid in valid:
            col.Add(eid)
        doc.Delete(col)
        return len(valid)
    except Exception:
        pass
    eliminados = 0
    for eid in valid:
        try:
            doc.Delete(eid)
            eliminados += 1
        except Exception:
            pass
    return eliminados


def _limpiar_grupos_legado_vista(doc, view):
    """Elimina grupos Arainco_ORIENT_MURO_* de versiones anteriores (opcional)."""
    base = _nombre_grupo_vista(_as_unicode(view.Name).strip())
    to_delete = []
    try:
        grupos = list(FilteredElementCollector(doc).OfClass(Group).ToElements())
    except Exception:
        return 0
    for grp in grupos:
        try:
            gt = grp.GroupType
            if gt is None:
                continue
            nombre = _as_unicode(gt.Name)
        except Exception:
            continue
        if not nombre.startswith(base):
            continue
        if not _grupo_es_de_vista(doc, view, grp):
            continue
        to_delete.append(grp.Id)
    eliminados = 0
    for gid in to_delete:
        try:
            if doc.GetElement(gid) is None:
                continue
            doc.Delete(gid)
            eliminados += 1
        except Exception:
            pass
    return eliminados


def _grupo_es_de_vista(doc, view, grp):
    if doc is None or view is None or grp is None:
        return False
    try:
        view_id = view.Id
        member_ids = grp.GetMemberIds()
    except Exception:
        return False
    if member_ids is None:
        return False
    try:
        n = int(member_ids.Count)
    except Exception:
        try:
            n = len(member_ids)
        except Exception:
            n = 0
    if n == 0:
        return False
    for eid in member_ids:
        try:
            el = doc.GetElement(eid)
            if el is None:
                continue
            owner_id = getattr(el, u"OwnerViewId", None)
            if owner_id is not None and _element_id_equals(owner_id, view_id):
                return True
        except Exception:
            continue
    return False


def _refresh_view(uidoc, doc):
    try:
        doc.Regenerate()
    except Exception:
        pass
    try:
        uidoc.RefreshActiveView()
    except Exception:
        pass


def resumen_vista(uidoc):
    """Texto de subtítulo para la UI: vista activa y conteo de muros."""
    if uidoc is None:
        return u"No hay documento activo."
    view = uidoc.ActiveView
    ok, msg = _vista_permite_detail_lines(view)
    if not ok:
        return msg
    walls = _muros_en_vista(uidoc.Document, view)
    try:
        vname = _as_unicode(view.Name)
    except Exception:
        vname = u"Vista"
    return u"Vista: {0} · {1} muro(s) visibles.".format(vname, len(walls))


def ejecutar_dibujar(uidoc, line_ids_previos=None, titulo=None, etiqueta_ok=None):
    """
    Dibuja helpers en la vista activa (sin agrupar).

    Returns:
        (ok, mensaje, ids_detail_lines_int)
    """
    titulo_tx = titulo or _TITULO
    etiqueta = etiqueta_ok or u"Helpers"
    if uidoc is None:
        return False, u"No hay documento activo.", []
    doc = uidoc.Document
    view = uidoc.ActiveView

    ok, msg_vista = _vista_permite_detail_lines(view)
    if not ok:
        return False, msg_vista, []

    plane = _plano_vista(view)
    if plane is None:
        return False, u"No se pudo obtener el plano de la vista activa.", []

    walls = _muros_en_vista(doc, view)
    if not walls:
        return False, u"No hay muros visibles en la vista activa.", []

    if line_ids_previos:
        t_del = Transaction(doc, u"Arainco: Limpiar helpers orientación muro")
        t_del.Start()
        try:
            _eliminar_detail_lines(doc, line_ids_previos)
            t_del.Commit()
        except Exception as ex:
            t_del.RollBack()
            return False, _as_unicode(ex), []

    t = Transaction(doc, titulo_tx)
    t.Start()
    try:
        dibujados, detail_ids = _dibujar_helpers(doc, view, walls, plane)
        if not detail_ids:
            t.RollBack()
            return False, u"No se creó ninguna detail line en la vista activa.", []
        t.Commit()
    except Exception as ex:
        t.RollBack()
        return False, _as_unicode(ex), []

    nuevos_ids = _ids_enteros_desde_element_ids(detail_ids)
    _refresh_view(uidoc, doc)
    msg = u"{0}: {1} muro(s), {2} detail line(s).".format(
        etiqueta, dibujados, len(nuevos_ids)
    )
    return True, msg, nuevos_ids


def ejecutar_actualizar_helpers(uidoc, line_ids_previos):
    """
    Redibuja helpers según la orientación actual de los muros en la vista.

    Returns:
        (ok, mensaje, ids_detail_lines_int)
    """
    if not line_ids_previos:
        return (
            False,
            u"No hay helpers en esta sesión. Use «Dibujar helpers» primero.",
            [],
        )
    if uidoc is not None:
        try:
            uidoc.Document.Regenerate()
        except Exception:
            pass
    return ejecutar_dibujar(
        uidoc,
        line_ids_previos,
        titulo=_TITULO_ACTUALIZAR,
        etiqueta_ok=u"Helpers actualizados",
    )


def ejecutar_eliminar_helpers(uidoc, line_ids, refrescar=True, incluir_legado=False):
    """
    Elimina detail lines de helpers de la sesión.

    Returns:
        (ok, mensaje)
    """
    if uidoc is None:
        return False, u"No hay documento activo."
    doc = uidoc.Document
    view = uidoc.ActiveView

    ok, msg_vista = _vista_permite_detail_lines(view)
    if not ok:
        return False, msg_vista

    t = Transaction(doc, u"Arainco: Limpiar helpers orientación muro")
    t.Start()
    try:
        n = _eliminar_detail_lines(doc, line_ids)
        if incluir_legado:
            n += _limpiar_grupos_legado_vista(doc, view)
        t.Commit()
    except Exception as ex:
        t.RollBack()
        return False, _as_unicode(ex)

    if refrescar:
        _refresh_view(uidoc, doc)
    if n == 0:
        return True, u"No había helpers que eliminar."
    return True, u"Helpers eliminados."


def run(revit):
    """Entrada pyRevit: abre la UI WPF."""
    uidoc = None
    try:
        uidoc = revit.ActiveUIDocument
    except Exception:
        pass
    if uidoc is None:
        mostrar_aviso(revit, u"No hay documento activo.")
        return
    if not es_vista_building_section(uidoc.ActiveView):
        instruction, content = texto_aviso_vista_building_section(uidoc.ActiveView)
        mostrar_aviso(revit, instruction, content)
        return

    from orientacion_muro_alzado_ui import show_orientacion_window

    show_orientacion_window(revit)
