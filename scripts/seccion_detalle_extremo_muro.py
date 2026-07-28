# -*- coding: utf-8 -*-
"""
Sección detalle extremo de muro — motor Revit API.

Revit 2024+ | pyRevit | IronPython 2.7 / 3.4

Flujo:
1. Vista padre = Elevation o Building Section.
2. Pick muro.
3. Tipo Detail desde parámetro «Section Filter» de la vista padre
   (nombre del ViewFamilyType Detail que contenga ese texto).
4. Detail View en el extremo: **corte horizontal mirando hacia abajo**
   (el muro se ve en planta) para detallar armadura de confinamiento.

Respaldo canónico: ``BIMTools.extension/scripts/``.
Sincronizar con ``05_SeccionDetalleExtremoMuro.pushbutton/scripts/``.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BoundingBoxXYZ,
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    FilterElement,
    StorageType,
    Transaction,
    Transform,
    UnitTypeId,
    UnitUtils,
    View,
    ViewFamily,
    ViewFamilyType,
    ViewSection,
    ViewType,
    Wall,
    XYZ,
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

try:
    from Autodesk.Revit.DB.Structure import Rebar
except Exception:
    Rebar = None

_DIALOG_TITLE = u"Arainco: Sección detalle extremo muro"
_TX_CREAR = u"Arainco: Sección detalle extremo muro"
_PARAM_SECTION_FILTER = u"Section Filter"
_FAR_CLIP_MM_DEF = 400.0
_LONGITUD_DETALLE_MM_DEF = 1000.0  # alcance del crop a lo largo del muro desde el extremo
_MARGEN_EXTREMO_MM_DEF = 50.0  # margen más allá del extremo (afuera)
_ALTURA_MARGEN_MM_DEF = 150.0
_NEAR_CLIP_MM = 50.0
_ANCHO_MARGEN_MM_DEF = 150.0  # margen extra al espesor (far clip / espesor)

_VISTA_BUILDING_MARKERS = (
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


def _mm(mm):
    return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)


def _to_mm(internal):
    return UnitUtils.ConvertFromInternalUnits(float(internal), UnitTypeId.Millimeters)


def _normalize_compare_name(name):
    s = _as_unicode(name).strip()
    if not s:
        return u""
    try:
        import unicodedata

        s = unicodedata.normalize(u"NFC", s)
    except Exception:
        pass
    return s


def _vector_unitario(v):
    if v is None:
        return None
    try:
        ln = float(v.GetLength())
        if ln < 1e-12:
            return None
        return XYZ(v.X / ln, v.Y / ln, v.Z / ln)
    except Exception:
        return None


def _enum_equals(a, b):
    try:
        return int(a) == int(b)
    except Exception:
        try:
            return a == b
        except Exception:
            return False


def mostrar_aviso(instruction, content=u"", uiapp=None):
    """Aviso WPF estándar; respaldo a TaskDialog."""
    try:
        from bimtools_instruction_dialog import show_message_dialog
        from revit_wpf_window_position import revit_main_hwnd

        hwnd = revit_main_hwnd(uiapp) if uiapp is not None else None
        show_message_dialog(
            _DIALOG_TITLE,
            instruction=_as_unicode(instruction),
            content=_as_unicode(content) if content else None,
            ok_text=u"Entendido",
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass
    msg = _as_unicode(instruction)
    if content:
        msg = u"{0}\n\n{1}".format(msg, _as_unicode(content))
    try:
        TaskDialog.Show(_DIALOG_TITLE, msg)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Vista padre
# ---------------------------------------------------------------------------


def _nombre_es_building_section(nombre):
    n = _as_unicode(nombre).strip().lower()
    if not n:
        return False
    for tok in _VISTA_BUILDING_MARKERS:
        if tok in n:
            return True
    return False


def _view_family_type_element(view):
    if view is None:
        return None
    try:
        tid = view.GetTypeId()
        if tid is None or tid == ElementId.InvalidElementId:
            return None
        return view.Document.GetElement(tid)
    except Exception:
        return None


def es_vista_partida_valida(view):
    """
    True si la vista activa es Elevation o Building Section
    (punto de partida del flujo).
    """
    if view is None:
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass
    try:
        vt = view.ViewType
    except Exception:
        return False
    if _enum_equals(vt, ViewType.Elevation):
        return True
    if _enum_equals(vt, ViewType.Detail):
        return False
    if not _enum_equals(vt, ViewType.Section):
        return False
    vft = _view_family_type_element(view)
    if vft is not None:
        try:
            if vft.ViewFamily != ViewFamily.Section:
                return False
        except Exception:
            pass
        try:
            if _nombre_es_building_section(vft.Name):
                return True
        except Exception:
            pass
    # Section sin marcador explícito de Detail → aceptar como Building Section
    try:
        n = _as_unicode(vft.Name if vft is not None else u"").lower()
        if u"detail" in n and u"building" not in n:
            return False
    except Exception:
        pass
    return True


def leer_section_filter_texto(document, view):
    """
    Obtiene el texto usable de «Section Filter» en la vista padre.

    Returns:
        (texto, None) o (None, mensaje_error)
    """
    if document is None or view is None:
        return None, u"Vista o documento no válidos."
    p = view.LookupParameter(_PARAM_SECTION_FILTER)
    if p is None:
        return None, (
            u"No se encontró el parámetro «Section Filter» en la vista activa. "
            u"Debe existir en la categoría Vistas (instancia)."
        )
    if not p.HasValue:
        return None, u"«Section Filter» está vacío en la vista activa."

    try:
        if p.StorageType == StorageType.ElementId:
            eid = p.AsElementId()
            if eid is None or eid == ElementId.InvalidElementId:
                return None, u"«Section Filter» no tiene referencia válida."
            el = document.GetElement(eid)
            if el is None:
                return None, u"No se encontró el elemento referenciado por «Section Filter»."
            if isinstance(el, ViewFamilyType):
                name = _normalize_compare_name(
                    getattr(el, "Name", None) or _view_family_type_display_name(el)
                )
                if not name:
                    return None, u"El ViewFamilyType de «Section Filter» no tiene nombre."
                return name, None
            if isinstance(el, FilterElement):
                fn = _normalize_compare_name(getattr(el, "Name", None))
                if not fn:
                    return None, u"El filtro referenciado por «Section Filter» no tiene nombre."
                return fn, None
            return None, (
                u"«Section Filter» apunta a un elemento que no es tipo de vista ni filtro."
            )

        if p.StorageType == StorageType.String:
            s = p.AsString()
            if s is None or not str(s).strip():
                return None, u"«Section Filter» (texto) está vacío."
            return _normalize_compare_name(s), None

        vs = None
        try:
            vs = p.AsValueString()
        except Exception:
            pass
        if vs and str(vs).strip():
            return _normalize_compare_name(vs), None
    except Exception as ex:
        return None, u"Error al leer «Section Filter»: {0}".format(_as_unicode(ex))

    return None, u"No se pudo interpretar «Section Filter»."


# ---------------------------------------------------------------------------
# ViewFamilyType Detail desde Section Filter
# ---------------------------------------------------------------------------


def _view_family_type_display_name(vft):
    if vft is None:
        return u""
    try:
        n = vft.Name
        if n:
            s = _as_unicode(n).strip()
            if s:
                return s
    except Exception:
        pass
    for bip in (
        BuiltInParameter.ALL_MODEL_TYPE_NAME,
        BuiltInParameter.SYMBOL_NAME_PARAM,
    ):
        try:
            p = vft.get_Parameter(bip)
            if p and p.HasValue:
                s = _as_unicode(p.AsString()).strip()
                if s:
                    return s
        except Exception:
            continue
    return u""


def _iter_view_family_types_detail(document):
    col = FilteredElementCollector(document).OfClass(ViewFamilyType)
    try:
        col = col.WhereElementIsElementType()
    except Exception:
        pass
    for vft in col:
        try:
            if vft is not None and vft.ViewFamily == ViewFamily.Detail:
                yield vft
        except Exception:
            continue


def find_view_family_type_detail_by_name(document, filter_text):
    """
    Busca ``ViewFamilyType`` con ``ViewFamily.Detail``:

    1. Nombre exacto (normalizado).
    2. Sin distinguir mayúsculas.
    3. El nombre del tipo **contiene** el texto del filtro
       (p. ej. ``Detail (01_ESTRUCTURA_A)`` para ``01_ESTRUCTURA_A``).
       Si hay varias, la de nombre más corto.
    """
    target = _normalize_compare_name(filter_text)
    if not target:
        return None, u"«Section Filter» no tiene texto válido para buscar el tipo Detail."
    for vft in _iter_view_family_types_detail(document):
        try:
            if _normalize_compare_name(_view_family_type_display_name(vft)) == target:
                return vft, None
        except Exception:
            continue
    tl = target.lower()
    for vft in _iter_view_family_types_detail(document):
        try:
            if _view_family_type_display_name(vft).lower() == tl:
                return vft, None
        except Exception:
            continue
    contains_matches = []
    for vft in _iter_view_family_types_detail(document):
        try:
            n = _view_family_type_display_name(vft)
            if not n:
                continue
            if tl in n.lower():
                contains_matches.append((len(n), vft))
        except Exception:
            continue
    if contains_matches:
        contains_matches.sort(key=lambda x: x[0])
        return contains_matches[0][1], None
    sample = []
    for vft in _iter_view_family_types_detail(document):
        try:
            n = _view_family_type_display_name(vft)
            if n:
                sample.append(n)
        except Exception:
            continue
    sample = sorted(set(sample))[:12]
    msg = (
        u"No se encontró un tipo Detail cuyo nombre sea «{0}» o contenga ese texto. "
        u"Cree un ViewFamilyType Detail acorde (p. ej. «Detail ({0})») "
        u"o ajuste «Section Filter»."
    ).format(target)
    if sample:
        msg += u" Ejemplos en el proyecto: {0}.".format(u", ".join(sample))
    return None, msg


def resolver_view_family_type_detail_desde_vista(document, view):
    """
    Resuelve el ``ViewFamilyType`` Detail a partir de «Section Filter» de ``view``.

    Returns:
        ``(ViewFamilyType, section_filter_texto, None)`` o
        ``(None, section_filter_texto_o_None, mensaje_error)``.
    """
    sf_text, err = leer_section_filter_texto(document, view)
    if sf_text is None:
        return None, None, err

    # Si el parámetro apuntaba a un VFT Detail directamente
    p = view.LookupParameter(_PARAM_SECTION_FILTER)
    try:
        if p is not None and p.StorageType == StorageType.ElementId:
            el = document.GetElement(p.AsElementId())
            if isinstance(el, ViewFamilyType):
                try:
                    if el.ViewFamily == ViewFamily.Detail:
                        return el, sf_text, None
                except Exception:
                    pass
    except Exception:
        pass

    vft, err2 = find_view_family_type_detail_by_name(document, sf_text)
    if vft is None:
        return None, sf_text, err2
    return vft, sf_text, None


# ---------------------------------------------------------------------------
# Geometría de muro
# ---------------------------------------------------------------------------


def _location_curve_wall(wall):
    if wall is None:
        return None
    try:
        loc = wall.Location
        if loc is None:
            return None
        return loc.Curve
    except Exception:
        return None


def wall_end_points(wall):
    """
    Returns:
        (p0, p1) XYZ o (None, None)
    """
    curve = _location_curve_wall(wall)
    if curve is None:
        return None, None
    try:
        return curve.GetEndPoint(0), curve.GetEndPoint(1)
    except Exception:
        return None, None


def wall_thickness_ft(wall):
    try:
        w = float(wall.Width)
        if w > 1e-9:
            return w
    except Exception:
        pass
    try:
        p = wall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)
        if p is not None and p.HasValue:
            return float(p.AsDouble())
    except Exception:
        pass
    return None


def wall_height_ft(wall):
    try:
        p = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)
        if p is not None and p.HasValue:
            h = float(p.AsDouble())
            if h > 1e-9:
                return h
    except Exception:
        pass
    try:
        bb = wall.get_BoundingBox(None)
        if bb is not None and bb.Min is not None and bb.Max is not None:
            return float(bb.Max.Z - bb.Min.Z)
    except Exception:
        pass
    return None


def wall_mark(wall):
    if wall is None:
        return u""
    for bip in (BuiltInParameter.ALL_MODEL_MARK, BuiltInParameter.DOOR_NUMBER):
        try:
            p = wall.get_Parameter(bip)
            if p is not None and p.HasValue:
                s = _as_unicode(p.AsString()).strip()
                if s:
                    return s
        except Exception:
            continue
    try:
        return u"Id{0}".format(int(wall.Id.IntegerValue))
    except Exception:
        return u"Muro"


def wall_type_name(wall):
    try:
        wt = wall.WallType
        if wt is not None:
            return _as_unicode(wt.Name).strip()
    except Exception:
        pass
    return u""


def collect_rebar_on_wall(document, wall):
    """Lista de Rebar (o elementos) hospedados en el muro."""
    out = []
    if document is None or wall is None or Rebar is None:
        return out
    try:
        wid = wall.Id
    except Exception:
        return out
    try:
        col = FilteredElementCollector(document).OfClass(Rebar).WhereElementIsNotElementType()
        for r in col:
            try:
                hid = r.GetHostId()
                if hid is not None and hid == wid:
                    out.append(r)
            except Exception:
                continue
    except Exception:
        pass
    return out


def wall_mid_height_z(wall):
    """Cota Z del corte horizontal (mitad de la altura del muro)."""
    try:
        bb = wall.get_BoundingBox(None)
        if bb is not None and bb.Min is not None and bb.Max is not None:
            return 0.5 * (float(bb.Min.Z) + float(bb.Max.Z))
    except Exception:
        pass
    p0, _p1 = wall_end_points(wall)
    ht = wall_height_ft(wall)
    if p0 is not None and ht is not None:
        return float(p0.Z) + 0.5 * float(ht)
    if p0 is not None:
        return float(p0.Z)
    return 0.0


def wall_section_preview_data(document, wall):
    """
    Datos para el canvas de **planta** (corte mirando hacia abajo).

    Returns:
        dict con espesor_mm, longitud_detalle_mm, marca, tipo, n_rebar, ...
    """
    th = wall_thickness_ft(wall)
    ht = wall_height_ft(wall)
    rebars = collect_rebar_on_wall(document, wall)
    n = len(rebars)
    n_vert = max(3, min(8, n if n > 0 else 4))
    return {
        u"marca": wall_mark(wall),
        u"tipo": wall_type_name(wall),
        u"espesor_mm": _to_mm(th) if th else 200.0,
        u"altura_mm": _to_mm(ht) if ht else 3000.0,
        u"longitud_detalle_mm": float(_LONGITUD_DETALLE_MM_DEF),
        u"n_rebar": n,
        u"n_vert_conf": n_vert,
        u"wall_id": int(wall.Id.IntegerValue),
    }


# ---------------------------------------------------------------------------
# Crear Detail View (corte horizontal → planta)
# ---------------------------------------------------------------------------


def _wall_axis_and_normal(wall, end_index):
    """
    Returns:
        (origin_xy XYZ, along_unit XYZ hacia el interior, normal_unit XYZ) o (None, err)
    """
    p0, p1 = wall_end_points(wall)
    if p0 is None or p1 is None:
        return None, None, None, u"El muro no tiene LocationCurve válida."
    if int(end_index) == 0:
        origin_xy = XYZ(p0.X, p0.Y, 0.0)
        along = p1.Subtract(p0)
    else:
        origin_xy = XYZ(p1.X, p1.Y, 0.0)
        along = p0.Subtract(p1)
    along_xy = XYZ(along.X, along.Y, 0.0)
    if float(along_xy.GetLength()) > 1e-9:
        along = along_xy
    along_u = _vector_unitario(along)
    if along_u is None:
        return None, None, None, u"No se pudo obtener la dirección del muro."

    normal = None
    try:
        orient = wall.Orientation
        o_xy = XYZ(orient.X, orient.Y, 0.0)
        if float(o_xy.GetLength()) > 1e-9:
            normal = _vector_unitario(o_xy)
    except Exception:
        normal = None
    if normal is None:
        normal = _vector_unitario(XYZ.BasisZ.CrossProduct(along_u))
    if normal is None:
        return None, None, None, u"No se pudo obtener la normal del muro."
    return origin_xy, along_u, normal, None


def _transform_detail_en_extremo(wall, end_index):
    """
    Triedro para Detail View: **corte horizontal mirando hacia abajo**
    (el extremo del muro se ve en planta — confinamiento).

    Como el marcador de sección horizontal en un alzado con flecha hacia abajo:

    - Origin = extremo del muro a media altura
    - BasisZ = (0, 0, -1) mirada hacia abajo
    - BasisY = hacia el interior del muro (arriba en la hoja ≈ eje del muro)
    - BasisX = normal al muro (espesor en la hoja)
    """
    origin_xy, along_u, normal, err = _wall_axis_and_normal(wall, end_index)
    if along_u is None:
        return None, err

    z_cut = wall_mid_height_z(wall)
    origin = XYZ(origin_xy.X, origin_xy.Y, z_cut)

    bz = XYZ(0.0, 0.0, -1.0)  # look down
    by = along_u  # up on sheet = into wall along axis
    bx = _vector_unitario(by.CrossProduct(bz))
    if bx is None:
        return None, u"Triedro del Detail View inválido."
    # Preferir que BasisX coincida con la normal del muro (o su opuesta)
    if float(bx.DotProduct(normal)) < 0.0:
        bx = XYZ(-bx.X, -bx.Y, -bx.Z)
    # Re-ortogonalizar Y
    by = _vector_unitario(bz.CrossProduct(bx))
    if by is None:
        return None, u"Triedro del Detail View inválido."
    # Mantener Y hacia el interior del muro
    if float(by.DotProduct(along_u)) < 0.0:
        bx = XYZ(-bx.X, -bx.Y, -bx.Z)
        by = _vector_unitario(bz.CrossProduct(bx))
    if by is None:
        return None, u"Triedro del Detail View inválido."

    tr = Transform.Identity
    tr.Origin = origin
    tr.BasisX = bx
    tr.BasisY = by
    tr.BasisZ = bz
    return tr, None


def _bbox_detail_para_muro(
    wall,
    transform,
    far_clip_mm=_FAR_CLIP_MM_DEF,
    longitud_mm=_LONGITUD_DETALLE_MM_DEF,
    margen_extremo_mm=_MARGEN_EXTREMO_MM_DEF,
    margen_espesor_mm=_ANCHO_MARGEN_MM_DEF,
):
    """
    Caja del Detail en planta (mirada hacia abajo):

    - X: espesor del muro (± mitad + margen) — transversal al eje
    - Y: desde fuera del extremo hasta ``longitud_mm`` hacia el interior
    - Z: profundidad del corte vertical (far clip); BasisZ apunta hacia abajo
    """
    th = wall_thickness_ft(wall)
    if th is None or th < 1e-9:
        return None, u"No se pudo obtener el espesor del muro."

    half = 0.5 * th + _mm(margen_espesor_mm)
    xmn = -half
    xmx = half

    ymn = -_mm(margen_extremo_mm)
    ymx = _mm(longitud_mm)

    near_clip = _mm(_NEAR_CLIP_MM)
    far_clip = _mm(far_clip_mm)
    if far_clip <= near_clip:
        far_clip = near_clip + _mm(100.0)
    # Local +Z = mirada (hacia abajo). Near por encima del plano, far por debajo.
    zmn = -near_clip
    zmx = far_clip

    box = BoundingBoxXYZ()
    box.Transform = transform
    box.Min = XYZ(xmn, ymn, zmn)
    box.Max = XYZ(xmx, ymx, zmx)
    for i in range(3):
        try:
            box.SetMinEnabled(i, True)
            box.SetMaxEnabled(i, True)
            continue
        except Exception:
            pass
        try:
            box.set_MinEnabled(i, True)
            box.set_MaxEnabled(i, True)
        except Exception:
            pass
    return box, None


def _nombre_unico_vista(document, base_name):
    base = _as_unicode(base_name).strip() or u"DET. MURO"
    used = set()
    for v in FilteredElementCollector(document).OfClass(View):
        try:
            if v and v.Name:
                used.add(_as_unicode(v.Name).strip())
        except Exception:
            continue
    if base not in used:
        return base
    for i in range(2, 200):
        cand = u"{0} ({1})".format(base, i)
        if cand not in used:
            return cand
    return u"{0} (x)".format(base)


def proposed_view_name(wall, end_index):
    tip = u"INI" if int(end_index) == 0 else u"TER"
    mark = wall_mark(wall) or u"MURO"
    return u"DET. MURO {0} {1}".format(mark, tip)


def crear_detail_extremo_muro(
    document,
    wall,
    end_index,
    vft_detail,
    far_clip_mm=_FAR_CLIP_MM_DEF,
    longitud_mm=_LONGITUD_DETALLE_MM_DEF,
):
    """
    Crea ``ViewSection`` Detail: corte horizontal mirando hacia abajo
    en el extremo del muro (planta de confinamiento).

    Returns:
        (ViewSection, None) o (None, mensaje_error)
    """
    if document is None:
        return None, u"No hay documento."
    if wall is None or not isinstance(wall, Wall):
        return None, u"No hay muro válido."
    if vft_detail is None:
        return None, u"No hay tipo Detail resuelto."
    try:
        if vft_detail.ViewFamily != ViewFamily.Detail:
            return None, u"El tipo resuelto no es ViewFamily.Detail."
    except Exception:
        pass

    tr, err = _transform_detail_en_extremo(wall, end_index)
    if tr is None:
        return None, err
    box, err = _bbox_detail_para_muro(
        wall, tr, far_clip_mm=far_clip_mm, longitud_mm=longitud_mm
    )
    if box is None:
        return None, err

    try:
        vs = ViewSection.CreateDetail(document, vft_detail.Id, box)
    except Exception as ex:
        return None, u"CreateDetail falló: {0}".format(_as_unicode(ex))

    try:
        vs.CropBoxActive = True
    except Exception:
        pass
    try:
        vs.CropBoxVisible = False
    except Exception:
        pass

    name = proposed_view_name(wall, end_index)
    try:
        vs.Name = _nombre_unico_vista(document, name)
    except Exception:
        pass

    return vs, None


def ejecutar_crear_detail_extremo(
    uidoc,
    wall,
    end_index,
    vft_detail,
    far_clip_mm=_FAR_CLIP_MM_DEF,
    longitud_mm=_LONGITUD_DETALLE_MM_DEF,
):
    """
    Transacción + CreateDetail (planta / mirada abajo) + activar vista.

    Returns:
        (True, mensaje) o (False, mensaje_error)
    """
    if uidoc is None:
        return False, u"No hay documento activo."
    doc = uidoc.Document
    t = Transaction(doc, _TX_CREAR)
    t.Start()
    try:
        vs, err = crear_detail_extremo_muro(
            doc,
            wall,
            end_index,
            vft_detail,
            far_clip_mm=far_clip_mm,
            longitud_mm=longitud_mm,
        )
        if vs is None:
            t.RollBack()
            return False, err or u"No se pudo crear el Detail View."
        t.Commit()
    except Exception as ex:
        try:
            t.RollBack()
        except Exception:
            pass
        return False, _as_unicode(ex)

    try:
        uidoc.ActiveView = vs
    except Exception:
        pass
    try:
        vname = _as_unicode(vs.Name)
    except Exception:
        vname = u"Detail"
    tip = u"inicio" if int(end_index) == 0 else u"término"
    return True, u"Creado «{0}» (confinamiento en {1}).".format(vname, tip)


# ---------------------------------------------------------------------------
# Selección
# ---------------------------------------------------------------------------


class _WallSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return elem is not None and isinstance(elem, Wall)

    def AllowReference(self, reference, point):
        return True


def pick_wall(uidoc, uiapp=None):
    """
    PickObject de un muro.

    Returns:
        Wall o None (cancelado)
    """
    if uidoc is None:
        return None
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            _WallSelectionFilter(),
            u"Seleccione un muro en el alzado / Building Section",
        )
    except Exception:
        return None
    try:
        return uidoc.Document.GetElement(ref.ElementId)
    except Exception:
        return None


def run(revit):
    """Entrada: valida vista → pick muro → UI."""
    uiapp = revit
    uidoc = None
    try:
        uidoc = revit.ActiveUIDocument
    except Exception:
        uidoc = None
    if uidoc is None:
        mostrar_aviso(u"No hay documento activo.", uiapp=uiapp)
        return

    view = uidoc.ActiveView
    if not es_vista_partida_valida(view):
        mostrar_aviso(
            u"Active un alzado (Elevation) o una Building Section.",
            content=u"La herramienta parte de esa vista para leer «Section Filter» "
            u"y crear el Detail View del extremo del muro.",
            uiapp=uiapp,
        )
        return

    vft, sf_text, err = resolver_view_family_type_detail_desde_vista(uidoc.Document, view)
    if vft is None:
        mostrar_aviso(
            u"No se pudo resolver el tipo Detail.",
            content=err or u"",
            uiapp=uiapp,
        )
        return

    wall = pick_wall(uidoc, uiapp)
    if wall is None:
        return

    from seccion_detalle_extremo_muro_ui import show_detalle_extremo_window

    show_detalle_extremo_window(
        uiapp,
        wall=wall,
        parent_view=view,
        vft_detail=vft,
        section_filter_text=sf_text,
    )
