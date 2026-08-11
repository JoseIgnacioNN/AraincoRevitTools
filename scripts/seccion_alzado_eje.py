# -*- coding: utf-8 -*-
"""
Sección de alzado por eje (Grid) — motor Revit API.

Revit 2024+ | pyRevit | IronPython 3.4

Respaldo de desarrollo en ``BIMTools.extension/scripts/``.
Tras editar aquí, sincronice con
``02_SeccionAlzadoEje.pushbutton/scripts/``.

Orientación del corte (planta): perpendicular al eje, 90° antihorario respecto
al trazo del Grid (punto 0 → punto 1). Equivale al lado izquierdo de la línea
del eje según la convención de la herramienta.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BoundingBoxXYZ,
    BuiltInCategory,
    BuiltInParameter,
    FilteredElementCollector,
    Grid,
    Transaction,
    Transform,
    UnitTypeId,
    UnitUtils,
    View,
    ViewFamily,
    ViewFamilyType,
    ViewSection,
    XYZ,
)

_DIALOG_TITLE = u"Arainco: Sección alzado por eje"
_TX_CREAR = u"Arainco: Sección alzado por eje"
_FAR_CLIP_OFFSET_MM = 300.0
_MARGEN_MM = 600.0
_NEAR_CLIP_MM = 100.0

_CATS_BBOX = (
    BuiltInCategory.OST_Walls,
    BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_StructuralFraming,
    BuiltInCategory.OST_StructuralColumns,
    BuiltInCategory.OST_StructuralFoundation,
    BuiltInCategory.OST_Roofs,
    BuiltInCategory.OST_Stairs,
    BuiltInCategory.OST_Ramps,
    BuiltInCategory.OST_GenericModel,
    BuiltInCategory.OST_Mass,
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


def _vector_unitario(v):
    if v is None:
        return None
    try:
        ln = v.GetLength()
        if ln < 1e-12:
            return None
        return XYZ(v.X / ln, v.Y / ln, v.Z / ln)
    except Exception:
        return None


def _punto_medio_curva(curve):
    if curve is None:
        return None
    try:
        return curve.Evaluate(0.5, True)
    except Exception:
        try:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            return XYZ(
                (p0.X + p1.X) * 0.5,
                (p0.Y + p1.Y) * 0.5,
                (p0.Z + p1.Z) * 0.5,
            )
        except Exception:
            return None


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


def _iter_view_family_types_section(document):
    col = FilteredElementCollector(document).OfClass(ViewFamilyType)
    try:
        col = col.WhereElementIsElementType()
    except Exception:
        pass
    for vft in col:
        try:
            if vft is not None and vft.ViewFamily == ViewFamily.Section:
                yield vft
        except Exception:
            continue


def _es_tipo_building_section(nombre):
    n = _as_unicode(nombre).strip().lower()
    if not n:
        return False
    tokens = (
        u"building section",
        u"sección de edificio",
        u"seccion de edificio",
        u"coupe de batiment",
    )
    for tok in tokens:
        if tok in n:
            return True
    return False


def listar_tipos_seccion_building(document):
    """
    ``ViewFamilyType`` de familia Section (Building Section) del proyecto.

    Returns:
        lista de ``(nombre, ViewFamilyType)`` ordenada por nombre.
    """
    building = []
    todos = []
    for vft in _iter_view_family_types_section(document):
        try:
            nombre = _view_family_type_display_name(vft)
            if not nombre:
                try:
                    nombre = u"Id {0}".format(vft.Id.IntegerValue)
                except Exception:
                    nombre = u"(sin nombre)"
            todos.append((nombre, vft))
            if _es_tipo_building_section(nombre):
                building.append((nombre, vft))
        except Exception:
            continue
    out = building if building else todos
    try:
        out.sort(key=lambda t: t[0].lower())
    except Exception:
        pass
    return out


def listar_ejes_modelo(document):
    """Todos los ``Grid`` del documento, ordenados por nombre."""
    ejes = []
    try:
        for g in FilteredElementCollector(document).OfClass(Grid):
            if g is None or not isinstance(g, Grid):
                continue
            try:
                nombre = _as_unicode(g.Name).strip()
            except Exception:
                nombre = u""
            if not nombre:
                try:
                    nombre = u"Id {0}".format(g.Id.IntegerValue)
                except Exception:
                    nombre = u"(sin nombre)"
            ejes.append((nombre, g))
    except Exception:
        pass
    try:
        ejes.sort(key=lambda t: t[0].lower())
    except Exception:
        pass
    return ejes


def _curva_mas_larga_grid(grid):
    if grid is None:
        return None
    candidatas = []
    try:
        c3 = grid.Curve
        if c3 is not None:
            candidatas.append(c3)
    except Exception:
        pass
    try:
        c0 = grid.Curve
        if c0 is not None and c0.IsBound:
            candidatas.append(c0)
    except Exception:
        pass
    if not candidatas:
        return None
    try:
        return max(candidatas, key=lambda c: float(c.Length))
    except Exception:
        return candidatas[0]


def _direccion_eje_y_origen(grid):
    """
    Dirección del eje en planta (0 → 1) y origen en el punto medio de la curva.

    Returns:
        (axis_dir, origen, nombre) o (None, None, mensaje_error)
    """
    if grid is None:
        return None, None, u"No se indicó un eje."
    curve = _curva_mas_larga_grid(grid)
    if curve is None:
        return None, None, u"No se pudo obtener la curva del eje."
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
    except Exception:
        return None, None, u"La curva del eje no tiene extremos válidos."
    axis_raw = XYZ(p1.X - p0.X, p1.Y - p0.Y, 0.0)
    axis_dir = _vector_unitario(axis_raw)
    origin = _punto_medio_curva(curve)
    if axis_dir is None or origin is None:
        return None, None, u"No se pudo definir dirección u origen del eje."
    try:
        nombre = _as_unicode(grid.Name).strip()
    except Exception:
        nombre = u""
    if not nombre:
        try:
            nombre = u"Id {0}".format(grid.Id.IntegerValue)
        except Exception:
            nombre = u"Eje"
    return axis_dir, origin, nombre


def direccion_corte_desde_eje(axis_dir):
    """
    Dirección de corte (BasisZ de la sección): 90° CCW al eje en planta.

    Regla visual: el alzado mira hacia el lado izquierdo del trazo del eje
    (de inicio a fin).
    """
    if axis_dir is None:
        return None
    horiz = _vector_unitario(XYZ(axis_dir.X, axis_dir.Y, 0.0))
    if horiz is None:
        return None
    return _vector_unitario(XYZ(-horiz.Y, horiz.X, 0.0))


def _construir_transform(origen, dir_corte):
    bz = dir_corte.Normalize()
    bx = XYZ.BasisZ.CrossProduct(bz)
    if bx.GetLength() < 1e-6:
        bx = XYZ.BasisX.CrossProduct(bz)
    if bx.GetLength() < 1e-6:
        return None
    bx = bx.Normalize()
    by = bz.CrossProduct(bx).Normalize()
    tr = Transform.Identity
    tr.Origin = origen
    tr.BasisX = bx
    tr.BasisY = by
    tr.BasisZ = bz
    return tr


def _esquinas_bbox(bb):
    mn, mx = bb.Min, bb.Max
    return [
        XYZ(mn.X, mn.Y, mn.Z),
        XYZ(mx.X, mn.Y, mn.Z),
        XYZ(mn.X, mx.Y, mn.Z),
        XYZ(mx.X, mx.Y, mn.Z),
        XYZ(mn.X, mn.Y, mx.Z),
        XYZ(mx.X, mn.Y, mx.Z),
        XYZ(mn.X, mx.Y, mx.Z),
        XYZ(mx.X, mx.Y, mx.Z),
    ]


def _union_bbox(a, b):
    if a is None:
        return b
    if b is None:
        return a
    out = BoundingBoxXYZ()
    out.Min = XYZ(
        min(a.Min.X, b.Min.X),
        min(a.Min.Y, b.Min.Y),
        min(a.Min.Z, b.Min.Z),
    )
    out.Max = XYZ(
        max(a.Max.X, b.Max.X),
        max(a.Max.Y, b.Max.Y),
        max(a.Max.Z, b.Max.Z),
    )
    return out


def _bbox_modelo(document):
    """Unión de bounding boxes de categorías estructurales / arquitectónicas."""
    merged = None
    for cat in _CATS_BBOX:
        try:
            for el in (
                FilteredElementCollector(document)
                .OfCategory(cat)
                .WhereElementIsNotElementType()
            ):
                try:
                    bb = el.get_BoundingBox(None)
                except Exception:
                    bb = None
                if bb is None or bb.Min is None or bb.Max is None:
                    continue
                merged = _union_bbox(merged, bb)
        except Exception:
            continue
    if merged is not None:
        return merged
    try:
        for el in FilteredElementCollector(document).WhereElementIsNotElementType():
            try:
                bb = el.get_BoundingBox(None)
            except Exception:
                bb = None
            if bb is None or bb.Min is None or bb.Max is None:
                continue
            merged = _union_bbox(merged, bb)
    except Exception:
        pass
    return merged


def _nombre_unico_vista(view, document, nombre_base):
    existentes = set()
    for v in FilteredElementCollector(document).OfClass(View):
        try:
            if v is None or v.Id == view.Id:
                continue
            n = v.Name
            if n:
                existentes.add(_as_unicode(n).strip().lower())
        except Exception:
            continue
    cand = _as_unicode(nombre_base).strip()
    if not cand:
        cand = u"Sección"
    k = 0
    while cand.lower() in existentes:
        k += 1
        cand = u"{0} ({1})".format(nombre_base, k)
    view.Name = cand


def _aplicar_far_clip_offset_mm(view, mm):
    try:
        val = _mm(mm)
    except Exception:
        return False
    p = None
    try:
        from Autodesk.Revit.DB import ParameterTypeId

        p = view.get_Parameter(ParameterTypeId.ViewerBoundOffsetFar)
    except Exception:
        pass
    if p is None:
        try:
            p = view.get_Parameter(BuiltInParameter.VIEWER_BOUND_OFFSET_FAR)
        except Exception:
            p = None
    if p is None or p.IsReadOnly:
        return False
    try:
        p.Set(val)
        return True
    except Exception:
        return False


def _extents_seccion_desde_puntos(tr, pts, margen_mm, far_clip_mm):
    if tr is None or not pts:
        return None
    ox = tr.Origin
    bx = tr.BasisX
    by = tr.BasisY
    bz = tr.BasisZ
    xs, ys, zs = [], [], []
    for p in pts:
        d = p - ox
        xs.append(float(d.DotProduct(bx)))
        ys.append(float(d.DotProduct(by)))
        zs.append(float(d.DotProduct(bz)))
    if not xs:
        return None
    m = _mm(margen_mm)
    near_clip = _mm(_NEAR_CLIP_MM)
    far_clip = max(max(zs) + m, _mm(far_clip_mm))
    xabs = max(abs(min(xs)), abs(max(xs))) + m
    ymn_raw = min(ys) - m
    ymx_raw = max(ys) + m
    ymid = 0.5 * (ymn_raw + ymx_raw)
    if abs(ymid) > 1e-9:
        tr.Origin = ox.Add(by.Multiply(ymid))
    yabs = max(abs(ymn_raw - ymid), abs(ymx_raw - ymid))
    return tr, (-xabs, xabs, -yabs, yabs, -near_clip, far_clip)


def crear_seccion_alzado(document, grid, vft_id, far_clip_mm=_FAR_CLIP_OFFSET_MM):
    """
    Crea una ``ViewSection`` de alzado a partir del ``Grid`` indicado.

    Returns:
        (ViewSection, None) o (None, mensaje_error)
    """
    if document is None:
        return None, u"No hay documento."
    if grid is None:
        return None, u"No se indicó un eje."
    if vft_id is None:
        return None, u"No se indicó el tipo de sección."

    axis_dir, origen, nombre_eje = _direccion_eje_y_origen(grid)
    if axis_dir is None:
        return None, _as_unicode(origen)

    dir_corte = direccion_corte_desde_eje(axis_dir)
    if dir_corte is None:
        return None, u"No se pudo calcular la orientación del corte."

    tr = _construir_transform(origen, dir_corte)
    if tr is None:
        return None, u"No se pudo construir la orientación de la sección."

    bb = _bbox_modelo(document)
    if bb is None:
        return None, u"No se pudo obtener la extensión del modelo para dimensionar la sección."

    pts = _esquinas_bbox(bb)
    result = _extents_seccion_desde_puntos(tr, pts, _MARGEN_MM, far_clip_mm)
    if result is None:
        return None, u"No se pudieron calcular los límites de la sección."
    tr, (xmn, xmx, ymn, ymx, zmn, zmx) = result

    box = BoundingBoxXYZ()
    box.Transform = tr
    box.Min = XYZ(xmn, ymn, zmn)
    box.Max = XYZ(xmx, ymx, zmx)

    try:
        vs = ViewSection.CreateSection(document, vft_id, box)
    except Exception as ex:
        return None, u"CreateSection falló: {0}".format(_as_unicode(ex))

    try:
        vs.CropBoxActive = False
    except Exception:
        pass
    try:
        vs.CropBoxVisible = False
    except Exception:
        pass

    _aplicar_far_clip_offset_mm(vs, far_clip_mm)

    label = u"ALZ. {0}".format(nombre_eje)
    try:
        _nombre_unico_vista(vs, document, label)
    except Exception:
        pass

    return vs, None


def ejecutar_crear_seccion(uidoc, grid, vft):
    """
    Transacción + creación. Opcionalmente activa la vista creada.

    Returns:
        (True, mensaje) o (False, mensaje_error)
    """
    if uidoc is None:
        return False, u"No hay documento activo."
    doc = uidoc.Document
    if grid is None:
        return False, u"Selecciona un eje (Grid)."
    if vft is None:
        return False, u"Selecciona un tipo de sección."

    t = Transaction(doc, _TX_CREAR)
    t.Start()
    try:
        vs, err = crear_seccion_alzado(doc, grid, vft.Id)
        if vs is None:
            t.RollBack()
            return False, err or u"No se pudo crear la sección."
        t.Commit()
    except Exception as ex:
        t.RollBack()
        return False, _as_unicode(ex)

    try:
        uidoc.ActiveView = vs
    except Exception:
        pass

    try:
        vname = _as_unicode(vs.Name)
    except Exception:
        vname = u"Sección"
    return True, u"Sección creada: «{0}». Far Clip Offset: {1} mm.".format(
        vname, int(_FAR_CLIP_OFFSET_MM)
    )


def run(revit):
    from seccion_alzado_eje_ui import show_seccion_alzado_window

    show_seccion_alzado_window(revit)
