# -*- coding: utf-8 -*-
"""
Vistas 3D por eje (Grid) — motor Revit API.

Revit 2024+ | pyRevit | IronPython 3.4

Respaldo de desarrollo en ``BIMTools.extension/scripts/``.
Tras editar aquí, sincronice con
``04_Vistas3DPorEje.pushbutton/scripts/``.

Cada vista 3D representa el eje en elevación: section box en el plano vertical
del Grid (ancho = trazo del eje, alto = extensión vertical 3D del Grid) y cámara
ortográfica mirando perpendicular al trazo (misma convención que «Sección alzado
por eje»). No se usa BBAA del modelo.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BoundingBoxXYZ,
    BuiltInParameter,
    FilteredElementCollector,
    Grid,
    Level,
    Options,
    Transaction,
    TransactionGroup,
    Transform,
    UnitTypeId,
    UnitUtils,
    View,
    View3D,
    ViewFamily,
    ViewFamilyType,
    ViewOrientation3D,
    XYZ,
)

_DIALOG_TITLE = u"Arainco: Vistas 3D por eje"
_TX_CREAR = u"Arainco: Vistas 3D por eje"
_ESPESOR_DEFAULT_MM = 2000.0
_MARGEN_PLANO_MM = 100.0
_CAMARA_DIST_FACTOR = 1.35


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


def listar_tipos_vista_3d(document):
    """``ViewFamilyType`` ThreeDimensional del proyecto, ordenados por nombre."""
    out = []
    col = FilteredElementCollector(document).OfClass(ViewFamilyType)
    try:
        col = col.WhereElementIsElementType()
    except Exception:
        pass
    for vft in col:
        try:
            if vft is None or vft.ViewFamily != ViewFamily.ThreeDimensional:
                continue
            nombre = _view_family_type_display_name(vft)
            if not nombre:
                try:
                    nombre = u"Id {0}".format(vft.Id.IntegerValue)
                except Exception:
                    nombre = u"(sin nombre)"
            out.append((nombre, vft))
        except Exception:
            continue
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


def direccion_alzado_desde_eje(axis_dir):
    """
    Dirección de mirada del alzado: 90° CCW al eje en planta.

    Equivale al lado izquierdo del trazo del Grid (0 → 1), misma regla que
    ``seccion_alzado_eje.direccion_corte_desde_eje``.
    """
    if axis_dir is None:
        return None
    horiz = _vector_unitario(XYZ(axis_dir.X, axis_dir.Y, 0.0))
    if horiz is None:
        return None
    return _vector_unitario(XYZ(-horiz.Y, horiz.X, 0.0))


def _construir_transform_section_box(origen, axis_dir):
    """
    Sistema local del section box = plano de elevación del Grid:

    - BasisZ: mirada del alzado (perpendicular al eje, 90° CCW) — espesor
    - BasisX: horizontal en elevación (Z × mirada), paralelo al eje
    - BasisY: vertical de la elevación

    Coincide con la orientación de una Building Section sobre el mismo eje.
    """
    bz = direccion_alzado_desde_eje(axis_dir)
    if bz is None:
        return None
    bx = _vector_unitario(XYZ.BasisZ.CrossProduct(bz))
    if bx is None:
        return None
    by = _vector_unitario(bz.CrossProduct(bx))
    if by is None:
        return None
    tr = Transform.Identity
    tr.Origin = origen
    tr.BasisX = bx
    tr.BasisY = by
    tr.BasisZ = bz
    return tr


def _extremos_horizontales_grid(grid):
    """Extremos en planta del trazo del Grid (ancho del plano de elevación)."""
    curve = _curva_mas_larga_grid(grid)
    if curve is None:
        return None, None
    try:
        return curve.GetEndPoint(0), curve.GetEndPoint(1)
    except Exception:
        return None, None


def _rango_z_desde_geometria_grid(grid):
    """Z min/max a partir de la geometría 3D del Grid."""
    zs = []
    try:
        opt = Options()
        opt.ComputeReferences = False
        try:
            opt.IncludeNonVisibleObjects = True
        except Exception:
            pass
        geom = grid.get_Geometry(opt)
    except Exception:
        geom = None
    if geom is not None:
        try:
            for go in geom:
                try:
                    # Line / Curve
                    if hasattr(go, u"GetEndPoint"):
                        for i in (0, 1):
                            try:
                                zs.append(float(go.GetEndPoint(i).Z))
                            except Exception:
                                pass
                    # Solid / Mesh: usar bbox del objeto
                    try:
                        bb = go.GetBoundingBox()
                    except Exception:
                        bb = None
                    if bb is not None and bb.Min is not None and bb.Max is not None:
                        zs.append(float(bb.Min.Z))
                        zs.append(float(bb.Max.Z))
                except Exception:
                    continue
        except Exception:
            pass
    if len(zs) >= 2 and (max(zs) - min(zs)) > 1e-6:
        return min(zs), max(zs)
    return None, None


def _rango_z_desde_bbox_grid(grid):
    try:
        bb = grid.get_BoundingBox(None)
    except Exception:
        bb = None
    if bb is None or bb.Min is None or bb.Max is None:
        return None, None
    z0 = float(bb.Min.Z)
    z1 = float(bb.Max.Z)
    if abs(z1 - z0) < 1e-6:
        return None, None
    return min(z0, z1), max(z0, z1)


def _rango_z_desde_niveles(document):
    zs = []
    try:
        for lvl in FilteredElementCollector(document).OfClass(Level):
            try:
                zs.append(float(lvl.Elevation))
            except Exception:
                continue
    except Exception:
        pass
    if len(zs) < 2:
        return None, None
    return min(zs), max(zs)


def _rango_vertical_plano_grid(grid, document):
    """
    Altura del plano generable del Grid: extensión vertical 3D del eje.

    Orden: geometría → bbox del Grid → niveles del proyecto.
    """
    z0, z1 = _rango_z_desde_geometria_grid(grid)
    if z0 is not None:
        return z0, z1
    z0, z1 = _rango_z_desde_bbox_grid(grid)
    if z0 is not None:
        return z0, z1
    return _rango_z_desde_niveles(document)


def _puntos_plano_grid(grid, document):
    """
    Cuatro esquinas del rectángulo vertical del Grid (plano de elevación).

    Ancho = trazo del eje; alto = extensión vertical del Grid.
    """
    p0, p1 = _extremos_horizontales_grid(grid)
    if p0 is None or p1 is None:
        return None, u"No se pudieron obtener los extremos del eje."
    zmin, zmax = _rango_vertical_plano_grid(grid, document)
    if zmin is None or zmax is None:
        return None, u"No se pudo obtener la extensión vertical del eje."
    if abs(zmax - zmin) < 1e-9:
        return None, u"La extensión vertical del eje es nula."
    pts = [
        XYZ(p0.X, p0.Y, zmin),
        XYZ(p0.X, p0.Y, zmax),
        XYZ(p1.X, p1.Y, zmin),
        XYZ(p1.X, p1.Y, zmax),
    ]
    return pts, None


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
        cand = u"3D"
    k = 0
    while cand.lower() in existentes:
        k += 1
        cand = u"{0} ({1})".format(nombre_base, k)
    view.Name = cand


def _extents_section_box_desde_plano_grid(tr, pts_plano, espesor_mm, margen_mm):
    """
    Ancho (X) y alto (Y) del section box = proyección del plano del Grid.
    Espesor (Z) = valor indicado, centrado en el plano del eje.
    """
    if tr is None or not pts_plano:
        return None
    ox = tr.Origin
    bx = tr.BasisX
    by = tr.BasisY
    xs, ys = [], []
    for p in pts_plano:
        d = p - ox
        xs.append(float(d.DotProduct(bx)))
        ys.append(float(d.DotProduct(by)))
    if not xs:
        return None
    m = _mm(margen_mm)
    half_z = 0.5 * _mm(espesor_mm)
    if half_z < 1e-9:
        return None
    return (
        min(xs) - m,
        max(xs) + m,
        min(ys) - m,
        max(ys) + m,
        -half_z,
        half_z,
    )


def _aplicar_orientacion_elevacion(view3d, tr, xmn, xmx, ymn, ymx, espesor_mm):
    """
    Cámara en elevación: mirada horizontal perpendicular al eje (BasisZ),
    arriba = vertical del alzado. Ortográfica (sin perspectiva).
    """
    if view3d is None or tr is None:
        return
    mid_x = 0.5 * (xmn + xmx)
    mid_y = 0.5 * (ymn + ymx)
    half_z = 0.5 * _mm(espesor_mm)
    span_x = max(abs(xmx - xmn), 1.0)
    span_y = max(abs(ymx - ymn), 1.0)
    dist = max(span_x, span_y) * _CAMARA_DIST_FACTOR + half_z
    # Plano del eje en Z local = 0; cámara delante del corte mirando +Z (alzado).
    target_local = XYZ(mid_x, mid_y, 0.0)
    eye_local = XYZ(mid_x, mid_y, -dist)
    target = tr.OfPoint(target_local)
    eye = tr.OfPoint(eye_local)
    forward = _vector_unitario(tr.BasisZ)
    up = _vector_unitario(tr.BasisY)
    if forward is None or up is None:
        forward = _vector_unitario(target - eye)
        up = XYZ.BasisZ
    if forward is None:
        return
    try:
        view3d.IsPerspective = False
    except Exception:
        pass
    try:
        ori = ViewOrientation3D(eye, up, forward)
        view3d.SetOrientation(ori)
    except Exception:
        pass


def crear_vista_3d_por_eje(
    document,
    grid,
    vft_id,
    espesor_mm=_ESPESOR_DEFAULT_MM,
):
    """
    Crea una ``View3D`` en elevación del ``Grid`` (section box + cámara alzado).

    Ancho y alto del section box = plano vertical del Grid (trazo × extensión Z).

    Returns:
        (View3D, None) o (None, mensaje_error)
    """
    if document is None:
        return None, u"No hay documento."
    if grid is None:
        return None, u"No se indicó un eje."
    if vft_id is None:
        return None, u"No se indicó el tipo de vista 3D."
    try:
        espesor = float(espesor_mm)
    except Exception:
        return None, u"Espesor inválido."
    if espesor <= 0:
        return None, u"El espesor del section box debe ser mayor que 0 mm."

    axis_dir, origen, nombre_eje = _direccion_eje_y_origen(grid)
    if axis_dir is None:
        return None, _as_unicode(origen)

    tr = _construir_transform_section_box(origen, axis_dir)
    if tr is None:
        return None, u"No se pudo construir la orientación del section box."

    pts_plano, err_plano = _puntos_plano_grid(grid, document)
    if pts_plano is None:
        return None, err_plano or u"No se pudo definir el plano del eje."

    extents = _extents_section_box_desde_plano_grid(
        tr, pts_plano, espesor, _MARGEN_PLANO_MM
    )
    if extents is None:
        return None, u"No se pudieron calcular los límites del section box."
    xmn, xmx, ymn, ymx, zmn, zmx = extents

    box = BoundingBoxXYZ()
    box.Transform = tr
    box.Min = XYZ(xmn, ymn, zmn)
    box.Max = XYZ(xmx, ymx, zmx)

    try:
        view3d = View3D.CreateIsometric(document, vft_id)
    except Exception as ex:
        return None, u"CreateIsometric falló: {0}".format(_as_unicode(ex))

    try:
        view3d.SetSectionBox(box)
        view3d.IsSectionBoxActive = True
    except Exception as ex:
        return None, u"No se pudo aplicar el section box: {0}".format(_as_unicode(ex))

    _aplicar_orientacion_elevacion(view3d, tr, xmn, xmx, ymn, ymx, espesor)

    label = u"3D ALZ. {0}".format(nombre_eje)
    try:
        _nombre_unico_vista(view3d, document, label)
    except Exception:
        pass

    return view3d, None


def ejecutar_crear_vistas(uidoc, grids, vft, espesor_mm=_ESPESOR_DEFAULT_MM):
    """
    Transacción grupal: una vista 3D por cada Grid.

    Returns:
        (True, mensaje, lista_vistas) o (False, mensaje_error, [])
    """
    if uidoc is None:
        return False, u"No hay documento activo.", []
    doc = uidoc.Document
    if not grids:
        return False, u"Selecciona al menos un eje (Grid).", []
    if vft is None:
        return False, u"Selecciona un tipo de vista 3D.", []

    creadas = []
    errores = []
    tg = TransactionGroup(doc, _TX_CREAR)
    tg.Start()
    try:
        t = Transaction(doc, _TX_CREAR)
        t.Start()
        try:
            for grid in grids:
                vs, err = crear_vista_3d_por_eje(
                    doc, grid, vft.Id, espesor_mm=espesor_mm
                )
                if vs is None:
                    try:
                        nombre = _as_unicode(grid.Name)
                    except Exception:
                        nombre = u"?"
                    errores.append(u"{0}: {1}".format(nombre, err or u"error"))
                else:
                    creadas.append(vs)
            if not creadas:
                t.RollBack()
                tg.RollBack()
                detalle = u"; ".join(errores) if errores else u"sin detalle"
                return False, u"No se creó ninguna vista. {0}".format(detalle), []
            t.Commit()
        except Exception as ex:
            t.RollBack()
            tg.RollBack()
            return False, _as_unicode(ex), []
        tg.Assimilate()
    except Exception as ex:
        try:
            tg.RollBack()
        except Exception:
            pass
        return False, _as_unicode(ex), []

    if creadas:
        try:
            uidoc.ActiveView = creadas[-1]
        except Exception:
            pass

    msg = u"Creadas {0} vista(s) 3D en elevación del eje.".format(len(creadas))
    if errores:
        msg += u" Omitidas {0}: {1}".format(len(errores), u"; ".join(errores[:5]))
    return True, msg, creadas


def mostrar_aviso(uiapp, instruction, content=u""):
    """Diálogo informativo WPF. Respaldo: TaskDialog."""
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
            _DIALOG_TITLE,
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
        TaskDialog.Show(_DIALOG_TITLE, body)
    except Exception:
        pass


def run(revit):
    from vistas_3d_por_eje_ui import show_vistas_3d_por_eje_window

    show_vistas_3d_por_eje_window(revit)
