# -*- coding: utf-8 -*-
"""
Spot elevation en vértices de losas (múltiples Floors).
Referencia de cota: **solo cara superior** del sólido de la losa.

Vértices (unión, luego deduplicados):
- Puntos del ``SlabShapeEditor`` / ``SlabShapeVertices`` (sub-elementos), proyectados a la cara superior.
- **Más** los vértices de aristas de **todas** las caras superiores del sólido (malla). Antes solo se
  leía una cara; el **perímetro** suele repartirse entre varias caras triangulares.
- Para colocar la referencia, si ``Face.Project`` falla en aristas (típico en perímetro), se reintenta
  con micro-desplazamientos en XY (``_NUDGE_FT``). El punto de cota es siempre un
  ``IntersectionResult.XYZPoint`` válido (no proyección solo al plano infinito). Los puntos del
  líder se desplazan en el **plano de la cara** para evitar «Spot Dimension does not lie on its reference».
- Líder en **L** (hombro horizontal + tramo al vértice) con ``View.RightDirection`` / ``UpDirection``;
  constantes ``_SPOT_LEADER_*_FT`` (pies internos).
- Los candidatos se deduplican **sin** proyectar todos al plano de una sola cara superior.
  Tolerancia ``_TOL_VERTEX_SPOT_FT`` (~1 mm) para no duplicar el mismo vértice desde slab y malla;
  tras refinar el punto de referencia se vuelve a filtrar por posición final de cota.
- Al crear cada cota se prueba una cadena de variantes (origen en cara / en vértice, líder en plano o en XY global, sin líder).
- Tras ``Face.Project``, el punto se refina hacia el **interior** del triángulo (hacia el centroide en el plano de la cara);
  se conserva el resultado del mayor desplazamiento válido. Último recurso: ``HostObjectUtils.GetTopFaces`` con la misma geometría de cota.

Se deduplican y se proyectan al plano de la cara. Si la cara superior está partida en varias
caras, por vértice se elige la ``PlanarFace`` cuyo ``Face.Project`` minimiza la distancia.

Opción de instancia: elevación mostrada Top o Bottom (según parámetros del tipo/proyecto).
Tipos: solo Spot Elevation (DimensionStyleType.SpotElevation).
UI alineada con ``bimtools_wpf_dark_theme`` (misma línea que Fundación aislada / Numerar fundaciones).

Creación masiva: las cotas se crean en **lotes de transacciones** dentro de un ``TransactionGroup``
(``Assimilate``) para evitar una sola transacción gigante — causa habitual de inestabilidad o cierre de Revit
con muchos vértices. No se llama a ``Regenerate`` en bucle.
"""

def main(revit):
    import clr

    clr.AddReference("RevitAPI")
    clr.AddReference("RevitAPIUI")
    clr.AddReference("PresentationFramework")
    clr.AddReference("PresentationCore")
    clr.AddReference("WindowsBase")
    clr.AddReference("System")

    import System
    from System.Collections.Generic import List
    from System.Windows.Markup import XamlReader
    from System.Windows import (
        MessageBox,
        MessageBoxButton,
        MessageBoxImage,
        RoutedEventHandler,
        WindowState,
    )
    from System.Windows.Input import Key, KeyEventHandler

    from Autodesk.Revit.DB import (
        BuiltInCategory,
        BuiltInParameter,
        DimensionStyleType,
        FilteredElementCollector,
        Floor,
        GeometryInstance,
        Options,
        ElementId,
        PlanarFace,
        Solid,
        SpotDimension,
        SpotDimensionType,
        StorageType,
        Transaction,
        TransactionGroup,
        ViewPlan,
        ViewType,
        View3D,
        XYZ,
        HostObjectUtils,
    )
    from Autodesk.Revit.UI import TaskDialog, ExternalEvent, IExternalEventHandler
    from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

    from revit_wpf_window_position import (
        position_wpf_window_top_left_at_active_view,
        revit_main_hwnd,
    )
    import os
    from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
    from bimtools_paths import get_logo_paths

    def _try_load_spot_logo(img):
        if img is None:
            return
        from System.IO import FileAccess, FileMode, FileStream
        from System.Windows.Media.Imaging import BitmapCacheOption, BitmapImage

        for path in get_logo_paths():
            if not path or not os.path.isfile(path):
                continue
            stream = None
            try:
                stream = FileStream(path, FileMode.Open, FileAccess.Read)
                bmp = BitmapImage()
                bmp.BeginInit()
                bmp.StreamSource = stream
                bmp.CacheOption = BitmapCacheOption.OnLoad
                bmp.EndInit()
                bmp.Freeze()
                img.Source = bmp
                return
            except Exception:
                continue
            finally:
                if stream is not None:
                    try:
                        stream.Dispose()
                    except Exception:
                        pass

    _APPDOMAIN_WINDOW_KEY = "BIMTools.SpotElevVerticesFloor.ActiveWindow"
    _TOOL_DIALOG_TITLE = u"BIMTools — Spot elevation vértices losa"

    # Desplazamiento del líder (pies internos) — mismo orden de magnitud que script RPS de referencia
    _OFFSET_X_FT = 0.5
    _OFFSET_Y_FT = 0.5
    # Líder en L (hombro horizontal + tramo al vértice) en coordenadas de vista
    _SPOT_LEADER_SHOULDER_FT = 0.55
    _SPOT_LEADER_DIAG_ALONG_FT = 0.38
    _SPOT_LEADER_DIAG_PERP_FT = 0.32
    _TOL_PT = 1e-6
    _TOL_DEDUPE_FT = 1e-4
    # Misma ubicación desde slab + malla o tras refinar: ~1 mm en pies (~0,0033 ft).
    # 1e-4 ft era demasiado estricto y generaba dos Spot Elevation por el mismo vértice.
    _TOL_VERTEX_SPOT_FT = 0.0033
    _TOL_NORMAL = 0.01
    # Lotes pequeños: una transacción con miles de NewSpotElevation agota memoria interna de Revit.
    _SPOT_BATCH_SIZE = 40
    # En aristas entre triángulos / perímetro, Face.Project a veces devuelve None; ~3 mm en pies.
    _NUDGE_FT = 0.01

    def _get_active_window():
        try:
            win = System.AppDomain.CurrentDomain.GetData(_APPDOMAIN_WINDOW_KEY)
        except Exception:
            return None
        if win is None:
            return None
        try:
            _ = win.Title
        except Exception:
            _clear_active_window()
            return None
        try:
            if hasattr(win, "IsLoaded") and (not win.IsLoaded):
                _clear_active_window()
                return None
        except Exception:
            pass
        return win


    def _set_active_window(win):
        try:
            System.AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, win)
        except Exception:
            pass


    def _clear_active_window():
        try:
            System.AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, None)
        except Exception:
            pass


    def _task_dialog_show(message, wpf_window=None):
        if wpf_window is not None:
            try:
                wpf_window.Topmost = False
            except Exception:
                pass
        try:
            TaskDialog.Show(_TOOL_DIALOG_TITLE, message)
        finally:
            if wpf_window is not None:
                try:
                    wpf_window.Topmost = True
                except Exception:
                    pass


    def _collect_floors_from_selection(doc, uidoc):
        """Devuelve lista de Floor desde la selección actual (sin duplicados)."""
        floors = []
        seen = set()
        try:
            for eid in uidoc.Selection.GetElementIds():
                elem = doc.GetElement(eid)
                if elem is None or elem.Category is None:
                    continue
                if int(elem.Category.Id.IntegerValue) != int(BuiltInCategory.OST_Floors):
                    continue
                iid = elem.Id.IntegerValue
                if iid in seen:
                    continue
                seen.add(iid)
                floors.append(elem)
        except Exception:
            pass
        return floors


    def _get_spot_elevation_types(doc):
        """Preferir Spot Elevation; si el proyecto no tiene ninguno filtrado, cualquier SpotDimensionType (script RPS)."""
        types = []
        for st in FilteredElementCollector(doc).OfClass(SpotDimensionType).ToElements():
            try:
                if st.StyleType == DimensionStyleType.SpotElevation:
                    types.append(st)
            except Exception:
                continue
        if not types:
            for st in FilteredElementCollector(doc).OfClass(SpotDimensionType).ToElements():
                types.append(st)
        try:
            types.sort(key=lambda t: t.Id.IntegerValue)
        except Exception:
            pass
        return types

    def _spot_dimension_type_display_name(st):
        """Nombre del tipo Spot Elevation para la UI (IronPython puede fallar con .Name)."""
        try:
            n = st.Name
            if n is not None:
                s = unicode(n).strip()
                if s:
                    return s
        except Exception:
            pass
        for bip in (BuiltInParameter.ALL_MODEL_TYPE_NAME, BuiltInParameter.SYMBOL_NAME_PARAM):
            try:
                p = st.get_Parameter(bip)
                if p is not None and p.HasValue:
                    s = p.AsString()
                    if s and s.strip():
                        return s.strip()
            except Exception:
                continue
        try:
            pi = st.GetType().GetProperty("Name")
            if pi is not None:
                v = pi.GetValue(st, None)
                if v is not None:
                    s = unicode(v).strip()
                    if s:
                        return s
        except Exception:
            pass
        try:
            return u"Spot elevation (Id {0})".format(int(st.Id.IntegerValue))
        except Exception:
            return u"Spot elevation"

    class FloorFilter(ISelectionFilter):
        def AllowElement(self, elem):
            try:
                if elem is None or elem.Category is None:
                    return False
                return int(elem.Category.Id.IntegerValue) == int(BuiltInCategory.OST_Floors)
            except Exception:
                return False

        def AllowReference(self, ref, pt):
            return False


    def _pts_equal(a, b, tol=_TOL_PT):
        return (
            abs(a.X - b.X) < tol
            and abs(a.Y - b.Y) < tol
            and abs(a.Z - b.Z) < tol
        )

    def _snap_xyz_to_planar_face(pt, planar_face):
        """Proyecta un punto al plano de la PlanarFace (refPt/origin sobre la referencia)."""
        if pt is None or planar_face is None or not isinstance(planar_face, PlanarFace):
            return pt
        try:
            n = planar_face.FaceNormal
            o = planar_face.Origin
            v = pt.Subtract(o)
            d = float(v.DotProduct(n))
            return pt.Subtract(n.Multiply(d))
        except Exception:
            return pt

    def _merge_unique_xyz(points, planar_face=None, tol=_TOL_DEDUPE_FT):
        """
        Deduplica por distancia 3D. Si ``planar_face`` se omite, no se proyecta al plano de
        una sola cara: proyectar todo a la primera cara superior desplazaba vértices fuera de los
        triángulos reales y Revit seguía rechazando la referencia.
        """
        merged = []
        for p in points:
            if p is None:
                continue
            if planar_face is None:
                s = p
            else:
                try:
                    s = _snap_xyz_to_planar_face(p, planar_face)
                except Exception:
                    s = p
            if s is None:
                continue
            if not any(_pts_equal(s, v, tol=tol) for v in merged):
                merged.append(s)
        return merged

    def _upward_planar_faces(solid):
        """Caras superiores (normal hacia arriba); en losas editadas suele haber varias (triángulos)."""
        out = []
        if solid is None:
            return out
        up = XYZ(0.0, 0.0, 1.0)
        for face in solid.Faces:
            if not isinstance(face, PlanarFace):
                continue
            try:
                n = face.FaceNormal
                if float(n.DotProduct(up)) > 0.985:
                    out.append(face)
            except Exception:
                continue
        if not out:
            for face in solid.Faces:
                if not isinstance(face, PlanarFace):
                    continue
                try:
                    if float(face.FaceNormal.Z) > 0.15:
                        out.append(face)
                except Exception:
                    continue
        return out

    def _nudge_xy_offsets():
        e = _NUDGE_FT
        h = 0.5 * e
        q = 0.25 * e
        return (
            (q, 0.0),
            (-q, 0.0),
            (0.0, q),
            (0.0, -q),
            (h, 0.0),
            (-h, 0.0),
            (0.0, h),
            (0.0, -h),
            (e, 0.0),
            (-e, 0.0),
            (0.0, e),
            (0.0, -e),
            (e, e),
            (-e, e),
            (e, -e),
            (-e, -e),
            (2.0 * e, 0.0),
            (0.0, 2.0 * e),
            (-2.0 * e, 0.0),
            (0.0, -2.0 * e),
        )

    def _resolve_ref_face_and_point(solid, pt, default_face, upward_faces):
        """
        Devuelve cara + punto **garantizado** sobre la cara (``IntersectionResult.XYZPoint``).
        Nunca usa solo proyección al plano: Revit rechaza «Spot Dimension does not lie on its
        reference» si el punto no está sobre la cara finita (triángulo).
        Se minimiza la distancia al vértice pedido entre todas las caras y micro-desplazamientos.
        """
        if pt is None:
            return None, None
        if solid is None or not upward_faces:
            try:
                return default_face, _snap_xyz_to_planar_face(pt, default_face)
            except Exception:
                return default_face, pt

        try:
            z = float(pt.Z)
        except Exception:
            z = pt.Z

        test_points = [pt]
        for dx, dy in _nudge_xy_offsets():
            test_points.append(XYZ(pt.X + dx, pt.Y + dy, z))

        best_face = None
        best_xyz = None
        best_d = None

        for test_pt in test_points:
            for face in upward_faces:
                if face is None:
                    continue
                try:
                    ir = face.Project(test_pt)
                except Exception:
                    continue
                if ir is None:
                    continue
                try:
                    xyz = ir.XYZPoint
                except Exception:
                    xyz = None
                if xyz is None:
                    continue
                try:
                    d = float(pt.DistanceTo(xyz))
                except Exception:
                    d = 0.0
                if best_d is None or d < best_d - 1e-12:
                    best_d = d
                    best_face = face
                    best_xyz = xyz

        if best_face is not None and best_xyz is not None:
            return best_face, best_xyz

        try:
            return default_face, _snap_xyz_to_planar_face(pt, default_face)
        except Exception:
            return default_face, pt


    def _view_accepts_spot_dimensions(view):
        """Hojas, programaciones, etc. no alojan Spot Elevation."""
        try:
            vt = view.ViewType
        except Exception:
            return True, u""
        if vt in (
            ViewType.DrawingSheet,
            ViewType.Schedule,
            ViewType.Legend,
            ViewType.Rendering,
        ):
            return False, (
                u"La vista activa no puede mostrar Spot Elevation (hoja, leyenda, etc.). "
                u"Abre una planta, alzado, sección o 3D y vuelve a generar."
            )
        return True, u""


    def _plan_view_for_floor(doc, floor):
        """Planta de trabajo cuyo GenLevel coincide con la losa (no plantilla)."""
        lev_int = None
        try:
            lev = doc.GetElement(floor.LevelId)
            if lev is not None:
                lev_int = lev.Id.IntegerValue
        except Exception:
            pass
        if lev_int is not None:
            for vp in FilteredElementCollector(doc).OfClass(ViewPlan).ToElements():
                try:
                    if vp.IsTemplate:
                        continue
                    gl = vp.GenLevel
                    if gl is not None and gl.Id.IntegerValue == lev_int:
                        return vp
                except Exception:
                    continue
        for vp in FilteredElementCollector(doc).OfClass(ViewPlan).ToElements():
            try:
                if not vp.IsTemplate:
                    return vp
            except Exception:
                continue
        return None


    def _resolve_view_for_spot_creation(doc, uidoc, floors):
        """
        Las Spot Elevation son anotaciones 2D por vista. En vista 3D la API suele
        devolver None o no mostrar la marca; usamos la planta del nivel de la losa.
        """
        active = uidoc.ActiveView
        if not floors:
            return active, u""
        try:
            use_3d = isinstance(active, View3D)
        except Exception:
            use_3d = type(active).__name__ == "View3D"
        if not use_3d:
            return active, u""
        pv = _plan_view_for_floor(doc, floors[0])
        if pv is None:
            return None, u""
        note = (
            u"Vista activa: 3D. Las Spot Elevation son de vista: se crearán en "
            u"«{0}» (planta del nivel de la losa) y se mostrará esa vista.\n\n"
        ).format(pv.Name)
        return pv, note


    def _vertices_from_face(planar_face):
        """Todos los extremos de aristas de la cara (p. ej. triangulación en losa editada)."""
        verts = []
        for edge_array in planar_face.EdgeLoops:
            for edge in edge_array:
                crv = edge.AsCurve()
                if crv is None:
                    continue
                for k in (0, 1):
                    try:
                        pt = crv.GetEndPoint(k)
                    except Exception:
                        continue
                    if not any(_pts_equal(pt, v) for v in verts):
                        verts.append(pt)
        return verts

    def _vertices_unique_from_upward_faces(solid):
        """
        Vértices de aristas de **todas** las caras superiores (triángulos de la malla).
        Una sola cara en ``_get_top_face_vertices`` omite el perímetro que cae en otras caras.
        """
        verts = []
        if solid is None:
            return verts
        for face in _upward_planar_faces(solid):
            for pt in _vertices_from_face(face):
                if not any(_pts_equal(pt, v) for v in verts):
                    verts.append(pt)
        return verts

    def _planar_face_centroid(planar_face):
        verts = _vertices_from_face(planar_face)
        if not verts:
            return None
        n = len(verts)
        sx = sum(v.X for v in verts) / n
        sy = sum(v.Y for v in verts) / n
        sz = sum(v.Z for v in verts) / n
        return XYZ(sx, sy, sz)

    def _refine_xyz_tangent_nudges(face, seed_xyz):
        """Pasos en el plano de la cara cuando no hay dirección clara hacia el centroide."""
        if face is None or not isinstance(face, PlanarFace) or seed_xyz is None:
            return seed_xyz
        try:
            n = face.FaceNormal
            refv = XYZ(0, 0, 1)
            u = n.CrossProduct(refv)
            if u.GetLength() < 1e-9:
                u = n.CrossProduct(XYZ(1, 0, 0))
            lu = u.GetLength()
            if lu < 1e-12:
                return seed_xyz
            u = u.Multiply(1.0 / lu)
            v = n.CrossProduct(u)
            lv = v.GetLength()
            if lv < 1e-12:
                return seed_xyz
            v = v.Multiply(1.0 / lv)
            best_t = None
            for eps in (
                1e-6,
                3e-6,
                1e-5,
                3e-5,
                1e-4,
                3e-4,
                0.001,
                0.003,
                0.006,
                0.01,
                0.02,
                0.04,
            ):
                for sx, sy in (
                    (1, 0),
                    (-1, 0),
                    (0, 1),
                    (0, -1),
                    (1, 1),
                    (-1, 1),
                    (1, -1),
                    (-1, -1),
                ):
                    test = seed_xyz.Add(u.Multiply(sx * eps)).Add(v.Multiply(sy * eps))
                    try:
                        ir = face.Project(test)
                    except Exception:
                        continue
                    if ir is None:
                        continue
                    try:
                        p = ir.XYZPoint
                    except Exception:
                        p = None
                    if p is not None:
                        best_t = p
            if best_t is not None:
                return best_t
            return seed_xyz
        except Exception:
            return seed_xyz

    def _refine_xyz_toward_face_interior(face, seed_xyz):
        """
        Desplaza ligeramente hacia el interior del polígono de la cara (hacia el centroide en el plano).
        En aristas, Face.Project es válido pero NewSpotElevation aún puede rechazar; un punto
        un poco más hacia dentro del triángulo suele pasar la validación.
        """
        if face is None or not isinstance(face, PlanarFace) or seed_xyz is None:
            return seed_xyz
        try:
            c = _planar_face_centroid(face)
            if c is None:
                return seed_xyz
            n = face.FaceNormal
            d = c.Subtract(seed_xyz)
            d = d.Subtract(n.Multiply(d.DotProduct(n)))
            ln = d.GetLength()
            if ln < 1e-12:
                return _refine_xyz_tangent_nudges(face, seed_xyz)
            d = d.Multiply(1.0 / ln)
            best = None
            for eps in (
                1e-6,
                3e-6,
                1e-5,
                3e-5,
                1e-4,
                3e-4,
                0.001,
                0.003,
                0.006,
                0.01,
                0.02,
                0.04,
                0.06,
                0.08,
            ):
                test = seed_xyz.Add(d.Multiply(eps))
                try:
                    ir = face.Project(test)
                except Exception:
                    continue
                if ir is None:
                    continue
                try:
                    p = ir.XYZPoint
                except Exception:
                    p = None
                if p is not None:
                    best = p
            if best is not None:
                return best
            ir0 = face.Project(seed_xyz)
            if ir0 is not None:
                try:
                    p0 = ir0.XYZPoint
                except Exception:
                    p0 = None
                if p0 is not None:
                    return p0
            return _refine_xyz_tangent_nudges(face, seed_xyz)
        except Exception:
            return seed_xyz

    def _vertices_xy_from_slab_shape(floor):
        """XY de SlabShapeVertices (mismo conjunto que en Modify Sub-elements / Shape Editing)."""
        if floor is None or not isinstance(floor, Floor):
            return []
        editor = None
        try:
            editor = floor.GetSlabShapeEditor()
            if editor is None or not editor.IsValidObject:
                return []
            coll = getattr(editor, "SlabShapeVertices", None)
            if coll is None:
                return []
            xy_pts = []
            seen = set()
            for sv in coll:
                try:
                    p = sv.Position
                except Exception:
                    continue
                if p is None:
                    continue
                key = (round(p.X, 6), round(p.Y, 6))
                if key not in seen:
                    seen.add(key)
                    xy_pts.append((p.X, p.Y))
            return xy_pts
        except Exception:
            return []
        finally:
            if editor is not None:
                try:
                    editor.Dispose()
                except Exception:
                    pass

    def _project_xy_to_planar_face_xyz(x, y, planar_face):
        """Proyecta (x,y) al plano de la PlanarFace (cara superior de referencia)."""
        if planar_face is None or not isinstance(planar_face, PlanarFace):
            return None
        try:
            n = planar_face.FaceNormal
            p0 = None
            for edge_array in planar_face.EdgeLoops:
                for edge in edge_array:
                    p0 = edge.AsCurve().GetEndPoint(0)
                    break
                if p0 is not None:
                    break
            if p0 is None:
                return None
            if abs(n.Z) > 1e-9:
                z = p0.Z - (n.X * (x - p0.X) + n.Y * (y - p0.Y)) / n.Z
                return XYZ(x, y, z)
            return XYZ(x, y, p0.Z)
        except Exception:
            return None

    def _spot_leader_bend_end(face, origin, ox, oy):
        """
        Puntos del líder en el plano de la PlanarFace. Offsets solo en XY + Z fijo pueden
        salir del plano en triángulos inclinados y contribuir a «does not lie on reference».
        """
        if face is None or not isinstance(face, PlanarFace):
            return (
                XYZ(origin.X + ox, origin.Y + oy, origin.Z),
                XYZ(origin.X + ox * 2, origin.Y + oy * 2, origin.Z),
            )
        try:
            n = face.FaceNormal
            refv = XYZ(0, 0, 1)
            u = n.CrossProduct(refv)
            lu = u.GetLength()
            if lu < 1e-9:
                u = n.CrossProduct(XYZ(1, 0, 0))
                lu = u.GetLength()
            if lu < 1e-12:
                return (
                    XYZ(origin.X + ox, origin.Y + oy, origin.Z),
                    XYZ(origin.X + ox * 2, origin.Y + oy * 2, origin.Z),
                )
            u = u.Multiply(1.0 / lu)
            v = n.CrossProduct(u)
            lv = v.GetLength()
            if lv < 1e-12:
                return (
                    XYZ(origin.X + ox, origin.Y + oy, origin.Z),
                    XYZ(origin.X + ox * 2, origin.Y + oy * 2, origin.Z),
                )
            v = v.Multiply(1.0 / lv)
            bend = origin.Add(u.Multiply(ox)).Add(v.Multiply(oy))
            end = origin.Add(u.Multiply(ox * 2)).Add(v.Multiply(oy * 2))
            return bend, end
        except Exception:
            return (
                XYZ(origin.X + ox, origin.Y + oy, origin.Z),
                XYZ(origin.X + ox * 2, origin.Y + oy * 2, origin.Z),
            )

    def _spot_leader_bend_world(origin, ox, oy):
        """Líder con offset en ejes globales XY (planta); fallback si el líder en plano de cara falla."""
        return (
            XYZ(origin.X + ox, origin.Y + oy, origin.Z),
            XYZ(origin.X + ox * 2, origin.Y + oy * 2, origin.Z),
        )

    def _normalize_xyz(vec):
        if vec is None:
            return None
        try:
            ln = vec.GetLength()
            if ln < 1e-12:
                return None
            return vec.Multiply(1.0 / ln)
        except Exception:
            return None

    def _spot_leader_l_shape_view(view, ref_pt):
        """
        Líder tipo «shoulder» + tramo al punto: ref_pt → bend (diagonal en el plano de la vista)
        → end (hombro horizontal hacia el lado del símbolo).

        Usa RightDirection / UpDirection de la vista (comportamiento habitual en planta).
        """
        if view is None or ref_pt is None:
            return None, None
        try:
            r = _normalize_xyz(view.RightDirection)
            u = _normalize_xyz(view.UpDirection)
            if r is None or u is None:
                return None, None
            # Desde el vértice: diagonal hacia donde suele ir el símbolo (-derecha, -arriba en vista)
            bend = ref_pt.Add(r.Multiply(-_SPOT_LEADER_DIAG_ALONG_FT)).Add(
                u.Multiply(-_SPOT_LEADER_DIAG_PERP_FT)
            )
            # Hombro horizontal adicional (leader shoulder) hacia el símbolo
            end = bend.Add(r.Multiply(-_SPOT_LEADER_SHOULDER_FT))
            return bend, end
        except Exception:
            return None, None

    def _set_spot_elevation_display_mode(spot, show_top):
        """
        Instancia de Spot Elevation: elevación mostrada (tipo/proyecto).
        Enumeración habitual (Elevation Origin / similar): 0=Actual, 1=Top, 2=Bottom, 3=Top&Bottom.
        Tras ello se prueban valores 0/1 por compatibilidad con tipos antiguos binarios.
        """
        if spot is None:
            return
        try:
            if not isinstance(spot, SpotDimension):
                return
        except Exception:
            return

        def _try_int(p, val):
            try:
                if (
                    p
                    and not p.IsReadOnly
                    and p.StorageType == StorageType.Integer
                ):
                    p.Set(int(val))
                    return True
            except Exception:
                pass
            return False

        def _try_parameter_values(p):
            if p is None:
                return False
            if show_top:
                for val in (1, 0, 3):
                    if _try_int(p, val):
                        return True
            else:
                for val in (2, 0, 1, 3):
                    if _try_int(p, val):
                        return True
            return False

        for bip_name in (
            "SPOT_ELEVATION_BASE",
            "SPOT_ELEVATION_BASE_INDEX",
            "SPOT_DIM_ELEVATION_BASE",
        ):
            bip = getattr(BuiltInParameter, bip_name, None)
            if bip is None:
                continue
            try:
                p = spot.get_Parameter(bip)
                if _try_parameter_values(p):
                    return
            except Exception:
                continue

        for nm in (
            u"Elevation Base",
            u"Base de elevación",
            u"Display Elevation",
            u"Elevación mostrada",
        ):
            try:
                p = spot.LookupParameter(nm)
                if _try_parameter_values(p):
                    return
            except Exception:
                continue

        try:
            for p in spot.Parameters:
                if p is None or p.IsReadOnly:
                    continue
                if p.StorageType != StorageType.Integer:
                    continue
                try:
                    dn = p.Definition.Name
                    dnl = (dn or u"").lower()
                except Exception:
                    continue
                if (u"elevation" in dnl or u"elevación" in dnl) and (
                    u"base" in dnl or u"display" in dnl
                ):
                    if _try_parameter_values(p):
                        return
        except Exception:
            pass

    def _get_top_face_vertices(floor, view):
        """
        1) Options.View = vista de creación, primer sólido, cara con normal Z ≈ +1 (cara superior).
        2) Si falla: geometría sin vista, sólido mayor, heurísticas (cara superior más alta).
        """

        def _extract_solid_reference_order(geom_elem, largest):
            """Si largest=False, como el script manual (primer sólido). Si True, el de mayor volumen."""
            if not geom_elem:
                return None
            if not largest:
                solid = None
                for obj in geom_elem:
                    if isinstance(obj, Solid) and obj.Volume > 0:
                        solid = obj
                        break
                    if isinstance(obj, GeometryInstance):
                        for sub in obj.GetInstanceGeometry():
                            if isinstance(sub, Solid) and sub.Volume > 0:
                                solid = sub
                                break
                        if solid is not None:
                            break
                return solid
            _sol_state = [None, -1.0]

            def _take_sol(s):
                if s is None or not isinstance(s, Solid):
                    return
                try:
                    v = float(s.Volume)
                except Exception:
                    return
                if v > _sol_state[1]:
                    _sol_state[1] = v
                    _sol_state[0] = s

            for obj in geom_elem:
                if isinstance(obj, Solid):
                    _take_sol(obj)
                elif isinstance(obj, GeometryInstance):
                    try:
                        for sub in obj.GetInstanceGeometry():
                            _take_sol(sub)
                    except Exception:
                        pass
            return _sol_state[0] if _sol_state[1] > 0 else None

        def _top_face_reference_style(solid):
            """Primera cara con normal hacia arriba (cara superior de la losa)."""
            for face in solid.Faces:
                try:
                    n = face.FaceNormal
                    if abs(n.Z - 1.0) < 0.01:
                        return face
                except Exception:
                    continue
            return None

        up = XYZ(0.0, 0.0, 1.0)

        def _max_z_on_face(face):
            mz = None
            try:
                for edge_array in face.EdgeLoops:
                    for edge in edge_array:
                        ptz = edge.AsCurve().GetEndPoint(0).Z
                        if mz is None or ptz > mz:
                            mz = ptz
            except Exception:
                return None
            return mz

        def _collect_candidates(solid, min_dot, require_up_z):
            out = []
            for face in solid.Faces:
                if not isinstance(face, PlanarFace):
                    continue
                try:
                    n = face.FaceNormal
                except Exception:
                    continue
                try:
                    dot = n.DotProduct(up)
                    if dot < min_dot:
                        continue
                    if require_up_z and n.Z < 0.15:
                        continue
                except Exception:
                    continue
                maxz = _max_z_on_face(face)
                if maxz is None:
                    continue
                out.append((maxz, face))
            return out

        def _fallback_top_face(solid):
            candidates = _collect_candidates(solid, 0.92, False)
            if not candidates:
                candidates = _collect_candidates(solid, 0.75, True)
            if not candidates:
                for face in solid.Faces:
                    if not isinstance(face, PlanarFace):
                        continue
                    try:
                        if abs(face.FaceNormal.Z) < 0.65:
                            continue
                    except Exception:
                        continue
                    maxz = _max_z_on_face(face)
                    if maxz is not None:
                        candidates.append((maxz, face))
            if not candidates:
                for face in solid.Faces:
                    if not isinstance(face, PlanarFace):
                        continue
                    maxz = _max_z_on_face(face)
                    if maxz is not None:
                        candidates.append((maxz, face))
            if not candidates:
                return None
            candidates.sort(key=lambda x: x[0], reverse=True)
            return candidates[0][1]

        # —— Paso A: vista de creación
        if view is not None:
            opts = Options()
            opts.ComputeReferences = True
            opts.IncludeNonVisibleObjects = False
            opts.View = view
            geom_elem = floor.get_Geometry(opts)
            solid = _extract_solid_reference_order(geom_elem, largest=False)
            if solid is not None:
                tf = _top_face_reference_style(solid)
                if tf is not None:
                    verts = _vertices_from_face(tf)
                    if verts:
                        return solid, tf, verts

        # —— Paso B: sin vista + sólido mayor + heurísticas
        opts2 = Options()
        opts2.ComputeReferences = True
        opts2.IncludeNonVisibleObjects = False
        geom_elem2 = floor.get_Geometry(opts2)
        solid2 = _extract_solid_reference_order(geom_elem2, largest=True)
        if solid2 is None:
            return None, None, None
        top_face = _top_face_reference_style(solid2)
        if top_face is None:
            top_face = _fallback_top_face(solid2)
        if top_face is None:
            return solid2, None, None
        verts2 = _vertices_from_face(top_face)
        if not verts2:
            return solid2, top_face, []
        return solid2, top_face, verts2


    def _prepare_spot_specs_for_floors(view, floors):
        """
        Resuelve geometría y puntos **sin** crear elementos (solo lectura).
        Retorna (lista de dicts listos para NewSpotElevation, advertencias).
        """
        specs = []
        advertencias = []
        for fl in floors:
            _solid, top_face, _ = _get_top_face_vertices(fl, view)
            if top_face is None:
                advertencias.append(
                    u"ID {}: sin cara superior.".format(fl.Id.IntegerValue)
                )
                continue
            if top_face.Reference is None:
                advertencias.append(
                    u"ID {}: sin referencia en cara superior.".format(fl.Id.IntegerValue)
                )
                continue

            upward_faces = _upward_planar_faces(_solid)

            xy_slab = _vertices_xy_from_slab_shape(fl)
            candidates = []
            if xy_slab:
                for x, y in xy_slab:
                    pxyz = _project_xy_to_planar_face_xyz(x, y, top_face)
                    if pxyz is not None:
                        candidates.append(pxyz)
            for pt in _vertices_unique_from_upward_faces(_solid):
                candidates.append(pt)
            vertices = _merge_unique_xyz(
                candidates,
                None,
                tol=_TOL_VERTEX_SPOT_FT,
            )

            if not vertices:
                advertencias.append(
                    u"ID {}: sin vértices (sub-elementos ni cara superior).".format(
                        fl.Id.IntegerValue
                    )
                )
                continue

            seen_ref_pts = []

            for pt in vertices:
                ref_face, use_pt = _resolve_ref_face_and_point(
                    _solid, pt, top_face, upward_faces
                )
                if ref_face is None or use_pt is None:
                    advertencias.append(
                        u"ID {}: no se pudo proyectar vértice ({:.2f},{:.2f},{:.2f}) a cara superior."
                        .format(fl.Id.IntegerValue, pt.X, pt.Y, pt.Z)
                    )
                    continue
                try:
                    ref = ref_face.Reference
                except Exception:
                    ref = None
                if ref is None:
                    advertencias.append(
                        u"ID {}: cara superior sin Reference en vértice ({:.2f},{:.2f},{:.2f})."
                        .format(fl.Id.IntegerValue, use_pt.X, use_pt.Y, use_pt.Z)
                    )
                    continue
                use_pt = _refine_xyz_toward_face_interior(ref_face, use_pt)
                ref_pt = use_pt
                if any(
                    _pts_equal(ref_pt, r, tol=_TOL_VERTEX_SPOT_FT) for r in seen_ref_pts
                ):
                    continue
                seen_ref_pts.append(ref_pt)
                bend, end = _spot_leader_l_shape_view(view, use_pt)
                if bend is None or end is None:
                    bend, end = _spot_leader_bend_end(
                        ref_face,
                        use_pt,
                        _OFFSET_X_FT,
                        _OFFSET_Y_FT,
                    )
                specs.append(
                    {
                        u"floor": fl,
                        u"ref": ref,
                        u"vertex_pt": pt,
                        u"origin": use_pt,
                        u"bend": bend,
                        u"end": end,
                        u"ref_pt": ref_pt,
                    }
                )
        return specs, advertencias


    def _create_spots(doc, view, floors, spot_type_id, show_top_elevation=False):
        """
        Crea spot elevations en lotes de transacciones bajo un TransactionGroup.
        Retorna (creados, advertencias).
        """
        specs, advertencias = _prepare_spot_specs_for_floors(view, floors)
        if not specs:
            return 0, advertencias

        creados = 0
        n = len(specs)
        tg = TransactionGroup(doc, u"Spot elevation vértices losa")
        try:
            tg.Start()
            for i in range(0, n, _SPOT_BATCH_SIZE):
                chunk = specs[i : i + _SPOT_BATCH_SIZE]
                i_end = min(i + _SPOT_BATCH_SIZE, n)
                t = Transaction(
                    doc,
                    u"Spot elevation vértices losa (lote {}–{})".format(
                        i + 1,
                        i_end,
                    ),
                )
                t.Start()
                try:
                    for spec in chunk:
                        fl = spec[u"floor"]
                        ref = spec[u"ref"]
                        origin_on_face = spec[u"origin"]
                        bend = spec[u"bend"]
                        end = spec[u"end"]
                        ref_pt = spec[u"ref_pt"]
                        vertex_pt = spec[u"vertex_pt"]
                        spot = None
                        last_ex = None
                        bend_w, end_w = _spot_leader_l_shape_view(view, ref_pt)
                        if bend_w is None or end_w is None:
                            bend_w, end_w = _spot_leader_bend_world(
                                ref_pt,
                                _OFFSET_X_FT,
                                _OFFSET_Y_FT,
                            )
                        for o, b, e, rp, hl in (
                            (origin_on_face, bend, end, ref_pt, True),
                            (vertex_pt, bend, end, ref_pt, True),
                            (origin_on_face, bend_w, end_w, ref_pt, True),
                            (vertex_pt, bend_w, end_w, ref_pt, True),
                            (ref_pt, bend_w, end_w, ref_pt, True),
                            (ref_pt, ref_pt, ref_pt, ref_pt, False),
                            (vertex_pt, vertex_pt, vertex_pt, ref_pt, False),
                        ):
                            try:
                                spot = doc.Create.NewSpotElevation(
                                    view,
                                    ref,
                                    o,
                                    b,
                                    e,
                                    rp,
                                    hl,
                                )
                                if spot is not None:
                                    break
                            except Exception as ex_try:
                                last_ex = ex_try
                                continue
                        if spot is None:
                            try:
                                top_refs = HostObjectUtils.GetTopFaces(fl)
                                if top_refs is not None:
                                    for tr in top_refs:
                                        if tr is None:
                                            continue
                                        for o, b, e, rp, hl in (
                                            (ref_pt, bend_w, end_w, ref_pt, True),
                                            (ref_pt, bend, end, ref_pt, True),
                                            (vertex_pt, bend_w, end_w, ref_pt, True),
                                            (ref_pt, ref_pt, ref_pt, ref_pt, False),
                                        ):
                                            try:
                                                spot = doc.Create.NewSpotElevation(
                                                    view,
                                                    tr,
                                                    o,
                                                    b,
                                                    e,
                                                    rp,
                                                    hl,
                                                )
                                                if spot is not None:
                                                    break
                                            except Exception as ex_try:
                                                last_ex = ex_try
                                                continue
                                        if spot is not None:
                                            break
                            except Exception as ex_host:
                                last_ex = ex_host
                        if spot is None:
                            if last_ex is None:
                                msg = u"NewSpotElevation devolvió None"
                            else:
                                try:
                                    msg = unicode(last_ex)
                                except Exception:
                                    try:
                                        msg = str(last_ex)
                                    except Exception:
                                        msg = u"Error al crear Spot Elevation"
                            advertencias.append(
                                u"ID {} vértice ({:.2f},{:.2f},{:.2f}): {}".format(
                                    fl.Id.IntegerValue,
                                    ref_pt.X,
                                    ref_pt.Y,
                                    ref_pt.Z,
                                    msg,
                                )
                            )
                            continue
                        try:
                            spot.ChangeTypeId(spot_type_id)
                        except Exception:
                            pass
                        _set_spot_elevation_display_mode(spot, show_top_elevation)
                        creados += 1
                    t.Commit()
                except Exception:
                    if t.HasStarted():
                        t.RollBack()
                    raise
            tg.Assimilate()
        except Exception:
            if tg.HasStarted():
                tg.RollBack()
            raise

        return creados, advertencias

    class _SpotElevGenerarHandler(IExternalEventHandler):
        """Generación en contexto API válido; payload por ids (dlg pierde estado tras ShowDialog/reload)."""

        def __init__(self):
            self.floor_ids = []
            self.spot_type_id_int = None
            self.show_top_elevation = False

        def GetName(self):
            return u"BIMTools — Spot elev. vértices losa (generar)"

        def Execute(self, uiapp):
            try:
                uidoc = uiapp.ActiveUIDocument
                if uidoc is None:
                    return
                doc = uidoc.Document
                if not self.floor_ids or self.spot_type_id_int is None:
                    return
                spot_type_id = ElementId(int(self.spot_type_id_int))
                floors = []
                for iid in self.floor_ids:
                    el = doc.GetElement(ElementId(int(iid)))
                    if el and el.IsValidObject:
                        floors.append(el)
                if not floors:
                    TaskDialog.Show(
                        _TOOL_DIALOG_TITLE,
                        u"No se pudieron resolver las losas en el documento.",
                    )
                    return
                view = uidoc.ActiveView
                if view is None:
                    TaskDialog.Show(_TOOL_DIALOG_TITLE, u"No hay vista activa.")
                    return
                ok_view, view_msg = _view_accepts_spot_dimensions(view)
                if not ok_view:
                    TaskDialog.Show(_TOOL_DIALOG_TITLE, view_msg)
                    return
                spot_view, note_3d = _resolve_view_for_spot_creation(doc, uidoc, floors)
                if spot_view is None:
                    TaskDialog.Show(
                        _TOOL_DIALOG_TITLE,
                        u"Estás en vista 3D y no hay una planta (ViewPlan) asociada al nivel de la "
                        u"losa. Crea o abre una planta de ese nivel y ejecuta la herramienta de nuevo "
                        u"(o ejecútala ya en planta).",
                    )
                    return
                ok_sv, sv_msg = _view_accepts_spot_dimensions(spot_view)
                if not ok_sv:
                    TaskDialog.Show(_TOOL_DIALOG_TITLE, sv_msg)
                    return

                creados, advertencias = _create_spots(
                    doc,
                    spot_view,
                    floors,
                    spot_type_id,
                    getattr(self, "show_top_elevation", False),
                )
                if note_3d:
                    try:
                        uidoc.RequestViewChange(spot_view)
                    except Exception:
                        pass
                msg = note_3d if note_3d else u""
                msg += u"Vista de creación: «{}» ({}).\n".format(
                    spot_view.Name, spot_view.ViewType
                )
                msg += u"Se crearon {} Spot Elevation(s).".format(creados)
                if advertencias:
                    msg += u"\n\nAvisos:" + u"\n".join(advertencias[:8])
                    if len(advertencias) > 8:
                        msg += u"\n... (+{} más)".format(len(advertencias) - 8)
                TaskDialog.Show(_TOOL_DIALOG_TITLE, msg)
            except Exception as ex:
                TaskDialog.Show(
                    _TOOL_DIALOG_TITLE + u" — Error",
                    u"Error: {}".format(str(ex)),
                )

    # ── XAML (tema oscuro compartido: Fundación aislada / bimtools_wpf_dark_theme) ─
    _SPOT_XAML = (
        u"""
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Arainco - Spot elevation vértices losa"
    Width="440"
    MaxHeight="720"
    WindowStartupLocation="Manual"
    Background="Transparent"
    AllowsTransparency="True"
    MinHeight="0"
    FontFamily="Segoe UI"
    WindowStyle="None"
    ResizeMode="NoResize"
    Topmost="True"
    UseLayoutRounding="True"
    SizeToContent="Height"
    >
  <Window.Resources>
"""
        + BIMTOOLS_DARK_STYLES_XML
        + u"""
  </Window.Resources>
  <Border x:Name="SpotRootChrome" CornerRadius="10" Background="#0A1A2F" Padding="12"
          BorderBrush="#1A3A4D" BorderThickness="1"
          HorizontalAlignment="Stretch" VerticalAlignment="Top" ClipToBounds="True">
    <Border.Effect>
      <DropShadowEffect Color="#000000" BlurRadius="16" ShadowDepth="0" Opacity="0.35"/>
    </Border.Effect>
    <Grid HorizontalAlignment="Stretch">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <Border x:Name="TitleBar" Grid.Row="0" Background="#0E1B32" CornerRadius="6" Padding="10,8" Margin="0,0,0,10"
              BorderBrush="#21465C" BorderThickness="1" HorizontalAlignment="Stretch">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <Image x:Name="ImgLogo" Width="40" Height="40" Grid.Column="0"
                 Stretch="Uniform" Margin="0,0,10,0" VerticalAlignment="Center"/>
          <StackPanel Grid.Column="1" VerticalAlignment="Center">
            <TextBlock Text="Spot elevation — vértices losa" FontSize="15" FontWeight="SemiBold"
                       Foreground="#E8F4F8"/>
            <TextBlock Text="Referencia: cara superior. Vértices: sub-elementos (Modify Sub-elements) si existen; si no, aristas de la cara superior."
                       FontSize="11" Foreground="#95B8CC" Margin="0,6,0,0" TextWrapping="Wrap"/>
          </StackPanel>
          <Button x:Name="BtnClose" Grid.Column="2" Style="{StaticResource BtnCloseX_MinimalNoBg}"
                  VerticalAlignment="Center" ToolTip="Cerrar"/>
        </Grid>
      </Border>

      <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" MaxHeight="520" Margin="0,0,0,0">
        <StackPanel HorizontalAlignment="Stretch">
          <StackPanel Margin="0,0,0,10" HorizontalAlignment="Stretch">
            <Button x:Name="BtnSeleccionar" Content="Seleccionar losas en modelo"
                    Style="{StaticResource BtnSelectOutline}"
                    HorizontalAlignment="Stretch"/>
            <TextBlock x:Name="TxtCount" Text="0 losa(s)" Style="{StaticResource LabelSmall}" Margin="0,8,0,0"/>
          </StackPanel>

          <GroupBox Style="{StaticResource GbParams}" HorizontalAlignment="Stretch">
            <GroupBox.Header>
              <TextBlock Text="Tipo de Spot Elevation" FontWeight="SemiBold" Foreground="#E8F4F8" FontSize="11"/>
            </GroupBox.Header>
            <ComboBox x:Name="CmbTipo" Style="{StaticResource Combo}" IsEditable="False" IsReadOnly="True"
                      HorizontalAlignment="Stretch">
              <ComboBox.ItemContainerStyle>
                <Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/>
              </ComboBox.ItemContainerStyle>
            </ComboBox>
          </GroupBox>

          <GroupBox Style="{StaticResource GbParams}" HorizontalAlignment="Stretch">
            <GroupBox.Header>
              <TextBlock Text="Elevación mostrada (instancia)" FontWeight="SemiBold" Foreground="#E8F4F8" FontSize="11"/>
            </GroupBox.Header>
            <StackPanel>
              <RadioButton x:Name="RbElevBottom" GroupName="SpotElevDisp" Content="Bottom elevation (inferior)"
                           IsChecked="True" Foreground="#E8F4F8" Margin="0,2,0,4"/>
              <RadioButton x:Name="RbElevTop" GroupName="SpotElevDisp" Content="Top elevation (superior)"
                           Foreground="#E8F4F8" Margin="0,2,0,2"/>
            </StackPanel>
          </GroupBox>
        </StackPanel>
      </ScrollViewer>

      <Grid Grid.Row="2" Margin="0,10,0,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="8"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Button x:Name="BtnCancelar" Grid.Column="0" Content="Cancelar" IsCancel="True"
                Style="{StaticResource BtnSelectOutline}" HorizontalAlignment="Stretch"/>
        <Button x:Name="BtnAceptar" Grid.Column="2" Content="Generar" IsDefault="True"
                Style="{StaticResource BtnPrimary}" HorizontalAlignment="Stretch"/>
      </Grid>
    </Grid>
  </Border>
</Window>
"""
    )


    class SpotElevVerticesDialog(object):
        def __init__(self, doc, uidoc, gen_handler, gen_event):
            self._doc = doc
            self._uidoc = uidoc
            self._gen_handler = gen_handler
            self._gen_event = gen_event
            self._win = XamlReader.Parse(_SPOT_XAML)
            self._floors = _collect_floors_from_selection(doc, uidoc)
            self._spot_types_list = []
            self._spot_type_id = None
            self._generation_accepted = False
            self._generar_lock = False

            self._cmb = self._win.FindName("CmbTipo")
            self._txt_count = self._win.FindName("TxtCount")
            _try_load_spot_logo(self._win.FindName("ImgLogo"))

            btn_sel = self._win.FindName("BtnSeleccionar")
            if btn_sel is not None:
                btn_sel.Click += RoutedEventHandler(self._on_seleccionar)
            # Lambda explícita: en IronPython el enlace a métodos de instancia a veces no dispara Click.
            btn_can = self._win.FindName("BtnCancelar")
            if btn_can is not None:
                btn_can.Click += RoutedEventHandler(
                    lambda s, e: self._on_cancelar(s, e)
                )
            btn_ok = self._win.FindName("BtnAceptar")
            if btn_ok is not None:
                btn_ok.Click += RoutedEventHandler(
                    lambda s, e: self._on_aceptar(s, e)
                )
            btn_close = self._win.FindName("BtnClose")
            if btn_close is not None:
                btn_close.Click += RoutedEventHandler(self._on_cancelar)
            self._win.KeyDown += KeyEventHandler(self._on_key_down)

            try:
                from System.Windows.Input import MouseButtonEventHandler

                title_bar = self._win.FindName("TitleBar")
                if title_bar is not None:
                    title_bar.MouseLeftButtonDown += MouseButtonEventHandler(
                        lambda s, e: self._win.DragMove()
                    )
                if btn_close is not None:
                    btn_close.MouseLeftButtonDown += MouseButtonEventHandler(
                        lambda s, e: setattr(e, "Handled", True)
                    )
            except Exception:
                pass

            # Combo del tema: ancho fijo 110 px; forzar ancho completo para nombres largos.
            try:

                def _on_win_loaded(sender, args):
                    try:
                        from System import Double
                        from System.Windows import HorizontalAlignment

                        if self._cmb:
                            self._cmb.Width = Double.NaN
                            self._cmb.MinWidth = 200
                            self._cmb.MaxWidth = Double.PositiveInfinity
                            self._cmb.HorizontalAlignment = HorizontalAlignment.Stretch
                    except Exception:
                        pass
                    if self._cmb and self._spot_types_list:
                        try:
                            self._cmb.SelectedIndex = 0
                        except Exception:
                            pass

                self._win.Loaded += RoutedEventHandler(_on_win_loaded)
            except Exception:
                pass

            self._load_spot_types()
            self._update_count()

        def _load_spot_types(self):
            self._spot_types_list = _get_spot_elevation_types(self._doc)
            items_net = List[object]()
            seen_labels = {}
            for st in self._spot_types_list:
                base = _spot_dimension_type_display_name(st)
                try:
                    iid = int(st.Id.IntegerValue)
                except BaseException:
                    iid = 0
                if base in seen_labels:
                    label = u"{0} (Id {1})".format(base, iid)
                else:
                    seen_labels[base] = True
                    label = base
                items_net.Add(System.String(label))
            if self._cmb:
                self._cmb.ItemsSource = items_net
                if items_net.Count > 0:
                    self._cmb.SelectedIndex = 0
                    try:
                        self._cmb.UpdateLayout()
                    except Exception:
                        pass
                    self._cmb.SelectedIndex = 0
            if not self._spot_types_list:
                MessageBox.Show(
                    u"No hay tipos de Spot Elevation en el proyecto.\n"
                    u"Carga o crea tipos de cota Spot Elevation (no Spot Coordinate ni Spot Slope).",
                    _TOOL_DIALOG_TITLE,
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning,
                )

        def _update_count(self):
            n = len(self._floors)
            if self._txt_count:
                self._txt_count.Text = u"{} losa(s)".format(n)

        def _on_seleccionar(self, sender, args):
            try:
                self._win.Hide()
                refs = self._uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    FloorFilter(),
                    u"Selecciona una o más losas (Floor)",
                )
                self._floors = []
                seen = set()
                for ref in refs:
                    elem = self._doc.GetElement(ref.ElementId)
                    if elem and elem.IsValidObject:
                        iid = elem.Id.IntegerValue
                        if iid not in seen:
                            seen.add(iid)
                            self._floors.append(elem)
                self._update_count()
            except Exception as ex:
                err = str(ex).lower()
                if "cancel" not in err and "operation" not in err:
                    MessageBox.Show(
                        str(ex),
                        _TOOL_DIALOG_TITLE,
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning,
                    )
            finally:
                self._win.Show()
                self._win.Activate()

        def _on_cancelar(self, sender, args):
            try:
                self._win.Close()
            except Exception:
                pass

        def _on_aceptar(self, sender, args):
            # Evitar doble ejecución: IsDefault + Key.Return en el Window disparaban dos veces.
            if self._generar_lock or self._generation_accepted:
                return
            if not self._floors:
                MessageBox.Show(
                    u"Selecciona al menos una losa Floor.",
                    _TOOL_DIALOG_TITLE,
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning,
                )
                return
            idx = self._cmb.SelectedIndex if self._cmb else -1
            try:
                idx = int(idx)
            except Exception:
                idx = -1
            # Plantilla WPF/IronPython: SelectedIndex puede seguir en -1 pese a mostrar texto.
            if (
                idx < 0
                and self._spot_types_list
                and len(self._spot_types_list) > 0
            ):
                idx = 0
                try:
                    if self._cmb:
                        self._cmb.SelectedIndex = 0
                except Exception:
                    pass
            if idx < 0 or idx >= len(self._spot_types_list):
                MessageBox.Show(
                    u"Selecciona un tipo de Spot Elevation.",
                    _TOOL_DIALOG_TITLE,
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning,
                )
                return
            try:
                self._spot_type_id = self._spot_types_list[idx].Id
            except BaseException:
                MessageBox.Show(
                    u"Tipo no válido.",
                    _TOOL_DIALOG_TITLE,
                    MessageBoxButton.OK,
                    MessageBoxImage.Error,
                )
                return
            self._generar_lock = True
            try:
                rb_top = self._win.FindName("RbElevTop")
                self._gen_handler.show_top_elevation = (
                    rb_top is not None and rb_top.IsChecked == True
                )
            except Exception:
                self._gen_handler.show_top_elevation = False
            try:
                self._gen_handler.floor_ids = [
                    int(f.Id.IntegerValue) for f in self._floors
                ]
                self._gen_handler.spot_type_id_int = int(
                    self._spot_type_id.IntegerValue
                )
            except Exception:
                self._generar_lock = False
                return
            try:
                self._gen_event.Raise()
            except Exception:
                self._generar_lock = False
                TaskDialog.Show(
                    _TOOL_DIALOG_TITLE,
                    u"No se pudo encolar la generación. Intenta de nuevo.",
                )
                return
            self._generation_accepted = True
            try:
                self._win.Close()
            except Exception:
                pass

        def _on_key_down(self, sender, args):
            if args.Key == Key.Escape:
                self._on_cancelar(sender, args)

        def show(self, revit):
            hwnd = None
            try:
                hwnd = revit_main_hwnd(revit.Application)
                # Owner=Revit a veces impide que ShowDialog termine bien al cerrar con DialogResult.
                # Posicionamos respecto a la vista; no fijamos Owner.
            except Exception:
                pass
            position_wpf_window_top_left_at_active_view(self._win, self._uidoc, hwnd)
            return self._win.ShowDialog()


    existing = _get_active_window()
    if existing is not None:
        try:
            if existing.WindowState == WindowState.Minimized:
                existing.WindowState = WindowState.Normal
        except Exception:
            pass
        try:
            existing.Show()
        except Exception:
            pass
        try:
            existing.Activate()
            existing.Focus()
        except Exception:
            pass
        _task_dialog_show(u"La herramienta ya esta en ejecucion.", existing)
        return

    uidoc = revit.ActiveUIDocument
    if uidoc is None:
        TaskDialog.Show(_TOOL_DIALOG_TITLE, u"No hay documento activo.")
        return
    doc = uidoc.Document
    view = doc.ActiveView
    if view is None:
        TaskDialog.Show(_TOOL_DIALOG_TITLE, u"No hay vista activa.")
        return

    _gen_handler = _SpotElevGenerarHandler()
    _gen_event = ExternalEvent.Create(_gen_handler)
    dlg = SpotElevVerticesDialog(doc, uidoc, _gen_handler, _gen_event)
    win = dlg._win

    def _on_closed(sender, args):
        _clear_active_window()

    win.Closed += _on_closed
    _set_active_window(win)
    try:
        dlg.show(revit)
    finally:
        _clear_active_window()

def run(revit):
    """Compat: delega en main."""
    main(revit)
