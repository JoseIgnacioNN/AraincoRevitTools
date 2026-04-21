# -*- coding: utf-8 -*-
"""
Secciones transversales — Fundacion aislada
============================================
Motor   : IronPython 2.7 / 3.4  (RevitPythonShell o pyRevit)
Revit   : 2023-2025

Uso en RPS
----------
Abre este archivo con File > Open o cópialo en la consola interactiva.
Al ejecutarlo, selecciona una sola fundacion aislada (StructuralFoundation)
en el modelo; el script crea dos ViewSection perpendiculares entre si que
cortan el elemento por su centro geometrico:

  Corte A-A  — plano que contiene el eje local Y y el eje Z global
               (se ve el ancho B de la fundacion)
  Corte B-B  — plano que contiene el eje local X y el eje Z global
               (se ve el largo L de la fundacion)

Ambas vistas quedan nombradas segun el parametro "Numeracion Fundacion"
si existe; de lo contrario, se usa el ElementId.

Parametros ajustables (bloque CONFIGURACION):
  MARGEN_MM       — expansion de la caja de vista alrededor del solido
  PROFUNDIDAD_MM  — ancho del corte (distancia Far Clip desde el plano)
  FAR_CLIP_MM     — Far Clip Offset de la vista resultante
  NOMBRE_PREFIJO  — prefijo de los nombres de vista generados
"""

from __future__ import print_function

# ---------------------------------------------------------------------------
# CONFIGURACION — ajustar segun proyecto
# ---------------------------------------------------------------------------
MARGEN_MM = 1200.0       # margen alrededor del solido en la vista (mm)
PROFUNDIDAD_MM = 800.0   # mitad del grosor del corte (mm)
FAR_CLIP_MM = 200.0      # Far Clip Offset de cada vista (mm)
NOMBRE_PREFIJO = u"BIMTools \u2014 Sec. Fund."  # prefijo de nombre de vista

# ---------------------------------------------------------------------------
# Imports Revit API
# En RPS los ensamblados ya estan cargados; clr.AddReference es inocuo.
# ---------------------------------------------------------------------------
import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BoundingBoxXYZ,
    BuiltInCategory,
    BuiltInParameter,
    FamilyInstance,
    FilteredElementCollector,
    GeometryInstance,
    Line,
    LocationPoint,
    Options,
    Solid,
    StorageType,
    Transaction,
    Transform,
    UnitTypeId,
    UnitUtils,
    View,
    ViewDetailLevel,
    ViewFamily,
    ViewFamilyType,
    ViewPlan,
    ViewSection,
    WallFoundation,
    XYZ,
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

# ---------------------------------------------------------------------------
# Contexto Revit  (variables inyectadas por RPS / pyRevit)
# ---------------------------------------------------------------------------
try:
    _uidoc = __revit__.ActiveUIDocument  # noqa: F821
    _doc = _uidoc.Document
except Exception:
    raise RuntimeError(
        "Este script debe ejecutarse dentro de RevitPythonShell o pyRevit."
    )


# ---------------------------------------------------------------------------
# Utilidades internas
# ---------------------------------------------------------------------------

def _mm(mm):
    """Convierte milimetros a unidades internas de Revit (pies)."""
    return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)


def _primer_vft_seccion(document):
    """Devuelve el Id del primer ViewFamilyType de tipo Section."""
    col = FilteredElementCollector(document).OfClass(ViewFamilyType)
    for vft in col:
        try:
            if vft is not None and vft.ViewFamily == ViewFamily.Section:
                return vft.Id
        except Exception:
            continue
    return None


def _nombre_unico(view, document, nombre_base):
    """Asigna a *view* un nombre derivado de *nombre_base* que no colisione."""
    existentes = set()
    for v in FilteredElementCollector(document).OfClass(View):
        try:
            if v is None or v.Id == view.Id:
                continue
            n = v.Name
            if n:
                existentes.add(str(n).strip().lower())
        except Exception:
            continue
    cand = nombre_base
    k = 0
    while cand.strip().lower() in existentes:
        k += 1
        cand = u"{0} ({1})".format(nombre_base, k)
    view.Name = cand


def _leer_numeracion_fundacion(elem):
    """
    Lee el parametro 'Numeracion Fundacion' del elemento.
    Retorna string o None.
    """
    param_names = (
        "Numeracion Fundacion",
        "Numeracion fundacion",
        "Foundation Numbering",
        "Numeracion",
    )
    for name in param_names:
        p = elem.LookupParameter(name)
        if p is None or not p.HasValue:
            continue
        try:
            if p.StorageType == StorageType.String:
                s = p.AsString()
                if s and str(s).strip():
                    return str(s).strip()
            elif p.StorageType == StorageType.Integer:
                return str(p.AsInteger())
            elif p.StorageType == StorageType.Double:
                return str(int(round(p.AsDouble())))
            vs = p.AsValueString()
            if vs and str(vs).strip():
                return str(vs).strip()
        except Exception:
            continue
    return None


def _puntos_geometria(elem):
    """
    Devuelve lista de XYZ (vertices de los solidos del elemento).
    Se usa para calcular los extents reales en el sistema local del corte.
    """
    pts = []
    opt = Options()
    try:
        opt.DetailLevel = ViewDetailLevel.Fine
    except Exception:
        pass
    try:
        geom = elem.get_Geometry(opt)
        if geom is None:
            return pts
        for obj in geom:
            solids = []
            if isinstance(obj, Solid):
                if obj.Volume > 1e-9:
                    solids.append(obj)
            elif isinstance(obj, GeometryInstance):
                try:
                    for sub in obj.GetInstanceGeometry():
                        if isinstance(sub, Solid) and sub.Volume > 1e-9:
                            solids.append(sub)
                except Exception:
                    pass
            for s in solids:
                try:
                    for face in s.Faces:
                        try:
                            mesh = face.Triangulate()
                            for v in mesh.Vertices:
                                pts.append(v)
                        except Exception:
                            pass
                except Exception:
                    pass
    except Exception:
        pass
    return pts


def _esquinas_bbox(bb):
    """Devuelve las 8 esquinas del BoundingBoxXYZ como lista de XYZ."""
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


def _construir_transform(origen, dir_corte):
    """
    Construye el Transform para una ViewSection.

    Convencion Revit:
      BasisZ  = dir_corte        (direccion de corte / profundidad)
      BasisX  = Z_global x BasisZ  (horizontal en la vista, hacia la derecha)
      BasisY  = BasisZ x BasisX    (vertical en la vista, hacia arriba ~ Z global)

    Esta es identica a la logica de _transform_seccion_transversal_punto_medio
    usada en vista_seccion_enfierrado_vigas.py.
    """
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


def _crear_seccion(document, vft_id, elem, origen, dir_corte, label):
    """
    Crea una ViewSection que corta *elem* en *origen* mirando hacia *dir_corte*.

    Calcula los extents proyectando los vertices de la geometria sobre el
    sistema local del corte y agrega MARGEN_MM en todas las direcciones.
    Usa las 8 esquinas del BoundingBox como respaldo si no hay vertices.
    """
    tr = _construir_transform(origen, dir_corte)
    if tr is None:
        return None, u"No se pudo construir la orientacion del corte."

    pts = _puntos_geometria(elem)
    if len(pts) < 4:
        bb_elem = elem.get_BoundingBox(None)
        if bb_elem is None:
            return None, u"Sin geometria ni bounding box."
        pts = _esquinas_bbox(bb_elem)

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

    m = _mm(MARGEN_MM)

    # Near-clip: valor pequeno fijo. Revit posiciona el HEAD de la anotacion en
    # la esquina (Max.X, Min.Y, Min.Z) del BoundingBox. Con near_clip grande
    # (= old half_prof) el HEAD aparecia muy desplazado del plano de corte en
    # planta. Con near_clip pequeno el HEAD queda casi sobre el extremo del corte.
    near_clip = _mm(100.0)  # 100 mm es suficiente para que Revit genere la seccion

    # Far-clip: cubre la proyeccion HACIA ADELANTE (bz > 0) de la geometria + holgura.
    # Esto asegura que la seccion muestre la fundacion completa al mirar hacia dentro.
    max_z_fwd = max(zs)
    far_clip = max(max_z_fwd + m, _mm(PROFUNDIDAD_MM) * 0.5)

    # Horizontalmente (BasisX): simetrico alrededor de 0 SIN mover el origen en planta.
    xabs = max(abs(min(xs)), abs(max(xs))) + m
    xmn = -xabs
    xmx = xabs

    # Verticalmente (BasisY ~ Z global): recentrar solo en esta direccion.
    ymn_raw = min(ys) - m
    ymx_raw = max(ys) + m
    ymid = 0.5 * (ymn_raw + ymx_raw)
    if abs(ymid) > 1e-9:
        tr.Origin = ox.Add(by.Multiply(ymid))
    yabs = max(abs(ymn_raw - ymid), abs(ymx_raw - ymid))
    ymn = -yabs
    ymx = yabs

    box = BoundingBoxXYZ()
    box.Transform = tr
    box.Min = XYZ(xmn, ymn, -near_clip)
    box.Max = XYZ(xmx, ymx, far_clip)

    try:
        vs = ViewSection.CreateSection(document, vft_id, box)
    except Exception as ex:
        return None, u"CreateSection fallo: {0}".format(ex)

    # Far Clip Offset
    try:
        p_far = vs.get_Parameter(BuiltInParameter.VIEWER_BOUND_OFFSET_FAR)
        if p_far is not None and not p_far.IsReadOnly:
            p_far.Set(_mm(FAR_CLIP_MM))
    except Exception:
        pass

    try:
        _nombre_unico(vs, document, label)
    except Exception:
        pass

    return vs, None


def _ejes_locales_fundacion(elem):
    """
    Devuelve (dir_x_local, dir_y_local) — vectores horizontales normalizados en
    coordenadas mundo que definen los ejes principales del plano de la fundacion.

    Orden de prioridad:
      1. Aristas horizontales del solido (mas robusto: usa geometria real).
      2. GetTotalTransform().BasisX del FamilyInstance.
      3. AABB: eje mas largo en planta.
    """
    # --- Metodo 1: aristas horizontales del solido ---
    try:
        opt = Options()
        try:
            opt.DetailLevel = ViewDetailLevel.Fine
        except Exception:
            pass
        geom = elem.get_Geometry(opt)
        candidatos = {}  # direccion (tuple) -> longitud acumulada

        def _acumular(solido):
            for edge in solido.Edges:
                try:
                    curva = edge.AsCurve()
                    if not isinstance(curva, Line):
                        continue
                    p0 = curva.GetEndPoint(0)
                    p1 = curva.GetEndPoint(1)
                    # Solo aristas horizontales (dZ pequeño)
                    if abs(float(p1.Z - p0.Z)) > 1e-3:
                        continue
                    v = XYZ(float(p1.X - p0.X), float(p1.Y - p0.Y), 0.0)
                    lg = float(v.GetLength())
                    if lg < 1e-6:
                        continue
                    d = v.Normalize()
                    # Canonizar direccion (primer cuadrante o eje +)
                    if float(d.X) < -1e-6 or (abs(float(d.X)) < 1e-6 and float(d.Y) < 0):
                        d = d.Negate()
                    key = (round(float(d.X), 4), round(float(d.Y), 4))
                    candidatos[key] = candidatos.get(key, 0.0) + lg
                except Exception:
                    continue

        for obj in geom:
            if isinstance(obj, Solid) and obj.Volume > 1e-9:
                _acumular(obj)
            elif isinstance(obj, GeometryInstance):
                try:
                    for sub in obj.GetInstanceGeometry():
                        if isinstance(sub, Solid) and sub.Volume > 1e-9:
                            _acumular(sub)
                except Exception:
                    pass

        if candidatos:
            # Ordenar por longitud acumulada (el eje mas largo = dir_x)
            ordenados = sorted(candidatos.items(), key=lambda kv: -kv[1])
            kx, _ = ordenados[0]
            dir_x = XYZ(kx[0], kx[1], 0.0).Normalize()
            # dir_y perpendicular a dir_x en planta
            dir_y = XYZ.BasisZ.CrossProduct(dir_x).Normalize()
            # Si hay un segundo eje distinto, verificar perpendicularity
            if len(ordenados) >= 2:
                k2, _ = ordenados[1]
                d2 = XYZ(k2[0], k2[1], 0.0).Normalize()
                if abs(float(dir_x.DotProduct(d2))) < 0.1:
                    dir_y = d2
            return dir_x, dir_y
    except Exception:
        pass

    # --- Metodo 2: GetTotalTransform del FamilyInstance ---
    if isinstance(elem, FamilyInstance):
        try:
            fi_tr = elem.GetTotalTransform()
            if fi_tr is not None:
                bx = fi_tr.BasisX
                if bx is not None and bx.GetLength() > 1e-9:
                    bx_h = XYZ(float(bx.X), float(bx.Y), 0.0)
                    if bx_h.GetLength() > 1e-6:
                        dir_x = bx_h.Normalize()
                        dir_y = XYZ.BasisZ.CrossProduct(dir_x).Normalize()
                        return dir_x, dir_y
        except Exception:
            pass

    # --- Metodo 3: AABB (eje mas largo en planta) ---
    dir_x = XYZ.BasisX
    dir_y = XYZ.BasisY
    try:
        bb = elem.get_BoundingBox(None)
        if bb is not None:
            sx = abs(float(bb.Max.X) - float(bb.Min.X))
            sy = abs(float(bb.Max.Y) - float(bb.Min.Y))
            if sy > sx:
                dir_x = XYZ.BasisY
                dir_y = XYZ.BasisX
    except Exception:
        pass
    return dir_x, dir_y


# ---------------------------------------------------------------------------
# Filtro de seleccion: solo fundaciones aisladas
# ---------------------------------------------------------------------------

_FOUNDATION_CAT_ID = int(BuiltInCategory.OST_StructuralFoundation)


class _FiltroFundacionAislada(ISelectionFilter):
    """Acepta solo FamilyInstance de categoria Structural Foundation (excluye zapata de muro)."""

    def AllowElement(self, elem):
        try:
            if elem is None:
                return False
            if isinstance(elem, WallFoundation):
                return False
            cat = elem.Category
            if cat is None:
                return False
            return int(cat.Id.IntegerValue) == _FOUNDATION_CAT_ID
        except Exception:
            return False

    def AllowReference(self, ref, point):
        return False


# ---------------------------------------------------------------------------
# Expansion automatica del recuadre de la vista de planta
# ---------------------------------------------------------------------------

def _expandir_vista_planta(doc, vista_planta, vistas_seccion, extra_m):
    """
    Expande el CropBox de la vista de planta activa para que los 4 marcadores
    de seccion queden dentro del recuadre visible. Solo opera si el CropBox
    esta activo y la vista es de tipo planta.
    """
    try:
        if not vista_planta or not vistas_seccion:
            return
        if not isinstance(vista_planta, ViewPlan):
            return
        if not vista_planta.CropBoxActive:
            return

        # Calcular las 4 posiciones de cabecera en coordenadas mundo
        head_world = []
        for vs in vistas_seccion:
            try:
                cb = vs.CropBox
                tr = cb.Transform
                xabs_local = abs(float(cb.Max.X))
                bx_world = tr.BasisX
                o_world = tr.Origin
                head_world.append(o_world + bx_world.Multiply(xabs_local))
                head_world.append(o_world - bx_world.Multiply(xabs_local))
            except Exception:
                continue

        if not head_world:
            return

        # Transformar a coordenadas locales de la vista de planta
        plan_cb = vista_planta.CropBox
        inv_tr = plan_cb.Transform.Inverse
        local_pts = [inv_tr.OfPoint(p) for p in head_world]
        lxs = [float(p.X) for p in local_pts]
        lys = [float(p.Y) for p in local_pts]

        old_min_x = float(plan_cb.Min.X)
        old_max_x = float(plan_cb.Max.X)
        old_min_y = float(plan_cb.Min.Y)
        old_max_y = float(plan_cb.Max.Y)

        fin_min_x = min(old_min_x, min(lxs) - extra_m)
        fin_max_x = max(old_max_x, max(lxs) + extra_m)
        fin_min_y = min(old_min_y, min(lys) - extra_m)
        fin_max_y = max(old_max_y, max(lys) + extra_m)

        eps = 1e-4
        needs = (fin_min_x < old_min_x - eps or fin_max_x > old_max_x + eps or
                 fin_min_y < old_min_y - eps or fin_max_y > old_max_y + eps)
        if not needs:
            return

        tx2 = Transaction(doc, u"BIMTools \u2014 Expandir vista planta")
        tx2.Start()
        try:
            new_cb = BoundingBoxXYZ()
            new_cb.Transform = plan_cb.Transform
            new_cb.Min = XYZ(fin_min_x, fin_min_y, float(plan_cb.Min.Z))
            new_cb.Max = XYZ(fin_max_x, fin_max_y, float(plan_cb.Max.Z))
            vista_planta.CropBox = new_cb
            tx2.Commit()
            print(u"Vista de planta expandida para mostrar los 4 marcadores.")
        except Exception as ex:
            try:
                tx2.RollBack()
            except Exception:
                pass
            print(u"No se pudo expandir la vista de planta: {0}".format(ex))
    except Exception as ex:
        print(u"_expandir_vista_planta error: {0}".format(ex))


# ---------------------------------------------------------------------------
# Punto de entrada
# ---------------------------------------------------------------------------

def main():
    vft_id = _primer_vft_seccion(_doc)
    if vft_id is None:
        TaskDialog.Show(
            u"Error",
            u"No hay ningun tipo de vista de tipo 'Section' en el proyecto.",
        )
        return

    # Seleccion interactiva
    try:
        ref = _uidoc.Selection.PickObject(
            ObjectType.Element,
            _FiltroFundacionAislada(),
            u"Seleccione una fundacion aislada (clic izquierdo, Esc para cancelar)",
        )
    except Exception:
        # El usuario presiono Esc o cancelo
        print("Seleccion cancelada.")
        return

    elem = _doc.GetElement(ref.ElementId)
    if elem is None:
        TaskDialog.Show(u"Error", u"No se pudo obtener el elemento seleccionado.")
        return

    # --- Centro geometrico ---
    # Estrategia prioritaria: AABB de los vertices del SOLIDO en coordenadas mundo.
    # Es mas robusto que AABB del elemento (puede estar inflado) y que LocationPoint
    # (puede estar en una esquina si la familia no tiene el origen en el centro).
    pts_centro = _puntos_geometria(elem)

    if len(pts_centro) >= 4:
        xs_c = [float(p.X) for p in pts_centro]
        ys_c = [float(p.Y) for p in pts_centro]
        zs_c = [float(p.Z) for p in pts_centro]
        cx = 0.5 * (min(xs_c) + max(xs_c))
        cy = 0.5 * (min(ys_c) + max(ys_c))
        cz = 0.5 * (min(zs_c) + max(zs_c))
        centro = XYZ(cx, cy, cz)
    else:
        # Fallback: AABB del elemento
        bb = elem.get_BoundingBox(None)
        if bb is None:
            TaskDialog.Show(u"Error", u"El elemento no tiene geometria ni BoundingBox.")
            return
        cx = 0.5 * (float(bb.Min.X) + float(bb.Max.X))
        cy = 0.5 * (float(bb.Min.Y) + float(bb.Max.Y))
        cz = 0.5 * (float(bb.Min.Z) + float(bb.Max.Z))
        centro = XYZ(cx, cy, cz)

    # Ejes locales del elemento
    dir_x, dir_y = _ejes_locales_fundacion(elem)

    # Etiqueta base segun "Numeracion Fundacion" o ElementId
    num_str = _leer_numeracion_fundacion(elem)
    if not num_str:
        try:
            num_str = str(int(elem.Id.IntegerValue))
        except Exception:
            num_str = u"?"

    label_aa = u"{0} {1} \u2014 A-A".format(NOMBRE_PREFIJO, num_str)
    label_bb = u"{0} {1} \u2014 B-B".format(NOMBRE_PREFIJO, num_str)

    avisos = []
    vistas = []

    # Guardar referencia a la vista de planta activa para expandirla despues
    vista_planta_activa = _uidoc.ActiveView

    tx = Transaction(_doc, u"BIMTools \u2014 Secciones fundacion aislada")
    tx.Start()
    try:
        # Corte A-A: plano local YZ  (dir_corte = eje X local)
        vs_aa, err_aa = _crear_seccion(_doc, vft_id, elem, centro, dir_x, label_aa)
        if vs_aa is not None:
            vistas.append(vs_aa)
        elif err_aa:
            avisos.append(u"A-A: " + err_aa)

        # Corte B-B: plano local XZ  (dir_corte = eje Y local)
        vs_bb, err_bb = _crear_seccion(_doc, vft_id, elem, centro, dir_y, label_bb)
        if vs_bb is not None:
            vistas.append(vs_bb)
        elif err_bb:
            avisos.append(u"B-B: " + err_bb)

        tx.Commit()
    except Exception as ex:
        try:
            tx.RollBack()
        except Exception:
            pass
        TaskDialog.Show(u"Error", u"Error al crear secciones:\n{0}".format(ex))
        return

    # Expandir la vista de planta para que los 4 marcadores sean visibles
    if vistas:
        _expandir_vista_planta(_doc, vista_planta_activa, vistas, _mm(500.0))

    # Resumen
    lineas = []
    for v in vistas:
        try:
            lineas.append(u"\u2022 {0}".format(v.Name))
        except Exception:
            lineas.append(u"\u2022 (vista creada)")
    if avisos:
        lineas.append(u"")
        lineas.append(u"Avisos:")
        for a in avisos:
            lineas.append(u"  " + a)

    if vistas:
        TaskDialog.Show(
            u"Secciones creadas ({0})".format(len(vistas)),
            u"\n".join(lineas),
        )
        try:
            _uidoc.ActiveView = vistas[-1]
        except Exception:
            pass
    else:
        TaskDialog.Show(
            u"Sin vistas",
            u"No se pudo crear ninguna seccion.\n" + u"\n".join(avisos),
        )


main()
