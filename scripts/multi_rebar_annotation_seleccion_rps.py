# -*- coding: utf-8 -*-
"""
Aplica Multi-Rebar Annotation a un set de barras (Rebar) seleccionadas en la vista activa.

Uso en RPS
----------
1. Selecciona una o varias barras (Rebar / RebarInSystem) en la vista activa.
2. Ejecuta este script.
3. Si el proyecto tiene más de un tipo MRA, el script muestra la lista disponible en consola
   y usa el que se indique en MRA_TYPE_NAME (o el primero si ese campo queda vacío).

Configuración opcional
----------------------
MRA_TYPE_NAME  : nombre exacto del tipo Multi-Rebar Annotation del proyecto.
                 Déjalo en "" para usar el primero disponible, o escribe el nombre
                 del tipo que quieras aplicar (insensible a mayúsculas).
OFFSET_MM      : distancia en mm entre el array de barras y la línea de cota MRA
                 (medida en la dirección perpendicular a la distribución).
                 None → se calcula automáticamente a partir del bounding-box del rebar.
TAG_HAS_LEADER : True/False → la etiqueta lleva o no línea de llamada.
"""

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    DimensionStyleType,
    ElementCategoryFilter,
    ElementId,
    FilteredElementCollector,
    IndependentTag,
    MultiReferenceAnnotation,
    MultiReferenceAnnotationOptions,
    MultiReferenceAnnotationType,
    StorageType,
    Transaction,
    View3D,
    XYZ,
)
from Autodesk.Revit.DB.Structure import Rebar, RebarInSystem
from Autodesk.Revit.UI import TaskDialog
from System.Collections.Generic import List

# ---------------------------------------------------------------------------
# CONFIGURACIÓN
# ---------------------------------------------------------------------------
MRA_TYPE_NAME = u""   # "" → primer tipo disponible; o ej. u"Recorrido Barras"
OFFSET_MM = None      # None → automático; o p.ej. 300 para 300 mm
TAG_HAS_LEADER = False
# +1 → offset en la dirección positiva de la perpendicular a la distribución.
# -1 → dirección negativa (prueba este valor si la MRA aparece dentro del elemento).
OFFSET_SIDE = 1
# ---------------------------------------------------------------------------


def _mm_to_ft(mm):
    return float(mm) / 304.8


def _norm(s):
    """Normaliza texto: unicode, sin espacios dobles, sin NBSP."""
    if s is None:
        return u""
    try:
        t = unicode(s)
    except Exception:
        t = str(s)
    return t.replace(u"\u00A0", u" ").strip()


def _type_names(t):
    """Todas las cadenas candidatas de nombre para un MultiReferenceAnnotationType."""
    out = []
    try:
        n = getattr(t, "Name", None)
        if n:
            out.append(_norm(n))
    except Exception:
        pass
    for bip in (BuiltInParameter.SYMBOL_NAME_PARAM, BuiltInParameter.ALL_MODEL_TYPE_NAME):
        try:
            p = t.get_Parameter(bip)
            if p is not None and p.HasValue and p.StorageType == StorageType.String:
                v = _norm(p.AsString())
                if v:
                    out.append(v)
        except Exception:
            pass
    seen = set()
    return [x for x in out if x and not (x in seen or seen.add(x))]


def _get_all_mra_types(doc):
    """Lista de todos los MultiReferenceAnnotationType del proyecto."""
    try:
        return list(FilteredElementCollector(doc).OfClass(MultiReferenceAnnotationType))
    except Exception:
        return []


def _find_mra_type(doc, type_name):
    """
    Busca un MultiReferenceAnnotationType por nombre.
    Si type_name es vacío devuelve el primero disponible.
    """
    all_types = _get_all_mra_types(doc)
    if not all_types:
        return None
    if not _norm(type_name):
        return all_types[0]
    target = _norm(type_name)
    target_lower = target.lower()
    exact = []
    ci = []
    for t in all_types:
        for cand in _type_names(t):
            if cand == target:
                exact.append(t)
                break
            if cand.lower() == target_lower:
                ci.append(t)
                break
    if exact:
        return exact[0]
    if ci:
        return ci[0]
    return None


def _vista_valida(view):
    if view is None:
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass
    try:
        if isinstance(view, View3D):
            return False
    except Exception:
        pass
    return True


def _centro_rebar(rebar, view):
    """Centro del bounding-box del rebar en la vista (o global)."""
    try:
        bb = rebar.get_BoundingBox(view)
        if bb is not None:
            return (bb.Min + bb.Max) * 0.5
    except Exception:
        pass
    try:
        bb = rebar.get_BoundingBox(None)
        if bb is not None:
            return (bb.Min + bb.Max) * 0.5
    except Exception:
        pass
    return None


def _spacing_direction(rebar, fallback_rd, view_vd, view_up):
    """
    Calcula la dirección de distribución (spacing) del set de barras.
    Devuelve (XYZ normalizado, str_metodo).

    Método A  : vector posición-0 → posición-n-1 del array (más fiable).
    Método B1 : bar_dir × vd  (funciona si la barra NO es paralela a vd).
    Método B2 : cuando bar_dir ≈ vd (barra longitudinal vista en corte/elevación),
                la distribución visible en la vista es v_up (vertical en pantalla).
    Método C  : fallback_rd (último recurso).
    """
    try:
        from Autodesk.Revit.DB.Structure import MultiplanarOption
        n = 0
        try:
            n = int(rebar.NumberOfBarPositions)
        except Exception:
            pass

        # Método A: posición 0 → posición n-1
        if n > 1:
            for mpo_name in ("IncludeAllMultiplanarCurves", "IncludeOnlyPlanarCurves"):
                mpo = getattr(MultiplanarOption, mpo_name, None)
                if mpo is None:
                    continue
                try:
                    cs0 = list(rebar.GetCenterlineCurves(False, False, False, mpo, 0))
                    csn = list(rebar.GetCenterlineCurves(False, False, False, mpo, n - 1))
                    if cs0 and csn:
                        c0 = cs0[0].Evaluate(0.5, True)
                        cn = csn[0].Evaluate(0.5, True)
                        v = cn - c0
                        if float(v.GetLength()) > 1e-6:
                            return v.Normalize(), u"A:pos0→posN (n={0}, mpo={1})".format(n, mpo_name)
                except Exception:
                    pass

        # Obtener la curva más larga de la barra para determinar bar_dir
        curves = []
        for mpo_name in ("IncludeAllMultiplanarCurves", "IncludeOnlyPlanarCurves"):
            mpo = getattr(MultiplanarOption, mpo_name, None)
            if mpo is None:
                continue
            try:
                cs = list(rebar.GetCenterlineCurves(False, False, False, mpo, 0))
                if cs:
                    curves = cs
                    break
            except Exception:
                pass
        if not curves:
            try:
                curves = list(rebar.GetCenterlineCurves(False, False, False))
            except Exception:
                pass

        if curves:
            longest = max(curves, key=lambda c: float(c.Length))
            p0 = longest.Evaluate(0.0, True)
            p1 = longest.Evaluate(1.0, True)
            bar_dir = p1 - p0
            if bar_dir.GetLength() > 1e-9:
                bar_dir = bar_dir.Normalize()

                # Método B2: barra ≈ paralela a vd (barra longitudinal "entrando en pantalla").
                # En vistas de corte/alzado la distribución visible es vertical (v_up).
                dot_vd = abs(float(bar_dir.DotProduct(view_vd)))
                if dot_vd > 0.8:
                    return view_up, u"B2:barParalelaAVd→v_up (dot={0:.3f})".format(dot_vd)

                # Método B1: bar_dir × vd
                spacing = bar_dir.CrossProduct(view_vd)
                if spacing.GetLength() > 1e-9:
                    return spacing.Normalize(), u"B1:barDir×vd (barDir=({0:.2f},{1:.2f},{2:.2f}))".format(
                        float(bar_dir.X), float(bar_dir.Y), float(bar_dir.Z))

        return fallback_rd, u"C:fallback_rd"
    except Exception as ex:
        return fallback_rd, u"C:fallback_rd (ex:{0})".format(ex)


def _offset_automatico(rebar, view, offset_dir):
    """
    Offset en pies desde el centro del rebar hasta la línea MRA,
    medido en la dirección `offset_dir` (perpendicular a la distribución).
    Se usa la mitad de la proyección del bbox del rebar en esa dirección,
    más un margen fijo de 500 mm (suficiente para quedar fuera del elemento host).
    """
    margen_ft = _mm_to_ft(500.0)
    try:
        bb = rebar.get_BoundingBox(view)
        if bb is None:
            bb = rebar.get_BoundingBox(None)
        if bb is not None:
            dim = abs(float((bb.Max - bb.Min).DotProduct(offset_dir)))
            return dim * 0.5 + margen_ft
    except Exception:
        pass
    return margen_ft


def _project_onto_view_plane(v, vd):
    """
    Proyecta el vector v sobre el plano de vista (elimina la componente en vd).
    Devuelve el vector normalizado, o None si queda degenerado (< 1e-9).
    """
    dot = float(v.DotProduct(vd))
    px = float(v.X) - dot * float(vd.X)
    py = float(v.Y) - dot * float(vd.Y)
    pz = float(v.Z) - dot * float(vd.Z)
    proj = XYZ(px, py, pz)
    if proj.GetLength() < 1e-9:
        return None
    return proj.Normalize()


def _fmt(v):
    """Formatea un XYZ para impresión."""
    try:
        return u"({0:.3f}, {1:.3f}, {2:.3f})".format(float(v.X), float(v.Y), float(v.Z))
    except Exception:
        return str(v)


def _crear_mra(doc, view, rebar, mrat_type, offset_ft, tag_has_leader, avisos, offset_side=1):
    """
    Intenta crear una MultiReferenceAnnotation para `rebar` en `view`.
    Devuelve True si tuvo éxito.
    """
    rid = rebar.Id.IntegerValue
    p_mid = _centro_rebar(rebar, view)
    if p_mid is None:
        avisos.append(u"  Id {0}: no se pudo obtener centro del rebar.".format(rid))
        return False
    try:
        vd  = view.ViewDirection.Normalize()
        rd  = view.RightDirection.Normalize()
        v_up = view.UpDirection.Normalize()
    except Exception as ex:
        avisos.append(u"  Id {0}: error obteniendo vectores de vista: {1}".format(rid, ex))
        return False

    # --- Diagnóstico: vectores de vista ---
    print(u"  Id {0}: vd={1}  rd={2}  v_up={3}".format(rid, _fmt(vd), _fmt(rd), _fmt(v_up)))
    print(u"  Id {0}: p_mid={1}".format(rid, _fmt(p_mid)))

    # 1. Dirección de distribución 3D → proyectada al plano de vista.
    #    DimensionLineDirection DEBE estar en el plano de la vista.
    spacing_dir_3d, metodo = _spacing_direction(rebar, rd, vd, v_up)
    print(u"  Id {0}: spacing_dir_3d={1}  método={2}".format(rid, _fmt(spacing_dir_3d), metodo))

    spacing_dir = _project_onto_view_plane(spacing_dir_3d, vd)
    if spacing_dir is None:
        # La distribución es ≈ perpendicular al plano de vista. Usar rd.
        spacing_dir = rd
        print(u"  Id {0}: spacing_dir (proyectado) → degenerado → fallback rd={1}".format(rid, _fmt(rd)))
    else:
        print(u"  Id {0}: spacing_dir (proyectado)={1}".format(rid, _fmt(spacing_dir)))

    # 2. Dirección de offset = perpendicular a spacing_dir en el plano de vista.
    #    Queremos desplazar la línea MRA "fuera" del array de barras.
    perp_dir = spacing_dir.CrossProduct(vd)
    if perp_dir.GetLength() < 1e-9:
        perp_dir = v_up
        print(u"  Id {0}: perp_dir degenerado → usando v_up={1}".format(rid, _fmt(v_up)))
    else:
        perp_dir = perp_dir.Normalize()

    print(u"  Id {0}: perp_dir={1}  (OFFSET_SIDE={2})".format(rid, _fmt(perp_dir), offset_side))

    try:
        opts = MultiReferenceAnnotationOptions(mrat_type)
    except Exception:
        try:
            opts = MultiReferenceAnnotationOptions()
            opts.MultiReferenceAnnotationType = mrat_type.Id
        except Exception as ex:
            avisos.append(u"  Id {0}: error creando options: {1}".format(rid, ex))
            return False

    try:
        opts.DimensionStyleType = DimensionStyleType.Linear
    except Exception:
        pass

    off = offset_ft if offset_ft is not None else _offset_automatico(rebar, view, perp_dir)
    side = 1 if offset_side >= 0 else -1
    p_line = p_mid + perp_dir.Multiply(float(off) * side)
    print(u"  Id {0}: off={1:.4f} ft ({2:.0f} mm)  p_line={3}".format(
        rid, off, off * 304.8, _fmt(p_line)))

    try:
        opts.DimensionPlaneNormal  = vd
        opts.DimensionLineDirection = spacing_dir
        opts.DimensionLineOrigin   = p_line
        opts.TagHeadPosition       = p_line
        opts.TagHasLeader          = bool(tag_has_leader)
    except Exception as ex:
        avisos.append(u"  Id {0}: error configurando options: {1}".format(rid, ex))
        return False

    ids = List[ElementId]()
    ids.Add(rebar.Id)
    try:
        opts.SetElementsToDimension(ids)
    except Exception as ex:
        avisos.append(u"  Id {0}: SetElementsToDimension falló: {1}".format(rebar.Id.IntegerValue, ex))
        return False

    try:
        if hasattr(opts, "ElementsMatchReferenceCategory"):
            if not opts.ElementsMatchReferenceCategory(doc):
                avisos.append(u"  Id {0}: elementos no válidos para el tipo MRA.".format(rebar.Id.IntegerValue))
                return False
    except Exception:
        pass

    try:
        mra = MultiReferenceAnnotation.Create(doc, view.Id, opts)
        if mra is None:
            avisos.append(u"  Id {0}: MultiReferenceAnnotation.Create retornó None.".format(rebar.Id.IntegerValue))
            return False
        # Forzar sin línea de llamada en la etiqueta dependiente si se pide.
        if not tag_has_leader:
            try:
                flt = ElementCategoryFilter(BuiltInCategory.OST_RebarTags)
                for did in mra.GetDependentElements(flt):
                    el = doc.GetElement(did)
                    if isinstance(el, IndependentTag) and el.HasLeader:
                        el.HasLeader = False
            except Exception:
                pass
        return True
    except Exception as ex:
        avisos.append(u"  Id {0}: Create falló: {1}".format(rebar.Id.IntegerValue, ex))
        return False


def run(revit_app):
    doc = revit_app.ActiveUIDocument.Document
    uidoc = revit_app.ActiveUIDocument
    view = uidoc.ActiveView

    # --- Validar vista --------------------------------------------------
    if not _vista_valida(view):
        TaskDialog.Show(
            u"Multi-Rebar Annotation",
            u"La vista activa no es válida para Multi-Rebar Annotations.\n"
            u"Use una vista de planta, alzado o sección (no plantilla, no 3D).",
        )
        return

    # --- Validar tipo MRA -----------------------------------------------
    all_types = _get_all_mra_types(doc)
    if not all_types:
        TaskDialog.Show(
            u"Multi-Rebar Annotation",
            u"No se encontró ningún tipo Multi-Rebar Annotation en el proyecto.\n"
            u"Cargue o cree al menos un tipo antes de ejecutar este script.",
        )
        return

    print(u"\n=== Multi-Rebar Annotation – tipos disponibles en el proyecto ===")
    for t in all_types:
        names = _type_names(t)
        print(u"  • {0}  (Id {1})".format(names[0] if names else u"<sin nombre>", t.Id.IntegerValue))

    mrat_type = _find_mra_type(doc, MRA_TYPE_NAME)
    if mrat_type is None:
        TaskDialog.Show(
            u"Multi-Rebar Annotation",
            u"No se encontró el tipo «{0}».\nRevisa el valor de MRA_TYPE_NAME en el script.".format(MRA_TYPE_NAME),
        )
        return

    used_name = (_type_names(mrat_type) or [u"<sin nombre>"])[0]
    print(u"\nTipo MRA a usar: «{0}»  (Id {1})".format(used_name, mrat_type.Id.IntegerValue))

    # --- Obtener selección de rebars ------------------------------------
    sel_ids = list(uidoc.Selection.GetElementIds())
    if not sel_ids:
        TaskDialog.Show(
            u"Multi-Rebar Annotation",
            u"No hay elementos seleccionados.\nSelecciona una o más barras (Rebar) y vuelve a ejecutar.",
        )
        return

    rebars = []
    ignorados = []
    for eid in sel_ids:
        el = doc.GetElement(eid)
        if isinstance(el, Rebar):
            rebars.append(el)
        elif isinstance(el, RebarInSystem):
            # RebarInSystem no puede recibir MRA directamente; informar al usuario.
            ignorados.append(eid.IntegerValue)
        else:
            ignorados.append(eid.IntegerValue)

    if not rebars:
        TaskDialog.Show(
            u"Multi-Rebar Annotation",
            u"Ninguno de los elementos seleccionados es una barra individual (Rebar).\n"
            u"Multi-Rebar Annotation requiere barras de armadura individuales, "
            u"no RebarInSystem ni otros elementos.",
        )
        return

    print(u"\nBarras a anotar : {0}".format(len(rebars)))
    if ignorados:
        print(u"Ignorados (no Rebar): {0}".format(ignorados))

    # --- Offset en pies -------------------------------------------------
    offset_ft = _mm_to_ft(OFFSET_MM) if OFFSET_MM is not None else None

    # --- Transacción ----------------------------------------------------
    avisos = []
    n_ok = 0
    t = Transaction(doc, u"BIMTools – Multi-Rebar Annotation")
    try:
        t.Start()
        for rb in rebars:
            if _crear_mra(doc, view, rb, mrat_type, offset_ft, TAG_HAS_LEADER, avisos, OFFSET_SIDE):
                n_ok += 1
        t.Commit()
    except Exception as ex:
        try:
            t.RollbackToCheckpoint()
        except Exception:
            pass
        try:
            t.RollbackToCheckpoint()
        except Exception:
            pass
        try:
            if t.HasStarted():
                t.RollBack()
        except Exception:
            pass
        TaskDialog.Show(u"Multi-Rebar Annotation", u"Error en la transacción:\n{0}".format(ex))
        return

    # --- Resumen --------------------------------------------------------
    print(u"\nResultado: {0}/{1} anotaciones creadas correctamente.".format(n_ok, len(rebars)))
    if avisos:
        print(u"\nAvisos:")
        for a in avisos:
            print(a)

    msg = u"Anotaciones creadas: {0} de {1} barras.".format(n_ok, len(rebars))
    if avisos:
        msg += u"\n\nAvisos:\n" + u"\n".join(avisos[:10])
        if len(avisos) > 10:
            msg += u"\n… (ver consola para más detalles)"
    TaskDialog.Show(u"Multi-Rebar Annotation", msg)


# Punto de entrada RPS
run(__revit__)
