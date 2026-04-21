# -*- coding: utf-8 -*-
"""
RPS / Revit Python Shell: cara seleccionada → sistema de coordenadas paramétrico y normal.

Tras una selección interactiva de una cara (ObjectType.Face), el script:
  - Resuelve la geometría Face del elemento host.
  - Toma un punto interior de la cara (centro del recinto UV del BoundingBox).
  - Evalúa ComputeDerivatives(uv): triedro local en espacio modelo
    (Origen, tangente U = BasisX, tangente V = BasisY, normal = BasisZ).

Variables globales de módulo tras éxito (para inspección en la consola RPS):
  CARA_REFERENCIA      — Reference
  CARA_HOST            — Element
  CARA_GEOM            — Face
  CARA_UV_CENTRO       — UV usado para la evaluación
  CARA_TRANSFORM_UV    — Transform devuelto por ComputeDerivatives
  CARA_NORMAL_MODELO   — XYZ normal unitaria (espacio modelo)
  CARA_ORIGEN_MODELO   — XYZ origen del triedro en el punto UV

Revit 2024+ | IronPython 2.7 / 3.x (según tu RPS).
"""

from __future__ import print_function

import clr
import math
import sys

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import Face, Options, PlanarFace, UV, ViewDetailLevel, XYZ
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType

CARA_REFERENCIA = None
CARA_HOST = None
CARA_GEOM = None
CARA_UV_CENTRO = None
CARA_TRANSFORM_UV = None
CARA_NORMAL_MODELO = None
CARA_ORIGEN_MODELO = None


def _xyz_fmt(p, dec=6):
    if p is None:
        return "(null)"
    return "({0:.{3}f}, {1:.{3}f}, {2:.{3}f})".format(p.X, p.Y, p.Z, dec)


def _norm_vec(v):
    if v is None:
        return None
    try:
        ln = v.GetLength()
    except Exception:
        try:
            ln = math.sqrt(v.X * v.X + v.Y * v.Y + v.Z * v.Z)
        except Exception:
            return None
    if ln < 1e-12:
        return None
    try:
        return v.Divide(ln)
    except Exception:
        return XYZ(v.X / ln, v.Y / ln, v.Z / ln)


def _uv_centro_en_cara(face):
    """Punto UV interior: centro del BoundingBox UV de la cara."""
    if face is None:
        return None
    bb = face.GetBoundingBox()
    if bb is None:
        return UV(0.0, 0.0)
    u = 0.5 * (bb.Min.U + bb.Max.U)
    v = 0.5 * (bb.Min.V + bb.Max.V)
    return UV(u, v)


def _compute_derivatives(face, uv):
    """Invoca Face.ComputeDerivatives con la firma disponible en tu versión de Revit."""
    if face is None or uv is None:
        return None
    try:
        return face.ComputeDerivatives(uv, True)
    except TypeError:
        pass
    except Exception:
        pass
    try:
        return face.ComputeDerivatives(uv)
    except Exception:
        return None


def _asignar_modulo(ref, host, geom, uv, tr, normal, origen):
    m = sys.modules[__name__]
    m.CARA_REFERENCIA = ref
    m.CARA_HOST = host
    m.CARA_GEOM = geom
    m.CARA_UV_CENTRO = uv
    m.CARA_TRANSFORM_UV = tr
    m.CARA_NORMAL_MODELO = normal
    m.CARA_ORIGEN_MODELO = origen


def ejecutar(uidoc, doc):
    if uidoc is None or doc is None:
        TaskDialog.Show(u"Cara — ejes y normal", u"No hay documento activo.")
        return

    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Face,
            u"Selecciona una cara del modelo.",
        )
    except OperationCanceledException:
        return
    except Exception as ex:
        TaskDialog.Show(u"Cara — ejes y normal", u"Selección cancelada o error:\n{}".format(ex))
        return

    if ref is None:
        return

    host = doc.GetElement(ref.ElementId)
    if host is None:
        TaskDialog.Show(u"Cara — ejes y normal", u"No se pudo obtener el elemento host de la referencia.")
        return

    opts = Options()
    opts.ComputeReferences = True
    opts.DetailLevel = ViewDetailLevel.Fine

    try:
        geom = host.GetGeometryObjectFromReference(ref)
    except Exception as ex:
        TaskDialog.Show(
            u"Cara — ejes y normal",
            u"No se pudo resolver la geometría desde la referencia:\n{}".format(ex),
        )
        return

    if not isinstance(geom, Face):
        TaskDialog.Show(
            u"Cara — ejes y normal",
            u"La geometría obtenida no es Face (tipo: {}).".format(type(geom).__name__),
        )
        return

    face = geom
    uv = _uv_centro_en_cara(face)
    tr = _compute_derivatives(face, uv)

    if tr is None:
        TaskDialog.Show(u"Cara — ejes y normal", u"ComputeDerivatives devolvió None para esta cara/UV.")
        _asignar_modulo(ref, host, face, uv, None, None, None)
        return

    origen = tr.Origin
    t_u = tr.BasisX
    t_v = tr.BasisY
    n_tr = tr.BasisZ

    n_unit = _norm_vec(n_tr)
    if n_unit is None:
        n_unit = n_tr

    # PlanarFace: la API expone también FaceNormal (debe alinearse con BasisZ salvo orientación).
    nota_plana = u""
    if isinstance(face, PlanarFace):
        fn = getattr(face, "FaceNormal", None)
        if fn is not None:
            fn_u = _norm_vec(fn)
            if fn_u is not None:
                dot = 0.0
                try:
                    dot = abs(fn_u.DotProduct(n_unit))
                except Exception:
                    pass
                nota_plana = u"\nPlanarFace.FaceNormal (unitaria): {0}\n".format(_xyz_fmt(fn_u))
                nota_plana += u"  |dot(FaceNormal, BasisZ de deriv)| ≈ {0:.6f}\n".format(dot)

    _asignar_modulo(ref, host, face, uv, tr, n_unit, origen)

    # Construcción explícita (IronPython no admite concatenación implícita
    # entre una string literal y el resultado de un .format()).
    msg = (
        u"Sistema en el punto UV (centro del recinto de la cara)\n"
        + u"  U = {0:.8f},  V = {1:.8f}\n\n".format(uv.U, uv.V)
        + u"Origen (modelo): {0}\n".format(_xyz_fmt(origen))
        + u"Tangente ∂r/∂u (BasisX): {0}  len≈{1:.6f}\n".format(
            _xyz_fmt(t_u), t_u.GetLength() if t_u else 0.0
        )
        + u"Tangente ∂r/∂v (BasisY): {0}  len≈{1:.6f}\n".format(
            _xyz_fmt(t_v), t_v.GetLength() if t_v else 0.0
        )
        + u"Normal (BasisZ del Transform, unitaria si el API normaliza): {0}\n".format(_xyz_fmt(n_unit))
        + nota_plana
        + u"\nVariables en el módulo: cara_sistema_coordenadas_rps.CARA_*"
    )
    TaskDialog.Show(u"Cara — ejes UV y normal", msg)

    # Consola RPS
    print(u"CARA_NORMAL_MODELO:", _xyz_fmt(CARA_NORMAL_MODELO))
    print(u"CARA_ORIGEN_MODELO:", _xyz_fmt(CARA_ORIGEN_MODELO))


def _main():
    try:
        doc = __revit__.ActiveUIDocument.Document
        uidoc = __revit__.ActiveUIDocument
    except NameError:
        TaskDialog.Show(
            u"Cara — ejes y normal",
            u"Define __revit__ (pyRevit/RPS) o llama ejecutar(uidoc, doc).",
        )
        return
    ejecutar(uidoc, doc)


_main()
