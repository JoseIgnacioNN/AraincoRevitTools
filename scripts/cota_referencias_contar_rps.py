# -*- coding: utf-8 -*-
"""
RPS / pyRevit: contabilizar referencias de una cota (Dimension) seleccionada.

- Selecciona una cota (Dimension).
- Lee Dimension.References (ReferenceArray).
- Intenta resolver cada Reference a elemento (host o link) y la agrupa para conteo.

Revit 2021+ (según API) | IronPython 2.7/3.x
"""

from __future__ import print_function

import clr
import sys

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import ElementId
from Autodesk.Revit.DB import Dimension
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

DIMENSION_SELECCIONADA = None  # Dimension
REFERENCIAS = None  # list[Reference]
RESUMEN_CONTEOS = None  # dict


class _DimensionSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            return isinstance(elem, Dimension)
        except Exception:
            return False

    def AllowReference(self, reference, position):
        return True


def _safe_int(eid):
    try:
        if eid is None:
            return None
        if isinstance(eid, ElementId):
            return eid.IntegerValue
        return int(eid)
    except Exception:
        return None


def _try_get_stable_rep(doc, ref):
    try:
        if doc is None or ref is None:
            return None
        s = ref.ConvertToStableRepresentation(doc)
        return s
    except Exception:
        return None


def _ref_target_info(doc, ref):
    """
    Devuelve dict con info de destino:
    - host_element_id / linked_element_id / link_instance_id
    - category_name / element_name
    - stable_rep (si disponible)
    """
    info = {
        "host_element_id": None,
        "link_instance_id": None,
        "linked_element_id": None,
        "category_name": None,
        "element_name": None,
        "stable_rep": None,
        "ref_type": None,
    }

    if doc is None or ref is None:
        return info

    # Stable representation ayuda mucho a depurar duplicados / tipos.
    info["stable_rep"] = _try_get_stable_rep(doc, ref)

    # Tipo de referencia (según versión de API).
    for attr in ("ElementReferenceType", "ReferenceType"):
        try:
            rt = getattr(ref, attr, None)
            if rt is not None:
                info["ref_type"] = str(rt)
                break
        except Exception:
            pass

    # Host element id (siempre que aplique)
    try:
        info["host_element_id"] = _safe_int(getattr(ref, "ElementId", None))
    except Exception:
        info["host_element_id"] = None

    # Referencia a link (si aplica)
    try:
        link_eid = getattr(ref, "LinkedElementId", None)
        if link_eid is not None and hasattr(link_eid, "IntegerValue"):
            info["linked_element_id"] = _safe_int(link_eid)
    except Exception:
        pass

    try:
        link_inst = getattr(ref, "ElementId", None)
        # En referencias linkadas, ref.ElementId suele ser la instancia del link;
        # ref.LinkedElementId es el elemento interno del link.
        # Lo dejamos en host_element_id; además lo exponemos como link_instance_id si hay linked_element_id.
        if info["linked_element_id"] is not None:
            info["link_instance_id"] = _safe_int(link_inst)
    except Exception:
        pass

    # Resuelve elemento para nombre/categoría:
    try:
        elem = None
        if info["linked_element_id"] is not None and info["link_instance_id"] is not None:
            # Intento best-effort: resolver la instancia del link para reportar al menos su categoría/nombre.
            elem = doc.GetElement(ElementId(info["link_instance_id"]))
        elif info["host_element_id"] is not None:
            elem = doc.GetElement(ElementId(info["host_element_id"]))

        if elem is not None:
            try:
                info["element_name"] = getattr(elem, "Name", None)
            except Exception:
                info["element_name"] = None
            try:
                cat = getattr(elem, "Category", None)
                info["category_name"] = getattr(cat, "Name", None) if cat else None
            except Exception:
                info["category_name"] = None
    except Exception:
        pass

    return info


def ejecutar(uidoc, doc):
    global DIMENSION_SELECCIONADA, REFERENCIAS, RESUMEN_CONTEOS

    if uidoc is None or doc is None:
        TaskDialog.Show(u"Contar referencias de cota", u"No hay documento activo.")
        return

    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            _DimensionSelectionFilter(),
            u"Selecciona una cota (Dimension).",
        )
    except OperationCanceledException:
        return
    except Exception as ex:
        TaskDialog.Show(u"Contar referencias de cota", u"Selección cancelada o error:\n{}".format(ex))
        return

    if ref is None:
        return

    dim = doc.GetElement(ref.ElementId)
    if not isinstance(dim, Dimension):
        TaskDialog.Show(
            u"Contar referencias de cota",
            u"El elemento seleccionado no es Dimension (tipo: {}).".format(type(dim).__name__),
        )
        return

    DIMENSION_SELECCIONADA = dim

    # References
    refs = []
    try:
        ra = dim.References
        if ra is not None:
            it = ra.GetEnumerator()
            while it.MoveNext():
                refs.append(it.Current)
    except Exception:
        # Algunos casos raros pueden fallar; lo dejamos vacío.
        refs = []

    REFERENCIAS = refs

    # Conteos
    conteo_total = len(refs)
    por_elemento = {}  # key -> count
    por_categoria = {}  # name -> count
    detalles = []  # filas para mostrar

    for i, r in enumerate(refs):
        info = _ref_target_info(doc, r)

        # Clave de agrupación (prioriza link interno si existe)
        if info["linked_element_id"] is not None and info["link_instance_id"] is not None:
            key = u"LINK inst:{} elem:{}".format(info["link_instance_id"], info["linked_element_id"])
        elif info["host_element_id"] is not None:
            key = u"HOST elem:{}".format(info["host_element_id"])
        else:
            key = u"(sin ElementId)"

        por_elemento[key] = por_elemento.get(key, 0) + 1

        cat = info.get("category_name") or u"(sin categoría)"
        por_categoria[cat] = por_categoria.get(cat, 0) + 1

        detalles.append(
            u"{idx:>2}. {key} | cat={cat} | name={name} | type={rt}".format(
                idx=i + 1,
                key=key,
                cat=cat,
                name=info.get("element_name") or u"(sin nombre)",
                rt=info.get("ref_type") or u"(?)",
            )
        )

    RESUMEN_CONTEOS = {
        "total": conteo_total,
        "por_elemento": por_elemento,
        "por_categoria": por_categoria,
    }

    # Construye mensaje
    lines = []
    lines.append(u"Cota seleccionada: id={}  nombre={}".format(_safe_int(dim.Id), getattr(dim, "Name", u"")))
    try:
        vt = getattr(dim, "ValueString", None)
        if vt:
            lines.append(u"Valor (ValueString): {}".format(vt))
    except Exception:
        pass

    lines.append(u"")
    lines.append(u"Total de referencias: {}".format(conteo_total))
    lines.append(u"")
    lines.append(u"Por categoría:")
    for k in sorted(por_categoria.keys(), key=lambda s: s.lower()):
        lines.append(u"  - {}: {}".format(k, por_categoria[k]))

    lines.append(u"")
    lines.append(u"Por elemento (agrupado):")
    for k in sorted(por_elemento.keys(), key=lambda s: s.lower()):
        lines.append(u"  - {}: {}".format(k, por_elemento[k]))

    if detalles:
        lines.append(u"")
        lines.append(u"Detalle (1..n):")
        lines.extend(detalles[:60])
        if len(detalles) > 60:
            lines.append(u"... ({} más)".format(len(detalles) - 60))

    lines.append(u"")
    lines.append(u"Variables del módulo: cota_referencias_contar_rps.DIMENSION_SELECCIONADA / REFERENCIAS / RESUMEN_CONTEOS")

    TaskDialog.Show(u"Contar referencias de cota", u"\n".join(lines))

    # Consola RPS
    print(u"Total referencias:", conteo_total)
    print(u"Por categoría:", por_categoria)
    print(u"Por elemento:", por_elemento)


def _main():
    try:
        doc = __revit__.ActiveUIDocument.Document
        uidoc = __revit__.ActiveUIDocument
    except NameError:
        TaskDialog.Show(
            u"Contar referencias de cota",
            u"Define __revit__ (pyRevit/RPS) o llama ejecutar(uidoc, doc).",
        )
        return
    ejecutar(uidoc, doc)


_main()

