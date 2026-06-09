# -*- coding: utf-8 -*-
"""
Detección de material estructural hormigón (Concrete) — subconjunto portable.

Extraído de ``geometria_colision_vigas`` para el paquete autocontenido de
``11_UnirGeometriaHormigonVista.pushbutton`` (sin depender de scripts/ de la extensión).
"""

from Autodesk.Revit.DB import BuiltInCategory

_CATS_ESCANEO_MATERIAL_ESTRUCTURAL = (
    BuiltInCategory.OST_StructuralColumns,
    BuiltInCategory.OST_StructuralFraming,
    BuiltInCategory.OST_Walls,
    BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_StructuralFoundation,
)


def material_estructural_es_concrete(elem):
    """
    True si el elemento se considera de **material estructural hormigón** (Concrete).

    Prioriza ``StructuralMaterialType.Concrete`` (vigas, pilares, zapatas, etc.).
    Respaldo: parámetro de material estructural por texto (idiomas / nombre de material).
    """
    if elem is None:
        return False
    try:
        from Autodesk.Revit.DB.Structure import StructuralMaterialType

        sm = elem.StructuralMaterialType
        return sm == StructuralMaterialType.Concrete
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import BuiltInParameter

        p = elem.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)
        if p is not None and p.HasValue:
            try:
                vs = p.AsValueString()
                if vs and _texto_material_indica_hormigon(vs):
                    return True
            except Exception:
                pass
            try:
                s = p.AsString()
                if s and _texto_material_indica_hormigon(s):
                    return True
            except Exception:
                pass
    except Exception:
        pass
    for key in (u"Structural Material", u"Material estructural"):
        try:
            p = elem.LookupParameter(key)
            if p and p.HasValue:
                try:
                    vs = p.AsValueString()
                except Exception:
                    vs = None
                if vs and _texto_material_indica_hormigon(vs):
                    return True
        except Exception:
            pass
    try:
        from Autodesk.Revit.DB import BuiltInParameter, ElementId, StorageType

        doc0 = elem.Document
        if doc0 is not None and elem is not None:
            tid = elem.GetTypeId()
            if tid is not None and tid != ElementId.InvalidElementId:
                et = doc0.GetElement(tid)
                if et is not None:
                    p2 = et.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)
                    if p2 is not None and p2.HasValue:
                        if p2.StorageType == StorageType.ElementId:
                            mid = p2.AsElementId()
                            if mid is not None and mid != ElementId.InvalidElementId:
                                m = doc0.GetElement(mid)
                                if m is not None and _mat_o_texto_sugiere_hormigon(
                                    m, p2, doc0
                                ):
                                    return True
                        for attr in (u"AsString", u"AsValueString"):
                            try:
                                t = getattr(p2, attr)()
                                if t and _texto_material_indica_hormigon(t):
                                    return True
                            except Exception:
                                pass
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import BuiltInParameter, ElementId, Floor, FloorType

        doc0 = elem.Document
        _bip_fs = getattr(BuiltInParameter, u"FLOOR_PARAM_IS_STRUCTURAL", None)
        if _bip_fs is not None and doc0 is not None and isinstance(elem, Floor):
            p_st = elem.get_Parameter(_bip_fs)
            if p_st is not None and p_st.HasValue:
                try:
                    if p_st.AsInteger() == 1:
                        return True
                except Exception:
                    pass
        if doc0 is not None and elem is not None:
            tid = elem.GetTypeId()
            if tid is not None and tid != ElementId.InvalidElementId:
                et = doc0.GetElement(tid)
                if isinstance(et, FloorType):
                    if _bip_fs is not None:
                        p_t = et.get_Parameter(_bip_fs)
                        if p_t is not None and p_t.HasValue:
                            try:
                                if p_t.AsInteger() == 1:
                                    return True
                            except Exception:
                                pass
                    if _texto_material_indica_hormigon(et.Name or u""):
                        return True
    except Exception:
        pass
    return False


def _mat_o_texto_sugiere_hormigon(material, param, _document):
    """Comprueba parámetro de material y ``Material`` del documento (nombre, etc.)."""
    if param is not None and param.HasValue:
        for attr in (u"AsString", u"AsValueString"):
            try:
                t = getattr(param, attr)()
                if t and _texto_material_indica_hormigon(t):
                    return True
            except Exception:
                pass
    if material is None:
        return False
    try:
        n = material.Name
    except Exception:
        n = None
    if n and _texto_material_indica_hormigon(n):
        return True
    try:
        from Autodesk.Revit.DB.Structure import StructuralMaterialType

        if hasattr(material, "StructuralMaterialType"):
            if material.StructuralMaterialType == StructuralMaterialType.Concrete:
                return True
    except Exception:
        pass
    return False


def _texto_material_indica_hormigon(texto):
    if not texto:
        return False
    try:
        t = unicode(texto).lower()
    except Exception:
        t = str(texto).lower()
    for s in (
        u"concrete",
        u"hormigón",
        u"hormigon",
        u"beton",
        u"betão",
    ):
        if s in t:
            return True
    return False
