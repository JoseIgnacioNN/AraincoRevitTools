# -*- coding: utf-8 -*-
"""
Intersección geométrica vigas ↔ columnas y muros de **material estructural hormigón**.

- Geometría de vigas seleccionadas: ``obtener_solidos_elemento``.
- Elementos con comportamiento **Concrete** en el proyecto: ``elementos_hormigon_en_proyecto``,
  ``geometria_elementos_hormigon_proyecto``.
- Evaluación frente a **columnas** (y muros) hormigón: filtros de intersección de la API
  más filtro de material; sin bounding box en la lógica principal.

Enfoque alineado con la API de Revit (nivel elemento y nivel ``Solid``):

1. **Filtros de intersección (elemento)** — eficientes para acotar candidatos:
   - ``ElementIntersectsSolidFilter``: elementos que intersectan un ``Solid`` concreto
     (cada sólido obtenido de la viga vía ``get_Geometry``).
   - ``ElementIntersectsElementFilter``: reserva si la viga no aporta sólidos útiles;
     intersección elemento–elemento según el motor geométrico de Revit.

2. **Intersección sólido–sólido** — utilidades ``solidos_intersectan_por_booleana`` /
   ``elementos_intersectan_por_solidos`` para comprobaciones puntuales si otro script las
   reutiliza; **no** se usan para descartar candidatos de los filtros (un contacto cara a
   cara puede tener volumen de intersección nulo).

3. **Otros** (no usados aquí; otros flujos): ``ReferenceIntersector`` (rayos),
   ``JoinGeometryUtils`` (uniones de geometría), ``ExtrusionAnalyzer``, etc.

Restricción de este módulo: **no** se usa ``get_BoundingBox``, ``Outline`` ni
``BoundingBoxIntersectsFilter`` en la lógica de evaluación.

Uso típico: a partir de vigas seleccionadas, obtener columnas y muros **de hormigón**
que colisionan con su volumen físico.
"""

from Autodesk.Revit.DB import (
    BooleanOperationsType,
    BooleanOperationsUtils,
    BuiltInCategory,
    ElementCategoryFilter,
    ElementIntersectsElementFilter,
    ElementIntersectsSolidFilter,
    FilteredElementCollector,
    GeometryInstance,
    Options,
    Solid,
    ViewDetailLevel,
)

_TOL_VOLUMEN_INTERSECCION_FT3 = 1e-6

# Categorías donde suele existir material estructural (hormigón / acero / madera).
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
    # Tipo (forjados, muros): a veces el parámetro de material estructural está en el **tipo**,
    # o como ElementId a un Material cuyo nombre/activo no llegó al instancia.
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
    # Suelo forjado: instancia estructural (``FLOOR_PARAM_IS_STRUCTURAL``; no depender del nombre "Structural" en la UI).
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


def elementos_hormigon_en_proyecto(document):
    """
    Recorre categorías estructurales habituales y devuelve instancias cuyo material
    estructural es **Concrete** (sin duplicar Id).
    """
    if document is None:
        return []
    seen = set()
    out = []
    for bic in _CATS_ESCANEO_MATERIAL_ESTRUCTURAL:
        try:
            elems = (
                FilteredElementCollector(document)
                .OfCategory(bic)
                .WhereElementIsNotElementType()
                .ToElements()
            )
        except Exception:
            continue
        for e in elems or []:
            if e is None or not e.IsValidObject:
                continue
            try:
                eid = e.Id.IntegerValue
            except Exception:
                continue
            if eid in seen:
                continue
            if not material_estructural_es_concrete(e):
                continue
            seen.add(eid)
            out.append(e)
    return out


def geometria_elementos_hormigon_proyecto(document):
    """
    Para cada elemento de hormigón del proyecto, obtiene la lista de ``Solid`` de su
    geometría (misma extracción que para vigas).

    Returns:
        ``list`` de tuplas ``(Element, list[Solid])``.
    """
    pairs = []
    for e in elementos_hormigon_en_proyecto(document):
        pairs.append((e, obtener_solidos_elemento(e)))
    return pairs


def obtener_solidos_elemento(elemento, options=None):
    """
    Extrae todos los ``Solid`` con volumen > 0 de la geometría del elemento.
    Incluye ``GeometryInstance`` aplicando ``GetInstanceGeometry(Transform)``.
    """
    if elemento is None:
        return []
    if options is None:
        options = Options()
        options.ComputeReferences = False
        try:
            options.DetailLevel = ViewDetailLevel.Fine
        except Exception:
            pass
        try:
            options.IncludeNonVisibleObjects = True
        except Exception:
            pass
    try:
        geom_elem = elemento.get_Geometry(options)
    except Exception:
        return []
    if geom_elem is None:
        return []
    solidos = []
    for obj in geom_elem:
        if obj is None:
            continue
        if isinstance(obj, Solid):
            try:
                if obj.Volume > 1e-12:
                    solidos.append(obj)
            except Exception:
                pass
        elif isinstance(obj, GeometryInstance):
            inst_geom = None
            try:
                inst_geom = obj.GetInstanceGeometry(obj.Transform)
            except Exception:
                pass
            if inst_geom is None:
                try:
                    inst_geom = obj.GetInstanceGeometry()
                except Exception:
                    inst_geom = None
            if inst_geom is None:
                continue
            for g in inst_geom:
                if isinstance(g, Solid):
                    try:
                        if g.Volume > 1e-12:
                            solidos.append(g)
                    except Exception:
                        pass
    return solidos


def solidos_intersectan_por_booleana(solid_a, solid_b, tol_volumen=_TOL_VOLUMEN_INTERSECCION_FT3):
    """
    True si el volumen común (intersección) entre dos ``Solid`` supera ``tol_volumen``.

    Implementado con ``BooleanOperationsUtils.ExecuteBooleanOperation`` (intersección
    booleana explícita entre volúmenes BRep).
    """
    if solid_a is None or solid_b is None:
        return False
    try:
        if solid_a.Volume <= 0 or solid_b.Volume <= 0:
            return False
    except Exception:
        return False
    try:
        inter = BooleanOperationsUtils.ExecuteBooleanOperation(
            solid_a, solid_b, BooleanOperationsType.Intersect
        )
        return inter is not None and inter.Volume > tol_volumen
    except Exception:
        return False


def elementos_intersectan_por_solidos(el_a, el_b, tol_volumen=_TOL_VOLUMEN_INTERSECCION_FT3):
    """Comprueba si algún sólido de ``el_a`` intersecta algún sólido de ``el_b`` (booleana)."""
    sa = obtener_solidos_elemento(el_a)
    sb = obtener_solidos_elemento(el_b)
    for a in sa:
        for b in sb:
            if solidos_intersectan_por_booleana(a, b, tol_volumen):
                return True
    return False


_CATS_COLISION = (
    BuiltInCategory.OST_StructuralColumns,
    BuiltInCategory.OST_Walls,
)


def _elemento_es_columna(elem):
    try:
        return elem is not None and elem.Category is not None and int(
            elem.Category.Id.IntegerValue
        ) == int(BuiltInCategory.OST_StructuralColumns)
    except Exception:
        return False


def _elemento_es_muro(elem):
    try:
        return elem is not None and elem.Category is not None and int(
            elem.Category.Id.IntegerValue
        ) == int(BuiltInCategory.OST_Walls)
    except Exception:
        return False


def columnas_hormigon_del_proyecto(document):
    """Lista de columnas estructurales con material hormigón."""
    out = []
    for e in elementos_hormigon_en_proyecto(document):
        if _elemento_es_columna(e):
            out.append(e)
    return out


def _recolectar_por_filtro_solid(document, solid, excluir_ids, resultado_por_id):
    """Añade a ``resultado_por_id`` elementos que pasan ``ElementIntersectsSolidFilter``."""
    if document is None or solid is None:
        return
    for bic in _CATS_COLISION:
        try:
            f_solid = ElementIntersectsSolidFilter(solid)
            f_cat = ElementCategoryFilter(bic)
            col = (
                FilteredElementCollector(document)
                .WherePasses(f_cat)
                .WherePasses(f_solid)
                .ToElements()
            )
        except Exception:
            continue
        for e in col:
            if e is None or not e.IsValidObject:
                continue
            try:
                eid = e.Id.IntegerValue
            except Exception:
                continue
            if eid in excluir_ids:
                continue
            if eid not in resultado_por_id:
                resultado_por_id[eid] = e


def _recolectar_por_filtro_elemento(document, elemento_viga, excluir_ids, resultado_por_id):
    """Reserva si la viga no aporta sólidos: ``ElementIntersectsElementFilter`` (geometría Revit)."""
    if document is None or elemento_viga is None:
        return
    for bic in _CATS_COLISION:
        try:
            f_el = ElementIntersectsElementFilter(elemento_viga)
            f_cat = ElementCategoryFilter(bic)
            col = (
                FilteredElementCollector(document)
                .WherePasses(f_cat)
                .WherePasses(f_el)
                .ToElements()
            )
        except Exception:
            continue
        for e in col:
            if e is None or not e.IsValidObject:
                continue
            try:
                eid = e.Id.IntegerValue
            except Exception:
                continue
            if eid in excluir_ids:
                continue
            if eid not in resultado_por_id:
                resultado_por_id[eid] = e


def columnas_y_muros_que_colisionan_con_vigas(document, elementos_vigas):
    """
    Determina columnas y muros **de material estructural hormigón** cuya geometría
    colisiona o toca alguna de las vigas (geometría de vigas vía ``get_Geometry``;
    candidatos por filtros de intersección de la API).

    No utiliza bounding box en ninguna fase de este módulo.

    Args:
        document: ``Document``.
        elementos_vigas: iterable de ``Element`` (vigas / Structural Framing).

    Returns:
        ``tuple``: (lista columnas hormigón, lista muros hormigón), sin duplicados.
    """
    vigas = [v for v in elementos_vigas if v is not None]
    if not vigas or document is None:
        return ([], [])

    excluir = set()
    for v in vigas:
        try:
            excluir.add(v.Id.IntegerValue)
        except Exception:
            pass

    candidatos_por_id = {}

    for viga in vigas:
        solidos = obtener_solidos_elemento(viga)
        if solidos:
            for s in solidos:
                _recolectar_por_filtro_solid(document, s, excluir, candidatos_por_id)
        # Siempre complementar con intersección elemento–elemento: detecta contacto en apoyos
        # donde el filtro por ``Solid`` puede ser menos completo que el motor global.
        _recolectar_por_filtro_elemento(document, viga, excluir, candidatos_por_id)

    columnas = []
    muros = []
    vistos_col = set()
    vistos_mur = set()

    # Candidatos por filtros de intersección; solo se conservan si el material es hormigón.
    for _eid, elem in candidatos_por_id.items():
        if _elemento_es_columna(elem):
            if not material_estructural_es_concrete(elem):
                continue
            try:
                i = elem.Id.IntegerValue
                if i not in vistos_col:
                    vistos_col.add(i)
                    columnas.append(elem)
            except Exception:
                pass
        elif _elemento_es_muro(elem):
            if not material_estructural_es_concrete(elem):
                continue
            try:
                i = elem.Id.IntegerValue
                if i not in vistos_mur:
                    vistos_mur.add(i)
                    muros.append(elem)
            except Exception:
                pass

    return (columnas, muros)


def evaluar_interseccion_cada_viga_con_columnas_hormigon(document, elementos_vigas):
    """
    Evalúa **cada viga** por separado frente a las columnas de hormigón que intersectan
    solo esa viga (misma lógica que ``columnas_y_muros_que_colisionan_con_vigas`` con un
    único elemento en el grupo).

    Returns:
        ``list`` de ``(Element viga, list columnas_hormigon)``.
    """
    out = []
    for v in elementos_vigas or []:
        if v is None:
            continue
        cols, _mur = columnas_y_muros_que_colisionan_con_vigas(document, [v])
        out.append((v, cols))
    return out
