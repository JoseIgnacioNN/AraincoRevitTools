"""
Script para Revit Python Shell (RPS) / IronPython 3.4
Crea un nuevo tipo de gancho de armadura (RebarHookType) en el documento activo.

Ejecutar: File > Run Script... y seleccionar este archivo .py
O en la consola: exec(open(r'RUTA_COMPLETA\crear_hook_type.py', encoding='utf-8').read())
El Hook Length en el nombre se obtiene del parámetro Hook Lengths de RebarBarType (Structural Rebar).
"""

import math
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, UnitUtils, UnitTypeId
from Autodesk.Revit.DB.Structure import RebarHookType, RebarBarType


def obtener_rebar_bar_types(documento):
    """Obtiene todos los RebarBarType del documento."""
    collector = FilteredElementCollector(documento)
    return list(collector.OfClass(RebarBarType))


def obtener_primer_rebar_bar_type(documento):
    """Obtiene el primer RebarBarType del documento para calcular Hook Length."""
    bar_types = obtener_rebar_bar_types(documento)
    return bar_types[0] if bar_types else None


def obtener_hook_length_desde_bar_type(bar_type, hook_type, documento):
    """
    Obtiene el Hook Length desde el parámetro Hook Lengths del RebarBarType.
    Si el gancho es nuevo y no está en la tabla, usa GetDefaultHookExtension (diámetro × multiplicador).
    """
    try:
        # Intentar obtener desde la tabla Hook Lengths del RebarBarType
        largo_interno = bar_type.GetHookLength(hook_type.Id)
    except Exception:
        # Gancho recién creado: puede no estar en la tabla. Usar cálculo por defecto.
        diametro = bar_type.BarModelDiameter
        largo_interno = hook_type.GetDefaultHookExtension(diametro)
    # Convertir de unidades internas (pies) a mm
    return UnitUtils.ConvertFromInternalUnits(largo_interno, UnitTypeId.Millimeters)


def obtener_nombres_hook_existentes(documento):
    """
    Obtiene los nombres de todos los RebarHookType existentes en el documento.
    Basado en la lógica de get_rebar_hook_types.py
    """
    collector = FilteredElementCollector(documento)
    hook_types = list(collector.OfClass(RebarHookType))
    return [ht.Name for ht in hook_types]


def crear_hook_type(nombre, angulo_grados, multiplicador_extension, largo_hook_mm=None, rebar_bar_type=None):
    """
    Crea un nuevo RebarHookType en el documento activo.
    El Hook Length se establece con el valor pasado (ej: calculado en función del espesor de un elemento).

    Args:
        nombre: Nombre del tipo de gancho (string)
        angulo_grados: Ángulo del gancho en grados (float, rango 0-180)
        multiplicador_extension: Multiplicador para la longitud del segmento recto (float, rango 0-99)
        largo_hook_mm: Hook Length en mm (variable calculada externamente; si None, usa 50 mm)
        rebar_bar_type: RebarBarType para el nombre (opcional, usa el primero del doc si None)

    Returns:
        RebarHookType creado o None si falla
    """
    # doc es la variable predefinida del documento activo en Revit Python Shell
    if largo_hook_mm is None:
        largo_hook_mm = 50.0
    bar_type = rebar_bar_type or obtener_primer_rebar_bar_type(doc)
    if not bar_type:
        raise Exception("No hay RebarBarType en el documento. Crea al menos un tipo de barra de armadura.")

    nombres_existentes = obtener_nombres_hook_existentes(doc)
    angulo_str = str(int(angulo_grados)) if angulo_grados == int(angulo_grados) else str(angulo_grados)

    t = Transaction(doc, "Crear tipo de gancho de armadura")
    t.Start()

    try:
        # Conversión de grados a radianes (la API requiere radianes, rango 0-PI)
        angulo_radianes = math.radians(angulo_grados)

        # Crear el RebarHookType (angle en radianes, multiplier para extensión)
        hook_type = RebarHookType.Create(doc, angulo_radianes, multiplicador_extension)

        # Desactivar auto-cálculo y establecer Hook Length con el valor calculado (variable)
        largo_interno = UnitUtils.ConvertToInternalUnits(largo_hook_mm, UnitTypeId.Millimeters)
        for bt in obtener_rebar_bar_types(doc):
            bt.SetAutoCalcHookLengths(hook_type.Id, False)
            bt.SetHookLength(hook_type.Id, largo_interno)

        largo_str = "{:.1f} mm".format(largo_hook_mm)

        # Formato: "Nombre - 90º - Largo" (ej: "Rebar Hook - 90º - 200.0 mm")
        nombre_base = "{} - {}º - {}".format(nombre, angulo_str, largo_str)
        nombre_final = nombre_base
        if nombre_base in nombres_existentes:
            contador = 1
            while nombre_final in nombres_existentes:
                nombre_final = "{} ({})".format(nombre_base, contador)
                contador += 1
            print("Nombre '{}' ya existe. Usando '{}' en su lugar.".format(nombre_base, nombre_final))

        # Asignar el nombre personalizado (propiedad Name de ElementType, write-only)
        hook_type.Name = nombre_final

        t.Commit()
        print("Tipo de gancho creado exitosamente: '{}' (ángulo: {}°, Hook Length: {} mm)".format(
            nombre_final, angulo_grados, largo_hook_mm))
        return hook_type

    except Exception as ex:
        t.RollBack()
        print("Error al crear el tipo de gancho: {}".format(str(ex)))
        raise


# Ejemplo de uso
if __name__ == "__main__" or "doc" in dir():
    # Sin largo_hook_mm: usa 50 mm por defecto
    crear_hook_type(
        nombre="Rebar Hook",
        angulo_grados=90.0,
        multiplicador_extension=12.0
    )
    # Con largo_hook_mm: variable calculada (ej: según espesor del elemento)
    # crear_hook_type("Rebar Hook", 90.0, 12.0, largo_hook_mm=70.0)
