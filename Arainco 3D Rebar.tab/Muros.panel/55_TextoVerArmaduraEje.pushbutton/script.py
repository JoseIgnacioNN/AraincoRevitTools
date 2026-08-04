# -*- coding: utf-8 -*-
"""
Texto Ver Armadura Eje — TextNote vertical en alzado o Building Section.

Revit 2024+ | pyRevit | IronPython 2.7 / 3.4

Flujo:
  1. Validar que la vista activa sea un alzado (Elevation) o
     una sección Building Section (ViewType.Section).
  2. Seleccionar un Grid (eje) visible en la vista.
  3. Indicar punto de inserción.
  4. Crear TextNote vertical: «VER ARMADURA EN EJE {nombre}».
  5. Repetir hasta cancelar (Esc).
"""

from __future__ import print_function

import math

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    Grid,
    HorizontalTextAlignment,
    Plane,
    SketchPlane,
    TextNote,
    TextNoteOptions,
    TextNoteType,
    Transaction,
    VerticalTextAlignment,
    ViewFamily,
    ViewType,
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

__title__ = u"Texto Ver\nArmadura Eje"
__author__ = u"BIMTools"
__doc__ = (
    u"En un alzado o Building Section, selecciona un eje (Grid) y un punto. "
    u"Genera texto vertical «VER ARMADURA EN EJE [nombre del eje]»."
)

TOOL_TITLE = u"Arainco: Texto Ver Armadura Eje"
TX_NAME = u"Arainco: Texto Ver Armadura Eje"
TEXT_TEMPLATE = u"VER ARMADURA EN EJE {0}"
TEXT_NOTE_TYPE_NAME = u"3.0mm Arial"
TEXT_NOTE_ROTATION_RAD = math.pi * 0.5


def _as_unicode(value):
    if value is None:
        return u""
    try:
        return unicode(value)
    except NameError:
        return str(value)


def _mostrar_aviso(uiapp, instruction, content=u""):
    """Diálogo WPF BIMTools; respaldo TaskDialog."""
    try:
        from bimtools_instruction_dialog import show_message_dialog
        from revit_wpf_window_position import revit_main_hwnd

        hwnd = revit_main_hwnd(uiapp) if uiapp is not None else None
        show_message_dialog(
            TOOL_TITLE,
            instruction=instruction,
            content=content,
            ok_text=u"Entendido",
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass
    try:
        msg = instruction
        if content:
            msg = instruction + u"\n\n" + content
        TaskDialog.Show(TOOL_TITLE, msg)
    except Exception:
        print(instruction)


class _GridFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, Grid)

    def AllowReference(self, reference, position):
        return True


def _enum_equals(valor, enum_obj):
    if valor is None or enum_obj is None:
        return False
    try:
        if valor == enum_obj:
            return True
    except Exception:
        pass
    try:
        a = _as_unicode(valor.ToString() if hasattr(valor, u"ToString") else valor)
        b = _as_unicode(
            enum_obj.ToString() if hasattr(enum_obj, u"ToString") else enum_obj
        )
        return a.lower() == b.lower()
    except Exception:
        return False


def _es_alzado(view):
    try:
        return _enum_equals(view.ViewType, ViewType.Elevation)
    except Exception:
        return False


def _nombre_tipo_vista(view):
    """Nombre del ViewFamilyType (p. ej. Building Section)."""
    if view is None:
        return u""
    try:
        tid = view.GetTypeId()
        if tid is not None and tid != ElementId.InvalidElementId:
            vft = view.Document.GetElement(tid)
            if vft is not None:
                nm = _as_unicode(getattr(vft, u"Name", u"") or u"").strip()
                if nm:
                    return nm
    except Exception:
        pass
    try:
        p = view.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)
        if p is not None:
            raw = _as_unicode(p.AsValueString() or p.AsString() or u"").strip()
            if u":" in raw:
                raw = raw.split(u":", 1)[1].strip()
            if raw:
                return raw
    except Exception:
        pass
    return u""


def _es_building_section(view):
    """True si la vista es sección de edificio (Building Section)."""
    if view is None:
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass
    try:
        if _enum_equals(view.ViewType, ViewType.Detail):
            return False
    except Exception:
        pass
    try:
        if not _enum_equals(view.ViewType, ViewType.Section):
            return False
    except Exception:
        return False

    vft = None
    try:
        tid = view.GetTypeId()
        if tid is not None and tid != ElementId.InvalidElementId:
            vft = view.Document.GetElement(tid)
    except Exception:
        vft = None

    nombre = u""
    if vft is not None:
        try:
            if not _enum_equals(vft.ViewFamily, ViewFamily.Section):
                return False
        except Exception:
            pass
        try:
            nombre = _as_unicode(vft.Name or u"")
        except Exception:
            nombre = u""
    if not nombre:
        nombre = _nombre_tipo_vista(view)

    key = nombre.lower().replace(u" ", u"")
    if key:
        if u"detail" in key or u"detalle" in key:
            return False
        if u"buildingsection" in key or u"secciondeedificio" in key:
            return True
        if u"building" in key and u"section" in key:
            return True

    # ViewType.Section + ViewFamily.Section sin marcadores de detalle
    if vft is not None:
        try:
            if _enum_equals(vft.ViewFamily, ViewFamily.Section):
                return True
        except Exception:
            pass
    return True


def _vista_soportada(view):
    """Alzado (Elevation) o sección Building Section."""
    return _es_alzado(view) or _es_building_section(view)


def _nombre_eje(grid):
    try:
        name = grid.Name
        if name and _as_unicode(name).strip():
            return _as_unicode(name).strip()
    except Exception:
        pass
    return u"?"


def _texto_armadura(grid):
    return TEXT_TEMPLATE.format(_nombre_eje(grid))


def _text_note_type_name(text_type):
    """Nombre visible del TextNoteType (``Name`` suele estar vacío)."""
    if text_type is None:
        return u""
    for bip in (
        BuiltInParameter.SYMBOL_NAME_PARAM,
        BuiltInParameter.ALL_MODEL_TYPE_NAME,
    ):
        try:
            p = text_type.get_Parameter(bip)
            if p is None or not p.HasValue:
                continue
            s = _as_unicode(p.AsString() or u"").strip()
            if s:
                return s
            s = _as_unicode(p.AsValueString() or u"").strip()
            if s:
                return s
        except Exception:
            pass
    try:
        from Autodesk.Revit.DB import Element

        s = _as_unicode(Element.Name.__get__(text_type) or u"").strip()
        if s:
            return s
    except Exception:
        pass
    try:
        return _as_unicode(text_type.Name or u"").strip()
    except Exception:
        return u""


def _canon_text_type_key(name):
    """Normaliza nombre de tipo para comparar (espacios, coma/punto)."""
    s = _as_unicode(name or u"").strip().lower()
    if not s:
        return u""
    s = s.replace(u",", u".")
    parts = s.split()
    return u" ".join(parts)


def _text_note_type(doc):
    """Busca el TextNoteType «3.0mm Arial» en el proyecto."""
    target = _canon_text_type_key(TEXT_NOTE_TYPE_NAME)
    if not target:
        return None
    fallback = None
    try:
        for tnt in FilteredElementCollector(doc).OfClass(TextNoteType):
            if tnt is None:
                continue
            name = _text_note_type_name(tnt)
            key = _canon_text_type_key(name)
            if not key:
                continue
            if key == target:
                return tnt
            # Coincidencia flexible: 3.0 / 3,0 + Arial
            if fallback is None and u"3.0" in key and u"arial" in key:
                fallback = tnt
    except Exception:
        pass
    return fallback


def _aplicar_opciones_texto(opts):
    try:
        opts.Rotation = float(TEXT_NOTE_ROTATION_RAD)
    except Exception:
        pass
    try:
        opts.HorizontalAlignment = HorizontalTextAlignment.Center
    except Exception:
        pass
    try:
        opts.VerticalAlignment = VerticalTextAlignment.Middle
    except Exception:
        pass
    try:
        opts.KeepRotatedTextReadable = False
    except Exception:
        pass


def _crear_texto_vertical(doc, view, origin, texto, text_type):
    tn = None
    try:
        opts = TextNoteOptions(text_type.Id)
        _aplicar_opciones_texto(opts)
        tn = TextNote.Create(doc, view.Id, origin, texto, opts)
    except Exception:
        tn = None
    if tn is None:
        opts = TextNoteOptions()
        opts.TypeId = text_type.Id
        _aplicar_opciones_texto(opts)
        tn = TextNote.Create(doc, view.Id, origin, texto, opts)
    try:
        tn.KeepRotatedTextReadable = False
    except Exception:
        pass
    return tn


def _ensure_view_sketch_plane(doc, view):
    """Activa el plano de la vista (ViewDirection + Origin) para PickPoint.

    Toda Building Section / Elevation tiene plano de corte; si no hay
    SketchPlane activo, se crea a partir de ese plano.
    """
    try:
        if view.SketchPlane is not None:
            return True
    except Exception:
        pass
    t = Transaction(doc, u"Arainco: Plano de trabajo vista")
    t.Start()
    try:
        plane = Plane.CreateByNormalAndOrigin(view.ViewDirection, view.Origin)
        sketch_plane = SketchPlane.Create(doc, plane)
        view.SketchPlane = sketch_plane
        t.Commit()
        return True
    except Exception:
        try:
            t.RollBack()
        except Exception:
            pass
        return False


def _pick_grid(uidoc):
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            _GridFilter(),
            u"Selecciona el eje (Grid) en el alzado o Building Section. "
            u"Esc para terminar.",
        )
    except OperationCanceledException:
        return None
    except Exception:
        return None
    if ref is None:
        return None
    elem = uidoc.Document.GetElement(ref.ElementId)
    if isinstance(elem, Grid):
        return elem
    return None


def _pick_punto(uidoc, nombre_eje, uiapp):
    try:
        return uidoc.Selection.PickPoint(
            u"Indica el punto de inserción para «VER ARMADURA EN EJE {0}». "
            u"Esc para cancelar este eje.".format(nombre_eje)
        )
    except OperationCanceledException:
        return None
    except Exception as ex:
        _mostrar_aviso(
            uiapp,
            u"No se pudo indicar el punto de inserción.",
            content=_as_unicode(ex),
        )
        return None


def main():
    from pyrevit import revit

    doc = revit.doc
    uidoc = revit.uidoc
    uiapp = __revit__
    vista = uidoc.ActiveView

    if not _vista_soportada(vista):
        tipo = _nombre_tipo_vista(vista)
        content = u"Vista activa: {0}".format(
            _as_unicode(getattr(vista, "Name", u""))
        )
        if tipo:
            content += u"\nTipo de vista: {0}".format(tipo)
        _mostrar_aviso(
            uiapp,
            u"Abre un alzado (Elevation) o una sección Building Section "
            u"antes de usar esta herramienta.",
            content=content,
        )
        return

    text_type = _text_note_type(doc)
    if text_type is None:
        disponibles = []
        try:
            for tnt in FilteredElementCollector(doc).OfClass(TextNoteType):
                nm = _text_note_type_name(tnt)
                if nm and nm not in disponibles:
                    disponibles.append(nm)
        except Exception:
            pass
        content = u"Créalo o renómbralo en Anotar > Texto (tipos de TextNote)."
        if disponibles:
            sample = u", ".join(disponibles[:8])
            content += u"\n\nTipos encontrados: {0}".format(sample)
        _mostrar_aviso(
            uiapp,
            u"No se encontró el tipo de texto «{0}».".format(TEXT_NOTE_TYPE_NAME),
            content=content,
        )
        return

    if not _ensure_view_sketch_plane(doc, vista):
        _mostrar_aviso(
            uiapp,
            u"No se pudo activar el plano de trabajo de la vista.",
            content=u"Vista: {0}".format(_as_unicode(getattr(vista, "Name", u""))),
        )
        return

    while True:
        grid = _pick_grid(uidoc)
        if grid is None:
            break

        texto = _texto_armadura(grid)
        punto = _pick_punto(uidoc, _nombre_eje(grid), uiapp)
        if punto is None:
            continue

        t = Transaction(doc, TX_NAME)
        t.Start()
        try:
            tn = _crear_texto_vertical(doc, vista, punto, texto, text_type)
            if tn is None:
                t.RollBack()
                _mostrar_aviso(
                    uiapp,
                    u"No se pudo crear el TextNote.",
                    content=texto,
                )
                continue
            t.Commit()
        except Exception as ex:
            try:
                t.RollBack()
            except Exception:
                pass
            _mostrar_aviso(
                uiapp,
                u"Error al crear el texto.",
                content=_as_unicode(ex),
            )
            break


# --- Validación acceso corporativo (RECURSOS COMPARTIDOS) ---
import os as _os_ac
import sys as _sys_ac

_tab_ac = _os_ac.path.dirname(_os_ac.path.abspath(__file__))
for _iac in range(16):
    if _os_ac.path.basename(_tab_ac).endswith(u".tab"):
        break
    _parent_ac = _os_ac.path.dirname(_tab_ac)
    if _parent_ac == _tab_ac:
        _tab_ac = None
        break
    _tab_ac = _parent_ac
if _tab_ac and _tab_ac not in _sys_ac.path:
    _sys_ac.path.insert(0, _tab_ac)
import bimtools_access_bootstrap as _bimtools_access

if _bimtools_access.require_tool_access(__file__, __revit__, __title__):
    main()
