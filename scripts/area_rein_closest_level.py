# -*- coding: utf-8 -*-
"""
Escribir el nombre del nivel del host en el parámetro "Closest Level" de Area Reinforcement.

Revit 2024+ | pyRevit / RPS (importable).
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from System.Collections.Generic import List
from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    Floor,
    StorageType,
    Transaction,
    UnitUtils,
    UnitTypeId,
    Wall,
)
from Autodesk.Revit.DB.Structure import AreaReinforcement
from Autodesk.Revit.UI import TaskDialog

CLOSEST_LEVEL_PARAM_NAME = u"Closest Level"
TRANSACTION_NAME = u"Arainco: Closest Level en Area Reinforcement"
DIALOG_TITLE = u"Arainco: Closest Level Area Reinforcement"


def _element_id_int(eid):
    if eid is None:
        return None
    try:
        return int(eid.Value)
    except Exception:
        try:
            return int(eid.IntegerValue)
        except Exception:
            return None


def _category_label(element):
    if element is None:
        return u"(sin categoría)"
    try:
        cat = element.Category
        if cat is not None and cat.Name:
            return cat.Name
    except Exception:
        pass
    return u"(sin categoría)"


def _type_name(document, element):
    if element is None:
        return u""
    try:
        type_id = element.GetTypeId()
        if type_id is None or type_id == ElementId.InvalidElementId:
            return u""
        sym = document.GetElement(type_id)
        if sym is None:
            return u""
        try:
            fam = sym.FamilyName
            typ = sym.Name
            if fam and typ:
                return u"{} : {}".format(fam, typ)
            return typ or fam or u""
        except Exception:
            return sym.Name or u""
    except Exception:
        return u""


def _host_thickness_mm(host):
    if host is None:
        return None
    try:
        if isinstance(host, Floor):
            p = host.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)
        elif isinstance(host, Wall):
            p = host.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)
        else:
            p = None
        if p is not None and p.HasValue:
            return UnitUtils.ConvertFromInternalUnits(p.AsDouble(), UnitTypeId.Millimeters)
    except Exception:
        pass
    try:
        p = host.LookupParameter("Default Thickness")
        if p is not None and p.HasValue:
            return UnitUtils.ConvertFromInternalUnits(p.AsDouble(), UnitTypeId.Millimeters)
    except Exception:
        pass
    return None


def _host_level_id(host):
    if host is None:
        return None
    try:
        lid = host.LevelId
        if lid is not None and lid != ElementId.InvalidElementId:
            try:
                if int(lid.IntegerValue) < 0:
                    pass
                else:
                    return lid
            except Exception:
                return lid
    except Exception:
        pass
    try:
        if isinstance(host, Wall):
            p = host.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)
            if p is not None and p.HasValue and p.StorageType == StorageType.ElementId:
                eid = p.AsElementId()
                if eid is not None and eid != ElementId.InvalidElementId:
                    return eid
    except Exception:
        pass
    for bip_name in (
        u"INSTANCE_REFERENCE_LEVEL_PARAM",
        u"LEVEL_PARAM",
        u"SCHEDULE_LEVEL_PARAM",
    ):
        try:
            bip = getattr(BuiltInParameter, bip_name, None)
            if bip is None:
                continue
            p = host.get_Parameter(bip)
            if p is None or not p.HasValue or p.StorageType != StorageType.ElementId:
                continue
            eid = p.AsElementId()
            if eid is not None and eid != ElementId.InvalidElementId:
                return eid
        except Exception:
            pass
    return None


def _host_level_name(document, host):
    lid = _host_level_id(host)
    if lid is None:
        return None, None
    try:
        level = document.GetElement(lid)
        if level is None:
            return _element_id_int(lid), None
        name = level.Name
        if name is None:
            return _element_id_int(lid), None
        return _element_id_int(lid), name.ToString()
    except Exception:
        return _element_id_int(lid), None


def _area_reinforcement_host_id(area_rein):
    try:
        hid = area_rein.GetHostId()
        if hid is not None and hid != ElementId.InvalidElementId:
            return hid
    except Exception:
        pass
    return None


def _set_closest_level_parameter(area_rein, level_name):
    if not level_name:
        return False, u"sin nombre de nivel"
    try:
        param = area_rein.LookupParameter(CLOSEST_LEVEL_PARAM_NAME)
    except Exception:
        param = None
    if param is None:
        return False, u'parámetro "{}" no encontrado'.format(CLOSEST_LEVEL_PARAM_NAME)
    if param.IsReadOnly:
        return False, u'parámetro "{}" es solo lectura'.format(CLOSEST_LEVEL_PARAM_NAME)
    if param.StorageType != StorageType.String:
        return False, u'parámetro "{}" no es texto (StorageType={})'.format(
            CLOSEST_LEVEL_PARAM_NAME,
            param.StorageType,
        )
    try:
        param.Set(level_name)
        return True, None
    except Exception as ex:
        return False, str(ex)


def describe_host(document, host):
    if host is None:
        return {
            u"host_id": None,
            u"host_class": None,
            u"host_category": None,
            u"host_type": None,
            u"host_name": None,
            u"host_thickness_mm": None,
            u"host_level_id": None,
            u"host_level_name": None,
        }
    level_id, level_name = _host_level_name(document, host)
    return {
        u"host_id": _element_id_int(host.Id),
        u"host_class": type(host).__name__,
        u"host_category": _category_label(host),
        u"host_type": _type_name(document, host),
        u"host_name": getattr(host, "Name", None) or u"",
        u"host_thickness_mm": _host_thickness_mm(host),
        u"host_level_id": level_id,
        u"host_level_name": level_name,
    }


def describe_area_reinforcement(document, area_rein):
    host_id = _area_reinforcement_host_id(area_rein)
    host = document.GetElement(host_id) if host_id is not None else None
    info = {
        u"area_rein_id": _element_id_int(area_rein.Id),
        u"area_rein_type": _type_name(document, area_rein),
        u"host_id": _element_id_int(host_id),
        u"host": host,
        u"area_rein": area_rein,
        u"closest_level_written": None,
        u"closest_level_error": None,
    }
    info.update(describe_host(document, host))
    return info


def format_result_line(info):
    ar_id = info.get(u"area_rein_id")
    if info.get(u"host_id") is None:
        return u"AreaReinforcement Id {} → sin host (GetHostId inválido).".format(ar_id)

    parts = [
        u"AreaReinforcement Id {} → Host Id {}".format(ar_id, info.get(u"host_id")),
        u"{} [{}]".format(info.get(u"host_category"), info.get(u"host_class")),
    ]
    if info.get(u"host_type"):
        parts.append(info.get(u"host_type"))
    if info.get(u"host_name"):
        parts.append(u'Nombre: "{}"'.format(info.get(u"host_name")))
    th = info.get(u"host_thickness_mm")
    if th is not None:
        parts.append(u"Espesor: {:.1f} mm".format(float(th)))
    level_name = info.get(u"host_level_name")
    if level_name:
        parts.append(u'Nivel: "{}"'.format(level_name))
    elif info.get(u"host_level_id") is not None:
        parts.append(u"Nivel Id {} (sin nombre)".format(info.get(u"host_level_id")))
    written = info.get(u"closest_level_written")
    if written:
        parts.append(u'Closest Level ← "{}"'.format(written))
    err = info.get(u"closest_level_error")
    if err:
        parts.append(u"Closest Level: {}".format(err))
    return u" | ".join(parts)


def collect_all_area_reinforcements(document):
    return list(FilteredElementCollector(document).OfClass(AreaReinforcement))


def set_element_selection(uidoc, element_ids):
    ids = List[ElementId]()
    for eid in element_ids:
        if eid is not None and eid != ElementId.InvalidElementId:
            ids.Add(eid)
    uidoc.Selection.SetElementIds(ids)


def apply_closest_level_to_targets(document, targets, write_closest_level=True):
    results = []
    missing_host = 0
    missing_level = 0

    for ar in targets:
        info = describe_area_reinforcement(document, ar)
        results.append(info)
        if info.get(u"host_id") is None:
            missing_host += 1
        elif not info.get(u"host_level_name"):
            missing_level += 1

    updated = 0
    write_failed = 0
    if not write_closest_level:
        return results, updated, write_failed, missing_host, missing_level

    trans = Transaction(document, TRANSACTION_NAME)
    trans.Start()
    try:
        for info in results:
            ar = info.get(u"area_rein")
            level_name = info.get(u"host_level_name")
            if ar is None:
                info[u"closest_level_error"] = u"Area Reinforcement no disponible"
                write_failed += 1
                continue
            if not level_name:
                info[u"closest_level_error"] = u"host sin nivel"
                write_failed += 1
                continue
            ok, err = _set_closest_level_parameter(ar, level_name)
            if ok:
                info[u"closest_level_written"] = level_name
                updated += 1
            else:
                info[u"closest_level_error"] = err
                write_failed += 1
        trans.Commit()
    except Exception:
        trans.RollBack()
        raise

    return results, updated, write_failed, missing_host, missing_level


def build_summary(results, updated, write_failed, missing_host, missing_level):
    summary = u"{} Area Reinforcement analizado(s).".format(len(results))
    if missing_host:
        summary += u" {} sin host.".format(missing_host)
    if missing_level:
        summary += u" {} sin nivel en el host.".format(missing_level)
    summary += u"\nClosest Level escrito: {}.".format(updated)
    if write_failed:
        summary += u" Fallos: {}.".format(write_failed)
    return summary


def run(uiapp, select_area_rein_after=True, show_dialog=True):
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        msg = u"No hay documento activo."
        print(u"Error:", msg)
        if show_dialog:
            TaskDialog.Show(DIALOG_TITLE, msg)
        return []

    document = uidoc.Document
    targets = collect_all_area_reinforcements(document)
    if not targets:
        msg = u"No hay Area Reinforcement en el proyecto."
        print(u"Error:", msg)
        if show_dialog:
            TaskDialog.Show(DIALOG_TITLE, msg)
        return []

    results, updated, write_failed, missing_host, missing_level = apply_closest_level_to_targets(
        document,
        targets,
        write_closest_level=True,
    )

    if select_area_rein_after:
        set_element_selection(uidoc, [ar.Id for ar in targets])

    lines = [format_result_line(info) for info in results]
    summary = build_summary(results, updated, write_failed, missing_host, missing_level)
    if select_area_rein_after:
        summary += u"\nSelección actualizada a todos los Area Reinforcement del proyecto."

    print(u"\n=== Closest Level — Area Reinforcement ===")
    for line in lines:
        print(line)
    print(summary)

    if show_dialog:
        detail = u"\n".join(lines)
        if len(detail) > 1800:
            detail = detail[:1800] + u"\n… (ver consola para el listado completo)"
        try:
            TaskDialog.Show(DIALOG_TITLE, summary + u"\n\n" + detail)
        except Exception:
            pass

    return results
