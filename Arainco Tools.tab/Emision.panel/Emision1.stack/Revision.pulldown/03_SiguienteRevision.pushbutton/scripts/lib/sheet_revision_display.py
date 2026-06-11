# -*- coding: utf-8 -*-
"""
Texto de revisión actual de una lámina para el DataGrid de Revisiones.
"""

from __future__ import print_function

try:
    unicode
except NameError:
    unicode = str

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import BuiltInParameter, ElementId


def _param_as_string(elem, bip):
    try:
        p = elem.get_Parameter(bip)
        if p is None:
            return u""
        return (p.AsString() or p.AsValueString() or u"").strip()
    except Exception:
        return u""


def _param_by_bip_name(elem, bip_name):
    try:
        name = bip_name
        try:
            if isinstance(bip_name, unicode):
                name = str(bip_name)
        except Exception:
            pass
        bip = getattr(BuiltInParameter, name, None)
        if bip is None:
            return u""
        return _param_as_string(elem, bip)
    except Exception:
        return u""


def sheet_revision_display(sheet, doc):
    """
    Texto de revisión para la grilla: GetCurrentRevision primero, luego parámetro de lámina.
    """
    try:
        rid = sheet.GetCurrentRevision()
        if rid is not None:
            try:
                if rid == ElementId.InvalidElementId:
                    rid = None
            except Exception:
                pass
        if rid is not None:
            rev = doc.GetElement(rid)
            if rev is not None:
                for attr in (u"RevisionNumber", u"Description", u"SequenceNumber"):
                    try:
                        val = getattr(rev, attr, None)
                        if val is None:
                            continue
                        s = unicode(val).strip()
                        if s:
                            return s
                    except Exception:
                        pass
    except Exception:
        pass
    t = _param_by_bip_name(sheet, u"SHEET_CURRENT_REVISION")
    if t:
        return t
    try:
        p = sheet.LookupParameter(u"Current Revision")
        if p is not None:
            t = (p.AsString() or p.AsValueString() or u"").strip()
            if t:
                return t
    except Exception:
        pass
    return u""
