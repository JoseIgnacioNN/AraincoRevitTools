# -*- coding: utf-8 -*-
"""
RevisionService — lógica de negocio para aplicar revisiones a láminas Revit.

Encapsula toda la lógica de:
- Detección de revisión siguiente / revisión 0.
- Escritura de parámetros Rnn en cajetín (dos convenciones de nombre).
- Actualización del índice de revisiones de la lámina (SetAdditionalRevisionIds).
- Manejo de nubes de revisión (CANTIDAD_REVISIONES).
- Progreso y bloqueo de comandos durante la emisión.
- Transacción agrupada con rollback ante primer error.
"""

from __future__ import print_function

try:
    unicode
except NameError:
    unicode = str

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from System.Collections.Generic import List as ClrList
from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    Revision,
    Transaction,
    ViewSheet,
)

try:
    from Autodesk.Revit.DB import RevisionCloud
except Exception:
    RevisionCloud = None

from siguiente_revision.constants import (
    TX_NAME,
    SUFFIX_NUM,
    SUFFIX_DES,
    SUFFIX_DIR,
    SUFFIX_DIB,
    SUFFIX_REV,
    SUFFIX_APR,
    SUFFIX_FCH,
    MAX_REVISION_SLOTS,
    PARAM_CANTIDAD_REVISIONES,
    LAYOUT_RNN_FIELD,
    LAYOUT_R01_ROW,
)
from siguiente_revision.services import parameter_service as ps
from siguiente_revision.services.sheet_service import sheet_display


# ---------------------------------------------------------------------------
# Helpers de ElementId
# ---------------------------------------------------------------------------

def _eid_int(eid):
    try:
        return int(eid.Value)
    except Exception:
        try:
            return int(eid.IntegerValue)
        except Exception:
            return 0


def _eid_from_int(val):
    from System import Int64
    try:
        return ElementId(Int64(int(val)))
    except Exception:
        try:
            return ElementId(int(val))
        except Exception:
            return ElementId.InvalidElementId


# ---------------------------------------------------------------------------
# Revisiones del proyecto
# ---------------------------------------------------------------------------

def get_ordered_revision_ids(doc):
    """
    Lista de ElementId de revisiones en el orden de Gestión de revisiones.
    """
    try:
        raw = Revision.GetAllRevisionIds(doc)
    except Exception:
        raw = None
    out = []
    if raw:
        for eid in raw:
            if eid is not None and eid != ElementId.InvalidElementId:
                out.append(eid)
    return out


def get_revisions_on_sheet(doc, sheet):
    """Revisiones de la lámina como lista de objetos Revision."""
    revs = []
    for eid in sheet.GetAdditionalRevisionIds():
        el = doc.GetElement(eid)
        if isinstance(el, Revision):
            revs.append(el)
    return revs


def index_in_project_order(project_ordered_ids, revision_element):
    """Índice (0-based) de una revisión en la lista del proyecto; -1 si no aparece."""
    t = _eid_int(revision_element.Id)
    for i, eid in enumerate(project_ordered_ids):
        if _eid_int(eid) == t:
            return i
    return -1


def furthest_sheet_revision(existing_revs, project_ordered_ids):
    """
    Revisión de la lámina con mayor índice en la secuencia del proyecto.

    Returns (revision_element, index) o (None, -1).
    """
    best = None
    best_i = -1
    for r in existing_revs:
        j = index_in_project_order(project_ordered_ids, r)
        if j < 0:
            continue
        if j > best_i:
            best_i = j
            best = r
    return best, best_i


def revision_number_display(rev_el):
    """Número o letra de la revisión (ej. 0, 1, A) desde API de Revit."""
    try:
        s = unicode(getattr(rev_el, u"RevisionNumber", None) or u"").strip()
        if s:
            return s
    except Exception:
        pass
    for nm in ("REVIT_REVISION_NUMBER", "REVISION_NUMBER"):
        bip = getattr(BuiltInParameter, nm, None)
        if bip is None:
            continue
        try:
            p = rev_el.get_Parameter(bip)
        except Exception:
            p = None
        if p is None:
            continue
        try:
            s = (p.AsString() or p.AsValueString() or u"").strip()
        except Exception:
            s = u""
        if s:
            return s
    return u""


def index_of_revision_display_number(doc, ordered_ids, want_display):
    """
    Índice (0-based) de la primera revisión cuyo número mostrado coincide con want_display.
    Soporta comparación numérica y de cadena.
    """
    w = unicode(want_display or u"").strip()
    if not w:
        return -1
    wl = w.lower()
    for i, eid in enumerate(ordered_ids):
        rev = doc.GetElement(eid)
        if not isinstance(rev, Revision):
            continue
        num = (revision_number_display(rev) or u"").strip()
        if not num:
            continue
        if num == w or num.lower() == wl:
            return i
        try:
            if int(num) == int(w):
                return i
        except Exception:
            pass
    return -1


# ---------------------------------------------------------------------------
# Detección de layout de cajetín
# ---------------------------------------------------------------------------

def detect_layout(sheet, doc):
    """
    Detecta la convención de nombres de parámetros del cajetín:
    - LAYOUT_RNN_FIELD: ``R02_01_NUM``, ``R02_02_DES``, …
    - LAYOUT_R01_ROW:  ``R01_02_NUM``, ``R01_02_DES``, …
    """
    if sheet is None or doc is None:
        return LAYOUT_RNN_FIELD
    _, p_r2 = ps.lookup(sheet, doc, u"R02_01_NUM")
    if p_r2 is not None:
        return LAYOUT_RNN_FIELD
    _, p_12 = ps.lookup(sheet, doc, u"R01_02_NUM")
    if p_12 is not None:
        return LAYOUT_R01_ROW
    _, p_11 = ps.lookup(sheet, doc, u"R01_01_NUM")
    if p_11 is not None:
        return LAYOUT_R01_ROW
    return LAYOUT_RNN_FIELD


def _r_slot_prefix(slot_1based):
    """«R09_» para slot_1based=9."""
    return u"R{:02d}_".format(int(slot_1based))


def revision_num_param_name(layout, slot_1based):
    s = int(slot_1based)
    if layout == LAYOUT_R01_ROW:
        return u"R01_{:02d}_NUM".format(s)
    return _r_slot_prefix(s) + SUFFIX_NUM


def revision_slot_display(layout, slot_1based):
    """Etiqueta legible para mensajes (ej. «R01_03» o «R03»)."""
    s = int(slot_1based)
    if layout == LAYOUT_R01_ROW:
        return u"R01_{:02d}".format(s)
    return u"R{:02d}".format(s)


# ---------------------------------------------------------------------------
# Slots de cajetín
# ---------------------------------------------------------------------------

def slot_1based_in_issue_list(ids_list, target_rev_id_int):
    """Posición 1-based del ElementId de revisión en la lista de índice de lámina."""
    try:
        n = int(ids_list.Count)
    except Exception:
        return 0
    want = int(target_rev_id_int)
    for i in range(n):
        try:
            eid = ids_list[i]
        except Exception:
            try:
                eid = ids_list.get_Item(i)
            except Exception:
                continue
        if _eid_int(eid) == want:
            return i + 1
    return 0


def first_empty_slot(sheet, doc, through_slot, layout):
    """
    Primera fila en [1 … through_slot] cuyo parámetro NUM existe y está vacío.
    Devuelve -1 si no hay ninguna.
    """
    last = max(1, min(int(through_slot), MAX_REVISION_SLOTS))
    for s in range(1, last + 1):
        param_name = revision_num_param_name(layout, s)
        _, p = ps.lookup(sheet, doc, param_name)
        if p is None:
            continue
        try:
            if not p.HasValue:
                return s
        except Exception:
            pass
        txt = ps.get_text(sheet, doc, param_name)
        if txt is None:
            continue
        if not txt:
            return s
    return -1


# ---------------------------------------------------------------------------
# Escritura en cajetín
# ---------------------------------------------------------------------------

def _set_dibujo_slot(sheet, doc, layout, slot_1based, value):
    """Escribe el campo Dibujó según layout (intenta DIR y DIB)."""
    row = int(slot_1based)
    if layout == LAYOUT_R01_ROW:
        ok_dir = ps.set_named(sheet, doc, u"R01_{:02d}_DIR".format(row), value)
        ok_dib = ps.set_named(sheet, doc, u"R01_{:02d}_DIB".format(row), value)
        return ok_dir or ok_dib
    pref = _r_slot_prefix(row)
    ok_dir = ps.set_named(sheet, doc, pref + SUFFIX_DIR, value)
    ok_dib = ps.set_named(sheet, doc, pref + SUFFIX_DIB, value)
    return ok_dir or ok_dib


def write_revision_slot(sheet, doc, layout, slot_1based, numero_str,
                        description, dibujo, reviso, aprobo, fecha):
    """
    Escribe los seis campos (NUM, DES, DIR/DIB, REV, APR, FCH) de una fila de revisión.

    Si dibujo está vacío, usa reviso como valor efectivo de Dibujó
    (idéntico al comportamiento Dynamo original).
    """
    row = int(slot_1based)
    dib_eff = (dibujo or u"").strip() or (reviso or u"").strip()
    if layout == LAYOUT_R01_ROW:
        ps.set_named(sheet, doc, u"R01_{:02d}_NUM".format(row), numero_str or u"")
        ps.set_named(sheet, doc, u"R01_{:02d}_DES".format(row), description or u"")
        _set_dibujo_slot(sheet, doc, layout, row, dib_eff)
        ps.set_named(sheet, doc, u"R01_{:02d}_REV".format(row), reviso or u"")
        ps.set_named(sheet, doc, u"R01_{:02d}_APR".format(row), aprobo or u"")
        ps.set_named(sheet, doc, u"R01_{:02d}_FCH".format(row), fecha or u"")
        return
    pref = _r_slot_prefix(row)
    ps.set_named(sheet, doc, pref + SUFFIX_NUM, numero_str or u"")
    ps.set_named(sheet, doc, pref + SUFFIX_DES, description or u"")
    _set_dibujo_slot(sheet, doc, layout, row, dib_eff)
    ps.set_named(sheet, doc, pref + SUFFIX_REV, reviso or u"")
    ps.set_named(sheet, doc, pref + SUFFIX_APR, aprobo or u"")
    ps.set_named(sheet, doc, pref + SUFFIX_FCH, fecha or u"")


# ---------------------------------------------------------------------------
# Nubes de revisión
# ---------------------------------------------------------------------------

def _iter_revision_clouds(doc, sheet_id):
    bic = getattr(BuiltInCategory, "OST_RevisionClouds", None)
    if bic is not None:
        try:
            for c in (FilteredElementCollector(doc, sheet_id)
                      .OfCategory(bic)
                      .WhereElementIsNotElementType()):
                yield c
            return
        except Exception:
            pass
    if RevisionCloud is None:
        return
    try:
        for c in (FilteredElementCollector(doc, sheet_id)
                  .OfClass(RevisionCloud)
                  .WhereElementIsNotElementType()):
            yield c
    except Exception:
        return


def update_revision_cloud_count(doc, sheet, count):
    """Actualiza CANTIDAD_REVISIONES en las nubes de revisión de la lámina."""
    from Autodesk.Revit.DB import StorageType
    for c in _iter_revision_clouds(doc, sheet.Id):
        p = c.LookupParameter(PARAM_CANTIDAD_REVISIONES)
        if p is None or p.IsReadOnly:
            continue
        try:
            if p.StorageType == StorageType.String:
                p.Set(unicode(count))
            elif p.StorageType == StorageType.Integer:
                p.Set(int(count))
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Preview de nueva revisión (para columna en DataGrid)
# ---------------------------------------------------------------------------

def compute_sel_enabled(doc, sheet, ordered, emit_rev0, ti_rev0):
    """
    False solo en modo Revisión 0 cuando la última revisión del índice de la
    lámina ya es la revisión 0 del proyecto.
    """
    if not emit_rev0:
        return True
    if ti_rev0 < 0 or not ordered:
        return True
    existing = get_revisions_on_sheet(doc, sheet)
    if existing:
        _rev, fi = furthest_sheet_revision(existing, ordered)
        if _rev is None or fi < 0:
            return True
    else:
        fi = -1
    return not (fi >= 0 and fi == ti_rev0)


def preview_next_revision(doc, sheet, ordered, emit_rev0, ti_rev0):
    """
    Número mostrado de la revisión que se aplicaría (texto para columna NuevaRevision).
    """
    if not ordered:
        return u"\u2014"
    existing = get_revisions_on_sheet(doc, sheet)
    if existing:
        furthest, fi = furthest_sheet_revision(existing, ordered)
        if furthest is None or fi < 0:
            return u"(Sin coincidencia)"
    else:
        fi = -1
    ti = fi + 1
    if emit_rev0:
        if ti_rev0 < 0:
            return u"(Sin rev. 0)"
        if fi >= 0 and fi == ti_rev0:
            return u"(Ya en rev. 0)"
        ti = ti_rev0
    if ti >= len(ordered):
        return u"(Sin siguiente)"
    tgt = doc.GetElement(ordered[ti])
    if not isinstance(tgt, Revision):
        return u"\u2014"
    s = (revision_number_display(tgt) or u"").strip()
    return s if s else u"\u2014"


# ---------------------------------------------------------------------------
# Resultado de la emisión
# ---------------------------------------------------------------------------

class ApplyResult(object):
    """Resultado de apply(): número de láminas procesadas y mensajes de advertencia."""
    def __init__(self, done, errors_list):
        self.done = done
        self.errors = errors_list

    @property
    def error_text(self):
        return u"\n".join(self.errors)

    def __repr__(self):
        return u"ApplyResult(done={}, errors={})".format(self.done, len(self.errors))


# ---------------------------------------------------------------------------
# FormData — datos del formulario para la emisión
# ---------------------------------------------------------------------------

class FormData(object):
    """Datos validados del formulario de emisión."""
    def __init__(self, description, dibujo, reviso, aprobo, fecha_str,
                 emit_rev0=False):
        self.description = description
        self.dibujo = dibujo
        self.reviso = reviso
        self.aprobo = aprobo
        self.fecha_str = fecha_str
        self.emit_rev0 = emit_rev0


# ---------------------------------------------------------------------------
# Aplicación principal de revisiones
# ---------------------------------------------------------------------------

def apply(doc, sheets, form_data, revit_uiapp=None):
    """
    Aplica la revisión a la lista de láminas usando una transacción agrupada.

    Si falla una lámina, se hace rollback del lote completo hasta ese punto.

    Args:
        doc: Revit Document.
        sheets: iterable de ViewSheet.
        form_data: FormData con los datos del formulario.
        revit_uiapp: UIApplication opcional para ProgressBar y bloqueo de comandos.

    Returns:
        ApplyResult con done (int) y errors (list of str).
    """
    errs = []
    ordered = get_ordered_revision_ids(doc)
    ti_rev0 = -1
    if form_data.emit_rev0:
        ti_rev0 = index_of_revision_display_number(doc, ordered, u"0")
        if ti_rev0 < 0:
            return ApplyResult(0, [
                u"No existe ninguna revisión con número 0 en Gestión de revisiones del proyecto."
            ])

    sheets_list = list(sheets)
    ntot = len(sheets_list)
    pb = None
    pb_ok = False
    blocker = None
    blocker_ok = False

    # --- ProgressBar + bloqueo de comandos ---
    try:
        if revit_uiapp is not None and ntot > 0:
            try:
                from infra.revit_window_blocker import BloquearComandosRevit
                if BloquearComandosRevit is not None:
                    blocker = BloquearComandosRevit(revit_uiapp)
                    blocker.__enter__()
                    blocker_ok = True
            except Exception:
                blocker = None
                blocker_ok = False
            try:
                from pyrevit import forms
                from siguiente_revision.constants import PBAR_TITLE_BASE, PBAR_ACCENT_RGB
                title0 = u"{} 0/{}".format(PBAR_TITLE_BASE, ntot)
                pb = forms.ProgressBar(title=title0, cancellable=False)
                try:
                    from System.Windows.Media import Color, SolidColorBrush
                    r, g, b = PBAR_ACCENT_RGB
                    pb.Resources[u"pyRevitAccentBrush"] = SolidColorBrush(Color.FromRgb(r, g, b))
                except Exception:
                    pass
                pb.__enter__()
                pb_ok = True
            except Exception:
                pb = None
                pb_ok = False
    except Exception:
        pass

    done = 0
    tx_mega = None
    pending_writes = 0
    aborted_tx = False

    try:
        for si, sheet in enumerate(sheets_list):
            try:
                if not ordered:
                    errs.append(
                        u"{}: el proyecto no define revisiones (Gestión de revisiones).".format(
                            sheet_display(sheet)
                        )
                    )
                    continue

                existing = get_revisions_on_sheet(doc, sheet)
                if existing:
                    furthest, fi = furthest_sheet_revision(existing, ordered)
                    if furthest is None or fi < 0:
                        errs.append(
                            u"{}: las revisiones de la lámina no coinciden con la secuencia del proyecto.".format(
                                sheet_display(sheet)
                            )
                        )
                        continue
                else:
                    fi = -1

                if form_data.emit_rev0:
                    if fi >= 0 and fi == ti_rev0:
                        errs.append(
                            u"{}: modo revisión 0 omitido: la última revisión del índice ya es la revisión 0.".format(
                                sheet_display(sheet)
                            )
                        )
                        continue

                ti = fi + 1
                if form_data.emit_rev0:
                    ti = ti_rev0

                if ti >= len(ordered):
                    errs.append(
                        u"{}: no hay revisión válida después de la última del índice.".format(
                            sheet_display(sheet)
                        )
                    )
                    continue

                target_rev_id = ordered[ti]

                if tx_mega is None:
                    tx_mega = Transaction(doc, TX_NAME)
                    tx_mega.Start()

                target_rev = doc.GetElement(target_rev_id)
                if not isinstance(target_rev, Revision):
                    raise Exception(u"ElementId objetivo no es una revisión.")

                ids = ClrList[ElementId]()
                for eid in sheet.GetAdditionalRevisionIds():
                    ids.Add(eid)

                on_sheet_ids = set()
                nc_ids = int(ids.Count)
                for i in range(nc_ids):
                    try:
                        xe = ids[i]
                    except Exception:
                        try:
                            xe = ids.get_Item(i)
                        except Exception:
                            continue
                    on_sheet_ids.add(_eid_int(xe))

                for j in range(fi + 1, ti + 1):
                    ej = ordered[j]
                    jid = _eid_int(ej)
                    if jid not in on_sheet_ids:
                        ids.Add(ej)
                        on_sheet_ids.add(jid)

                ni_target = _eid_int(target_rev_id)
                geom_slot = slot_1based_in_issue_list(ids, ni_target)
                if geom_slot < 1:
                    raise Exception(u"No se pudo ubicar la revisión en la lista del índice de lámina.")

                layout = detect_layout(sheet, doc)

                if form_data.emit_rev0:
                    slot_write = first_empty_slot(sheet, doc, MAX_REVISION_SLOTS, layout)
                    if slot_write < 1:
                        raise Exception(
                            u"No hay ninguna fila R01–R{} con NUM vacío en esta lámina/cajetín.".format(
                                MAX_REVISION_SLOTS
                            )
                        )
                    _, p_chk = ps.lookup(sheet, doc, revision_num_param_name(layout, slot_write))
                    if p_chk is None:
                        raise Exception(
                            u"No existe el parámetro {} en lámina ni cajetín.".format(
                                revision_num_param_name(layout, slot_write)
                            )
                        )
                else:
                    slot_write = first_empty_slot(sheet, doc, MAX_REVISION_SLOTS, layout)
                    if slot_write < 1:
                        raise Exception(
                            u"No hay ninguna fila R01–R{} con NUM vacío en esta lámina/cajetín.".format(
                                MAX_REVISION_SLOTS
                            )
                        )
                    _, p_auto = ps.lookup(sheet, doc, revision_num_param_name(layout, slot_write))
                    if p_auto is None:
                        raise Exception(
                            u"No existe el parámetro {} en lámina ni cajetín.".format(
                                revision_num_param_name(layout, slot_write)
                            )
                        )

                sheet.SetAdditionalRevisionIds(ids)
                numero_str = revision_number_display(target_rev)
                write_revision_slot(
                    sheet, doc, layout, slot_write,
                    numero_str,
                    form_data.description,
                    form_data.dibujo,
                    form_data.reviso,
                    form_data.aprobo,
                    form_data.fecha_str,
                )

                if geom_slot != slot_write:
                    errs.append(
                        u"{}: revisión en posición {} del índice Revit; "
                        u"datos del formulario en {} (primera fila libre del cajetín).".format(
                            sheet_display(sheet),
                            geom_slot,
                            revision_slot_display(layout, slot_write),
                        )
                    )

                update_revision_cloud_count(doc, sheet, ids.Count)
                pending_writes += 1

            except Exception as ex:
                pending_writes = 0
                aborted_tx = True
                errs.append(u"{}: {}".format(sheet_display(sheet), unicode(ex)))
                try:
                    if (tx_mega is not None
                            and tx_mega.HasStarted()
                            and not tx_mega.HasEnded()):
                        tx_mega.RollBack()
                except Exception:
                    pass
                break
            finally:
                if pb_ok and pb is not None:
                    try:
                        c = int(ntot) if ntot else 1
                        i_step = int(si) + 1
                        try:
                            pb.update_progress(i_step, max_value=c)
                        except TypeError:
                            try:
                                pb.update_progress(i_step, max=c)
                            except Exception:
                                pass
                        try:
                            pb.title = u"{} {}/{}".format(
                                PBAR_TITLE_BASE, i_step, c
                            )
                        except Exception:
                            pass
                    except Exception:
                        pass

        if (not aborted_tx
                and tx_mega is not None
                and tx_mega.HasStarted()
                and not tx_mega.HasEnded()):
            try:
                tx_mega.Commit()
                done = pending_writes
            except Exception as ex_commit:
                errs.append(u"No se pudo confirmar la transacción: {0}".format(unicode(ex_commit)))
                try:
                    if tx_mega.HasStarted() and not tx_mega.HasEnded():
                        tx_mega.RollBack()
                except Exception:
                    pass

    finally:
        if pb_ok and pb is not None:
            try:
                pb.__exit__(None, None, None)
            except Exception:
                pass
        if blocker_ok and blocker is not None:
            try:
                blocker.__exit__(None, None, None)
            except Exception:
                pass

    return ApplyResult(done, errs)
