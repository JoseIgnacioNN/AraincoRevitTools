# -*- coding: utf-8 -*-
"""
Vista de dibujo (Drafting View) con despiece por línea de fierro:
croquis escalonado + etiqueta ``Ø L=total (parcial+…)`` por tramo.
"""

from __future__ import print_function

import math

import clr
clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    CurveElement,
    DetailCurve,
    ElementId,
    FilteredElementCollector,
    HorizontalTextAlignment,
    Line,
    TextNote,
    TextNoteOptions,
    TextNoteType,
    Transaction,
    VerticalTextAlignment,
    View,
    ViewDrafting,
    ViewFamily,
    ViewFamilyType,
    XYZ,
)

from column_reinforcement.linea_fierro import (
    _arma_len_mm_round_from_internal_ft,
    etiqueta_despiece_mm,
    fingerprint_seg_linea_fierro,
    linea_fierro_key_from_seg_jobs,
    linea_fierro_label_map_from_keys,
)

DRAFTING_VIEW_NAME = u"Arainco Despiece Columnas"
TEXT_NOTE_TYPE_NAME = u"2.5mm Arial"

GRID_COLS = 3
CELL_W_MM = 95.0
CELL_H_MIN_MM = 85.0
MARGIN_MM = 12.0
ROW_GAP_MM = 18.0
LAP_STAGGER_MM = 10.0
LABEL_GAP_MM = 5.0
TITLE_GAP_MM = 8.0
HOOK_DRAW_MAX_MM = 35.0


def pata_flags_from_fingerprint(fp):
    try:
        k = int(fp[0])
    except Exception:
        return False, False
    if k == 1:
        return True, False
    if k == 2:
        return False, True
    if k == 3:
        return True, True
    return False, False


def build_seg_pata_flags_by_line_idx(line_plans, line_rb_accum):
    """``line_rb_accum[line_idx]`` → ``[(troceo_ui_i, seg_i, rebar, fp), ...]``."""
    out = {}
    for lp in line_plans or []:
        line_idx = lp.get("line_idx")
        items = sorted(line_rb_accum.get(line_idx) or [], key=lambda x: (x[0], x[1]))
        if items:
            out[line_idx] = [pata_flags_from_fingerprint(x[3]) for x in items]
    return out


def _mm_to_ft(mm):
    return float(mm) / 304.8


def _first_drafting_view_family_type(document):
    for vft in FilteredElementCollector(document).OfClass(ViewFamilyType):
        try:
            if vft and vft.ViewFamily == ViewFamily.Drafting:
                return vft
        except Exception:
            pass
    return None


def _find_drafting_view_by_name(document, view_name):
    for v in FilteredElementCollector(document).OfClass(ViewDrafting):
        try:
            if v and v.Name == view_name:
                return v
        except Exception:
            pass
    return None


def _create_or_get_drafting_view(document, vft, view_name):
    v = _find_drafting_view_by_name(document, view_name)
    if v:
        return v
    for v in FilteredElementCollector(document).OfClass(View):
        try:
            if v and v.Name == view_name and not isinstance(v, ViewDrafting):
                raise RuntimeError(
                    u'Ya existe una vista "{}" que no es de dibujo.'.format(view_name)
                )
        except RuntimeError:
            raise
        except Exception:
            pass
    v = ViewDrafting.Create(document, vft.Id)
    v.Name = view_name
    return v


def _clear_detail_curves_in_view(document, drafting_view):
    view_id = drafting_view.Id
    to_delete = []
    for ce in FilteredElementCollector(document).OfClass(CurveElement):
        try:
            if ce is None or ce.OwnerViewId != view_id:
                continue
            if not isinstance(ce, DetailCurve):
                continue
            to_delete.append(ce.Id)
        except Exception:
            pass
    for did in to_delete:
        try:
            document.Delete(did)
        except Exception:
            pass


def _clear_text_notes_in_view(document, view):
    view_id = view.Id
    for el in FilteredElementCollector(document, view_id):
        try:
            if isinstance(el, TextNote):
                document.Delete(el.Id)
        except Exception:
            pass


def _clear_drafting_view_content(document, drafting_view):
    _clear_detail_curves_in_view(document, drafting_view)
    _clear_text_notes_in_view(document, drafting_view)


def _map_local_to_view_plane(view, x_local, y_local):
    right = view.RightDirection
    up = view.UpDirection
    o = view.Origin
    return XYZ(
        o.X + right.X * x_local + up.X * y_local,
        o.Y + right.Y * x_local + up.Y * y_local,
        o.Z + right.Z * x_local + up.Z * y_local,
    )


def _draw_detail_line(document, view, p1, p2):
    line = Line.CreateBound(p1, p2)
    if line is None:
        return None
    try:
        return document.Create.NewDetailCurve(view, line)
    except Exception:
        return None


def _norm_upper(s):
    try:
        return (s or u"").strip().upper()
    except Exception:
        return u""


def _first_text_note_type(document):
    for tnt in FilteredElementCollector(document).OfClass(TextNoteType):
        if tnt:
            return tnt
    return None


def _find_text_note_type_named(document, exact_name):
    en = (exact_name or u"").strip()
    if not en:
        return None
    for tnt in FilteredElementCollector(document).OfClass(TextNoteType):
        try:
            if tnt and (tnt.Name or u"").strip() == en:
                return tnt
        except Exception:
            pass
    return None


def _ensure_text_note_type(document):
    t = _find_text_note_type_named(document, TEXT_NOTE_TYPE_NAME)
    if t:
        return t
    for tnt in FilteredElementCollector(document).OfClass(TextNoteType):
        try:
            nu = _norm_upper(tnt.Name)
            if u"2.5" in nu and u"ARIAL" in nu:
                return tnt
        except Exception:
            pass
    base = _first_text_note_type(document)
    if base is None:
        return None
    try:
        new_id = base.Duplicate(TEXT_NOTE_TYPE_NAME)
        if new_id is None or new_id == ElementId.InvalidElementId:
            return base
        nt = document.GetElement(new_id)
        return nt if nt is not None else base
    except Exception:
        return base


def _create_text_note(document, view, lx, ly, txt, text_note_type, h_center=True):
    if not txt or text_note_type is None:
        return
    origin = _map_local_to_view_plane(view, lx, ly)
    try:
        opts = TextNoteOptions(text_note_type.Id)
        if h_center:
            try:
                opts.HorizontalAlignment = HorizontalTextAlignment.Center
            except Exception:
                pass
            try:
                opts.VerticalAlignment = VerticalTextAlignment.Top
            except Exception:
                pass
        TextNote.Create(document, view.Id, origin, txt, opts)
    except Exception:
        try:
            opts = TextNoteOptions()
            opts.TypeId = text_note_type.Id
            TextNote.Create(document, view.Id, origin, txt, opts)
        except Exception:
            pass


def _seg_height_mm(sj):
    return max(
        1.0,
        float(_arma_len_mm_round_from_internal_ft(sj.get("span_seg", 0.0))),
    )


def _hook_draw_mm(sj):
    lp = float(_arma_len_mm_round_from_internal_ft(sj.get("pata_hook_ft_seg", 0.0)))
    if lp < 1e-6:
        return 0.0
    return min(lp, HOOK_DRAW_MAX_MM)


def _draw_segment_schematic(
    document,
    view,
    x_center_ft,
    y_bottom_ft,
    height_ft,
    hook_mm,
    want_bot,
    want_top,
):
    x = float(x_center_ft)
    y0 = float(y_bottom_ft)
    y1 = y0 + float(height_ft)
    p_bot = _map_local_to_view_plane(view, x, y0)
    p_top = _map_local_to_view_plane(view, x, y1)
    _draw_detail_line(document, view, p_bot, p_top)
    h_ft = _mm_to_ft(hook_mm) if hook_mm > 1e-6 else 0.0
    if h_ft > 1e-9:
        if want_bot:
            p_h0 = _map_local_to_view_plane(view, x, y0)
            p_h1 = _map_local_to_view_plane(view, x + h_ft, y0)
            _draw_detail_line(document, view, p_h0, p_h1)
        if want_top:
            p_h0 = _map_local_to_view_plane(view, x, y1)
            p_h1 = _map_local_to_view_plane(view, x + h_ft, y1)
            _draw_detail_line(document, view, p_h0, p_h1)


def _representative_plan_for_key(line_plans, key, seg_pata_flags):
    for lp in line_plans or []:
        segs = sorted(lp.get("seg_jobs") or [], key=lambda s: int(s["seg_i"]))
        if not segs:
            continue
        flags = seg_pata_flags.get(lp.get("line_idx"))
        if flags is None:
            flags = [
                (bool(s.get("want_bot_pata")), bool(s.get("want_top_pata")))
                for s in segs
            ]
        did_b = [f[0] for f in flags]
        did_t = [f[1] for f in flags]
        k = linea_fierro_key_from_seg_jobs(segs, did_b, did_t)
        if k == key:
            return lp, segs, did_b, did_t
    return None, [], [], []


def _cell_content_height_mm(segs, did_b, did_t):
    total = 0.0
    for sj in segs:
        total += _seg_height_mm(sj)
    return max(CELL_H_MIN_MM, total)


def _draw_linea_fierro_cell(
    document,
    view,
    text_note_type,
    cell_x_mm,
    cell_y_mm,
    cell_w_mm,
    cell_h_mm,
    letter,
    segs,
    did_b,
    did_t,
    diam_resolver,
):
    n = len(segs)
    if n < 1:
        return

    stagger_ft = _mm_to_ft(LAP_STAGGER_MM)
    heights_mm = [_seg_height_mm(sj) for sj in segs]
    total_h_mm = sum(heights_mm)
    scale = 1.0
    usable_h = max(CELL_H_MIN_MM - TITLE_GAP_MM - LABEL_GAP_MM * 2.0, 40.0)
    if total_h_mm > usable_h:
        scale = float(usable_h) / float(total_h_mm)

    cx_cell = _mm_to_ft(cell_x_mm + cell_w_mm * 0.5)
    y_base_ft = _mm_to_ft(cell_y_mm + TITLE_GAP_MM + LABEL_GAP_MM)

    x_centers = []
    if n == 1:
        x_centers.append(cx_cell)
    else:
        span_stagger = stagger_ft * (n - 1)
        x0 = cx_cell - span_stagger * 0.5
        for i in range(n):
            x_centers.append(x0 + i * stagger_ft)

    y_cursor = y_base_ft
    for i, sj in enumerate(segs):
        h_mm = heights_mm[i] * scale
        h_ft = _mm_to_ft(h_mm)
        db = did_b[i] if i < len(did_b) else False
        dt = did_t[i] if i < len(did_t) else False
        hook_mm = _hook_draw_mm(sj) if (db or dt) else 0.0
        _draw_segment_schematic(
            document,
            view,
            x_centers[i],
            y_cursor,
            h_ft,
            hook_mm,
            db,
            dt,
        )
        fp = fingerprint_seg_linea_fierro(
            sj["span_seg"],
            db,
            dt,
            sj.get("pata_hook_ft_seg", 0.0),
        )
        d_mm = diam_resolver(sj)
        lbl = etiqueta_despiece_mm(d_mm, fp)
        lbl_y = y_cursor - _mm_to_ft(LABEL_GAP_MM + 2.0)
        _create_text_note(
            document,
            view,
            x_centers[i],
            lbl_y,
            lbl,
            text_note_type,
            h_center=True,
        )
        y_cursor += h_ft

    title = u"linea de fierro {0}".format(letter)
    title_y = _mm_to_ft(cell_y_mm + 2.0)
    _create_text_note(
        document,
        view,
        cx_cell,
        title_y,
        title,
        text_note_type,
        h_center=True,
    )


def generate_despiece_drafting_view(
    doc,
    model_groups,
    label_map,
    rebar_nominal_diameter_mm_fn,
    manage_transaction=True,
):
    """
    Vista de despiece alineada al modelo: ``model_groups`` y ``label_map`` deben
    provenir de ``collect_linea_fierro_model_groups``.
    """
    if not model_groups or not label_map:
        return None

    vft = _first_drafting_view_family_type(doc)
    if vft is None:
        return None

    t = None
    if manage_transaction:
        t = Transaction(doc, u"Arainco: Despiece líneas fierro columnas")
        t.Start()
    try:
        view = _create_or_get_drafting_view(doc, vft, DRAFTING_VIEW_NAME)
        _clear_drafting_view_content(doc, view)
        tnt = _ensure_text_note_type(doc)

        def _diam(sj):
            bt = sj.get("layout_bar_type_seg")
            if bt is None:
                return 12.0
            try:
                v = rebar_nominal_diameter_mm_fn(bt)
                return float(v) if v is not None else 12.0
            except Exception:
                return 12.0

        keys = [grp["key"] for grp in model_groups]
        n_cells = len(keys)
        cols = max(1, int(GRID_COLS))
        rows = int(math.ceil(float(n_cells) / float(cols)))

        cell_heights = []
        for grp in model_groups:
            segs = grp.get("seg_jobs") or []
            pata_flags = grp.get("pata_flags") or []
            did_b = [f[0] for f in pata_flags]
            did_t = [f[1] for f in pata_flags]
            cell_heights.append(_cell_content_height_mm(segs, did_b, did_t))

        row_heights = []
        for r in range(rows):
            row_keys = keys[r * cols : (r + 1) * cols]
            if not row_keys:
                continue
            idx0 = r * cols
            mx = max(cell_heights[idx0 : idx0 + len(row_keys)])
            row_heights.append(mx + TITLE_GAP_MM + LABEL_GAP_MM * 3.0)

        x0 = MARGIN_MM
        y_cursor = MARGIN_MM

        for idx, grp in enumerate(model_groups):
            col = idx % cols
            row = idx // cols
            if col == 0 and row > 0:
                y_cursor += row_heights[row - 1] + ROW_GAP_MM
            cell_x = x0 + col * (CELL_W_MM + MARGIN_MM)
            cell_y = y_cursor
            k = grp["key"]
            letter = label_map.get(k, u"?")
            segs = grp.get("seg_jobs") or []
            pata_flags = grp.get("pata_flags") or []
            did_b = [f[0] for f in pata_flags]
            did_t = [f[1] for f in pata_flags]
            ch = cell_heights[idx]
            _draw_linea_fierro_cell(
                doc,
                view,
                tnt,
                cell_x,
                cell_y,
                CELL_W_MM,
                ch,
                letter,
                segs,
                did_b,
                did_t,
                _diam,
            )

        if manage_transaction and t is not None:
            t.Commit()
        return view
    except Exception:
        if (
            manage_transaction
            and t is not None
            and t.HasStarted()
            and not t.HasEnded()
        ):
            t.RollBack()
        raise
