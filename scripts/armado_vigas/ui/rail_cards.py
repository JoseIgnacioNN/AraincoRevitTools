# -*- coding: utf-8 -*-
"""Rail derecho mockup Opción D — pestañas SUP / INF / LAT / CONF + tarjeta SUPLE aparte.

Solo estado de UI/preview. No dispara colocación de Rebar.
"""

from __future__ import division

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import (
    FontWeights,
    GridLength,
    GridUnitType,
    HorizontalAlignment,
    TextWrapping,
    Thickness,
    VerticalAlignment,
)
from System.Windows import RoutedEventHandler
from System.Windows.Controls import (
    Border,
    Button,
    ColumnDefinition,
    DockPanel,
    Grid,
    Orientation,
    StackPanel,
    TextBlock,
    WrapPanel,
)
from System.Windows.Controls import Dock
from System.Windows.Input import Cursors

from armado_vigas.domain.confinement import (
    conf_draft_label,
    ensure_beam_confinement,
    get_conf_draft,
    is_conf_draft_defined,
)
from armado_vigas.domain.constants import (
    BAR_COUNT_MAX,
    BAR_COUNT_MIN,
    CAPAS_MAX,
    CAPAS_MIN,
    CONCRETE_GRADE_DEFAULT,
    CONCRETE_GRADES,
    ESTRIBO_SPACING_DEFAULT_CENT as _SP_CENT,
    ESTRIBO_SPACING_DEFAULT_EXT as _SP_EXT,
    LONG_DIAM_OPTS,
    normalize_concrete_grade,
)
from armado_vigas.domain.laterales import (
    LATERALES_COUNT_MAX,
    LATERALES_COUNT_MIN,
    LATERALES_DIAM_DEFAULT,
    session_n_laterales,
)
from armado_vigas.domain.layers import (
    beam_n_capas_inf,
    beam_n_capas_sup,
    clamp_bar_count,
    ensure_beam_layers,
    is_global_layer_sync_field,
    layer_keys,
    set_first_layer_bar_count,
    sync_layer_field_all_beams,
)
from armado_vigas.domain.stirrups import (
    STIRRUP_ZONE_AUTO,
    compute_stirrup_zones,
    ensure_beam_stirrup_zone_mode,
    normalize_stirrup_zone_mode,
    recommended_stirrup_zone_mode,
    section_height_mm,
    stirrup_zone_mode_label,
    stirrup_zone_mode_labels,
)
from armado_vigas.domain.suple_inferior import (
    beam_suple_inf_enabled,
    ensure_beam_suple_inferior,
)
from armado_vigas.domain.suple_superior import (
    adjacent_beams_for_apoyo,
    active_apoyos_suple_sup,
    beam_suple_sup_enabled,
    clear_all_suple_sup_apoyos,
    ensure_beam_suple_superior,
    ensure_session_suple_sup,
    get_apoyo_suple_sup_arm,
    is_apoyo_suple_sup_on,
    set_apoyo_suple_sup,
    set_apoyo_suple_sup_arm,
    sync_beams_suple_arm_from_apoyos,
    sync_beams_suple_from_apoyo_set,
    toggle_apoyo_suple_sup,
)
from armado_vigas.ui import theme as th
from armado_vigas.ui import typography as typo
from armado_vigas.ui.theme import make_role_badge
from armado_vigas.ui.wpf_controls import (
    brush_hex,
    label_small,
    make_bar_count_combo,
    make_capas_combo,
    make_diam_combo,
    make_spacing_input,
    make_int_combo,
    make_string_combo,
    make_yesno_toggle,
)

RAIL_TABS = (
    (u"sup", u"SUP", u"#22d3ee"),
    (u"inf", u"INF", u"#fb7185"),
    (u"lat", u"LAT", u"#a78bfa"),
    (u"conf", u"CONF", u"#5bb8d4"),
)

ESTRIBO_DIAM_OPTS = (8, 10, 12, 16)


def _corner(border, r=4.0):
    try:
        from System.Windows import CornerRadius

        border.CornerRadius = CornerRadius(float(r))
    except Exception:
        pass
    return border


def _layer_ordinal_label(layer_num_1based):
    """Etiqueta ordinal tipo Armado Muros: «1ª C.»."""
    return u"{0}ª C.".format(int(layer_num_1based))


def _layer_summary_text(n_bars, diam_mm):
    """Resumen compacto tipo Muros: «ø16 × 3»."""
    try:
        nb = int(n_bars)
    except Exception:
        nb = BAR_COUNT_MIN
    try:
        dm = int(diam_mm)
    except Exception:
        dm = 16
    return u"ø{0} × {1}".format(dm, nb)


def _field_stack(label, control):
    """Etiqueta arriba + control debajo (legible en rail estrecho)."""
    sp = StackPanel()
    sp.Orientation = Orientation.Vertical
    sp.HorizontalAlignment = HorizontalAlignment.Stretch
    tb = TextBlock()
    tb.Text = label or u""
    tb.Foreground = th.brush_fg_mid()
    tb.FontSize = typo.LABEL_FONT_PX
    tb.FontWeight = FontWeights.SemiBold
    tb.Margin = Thickness(0, 0, 0, 3)
    sp.Children.Add(tb)
    try:
        control.HorizontalAlignment = HorizontalAlignment.Stretch
        control.VerticalAlignment = VerticalAlignment.Center
        control.Margin = Thickness(0)
    except Exception:
        pass
    sp.Children.Add(control)
    return sp


def _label_control_row(label, control, label_w=28.0):
    """Fila etiqueta + control (legacy / rows horizontales densas)."""
    dp = DockPanel()
    dp.LastChildFill = True
    dp.Margin = Thickness(0)
    tb = TextBlock()
    tb.Text = label or u""
    tb.Width = float(label_w)
    tb.VerticalAlignment = VerticalAlignment.Center
    tb.Foreground = th.brush_fg_lo()
    tb.FontSize = typo.LABEL_FONT_PX
    DockPanel.SetDock(tb, Dock.Left)
    dp.Children.Add(tb)
    try:
        control.HorizontalAlignment = HorizontalAlignment.Stretch
        control.VerticalAlignment = VerticalAlignment.Center
        control.Margin = Thickness(0)
    except Exception:
        pass
    dp.Children.Add(control)
    return dp


def _build_layer_config_card(
    cv, owner, tramo_beams, layer_num, is_sup, accent, highlighted=False,
    face=None, tramos=None, primary_tramo=None,
):
    """Tarjeta por capa: ordinal + resumen · Barras | Diámetro en 2 columnas."""
    from armado_vigas.domain.tramo_armado import owner_display_value

    k = layer_keys(layer_num)
    qty_field = k["nSup"] if is_sup else k["nInf"]
    diam_field = k["diamSup"] if is_sup else k["diamInf"]
    face = face or (u"sup" if is_sup else u"inf")
    session = _session_for_cv(cv)
    try:
        if layer_num == 1:
            from armado_vigas.domain.tramo_armado import resolve_first_layer_n_linked

            n_bars = resolve_first_layer_n_linked(
                session, face, primary_tramo, owner
            )
        else:
            n_raw = owner_display_value(session, face, primary_tramo, qty_field, owner)
            n_bars = clamp_bar_count(n_raw if n_raw is not None else BAR_COUNT_MIN)
    except Exception:
        n_bars = BAR_COUNT_MIN
    try:
        d_raw = owner_display_value(session, face, primary_tramo, diam_field, owner)
        diam_val = int(d_raw if d_raw is not None else 16)
    except Exception:
        diam_val = 16

    card = Border()
    card.Background = brush_hex(u"#0c1824", 255)
    # Contorno suave; 1ª capa solo acento lateral (menos ruido).
    card.BorderBrush = th.brush_border_muted() if not highlighted else brush_hex(accent, 140)
    card.BorderThickness = Thickness(1)
    card.Padding = Thickness(8, 8, 8, 8)
    card.Margin = Thickness(0, 0, 0, 8)
    card.HorizontalAlignment = HorizontalAlignment.Stretch
    _corner(card, 4.0)

    root = Grid()
    root.HorizontalAlignment = HorizontalAlignment.Stretch
    if highlighted:
        try:
            cd_acc = ColumnDefinition()
            cd_acc.Width = GridLength(3.0, GridUnitType.Pixel)
            cd_body = ColumnDefinition()
            cd_body.Width = GridLength(1.0, GridUnitType.Star)
            root.ColumnDefinitions.Add(cd_acc)
            root.ColumnDefinitions.Add(cd_body)
            accent_bar = Border()
            accent_bar.Background = brush_hex(accent, 200)
            accent_bar.Margin = Thickness(0, 0, 8, 0)
            accent_bar.HorizontalAlignment = HorizontalAlignment.Stretch
            accent_bar.VerticalAlignment = VerticalAlignment.Stretch
            _corner(accent_bar, 1.5)
            Grid.SetColumn(accent_bar, 0)
            root.Children.Add(accent_bar)
            body_col = 1
        except Exception:
            body_col = 0
    else:
        cd_body = ColumnDefinition()
        cd_body.Width = GridLength(1.0, GridUnitType.Star)
        root.ColumnDefinitions.Add(cd_body)
        body_col = 0

    body = StackPanel()
    body.HorizontalAlignment = HorizontalAlignment.Stretch

    # Cabecera: ordinal + chip resumen
    hdr = DockPanel()
    hdr.LastChildFill = True
    hdr.Margin = Thickness(0, 0, 0, 8)

    sum_chip = Border()
    sum_chip.Padding = Thickness(6, 1, 6, 1)
    sum_chip.Background = brush_hex(accent, 28)
    sum_chip.BorderBrush = brush_hex(accent, 90)
    sum_chip.BorderThickness = Thickness(1)
    _corner(sum_chip, 3.0)
    DockPanel.SetDock(sum_chip, Dock.Right)
    summary = TextBlock()
    summary.Text = _layer_summary_text(n_bars, diam_val)
    summary.Foreground = brush_hex(accent)
    summary.FontSize = 10.0
    summary.FontWeight = FontWeights.SemiBold
    summary.VerticalAlignment = VerticalAlignment.Center
    sum_chip.Child = summary
    hdr.Children.Add(sum_chip)

    title = TextBlock()
    title.Text = _layer_ordinal_label(layer_num)
    title.FontSize = 11.0
    title.FontWeight = FontWeights.Bold
    title.Foreground = th.brush_fg_hi()
    title.VerticalAlignment = VerticalAlignment.Center
    hdr.Children.Add(title)
    body.Children.Add(hdr)

    if layer_num == 1:
        note = TextBlock()
        note.Text = u"n 1ª capa ligada SUP↔INF (conf. E/T) · ø propio de cada cara"
        note.Foreground = th.brush_fg_lo()
        note.FontSize = typo.META_FONT_PX
        note.Margin = Thickness(0, 0, 0, 6)
        try:
            from System.Windows import TextWrapping

            note.TextWrapping = TextWrapping.Wrap
        except Exception:
            pass
        body.Children.Add(note)

    n_cb = make_bar_count_combo(
        _win(cv),
        n_bars,
        lambda v, tb=tramo_beams, o=owner, f=qty_field, s=summary, d=diam_val, fc=face, tr=tramos: (
            _apply_layer_and_refresh_summary(
                cv, tb, o, f, clamp_bar_count(v), s, d, is_qty=True, face=fc, tramos=tr,
            )
        ),
        compact=True,
        stretch=True,
    )
    diam_cb = make_diam_combo(
        _win(cv),
        diam_val,
        _session_bar_opts(cv, diam_val),
        lambda v, tb=tramo_beams, o=owner, f=diam_field, s=summary, n0=n_bars, fc=face, tr=tramos: (
            _apply_layer_and_refresh_summary(
                cv, tb, o, f, v, s, n0, is_qty=False, face=fc, tramos=tr,
            )
        ),
        compact=True,
        stretch=True,
    )

    row = Grid()
    row.HorizontalAlignment = HorizontalAlignment.Stretch
    cd_n = ColumnDefinition()
    # n ocupa ~36 %; ø el resto (texto ø16 más ancho)
    cd_n.Width = GridLength(0.36, GridUnitType.Star)
    cd_g = ColumnDefinition()
    cd_g.Width = GridLength(10.0, GridUnitType.Pixel)
    cd_d = ColumnDefinition()
    cd_d.Width = GridLength(0.64, GridUnitType.Star)
    row.ColumnDefinitions.Add(cd_n)
    row.ColumnDefinitions.Add(cd_g)
    row.ColumnDefinitions.Add(cd_d)

    left = _field_stack(u"Barras (n)", n_cb)
    right = _field_stack(u"Diámetro", diam_cb)
    Grid.SetColumn(left, 0)
    Grid.SetColumn(right, 2)
    row.Children.Add(left)
    row.Children.Add(right)
    body.Children.Add(row)

    Grid.SetColumn(body, body_col)
    root.Children.Add(body)
    card.Child = root
    return card


def _apply_layer_and_refresh_summary(
    cv, tramo_beams, owner, field, value, summary_tb, other_val, is_qty=True,
    face=None, tramos=None,
):
    """Aplica cambio de capa y actualiza el texto resumen de la tarjeta (sin segunda lógica de dominio)."""
    _apply_layer_field(cv, tramo_beams, owner, field, value, face=face, tramos=tramos)
    try:
        if summary_tb is None:
            return
        if is_qty:
            n_bars = clamp_bar_count(value)
            diam_mm = int(other_val or 16)
        else:
            diam_mm = int(value or 16)
            n_bars = clamp_bar_count(other_val or BAR_COUNT_MIN)
        # Tras apply, preferir config del tramo primario.
        from armado_vigas.domain.tramo_armado import owner_display_value

        session = _session_for_cv(cv)
        face = face or (u"inf" if (field or u"").startswith((u"nInf", u"diamInf")) else u"sup")
        primary = None
        if tramos is not None:
            sel = _selected_tramos_list(cv, face, tramos)
            primary = sel[0] if sel else None
            if not primary and tramos:
                primary = tramos[0]
        if primary is not None:
            if field in (u"nSup", u"nInf", u"nSup2", u"nInf2", u"nSup3", u"nInf3"):
                n_v = owner_display_value(session, face, primary, field, owner)
                n_bars = clamp_bar_count(n_v if n_v is not None else n_bars)
            if field.startswith(u"diam"):
                d_v = owner_display_value(session, face, primary, field, owner)
                diam_mm = int(d_v if d_v is not None else diam_mm)
                qty_peer = {
                    u"diamSup": u"nSup",
                    u"diamInf": u"nInf",
                    u"diamSup2": u"nSup2",
                    u"diamInf2": u"nInf2",
                    u"diamSup3": u"nSup3",
                    u"diamInf3": u"nInf3",
                }.get(field)
                if qty_peer:
                    n_v = owner_display_value(session, face, primary, qty_peer, owner)
                    n_bars = clamp_bar_count(n_v if n_v is not None else n_bars)
        elif owner is not None:
            if field in (u"nSup", u"nInf", u"nSup2", u"nInf2", u"nSup3", u"nInf3"):
                n_bars = clamp_bar_count(owner.get(field) or n_bars)
            if field.startswith(u"diam"):
                diam_mm = int(owner.get(field) or diam_mm)
        summary_tb.Text = _layer_summary_text(n_bars, diam_mm)
    except Exception:
        pass


def _build_capas_block(cv, owner, tramo_beams, is_sup, accent, tramos=None):
    """Bloque Capas del rail SUP/INF — contador + una card por capa activa."""
    block = StackPanel()
    block.Margin = Thickness(0, 0, 0, 0)
    block.HorizontalAlignment = HorizontalAlignment.Stretch

    side = u"sup" if is_sup else u"inf"
    face = side
    n_capas = beam_n_capas_sup(owner) if is_sup else beam_n_capas_inf(owner)
    primary = None
    if tramos:
        sel = _selected_tramos_list(cv, face, tramos)
        primary = sel[0] if sel else (tramos[0] if tramos else None)

    # Fila compacta: "Nº capas" a la izquierda, combo fijo a la derecha.
    cap_row = DockPanel()
    cap_row.LastChildFill = True
    cap_row.Margin = Thickness(0, 0, 0, 10)
    cap_cb = make_capas_combo(
        _win(cv),
        n_capas,
        lambda v, tb=tramo_beams, s=side: _apply_capas(cv, tb, s, v),
        compact=True,
        stretch=False,
    )
    try:
        cap_cb.Width = 52.0
        cap_cb.MinWidth = 52.0
        cap_cb.MaxWidth = 52.0
        cap_cb.HorizontalAlignment = HorizontalAlignment.Right
    except Exception:
        pass
    DockPanel.SetDock(cap_cb, Dock.Right)
    cap_row.Children.Add(cap_cb)

    cap_lbl = TextBlock()
    cap_lbl.Text = u"Nº de capas"
    cap_lbl.FontWeight = FontWeights.SemiBold
    cap_lbl.Foreground = th.brush_fg_mid()
    cap_lbl.FontSize = typo.LABEL_FONT_PX
    cap_lbl.VerticalAlignment = VerticalAlignment.Center
    cap_row.Children.Add(cap_lbl)
    block.Children.Add(cap_row)

    # Solo capas activas (sin filas L* inactivas ni cabecera ARMADO redundante).
    for layer_num in range(1, n_capas + 1):
        if layer_num > CAPAS_MAX:
            break
        card = _build_layer_config_card(
            cv, owner, tramo_beams, layer_num, is_sup, accent,
            highlighted=(layer_num == 1),
            face=face, tramos=tramos, primary_tramo=primary,
        )
        block.Children.Add(card)

    return block


def _section_label(text):
    tb = TextBlock()
    tb.Text = (text or u"").upper()
    tb.Foreground = th.brush_fg_lo()
    tb.FontSize = typo.META_FONT_PX
    tb.FontWeight = FontWeights.Bold
    tb.Margin = Thickness(0, 0, 0, 6)
    return tb


def _hint(text):
    tb = TextBlock()
    tb.Text = text or u""
    tb.Foreground = th.brush_fg_lo()
    tb.FontSize = typo.META_FONT_PX
    try:
        from System.Windows import TextWrapping

        tb.TextWrapping = TextWrapping.Wrap
    except Exception:
        pass
    tb.Margin = Thickness(0, 0, 0, 8)
    return tb


def _status(cv, msg):
    cb = getattr(cv, "_cb", None) or {}
    cb.get("on_status", lambda _m: None)(msg)


def _redraw(cv):
    cb = getattr(cv, "_cb", None) or {}
    cb.get("on_redraw", lambda: None)()


def _win(cv):
    return getattr(cv, "_win", None)


def _session_bar_opts(cv, current_mm=None):
    session = getattr(cv, "_last_session", None)
    opts = getattr(session, "bar_diameters_mm", None) if session is not None else None
    opts = opts or LONG_DIAM_OPTS
    if current_mm is None:
        return opts
    try:
        cur = int(round(float(current_mm)))
    except Exception:
        return opts
    if cur in opts:
        return opts
    return tuple(sorted(set(list(opts) + [cur])))


def ensure_rail_state(cv):
    """Estado de pestañas / fases en el canvas view (persistente entre redraw)."""
    if not hasattr(cv, "rail_card") or cv.rail_card not in (u"sup", u"inf", u"lat", u"conf"):
        cv.rail_card = u"sup"
    if not hasattr(cv, "card_on_sup"):
        cv.card_on_sup = True
    if not hasattr(cv, "card_on_inf"):
        cv.card_on_inf = True
    if not hasattr(cv, "card_on_lat"):
        cv.card_on_lat = True
    if not hasattr(cv, "card_on_conf"):
        cv.card_on_conf = True
    if not hasattr(cv, "conf_face"):
        cv.conf_face = u"sup"
    if not hasattr(cv, "selected_tramo_ids_sup"):
        cv.selected_tramo_ids_sup = set()
    if not hasattr(cv, "selected_tramo_ids_inf"):
        cv.selected_tramo_ids_inf = set()


def _card_enabled(cv, card_id):
    return {
        u"sup": bool(getattr(cv, "card_on_sup", True)),
        u"inf": bool(getattr(cv, "card_on_inf", True)),
        u"lat": bool(getattr(cv, "card_on_lat", True)),
        u"conf": bool(getattr(cv, "card_on_conf", True)),
    }.get(card_id, True)


def _set_card_enabled(cv, card_id, value):
    on = bool(value)
    if card_id == u"sup":
        cv.card_on_sup = on
        _status(cv, u"SUP activado" if on else u"SUP desactivado · no se colocará")
    elif card_id == u"inf":
        cv.card_on_inf = on
        _status(cv, u"INF activado" if on else u"INF desactivado · no se colocará")
    elif card_id == u"lat":
        cv.card_on_lat = on
        session = getattr(cv, "_last_session", None)
        if session is not None:
            session.lateralesEnabled = on
        _status(cv, u"LAT activado" if on else u"LAT desactivado · no se colocará")
    elif card_id == u"conf":
        cv.card_on_conf = on
        _status(cv, u"CONF activado" if on else u"CONF desactivado · no se colocará")
    _redraw(cv)


def _select_rail_card(cv, card_id):
    ensure_rail_state(cv)
    if card_id not in (u"sup", u"inf", u"lat", u"conf"):
        return
    prev = getattr(cv, u"rail_card", None)
    cv.rail_card = card_id
    if card_id in (u"sup", u"inf"):
        cv.conf_face = card_id
    if prev != card_id:
        try:
            cv._conf_pending = None
            cv._conf_hover = None
            cv._conf_origin = None
            cv._conf_cursor = None
        except Exception:
            pass
    _status(cv, u"Panel {0}".format(card_id.upper()))
    _redraw(cv)


def _normalize_tramo_multi(cv, face, tramos):
    ids_attr = u"selected_tramo_ids_sup" if face == u"sup" else u"selected_tramo_ids_inf"
    primary_attr = u"selected_tramo_sup_id" if face == u"sup" else u"selected_tramo_inf_id"
    ordered = [t.get("id") for t in (tramos or []) if t.get("id") is not None]
    cur = getattr(cv, ids_attr, None) or set()
    valid = {tid for tid in cur if tid in ordered}
    primary = getattr(cv, primary_attr, None)
    if not valid:
        if primary in ordered:
            valid = {primary}
        elif ordered:
            valid = {ordered[0]}
            primary = ordered[0]
    if primary not in valid and valid:
        primary = min(valid)
    setattr(cv, ids_attr, valid)
    setattr(cv, primary_attr, primary)
    return valid, primary


def _select_tramo_multi(cv, face, tramo_id, ctrl=False):
    """Selecciona tramo(s): clic = uno; Ctrl+clic = toggle.

    - SUP: solo tramos de cara superior (no cambia de pestaña).
    - INF: solo tramos inferiores (no cambia a CONF).
    - CONF: tramos sup o inf; se mantiene en CONF.
    """
    ensure_rail_state(cv)
    card = getattr(cv, u"rail_card", None) or u"sup"
    if card not in (u"sup", u"inf", u"conf"):
        return
    # Cara acotada a la pestaña (salvo CONF, que admite ambas).
    if card == u"sup" and face != u"sup":
        return
    if card == u"inf" and face != u"inf":
        return
    ids_attr = u"selected_tramo_ids_sup" if face == u"sup" else u"selected_tramo_ids_inf"
    primary_attr = u"selected_tramo_sup_id" if face == u"sup" else u"selected_tramo_inf_id"
    selected = set(getattr(cv, ids_attr, None) or set())
    primary = getattr(cv, primary_attr, None)
    if ctrl:
        # Si el set aún no estaba hidratado, partir del primario actual.
        if not selected and primary is not None:
            selected.add(primary)
        if tramo_id in selected:
            if len(selected) > 1:
                selected.discard(tramo_id)
                if primary == tramo_id or primary not in selected:
                    primary = min(selected)
            # No deseleccionar el último.
        else:
            selected.add(tramo_id)
            primary = tramo_id
        setattr(cv, primary_attr, primary)
    else:
        selected = {tramo_id}
        setattr(cv, primary_attr, tramo_id)
    setattr(cv, ids_attr, selected)
    # Recordar cara en CONF; no forzar navegación a CONF.
    if card == u"conf":
        cv.conf_face = face
    elif card == u"sup":
        cv.rail_card = u"sup"
        cv.conf_face = u"sup"
    elif card == u"inf":
        cv.rail_card = u"inf"
        cv.conf_face = u"inf"
    cb = getattr(cv, "_cb", None) or {}
    cb.get("on_select_tramo", lambda _t, _f: None)(tramo_id, face)
    _redraw(cv)


def _beam_indices_for_tramos(tramos, selected_ids):
    want = set(selected_ids or [])
    out = []
    for t in tramos or []:
        if t.get("id") not in want:
            continue
        for i in t.get("beamIndices") or []:
            try:
                ii = int(i)
            except Exception:
                continue
            if ii not in out:
                out.append(ii)
    return out


def _tramo_beams(cv, tramos, selected_ids, beams):
    indices = _beam_indices_for_tramos(tramos, selected_ids)
    out = []
    for i in indices:
        if 0 <= i < len(beams):
            out.append(beams[i])
    return out


def _owner_for_face(cv, tramos, selected_ids, beams, primary_id):
    tb = _tramo_beams(cv, tramos, selected_ids, beams)
    if tb:
        # Prefer owner del tramo primario
        for t in tramos or []:
            if t.get("id") == primary_id:
                for i in t.get("beamIndices") or []:
                    try:
                        ii = int(i)
                    except Exception:
                        continue
                    if 0 <= ii < len(beams):
                        return beams[ii], tb
        return tb[0], tb
    if beams:
        idx = getattr(cv, "selected_beam_idx", 0)
        if not (0 <= idx < len(beams)):
            idx = 0
        return beams[idx], [beams[idx]]
    return None, []


def build_rail_tabs(cv):
    """Pestañas SUP/INF/LAT/CONF a ancho completo del rail (4 columnas * iguales)."""
    ensure_rail_state(cv)
    grid = Grid()
    grid.Margin = Thickness(0, 0, 0, 10)
    grid.HorizontalAlignment = HorizontalAlignment.Stretch
    n = len(RAIL_TABS)
    for _ in range(n):
        cd = ColumnDefinition()
        cd.Width = GridLength(1.0, GridUnitType.Star)
        grid.ColumnDefinitions.Add(cd)

    active = cv.rail_card
    for i, (card_id, label, color) in enumerate(RAIL_TABS):
        sel = active == card_id
        on = _card_enabled(cv, card_id)
        btn = Button()
        btn.Content = (label + (u" · off" if not on else u""))
        btn.Padding = Thickness(4, 6, 4, 6)
        btn.FontSize = 11.0
        btn.FontWeight = FontWeights.Bold if sel else FontWeights.Normal
        btn.Cursor = Cursors.Hand
        btn.HorizontalAlignment = HorizontalAlignment.Stretch
        btn.HorizontalContentAlignment = HorizontalAlignment.Center
        # Separación entre celdas (sin hueco muerto a la derecha del último).
        if i < n - 1:
            btn.Margin = Thickness(0, 0, 4, 0)
        else:
            btn.Margin = Thickness(0)
        try:
            btn.MinWidth = 0.0
        except Exception:
            pass
        if sel:
            btn.BorderBrush = brush_hex(color)
            btn.BorderThickness = Thickness(1.5)
            btn.Background = brush_hex(u"#0E1B32")
            btn.Foreground = th.brush_fg_hi()
            btn.Opacity = 1.0 if on else 0.55
        else:
            btn.BorderBrush = th.brush_border_muted()
            btn.BorderThickness = Thickness(1)
            btn.Background = th.brush_app()
            btn.Foreground = th.brush_fg_lo()
            btn.Opacity = 0.38 if on else 0.22

        def _click(sender, args, cid=card_id):
            _select_rail_card(cv, cid)

        try:
            btn.Click += RoutedEventHandler(_click)
        except Exception:
            pass
        Grid.SetColumn(btn, i)
        grid.Children.Add(btn)
    return grid


def _rail_pill(label, color):
    return make_role_badge(label, u"suple" if u"SUPLE" in (label or u"").upper() else u"cent")


def _card_shell(enabled, accent):
    block = Border()
    block.Margin = Thickness(0, 0, 0, 0)
    block.Padding = Thickness(10)
    block.Background = th.brush_panel()
    block.BorderBrush = brush_hex(accent) if enabled else th.brush_border()
    block.BorderThickness = Thickness(1.5 if enabled else 1)
    block.Opacity = 1.0 if enabled else 0.55
    _corner(block, 4)
    return block


def _hdr_row(title, pills, toggle_widget=None):
    """Cabecera de card: título (wrap) + pills + toggle a la derecha."""
    dock = DockPanel()
    dock.LastChildFill = True
    dock.Margin = Thickness(0, 0, 0, 8)

    right = StackPanel()
    right.Orientation = Orientation.Horizontal
    right.VerticalAlignment = VerticalAlignment.Center
    for p in pills or []:
        if p is not None:
            p.Margin = Thickness(4, 0, 0, 0)
            p.VerticalAlignment = VerticalAlignment.Center
            right.Children.Add(p)
    if toggle_widget is not None:
        toggle_widget.Margin = Thickness(8, 0, 0, 0)
        toggle_widget.VerticalAlignment = VerticalAlignment.Center
        right.Children.Add(toggle_widget)
    DockPanel.SetDock(right, Dock.Right)
    dock.Children.Add(right)

    tb = TextBlock()
    tb.Text = title or u""
    tb.Foreground = th.brush_fg_hi()
    tb.FontSize = 12.0
    tb.FontWeight = FontWeights.SemiBold
    tb.VerticalAlignment = VerticalAlignment.Center
    try:
        from System.Windows import TextWrapping

        tb.TextWrapping = TextWrapping.Wrap
    except Exception:
        pass
    dock.Children.Add(tb)
    return dock


def _session_or_default(cv, session=None):
    if session is not None:
        return session
    return getattr(cv, "_last_session", None)


def _set_concrete_grade_from_rail(cv, grade):
    """Preferencia global de sesión (G25/G35/G45) · actualiza tablas y preview."""
    g = normalize_concrete_grade(grade)
    session = _session_or_default(cv)
    try:
        if session is not None and hasattr(session, "set_concrete_grade"):
            g = session.set_concrete_grade(g)
        else:
            from armado_vigas.revit.session import SESSION

            g = SESSION.set_concrete_grade(g)
    except Exception:
        pass
    _status(
        cv,
        u"Dosificación {0} · tablas traslape / empotramiento / pata L actualizadas.".format(
            g
        ),
    )
    _redraw(cv)
    return g


def _dosif_active_colors(grade):
    """(bg_hex, border_hex, fg_hex) para el chip activo de dosificación."""
    g = normalize_concrete_grade(grade)
    if g == u"G35":
        return (u"#4c1d95", u"#a78bfa", u"#ede9fe")
    if g == u"G45":
        return (u"#9d174d", u"#f472b6", u"#fce7f3")
    return (u"#0c4a6e", u"#38bdf8", u"#e0f2fe")


def build_dosificacion_segment(cv, session=None):
    """Segmented G25 | G35 | G45 — preferencia de lote en rail (variante A)."""
    session = _session_or_default(cv, session)
    try:
        current = normalize_concrete_grade(
            getattr(session, "concreteGrade", None) if session is not None else None
        )
    except Exception:
        current = CONCRETE_GRADE_DEFAULT

    root = StackPanel()
    root.Margin = Thickness(0, 10, 0, 0)

    lbl = TextBlock()
    lbl.Text = u"Dosificación del hormigón · lote"
    lbl.Foreground = th.brush_fg_lo()
    lbl.FontSize = typo.LABEL_FONT_PX
    lbl.FontWeight = FontWeights.SemiBold
    lbl.Margin = Thickness(0, 0, 0, 5)
    root.Children.Add(lbl)

    grid = Grid()
    grid.HorizontalAlignment = HorizontalAlignment.Stretch
    for _ in range(3):
        cd = ColumnDefinition()
        cd.Width = GridLength(1.0, GridUnitType.Star)
        grid.ColumnDefinitions.Add(cd)

    for i, lab in enumerate(CONCRETE_GRADES):
        active = lab == current
        btn = Button()
        btn.Content = lab
        btn.Tag = lab
        btn.Height = 28.0
        btn.Padding = Thickness(2, 4, 2, 4)
        btn.FontSize = typo.CTRL_FONT_PX
        btn.FontWeight = FontWeights.Bold
        btn.Margin = Thickness(0 if i == 0 else 3, 0, 0, 0)
        btn.HorizontalAlignment = HorizontalAlignment.Stretch
        try:
            btn.Cursor = Cursors.Hand
        except Exception:
            pass
        if active:
            bg, br, fg = _dosif_active_colors(lab)
            btn.Background = brush_hex(bg)
            btn.BorderBrush = brush_hex(br)
            btn.Foreground = brush_hex(fg)
            btn.BorderThickness = Thickness(1.2)
        else:
            btn.Background = brush_hex(u"#0a1520")
            btn.BorderBrush = th.brush_border()
            btn.Foreground = th.brush_fg_mid()
            btn.BorderThickness = Thickness(1)
        try:
            btn.ToolTip = (
                u"Dosificación {0} · tablas de traslape, empotramiento y pata L "
                u"para todo el lote"
            ).format(lab)
        except Exception:
            pass

        def _on(s, e, g=lab):
            sess = _session_or_default(cv)
            cur = normalize_concrete_grade(
                getattr(sess, "concreteGrade", None) if sess is not None else None
            )
            if normalize_concrete_grade(g) == cur:
                return
            _set_concrete_grade_from_rail(cv, g)

        try:
            btn.Click += RoutedEventHandler(_on)
        except Exception:
            pass
        Grid.SetColumn(btn, i)
        grid.Children.Add(btn)

    root.Children.Add(grid)

    hint = TextBlock()
    hint.Text = u"Global · actualiza traslape, empotramiento y pata L"
    hint.Foreground = th.brush_fg_lo()
    hint.FontSize = 9.0
    hint.Margin = Thickness(0, 5, 0, 0)
    try:
        hint.TextWrapping = TextWrapping.Wrap
    except Exception:
        pass
    root.Children.Add(hint)
    return root


def build_config_viga_header(cv, beam, beams, session=None):
    block = Border()
    block.Margin = Thickness(0, 0, 0, 10)
    block.Padding = Thickness(10)
    block.Background = th.brush_panel()
    block.BorderBrush = th.brush_border()
    block.BorderThickness = Thickness(1)
    _corner(block, 4)
    sp = StackPanel()

    row = StackPanel()
    row.Orientation = Orientation.Horizontal
    row.Margin = Thickness(0, 0, 0, 6)
    title = TextBlock()
    title.Text = u"Configuración viga"
    title.Foreground = th.brush_fg_hi()
    title.FontSize = 12.0
    title.FontWeight = FontWeights.SemiBold
    row.Children.Add(title)
    pill = make_role_badge(u"VIGA", u"confin")
    pill.Margin = Thickness(8, 0, 0, 0)
    row.Children.Add(pill)
    sp.Children.Add(row)

    meta = Border()
    meta.Background = brush_hex(u"#0E1B32")
    meta.BorderBrush = th.brush_border()
    meta.BorderThickness = Thickness(1)
    meta.Padding = Thickness(8, 3, 8, 3)
    meta.Margin = Thickness(0, 0, 0, 0)
    _corner(meta, 3)
    mt = TextBlock()
    if beam:
        idx = getattr(cv, "selected_beam_idx", 0)
        if not (0 <= idx < len(beams or [])):
            idx = 0
        n_sel = len(getattr(cv, "selected_beam_indices", None) or set() or {idx})
        extra = u""
        if n_sel > 1:
            extra = u" · lote {0}".format(n_sel)
        mt.Text = u"Viga {0} · {1} · {2} · {3:.1f} m{4}".format(
            idx + 1,
            beam.get("id") or u"—",
            beam.get("type") or u"—",
            float(beam.get("len") or 0.0),
            extra,
        )
    else:
        mt.Text = u"Sin vigas en la selección"
    mt.Foreground = th.brush_fg_mid()
    mt.FontSize = 10.0
    try:
        mt.TextWrapping = TextWrapping.Wrap
    except Exception:
        pass
    meta.Child = mt
    sp.Children.Add(meta)

    # Variante A: dosificación siempre visible bajo la meta de viga.
    sp.Children.Add(build_dosificacion_segment(cv, session))

    block.Child = sp
    return block


def build_face_armadura_card(cv, face, beams, tramos, session):
    ensure_rail_state(cv)
    is_sup = face == u"sup"
    accent = u"#22d3ee" if is_sup else u"#fb7185"
    enabled = _card_enabled(cv, face)
    selected_ids, primary_id = _normalize_tramo_multi(cv, face, tramos)
    owner, tramo_beams = _owner_for_face(cv, tramos, selected_ids, beams, primary_id)
    if owner is not None:
        ensure_beam_layers(owner)

    shell = _card_shell(enabled, accent)
    root = StackPanel()

    pill_face = make_role_badge(u"SUP" if is_sup else u"INF", u"cent" if is_sup else u"ext")
    toggle = make_yesno_toggle(
        _win(cv),
        enabled,
        lambda v, f=face: _set_card_enabled(cv, f, v),
        compact=True,
    )
    root.Children.Add(
        _hdr_row(
            u"Armadura superior" if is_sup else u"Armadura inferior",
            [pill_face],
            toggle,
        )
    )

    body = StackPanel()
    body.IsEnabled = bool(enabled)
    if not enabled:
        body.Opacity = 0.72

    # Chip de contexto Tn (alzado) — una línea scannable, sin parágrafo largo.
    indices = _beam_indices_for_tramos(tramos, selected_ids)
    ctx = Border()
    ctx.Background = brush_hex(u"#0E1B32")
    ctx.BorderBrush = th.brush_border()
    ctx.BorderThickness = Thickness(1)
    ctx.Padding = Thickness(8, 5, 8, 5)
    ctx.Margin = Thickness(0, 0, 0, 10)
    _corner(ctx, 3)
    info = TextBlock()
    if primary_id is not None:
        labels = u" · ".join(u"V{0}".format(i + 1) for i in indices) if indices else u"—"
        if len(selected_ids) > 1:
            tids = u"+".join(
                u"T{0}".format(tid) for tid in sorted(selected_ids)
            )
            info.Text = u"{0}  ·  {1} viga(s)  ·  {2}".format(
                tids, len(indices), labels,
            )
        else:
            info.Text = u"Tn T{0}  ·  {1} viga(s)  ·  {2}".format(
                primary_id, len(indices), labels,
            )
    else:
        info.Text = u"Sin Tn · elige una banda en el alzado (Ctrl+clic multi)"
    info.Foreground = th.brush_fg_mid()
    info.FontSize = typo.LABEL_FONT_PX
    try:
        from System.Windows import TextWrapping

        info.TextWrapping = TextWrapping.Wrap
    except Exception:
        pass
    ctx.Child = info
    body.Children.Add(ctx)

    if owner is not None:
        body.Children.Add(
            _build_capas_block(cv, owner, tramo_beams, is_sup, accent, tramos=tramos)
        )
    else:
        empty = TextBlock()
        empty.Text = u"Selecciona un tramo en el alzado para editar capas."
        empty.Foreground = th.brush_fg_lo()
        empty.FontSize = 10.0
        empty.Margin = Thickness(0, 4, 0, 0)
        try:
            from System.Windows import TextWrapping

            empty.TextWrapping = TextWrapping.Wrap
        except Exception:
            pass
        body.Children.Add(empty)

    root.Children.Add(body)
    shell.Child = root
    return shell


def _apply_capas(cv, tramo_beams, side, n_capas):
    field = u"nCapasSup" if side == u"sup" else u"nCapasInf"
    n_val = max(CAPAS_MIN, min(CAPAS_MAX, int(n_capas)))
    all_beams = getattr(cv, "_last_beams", None) or list(tramo_beams or [])
    sync_layer_field_all_beams(all_beams, field, n_val)
    _redraw(cv)


def _session_for_cv(cv):
    session = getattr(cv, u"_last_session", None)
    if session is not None:
        return session
    try:
        from armado_vigas.revit.session import SESSION as _S

        return _S
    except Exception:
        return None


def _selected_tramos_list(cv, face, tramos):
    selected_ids, _primary = _normalize_tramo_multi(cv, face, tramos)
    want = set(selected_ids or [])
    return [t for t in (tramos or []) if t.get(u"id") in want]


def _apply_layer_field(
    cv, tramo_beams, owner, field, value, face=None, tramos=None, force_tramos=None,
):
    """n / ø por tramo seleccionado; capas siguen siendo globales al lote.

    ``force_tramos``: lista explícita (p. ej. panel de un Tn) en lugar de la
    multi-selección del canvas.
    """
    from armado_vigas.domain.tramo_armado import (
        apply_tramo_arm_to_beams,
        set_tramo_arm_field,
    )
    from armado_vigas.domain.layers import LAYER_CAPAS_FIELDS, LAYER_QTY_FIELDS

    all_beams = getattr(cv, "_last_beams", None) or list(tramo_beams or [])
    session = _session_for_cv(cv)
    face = face or (u"inf" if field.startswith(u"nInf") or field.startswith(u"diamInf") else u"sup")

    if field in LAYER_CAPAS_FIELDS or is_global_layer_sync_field(field):
        sync_layer_field_all_beams(all_beams, field, value)
        _redraw(cv)
        return

    # n / ø: config por tramo (force o multi-selección).
    selected = []
    if force_tramos is not None:
        selected = list(force_tramos or [])
    elif tramos is not None:
        selected = _selected_tramos_list(cv, face, tramos)

    if session is not None and selected:
        set_tramo_arm_field(session, face, selected, field, value)
        apply_tramo_arm_to_beams(selected, all_beams, face, field, value)
        try:
            from armado_vigas.domain.tramo_armado import merge_armado_onto_tramos

            # 1ª capa n: actualizar merge en ambas caras (n ligada SUP↔INF).
            faces_merge = (u"sup", u"inf") if field in (u"nSup", u"nInf") else (face,)
            for f in faces_merge:
                tr_list = (
                    getattr(session, u"tramos_inf" if f == u"inf" else u"tramos_sup", None)
                    or (tramos if f == face else None)
                    or []
                )
                merge_armado_onto_tramos(session, f, tr_list, all_beams)
        except Exception:
            pass
    elif owner is not None:
        if field in (u"nSup", u"nInf"):
            set_first_layer_bar_count(owner, value)
        else:
            owner[field] = value
        for b in tramo_beams or []:
            if b is owner:
                continue
            if field in (u"nSup", u"nInf"):
                set_first_layer_bar_count(b, value)
            elif field.startswith(u"diam") or field in LAYER_QTY_FIELDS:
                b[field] = value
        for b in tramo_beams or []:
            ensure_beam_layers(b)
            ensure_beam_confinement(b)
    _redraw(cv)


def _build_suple_arm_config_card(
    cv,
    accent,
    title,
    note,
    n_bars,
    diam_val,
    on_n,
    on_diam,
    on_quitar=None,
    quitar_label=u"Quitar",
):
    """
    Tarjeta anidada al estilo «1ª C.» de Armadura superior:
    título + chip ø×n · nota · Barras (n) | Diámetro.
    """
    card = Border()
    card.Background = brush_hex(u"#0c1824", 255)
    card.BorderBrush = brush_hex(accent, 140)
    card.BorderThickness = Thickness(1)
    card.Padding = Thickness(8, 8, 8, 8)
    card.Margin = Thickness(0, 0, 0, 0)
    card.HorizontalAlignment = HorizontalAlignment.Stretch
    _corner(card, 4.0)

    root = Grid()
    root.HorizontalAlignment = HorizontalAlignment.Stretch
    try:
        cd_acc = ColumnDefinition()
        cd_acc.Width = GridLength(3.0, GridUnitType.Pixel)
        cd_body = ColumnDefinition()
        cd_body.Width = GridLength(1.0, GridUnitType.Star)
        root.ColumnDefinitions.Add(cd_acc)
        root.ColumnDefinitions.Add(cd_body)
        accent_bar = Border()
        accent_bar.Background = brush_hex(accent, 200)
        accent_bar.Margin = Thickness(0, 0, 8, 0)
        accent_bar.HorizontalAlignment = HorizontalAlignment.Stretch
        accent_bar.VerticalAlignment = VerticalAlignment.Stretch
        _corner(accent_bar, 1.5)
        Grid.SetColumn(accent_bar, 0)
        root.Children.Add(accent_bar)
        body_col = 1
    except Exception:
        cd_body = ColumnDefinition()
        cd_body.Width = GridLength(1.0, GridUnitType.Star)
        root.ColumnDefinitions.Add(cd_body)
        body_col = 0

    body = StackPanel()
    body.HorizontalAlignment = HorizontalAlignment.Stretch

    hdr = DockPanel()
    hdr.LastChildFill = True
    hdr.Margin = Thickness(0, 0, 0, 8)

    if on_quitar is not None:
        btn_off = Button()
        btn_off.Content = quitar_label or u"Quitar"
        btn_off.Padding = Thickness(6, 3, 6, 3)
        btn_off.Margin = Thickness(6, 0, 0, 0)
        btn_off.FontSize = 9.0
        btn_off.Cursor = Cursors.Hand
        btn_off.BorderBrush = th.brush_border()
        btn_off.Background = th.brush_input()
        btn_off.Foreground = th.brush_fg_lo()
        btn_off.BorderThickness = Thickness(1)
        try:
            btn_off.Click += RoutedEventHandler(
                lambda sender, args: on_quitar()
            )
        except Exception:
            pass
        DockPanel.SetDock(btn_off, Dock.Right)
        hdr.Children.Add(btn_off)

    sum_chip = Border()
    sum_chip.Padding = Thickness(6, 1, 6, 1)
    sum_chip.Background = brush_hex(accent, 28)
    sum_chip.BorderBrush = brush_hex(accent, 90)
    sum_chip.BorderThickness = Thickness(1)
    _corner(sum_chip, 3.0)
    DockPanel.SetDock(sum_chip, Dock.Right)
    summary = TextBlock()
    summary.Text = _layer_summary_text(n_bars, diam_val)
    summary.Foreground = brush_hex(accent)
    summary.FontSize = 10.0
    summary.FontWeight = FontWeights.SemiBold
    summary.VerticalAlignment = VerticalAlignment.Center
    sum_chip.Child = summary
    hdr.Children.Add(sum_chip)

    title_tb = TextBlock()
    title_tb.Text = title or u"Suple"
    title_tb.FontSize = 11.0
    title_tb.FontWeight = FontWeights.Bold
    title_tb.Foreground = th.brush_fg_hi()
    title_tb.VerticalAlignment = VerticalAlignment.Center
    hdr.Children.Add(title_tb)
    body.Children.Add(hdr)

    if note:
        note_tb = TextBlock()
        note_tb.Text = note
        note_tb.Foreground = th.brush_fg_lo()
        note_tb.FontSize = typo.META_FONT_PX
        note_tb.Margin = Thickness(0, 0, 0, 6)
        try:
            note_tb.TextWrapping = TextWrapping.Wrap
        except Exception:
            pass
        body.Children.Add(note_tb)

    def _refresh_summary_n(v):
        try:
            on_n(v)
        finally:
            try:
                summary.Text = _layer_summary_text(clamp_bar_count(v), diam_val)
            except Exception:
                pass

    def _refresh_summary_d(v):
        try:
            on_diam(v)
        finally:
            try:
                summary.Text = _layer_summary_text(n_bars, int(v or 16))
            except Exception:
                pass

    n_cb = make_bar_count_combo(
        _win(cv),
        n_bars,
        lambda v: _refresh_summary_n(v),
        compact=True,
        stretch=True,
    )
    diam_cb = make_diam_combo(
        _win(cv),
        diam_val,
        _session_bar_opts(cv, diam_val),
        lambda v: _refresh_summary_d(v),
        compact=True,
        stretch=True,
    )

    row = Grid()
    row.HorizontalAlignment = HorizontalAlignment.Stretch
    cd_n = ColumnDefinition()
    cd_n.Width = GridLength(0.36, GridUnitType.Star)
    cd_g = ColumnDefinition()
    cd_g.Width = GridLength(10.0, GridUnitType.Pixel)
    cd_d = ColumnDefinition()
    cd_d.Width = GridLength(0.64, GridUnitType.Star)
    row.ColumnDefinitions.Add(cd_n)
    row.ColumnDefinitions.Add(cd_g)
    row.ColumnDefinitions.Add(cd_d)

    left = _field_stack(u"Barras (n)", n_cb)
    right = _field_stack(u"Diámetro", diam_cb)
    Grid.SetColumn(left, 0)
    Grid.SetColumn(right, 2)
    row.Children.Add(left)
    row.Children.Add(right)
    body.Children.Add(row)

    Grid.SetColumn(body, body_col)
    root.Children.Add(body)
    card.Child = root
    return card


def _suple_context_chip(text):
    """Chip de contexto scannable (mismo patrón que Tn en Armadura superior)."""
    ctx = Border()
    ctx.Background = brush_hex(u"#0E1B32")
    ctx.BorderBrush = th.brush_border()
    ctx.BorderThickness = Thickness(1)
    ctx.Padding = Thickness(8, 5, 8, 5)
    ctx.Margin = Thickness(0, 0, 0, 10)
    _corner(ctx, 3)
    info = TextBlock()
    info.Text = text or u""
    info.Foreground = th.brush_fg_mid()
    info.FontSize = typo.LABEL_FONT_PX
    try:
        info.TextWrapping = TextWrapping.Wrap
    except Exception:
        pass
    ctx.Child = info
    return ctx


def build_suple_face_card(cv, face, beams):
    ensure_rail_state(cv)
    if not _card_enabled(cv, face):
        return None
    is_sup = face == u"sup"
    accent = u"#a78bfa"
    beams = list(beams or [])
    n = len(beams)
    if n < 1:
        return None

    # SUP por apoyo; INF por viga seleccionada.
    sel = set(getattr(cv, "selected_beam_indices", None) or set())
    sel = {i for i in sel if 0 <= i < n}
    primary = getattr(cv, "selected_beam_idx", 0)
    if not sel:
        if 0 <= primary < n:
            sel = {primary}
        else:
            sel = {0}
            primary = 0
    if primary not in sel and sel:
        primary = min(sel)
    indices = sorted(sel)
    beam = beams[primary] if 0 <= primary < n else beams[0]

    session = getattr(cv, u"_last_session", None)
    try:
        from armado_vigas.revit.session import SESSION as _S

        if session is None:
            session = _S
    except Exception:
        pass

    n_capas = 1
    if is_sup:
        for b in beams:
            ensure_beam_suple_superior(b)
        ensure_session_suple_sup(session)
        try:
            sync_beams_suple_from_apoyo_set(session, beams)
        except Exception:
            pass
        on_ap = bool(
            getattr(session, u"suple_sup_apoyo_ids", None)
            and (session.suple_sup_apoyo_ids or set())
        )
        on_flag = any(beam_suple_sup_enabled(b) for b in beams)
        on = bool(on_ap or on_flag)
        n_suple, diam = 2, 16
        for b in beams:
            try:
                n_suple = int(b.get("nSupleSup") or n_suple)
                diam = int(b.get("diamSupleSup") or diam)
                break
            except Exception:
                continue
    else:
        ensure_beam_suple_inferior(beam)
        on = beam_suple_inf_enabled(beam)
        n_suple = int(beam.get("nSupleInf") or 2)
        diam = int(beam.get("diamSupleInf") or 16)
        n_capas = beam_n_capas_inf(beam)

    shell = _card_shell(on, accent if on else th.BORDER)
    shell.Margin = Thickness(0, 10, 0, 0)
    root = StackPanel()

    pill_face = make_role_badge(u"SUP" if is_sup else u"INF", u"cent" if is_sup else u"ext")
    pill_suple = make_role_badge(u"SUPLE SUP" if is_sup else u"SUPLE INF", u"suple")
    toggle = make_yesno_toggle(
        _win(cv),
        on,
        lambda v, b=beam, f=face: _toggle_suple_master(cv, b, f, v),
        compact=True,
    )
    root.Children.Add(
        _hdr_row(
            u"Suple superior" if is_sup else u"Suple inferior",
            [pill_face, pill_suple],
            toggle,
        )
    )

    body = StackPanel()
    body.IsEnabled = bool(on)
    if not on:
        body.Opacity = 0.72

    if is_sup:
        apoyos_all = getattr(session, u"apoyos", None) if session else []
        activos = active_apoyos_suple_sup(session, apoyos_all) if on else []
        activos_ids = set(a[u"id"] for a in activos)
        sel_ap = None
        try:
            sel_ap = getattr(session, u"selected_suple_apoyo_id", None)
            if sel_ap:
                sel_ap = unicode(sel_ap)
        except Exception:
            sel_ap = None
        if on and activos and (not sel_ap or sel_ap not in activos_ids):
            sel_ap = activos[0][u"id"]
            try:
                session.selected_suple_apoyo_id = sel_ap
            except Exception:
                pass

        n_a, d_a = 2, 16
        if sel_ap:
            try:
                n_a, d_a = get_apoyo_suple_sup_arm(session, sel_ap)
            except Exception:
                pass

        # Chip de contexto (estilo Tn de Armadura superior).
        if not on:
            ctx_txt = u"Suple SUP off · active el toggle para definir apoyos"
        elif not sel_ap or sel_ap not in activos_ids:
            ctx_txt = u"Sin apoyo · clic en columna/muro del alzado"
        else:
            n_act = len(activos)
            ctx_txt = (
                u"Apoyo {0}".format(sel_ap)
                if n_act <= 1
                else u"Apoyo {0}  ·  {1} activos".format(sel_ap, n_act)
            )
            adj = adjacent_beams_for_apoyo(
                getattr(session, u"domain_beams", None) or beams, sel_ap
            )
            if adj:
                parts = []
                for r in adj[:4]:
                    bid = (r.get(u"beam") or {}).get(u"id") or u"?"
                    span = int(r.get(u"span_mm") or 0)
                    parts.append(u"{0} {1}mm".format(bid, span))
                ctx_txt += u"  ·  L/3  ·  " + u" · ".join(parts)
        body.Children.Add(_suple_context_chip(ctx_txt))

        if not on:
            empty = TextBlock()
            empty.Text = u"Active Suple SUP para editar Barras y Diámetro por apoyo."
            empty.Foreground = th.brush_fg_lo()
            empty.FontSize = 10.0
            empty.Margin = Thickness(0, 4, 0, 0)
            try:
                empty.TextWrapping = TextWrapping.Wrap
            except Exception:
                pass
            body.Children.Add(empty)
        elif not sel_ap or sel_ap not in activos_ids:
            empty = TextBlock()
            empty.Text = (
                u"Selecciona un apoyo en el alzado para editar "
                u"Barras y Diámetro (mismo layout que capas de armadura)."
            )
            empty.Foreground = th.brush_fg_lo()
            empty.FontSize = 10.0
            empty.Margin = Thickness(0, 4, 0, 0)
            try:
                empty.TextWrapping = TextWrapping.Wrap
            except Exception:
                pass
            body.Children.Add(empty)
        else:
            def _quitar_apoyo(apoyo_id=sel_ap):
                try:
                    beams_all = list(
                        getattr(session, u"domain_beams", None) or beams or []
                    )
                    set_apoyo_suple_sup(
                        session,
                        apoyo_id,
                        False,
                        beams=beams_all,
                        apoyos=getattr(session, u"apoyos", None),
                    )
                    _status(cv, u"Suple SUP · {0} OFF".format(apoyo_id))
                    _redraw(cv)
                except Exception:
                    pass

            body.Children.Add(
                _build_suple_arm_config_card(
                    cv,
                    accent,
                    title=u"Suple",
                    note=(
                        u"L/3 en vigas adyacentes · n·ø propios de este apoyo · "
                        u"Ctrl+clic en alzado quita · fusión L/3+L/3 si Fin+Ini "
                        u"comparten nudo"
                    ),
                    n_bars=n_a,
                    diam_val=d_a,
                    on_n=lambda v, aid=sel_ap: _set_suple_sup_apoyo_field(
                        cv, aid, u"n", clamp_bar_count(v)
                    ),
                    on_diam=lambda v, aid=sel_ap: _set_suple_sup_apoyo_field(
                        cv, aid, u"diam", v
                    ),
                    on_quitar=_quitar_apoyo,
                    quitar_label=u"Quitar",
                )
            )
    else:
        # INF: chip de contexto + selector de vigas + tarjeta estilo capa.
        if len(indices) > 1:
            ctx_txt = u"{0}  ·  {1} viga(s)".format(
                u"+".join(u"V{0}".format(i + 1) for i in indices),
                len(indices),
            )
        else:
            ctx_txt = u"V{0}  ·  1 viga".format((primary or 0) + 1)
        if beam.get("id"):
            ctx_txt += u"  ·  {0}".format(beam.get("id"))
        if on:
            ctx_txt += u"  ·  capa {0} (n+1)".format(n_capas + 1)
        else:
            ctx_txt += u"  ·  suple off"
        body.Children.Add(_suple_context_chip(ctx_txt))

        vwrap = WrapPanel()
        vwrap.Margin = Thickness(0, 0, 0, 10)
        for bi, b in enumerate(beams):
            is_sel = bi in sel
            is_prim = bi == primary and is_sel
            btn = Button()
            btn.Content = u"V{0}".format(bi + 1)
            btn.Padding = Thickness(8, 3, 8, 3)
            btn.Margin = Thickness(0, 0, 6, 6)
            btn.FontSize = 10.0
            btn.FontWeight = FontWeights.Bold if is_prim else FontWeights.SemiBold
            btn.Cursor = Cursors.Hand
            if is_sel:
                btn.BorderBrush = brush_hex(accent)
                btn.Background = brush_hex(u"#a78bfa", 38)
                btn.Foreground = brush_hex(u"#c4b5fd")
            else:
                btn.BorderBrush = th.brush_border()
                btn.Background = th.brush_input()
                btn.Foreground = th.brush_fg_lo()
            btn.BorderThickness = Thickness(1)

            def _pick_beam(sender, args, idx=bi):
                try:
                    cv._handle_beam_select(idx, args, update_zone=False)
                except Exception:
                    pass

            try:
                btn.Click += RoutedEventHandler(_pick_beam)
            except Exception:
                pass
            vwrap.Children.Add(btn)
        body.Children.Add(vwrap)

        if not on:
            empty = TextBlock()
            empty.Text = u"Active Suple INF para editar Barras y Diámetro."
            empty.Foreground = th.brush_fg_lo()
            empty.FontSize = 10.0
            empty.Margin = Thickness(0, 4, 0, 0)
            try:
                empty.TextWrapping = TextWrapping.Wrap
            except Exception:
                pass
            body.Children.Add(empty)
        else:
            body.Children.Add(
                _build_suple_arm_config_card(
                    cv,
                    accent,
                    title=u"{0}ª C. (suple)".format(n_capas + 1),
                    note=(
                        u"Capa extra (n+1) · zona ~80 % central · "
                        u"independiente de bandas Tn · Ctrl multi-sel."
                    ),
                    n_bars=n_suple,
                    diam_val=diam,
                    on_n=lambda v, b=beam, f=face: _set_suple_field(
                        cv, b, f, u"n", clamp_bar_count(v)
                    ),
                    on_diam=lambda v, b=beam, f=face: _set_suple_field(
                        cv, b, f, u"diam", v
                    ),
                    on_quitar=None,
                )
            )

    root.Children.Add(body)
    shell.Child = root
    return shell



def _set_suple_sup_apoyo_field(cv, apoyo_id, kind, value):
    """n / ø por apoyo activo (suple SUP)."""
    session = getattr(cv, u"_last_session", None)
    try:
        from armado_vigas.revit.session import SESSION as _S

        if session is None:
            session = _S
    except Exception:
        pass
    if session is None or not apoyo_id:
        return
    ensure_session_suple_sup(session)
    try:
        session.selected_suple_apoyo_id = apoyo_id
    except Exception:
        pass
    n = diam = None
    if kind == u"n":
        n = clamp_bar_count(value)
    else:
        try:
            diam = int(value)
        except Exception:
            diam = value
    n2, d2 = set_apoyo_suple_sup_arm(session, apoyo_id, n=n, diam=diam)
    beams_all = list(getattr(session, u"domain_beams", None) or [])
    try:
        from armado_vigas.domain.suple_superior import apply_apoyo_arm_to_adjacent_beams

        apply_apoyo_arm_to_adjacent_beams(session, apoyo_id, beams=beams_all)
    except Exception:
        try:
            sync_beams_suple_arm_from_apoyos(session, beams_all)
        except Exception:
            pass
    _redraw(cv)
    _status(
        cv,
        u"Suple SUP · {0} · n={1} · ø{2}".format(apoyo_id, n2, d2),
    )


def _toggle_suple_master(cv, beam, face, on):
    if face == u"sup":
        session = getattr(cv, u"_last_session", None)
        try:
            from armado_vigas.revit.session import SESSION as _S

            if session is None:
                session = _S
        except Exception:
            pass
        beams_all = list(getattr(session, u"domain_beams", None) or [])
        if not beams_all and beam is not None:
            beams_all = [beam]
        if not on:
            clear_all_suple_sup_apoyos(session, beams=beams_all)
            for b in beams_all or []:
                try:
                    cv._set_beam_suple_field(b, u"supleSupEnabled", False)
                except Exception:
                    ensure_beam_suple_superior(b)
                    b[u"supleSupEnabled"] = False
                    b[u"supleSupStartEnabled"] = False
                    b[u"supleSupEndEnabled"] = False
            _redraw(cv)
            _status(cv, u"Suple SUP · off · sin apoyos activos")
            return
        # Master ON: habilita UI; tramos solo con apoyos seleccionados.
        ensure_session_suple_sup(session)
        for b in beams_all or []:
            ensure_beam_suple_superior(b)
            b[u"supleSupEnabled"] = True
        try:
            sync_beams_suple_from_apoyo_set(session, beams_all)
            sync_beams_suple_arm_from_apoyos(session, beams_all)
        except Exception:
            pass
        _redraw(cv)
        _status(
            cv,
            u"Suple SUP · on · añada apoyos o clic en columnas/muros del alzado",
        )
    else:
        # Suple INF: por viga (y multi-sel vía _targets_for_beam_edit).
        try:
            cv._set_beam_suple_field(beam, u"supleInfEnabled", bool(on))
        except Exception:
            if beam is not None:
                ensure_beam_suple_inferior(beam)
                beam[u"supleInfEnabled"] = bool(on)
            try:
                cv.invalidate_elev_cache()
            except Exception:
                pass
            _redraw(cv)
        _status(cv, u"Suple INF · {0}".format(u"on · ~80% central" if on else u"off"))


def _set_suple_field(cv, beam, face, kind, value):
    """INF: solo a la viga dada. SUP legacy (no usado por card nueva)."""
    if face == u"sup":
        # Compat: si llega sin apoyo, aplica al apoyo seleccionado.
        session = getattr(cv, u"_last_session", None)
        try:
            from armado_vigas.revit.session import SESSION as _S

            if session is None:
                session = _S
        except Exception:
            pass
        aid = getattr(session, u"selected_suple_apoyo_id", None) if session else None
        if aid:
            _set_suple_sup_apoyo_field(cv, aid, kind, value)
            return
        field = u"diamSupleSup" if kind == u"diam" else u"nSupleSup"
        targets = list(getattr(session, u"domain_beams", None) or [])
        if not targets and beam is not None:
            targets = [beam]
        for b in targets:
            if b is None:
                continue
            ensure_beam_suple_superior(b)
            b[field] = value
        _redraw(cv)
        _status(cv, u"Suple SUP · {0}={1}".format(field, value))
        return
    field = u"diamSupleInf" if kind == u"diam" else u"nSupleInf"
    try:
        cv._set_beam_suple_field(beam, field, value)
    except Exception:
        beam[field] = value
        _redraw(cv)
    _status(cv, u"Suple · {0}={1}".format(field, value))


def build_laterales_card(cv, session):
    ensure_rail_state(cv)
    enabled = _card_enabled(cv, u"lat")
    n_lat = session_n_laterales(session, 0)
    d0 = int(LATERALES_DIAM_DEFAULT)
    diam = int(getattr(session, "diamLaterales", d0) or d0) if session is not None else d0
    shell = _card_shell(enabled, u"#a78bfa")
    root = StackPanel()
    pill = make_role_badge(u"LAT", u"suple")
    toggle = make_yesno_toggle(
        _win(cv),
        enabled,
        lambda v: _set_card_enabled(cv, u"lat", v),
        compact=True,
    )
    root.Children.Add(_hdr_row(u"Laterales", [pill], toggle))
    root.Children.Add(_hint(u"Barras laterales del alma · preview de sección."))

    body = StackPanel()
    body.IsEnabled = bool(enabled)
    if not enabled:
        body.Opacity = 0.72

    # Mismo layout de controladores que capas SUP/INF: etiqueta arriba + combo
    # a ancho de columna (n ~36 %, ø ~64 %).
    n_max_ui = min(int(LATERALES_COUNT_MAX), 12)
    n_cb = make_int_combo(
        _win(cv),
        n_lat,
        LATERALES_COUNT_MIN,
        n_max_ui,
        lambda v: _set_laterales(cv, session, u"n", v),
        compact=True,
        stretch=True,
    )
    diam_cb = make_diam_combo(
        _win(cv),
        diam,
        _session_bar_opts(cv, diam),
        lambda v: _set_laterales(cv, session, u"diam", v),
        compact=True,
        stretch=True,
    )

    row = Grid()
    row.HorizontalAlignment = HorizontalAlignment.Stretch
    row.Margin = Thickness(0, 2, 0, 4)
    cd_n = ColumnDefinition()
    cd_n.Width = GridLength(0.36, GridUnitType.Star)
    cd_g = ColumnDefinition()
    cd_g.Width = GridLength(10.0, GridUnitType.Pixel)
    cd_d = ColumnDefinition()
    cd_d.Width = GridLength(0.64, GridUnitType.Star)
    row.ColumnDefinitions.Add(cd_n)
    row.ColumnDefinitions.Add(cd_g)
    row.ColumnDefinitions.Add(cd_d)

    left = _field_stack(u"Barras (n)", n_cb)
    right = _field_stack(u"Diámetro", diam_cb)
    Grid.SetColumn(left, 0)
    Grid.SetColumn(right, 2)
    row.Children.Add(left)
    row.Children.Add(right)
    body.Children.Add(row)

    foot = TextBlock()
    foot.Text = u"n y ø de laterales del alma · mismas opciones de diámetro que SUP/INF"
    foot.Foreground = th.brush_fg_lo()
    foot.FontSize = 9.0
    foot.Margin = Thickness(0, 4, 0, 0)
    try:
        from System.Windows import TextWrapping

        foot.TextWrapping = TextWrapping.Wrap
    except Exception:
        pass
    body.Children.Add(foot)

    root.Children.Add(body)
    shell.Child = root
    return shell


def _set_laterales(cv, session, kind, value):
    if session is None:
        return
    if kind == u"n":
        session.nLaterales = max(LATERALES_COUNT_MIN, min(LATERALES_COUNT_MAX, int(value)))
        _status(cv, u"Laterales · n={0}".format(session.nLaterales))
    else:
        session.diamLaterales = int(value)
        _status(cv, u"Laterales · Ø{0}".format(session.diamLaterales))
    _redraw(cv)


def build_confinamiento_card(cv, beam, session):
    ensure_rail_state(cv)
    enabled = _card_enabled(cv, u"conf")
    shell = _card_shell(enabled, u"#5bb8d4")
    root = StackPanel()
    pill = make_role_badge(u"CONF", u"confin")
    toggle = make_yesno_toggle(
        _win(cv),
        enabled,
        lambda v: _set_card_enabled(cv, u"conf", v),
        compact=True,
    )
    root.Children.Add(_hdr_row(u"Confinamiento / estribos", [pill], toggle))
    n_sel = 1
    try:
        n_sel = max(1, len(getattr(cv, u"selected_beam_indices", None) or set()) or 1)
    except Exception:
        n_sel = 1
    root.Children.Add(
        _hint(
            u"Ctrl+clic en alzado: multi-viga · ø/@ y dibujo E/T se aplican al lote "
            u"({0} viga{1}). Sección: preview de la primaria.".format(
                n_sel, u"s" if n_sel != 1 else u"",
            )
        )
    )

    body = StackPanel()
    body.IsEnabled = bool(enabled)
    if not enabled:
        body.Opacity = 0.72

    if beam is None:
        body.Children.Add(_hint(u"Seleccione una viga."))
        root.Children.Add(body)
        shell.Child = root
        return shell

    ensure_beam_layers(beam)
    ensure_beam_confinement(beam)
    ensure_beam_stirrup_zone_mode(beam)
    plan = compute_stirrup_zones(beam)
    h_mm = section_height_mm(beam.get("type"))
    l_ext = int(round(plan.get("L_ext_each") or (2.0 * h_mm)))
    l_cent = int(round(plan.get("L_cent") or 0))

    # Zonificación: solo Extremos+Centro | Único (auto se aplica y persiste).
    mode_labels = stirrup_zone_mode_labels()
    mode_cur = stirrup_zone_mode_label(beam.get("estZonasMode"))

    def _on_zonas_mode(v, b=beam):
        mode = normalize_stirrup_zone_mode(v)
        if mode == STIRRUP_ZONE_AUTO:
            mode = recommended_stirrup_zone_mode(b)
        _set_beam_simple(cv, b, u"estZonasMode", mode)

    body.Children.Add(
        _field_stack(
            u"Zonificación estribos",
            make_string_combo(
                _win(cv),
                mode_labels,
                mode_cur,
                _on_zonas_mode,
                compact=True,
                stretch=True,
            ),
        )
    )

    auto_kind = plan.get("autoMode") or u"unico"
    eff_kind = plan.get("effectiveMode") or auto_kind
    auto_txt = u"Ext+Cent" if auto_kind == u"ext_cent" else u"Único"
    eff_txt = u"Ext+Cent" if eff_kind == u"ext_cent" else u"Único"
    if plan.get("isOverride"):
        cond_txt = u"{0} · sugerido {1}".format(eff_txt, auto_txt)
    else:
        cond_txt = eff_txt

    meta = TextBlock()
    try:
        sel_idxs = sorted(getattr(cv, u"selected_beam_indices", None) or set())
        beams_all = list(getattr(cv, u"_last_beams", None) or [])
    except Exception:
        sel_idxs = []
        beams_all = []
    if len(sel_idxs) > 1 and beams_all:
        labels = []
        for ii in sel_idxs:
            try:
                if 0 <= int(ii) < len(beams_all):
                    labels.append(
                        unicode(beams_all[int(ii)].get(u"id") or u"V{0}".format(int(ii) + 1))
                    )
            except Exception:
                continue
        lote_txt = u"{0} vigas · {1}".format(
            len(sel_idxs), u" + ".join(labels[:6]) + (u"…" if len(labels) > 6 else u""),
        )
    else:
        lote_txt = unicode(beam.get("id") or u"—")
    meta.Text = u"{0} · {1} · h={2} mm · Lext {3} mm ×2 · Lcent {4} mm".format(
        lote_txt, cond_txt, h_mm, l_ext, l_cent
    )
    meta.Foreground = th.brush_fg_mid()
    meta.FontSize = 10.0
    meta.Margin = Thickness(0, 4, 0, 10)
    try:
        from System.Windows import TextWrapping

        meta.TextWrapping = TextWrapping.Wrap
    except Exception:
        pass
    body.Children.Add(meta)

    sz = getattr(cv, "selected_stirrup_zone", None) or {}
    role = sz.get("role") or u"cent"

    def _zone_block(title, zone_role, accent, diam_field, sp_field, spacing_default, hint):
        sel = role == zone_role
        box = Border()
        box.Margin = Thickness(0, 0, 0, 8)
        box.Padding = Thickness(8)
        box.BorderBrush = brush_hex(accent) if sel else th.brush_border()
        box.BorderThickness = Thickness(1.5 if sel else 1)
        box.Background = brush_hex(u"#0E1B32") if sel else th.brush_input()
        _corner(box, 4)
        sp = StackPanel()
        hdr = Button()
        hdr.Content = u"{0} · {1}".format(title, hint)
        hdr.Padding = Thickness(4, 4, 4, 4)
        hdr.FontSize = 11.0
        hdr.FontWeight = FontWeights.SemiBold
        hdr.Cursor = Cursors.Hand
        hdr.Background = brush_hex(u"#000000", 0)
        hdr.BorderThickness = Thickness(0)
        hdr.Foreground = brush_hex(accent) if sel else th.brush_fg_mid()
        hdr.HorizontalContentAlignment = HorizontalAlignment.Left

        def _sel(sender, args, r=zone_role):
            idx = getattr(cv, "selected_beam_idx", 0)
            cv.selected_stirrup_zone = {u"idx": idx, u"role": r}
            cb = getattr(cv, "_cb", None) or {}
            cb.get("on_select_stirrup_zone", lambda _i, _r: None)(idx, r)
            _redraw(cv)

        try:
            hdr.Click += RoutedEventHandler(_sel)
        except Exception:
            pass
        sp.Children.Add(hdr)

        # Mismo patrón SUP/INF: etiqueta arriba + control a ancho de columna.
        diam_val = int(beam.get(diam_field) or 10)
        diam_cb = make_diam_combo(
            _win(cv),
            diam_val,
            ESTRIBO_DIAM_OPTS,
            lambda v, f=diam_field: _set_beam_simple(cv, beam, f, v),
            compact=True,
            stretch=True,
        )
        sp_val = int(beam.get(sp_field) or spacing_default)
        sp_tb = make_spacing_input(
            _win(cv),
            sp_val,
            lambda v, f=sp_field: _set_beam_simple(cv, beam, f, v),
            compact=True,
            stretch=True,
        )

        row = Grid()
        row.HorizontalAlignment = HorizontalAlignment.Stretch
        row.Margin = Thickness(0, 6, 0, 0)
        cd_d = ColumnDefinition()
        cd_d.Width = GridLength(0.50, GridUnitType.Star)
        cd_g = ColumnDefinition()
        cd_g.Width = GridLength(10.0, GridUnitType.Pixel)
        cd_s = ColumnDefinition()
        cd_s.Width = GridLength(0.50, GridUnitType.Star)
        row.ColumnDefinitions.Add(cd_d)
        row.ColumnDefinitions.Add(cd_g)
        row.ColumnDefinitions.Add(cd_s)

        left = _field_stack(u"Diámetro (ø)", diam_cb)
        right = _field_stack(u"Espaciado (@)", sp_tb)
        Grid.SetColumn(left, 0)
        Grid.SetColumn(right, 2)
        row.Children.Add(left)
        row.Children.Add(right)
        sp.Children.Add(row)
        box.Child = sp
        return box

    if plan.get("mode") == u"single":
        r = u"uni" if plan.get("singleKind") == u"merge" else u"cent"
        body.Children.Add(
            _zone_block(
                u"Único" if r == u"uni" else u"Cent",
                r,
                u"#34d399",
                u"estCentDiam",
                u"estCentSpacing",
                _SP_CENT,
                u"L {0} mm".format(l_cent or l_ext),
            )
        )
    else:
        body.Children.Add(
            _zone_block(
                u"Cent",
                u"cent",
                u"#34d399",
                u"estCentDiam",
                u"estCentSpacing",
                _SP_CENT,
                u"L {0} mm".format(l_cent),
            )
        )
        body.Children.Add(
            _zone_block(
                u"Ext",
                u"ext",
                u"#fbbf24",
                u"estExtDiam",
                u"estExtSpacing",
                _SP_EXT,
                u"L {0} mm ×2".format(l_ext),
            )
        )

    # Diagrama E/T: editar en preview de sección o Zoom (tools no van en el rail).
    body.Children.Add(_section_label(u"Diagrama E"))
    d = get_conf_draft(beam) if beam is not None else {}
    defined = is_conf_draft_defined(beam) if beam is not None else False
    draft_tb = TextBlock()
    if defined:
        draft_tb.Text = conf_draft_label(d) or u"Definido"
        draft_tb.Foreground = brush_hex(u"#5bb8d4")
    else:
        draft_tb.Text = u"Sin estribos/trabas · preview o Zoom para dibujar"
        draft_tb.Foreground = th.brush_fg_mid()
    draft_tb.FontSize = 10.0
    draft_tb.Margin = Thickness(0, 2, 0, 2)
    try:
        draft_tb.TextWrapping = TextWrapping.Wrap
    except Exception:
        pass
    body.Children.Add(draft_tb)
    body.Children.Add(
        _hint(u"Dibujo 135° en sección (arriba) o Zoom · clic 1 ancla · clic 2 cierra.")
    )

    root.Children.Add(body)
    shell.Child = root
    return shell


def _set_beam_simple(cv, beam, field, value):
    try:
        cv._set_beam_field(beam, field, value)
    except Exception:
        beam[field] = value
        if field == u"estConfin":
            ensure_beam_confinement(beam)
        elif field == u"estZonasMode":
            ensure_beam_stirrup_zone_mode(beam)
        _redraw(cv)


def populate_section_rail(cv, beams, session):
    """Rellena PnlSectionCtrls con pestañas + tarjetas del mockup (debajo del preview)."""
    ensure_rail_state(cv)
    pnl = getattr(cv, "_pnl_section_ctrls", None)
    if pnl is None:
        return
    pnl.Children.Clear()

    if not beams:
        empty = TextBlock()
        empty.Text = u"Seleccione vigas de hormigón para previsualizar."
        empty.Foreground = th.brush_fg_lo()
        empty.FontSize = 10.0
        pnl.Children.Add(empty)
        return

    idx = getattr(cv, "selected_beam_idx", 0)
    if not (0 <= idx < len(beams)):
        idx = 0
        cv.selected_beam_idx = 0
        cv.selected_beam_indices = {0}
    beam = beams[idx]

    tramos_sup = list(getattr(session, "tramos_sup", None) or []) if session else []
    tramos_inf = list(getattr(session, "tramos_inf", None) or []) if session else []
    _normalize_tramo_multi(cv, u"sup", tramos_sup)
    _normalize_tramo_multi(cv, u"inf", tramos_inf)

    card = getattr(cv, "rail_card", u"sup") or u"sup"
    pnl.Children.Add(build_config_viga_header(cv, beam, beams, session))
    pnl.Children.Add(build_rail_tabs(cv))

    if card == u"sup":
        pnl.Children.Add(build_face_armadura_card(cv, u"sup", beams, tramos_sup, session))
        suple = build_suple_face_card(cv, u"sup", beams)
        if suple is not None:
            pnl.Children.Add(suple)
    elif card == u"inf":
        pnl.Children.Add(build_face_armadura_card(cv, u"inf", beams, tramos_inf, session))
        suple = build_suple_face_card(cv, u"inf", beams)
        if suple is not None:
            pnl.Children.Add(suple)
    elif card == u"lat":
        pnl.Children.Add(build_laterales_card(cv, session))
    else:
        # CONF: siempre la viga seleccionada en alzado (config por viga, no owner Tn).
        conf_beam = beam
        try:
            idx = int(getattr(cv, "selected_beam_idx", 0) or 0)
            if 0 <= idx < len(beams):
                conf_beam = beams[idx]
        except Exception:
            conf_beam = beam
        pnl.Children.Add(build_confinamiento_card(cv, conf_beam or beam, session))

    footnote = TextBlock()
    footnote.Text = u"Vista previa UI · motor de colocación de armadura no incluido."
    footnote.Foreground = th.brush_fg_lo()
    footnote.FontSize = 9.0
    footnote.Margin = Thickness(0, 12, 0, 0)
    try:
        from System.Windows import TextWrapping

        footnote.TextWrapping = TextWrapping.Wrap
    except Exception:
        pass
    pnl.Children.Add(footnote)
