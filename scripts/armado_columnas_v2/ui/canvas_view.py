# -*- coding: utf-8 -*-
"""Canvas de elevación fiel a la vista activa + rail de sección."""

from __future__ import division

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import (
    CornerRadius,
    FontWeights,
    Thickness,
    TextWrapping,
)
from System.Windows.Controls import Border, Canvas, TextBlock
from System.Windows.Input import Cursors, Keyboard, ModifierKeys, MouseButtonEventHandler
from System.Windows.Media import DoubleCollection
from System.Windows.Shapes import Line, Rectangle

from armado_columnas_v2.ui import layout as lay
from armado_columnas_v2.ui.section_preview import draw_section_preview, section_meta_lines
from armado_columnas_v2.ui.wpf_brushes import brush_hex
from armado_columnas_v2.session import ColumnArmadoSession

# Estilos por kind: (stroke, fill, sel_stroke, label_prefix)
_KIND_STYLE = {
    u"column": (u"#5bb8d4", u"#0f2a38", u"#6eb8c8", u"C"),
    u"foundation": (u"#d4a574", u"#2a1f14", u"#c4a070", u"F"),
    u"beam": (u"#7d9b8a", u"#142018", u"#8fad9a", u"V"),
    u"floor": (u"#8b9cb3", u"#121a24", u"#a0b0c4", u"L"),
}
# Fondo → frente
_DRAW_KIND_ORDER = (u"floor", u"foundation", u"beam", u"column")


class ArmadoColumnasCanvasView(object):
    def __init__(self, win, callbacks):
        """
        callbacks:
          on_status(msg), on_redraw(), on_select_column(member_id, multi)
        """
        self._win = win
        self._cb = callbacks or {}
        self._host = win.FindName(u"PnlCanvasHost") if win is not None else None
        self._scr = win.FindName(u"ScrCanvas") if win is not None else None
        self._preview = win.FindName(u"CnvSectionPreview") if win is not None else None
        self._txt_meta = win.FindName(u"TxtSectionMeta") if win is not None else None
        self._txt_lote = win.FindName(u"TxtLoteSummary") if win is not None else None
        self._txt_levels = win.FindName(u"TxtLevelsSummary") if win is not None else None
        self._rail_ctrls = win.FindName(u"PnlSectionCtrls") if win is not None else None

    def _viewport_w(self):
        try:
            if self._scr is not None and self._scr.ActualWidth > 40:
                return float(self._scr.ActualWidth) - 8.0
        except Exception:
            pass
        try:
            if self._win is not None and self._win.ActualWidth > 200:
                return max(
                    320.0,
                    float(self._win.ActualWidth) - lay.SECTION_RAIL_WIDTH_PX - 80.0,
                )
        except Exception:
            pass
        return 720.0

    def _viewport_h(self):
        try:
            if self._scr is not None and self._scr.ActualHeight > 40:
                return float(self._scr.ActualHeight) - 8.0
        except Exception:
            pass
        return 520.0

    def redraw(self, session):
        if self._host is None:
            return
        members = list(session.domain_members or session.domain_columns or [])
        selected = session.selected_ids or set()
        if hasattr(session, "preview_member"):
            preview = session.preview_member()
        else:
            preview = session.preview_column()

        self._update_summaries(session, members, selected)
        self._draw_elevation(members, selected, session)
        self._draw_section(preview)
        self._draw_rail_placeholder(preview, session)

    def _update_summaries(self, session, members, selected):
        n_c = len(getattr(session, "domain_columns", None) or [])
        n_f = len(getattr(session, "domain_foundations", None) or [])
        n_v = len(getattr(session, "domain_beams", None) or [])
        n_l = len(getattr(session, "domain_floors", None) or [])
        if not (n_c or n_f or n_v or n_l) and members:
            for m in members:
                k = m.get("kind")
                if k == u"column":
                    n_c += 1
                elif k == u"foundation":
                    n_f += 1
                elif k == u"beam":
                    n_v += 1
                elif k == u"floor":
                    n_l += 1
        ns = len(selected)
        if self._txt_lote is not None:
            vista = getattr(session, "view_name", None) or u"vista activa"
            self._txt_lote.Text = (
                u"{0} col · {1} fund · {2} viga · {3} losa · {4} sel · «{5}»".format(
                    n_c, n_f, n_v, n_l, ns, vista
                )
            )
        if self._txt_levels is not None:
            if not members:
                self._txt_levels.Text = u"Sin elementos proyectados."
            else:
                self._txt_levels.Text = (
                    u"Layout a escala de vista · Right→X · Up→Y · "
                    u"contexto (viga/losa/fund) recortado ±2000 mm del centroide de columnas."
                )

    def _draw_section(self, member):
        meta = draw_section_preview(self._preview, member)
        if self._txt_meta is not None:
            self._txt_meta.Text = meta or section_meta_lines(member)

    def _draw_rail_placeholder(self, member, session):
        if self._rail_ctrls is None:
            return
        self._rail_ctrls.Children.Clear()
        tip = TextBlock()
        tip.Text = (
            u"Controles de dosificación del rail (Ø long., n° barras/cara, "
            u"estribos, empalmes) se definirán en la siguiente iteración."
        )
        tip.Foreground = brush_hex(u"#64748b")
        tip.FontSize = 10.0
        tip.TextWrapping = TextWrapping.Wrap
        tip.Margin = Thickness(0, 4, 0, 0)
        self._rail_ctrls.Children.Add(tip)

        if member:
            chip = Border()
            chip.Background = brush_hex(u"#0d2430")
            chip.BorderBrush = brush_hex(u"#21465C")
            chip.BorderThickness = Thickness(1)
            chip.CornerRadius = CornerRadius(4)
            chip.Padding = Thickness(8, 6, 8, 6)
            chip.Margin = Thickness(0, 10, 0, 0)
            kind = member.get("kind") or u"column"
            rol = ColumnArmadoSession.kind_label_es(kind).capitalize()
            tb = TextBlock()
            tb.Text = u"{0}: {1}".format(
                rol, member.get("label") or member.get("id") or u"—"
            )
            tb.Foreground = brush_hex(u"#7eb8d0")
            tb.FontSize = 10.0
            tb.FontWeight = FontWeights.SemiBold
            chip.Child = tb
            self._rail_ctrls.Children.Add(chip)

    def _draw_elevation(self, members, selected, session):
        self._host.Children.Clear()
        if not members:
            empty = TextBlock()
            empty.Text = (
                u"Sin elementos. Ejecute la herramienta y seleccione "
                u"columnas, fundaciones, vigas y losas en la vista activa."
            )
            empty.Foreground = brush_hex(u"#64748b")
            empty.Margin = Thickness(16)
            empty.FontSize = 12.0
            empty.TextWrapping = TextWrapping.Wrap
            self._host.Children.Add(empty)
            return

        vw = self._viewport_w()
        vh = self._viewport_h()

        # Recorte de contexto: ±2000 mm del centroide de columnas (seleccionadas)
        display_members, clip_meta = lay.clip_context_members_horizontal(
            members, selected_ids=selected,
        )
        if not display_members:
            display_members = list(members)

        has_model = all(
            m.get("uMin") is not None and m.get("vMin") is not None
            for m in display_members
        )
        if has_model:
            meta = lay.compute_elevation_layout(display_members, vw, vh)
        else:
            meta = lay.compute_column_slots(display_members, vw)

        layouts = meta.get("layouts") or []
        content_w = float(meta.get("contentWidthPx") or vw)
        content_h = float(meta.get("contentHeightPx") or vh)

        by_id = {}
        for m in display_members:
            by_id[m.get("id")] = m

        root = Canvas()
        root.Width = content_w
        root.Height = content_h
        root.Background = brush_hex(u"#071018", 0)
        self._host.Children.Add(root)

        hdr = TextBlock()
        scale = meta.get("scalePxPerFt")
        if meta.get("modelPositions") and scale:
            if clip_meta.get("applied"):
                hdr.Text = (
                    u"Elevación · escala vista ({0:.1f} px/ft) · "
                    u"contexto recortado ±{1:.0f} mm".format(
                        float(scale), float(clip_meta.get("halfMm") or 2000)
                    )
                )
            else:
                hdr.Text = u"Elevación · escala vista ({0:.1f} px/ft)".format(
                    float(scale)
                )
        else:
            hdr.Text = u"Elevación"
        hdr.Foreground = brush_hex(u"#64748b")
        hdr.FontSize = 9.0
        hdr.FontWeight = FontWeights.SemiBold
        Canvas.SetLeft(hdr, 12.0)
        Canvas.SetTop(hdr, 6.0)
        root.Children.Add(hdr)

        if meta.get("modelPositions"):
            self._draw_extent_frame(root, meta)

        kind_rank = dict((k, i) for i, k in enumerate(_DRAW_KIND_ORDER))
        ordered = sorted(
            layouts,
            key=lambda lay_i: (
                kind_rank.get(
                    (by_id.get(lay_i.get("id")) or {}).get("kind"), 50
                ),
                float(lay_i.get("topPx") or 0),
                float(lay_i.get("leftPx") or 0),
            ),
        )
        for layout in ordered:
            mid = layout.get("id")
            member = by_id.get(mid)
            if member is None:
                continue
            self._draw_member(root, member, layout, mid in selected)

        if layouts and meta.get("modelPositions"):
            elev_top = float(meta.get("elevTopPx") or 0)
            elev_bot = float(meta.get("elevBottomPx") or elev_top)
            axis_x = max(8.0, float(meta.get("originXPx") or 0) - 18.0)
            axis = Line()
            axis.X1 = axis_x
            axis.Y1 = elev_top
            axis.X2 = axis_x
            axis.Y2 = elev_bot
            axis.Stroke = brush_hex(u"#334155", 180)
            axis.StrokeThickness = 1.0
            root.Children.Add(axis)

            try:
                h_m = (
                    float(meta["modelVMax"]) - float(meta["modelVMin"])
                ) * 0.3048
                tb = TextBlock()
                tb.Text = u"{0:.2f} m".format(h_m)
                tb.Foreground = brush_hex(u"#475569")
                tb.FontSize = 8.0
                Canvas.SetLeft(tb, max(2.0, axis_x - 28.0))
                Canvas.SetTop(tb, (elev_top + elev_bot) * 0.5 - 6.0)
                root.Children.Add(tb)
            except Exception:
                pass

    def _draw_extent_frame(self, root, meta):
        try:
            x0 = float(meta["originXPx"])
            y0 = float(meta["elevTopPx"])
            x1 = x0 + (float(meta["modelUMax"]) - float(meta["modelUMin"])) * float(
                meta["scalePxPerFt"]
            )
            y1 = float(meta["elevBottomPx"])
        except Exception:
            return
        frame = Rectangle()
        frame.Width = max(1.0, x1 - x0)
        frame.Height = max(1.0, y1 - y0)
        Canvas.SetLeft(frame, x0)
        Canvas.SetTop(frame, y0)
        frame.Stroke = brush_hex(u"#1e3344", 120)
        frame.StrokeThickness = 1.0
        frame.StrokeDashArray = DoubleCollection()
        frame.StrokeDashArray.Add(3)
        frame.StrokeDashArray.Add(3)
        frame.Fill = brush_hex(u"#0a1620", 40)
        frame.IsHitTestVisible = False
        root.Children.Add(frame)

    def _draw_member(self, root, member, layout, is_selected):
        left = float(layout["leftPx"])
        top = float(layout["topPx"])
        width = float(layout["widthPx"])
        height = float(layout["heightPx"])
        kind = member.get("kind") or u"column"
        selectable = kind == u"column"
        # Solo columnas reciben resalte de selección
        show_sel = bool(is_selected and selectable)
        style = _KIND_STYLE.get(kind) or _KIND_STYLE[u"column"]
        stroke_base, fill_base, stroke_sel, prefix = style
        mid = member.get("id")

        if selectable:
            pad = 2.0
            hit = Border()
            hit.Width = width + pad * 2.0
            hit.Height = height + pad * 2.0
            hit.Background = brush_hex(u"#0a1620", 1)
            hit.Cursor = Cursors.Hand
            hit.BorderThickness = Thickness(0)
            hit.CornerRadius = CornerRadius(0)
            Canvas.SetLeft(hit, left - pad)
            Canvas.SetTop(hit, top - pad)
            root.Children.Add(hit)

            def _on_click(sender, args, _id=mid):
                multi = False
                try:
                    multi = Keyboard.Modifiers == ModifierKeys.Control
                except Exception:
                    multi = False
                cb = self._cb.get("on_select_column")
                if cb:
                    cb(_id, multi)

            try:
                hit.MouseLeftButtonUp += MouseButtonEventHandler(_on_click)
            except Exception:
                pass

        # Resalte sutil en el cuerpo (sin marco extra)
        if show_sel:
            fill = brush_hex(u"#1a3a4c", 235)
            stroke_brush = brush_hex(stroke_base, 230)
            stroke_th = 1.35
        else:
            stroke_brush = brush_hex(stroke_base, 200)
            stroke_th = 1.2
            fill = brush_hex(fill_base, 210)

        body = Rectangle()
        body.Width = width
        body.Height = height
        Canvas.SetLeft(body, left)
        Canvas.SetTop(body, top)
        body.RadiusX = 0
        body.RadiusY = 0
        body.Stroke = stroke_brush
        body.StrokeThickness = stroke_th
        body.Fill = fill
        body.IsHitTestVisible = False
        root.Children.Add(body)

        lbl = TextBlock()
        lbl.Text = member.get("label") or u"{0}?".format(prefix)
        lbl.Foreground = brush_hex(u"#b8d0dc" if show_sel else u"#95b8cc")
        lbl.FontSize = 10.0
        lbl.FontWeight = FontWeights.Normal
        lbl.IsHitTestVisible = False
        Canvas.SetLeft(lbl, left)
        Canvas.SetTop(lbl, top + height + 3.0)
        root.Children.Add(lbl)

        sub = TextBlock()
        try:
            elev_w_cm = float(member.get("spanU_ft") or 0) * 30.48
            elev_h_cm = float(member.get("spanV_ft") or 0) * 30.48
            if kind == u"floor":
                h_txt = u"{0:.0f}×{1:.0f} cm vista".format(elev_w_cm, elev_h_cm)
            elif kind == u"beam":
                h_txt = u"{0:.0f}×{1:.0f} cm vista".format(elev_w_cm, elev_h_cm)
            elif kind == u"foundation":
                h_txt = u"{0:.0f}×{1:.0f} cm vista".format(elev_w_cm, elev_h_cm)
            else:
                h_txt = u"{0:.0f} cm · {1:.2f} m".format(
                    float(member.get("widthCm") or 0),
                    float(member.get("heightM") or 0),
                )
        except Exception:
            h_txt = member.get("typeName") or u""
        sub.Text = h_txt
        sub.Foreground = brush_hex(u"#64748b")
        sub.FontSize = 8.5
        sub.IsHitTestVisible = False
        Canvas.SetLeft(sub, left)
        Canvas.SetTop(sub, top + height + 16.0)
        root.Children.Add(sub)
