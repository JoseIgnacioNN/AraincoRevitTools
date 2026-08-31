# -*- coding: utf-8 -*-
"""Batch de geometría ortogonal (alzado) — un StreamGeometry/Path por estilo.

Pre-procesa segmentos en Python puro y cruza a .NET/WPF una sola vez por
grupo (stroke/fill, grosor, dash, z-index). Elimina cientos de Line/Rectangle
por redibujado sin alterar las coordenadas ya calculadas por el caller.
"""

from __future__ import print_function

import clr

clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System.Windows import Point
from System.Windows.Controls import Canvas
from System.Windows.Media import (
    DoubleCollection,
    EdgeMode,
    PenLineCap,
    PenLineJoin,
    RenderOptions,
    StreamGeometry,
)
from System.Windows.Shapes import Path

from armado_vigas.ui.net_ui import freeze_freezable


def _dash_key(dash):
    if not dash:
        return None
    try:
        return tuple(float(x) for x in dash)
    except Exception:
        return None


def _make_dash_collection(dash):
    if not dash:
        return None
    try:
        dc = DoubleCollection()
        for x in dash:
            dc.Add(float(x))
        freeze_freezable(dc)
        return dc
    except Exception:
        try:
            return DoubleCollection(list(dash))
        except Exception:
            return None


class ElevGeomBatch(object):
    """Acumula líneas y rects; ``flush(canvas)`` emite Path consolidados.

    Keys de agrupación =
      (id(brush), thickness, dash_key, zindex)  para strokes
      (id(brush), zindex)                         para fills
    ``id(brush)`` es estable con brushes Freezados cacheados (``brush_hex``).
    """

    def __init__(self):
        # style_key -> (stroke, thickness, dash, zindex, segs)
        # segs: list of (x1, y1, x2, y2)
        self._lines = {}
        # style_key -> (fill, zindex, rects)
        # rects: list of (x, y, w, h)
        self._fill_rects = {}
        # style_key -> (stroke, thickness, dash, zindex, rects)
        self._stroke_rects = {}
        self._count = 0

    @property
    def primitive_count(self):
        return int(self._count)

    def add_line(self, x1, y1, x2, y2, stroke, thickness=0.9, dash=None, zindex=0):
        if stroke is None:
            return
        try:
            t = float(thickness)
        except Exception:
            t = 0.9
        z = int(zindex or 0)
        dk = _dash_key(dash)
        key = (id(stroke), t, dk, z)
        bucket = self._lines.get(key)
        if bucket is None:
            self._lines[key] = (stroke, t, dash, z, [])
            bucket = self._lines[key]
        bucket[4].append((float(x1), float(y1), float(x2), float(y2)))
        self._count += 1

    def add_rect(
        self,
        x,
        y,
        w,
        h,
        fill=None,
        stroke=None,
        thickness=0.9,
        dash=None,
        zindex=0,
    ):
        try:
            xf, yf, wf, hf = float(x), float(y), float(w), float(h)
        except Exception:
            return
        if wf <= 0.0 or hf <= 0.0:
            return
        z = int(zindex or 0)
        if fill is not None:
            key = (id(fill), z)
            bucket = self._fill_rects.get(key)
            if bucket is None:
                self._fill_rects[key] = (fill, z, [])
                bucket = self._fill_rects[key]
            bucket[2].append((xf, yf, wf, hf))
            self._count += 1
        if stroke is not None:
            try:
                t = float(thickness)
            except Exception:
                t = 0.9
            dk = _dash_key(dash)
            key = (id(stroke), t, dk, z)
            bucket = self._stroke_rects.get(key)
            if bucket is None:
                self._stroke_rects[key] = (stroke, t, dash, z, [])
                bucket = self._stroke_rects[key]
            bucket[4].append((xf, yf, wf, hf))
            self._count += 1

    def flush(self, canvas):
        """Materializa un Path por grupo de estilo sobre ``canvas``."""
        if canvas is None:
            return 0
        n_paths = 0
        # Fills primero (suelen ir detrás de strokes del mismo z).
        for _key, (fill, z, rects) in self._fill_rects.items():
            if not rects:
                continue
            path = _stream_fill_rects(rects, fill, z)
            if path is not None:
                canvas.Children.Add(path)
                n_paths += 1
        for _key, (stroke, t, dash, z, segs) in self._lines.items():
            if not segs:
                continue
            path = _stream_open_lines(segs, stroke, t, dash, z)
            if path is not None:
                canvas.Children.Add(path)
                n_paths += 1
        for _key, (stroke, t, dash, z, rects) in self._stroke_rects.items():
            if not rects:
                continue
            path = _stream_stroke_rects(rects, stroke, t, dash, z)
            if path is not None:
                canvas.Children.Add(path)
                n_paths += 1
        self.clear()
        return n_paths

    def clear(self):
        self._lines.clear()
        self._fill_rects.clear()
        self._stroke_rects.clear()
        self._count = 0


def _stream_open_lines(segs, stroke, thickness, dash, zindex):
    sg = StreamGeometry()
    try:
        ctx = sg.Open()
        for x1, y1, x2, y2 in segs:
            ctx.BeginFigure(Point(x1, y1), False, False)
            ctx.LineTo(Point(x2, y2), True, False)
        ctx.Close()
    except Exception:
        return None
    freeze_freezable(sg)
    path = Path()
    path.Data = sg
    path.Stroke = stroke
    path.StrokeThickness = float(thickness)
    path.Fill = None
    try:
        path.StrokeStartLineCap = PenLineCap.Flat
        path.StrokeEndLineCap = PenLineCap.Flat
        path.StrokeLineJoin = PenLineJoin.Miter
        path.SnapsToDevicePixels = True
    except Exception:
        pass
    dc = _make_dash_collection(dash)
    if dc is not None:
        path.StrokeDashArray = dc
    if zindex:
        try:
            Canvas.SetZIndex(path, int(zindex))
        except Exception:
            pass
    try:
        RenderOptions.SetEdgeMode(path, EdgeMode.Aliased)
    except Exception:
        pass
    return path


def _append_rect_figure(ctx, x, y, w, h, filled):
    # Contorno CCW cerrado; filled=True → IsFilled en BeginFigure.
    ctx.BeginFigure(Point(x, y), filled, True)
    ctx.LineTo(Point(x + w, y), True, False)
    ctx.LineTo(Point(x + w, y + h), True, False)
    ctx.LineTo(Point(x, y + h), True, False)


def _stream_fill_rects(rects, fill, zindex):
    sg = StreamGeometry()
    try:
        ctx = sg.Open()
        for x, y, w, h in rects:
            _append_rect_figure(ctx, x, y, w, h, True)
        ctx.Close()
    except Exception:
        return None
    freeze_freezable(sg)
    path = Path()
    path.Data = sg
    path.Fill = fill
    path.Stroke = None
    try:
        path.SnapsToDevicePixels = True
    except Exception:
        pass
    if zindex:
        try:
            Canvas.SetZIndex(path, int(zindex))
        except Exception:
            pass
    try:
        RenderOptions.SetEdgeMode(path, EdgeMode.Aliased)
    except Exception:
        pass
    return path


def _stream_stroke_rects(rects, stroke, thickness, dash, zindex):
    sg = StreamGeometry()
    try:
        ctx = sg.Open()
        for x, y, w, h in rects:
            _append_rect_figure(ctx, x, y, w, h, False)
        ctx.Close()
    except Exception:
        return None
    freeze_freezable(sg)
    path = Path()
    path.Data = sg
    path.Stroke = stroke
    path.StrokeThickness = float(thickness)
    path.Fill = None
    try:
        path.StrokeStartLineCap = PenLineCap.Flat
        path.StrokeEndLineCap = PenLineCap.Flat
        path.StrokeLineJoin = PenLineJoin.Miter
        path.SnapsToDevicePixels = True
    except Exception:
        pass
    dc = _make_dash_collection(dash)
    if dc is not None:
        path.StrokeDashArray = dc
    if zindex:
        try:
            Canvas.SetZIndex(path, int(zindex))
        except Exception:
            pass
    try:
        RenderOptions.SetEdgeMode(path, EdgeMode.Aliased)
    except Exception:
        pass
    return path


def apply_aliased_render(element):
    """Desactiva AA en líneas/puntos del elemento (ahorro GPU ortogonal)."""
    if element is None:
        return
    try:
        RenderOptions.SetEdgeMode(element, EdgeMode.Aliased)
    except Exception:
        pass
    try:
        element.SnapsToDevicePixels = True
    except Exception:
        pass
