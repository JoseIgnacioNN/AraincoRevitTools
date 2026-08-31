# -*- coding: utf-8 -*-
"""Puente CPython ↔ .NET 8 / WPF — colecciones tipadas, brushes Freezable, Dispatcher.

Todo el trabajo de UI y API Revit permanece en el hilo del Dispatcher (no
dispara hilos secundarios). Optimiza el puente pythonnet evitando que WPF
recorra listas nativas de Python o asigne brushes mutables sin freeze.
"""

from __future__ import print_function

import clr

clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")
clr.AddReference("System")

import System
from System import Action, String
from System.Collections.Generic import List as NetList
from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Media import SolidColorBrush, Color
from System.Windows.Threading import Dispatcher, DispatcherPriority

# Cache (hex, alpha) → SolidColorBrush congelado. Compartido UI + canvas.
_BRUSH_CACHE = {}


def freeze_freezable(obj):
    """Congela cualquier Freezable WPF (Brush, Pen, Geometry, DoubleCollection…)."""
    if obj is None:
        return None
    try:
        if getattr(obj, "CanFreeze", False) and not getattr(obj, "IsFrozen", True):
            obj.Freeze()
    except Exception:
        pass
    return obj


def freeze_brush(brush):
    """Congela un Freezable si es posible (seguro para compartir entre visuales)."""
    return freeze_freezable(brush)


# Cache (id(brush), thickness, dash_key) → Pen congelado.
_PEN_CACHE = {}


def pen_cached(brush, thickness=1.0, dash=None):
    """Pen Freezado reutilizable (DrawingContext / bajo nivel)."""
    if brush is None:
        return None
    try:
        t = float(thickness)
    except Exception:
        t = 1.0
    dk = None
    if dash:
        try:
            dk = tuple(float(x) for x in dash)
        except Exception:
            dk = None
    key = (id(brush), t, dk)
    cached = _PEN_CACHE.get(key)
    if cached is not None:
        return cached
    try:
        from System.Windows.Media import (
            DashStyle,
            DoubleCollection,
            Pen,
            PenLineCap,
        )

        pen = Pen(brush, t)
        pen.StartLineCap = PenLineCap.Flat
        pen.EndLineCap = PenLineCap.Flat
        if dk is not None:
            dc = DoubleCollection()
            for x in dk:
                dc.Add(float(x))
            freeze_freezable(dc)
            ds = DashStyle(dc, 0.0)
            freeze_freezable(ds)
            pen.DashStyle = ds
        freeze_freezable(pen)
        _PEN_CACHE[key] = pen
        return pen
    except Exception:
        return None


def clear_pen_cache():
    _PEN_CACHE.clear()


def brush_hex(hx, alpha=255):
    """SolidColorBrush cacheado y Freezado (alpha 0–255)."""
    h = (hx or u"#64748b").strip().lstrip(u"#")
    if len(h) < 6:
        h = u"64748b"
    else:
        h = h[0:6]
    try:
        aa = max(0, min(255, int(alpha)))
    except Exception:
        aa = 255
    key = (h.lower(), aa)
    cached = _BRUSH_CACHE.get(key)
    if cached is not None:
        return cached
    try:
        rr = int(h[0:2], 16)
        gg = int(h[2:4], 16)
        bb = int(h[4:6], 16)
    except Exception:
        rr, gg, bb = 0x64, 0x74, 0x8B
    brush = SolidColorBrush(Color.FromArgb(aa, rr, gg, bb))
    freeze_brush(brush)
    _BRUSH_CACHE[key] = brush
    return brush


def clear_brush_cache():
    """Libera entrada de caché (tests / unload). No afecta brushes ya asignados."""
    _BRUSH_CACHE.clear()


def to_net_int_list(iterable):
    """System.Collections.Generic.List[int] desde secuencia Python numérica."""
    try:
        net = NetList[System.Int32]()
    except Exception:
        try:
            net = NetList[int]()
        except Exception:
            # Fallback: lista Python (caller tolera for-in).
            out = []
            for item in iterable or []:
                try:
                    out.append(int(item))
                except Exception:
                    pass
            return out
    if not iterable:
        return net
    for item in iterable:
        try:
            net.Add(int(item))
        except Exception:
            continue
    return net


def to_net_string_list(iterable):
    """System.Collections.Generic.List[str] desde lista/tupla Python."""
    try:
        net = NetList[String]()
    except Exception:
        try:
            net = NetList[str]()
        except Exception:
            return list(iterable or [])
    if not iterable:
        return net
    for item in iterable:
        try:
            net.Add(System.Convert.ToString(item) if item is not None else u"")
        except Exception:
            try:
                net.Add(unicode(item))
            except NameError:
                net.Add(str(item))
    return net


def to_observable_strings(iterable):
    """ObservableCollection[str] para ItemsSource de ListBox/ComboBox tipados."""
    col = ObservableCollection[String]()
    if not iterable:
        return col
    for item in iterable:
        try:
            col.Add(System.Convert.ToString(item) if item is not None else u"")
        except Exception:
            try:
                col.Add(unicode(item))
            except NameError:
                col.Add(str(item))
    return col


def dispatcher_of(element):
    """Dispatcher del elemento WPF o del Application actual."""
    try:
        if element is not None and getattr(element, "Dispatcher", None) is not None:
            return element.Dispatcher
    except Exception:
        pass
    try:
        from System.Windows import Application

        app = Application.Current
        if app is not None:
            return app.Dispatcher
    except Exception:
        pass
    return None


class CoalescedUiAction(object):
    """Fusiona N solicitudes de redibujado en un único BeginInvoke (hilo UI).

    Uso típico::
        self._coalesce = CoalescedUiAction(window, self._redraw_canvas)
        self._coalesce.request()  # múltiples llamadas → 1 ejecución
    """

    def __init__(self, element, action, priority=None):
        self._element = element
        self._action = action
        self._priority = priority if priority is not None else DispatcherPriority.Background
        self._pending = False
        self._running = False
        self._requeue = False

    def request(self):
        if self._pending and not self._running:
            return
        if self._running:
            self._requeue = True
            return
        self._pending = True
        disp = dispatcher_of(self._element)
        if disp is None:
            self._pending = False
            try:
                self._action()
            except Exception:
                pass
            return
        try:
            disp.BeginInvoke(Action(self._run), self._priority)
        except Exception:
            self._pending = False
            try:
                self._action()
            except Exception:
                pass

    def _run(self):
        self._pending = False
        self._running = True
        try:
            self._action()
        finally:
            self._running = False
            if self._requeue:
                self._requeue = False
                self.request()


class BatchedUiWork(object):
    """Empaqueta trabajo visual en lotes vía Dispatcher.InvokeAsync / BeginInvoke.

    Cada lote se ejecuta en el hilo del Dispatcher (API Revit si se invoca
    desde el lote: solo hilo principal — el caller debe asegurar datos ya leídos).
    """

    def __init__(self, element, batch_size=24, priority=None):
        self._element = element
        self._batch_size = max(1, int(batch_size))
        self._priority = priority if priority is not None else DispatcherPriority.Background
        self._queue = []
        self._scheduled = False

    def enqueue(self, work):
        """``work`` es un callable sin args a ejecutar en el hilo UI."""
        if work is None:
            return
        self._queue.append(work)
        self._schedule()

    def enqueue_many(self, works):
        if not works:
            return
        self._queue.extend(works)
        self._schedule()

    def _schedule(self):
        if self._scheduled or not self._queue:
            return
        self._scheduled = True
        disp = dispatcher_of(self._element)
        if disp is None:
            self._scheduled = False
            self.flush_sync()
            return
        try:
            # InvokeAsync en .NET 4.5+; fallback BeginInvoke.
            try:
                disp.InvokeAsync(Action(self._run_batch), self._priority)
            except Exception:
                disp.BeginInvoke(Action(self._run_batch), self._priority)
        except Exception:
            self._scheduled = False
            self.flush_sync()

    def _run_batch(self):
        self._scheduled = False
        n = 0
        while self._queue and n < self._batch_size:
            fn = self._queue.pop(0)
            n += 1
            try:
                fn()
            except Exception:
                pass
        if self._queue:
            self._schedule()

    def flush_sync(self):
        """Drena la cola de forma síncrona (carga inicial / shutdown)."""
        while self._queue:
            fn = self._queue.pop(0)
            try:
                fn()
            except Exception:
                pass
        self._scheduled = False


class DebouncedUiAction(object):
    """Retrasa la acción hasta ``delay_ms`` sin nuevas solicitudes (SizeChanged)."""

    def __init__(self, element, action, delay_ms=80, priority=None):
        self._element = element
        self._action = action
        self._delay_ms = max(16, int(delay_ms))
        self._priority = priority if priority is not None else DispatcherPriority.Background
        self._token = 0
        self._timer = None

    def request(self):
        self._token += 1
        token = self._token
        disp = dispatcher_of(self._element)
        if disp is None:
            try:
                self._action()
            except Exception:
                pass
            return

        def _fire():
            if token != self._token:
                return
            try:
                self._action()
            except Exception:
                pass

        try:
            from System import TimeSpan
            from System.Windows.Threading import DispatcherTimer

            if self._timer is not None:
                try:
                    self._timer.Stop()
                except Exception:
                    pass
                self._timer = None

            timer = DispatcherTimer()
            timer.Interval = TimeSpan.FromMilliseconds(self._delay_ms)
            try:
                timer.Priority = self._priority
            except Exception:
                pass

            def _on_tick(sender, args):
                try:
                    timer.Stop()
                except Exception:
                    pass
                if self._timer is timer:
                    self._timer = None
                _fire()

            timer.Tick += _on_tick
            self._timer = timer
            timer.Start()
        except Exception:
            try:
                disp.BeginInvoke(Action(_fire), self._priority)
            except Exception:
                _fire()
