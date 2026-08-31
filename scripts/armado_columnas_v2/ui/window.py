# -*- coding: utf-8 -*-
"""Ventana principal WPF — instancia única (shell Armado vigas)."""

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")

import System
from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage
from System.Windows.Markup import XamlReader

from Autodesk.Revit.UI import TaskDialog

from armado_columnas_v2.session import SESSION
from armado_columnas_v2.session import ColumnArmadoSession
from armado_columnas_v2.ui.canvas_view import ArmadoColumnasCanvasView
from armado_columnas_v2.ui.xaml import build_armado_columnas_v2_xaml

_APP_DOMAIN_KEY = u"BIMTools_ArmadoColumnasV2_Window"
_DIALOG = u"Arainco: Armado columnas V2"


def _clear_window_ref(_window=None):
    try:
        System.AppDomain.CurrentDomain.SetData(_APP_DOMAIN_KEY, None)
    except Exception:
        pass


def _get_existing_window():
    try:
        w = System.AppDomain.CurrentDomain.GetData(_APP_DOMAIN_KEY)
        if w is None:
            return None
        win = getattr(w, "_win", None)
        if win is None:
            _clear_window_ref()
            return None
        try:
            _ = win.Title
        except Exception:
            _clear_window_ref()
            return None
        try:
            if not bool(w.IsLoaded):
                _clear_window_ref()
                return None
        except Exception:
            _clear_window_ref()
            return None
        return w
    except Exception:
        return None


def _focus_existing_window(existing):
    if existing is None:
        return
    try:
        if existing.WindowState == System.Windows.WindowState.Minimized:
            existing.WindowState = System.Windows.WindowState.Normal
    except Exception:
        pass
    try:
        wpf = getattr(existing, "_win", None)
        if wpf is not None and not wpf.IsVisible:
            wpf.Show()
        # Reaplicar maximizado en el monitor de Revit al reenfocar
        if wpf is not None and hasattr(existing, "_uiapp"):
            from revit_wpf_window_position import (
                fill_wpf_window_on_work_area_px,
                revit_main_hwnd,
            )
            from revit_wpf_window_position import _revit_monitor_work_area

            hwnd = revit_main_hwnd(existing._uiapp)
            area = _revit_monitor_work_area(hwnd)
            if area is not None:
                left_px, top_px, width_px, height_px = area
                fill_wpf_window_on_work_area_px(
                    wpf, left_px, top_px, width_px, height_px, hwnd,
                )
    except Exception:
        pass
    try:
        existing.Activate()
    except Exception:
        pass


class ArmadoColumnasV2Window(object):
    """Controlador de la ventana WPF (``self._win`` es el ``Window`` parseado)."""

    def __init__(self, uiapp, pushbutton_dir=None):
        self._uiapp = uiapp
        self._pushbutton_dir = pushbutton_dir
        self._win = None
        self._canvas = None
        self._build_ui()

    @property
    def IsLoaded(self):
        try:
            return self._win is not None and self._win.IsLoaded
        except Exception:
            return False

    def Activate(self):
        if self._win is not None:
            self._win.Activate()

    @property
    def WindowState(self):
        return self._win.WindowState if self._win is not None else None

    @WindowState.setter
    def WindowState(self, value):
        if self._win is not None:
            self._win.WindowState = value

    def Close(self):
        if self._win is not None:
            self._win.Close()

    def _build_ui(self):
        try:
            self._win = XamlReader.Parse(build_armado_columnas_v2_xaml())
        except Exception as ex:
            try:
                msg = unicode(ex)
            except NameError:
                msg = str(ex)
            TaskDialog.Show(_DIALOG, u"No se cargó la ventana WPF:\n{0}".format(msg))
            return

        callbacks = {
            "on_redraw": lambda: self._dispatch_ui(self._redraw_canvas),
            "on_status": self.set_status,
            "on_select_column": self._on_select_column,
        }
        self._canvas = ArmadoColumnasCanvasView(self._win, callbacks)

        try:
            from System import EventHandler as _EH_clr
            from System.Windows import RoutedEventHandler as _REH

            self._win.Closed += _EH_clr(self._on_closed)
            self._win.Loaded += _REH(lambda s, e: self._dispatch_ui(self._redraw_canvas))
            self._win.SizeChanged += _REH(lambda s, e: self._dispatch_ui(self._redraw_canvas))
        except Exception:
            pass

        self._wire_controls()
        self._redraw_canvas()
        self.set_status(SESSION.last_message or u"Configure el armado del lote.")

    def _wire_controls(self):
        try:
            from System.Windows import RoutedEventHandler as _REH
        except Exception:
            return

        btn_col = self._win.FindName(u"BtnColocar")
        if btn_col is not None:
            btn_col.Click += _REH(lambda s, e: self._on_colocar())

        btn_cancel = self._win.FindName(u"BtnCancelar")
        if btn_cancel is not None:
            btn_cancel.Click += _REH(lambda s, e: self.Close())

    def _on_colocar(self):
        self.set_status(
            u"Colocación deshabilitada en esta fase (solo diseño UI)."
        )

    def _on_select_column(self, col_id, multi=False):
        SESSION.select_column(col_id, multi=multi)
        col = SESSION.preview_member()
        name = (col or {}).get("label") or col_id or u"—"
        kind = (col or {}).get("kind") or u"column"
        rol = ColumnArmadoSession.kind_label_es(kind)
        n = len(SESSION.selected_ids or [])
        if n > 1:
            self.set_status(u"{0} elementos · preview {1} ({2}).".format(n, name, rol))
        else:
            self.set_status(u"{0} · {1} · preview sección.".format(name, rol))
        self._redraw_canvas()

    def _dispatch_ui(self, action):
        if self._win is None:
            return
        try:
            from System import Action
            from System.Windows.Threading import DispatcherPriority

            self._win.Dispatcher.BeginInvoke(
                Action(action),
                DispatcherPriority.Normal,
            )
        except Exception:
            try:
                action()
            except Exception as ex:
                self.set_status(self._format_error(ex))

    @staticmethod
    def _format_error(ex):
        try:
            msg = unicode(ex)
        except NameError:
            msg = str(ex)
        return u"Error: {0}".format(msg)

    def _redraw_canvas(self):
        if self._canvas is None:
            return False
        try:
            self._canvas.redraw(SESSION)
            return True
        except Exception as ex:
            self.set_status(self._format_error(ex))
            return False

    def set_status(self, text):
        try:
            tb = self._win.FindName(u"TxtEstado") if self._win else None
            if tb is not None:
                tb.Text = text or u""
        except Exception:
            pass

    def _on_closed(self, sender, args):
        _clear_window_ref(self)

    def Show(self):
        if self._win is None:
            return
        try:
            from revit_wpf_window_position import (
                bind_maximize_wpf_on_revit_monitor,
                fill_wpf_window_on_work_area_px,
                preposition_wpf_window_on_work_area,
                revit_main_hwnd,
            )

            hwnd = revit_main_hwnd(self._uiapp)
            # Área de trabajo del monitor donde corre Revit (no el secundario).
            try:
                from revit_wpf_window_position import _revit_monitor_work_area

                area = _revit_monitor_work_area(hwnd)
            except Exception:
                area = None

            if area is not None:
                left_px, top_px, width_px, height_px = area
                # Antes de Show: anclar al monitor correcto (multi-monitor).
                preposition_wpf_window_on_work_area(
                    self._win, left_px, top_px, width_px, height_px, hwnd,
                )
                # Al maximizar manualmente, permanece en el monitor de Revit.
                bind_maximize_wpf_on_revit_monitor(self._win, hwnd)

                applied = [False]

                def _apply_fill(sender, args):
                    if applied[0]:
                        return
                    if fill_wpf_window_on_work_area_px(
                        self._win, left_px, top_px, width_px, height_px, hwnd,
                    ):
                        applied[0] = True

                try:
                    from System.Windows import RoutedEventHandler

                    h = RoutedEventHandler(_apply_fill)
                    self._win.SourceInitialized += h
                    self._win.Loaded += h
                    try:
                        self._win.ContentRendered += h
                    except Exception:
                        pass
                except Exception:
                    fill_wpf_window_on_work_area_px(
                        self._win, left_px, top_px, width_px, height_px, hwnd,
                    )
            else:
                # Fallback: maximizar WPF clásico + bind al monitor de Revit si se puede
                try:
                    from System.Windows import WindowState

                    bind_maximize_wpf_on_revit_monitor(self._win, hwnd)
                    self._win.WindowState = WindowState.Maximized
                except Exception:
                    pass
        except Exception:
            try:
                from System.Windows import WindowState

                self._win.WindowState = WindowState.Maximized
            except Exception:
                pass
        self._win.Show()


def get_existing_armado_columnas_v2_window():
    return _get_existing_window()


def show_armado_columnas_v2_window(uiapp, pushbutton_dir=None):
    existing = _get_existing_window()
    if existing is not None:
        _focus_existing_window(existing)
        MessageBox.Show(
            u"La herramienta ya está en ejecución.",
            _DIALOG,
            MessageBoxButton.OK,
            MessageBoxImage.Information,
        )
        return existing

    win = ArmadoColumnasV2Window(uiapp, pushbutton_dir)
    if win._win is None:
        return None
    win.Show()
    try:
        System.AppDomain.CurrentDomain.SetData(_APP_DOMAIN_KEY, win)
    except Exception:
        pass
    return win
