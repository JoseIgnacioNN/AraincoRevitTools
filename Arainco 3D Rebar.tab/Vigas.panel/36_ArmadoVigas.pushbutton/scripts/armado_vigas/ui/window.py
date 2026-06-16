# -*- coding: utf-8 -*-
"""Ventana principal WPF/XAML — instancia única (patrón Armado Muros)."""

import weakref

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")

import System
from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage
from System.Windows.Markup import XamlReader

from Autodesk.Revit.UI import ExternalEvent, TaskDialog

from armado_vigas.domain.suple_inferior import beam_suple_inf_enabled
from armado_vigas.domain.tramos import build_session_tramos, sort_beams
from armado_vigas.revit.colocar import ColocarArmaduraHandler
from armado_vigas.revit.direction_overlay import ClearDirectionOverlayHandler
from armado_vigas.revit.session import SESSION
from armado_vigas.ui.canvas_view import ArmadoVigasCanvasView
from armado_vigas.ui import layout as lay
from armado_vigas.ui.xaml import build_armado_vigas_xaml

_APP_DOMAIN_KEY = u"BIMTools_ArmadoVigas_Window"


def _get_existing_window():
    try:
        w = System.AppDomain.CurrentDomain.GetData(_APP_DOMAIN_KEY)
        if w is not None and hasattr(w, "IsLoaded") and w.IsLoaded:
            return w
    except Exception:
        pass
    return None


def _clear_window_ref(window):
    try:
        System.AppDomain.CurrentDomain.SetData(_APP_DOMAIN_KEY, None)
    except Exception:
        pass


class ArmadoVigasWindow(object):
    """Controlador de la ventana WPF (``self._win`` es el ``Window`` parseado)."""

    def __init__(self, uiapp, pushbutton_dir=None):
        self._uiapp = uiapp
        self._pushbutton_dir = pushbutton_dir
        self._win = None
        self._canvas = None
        self._colocar_handler = ColocarArmaduraHandler(weakref.ref(self))
        self._colocar_event = ExternalEvent.Create(self._colocar_handler)
        self._clear_overlay_handler = ClearDirectionOverlayHandler()
        self._clear_overlay_event = ExternalEvent.Create(self._clear_overlay_handler)
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

    def request_close(self):
        """Cierra la ventana en el hilo WPF (desde ExternalEvent)."""
        self._dispatch_ui(self.Close)

    def _build_ui(self):
        try:
            self._win = XamlReader.Parse(build_armado_vigas_xaml())
        except Exception as ex:
            try:
                msg = unicode(ex)
            except NameError:
                msg = str(ex)
            TaskDialog.Show(u"Arainco: Armado vigas", u"No se cargó la ventana WPF:\n{0}".format(msg))
            return

        callbacks = {
            "on_redraw": lambda: self._dispatch_ui(self._redraw_canvas),
            "on_status": self.set_status,
            "on_toggle_empalme": self._toggle_empalme,
            "on_select_tramo": self._on_select_tramo,
            "on_select_beam": self._on_select_beam,
            "on_select_stirrup_zone": self._on_select_stirrup_zone,
        }
        self._canvas = ArmadoVigasCanvasView(self._win, callbacks)

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
        self.set_status(SESSION.last_message or u"Configure el armado del lote seleccionado.")

    def _wire_controls(self):
        try:
            from System.Windows import RoutedEventHandler as _REH
        except Exception:
            return

        btn_col = self._win.FindName(u"BtnColocar")
        if btn_col is not None:
            btn_col.Click += _REH(lambda s, e: self.raise_colocar())

        btn_cancel = self._win.FindName(u"BtnCancelar")
        if btn_cancel is not None:
            btn_cancel.Click += _REH(lambda s, e: self.Close())

    def _active_view(self):
        try:
            uidoc = self._uiapp.ActiveUIDocument if self._uiapp else None
            return uidoc.ActiveView if uidoc is not None else None
        except Exception:
            return None

    def _apply_view_order(self):
        from armado_vigas.revit.view_order import assign_beam_view_order, assign_beam_col_endpoints

        view = self._active_view()
        assign_beam_view_order(SESSION.domain_beams, view)
        assign_beam_col_endpoints(SESSION.domain_beams, SESSION.apoyos, view)

    def _rebuild_tramos(self):
        self._apply_view_order()
        beams = sort_beams(list(SESSION.domain_beams or []))
        SESSION.tramos_sup, SESSION.tramos_inf = build_session_tramos(
            beams,
            empalme_beam_ids_sup=SESSION.empalme_beam_ids_sup,
            empalme_beam_ids_inf=SESSION.empalme_beam_ids_inf,
            split_empalme=SESSION.split_empalme,
        )
        SESSION.tramos = SESSION.tramos_sup

    def _toggle_empalme(self, beam_id, face=u"inf"):
        if not beam_id:
            return
        is_sup = face == u"sup"
        target = SESSION.empalme_beam_ids_sup if is_sup else SESSION.empalme_beam_ids_inf
        cara = u"superior" if is_sup else u"inferior"
        if beam_id in target:
            target.discard(beam_id)
            self.set_status(
                u"Traslapo {0} desmarcado · {1} · tramos recalculados.".format(cara, beam_id)
            )
        else:
            target.add(beam_id)
            self.set_status(
                u"Traslapo @ mitad · fibra {0} · {1} · tramos recalculados.".format(cara, beam_id)
            )
        self._rebuild_tramos()

    def _on_select_tramo(self, tramo_id, face=u"sup"):
        cara = u"superior" if face == u"sup" else u"inferior"
        self.set_status(u"Tramo T{0} · cara {1}.".format(tramo_id, cara))

    def _on_select_beam(self, idx):
        beams = sort_beams(list(SESSION.domain_beams or []))
        if 0 <= idx < len(beams):
            self.set_status(u"{0} · preview sección.".format(lay.beam_canvas_label(idx)))

    def _on_select_stirrup_zone(self, idx, role):
        beams = sort_beams(list(SESSION.domain_beams or []))
        if 0 <= idx < len(beams):
            if role == u"confin":
                self.set_status(u"Confin. global · {0}.".format(beams[idx].get("id")))
            elif role == u"suple":
                beam = beams[idx]
                on = beam_suple_inf_enabled(beam)
                self.set_status(
                    u"Suple inf. · {0} · {1} · ø{2} · n={3}.".format(
                        beam.get("id"),
                        u"Sí" if on else u"No",
                        int(beam.get("diamSupleInf") or 16),
                        int(beam.get("nSupleInf") or 2),
                    )
                )
            elif role == u"laterales":
                self.set_status(
                    u"Laterales · lote · {0} · n={1} · ø{2}.".format(
                        u"Sí" if getattr(SESSION, "lateralesEnabled", False) else u"No",
                        int(getattr(SESSION, "nLaterales", 1) or 1),
                        int(getattr(SESSION, "diamLaterales", 16) or 16),
                    )
                )
            else:
                labels = {"ext": u"Ext ini/fin", "cent": u"Cent", "uni": u"Único"}
                self.set_status(u"{0} · {1}.".format(labels.get(role, role), beams[idx].get("id")))

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
            self._rebuild_tramos()
            self._canvas.redraw(SESSION)
            return True
        except Exception as ex:
            self.set_status(self._format_error(ex))
            return False

    def raise_colocar(self):
        self.set_status(u"Colocando armadura…")
        try:
            if self._colocar_event is not None:
                self._colocar_event.Raise()
            else:
                self.set_status(u"Error: evento de colocación no disponible.")
        except Exception as ex:
            self.set_status(self._format_error(ex))

    def set_status(self, text):
        try:
            tb = self._win.FindName(u"TxtEstado") if self._win else None
            if tb is not None:
                tb.Text = text or u""
        except Exception:
            pass

    def _on_closed(self, sender, args):
        try:
            self._clear_overlay_event.Raise()
        except Exception:
            pass
        _clear_window_ref(self)

    def Show(self):
        if self._win is None:
            return
        try:
            from revit_wpf_window_position import (
                bind_maximize_wpf_on_secondary_monitor,
                position_wpf_window_top_left_at_active_view,
                revit_main_hwnd,
            )

            hwnd = revit_main_hwnd(self._uiapp)
            uidoc = self._uiapp.ActiveUIDocument if self._uiapp else None
            if not bind_maximize_wpf_on_secondary_monitor(self._win, hwnd):
                position_wpf_window_top_left_at_active_view(self._win, uidoc, hwnd)
        except Exception:
            pass
        self._win.Show()


def get_existing_armado_vigas_window():
    return _get_existing_window()


def show_armado_vigas_window(uiapp, pushbutton_dir=None):
    existing = _get_existing_window()
    if existing is not None:
        try:
            if existing.WindowState == System.Windows.WindowState.Minimized:
                existing.WindowState = System.Windows.WindowState.Normal
            existing.Activate()
        except Exception:
            pass
        MessageBox.Show(
            u"La herramienta ya está en ejecución.",
            u"Arainco: Armado vigas",
            MessageBoxButton.OK,
            MessageBoxImage.Information,
        )
        return existing

    win = ArmadoVigasWindow(uiapp, pushbutton_dir)
    try:
        System.AppDomain.CurrentDomain.SetData(_APP_DOMAIN_KEY, win)
    except Exception:
        pass
    win.Show()
    return win
