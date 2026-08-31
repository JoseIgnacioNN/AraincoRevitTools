# -*- coding: utf-8 -*-
"""Ventana principal WPF/XAML — Armado vigas (preview + colocación Rebar)."""

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")

import System
from System.Windows.Markup import XamlReader

from Autodesk.Revit.UI import ExternalEvent

from armado_vigas.domain.constants import (
    CONCRETE_GRADE_DEFAULT,
    normalize_concrete_grade,
)
from armado_vigas.domain.suple_inferior import beam_suple_inf_enabled
from armado_vigas.domain.suple_superior import (
    beam_suple_sup_enabled,
    beam_suple_sup_side_enabled,
)
from armado_vigas.domain.tramos import build_session_tramos, sort_beams
from armado_vigas.revit.colocar import (
    clear_colocar_target,
    ensure_colocar_event,
    raise_colocar_armadura,
)
from armado_vigas.revit.direction_overlay import ClearDirectionOverlayHandler
from armado_vigas.revit.session import SESSION
from armado_vigas.ui.canvas_view import ArmadoVigasCanvasView
from armado_vigas.ui import layout as lay
from armado_vigas.ui.instruction_dialog import DIALOG_TITLE, show_info
from armado_vigas.ui.xaml import build_armado_vigas_xaml
from armado_vigas.ui.net_ui import CoalescedUiAction, DebouncedUiAction

_APP_DOMAIN_KEY = u"Arainco_ArmadoVigas_Window"
_DIALOG = DIALOG_TITLE


def _clear_window_ref(_window=None):
    try:
        System.AppDomain.CurrentDomain.SetData(_APP_DOMAIN_KEY, None)
    except Exception:
        pass


def _get_existing_window():
    """Instancia viva en AppDomain; limpia referencias stale o ventana cerrada."""
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
    """Restaura ventana y la trae maximizada al frente."""
    if existing is None:
        return
    try:
        wpf = getattr(existing, "_win", None)
        if wpf is not None:
            try:
                from System import Double

                wpf.MaxWidth = Double.PositiveInfinity
                wpf.MaxHeight = Double.PositiveInfinity
            except Exception:
                pass
            existing.WindowState = System.Windows.WindowState.Maximized
            if not wpf.IsVisible:
                wpf.Show()
        else:
            existing.WindowState = System.Windows.WindowState.Maximized
    except Exception:
        try:
            if existing.WindowState == System.Windows.WindowState.Minimized:
                existing.WindowState = System.Windows.WindowState.Normal
        except Exception:
            pass
    try:
        existing.Activate()
    except Exception:
        pass


class ArmadoVigasWindow(object):
    """Controlador de la ventana WPF (``self._win`` es el ``Window`` parseado)."""

    def __init__(self, uiapp, pushbutton_dir=None):
        self._uiapp = uiapp
        self._pushbutton_dir = pushbutton_dir
        self._win = None
        self._canvas = None
        self._clear_overlay_handler = ClearDirectionOverlayHandler()
        self._clear_overlay_event = ExternalEvent.Create(self._clear_overlay_handler)
        # ExternalEvent de colocación: singleton de proceso (no recrear por ventana).
        ensure_colocar_event()
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
        self._dispatch_ui(self.Close)

    def _build_ui(self):
        try:
            self._win = XamlReader.Parse(build_armado_vigas_xaml())
        except Exception as ex:
            try:
                msg = unicode(ex)
            except NameError:
                msg = str(ex)
            show_info(
                u"No se cargó la ventana WPF.",
                msg,
                title=_DIALOG,
                uiapp=getattr(self, u"_uiapp", None),
            )
            return

        callbacks = {
            # Edits de rail no cambian topología de tramos (empalme sí lo hace en su handler).
            "on_redraw": lambda: self.request_redraw(
                rebuild_session=False, refresh_view=False
            ),
            "on_status": self.set_status,
            "on_toggle_empalme": self._toggle_empalme,
            "on_select_tramo": self._on_select_tramo,
            "on_select_beam": self._on_select_beam,
            "on_select_stirrup_zone": self._on_select_stirrup_zone,
        }
        self._canvas = ArmadoVigasCanvasView(self._win, callbacks)

        # Redibujado coalescido (N clicks UI → 1 paint). Solo hilo Dispatcher.
        from System.Windows.Threading import DispatcherPriority

        self._redraw_coalesce = CoalescedUiAction(
            self._win, self._redraw_canvas_work, priority=DispatcherPriority.Normal,
        )
        self._resize_debounce = DebouncedUiAction(
            self._win,
            self._on_size_changed_redraw,
            delay_ms=90,
            priority=DispatcherPriority.Background,
        )
        self._redraw_opts = {"rebuild_session": True, "refresh_view": False}

        try:
            from System import EventHandler as _EH_clr
            from System.Windows import RoutedEventHandler as _REH
            from System.Windows import SizeChangedEventHandler as _SCEH

            self._win.Closed += _EH_clr(self._on_closed)
            self._win.Loaded += _REH(
                lambda s, e: self.request_redraw(rebuild_session=True, refresh_view=True)
            )
            try:
                self._win.SizeChanged += _SCEH(lambda s, e: self._resize_debounce.request())
            except Exception:
                self._win.SizeChanged += _REH(lambda s, e: self._resize_debounce.request())
        except Exception:
            pass

        self._wire_controls()
        # Primer paint solo en Loaded (viewport válido).
        self.set_status(
            SESSION.last_message
            or u"Configure SUP/INF/LAT/CONF · Colocar crea Rebar en el modelo."
        )

    def _wire_controls(self):
        try:
            from System.Windows import RoutedEventHandler as _REH
        except Exception:
            return

        btn_col = self._win.FindName(u"BtnColocar")
        if btn_col is not None:
            try:
                btn_col.Content = u"Colocar armadura"
                btn_col.ToolTip = (
                    u"Crea Rebar según toggles SUP / INF / LAT / CONF del rail."
                )
            except Exception:
                pass
            btn_col.Click += _REH(lambda s, e: self.raise_colocar())

        btn_cancel = self._win.FindName(u"BtnCancelar")
        if btn_cancel is not None:
            btn_cancel.Click += _REH(lambda s, e: self.Close())

        self._wire_dosificacion()

    def _wire_dosificacion(self):
        """Asegura grado de sesión; el control UI vive en el rail (Configuración viga)."""
        grade = normalize_concrete_grade(
            getattr(SESSION, "concreteGrade", CONCRETE_GRADE_DEFAULT)
        )
        SESSION.set_concrete_grade(grade)
        self._cmb_dosif = None

    def _active_view(self):
        try:
            uidoc = self._uiapp.ActiveUIDocument if self._uiapp else None
            return uidoc.ActiveView if uidoc is not None else None
        except Exception:
            return None

    def _apply_view_order(self):
        from armado_vigas.revit.view_order import assign_beam_view_order

        view = self._active_view()
        assign_beam_view_order(SESSION.domain_beams, view)
        try:
            from armado_vigas.revit.elev_geometry import (
                assign_beam_supports_by_proximity,
                enrich_session_elev_geometry,
            )

            doc = None
            try:
                uidoc = self._uiapp.ActiveUIDocument if self._uiapp else None
                doc = uidoc.Document if uidoc is not None else None
            except Exception:
                doc = None
            enrich_session_elev_geometry(
                SESSION.domain_beams, SESSION.apoyos, view, document=doc
            )
            assign_beam_supports_by_proximity(SESSION.domain_beams, SESSION.apoyos, view)
        except Exception:
            from armado_vigas.revit.view_order import assign_beam_col_endpoints

            assign_beam_col_endpoints(SESSION.domain_beams, SESSION.apoyos, view)

    def _rebuild_tramos(self, refresh_view=False):
        """Recalcula tramos. ``refresh_view`` relee geometría de vista (Revit API, hilo UI).

        En redibujados de UI (capas, selección, resize) se omite re-enriquecer
        geometría del alzado — mismo resultado visual, menos round-trips a Revit.
        """
        if refresh_view:
            self._apply_view_order()
        beams = sort_beams(list(SESSION.domain_beams or []))
        SESSION.tramos_sup, SESSION.tramos_inf = build_session_tramos(
            beams,
            empalme_beam_ids_sup=SESSION.empalme_beam_ids_sup,
            empalme_beam_ids_inf=SESSION.empalme_beam_ids_inf,
            split_empalme=SESSION.split_empalme,
        )
        try:
            from armado_vigas.domain.tramo_armado import merge_armado_onto_tramos

            merge_armado_onto_tramos(SESSION, u"sup", SESSION.tramos_sup, beams)
            merge_armado_onto_tramos(SESSION, u"inf", SESSION.tramos_inf, beams)
        except Exception:
            pass
        SESSION.tramos = SESSION.tramos_sup

    def request_redraw(self, rebuild_session=True, refresh_view=False):
        """Solicita paint coalescido en el Dispatcher (sin hilos secundarios)."""
        prev = self._redraw_opts or {}
        self._redraw_opts = {
            "rebuild_session": bool(prev.get("rebuild_session")) or bool(rebuild_session),
            "refresh_view": bool(prev.get("refresh_view")) or bool(refresh_view),
        }
        try:
            self._redraw_coalesce.request()
        except Exception:
            self._redraw_canvas_work()

    def _on_size_changed_redraw(self):
        # Solo re-layout WPF: sin rebuild de tramos ni relectura de geometría Revit.
        self.request_redraw(rebuild_session=False, refresh_view=False)

    def _redraw_canvas_work(self):
        opts = self._redraw_opts or {}
        self._redraw_opts = {"rebuild_session": False, "refresh_view": False}
        rebuild = bool(opts.get("rebuild_session"))
        refresh_view = bool(opts.get("refresh_view"))
        self._redraw_canvas(rebuild_session=rebuild, refresh_view=refresh_view)

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
        self._rebuild_tramos(refresh_view=False)
        # Fuerza rebuild de alzado (topología Tn cambió); rail ya no lo pide.
        try:
            if self._canvas is not None:
                self._canvas.invalidate_elev_cache()
        except Exception:
            pass
        self.request_redraw(rebuild_session=False, refresh_view=False)

    def _on_select_tramo(self, tramo_id, face=u"sup"):
        cara = u"superior" if face == u"sup" else u"inferior"
        cv = getattr(self, "_canvas", None) or getattr(self, "canvas", None)
        n = 1
        try:
            ids_attr = (
                u"selected_tramo_ids_sup" if face == u"sup" else u"selected_tramo_ids_inf"
            )
            ids = getattr(cv, ids_attr, None) if cv is not None else None
            if ids:
                n = len(ids)
        except Exception:
            n = 1
        if n > 1:
            self.set_status(
                u"{0} tramos · foco T{1} · cara {2}.".format(n, tramo_id, cara)
            )
        else:
            self.set_status(u"Tramo T{0} · cara {1}.".format(tramo_id, cara))

    def _on_select_beam(self, idx, n_selected=1):
        beams = sort_beams(list(SESSION.domain_beams or []))
        if 0 <= idx < len(beams):
            if n_selected > 1:
                self.set_status(
                    u"{0} vigas seleccionadas · preview {1} · cambios en lote.".format(
                        n_selected, lay.beam_canvas_label(idx),
                    )
                )
            else:
                self.set_status(u"{0} · preview sección.".format(lay.beam_canvas_label(idx)))

    def _on_select_stirrup_zone(self, idx, role):
        beams = sort_beams(list(SESSION.domain_beams or []))
        if 0 <= idx < len(beams):
            if role == u"confin":
                self.set_status(u"Confin. · {0}.".format(beams[idx].get("id")))
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
            elif role == u"supleSup":
                beam = beams[idx]
                on = beam_suple_sup_enabled(beam)
                n_ap = 0
                try:
                    n_ap = len(getattr(SESSION, u"suple_sup_apoyo_ids", None) or set())
                except Exception:
                    n_ap = 0
                self.set_status(
                    u"Suple SUP por apoyo · {0} apoyo(s) · L/3 · "
                    u"{1} Ini {2}/{3} · ø{4} · n={5}.".format(
                        n_ap,
                        beam.get("id"),
                        beam.get("colStart") or u"—",
                        beam.get("colEnd") or u"—",
                        int(beam.get("diamSupleSup") or 16),
                        int(beam.get("nSupleSup") or 2),
                    )
                    if on or n_ap
                    else u"Suple SUP · off · clic en columnas/muros del alzado (L/3)."
                )
            elif role == u"laterales":
                from armado_vigas.domain.laterales import LATERALES_DIAM_DEFAULT

                d0 = int(LATERALES_DIAM_DEFAULT)
                self.set_status(
                    u"Laterales · lote · {0} · n={1} · ø{2}.".format(
                        u"Sí" if getattr(SESSION, "lateralesEnabled", False) else u"No",
                        int(getattr(SESSION, "nLaterales", 0) if getattr(SESSION, "nLaterales", None) is not None else 0),
                        int(getattr(SESSION, "diamLaterales", d0) or d0),
                    )
                )
            else:
                labels = {"ext": u"Ext ini/fin", "cent": u"Cent", "uni": u"Único"}
                self.set_status(u"{0} · {1}.".format(labels.get(role, role), beams[idx].get("id")))

    def _dispatch_ui(self, action, priority=None):
        """Ejecuta en el Dispatcher de la ventana (hilo UI / API Revit)."""
        if self._win is None:
            return
        try:
            from System import Action
            from System.Windows.Threading import DispatcherPriority

            pri = priority if priority is not None else DispatcherPriority.Normal
            # InvokeAsync (.NET) cuando existe; fallback BeginInvoke.
            try:
                self._win.Dispatcher.InvokeAsync(Action(action), pri)
            except Exception:
                self._win.Dispatcher.BeginInvoke(Action(action), pri)
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

    def _redraw_canvas(self, rebuild_session=True, refresh_view=False):
        if self._canvas is None:
            return False
        try:
            if rebuild_session:
                self._rebuild_tramos(refresh_view=refresh_view)
            self._canvas.redraw(SESSION)
            return True
        except Exception as ex:
            self.set_status(self._format_error(ex))
            return False

    def _sync_place_flags_from_ui(self):
        """Copia toggles del rail a SESSION antes de colocar."""
        cv = self._canvas
        if cv is None:
            return
        try:
            from armado_vigas.ui.rail_cards import ensure_rail_state

            ensure_rail_state(cv)
        except Exception:
            pass
        SESSION.placeSup = bool(getattr(cv, "card_on_sup", True))
        SESSION.placeInf = bool(getattr(cv, "card_on_inf", True))
        SESSION.placeConf = bool(getattr(cv, "card_on_conf", True))
        SESSION.lateralesEnabled = bool(getattr(cv, "card_on_lat", True))

    def raise_colocar(self):
        """Oculta UI y encola ExternalEvent de colocación (hilo Revit API)."""
        self._sync_place_flags_from_ui()
        if not (
            SESSION.placeSup
            or SESSION.placeInf
            or SESSION.placeConf
            or SESSION.lateralesEnabled
        ):
            self.set_status(u"Nada que colocar · active SUP, INF, LAT o CONF.")
            try:
                show_info(
                    u"Todos los paneles están desactivados.",
                    u"Active al menos SUP, INF, LAT o CONF en el rail.",
                    title=_DIALOG,
                    uiapp=self._uiapp,
                )
            except Exception:
                pass
            return
        if not SESSION.framing_elements:
            self.set_status(u"No hay vigas en el lote.")
            try:
                show_info(
                    u"No hay vigas en el lote.",
                    u"No se puede colocar armadura sin vigas seleccionadas.",
                    title=_DIALOG,
                    uiapp=self._uiapp,
                )
            except Exception:
                pass
            return
        parts = []
        if SESSION.placeSup:
            parts.append(u"SUP")
        if SESSION.placeInf:
            parts.append(u"INF")
        if SESSION.placeConf:
            parts.append(u"CONF")
        if SESSION.lateralesEnabled:
            parts.append(u"LAT")
        self.set_status(u"Colocando {0}…".format(u"+".join(parts)))
        try:
            # Hide sync + Raise: Revit no procesa ExternalEvent con WPF maximizado.
            raise_colocar_armadura(self)
        except Exception as ex:
            try:
                msg = unicode(ex)
            except NameError:
                msg = str(ex)
            self.set_status(u"No se pudo iniciar colocación: {0}".format(msg))
            try:
                self.restore_after_colocar()
            except Exception:
                pass
            try:
                show_info(
                    u"No se pudo iniciar la colocación.",
                    msg,
                    title=_DIALOG,
                    uiapp=self._uiapp,
                )
            except Exception:
                pass

    def raise_colocar_stub(self):
        """Compat: redirige al motor de colocación."""
        self.raise_colocar()

    def hide_for_colocar(self):
        """Oculta de inmediato (sync) para liberar idle de Revit."""
        try:
            if self._win is not None:
                self._win.Hide()
        except Exception:
            pass

    def hide_for_colocar_on_ui(self):
        # Preferir sync: el async deja Raise sin procesar.
        self.hide_for_colocar()

    def restore_after_colocar(self):
        def _restore():
            try:
                if self._win is None:
                    return
                self._win.Show()
                try:
                    from System.Windows import WindowState

                    self._win.WindowState = WindowState.Maximized
                except Exception:
                    pass
                self._win.Activate()
            except Exception:
                pass

        # Si ya estamos en UI thread, restaurar sync; si no, Dispatcher.
        try:
            if self._win is not None and self._win.Dispatcher.CheckAccess():
                _restore()
                return
        except Exception:
            pass
        self._dispatch_ui(_restore)

    def set_status(self, text):
        try:
            tb = self._win.FindName(u"TxtEstado") if self._win else None
            if tb is not None:
                tb.Text = text or u""
        except Exception:
            pass

    def _on_closed(self, sender, args):
        try:
            clear_colocar_target(self)
        except Exception:
            pass
        try:
            self._clear_overlay_event.Raise()
        except Exception:
            pass
        _clear_window_ref(self)

    def Show(self):
        if self._win is None:
            return
        try:
            from System import Double
            from System.Windows import SizeToContent, WindowStartupLocation, WindowState
            from revit_wpf_window_position import (
                bind_maximize_wpf_on_revit_monitor,
                revit_main_hwnd,
            )

            hwnd = revit_main_hwnd(self._uiapp)
            # Sin tope MaxWidth del XAML.
            try:
                self._win.MaxWidth = Double.PositiveInfinity
                self._win.MaxHeight = Double.PositiveInfinity
            except Exception:
                pass
            try:
                self._win.SizeToContent = SizeToContent.Manual
            except Exception:
                pass
            try:
                self._win.WindowStartupLocation = WindowStartupLocation.Manual
            except Exception:
                pass
            # Primera apertura en el monitor de Revit; si el usuario restaura y
            # arrastra a otro monitor, el maximizar siguiente usa ese monitor.
            try:
                from revit_wpf_window_position import (
                    _monitor_work_area_px,
                    _primary_work_area_px,
                    preposition_wpf_window_on_work_area,
                )

                area = _monitor_work_area_px(hwnd)
                if area is None:
                    area = _primary_work_area_px()
                if area is not None:
                    preposition_wpf_window_on_work_area(
                        self._win, area[0], area[1], area[2], area[3], hwnd,
                    )
            except Exception:
                pass
            bind_maximize_wpf_on_revit_monitor(self._win, hwnd)
            self._win.WindowState = WindowState.Maximized
        except Exception:
            try:
                from System.Windows import WindowState

                self._win.WindowState = WindowState.Maximized
            except Exception:
                pass
        self._win.Show()


def get_existing_armado_vigas_window():
    return _get_existing_window()


def show_armado_vigas_window(uiapp, pushbutton_dir=None):
    existing = _get_existing_window()
    if existing is not None:
        _focus_existing_window(existing)
        show_info(
            u"La herramienta ya está en ejecución.",
            u"Se enfoca la ventana abierta de Armado vigas.",
            title=_DIALOG,
            uiapp=uiapp,
        )
        return existing

    win = ArmadoVigasWindow(uiapp, pushbutton_dir)
    if win._win is None:
        return None
    win.Show()
    try:
        System.AppDomain.CurrentDomain.SetData(_APP_DOMAIN_KEY, win)
    except Exception:
        pass
    return win
