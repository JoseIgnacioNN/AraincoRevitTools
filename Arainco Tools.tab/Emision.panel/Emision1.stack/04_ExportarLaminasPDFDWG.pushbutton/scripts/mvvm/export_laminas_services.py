# -*- coding: utf-8 -*-
"""
Capa de servicios para la herramienta Exportar Láminas.

BloquearComandosRevit – context manager: EnableWindow(False/True) sobre la ventana de Revit
                       (copia autocontenida; no usa join_geometry_concrete_vista).
RevitWindowService   – bloquea la ventana de Revit y muestra diálogos WPF BIMTools.
FolderBrowserService – envuelve FolderBrowserDialog de WinForms con propietario Win32.
ProgressService      – gestiona barras de progreso de pyRevit por fases de exportación.
"""

import os

import clr  # noqa: E402

clr.AddReference("System.Windows.Forms")

from ui.export_laminas_instruction_dialog import (  # noqa: E402
    show_message_dialog,
    show_ok_cancel_dialog,
)

_TASK_TITLE = u"Arainco: Exportar Láminas"


# ---------------------------------------------------------------------------
# Bloqueo de la ventana principal de Revit (autocontenido en el botón)
# ---------------------------------------------------------------------------

def _hwnd_to_int(hwnd):
    if hwnd is None:
        return None
    try:
        if hasattr(hwnd, u"ToInt32"):
            return int(hwnd.ToInt32())
    except Exception:
        pass
    try:
        if hasattr(hwnd, u"ToInt64"):
            return int(hwnd.ToInt64())
    except Exception:
        pass
    try:
        return int(hwnd)
    except Exception:
        return None


def _revit_main_window_set_enabled(revit, enable):
    """Habilita o deshabilita la ventana principal de Revit (EnableWindow Win32)."""
    if revit is None:
        return
    try:
        from infra.revit_wpf_window_position import revit_main_hwnd
    except Exception:
        return
    hwnd = revit_main_hwnd(revit)
    h = _hwnd_to_int(hwnd)
    if h is None or h == 0:
        return
    try:
        import ctypes
    except Exception:
        return
    try:
        ctypes.windll.user32.EnableWindow(h, 1 if enable else 0)
    except Exception:
        pass


class BloquearComandosRevit(object):
    """
    Context manager: deshabilita la ventana principal de Revit y la restaura al salir.
    No depende de otros módulos de la extensión.
    """

    def __init__(self, revit):
        self._revit = revit
        self._touched = False

    def __enter__(self):
        _revit_main_window_set_enabled(self._revit, False)
        self._touched = True
        return self

    def __exit__(self, _exc_type, _exc, _tb):
        if self._touched:
            _revit_main_window_set_enabled(self._revit, True)
        self._touched = False
        return False


# Alias con prefijo (compatibilidad con script.py histórico)
_BloquearComandosRevit = BloquearComandosRevit


# ---------------------------------------------------------------------------
# RevitWindowService
# ---------------------------------------------------------------------------

class RevitWindowService(object):
    """
    Abstrae interacciones con la ventana principal de Revit:
    – Bloqueo durante exportación (_BloquearComandosRevit).
    – Presentación de diálogos WPF BIMTools (mensajes y confirmaciones).
    """

    def __init__(self, revit, bloquear_cls):
        """
        revit        – referencia a __revit__ (para _BloquearComandosRevit).
        bloquear_cls – clase :class:`BloquearComandosRevit` de este módulo (puede ser None).
        """
        self._revit = revit
        self._bloquear_cls = bloquear_cls
        self._blocker = None

    # -- Bloqueo de la ventana de Revit ------------------------------------

    def block_revit(self):
        """Inhabilita la ventana principal de Revit si _BloquearComandosRevit está disponible."""
        if self._blocker is not None or self._bloquear_cls is None:
            return
        try:
            b = self._bloquear_cls(self._revit)
            b.__enter__()
            self._blocker = b
        except Exception:
            self._blocker = None

    def unblock_revit(self):
        """Restaura la ventana principal de Revit."""
        b = self._blocker
        self._blocker = None
        if b is not None:
            try:
                b.__exit__(None, None, None)
            except Exception:
                pass

    # -- Diálogos WPF BIMTools --------------------------------------------

    def _revit_dialog_context(self):
        uiapp = None
        hwnd = None
        try:
            from infra.revit_wpf_window_position import revit_main_hwnd

            revit = self._revit
            try:
                uiapp = revit.Application if revit is not None else None
            except Exception:
                uiapp = None
            if uiapp is None:
                uiapp = revit
            hwnd = revit_main_hwnd(uiapp)
        except Exception:
            pass
        return uiapp, hwnd

    @staticmethod
    def run_above_wpf(callback, wpf_win=None):
        """Ejecuta callback; baja Topmost de la ventana WPF si aplica."""
        top = None
        if wpf_win is not None:
            try:
                top = wpf_win.Topmost
                wpf_win.Topmost = False
            except Exception:
                top = None
        try:
            callback()
        finally:
            if wpf_win is not None and top is not None:
                try:
                    wpf_win.Topmost = top
                except Exception:
                    pass

    def show_ok(self, main_instruction, wpf_win=None):
        """Diálogo informativo (solo Aceptar), estilo BIMTools."""
        uiapp, hwnd = self._revit_dialog_context()

        def _cb():
            show_message_dialog(
                _TASK_TITLE,
                main_instruction,
                u"",
                ok_text=u"Entendido",
                hwnd_revit=hwnd,
                uiapp=uiapp,
            )

        self.run_above_wpf(_cb, wpf_win)

    def show_errors(self, main_instruction, errors, wpf_win=None):
        """Diálogo con lista de errores (máximo 20 líneas), estilo BIMTools."""
        uiapp, hwnd = self._revit_dialog_context()
        err_txt = u"\n".join(errors[:20])
        if len(errors) > 20:
            err_txt += u"\n…"

        def _cb():
            show_message_dialog(
                _TASK_TITLE,
                main_instruction,
                err_txt,
                ok_text=u"Entendido",
                hwnd_revit=hwnd,
                uiapp=uiapp,
            )

        self.run_above_wpf(_cb, wpf_win)

    def ask_yes_no(self, main_instruction, content, wpf_win=None):
        """Diálogo Sí/No; devuelve True si el usuario confirma."""
        uiapp, hwnd = self._revit_dialog_context()
        result = [False]

        def _cb():
            result[0] = show_ok_cancel_dialog(
                _TASK_TITLE,
                main_instruction,
                content,
                ok_text=u"Sí",
                cancel_text=u"No",
                hwnd_revit=hwnd,
                uiapp=uiapp,
            )

        self.run_above_wpf(_cb, wpf_win)
        return result[0]


# ---------------------------------------------------------------------------
# FolderBrowserService
# ---------------------------------------------------------------------------

class FolderBrowserService(object):
    """
    Abre un FolderBrowserDialog de WinForms modalizado bajo la ventana WPF
    (propietario Win32 resuelto automáticamente).
    """

    def __init__(self, revit_application):
        """
        revit_application – Application de Revit (__revit__.Application),
                            usado para obtener el HWND de la ventana principal.
        """
        self._app = revit_application

    def browse(self, current_path=u"", wpf_win=None):
        """
        Muestra el diálogo de selección de carpeta.

        Devuelve la ruta elegida como str, o None si se cancela.
        """
        from System.Windows.Forms import FolderBrowserDialog, NativeWindow
        from infra.revit_wpf_window_position import revit_main_hwnd

        class _Win32FolderOwner(NativeWindow):
            """IWin32Window anónimo para modalizar el diálogo bajo la ventana WPF."""
            pass

        dlg = FolderBrowserDialog()
        dlg.Description = (
            u"Carpeta base para la entrega "
            u"(se añade la subcarpeta de fecha si aplica)."
        )

        # Pre-seleccionar carpeta actual o su padre
        sel = u""
        if current_path:
            try:
                if os.path.isdir(current_path):
                    sel = current_path
                else:
                    par = os.path.dirname(current_path.rstrip(u"\\/"))
                    if par and os.path.isdir(par):
                        sel = par
            except Exception:
                sel = u""
        if sel:
            try:
                dlg.SelectedPath = sel
            except Exception:
                pass

        # Resolver propietario Win32 (Revit → WPF Tool Win)
        owner_wrap = None
        top_prev = None
        if wpf_win is not None:
            try:
                top_prev = wpf_win.Topmost
                wpf_win.Topmost = False
            except Exception:
                top_prev = None
        try:
            _r_hwnd = None
            try:
                _r_hwnd = revit_main_hwnd(self._app)
            except Exception:
                pass
            if _r_hwnd is not None:
                try:
                    if _r_hwnd.ToInt64() != 0:
                        owner_wrap = _Win32FolderOwner()
                        owner_wrap.AssignHandle(_r_hwnd)
                except Exception:
                    owner_wrap = None
            if owner_wrap is None and wpf_win is not None:
                try:
                    from System.Windows.Interop import WindowInteropHelper
                    hwnd = WindowInteropHelper(wpf_win).Handle
                    if hwnd.ToInt64() != 0:
                        owner_wrap = _Win32FolderOwner()
                        owner_wrap.AssignHandle(hwnd)
                except Exception:
                    owner_wrap = None
        except Exception:
            owner_wrap = None

        try:
            dr = (
                dlg.ShowDialog(owner_wrap)
                if owner_wrap is not None
                else dlg.ShowDialog()
            )
        finally:
            if owner_wrap is not None:
                try:
                    owner_wrap.ReleaseHandle()
                except Exception:
                    pass
            if wpf_win is not None and top_prev is not None:
                try:
                    wpf_win.Topmost = top_prev
                except Exception:
                    pass

        if not self._dialog_accepted(dr):
            return None
        try:
            base = unicode(dlg.SelectedPath).strip()
        except Exception:
            base = u""
        return base if base else None

    @staticmethod
    def _dialog_accepted(dialog_result):
        from System.Windows.Forms import DialogResult
        try:
            if dialog_result == DialogResult.OK:
                return True
        except Exception:
            pass
        try:
            if int(dialog_result) == int(DialogResult.OK):
                return True
        except Exception:
            pass
        try:
            return unicode(str(dialog_result)).strip().upper() == u"OK"
        except Exception:
            return False


# ---------------------------------------------------------------------------
# ProgressService
# ---------------------------------------------------------------------------

class ProgressService(object):
    """
    Gestiona la barra de progreso de pyRevit para las fases de exportación.
    Cada fase (DWG / PDF / Listado) se activa con begin_phase, se avanza con
    step y se cierra con end_phase (o automáticamente al comenzar la siguiente).
    """

    _ACCENT_RGB = (91, 192, 222)

    def __init__(self, pyrevit_forms):
        """pyrevit_forms – módulo pyrevit.forms (puede ser None)."""
        self._forms = pyrevit_forms
        self._pb = None
        self._pb_active = False

    @staticmethod
    def phase_title(base_title, total):
        """Título inicial de fase: «base 0/N»."""
        try:
            t = max(int(total), 1)
        except Exception:
            t = 1
        return u"{} 0/{}".format(base_title, t)

    def begin_phase(self, title, count):
        """Finaliza la fase activa (si existe) e inicia una nueva."""
        self._end_current()
        if self._forms is None:
            return
        try:
            c = int(count)
        except Exception:
            c = 0
        if c < 1:
            return
        try:
            pb = self._forms.ProgressBar(title=title, cancellable=False)
            try:
                from System.Windows.Media import Color, SolidColorBrush
                r, g, b = self._ACCENT_RGB
                pb.Resources[u"pyRevitAccentBrush"] = SolidColorBrush(
                    Color.FromRgb(r, g, b)
                )
            except Exception:
                pass
            pb.__enter__()
            self._pb = pb
            self._pb_active = True
        except Exception:
            self._pb = None
            self._pb_active = False

    def step(self, step_index, total, base_title):
        """Avanza un paso dentro de la fase activa."""
        if not self._pb_active or self._pb is None:
            return
        try:
            c = max(int(total), 1)
            i = int(step_index) + 1
            try:
                self._pb.update_progress(i, max_value=c)
            except TypeError:
                try:
                    self._pb.update_progress(i, max=c)
                except Exception:
                    pass
            except Exception:
                pass
            try:
                self._pb.title = u"{} {}/{}".format(base_title, i, c)
            except Exception:
                pass
        except Exception:
            pass

    def end_phase(self):
        """Cierra la fase activa."""
        self._end_current()

    def _end_current(self):
        if self._pb_active and self._pb is not None:
            try:
                self._pb.__exit__(None, None, None)
            except Exception:
                pass
        self._pb = None
        self._pb_active = False
