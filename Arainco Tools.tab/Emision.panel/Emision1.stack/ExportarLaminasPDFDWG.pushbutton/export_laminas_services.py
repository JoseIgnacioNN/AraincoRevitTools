# -*- coding: utf-8 -*-
"""
Capa de servicios para la herramienta Exportar Láminas.

BloquearComandosRevit – context manager: EnableWindow(False/True) sobre la ventana de Revit
                       (copia autocontenida; no usa join_geometry_concrete_vista).
RevitWindowService   – bloquea la ventana de Revit y muestra TaskDialogs sobre la
                       ventana WPF (Topmost).
FolderBrowserService – envuelve FolderBrowserDialog de WinForms con propietario Win32.
ProgressService      – gestiona barras de progreso de pyRevit por fases de exportación.
"""

import os
import sys

_pb = os.path.dirname(os.path.abspath(__file__))
if _pb not in sys.path:
    sys.path.insert(0, _pb)

import clr  # noqa: E402

clr.AddReference("RevitAPIUI")
clr.AddReference("System.Windows.Forms")

from Autodesk.Revit.UI import (  # noqa: E402
    TaskDialog,
    TaskDialogCommonButtons,
    TaskDialogResult,
)

_TASK_TITLE = u"Exportar Láminas"


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
        from revit_wpf_window_position import revit_main_hwnd
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
    – Presentación de TaskDialogs por encima de la ventana WPF Topmost.
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

    # -- TaskDialogs sobre la ventana WPF ---------------------------------

    @staticmethod
    def run_above_wpf(callback, wpf_win=None):
        """
        Baja Topmost de la ventana WPF, ejecuta callback y lo restaura.
        Garantiza que los TaskDialog de Revit queden visibles sobre la ventana.
        """
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
        """TaskDialog informativo (botón OK) mostrado por encima de la ventana WPF."""
        def _cb():
            td = TaskDialog(_TASK_TITLE)
            try:
                td.TitleAutoPrefix = False
            except Exception:
                pass
            td.MainInstruction = main_instruction
            td.CommonButtons = TaskDialogCommonButtons.Ok
            td.DefaultButton = TaskDialogResult.Ok
            td.Show()
        self.run_above_wpf(_cb, wpf_win)

    def show_errors(self, main_instruction, errors, wpf_win=None):
        """TaskDialog con lista de errores (máximo 20 líneas)."""
        def _cb():
            err_txt = u"\n".join(errors[:20])
            if len(errors) > 20:
                err_txt += u"\n…"
            td = TaskDialog(_TASK_TITLE)
            try:
                td.TitleAutoPrefix = False
            except Exception:
                pass
            td.MainInstruction = main_instruction
            td.MainContent = err_txt
            td.CommonButtons = TaskDialogCommonButtons.Ok
            td.DefaultButton = TaskDialogResult.Ok
            td.Show()
        self.run_above_wpf(_cb, wpf_win)

    def ask_yes_no(self, main_instruction, content, wpf_win=None):
        """TaskDialog Sí/No; devuelve True si el usuario elige Sí."""
        result = [False]

        def _cb():
            td = TaskDialog(_TASK_TITLE)
            try:
                td.TitleAutoPrefix = False
            except Exception:
                pass
            td.MainInstruction = main_instruction
            td.MainContent = content
            td.CommonButtons = (
                TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
            )
            td.DefaultButton = TaskDialogResult.Yes
            result[0] = td.Show() == TaskDialogResult.Yes

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
        from revit_wpf_window_position import revit_main_hwnd

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
