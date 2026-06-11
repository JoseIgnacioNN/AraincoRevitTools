# -*- coding: utf-8 -*-
"""
Alinea ventanas WPF con la esquina superior izquierda del área de dibujo
de la vista activa de Revit (UIView.GetWindowRectangle), con conversión DPI.

Fallback: CenterScreen si no hay rectángulo o HWND.
"""

import clr

clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")


def revit_main_hwnd(uiapp):
    """HWND de la ventana principal de Revit (IntPtr o compatible)."""
    hwnd = None
    try:
        if uiapp is not None and hasattr(uiapp, "MainWindowHandle"):
            hwnd = uiapp.MainWindowHandle
    except Exception:
        pass
    if not hwnd or (hasattr(hwnd, "ToInt64") and hwnd.ToInt64() == 0):
        try:
            import System.Diagnostics

            proc = System.Diagnostics.Process.GetCurrentProcess()
            hwnd = proc.MainWindowHandle
        except Exception:
            pass
    if hwnd and (not hasattr(hwnd, "ToInt64") or hwnd.ToInt64() != 0):
        return hwnd
    return None


def _get_active_ui_view_screen_rect(uidoc):
    if uidoc is None:
        return None
    try:
        active_id = uidoc.ActiveView.Id
        try:
            active_id_int = int(active_id.Value)
        except Exception:
            try:
                active_id_int = int(active_id.IntegerValue)
            except Exception:
                active_id_int = None
        for uiv in uidoc.GetOpenUIViews():
            try:
                if active_id_int is not None:
                    try:
                        vid = uiv.ViewId
                        try:
                            vid_int = int(vid.Value)
                        except Exception:
                            vid_int = int(vid.IntegerValue)
                        if vid_int == active_id_int:
                            return uiv.GetWindowRectangle()
                    except Exception:
                        pass
                # Fallback: igualdad directa (algunas versiones funcionan bien así)
                if uiv.ViewId == active_id:
                    return uiv.GetWindowRectangle()
            except Exception:
                continue
    except Exception:
        pass
    return None


def _screen_pixels_to_wpf_dip(left_px, top_px, hwnd):
    try:
        # Preferir DPI por-monitor (vista puede estar en otro monitor con escala distinta).
        dpi_x, dpi_y = _get_monitor_dpi_for_point(float(left_px), float(top_px))
        if dpi_x and dpi_y:
            # WPF usa DIP: 96 DIP = 96 px a 100% (96 DPI).
            return (float(left_px) * 96.0 / float(dpi_x)), (float(top_px) * 96.0 / float(dpi_y))

        # Fallback: usar transform del hwnd (normalmente Revit main window).
        from System.Windows import Point
        from System.Windows.Interop import HwndSource

        hs = HwndSource.FromHwnd(hwnd)
        if hs is None or hs.CompositionTarget is None:
            return float(left_px), float(top_px)
        pt = hs.CompositionTarget.TransformFromDevice.Transform(Point(float(left_px), float(top_px)))
        return float(pt.X), float(pt.Y)
    except Exception:
        return float(left_px), float(top_px)


def _get_monitor_dpi_for_point(x_px, y_px):
    """Devuelve (dpiX, dpiY) del monitor que contiene el punto (screen px).

    Usa GetDpiForMonitor (Shcore.dll) cuando está disponible (Windows 8.1+).
    Retorna (None, None) si falla.
    """
    try:
        import System
        from System import IntPtr
        from System.Runtime.InteropServices import Marshal

        # POINT struct (int x, int y)
        class _POINT(System.ValueType):
            _fields_ = [("x", System.Int32), ("y", System.Int32)]

        # MonitorFromPoint via user32
        user32 = System.Runtime.InteropServices.DllImportAttribute("user32.dll")
    except Exception:
        # IronPython no soporta declarar structs así; usar P/Invoke via ctypes si disponible.
        pass

    # Implementación robusta con ctypes (disponible en IronPython 2.7).
    try:
        import ctypes
        from ctypes import wintypes

        user32 = ctypes.WinDLL("user32", use_last_error=True)
        shcore = ctypes.WinDLL("Shcore", use_last_error=True)

        class POINT(ctypes.Structure):
            _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]

        MONITOR_DEFAULTTONEAREST = 2
        MDT_EFFECTIVE_DPI = 0

        MonitorFromPoint = user32.MonitorFromPoint
        MonitorFromPoint.argtypes = [POINT, wintypes.DWORD]
        MonitorFromPoint.restype = wintypes.HMONITOR

        GetDpiForMonitor = shcore.GetDpiForMonitor
        GetDpiForMonitor.argtypes = [wintypes.HMONITOR, wintypes.INT, ctypes.POINTER(wintypes.UINT), ctypes.POINTER(wintypes.UINT)]
        GetDpiForMonitor.restype = wintypes.HRESULT

        pt = POINT(int(round(float(x_px))), int(round(float(y_px))))
        hmon = MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST)
        if not hmon:
            return None, None
        dpi_x = wintypes.UINT()
        dpi_y = wintypes.UINT()
        hr = GetDpiForMonitor(hmon, MDT_EFFECTIVE_DPI, ctypes.byref(dpi_x), ctypes.byref(dpi_y))
        if int(hr) != 0:
            return None, None
        dx = int(dpi_x.value) if int(dpi_x.value) > 0 else None
        dy = int(dpi_y.value) if int(dpi_y.value) > 0 else None
        return dx, dy
    except Exception:
        return None, None


def position_wpf_window_top_left_at_active_view(win, uidoc, hwnd, match_active_view_width=False):
    """Alinea esquina superior izquierda del formulario con la de la vista activa.

    Si ``match_active_view_width`` es True, además fija el ancho del ``Window`` al del
    rectángulo de la vista (UIView) en píxeles de pantalla convertidos a DIP) y usa
    ``SizeToContent.Height`` para que la altura siga al contenido sin recortar.

    Si no hay rectángulo y ``match_active_view_width`` es True, se restaura
    ``SizeToContent.WidthAndHeight`` y se centra en pantalla.
    """
    try:
        from System.Windows import SizeToContent, WindowStartupLocation
    except Exception:
        SizeToContent = None
        WindowStartupLocation = None
    rect = _get_active_ui_view_screen_rect(uidoc)
    if rect is None:
        if match_active_view_width and SizeToContent is not None:
            try:
                win.SizeToContent = SizeToContent.WidthAndHeight
            except Exception:
                pass
        if WindowStartupLocation is not None:
            try:
                win.WindowStartupLocation = WindowStartupLocation.CenterScreen
            except Exception:
                pass
        return
    try:
        left_px = float(rect.Left)
        top_px = float(rect.Top)
        right_px = float(rect.Right)
        bottom_px = float(rect.Bottom)
    except Exception:
        if match_active_view_width and SizeToContent is not None:
            try:
                win.SizeToContent = SizeToContent.WidthAndHeight
            except Exception:
                pass
        if WindowStartupLocation is not None:
            try:
                win.WindowStartupLocation = WindowStartupLocation.CenterScreen
            except Exception:
                pass
        return
    width_px = max(0.0, right_px - left_px)
    # Nota: hwnd puede ser None en algunos contextos; la conversión usa primero DPI por monitor.
    left_dip, top_dip = _screen_pixels_to_wpf_dip(left_px, top_px, hwnd)
    right_dip, bottom_dip = _screen_pixels_to_wpf_dip(right_px, bottom_px, hwnd)
    width_dip = max(0.0, right_dip - left_dip)
    if match_active_view_width and SizeToContent is not None and width_dip > 0.0:
        try:
            win.SizeToContent = SizeToContent.Height
        except Exception:
            pass
        try:
            min_w = float(win.MinWidth)
        except Exception:
            min_w = 0.0
        try:
            from System.Windows import SystemParameters

            max_w = float(SystemParameters.WorkArea.Width) - 16.0
        except Exception:
            max_w = width_dip
        try:
            win.Width = max(min_w, min(width_dip, max_w))
        except Exception:
            pass
    if WindowStartupLocation is not None:
        try:
            win.WindowStartupLocation = WindowStartupLocation.Manual
        except Exception:
            pass
    try:
        win.Left = left_dip
        win.Top = top_dip
    except Exception:
        pass


def position_wpf_window_center_on_active_view(win, uidoc, hwnd, width_dip, height_dip):
    """Centra el rectángulo del formulario (en DIP) en el área de dibujo de la vista activa.

    Usa el mismo rectángulo y conversión DPI que ``position_wpf_window_top_left_at_active_view``.
    Si no hay vista activa o rectángulo, ``CenterScreen``.
    """
    try:
        from System.Windows import WindowStartupLocation
    except Exception:
        WindowStartupLocation = None
    rect = _get_active_ui_view_screen_rect(uidoc)
    if rect is None:
        if WindowStartupLocation is not None:
            try:
                win.WindowStartupLocation = WindowStartupLocation.CenterScreen
            except Exception:
                pass
        return
    try:
        left_px = float(rect.Left)
        top_px = float(rect.Top)
        right_px = float(rect.Right)
        bottom_px = float(rect.Bottom)
    except Exception:
        if WindowStartupLocation is not None:
            try:
                win.WindowStartupLocation = WindowStartupLocation.CenterScreen
            except Exception:
                pass
        return
    left_dip, top_dip = _screen_pixels_to_wpf_dip(left_px, top_px, hwnd)
    right_dip, bottom_dip = _screen_pixels_to_wpf_dip(right_px, bottom_px, hwnd)
    vw = max(0.0, right_dip - left_dip)
    vh = max(0.0, bottom_dip - top_dip)
    cx = left_dip + vw * 0.5
    cy = top_dip + vh * 0.5
    fw = max(1.0, float(width_dip))
    fh = max(1.0, float(height_dip))
    if WindowStartupLocation is not None:
        try:
            win.WindowStartupLocation = WindowStartupLocation.Manual
        except Exception:
            pass
    try:
        win.Left = cx - fw * 0.5
        win.Top = cy - fh * 0.5
    except Exception:
        pass


def _hwnd_to_int(hwnd):
    if hwnd is None:
        return 0
    try:
        return int(hwnd.ToInt64())
    except Exception:
        try:
            return int(hwnd)
        except Exception:
            return 0


def _wpf_window_hwnd(win):
    try:
        from System import IntPtr
        from System.Windows.Interop import WindowInteropHelper

        h = WindowInteropHelper(win).Handle
        if h == IntPtr.Zero:
            return 0
        return _hwnd_to_int(h)
    except Exception:
        return 0


def _primary_work_area_px():
    try:
        import ctypes
        from ctypes import wintypes

        class RECT(ctypes.Structure):
            _fields_ = [
                ("left", wintypes.LONG),
                ("top", wintypes.LONG),
                ("right", wintypes.LONG),
                ("bottom", wintypes.LONG),
            ]

        rect = RECT()
        if not ctypes.windll.user32.SystemParametersInfoW(
            0x0030, 0, ctypes.byref(rect), 0,
        ):
            return None
        return (
            float(rect.left),
            float(rect.top),
            float(rect.right - rect.left),
            float(rect.bottom - rect.top),
        )
    except Exception:
        return None


def _monitor_work_area_px(hwnd):
    if hwnd is not None:
        try:
            import ctypes
            from ctypes import wintypes

            user32 = ctypes.WinDLL("user32", use_last_error=True)
            MONITOR_DEFAULTTONEAREST = 2

            class RECT(ctypes.Structure):
                _fields_ = [
                    ("left", wintypes.LONG),
                    ("top", wintypes.LONG),
                    ("right", wintypes.LONG),
                    ("bottom", wintypes.LONG),
                ]

            class MONITORINFO(ctypes.Structure):
                _fields_ = [
                    ("cbSize", wintypes.DWORD),
                    ("rcMonitor", RECT),
                    ("rcWork", RECT),
                    ("dwFlags", wintypes.DWORD),
                ]

            MonitorFromWindow = user32.MonitorFromWindow
            MonitorFromWindow.argtypes = [wintypes.HWND, wintypes.DWORD]
            MonitorFromWindow.restype = wintypes.HMONITOR
            GetMonitorInfoW = user32.GetMonitorInfoW
            GetMonitorInfoW.argtypes = [ctypes.c_void_p, ctypes.POINTER(MONITORINFO)]
            GetMonitorInfoW.restype = wintypes.BOOL

            h = _hwnd_to_int(hwnd)
            if h:
                hmon = MonitorFromWindow(h, MONITOR_DEFAULTTONEAREST)
                if hmon:
                    mi = MONITORINFO()
                    mi.cbSize = ctypes.sizeof(MONITORINFO)
                    if GetMonitorInfoW(hmon, ctypes.byref(mi)):
                        rc = mi.rcWork
                        return (
                            float(rc.left),
                            float(rc.top),
                            float(rc.right - rc.left),
                            float(rc.bottom - rc.top),
                        )
        except Exception:
            pass
    return None


def _work_area_dip_rect(left_px, top_px, width_px, height_px, hwnd_ref=None):
    left_dip, top_dip = _screen_pixels_to_wpf_dip(left_px, top_px, hwnd_ref)
    right_dip, bottom_dip = _screen_pixels_to_wpf_dip(
        left_px + width_px, top_px + height_px, hwnd_ref,
    )
    return (
        left_dip,
        top_dip,
        max(1.0, right_dip - left_dip),
        max(1.0, bottom_dip - top_dip),
    )


def position_wpf_window_center_work_area(win):
    try:
        from System.Windows import SystemParameters, WindowStartupLocation

        wa = SystemParameters.WorkArea
        win.WindowStartupLocation = WindowStartupLocation.Manual
        w = float(win.Width)
        h = float(win.Height)
        if w <= 1.0:
            try:
                fw = float(win.ActualWidth)
                if fw > 1.0:
                    w = fw
            except Exception:
                pass
        if h <= 1.0:
            try:
                fh = float(win.ActualHeight)
                if fh > 1.0:
                    h = fh
            except Exception:
                pass
        wa_left = float(wa.Left)
        wa_top = float(wa.Top)
        wa_width = float(wa.Width)
        wa_height = float(wa.Height)
        left = wa_left + (wa_width - w) / 2.0
        top = wa_top + (wa_height - h) / 2.0
        if left < wa_left:
            left = wa_left
        if top < wa_top:
            top = wa_top
        max_left = wa_left + wa_width - w
        max_top = wa_top + wa_height - h
        if max_left < wa_left:
            max_left = wa_left
        if max_top < wa_top:
            max_top = wa_top
        if left > max_left:
            left = max_left
        if top > max_top:
            top = max_top
        win.Left = left
        win.Top = top
    except Exception:
        pass


def position_wpf_window_center_on_monitor(win, hwnd=None):
    if win is None:
        return False
    area = _monitor_work_area_px(hwnd)
    if area is None:
        area = _primary_work_area_px()
    if area is None:
        position_wpf_window_center_work_area(win)
        return False
    left_px, top_px, width_px, height_px = area
    left_dip, top_dip, wa_w_dip, wa_h_dip = _work_area_dip_rect(
        left_px, top_px, width_px, height_px, hwnd,
    )
    try:
        from System.Windows import WindowStartupLocation

        w = float(win.Width)
        h = float(win.Height)
        if w <= 1.0:
            try:
                fw = float(win.ActualWidth)
                if fw > 1.0:
                    w = fw
            except Exception:
                w = 400.0
        if h <= 1.0:
            try:
                fh = float(win.ActualHeight)
                if fh > 1.0:
                    h = fh
            except Exception:
                h = 180.0
        win.WindowStartupLocation = WindowStartupLocation.Manual
        left = left_dip + max(0.0, (wa_w_dip - w) * 0.5)
        top = top_dip + max(0.0, (wa_h_dip - h) * 0.5)
        max_left = left_dip + max(0.0, wa_w_dip - w)
        max_top = top_dip + max(0.0, wa_h_dip - h)
        win.Left = min(left, max_left)
        win.Top = min(top, max_top)
        return True
    except Exception:
        return False


def bind_center_wpf_on_revit_monitor(win, hwnd_revit=None):
    if win is None:
        return

    def _apply(sender, args):
        position_wpf_window_center_on_monitor(win, hwnd_revit)

    try:
        from System.Windows import RoutedEventHandler

        h = RoutedEventHandler(_apply)
        win.Loaded += h
        try:
            win.ContentRendered += h
        except Exception:
            pass
    except Exception:
        position_wpf_window_center_on_monitor(win, hwnd_revit)


def _revit_monitor_work_area(hwnd_revit=None):
    area = _monitor_work_area_px(hwnd_revit)
    if area is None:
        area = _primary_work_area_px()
    return area


def _fill_wpf_window_work_area_wpf(win, left_px, top_px, width_px, height_px, hwnd_ref=None):
    if win is None:
        return False
    try:
        from System.Windows import WindowStartupLocation, WindowState

        try:
            from System import Double

            win.MaxWidth = Double.PositiveInfinity
            win.MaxHeight = Double.PositiveInfinity
        except Exception:
            pass
        left_dip, top_dip, w_dip, h_dip = _work_area_dip_rect(
            left_px, top_px, width_px, height_px, hwnd_ref,
        )
        win.WindowStartupLocation = WindowStartupLocation.Manual
        win.WindowState = WindowState.Normal
        win.Left = left_dip
        win.Top = top_dip
        win.Width = w_dip
        win.Height = h_dip
        return True
    except Exception:
        return False


def bind_maximize_wpf_on_revit_monitor(win, hwnd_revit=None):
    if win is None:
        return False

    def _on_state_changed(sender, args):
        try:
            from System import Double
            from System.Windows import WindowStartupLocation, WindowState
        except Exception:
            return
        try:
            if win.WindowState != WindowState.Maximized:
                return
        except Exception:
            return
        area = _revit_monitor_work_area(hwnd_revit)
        if area is None:
            return
        left_px, top_px, width_px, height_px = area
        try:
            win.MaxWidth = Double.PositiveInfinity
            win.MaxHeight = Double.PositiveInfinity
            win.WindowStartupLocation = WindowStartupLocation.Manual
        except Exception:
            pass
        wh = _wpf_window_hwnd(win)
        if wh:
            try:
                import ctypes

                user32 = ctypes.windll.user32
                SWP_NOZORDER = 0x0004
                if user32.SetWindowPos(
                    wh,
                    0,
                    int(round(left_px)),
                    int(round(top_px)),
                    max(320, int(round(width_px))),
                    max(200, int(round(height_px))),
                    SWP_NOZORDER,
                ):
                    return
            except Exception:
                pass
        if _fill_wpf_window_work_area_wpf(
            win, left_px, top_px, width_px, height_px, hwnd_revit,
        ):
            try:
                win.WindowState = WindowState.Maximized
            except Exception:
                pass

    try:
        from System import EventHandler

        win.StateChanged += EventHandler(_on_state_changed)
        return True
    except Exception:
        return False
