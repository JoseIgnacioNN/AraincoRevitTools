# -*- coding: utf-8 -*-
"""
Utilidades de integración WPF para la capa ui/.

Carga XAML con rutas relativas a la carpeta ui/ del .pushbutton,
gestiona singletons de ventana y posicionado de ventanas WPF.

REGLA DE CAPA:
- No toca Revit API DB (solo AppDomain y UI de Windows).
- No importa módulos de core/ ni creators/.
"""
import os
import sys

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System import AppDomain
from System.IO import StreamReader, StringReader
from System.Windows import Application, Window
from System.Windows.Markup import XamlReader

try:
    from Autodesk.Revit.UI import TaskDialog
except Exception:
    TaskDialog = None

# Directorio de este archivo (siempre = <pushbutton>/ui/)
_UI_DIR = os.path.dirname(os.path.abspath(__file__))


# ---------------------------------------------------------------------------
# Carga de XAML
# ---------------------------------------------------------------------------

def load_xaml_from_ui_folder(filename):
    """
    Carga un .xaml desde la carpeta ui/ del .pushbutton.

    ``filename``: nombre de archivo (p.ej. "wizard_paso1_rejilla.xaml").
    Devuelve el objeto raíz WPF parseado, o lanza Exception si falla.
    """
    path = os.path.join(_UI_DIR, filename)
    if not os.path.isfile(path):
        raise IOError(u"XAML no encontrado: {}".format(path))
    with open(path, "r") as f:
        xaml_text = f.read()
    try:
        return XamlReader.Parse(xaml_text)
    except Exception:
        sr = StringReader(xaml_text)
        return XamlReader.Load(sr)


# ---------------------------------------------------------------------------
# Tema oscuro
# ---------------------------------------------------------------------------

def load_dark_theme_xml():
    """Devuelve el string XAML del tema oscuro o None si no existe."""
    path = os.path.join(_UI_DIR, "dark_theme.xaml")
    if not os.path.isfile(path):
        return None
    try:
        with open(path, "r") as f:
            return f.read()
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Singleton de ventana por clave AppDomain
# ---------------------------------------------------------------------------

def get_singleton_window(key):
    """
    Devuelve la instancia WPF activa con la clave dada, o None si no existe
    (o fue garbage-collected).
    """
    try:
        w = AppDomain.CurrentDomain.GetData(key)
        if w is None:
            return None
        if not getattr(w, "IsLoaded", True):
            AppDomain.CurrentDomain.SetData(key, None)
            return None
        return w
    except Exception:
        return None


def register_singleton_window(key, window):
    """
    Registra ``window`` con la clave ``key`` en el AppDomain y conecta
    un handler Closed para limpiar la referencia automáticamente.
    """
    AppDomain.CurrentDomain.SetData(key, window)

    def _on_closed(s, e):
        try:
            cur = AppDomain.CurrentDomain.GetData(key)
            if cur is window:
                AppDomain.CurrentDomain.SetData(key, None)
        except Exception:
            pass

    try:
        window.Closed += _on_closed
    except Exception:
        pass


def show_singleton_or_new(key, factory_fn):
    """
    Si ya existe una ventana activa con ``key``, la activa y muestra TaskDialog.
    En caso contrario llama ``factory_fn()`` para crear la ventana, la registra
    y la muestra (ShowDialog).

    Devuelve el resultado de ShowDialog, o None si ya estaba abierta.
    """
    existing = get_singleton_window(key)
    if existing is not None:
        try:
            existing.Activate()
            if hasattr(existing, "WindowState"):
                from System.Windows import WindowState
                if existing.WindowState == WindowState.Minimized:
                    existing.WindowState = WindowState.Normal
        except Exception:
            pass
        if TaskDialog is not None:
            TaskDialog.Show(u"Armado Columnas", u"La herramienta ya está en ejecución.")
        return None

    window = factory_fn()
    register_singleton_window(key, window)
    return window.ShowDialog()


# ---------------------------------------------------------------------------
# Posicionado de ventana
# ---------------------------------------------------------------------------

def center_window_on_screen(window):
    """Centra la ventana en la pantalla principal."""
    try:
        from System.Windows import SystemParameters, WindowStartupLocation
        window.WindowStartupLocation = WindowStartupLocation.CenterScreen
    except Exception:
        pass


def position_window_near_cursor(window):
    """Intenta posicionar la ventana cerca del cursor del ratón."""
    try:
        from System.Windows import SystemParameters
        from System.Windows.Forms import Cursor
        pos = Cursor.Position
        screen_w = float(SystemParameters.PrimaryScreenWidth)
        screen_h = float(SystemParameters.PrimaryScreenHeight)
        win_w = float(window.Width) if window.Width > 0 else 800.0
        win_h = float(window.Height) if window.Height > 0 else 600.0
        left = min(float(pos.X), screen_w - win_w)
        top  = min(float(pos.Y), screen_h - win_h)
        window.Left = max(0.0, left)
        window.Top  = max(0.0, top)
    except Exception:
        center_window_on_screen(window)
