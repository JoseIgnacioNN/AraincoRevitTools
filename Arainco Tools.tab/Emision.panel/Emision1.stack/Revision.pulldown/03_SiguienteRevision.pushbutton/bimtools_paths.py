# -*- coding: utf-8 -*-
"""
Rutas de logo corporativo para WPF — versión empaquetada en el pushbutton.

Prioridad:
  1) Carpeta del pushbutton registrada con set_pushbutton_dir (desde script.py)
  2) <raíz de extensión *.extension>/assets/ y /branding/
  3) icon.png al final como último recurso visual

Este archivo vive junto a script.py del botón Revisiones; al copiar la carpeta
.pushbutton a otra extensión, coloque logo.png / empresa_logo.png / icon.png aquí.
"""

import os

_LOGO_NAMES = ("empresa_logo.png", "logo_empresa.png", "logo.png")
_pushbutton_dir = None


def set_pushbutton_dir(path):
    global _pushbutton_dir
    if path and os.path.isdir(os.path.normpath(path)):
        _pushbutton_dir = os.path.normpath(os.path.abspath(path))
    else:
        _pushbutton_dir = None


def get_pushbutton_dir():
    return _pushbutton_dir


def default_extension_root():
    """
    Raíz *.extension: sube desde la carpeta de este módulo (pushbutton).
    """
    here = os.path.dirname(os.path.abspath(__file__))
    d = here
    for _ in range(24):
        base = os.path.basename(d)
        try:
            if base.lower().endswith(".extension"):
                return d
        except Exception:
            pass
        parent = os.path.dirname(d)
        if parent == d:
            break
        d = parent
    return os.path.dirname(here)


def get_logo_paths(extension_root=None):
    if extension_root is None:
        extension_root = default_extension_root()
    else:
        extension_root = os.path.normpath(os.path.abspath(extension_root))

    out = []
    seen = set()
    icon_last = []
    icon_seen = set()

    def _add(p):
        p = os.path.normpath(os.path.abspath(p))
        if p not in seen:
            seen.add(p)
            out.append(p)

    def _queue_icon(p):
        p = os.path.normpath(os.path.abspath(p))
        if p not in icon_seen:
            icon_seen.add(p)
            icon_last.append(p)

    if _pushbutton_dir:
        for name in _LOGO_NAMES:
            _add(os.path.join(_pushbutton_dir, name))
        _queue_icon(os.path.join(_pushbutton_dir, "icon.png"))
    for sub in ("assets", "branding"):
        base = os.path.join(extension_root, sub)
        for name in _LOGO_NAMES:
            _add(os.path.join(base, name))
        _queue_icon(os.path.join(base, "icon.png"))

    for p in icon_last:
        _add(p)
    return out


def load_logo_bitmap_image():
    try:
        import clr

        clr.AddReference("PresentationCore")
        clr.AddReference("System")
        from System.IO import FileAccess, FileMode, FileStream
        from System.Windows.Media.Imaging import BitmapCacheOption, BitmapImage
    except Exception:
        return None

    for path in get_logo_paths():
        if not path or not os.path.isfile(path):
            continue
        stream = None
        try:
            stream = FileStream(path, FileMode.Open, FileAccess.Read)
            bmp = BitmapImage()
            bmp.BeginInit()
            bmp.StreamSource = stream
            bmp.CacheOption = BitmapCacheOption.OnLoad
            bmp.EndInit()
            bmp.Freeze()
            return bmp
        except Exception:
            continue
        finally:
            if stream is not None:
                try:
                    stream.Dispose()
                except Exception:
                    pass
    return None
