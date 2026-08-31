# -*- coding: utf-8 -*-
"""Pinceles WPF mínimos (IronPython)."""

import clr

clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows.Media import SolidColorBrush, Color


def brush_hex(hx, alpha=255):
    h = (hx or u"#64748b").strip().lstrip(u"#")
    if len(h) < 6:
        h = u"64748b"
    rr = int(h[0:2], 16)
    gg = int(h[2:4], 16)
    bb = int(h[4:6], 16)
    aa = max(0, min(255, int(alpha)))
    return SolidColorBrush(Color.FromArgb(aa, rr, gg, bb))
