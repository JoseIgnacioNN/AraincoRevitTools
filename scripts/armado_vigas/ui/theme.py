# -*- coding: utf-8 -*-
"""Tokens visuales homogeneidad — mockup opcion_d_homogeneidad."""

from System.Windows import FontWeights, Thickness
from System.Windows.Controls import Border, TextBlock
from System.Windows.Media import SolidColorBrush, Color

from armado_vigas.ui import typography as typo


def _brush_hex(hx, alpha=255):
    h = (hx or u"#64748b").strip().lstrip(u"#")
    if len(h) < 6:
        h = u"64748b"
    rr = int(h[0:2], 16)
    gg = int(h[2:4], 16)
    bb = int(h[4:6], 16)
    aa = max(0, min(255, int(alpha)))
    return SolidColorBrush(Color.FromArgb(aa, rr, gg, bb))


TRAMO_SOFT_HALO = 31

# ── Surfaces ──
BG_APP = u"#071018"
BG_PANEL = u"#0a1620"
BG_INPUT = u"#071018"

# ── Borders ──
BORDER = u"#21465C"
BORDER_MUTED = u"#2d4455"
BORDER_INPUT = u"#1e3344"

# ── Typography colors ──
FG_HI = u"#e8f4f8"
FG_MID = u"#95b8cc"
FG_LO = u"#64748b"

# ── Primary accent (selección / foco) ──
ACCENT = u"#22d3ee"

# ── Semantic (badges) ──
SEM_EMPALME = u"#fbbf24"
SEM_SUPLE = u"#a78bfa"
SEM_CENT = u"#6ee7b7"
SEM_EXT = u"#fcd34d"
SEM_CONFIN = u"#7dd3fc"

_SEM_BADGE = {
    u"suple": (SEM_SUPLE, 26),
    u"empalme": (SEM_EMPALME, 31),
    u"cent": (SEM_CENT, 31),
    u"ext": (SEM_EXT, 31),
    u"confin": (SEM_CONFIN, 31),
}


def brush_panel(alpha=255):
    return _brush_hex(BG_PANEL, alpha)


def brush_app(alpha=255):
    return _brush_hex(BG_APP, alpha)


def brush_input(alpha=255):
    return _brush_hex(BG_INPUT, alpha)


def brush_border(alpha=255):
    return _brush_hex(BORDER, alpha)


def brush_border_muted(alpha=255):
    return _brush_hex(BORDER_MUTED, alpha)


def brush_border_input(alpha=255):
    return _brush_hex(BORDER_INPUT, alpha)


def brush_accent(alpha=255):
    return _brush_hex(ACCENT, alpha)


def brush_fg_hi(alpha=255):
    return _brush_hex(FG_HI, alpha)


def brush_fg_mid(alpha=255):
    return _brush_hex(FG_MID, alpha)


def brush_fg_lo(alpha=255):
    return _brush_hex(FG_LO, alpha)


def selection_border_brush(selected):
    return brush_accent() if selected else brush_border(115)


def selection_background_brush(selected):
    return brush_accent(TRAMO_SOFT_HALO) if selected else brush_panel(0)


def brush_sem(hex_color, alpha=255):
    return _brush_hex(hex_color, alpha)


def apply_panel_chrome(border, selected=False, padding=8):
    """Chrome unificado panel-card / slot neutro."""
    border.Background = brush_panel()
    border.BorderBrush = selection_border_brush(selected)
    border.BorderThickness = Thickness(1)
    border.Padding = Thickness(float(padding))
    if selected:
        border.Background = selection_background_brush(True)


def make_role_badge(label, role=u"suple"):
    """Badge semántico pequeño (Suple, Empalme, Cant…)."""
    color, bg_alpha = _SEM_BADGE.get(role, (ACCENT, 31))
    badge = Border()
    badge.Padding = Thickness(4, 1, 4, 1)
    badge.Background = _brush_hex(color, bg_alpha)
    badge.BorderBrush = _brush_hex(color, 107)
    badge.BorderThickness = Thickness(1)
    try:
        from System.Windows import CornerRadius
        badge.CornerRadius = CornerRadius(3.0)
    except Exception:
        pass
    tb = TextBlock()
    tb.Text = (label or u"").upper()
    tb.FontSize = typo.META_FONT_PX
    tb.FontWeight = FontWeights.Bold
    tb.Foreground = _brush_hex(color)
    badge.Child = tb
    return badge
