# -*- coding: utf-8 -*-
"""
Tokens visuales por defecto — BIMTools / Arainco.

Estilo general escalable para **herramientas nuevas** con formulario WPF.
No aplicar retroactivamente a UIs existentes sin armonización explícita.

La cinta blanca superior es la barra nativa de Windows (Window estándar, sin
``WindowStyle="None"``). El cuerpo de la ventana usa el tema oscuro.

Implementación: ``bimtools_wpf_shell.py`` · Regla: ``.cursor/rules/bimtools-ui.mdc``
"""

# ── Tipografía ────────────────────────────────────────────────────────────────

FONT_FAMILY = u"Segoe UI"
FONT_SIZE_BASE = 12
FONT_SIZE_TITLE = 18
FONT_SIZE_SUBTITLE = 11
FONT_SIZE_BODY = 11
FONT_SIZE_HINT = 10
FONT_SIZE_STATUS = 10

FONT_WEIGHT_TITLE = u"Bold"
FONT_WEIGHT_LABEL = u"SemiBold"

# ── Superficies ───────────────────────────────────────────────────────────────

BG_APP = u"#071018"
BG_PANEL = u"#0a1620"
BG_PANEL_ELEVATED = u"#0E1B32"
BG_GROUP_HEADER = u"#11253D"
BG_INPUT = u"#050E18"

# ── Bordes ────────────────────────────────────────────────────────────────────

BORDER = u"#21465C"
BORDER_INPUT = u"#1A3A4D"

# ── Texto ─────────────────────────────────────────────────────────────────────

FG_TITLE = u"#E8F4F8"
FG_BODY = u"#95B8CC"
FG_MUTED = u"#64748b"
FG_ON_ACCENT = u"#0A1A2F"

# ── Acentos ───────────────────────────────────────────────────────────────────

ACCENT_PRIMARY = u"#5BC0DE"
ACCENT_SLIDER = u"#22D3EE"
BTN_MANUAL = u"#2A5C3D"

# ── Espaciado y forma ─────────────────────────────────────────────────────────

PAD_WINDOW = 18
PAD_PANEL = 12
PAD_PANEL_COMPACT = 10
CORNER_PANEL = 4
CORNER_GROUP = 6

# ── Ventana ───────────────────────────────────────────────────────────────────

WINDOW_CHROME_TITLE = u"Arainco"
WINDOW_SHOW_IN_TASKBAR = False
WINDOW_RESIZE_DEFAULT = u"CanResize"
