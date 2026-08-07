# -*- coding: utf-8 -*-
"""
Shell WPF reutilizable — tema oscuro BIMTools con cinta blanca nativa.

**Solo para herramientas nuevas.** Las UIs existentes no se migran con este módulo;
su armonización se hará caso a caso.

Ventana estándar de Windows (sin ``WindowStyle="None"``): la barra de título
del SO queda visible; el contenido usa la paleta de ``bimtools_ui_tokens``.

Regla: ``.cursor/rules/bimtools-ui.mdc``
Mockup: ``canvases/bimtools-ui-default-style-mockup.canvas.tsx``
"""

from __future__ import print_function

from bimtools_ui_tokens import (
    BG_APP,
    BG_PANEL,
    BORDER,
    CORNER_PANEL,
    FG_BODY,
    FG_MUTED,
    FG_TITLE,
    FONT_FAMILY,
    FONT_SIZE_BASE,
    FONT_SIZE_HINT,
    FONT_SIZE_STATUS,
    FONT_SIZE_SUBTITLE,
    FONT_SIZE_TITLE,
    FONT_WEIGHT_TITLE,
    PAD_PANEL,
    PAD_WINDOW,
    WINDOW_SHOW_IN_TASKBAR,
    WINDOW_CHROME_TITLE,
)

_PLACEHOLDER_STYLES = u"__BIMTOOLS_DARK_STYLES__"
_PLACEHOLDER_WINDOW_TITLE = u"__SHELL_WINDOW_TITLE__"
_PLACEHOLDER_TOOL_TITLE = u"__SHELL_TOOL_TITLE__"
_PLACEHOLDER_WIDTH = u"__SHELL_WIDTH__"
_PLACEHOLDER_MIN_WIDTH = u"__SHELL_MIN_WIDTH__"
_PLACEHOLDER_HEIGHT = u"__SHELL_HEIGHT__"
_PLACEHOLDER_MIN_HEIGHT = u"__SHELL_MIN_HEIGHT__"
_PLACEHOLDER_RESIZE = u"__SHELL_RESIZE__"
_PLACEHOLDER_SIZE_TO_CONTENT = u"__SHELL_SIZE_TO_CONTENT__"
_PLACEHOLDER_BODY = u"__SHELL_BODY__"
_PLACEHOLDER_FOOTER_HINT = u"__SHELL_FOOTER_HINT__"
_PLACEHOLDER_FOOTER_ACTIONS = u"__SHELL_FOOTER_ACTIONS__"

_SIMPLE_TOOL_XAML = u"""
<Window
  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  Title="__SHELL_WINDOW_TITLE__"
  Width="__SHELL_WIDTH__"
  MinWidth="__SHELL_MIN_WIDTH__"
  Height="__SHELL_HEIGHT__"
  MinHeight="__SHELL_MIN_HEIGHT__"
  SizeToContent="__SHELL_SIZE_TO_CONTENT__"
  ResizeMode="__SHELL_RESIZE__"
  WindowStartupLocation="Manual"
  Background="{bg_app}"
  FontFamily="{font_family}"
  FontSize="{font_base}"
  ShowInTaskbar="{show_taskbar}">
  <Window.Resources>
__BIMTOOLS_DARK_STYLES__
  </Window.Resources>
  <Border Background="{bg_app}" BorderBrush="{border}" BorderThickness="1"
          Padding="{pad_window}">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <StackPanel Grid.Row="0" Margin="0,0,0,10">
        <TextBlock x:Name="TxtTitle" Text="__SHELL_TOOL_TITLE__"
                   Foreground="{fg_title}" FontSize="{font_title}"
                   FontWeight="{font_weight_title}"/>
        <TextBlock x:Name="TxtSubtitle" Margin="0,6,0,0"
                   Foreground="{fg_body}" FontSize="{font_subtitle}"
                   TextWrapping="Wrap"/>
      </StackPanel>

      <Border Grid.Row="1" Margin="0,0,0,0" Background="{bg_panel}"
              BorderBrush="{border}" BorderThickness="1"
              CornerRadius="{corner_panel}" Padding="{pad_panel}">
        __SHELL_BODY__
      </Border>

      __SHELL_FOOTER_HINT__

      <Grid Grid.Row="3" Margin="0,14,0,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="TxtStatus" Grid.Column="0" VerticalAlignment="Center"
                   Foreground="{fg_muted}" FontSize="{font_status}"
                   TextWrapping="Wrap" Margin="0,0,12,0"/>
        <StackPanel Grid.Column="1" Orientation="Horizontal"
                    HorizontalAlignment="Right">
          __SHELL_FOOTER_ACTIONS__
        </StackPanel>
      </Grid>
    </Grid>
  </Border>
</Window>
""".format(
    bg_app=BG_APP,
    border=BORDER,
    pad_window=PAD_WINDOW,
    font_family=FONT_FAMILY,
    font_base=FONT_SIZE_BASE,
    show_taskbar=u"False" if not WINDOW_SHOW_IN_TASKBAR else u"True",
    fg_title=FG_TITLE,
    font_title=FONT_SIZE_TITLE,
    font_weight_title=FONT_WEIGHT_TITLE,
    fg_body=FG_BODY,
    font_subtitle=FONT_SIZE_SUBTITLE,
    bg_panel=BG_PANEL,
    corner_panel=CORNER_PANEL,
    pad_panel=PAD_PANEL,
    fg_muted=FG_MUTED,
    font_status=FONT_SIZE_STATUS,
)


def _escape_xaml(text):
    s = text if text is not None else u""
    try:
        s = unicode(s)
    except NameError:
        s = str(s)
    return (
        s.replace(u"&", u"&amp;")
        .replace(u"<", u"&lt;")
        .replace(u">", u"&gt;")
        .replace(u'"', u"&quot;")
    )


def build_simple_tool_xaml(
    title,
    styles_xml,
    body_xaml,
    footer_actions_xaml=u"",
    footer_hint_xaml=u"",
    width=520,
    min_width=0,
    height=0,
    min_height=0,
    resize_mode=u"CanResize",
    size_to_content_height=False,
):
    """
    Genera XAML de ventana simple con shell estándar.

    ``title``: título de la herramienta en el cuerpo (TxtTitle), p. ej.
    ``Arainco: Mi herramienta``. La cinta blanca del SO usa siempre
    ``WINDOW_CHROME_TITLE`` (``Arainco``).

    ``body_xaml``: contenido del panel central (TextBlock, StackPanel, etc.).
    ``footer_actions_xaml``: botones alineados a la derecha del footer.
    ``footer_hint_xaml``: bloque opcional entre el panel y el footer (fila Grid).
    """
    xaml = _SIMPLE_TOOL_XAML
    xaml = xaml.replace(_PLACEHOLDER_STYLES, styles_xml or u"")
    xaml = xaml.replace(
        _PLACEHOLDER_WINDOW_TITLE,
        _escape_xaml(WINDOW_CHROME_TITLE),
    )
    xaml = xaml.replace(_PLACEHOLDER_TOOL_TITLE, _escape_xaml(title))

    if width and int(width) > 0:
        xaml = xaml.replace(_PLACEHOLDER_WIDTH, unicode(int(width)))
    else:
        xaml = xaml.replace(u' Width="__SHELL_WIDTH__"', u"")

    if min_width and int(min_width) > 0:
        xaml = xaml.replace(_PLACEHOLDER_MIN_WIDTH, unicode(int(min_width)))
    else:
        xaml = xaml.replace(u' MinWidth="__SHELL_MIN_WIDTH__"', u"")

    if height and int(height) > 0:
        xaml = xaml.replace(_PLACEHOLDER_HEIGHT, unicode(int(height)))
    else:
        xaml = xaml.replace(u' Height="__SHELL_HEIGHT__"', u"")

    if min_height and int(min_height) > 0:
        xaml = xaml.replace(_PLACEHOLDER_MIN_HEIGHT, unicode(int(min_height)))
    else:
        xaml = xaml.replace(u' MinHeight="__SHELL_MIN_HEIGHT__"', u"")

    xaml = xaml.replace(_PLACEHOLDER_RESIZE, resize_mode or u"CanResize")

    if size_to_content_height:
        xaml = xaml.replace(_PLACEHOLDER_SIZE_TO_CONTENT, u"Height")
    else:
        xaml = xaml.replace(u' SizeToContent="__SHELL_SIZE_TO_CONTENT__"', u"")

    xaml = xaml.replace(_PLACEHOLDER_BODY, body_xaml or u"")
    xaml = xaml.replace(_PLACEHOLDER_FOOTER_ACTIONS, footer_actions_xaml or u"")

    if footer_hint_xaml and footer_hint_xaml.strip():
        hint = (
            u'<TextBlock x:Name="TxtFooterHint" Grid.Row="2" '
            u'Foreground="{fg}" FontSize="{fs}" TextWrapping="Wrap" '
            u'Margin="0,8,0,0">{content}</TextBlock>'
        ).format(fg=FG_MUTED, fs=FONT_SIZE_HINT, content=footer_hint_xaml)
        xaml = xaml.replace(_PLACEHOLDER_FOOTER_HINT, hint)
    else:
        xaml = xaml.replace(_PLACEHOLDER_FOOTER_HINT, u"")

    return xaml
