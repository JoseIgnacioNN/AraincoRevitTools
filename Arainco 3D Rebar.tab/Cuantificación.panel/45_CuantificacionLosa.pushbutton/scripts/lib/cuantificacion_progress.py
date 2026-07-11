# -*- coding: utf-8 -*-
"""Barra de progreso pyRevit — Cuantificación losa por nivel."""

from __future__ import print_function


def _pbar_enabled():
    try:
        from pyrevit import forms as _forms  # noqa: F401

        return True
    except Exception:
        return False


class CuantificacionLosaProgress(object):
    """Context manager no-op si pyRevit ProgressBar no está disponible."""

    def __init__(self, total, title_prefix=None):
        self._total = max(1, int(total or 1))
        self._index = 0
        self._pb = None
        self._open = False
        self._title_prefix = title_prefix or u"Arainco: Cuantificación losa"

    def __enter__(self):
        if not _pbar_enabled():
            return self
        try:
            from pyrevit import forms as _pyrevit_forms

            self._pb = _pyrevit_forms.ProgressBar(
                title=self._title(0),
                cancellable=False,
            )
            try:
                from System.Windows.Media import Color, SolidColorBrush

                self._pb.Resources[u"pyRevitAccentBrush"] = SolidColorBrush(
                    Color.FromRgb(91, 192, 222),
                )
            except Exception:
                pass
            self._pb.__enter__()
            self._open = True
        except Exception:
            self._pb = None
            self._open = False
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._open and self._pb is not None:
            try:
                self._pb.__exit__(exc_type, exc_val, exc_tb)
            except Exception:
                pass
        self._open = False
        self._pb = None
        return False

    def _title(self, index):
        return u"{0} {1}/{2}".format(
            self._title_prefix,
            int(index) + 1,
            int(self._total),
        )

    def step(self, phase_label):
        """Avanza un paso y actualiza título de la barra."""
        if self._pb is None:
            return
        i = int(self._index)
        if i >= self._total:
            i = self._total - 1
        self._index = i + 1
        label = phase_label or u""
        base = u"{0} — {1}".format(self._title(i), label) if label else self._title(i)
        try:
            if hasattr(self._pb, u"update_progress"):
                try:
                    self._pb.update_progress(i + 1, max_value=self._total)
                except TypeError:
                    try:
                        self._pb.update_progress(i + 1, max=self._total)
                    except Exception:
                        pass
        except Exception:
            pass
        try:
            self._pb.title = base
        except Exception:
            pass
