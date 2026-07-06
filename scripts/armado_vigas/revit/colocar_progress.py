# -*- coding: utf-8 -*-
"""Barra de progreso pyRevit durante colocación de armadura (Armado vigas)."""

from __future__ import print_function


def _pbar_enabled():
    try:
        from pyrevit import forms as _forms  # noqa: F401
        return True
    except Exception:
        return False


def _phase_count(session):
    """Fases visibles en la barra (laterales opcional)."""
    n = 7
    if getattr(session, u"lateralesEnabled", False):
        n += 1
    return n


class ColocarArmaduraProgress(object):
    """Context manager no-op si pyRevit ProgressBar no está disponible."""

    def __init__(self, session):
        self._session = session
        self._total = _phase_count(session)
        self._index = 0
        self._pb = None
        self._open = False

    def __enter__(self):
        if not _pbar_enabled() or self._total < 1:
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
        return u"Arainco: Armado vigas {0}/{1}".format(
            int(index) + 1,
            int(self._total),
        )

    def step(self, phase_label):
        """Avanza una fase y actualiza título de la barra."""
        if self._pb is None:
            return
        i = int(self._index)
        self._index += 1
        base = u"{0} — {1}".format(self._title(i), phase_label)
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
