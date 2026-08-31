# -*- coding: utf-8 -*-
"""
OSNAP de canvas (extremo, medio, perpendicular, proyección, tracking).

Delegado en ``area_rein_losa_sketch`` — misma lógica que 63_AreaReinLosaSketch.
"""

from __future__ import print_function

from area_rein_losa_sketch import (
    _append_ring_snap,
    _build_snap_typed_index,
    _guides_from_tracks_mm,
    _merge_acquired_tracks,
    _merge_ot_points,
    _osnap_status_label,
    _snap_point_mm,
    _tracking_paths_from_ot,
    _SNAP_ACQUIRE_MAX,
)

DRAW_RECT = u"rect"
DRAW_POLY = u"poly"


class CanvasSketchOsnap(object):
    """Índice OSNAP + tracks adquiridos + guías de dibujo."""

    def __init__(self):
        self._dirty = True
        self._ends = []
        self._mids = []
        self._ints = []
        self._centers = []
        self._segs = []
        self._cell_index = None
        self._acquired_tracks = []
        self._ot_points = []
        self.guides = None
        self._ring_count = 0

    def mark_dirty(self):
        self._dirty = True

    def clear_tracks(self):
        self._acquired_tracks = []
        self._ot_points = []
        self.guides = None

    def rebuild(self, ring_polys, draft_pts=None, track_origin=None):
        ends = []
        mids = []
        ints = []
        centers = []
        segs = []
        ring_count = 0

        for poly in ring_polys or []:
            if not poly or len(poly) < 3:
                continue
            ring_count += 1
            try:
                _append_ring_snap(
                    ends,
                    segs,
                    poly,
                    include_midpoints=True,
                    mids=mids,
                    centers=centers,
                )
            except Exception:
                pass

        # Aristas del borrador en curso (misma lógica que Area Rein. Losa Sketch)
        pick = list(draft_pts or [])
        if pick:
            for i, p in enumerate(pick):
                try:
                    ends.append((float(p[0]), float(p[1])))
                except Exception:
                    continue
                if i > 0:
                    try:
                        a = pick[i - 1]
                        b = p
                        segs.append(
                            (
                                (float(a[0]), float(a[1])),
                                (float(b[0]), float(b[1])),
                            )
                        )
                    except Exception:
                        pass

        origin = track_origin
        if origin is not None:
            try:
                ax, ay = float(origin[0]), float(origin[1])
                span = 50000.0
                segs.append(((ax - span, ay), (ax + span, ay)))
                segs.append(((ax, ay - span), (ax, ay + span)))
                ends.append((ax, ay))
            except Exception:
                pass

        self._ends = ends
        self._mids = mids
        self._ints = ints
        self._centers = centers
        self._segs = segs
        self._verts = list(ends) + list(mids) + list(ints) + list(centers)
        self._ring_count = ring_count
        try:
            self._cell_index = _build_snap_typed_index(
                ends, mids, ints, centers, segs
            )
        except Exception:
            self._cell_index = None
        self._dirty = False

    def ensure_ready(self, ring_polys, draft_pts=None, track_origin=None):
        if self._dirty:
            self.rebuild(ring_polys, draft_pts=draft_pts, track_origin=track_origin)

    def resolve(self, pt_mm, track_origin, thresh_mm):
        if pt_mm is None:
            self.guides = None
            return None, None
        try:
            px, py = float(pt_mm[0]), float(pt_mm[1])
        except Exception:
            self.guides = None
            return None, None
        thresh = float(thresh_mm or 0.0)
        if thresh <= 0.0:
            self.guides = None
            return (px, py), None

        acq = list(self._acquired_tracks or [])
        ot = list(self._ot_points or [])
        try:
            result = _snap_point_mm(
                (px, py),
                getattr(self, u"_verts", None) or self._ends,
                self._segs,
                thresh,
                self._cell_index,
                track_origin=track_origin,
                acquired=acq,
                ends=self._ends,
                mids=self._mids,
                ints=self._ints,
                centers=self._centers,
                ot_points=ot,
            )
        except Exception:
            self.guides = None
            return (px, py), None

        if result is None or len(result) < 3:
            self.guides = None
            return (px, py), None

        snapped = result[0]
        kind = result[1]
        guides_hit = result[2]
        for_ln = result[3] if len(result) > 3 else []
        for_pt = result[4] if len(result) > 4 else []

        self._acquired_tracks = _merge_acquired_tracks(
            acq, for_ln, _SNAP_ACQUIRE_MAX
        )
        self._ot_points = _merge_ot_points(ot, for_pt, _SNAP_ACQUIRE_MAX)

        if snapped is None:
            snapped = (px, py)
        center = snapped
        tracks = list(self._acquired_tracks)
        try:
            for path in _tracking_paths_from_ot(
                self._ot_points,
                last_pt=track_origin,
                include_polar=False,
            ):
                tracks.append(path)
        except Exception:
            pass
        try:
            sticky = _guides_from_tracks_mm(tracks, center)
        except Exception:
            sticky = None
        self.guides = sticky if sticky else guides_hit
        return snapped, kind

    def stats(self):
        return {
            u"rings": int(getattr(self, u"_ring_count", 0) or 0),
            u"ends": len(self._ends or []),
            u"mids": len(self._mids or []),
            u"segs": len(self._segs or []),
            u"tracks": len(self._acquired_tracks or []),
            u"ot_pts": len(self._ot_points or []),
            u"dirty": bool(getattr(self, u"_dirty", False)),
            u"has_index": self._cell_index is not None,
        }


def osnap_status_label(kind):
    return _osnap_status_label(kind)
