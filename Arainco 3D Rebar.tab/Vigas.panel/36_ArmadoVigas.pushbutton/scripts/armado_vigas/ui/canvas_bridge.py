# -*- coding: utf-8 -*-
"""Serialización dominio → JSON (referencia mockup HTML; UI producción usa WPF)."""

from armado_vigas.domain.constants import CAPAS_DEFAULT
from armado_vigas.domain.layers import ensure_beam_layers
from armado_vigas.domain.tramos import sort_beams


def _beam_to_canvas_dict(beam, index):
    ensure_beam_layers(beam)
    out = {
        "id": beam.get("id") or u"V-{0}".format(index),
        "type": beam.get("type") or u"30×60",
        "len": float(beam.get("len") or 0.0),
        "u": int(beam.get("u", index)),
        "nCapas": int(beam.get("nCapas") or CAPAS_DEFAULT),
        "nCapasSup": int(beam.get("nCapasSup") or beam.get("nCapas") or CAPAS_DEFAULT),
        "nCapasInf": int(beam.get("nCapasInf") or beam.get("nCapas") or CAPAS_DEFAULT),
        "nSup": int(beam.get("nSup") or 2),
        "nInf": int(beam.get("nInf") or 2),
        "diamSup": int(beam.get("diamSup") or 16),
        "diamInf": int(beam.get("diamInf") or 16),
        "estExtDiam": int(beam.get("estExtDiam") or 10),
        "estExtSpacing": int(beam.get("estExtSpacing") or 125),
        "estCentDiam": int(beam.get("estCentDiam") or 8),
        "estCentSpacing": int(beam.get("estCentSpacing") or 200),
        "estConfin": beam.get("estConfin") or u"Perimetral",
        "supleInfEnabled": bool(beam.get("supleInfEnabled")),
        "diamSupleInf": int(beam.get("diamSupleInf") or 16),
        "nSupleInf": int(beam.get("nSupleInf") or 2),
        "colStart": beam.get("colStart") or u"",
        "colEnd": beam.get("colEnd") or u"",
    }
    for layer_num in (2, 3):
        ns = "nSup{0}".format(layer_num)
        ni = "nInf{0}".format(layer_num)
        ds = "diamSup{0}".format(layer_num)
        di = "diamInf{0}".format(layer_num)
        if beam.get(ns) is not None:
            out[ns] = int(beam[ns])
        if beam.get(ni) is not None:
            out[ni] = int(beam[ni])
        if beam.get(ds) is not None:
            out[ds] = int(beam[ds])
        if beam.get(di) is not None:
            out[di] = int(beam[di])
    return out


def session_to_json_string(session):
    beams = sort_beams(list(session.domain_beams or []))
    payload = {
        "beams": [_beam_to_canvas_dict(b, i) for i, b in enumerate(beams)],
        "apoyosLoaded": bool(getattr(session, "apoyos_loaded", False)),
        "message": session.last_message or u"",
    }
    try:
        clr = __import__("clr")
        clr.AddReference("System.Web.Extensions")
        from System.Web.Script.Serialization import JavaScriptSerializer

        return JavaScriptSerializer().Serialize(payload)
    except Exception:
        return _manual_json(payload)


def _manual_json(payload):
    import json

    try:
        return json.dumps(payload, ensure_ascii=False)
    except Exception:
        parts = ['{"beams":[']
        beam_parts = []
        for b in payload.get("beams") or []:
            fields = []
            for k, v in b.items():
                if isinstance(v, (int, float)):
                    fields.append('"{0}":{1}'.format(k, v))
                else:
                    fields.append(
                        u'"{0}":"{1}"'.format(
                            k, unicode(v).replace('"', '\\"')
                        )
                    )
            beam_parts.append("{" + ",".join(fields) + "}")
        parts.append(",".join(beam_parts))
        parts.append('],"apoyosLoaded":')
        parts.append("true" if payload.get("apoyosLoaded") else "false")
        parts.append(',"message":"')
        parts.append(unicode(payload.get("message") or "").replace('"', '\\"'))
        parts.append('"}')
        return u"".join(parts)
