# -*- coding: utf-8 -*-
"""
Post-proceso para **Armado Muros Nodo** (un solo ``RemoveAreaReinforcementSystem``):

* Tras L+135 / pata: en cada ``Rebar`` con reparto (≥3 posiciones), se **desactivan la primera
  y la última barra** del set (``includeFirstBar`` / ``includeLastBar`` = False en la API de layout).
* Barras **horizontales en planta** (criterio arearein_exterior / interior) →
  L + 135° y ganchos, como ``arearein_ambas_caras_h_l135_rps``.
* **Verticales** L+135: pata L en **un** extremo y gancho 135 en el trazado; el otro extremo
  del tramo largo: gancho 135 (sin segunda pata L). La pata L sigue el mismo criterio que
  en **horizontales** respecto al boceto: **cara exterior** → pata al **inicio** del boceto
  (``pata_en_extremo_final=False``); **cara interior** → pata al **final** (``True``), coherente
  con el sentido de malla en cada cara. Si hace falta anclaje en cabeza **y** pie
  (sin forjados), **una** pata L en el **pie** y en cabeza solo 135 (no
  ``extender_doble_pata_135_y_reemplazar``).
* Si el boceto llevó fundación (``h_fund_internal`` > 0), las barras **verticales** pueden
  recibir pata 90° al final del boceto (``arearein_verticales_pata_90_fin_boceto_rps``) cuando
  hay forjado unido a la **cara inferior** (y no aplica el L+135 en el pie por falta de unión).

* **Verticales, caras sup. e inf.:** se evalúan **por separado** (ninguna asume el resultado
  de la otra). **Cabeza:** ``hay_suelo_unido_cara_superior``. **Pie (L+135):**
  ``hay_suelo_unido_cara_inferior_excluyendo_fundacion`` (forjado estructural) **o**
  ``hay_muro_unido_cara_inferior`` (muro apilado/unido bajo el host). Si hay cualquiera de
  esos empalmes en la cara inferior, **no** se agrega pata L ni gancho 135 en el pie.
  Con ``hay_muro_unido_cara_superior`` (muro apilado en la **cara superior**), en verticales
  de **cara interior** o **cara exterior** no se aplica pata L ni gancho 135 en el **extremo
  inicial** hacia el encuentro con el muro superior (criterio análogo al muro bajo y el pie);
  en los 135° que sí se generan se parte de ``GANCHOS_135_MURO_CARA_SUP_*`` y se ajusta
  el par **Left/Right** y **180°** en el extremo de la pata (``hook_rot_end0`` ext.,
  ``hook_rot_end1`` int.);
  en ese extremo, el boceto se **estira** hacia arriba con *L* = tabla
  (``traslape_mm_from_nominal_diameter_mm`` / grado en
  ``arearein_verticales_empotramiento_rps``) **+ 25 mm fijos** (``EMBED_EXTRA_TABLE_MM``).
  Con muro en la **cara inferior** (``hay_muro_unido_cara_inferior``) el extremo de **menor Z**
  sin pata L ni gancho 135 (empalme hacia abajo) se alarga **25 mm** fijos
  (``extender_vertical_pie_emp_muro_inf_mm`` / ``EMBED_MURO_INFERIOR_MM``), en el mismo
  orden «malla lisa → L+135».
  La cimentación unida no anula L+135 en el pie (alineado a excl. de fund. en la cara superior).
  En cada extremo, si *no* hay ese empalme, **pata L + gancho 135°** (sentido: ``rebar_extender_…``).
  Si aplica en **cabeza y pie**: un solo ``extender_l_asignar_ganchos_135_y_reemplazar`` con
  pata hacia el **pie** según cara (ext. inicio / int. final); cabeza con gancho 135 sin
  segunda pata L. (El doble ``extender_doble_pata_135_y_reemplazar`` añadía L en ambos
  lados y se retiró.) Solo **cabeza** (sin requisito de pie) usa
  ``pata_en_extremo_final_para_cabeza_por_elevación`` para anclar arriba.
* Pata 90° a fundación: boceto con fundación, módulo pata 90, y ``hay_suelo_unido_cara_inferior``
  **completo** (incl. fundación) — criterio distinto del L+135 en el pie; si aplica, ese bloque
  trata de la pata 90 y no reemplaza el criterio anterior salvo en el flujo concreto.

Revit 2024+, pyRevit. Los módulos referenciados deben estar en el mismo path ``scripts/``.
"""
from __future__ import print_function

import os
import sys

_L135_ERR = u""
_AREX_ERR = u""
_ARIN_ERR = u""
_P90_ERR = u""
l135 = arex = arin = pata90 = vert_emp = None
try:
    import rebar_extender_l_ganchos_135_rps as l135
except Exception as ex:  # noqa: BLE001
    _L135_ERR = u"{0!s}".format(ex)
try:
    import arearein_exterior_h_l135_rps as arex
except Exception as ex:  # noqa: BLE001
    _AREX_ERR = u"{0!s}".format(ex)
try:
    import arearein_interior_h_l135_rps as arin
except Exception as ex:  # noqa: BLE001
    _ARIN_ERR = u"{0!s}".format(ex)
try:
    import arearein_verticales_pata_90_fin_boceto_rps as pata90
except Exception as ex:  # noqa: BLE001
    _P90_ERR = u"{0!s}".format(ex)
try:
    import arearein_verticales_empotramiento_rps as vert_emp
except Exception:  # noqa: BLE001
    vert_emp = None
try:
    from wall_node_boolean_section_rps import (
        hay_suelo_unido_cara_inferior,
        hay_suelo_unido_cara_inferior_excluyendo_fundacion,
        hay_muro_unido_cara_inferior,
        hay_muro_unido_cara_superior,
        hay_suelo_unido_cara_superior,
    )
except Exception:  # noqa: BLE001
    # Degradado: no se puede comprobar unión; se asume *sin* forjado en cada cara (como
    # ``False`` de las funciones reales) para no bloquear L+135; conviene reparar el import.
    def hay_suelo_unido_cara_superior(_doc, _host):
        return False

    def hay_suelo_unido_cara_inferior(_doc, _host):
        return False

    def hay_suelo_unido_cara_inferior_excluyendo_fundacion(_doc, _host):
        return False

    def hay_muro_unido_cara_inferior(_doc, _host):
        return False

    def hay_muro_unido_cara_superior(_doc, _host):
        return False

from Autodesk.Revit.DB import Transaction, Wall
from Autodesk.Revit.DB.Structure import (
    AreaReinforcement,
    MultiplanarOption,
    Rebar,
    RebarHookOrientation,
)

from armado_muros_nodo_shared import desactivar_extremos_rebar_set

# Criterio alineado con arearein_verticales_pata_90 (tangente 1.º tramo, muro).
_MURO_VERT_MIN_ABS_TZ = 0.45


def _import_error_mensaje():
    partes = []
    if l135 is None:
        partes.append(u"rebar_extender: {0}".format(_L135_ERR or u"—"))
    if arex is None:
        partes.append(u"arearein_exterior_h_l135: {0}".format(_AREX_ERR or u"—"))
    if arin is None:
        partes.append(u"arearein_interior_h_l135: {0}".format(_ARIN_ERR or u"—"))
    if pata90 is None and _P90_ERR:
        partes.append(u"pata90: {0}".format(_P90_ERR or u"—"))
    if not partes:
        return u""
    return u"\n".join(partes)


def _eid(ei):
    if ei is None:
        return 0
    try:
        return int(ei.Value)
    except Exception:
        try:
            return int(ei.IntegerValue)
        except Exception:
            return 0


def _rebar_es_vertical_muro_criterio(r, host, pos_idx):
    u"""
    True si el refuerzo es vertical de muro (L+135 vert. sin depender del import pata 90).
    """
    if pata90 is not None:
        try:
            return bool(
                pata90._rebar_es_vertical_por_criterio(r, host, int(pos_idx))
            )
        except Exception:  # noqa: BLE001
            pass
    if r is None or not isinstance(host, Wall):
        return False
    try:
        crvs = r.GetCenterlineCurves(
            False,
            False,
            False,
            MultiplanarOption.IncludeAllMultiplanarCurves,
            int(pos_idx),
        )
        if crvs is None or crvs.Count < 1:
            return False
        c0 = crvs[0]
        p0 = c0.GetEndPoint(0)
        p1 = c0.GetEndPoint(1)
        v = p1 - p0
        if v.GetLength() < 1e-12:
            return False
        return abs(float(v.Z)) >= float(_MURO_VERT_MIN_ABS_TZ)
    except Exception:  # noqa: BLE001
        return False


def _gancho_135_kwargs_muro_unido_cara_sup(activo, cara=None):
    u"""
    Con muro unido a la **cara superior** del host, pasa a ``extender_l`` el perfil
    de ganchos. **180°** en el extremo de la pata: ``_assign_135_hook_solo_pata`` usa
    ``rot0`` si la pata es ext.0 (cara ext.) o ``rot1`` si la pata es ext.1 (cara int.),
    por eso **exterior** pone 180 en ``hook_rot_end0`` (*Hook Rotation at Start* en
    muchos bocetos) e **interior** 180 en ``hook_rot_end1`` (gancho en el otro extremo).
    In-plane Left/Right distinto ext. vs int.
    """
    if not activo or l135 is None:
        return {}
    try:
        o0 = l135.GANCHOS_135_MURO_CARA_SUP_ORIENT0
        o1 = l135.GANCHOS_135_MURO_CARA_SUP_ORIENT1
        r0, r1 = 0.0, 0.0
        if cara == u"ext":
            o0, o1 = RebarHookOrientation.Left, RebarHookOrientation.Right
            r0, r1 = 180.0, 0.0
        elif cara == u"int":
            o0, o1 = RebarHookOrientation.Right, RebarHookOrientation.Left
            r0, r1 = 0.0, 180.0
        return {
            u"hook_orient_end0": o0,
            u"hook_orient_end1": o1,
            u"hook_rot_end0_deg": r0,
            u"hook_rot_end1_deg": r1,
        }
    except Exception:  # noqa: BLE001
        return {}


def _face_cara_ext_o_int_desde_pata_o_arex(r, host):
    u"""ext / int / None; con o sin módulo pata 90."""
    if pata90 is not None:
        try:
            return pata90._face_cara_ext_o_int(r, host)
        except Exception:  # noqa: BLE001
            pass
    ex = arex._rebar_solo_cara_exterior(r, host)
    inn = arin._rebar_solo_cara_interior(r, host)
    if ex and not inn:
        return u"ext"
    if inn and not ex:
        return u"int"
    return None


def post_malla_nudo_tras_crear(
    doc, area_rein, h_fund_internal
):
    """
    :returns: ``(rebar_finales_ids, mensajes)`` con ``rebar_finales_ids`` = lista
    de ``ElementId`` de barra a marcar visibles. Si no se pudo importar lógica
    horizontal, devuelve ``(None, [msg])`` y el caller debe conservar el area reinf.
    """
    if doc is None or not isinstance(area_rein, AreaReinforcement):
        return (
            [],
            [u"post_malla: documento o AreaReinforcement inválido."],
        )
    if l135 is None or arex is None or arin is None:
        return (None, [u"Faltan módulos L+135 (horizontal):\n{0}".format(_import_error_mensaje())])
    # `vert_emp` se importó con el 1.º carga de este módulo: pyRevit mantiene
    # ``sys.modules``; sin reload, sigue en uso ``extender_vertical_…`` antiguo
    # (p. ej. sin +25 mm) tras editar ``arearein_verticales_empotramiento_rps``.
    global vert_emp
    if vert_emp is not None:
        try:
            _vn = getattr(vert_emp, u"__name__", u"arearein_verticales_empotramiento_rps")
            if _vn in sys.modules:
                try:
                    import importlib
                    vert_emp = importlib.reload(sys.modules[_vn])
                except Exception:  # noqa: BLE001
                    try:
                        import imp
                        vert_emp = imp.reload(sys.modules[_vn])
                    except Exception:
                        pass
        except Exception:  # noqa: BLE001
            pass
    t = Transaction(doc, u"BIMTools: quitar area (nudo) → L+135 / pata 90 o vert. L+135")
    t.Start()
    try:
        new_ids = AreaReinforcement.RemoveAreaReinforcementSystem(doc, area_rein)
    except Exception as ex:  # noqa: BLE001
        t.RollBack()
        return [], [u"RemoveAreaReinforcementSystem: {0!s}".format(ex)]
    t.Commit()
    # Sin Regenerate aquí: tras Remove, varias rebars quedan en estado intermedio
    # y un Regenerate() puede disparar errores de forma/ganchos antes de L+135 / pata 90.

    created = arex._iter_ids(new_ids)
    ms = []
    rebar_finales = []
    inv_ext = bool(l135.INVERTIR_DIRECCION_PATA)
    inv_int = not inv_ext
    px, tz = arex.PLAN_PREFER_X_FOR_HORIZONTAL, arex.HORIZONTAL_MAX_ABS_TZ
    hay_fund = h_fund_internal is not None and float(h_fund_internal) > 1e-6

    w_host = None
    try:
        w_host = doc.GetElement(area_rein.GetHostId())
    except Exception:  # noqa: BLE001
        w_host = None
    if w_host is None or not isinstance(w_host, Wall):
        for _eid0 in created:
            _r0 = doc.GetElement(_eid0)
            if _r0 is None or not isinstance(_r0, Rebar):
                continue
            try:
                _h0 = doc.GetElement(_r0.GetHostId())
            except Exception:  # noqa: BLE001
                _h0 = None
            if _h0 is not None and isinstance(_h0, Wall):
                w_host = _h0
                break
    # Forjado sup / inf.: consultas **independientes** (mismo criterio de error: fallo = no
    # asumir unión en esa cara → permitir L+135 en el extremo correspondiente).
    # *Inferior, dos lecturas:* ``hay_suelo_unido_cara_inferior`` incl. fund. (pata 90);
    # ``_excluyendo_fundacion`` solo forjado estructural, alineado a la cara superior, para
    # ``need_l135_pie`` (ciment. unida no debe suprimir L+135 en el pie).
    hay_forjado_sup = False
    hay_forjado_inf = False
    hay_forjado_inf_sin_fund = False
    hay_muro_inf = False
    hay_muro_sup = False
    if w_host is not None and isinstance(w_host, Wall):
        try:
            hay_forjado_sup = bool(hay_suelo_unido_cara_superior(doc, w_host))
        except Exception:  # noqa: BLE001
            hay_forjado_sup = False
        try:
            hay_forjado_inf = bool(hay_suelo_unido_cara_inferior(doc, w_host))
        except Exception:  # noqa: BLE001
            hay_forjado_inf = False
        try:
            hay_forjado_inf_sin_fund = bool(
                hay_suelo_unido_cara_inferior_excluyendo_fundacion(doc, w_host)
            )
        except Exception:  # noqa: BLE001
            hay_forjado_inf_sin_fund = False
        try:
            hay_muro_inf = bool(hay_muro_unido_cara_inferior(doc, w_host))
        except Exception:  # noqa: BLE001
            hay_muro_inf = False
        try:
            hay_muro_sup = bool(hay_muro_unido_cara_superior(doc, w_host))
        except Exception:  # noqa: BLE001
            hay_muro_sup = False
    need_l135_cabeza = not hay_forjado_sup
    # Pie: si hay forjado estructural inferior (sin fund.) o hay muro unido bajo el host,
    # NO aplicar L+135 (pata+gancho) en el extremo inferior.
    need_l135_pie = (not hay_forjado_inf_sin_fund) and (not hay_muro_inf)
    # Si hay muro unido en la cara inferior, al aplicar L+135 en **cabeza** no debe quedar
    # gancho 135° en el pie (extremo opuesto al de la pata L).
    gancho_solo_extremo_pata_por_muro_inf = bool(hay_muro_inf)

    c_not_rebar = c_horiz = c_horiz_ext = c_horiz_int = 0
    c_horiz_ambig = c_l135_ok = c_l135_fail = 0
    c_vert_l135 = c_vert_l135_ok = c_vert_l135_fail = 0
    c_vert_pata = c_vert_pata_ok = c_vert_pata_fail = 0
    c_unchanged = 0

    for eid in created:
        r = doc.GetElement(eid)
        if r is None:
            continue
        _embed_muro_sup_ok_en_path_l135 = False
        if not isinstance(r, Rebar):
            c_not_rebar += 1
            rebar_finales.append(eid)
            ms.append(u"— id {0}: no Rebar, sin cambio.".format(_eid(eid)))
            continue
        host = doc.GetElement(r.GetHostId())
        if host is None:
            rebar_finales.append(eid)
            continue

        is_hor = arex._rebar_horizontal_en_plano(r, px, tz, host)
        if is_hor:
            c_horiz += 1
        if is_hor:
            if arex._rebar_solo_cara_exterior(r, host):
                invert, pata_en_final = inv_ext, False
                c_horiz_ext += 1
            elif arin._rebar_solo_cara_interior(r, host):
                invert, pata_en_final = inv_int, True
                c_horiz_int += 1
            else:
                invert, pata_en_final = None, None
                c_horiz_ambig += 1
            if invert is not None:
                try:
                    largo_p = l135.largo_pata_mm_desde_espesor_host(doc, host)
                except Exception:  # noqa: BLE001
                    largo_p = float(getattr(l135, "LARGO_PATA_MM", 200.0))
                ok, msg, new_rb = l135.extender_l_asignar_ganchos_135_y_reemplazar(
                    doc,
                    r,
                    largo_p,
                    invert,
                    l135.INDICE_POSICION,
                    pata_en_final,
                )
                if ok and new_rb is not None:
                    c_l135_ok += 1
                    rebar_finales.append(new_rb.Id)
                    # No aplicar .format a msg: puede contener "{" (Revit) y colisionar;
                    # el id se concatena.
                    mid = _eid(new_rb.Id)
                    ms.append(
                        (msg or u"OK horiz. L+135") + (u"  [nuevo: {0}]".format(mid))
                    )
                else:
                    c_l135_fail += 1
                    rebar_finales.append(eid)
                    ms.append(
                        u"FALLO horiz. L+135 id {0}: {1}".format(
                            _eid(eid), (msg or u"")
                        )
                    )
                continue
        if need_l135_cabeza or need_l135_pie:
            if _rebar_es_vertical_muro_criterio(r, host, 0):
                cara = _face_cara_ext_o_int_desde_pata_o_arex(r, host)
                if cara is not None:
                    pata_en_malla = None
                    if arex._rebar_solo_cara_exterior(r, host):
                        invert = inv_ext
                        # Mismo criterio que horizontales: ext → pata al inicio del boceto.
                        pata_en_malla = False
                    elif arin._rebar_solo_cara_interior(r, host):
                        invert = inv_int
                        pata_en_malla = True
                    else:
                        invert = None
                    if invert is not None and pata_en_malla is not None:
                        c_vert_l135 += 1
                        try:
                            largo_p = None
                            fn = getattr(
                                l135, "largo_pata_mm_muro_vertical_entre_caras", None
                            )
                            if callable(fn):
                                largo_p = fn(
                                    doc,
                                    host,
                                    r,
                                    l135.INDICE_POSICION,
                                )
                            if largo_p is None:
                                largo_p = l135.largo_pata_mm_desde_espesor_host(
                                    doc, host
                                )
                        except Exception:  # noqa: BLE001
                            try:
                                largo_p = l135.largo_pata_mm_desde_espesor_host(
                                    doc, host
                                )
                            except Exception:  # noqa: BLE001
                                largo_p = float(
                                    getattr(l135, "LARGO_PATA_MM", 200.0)
                                )
                        # Cabeza + pie: **una** pata L (cara ext. inicio / int. final, como
                        # horizontales) + tramo + gancho 135 en el otro extremo.
                        # ``extender_doble_pata_135`` generaba pata L en ambos extremos.
                        r_work = r
                        muro_emp_preaplicado = False
                        _preembed_msg = u""
                        if hay_muro_sup and vert_emp is not None:
                            # Antes de L+135: boceto liso. Después, CreateFromRebar del
                            # extensor sobre final_rb dispara suelo InternalException (log H5).
                            _fn_pre = getattr(
                                vert_emp,
                                u"extender_vertical_cabeza_empotramiento_por_diam",
                                None,
                            )
                            if callable(_fn_pre):
                                _okp, _msgp, _rbp = _fn_pre(
                                    doc,
                                    r_work,
                                    int(
                                        getattr(
                                            l135,
                                            u"INDICE_POSICION",
                                            0,
                                        )
                                    ),
                                )
                                if _okp and _rbp is not None:
                                    r_work = _rbp
                                    muro_emp_preaplicado = True
                                    if _msgp:
                                        _preembed_msg = (u"" + _msgp)[:500]
                        if hay_muro_inf and vert_emp is not None:
                            _fn_pie0 = getattr(
                                vert_emp,
                                u"extender_vertical_pie_emp_muro_inf_mm",
                                None,
                            )
                            if callable(_fn_pie0):
                                _okpi, _msgpi, _rbpi = _fn_pie0(
                                    doc,
                                    r_work,
                                    int(
                                        getattr(
                                            l135,
                                            u"INDICE_POSICION",
                                            0,
                                        )
                                    ),
                                )
                                if _okpi and _rbpi is not None:
                                    r_work = _rbpi
                                    if _msgpi:
                                        _m_pi = (u"" + _msgpi)[:500]
                                        _preembed_msg = (
                                            (u"  " + _m_pi)
                                            if _preembed_msg
                                            else _m_pi
                                        )
                        ok_all = True
                        msg_last = _preembed_msg
                        final_rb = None
                        if need_l135_cabeza and need_l135_pie:
                            # Muro en cabeza: sin 135 hacia el empalme superior. Interior: pata
                            # según malla (final); exterior: pata forzada al pie; en ambos
                            # ``gancho_135_solo`` en el extremo de la pata.
                            if hay_muro_sup:
                                if pata_en_malla:
                                    _p6, _p7 = pata_en_malla, True
                                else:
                                    _p6 = (
                                        l135.pata_en_extremo_final_para_pie_por_elevacion(
                                            r_work, l135.INDICE_POSICION
                                        )
                                    )
                                    _p7 = True
                            else:
                                _p6, _p7 = pata_en_malla, False
                            _gkw = _gancho_135_kwargs_muro_unido_cara_sup(
                                bool(hay_muro_sup),
                                cara,
                            )
                            _ok1, _msg1, _nb1 = (
                                l135.extender_l_asignar_ganchos_135_y_reemplazar(
                                    doc,
                                    r_work,
                                    largo_p,
                                    invert,
                                    l135.INDICE_POSICION,
                                    _p6,
                                    _p7,
                                    **_gkw
                                )
                            )
                            msg_last = _msg1 or msg_last
                            if not _ok1 or _nb1 is None:
                                ok_all = False
                            else:
                                r_work = _nb1
                                final_rb = _nb1
                        elif need_l135_cabeza:
                            # Muro apilado arriba: no L+135 de anclaje análogo a cabeza; la malla
                            # continúa hacia el muro (ext. o int., extremo hacia el empalme).
                            if hay_muro_sup:
                                ok_cab, msg_cab, nb_c = (True, u"", r_work)
                            else:
                                _pe_cab = (
                                    l135.pata_en_extremo_final_para_cabeza_por_elevacion(
                                        r_work, l135.INDICE_POSICION
                                    )
                                )
                                # Misma lógica que horizontales: pata L en un extremo y ganchos
                                # 135° en **ambos** (no ``gancho_135_solo``), salvo muro int. inf.
                                _gkw2 = _gancho_135_kwargs_muro_unido_cara_sup(
                                    bool(hay_muro_sup),
                                    cara,
                                )
                                ok_cab, msg_cab, nb_c = (
                                    l135.extender_l_asignar_ganchos_135_y_reemplazar(
                                        doc,
                                        r_work,
                                        largo_p,
                                        invert,
                                        l135.INDICE_POSICION,
                                        _pe_cab,
                                        gancho_solo_extremo_pata_por_muro_inf,
                                        **_gkw2
                                    )
                                )
                            msg_last = msg_cab or msg_last
                            if not ok_cab or nb_c is None:
                                ok_all = False
                            else:
                                r_work = nb_c
                                final_rb = nb_c
                        elif need_l135_pie:
                            # Con muro en cabeza, sin 135 en el extremo hacia el muro; ambas caras.
                            _gancho_solo_por_muro_sup = bool(hay_muro_sup)
                            _gkw3 = _gancho_135_kwargs_muro_unido_cara_sup(
                                _gancho_solo_por_muro_sup,
                                cara,
                            )
                            ok_pie, msg_pie, nb_p = (
                                l135.extender_l_asignar_ganchos_135_y_reemplazar(
                                    doc,
                                    r_work,
                                    largo_p,
                                    invert,
                                    l135.INDICE_POSICION,
                                    pata_en_malla,
                                    _gancho_solo_por_muro_sup,
                                    **_gkw3
                                )
                            )
                            msg_last = msg_pie or msg_last
                            if not ok_pie or nb_p is None:
                                ok_all = False
                            else:
                                r_work = nb_p
                                final_rb = nb_p
                        if ok_all and final_rb is not None:
                            if (
                                hay_muro_sup
                                and vert_emp is not None
                                and (not muro_emp_preaplicado)
                            ):
                                try:
                                    _fn_emp = getattr(
                                        vert_emp,
                                        u"extender_vertical_cabeza_empotramiento_por_diam",
                                        None,
                                    )
                                    if callable(_fn_emp):
                                        _ok_emp, _msg_emp, _rb_emp = _fn_emp(
                                            doc,
                                            final_rb,
                                            int(
                                                getattr(
                                                    l135,
                                                    u"INDICE_POSICION",
                                                    0,
                                                )
                                            ),
                                        )
                                        if _ok_emp and _rb_emp is not None:
                                            final_rb = _rb_emp
                                            _embed_muro_sup_ok_en_path_l135 = (
                                                True
                                            )
                                            if _msg_emp:
                                                msg_last = (
                                                    (msg_last or u"")
                                                    + u"  "
                                                    + _msg_emp
                                                )
                                except Exception as ex_emp:  # noqa: BLE001
                                    msg_last = (msg_last or u"") + (
                                        u"  [vert. empotr. muro sup.: {0!s}]"
                                    ).format(ex_emp)
                            c_vert_l135_ok += 1
                            rebar_finales.append(final_rb.Id)
                            midv = _eid(final_rb.Id)
                            if need_l135_cabeza and need_l135_pie:
                                tag = u"OK vert. L+135 (cabeza+pie: L pie, 135 cabeza)"
                            elif need_l135_cabeza:
                                tag = u"OK vert. L+135 (cabeza, sin forj. sup.)"
                            else:
                                tag = u"OK vert. L+135 (pie, sin unión inf.)"
                            ms.append(
                                (msg_last or tag)
                                + (u"  [nuevo: {0}]".format(midv))
                            )
                        else:
                            c_vert_l135_fail += 1
                            # Si el 1.er extender tuvo éxito, el id de la malla original ya no
                            # existe; el 2.º paso puede fallar (p. ej. InternalException) y no
                            # debemos devolver eid obsoleto a rebar_finales (visibilidad / «perdida»).
                            _append_id = eid
                            try:
                                if doc.GetElement(eid) is not None:
                                    _append_id = eid
                                elif (
                                    r_work is not None
                                    and doc.GetElement(r_work.Id) is not None
                                ):
                                    _append_id = r_work.Id
                            except Exception:  # noqa: BLE001
                                try:
                                    if (
                                        r_work is not None
                                        and doc.GetElement(r_work.Id) is not None
                                    ):
                                        _append_id = r_work.Id
                                except Exception:  # noqa: BLE001
                                    pass
                            rebar_finales.append(_append_id)
                            ms.append(
                                u"FALLO vert. L+135 id {0}: {1}".format(
                                    _eid(eid), (msg_last or u"")
                                )
                            )
                        continue
        # Sin requisito L+135 (p. ej. forjado en sup.): muro apilado aun pide
        # estirar cabeza con tabla (sketch).
        if (
            hay_muro_sup
            and vert_emp is not None
            and (not _embed_muro_sup_ok_en_path_l135)
            and (not (need_l135_cabeza or need_l135_pie))
            and _rebar_es_vertical_muro_criterio(r, host, 0)
        ):
            _c_emp_sup = _face_cara_ext_o_int_desde_pata_o_arex(r, host)
            if _c_emp_sup is not None:
                _fn_emp2 = getattr(
                    vert_emp,
                    u"extender_vertical_cabeza_empotramiento_por_diam",
                    None,
                )
                if callable(_fn_emp2):
                    try:
                        _o2, _m2, _rb2 = _fn_emp2(
                            doc,
                            r,
                            int(
                                getattr(l135, u"INDICE_POSICION", 0)
                            ),
                        )
                        if _o2 and _rb2 is not None:
                            r = _rb2
                            if _m2:
                                ms.append(
                                    u"OK vert. empotram. muro sup.  {0}".format(
                                        _m2
                                    )[:500]
                                )
                    except Exception:  # noqa: BLE001
                        pass
        if (
            hay_muro_inf
            and vert_emp is not None
            and (not (need_l135_cabeza or need_l135_pie))
            and _rebar_es_vertical_muro_criterio(r, host, 0)
        ):
            _c_emp_inf = _face_cara_ext_o_int_desde_pata_o_arex(r, host)
            if _c_emp_inf is not None:
                _fn_emp3 = getattr(
                    vert_emp,
                    u"extender_vertical_pie_emp_muro_inf_mm",
                    None,
                )
                if callable(_fn_emp3):
                    try:
                        _o3, _m3, _rb3 = _fn_emp3(
                            doc,
                            r,
                            int(
                                getattr(l135, u"INDICE_POSICION", 0)
                            ),
                        )
                        if _o3 and _rb3 is not None:
                            r = _rb3
                            if _m3:
                                ms.append(
                                    u"OK vert. 25 mm pie muro inf.  {0}".format(
                                        _m3
                                    )[:500]
                                )
                    except Exception:  # noqa: BLE001
                        pass
        if hay_fund and pata90 is not None and hay_forjado_inf:
            cara = _face_cara_ext_o_int_desde_pata_o_arex(r, host)
            if cara is not None and _rebar_es_vertical_muro_criterio(r, host, 0):
                c_vert_pata += 1
                Lp = pata90._largo_pata_mm_resuelto(doc, host)
                invert = bool(
                    getattr(pata90, "INVERTIR_DIRECCION_PATA", False)
                )
                if cara == u"int":
                    invert = not invert
                ok, msg, new_rb = pata90._add_pata_90_fin_a_rebar(
                    doc, r, Lp, invert, getattr(pata90, "INDICE_POSICION", 0)
                )
                if ok and new_rb is not None:
                    c_vert_pata_ok += 1
                    rebar_finales.append(new_rb.Id)
                    ms.append(
                        (msg or u"OK vert. pata 90 id {0}").format(
                            _eid(new_rb.Id)
                        )
                    )
                else:
                    c_vert_pata_fail += 1
                    rebar_finales.append(eid)
                    ms.append(
                        u"FALLO vert. pata 90 id {0}: {1}".format(
                            _eid(eid), (msg or u"")
                        )
                    )
                continue
        c_unchanged += 1
        _eid_cierre = eid
        try:
            if r is not None and isinstance(r, Rebar):
                if doc.GetElement(eid) is None and doc.GetElement(r.Id) is not None:
                    _eid_cierre = r.Id
        except Exception:  # noqa: BLE001
            pass
        rebar_finales.append(_eid_cierre)

    t_ex = Transaction(doc, u"BIMTools: malla nudo — excluir 1.ª y últ. barra (cada set)")
    t_ex.Start()
    try:
        n_ex = 0
        _vistos = set()
        for eid0 in rebar_finales:
            k = _eid(eid0)
            if k in _vistos:
                continue
            _vistos.add(k)
            rb = doc.GetElement(eid0)
            if rb is not None and isinstance(rb, Rebar) and desactivar_extremos_rebar_set(rb, doc):
                n_ex += 1
        t_ex.Commit()
        if n_ex:
            ms.append(
                u"Excl. 1.ª/últ. posición: {0} rebar con extremos desactivados.".format(
                    n_ex
                )
            )
    except Exception as exex:  # noqa: BLE001
        try:
            t_ex.RollBack()
        except Exception:
            pass
        ms.append(
            u"Excl. extremos rebar: {0!s}".format(exex)
        )

    if hay_fund and pata90 is None:
        ms.append(
            u"(Aviso) Fundación: no se importó pata 90° vertical: {0}".format(
                _P90_ERR or u"—"
            )
        )
    return rebar_finales, ms
