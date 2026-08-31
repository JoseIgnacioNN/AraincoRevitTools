# -*- coding: utf-8 -*-
"""Configuración de ø y n por tramo Tn (no global al lote).

Varios Tn pueden compartir vigas en el empalme; la fuente de verdad es
``session.tramo_armado`` indexada por topología (índices + edges).
No se escribe ø/n en las vigas del tramo (evita que el último Tn pise al vecino).
"""

from armado_vigas.domain.constants import BAR_COUNT_MIN
from armado_vigas.domain.layers import (
    beam_n_capas_inf,
    beam_n_capas_sup,
    clamp_bar_count,
    ensure_beam_layers,
    layer_keys,
    set_first_layer_bar_count,
)

# Campos de armado longitudinal por tramo (capas independientes).
TRAMO_ARM_FIELDS = (
    u"nSup",
    u"nInf",
    u"nSup2",
    u"nInf2",
    u"nSup3",
    u"nInf3",
    u"diamSup",
    u"diamInf",
    u"diamSup2",
    u"diamInf2",
    u"diamSup3",
    u"diamInf3",
)

TRAMO_ARM_FIELDS_SUP = (
    u"nSup",
    u"nSup2",
    u"nSup3",
    u"diamSup",
    u"diamSup2",
    u"diamSup3",
)

TRAMO_ARM_FIELDS_INF = (
    u"nInf",
    u"nInf2",
    u"nInf3",
    u"diamInf",
    u"diamInf2",
    u"diamInf3",
)


def _face_tag(face):
    if face in (u"inf", u"INF", True, 1, u"inferior"):
        return u"inf"
    return u"sup"


def face_arm_fields(face):
    return TRAMO_ARM_FIELDS_INF if _face_tag(face) == u"inf" else TRAMO_ARM_FIELDS_SUP


def ensure_session_tramo_armado(session):
    if session is None:
        return {u"sup": {}, u"inf": {}}
    store = getattr(session, u"tramo_armado", None)
    if not isinstance(store, dict):
        store = {u"sup": {}, u"inf": {}}
        session.tramo_armado = store
    if u"sup" not in store or not isinstance(store.get(u"sup"), dict):
        store[u"sup"] = {}
    if u"inf" not in store or not isinstance(store.get(u"inf"), dict):
        store[u"inf"] = {}
    return store


def tramo_topo_key(face, tramo):
    """Clave estable aunque el id Tn renumere al recalcular."""
    if not tramo:
        return None
    idxs = tramo.get(u"beamIndices") or []
    try:
        idx_s = u",".join(str(int(i)) for i in idxs)
    except Exception:
        idx_s = u",".join(str(i) for i in idxs)
    es = tramo.get(u"edgeStart") or u""
    ee = tramo.get(u"edgeEnd") or u""
    return u"{0}|{1}|{2}|{3}".format(_face_tag(face), idx_s, es, ee)


def _parse_topo_key(key):
    """``(face, idxs_set, edgeStart, edgeEnd)`` o ``None``."""
    if not key:
        return None
    try:
        parts = _safe_split_key(key)
        if len(parts) < 4:
            return None
        face, idx_s, es, ee = parts[0], parts[1], parts[2], parts[3]
        idxs = set()
        if idx_s:
            for tok in idx_s.split(u","):
                if tok == u"":
                    continue
                idxs.add(int(tok))
        return face, idxs, es or u"", ee or u""
    except Exception:
        return None


def _safe_split_key(key):
    s = u"{0}".format(key)
    return s.split(u"|")


def get_tramo_arm_cfg(session, face, tramo, create=False):
    """Devuelve dict de config del tramo (o vacío)."""
    store = ensure_session_tramo_armado(session)
    key = tramo_topo_key(face, tramo)
    if not key:
        return {}
    face_map = store[_face_tag(face)]
    cfg = face_map.get(key)
    if cfg is None:
        if not create:
            return {}
        cfg = {}
        face_map[key] = cfg
    return cfg


def set_tramo_arm_field(session, face, tramos, field, value):
    """Escribe ``field`` en la config de cada tramo de la lista.

    ``nSup`` / ``nInf`` (1ª capa) se escriben **ligados** en SUP e INF para las
    vigas del tramo (misma n en ambas caras → conf. E/T alineado).
    """
    if field not in TRAMO_ARM_FIELDS:
        return
    if field in (u"nSup", u"nInf"):
        set_first_layer_n_linked(session, tramos, value, source_face=face)
        return
    for tramo in tramos or []:
        cfg = get_tramo_arm_cfg(session, face, tramo, create=True)
        if field in (u"nSup2", u"nInf2", u"nSup3", u"nInf3"):
            cfg[field] = clamp_bar_count(value)
        else:
            try:
                cfg[field] = int(round(float(value)))
            except Exception:
                cfg[field] = value
        try:
            cfg[u"_tramo_id"] = tramo.get(u"id")
        except Exception:
            pass


def _beam_indices_of_tramos(tramos):
    out = set()
    for t in tramos or []:
        for raw in t.get(u"beamIndices") or []:
            try:
                out.add(int(raw))
            except Exception:
                continue
    return out


def set_first_layer_n_linked(session, tramos, value, source_face=None):
    """Liga n 1ª capa SUP↔INF en tramos que comparten las vigas de ``tramos``.

    El confinamiento indexa columnas 0…n−1 en sup e inf: n debe ser la misma.
    """
    n = clamp_bar_count(value)
    idxs = _beam_indices_of_tramos(tramos)
    if not idxs and not tramos:
        return n
    ensure_session_tramo_armado(session)

    def _write_cfg(face_key, tramo):
        cfg = get_tramo_arm_cfg(session, face_key, tramo, create=True)
        cfg[u"nSup"] = n
        cfg[u"nInf"] = n
        try:
            cfg[u"_tramo_id"] = tramo.get(u"id")
            cfg[u"_n1_linked"] = True
        except Exception:
            pass

    for face_key in (u"sup", u"inf"):
        attr = u"tramos_sup" if face_key == u"sup" else u"tramos_inf"
        face_list = list(getattr(session, attr, None) or [])
        written = set()
        for t in face_list:
            t_idxs = _beam_indices_of_tramos([t])
            if idxs and (t_idxs & idxs):
                _write_cfg(face_key, t)
                written.add(id(t))
        # Garantiza escritura en la selección (aunque no figurese en session aún).
        for t in tramos or []:
            if id(t) in written:
                continue
            t_idxs = _beam_indices_of_tramos([t])
            if idxs and t_idxs and not (t_idxs & idxs):
                continue
            _write_cfg(face_key, t)
    return n


def resolve_first_layer_n_linked(session, face, tramo, fallback_beam=None):
    """n 1ª capa efectiva (SUP↔INF): max de ambos si difieren; fallback viga."""
    n_s = tramo_layer_bar_count(session, u"sup", tramo, 1, False, fallback_beam) if tramo else None
    n_i = tramo_layer_bar_count(session, u"inf", tramo, 1, True, fallback_beam) if tramo else None
    # Si el tramo de esta cara no tiene la otra, buscar por índices en la otra cara.
    if tramo is not None and session is not None:
        idxs = _beam_indices_of_tramos([tramo])
        other = u"inf" if _face_tag(face) == u"sup" else u"sup"
        attr = u"tramos_inf" if other == u"inf" else u"tramos_sup"
        for t in list(getattr(session, attr, None) or []):
            if not (_beam_indices_of_tramos([t]) & idxs):
                continue
            if other == u"sup":
                n_s = tramo_layer_bar_count(session, u"sup", t, 1, False, fallback_beam)
            else:
                n_i = tramo_layer_bar_count(session, u"inf", t, 1, True, fallback_beam)
            break
    vals = []
    for v in (n_s, n_i):
        if v is not None:
            try:
                vals.append(clamp_bar_count(v))
            except Exception:
                pass
    if vals:
        return max(vals)
    if fallback_beam is not None:
        from armado_vigas.domain.layers import first_layer_bar_count

        return first_layer_bar_count(fallback_beam)
    return BAR_COUNT_MIN


def apply_tramo_arm_to_beams(tramos, beams, face, field, value):
    """
    Espejo de ø/n en las vigas de los tramos (preview de sección / conf.).

    Así, aunque el paint no resuelva el Tn, la viga no se queda en defaults.
    En empalme, cada viga pertenece a un solo Tn por cara (no se pisan vecinos).
    """
    if not beams or not tramos:
        return
    seen = set()
    for tramo in tramos or []:
        for raw in tramo.get(u"beamIndices") or []:
            try:
                ii = int(raw)
            except Exception:
                continue
            if ii < 0 or ii >= len(beams) or ii in seen:
                continue
            seen.add(ii)
            beam = beams[ii]
            if field in (u"nSup", u"nInf"):
                set_first_layer_bar_count(beam, value)
            else:
                beam[field] = value
            ensure_beam_layers(beam)


def tramo_layer_diam(session, face, tramo, layer_num, es_cara_inferior, fallback_beam=None):
    """ø (mm) de capa del tramo; fallback a viga si no hay config."""
    k = layer_keys(layer_num)
    field = k[u"diamInf"] if es_cara_inferior else k[u"diamSup"]
    cfg = get_tramo_arm_cfg(session, face, tramo, create=False)
    if field in cfg and cfg.get(field) is not None:
        try:
            return int(cfg[field])
        except Exception:
            pass
    beam = fallback_beam
    if beam is not None:
        ensure_beam_layers(beam)
        v = beam.get(field)
        if v is not None:
            try:
                return int(v)
            except Exception:
                pass
        return int(beam.get(u"diamInf" if es_cara_inferior else u"diamSup") or 16)
    return 16


def tramo_layer_bar_count(session, face, tramo, layer_num, es_cara_inferior, fallback_beam=None):
    """n barras de capa del tramo; fallback a viga."""
    k = layer_keys(layer_num)
    field = k[u"nInf"] if es_cara_inferior else k[u"nSup"]
    cfg = get_tramo_arm_cfg(session, face, tramo, create=False)
    if field in cfg and cfg.get(field) is not None:
        return clamp_bar_count(cfg[field])
    beam = fallback_beam
    if beam is not None:
        ensure_beam_layers(beam)
        return clamp_bar_count(beam.get(field) or BAR_COUNT_MIN)
    return BAR_COUNT_MIN


def owner_display_value(session, face, tramo, field, fallback_beam=None):
    """Valor a mostrar en el rail / panel para el tramo."""
    cfg = get_tramo_arm_cfg(session, face, tramo, create=False)
    if field in cfg and cfg.get(field) is not None:
        return cfg[field]
    if fallback_beam is not None:
        return fallback_beam.get(field)
    return None


def _cfg_has_face_values(cfg, face):
    if not cfg:
        return False
    return any(f in cfg and cfg.get(f) is not None for f in face_arm_fields(face))


def _match_score(new_idxs, new_es, new_ee, old_key):
    """Mayor = mejor legado al trocear empalme (subconjunto / jaccard + edges)."""
    parsed = _parse_topo_key(old_key)
    if not parsed:
        return -1.0
    _face, old_idxs, old_es, old_ee = parsed
    if not new_idxs and not old_idxs:
        return 0.0
    new_set = set(new_idxs)
    inter = new_set & old_idxs
    if not inter:
        return -1.0
    union = new_set | old_idxs
    jacc = float(len(inter)) / float(len(union) or 1)
    # Preferir solape fuerte y que el nuevo sea subconjunto del viejo (troceo).
    subset_bonus = 0.35 if new_set <= old_idxs else 0.0
    edge_bonus = 0.0
    if (new_es or u"") == (old_es or u""):
        edge_bonus += 0.08
    if (new_ee or u"") == (old_ee or u""):
        edge_bonus += 0.08
    # Preferir el run antiguo más grande del que se troceó (más índices).
    size_bonus = 0.02 * min(len(old_idxs), 20)
    return jacc + subset_bonus + edge_bonus + size_bonus


def _best_legacy_cfg(face_map, face, tramo, current_key):
    """Busca cfg de una clave antigua del mismo troceo para heredar ø/n."""
    idxs = tramo.get(u"beamIndices") or []
    try:
        new_idxs = [int(i) for i in idxs]
    except Exception:
        new_idxs = list(idxs)
    new_es = tramo.get(u"edgeStart") or u""
    new_ee = tramo.get(u"edgeEnd") or u""
    best_score = 0.15  # umbral mínimo (exige algo de overlap)
    best_cfg = None
    best_key = None
    for old_key, old_cfg in (face_map or {}).items():
        if old_key == current_key or not old_cfg:
            continue
        if not _cfg_has_face_values(old_cfg, face):
            continue
        sc = _match_score(new_idxs, new_es, new_ee, old_key)
        if sc > best_score:
            best_score = sc
            best_cfg = old_cfg
            best_key = old_key
    return best_key, best_cfg, best_score


def _copy_face_fields(src_cfg, dst_cfg, face):
    for field in face_arm_fields(face):
        if field in src_cfg and src_cfg.get(field) is not None:
            dst_cfg[field] = src_cfg[field]


def seed_tramo_arm_cfg(session, face, tramo, beam=None):
    """
    Si el tramo no tiene cfg de su cara: hereda de clave legada (troceo empalme)
    o, en última instancia, de la viga owner **solo campos de esa cara**.
    """
    if tramo is None:
        return
    face = _face_tag(face)
    store = ensure_session_tramo_armado(session)
    face_map = store[face]
    key = tramo_topo_key(face, tramo)
    if not key:
        return
    cfg = face_map.get(key)
    if cfg is None:
        cfg = {}
        face_map[key] = cfg
    if _cfg_has_face_values(cfg, face):
        return

    # 1) Migrar desde clave del run anterior (p. ej. full → half/half).
    leg_key, leg_cfg, _score = _best_legacy_cfg(face_map, face, tramo, key)
    if leg_cfg is not None:
        _copy_face_fields(leg_cfg, cfg, face)
        try:
            cfg[u"_tramo_id"] = tramo.get(u"id")
            cfg[u"_migrated_from"] = leg_key
        except Exception:
            pass
        return

    # 2) Seed desde viga: solo campos de la cara (no contaminar diam de la otra).
    if beam is not None:
        ensure_beam_layers(beam)
        for field in face_arm_fields(face):
            if beam.get(field) is not None:
                cfg[field] = beam[field]
        try:
            cfg[u"_tramo_id"] = tramo.get(u"id")
        except Exception:
            pass


def seed_tramo_arm_from_beam(session, face, tramo, beam):
    """Compat: redirige a :func:`seed_tramo_arm_cfg`."""
    seed_tramo_arm_cfg(session, face, tramo, beam)


def prune_tramo_arm_store(session, face, tramos):
    """Elimina claves huérfanas que ya no corresponden a un Tn activo."""
    store = ensure_session_tramo_armado(session)
    face = _face_tag(face)
    face_map = store[face]
    active = set()
    for t in tramos or []:
        k = tramo_topo_key(face, t)
        if k:
            active.add(k)
    for k in list(face_map.keys()):
        if k not in active:
            try:
                del face_map[k]
            except Exception:
                pass


def merge_armado_onto_tramos(session, face, tramos, beams=None, prune=True):
    """Siembra / migra cfg, poda huérfanas y adjunta ``armado`` a cada Tn."""
    face = _face_tag(face)
    for tramo in tramos or []:
        owner = None
        idxs = tramo.get(u"beamIndices") or []
        if beams and idxs:
            try:
                ii = int(idxs[0])
                if 0 <= ii < len(beams):
                    owner = beams[ii]
            except Exception:
                owner = None
        seed_tramo_arm_cfg(session, face, tramo, owner)
        cfg = get_tramo_arm_cfg(session, face, tramo, create=False)
        tramo[u"armado"] = dict(cfg) if cfg else {}
    if prune:
        prune_tramo_arm_store(session, face, tramos)


def _tramos_containing_beam(tramos, beam_idx):
    out = []
    if beam_idx is None:
        return out
    try:
        bi = int(beam_idx)
    except Exception:
        return out
    for t in tramos or []:
        for raw in t.get(u"beamIndices") or []:
            try:
                if int(raw) == bi:
                    out.append(t)
                    break
            except Exception:
                continue
    return out


def _find_tramo_by_id(tramos, tramo_id):
    if tramo_id is None:
        return None
    for t in tramos or []:
        if t.get(u"id") == tramo_id:
            return t
    return None


def _pick_tramo_for_beam(
    tramos, beam_idx, preferred_tramo=None, preferred_id=None, session=None, face=None,
):
    """Tramo de la cara que incluye la viga.

    Prioridad: preferido (obj/id) si contiene la viga → cand. con cfg de armado
    en esa cara → primer candidato.
    """
    cands = _tramos_containing_beam(tramos, beam_idx)
    if not cands:
        return None
    preferred = preferred_tramo
    if preferred is None and preferred_id is not None:
        preferred = _find_tramo_by_id(tramos, preferred_id)
    if preferred is not None:
        pid = preferred.get(u"id")
        for t in cands:
            if t is preferred or (pid is not None and t.get(u"id") == pid):
                return t
    # Preferir tramo con armado ya configurado en SUP/INF.
    if session is not None and face is not None:
        for t in cands:
            try:
                cfg = get_tramo_arm_cfg(session, face, t, create=False)
                if _cfg_has_face_values(cfg, face):
                    return t
            except Exception:
                continue
    return cands[0]


def _arm_cfg_for_beam_idx(session, face, beam_idx):
    """Fallback: cfg en ``tramo_armado`` cuya topología contiene la viga."""
    if session is None or beam_idx is None:
        return {}
    try:
        bi = int(beam_idx)
    except Exception:
        return {}
    store = ensure_session_tramo_armado(session)
    face_map = store.get(_face_tag(face)) or {}
    best_cfg = None
    best_score = -1.0
    for key, cfg in face_map.items():
        if not cfg or not isinstance(cfg, dict):
            continue
        if not _cfg_has_face_values(cfg, face):
            continue
        parsed = _parse_topo_key(key)
        if not parsed:
            continue
        _f, idxs, _es, _ee = parsed
        if bi not in idxs:
            continue
        # Preferir run más específico (menos vigas = más local al tramo).
        score = 100.0 / float(max(1, len(idxs)))
        if score > best_score:
            best_score = score
            best_cfg = cfg
    return best_cfg or {}


def _layer_n_from_sources(session, face, tramo, layer_num, es_inf, beam, beam_idx):
    """n de capa: cfg del Tn → store por índice → viga."""
    k = layer_keys(layer_num)
    field = k[u"nInf"] if es_inf else k[u"nSup"]
    if tramo is not None:
        return tramo_layer_bar_count(
            session, face, tramo, layer_num, es_inf, beam
        )
    cfg = _arm_cfg_for_beam_idx(session, face, beam_idx)
    if field in cfg and cfg.get(field) is not None:
        return clamp_bar_count(cfg[field])
    if beam is not None:
        ensure_beam_layers(beam)
        return clamp_bar_count(beam.get(field) or BAR_COUNT_MIN)
    return BAR_COUNT_MIN


def _layer_diam_from_sources(session, face, tramo, layer_num, es_inf, beam, beam_idx):
    """ø de capa: cfg del Tn → store por índice → viga."""
    k = layer_keys(layer_num)
    field = k[u"diamInf"] if es_inf else k[u"diamSup"]
    if tramo is not None:
        return tramo_layer_diam(
            session, face, tramo, layer_num, es_inf, beam
        )
    cfg = _arm_cfg_for_beam_idx(session, face, beam_idx)
    if field in cfg and cfg.get(field) is not None:
        try:
            return int(cfg[field])
        except Exception:
            pass
    if beam is not None:
        ensure_beam_layers(beam)
        v = beam.get(field)
        if v is not None:
            try:
                return int(v)
            except Exception:
                pass
        return int(beam.get(u"diamInf" if es_inf else u"diamSup") or 16)
    return 16


def beam_section_arm_for_preview(
    session,
    beam,
    beams=None,
    beam_idx=None,
    preferred_tramo_sup=None,
    preferred_tramo_inf=None,
    preferred_tramo_sup_id=None,
    preferred_tramo_inf_id=None,
):
    """
    Vista de n/ø longitudinales efectivos para dibujar la sección de una viga.

    Fuente de verdad: ``session.tramo_armado`` (lo configurado en pestañas SUP/INF
    por tramo Tn). Capas sup desde SUP, inf desde INF. 1ª capa n queda ligada
    (max) para columnas de confinamiento E/T.

    Devuelve un dict listo para pintar (no es un alias del beam). Opcionalmente
    se pueden sincronizar n/ø en el beam de dominio con
    :func:`sync_resolved_arm_onto_beam`.
    """
    if beam is None:
        return None
    if session is None:
        try:
            from armado_vigas.revit.session import SESSION as _S

            session = _S
        except Exception:
            session = None

    ensure_beam_layers(beam)
    view = dict(beam)
    beams = list(beams or [])
    if beam_idx is None and beams:
        try:
            beam_idx = beams.index(beam)
        except ValueError:
            try:
                bid = beam.get(u"id")
                for i, b in enumerate(beams):
                    if b is beam or (bid is not None and b.get(u"id") == bid):
                        beam_idx = i
                        break
            except Exception:
                beam_idx = None

    tramos_sup = list(getattr(session, u"tramos_sup", None) or []) if session else []
    tramos_inf = list(getattr(session, u"tramos_inf", None) or []) if session else []

    t_sup = _pick_tramo_for_beam(
        tramos_sup,
        beam_idx,
        preferred_tramo=preferred_tramo_sup,
        preferred_id=preferred_tramo_sup_id,
        session=session,
        face=u"sup",
    )
    t_inf = _pick_tramo_for_beam(
        tramos_inf,
        beam_idx,
        preferred_tramo=preferred_tramo_inf,
        preferred_id=preferred_tramo_inf_id,
        session=session,
        face=u"inf",
    )

    # Si el preferido no contiene la viga (p. ej. otro Tn activo), buscar
    # cualquier Tn de esa cara con armado que sí la contenga.
    if t_sup is None and preferred_tramo_sup is not None:
        t_sup = _pick_tramo_for_beam(
            tramos_sup, beam_idx, session=session, face=u"sup",
        )
    if t_inf is None and preferred_tramo_inf is not None:
        t_inf = _pick_tramo_for_beam(
            tramos_inf, beam_idx, session=session, face=u"inf",
        )

    if session is not None:
        if t_sup is not None:
            seed_tramo_arm_cfg(session, u"sup", t_sup, beam)
        if t_inf is not None:
            seed_tramo_arm_cfg(session, u"inf", t_inf, beam)

    n_cap_s = beam_n_capas_sup(beam)
    n_cap_i = beam_n_capas_inf(beam)
    view[u"nCapasSup"] = n_cap_s
    view[u"nCapasInf"] = n_cap_i

    for layer in range(1, n_cap_s + 1):
        k = layer_keys(layer)
        view[k[u"nSup"]] = _layer_n_from_sources(
            session, u"sup", t_sup, layer, False, beam, beam_idx,
        )
        view[k[u"diamSup"]] = _layer_diam_from_sources(
            session, u"sup", t_sup, layer, False, beam, beam_idx,
        )

    for layer in range(1, n_cap_i + 1):
        k = layer_keys(layer)
        view[k[u"nInf"]] = _layer_n_from_sources(
            session, u"inf", t_inf, layer, True, beam, beam_idx,
        )
        view[k[u"diamInf"]] = _layer_diam_from_sources(
            session, u"inf", t_inf, layer, True, beam, beam_idx,
        )

    # 1ª capa ligada: columnas E/T (mismo nº sup/inf).
    # Preferir valor ligado explícito si ya se escribió en store.
    n1 = None
    if session is not None and (t_sup is not None or t_inf is not None):
        try:
            n1 = resolve_first_layer_n_linked(
                session, u"sup" if t_sup is not None else u"inf",
                t_sup or t_inf, beam,
            )
        except Exception:
            n1 = None
    if n1 is None:
        ns = clamp_bar_count(view.get(u"nSup") or BAR_COUNT_MIN)
        ni = clamp_bar_count(view.get(u"nInf") or BAR_COUNT_MIN)
        n1 = max(ns, ni, BAR_COUNT_MIN)
    else:
        n1 = max(clamp_bar_count(n1), BAR_COUNT_MIN)
    view[u"nSup"] = n1
    view[u"nInf"] = n1
    return view


def sync_resolved_arm_onto_beam(beam, view):
    """Copia n/ø/capas resueltos al beam de dominio (conf. E/T y fallbacks)."""
    if beam is None or not view:
        return beam
    try:
        n_cap_s = int(view.get(u"nCapasSup") or beam_n_capas_sup(beam))
        n_cap_i = int(view.get(u"nCapasInf") or beam_n_capas_inf(beam))
    except Exception:
        n_cap_s = beam_n_capas_sup(beam)
        n_cap_i = beam_n_capas_inf(beam)
    beam[u"nCapasSup"] = n_cap_s
    beam[u"nCapasInf"] = n_cap_i
    for layer in range(1, n_cap_s + 1):
        k = layer_keys(layer)
        if view.get(k[u"nSup"]) is not None:
            beam[k[u"nSup"]] = clamp_bar_count(view[k[u"nSup"]])
        if view.get(k[u"diamSup"]) is not None:
            try:
                beam[k[u"diamSup"]] = int(view[k[u"diamSup"]])
            except Exception:
                pass
    for layer in range(1, n_cap_i + 1):
        k = layer_keys(layer)
        if view.get(k[u"nInf"]) is not None:
            beam[k[u"nInf"]] = clamp_bar_count(view[k[u"nInf"]])
        if view.get(k[u"diamInf"]) is not None:
            try:
                beam[k[u"diamInf"]] = int(view[k[u"diamInf"]])
            except Exception:
                pass
    # 1ª capa n ligada
    try:
        set_first_layer_bar_count(
            beam, view.get(u"nSup") or view.get(u"nInf") or BAR_COUNT_MIN
        )
    except Exception:
        pass
    ensure_beam_layers(beam)
    return beam
