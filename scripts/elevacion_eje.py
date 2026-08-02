# -*- coding: utf-8 -*-
"""
Elevación Eje — motor Revit API.

Revit 2024+ | pyRevit | IronPython 2.7 / 3.4

Por cada Grid seleccionado crea una ViewSection de tipo Building Section
(perpendicular al eje, misma convención que Sección alzado por eje) y aplica
la escala elegida en la UI (``View.Scale``).

Siempre genera el contorno de hormigón (unión booleana → corte por
plano del eje → detail lines agrupadas con el nombre de la vista).

En cada elevación etiqueta:
  - muros Concrete paralelos al plano (``EST_A_WALL TAG_ELEVACION_MHA`` / ``Espesor Muro``);
  - vigas Concrete paralelas al plano (``EST_A_STRUCTURAL FRAMING TAG_ELEVACION`` / ``Tag Viga``),
    con cabeza en la cara superior del bbox.

El ``ViewFamilyType`` se elige leyendo «Section Filter» de la planta activa
y buscando un tipo Building Section cuyo nombre contenga ese valor.

Transacciones: ``TransactionGroup`` (``Arainco: Elevación Eje``) + una
``Transaction`` por eje + ``Assimilate()`` → un solo Undo; si un eje falla,
se continúa con el resto.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    ElementId,
    FailureProcessingResult,
    FailureResolutionType,
    FailureSeverity,
    FilteredElementCollector,
    IFailuresPreprocessor,
    StorageType,
    Transaction,
    TransactionGroup,
    TransactionStatus,
    View,
    ViewSection,
)

from seccion_alzado_eje import (
    bbox_modelo_documento,
    crear_seccion_alzado,
    listar_tipos_seccion_building,
    recopilar_nombres_vistas,
)
from seccion_detalle_extremo_muro import leer_section_filter_texto

_TOOL_TITLE = u"Arainco: Elevación Eje"
_TX_NAME = u"Arainco: Elevación Eje"
_VISTA_MID = u"ELEVACION EJE"
_PARAM_ARMADURA_EJE = u"Armadura_Eje"
_PROGRESS_ACCENT_RGB = (91, 192, 222)
_WARNING_SWALLOWER = None


def _failure_description_text(fma):
    if fma is None:
        return u""
    for meth in (
        u"GetDescriptionText",
        u"GetDescriptionString",
        u"GetDescription",
    ):
        try:
            t = getattr(fma, meth)()
            if t:
                return _as_unicode(t)
        except Exception:
            pass
    try:
        t = fma.GetDefaultResolutionCaption()
        if t:
            return _as_unicode(t)
    except Exception:
        pass
    return u""


def _failure_is_line_too_short(fma):
    """Detecta el error/warning «Line is too short.» (y equivalentes)."""
    s = _failure_description_text(fma).lower()
    if not s:
        return False
    if u"line is too short" in s:
        return True
    if u"too short" in s and (u"line" in s or u"curve" in s or u"línea" in s or u"linea" in s):
        return True
    if u"línea demasiado corta" in s or u"linea demasiado corta" in s:
        return True
    if u"curva demasiado corta" in s:
        return True
    return False


def _try_resolve_delete_elements(failures_accessor, fma):
    """
    Aplica la resolución «Delete Element(s)» del diálogo de Revit.
    Devuelve True si se resolvió el fallo.
    """
    if failures_accessor is None or fma is None:
        return False

    def _permitted(rt):
        try:
            return bool(failures_accessor.IsFailureResolutionPermitted(fma, rt))
        except Exception:
            pass
        try:
            if hasattr(fma, u"HasResolutionOfType"):
                return bool(fma.HasResolutionOfType(rt))
        except Exception:
            pass
        return False

    rt_del = getattr(FailureResolutionType, u"DeleteElements", None)
    if rt_del is not None and _permitted(rt_del):
        try:
            failures_accessor.SetCurrentResolutionType(fma, rt_del)
            failures_accessor.ResolveFailure(fma)
            return True
        except Exception:
            pass

    # Resolución por defecto (suele ser Delete Elements en este error).
    try:
        has_res = False
        try:
            if hasattr(fma, u"HasResolutions"):
                has_res = bool(fma.HasResolutions())
        except Exception:
            has_res = False
        if has_res:
            try:
                if hasattr(failures_accessor, u"IsFailureResolutionPermitted"):
                    if not bool(failures_accessor.IsFailureResolutionPermitted(fma)):
                        return False
            except (Exception, TypeError):
                pass
            failures_accessor.ResolveFailure(fma)
            return True
    except Exception:
        pass
    return False


class _ElevacionEjeWarningSwallower(IFailuresPreprocessor):
    """
    Silencia warnings y resuelve «Line is too short» con Delete Element(s).
    """

    def _iter_failure_msgs(self, failures_accessor):
        if failures_accessor is None:
            return
        try:
            fmsgs = failures_accessor.GetFailureMessages()
        except Exception:
            return
        if fmsgs is None:
            return
        try:
            n = int(fmsgs.Count)
        except Exception:
            n = 0
        for i in range(n):
            f = None
            try:
                f = fmsgs.get_Item(i)
            except Exception:
                try:
                    f = fmsgs[i]
                except Exception:
                    f = None
            if f is not None:
                yield f

    def PreprocessFailures(self, failures_accessor):
        if failures_accessor is None:
            return FailureProcessingResult.Continue
        # Recolectar primero: mutar mientras se itera puede tumbar Revit.
        msgs = list(self._iter_failure_msgs(failures_accessor))
        resolved_or_cleared = False
        for f in msgs:
            try:
                sev = f.GetSeverity()
            except Exception:
                continue
            if sev == FailureSeverity.Warning:
                # Algunos «too short» llegan como Warning; otros como Error.
                if _failure_is_line_too_short(f):
                    if _try_resolve_delete_elements(failures_accessor, f):
                        resolved_or_cleared = True
                        continue
                try:
                    failures_accessor.DeleteWarning(f)
                    resolved_or_cleared = True
                except Exception:
                    pass
                continue
            if sev == FailureSeverity.Error and _failure_is_line_too_short(f):
                if _try_resolve_delete_elements(failures_accessor, f):
                    resolved_or_cleared = True
                    continue
                try:
                    # Respaldo: quitar el error de la cola si la API lo permite.
                    failures_accessor.DeleteError(f)
                    resolved_or_cleared = True
                except Exception:
                    pass
        # Continue mostraría el diálogo aunque el fallo se haya tratado.
        if resolved_or_cleared:
            return FailureProcessingResult.ProceedWithCommit
        return FailureProcessingResult.Continue


def _attach_warning_swallower(txn):
    """Adjunta el preprocessor de fallos a la ``Transaction`` (singleton CLR)."""
    global _WARNING_SWALLOWER
    if txn is None or not isinstance(txn, Transaction):
        return False
    try:
        if _WARNING_SWALLOWER is None:
            _WARNING_SWALLOWER = _ElevacionEjeWarningSwallower()
        opts = txn.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(_WARNING_SWALLOWER)
        try:
            opts.SetClearAfterRollback(True)
        except Exception:
            pass
        txn.SetFailureHandlingOptions(opts)
        return True
    except Exception:
        return False


def _dispose_revit_scope(obj):
    """Libera ``Transaction`` / ``TransactionGroup`` (IDisposable)."""
    if obj is None:
        return
    try:
        if hasattr(obj, u"Dispose"):
            obj.Dispose()
    except Exception:
        pass


def _txn_has_started(txn):
    if txn is None:
        return False
    try:
        return bool(txn.HasStarted())
    except Exception:
        return False


def _txn_has_ended(txn):
    if txn is None:
        return True
    try:
        return bool(txn.HasEnded())
    except Exception:
        return True


def _safe_rollback_txn(txn):
    if txn is None:
        return
    try:
        if _txn_has_started(txn) and not _txn_has_ended(txn):
            txn.RollBack()
    except Exception:
        pass


def _safe_commit_txn(txn):
    """
    Commit con comprobación de estado.

    Returns:
        True si ``TransactionStatus.Committed``.
    """
    if txn is None:
        return False
    try:
        status = txn.Commit()
    except Exception:
        _safe_rollback_txn(txn)
        return False
    try:
        if status == TransactionStatus.Committed:
            return True
    except Exception:
        pass
    _safe_rollback_txn(txn)
    return False

# Tipos de planta permitidos (str(ViewType); evita conflictos IronPython con el enum).
_VISTAS_PLANTA = frozenset(
    (u"FloorPlan", u"StructuralPlan", u"EngineeringPlan", u"CeilingPlan")
)


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _view_type_name(view):
    try:
        vt = view.ViewType
        s = vt.ToString() if hasattr(vt, u"ToString") else str(vt)
    except Exception:
        return u""
    s = (s or u"").strip()
    if u"." in s:
        s = s.split(u".")[-1]
    return s


def vista_permitida(view):
    """
    La herramienta solo opera en vistas de planta (no plantillas, no cielos).

    Returns:
        (True, u"") o (False, mensaje)
    """
    if view is None:
        return False, u"No hay vista activa."
    try:
        if view.IsTemplate:
            return False, u"No se puede usar sobre una plantilla de vista."
    except Exception:
        pass
    vt = _view_type_name(view)
    if vt in _VISTAS_PLANTA:
        return True, u""
    return False, (
        u"Esta herramienta solo se puede usar en una vista de planta "
        u"(Floor Plan / Structural Plan / Ceiling Plan)."
    )


def _normalize_name(name):
    s = _as_unicode(name).strip()
    if not s:
        return u""
    try:
        import unicodedata

        s = unicodedata.normalize(u"NFC", s)
    except Exception:
        pass
    return s


def find_building_section_type_by_section_filter(document, filter_text):
    """
    Busca ``ViewFamilyType`` Building Section cuyo nombre contenga
    el texto de «Section Filter».

    Prioridad:
      1. Nombre exacto (normalizado)
      2. Igual sin distinguir mayúsculas
      3. El nombre **contiene** el texto (si hay varias, la más corta)

    Returns:
        (ViewFamilyType, None) o (None, mensaje_error)
    """
    target = _normalize_name(filter_text)
    if not target:
        return None, (
            u"«Section Filter» no tiene texto válido para buscar el tipo "
            u"Building Section."
        )

    tipos = listar_tipos_seccion_building(document)
    if not tipos:
        return None, (
            u"No hay tipos de sección (Building Section) en el proyecto."
        )

    for nombre, vft in tipos:
        if _normalize_name(nombre) == target:
            return vft, None

    tl = target.lower()
    for nombre, vft in tipos:
        if _as_unicode(nombre).strip().lower() == tl:
            return vft, None

    contains_matches = []
    for nombre, vft in tipos:
        n = _as_unicode(nombre).strip()
        if not n:
            continue
        if tl in n.lower():
            contains_matches.append((len(n), nombre, vft))
    if contains_matches:
        contains_matches.sort(key=lambda x: (x[0], x[1].lower()))
        return contains_matches[0][2], None

    sample = [n for n, _v in tipos if n][:12]
    msg = (
        u"No se encontró un tipo Building Section cuyo nombre contenga "
        u"«{0}» (valor de Section Filter en la vista activa)."
    ).format(target)
    if sample:
        msg += u" Tipos disponibles: {0}.".format(u", ".join(sample))
    return None, msg


def resolver_tipo_building_section(document, view):
    """
    Lee «Section Filter» de ``view`` y resuelve el ``ViewFamilyType`` Building Section.

    Returns:
        (ViewFamilyType, section_filter_texto, None) o
        (None, section_filter_texto_o_None, mensaje_error)
    """
    if document is None or view is None:
        return None, None, u"Vista o documento no válidos."

    sf_text, err = leer_section_filter_texto(document, view)
    if sf_text is None:
        return None, None, err

    vft, err_vft = find_building_section_type_by_section_filter(document, sf_text)
    if vft is None:
        return None, sf_text, err_vft
    return vft, sf_text, None


def _aplicar_escala(view, scale_ratio):
    if view is None:
        return False
    try:
        view.Scale = int(scale_ratio)
        return True
    except Exception:
        return False


def estampar_armadura_eje(view, nombre_eje):
    """
    Escribe el nombre del eje en el parámetro de instancia ``Armadura_Eje``.

    Returns:
        True si se escribió; False si falta el parámetro o no se pudo setear.
    """
    if view is None:
        return False
    valor = _as_unicode(nombre_eje).strip()
    if not valor:
        return False
    p = None
    try:
        p = view.LookupParameter(_PARAM_ARMADURA_EJE)
    except Exception:
        p = None
    if p is None or p.IsReadOnly:
        return False
    try:
        if p.StorageType == StorageType.String:
            p.Set(valor)
            return True
    except Exception:
        pass
    try:
        p.SetValueString(valor)
        return True
    except Exception:
        return False


def nombre_vista_elevacion(section_filter_text, nombre_eje):
    """
    View.Name = Section Filter + "_" + "ELEVACION EJE" + nombre del eje.

    Ejemplo: ``02_MA_ELEVACION EJE 1``
    """
    sf = _normalize_name(section_filter_text)
    eje = _as_unicode(nombre_eje).strip() or u"Eje"
    if not sf:
        return u"{0} {1}".format(_VISTA_MID, eje).strip()
    return u"{0}_{1} {2}".format(sf, _VISTA_MID, eje).strip()


def _clave_nombre_eje(nombre):
    """Clave comparable de nombre de eje (sin distinguir mayúsculas)."""
    return _normalize_name(nombre).lower()


def leer_armadura_eje(view):
    """Valor del parámetro ``Armadura_Eje`` en la vista, o cadena vacía."""
    if view is None:
        return u""
    p = None
    try:
        p = view.LookupParameter(_PARAM_ARMADURA_EJE)
    except Exception:
        p = None
    if p is None or not p.HasValue:
        return u""
    try:
        if p.StorageType == StorageType.String:
            return _as_unicode(p.AsString()).strip()
    except Exception:
        pass
    try:
        return _as_unicode(p.AsValueString()).strip()
    except Exception:
        return u""


def _element_id_igual(a, b):
    if a is None or b is None:
        return False
    try:
        if a == ElementId.InvalidElementId or b == ElementId.InvalidElementId:
            return False
    except Exception:
        pass
    try:
        return int(a.IntegerValue) == int(b.IntegerValue)
    except Exception:
        try:
            return a == b
        except Exception:
            return False


def _vista_es_elevacion_eje(view):
    """True si el nombre indica una elevación creada por esta herramienta."""
    if view is None:
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass
    try:
        nombre = _as_unicode(view.Name)
    except Exception:
        nombre = u""
    return _VISTA_MID.upper() in nombre.upper()


def _nombre_eje_de_vista_elevacion(view):
    """
    Nombre del eje asociado a una vista Elevación Eje.

    Prioriza ``Armadura_Eje``; si falta, parsea el sufijo tras «ELEVACION EJE».
    """
    arm = leer_armadura_eje(view)
    if arm:
        return arm
    try:
        nombre = _as_unicode(view.Name)
    except Exception:
        return u""
    upper = nombre.upper()
    marker = _VISTA_MID.upper()
    idx = upper.find(marker)
    if idx < 0:
        return u""
    return nombre[idx + len(marker) :].strip()


def nombres_ejes_ya_elevados(document, vft):
    """
    Nombres de ejes que ya tienen una Elevación Eje del ``ViewFamilyType`` dado.

    El filtro es por tipo Building Section (``view.GetTypeId()``), no global:
    el mismo eje puede ofrecerse de nuevo con otro tipo Building Section.

    Returns:
        set de claves ``_clave_nombre_eje`` (minúsculas normalizadas).
    """
    out = set()
    if document is None or vft is None:
        return out
    try:
        vft_id = vft.Id
    except Exception:
        return out

    try:
        collector = FilteredElementCollector(document).OfClass(ViewSection)
    except Exception:
        try:
            collector = FilteredElementCollector(document).OfClass(View)
        except Exception:
            return out

    for view in collector:
        if view is None:
            continue
        try:
            if not isinstance(view, ViewSection):
                if _view_type_name(view) != u"Section":
                    continue
        except Exception:
            continue
        if not _vista_es_elevacion_eje(view):
            continue
        try:
            tid = view.GetTypeId()
        except Exception:
            tid = None
        if not _element_id_igual(tid, vft_id):
            continue
        eje = _nombre_eje_de_vista_elevacion(view)
        clave = _clave_nombre_eje(eje)
        if clave:
            out.add(clave)
    return out


def _pbar_enabled():
    try:
        from pyrevit import forms as _forms  # noqa: F401
    except Exception:
        return False
    return True


class ElevacionEjeProgress(object):
    """ProgressBar pyRevit (acento BIMTools); no-op si no está disponible."""

    def __init__(self, total, title_prefix=None):
        self._total = max(1, int(total or 1))
        self._pb = None
        self._open = False
        self._title_prefix = title_prefix or _TOOL_TITLE

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

                r, g, b = _PROGRESS_ACCENT_RGB
                self._pb.Resources[u"pyRevitAccentBrush"] = SolidColorBrush(
                    Color.FromRgb(r, g, b),
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

    def _title(self, current):
        cur = max(0, int(current))
        if cur < 1:
            return u"{0} — Creando 0/{1}…".format(
                self._title_prefix, int(self._total),
            )
        return u"{0} — Creando {1}/{2}…".format(
            self._title_prefix, cur, int(self._total),
        )

    def update(self, current, label=None):
        if self._pb is None:
            return
        c = max(0, min(int(current), int(self._total)))
        base = self._title(c if c > 0 else 0)
        if label:
            base = u"{0} ({1})".format(base, _as_unicode(label))
        try:
            if hasattr(self._pb, u"update_progress") and c > 0:
                try:
                    self._pb.update_progress(c, max_value=self._total)
                except TypeError:
                    try:
                        self._pb.update_progress(c, max=self._total)
                    except Exception:
                        pass
        except Exception:
            pass
        try:
            self._pb.title = base
        except Exception:
            pass


def ejecutar_crear_elevaciones(uidoc, ejes, scale_ratio, crear_contorno=True):
    """
    Crea una Building Section por cada eje, con contorno e etiquetas.

    Transacciones (Opción B):
      ``TransactionGroup`` + una ``Transaction`` por eje + ``Assimilate()``.
      Undo: una sola entrada ``Arainco: Elevación Eje``.
      Si un eje falla, se continúa con el resto.

    Args:
        uidoc: UIDocument
        ejes: iterable de ``(nombre, Grid)``
        scale_ratio: denominador View.Scale (p. ej. 50 → 1:50)
        crear_contorno: ignorado (el contorno siempre se genera; se mantiene
            por compatibilidad de firma)

    Returns:
        (ok, mensaje, vistas_creadas)
        ok True si se creó al menos una vista.
    """
    if uidoc is None:
        return False, u"No hay documento activo.", []
    doc = uidoc.Document
    if doc is None:
        return False, u"No hay documento activo.", []

    active_view = uidoc.ActiveView
    ok_vista, msg_vista = vista_permitida(active_view)
    if not ok_vista:
        return False, msg_vista, []

    pares = []
    for item in ejes or []:
        try:
            nombre, grid = item
        except Exception:
            continue
        if grid is None:
            continue
        pares.append((_as_unicode(nombre).strip() or u"Eje", grid))

    if not pares:
        return False, u"Marca al menos un eje para continuar.", []

    try:
        scale = int(scale_ratio)
    except Exception:
        scale = 50
    if scale <= 0:
        scale = 50

    vft, sf_text, err_vft = resolver_tipo_building_section(doc, active_view)
    if vft is None:
        return False, err_vft or u"Tipo Building Section no disponible.", []

    try:
        tipo_nombre = _as_unicode(vft.Name).strip()
    except Exception:
        tipo_nombre = sf_text or u"Building Section"

    ya_elevados = nombres_ejes_ya_elevados(doc, vft)
    omitidos = []
    if ya_elevados:
        pares_filtrados = []
        for nombre, grid in pares:
            if _clave_nombre_eje(nombre) in ya_elevados:
                omitidos.append(nombre)
            else:
                pares_filtrados.append((nombre, grid))
        pares = pares_filtrados

    if not pares:
        if omitidos:
            return False, (
                u"Los ejes seleccionados ya tienen elevación con el tipo "
                u"Building Section «{0}»."
            ).format(tipo_nombre), []
        return False, u"Marca al menos un eje para continuar.", []

    try:
        from contorno_hormigon_elevacion_eje import (
            generar_contorno_model_lines,
            resolver_medium_lines_style_id,
        )
        from contorno_hormigon_eje import recopilar_nombres_grupos
    except Exception as ex_imp:
        return False, (
            u"No se pudo cargar el módulo de contorno: {0}"
        ).format(_as_unicode(ex_imp)), []

    try:
        from elevacion_eje_wall_tags import (
            etiquetar_muros_concrete_paralelos,
        )
        from armado_muros_wall_tags_rebase import (
            resolve_wall_espesor_tag_symbol,
        )
        from elevacion_eje_beam_tags import (
            etiquetar_vigas_concrete_paralelas,
            resolve_beam_tag_symbol,
        )
        from elevacion_eje_collect import (
            MaterialConcreteCache,
            indexar_tags_por_host,
            recoger_concrete_en_vista,
        )
    except Exception as ex_tag:
        return False, (
            u"No se pudo cargar el módulo de etiquetas: {0}"
        ).format(_as_unicode(ex_tag)), []

    # Lecturas fuera de transacción (documento no modificable).
    bbox = bbox_modelo_documento(doc)
    if bbox is None:
        return False, (
            u"No se pudo obtener la extensión del modelo "
            u"para dimensionar las secciones."
        ), []
    nombres_vistas = recopilar_nombres_vistas(doc)
    material_cache = MaterialConcreteCache()
    style_id = None
    nombres_grupos = set()
    try:
        style_id = resolver_medium_lines_style_id(doc)
    except Exception:
        style_id = None
    try:
        nombres_grupos = recopilar_nombres_grupos(doc)
    except Exception:
        nombres_grupos = set()
    wall_sym, err_sym = resolve_wall_espesor_tag_symbol(doc)
    tags_symbol_err = None
    if wall_sym is None:
        tags_symbol_err = err_sym or u"Tipo de etiqueta muro no encontrado."
    beam_sym, err_beam = resolve_beam_tag_symbol(doc)
    beam_tags_symbol_err = None
    if beam_sym is None:
        beam_tags_symbol_err = err_beam or u"Tipo de etiqueta viga no encontrado."

    creadas = []
    fallos = []
    sin_param = []
    contorno_ok = 0
    contorno_fallos = []
    tags_ok = 0
    tags_fail = 0
    tags_skip = 0
    tags_fallos = []
    beam_tags_ok = 0
    beam_tags_fail = 0
    beam_tags_skip = 0
    total_steps = len(pares) * 3
    step = 0

    tg = TransactionGroup(doc, _TX_NAME)
    try:
        tg.Start()
        with ElevacionEjeProgress(total_steps, title_prefix=_TOOL_TITLE) as progress:
            progress.update(0)
            for i_eje, (nombre, grid) in enumerate(pares):
                step_end = (i_eje + 1) * 3
                step += 1
                progress.update(step, label=nombre)
                label = nombre_vista_elevacion(sf_text, nombre)
                t = Transaction(doc, _TX_NAME)
                _attach_warning_swallower(t)
                try:
                    t.Start()
                    # Activate exige TX abierta (resolución previa es solo lectura).
                    if wall_sym is not None:
                        try:
                            if not wall_sym.IsActive:
                                wall_sym.Activate()
                        except Exception:
                            pass
                    if beam_sym is not None:
                        try:
                            if not beam_sym.IsActive:
                                beam_sym.Activate()
                        except Exception:
                            pass
                    vs, err = crear_seccion_alzado(
                        doc,
                        grid,
                        vft.Id,
                        nombre_vista=label,
                        bbox_modelo=bbox,
                        nombres_vistas=nombres_vistas,
                        desactivar_crop=False,
                    )
                    if vs is None:
                        fallos.append(
                            u"{0}: {1}".format(nombre, err or u"error")
                        )
                        _safe_rollback_txn(t)
                        step = step_end
                        progress.update(step, label=nombre)
                        continue

                    _aplicar_escala(vs, scale)
                    if not estampar_armadura_eje(vs, nombre):
                        sin_param.append(nombre)

                    try:
                        doc.Regenerate()
                    except Exception:
                        pass

                    packed = recoger_concrete_en_vista(doc, vs, material_cache)
                    hosts = list(packed.get(u"hosts") or [])
                    muros = list(packed.get(u"muros") or [])
                    vigas = list(packed.get(u"vigas") or [])

                    step += 1
                    progress.update(step, label=u"{0} · contorno".format(nombre))
                    try:
                        ok_c, info_c = generar_contorno_model_lines(
                            doc,
                            vs,
                            grid,
                            group_name=label,
                            style_id=style_id,
                            nombres_grupos=nombres_grupos,
                            regenerate=False,
                            elementos=hosts,
                        )
                    except Exception as ex_c:
                        ok_c, info_c = False, _as_unicode(ex_c)
                    if ok_c:
                        contorno_ok += 1
                        try:
                            sid = info_c.get(u"style_id")
                        except Exception:
                            sid = None
                        if sid is not None:
                            style_id = sid
                    else:
                        contorno_fallos.append(
                            u"{0}: {1}".format(nombre, _as_unicode(info_c))
                        )

                    step += 1
                    progress.update(step, label=u"{0} · tags".format(nombre))
                    tagged_hosts = indexar_tags_por_host(doc, vs)

                    if wall_sym is not None:
                        try:
                            tres = etiquetar_muros_concrete_paralelos(
                                doc,
                                vs,
                                symbol=wall_sym,
                                muros=muros,
                                tagged_hosts=tagged_hosts,
                            )
                        except Exception as ex_t:
                            tags_fail += 1
                            tags_fallos.append(
                                u"{0} muro: {1}".format(
                                    nombre, _as_unicode(ex_t),
                                )
                            )
                            tres = None
                        if tres is not None:
                            tags_ok += int(tres.get(u"n_ok", 0) or 0)
                            tags_skip += int(tres.get(u"n_skip", 0) or 0)
                            tags_fail += int(tres.get(u"n_fail", 0) or 0)
                    else:
                        tags_fail += 1

                    if beam_sym is not None:
                        try:
                            bres = etiquetar_vigas_concrete_paralelas(
                                doc,
                                vs,
                                symbol=beam_sym,
                                vigas=vigas,
                                tagged_hosts=tagged_hosts,
                            )
                        except Exception as ex_b:
                            beam_tags_fail += 1
                            tags_fallos.append(
                                u"{0} viga: {1}".format(
                                    nombre, _as_unicode(ex_b),
                                )
                            )
                            bres = None
                        if bres is not None:
                            beam_tags_ok += int(bres.get(u"n_ok", 0) or 0)
                            beam_tags_skip += int(bres.get(u"n_skip", 0) or 0)
                            beam_tags_fail += int(bres.get(u"n_fail", 0) or 0)
                    else:
                        beam_tags_fail += 1

                    try:
                        vs.CropBoxActive = False
                    except Exception:
                        pass

                    if _safe_commit_txn(t):
                        creadas.append(vs)
                    else:
                        fallos.append(
                            u"{0}: commit de transacción falló".format(nombre)
                        )
                except Exception as ex_eje:
                    _safe_rollback_txn(t)
                    fallos.append(
                        u"{0}: {1}".format(nombre, _as_unicode(ex_eje))
                    )
                    step = step_end
                    progress.update(step, label=nombre)
                finally:
                    _dispose_revit_scope(t)

            progress.update(total_steps)

        if creadas:
            try:
                tg.Assimilate()
            except Exception as ex_as:
                try:
                    if not _txn_has_ended(tg):
                        tg.RollBack()
                except Exception:
                    pass
                return False, (
                    u"No se pudo asimilar el grupo de transacciones: {0}"
                ).format(_as_unicode(ex_as)), []
        else:
            try:
                if not _txn_has_ended(tg):
                    tg.RollBack()
            except Exception:
                pass
            detalle = u"; ".join(fallos[:5])
            if len(fallos) > 5:
                detalle = detalle + u"…"
            return False, u"No se pudo crear ninguna elevación. {0}".format(
                detalle,
            ), []
    except Exception as ex:
        try:
            if not _txn_has_ended(tg):
                tg.RollBack()
        except Exception:
            pass
        return False, _as_unicode(ex), []
    finally:
        _dispose_revit_scope(tg)

    n = len(creadas)
    msg = u"{0} elevación(es) creadas · Tipo «{1}» · Escala 1:{2}.".format(
        n, tipo_nombre, scale,
    )
    if not sin_param:
        msg = msg + u" Armadura_Eje estampado."
    else:
        msg = msg + (
            u" Sin «Armadura_Eje» en {0} vista(s) "
            u"(parámetro ausente o no escribible)."
        ).format(len(sin_param))
    msg = msg + u" Contorno: {0}/{1}.".format(contorno_ok, n)
    if contorno_fallos:
        preview = u"; ".join(contorno_fallos[:3])
        if len(contorno_fallos) > 3:
            preview = preview + u"…"
        msg = msg + u" Fallos contorno: {0}.".format(preview)
    if tags_symbol_err:
        msg = msg + u" Etiquetas muro: {0}.".format(tags_symbol_err)
    else:
        msg = msg + u" Etiquetas muro: {0} ok".format(tags_ok)
        if tags_skip:
            msg = msg + u", {0} ya etiquetados".format(tags_skip)
        if tags_fail:
            msg = msg + u", {0} fallos".format(tags_fail)
        msg = msg + u"."
    if beam_tags_symbol_err:
        msg = msg + u" Etiquetas viga: {0}.".format(beam_tags_symbol_err)
    else:
        msg = msg + u" Etiquetas viga: {0} ok".format(beam_tags_ok)
        if beam_tags_skip:
            msg = msg + u", {0} ya etiquetadas".format(beam_tags_skip)
        if beam_tags_fail:
            msg = msg + u", {0} fallos".format(beam_tags_fail)
        msg = msg + u"."
    if tags_fallos:
        preview = u"; ".join(tags_fallos[:3])
        if len(tags_fallos) > 3:
            preview = preview + u"…"
        msg = msg + u" Detalle tags: {0}.".format(preview)
    if omitidos:
        msg = msg + u" Omitidos (ya elevados en este tipo): {0}.".format(
            len(omitidos),
        )
    if fallos:
        msg = msg + u" Fallos creación: {0}.".format(len(fallos))
    return True, msg, creadas
