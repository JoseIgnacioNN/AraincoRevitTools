# -*- coding: utf-8 -*-
# Ejecutar en RPS: File > Run script (no pegar línea a línea).
"""
Longitudinales como ``Structural Rebar``: rejilla regular A × B respecto recubrimiento.

Lado A = siempre el lado más corto del rectángulo de sección en planta; lado B = el más largo.
El reparto en planta usa los **ejes locales del símbolo** (`GetTransform`: lado corto/largo según
extensión real) cuando están disponibles; si no, el marco **X/Y del proyecto** como hasta ahora.
La **sección en planta** sale de parámetros de tipo/instancia (``b``/``h``, etc.) o de **geometría sólida**
(extensión orientada ⟂ eje del pilar, o **prioritaria**: ``FamilyInstance.GetTransform`` + mallas
``Face.Triangulate`` en ejes locales del símbolo); **no** se usa ``Element.GetBoundingBox`` (BBAA elemento) para ese análisis.

- Checkbox desmarcado: sólo el contorno exterior de esa rejilla
  (= 2·A + 2·B − 4 hilos — estribo verde típico).
- Checkbox marcado: contorno exterior + contorno interior alineado, sin relleno intermedio
  (estribo rojo típico): puntos índices 1…A−2 y 1…B−2 que están en la arista interior
  (perímetro de la rejilla interior (A−2)×(B−2)). Ej. 4×6 → 16 + 8 = 24 sin “tercer nivel” visual.

Con varios pilares seleccionados, **antes de la fusión** se agrupan por **sección en planta** (lados en mm deducidos como arriba; el par corto × largo hace estable 75 × 100 con 100 × 75). Por cada tamaño distinto aparece **un cuadro** para rejilla lado corto × largo y contorno interior. Luego los (X,Y) en proyecto siguen agrupándose como hasta ahora (**una barra** por eje común/tramo).

En cada **barra vertical fusionada**, **Comentarios** llevan **A**/ **B**/ **IA**/ **IB** según un **orden horario único**

por el perímetro de la rejilla: **alternancia A–B–A–B…** desde la esquina ``(índice corto = 0, índice largo = último)`` (equivalente al esquema de planta habitual).

Las **IA**/**IB** siguen la misma regla sobre el perímetro interior. Los marcadores XY usan «A-type» (**A**/**IA**)→trazo ±X proyecto y «B-type» (**B**/**IB**)→±Y.

Opcional: **columnas de referencia** definen cotas de troceo. Para cada una, el plano de **tipo A**
``_plane_from_column_location_curve`` pasa por el **punto inicial** del ``LocationCurve`` (**startpoint**, extremo ``0``) y **ajusta la cota Z** al borde inferior del sólido (vértices; sin BB elemento), o segundo plano por **borde superior** si otra referencia repetía base.
**Hilos A y IA** (anillo exterior e interior): cortan sólo contra esos planos (**sin desplazar** con empalme).
**Hilos B e IB**: la **misma** construcción de plano pero **trasladada** ``L(Ø)`` **de empalme/traslape tabular** a lo largo del eje de esa columna (start→end): el **punto de troceo físico queda corrido respecto del de A**.
**Los tramos modelo** siguen fusionados + fundación −Z antes de cortar (no se antepone el alargamiento +Z de empotramiento al intervalo de troceo); **tras** el troceo se aplica la política de +L en último/primer tramo e intermedios según colisión y esquema.

Tras la fusión, el trazo se alarga ``+Z`` según ``traslape_mm_from_nominal_diameter_mm`` (traslape/empotramiento

por diámetro en ``bimtools_rebar_hook_lengths``); Ø nominal interino configurable (actualmente **12 mm**).

Tras ese estiramiento, un **prisma vertical** igual al segmento modelo de **empotramiento** (entre ``Z`` techo fusionado

y extremo tras ``+ L(Ø)``) se contrasta contra **los sólidos** de las ``Structural Columns`` del proyecto

(**booleano** ``Solid``–prisma). No se usa **BBAA elemento** como sustituto de ese choque ni para dimensiones/sección.

En paralelo se evalúa un prisma **−Z** desde ``Z`` suelo fusionado (``base_xyz.Z``) y **misma ``L(Ø)``**,

**sólo en hilos que no se estiren por fundación unida**: si ese hilo lleva −Z geométrico por fundación, no participa esta prueba.


Entre pilares que comparten el mismo XY fusionado, el que tiene **menor** cota **Z máxima de la geometría sólida**

se trata como pieza típica del **primer tramo inferior** del apilamiento en ese eje: **no cuenta** esa instancia

(respaldo ante solapes geométricos residuales con el propio primer pilar).


Las aristas fuera de obra de otros pilares (p. ej. pilar más estrecho encima sin coincidencia de rejilla) siguen evaluándose por geometría sólida.

Si en un hilo **no** se conserva ese ``+ L(Ø)``, el trazo vuelve al fusionado y se acorta hacia −Z

``25 mm + (Ø nominal)/2`` desde el máximo fusionado.

En ese caso también se crea una **pata** (barra horizontal en planta) desde ese extremo superior

hacia el centro en planta (según ``Location`` / geometría del pilar) de la **contribuyente más cercana** en XY;

el largo de la pata sale de ``hook_length_mm_from_nominal_diameter_mm`` / tabla BIMTools pata−Ø.


Si alguna columna **seleccionada** tiene **fundación estructural** unida por *Unir geometría*

con la **base** del pilar (cara superior de la fundación alineada con la inferior del sólido del pilar),

cada barra fusionada cuyos contribuyentes incluyan esa columna **alarga hacia −Z** desde el punto de arranque

``altura geométrica fundación − 50 mm`` (máximo entre pilares contribuyentes). El extremo **superior** del trazo fusionado/empotr. no cambia.

En ese mismo arranque **inferior** (``startpoint`` de la barra vertical fusionada) puede ir una **pata** en el **mismo ``Rebar``** (tramo horizontal coplanar con el longitudinal) hacia el centro en planta

de la contribuyente más cercana en XY (misma longitud tabulada hook/pata Ø que las patas por reverso +Z).


Compatible con RevitPythonShell (RPS) y pyRevit. Se admite selección múltiple (PickObjects).

"""

from __future__ import print_function

import os
import sys
import math
from collections import defaultdict, Counter
import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from Autodesk.Revit.DB import (  # noqa: E402
    BooleanOperationsType,
    BooleanOperationsUtils,
    BuiltInCategory,
    BuiltInParameter,
    Category,
    Curve,
    CurveElement,
    CurveLoop,
    ElementId,
    FilteredElementCollector,
    GeometryCreationUtilities,
    GeometryInstance,
    GraphicsStyle,
    JoinGeometryUtils,
    Line,
    LocationCurve,
    LocationPoint,
    Options,
    Plane,
    SketchPlane,
    Solid,
    StorageType,
    Transaction,
    TransactionGroup,
    Transform,
    UnitTypeId,
    UnitUtils,
    ViewDetailLevel,
    XYZ,
)
from Autodesk.Revit.DB.Structure import (  # noqa: E402
    Rebar,
    RebarBarType,
    RebarHookOrientation,
    RebarShape,
    RebarStyle,
)
from Autodesk.Revit.UI import TaskDialog  # noqa: E402
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType  # noqa: E402
from Autodesk.Revit.Exceptions import OperationCanceledException  # noqa: E402
import System  # noqa: E402

from System.Windows.Forms import (  # noqa: E402
    Button,
    CheckBox,
    DialogResult,
    Form,
    FormBorderStyle,
    FormStartPosition,
    FormWindowState,
    HorizontalAlignment,
    Label,
    NumericUpDown,
    Padding,
)
from System.Drawing import Point, Size, SystemColors, SystemFonts  # noqa: E402
from System import AppDomain  # noqa: E402
from System import Convert  # noqa: E402


def _prepend_scripts_search_path():
    """Inserta rutas candidatas al frente de ``sys.path``.

    Se recorre ``dirs`` al revés al insertar, para que la **primera** entrada de
    ``dirs`` quede finalmente en ``sys.path[0]``: directorio de este archivo (carpeta
    ``scripts/`` del pushbutton o de la extensión), luego cwd, luego heurística
    BIMTools bajo ``%USERPROFILE%``. Así las copias empaquetadas en el ``.pushbutton``
    tienen prioridad sobre la extensión global del usuario.
    """
    dirs = []
    try:
        d0 = os.path.dirname(os.path.abspath(__file__))
        if d0:
            dirs.append(d0)
    except NameError:
        pass
    try:
        cwd = os.getcwd()
        if cwd:
            dirs.append(cwd)
    except Exception:
        pass
    try:
        home = os.path.expanduser("~")
        guess = os.path.join(
            home,
            "CustomRevitExtensions",
            "BIMTools.extension",
            "scripts",
        )
        if os.path.isdir(guess):
            dirs.append(guess)
    except Exception:
        pass
    seen_nd = set()
    for d in reversed(dirs):
        if not d:
            continue
        try:
            nd = os.path.normpath(d)
        except Exception:
            continue
        if nd in seen_nd:
            continue
        seen_nd.add(nd)
        if os.path.isdir(nd) and nd not in sys.path:
            sys.path.insert(0, nd)


_prepend_scripts_search_path()

from bimtools_rebar_3d_visibility import apply_rebar_unobscured_in_3d_views  # noqa: E402

try:
    from bimtools_rebar_hook_lengths import traslape_mm_from_nominal_diameter_mm
except Exception:
    traslape_mm_from_nominal_diameter_mm = None

try:
    from bimtools_rebar_hook_lengths import hook_length_mm_from_nominal_diameter_mm
except Exception:
    hook_length_mm_from_nominal_diameter_mm = None

try:
    from column_reinforcement.geometry.schemes import (
        is_a_split_scheme,
        is_b_split_scheme,
        is_lap_extension_scheme,
    )
except Exception:
    def is_a_split_scheme(tag):
        return str(tag or "").strip() in ("A", "IA")

    def is_b_split_scheme(tag):
        return str(tag or "").strip() in ("B", "IB")

    def is_lap_extension_scheme(tag):
        return str(tag or "").strip() in ("A", "IA", "B", "IB")


# Copia mínima de tablas BIMTools si no se importó ``bimtools_rebar_hook_lengths``.
_COL_EMB_BASE = (
    (8, 570),
    (10, 710),
    (12, 860),
    (16, 1140),
    (18, 1290),
    (22, 1960),
    (25, 2230),
    (28, 2500),
    (32, 2850),
    (36, 3210),
)
_COL_EMB_G25 = (
    (8, 560),
    (10, 690),
    (12, 840),
    (16, 1110),
    (18, 1240),
    (22, 1890),
    (25, 2150),
    (28, 2410),
    (32, 2750),
    (36, 3090),
)
_COL_EMB_G35 = (
    (8, 470),
    (10, 590),
    (12, 710),
    (16, 940),
    (18, 1060),
    (22, 1600),
    (25, 1810),
    (28, 2030),
    (32, 2320),
    (36, 2620),
)
_COL_EMB_G45 = (
    (8, 420),
    (10, 520),
    (12, 630),
    (16, 820),
    (18, 930),
    (22, 1410),
    (25, 1600),
    (28, 1800),
    (32, 2050),
    (36, 2310),
)

# Respaldos BIMTools para largo de **pata** (mm por Ø nominal) si no hay import hook.
_HOOK_PATA_MM_TBL = (
    (8, 160),
    (10, 200),
    (12, 240),
    (16, 320),
    (18, 360),
    (22, 440),
    (25, 500),
    (28, 570),
    (32, 650),
    (36, 720),
)


def _column_emb_grade_norm(concrete_grade):
    if concrete_grade is None:
        return None
    try:
        s = str(concrete_grade).strip().upper()
    except Exception:
        return None
    if s in ("G25", "G35", "G45"):
        return s
    return None


def _column_emb_tbl(grade_norm):
    if grade_norm == "G25":
        return _COL_EMB_G25
    if grade_norm == "G35":
        return _COL_EMB_G35
    if grade_norm == "G45":
        return _COL_EMB_G45
    return _COL_EMB_BASE


def _column_emb_interpolate_mm(d_int, tbl):
    d = float(d_int)
    d0, L0 = tbl[0]
    if d <= float(d0):
        return float(L0)
    d_last, L_last = tbl[-1]
    if d >= float(d_last):
        return float(L_last)
    for i in range(len(tbl) - 1):
        da, La = tbl[i]
        db, Lb = tbl[i + 1]
        da, db = float(da), float(db)
        if da <= d <= db:
            if abs(db - da) < 1e-9:
                return float(La)
            t = (d - da) / (db - da)
            return float(La) + t * (float(Lb) - float(La))
    return float(L_last)


def _traslape_embed_mm_local_fallback(diameter_mm, concrete_grade=None):
    """Misma lógica que ``traslape_mm_from_nominal_diameter_mm`` (tablas arriba)."""
    try:
        d = float(diameter_mm)
    except Exception:
        return None
    if d <= 0.0 or d != d:
        return None
    di = int(round(d))
    g = _column_emb_grade_norm(concrete_grade)
    tbl = _column_emb_tbl(g)
    return _column_emb_interpolate_mm(di, tbl)


def _resolved_traslape_embed_mm(diameter_mm, concrete_grade):
    if traslape_mm_from_nominal_diameter_mm is not None:
        try:
            val = traslape_mm_from_nominal_diameter_mm(diameter_mm, concrete_grade)
            if val is not None:
                return float(val)
        except Exception:
            pass
    val2 = _traslape_embed_mm_local_fallback(diameter_mm, concrete_grade)
    return float(val2) if val2 is not None else None


def _resolved_pata_hook_mm_for_revert(diameter_mm, concrete_grade):
    """Largo de pata (mm) después de reverte empotramiento Ø: tabla BIMTools gancho/hook."""
    if hook_length_mm_from_nominal_diameter_mm is not None:
        try:
            v = hook_length_mm_from_nominal_diameter_mm(
                diameter_mm, concrete_grade
            )
            if v is not None:
                return max(float(v), 0.0)
        except Exception:
            pass
    try:
        di = int(round(float(diameter_mm)))
    except Exception:
        di = 12
    return float(_column_emb_interpolate_mm(di, _HOOK_PATA_MM_TBL))


def _p(x, y):
    """WinForms Point requiere int (Int32); evita floats por división /. """
    return Point(int(round(float(x))), int(round(float(y))))


def _s(w, h):
    """WinForms Size requiere anchos/alto enteros."""
    return Size(int(round(float(w))), int(round(float(h))))


# ── Document / UIDocument ──────────────────────────────────────────────
try:
    doc = __revit__.ActiveUIDocument.Document  # noqa: F821
    uidoc = __revit__.ActiveUIDocument  # noqa: F821
except NameError:
    try:
        doc
    except NameError:
        doc = None
    try:
        uidoc
    except NameError:
        uidoc = None


class StructuralColumnFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            return (
                elem.Category
                and elem.Category.Id.IntegerValue
                == int(BuiltInCategory.OST_StructuralColumns)
            )
        except Exception:
            return False

    def AllowReference(self, reference, point):
        return False


def pick_structural_columns(uidoc):
    prompt = (
        "Selecciona una o más columnas estructurales. Completa la selección "
        "(Finalizar)."
    )
    ref_col = uidoc.Selection.PickObjects(
        ObjectType.Element,
        StructuralColumnFilter(),
        prompt,
    )
    if ref_col is None:
        return []
    return list(ref_col)


def pick_structural_columns_optional(uidoc, prompt):
    """
    Misma selección que ``pick_structural_columns`` pero **Esc** / cancelación
    devuelve lista vacía (no error).
    """
    try:
        ref_col = uidoc.Selection.PickObjects(
            ObjectType.Element,
            StructuralColumnFilter(),
            str(prompt),
        )
    except OperationCanceledException:
        return []
    if ref_col is None:
        return []
    return list(ref_col)


def build_column_elements_ordered(doc, refs):
    seen_iv = set()
    out = []

    for ref in refs:
        if ref is None:
            continue

        rid = ref.ElementId
        iv = int(rid.IntegerValue)
        if iv in seen_iv:
            continue

        elem = doc.GetElement(rid)
        if elem is None:
            continue

        seen_iv.add(iv)
        out.append(elem)

    if not out:
        raise Exception(
            "No quedaron columnas estructurales tras deduplicar la selección."
        )

    return out


def _column_base_z_ft_for_sort(col):
    u"""Cota Z del extremo inferior del pilar (pies internos), para ordenar de más bajo a más alto."""
    if col is None:
        return 0.0
    try:
        loc = col.Location
        crv = getattr(loc, "Curve", None)
        if crv is not None:
            p0 = crv.GetEndPoint(0)
            p1 = crv.GetEndPoint(1)
            return float(min(p0.Z, p1.Z))
        pt = getattr(loc, "Point", None)
        if pt is not None:
            return float(pt.Z)
    except Exception:
        pass
    try:
        bb = col.get_BoundingBox(None)
        if bb is not None:
            return float(bb.Min.Z)
    except Exception:
        pass
    return 0.0


class BarInputForm(Form):
    def __init__(self, section_heading=None):

        """

        ``section_heading``: texto opcional sobre la primera fila—p. ej. sección BB en mm + recuento de pilares.


        """

        Form.__init__(self)

        pad = Padding(24, 20, 24, 20)
        self.Padding = pad
        self.Font = SystemFonts.MessageBoxFont

        self.Text = (
            u"Configuración de rejilla por sección — layout columna"
            if section_heading
            else u"Configuración de barras (columna)"
        )
        header_drop = 0
        left_x = pad.Left
        label_w = 252
        hdr_w = label_w + 100 + 12

        if section_heading:

            lh = Label()

            lh.Text = str(section_heading)

            lh.Location = _p(left_x, pad.Top + 6)

            lh.Size = _s(hdr_w, 44)

            self.Controls.Add(lh)

            header_drop = 50

        self.ClientSize = _s(472, int(448 + header_drop))

        self.StartPosition = FormStartPosition.CenterScreen

        self.FormBorderStyle = FormBorderStyle.FixedDialog

        self.MaximizeBox = False

        self.MinimizeBox = False

        self.TopMost = True

        self.bars_a = None

        self.bars_b = None

        self.include_inner_outline = True

        nud_x = left_x + label_w + 12

        nud_w = 100

        nud_h = 28

        row_gap = 40

        y0 = pad.Top + 20 + header_drop

        la = Label()
        la.Text = (
            "Cantidad lado A (lado corto):\n"
            "Rejilla según la dimensión menor de la sección en planta"
        )
        la.Location = _p(left_x, y0 - 6)
        la.Size = _s(label_w, 36)
        self.Controls.Add(la)

        self.num_a = NumericUpDown()
        self.num_a.Location = _p(nud_x, y0 + 8)
        self.num_a.Size = _s(nud_w, nud_h)
        self.num_a.DecimalPlaces = 0
        self.num_a.Minimum = 2
        self.num_a.Maximum = 20
        self.num_a.Value = 4
        self.num_a.TabIndex = 0
        self.num_a.TextAlign = HorizontalAlignment.Center
        self.Controls.Add(self.num_a)

        y1 = y0 + row_gap + 22

        lb = Label()
        lb.Text = (
            "Cantidad lado B (lado largo):\n"
            "Rejilla según la dimensión mayor de la sección en planta"
        )
        lb.Location = _p(left_x, y1 - 6)
        lb.Size = _s(label_w, 36)
        self.Controls.Add(lb)

        self.num_b = NumericUpDown()
        self.num_b.Location = _p(nud_x, y1 + 8)
        self.num_b.Size = _s(nud_w, nud_h)
        self.num_b.DecimalPlaces = 0
        self.num_b.Minimum = 2
        self.num_b.Maximum = 20
        self.num_b.Value = 6
        self.num_b.TabIndex = 1
        self.num_b.TextAlign = HorizontalAlignment.Center
        self.Controls.Add(self.num_b)

        y2 = y1 + row_gap + 34

        self.second_ring_checkbox = CheckBox()
        self.second_ring_checkbox.Text = (
            "Añadir contorno interior (estribo rojo, sin más anillos"
            " intermedios)"
        )
        self.second_ring_checkbox.AutoSize = True
        self.second_ring_checkbox.Location = _p(left_x, y2)
        self.second_ring_checkbox.TabIndex = 2
        self.second_ring_checkbox.Checked = True
        self.Controls.Add(self.second_ring_checkbox)

        chk_help = Label()
        chk_help.Text = (
            "Marcado = perímetro exterior + perímetro rejilla interior (A−2)×(B−2)\n"
            "alineada (p.ej. 4×6 → 16+8 hilos; sin barras entre ambos).\n"
            "Desmarcado = sólo perímetro exterior (p.ej. 16 hilos)."
        )
        chk_help.ForeColor = SystemColors.GrayText
        chk_help.Location = _p(left_x + 22, y2 + 38)
        chk_help.Size = _s(label_w + nud_w + 12 - 12, 58)
        self.Controls.Add(chk_help)

        y3 = y2 + 118

        tip = Label()
        tip.Text = (
            "BBox en planta: el script asigna A al lado más corto (min ancho,fondo)\n"
            "y B al más largo, y proyecta ese reparto sobre los ejes X/Y del proyecto."
        )
        tip.Location = _p(left_x, y3)
        tip.ForeColor = SystemColors.GrayText
        tip.Size = _s(label_w + nud_w + 12, 40)
        self.Controls.Add(tip)

        bw = 108
        bh = 32
        bottom_y = self.ClientSize.Height - pad.Bottom - bh - 4
        ok_button = Button()
        ok_button.Text = "Generar"
        btn_offs = (
            label_w + nud_w + 12 - (2 * bw) - 14
        ) // 2  # debe ser entero para Point (Int32)
        ok_button.Location = _p(left_x + btn_offs, bottom_y)
        ok_button.Size = _s(bw, bh)
        ok_button.TabIndex = 3
        ok_button.Click += self.ok_clicked
        self.Controls.Add(ok_button)

        cancel_button = Button()
        cancel_button.Text = "Cancelar"
        cancel_button.Location = _p(ok_button.Right + 14, bottom_y)
        cancel_button.Size = _s(bw, bh)
        cancel_button.TabIndex = 4
        cancel_button.Click += self.cancel_clicked
        self.Controls.Add(cancel_button)

        self.AcceptButton = ok_button
        self.CancelButton = cancel_button

    def ok_clicked(self, sender, args):
        self.bars_a = Convert.ToInt32(self.num_a.Value)
        self.bars_b = Convert.ToInt32(self.num_b.Value)
        self.include_inner_outline = self.second_ring_checkbox.Checked
        self.DialogResult = DialogResult.OK
        self.Close()

    def cancel_clicked(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()


def get_column_dimensions(column):
    """
    Dimensiones y centro base sin ``Element.GetBoundingBox``.

    Preferencia: **ejes locales del símbolo** (``GetTransform``) y extensión desde **mallas
    trianguladas** de caras (geometría real). Si no aplica o falla, se usa el respaldo
    (parámetros b/h y aristas de sólidos / curva de ubicación).
    """
    doc = getattr(column, "Document", None)
    tol_curve = 1e-9
    try:
        if doc is not None:
            tol_curve = float(doc.Application.ShortCurveTolerance)
    except Exception:
        pass
    tol_curve = max(tol_curve, 1e-12)

    mesh_dims = _try_column_dimensions_transform_mesh_ft(column, tol_curve)
    if mesh_dims is not None:
        return mesh_dims  # 6-tuple incl. ejes rejilla locales

    p0, p1, axis_u, curve_len = _column_curve_endpoints_axis_ft(column, tol_curve)

    opts = _geometry_options_structure_solids()
    rng = _solid_aggregate_vertex_ranges_ft(column, opts)
    if rng is None:
        raise Exception(
            u"No se pudo obtener geometría sólida de la columna "
            u"(¿sin sólidos o elemento mal configurado?)."
        )

    min_x, max_x, min_y, max_y, min_z, max_z = rng

    w_par, d_par = _lookup_column_section_width_depth_ft(column)

    verts = []
    for solid in _iter_solids_revit_element(column, opts):
        for pt in _iter_vertices_from_solid(solid):
            verts.append(pt)

    origin_for_section = p0
    if origin_for_section is None:
        origin_for_section = XYZ(
            0.5 * (float(min_x) + float(max_x)),
            0.5 * (float(min_y) + float(max_y)),
            0.5 * (float(min_z) + float(max_z)),
        )

    if axis_u is None:
        axis_u = XYZ.BasisZ
    try:
        axis_u = axis_u.Normalize()
    except Exception:
        axis_u = XYZ.BasisZ

    if (
        w_par is not None
        and d_par is not None
        and float(w_par) > 1e-12
        and float(d_par) > 1e-12
    ):
        width, depth = float(w_par), float(d_par)
    else:
        if not verts:
            raise Exception(
                u"No se pudieron medir lados de sección: sin vértices en sólidos "
                u"y sin parámetros de ancho/profundidad (b/h, Width/Depth, etc.)."
            )
        width, depth = _oriented_section_width_depth_ft(
            origin_for_section,
            axis_u,
            verts,
        )
        if width <= 1e-12 or depth <= 1e-12:
            raise Exception(
                u"No se pudo determinar la sección en planta desde la geometría sólida."
            )

    if curve_len > tol_curve:
        height = float(curve_len)
    else:
        height = None
        if verts:
            o = origin_for_section
            dots = []
            for p in verts:
                try:
                    dots.append(float((p - o).DotProduct(axis_u)))
                except Exception:
                    continue
            if dots:
                cand = float(max(dots) - min(dots))
                if cand > tol_curve:
                    height = cand
        if height is None or height <= tol_curve:
            height = abs(float(max_z) - float(min_z))
        if height <= tol_curve:
            raise Exception(u"No se pudo determinar la altura del pilar.")

    if p0 is not None and p1 is not None:
        mid = p0 + 0.5 * (p1 - p0)
        cx, cy = float(mid.X), float(mid.Y)
    else:
        cx = 0.5 * (float(min_x) + float(max_x))
        cy = 0.5 * (float(min_y) + float(max_y))

    center = XYZ(cx, cy, float(min_z))
    vs, vl = _plan_axes_from_family_transform_ft(column, width, depth)
    return width, depth, height, center, vs, vl


def _column_reference_level_name(elem):
    """Nombre del nivel de referencia de la instancia, o ``None``."""
    if elem is None:
        return None
    try:
        doc = elem.Document
        lid = elem.LevelId
        if lid is None:
            return None
        try:
            if int(lid.IntegerValue) < 0:
                return None
        except Exception:
            pass
        lv = doc.GetElement(lid)
        if lv is None:
            return None
        nm = lv.Name
        if nm is None:
            return None
        return nm.ToString()
    except Exception:
        return None


_NUMERACION_COLUMNA_PARAM_CANDIDATES = (
    u"Numeracion Columna",
    u"Numeración Columna",
)


def _raw_numeracion_columna_value(elem):
    """Valor en bruto del parámetro de instancia (texto o número como string)."""
    if elem is None:
        return None
    for pname in _NUMERACION_COLUMNA_PARAM_CANDIDATES:
        try:
            p = elem.LookupParameter(pname)
        except Exception:
            p = None
        if p is None:
            continue
        try:
            if not p.HasValue:
                continue
        except Exception:
            pass
        raw = None
        try:
            st = p.StorageType
            if st == StorageType.String:
                raw = p.AsString()
                if raw is None:
                    raw = p.AsValueString()
            elif st == StorageType.Integer:
                raw = unicode(int(p.AsInteger()))
            elif st == StorageType.Double:
                raw = u"{0:.0f}".format(float(p.AsDouble()))
            else:
                raw = p.AsValueString()
        except Exception:
            try:
                raw = p.AsValueString()
            except Exception:
                raw = None
        if raw is None:
            continue
        try:
            t = raw.strip()
        except Exception:
            try:
                t = unicode(raw).strip()
            except Exception:
                t = u"{0}".format(raw).strip()
        if t:
            return t
    return None


def _format_pilar_label_from_numeracion_token(token):
    """Ej. '5' -> 'Pilar P5'; 'P5' -> 'Pilar P5'."""
    if not token:
        return None
    try:
        s = unicode(token).strip()
    except Exception:
        s = u"{0}".format(token).strip()
    if not s:
        return None
    su = s.upper()
    if not su.startswith(u"P"):
        s = u"P{0}".format(s)
    return u"Pilar {0}".format(s)


def _column_pilar_conjunto_label(elem):
    """Etiqueta UI a partir de **Numeracion Columna** (instancia)."""
    tok = _raw_numeracion_columna_value(elem)
    return _format_pilar_label_from_numeracion_token(tok)


def _build_troceo_scheme_rows(columns_ordered):
    """
    Tuplas ``(elemento, z_mm, id, height_mm, level_name, pilar_label)`` ordenadas por ``z_mm``
    ascendente (base del sólido en ``get_column_dimensions`` → ``center.Z``).

    ``height_mm``, ``level_name`` y ``pilar_label`` pueden ser ``None``; la UI infiere alturas.
    En el esquema vertical las cotas se muestran como **metros (3 decimales)** a partir de ``z_mm``;
    no se usan en pantalla el nombre de nivel ni la numeración de pilar.
    """
    rows = []
    for col in columns_ordered or []:
        z_ft = 0.0
        h_ft = None
        try:
            dims = get_column_dimensions(col)
            z_ft = float(dims[3].Z)
            h_ft = float(dims[2])
        except Exception:
            pass
        try:
            z_mm = UnitUtils.ConvertFromInternalUnits(z_ft, UnitTypeId.Millimeters)
        except Exception:
            z_mm = float(z_ft) * 304.8
        h_mm = None
        if h_ft is not None:
            try:
                h_mm = UnitUtils.ConvertFromInternalUnits(h_ft, UnitTypeId.Millimeters)
            except Exception:
                try:
                    h_mm = float(h_ft) * 304.8
                except Exception:
                    h_mm = None
        eid = _element_id_iv(col)
        if eid < 0:
            continue
        lvl = _column_reference_level_name(col)
        pilar = _column_pilar_conjunto_label(col)
        rows.append((col, float(z_mm), eid, h_mm, lvl, pilar))
    rows.sort(key=lambda r: r[1])
    if not rows and columns_ordered:
        for col in columns_ordered:
            eid = _element_id_iv(col)
            if eid < 0:
                continue
            rows.append((col, 0.0, eid, None, _column_reference_level_name(col), _column_pilar_conjunto_label(col)))
        rows.sort(key=lambda r: r[1])
    return rows


def _troceo_ubicacion_label_sort_key(lab):
    u"""Orden UI troceo: A, B, luego IA, IB, despu\u00e9s el resto alfab\u00e9tico."""
    try:
        s = unicode(lab).strip()
    except Exception:
        try:
            s = u"{0}".format(lab).strip()
        except Exception:
            s = u""
    pri = {u"A": 0, u"B": 1, u"IA": 2, u"IB": 3}
    if s in pri:
        return (pri[s], s)
    # Etiquetas externas: interiores (prefijo «I») van antes que el resto; luego alfab\u00e9tico.
    is_inner = s.startswith(u"I")
    return (10 if is_inner else 20, s)


def _troceo_fused_line_groups_core(
    doc,
    columns_ordered,
    section_grid_config,
    stirrup_configs=None,
    stirrup_bar_type_by_column_id=None,
    default_long_bar_diam_mm=12.0,
    cover=0.15,
):
    u"""
    Núcleo compartido: construye fused-world y agrupa líneas por huella
    ``(esquema_revit, span_mm_bucket)`` para producir etiquetas secuenciales
    (A, B, C, …) que coinciden con ``Armadura_Ubicacion`` real.

    Devuelve ``(labels_sorted, scheme_by_label)`` donde:
    - ``labels_sorted``: lista de letras en orden de prioridad.
    - ``scheme_by_label``: ``{letra: esquema_revit}`` (p.ej. ``{"C": "IA"}``),
      usado para calcular largos de barras internamente sin perder el tipo de corte.
    """
    if doc is None or not columns_ordered or not section_grid_config:
        return [], {}
    stirrup_configs = stirrup_configs or {}
    stirrup_bar_type_by_column_id = stirrup_bar_type_by_column_id or {}
    dims_cache = {}
    for col in columns_ordered:
        try:
            width, depth, height, center_chk, grid_vs, grid_vl = get_column_dimensions(
                col,
            )
        except Exception:
            continue
        iv = _element_id_iv(col)
        if iv < 0:
            continue
        dims_cache[iv] = (width, depth, height, center_chk, grid_vs, grid_vl)
    _column_bar_geometry = None
    try:
        from column_stirrup_creator import column_bar_geometry as _column_bar_geometry
    except Exception:
        _column_bar_geometry = None
    _early_bar_type, _, _ = _resolve_rebar_bar_type_by_diameter_mm(
        doc,
        float(default_long_bar_diam_mm),
    )
    _long_bar_model_diam_mm = float(default_long_bar_diam_mm)
    if _early_bar_type is not None:
        try:
            _long_bar_model_diam_mm = float(
                UnitUtils.ConvertFromInternalUnits(
                    float(_early_bar_type.BarModelDiameter),
                    UnitTypeId.Millimeters,
                )
            )
        except Exception:
            try:
                _nd = _rebar_nominal_diameter_mm(_early_bar_type)
                if _nd is not None:
                    _long_bar_model_diam_mm = float(_nd)
            except Exception:
                pass
    jobs = []
    for col in columns_ordered:
        try:
            iv = _element_id_iv(col)
            if iv < 0 or iv not in dims_cache:
                continue
            width, depth, height, center, grid_vs, grid_vl = dims_cache[iv]
            sk_col = _canonical_section_mm_key(width, depth)
            if sk_col not in section_grid_config:
                continue
            cfg = section_grid_config[sk_col]
            ba = int(cfg["bars_a"])
            bb = int(cfg["bars_b"])
            inc_in = bool(cfg["include_inner_outline"])
            side_short = min(width, depth)
            side_long = max(width, depth)
            short_on_x = width <= depth
            _stir_geom = None
            if _column_bar_geometry is not None:
                try:
                    _scfg = stirrup_configs.get(sk_col)
                    _sbt = getattr(_scfg, "stirrup_bar_type", None) if _scfg else None
                    _ov_bt_geom = _stirrup_bar_type_override_for_column(
                        stirrup_bar_type_by_column_id,
                        col,
                    )
                    if _ov_bt_geom is not None:
                        _sbt = _ov_bt_geom
                    if _sbt is None:
                        _sbt = _early_bar_type
                    _stir_geom = _column_bar_geometry(
                        col,
                        stirrup_bar_type=_sbt,
                        long_bar_diam_mm=_long_bar_model_diam_mm,
                    )
                except Exception:
                    _stir_geom = None
            if _stir_geom is not None:
                _pt_center, _lx, _ly, _sa, _sb, _off_long = _stir_geom
                pts = generate_bar_points(
                    _pt_center,
                    _sa,
                    _sb,
                    True,
                    ba,
                    bb,
                    _off_long,
                    inc_in,
                    _lx,
                    _ly,
                )
            else:
                pts = generate_bar_points(
                    center,
                    side_short,
                    side_long,
                    short_on_x,
                    ba,
                    bb,
                    cover,
                    inc_in,
                    grid_vs,
                    grid_vl,
                )
            jobs.append(
                dict(
                    height=height,
                    nominal_n=len(pts),
                    raw_pts=pts,
                    width=width,
                    depth=depth,
                    short_on_x=short_on_x,
                    elem=col,
                    section_key_mm=sk_col,
                    bars_a=ba,
                    bars_b=bb,
                    include_inner_outline=inc_in,
                )
            )
        except Exception:
            continue
    if not jobs:
        return [], {}
    try:
        tol = doc.Application.ShortCurveTolerance
    except Exception:
        tol = 1.0 / 304.8
    fused = fuse_vertical_world_intervals_from_jobs(jobs, tol)
    # ── Agrupación por huella (esquema_revit, span_mm_bucket) ───────────────
    # Cada línea fusionada se identifica por su tipo de corte (A/B/IA/IB…) Y
    # por su longitud total (redondeada a 10 mm).  Líneas del mismo tipo pero
    # con distinta altura —p.ej. en apilamientos de secciones mixtas— obtienen
    # letras distintas (C, D…) igual que lo hace el layout real en
    # Armadura_Ubicacion, a diferencia del parámetro Comments (A/B/IA/IB).
    _SCHEME_PRI = {u"A": 0, u"B": 1, u"IA": 2, u"IB": 3}
    group_sort_map = {}  # {(scheme_part, span_mm): sort_tuple}
    group_scheme_map = {}  # {(scheme_part, span_mm): scheme_str}
    for line_idx, fused_tup in enumerate(fused):
        span_ft = float(fused_tup[1]) if len(fused_tup) > 1 else 0.0
        agg_raw = fused_tup[3] if len(fused_tup) > 3 else u""
        try:
            agg = unicode(agg_raw).strip() if agg_raw else u""
        except Exception:
            try:
                agg = u"{0}".format(agg_raw).strip() if agg_raw is not None else u""
            except Exception:
                agg = u""
        if not agg:
            agg = _linea_fierro_nombre_alfabetico(line_idx)
        parts = (
            [p.strip() for p in agg.split(u"|") if p.strip()]
            if u"|" in agg
            else [agg]
        )
        span_mm = int(round(_arma_len_mm_round_from_internal_ft(span_ft) / 10.0) * 10)
        for p in parts:
            if not p:
                continue
            key = (p, span_mm)
            if key not in group_sort_map:
                pri = _SCHEME_PRI.get(p, 10)
                if pri == 10:
                    try:
                        pri = 10 + ord(unicode(p)[0])
                    except Exception:
                        pri = 99
                group_sort_map[key] = (pri, -span_mm, p)
            group_scheme_map[key] = p
    # Asignar letras secuenciales ordenadas por (prioridad_esquema, −span_mm)
    sorted_keys = sorted(group_sort_map.keys(), key=lambda k: group_sort_map[k])
    labels = []
    scheme_by_label = {}
    for i, gk in enumerate(sorted_keys):
        lab = _linea_fierro_nombre_alfabetico(i)
        labels.append(lab)
        scheme_by_label[lab] = group_scheme_map.get(gk, gk[0])
    try:
        labels.sort(key=_troceo_ubicacion_label_sort_key)
    except Exception:
        pass
    return labels, scheme_by_label


def troceo_fused_longitudinal_line_labels(
    doc,
    columns_ordered,
    section_grid_config,
    stirrup_configs=None,
    stirrup_bar_type_by_column_id=None,
    default_long_bar_diam_mm=12.0,
    cover=0.15,
):
    u"""Retro-compatibilidad: devuelve solo la lista de etiquetas."""
    labels, _ = _troceo_fused_line_groups_core(
        doc,
        columns_ordered,
        section_grid_config,
        stirrup_configs,
        stirrup_bar_type_by_column_id,
        default_long_bar_diam_mm,
        cover,
    )
    return labels


def troceo_fused_line_labels_and_scheme_map(
    doc,
    columns_ordered,
    section_grid_config,
    stirrup_configs=None,
    stirrup_bar_type_by_column_id=None,
    default_long_bar_diam_mm=12.0,
    cover=0.15,
):
    u"""
    Devuelve ``(labels, scheme_by_label)`` para el panel de previsualización.

    ``labels``: lista de etiquetas secuenciales (A, B, C…) que coinciden con
    ``Armadura_Ubicacion`` del layout real — incluyendo líneas C, D… cuando hay
    barras del mismo tipo con distintos spans (secciones mixtas).

    ``scheme_by_label``: ``{letra: esquema_revit}`` p.ej. ``{"C": "IA"}`` —
    imprescindible para calcular largos correctamente en ``_format_long_bar_length_estimate_mm``
    sin perder la información del tipo de corte (A, B, IA, IB).
    """
    return _troceo_fused_line_groups_core(
        doc,
        columns_ordered,
        section_grid_config,
        stirrup_configs,
        stirrup_bar_type_by_column_id,
        default_long_bar_diam_mm,
        cover,
    )


def _element_id_iv(elem):
    u"""Id numérico estable para ``ElementId`` (API 2024+ ``Value`` si es fiable; si no, ``IntegerValue``).

    En IronPython se ha visto ``Value`` presente pero devolviendo **0** con ``IntegerValue`` correcto;
    usar solo ``Value`` cuando ``int(Value) != 0`` para no colapsar todas las instancias en id 0.
    """
    if elem is None:
        return -1
    try:
        rid = elem.Id
    except Exception:
        return -1
    try:
        v = getattr(rid, "Value", None)
        if v is not None:
            iv = int(v)
            if iv != 0:
                return iv
    except Exception:
        pass
    try:
        return int(rid.IntegerValue)
    except Exception:
        return -1


def _stirrup_spacing_mm_override_for_column(stirrup_spacing_by_column_id, col):
    u"""Busca mm en el mapa por ``Value`` y, si no hay, por ``IntegerValue`` (claves antiguas/UI)."""
    d = stirrup_spacing_by_column_id or {}
    if not d or col is None:
        return None
    iv_v = _element_id_iv(col)
    try:
        if iv_v >= 0 and iv_v in d:
            return float(d[iv_v])
    except Exception:
        pass
    try:
        iv_i = int(col.Id.IntegerValue)
        if iv_i in d:
            return float(d[iv_i])
    except Exception:
        pass
    return None


def _stirrup_bar_type_override_for_column(stirrup_bar_type_by_column_id, col):
    u"""``RebarBarType`` por columna desde el esquema (claves ``Value`` e ``IntegerValue``)."""
    d = stirrup_bar_type_by_column_id or {}
    if not d or col is None:
        return None
    iv_v = _element_id_iv(col)
    try:
        if iv_v >= 0 and iv_v in d:
            return d[iv_v]
    except Exception:
        pass
    try:
        iv_i = int(col.Id.IntegerValue)
        if iv_i in d:
            return d[iv_i]
    except Exception:
        pass
    return None


def _element_id_from_int(iv):
    """``ElementId`` desde entero de API (Value / IntegerValue)."""
    if iv is None:
        return None
    try:
        return ElementId(int(iv))
    except Exception:
        try:
            return ElementId(Convert.ToInt64(int(iv)))
        except Exception:
            return None


def _troceo_nominal_diam_mm_for_seg_index(troceo_segment_diams, seg_i, fallback_mm):
    """Ø nominal (mm) del tramo ``seg_i`` (0 = más bajo); si sobran tramos usa el último definido."""
    if not troceo_segment_diams:
        return float(fallback_mm)
    try:
        if seg_i < len(troceo_segment_diams):
            return float(troceo_segment_diams[seg_i])
        return float(troceo_segment_diams[-1])
    except Exception:
        return float(fallback_mm)


def _troceo_sorted_cut_z_ft_from_planes(a_bar_cut_planes):
    """
    Cotas Z (unidades internas) de los planos de troceo **tipo A**, orden ascendentes.
    Se usa ``Origin.Z`` en planos casi horizontales; en otro caso también el origen.
    """
    zs = []
    for pl in a_bar_cut_planes or []:
        if pl is None:
            continue
        try:
            zs.append(float(pl.Origin.Z))
        except Exception:
            pass
    zs.sort()
    out = []
    merge_eps = 1e-6
    for z in zs:
        if not out or z > out[-1] + merge_eps:
            out.append(z)
    return out


def _troceo_ui_segment_index_for_z_mid(
    z_mid_ft,
    z_cuts_sorted_ft,
    n_ui_segments,
    fallback_index,
):
    """
    Índice de tramo de la UI (0 = bajo → N) según la cota media ``z_mid_ft`` del tramo modelo.

    ``z_cuts_sorted_ft`` tiene N valores (cotas de los N planos A); la UI define N+1 tipos.
    Si la geometría no cuadra con la UI, devuelve ``fallback_index`` (comportamiento previo por índice local).
    """
    try:
        n_ui = int(n_ui_segments)
        bc = z_cuts_sorted_ft or []
        if n_ui < 1:
            return int(fallback_index)
        if not bc:
            return 0 if n_ui == 1 else int(fallback_index)
        if len(bc) + 1 != n_ui:
            return int(fallback_index)
        z = float(z_mid_ft)
        if z < float(bc[0]):
            return 0
        for i in range(1, len(bc)):
            if z < float(bc[i]):
                return int(i)
        return int(len(bc))
    except Exception:
        return int(fallback_index)


def get_positions(length, count, edge_cover):
    if count < 2:
        return []

    if count == 2:
        return [-length / 2.0 + edge_cover, length / 2.0 - edge_cover]

    span = length - (2.0 * edge_cover)
    spacing = span / float(count - 1)
    return [
        -length / 2.0 + edge_cover + (i * spacing) for i in range(count)
    ]


def perimeter_outer_bar_count(bars_a, bars_b):
    if bars_a < 2 or bars_b < 2:
        return 0
    return (2 * bars_a) + (2 * bars_b) - 4


def perimeter_inner_outline_count(bars_a, bars_b):
    if bars_a < 4 or bars_b < 4:
        return 0
    ia = bars_a - 2
    ib = bars_b - 2
    if ia < 2 or ib < 2:
        return 0
    return (2 * ia) + (2 * ib) - 4


def _canonical_section_mm_key(width_ft, depth_ft):
    """
    Tupla estable (mm) ``(lado_corto_mm, lado_largo_mm)`` en planta, independiente del eje proyecto.
    """
    try:
        wx = UnitUtils.ConvertFromInternalUnits(abs(float(width_ft)), UnitTypeId.Millimeters)
        dx = UnitUtils.ConvertFromInternalUnits(abs(float(depth_ft)), UnitTypeId.Millimeters)
    except Exception:
        wx = abs(float(width_ft)) * 304.8
        dx = abs(float(depth_ft)) * 304.8
    s = float(min(wx, dx))
    L = float(max(wx, dx))
    return int(round(s)), int(round(L))


def hilos_esperados_una_columna(bars_a, bars_b, include_inner_outline):
    outer_n = perimeter_outer_bar_count(bars_a, bars_b)
    inner_n = (
        perimeter_inner_outline_count(bars_a, bars_b)
        if include_inner_outline
        else 0
    )
    return int(outer_n + inner_n)


_BAR_LAYOUT_FORM_SINGLETON_KEY = (
    "Arainco.column_reinforcement_layout_rps.BarInputFormSingleton"
)


def _show_bar_layout_form_singleton(win_form):

    """

    Una sola ventana WinForms por herramienta (AppDomain). Devuelve ``DialogResult``

    ó ``None`` si ya había una instancia activa (no abre otra).

    """

    ad = AppDomain.CurrentDomain
    prev = ad.GetData(_BAR_LAYOUT_FORM_SINGLETON_KEY)

    try:
        if prev is not None:

            alive = getattr(prev, "IsDisposed", False) is False
            visible = getattr(prev, "Visible", False)
            if alive and visible:

                prev.Activate()
                try:
                    if getattr(prev, "WindowState", None) == FormWindowState.Minimized:
                        prev.WindowState = FormWindowState.Normal
                except Exception:
                    pass

                TaskDialog.Show(
                    "Layout columna",
                    u"La herramienta ya esta en ejecucion.",
                )
                return None
    except Exception:
        pass

    ad.SetData(_BAR_LAYOUT_FORM_SINGLETON_KEY, win_form)

    def _on_closed(s, evt):
        try:
            cur = ad.GetData(_BAR_LAYOUT_FORM_SINGLETON_KEY)
            if cur is win_form:

                ad.SetData(_BAR_LAYOUT_FORM_SINGLETON_KEY, None)

        except Exception:
            pass

    try:

        win_form.FormClosed += _on_closed

    except Exception:

        pass

    return win_form.ShowDialog()


def _perimeter_ij_clockwise_corner_top_left_ix(nx, ny):
    """
    Perímetro ``nx × ny`` en sentido **horario**, empezando en ``(0, ny-1)``
    (arista “superior” izquierda→derecha, luego derecha, inferior derecha→izquierda, izquierda).
    """
    if nx < 2 or ny < 2:
        return []
    out = []
    for ix in range(nx):
        out.append((ix, ny - 1))
    for iy in range(ny - 2, -1, -1):
        out.append((nx - 1, iy))
    for ix in range(nx - 2, -1, -1):
        out.append((ix, 0))
    for iy in range(1, ny - 1):
        out.append((0, iy))
    return out


def _outer_outline_ij_ordered(bars_a, bars_b):
    return _perimeter_ij_clockwise_corner_top_left_ix(bars_a, bars_b)


def _inner_outline_ij_ordered(bars_a, bars_b):
    nx = int(bars_a) - 2
    ny = int(bars_b) - 2
    if nx < 2 or ny < 2:
        return []
    return [
        (ix + 1, iy + 1)
        for ix, iy in _perimeter_ij_clockwise_corner_top_left_ix(nx, ny)
    ]


def generate_bar_points(
    center,
    side_short,
    side_long,
    short_on_x,
    bars_a,
    bars_b,
    cover,
    include_inner_outline,
    v_short=None,
    v_long=None,
):
    """
    ``dict`` ``pt`` + ``bar_enum``: perímetro exterior **A/B alternos** en orden horario desde
    ``(0, B-1)``; perímetro interior **IA/IB** con la misma regla.

    Si ``v_short`` y ``v_long`` son unitarios en modelo (ejes del símbolo, ⟂ al eje del pilar),
    los desplazamientos son ``da * v_short + db * v_long`` (rejilla alineada al elemento rotado).
    Si son ``None``, se mantiene el criterio histórico ±X/±Y proyecto según ``short_on_x``.
    """
    offs_a = get_positions(side_short, bars_a, cover)
    offs_b = get_positions(side_long, bars_b, cover)

    if len(offs_a) != bars_a or len(offs_b) != bars_b:
        raise Exception("Error interno en reparto de posiciones rejilla.")

    def pt_at(ix, iy):
        da = offs_a[ix]
        db = offs_b[iy]
        if v_short is not None and v_long is not None:
            try:
                dxy = v_short.Multiply(float(da)).Add(v_long.Multiply(float(db)))
                return center.Add(dxy)
            except Exception:
                pass
        if short_on_x:
            return XYZ(center.X + da, center.Y + db, center.Z)
        return XYZ(center.X + db, center.Y + da, center.Z)

    points = []
    for k, (ix, iy) in enumerate(_outer_outline_ij_ordered(bars_a, bars_b)):
        points.append(
            dict(pt=pt_at(ix, iy), bar_enum=("A" if (k % 2 == 0) else "B"))
        )

    if not include_inner_outline:
        return points

    if bars_a < 4 or bars_b < 4:
        return points

    for ik, (ix, iy) in enumerate(_inner_outline_ij_ordered(bars_a, bars_b)):
        points.append(
            dict(
                pt=pt_at(ix, iy),
                bar_enum=("IA" if ik % 2 == 0 else "IB"),
            )
        )

    return points


XY_KEY_DECIMALS_DEFAULT = 9


# Desarrollo: hasta leer el ``RebarBarType``. Estiramiento +Z tras fusión = tabla empalme/empotramiento.
LAYOUT_BAR_NOMINAL_DIAM_MM = 12.0

# Tabla BIMTools para traslape/empotramiento: ``None`` = base proyecto; ``"G25"``/``"G35"``/``"G45"``.
LAYOUT_EMBED_CONCRETE_GRADE = None

# ``RebarShape`` del proyecto (nombre visible) para L = tronco recto + pata en un solo extremo.
COLUMN_REBAR_L_SHAPE_DISPLAY_NAME = u"02"

# Instancia ``Rebar``: identificador de línea de fierro (troceo + patas), A/B/… (ver ``_apply_linea_fierro_armadura_ubicacion``).
COLUMN_ARMA_UBICACION_PARAM = u"Armadura_Ubicacion"


# Marcadores de verificación (**líneas modelo** horizontales, mitad Z del hilo fusionado):

# Dos **GraphicsStyle** (subcategorías de líneas proyecto) por hints de nombre más **trazo X vs Y**.
_SCHEME_VERIFY_MARKER_ENABLED = False
_SCHEME_VERIFY_MARKER_HALF_MM = 90.0
_SCHEME_VERIFY_HINT_CORNER_LINES = ("ARAINCOVERIFYA", "THIN", "DELGADA")
_SCHEME_VERIFY_HINT_EDGE_LINES = ("ARAINCOVERIFYB", "WIDE", "ANCHA", "HIDDEN")

# Visualización de **planos de corte** barras A: origen (``Location`` start), normal (segmento) y trazo en plano.
_CUT_PLANE_MARKER_ENABLED = False
_CUT_PLANE_NORMAL_DISPLAY_MM = 700.0
_CUT_PLANE_INPLANE_LEG_MM = 350.0
_CUT_PLANE_HINT_LINESTYLE = (
    "ARAINCOCUTPLANE",
    "CUTPLANE",
    "DASH",
    "DASHED",
)


# Prisma +Z: desde ``Z`` techo fusionado hasta ``Z_topo + L(Ø)``. Prisma −Z (arranque): desde ``Z_suelo − L(Ø)`` hasta ``Z_suelo`` fusionado
# excluye hilos cuyo arranque ya se estira por fundación unida. En XY: **radio nominal** ``Ø/2`` más margen fijo (rejilla/redondeos).
_EMBED_PROBE_XY_MARGIN_MM = 1.0
_EMBED_PROBE_MIN_HALF_SIDE_MM = 2.0
_TOL_VOL_INTERSECCION_EMBED_FT3 = 5e-8

# Tras revertir el estiramiento por tabla Ø (sin colisión), acortamiento adicional del fusionado en −Z:
# ``_REVOKE_EMBED_EXTRA_SHRINK_MM`` + mitad del Ø nominal (``LAYOUT_BAR_NOMINAL_DIAM_MM / 2``).
_REVOKE_EMBED_EXTRA_SHRINK_MM = 25.0


# --- Fundaciones unidas bajo cara inferior columnas seleccionadas (Join Geometry, BB alineadas) -----
# Estir −Z desde el arranque fusionado: ``altura_BB_fundación − _FOUNDATION_STRETCH_DEDUCTION_MM`` [mm].
_FOUNDATION_STRETCH_DEDUCTION_MM = 50.0
_FOUNDATION_JOIN_FACE_Z_TOLERANCE_MM = 35.0
_FOUNDATION_JOIN_OVERLAP_XY_MM = 75.0
_CAT_STRUCT_FOUNDATION_IV = int(BuiltInCategory.OST_StructuralFoundation)


def _coerce_icollection_join_to_element_ids(raw):
    if raw is None:
        return []
    out = []
    try:
        for jid in raw:
            if jid is not None and jid != ElementId.InvalidElementId:
                out.append(jid)
    except Exception:
        pass
    if out:
        return out
    try:
        n = int(raw.Count)
    except Exception:
        n = 0
    for i in range(n):
        jid = None
        try:
            jid = raw[i]
        except Exception:
            try:
                jid = raw.get_Item(i)
            except Exception:
                jid = None
        if jid is not None and jid != ElementId.InvalidElementId:
            out.append(jid)
    return out


def _joined_elem_ids_revit(doc, host):
    if doc is None or host is None:
        return []
    raw = None
    for getter in (
        lambda: JoinGeometryUtils.GetJoinedElements(doc, host),
        lambda: JoinGeometryUtils.GetJoinedElements(doc, host.Id),
    ):
        try:
            raw = getter()
        except Exception:
            raw = None
        if raw is not None:
            break
    return _coerce_icollection_join_to_element_ids(raw)


def _coerce_element_id_int(eid_or_iv):
    u"""Entero estable desde ``ElementId`` o compatible (no confundir con ``_element_id_iv(elem)`` para instancias)."""
    if eid_or_iv is None:
        return 0
    try:
        return int(eid_or_iv.IntegerValue)
    except Exception:
        pass
    try:
        return int(getattr(eid_or_iv, "Value"))
    except Exception:
        pass
    try:
        return int(eid_or_iv)
    except Exception:
        return 0


def _vertex_ranges_xy_overlap_padded(rng_a, rng_b, pad_ft):
    """Solape en XY entre rangos ``(min_x, max_x, min_y, max_y, ...)`` con margen ``pad_ft``."""
    if rng_a is None or rng_b is None:
        return False
    try:
        if rng_a[0] is None or rng_b[0] is None:
            return False
    except Exception:
        return False
    p = abs(float(pad_ft))
    try:
        min_ax = float(rng_a[0])
        max_ax = float(rng_a[1])
        min_ay = float(rng_a[2])
        max_ay = float(rng_a[3])
        min_bx = float(rng_b[0])
        max_bx = float(rng_b[1])
        min_by = float(rng_b[2])
        max_by = float(rng_b[3])
        if max_ax + p < min_bx - p:
            return False
        if max_bx + p < min_ax - p:
            return False
        if max_ay + p < min_by - p:
            return False
        if max_by + p < min_ay - p:
            return False
    except Exception:
        return False
    return True


def _elem_is_structural_foundation(elem):
    if elem is None:
        return False
    try:
        cat = elem.Category
        if cat is None:
            return False
        return int(cat.Id.IntegerValue) == _CAT_STRUCT_FOUNDATION_IV
    except Exception:
        return False


def column_bottom_joined_foundation_stretch_down_mm(doc, column):
    """
    ``JoinGeometryUtils``: fundaciones estructurales unidas donde la cara superior de la fundación
    (``Z`` máximo del sólido) está alineada con la base del pilar (``Z`` mínimo del sólido),
    con solape XY deducido de vértices de sólidos (no BB elemento).
    Retorna mm a alargar hacia −Z (``altura géom. − 50``). ``0`` si no aplica.
    """
    if doc is None or column is None:
        return 0.0
    rng_c = _solid_aggregate_vertex_ranges_ft(column)
    if rng_c is None:
        return 0.0
    try:
        z_col_min = float(rng_c[4])
    except Exception:
        return 0.0
    try:
        tol_z_ft = UnitUtils.ConvertToInternalUnits(
            float(_FOUNDATION_JOIN_FACE_Z_TOLERANCE_MM),
            UnitTypeId.Millimeters,
        )
        pad_xy_ft = UnitUtils.ConvertToInternalUnits(
            float(_FOUNDATION_JOIN_OVERLAP_XY_MM),
            UnitTypeId.Millimeters,
        )
    except Exception:
        tol_z_ft = 35.0 / 3048.0
        pad_xy_ft = 75.0 / 3048.0
    best_mm = 0.0
    jids = _joined_elem_ids_revit(doc, column)
    if not jids:
        return 0.0
    for jid in jids:
        try:
            el = doc.GetElement(jid)
        except Exception:
            el = None
        if el is None or not _elem_is_structural_foundation(el):
            continue
        rng_f = _solid_aggregate_vertex_ranges_ft(el)
        if rng_f is None:
            continue
        if not _vertex_ranges_xy_overlap_padded(rng_c, rng_f, pad_xy_ft):
            continue
        try:
            z_f_max = float(rng_f[5])
            h_ft = float(rng_f[5]) - float(rng_f[4])
        except Exception:
            continue
        if abs(z_f_max - z_col_min) > float(tol_z_ft):
            continue
        if h_ft <= 1e-9:
            continue
        try:
            h_mm = float(
                UnitUtils.ConvertFromInternalUnits(h_ft, UnitTypeId.Millimeters)
            )
        except Exception:
            continue
        s_mm = h_mm - float(_FOUNDATION_STRETCH_DEDUCTION_MM)
        if s_mm > best_mm:
            best_mm = s_mm
    return max(best_mm, 0.0)


def build_selected_columns_foundation_down_ft(doc, columns_seq):
    """
    Map ``IntegerValue`` (~``Value`` respaldo) de columna → estir −Z interior [pie] desde arranque.
    Solo columnas donde ``column_bottom_joined_foundation_stretch_down_mm > 0``.
    """
    m = {}
    if doc is None or not columns_seq:
        return m
    for col in columns_seq:
        if col is None:
            continue
        mm_ext = column_bottom_joined_foundation_stretch_down_mm(doc, col)
        if mm_ext <= 1e-6:
            continue
        try:
            ft_down = UnitUtils.ConvertToInternalUnits(
                float(mm_ext),
                UnitTypeId.Millimeters,
            )
            iv = _coerce_element_id_int(col.Id)
            if iv == 0:
                continue
            m[iv] = max(float(m.get(iv, 0.0)), float(ft_down))
        except Exception:
            continue
    return m


def _contrib_max_foundation_down_ft(contrib_elem_ids, col_iv_to_ft):
    mx = 0.0
    if not col_iv_to_ft or not contrib_elem_ids:
        return mx
    for cid in contrib_elem_ids:
        iv = _coerce_element_id_int(cid)
        mx = max(mx, float(col_iv_to_ft.get(iv, 0.0)))
    return mx


def _geometry_options_structure_solids():
    opts = Options()
    try:
        opts.ComputeReferences = False
    except Exception:
        pass
    try:
        opts.DetailLevel = ViewDetailLevel.Fine
    except Exception:
        pass
    try:
        opts.IncludeNonVisibleObjects = True
    except Exception:
        pass
    return opts


def _iter_solids_revit_element(elem, opts):
    if elem is None:
        return
    try:
        ge = elem.get_Geometry(opts)
    except Exception:
        return
    if ge is None:
        return
    for obj in ge:
        if obj is None:
            continue
        if isinstance(obj, Solid):
            try:
                if float(obj.Volume) < 1e-11:
                    continue
            except Exception:
                continue
            yield obj
        elif isinstance(obj, GeometryInstance):
            try:
                sub = obj.GetInstanceGeometry()
            except Exception:
                continue
            if sub is None:
                continue
            for g2 in sub:
                if isinstance(g2, Solid):
                    try:
                        if float(g2.Volume) < 1e-11:
                            continue
                    except Exception:
                        continue
                    yield g2


_LOOKUP_COL_WIDTH_NAMES = (
    "b",
    "B",
    "Width",
    "Ancho",
    "Ancho nominal",
    "width",
)
_LOOKUP_COL_DEPTH_NAMES = (
    "h",
    "H",
    "Depth",
    "Profundidad",
    "depth",
)


def _lookup_column_section_width_depth_ft(elem):
    """Lee ancho/profundidad de sección en **pies internas** desde instancia y tipo."""
    if elem is None:
        return None, None
    doc = getattr(elem, "Document", None)
    et = None
    try:
        tid = elem.GetTypeId()
        if doc is not None and tid is not None and tid != ElementId.InvalidElementId:
            et = doc.GetElement(tid)
    except Exception:
        et = None
    w = d = None
    for target in (elem, et):
        if target is None:
            continue
        if w is None:
            for n in _LOOKUP_COL_WIDTH_NAMES:
                try:
                    p = target.LookupParameter(n)
                    if p is not None and p.HasValue:
                        w = float(p.AsDouble())
                        break
                except Exception:
                    pass
        if d is None:
            for n in _LOOKUP_COL_DEPTH_NAMES:
                try:
                    p = target.LookupParameter(n)
                    if p is not None and p.HasValue:
                        d = float(p.AsDouble())
                        break
                except Exception:
                    pass
        if w is not None and d is not None:
            break
    return w, d


def _column_curve_endpoints_axis_ft(column, tol_ft):
    """``(p0, p1, eje_unitario, longitud_curva)``. Sin curva válida: ``BasisZ`` y longitud ``0``."""
    tol = max(abs(float(tol_ft)), 1e-12)
    if column is None:
        return None, None, XYZ.BasisZ, 0.0
    try:
        loc = column.Location
    except Exception:
        return None, None, XYZ.BasisZ, 0.0
    cr = getattr(loc, "Curve", None)
    if cr is not None:
        try:
            p0 = cr.GetEndPoint(0)
            p1 = cr.GetEndPoint(1)
            v = p1 - p0
            ln = float(v.GetLength())
            if ln >= tol:
                return p0, p1, v.Normalize(), ln
            return p0, p1, XYZ.BasisZ, ln
        except Exception:
            pass
    pt = getattr(loc, "Point", None)
    if pt is not None:
        return pt, pt, XYZ.BasisZ, 0.0
    return None, None, XYZ.BasisZ, 0.0


def _iter_vertices_from_solid(solid):
    if solid is None:
        return
    try:
        if float(solid.Volume) < 1e-11:
            return
    except Exception:
        pass
    try:
        edges = solid.Edges
        ne = int(edges.Size)
    except Exception:
        return
    for i in range(ne):
        try:
            edge = edges.get_Item(i)
            crv = edge.AsCurve()
            if crv is None:
                continue
            for k in (0, 1):
                try:
                    yield crv.GetEndPoint(k)
                except Exception:
                    pass
        except Exception:
            continue


def _solid_aggregate_vertex_ranges_ft(elem, opts=None):
    """
    ``(min_x, max_x, min_y, max_y, min_z, max_z)`` en coordenadas modelo desde aristas de sólidos,
    o ``None`` si no hay vértices.
    """
    if elem is None:
        return None
    opts = opts if opts is not None else _geometry_options_structure_solids()
    min_x = min_y = min_z = None
    max_x = max_y = max_z = None
    count = 0
    for solid in _iter_solids_revit_element(elem, opts):
        for pt in _iter_vertices_from_solid(solid):
            count += 1
            x, y, z = float(pt.X), float(pt.Y), float(pt.Z)
            min_x = x if min_x is None else min(min_x, x)
            max_x = x if max_x is None else max(max_x, x)
            min_y = y if min_y is None else min(min_y, y)
            max_y = y if max_y is None else max(max_y, y)
            min_z = z if min_z is None else min(min_z, z)
            max_z = z if max_z is None else max(max_z, z)
    if count == 0:
        return None
    return (min_x, max_x, min_y, max_y, min_z, max_z)


def _oriented_section_width_depth_ft(origin, axis_unit, points):
    """
    Extents en pies en el plano ⟂ ``axis_unit`` que pasan por ``origin``.
    Devuelve ``(ancho, fondo)`` como máximo−mínimo sobre dos ejes ortogonales arbitrarios en ese plano.
    """
    try:
        ax = axis_unit.Normalize()
    except Exception:
        ax = XYZ.BasisZ
    ref = XYZ.BasisZ
    if abs(float(ax.DotProduct(ref))) > 0.99:
        ref = XYZ.BasisX
    try:
        v1 = ref.CrossProduct(ax)
        if v1.GetLength() < 1e-9:
            v1 = XYZ.BasisY.CrossProduct(ax)
        if v1.GetLength() < 1e-9:
            return 0.0, 0.0
        v1 = v1.Normalize()
        v2 = ax.CrossProduct(v1).Normalize()
    except Exception:
        return 0.0, 0.0
    dots1 = []
    dots2 = []
    o = origin
    for p in points:
        try:
            r = p - o
            dots1.append(float(r.DotProduct(v1)))
            dots2.append(float(r.DotProduct(v2)))
        except Exception:
            continue
    if len(dots1) < 2:
        return 0.0, 0.0
    w = float(max(dots1) - min(dots1))
    d = float(max(dots2) - min(dots2))
    return abs(w), abs(d)


def _geometry_options_mesh_solids():
    """Opciones para obtener mallas (`Face.Triangulate`) como el script de geometría real."""
    opts = Options()
    try:
        opts.ComputeReferences = True
    except Exception:
        pass
    try:
        opts.IncludeNonVisibleObjects = True
    except Exception:
        pass
    try:
        opts.DetailLevel = ViewDetailLevel.Fine
    except Exception:
        pass
    return opts


def _column_insertion_point_ft(column):
    """Origen de referencia para coords. locales: punto de instancia o extremo 0 de ``LocationCurve``."""
    if column is None:
        return None
    try:
        loc = column.Location
    except Exception:
        return None
    if isinstance(loc, LocationCurve):
        try:
            cr = loc.Curve
            if cr is not None:
                return cr.GetEndPoint(0)
        except Exception:
            return None
    if isinstance(loc, LocationPoint):
        try:
            return loc.Point
        except Exception:
            return None
    return None


def _normalize_xyz_safe(v, fallback):
    if v is None:
        return fallback
    try:
        if float(v.GetLength()) < 1e-12:
            return fallback
        return v.Normalize()
    except Exception:
        return fallback


def _plan_axes_from_family_transform_ft(column, width, depth):
    """
    Para rejilla en planta: ``v_short`` y ``v_long`` unitarios según ``BasisX``/``BasisY`` del
    ``FamilyInstance`` y convención ``width`` ⟺ lado medido en ``BasisX`` cuando ``width<=depth``.
    """
    if column is None:
        return None, None
    try:
        transform = column.GetTransform()
    except Exception:
        return None, None
    if transform is None:
        return None, None
    lx = _normalize_xyz_safe(transform.BasisX, None)
    ly = _normalize_xyz_safe(transform.BasisY, None)
    if lx is None or ly is None:
        return None, None
    try:
        if float(width) <= float(depth):
            return lx, ly
        return ly, lx
    except Exception:
        return None, None


def _try_column_dimensions_transform_mesh_ft(column, tol_ft):
    """
    Dimensiones en **pies internas** y centro en modelo desde la transformación del ``FamilyInstance``
    y vértices de mallas de todas las caras de los sólidos (sin BB elemento).

    Equivalente funcional al script de referencia (Location + GetTransform + Triangulate).
    Devuelve ``None`` si no hay ``GetTransform``, inserción, sólidos o extensión válida.

    Cuando tiene éxito, devuelve
    ``(side_lx, side_ly, height, center, v_short, v_long)`` en pies internas / ``XYZ`` modelo,
    con ``v_short``/``v_long`` unitarios ⟂ al eje del pilar para colocar la rejilla alineada al giro.
    """
    tol = max(abs(float(tol_ft)), 1e-12)
    if column is None:
        return None
    insertion_pt = _column_insertion_point_ft(column)
    if insertion_pt is None:
        return None
    try:
        transform = column.GetTransform()
    except Exception:
        return None
    if transform is None:
        return None

    lx = _normalize_xyz_safe(transform.BasisX, XYZ.BasisX)
    ly = _normalize_xyz_safe(transform.BasisY, XYZ.BasisY)
    lz = _normalize_xyz_safe(transform.BasisZ, XYZ.BasisZ)

    mesh_opts = _geometry_options_mesh_solids()
    solids_found = False
    min_x = min_y = min_z = None
    max_x = max_y = max_z = None
    vtx_count = 0

    for solid in _iter_solids_revit_element(column, mesh_opts):
        solids_found = True
        try:
            faces = solid.Faces
            nf = int(faces.Size)
        except Exception:
            continue
        for fi in range(nf):
            try:
                face = faces.get_Item(fi)
                mesh = face.Triangulate()
                if mesh is None:
                    continue
                verts = mesh.Vertices
                nv = int(verts.Count)
            except Exception:
                continue
            for vi in range(nv):
                try:
                    try:
                        pt = verts[vi]
                    except Exception:
                        pt = verts.get_Item(vi)
                    vec = pt - insertion_pt
                    xv = float(vec.DotProduct(lx))
                    yv = float(vec.DotProduct(ly))
                    zv = float(vec.DotProduct(lz))
                    vtx_count += 1
                    min_x = xv if min_x is None else min(min_x, xv)
                    max_x = xv if max_x is None else max(max_x, xv)
                    min_y = yv if min_y is None else min(min_y, yv)
                    max_y = yv if max_y is None else max(max_y, yv)
                    min_z = zv if min_z is None else min(min_z, zv)
                    max_z = zv if max_z is None else max(max_z, zv)
                except Exception:
                    continue

    if not solids_found or vtx_count == 0:
        return None
    if None in (min_x, max_x, min_y, max_y, min_z, max_z):
        return None

    side_a = float(max_x) - float(min_x)
    side_b = float(max_y) - float(min_y)
    height = float(max_z) - float(min_z)

    if side_a <= tol or side_b <= tol or height <= tol:
        return None

    cx_loc = 0.5 * (float(min_x) + float(max_x))
    cy_loc = 0.5 * (float(min_y) + float(max_y))
    cz_loc = float(min_z)

    if side_a <= side_b:
        v_short = lx
        v_long = ly
    else:
        v_short = ly
        v_long = lx

    try:
        center = (
            insertion_pt
            + lx.Multiply(cx_loc)
            + ly.Multiply(cy_loc)
            + lz.Multiply(cz_loc)
        )
    except Exception:
        return None

    return side_a, side_b, height, center, v_short, v_long


def _solidos_intersectan_volumen(solid_a, solid_b, tol_volumen=_TOL_VOL_INTERSECCION_EMBED_FT3):
    if solid_a is None or solid_b is None:
        return False
    try:
        va = float(solid_a.Volume)
        vb = float(solid_b.Volume)
    except Exception:
        return False
    if va <= 1e-12 or vb <= 1e-12:
        return False
    try:
        inter = BooleanOperationsUtils.ExecuteBooleanOperation(
            solid_a,
            solid_b,
            BooleanOperationsType.Intersect,
        )
    except Exception:
        return False
    if inter is None:
        return False
    try:
        return float(inter.Volume) > float(tol_volumen)
    except Exception:
        return False


def _contrib_elem_id_menor_bbox_max_z(doc, contrib_elem_ids):
    """
    Entre pilares contribuyentes a este XY fusionado: ``ElementId`` con menor ``Z`` máximo de la
    geometría sólida (vértices), típico del pilar ``inferior`` en un apilamiento.

    Sólo **clasifica qué instancia omitir** en la evaluación cuando el prisma **se solapa accidentalmente**

    con volumen del tramo inferior de apilamiento; **la colisión válida sigue siendo booleana prisma–solid**.
    """
    if doc is None or not contrib_elem_ids:
        return None
    best_eid = None
    best_mz = None
    for eid in contrib_elem_ids:
        try:
            el = doc.GetElement(eid)
        except Exception:
            el = None
        if el is None:
            continue
        rng = _solid_aggregate_vertex_ranges_ft(el)
        if rng is None:
            continue
        try:
            mz = float(rng[5])
        except Exception:
            continue
        if best_mz is None or mz < best_mz:
            best_mz = mz
            try:
                best_eid = el.Id
            except Exception:
                best_eid = None
    return best_eid


def _element_ids_coinciden(id_a, id_b):
    if id_a is None or id_b is None:
        return False
    try:
        return int(getattr(id_a, "IntegerValue", id_a)) == int(getattr(id_b, "IntegerValue", id_b))
    except Exception:
        try:
            return id_a == id_b
        except Exception:
            return False


def _build_vertical_square_prism_solid(px, py, z_start_ft, half_side_ft, height_ft):
    """
    Extrusión +Z desde ``z_start_ft``: cuadrado en XY centrado en (px, py).
    """
    hw = abs(float(half_side_ft))
    hgt = abs(float(height_ft))
    hs = XYZ(float(px), float(py), float(z_start_ft))
    p1 = XYZ(hs.X - hw, hs.Y - hw, hs.Z)
    p2 = XYZ(hs.X + hw, hs.Y - hw, hs.Z)
    p3 = XYZ(hs.X + hw, hs.Y + hw, hs.Z)
    p4 = XYZ(hs.X - hw, hs.Y + hw, hs.Z)
    try:
        loop = CurveLoop.Create(
            [
                Line.CreateBound(p1, p2),
                Line.CreateBound(p2, p3),
                Line.CreateBound(p3, p4),
                Line.CreateBound(p4, p1),
            ]
        )
    except Exception:
        return None
    try:
        sol = GeometryCreationUtilities.CreateExtrusionGeometry(
            [loop],
            XYZ.BasisZ,
            hgt,
        )
    except Exception:
        return None
    if sol is None or float(sol.Volume) < 1e-15:
        return None
    return sol


def embed_stretch_collides_any_column_solids(
    doc,
    xyz_base_xyz,
    fused_span_ft,
    dz_embed_ft,
    bar_nominal_mm,
    column_instances,
    geom_opts,
    contrib_elem_ids,
):
    """
    ``True`` si el prisma de ensayo coincide con **el segmento geométrico del empotramiento +Z**:

    desde el techo ``Z_topo`` **del trazo ya fusionado (sin tabla)** hasta ``Z_topo + L``

    donde ``L`` es el estiramiento por diámetro (mismo alcance que el tramo modelo antes de aplicar reverte),

    tiene intersección booleana volumétrica contra **algún** ``Solid`` ``Structural Columns``,

    omitiendo sólo la instancia contribuyente identificada con menor ``BBox.Max.Z`` (**clasificación inferior

    típico en apilamiento**; no usa BB como test de golpe).
    """
    dz_e = abs(float(dz_embed_ft))
    if doc is None or dz_e <= 1e-12:
        return False
    span = float(fused_span_ft)
    z_top_fused = float(xyz_base_xyz.Z) + span
    z0 = float(z_top_fused)
    h_extr = dz_e
    if h_extr <= 1e-12:
        return False
    half_w_mm = (
        float(bar_nominal_mm) / 2.0
        + float(_EMBED_PROBE_XY_MARGIN_MM)
    )
    half_w_mm = max(half_w_mm, float(_EMBED_PROBE_MIN_HALF_SIDE_MM))
    half_w_ft = UnitUtils.ConvertToInternalUnits(half_w_mm, UnitTypeId.Millimeters)
    probe = _build_vertical_square_prism_solid(
        float(xyz_base_xyz.X),
        float(xyz_base_xyz.Y),
        z0,
        half_w_ft,
        h_extr,
    )
    if probe is None:
        return False
    omit_eid = _contrib_elem_id_menor_bbox_max_z(doc, contrib_elem_ids)
    try:
        for col in column_instances or []:
            if col is None:
                continue
            if omit_eid is not None and _element_ids_coinciden(col.Id, omit_eid):
                continue
            for sd in _iter_solids_revit_element(col, geom_opts):
                if _solidos_intersectan_volumen(probe, sd):
                    return True
    except Exception:
        return False
    return False


def embed_start_collides_any_column_solids(
    doc,
    xyz_floor_fused_xyz,
    dz_embed_ft,
    bar_nominal_mm,
    column_instances,
    geom_opts,
    contrib_elem_ids,
):
    """

    ``True`` si el prisma de ensayo desde ``Z_inferior`` **hasta el arranque fusionado** ``Z_suelo`` coincide

    con **el segmento geométrico del empotramiento −Z** (tabla igual que +Z): ``Z_inferior = Z_suelo − L``,

    ``Z_suelo = Z`` del resultado fusionado (``base_xyz.Z``, suelo nominal del conjunto −Z antes de fundación).


    Omit sólo las columna identificada con **menor** ``BBox.Max.Z`` (**misma clasificación inferior en apilamiento** que el prisma +Z).

    """

    dz_e = abs(float(dz_embed_ft))
    if doc is None or dz_e <= 1e-12:
        return False

    z_suelo = float(xyz_floor_fused_xyz.Z)

    z0 = z_suelo - dz_e
    if dz_e <= 1e-12:
        return False
    half_w_mm = (
        float(bar_nominal_mm) / 2.0
        + float(_EMBED_PROBE_XY_MARGIN_MM)
    )
    half_w_mm = max(half_w_mm, float(_EMBED_PROBE_MIN_HALF_SIDE_MM))
    half_w_ft = UnitUtils.ConvertToInternalUnits(half_w_mm, UnitTypeId.Millimeters)
    probe = _build_vertical_square_prism_solid(
        float(xyz_floor_fused_xyz.X),
        float(xyz_floor_fused_xyz.Y),
        z0,
        half_w_ft,
        dz_e,
    )
    if probe is None:
        return False
    omit_eid = _contrib_elem_id_menor_bbox_max_z(doc, contrib_elem_ids)
    try:
        for col in column_instances or []:
            if col is None:
                continue
            if omit_eid is not None and _element_ids_coinciden(col.Id, omit_eid):
                continue
            for sd in _iter_solids_revit_element(col, geom_opts):
                if _solidos_intersectan_volumen(probe, sd):
                    return True
    except Exception:
        return False
    return False


def fuse_vertical_world_intervals_from_jobs(
    jobs,
    short_curve_tolerance,

    xy_decimals=XY_KEY_DECIMALS_DEFAULT,
):
    """

    Para cada punto de rejilla de cada columna, intervalo vertical en mundo ``[Zb, Ze]``.

    Claves XY redondeadas: se unen todos los intervalos que compartan la misma proyección

    **entre columnas**. Resultado: una línea modelo continua desde ``min(Z)`` a ``max(Z)``

    para ese eje, atravesando físicamente todo el volumen combinado donde haya soporte.


    Salida por hilo fusionado ``(XYZ base, span, contribuyentes, bar_enum_agg)``; si varias contribuciones etiquetan distinto el mismo XY, ``bar_enum_agg`` concatena ordenado con ``|``.

    """

    tol = abs(float(short_curve_tolerance)) + 1e-12

    buckets = {}

    key_order = []

    contrib_by_k = defaultdict(set)

    enum_by_k = defaultdict(set)

    for jb in jobs:
        dz_col = float(jb["height"])
        elem = jb.get("elem") or jb.get("column")
        if elem is None:
            continue
        try:
            eid = elem.Id
        except Exception:
            continue

        for raw in jb["raw_pts"]:
            bar_lbl_piece = ""

            if isinstance(raw, dict):
                xyzp = raw["pt"]
                bar_lbl_piece = str(raw.get("bar_enum") or "").strip()
            else:
                xyzp = raw

            xf = float(xyzp.X)

            yf = float(xyzp.Y)

            zb = float(xyzp.Z)

            ze = zb + dz_col

            k = (
                round(xf, int(xy_decimals)),
                round(yf, int(xy_decimals)),
            )

            contrib_by_k[k].add(eid)

            if bar_lbl_piece:
                enum_by_k[k].add(bar_lbl_piece)



            if k not in buckets:
                buckets[k] = dict(rx=xf, ry=yf, z0=zb, z1=ze)
                key_order.append(k)

            else:
                bk = buckets[k]
                bk["z0"] = min(float(bk["z0"]), zb)
                bk["z1"] = max(float(bk["z1"]), ze)

    fused = []

    for k in sorted(key_order, key=lambda t: (t[0], t[1])):
        bk = buckets[k]

        span = float(bk["z1"]) - float(bk["z0"])

        if span < tol:
            continue

        scheme_tags = set(enum_by_k.get(k, ()))
        agg_enum = ""

        try:
            nt = len(scheme_tags)
        except Exception:
            nt = 0

        if nt == 1:
            agg_enum = next(iter(scheme_tags))


        elif nt > 1:
            agg_enum = "|".join(sorted(scheme_tags))

        contrib_fset = frozenset(contrib_by_k[k])
        fused.append(
            (
                XYZ(float(bk["rx"]), float(bk["ry"]), float(bk["z0"])),
                span,
                contrib_fset,
                agg_enum,
            )
        )


    return fused


_sketch_plane_cache = {}  # noqa: PLC2401 Revit-safe


def _curve_set_bar_scheme_comment(document, curve_element, scheme_label):
    """

    Etiquetas **A**/ **B**/ **IA**/ **IB** en ``Comentarios`` de línea modelo (metadatos de esquema esquinas vs lados).

    """

    if document is None or curve_element is None:
        return
    txt = str(scheme_label).strip()
    if not txt:
        return
    try:
        for pname in ("Comments", "Comentarios"):
            rp = curve_element.LookupParameter(str(pname))
            if rp is not None and not rp.IsReadOnly:
                rp.Set(txt)
                return
        rp_fb = curve_element.get_Parameter(
            BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS
        )
        if rp_fb is not None and not rp_fb.IsReadOnly:
            rp_fb.Set(txt)



    except Exception:

        pass


def create_vertical_model_line(document, pt, span_along_z):
    """

    Crea una **línea modelo** en 3D mediante ``Document.Create.NewModelCurve``:
    el resultado es un ``ModelCurve`` (en el UI suele verse en categoría Lines / Model Lines).


    Tras la **fusión global** entre todas las columnas seleccionadas, cada XY proyecto

    común usa **una** curva modelo con altura Z combinada.

    Returns:
      El ``ModelCurve`` creado, o ``None`` si ``|dz|`` es menor que ``ShortCurveTolerance``.

    """
    dz = float(span_along_z)
    tol = document.Application.ShortCurveTolerance
    if abs(dz) < tol:
        return None

    start_pt = XYZ(pt.X, pt.Y, pt.Z)
    end_pt = XYZ(pt.X, pt.Y, pt.Z + dz)
    curve = Line.CreateBound(start_pt, end_pt)

    key = ("col_layout", round(start_pt.Y, 9))
    sketch_plane = _sketch_plane_cache.get(key)

    if sketch_plane is None or not sketch_plane.IsValidObject:
        plane = Plane.CreateByNormalAndOrigin(XYZ.BasisY, start_pt)
        sketch_plane = SketchPlane.Create(document, plane)
        _sketch_plane_cache[key] = sketch_plane

    return document.Create.NewModelCurve(curve, sketch_plane)


def _rebar_nominal_diameter_mm(bar_type):
    if bar_type is None:
        return None
    try:
        return UnitUtils.ConvertFromInternalUnits(
            float(bar_type.BarNominalDiameter),
            UnitTypeId.Millimeters,
        )
    except Exception:
        return None


def _resolve_rebar_bar_type_by_diameter_mm(document, target_mm):
    """
    ``RebarBarType`` más cercano al Ø nominal pedido. Devuelve
    ``(bar_type, exact_match, delta_mm)``.
    """
    if document is None:
        return None, False, None
    best = None
    best_delta = None
    target = float(target_mm)
    try:
        col = FilteredElementCollector(document).OfClass(RebarBarType)
    except Exception:
        col = []
    for bt in col:
        dmm = _rebar_nominal_diameter_mm(bt)
        if dmm is None:
            continue
        delta = abs(float(dmm) - target)
        if best is None or delta < best_delta:
            best = bt
            best_delta = delta
    if best is None:
        return None, False, None
    return best, (float(best_delta) <= 0.25), float(best_delta)


def _curve_ilist_for_rebar(curves):
    """Construye ``List[Curve]`` (implementa ``IList<Curve>``) para ``CreateFromCurves``."""
    from System.Collections.Generic import List

    lst = List[Curve]()
    for crv in curves:
        lst.Add(crv)
    return lst


def _curve_clr_array_curve_host(curves):
    """``Curve[]`` (.NET también expone ``IList<Curve>``): respaldo por enlace IronPython."""
    n = len(curves)
    clr_arr = System.Array.CreateInstance(Curve, n)
    for i, crv in enumerate(curves):
        clr_arr[i] = crv
    return clr_arr


def _nearest_contrib_column_for_xyz(document, bar_x, bar_y, z_mid, contrib_elem_ids):
    """Host de columna contribuyente más compatible con el punto medio de la barra (rangos por sólidos)."""
    best = None
    best_score = None
    bx = float(bar_x)
    by = float(bar_y)
    bz = float(z_mid)
    tol = 0.0
    try:
        tol = float(document.Application.ShortCurveTolerance) * 10.0
    except Exception:
        tol = 1e-6
    tol_curve = max(float(tol), 1e-12)
    for eid in contrib_elem_ids or []:
        try:
            el = document.GetElement(eid)
        except Exception:
            el = None
        if el is None:
            continue
        rng = _solid_aggregate_vertex_ranges_ft(el)
        if rng is None:
            continue
        min_x, max_x, min_y, max_y, min_z, max_z = rng
        score = None
        try:
            inside_xy = (
                bx >= float(min_x) - tol
                and bx <= float(max_x) + tol
                and by >= float(min_y) - tol
                and by <= float(max_y) + tol
            )
            inside_z = bz >= float(min_z) - tol and bz <= float(max_z) + tol
            p0, p1, _, _ = _column_curve_endpoints_axis_ft(el, tol_curve)
            if p0 is not None and p1 is not None:
                mid = p0 + 0.5 * (p1 - p0)
                cx = float(mid.X)
                cy = float(mid.Y)
            else:
                cx = 0.5 * (float(min_x) + float(max_x))
                cy = 0.5 * (float(min_y) + float(max_y))
            cz = 0.5 * (float(min_z) + float(max_z))
            dx = bx - cx
            dy = by - cy
            dz = bz - cz
            score = dx * dx + dy * dy + dz * dz
            if inside_xy and inside_z:
                score -= 1e9
            elif inside_xy:
                score -= 1e6
        except Exception:
            score = None
        if score is None:
            score = 1e30
        if best is None or score < best_score:
            best = el
            best_score = score
    return best


def _set_rebar_comment_text(document, rebar_elem, txt):
    _curve_set_bar_scheme_comment(document, rebar_elem, txt)


def _arma_len_mm_round_from_internal_ft(length_ft):
    """Longitud en ft internas → mm redondeados (firma estable entre instancias)."""
    try:
        return round(
            float(
                UnitUtils.ConvertFromInternalUnits(
                    float(length_ft), UnitTypeId.Millimeters
                )
            ),
            3,
        )
    except Exception:
        try:
            return round(float(length_ft) * 304.8, 3)
        except Exception:
            return 0.0


def _fingerprint_seg_linea_fierro(span_seg_ft, ok_pat_bot, ok_pat_top, pata_len_ft):
    """
    Huella de **un** tramo de la línea tras troceo: recto (solo Lz) o L con pata en bajo/arriba/ambos.
    Tuplas ordenables y hashables para agrupar igual número y largos efectivos.
    """
    lz = _arma_len_mm_round_from_internal_ft(span_seg_ft)
    lp = _arma_len_mm_round_from_internal_ft(pata_len_ft)
    try:
        bot = bool(ok_pat_bot)
        top = bool(ok_pat_top)
    except Exception:
        bot = False
        top = False
    if bot and top:
        return (3, lz, lp, lp)
    if bot:
        return (1, lz, lp)
    if top:
        return (2, lz, lp)
    return (0, lz)


def _linea_fierro_nombre_alfabetico(indice_cero):
    """0 → 'A', 25 → 'Z', 26 → 'AA' (estilo columna Excel)."""
    try:
        _chr = unichr
    except NameError:
        _chr = chr  # noqa: F821 IronPython tiene ``unichr``; CPyth 3 sólo ``chr``.
    n = int(indice_cero) + 1
    out = u""
    while n > 0:
        n, r = divmod(n - 1, 26)
        out = _chr(65 + r) + out
    return out


def _aplicar_armadura_ubicacion_si_existe(rebar_element, valor_texto):
    """Escribe ``Armadura_Ubicacion`` en el ``Rebar`` si el parámetro existe (mismo criterio que fundación/vigas)."""
    if rebar_element is None or valor_texto is None:
        return
    try:
        txt = unicode(valor_texto)
    except Exception:
        try:
            txt = u"{0}".format(valor_texto)
        except Exception:
            return
    try:
        p = rebar_element.LookupParameter(COLUMN_ARMA_UBICACION_PARAM)
        if p is None or p.IsReadOnly:
            return
        p.Set(txt)
    except Exception:
        pass


def _apply_linea_fierro_armadura_ubicacion(assignments_list):
    """
    ``assignments_list``: ``[(rebar, key), ...]`` donde ``key`` discrimina líneas de fierro
    (mismo ``key`` → misma letra A,B,…). Se asigna antes de ``Commit`` dentro de la misma transacción.
    """
    if not assignments_list:
        return
    nombre_por_key = _linea_fierro_label_map_from_assignments(assignments_list)
    if not nombre_por_key:
        return
    for rb, k in assignments_list:
        try:
            nm = nombre_por_key.get(k)
            if nm is None:
                continue
            _aplicar_armadura_ubicacion_si_existe(rb, nm)
        except Exception:
            continue


def _linea_fierro_label_map_from_assignments(assignments_list):
    """``key`` → etiqueta texto.

    ``k`` es ``(n_segmentos, tupla_de_huellas)``. Orden de letras A, B, C…:
    primero por **más segmentos** (descendente); empate por orden lexicográfico de
    la tupla de huellas (igual que antes entre grupos con el mismo ``n``).
    """
    if not assignments_list:
        return {}
    try:
        uniq = sorted(
            {k for _rb, k in assignments_list},
            key=lambda kk: (-kk[0], kk[1]),
        )
    except Exception:
        return {}
    return {k: _linea_fierro_nombre_alfabetico(i) for i, k in enumerate(uniq)}




def _shape_display_name_normalized_column(value):
    if value is None:
        return u""
    try:
        t = unicode(value)
    except Exception:
        try:
            t = System.Convert.ToString(value)
        except Exception:
            return u""
    try:
        return t.replace(u"\u00A0", u" ").strip()
    except Exception:
        return u""


def _rebar_shape_visible_label_column(shape):
    if shape is None:
        return u""
    for bip in (
        BuiltInParameter.SYMBOL_NAME_PARAM,
        BuiltInParameter.ALL_MODEL_TYPE_NAME,
    ):
        try:
            p = shape.get_Parameter(bip)
            if p is not None and p.HasValue:
                s = _shape_display_name_normalized_column(p.AsString())
                if s:
                    return s
        except Exception:
            continue
    try:
        return _shape_display_name_normalized_column(getattr(shape, "Name", None))
    except Exception:
        return u""


def _document_rebar_shape_by_visible_name(document, nombre_visible):
    """
    ``RebarShape`` del documento por nombre en UI — misma política flexible que geometría BIMTools:
    igualdad exacta, sin mayúsculas, o coincidencia por dígitos (p. ej. «02»).
    """
    if document is None or not nombre_visible:
        return None
    key = _shape_display_name_normalized_column(nombre_visible)
    if not key:
        return None
    try:
        key_lower = key.lower()
    except Exception:
        key_lower = key
    key_digits = u"".join(ch for ch in key if ch in u"0123456789")
    candidates = []
    try:
        for sh in FilteredElementCollector(document).OfClass(RebarShape):
            try:
                sn = _rebar_shape_visible_label_column(sh)
                if not sn:
                    continue
                try:
                    sn_low = sn.lower()
                except Exception:
                    sn_low = sn
                dig = u"".join(ch for ch in sn if ch in u"0123456789")
                candidates.append((sh, sn, sn_low, dig))
            except Exception:
                continue
    except Exception:
        return None
    for sh, sn, _, _dig in candidates:
        if sn == key:
            return sh
    for sh, _, sn_low, _dig in candidates:
        if sn_low == key_lower:
            return sh
    for sh, sn, _, dig in candidates:
        if dig and dig == key:
            return sh
    for sh, _sn, _sn_low, dig in candidates:
        if key_digits and dig == key_digits:
            return sh
    return None


def _try_create_column_rebar_from_curves_project_shape_no_hooks(
    document, host, bar_type, curves, normal_vec, shape_display_name
):
    """``CreateFromCurvesAndShape`` con forma de catálogo (sin definir nueva por barra)."""
    if (
        document is None
        or host is None
        or bar_type is None
        or not curves
        or shape_display_name is None
        or shape_display_name == u""
    ):
        return None
    shape = _document_rebar_shape_by_visible_name(document, shape_display_name)
    if shape is None:
        return None
    try:
        cl = _curve_ilist_for_rebar(curves)
    except Exception:
        return None
    orient_pairs = (
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
        (RebarHookOrientation.Right, RebarHookOrientation.Right),
        (RebarHookOrientation.Left, RebarHookOrientation.Left),
    )
    norms = [normal_vec]
    try:
        norms.append(normal_vec.Negate())
    except Exception:
        pass
    invalid = ElementId.InvalidElementId
    for nvec in norms:
        if nvec is None:
            continue
        for so, eo in orient_pairs:
            rb = None
            try:
                rb = Rebar.CreateFromCurvesAndShape(
                    document,
                    shape,
                    bar_type,
                    None,
                    None,
                    host,
                    nvec,
                    cl,
                    so,
                    eo,
                    0.0,
                    0.0,
                    invalid,
                    invalid,
                )
            except Exception:
                rb = None
            if rb is None:
                try:
                    rb = Rebar.CreateFromCurvesAndShape(
                        document,
                        shape,
                        bar_type,
                        None,
                        None,
                        host,
                        nvec,
                        cl,
                        so,
                        eo,
                    )
                except Exception:
                    rb = None
            if rb is not None:
                try:
                    rb.Style = RebarStyle.Standard
                except Exception:
                    pass
                return rb
    return None


def _create_rebar_from_curves_no_hooks(document, host, bar_type, curves, normal_vec):
    """
    Una sola llamada por barra: las ``Curve`` van juntas en ``List[Curve]``
    (.NET ``IList<Curve>``) hacia ``Rebar.CreateFromCurves``.

    Con polilínea (pata + vertical), se prueba primero **sin** reutilizar forma de catálogo
    (`useExistingShapeIfPossible=False`), porque ``True`` a veces simplifica la polilínea
    y separa mal el dibujo respecto de la geometría de varias ``Curve``.
    """
    if document is None or host is None or bar_type is None or not curves:
        return None
    try:
        arr_list = _curve_ilist_for_rebar(curves)
    except Exception:
        return None
    n_cv = len(curves)

    curve_inputs = [(arr_list, "List[Curve]")]
    if n_cv > 1:
        try:
            curve_inputs.append(
                (_curve_clr_array_curve_host(curves), "Curve[]")
            )
        except Exception:
            pass

    if n_cv > 1:
        use_create_pairs = (
            (False, False),
            (False, True),
            (True, False),
            (True, True),
        )
    else:
        use_create_pairs = (
            (True, True),
            (True, False),
            (False, True),
            (False, False),
        )
    norms = [normal_vec]
    try:
        norms.append(normal_vec.Negate())
    except Exception:
        pass
    orient_pairs = (
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
        (RebarHookOrientation.Right, RebarHookOrientation.Right),
        (RebarHookOrientation.Left, RebarHookOrientation.Left),
    )
    for curve_arg, _ctype_lbl in curve_inputs:
        for use_existing, create_new in use_create_pairs:
            for nvec in norms:
                if nvec is None:
                    continue
                for so, eo in orient_pairs:
                    try:
                        rb = Rebar.CreateFromCurves(
                            document,
                            RebarStyle.Standard,
                            bar_type,
                            None,
                            None,
                            host,
                            nvec,
                            curve_arg,
                            so,
                            eo,
                            use_existing,
                            create_new,
                        )
                        if rb is not None:
                            return rb
                    except Exception:
                        continue

    return None


def create_vertical_rebar(document, pt, span_along_z, host, bar_type, comment_text=None):
    """Un tramo solo vertical sin patas."""
    dz = float(span_along_z)
    tol = document.Application.ShortCurveTolerance
    if abs(dz) < tol:
        return None
    start_pt = XYZ(pt.X, pt.Y, pt.Z)
    end_pt = XYZ(pt.X, pt.Y, pt.Z + dz)
    try:
        curve = Line.CreateBound(start_pt, end_pt)
    except Exception:
        return None
    rb = _create_rebar_from_curves_no_hooks(
        document,
        host,
        bar_type,
        [curve],
        XYZ.BasisX,
    )
    if rb is not None and comment_text:
        _set_rebar_comment_text(document, rb, comment_text)
    return rb


def _pata_horizontal_leg_at_z(document, bx, by, z_level, contrib_elem_ids, pata_len_ft):
    """Pata en planta: desde la rejilla ``(bx,by)`` hacia el centro de columna contribuyente.
    Devuelve ``(pa,pb,ux,uy)`` con ``pa`` en el pie del eje principal y ``pb`` la punta horizontal.
    ``None`` si no aplica."""
    bx = float(bx)
    by = float(by)
    zr = float(z_level)
    lf = float(pata_len_ft)
    if lf <= 1e-12:
        return None
    nxy = _nearest_contrib_column_plan_center_xy(
        document,
        bx,
        by,
        contrib_elem_ids,
    )
    if nxy is None:
        return None
    cx, cy = nxy
    dx = float(cx) - bx
    dy = float(cy) - by
    hpl = math.hypot(dx, dy)
    if hpl <= 1e-9:
        return None
    ux = dx / hpl
    uy = dy / hpl
    pa = XYZ(bx, by, zr)
    pb = XYZ(bx + ux * lf, by + uy * lf, zr)
    return (pa, pb, ux, uy)


def _rebar_plane_normal_vertical_with_xy_leg(ux, uy):
    """Normal al plano que contiene un tramo horizontal en XY y un tramo paralelo a Z."""
    try:
        n = XYZ(float(uy), -float(ux), 0.0)
        lm = float(n.GetLength())
        if lm < 1e-12:
            return None
        return n.Normalize()
    except Exception:
        return None


def create_longitudinal_rebar_with_optional_patas(
    document,
    bx,
    by,
    zs,
    span_seg_z,
    host,
    bar_type,
    comment_text,
    want_bottom_pata,
    want_top_pata,
    pata_len_ft,
    contrib_ids,
):
    """
    Un solo ``Rebar``: tramo vertical opcionalmente con pata inferior y/o superior
    en la **misma** cadena de curvas.

    Con **solo una pata** (dos segmentos: tronco largo primero y pata después)
    intenta ``Rebar.CreateFromCurvesAndShape`` con la forma «02» ya definida
    en el proyecto; si no aplica o falla el API, usa ``CreateFromCurves``.
    Con **dos patas** tres segmentos, sólo ``CreateFromCurves``.
    """
    bx = float(bx)
    by = float(by)
    zs = float(zs)
    span = float(span_seg_z)
    tol = abs(float(document.Application.ShortCurveTolerance))
    if span <= tol:
        return None, False, False

    z_top = zs + span
    foot = XYZ(bx, by, zs)
    head = XYZ(bx, by, z_top)

    bot_leg = None
    top_leg = None
    if want_bottom_pata:
        bot_leg = _pata_horizontal_leg_at_z(
            document,
            bx,
            by,
            zs,
            contrib_ids,
            pata_len_ft,
        )
    if want_top_pata:
        top_leg = _pata_horizontal_leg_at_z(
            document,
            bx,
            by,
            z_top,
            contrib_ids,
            pata_len_ft,
        )

    curves = []
    ux_plane = None
    uy_plane = None

    if bot_leg and top_leg:
        pa_b, pb_b, ux_b, uy_b = bot_leg
        pa_t, pb_t, ux_t, uy_t = top_leg
        ux_plane = ux_b
        uy_plane = uy_b
        try:
            curves.append(Line.CreateBound(pb_b, pa_b))
            curves.append(Line.CreateBound(pa_b, head))
            curves.append(Line.CreateBound(pa_t, pb_t))
        except Exception:
            return None, False, False
    elif bot_leg:
        pa_b, pb_b, ux_b, uy_b = bot_leg
        ux_plane = ux_b
        uy_plane = uy_b
        try:
            curves.append(Line.CreateBound(head, pa_b))
            curves.append(Line.CreateBound(pa_b, pb_b))
            foot = pa_b
        except Exception:
            bot_leg = None
            curves = []
    elif top_leg:
        pa_t, pb_t, ux_t, uy_t = top_leg
        ux_plane = ux_t
        uy_plane = uy_t
        try:
            curves.append(Line.CreateBound(foot, head))
            curves.append(Line.CreateBound(pa_t, pb_t))
        except Exception:
            top_leg = None
            curves = []

    if not curves:
        foot = XYZ(bx, by, zs)
        try:
            curves.append(Line.CreateBound(foot, head))
        except Exception:
            return None, False, False

    nvec = XYZ.BasisX
    if ux_plane is not None and uy_plane is not None:
        n_pl = _rebar_plane_normal_vertical_with_xy_leg(ux_plane, uy_plane)
        if n_pl is not None:
            nvec = n_pl

    rb = None
    if (
        len(curves) == 2
        and ((bot_leg is not None) ^ (top_leg is not None))
    ):
        rb = _try_create_column_rebar_from_curves_project_shape_no_hooks(
            document,
            host,
            bar_type,
            curves,
            nvec,
            COLUMN_REBAR_L_SHAPE_DISPLAY_NAME,
        )
    if rb is None:
        rb = _create_rebar_from_curves_no_hooks(
            document, host, bar_type, curves, nvec
        )
    if rb is not None and comment_text:
        _set_rebar_comment_text(document, rb, comment_text)

    did_bot = bool(bot_leg is not None and want_bottom_pata)
    did_top = bool(top_leg is not None and want_top_pata)
    return rb, did_bot, did_top


def _column_solid_vertex_z_min_max(col):
    rng = _solid_aggregate_vertex_ranges_ft(col)
    if rng is None:
        return None, None
    try:
        return float(rng[4]), float(rng[5])
    except Exception:
        return None, None


def _horizontal_plane_z_snap_to_bbox(pl, col, use_top):
    """
    Para planos **horizontales** (corte en Z constante), alinea ``Origin.Z`` con el borde
    inferior o superior del volumen modelizado (**vértices de sólidos**), misma referencia
    que ``get_column_dimensions``. ``Location.Curve`` puede no coincidir con el sólido.
    """
    if pl is None or col is None:
        return pl
    try:
        n = pl.Normal
        nx = float(n.X)
        ny = float(n.Y)
        nz = float(n.Z)
        if abs(nz) < 0.99:
            return pl
        if (nx * nx + ny * ny) > 0.02:
            return pl
        zm, zM = _column_solid_vertex_z_min_max(col)
        if zm is None or zM is None:
            return pl
        zt = float(zM if use_top else zm)
        o = pl.Origin
        return Plane.CreateByNormalAndOrigin(
            XYZ.BasisZ,
            XYZ(float(o.X), float(o.Y), zt),
        )
    except Exception:
        return pl


def _plane_from_column_location_curve(col, short_curve_tolerance_ft):
    """
    Plano **perpendicular al eje** ``Location.Curve`` de la columna, que pasa por
    ``GetEndPoint(0)``: ``normal = normalize(p1 - p0)`` (corte típico horizontal si el pilar es vertical).

    Si la familia usa ``LocationPoint`` (sin curva), se asume pilar vertical: plano horizontal
    por ese punto con ``normal = Z``.

    Si la curva existe pero es más corta que la tolerancia, se usa el mismo criterio que un punto base.

    Planos horizontales: la cota Z del plano se alinea al **borde inferior del sólido** del pilar,
    igual que la rejilla de barras (no al ``Location`` puro, que puede no coincidir con el sólido).
    """
    tol = abs(float(short_curve_tolerance_ft))
    if tol < 1e-12:
        tol = 1e-12
    if col is None:
        return None
    pl = None
    try:
        loc = col.Location
        cr = getattr(loc, "Curve", None)
        if cr is not None:
            p0 = cr.GetEndPoint(0)
            p1 = cr.GetEndPoint(1)
            v = p1 - p0
            if v.GetLength() >= tol:
                n = v.Normalize()
                pl = Plane.CreateByNormalAndOrigin(n, p0)
            else:
                pl = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, p0)
        else:
            pt_loc = getattr(loc, "Point", None)
            if pt_loc is not None:
                pl = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, pt_loc)
    except Exception:
        return None
    if pl is None:
        return None
    return _horizontal_plane_z_snap_to_bbox(pl, col, use_top=False)


def _plane_from_column_location_curve_at_end(col, short_curve_tolerance_ft):
    """
    Igual criterio de normal que ``_plane_from_column_location_curve``, pero el plano pasa por
    el **extremo superior** de la curva (``GetEndPoint(1)``). Con ``LocationPoint``, cota Z =
    máximo ``Z`` de vértices del sólido (alineada a la rejilla).

    Planos horizontales: ``Origin.Z`` se ajusta al máximo ``Z`` del sólido (igual criterio que la base).

    Sirve como segundo plano distinto cuando varias columnas comparten arranque en planta y el primer plano
    horizontal quedaría duplicado en Z.
    """
    tol = abs(float(short_curve_tolerance_ft))
    if tol < 1e-12:
        tol = 1e-12
    if col is None:
        return None
    pl = None
    try:
        loc = col.Location
        cr = getattr(loc, "Curve", None)
        if cr is not None:
            p0 = cr.GetEndPoint(0)
            p1 = cr.GetEndPoint(1)
            v = p1 - p0
            if v.GetLength() >= tol:
                n = v.Normalize()
                pl = Plane.CreateByNormalAndOrigin(n, p1)
            else:
                pl = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, p1)
        else:
            pt_loc = getattr(loc, "Point", None)
            if pt_loc is not None:
                _, zM = _column_solid_vertex_z_min_max(col)
                if zM is None:
                    return None
                pl = Plane.CreateByNormalAndOrigin(
                    XYZ.BasisZ,
                    XYZ(float(pt_loc.X), float(pt_loc.Y), zM),
                )
    except Exception:
        return None
    if pl is None:
        return None
    return _horizontal_plane_z_snap_to_bbox(pl, col, use_top=True)


def _z_cut_vertical_bar_xy_plane(bx, by, z_lo, z_hi, plane, tol_ft):
    """
    Hilo vertical en ``(bx,by)`` entre ``z_lo`` y ``z_hi``; intersección con un ``Plane`` [Revit].

    Si el plano es casi vertical (``|n_z|`` pequeño) y la barra no cae en el plano, no hay corte.
    Si la barra yace en un plano vertical compartido, no se devuelve corte puntual.
    """
    if plane is None:
        return None
    try:
        n = plane.Normal
        ox = float(plane.Origin.X)
        oy = float(plane.Origin.Y)
        oz = float(plane.Origin.Z)
        nx = float(n.X)
        ny = float(n.Y)
        nz = float(n.Z)
    except Exception:
        return None
    tt = abs(float(tol_ft))
    if tt < 1e-12:
        tt = 1e-12
    bx = float(bx)
    by = float(by)
    z_lo = float(z_lo)
    z_hi = float(z_hi)
    if z_hi <= z_lo + tt:
        return None
    # Plano casi vertical: sin solución única en Z para hilo paralelo a Z.
    if abs(nz) < tt * 50.0:
        d_xy = nx * (bx - ox) + ny * (by - oy)
        if abs(d_xy) < tt * 200.0:
            return None
        return None
    z_int = oz - (nx * (bx - ox) + ny * (by - oy)) / nz
    if z_int <= z_lo + tt or z_int >= z_hi - tt:
        return None
    return z_int


def _split_z_span_by_planes(bx, by, z_lo, z_hi, planes, tol_ft):
    """
    Lista ``(z_start, dz)`` positivos; sin planos o sin cortes válidos devuelve un solo tramo.
    """
    z_lo = float(z_lo)
    z_hi = float(z_hi)
    tt = abs(float(tol_ft))
    if tt < 1e-12:
        tt = 1e-12
    merge_eps = max(tt * 4.0, tt)
    if z_hi <= z_lo + tt:
        return []
    if not planes:
        return [(z_lo, z_hi - z_lo)]
    cuts = [z_lo, z_hi]
    for pl in planes:
        zi = _z_cut_vertical_bar_xy_plane(bx, by, z_lo, z_hi, pl, tt)
        if zi is None:
            continue
        cuts.append(float(zi))
    cuts.sort()
    merged = []
    for c in cuts:
        if not merged or c > merged[-1] + merge_eps:
            merged.append(c)
    if len(merged) < 2:
        return [(z_lo, z_hi - z_lo)]
    out = []
    for i in range(len(merged) - 1):
        a = merged[i]
        b = merged[i + 1]
        dz = b - a
        if dz > tt * 4.0:
            out.append((a, dz))
    if not out:
        return [(z_lo, z_hi - z_lo)]
    return out


def build_column_cut_planes_from_elements(columns, tol_ft):
    """
    Un ``Plane`` por columna seleccionada (política **troceo A / IA** exterior e interior).

    Origen habitual: ``LocationCurve.GetEndPoint(0)``, normal = eje; en horizontales se alinea ``Z`` al
    borde inferior del sólido (**start** de pilar para la cota efectiva del corte). Si dos columnas repetían esa Z
    (~2 mm), la segunda puede usar borde **superior** del sólido para un segundo nivel de corte.
    """
    planes = []
    tol = abs(float(tol_ft))
    if tol < 1e-12:
        tol = 1e-12
    try:
        z_dup_ft = UnitUtils.ConvertToInternalUnits(2.0, UnitTypeId.Millimeters)
    except Exception:
        z_dup_ft = tol * 500.0
    used_z = []
    for col in columns or []:
        pl = _plane_from_column_location_curve(col, tol)
        if pl is None:
            continue
        try:
            n = pl.Normal
            oz = float(pl.Origin.Z)
            nx = float(n.X)
            ny = float(n.Y)
            nz = float(n.Z)
            is_horiz = abs(nz) >= 0.99 and (nx * nx + ny * ny) < 0.02
            if is_horiz:
                dup = any(abs(oz - uz) < z_dup_ft for uz in used_z)
                if dup:
                    pl2 = _plane_from_column_location_curve_at_end(col, tol)
                    if pl2 is not None:
                        try:
                            oz2 = float(pl2.Origin.Z)
                        except Exception:
                            oz2 = oz
                        if not any(abs(oz2 - uz) < z_dup_ft for uz in used_z):
                            pl = pl2
                            oz = oz2
                used_z.append(oz)
            else:
                used_z.append(oz)
        except Exception:
            pass
        planes.append(pl)
    return planes


def _column_location_start_end_unit_vector(col, short_curve_tolerance_ft):
    """Vector unitario del eje ``Location`` (**end − start**); sin curva válida → ``BasisZ``."""
    tol = abs(float(short_curve_tolerance_ft))
    if tol < 1e-12:
        tol = 1e-12
    if col is None:
        return None
    try:
        loc = col.Location
        cr = getattr(loc, "Curve", None)
        if cr is not None:
            p0 = cr.GetEndPoint(0)
            p1 = cr.GetEndPoint(1)
            v = p1 - p0
            if v.GetLength() >= tol:
                return v.Normalize()
            return XYZ.BasisZ
        if getattr(loc, "Point", None) is not None:
            return XYZ.BasisZ
    except Exception:
        pass
    return None


def _offset_cut_plane_along_column_axis(plane, col, offset_internal_ft, tol_ft):
    """
    Traslada el origen del plano de corte ``offset_internal_ft`` en la dirección del eje
    de la columna (start → end). Plano horizontal tras snap sigue siendo horizontal.
    """
    if plane is None or col is None:
        return plane
    if abs(float(offset_internal_ft)) <= 1e-12:
        return plane
    v_hat = _column_location_start_end_unit_vector(col, tol_ft)
    if v_hat is None:
        return plane
    off = float(offset_internal_ft)
    try:
        ox = float(plane.Origin.X) + off * float(v_hat.X)
        oy = float(plane.Origin.Y) + off * float(v_hat.Y)
        oz = float(plane.Origin.Z) + off * float(v_hat.Z)
        n = plane.Normal
        nx = float(n.X)
        ny = float(n.Y)
        nz = float(n.Z)
        is_horiz = abs(nz) >= 0.99 and (nx * nx + ny * ny) < 0.02
        if is_horiz:
            return Plane.CreateByNormalAndOrigin(
                XYZ.BasisZ,
                XYZ(ox, oy, oz),
            )
        return Plane.CreateByNormalAndOrigin(n, XYZ(ox, oy, oz))
    except Exception:
        return plane


def build_b_bar_cut_planes_from_elements(columns, tol_ft, offset_internal_ft):
    """
    **Troceo B / IB** (anillo exterior ``B`` y anillo interior ``IB``): misma secuencia de planos que
    ``build_column_cut_planes_from_elements`` (punto inicial de eje + cota Z al borde del sólido), pero cada plano
    se mueve ``offset_internal_ft`` en dirección columna ``start→end`` — en layout columnas ese valor es
    ``L(Ø)`` de empalme** (tabulado), para que el corte de barras tipo B no coincida con el de A/IA.

    Si ``offset_internal_ft`` es ~0 el resultado coincide con los planos A.
    """
    tol = abs(float(tol_ft))
    if tol < 1e-12:
        tol = 1e-12
    if abs(float(offset_internal_ft)) <= 1e-12:
        return build_column_cut_planes_from_elements(columns, tol_ft)
    planes = []
    try:
        z_dup_ft = UnitUtils.ConvertToInternalUnits(2.0, UnitTypeId.Millimeters)
    except Exception:
        z_dup_ft = tol * 500.0
    used_z = []
    for col in columns or []:
        pl = _plane_from_column_location_curve(col, tol)
        if pl is None:
            continue
        try:
            n = pl.Normal
            oz = float(pl.Origin.Z)
            nx = float(n.X)
            ny = float(n.Y)
            nz = float(n.Z)
            is_horiz = abs(nz) >= 0.99 and (nx * nx + ny * ny) < 0.02
            if is_horiz:
                dup = any(abs(oz - uz) < z_dup_ft for uz in used_z)
                if dup:
                    pl2 = _plane_from_column_location_curve_at_end(col, tol)
                    if pl2 is not None:
                        try:
                            oz2 = float(pl2.Origin.Z)
                        except Exception:
                            oz2 = oz
                        if not any(abs(oz2 - uz) < z_dup_ft for uz in used_z):
                            pl = pl2
                            oz = oz2
                used_z.append(oz)
            else:
                used_z.append(oz)
        except Exception:
            pass
        pl = _offset_cut_plane_along_column_axis(
            pl,
            col,
            float(offset_internal_ft),
            tol,
        )
        planes.append(pl)
    return planes


def _internal_xyz_to_mm_comment_str(xyz, nd=1):
    """``x,y,z`` en mm (proyecto) para texto de comentario."""
    try:
        xm = UnitUtils.ConvertFromInternalUnits(float(xyz.X), UnitTypeId.Millimeters)
        ym = UnitUtils.ConvertFromInternalUnits(float(xyz.Y), UnitTypeId.Millimeters)
        zm = UnitUtils.ConvertFromInternalUnits(float(xyz.Z), UnitTypeId.Millimeters)
        return u"{0:.{1}f};{2:.{1}f};{3:.{1}f}".format(xm, int(nd), ym, zm)
    except Exception:
        return u"?"


def _model_curve_set_comment_text(document, model_curve, text):
    """Escribe en Comentarios / ``ALL_MODEL_INSTANCE_COMMENTS`` (misma convención que esquema A/B)."""
    if document is None or model_curve is None:
        return
    txt = str(text).strip()
    if not txt:
        return
    try:
        for pname in ("Comments", "Comentarios"):
            rp = model_curve.LookupParameter(str(pname))
            if rp is not None and not rp.IsReadOnly:
                rp.Set(txt)
                return
        rp_fb = model_curve.get_Parameter(
            BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS
        )
        if rp_fb is not None and not rp_fb.IsReadOnly:
            rp_fb.Set(txt)
    except Exception:
        pass


def _resolve_cut_plane_linestyle_id(document):
    """``GraphicsStyle`` subcategoría Líneas cuyo nombre contiene un hint (p. ej. ``ARAINCOCUTPLANE``)."""
    tup = sorted(
        _scheme_verify_collect_lines_graphics_styles(document),
        key=lambda t: str(t[0]),
    )
    for hint in _CUT_PLANE_HINT_LINESTYLE:
        hh = str(hint).upper()
        for nm, gid in tup:
            if hh in str(nm or "").upper():
                return gid
    return None


def _create_model_curve_bound_xyz(document, p0, p1, linestyle_id=None):
    """``ModelCurve`` arbitrario en 3D (plano esquicio contiene el segmento ``p0``–``p1``)."""
    if document is None or p0 is None or p1 is None:
        return None
    tol = document.Application.ShortCurveTolerance
    try:
        v = p1 - p0
        ln = float(v.GetLength())
    except Exception:
        return None
    if ln < float(tol) * 2.0:
        return None
    try:
        line = Line.CreateBound(p0, p1)
        d = v.Normalize()
    except Exception:
        return None
    try:
        aux = XYZ.BasisZ
        if abs(float(d.DotProduct(aux))) > 0.99:
            aux = XYZ.BasisX
        ap = d.CrossProduct(aux)
        if ap.GetLength() < float(tol) * 2.0:
            ap = d.CrossProduct(XYZ.BasisY)
        if ap.GetLength() < float(tol) * 2.0:
            return None
        plane_normal = ap.Normalize()
        pln = Plane.CreateByNormalAndOrigin(plane_normal, p0)
        sk = SketchPlane.Create(document, pln)
        mc = document.Create.NewModelCurve(line, sk)
        _scheme_verify_apply_model_line_linestyle(mc, linestyle_id)
        return mc
    except Exception:
        return None


def _draw_cut_plane_markers_for_columns(
    document,
    columns,
    cut_planes,
    curve_tol,
    normal_len_ft,
    inplane_leg_ft,
    linestyle_id,
):
    """
    Por cada columna de referencia: línea a lo largo de la **normal** del plano (sentido corte)
    y línea corta **en el plano** (referencia visual). Comentarios con origen en mm y vector ``n`` unitario.

    ``cut_planes`` debe ser la misma lista que la usada para trocear (p. ej. salida de
    ``build_column_cut_planes_from_elements``), para que el marcado coincida con el corte real.
    """
    n_drawn = 0
    tt = abs(float(curve_tol))
    if tt < 1e-12:
        tt = 1e-12
    nlf = abs(float(normal_len_ft))
    ilf = abs(float(inplane_leg_ft))
    cols = columns or []
    cpl = cut_planes or []
    for i in range(min(len(cols), len(cpl))):
        col = cols[i]
        pl = cpl[i]
        if pl is None:
            continue
        try:
            cid = int(col.Id.IntegerValue)
        except Exception:
            cid = 0
        try:
            O = pl.Origin
            N = pl.Normal
            nx = float(N.X)
            ny = float(N.Y)
            nz = float(N.Z)
            nm = math.sqrt(nx * nx + ny * ny + nz * nz)
            if nm < 1e-12:
                continue
            nx /= nm
            ny /= nm
            nz /= nm
        except Exception:
            continue
        nvec = XYZ(nx, ny, nz)
        p_n = XYZ(
            float(O.X) + nx * nlf,
            float(O.Y) + ny * nlf,
            float(O.Z) + nz * nlf,
        )
        omm = _internal_xyz_to_mm_comment_str(O)
        c_norm = (
            u"Arainco: plano corte | col.id={} | O(mm)={} | n=({:.4f},{:.4f},{:.4f}) [segmento=normal]"
        ).format(int(cid), omm, nx, ny, nz)
        mc_n = _create_model_curve_bound_xyz(document, O, p_n, linestyle_id)
        if mc_n is not None:
            _model_curve_set_comment_text(document, mc_n, c_norm)
            n_drawn += 1
        taux = nvec.CrossProduct(XYZ.BasisX)
        if taux.GetLength() < tt * 20.0:
            taux = nvec.CrossProduct(XYZ.BasisY)
        if taux.GetLength() < tt * 20.0:
            taux = nvec.CrossProduct(XYZ.BasisZ)
        if taux.GetLength() < tt * 20.0:
            continue
        try:
            tuv = taux.Normalize()
        except Exception:
            continue
        p_t = XYZ(
            float(O.X) + float(tuv.X) * ilf,
            float(O.Y) + float(tuv.Y) * ilf,
            float(O.Z) + float(tuv.Z) * ilf,
        )
        c_tan = (
            u"Arainco: plano corte | col.id={} | trazo en plano ⟂ normal (referencia)"
        ).format(int(cid))
        mc_t = _create_model_curve_bound_xyz(document, O, p_t, linestyle_id)
        if mc_t is not None:
            _model_curve_set_comment_text(document, mc_t, c_tan)
            n_drawn += 1
    return n_drawn


def _nearest_contrib_column_plan_center_xy(doc, bar_x, bar_y, contrib_elem_ids):
    """(cx, cy) en planta del centro geométrico de referencia de la columna contribuyente más cercana."""
    best_cx = None
    best_cy = None
    best_d2 = None
    bx = float(bar_x)
    by = float(bar_y)
    for eid in contrib_elem_ids or []:
        try:
            el = doc.GetElement(eid)
        except Exception:
            el = None
        if el is None:
            continue
        try:
            _, _, _, cen, _, _ = get_column_dimensions(el)
        except Exception:
            continue
        try:
            cx = float(cen.X)
            cy = float(cen.Y)
        except Exception:
            continue
        ddx = cx - bx
        ddy = cy - by
        d2 = ddx * ddx + ddy * ddy
        if best_d2 is None or d2 < best_d2:
            best_d2 = d2
            best_cx = cx
            best_cy = cy
    if best_cx is None:
        return None
    return best_cx, best_cy


def _scheme_verify_apply_model_line_linestyle(curve_element, linestyle_element_id):
    if curve_element is None or linestyle_element_id is None:
        return
    try:
        if linestyle_element_id == ElementId.InvalidElementId:
            return
        curve_element.LineStyleId = linestyle_element_id
    except Exception:
        pass


def _scheme_verify_collect_lines_graphics_styles(document):
    """``(nombre subcategoría UPPER, GraphicStyle.ElementId)`` proyectados bajo Lines."""
    out = []
    try:
        lc = Category.GetCategory(document, BuiltInCategory.OST_Lines)
    except Exception:
        lc = None
    if lc is None:
        try:
            lc = document.Settings.Categories.get_Item(BuiltInCategory.OST_Lines)
        except Exception:
            lc = None
    if lc is None:
        return out
    p_iv = lc.Id.IntegerValue
    for gs in FilteredElementCollector(document).OfClass(GraphicsStyle):
        try:
            cg = getattr(gs, "GraphicsStyleCategory", None)
            if cg is None:
                cg = getattr(gs, "Category", None)
            if cg is None:
                continue
            pc = getattr(cg, "Parent", None)
            if pc is None:
                continue
            try:
                if pc.Id.IntegerValue != p_iv:
                    continue
            except Exception:
                continue
            nm_u = ""
            try:
                nm_u = str(getattr(cg, "Name", "") or "").upper()
            except Exception:
                nm_u = ""
            out.append((nm_u, gs.Id))
        except Exception:
            continue
    return out


def _scheme_verify_resolve_pair_linestyle_ids(doc):
    tup = sorted(
        _scheme_verify_collect_lines_graphics_styles(doc),
        key=lambda t: str(t[0]),
    )
    uniq = []
    ivdone = set()
    for nm, gid in tup:
        try:
            iv = int(gid.IntegerValue)
        except Exception:
            continue
        if iv in ivdone:
            continue
        ivdone.add(iv)
        uniq.append((nm, gid))
    if len(uniq) < 2:
        return None, None

    def grab(hints):
        for hint in hints:
            hh = str(hint).upper()
            for nm, gid in uniq:
                if hh in nm:
                    return gid
        return None

    cid = grab(_SCHEME_VERIFY_HINT_CORNER_LINES)
    eid = grab(_SCHEME_VERIFY_HINT_EDGE_LINES)
    if cid is None:
        cid = uniq[0][1]
    if eid is None:
        try:
            eid = uniq[1][1]
        except Exception:
            eid = None
    try:
        if (
            cid is not None
            and eid is not None
            and cid.IntegerValue == eid.IntegerValue
            and len(uniq) > 1
        ):
            eid = uniq[1][1]
    except Exception:
        pass
    return cid, eid


def create_horizontal_xy_model_line(document, p0, p1, linestyle_id=None):

    """

    ``ModelCurve`` horizontal (``Z`` de ``p0``). ``linestyle_id`` opcional (**GraphicsStyle** subcategoría Líneas).

    """
    if p0 is None or p1 is None:
        return None
    tol = document.Application.ShortCurveTolerance
    try:
        v = p1 - p0
        if float(v.GetLength()) < float(tol):
            return None
    except Exception:
        return None
    try:
        curve = Line.CreateBound(p0, p1)
        plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, p0)
        sketch_plane = SketchPlane.Create(document, plane)
        mc = document.Create.NewModelCurve(curve, sketch_plane)
        _scheme_verify_apply_model_line_linestyle(mc, linestyle_id)
        return mc
    except Exception:
        return None


def _scheme_verify_draw_scheme_marker_horizontal(
    document,
    bx,
    by,
    zm_ft,
    half_len_ft,
    curve_tol_ft,
    linestyle_corner_id,
    linestyle_edge_id,
    scheme_label,
):
    """
    Dos **familias visuales** en planta sobre ``zm_ft``:

    - **A / IA**: trazo en **±X**, estilo hint esquinas;
    - **B / IB**: trazo en **±Y**, estilo hint lados.


    Sin fusiones ambiguas (etiquetas con ``|`` no dibujan).
    """

    tg = str(scheme_label or "").strip().upper()
    if not tg or "|" in tg:
        return False

    xf = float(bx)
    yf = float(by)
    zf = float(zm_ft)
    hf = max(float(half_len_ft), float(curve_tol_ft) * 4.0)

    if tg in ("A", "IA"):
        linestyle = linestyle_corner_id
        pa = XYZ(xf - hf, yf, zf)
        pb = XYZ(xf + hf, yf, zf)
    elif tg in ("B", "IB"):
        linestyle = linestyle_edge_id
        pa = XYZ(xf, yf - hf, zf)
        pb = XYZ(xf, yf + hf, zf)
    else:
        return False

    mc = create_horizontal_xy_model_line(document, pa, pb, linestyle)
    return mc is not None


def _resolve_hook_135(doc):
    u"""
    Devuelve el ``RebarHookType`` con ángulo 135°.

    Si no existe ninguno con ese ángulo, devuelve el primer ``RebarHookType``
    disponible (con cualquier ángulo) para permitir la creación igualmente.
    Devuelve ``None`` sólo si no existe ningún ``RebarHookType`` en el proyecto.
    """
    import math
    from Autodesk.Revit.DB import FilteredElementCollector, BuiltInParameter
    from Autodesk.Revit.DB.Structure import RebarHookType as _RHT
    try:
        hooks = list(
            FilteredElementCollector(doc)
            .OfClass(_RHT)
            .WhereElementIsElementType()
            .ToElements()
        )
    except Exception:
        return None
    if not hooks:
        return None
    # First pass: look for "135" in the type name
    for h in hooks:
        for bip in (BuiltInParameter.SYMBOL_NAME_PARAM, BuiltInParameter.ALL_MODEL_TYPE_NAME):
            try:
                p = h.get_Parameter(bip)
                if p and p.HasValue:
                    s = p.AsString()
                    if s and u"135" in s:
                        return h
            except Exception:
                pass
    # Second pass: match by hook angle (tolerance ±2°)
    for h in hooks:
        try:
            angle_deg = float(h.HookAngle) * 180.0 / math.pi
            if abs(angle_deg - 135.0) < 2.0:
                return h
        except Exception:
            pass
    # Fallback: return any available hook type
    return hooks[0]


def _column_layout_pbar_phase_title(base_title, total):
    """Título inicial 0/N, igual que ``ProgressService.phase_title`` (exportar láminas)."""
    try:
        t = max(int(total), 1)
    except Exception:
        t = 1
    return u"{} 0/{}".format(base_title, t)


def _column_layout_pbar_start(title, count):
    u"""
    ``forms.ProgressBar`` de pyRevit con acento BIMTools (mismo estilo que
    ``export_laminas_services.ProgressService`` — ``pyRevitAccentBrush`` 91,192,222).
    """
    if count is None or int(count) < 1:
        return None
    try:
        from pyrevit import forms as _pyrevit_forms

        pb = _pyrevit_forms.ProgressBar(
            title=title,
            cancellable=False,
        )
        try:
            from System.Windows.Media import Color, SolidColorBrush

            r, g, b = (91, 192, 222)
            pb.Resources[u"pyRevitAccentBrush"] = SolidColorBrush(Color.FromRgb(r, g, b))
        except Exception:
            pass
        return pb
    except Exception:
        return None


def _column_layout_pbar_step(pb, current_index, count, base_title):
    u"""*current_index*: 0…count-1. Igual convención que ``ProgressService.step`` (láminas)."""
    if pb is None:
        return
    c = int(count) if count else 0
    if c < 1:
        c = 1
    i = int(current_index) + 1
    try:
        if hasattr(pb, u"update_progress"):
            try:
                pb.update_progress(i, max_value=c)
            except TypeError:
                try:
                    pb.update_progress(i, max=c)
                except Exception:
                    pass
    except Exception:
        pass
    try:
        pb.title = u"{} {}/{}".format(base_title, i, c)
    except Exception:
        pass


def main():
    global doc, uidoc

    try:
        if doc is None and uidoc is not None:
            doc = uidoc.Document
        if uidoc is None and doc is not None:
            uiapp = doc.Application.ActiveUIDocument
            if uiapp is not None:
                uidoc = uiapp
    except Exception:
        pass

    if doc is None or uidoc is None:
        raise Exception(
            "No hay Document/UIDocument. Abre un proyecto activo "
            "(RPS doc/uidoc o pyRevit __revit__)."
        )

    try:
        refs = pick_structural_columns(uidoc)
    except OperationCanceledException:
        TaskDialog.Show("Cancelado", "Selección cancelada.")
        return

    if not refs:
        TaskDialog.Show("Cancelado", "No seleccionaste ninguna columna.")
        return

    columns_ordered = build_column_elements_ordered(doc, refs)

    section_buckets = defaultdict(list)
    dims_cache = {}
    omitidas_msgs = []

    for col in columns_ordered:
        try:
            width, depth, height, center_chk, grid_vs, grid_vl = get_column_dimensions(col)
        except Exception as ex:
            cid = getattr(col.Id, "IntegerValue", "?")
            omitidas_msgs.append(u"Id {} — {}".format(cid, str(ex)))
            continue
        iv = _element_id_iv(col)
        if iv < 0:
            continue
        sk = _canonical_section_mm_key(width, depth)
        dims_cache[iv] = (width, depth, height, center_chk, grid_vs, grid_vl)
        section_buckets[sk].append(col)

    if not section_buckets:
        det = u"\n".join(omitidas_msgs[:12])
        if len(omitidas_msgs) > 12:
            det += u"\n…"
        msg = (
            u"Ninguna columna pudo leerse la geometría de sección para agrupar."
            u"\n{}".format(det)
        )
        TaskDialog.Show(u"Layout columna", msg)
        return

    ordered_section_keys = sorted(
        section_buckets.keys(),
        key=lambda k: (-len(section_buckets[k]), -int(k[1]), -int(k[0])),
    )

    _wizard_import_error = None
    try:
        from column_reinforcement.ui.column_layout_wizard_window import (
            show_column_layout_wizard_singleton,
        )
        from column_reinforcement.ui.troceo_scheme_window import TroceoSchemeOutcome
    except Exception as ex:
        _wizard_import_error = ex
        show_column_layout_wizard_singleton = None
        TroceoSchemeOutcome = None

    if show_column_layout_wizard_singleton is None or TroceoSchemeOutcome is None:
        msg = u"No se pudo cargar el asistente WPF «Arainco: Armado Columnas»."
        if _wizard_import_error is not None:
            try:
                msg += u"\n\n{}".format(_wizard_import_error)
            except Exception:
                pass
        TaskDialog.Show(u"Layout columna", msg)
        return

    section_wizard_meta = []
    for sk in ordered_section_keys:
        n_in_sec = len(section_buckets[sk])
        sec_title = u"{0} pilar{1} Revit (misma sección)".format(
            n_in_sec,
            u"es" if n_in_sec != 1 else u"",
        )
        section_wizard_meta.append((sk, sec_title))

    troceo_rows = _build_troceo_scheme_rows(columns_ordered)

    wiz = show_column_layout_wizard_singleton(
        section_wizard_meta,
        troceo_rows,
        uidoc.Application,
        uidoc,
        doc,
        float(LAYOUT_BAR_NOMINAL_DIAM_MM),
    )
    if wiz is None:
        return
    if getattr(wiz, "already_running", False):
        TaskDialog.Show(
            u"Layout columna",
            u"La herramienta ya esta en ejecucion.",
        )
        return
    if wiz.cancelled:
        return

    section_grid_config = wiz.section_grid_config
    stirrup_configs = getattr(wiz, "stirrup_configs", None) or {}
    stirrup_spacing_by_column_id = getattr(
        wiz,
        "stirrup_spacing_by_column_id",
        None,
    ) or {}
    stirrup_bar_type_by_column_id = getattr(
        wiz,
        "stirrup_bar_type_by_column_id",
        None,
    ) or {}
    wizard_troceo_outcome = wiz.troceo_outcome
    if wizard_troceo_outcome is None:
        wizard_troceo_outcome = TroceoSchemeOutcome(
            skip_no_cut=True,
            columns=[],
            segment_rebar_bar_type_ids=None,
        )

    cover = 0.15

    # Resolver RebarBarType antes del bucle de jobs para que el diámetro modelo
    # sea idéntico en la rejilla de barras longitudinales y en la de estribos.
    _early_bar_type, _, _ = _resolve_rebar_bar_type_by_diameter_mm(
        doc, float(LAYOUT_BAR_NOMINAL_DIAM_MM)
    )
    _long_bar_model_diam_mm = float(LAYOUT_BAR_NOMINAL_DIAM_MM)
    if _early_bar_type is not None:
        try:
            _long_bar_model_diam_mm = float(
                UnitUtils.ConvertFromInternalUnits(
                    float(_early_bar_type.BarModelDiameter),
                    UnitTypeId.Millimeters,
                )
            )
        except Exception:
            try:
                _nd = _rebar_nominal_diameter_mm(_early_bar_type)
                if _nd is not None:
                    _long_bar_model_diam_mm = float(_nd)
            except Exception:
                pass

    # Importar geometría compartida del pipeline de estribos.
    _column_bar_geometry = None
    try:
        import sys as _sys_cbg
        if "column_stirrup_creator" in _sys_cbg.modules:
            del _sys_cbg.modules["column_stirrup_creator"]
        from column_stirrup_creator import column_bar_geometry as _column_bar_geometry
    except Exception:
        _column_bar_geometry = None

    cfgs_set = {
        (
            int(c["bars_a"]),
            int(c["bars_b"]),
            bool(c["include_inner_outline"]),
        )
        for c in section_grid_config.values()
    }

    grid_uniform = len(cfgs_set) == 1

    if grid_uniform:
        bars_a, bars_b, include_inner_outline = next(iter(cfgs_set))
        outer_n = perimeter_outer_bar_count(bars_a, bars_b)
        inner_n = (
            perimeter_inner_outline_count(bars_a, bars_b)
            if include_inner_outline
            else 0
        )
        hilos_esperados_por_columna = hilos_esperados_una_columna(
            bars_a, bars_b, include_inner_outline,

        )
        rejilla_resumen = u"A:{} × B:{} ({} hilos por columna)".format(
            bars_a,
            bars_b,
            hilos_esperados_por_columna,

        )
    else:
        bars_a = None
        bars_b = None
        include_inner_outline = None
        outer_n = None
        inner_n = None
        hilos_esperados_por_columna = None
        rejilla_resumen = (
            u"Varias configuraciones por sección (revisar asistente — paso Rejilla)"
        )


    esperado_si_todas = 0
    jobs = []

    for col in columns_ordered:
        iv = _element_id_iv(col)
        if iv < 0 or iv not in dims_cache:
            continue
        width, depth, height, center, grid_vs, grid_vl = dims_cache[iv]
        sk_col = _canonical_section_mm_key(width, depth)
        cfg = section_grid_config[sk_col]
        ba = int(cfg["bars_a"])
        bb = int(cfg["bars_b"])
        inc_in = bool(cfg["include_inner_outline"])
        esperado_si_todas += hilos_esperados_una_columna(ba, bb, inc_in)

        side_short = min(width, depth)
        side_long = max(width, depth)
        short_on_x = width <= depth

        # Geometría alineada al pipeline de estribos: mismo plan_anchor, ejes lx/ly,
        # dimensiones sa/sb y offset_long (cover + Ø_estribo + r_barra longitudinal).
        _stir_geom = None
        if _column_bar_geometry is not None:
            try:
                _scfg = stirrup_configs.get(sk_col)
                _sbt = getattr(_scfg, "stirrup_bar_type", None) if _scfg else None
                _ov_bt_geom = _stirrup_bar_type_override_for_column(
                    stirrup_bar_type_by_column_id,
                    col,
                )
                if _ov_bt_geom is not None:
                    _sbt = _ov_bt_geom
                if _sbt is None:
                    _sbt = _early_bar_type
                _stir_geom = _column_bar_geometry(
                    col,
                    stirrup_bar_type=_sbt,
                    long_bar_diam_mm=_long_bar_model_diam_mm,
                )
            except Exception:
                _stir_geom = None

        if _stir_geom is not None:
            _pt_center, _lx, _ly, _sa, _sb, _off_long = _stir_geom
            pts = generate_bar_points(
                _pt_center,
                _sa,
                _sb,
                True,
                ba,
                bb,
                _off_long,
                inc_in,
                _lx,
                _ly,
            )
        else:
            # Fallback: stirrup creator no disponible; usar ejes del transform y
            # cover conservador (cubre cualquier Ø_estribo razonable).
            pts = generate_bar_points(
                center,
                side_short,
                side_long,
                short_on_x,
                ba,
                bb,
                cover,
                inc_in,
                grid_vs,
                grid_vl,
            )

        jobs.append(dict(
            height=height,
            nominal_n=len(pts),
            raw_pts=pts,
            width=width,
            depth=depth,
            short_on_x=short_on_x,
            elem=col,
            section_key_mm=sk_col,
            bars_a=ba,
            bars_b=bb,
            include_inner_outline=inc_in,
        ))


    if not jobs:
        det = "\n".join(omitidas_msgs[:12])
        if len(omitidas_msgs) > 12:
            det += "\n…"
        msg = (
            "Ninguna columna pudo generarse (sección/rejilla)."
            "\n{}".format(det)
        )
        TaskDialog.Show("Layout columna", msg)
        return


    nominal_pts_total = sum(int(jb["nominal_n"]) for jb in jobs)

    curve_tol = doc.Application.ShortCurveTolerance

    fused_world = fuse_vertical_world_intervals_from_jobs(jobs, curve_tol)

    embed_extend_mm = _resolved_traslape_embed_mm(
        LAYOUT_BAR_NOMINAL_DIAM_MM,
        LAYOUT_EMBED_CONCRETE_GRADE,
    )

    dz_extend_top = 0.0
    if embed_extend_mm is not None and float(embed_extend_mm) > 1e-9:
        dz_extend_top = UnitUtils.ConvertToInternalUnits(
            float(embed_extend_mm),
            UnitTypeId.Millimeters,
        )

    revoke_shrink_mm_total = (
        float(_REVOKE_EMBED_EXTRA_SHRINK_MM)
        + float(LAYOUT_BAR_NOMINAL_DIAM_MM) / 2.0
    )

    dz_revoke_extra_shrink = 0.0
    if dz_extend_top > 1e-12:
        dz_revoke_extra_shrink = UnitUtils.ConvertToInternalUnits(
            float(revoke_shrink_mm_total),
            UnitTypeId.Millimeters,
        )

    if not fused_world:
        TaskDialog.Show(
            "Layout columna",

            "No se generó ninguna barra (revisar tolerancia de curvas "

            "o posiciones de rejilla).",
        )
        return

    layout_bar_type, layout_bar_type_exact, layout_bar_type_delta_mm = (
        _resolve_rebar_bar_type_by_diameter_mm(
            doc,
            float(LAYOUT_BAR_NOMINAL_DIAM_MM),
        )
    )
    if layout_bar_type is None:
        TaskDialog.Show(
            "Layout columna",
            u"No se encontró ningún RebarBarType en el proyecto. "
            u"No se crearán barras de armadura.",
        )
        return

    a_bar_cut_planes = []
    b_bar_cut_planes = []
    cols_plane = []
    n_reference_columns_cut_planes = 0
    troceo_segment_diams = None
    troceo_segment_bar_type_ids = None
    outcome_tc = wizard_troceo_outcome
    try:
        if outcome_tc.skip_no_cut:
            cols_plane = []
            n_reference_columns_cut_planes = 0
            a_bar_cut_planes = []
            b_bar_cut_planes = []
            troceo_segment_diams = None
            troceo_segment_bar_type_ids = None
        else:
            cols_plane = (
                list(outcome_tc.columns) if getattr(outcome_tc, "columns", None) else []
            )
            troceo_segment_bar_type_ids = getattr(
                outcome_tc,
                "segment_rebar_bar_type_ids",
                None,
            )
            troceo_segment_diams = None
            if troceo_segment_bar_type_ids:
                troceo_segment_diams = []
                for tid in troceo_segment_bar_type_ids:
                    eid = _element_id_from_int(tid)
                    el = doc.GetElement(eid) if eid is not None else None
                    dm = _rebar_nominal_diameter_mm(el)
                    troceo_segment_diams.append(
                        float(dm)
                        if dm is not None
                        else float(LAYOUT_BAR_NOMINAL_DIAM_MM)
                    )
            n_reference_columns_cut_planes = len(cols_plane)
            a_bar_cut_planes = []
            b_bar_cut_planes = []
            if cols_plane:
                a_bar_cut_planes = build_column_cut_planes_from_elements(
                    cols_plane,
                    curve_tol,
                )
                b_bar_cut_planes = build_b_bar_cut_planes_from_elements(
                    cols_plane,
                    curve_tol,
                    float(dz_extend_top),
                )
    except Exception as ex:
        try:
            TaskDialog.Show(
                u"Arainco: Esquema de troceo",
                u"No se aplicaron planos de troceo.\n\n{}".format(ex),
            )
        except Exception:
            pass
        a_bar_cut_planes = []
        b_bar_cut_planes = []
        cols_plane = []
        n_reference_columns_cut_planes = 0
        troceo_segment_diams = None
        troceo_segment_bar_type_ids = None

    if (
        troceo_segment_diams
        and cols_plane
        and n_reference_columns_cut_planes > 0
    ):
        try:
            d_max_mm = max(
                float(x) for x in troceo_segment_diams if float(x) > 1e-9
            )
            emb_b = _resolved_traslape_embed_mm(
                d_max_mm,
                LAYOUT_EMBED_CONCRETE_GRADE,
            )
            dz_b = 0.0
            if emb_b is not None and float(emb_b) > 1e-9:
                dz_b = UnitUtils.ConvertToInternalUnits(
                    float(emb_b),
                    UnitTypeId.Millimeters,
                )
            b_bar_cut_planes = build_b_bar_cut_planes_from_elements(
                cols_plane,
                curve_tol,
                float(dz_b),
            )
        except Exception:
            pass

    troceo_z_cuts_ft_sorted = _troceo_sorted_cut_z_ft_from_planes(a_bar_cut_planes)
    troceo_n_ui_segs = 0
    if troceo_segment_bar_type_ids:
        troceo_n_ui_segs = len(troceo_segment_bar_type_ids)
    elif troceo_segment_diams:
        troceo_n_ui_segs = len(troceo_segment_diams)
    troceo_use_z_mid_for_ui_seg = (
        troceo_n_ui_segs > 0
        and len(troceo_z_cuts_ft_sorted) + 1 == troceo_n_ui_segs
    )

    col_iv_fund_down_ft = build_selected_columns_foundation_down_ft(
        doc,
        columns_ordered,
    )

    pata_hook_mm = float(
        _resolved_pata_hook_mm_for_revert(
            LAYOUT_BAR_NOMINAL_DIAM_MM,
            LAYOUT_EMBED_CONCRETE_GRADE,
        )
    )
    pata_hook_ft = 0.0
    if pata_hook_mm > 1e-6:
        pata_hook_ft = UnitUtils.ConvertToInternalUnits(
            float(pata_hook_mm),
            UnitTypeId.Millimeters,
        )

    _sketch_plane_cache.clear()

    try:
        column_instances_embed = list(
            FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_StructuralColumns)
            .WhereElementIsNotElementType()
        )
    except Exception:
        column_instances_embed = []

    embed_geom_opts = _geometry_options_structure_solids()

    hilos_totales_real = 0

    n_kept_embed_collision = 0
    n_reverted_embed_air = 0

    n_kept_embed_start_collision = 0
    n_reverted_embed_start_air = 0

    n_hilos_foundation_down_stretch = 0

    n_patitas_revert_embed = 0
    n_patitas_foundation_down = 0

    scheme_histogram = Counter()

    vrf_marker_half_ft = 0.0
    ls_scheme_corner_id = None
    ls_scheme_edge_id = None
    if _SCHEME_VERIFY_MARKER_ENABLED:
        try:
            vrf_marker_half_ft = UnitUtils.ConvertToInternalUnits(
                float(_SCHEME_VERIFY_MARKER_HALF_MM),
                UnitTypeId.Millimeters,
            )
        except Exception:
            vrf_marker_half_ft = 0.0
        ls_scheme_corner_id, ls_scheme_edge_id = (
            _scheme_verify_resolve_pair_linestyle_ids(doc)
        )

    ls_cut_plane_id = None
    if _CUT_PLANE_MARKER_ENABLED:
        try:
            ls_cut_plane_id = _resolve_cut_plane_linestyle_id(doc)
        except Exception:
            ls_cut_plane_id = None

    n_scheme_verify_markers_drawn = 0

    n_cut_plane_marker_elems = 0

    n_a_bar_extra_segments_from_cut_planes = 0

    n_b_bar_extra_segments_from_cut_planes = 0

    bar_type_cache = {}
    line_plans = []
    _pb_layout = None
    _pbar_layout_open = False
    for line_idx, (base_xyz, span_z, contrib_ids, bar_enum_lab) in enumerate(
        fused_world
    ):
        min_len_z = float(curve_tol) * 4.0
        span_draw_after_top = float(span_z)
        reverted_embed = False
        kept_top_embed = False
        if dz_extend_top > 1e-12:
            if embed_stretch_collides_any_column_solids(
                doc,
                base_xyz,
                float(span_z),
                dz_extend_top,
                float(LAYOUT_BAR_NOMINAL_DIAM_MM),
                column_instances_embed,
                embed_geom_opts,
                contrib_ids,
            ):
                kept_top_embed = True
                span_draw_after_top = float(span_z) + float(dz_extend_top)
                n_kept_embed_collision += 1
            else:
                reverted_embed = True
                n_reverted_embed_air += 1
                span_draw_after_top = max(
                    float(span_z) - float(dz_revoke_extra_shrink),
                    min_len_z,
                )

        dz_fdown_ft = _contrib_max_foundation_down_ft(
            contrib_ids,
            col_iv_fund_down_ft,
        )
        if dz_fdown_ft > 1e-12:
            n_hilos_foundation_down_stretch += 1

        z_lo_core = float(base_xyz.Z) - float(dz_fdown_ft)
        core_height = float(dz_fdown_ft) + float(span_z)
        z_hi_core = z_lo_core + core_height

        bx = float(base_xyz.X)
        by = float(base_xyz.Y)
        btag = str(bar_enum_lab or "").strip()

        segments = [(z_lo_core, core_height)]
        if a_bar_cut_planes and is_a_split_scheme(btag):
            segs_try = _split_z_span_by_planes(
                bx,
                by,
                z_lo_core,
                z_hi_core,
                a_bar_cut_planes,
                float(curve_tol),
            )
            if segs_try:
                segments = segs_try
                if len(segs_try) > 1:
                    n_a_bar_extra_segments_from_cut_planes += int(len(segs_try) - 1)

        if b_bar_cut_planes and is_b_split_scheme(btag):
            segs_b = _split_z_span_by_planes(
                bx,
                by,
                z_lo_core,
                z_hi_core,
                b_bar_cut_planes,
                float(curve_tol),
            )
            if segs_b:
                segments = segs_b
                if len(segs_b) > 1:
                    n_b_bar_extra_segments_from_cut_planes += int(len(segs_b) - 1)

        seg_list = [[float(s[0]), float(s[1])] for s in segments]

        if dz_extend_top > 1e-12 and seg_list:
            li = len(seg_list) - 1
            zs_l = seg_list[li][0]
            dsz_l = seg_list[li][1]
            if kept_top_embed:
                seg_list[li][1] = dsz_l + float(dz_extend_top)
            elif reverted_embed:
                delta_h = float(span_z) - float(span_draw_after_top)
                if delta_h > 1e-12:
                    seg_list[li][1] = max(dsz_l - delta_h, min_len_z)

        if dz_extend_top > 1e-12 and dz_fdown_ft <= 1e-12 and seg_list:
            if embed_start_collides_any_column_solids(
                doc,
                base_xyz,
                dz_extend_top,
                float(LAYOUT_BAR_NOMINAL_DIAM_MM),
                column_instances_embed,
                embed_geom_opts,
                contrib_ids,
            ):
                seg_list[0][0] -= float(dz_extend_top)
                seg_list[0][1] += float(dz_extend_top)
                n_kept_embed_start_collision += 1
            else:
                span_draw_eff = float(span_draw_after_top)
                span_draw_floor = max(
                    span_draw_eff - float(dz_revoke_extra_shrink),
                    min_len_z,
                )
                revoke_delta = span_draw_eff - span_draw_floor
                if revoke_delta > 1e-12:
                    seg_list[0][0] += revoke_delta
                    seg_list[0][1] -= revoke_delta
                n_reverted_embed_start_air += 1

        if not seg_list:
            continue
        z_line_start = seg_list[0][0]
        z_top_whole = seg_list[-1][0] + seg_list[-1][1]
        span_vertical = z_top_whole - z_line_start

        n_seg_total = len(seg_list)
        seg_jobs = []
        for seg_i, row in enumerate(seg_list):
            zs = row[0]
            dsz = row[1]
            span_seg = float(dsz)
            z_mid_seg = float(zs) + 0.5 * float(span_seg)
            if troceo_use_z_mid_for_ui_seg:
                troceo_ui_i = _troceo_ui_segment_index_for_z_mid(
                    z_mid_seg,
                    troceo_z_cuts_ft_sorted,
                    troceo_n_ui_segs,
                    seg_i,
                )
            else:
                troceo_ui_i = int(seg_i)
            if troceo_segment_bar_type_ids:
                _tid = (
                    troceo_segment_bar_type_ids[troceo_ui_i]
                    if troceo_ui_i < len(troceo_segment_bar_type_ids)
                    else troceo_segment_bar_type_ids[-1]
                )
                if _tid not in bar_type_cache:
                    _eid = _element_id_from_int(_tid)
                    _el = doc.GetElement(_eid) if _eid is not None else None
                    bar_type_cache[_tid] = (
                        _el if _el is not None else layout_bar_type
                    )
                layout_bar_type_seg = bar_type_cache[_tid]
                d_mm_seg = _rebar_nominal_diameter_mm(layout_bar_type_seg)
                if d_mm_seg is None:
                    d_mm_seg = float(LAYOUT_BAR_NOMINAL_DIAM_MM)
                else:
                    d_mm_seg = float(d_mm_seg)
            else:
                d_mm_seg = _troceo_nominal_diam_mm_for_seg_index(
                    troceo_segment_diams,
                    troceo_ui_i,
                    LAYOUT_BAR_NOMINAL_DIAM_MM,
                )
                if d_mm_seg not in bar_type_cache:
                    bt_seg, _, _ = _resolve_rebar_bar_type_by_diameter_mm(
                        doc,
                        float(d_mm_seg),
                    )
                    bar_type_cache[d_mm_seg] = (
                        bt_seg if bt_seg is not None else layout_bar_type
                    )
                layout_bar_type_seg = bar_type_cache[d_mm_seg]
            embed_mm_seg = _resolved_traslape_embed_mm(
                d_mm_seg,
                LAYOUT_EMBED_CONCRETE_GRADE,
            )
            dz_extend_seg = 0.0
            if embed_mm_seg is not None and float(embed_mm_seg) > 1e-9:
                dz_extend_seg = UnitUtils.ConvertToInternalUnits(
                    float(embed_mm_seg),
                    UnitTypeId.Millimeters,
                )
            if (
                n_seg_total > 1
                and is_lap_extension_scheme(btag)
                and dz_extend_seg > 1e-12
                and seg_i < n_seg_total - 1
            ):
                span_seg += float(dz_extend_seg)
            pata_hook_mm_seg = float(
                _resolved_pata_hook_mm_for_revert(
                    d_mm_seg,
                    LAYOUT_EMBED_CONCRETE_GRADE,
                )
            )
            pata_hook_ft_seg = 0.0
            if pata_hook_mm_seg > 1e-6:
                pata_hook_ft_seg = UnitUtils.ConvertToInternalUnits(
                    float(pata_hook_mm_seg),
                    UnitTypeId.Millimeters,
                )
            want_bot_pata = (
                seg_i == 0
                and dz_fdown_ft > 1e-12
                and pata_hook_ft_seg > 1e-12
            )
            want_top_pata = (
                seg_i == n_seg_total - 1
                and reverted_embed
                and pata_hook_ft_seg > 1e-12
            )
            z_for_host = float(zs) + float(span_seg) * 0.5
            if want_bot_pata and want_top_pata:
                pass
            elif want_bot_pata:
                _band = min(
                    max(float(span_seg) * 0.12, float(curve_tol) * 24.0),
                    float(span_seg) * 0.45,
                )
                z_for_host = float(zs) + _band
            elif want_top_pata:
                _band = min(
                    max(float(span_seg) * 0.12, float(curve_tol) * 24.0),
                    float(span_seg) * 0.45,
                )
                z_for_host = float(zs) + float(span_seg) - _band
            host_col = _nearest_contrib_column_for_xyz(
                doc,
                bx,
                by,
                z_for_host,
                contrib_ids,
            )
            seg_jobs.append(
                {
                    "troceo_ui_i": int(troceo_ui_i),
                    "seg_i": int(seg_i),
                    "zs": float(zs),
                    "span_seg": float(span_seg),
                    "host_col": host_col,
                    "layout_bar_type_seg": layout_bar_type_seg,
                    "bar_enum_lab": bar_enum_lab,
                    "want_bot_pata": want_bot_pata,
                    "want_top_pata": want_top_pata,
                    "pata_hook_ft_seg": float(pata_hook_ft_seg),
                    "contrib_ids": contrib_ids,
                }
            )
        line_plans.append(
            {
                "line_idx": line_idx,
                "bx": bx,
                "by": by,
                "z_line_start": z_line_start,
                "span_vertical": span_vertical,
                "n_seg_total": n_seg_total,
                "btag": btag,
                "bar_enum_lab": bar_enum_lab,
                "seg_jobs": seg_jobs,
            }
        )

    # ── Tabla de largos reales por (bar_enum, troceo_ui_i) ──────────────────
    # Calculada desde seg_jobs —mismas curvas que se usarán en CreateFromCurves—
    # sin crear ninguna barra ni abrir transacciones.
    # Clave: (bar_enum_lab_str, troceo_ui_i_int)
    # Valor: (lz_mm, lp_mm, has_bot_pata, has_top_pata)
    #   lz_mm = tronco vertical; lp_mm = largo de pata (si aplica, igual en ambos extremos).
    preview_bar_lengths_by_ubic_tramo = {}
    for lp_pbl in line_plans:
        lab_pbl = str(lp_pbl.get("bar_enum_lab") or "").strip()
        for sj_pbl in lp_pbl.get("seg_jobs") or []:
            _lz = _arma_len_mm_round_from_internal_ft(sj_pbl["span_seg"])
            _lp = _arma_len_mm_round_from_internal_ft(sj_pbl["pata_hook_ft_seg"])
            _kb = bool(sj_pbl["want_bot_pata"])
            _kt = bool(sj_pbl["want_top_pata"])
            _key = (lab_pbl, int(sj_pbl["troceo_ui_i"]))
            # Si varios hilos con la misma etiqueta/tramo difieren, conservar el mayor tronco
            existing = preview_bar_lengths_by_ubic_tramo.get(_key)
            if existing is None or float(_lz) > float(existing[0]):
                preview_bar_lengths_by_ubic_tramo[_key] = (_lz, _lp, _kb, _kt)
    # ────────────────────────────────────────────────────────────────────────

    max_tramo_ui = -1
    for lp in line_plans:
        for sj in lp["seg_jobs"]:
            if sj["troceo_ui_i"] > max_tramo_ui:
                max_tramo_ui = sj["troceo_ui_i"]

    n_tramos = max(0, max_tramo_ui + 1)
    has_stirrups = any(
        not getattr(cfg, "skip", True) for cfg in stirrup_configs.values()
    )
    total_pb_steps = 1 + n_tramos + 1
    if total_pb_steps < 1:
        total_pb_steps = 1

    _COL_PBAR_BASE = u"Arainco: Columnas armado longitudinal"
    _pb_layout = _column_layout_pbar_start(
        _column_layout_pbar_phase_title(_COL_PBAR_BASE, total_pb_steps),
        total_pb_steps,
    )
    _pbar_layout_open = False
    if _pb_layout is not None:
        try:
            _pb_layout.__enter__()
            _pbar_layout_open = True
        except Exception:
            _pb_layout = None


    rebars_for_3d_visibility = []

    tg_layout = TransactionGroup(
        doc, u"Arainco: Armado longitudinal columnas por tramos"
    )
    tg_layout.Start()

    try:
        _column_layout_pbar_step(
            _pb_layout,
            0,
            total_pb_steps,
            u"Arainco: Columnas (marcadores planos corte)",
        )
        txn_markers = Transaction(
            doc, u"Arainco: Marcadores planos corte fusion global columnas"
        )
        txn_markers.Start()
        try:
            if (
                _CUT_PLANE_MARKER_ENABLED
                and cols_plane
            ):
                try:
                    nlen_ft = UnitUtils.ConvertToInternalUnits(
                        float(_CUT_PLANE_NORMAL_DISPLAY_MM),
                        UnitTypeId.Millimeters,
                    )
                    ileg_ft = UnitUtils.ConvertToInternalUnits(
                        float(_CUT_PLANE_INPLANE_LEG_MM),
                        UnitTypeId.Millimeters,
                    )
                except Exception:
                    nlen_ft = 0.0
                    ileg_ft = 0.0
                if nlen_ft > 1e-12 and ileg_ft > 1e-12:
                    try:
                        n_cut_plane_marker_elems = _draw_cut_plane_markers_for_columns(
                            doc,
                            cols_plane,
                            a_bar_cut_planes,
                            curve_tol,
                            nlen_ft,
                            ileg_ft,
                            ls_cut_plane_id,
                        )
                    except Exception:
                        n_cut_plane_marker_elems = 0
            txn_markers.Commit()
        except Exception:
            if txn_markers.HasStarted():
                txn_markers.RollBack()
            raise

        line_rb_accum = defaultdict(list)
        line_marker_done = set()

        for tramo_k in range(0, max_tramo_ui + 1):
            _column_layout_pbar_step(
                _pb_layout,
                1 + tramo_k,
                total_pb_steps,
                u"Arainco: Columnas (Structural Rebar por tramo)",
            )
            txn_bar = Transaction(
                doc,
                u"Arainco: Structural Rebar columnas tramo {}".format(tramo_k + 1),
            )
            txn_bar.Start()
            try:
                for lp in line_plans:
                    line_idx = lp["line_idx"]
                    bx = lp["bx"]
                    by = lp["by"]
                    z_line_start = lp["z_line_start"]
                    span_vertical = lp["span_vertical"]
                    btag = lp["btag"]
                    bar_enum_lab = lp["bar_enum_lab"]
                    for sj in lp["seg_jobs"]:
                        if sj["troceo_ui_i"] != tramo_k:
                            continue
                        vert_rb, ok_pat_bot, ok_pat_top = (
                            create_longitudinal_rebar_with_optional_patas(
                                doc,
                                bx,
                                by,
                                sj["zs"],
                                sj["span_seg"],
                                sj["host_col"],
                                sj["layout_bar_type_seg"],
                                sj["bar_enum_lab"],
                                sj["want_bot_pata"],
                                sj["want_top_pata"],
                                sj["pata_hook_ft_seg"],
                                sj["contrib_ids"],
                            )
                        )
                        if vert_rb is None:
                            continue
                        hilos_totales_real += 1
                        try:
                            rebars_for_3d_visibility.append(vert_rb)
                        except Exception:
                            pass
                        try:
                            fp_seg = _fingerprint_seg_linea_fierro(
                                sj["span_seg"],
                                ok_pat_bot,
                                ok_pat_top,
                                sj["pata_hook_ft_seg"],
                            )
                            line_rb_accum[line_idx].append(
                                (
                                    sj["troceo_ui_i"],
                                    sj["seg_i"],
                                    vert_rb,
                                    fp_seg,
                                )
                            )
                        except Exception:
                            pass
                        if ok_pat_bot:
                            n_patitas_foundation_down += 1
                        if ok_pat_top:
                            n_patitas_revert_embed += 1
                        if btag:
                            scheme_histogram[btag] += 1
                        if (
                            line_idx not in line_marker_done
                            and _SCHEME_VERIFY_MARKER_ENABLED
                            and vrf_marker_half_ft > 1e-12
                        ):
                            zm = float(z_line_start) + float(span_vertical) * 0.5
                            mk = _scheme_verify_draw_scheme_marker_horizontal(
                                doc,
                                bx,
                                by,
                                zm,
                                vrf_marker_half_ft,
                                float(curve_tol),
                                ls_scheme_corner_id,
                                ls_scheme_edge_id,
                                bar_enum_lab,
                            )
                            if mk:
                                n_scheme_verify_markers_drawn += 1
                            line_marker_done.add(line_idx)
                txn_bar.Commit()
            except Exception:
                if txn_bar.HasStarted():
                    txn_bar.RollBack()
                raise
            if uidoc is not None:
                try:
                    uidoc.RefreshActiveView()
                except Exception:
                    pass

        _column_layout_pbar_step(
            _pb_layout,
            1 + n_tramos,
            total_pb_steps,
            u"Arainco: Columnas (parámetro línea fierro)",
        )

        arma_line_assignments = []
        for lp in line_plans:
            line_idx = lp["line_idx"]
            n_seg_total = lp["n_seg_total"]
            items = line_rb_accum.get(line_idx)
            if not items:
                continue
            items_sorted = sorted(items, key=lambda x: (x[0], x[1]))
            line_rebars_fierro = [x[2] for x in items_sorted]
            line_fps_fierro = [x[3] for x in items_sorted]
            if (
                line_fps_fierro
                and len(line_fps_fierro) == len(line_rebars_fierro)
                and len(line_fps_fierro) == n_seg_total
            ):
                try:
                    _k_lf = (len(line_fps_fierro), tuple(line_fps_fierro))
                    for _rb_lf in line_rebars_fierro:
                        arma_line_assignments.append((_rb_lf, _k_lf))
                except Exception:
                    pass

        txn_lf = Transaction(  
            doc, u"Arainco: Parámetro línea fierro armadura ubicación columnas"
        )
        txn_lf.Start()
        try:
            _apply_linea_fierro_armadura_ubicacion(arma_line_assignments)
            txn_lf.Commit()
        except Exception:
            if txn_lf.HasStarted():
                txn_lf.RollBack()
            raise

        tg_layout.Assimilate()
    except Exception:
        tg_layout.RollBack()
        raise
    finally:
        if _pbar_layout_open and _pb_layout is not None:
            try:
                _pb_layout.__exit__(None, None, None)
            except Exception:
                pass

    if has_stirrups:
        try:
            import sys as _sys
            if "column_stirrup_creator" in _sys.modules:
                del _sys.modules["column_stirrup_creator"]
        except Exception:
            pass
        _stir_import_err = None
        try:
            from column_stirrup_creator import create_stirrups_for_column as _csc
        except Exception as _e:
            _csc = None
            _stir_import_err = str(_e)
        if _csc is None:
            TaskDialog.Show(
                u"Arainco: Armado Columnas",
                u"No se pudo importar column_stirrup_creator.\n\n{}".format(
                    _stir_import_err or u""
                ),
            )
        else:
            hook_135 = _resolve_hook_135(doc)
            # Usar el mismo diámetro modelo resuelto antes del bucle de jobs
            # para que la rejilla de estribos sea idéntica a la de barras longitudinales.
            _n_stir_total = 0
            _stir_errors = []
            _stir_col_list = []
            for _col_s in columns_ordered:
                _iv_s = _element_id_iv(_col_s)
                if _iv_s < 0 or _iv_s not in dims_cache:
                    continue
                _w_s, _d_s = dims_cache[_iv_s][0], dims_cache[_iv_s][1]
                _sk_s = _canonical_section_mm_key(_w_s, _d_s)
                _cfg_s = stirrup_configs.get(_sk_s)
                if _cfg_s is None or getattr(_cfg_s, "skip", True):
                    continue
                _stir_col_list.append(_col_s)

            def _stir_col_sort_key(cm):
                try:
                    z0 = _column_base_z_ft_for_sort(cm)
                except Exception:
                    z0 = 0.0
                try:
                    eid = int(cm.Id.IntegerValue)
                except Exception:
                    try:
                        eid = int(cm.Id.Value)
                    except Exception:
                        eid = 0
                return (z0, eid)

            _stir_col_list.sort(key=_stir_col_sort_key)

            n_stir_cols = len(_stir_col_list)
            _COL_STIR_PBAR_BASE = u"Arainco: Estribos columnas"
            _pb_stir = None
            _pb_stir_open = False
            if n_stir_cols > 0:
                _pb_stir = _column_layout_pbar_start(
                    _column_layout_pbar_phase_title(
                        _COL_STIR_PBAR_BASE, n_stir_cols
                    ),
                    n_stir_cols,
                )
                if _pb_stir is not None:
                    try:
                        _pb_stir.__enter__()
                        _pb_stir_open = True
                    except Exception:
                        _pb_stir = None

            try:
                tg_stir = TransactionGroup(doc, u"Arainco: Estribos columnas")
                tg_stir.Start()
                try:
                    for _i_stir, _col_s in enumerate(_stir_col_list):
                        _column_layout_pbar_step(
                            _pb_stir,
                            _i_stir,
                            n_stir_cols,
                            _COL_STIR_PBAR_BASE,
                        )
                        _iv_s = _element_id_iv(_col_s)
                        if _iv_s < 0:
                            continue
                        _w_s, _d_s = dims_cache[_iv_s][0], dims_cache[_iv_s][1]
                        _sk_s = _canonical_section_mm_key(_w_s, _d_s)
                        _cfg_s = stirrup_configs.get(_sk_s)
                        _gcfg = section_grid_config.get(_sk_s, {})
                        _val_a = int(_gcfg.get("bars_a", 4))
                        _val_b = int(_gcfg.get("bars_b", 6))
                        _sbt = getattr(_cfg_s, "stirrup_bar_type", None)
                        _ov_bt = _stirrup_bar_type_override_for_column(
                            stirrup_bar_type_by_column_id,
                            _col_s,
                        )
                        if _ov_bt is not None:
                            _sbt = _ov_bt
                        _spacing_mm = float(
                            getattr(_cfg_s, "spacing_mm", 200.0)
                        )
                        _ov_sp = _stirrup_spacing_mm_override_for_column(
                            stirrup_spacing_by_column_id,
                            _col_s,
                        )
                        if _ov_sp is not None:
                            _spacing_mm = float(_ov_sp)
                        if _sbt is None:
                            _sbt = _early_bar_type
                        if _sbt is None:
                            _stir_errors.append(
                                u"Columna Id {}: sin tipo RebarBarType para estribo (def\u00edna \u00d8 en Troceo o cargue tipos en el proyecto).".format(
                                    _col_s.Id.IntegerValue
                                )
                            )
                            continue
                        txn_s = Transaction(
                            doc,
                            u"Arainco: Estribos columna {}".format(
                                _col_s.Id.IntegerValue
                            ),
                        )
                        txn_s.Start()
                        try:
                            _n = _csc(
                                doc,
                                _col_s,
                                _val_a,
                                _val_b,
                                str(getattr(_cfg_s, "sel_a_text", u"")),
                                str(getattr(_cfg_s, "sel_b_text", u"")),
                                _sbt,
                                hook_135,
                                float(_spacing_mm),
                                cover_mm=25.0,
                                long_bar_diam_mm=_long_bar_model_diam_mm,
                                collect_rebars=rebars_for_3d_visibility,
                            )
                            _n_stir_total += _n
                            if _n == 0:
                                _stir_errors.append(
                                    u"Columna Id {}: 0 estribos creados".format(
                                        _col_s.Id.IntegerValue
                                    )
                                )
                            txn_s.Commit()
                            if uidoc is not None:
                                try:
                                    uidoc.RefreshActiveView()
                                except Exception:
                                    pass
                        except Exception as _es:
                            if txn_s.HasStarted():
                                txn_s.RollBack()
                            _stir_errors.append(
                                u"Columna Id {}: {}".format(
                                    _col_s.Id.IntegerValue, str(_es)
                                )
                            )
                    tg_stir.Assimilate()
                except Exception as _eg:
                    try:
                        tg_stir.RollBack()
                    except Exception:
                        pass
                    TaskDialog.Show(
                        u"Arainco: Armado Columnas",
                        u"Error en TransactionGroup de estribos:\n{}".format(
                            str(_eg)
                        ),
                    )
                else:
                    if _stir_errors:
                        det = u"\n".join(_stir_errors[:10])
                        if len(_stir_errors) > 10:
                            det += u"\n…"
                        TaskDialog.Show(
                            u"Arainco: Armado Columnas",
                            u"{} estribo(s) creados. Advertencias:\n\n{}".format(
                                _n_stir_total, det
                            ),
                        )
            finally:
                if _pb_stir_open and _pb_stir is not None:
                    try:
                        _pb_stir.__exit__(None, None, None)
                    except Exception:
                        pass

    if rebars_for_3d_visibility:
        _txn_3d = Transaction(
            doc,
            u"Arainco: Armadura columnas visible en vistas 3D",
        )
        try:
            _txn_3d.Start()
            try:
                apply_rebar_unobscured_in_3d_views(
                    doc, rebars_for_3d_visibility
                )
                _txn_3d.Commit()
            except Exception:
                if _txn_3d.HasStarted():
                    _txn_3d.RollBack()
        except Exception:
            pass


def run_pyrevit(revit_app):
    """Compatibilidad pyRevit / ``__main__``: misma secuencia que el pushbutton.

    Delega en ``column_reinforcement.runner.run_pyrevit`` (reload del módulo legado,
    inyección doc/uidoc/__revit__, luego main).
    """
    from column_reinforcement.runner import run_pyrevit as _runner_run_pyrevit

    return _runner_run_pyrevit(revit_app)


if __name__ == "__main__":
    try:
        try:
            _revit_app = __revit__  # noqa: F821
        except NameError:
            _revit_app = None
        if _revit_app is not None:
            run_pyrevit(_revit_app)
        else:
            main()
    except Exception as ex:
        TaskDialog.Show("Error – layout columna", str(ex))

