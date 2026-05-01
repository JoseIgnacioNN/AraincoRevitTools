# -*- coding: utf-8 -*-
"""
Extracción de láminas (ViewSheet) y revisiones hacia plantilla Excel ARAINCO / TemplateListado.

Estructura del template (confirmada con PowerShell + script Dynamo de referencia):

LISTADO DE PLANOS (fila 12 = encabezado, datos desde fila 13):
  A: N° LISTADO DE PLANOS  → SheetNumber
  B: REV.                  → parámetro compartido _DES  (o revisión actual)
  C: EMISION               → parámetro compartido _REV  (o "Issued To")
  D: FECHA                 → parámetro compartido _FCH  (o fecha de revisión)
  E: DESCRIPCION           → SheetName

HISTORIAL DE REVISIONES (fila 7 = encabezado, datos desde fila 8):
  A: Numero de Plano       → SheetNumber
  B: Contenido             → SheetName
  Bloques de 6 columnas desde col C (1-based=3), paso 6:
    [Revision, Descripcion, Dibujo, Reviso, Aprobo, Fecha]
  Bloque 1: C-H  (revisión más antigua o actual)
  Bloque 2: I-N  etc.

Parámetros compartidos ARAINCO en láminas:
  _DES = Designación / etiqueta de revisión (va a REV. en LISTADO)
  _REV = Referencia de emisión (va a EMISION en LISTADO / Reviso en HISTORIAL)
  _FCH = Fecha de la revisión
  _DIB = Dibujó (iniciales)
  _APR = Aprobó (iniciales)
"""

from __future__ import print_function

import datetime
import shutil
import clr

clr.AddReference("RevitAPI")
clr.AddReference("System")

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    ModelPathUtils,
    ViewSheet,
)

# ---------------------------------------------------------------------------
# Constantes de posición (1-based, confirmadas con template real)
# ---------------------------------------------------------------------------

SHEET_LISTADO   = u"LISTADO DE PLANOS"
SHEET_HISTORIAL = u"HISTORIAL DE REVISIONES"

# LISTADO DE PLANOS
FIRST_DATA_ROW_LISTADO = 13   # fila 12 = encabezado

# HISTORIAL DE REVISIONES
FIRST_DATA_ROW_HISTORIAL = 8  # fila 7 = encabezado
BLOCK_START_COL = 3           # col C (1-based): inicio del primer bloque de revisión
BLOCK_STEP      = 6           # 6 cols por bloque: Revision, Descripcion, Dibujo, Reviso, Aprobo, Fecha
MAX_REVISION_BLOCKS = 15      # slots R01-R15

# Metadatos de proyecto en LISTADO (valor en col B, 1-based=2)
META_ROW_PROYECTO  = 7
META_ROW_CLIENTE   = 8
META_ROW_ORDEN     = 9
META_ROW_FECHA_ENV = 10

# Parámetros compartidos ARAINCO en láminas (Data group)
# Cada revisión tiene 6 campos: R{nn}_01_NUM … R{nn}_06_FCH
# Ejemplo: R01_01_NUM="A", R01_02_DES="PRELIMINAR", R01_03_DIB="J.N.N.",
#          R01_04_REV="C.M.Q.", R01_05_APR="P.A.V.", R01_06_FCH="29.08.25"
_PARAM_SLOTS = [u"R{:02d}_{:02d}_{}".format(n, f, s)
                for n in range(1, MAX_REVISION_BLOCKS + 1)
                for f, s in [(1, u"NUM"), (2, u"DES"), (3, u"DIB"),
                             (4, u"REV"), (5, u"APR"), (6, u"FCH")]]


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def default_listado_workbook_filename(doc):
    """
    Nombre de archivo del listado: ``YYYY.MM.DD_<nombre del modelo central>.xlsx``.

    Con trabajo compartido se usa el nombre base del archivo del modelo central (ruta
    devuelta por Revit). Sin central accesible, se intenta la ruta local del documento,
    ``Document.Title`` o el nombre en información del proyecto.
    """
    import os
    import re

    d = datetime.date.today()
    date_prefix = u"{:04d}.{:02d}.{:02d}".format(d.year, d.month, d.day)
    name_hint = u""
    try:
        if doc.IsWorkshared:
            cm = doc.GetWorksharingCentralModelPath()
            if cm is not None:
                vis = ModelPathUtils.ConvertModelPathToUserVisiblePath(cm)
                if vis:
                    try:
                        vis_s = unicode(vis)
                    except Exception:
                        vis_s = str(vis)
                    base = os.path.splitext(os.path.basename(vis_s))[0]
                    name_hint = (base or u"").strip()
    except Exception:
        pass
    if not name_hint:
        try:
            pn = doc.PathName
            if pn:
                try:
                    pn_s = unicode(pn)
                except Exception:
                    pn_s = str(pn)
                if pn_s:
                    name_hint = os.path.splitext(os.path.basename(pn_s))[0].strip()
        except Exception:
            pass
    if not name_hint:
        try:
            name_hint = (doc.Title or u"").strip()
        except Exception:
            name_hint = u""
    if not name_hint:
        try:
            pinf = doc.ProjectInformation
            if pinf is not None:
                name_hint = (pinf.Name or u"").strip()
        except Exception:
            pass
    if not name_hint:
        name_hint = u"Modelo"
    safe = re.sub(r'[<>:"/\\|?*]', u"_", name_hint).strip()
    if not safe:
        safe = u"Modelo"
    return date_prefix + u"_" + safe + u".xlsx"


def _param_string(elem, bip):
    try:
        p = elem.get_Parameter(bip)
        if p is None:
            return u""
        return (p.AsString() or p.AsValueString() or u"").strip()
    except Exception:
        return u""


def _lookup_param(elem, name):
    """Lee un parámetro compartido/de instancia por nombre."""
    try:
        p = elem.LookupParameter(name)
        if p is None:
            return u""
        return (p.AsString() or p.AsValueString() or u"").strip()
    except Exception:
        return u""


def _short_revision_number(rev_el):
    """Devuelve solo el número/letra corto de revisión (p. ej. '0', 'A')."""
    try:
        s = unicode(getattr(rev_el, u"RevisionNumber", None) or u"").strip()
        if s:
            return s
    except Exception:
        pass
    return u""


def _last_non_empty_fecha_from_blocks(blocks):
    """Recorre los bloques R01… de la más reciente a la más antigua; primera fecha no vacía."""
    if not blocks:
        return u""
    for b in reversed(blocks):
        try:
            f = (b.get(u"fecha") or u"").strip()
            if f:
                return f
        except Exception:
            pass
    return u""


def _revision_date_from_current_revision(sheet, doc):
    """Fecha de la revisión actual del plano (API Revit ``Revision.RevisionDate``)."""
    try:
        rid = sheet.GetCurrentRevision()
        if rid is None:
            return u""
        try:
            if rid == ElementId.InvalidElementId:
                return u""
        except Exception:
            pass
        rev = doc.GetElement(rid)
        if rev is None:
            return u""
        try:
            rd = getattr(rev, u"RevisionDate", None)
            if rd is not None:
                return unicode(rd).strip()
        except Exception:
            pass
    except Exception:
        pass
    return u""


# ---------------------------------------------------------------------------
# Lectura de datos de lámina — parámetros compartidos R{nn}_{campo}
# ---------------------------------------------------------------------------

def _read_revision_slots(sheet):
    """
    Lee los slots R01-R{MAX} de la lámina y devuelve lista de dicts con los
    6 campos por revisión, omitiendo slots vacíos (sin NUM ni DES).
    """
    blocks = []
    for n in range(1, MAX_REVISION_BLOCKS + 1):
        prefix = u"R{:02d}_".format(n)
        num = _lookup_param(sheet, prefix + u"01_NUM")
        des = _lookup_param(sheet, prefix + u"02_DES")
        dib = _lookup_param(sheet, prefix + u"03_DIB")
        rev = _lookup_param(sheet, prefix + u"04_REV")
        apr = _lookup_param(sheet, prefix + u"05_APR")
        fch = _lookup_param(sheet, prefix + u"06_FCH")
        if not num and not des:
            continue
        blocks.append({
            u"revision":    num,
            u"descripcion": des,
            u"dibujo":      dib,
            u"reviso":      rev,
            u"aprobo":      apr,
            u"fecha":       fch,
        })
    return blocks


def _current_rev_from_slots(blocks):
    """Retorna el último bloque no vacío (revisión actual) o None."""
    return blocks[-1] if blocks else None


def _sheet_rev_display_fallback(sheet, doc):
    """Fallback a built-in de Revit si no existen parámetros R{nn}."""
    try:
        rid = sheet.GetCurrentRevision()
        if rid is not None and rid != ElementId.InvalidElementId:
            rev = doc.GetElement(rid)
            if rev is not None:
                n = _short_revision_number(rev)
                if n:
                    return n
    except Exception:
        pass
    return _param_string(sheet, BuiltInParameter.SHEET_CURRENT_REVISION)


def export_row_for_sheet(sheet, doc, listado_fecha_column_override=None):
    """Una fila en el formato esperado por ``fill_template_excel`` / JSON de PowerShell.

    ``listado_fecha_column_override``: si no está vacío, sustituye el valor de la columna FECHA
    del listado (p. ej. fecha de emisión elegida en Exportar láminas).
    """
    blocks = _read_revision_slots(sheet)
    cur = _current_rev_from_slots(blocks)
    if cur:
        rev_display = cur[u"revision"]
        emision = cur[u"descripcion"]
        fecha = (cur[u"fecha"] or u"").strip()
    else:
        rev_display = _sheet_rev_display_fallback(sheet, doc)
        emision = u""
        fecha = u""
    if not fecha:
        fecha = _revision_date_from_current_revision(sheet, doc)
    if not fecha:
        fecha = _last_non_empty_fecha_from_blocks(blocks)
    try:
        o = (
            (listado_fecha_column_override or u"").strip()
            if listado_fecha_column_override is not None
            else u""
        )
    except Exception:
        o = u""
    if o:
        fecha = o
    return {
        u"sheet_number": sheet.SheetNumber or u"",
        u"sheet_name": sheet.Name or u"",
        u"rev_display": rev_display,
        u"emision": emision,
        u"fecha": fecha,
        u"blocks": blocks,
    }


def collect_export_rows_for_sheets(
    doc, sheets, listado_fecha_column_override=None
):
    """
    Misma estructura que ``collect_export_rows`` pero solo las láminas indicadas,
    conservando el orden de la lista ``sheets``.
    """
    rows = []
    for sh in sheets:
        if sh is None:
            continue
        try:
            rows.append(
                export_row_for_sheet(
                    sh, doc, listado_fecha_column_override=listado_fecha_column_override
                )
            )
        except Exception:
            pass
    return rows


def collect_export_rows(doc):
    sheets = list(FilteredElementCollector(doc).OfClass(ViewSheet).ToElements())
    sheets.sort(key=lambda s: ((s.SheetNumber or u"").upper(), (s.Name or u"").upper()))
    return collect_export_rows_for_sheets(doc, sheets)


def project_metadata(doc):
    meta = {
        u"proyecto":    u"",
        u"cliente":     u"",
        u"orden":       u"",
        u"fecha_envio": datetime.date.today().strftime(u"%d.%m.%Y"),
    }
    try:
        pi = doc.ProjectInformation
        try:
            meta[u"proyecto"] = (pi.Name or u"").strip()
        except Exception:
            pass
        try:
            meta[u"cliente"] = (pi.ClientName or u"").strip()
        except Exception:
            pass
        try:
            meta[u"orden"] = (pi.Number or u"").strip()
        except Exception:
            pass
    except Exception:
        pass
    return meta


# ---------------------------------------------------------------------------
# Serialización JSON (sin dependencias externas)
# ---------------------------------------------------------------------------

def _esc(s):
    s = s.replace(u"\\", u"\\\\")
    s = s.replace(u'"', u'\\"')
    s = s.replace(u"\r", u"\\r")
    s = s.replace(u"\n", u"\\n")
    s = s.replace(u"\t", u"\\t")
    return s


def _to_json(v):
    if v is None:
        return u"null"
    if isinstance(v, bool):
        return u"true" if v else u"false"
    if isinstance(v, (int, float)):
        return unicode(v)
    if isinstance(v, (str, unicode)):
        return u'"' + _esc(v) + u'"'
    if isinstance(v, dict):
        return u"{" + u",".join(u'"' + _esc(k) + u'":' + _to_json(val) for k, val in v.items()) + u"}"
    if isinstance(v, list):
        return u"[" + u",".join(_to_json(i) for i in v) + u"]"
    return u'"' + _esc(unicode(v)) + u'"'


def build_json_payload(meta, rows):
    payload = {
        u"meta": meta,
        u"rows": rows,
        u"C": {
            u"sheet_listado":             SHEET_LISTADO,
            u"sheet_historial":           SHEET_HISTORIAL,
            u"first_data_row_listado":    FIRST_DATA_ROW_LISTADO,
            u"first_data_row_historial":  FIRST_DATA_ROW_HISTORIAL,
            u"block_start_col":           BLOCK_START_COL,
            u"block_step":                BLOCK_STEP,
            u"max_revision_blocks":       MAX_REVISION_BLOCKS,
            u"meta_row_proyecto":         META_ROW_PROYECTO,
            u"meta_row_cliente":          META_ROW_CLIENTE,
            u"meta_row_orden":            META_ROW_ORDEN,
            u"meta_row_fecha_env":        META_ROW_FECHA_ENV,
        },
    }
    return _to_json(payload)


# ---------------------------------------------------------------------------
# PowerShell script de escritura
# ---------------------------------------------------------------------------

_PS_TEMPLATE = u"""
param([string]$JsonFile, [string]$OutputXlsx)
$ErrorActionPreference = 'Stop'

function Col-Letter([int]$n) {
    $s = ''
    while ($n -gt 0) {
        $n--
        $s = [char](65 + ($n % 26)) + $s
        $n = [math]::Floor($n / 26)
    }
    return $s
}

$raw  = [System.IO.File]::ReadAllText($JsonFile, [System.Text.Encoding]::UTF8)
$data = $raw | ConvertFrom-Json
$meta = $data.meta
$rows = $data.rows
$C    = $data.C

$xl = New-Object -ComObject Excel.Application
$xl.Visible       = $false
$xl.DisplayAlerts = $false

try {
    $wb  = $xl.Workbooks.Open($OutputXlsx)
    $wsL = $wb.Worksheets.Item($C.sheet_listado)
    $wsH = $wb.Worksheets.Item($C.sheet_historial)

    # --- Metadatos LISTADO (valor en col B) ---
    $wsL.Cells.Item($C.meta_row_proyecto,  2).Value2 = [string]$meta.proyecto
    $wsL.Cells.Item($C.meta_row_cliente,   2).Value2 = [string]$meta.cliente
    $wsL.Cells.Item($C.meta_row_orden,     2).Value2 = [string]$meta.orden
    $wsL.Cells.Item($C.meta_row_fecha_env, 2).Value2 = [string]$meta.fecha_envio

    # --- Limpiar datos previos (sin tocar encabezados) ---
    $lastRow = $C.first_data_row_listado + 3000
    $wsL.Range("A$($C.first_data_row_listado):E$lastRow").ClearContents()

    $lastHRow = $C.first_data_row_historial + 3000
    $endHCol  = Col-Letter ($C.block_start_col + $C.block_step * $C.max_revision_blocks - 1)
    $wsH.Range("A$($C.first_data_row_historial):$($endHCol)$lastHRow").ClearContents()

    # Forzar formato texto en columnas con valores que pueden tener ceros iniciales:
    # LISTADO: col A (Número plano), col B (REV), col D (Fecha)
    $wsL.Columns.Item(1).NumberFormat = "@"
    $wsL.Columns.Item(2).NumberFormat = "@"
    $wsL.Columns.Item(4).NumberFormat = "@"
    # HISTORIAL: col A (Número plano), y cada col Revision y Fecha en los bloques
    $wsH.Columns.Item(1).NumberFormat = "@"
    for ($i = 0; $i -lt $C.max_revision_blocks; $i++) {
        $c0 = $C.block_start_col + $C.block_step * $i
        $wsH.Columns.Item($c0 + 0).NumberFormat = "@"   # Revision
        $wsH.Columns.Item($c0 + 5).NumberFormat = "@"   # Fecha
    }

    # --- HOJA: LISTADO DE PLANOS ---
    # Cols: A=Numero, B=REV, C=EMISION, D=FECHA, E=DESCRIPCION
    $r = $C.first_data_row_listado
    foreach ($item in $rows) {
        $wsL.Cells.Item($r, 1).Value2 = [string]$item.sheet_number
        $wsL.Cells.Item($r, 2).Value2 = [string]$item.rev_display
        $wsL.Cells.Item($r, 3).Value2 = [string]$item.emision
        $wsL.Cells.Item($r, 4).Value2 = [string]$item.fecha
        $wsL.Cells.Item($r, 5).Value2 = [string]$item.sheet_name
        $r++
    }

    # --- HOJA: HISTORIAL DE REVISIONES ---
    # Cols A=Numero, B=Contenido
    # Bloques de 6 cols desde col C: [Revision, Descripcion, Dibujo, Reviso, Aprobo, Fecha]
    $r = $C.first_data_row_historial
    $truncated = $false
    foreach ($item in $rows) {
        $wsH.Cells.Item($r, 1).Value2 = [string]$item.sheet_number
        $wsH.Cells.Item($r, 2).Value2 = [string]$item.sheet_name
        $blocks = @($item.blocks)
        if ($blocks.Count -gt $C.max_revision_blocks) {
            $blocks    = $blocks[0..($C.max_revision_blocks - 1)]
            $truncated = $true
        }
        for ($i = 0; $i -lt $blocks.Count; $i++) {
            $b  = $blocks[$i]
            $c0 = $C.block_start_col + $C.block_step * $i   # 1-based col index
            $wsH.Cells.Item($r, $c0 + 0).Value2 = [string]$b.revision
            $wsH.Cells.Item($r, $c0 + 1).Value2 = [string]$b.descripcion
            $wsH.Cells.Item($r, $c0 + 2).Value2 = [string]$b.dibujo
            $wsH.Cells.Item($r, $c0 + 3).Value2 = [string]$b.reviso
            $wsH.Cells.Item($r, $c0 + 4).Value2 = [string]$b.aprobo
            $wsH.Cells.Item($r, $c0 + 5).Value2 = [string]$b.fecha
        }
        $r++
    }

    $wb.Save()
    $wb.Close($false)

    if ($truncated) { Write-Output 'TRUNCATED' }
    Write-Output 'OK'
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
"""


# ---------------------------------------------------------------------------
# Ejecución vía PowerShell
# ---------------------------------------------------------------------------

def fill_template_excel(output_xlsx, doc, rows):
    import System
    import System.IO
    import System.Diagnostics
    import System.Text

    meta = project_metadata(doc)
    json_payload = build_json_payload(meta, rows)

    json_path = output_xlsx + u".bimtools_tmp.json"
    ps1_path  = output_xlsx + u".bimtools_tmp.ps1"
    try:
        System.IO.File.WriteAllText(json_path, json_payload, System.Text.Encoding.UTF8)
        System.IO.File.WriteAllText(ps1_path,  _PS_TEMPLATE,  System.Text.Encoding.UTF8)

        args = (
            u'-NoProfile -ExecutionPolicy Bypass -File "{0}" '
            u'-JsonFile "{1}" -OutputXlsx "{2}"'
        ).format(ps1_path, json_path, output_xlsx)

        psi = System.Diagnostics.ProcessStartInfo(u"powershell.exe", args)
        psi.CreateNoWindow = True
        psi.UseShellExecute = False
        psi.RedirectStandardOutput = True
        psi.RedirectStandardError = True
        psi.StandardOutputEncoding = System.Text.Encoding.UTF8
        psi.StandardErrorEncoding  = System.Text.Encoding.UTF8

        proc = System.Diagnostics.Process.Start(psi)
        if not proc.WaitForExit(120000):
            try:
                proc.Kill()
            except Exception:
                pass
            raise IOError(
                u"PowerShell tardó más de 120 s. "
                u"Compruebe que Excel no quedó bloqueado."
            )

        stdout = (proc.StandardOutput.ReadToEnd() or u"").strip()
        stderr = (proc.StandardError.ReadToEnd() or u"").strip()

        if proc.ExitCode != 0:
            detail = stderr if stderr else stdout
            raise IOError(
                u"PowerShell terminó con código {0}.\n{1}".format(proc.ExitCode, detail)
            )

        return u"TRUNCATED" in stdout

    finally:
        for p in (json_path, ps1_path):
            try:
                System.IO.File.Delete(p)
            except Exception:
                pass


def run_export(revit, template_path, output_path):
    doc = revit.ActiveUIDocument.Document
    rows = collect_export_rows(doc)
    shutil.copy2(template_path, output_path)
    truncated = fill_template_excel(output_path, doc, rows)
    return len(rows), truncated


def run_export_sheets(
    revit,
    template_path,
    output_path,
    sheets,
    listado_fecha_column_override=None,
):
    """
    Igual que ``run_export`` pero solo las láminas en ``sheets`` (``ViewSheet``),
    en el orden dado (p. ej. selección de Exportar láminas).

    ``listado_fecha_column_override``: texto para la columna FECHA de todas las filas
    (selección por fecha de emisión en la UI de exportación).
    """
    doc = revit.ActiveUIDocument.Document
    rows = collect_export_rows_for_sheets(
        doc, sheets, listado_fecha_column_override=listado_fecha_column_override
    )
    shutil.copy2(template_path, output_path)
    truncated = fill_template_excel(output_path, doc, rows)
    return len(rows), truncated
