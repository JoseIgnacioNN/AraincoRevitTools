# -*- coding: utf-8 -*-
"""
Cargar familias desde carpeta — carga masiva de .rfa con sobrescritura.

Revit 2024+ | pyRevit | IronPython 3.4

Flujo:
  1. Seleccionar carpeta con FolderBrowserDialog.
  2. Buscar archivos .rfa en el primer nivel de la carpeta.
  3. Cargar cada familia; si ya existe en el proyecto, sobrescribirla.
  4. Resumen de cargadas, omitidas y errores.
"""

from __future__ import print_function

import os

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System.Windows.Forms")

from Autodesk.Revit.DB import (
    Family,
    FamilySource,
    IFamilyLoadOptions,
    Transaction,
)
from System.Windows.Forms import DialogResult, FolderBrowserDialog, NativeWindow

from pyrevit import forms

_TITULO = u"Arainco: Cargar familias desde carpeta"
_RFA_EXT = u".rfa"


class _OverwriteFamilyLoadOptions(IFamilyLoadOptions):
    """Sobrescribe familias existentes y sus valores de parámetros."""

    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues = True
        return True

    def OnSharedFamilyFound(
        self, sharedFamily, familyInUse, source, overwriteParameterValues
    ):
        source = FamilySource.Project
        overwriteParameterValues = True
        return True


class _RevitDialogOwner(NativeWindow):
    """Propietario Win32 para modalizar diálogos bajo Revit."""


def _pick_folder(uiapp):
    """Abre FolderBrowserDialog modal bajo la ventana principal de Revit."""
    dlg = FolderBrowserDialog()
    dlg.Description = u"Seleccione la carpeta con archivos de familia (.rfa)."
    dlg.ShowNewFolderButton = False

    owner = None
    try:
        hwnd = uiapp.MainWindowHandle
        if hwnd is not None and hwnd.ToInt64() != 0:
            owner = _RevitDialogOwner()
            owner.AssignHandle(hwnd)
    except Exception:
        owner = None

    try:
        if owner is not None:
            result = dlg.ShowDialog(owner)
        else:
            result = dlg.ShowDialog()
    finally:
        if owner is not None:
            try:
                owner.ReleaseHandle()
            except Exception:
                pass

    if result != DialogResult.OK:
        return None

    try:
        return dlg.SelectedPath
    except Exception:
        return None


def _list_rfa_files(folder_path):
    """Devuelve rutas absolutas de .rfa en el primer nivel de la carpeta."""
    files = []
    try:
        names = os.listdir(folder_path)
    except Exception:
        return files

    for name in names:
        if not name.lower().endswith(_RFA_EXT):
            continue
        full_path = os.path.join(folder_path, name)
        if os.path.isfile(full_path):
            files.append(os.path.abspath(full_path))

    files.sort(key=lambda p: os.path.basename(p).lower())
    return files


def _load_family(doc, path, load_options):
    """Carga una familia desde disco. Devuelve (ok, mensaje_error)."""
    fam_ref = clr.Reference[Family]()
    try:
        loaded = doc.LoadFamily(path, load_options, fam_ref)
    except Exception as ex:
        try:
            err = unicode(ex)
        except NameError:
            err = str(ex)
        return False, err

    if not loaded:
        return False, u"LoadFamily devolvió False."
    return True, None


def _ejecutar(uidoc):
    doc = uidoc.Document

    if doc.IsFamilyDocument:
        forms.alert(
            u"Esta herramienta no aplica a documentos de familia (.rfa).",
            title=_TITULO,
        )
        return

    folder = _pick_folder(uidoc.Application)
    if not folder:
        return

    rfa_files = _list_rfa_files(folder)
    if not rfa_files:
        forms.alert(
            u"No se encontraron archivos .rfa en la carpeta:\n\n{0}".format(folder),
            title=_TITULO,
        )
        return

    load_options = _OverwriteFamilyLoadOptions()
    loaded_ok = []
    failed = []

    t = Transaction(doc, _TITULO)
    t.Start()
    try:
        for path in rfa_files:
            name = os.path.basename(path)
            ok, reason = _load_family(doc, path, load_options)
            if ok:
                loaded_ok.append(name)
            else:
                failed.append((name, reason))
        t.Commit()
    except Exception as ex:
        t.RollBack()
        try:
            msg = unicode(ex)
        except NameError:
            msg = str(ex)
        forms.alert(u"Error al cargar familias:\n\n{0}".format(msg), title=_TITULO)
        raise

    msg = u"Carpeta: {0}\n\nFamilias cargadas: {1} de {2}.".format(
        folder,
        len(loaded_ok),
        len(rfa_files),
    )
    if failed:
        msg += u"\n\nCon error ({0}):".format(len(failed))
        for name, reason in failed[:12]:
            msg += u"\n  • {0}: {1}".format(name, reason)
        if len(failed) > 12:
            msg += u"\n  … y {0} más.".format(len(failed) - 12)

    forms.alert(msg, title=_TITULO)


def run(revit):
    """Entrada pyRevit: selección de carpeta y carga de familias."""
    uidoc = revit.ActiveUIDocument
    if uidoc is None:
        forms.alert(u"No hay documento activo.", title=_TITULO)
        return
    _ejecutar(uidoc)
