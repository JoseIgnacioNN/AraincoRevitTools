# -*- coding: utf-8 -*-
"""
RevisionWindow — carga y gestión del XAML de la ventana Revisiones.

Responsabilidades:
- Cargar revision_window.xaml desde disco e inyectar BIMTOOLS_DARK_STYLES_XML.
- Conectar todos los event handlers de la UI al ViewModel.
- Posicionar la ventana en el monitor de Revit (misma línea visual que Exportar Láminas).
- Gestionar el DataGrid (scrollbars, filtro, selección por rango).
"""

from __future__ import print_function

try:
    unicode
except NameError:
    unicode = str

import os
import codecs

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")
clr.AddReference("System.Data")

from System import Boolean, DateTime, Int32, Int64
from System.Collections.Generic import List as ClrList
from System.Data import DataRowChangeEventHandler
from System.Globalization import CultureInfo
from System.Windows import RoutedEventHandler, Visibility
from System.Windows.Markup import XamlReader

from siguiente_revision.constants import (
    DESCRIPCIONES, DATE_FORMAT, DATE_DAYS_BEFORE, DATE_DAYS_AFTER,
    ISSUES_DIR,
)
from siguiente_revision.services import revision_service, sheet_service
from siguiente_revision.services.people_service import PERSONAS_FILE

_XAML_DIR = os.path.dirname(os.path.abspath(__file__))
_XAML_PATH = os.path.join(_XAML_DIR, "revision_window.xaml")

# ---------------------------------------------------------------------------
# Carga del XAML
# ---------------------------------------------------------------------------

def load_xaml():
    """
    Lee revision_window.xaml desde disco, aplica sustituciones de placeholders
    e instancia la Window con XamlReader.Parse().
    """
    try:
        from infra.bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
    except Exception:
        BIMTOOLS_DARK_STYLES_XML = u""

    with codecs.open(_XAML_PATH, "r", "utf-8") as f:
        xaml_str = f.read()

    xaml_str = xaml_str.replace(u"__BIMTOOLS_DARK_STYLES__", BIMTOOLS_DARK_STYLES_XML)

    return XamlReader.Parse(xaml_str)


# ---------------------------------------------------------------------------
# Scrollbars oscuros
# ---------------------------------------------------------------------------

def _apply_scrollbars(root_visual, resources_owner):
    try:
        from System.Windows.Controls.Primitives import ScrollBar
        from System.Windows.Media import VisualTreeHelper
        from System.Windows import FrameworkElement

        sb_type = clr.GetClrType(ScrollBar)
        st = None
        if resources_owner is not None:
            for _key in (u"ExpLamScrollBarDark", u"BimToolsScrollBarDark"):
                try:
                    st = resources_owner.TryFindResource(_key)
                except Exception:
                    st = None
                if st is not None:
                    break
            if st is None:
                try:
                    st = resources_owner.TryFindResource(sb_type)
                except Exception:
                    st = None
        if st is None:
            return

        def _walk(o, depth):
            if depth > 60 or o is None:
                return
            try:
                n = VisualTreeHelper.GetChildrenCount(o)
                for i in range(n):
                    ch = VisualTreeHelper.GetChild(o, i)
                    try:
                        if ch.GetType().Equals(sb_type):
                            try:
                                ch.ClearValue(FrameworkElement.StyleProperty)
                            except Exception:
                                pass
                            ch.Style = st
                    except Exception:
                        pass
                    _walk(ch, depth + 1)
            except Exception:
                pass

        _walk(root_visual, 0)
    except Exception:
        pass


def _schedule_scrollbars(win):
    try:
        from System import Action
        from System.Windows.Threading import DispatcherPriority

        def _go():
            _apply_scrollbars(win, win)

        _go()
        win.Dispatcher.BeginInvoke(DispatcherPriority.Loaded,      Action(_go))
        win.Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle,  Action(_go))
        win.Dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle, Action(_go))
    except Exception:
        try:
            _apply_scrollbars(win, win)
        except Exception:
            pass


def _schedule_grid_scrollbars(win, dg):
    try:
        from System import Action
        from System.Windows.Threading import DispatcherPriority

        def _go():
            _apply_scrollbars(dg, win)
            _apply_scrollbars(win, win)

        _go()
        win.Dispatcher.BeginInvoke(DispatcherPriority.Loaded,      Action(_go))
        win.Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle,  Action(_go))
        win.Dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle, Action(_go))
    except Exception:
        try:
            _apply_scrollbars(dg, win)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Posición y propietario de ventana (Exportar Láminas)
# ---------------------------------------------------------------------------

def _attach_revit_owner(win, revit_app):
    try:
        from System.Windows.Interop import WindowInteropHelper
        from infra.revit_wpf_window_position import revit_main_hwnd

        uiapp = None
        try:
            uiapp = revit_app if revit_app is not None else None
        except Exception:
            uiapp = None
        hwnd = revit_main_hwnd(uiapp)
        if hwnd is not None:
            try:
                if hwnd.ToInt64() != 0:
                    WindowInteropHelper(win).Owner = hwnd
            except Exception:
                WindowInteropHelper(win).Owner = hwnd
    except Exception:
        pass


def _prepare_window_bounds(win):
    try:
        from System import Double

        win.MaxWidth = Double.PositiveInfinity
        win.MaxHeight = Double.PositiveInfinity
    except Exception:
        pass


def _position_window(win, revit_app):
    try:
        from infra.revit_wpf_window_position import (
            bind_center_wpf_on_revit_monitor,
            bind_maximize_wpf_on_revit_monitor,
            position_wpf_window_center_on_monitor,
            revit_main_hwnd,
        )

        hwnd = revit_main_hwnd(revit_app)
        bind_center_wpf_on_revit_monitor(win, hwnd)
        bind_maximize_wpf_on_revit_monitor(win, hwnd)
        position_wpf_window_center_on_monitor(win, hwnd)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Datos del formulario — combos de personas y fechas
# ---------------------------------------------------------------------------

def fill_persona_combos(win, vm):
    """Rellena los combos Dibujó / Revisó / Aprobó desde el ViewModel."""
    cb_d = win.FindName("CbDibujo")
    cb_r = win.FindName("CbReviso")
    cb_a = win.FindName("CbAprobo")
    if cb_d is None or cb_r is None or cb_a is None:
        return

    prev_d = unicode(cb_d.SelectedItem or u"")
    prev_r = unicode(cb_r.SelectedItem or u"")
    prev_a = unicode(cb_a.SelectedItem or u"")

    cb_d.Items.Clear()
    cb_r.Items.Clear()
    cb_a.Items.Clear()

    dib_items, ing_items = vm.load_people()

    for x in dib_items:
        cb_d.Items.Add(x)
    for x in ing_items:
        cb_r.Items.Add(x)
        cb_a.Items.Add(x)

    # Restaurar selección anterior si el ítem sigue existiendo; si no, primero de la lista.
    def _restore(cb, prev):
        if prev and prev in [unicode(cb.Items[i]) for i in range(cb.Items.Count)]:
            cb.SelectedItem = prev
        elif cb.Items.Count > 0:
            cb.SelectedIndex = 0

    _restore(cb_d, prev_d)
    _restore(cb_r, prev_r)
    _restore(cb_a, prev_a)


def fill_fecha_combo(win):
    """Rellena el combo Fecha con los 26 valores y selecciona hoy."""
    today = DateTime.Today
    inv   = CultureInfo.InvariantCulture
    opts  = []
    for i in range(-DATE_DAYS_BEFORE, DATE_DAYS_AFTER + 1):
        d = today.AddDays(i)
        opts.append(d.ToString(DATE_FORMAT, inv))

    cb = win.FindName("CbFecha")
    if cb is None:
        return
    cb.Items.Clear()
    for fx in opts:
        cb.Items.Add(fx)
    cb.SelectedIndex = DATE_DAYS_BEFORE  # hoy = índice 5


# ---------------------------------------------------------------------------
# Grid: helpers de estado/selección
# ---------------------------------------------------------------------------

def _eid_from_table_cell(eid_val):
    from Autodesk.Revit.DB import ElementId
    try:
        return ElementId(Int64(int(eid_val)))
    except Exception:
        try:
            return ElementId(int(eid_val))
        except Exception:
            return ElementId.InvalidElementId


def _nullable_bool(wpf_nb):
    try:
        if wpf_nb is None:
            return False
        if hasattr(wpf_nb, "HasValue"):
            return bool(wpf_nb.HasValue and wpf_nb.Value)
        return unicode(wpf_nb).strip().lower() == u"true"
    except Exception:
        return False


def _row_selected(rv):
    if rv is None:
        return False
    try:
        if bool(rv.Row[u"Sel"]):
            return True
    except Exception:
        pass
    try:
        return _nullable_bool(rv[u"Sel"])
    except Exception:
        return False


def _row_enabled(rv):
    try:
        if rv is None:
            return True
        try:
            return bool(rv.Row[u"SelEnabled"])
        except Exception:
            pass
        try:
            return bool(rv[u"SelEnabled"])
        except Exception:
            pass
    except Exception:
        pass
    return True


def _grid_index_of_rowview(grid, drv):
    if grid is None or drv is None:
        return -1
    try:
        idx = int(grid.Items.IndexOf(drv))
        if idx >= 0:
            return idx
    except Exception:
        pass
    try:
        want = int(drv.Row[u"IdInt"])
    except Exception:
        try:
            want = int(drv[u"IdInt"])
        except Exception:
            return -1
    try:
        n = int(grid.Items.Count)
    except Exception:
        return -1
    for i in range(n):
        try:
            it = grid.Items[i]
            if int(it.Row[u"IdInt"]) == want:
                return i
        except Exception:
            continue
    return -1


# ---------------------------------------------------------------------------
# Filtro de búsqueda
# ---------------------------------------------------------------------------

def apply_buscar_filter(win, vm):
    tbl = vm.sheet_table
    if win is None or tbl is None:
        return
    dv = tbl.DefaultView
    try:
        t = unicode(win.FindName("TxtBuscar").Text).strip()
    except Exception:
        t = u""
    if not t:
        dv.RowFilter = u""
    else:
        esc = t.replace(u"'", u"''")
        dv.RowFilter = (
            u"[SheetNumber] LIKE '%{0}%' OR [SheetName] LIKE '%{0}%' OR "
            u"[Revision] LIKE '%{0}%' OR [NuevaRevision] LIKE '%{0}%'"
        ).format(esc)
    _sync_sel_all_hdr(win, vm)


def sync_buscar_wm(win):
    try:
        tb = win.FindName("TxtBuscar")
        wm = win.FindName("TxtBuscarWatermark")
        if tb is None or wm is None:
            return
        txt = unicode(tb.Text or u"").strip()
        foc = False
        try:
            foc = bool(tb.IsFocused)
        except Exception:
            pass
        show = not txt and not foc
        wm.Visibility = Visibility.Visible if show else Visibility.Collapsed
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Estado del listado
# ---------------------------------------------------------------------------

def refresh_estado(win, vm):
    tbl = vm.sheet_table
    tb  = win.FindName("TxtEstadoLaminas") if win is not None else None
    if tbl is None or tb is None:
        return
    n = int(tbl.Rows.Count)
    ns = 0
    for i in range(n):
        try:
            if bool(tbl.Rows[i][u"Sel"]):
                ns += 1
        except Exception:
            pass
    try:
        tb.Text = u"{0} láminas  |  {1} seleccionadas".format(n, ns)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Columna «Nueva revisión» y SelEnabled
# ---------------------------------------------------------------------------

def refresh_nueva_revision_column(doc, vm):
    tbl = vm.sheet_table
    if tbl is None or doc is None:
        return
    from Autodesk.Revit.DB import ViewSheet
    ordered = revision_service.get_ordered_revision_ids(doc)
    ti_rev0 = revision_service.index_of_revision_display_number(doc, ordered, u"0")
    n = int(tbl.Rows.Count)
    for i in range(n):
        row = tbl.Rows[i]
        try:
            eid_val = row[u"IdInt"]
        except Exception:
            try:
                row[u"NuevaRevision"] = u"\u2014"
                row[u"SelEnabled"]    = True
            except Exception:
                pass
            continue
        try:
            el = doc.GetElement(_eid_from_table_cell(eid_val))
        except Exception:
            el = None
        if not isinstance(el, ViewSheet):
            try:
                row[u"NuevaRevision"] = u"\u2014"
                row[u"SelEnabled"]    = True
            except Exception:
                pass
            continue
        emit_rev0 = vm.emit_rev0
        try:
            row[u"NuevaRevision"] = revision_service.preview_next_revision(
                doc, el, ordered, emit_rev0, ti_rev0
            )
        except Exception:
            try:
                row[u"NuevaRevision"] = u"\u2014"
            except Exception:
                pass
        try:
            en = revision_service.compute_sel_enabled(doc, el, ordered, emit_rev0, ti_rev0)
            row[u"SelEnabled"] = bool(en)
            if not en:
                row[u"Sel"] = False
        except Exception:
            try:
                row[u"SelEnabled"] = True
            except Exception:
                pass


def refresh_nueva_revision_from_ui(win, vm, doc):
    if doc is None or vm.sheet_table is None:
        return
    r_punt = win.FindName("RadRevisionPuntual")
    if r_punt is not None:
        vm._emit_rev0 = _nullable_bool(r_punt.IsChecked)
    refresh_nueva_revision_column(doc, vm)
    refresh_estado(win, vm)
    _sync_sel_all_hdr(win, vm)


# ---------------------------------------------------------------------------
# Checkbox «marcar todas»
# ---------------------------------------------------------------------------

def _find_hdr_checkbox(original_source):
    try:
        from System.Windows.Controls import CheckBox
        from System.Windows.Media import VisualTreeHelper
    except Exception:
        return None
    d = original_source
    for _ in range(48):
        if d is None:
            break
        try:
            if isinstance(d, CheckBox):
                tg = d.Tag
                if tg is not None and unicode(str(tg)).strip() == u"HdrSelectAll":
                    return d
        except Exception:
            pass
        try:
            d = VisualTreeHelper.GetParent(d)
        except Exception:
            break
    return None


def _wire_select_all_hdr(win, vm):
    try:
        from System.Windows.Media import VisualTreeHelper
        from System.Windows.Controls import CheckBox
    except Exception:
        return
    grid = vm._grid
    if grid is None:
        return

    def walk_find(comp):
        try:
            n = VisualTreeHelper.GetChildrenCount(comp)
        except Exception:
            return None
        for i in range(int(n)):
            try:
                ch = VisualTreeHelper.GetChild(comp, i)
            except Exception:
                continue
            try:
                if isinstance(ch, CheckBox):
                    tg = ch.Tag
                    if tg is not None and unicode(str(tg)) == u"HdrSelectAll":
                        return ch
            except Exception:
                pass
            got = walk_find(ch)
            if got is not None:
                return got
        return None

    chk = walk_find(grid)
    if chk is not None:
        vm._hdr_chk = chk


def _sync_sel_all_hdr(win, vm):
    chk = vm._hdr_chk
    if chk is None:
        try:
            _wire_select_all_hdr(win, vm)
        except Exception:
            pass
        chk = vm._hdr_chk
    if chk is None:
        return
    tbl = vm.sheet_table
    if tbl is None:
        return
    dv = tbl.DefaultView
    try:
        n = int(dv.Count)
    except Exception:
        n = 0
    vm._syncing_sel_all = True
    try:
        try:
            chk.IsThreeState = True
        except Exception:
            pass
        if n == 0:
            chk.IsChecked = False
            return
        n_el = 0
        n_sel = 0
        for i in range(n):
            try:
                rv = dv[i]
                if not _row_enabled(rv):
                    continue
                n_el += 1
                if _row_selected(rv):
                    n_sel += 1
            except Exception:
                pass
        if n_el == 0:
            chk.IsChecked = False
            return
        if n_sel == 0:
            chk.IsChecked = False
        elif n_sel == n_el:
            chk.IsChecked = True
        else:
            chk.IsChecked = None
    finally:
        vm._syncing_sel_all = False


def _apply_select_all_visible(win, vm):
    tbl  = vm.sheet_table
    grid = vm._grid
    if tbl is None:
        return
    dv = tbl.DefaultView
    try:
        n = int(dv.Count)
    except Exception:
        return
    try:
        from System import Boolean as ClrBool
    except Exception:
        return
    if n == 0:
        refresh_estado(win, vm)
        _sync_sel_all_hdr(win, vm)
        return
    n_el = 0
    n_sel = 0
    for i in range(n):
        try:
            rv = grid.Items[i] if grid is not None else dv[i]
            if not _row_enabled(rv):
                continue
            n_el += 1
            if _row_selected(rv):
                n_sel += 1
        except Exception:
            pass
    new_val = True
    if n_el > 0 and n_sel == n_el:
        new_val = False
    vm._syncing_sel_all = True
    try:
        for i in range(n):
            try:
                rv = grid.Items[i] if grid is not None else dv[i]
                if not _row_enabled(rv):
                    continue
                if grid is not None:
                    grid.Items[i][u"Sel"] = ClrBool(new_val)
                else:
                    dv[i][u"Sel"] = ClrBool(new_val)
            except Exception:
                pass
    finally:
        vm._syncing_sel_all = False
    refresh_estado(win, vm)
    _sync_sel_all_hdr(win, vm)


# ---------------------------------------------------------------------------
# Handlers del DataGrid
# ---------------------------------------------------------------------------

def _on_grid_loaded(win, vm, sender, args):
    g = sender
    try:
        if g.Columns.Count > 0:
            g.Columns[g.Columns.Count - 1].Visibility = Visibility.Collapsed
    except Exception:
        pass
    try:
        from System import Action
        from System.Windows.Threading import DispatcherPriority

        def _wire():
            try:
                _wire_select_all_hdr(win, vm)
                _sync_sel_all_hdr(win, vm)
            except Exception:
                pass

        _wire()
        g.Dispatcher.BeginInvoke(DispatcherPriority.Loaded,     Action(_wire))
        g.Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, Action(_wire))
    except Exception:
        try:
            _wire_select_all_hdr(win, vm)
            _sync_sel_all_hdr(win, vm)
        except Exception:
            pass
    try:
        _schedule_grid_scrollbars(win, g)
    except Exception:
        pass


def _on_grid_preview_mouse(win, vm, sender, e):
    try:
        from System.Windows.Controls import DataGridCell, DataGridCheckBoxColumn
        from System.Windows.Input import Keyboard, ModifierKeys
        from System.Windows.Media import VisualTreeHelper
        from System import Boolean as ClrBool
    except Exception:
        return
    tbl  = vm.sheet_table
    grid = vm._grid
    if tbl is None or grid is None:
        return

    if _find_hdr_checkbox(e.OriginalSource) is not None:
        if not vm._syncing_sel_all:
            _apply_select_all_visible(win, vm)
        try:
            e.Handled = True
        except Exception:
            pass
        return

    d = e.OriginalSource
    cell = None
    for _ in range(40):
        if d is None:
            break
        try:
            if isinstance(d, DataGridCell):
                cell = d
                break
        except Exception:
            pass
        try:
            d = VisualTreeHelper.GetParent(d)
        except Exception:
            break
    if cell is None:
        return
    try:
        if not isinstance(cell.Column, DataGridCheckBoxColumn):
            return
    except Exception:
        return
    drv = cell.DataContext
    if drv is None:
        return
    if not _row_enabled(drv):
        try:
            e.Handled = True
        except Exception:
            pass
        return
    idx = _grid_index_of_rowview(grid, drv)
    if idx < 0:
        return

    shift_down = False
    try:
        shift_down = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift
    except Exception:
        pass
    anchor = vm._sel_anchor

    if shift_down and anchor is not None:
        try:
            cur_on  = _row_selected(drv)
            new_val = not cur_on
            i0 = min(int(anchor), idx)
            i1 = max(int(anchor), idx)
            for j in range(i0, i1 + 1):
                try:
                    rvj = grid.Items[j]
                    if not _row_enabled(rvj):
                        continue
                    grid.Items[j][u"Sel"] = ClrBool(new_val)
                except Exception:
                    pass
        except Exception:
            pass
        try:
            e.Handled = True
        except Exception:
            pass
        refresh_estado(win, vm)
        _sync_sel_all_hdr(win, vm)
        return

    def _bulk():
        try:
            cur_on  = _row_selected(drv)
            new_val = not cur_on
            vm._sel_anchor = idx
            indices = []
            seen    = set()
            try:
                coll = grid.SelectedItems
                if coll is not None:
                    for i in range(int(coll.Count)):
                        try:
                            it = coll[i]
                            j  = _grid_index_of_rowview(grid, it)
                            if j >= 0 and j not in seen:
                                seen.add(j)
                                indices.append(j)
                        except Exception:
                            pass
            except Exception:
                pass
            if idx not in seen:
                indices.append(idx)
            for j in indices:
                try:
                    rvj = grid.Items[j]
                    if not _row_enabled(rvj):
                        continue
                    grid.Items[j][u"Sel"] = ClrBool(new_val)
                except Exception:
                    pass
        except Exception:
            pass
        refresh_estado(win, vm)
        _sync_sel_all_hdr(win, vm)

    try:
        from System import Action
        from System.Windows.Threading import DispatcherPriority
        grid.Dispatcher.BeginInvoke(DispatcherPriority.Input, Action(_bulk))
    except Exception:
        _bulk()
    try:
        e.Handled = True
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Rebind (botón Actualizar)
# ---------------------------------------------------------------------------

def rebind_sheets_grid(win, vm, doc, reset_buscar=False):
    sheets_all = sheet_service.collect_sheets(doc)
    tbl        = sheet_service.build_selection_table(doc, sheets_all)
    _detach_row_changed(vm)
    vm.sheet_table = tbl
    vm._row_delegate = _make_row_delegate(win, vm)
    tbl.RowChanged += vm._row_delegate
    grid = vm._grid
    grid.ItemsSource = tbl.DefaultView
    if reset_buscar:
        try:
            win.FindName("TxtBuscar").Text = u""
        except Exception:
            pass
    apply_buscar_filter(win, vm)
    sync_buscar_wm(win)
    refresh_estado(win, vm)
    refresh_nueva_revision_from_ui(win, vm, doc)


def _detach_row_changed(vm):
    tbl = vm.sheet_table
    d   = vm._row_delegate
    if tbl is not None and d is not None:
        try:
            tbl.RowChanged -= d
        except Exception:
            pass


def _make_row_delegate(win, vm):
    def _fn(sender, args):
        if vm._syncing_sel_all:
            return
        refresh_estado(win, vm)
        _sync_sel_all_hdr(win, vm)
    return DataRowChangeEventHandler(_fn)


# ---------------------------------------------------------------------------
# Punto de entrada: crear y cablear la ventana completa
# ---------------------------------------------------------------------------

def build_and_wire(win, vm, doc, uidoc, revit_app, idx_rev0, sheets_all):
    """
    Cablea todos los event handlers de la ventana WPF al ViewModel.

    Args:
        win:        Window WPF ya instanciada vía load_xaml().
        vm:         RevisionViewModel.
        doc:        Revit Document.
        uidoc:      UIDocument (para gestionar personas con owner).
        revit_app:  Application de Revit.
        idx_rev0:   Índice de revisión 0 en el proyecto (-1 si no existe).
        sheets_all: Lista de ViewSheet inicial.
    """
    vm._win = win
    vm._has_revision_zero = idx_rev0 >= 0

    # --- Combos formulario ---
    for d in DESCRIPCIONES:
        win.FindName("CbDescripcion").Items.Add(d)
    win.FindName("CbDescripcion").SelectedIndex = 0
    fill_persona_combos(win, vm)
    fill_fecha_combo(win)

    # --- Deshabilitar «Revisión 0» si no existe ---
    if idx_rev0 < 0:
        rp = win.FindName("RadRevisionPuntual")
        if rp is not None:
            try:
                rp.IsEnabled = False
            except Exception:
                pass

    # --- Grid de láminas ---
    grid_sh = win.FindName("GridSheets")
    vm._grid = grid_sh
    tbl      = sheet_service.build_selection_table(doc, sheets_all)
    _detach_row_changed(vm)
    vm.sheet_table     = tbl
    vm._row_delegate   = _make_row_delegate(win, vm)
    tbl.RowChanged    += vm._row_delegate
    grid_sh.ItemsSource = tbl.DefaultView
    apply_buscar_filter(win, vm)
    sync_buscar_wm(win)
    refresh_estado(win, vm)
    refresh_nueva_revision_from_ui(win, vm, doc)

    # --- Gestionar personas ---
    def _on_gestionar_personas(_s, _e):
        try:
            from ui.gestionar_personas_wpf import GestionarPersonasDialog, load_personas_list
            from System.Collections.ObjectModel import ObservableCollection
            from System.IO import Directory
        except Exception:
            from Autodesk.Revit.UI import TaskDialog
            TaskDialog.Show(u"Revisiones", u"No se pudo cargar el módulo de gestión de personas.")
            return
        oc = ObservableCollection[object]()
        for p in load_personas_list(PERSONAS_FILE):
            oc.Add(p)
        try:
            Directory.CreateDirectory(ISSUES_DIR)
        except Exception:
            pass
        prev_top = None
        try:
            prev_top = win.Topmost
            win.Topmost = False
        except Exception:
            pass
        try:
            GestionarPersonasDialog(
                oc,
                ISSUES_DIR,
                PERSONAS_FILE,
                uidoc=uidoc,
                revit_app=revit_app,
                owner=win,
            )
        except Exception as ex:
            from Autodesk.Revit.UI import TaskDialog
            TaskDialog.Show(
                u"Revisiones",
                u"No se pudo abrir el directorio de personas:\n\n{0}".format(str(ex)),
            )
        finally:
            if prev_top is not None:
                try:
                    win.Topmost = prev_top
                except Exception:
                    pass
        fill_persona_combos(win, vm)

    btn_gest = win.FindName("BtnGestionarPersonas")
    if btn_gest is not None:
        btn_gest.Click += RoutedEventHandler(_on_gestionar_personas)

    # --- Modo Revisión ---
    def _on_destino_changed(_snd, _evt):
        refresh_nueva_revision_from_ui(win, vm, doc)

    rad_auto = win.FindName("RadRevAutomatica")
    rad_punt = win.FindName("RadRevisionPuntual")
    if rad_auto is not None:
        rad_auto.Checked += RoutedEventHandler(_on_destino_changed)
    if rad_punt is not None:
        rad_punt.Checked += RoutedEventHandler(_on_destino_changed)

    # --- Búsqueda ---
    tb_search = win.FindName("TxtBuscar")
    if tb_search is not None:
        from System.Windows.Controls import TextChangedEventHandler
        tb_search.TextChanged += TextChangedEventHandler(
            lambda s, e: (sync_buscar_wm(win), apply_buscar_filter(win, vm))
        )
        tb_search.GotFocus  += RoutedEventHandler(lambda s, e: sync_buscar_wm(win))
        tb_search.LostFocus += RoutedEventHandler(lambda s, e: sync_buscar_wm(win))

    # --- Botón Actualizar ---
    btn_ref = win.FindName("BtnRefrescar")
    if btn_ref is not None:
        btn_ref.Click += RoutedEventHandler(
            lambda s, e: rebind_sheets_grid(win, vm, doc, reset_buscar=True)
        )

    # --- Grid eventos ---
    grid_sh.Loaded += RoutedEventHandler(
        lambda s, e: _on_grid_loaded(win, vm, s, e)
    )
    try:
        from System.Windows.Input import MouseButtonEventHandler
        grid_sh.PreviewMouseLeftButtonDown += MouseButtonEventHandler(
            lambda s, e: _on_grid_preview_mouse(win, vm, s, e)
        )
    except Exception:
        pass

    # --- Botones OK / Cancelar ---
    btn_cancel = win.FindName("BtnCancel")
    if btn_cancel is not None:
        btn_cancel.Click += RoutedEventHandler(lambda s, e: vm.cancel(win))

    btn_ok = win.FindName("BtnOk")
    if btn_ok is not None:
        btn_ok.Click += RoutedEventHandler(lambda s, e: vm.accept(win))

    # --- Teclado: Escape cierra ---
    try:
        from System.Windows.Input import (
            ApplicationCommands, CommandBinding, ExecutedRoutedEventHandler,
            KeyBinding, Key, ModifierKeys,
        )
        win.CommandBindings.Add(
            CommandBinding(
                ApplicationCommands.Close,
                ExecutedRoutedEventHandler(lambda _s, _e: vm.cancel(win)),
            )
        )
        win.InputBindings.Add(
            KeyBinding(
                ApplicationCommands.Close,
                Key.Escape,
                getattr(ModifierKeys, "None", 0),
            )
        )
    except Exception:
        pass

    def _on_loaded_all(_s, _e):
        _attach_revit_owner(win, revit_app)
        _prepare_window_bounds(win)
        _position_window(win, revit_app)
        try:
            _schedule_scrollbars(win)
        except Exception:
            pass
        try:
            gd = win.FindName("GridSheets")
            if gd is not None:
                _schedule_grid_scrollbars(win, gd)
        except Exception:
            pass

    win.Loaded += RoutedEventHandler(_on_loaded_all)
