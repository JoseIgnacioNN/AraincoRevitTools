# -*- coding: utf-8 -*-
__title__ = u"Buscar en\nProyecto"

import os
import sys

_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
if _pushbutton_dir not in sys.path:
    sys.path.insert(0, _pushbutton_dir)

import bimtools_access_bootstrap as _bimtools_access

if not _bimtools_access.require_tool_access(__file__, __revit__, __title__):
    raise SystemExit

from pyrevit import revit, DB, forms
from System.Collections.Generic import List
doc = revit.doc
uidoc = revit.uidoc

def quitar_tildes(texto):
    """Elimina las tildes de un texto para hacer la búsqueda más flexible."""
    if not texto: return ""
    reemplazos = (("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u"))
    texto = texto.lower()
    for a, b in reemplazos:
        texto = texto.replace(a, b)
    return texto

def obtener_texto_elemento(elemento):
    """Extrae el texto de un elemento dependiendo de si es Nota, Etiqueta o Cota."""
    texto = ""
    if isinstance(elemento, DB.TextNote):
        texto = elemento.Text
    elif isinstance(elemento, DB.IndependentTag):
        try:
            texto = elemento.TagText
        except AttributeError:
            pass 
    elif hasattr(DB.Architecture, "RoomTag") and isinstance(elemento, DB.Architecture.RoomTag):
        try:
            room = elemento.Room
            if room:
                texto = "{} {}".format(room.Number, room.Name)
        except:
            pass
    elif isinstance(elemento, DB.Dimension):
        try:
            partes = [elemento.Prefix, elemento.Suffix, elemento.TextAbove, elemento.TextBelow, elemento.ValueOverride]
            texto = " ".join([p for p in partes if p and isinstance(p, str)])
        except AttributeError:
            pass
    elif isinstance(elemento, DB.Viewport):
        view_id = elemento.ViewId
        if view_id != DB.ElementId.InvalidElementId:
            try:
                view = doc.GetElement(view_id)
                if view:
                    param_titulo = view.get_Parameter(DB.BuiltInParameter.VIEW_DESCRIPTION)
                    titulo_plano = param_titulo.AsString() if param_titulo else ""
                    texto = titulo_plano if titulo_plano else view.Name
            except:
                pass
            
    return texto if texto else ""

def ajustar_camara_a_vista(vista):
    """Hace un 'Zoom To Fit' para ver la vista completa."""
    for uv in uidoc.GetOpenUIViews():
        if uv.ViewId == vista.Id:
            uv.ZoomToFit()
            break

class SearchForm(forms.WPFWindow):
    def __init__(self, xaml_file_name):
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.search_string = ""
        self.match_case = False
        self.exact_match = False
        self.SearchTextBox.Focus()

    def search_click(self, sender, args):
        self.search_string = self.SearchTextBox.Text
        self.match_case = self.MatchCaseCheckBox.IsChecked
        self.exact_match = self.WholeWordCheckBox.IsChecked
        self.Close()

def buscar_texto_global():
    # 1. Mostrar la ventana de búsqueda avanzada
    import os
    import re
    xaml_path = os.path.join(os.path.dirname(__file__), 'SearchWindow.xaml')
    form = SearchForm(xaml_path)
    form.show_dialog()
    
    search_string = form.search_string
    if not search_string or not search_string.strip(): return
    
    match_case = form.match_case
    exact_match = form.exact_match

    # Preparar el texto de búsqueda
    busqueda_original = search_string.strip()
    if not match_case:
        busqueda_procesada = quitar_tildes(busqueda_original.lower())
    else:
        busqueda_procesada = busqueda_original
        
    patron_regex = ""
    if exact_match:
        patron_regex = r'(?:^|\s)' + re.escape(busqueda_procesada) + r'(?:\s|$)'

    # 2. Recopilar TODOS los textos, etiquetas y cotas del proyecto completo
    # Al no pasarle el ID de una vista al FilteredElementCollector, busca en todo el documento
    text_notes = DB.FilteredElementCollector(doc).OfClass(DB.TextNote).WhereElementIsNotElementType().ToElements()
    tags = DB.FilteredElementCollector(doc).OfClass(DB.IndependentTag).WhereElementIsNotElementType().ToElements()
    dimensions = DB.FilteredElementCollector(doc).OfClass(DB.Dimension).WhereElementIsNotElementType().ToElements()
    viewports = DB.FilteredElementCollector(doc).OfClass(DB.Viewport).WhereElementIsNotElementType().ToElements()
    
    try:
        room_tags = DB.FilteredElementCollector(doc).OfClass(DB.Architecture.RoomTag).WhereElementIsNotElementType().ToElements()
    except:
        room_tags = []

    import itertools
    todos_los_elementos = itertools.chain(text_notes, tags, room_tags, dimensions, viewports)
    
    # Diccionario para agrupar los resultados por el ID de la vista a la que pertenecen
    resultados_dict = {}
    cache_vistas_visibles = {}

    # 3. Filtrar y agrupar por vista
    for elemento in todos_los_elementos:
        texto_elemento = obtener_texto_elemento(elemento)
        if not texto_elemento: continue
        
        # Procesar texto objetivo según opciones
        objetivo = texto_elemento if match_case else quitar_tildes(texto_elemento.lower())
        
        coincide = False
        if exact_match:
            if re.search(patron_regex, objetivo, re.UNICODE):
                coincide = True
        else:
            if busqueda_procesada in objetivo:
                coincide = True
                
        if coincide:
            # OwnerViewId nos dice en qué vista "vive" este elemento anotativo
            view_id = elemento.OwnerViewId
            if view_id != DB.ElementId.InvalidElementId:
                # Comprobar si el elemento está realmente visible (no oculto por Crop View ni oculto manualmente)
                if not isinstance(elemento, DB.Viewport):
                    if view_id.IntegerValue not in cache_vistas_visibles:
                        try:
                            # Obtener todos los IDs de elementos que Revit renderiza/ve en esta vista específica
                            visibles = DB.FilteredElementCollector(doc, view_id).ToElementIds()
                            cache_vistas_visibles[view_id.IntegerValue] = set([e.IntegerValue for e in visibles])
                        except:
                            cache_vistas_visibles[view_id.IntegerValue] = set()
                    
                    if elemento.Id.IntegerValue not in cache_vistas_visibles[view_id.IntegerValue]:
                        continue # El elemento fue recortado o está oculto en la vista
                
                if view_id not in resultados_dict:
                    resultados_dict[view_id] = []
                resultados_dict[view_id].append(elemento.Id)

    # 3.5 Filtrar solo las vistas que están en láminas (o que son láminas)
    vistas_en_laminas_ids = set(vp.ViewId for vp in DB.FilteredElementCollector(doc).OfClass(DB.Viewport).ToElements())
    
    # Agregar las vistas primarias de las dependientes que estén en láminas
    for v_id in list(vistas_en_laminas_ids):
        vista_vp = doc.GetElement(v_id)
        if vista_vp and hasattr(vista_vp, 'GetPrimaryViewId'):
            primary_id = vista_vp.GetPrimaryViewId()
            if primary_id != DB.ElementId.InvalidElementId:
                vistas_en_laminas_ids.add(primary_id)

    resultados_dict_filtrado = {}
    for v_id, ids in resultados_dict.items():
        vista = doc.GetElement(v_id)
        if not vista: continue
        
        # Mantener si es una lámina o si está en una lámina
        if isinstance(vista, DB.ViewSheet) or (v_id in vistas_en_laminas_ids):
            resultados_dict_filtrado[v_id] = ids

    resultados_dict = resultados_dict_filtrado

    # 4. Procesar resultados
    if not resultados_dict:
        forms.alert("No se encontró '{}' en láminas del proyecto.".format(search_string), title="Sin resultados")
        return

    # 5. Lógica de selección de vista
    if len(resultados_dict) == 1:
        # Solo se encontró en una vista
        vista_id = list(resultados_dict.keys())[0]
        vista_destino = doc.GetElement(vista_id)
        ids_encontrados = resultados_dict[vista_id]
    else:
        # Se encontró en múltiples vistas, armar el menú
        opciones_dict = {}
        for v_id, ids in resultados_dict.items():
            vista = doc.GetElement(v_id)
            if not vista: continue  # Seguridad: omitir si la vista es nula o inválida
            
            cantidad = len(ids)
            texto_coincidencia = "1 coincidencia" if cantidad == 1 else "{} coincidencias".format(cantidad)
            
            if isinstance(vista, DB.ViewSheet):
                etiqueta_menu = "📝 LÁMINA: {} - {} ({})".format(vista.SheetNumber, vista.Name, texto_coincidencia)
            else:
                nombre_vista = vista.Name
                param_titulo = vista.get_Parameter(DB.BuiltInParameter.VIEW_DESCRIPTION)
                titulo_plano = param_titulo.AsString() if param_titulo else ""
                
                if titulo_plano:
                    # Mostrar solo el título en la lámina como pidió el usuario
                    etiqueta_menu = "{} ({})".format(titulo_plano, texto_coincidencia)
                else:
                    etiqueta_menu = "{} ({})".format(nombre_vista, texto_coincidencia)

            opciones_dict[etiqueta_menu] = (vista, ids)

        # Calcular ancho dinámico basado en el texto más largo
        max_longitud = max(len(k) for k in opciones_dict.keys())
        ancho_calculado = max(200, max_longitud * 7 + 50) # 7 píxeles por letra

        # Mostrar el menú al usuario
        seleccion = forms.SelectFromList.show(
            opciones_dict.keys(),
            title="Resultados",
            width=ancho_calculado,
            button_name="Seleccionar"
        )
        
        if not seleccion: return
        
        vista_destino = opciones_dict[seleccion][0]
        ids_encontrados = opciones_dict[seleccion][1]

    # 6. Ejecutar salto, selección y centrado
    if doc.ActiveView.Id != vista_destino.Id:
        uidoc.ActiveView = vista_destino
        uidoc.RefreshActiveView() # Asegurar que la interfaz registre el cambio de vista

    element_ids = List[DB.ElementId](ids_encontrados)
    uidoc.Selection.SetElementIds(element_ids)
    ajustar_camara_a_vista(vista_destino)

# Ejecutar el script
buscar_texto_global()