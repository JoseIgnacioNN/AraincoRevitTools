# -*- coding: utf-8 -*-
import os, codecs
from pyrevit import forms

form_top = None
form_left = None
RUTINAS = ["Armadura sobre apoyo.py", "Armadura de borde.py", "Malla en 1 dirección.py", "Malla en 2 direcciones.py"]
rutina_actual_idx = 0 # La rutina que se abrirá por defecto

while True:
    nueva_rutina = RUTINAS[rutina_actual_idx]
    script_path = os.path.join(os.path.dirname(__file__), nueva_rutina)
    
    if not os.path.exists(script_path):
        forms.alert("No se encontró el archivo: {}".format(nueva_rutina))
        break

    # Ejecutar la rutina de armado
    try:
        with codecs.open(script_path, 'r', encoding='utf-8') as archivo_py: # Leemos el archivo forzando estrictamente el formato UTF-8
            codigo_fuente = archivo_py.read()
        variables_finales = {'__file__': script_path, '__name__': '__main__', 'saved_form_top': form_top, 'saved_form_left': form_left} # Creamos la caja de memoria (namespace) para la rutina de armado
        exec(codigo_fuente, variables_finales) # Ejecutamos el código en la memoria
    except Exception as e:
        forms.alert("Error interno al ejecutar '{}':\n\n{}".format(nueva_rutina, str(e)))
        break

    # Obtención de la nueva rutina de armado, luego de cerrarse la actual
    nueva_rutina_idx = variables_finales.get('nueva_rutina_idx', rutina_actual_idx) # Busca la variable "nueva_rutina_idx" en la rutina de armado que acaba de cerrarse
    if nueva_rutina_idx == rutina_actual_idx:
        break # El usuario no cambió la rutina (apretó ESC).
    else:
        rutina_actual_idx = nueva_rutina_idx # Actualizamos el índice y el while True reabrirá la nueva rutina de armado escogida
    
    # Guardar la última posición del formulario en memoria
    form_top = variables_finales.get('form_top', form_top) 
    form_left = variables_finales.get('form_left', form_left)