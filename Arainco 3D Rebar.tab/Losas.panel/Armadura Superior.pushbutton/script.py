# -*- coding: utf-8 -*-
import os, runpy
from pyrevit import forms

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
        namespace = {'__file__': script_path, '__name__': '__main__'}
        variables_finales = runpy.run_path(script_path, init_globals=namespace, run_name='__main__') # runpy.run_path ejecuta el script y devuelve todas sus variables
    except Exception as e:
        forms.alert("Error interno al ejecutar '{}':\n\n{}".format(nueva_rutina, str(e)))
        break

    # Obtención de la nueva rutina de armado, luego de cerrarse la actual
    nueva_rutina_idx = variables_finales.get('nueva_rutina_idx', rutina_actual_idx) # Busca la variable "nueva_rutina_idx" en la rutina de armado que acaba de cerrarse
    if nueva_rutina_idx == rutina_actual_idx:
        break # El usuario no cambió la rutina (apretó ESC).
    else:
        rutina_actual_idx = nueva_rutina_idx # Actualizamos el índice y el while True reabrirá la nueva rutina de armado escogida