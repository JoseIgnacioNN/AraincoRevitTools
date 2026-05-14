# -*- coding: utf-8 -*-
import os
from pyrevit import forms
import runpy

dir_path = os.path.dirname(__file__)

# Estrategia de búsqueda de archivos
lista_archivos = os.listdir(dir_path)
name_0 = next((f for f in lista_archivos if f.startswith("Armadura sobre apoyo") and f.endswith(".py")), "Armadura sobre apoyo.py")
name_1 = next((f for f in lista_archivos if f.startswith("Armadura de borde") and f.endswith(".py")), "Armadura de borde.py")
name_2 = next((f for f in lista_archivos if f.startswith("Malla en 1") and f.endswith(".py")), "Malla en 1 dirección.py")
name_3 = next((f for f in lista_archivos if f.startswith("Malla en 2") and f.endswith(".py")), "Malla en 2 direcciones.py")
RUTINAS = [name_0, name_1, name_2, name_3]

rutina_actual_idx = 0

while True:
    # 1. Definir qué rutina abrir basada en el índice actual
    if 0 <= rutina_actual_idx < len(RUTINAS):
        nueva_rutina = RUTINAS[rutina_actual_idx]
    else:
        nueva_rutina = RUTINAS[0]
        rutina_actual_idx = 0
    
    script_path = os.path.join(dir_path, nueva_rutina)
    
    if not os.path.exists(script_path):
        forms.alert("No se encontró el archivo: {}".format(nueva_rutina))
        break

    # 2. Ejecutar la rutina y capturar sus variables al finalizar
    try:
        namespace = {'__file__': script_path, '__name__': '__main__'}
        # runpy.run_path ejecuta el script y DEVUELVE todas sus variables
        variables_finales = runpy.run_path(script_path, init_globals=namespace, run_name='__main__')
    except Exception as e:
        forms.alert("Error interno al ejecutar '{}':\n\n{}".format(nueva_rutina, str(e)))
        break

    # 3. Leer la variable directamente de la memoria del script que acaba de cerrarse
    # Si el script no definió 'NUEVA_RUTINA_IDX', asumimos que el usuario no cambió la opción
    rutina_post_cierre_idx = variables_finales.get('NUEVA_RUTINA_IDX', rutina_actual_idx)

    # 4. Evaluar si cerramos el orquestador o iteramos de nuevo
    if rutina_post_cierre_idx == rutina_actual_idx:
        break # El usuario no cambió la rutina (ej. apretó ESC o aplicó la armadura), cerramos todo.
    else:
        rutina_actual_idx = rutina_post_cierre_idx # Actualizamos el índice y el while True reabrirá el nuevo script