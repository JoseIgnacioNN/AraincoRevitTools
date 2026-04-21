# -*- coding: utf-8 -*-
"""
Orquestador de Armadura Superior.
Ejecuta la rutina secundaria y re-evalúa si el usuario solicitó cambiar de formulario.
"""
import os
import sys
import json
import tempfile
from pyrevit import forms

# Archivo temporal para comunicar el estado del combobox de rutina
temp_file = os.path.join(tempfile.gettempdir(), 'arainco_rutina.json')
dir_path = os.path.dirname(__file__)

# Estrategia a prueba de balas para errores de codificación (UTF-8 vs CP1252):
# En lugar de leer el "string" que escupe la interfaz XAML (que traía la letra ó defectuosa),
# vamos a guiarnos por el Índice del Combobox (0, 1, 2, 3) y buscamos el archivo exacto en la memoria.
lista_archivos = os.listdir(dir_path)
name_0 = next((f for f in lista_archivos if f.startswith("Armadura sobre apoyo") and f.endswith(".py")), "Armadura sobre apoyo.py")
name_1 = next((f for f in lista_archivos if f.startswith("Armadura de borde") and f.endswith(".py")), "Armadura de borde.py")
name_2 = next((f for f in lista_archivos if f.startswith("Malla en 1") and f.endswith(".py")), "Malla en 1 dirección.py")
name_3 = next((f for f in lista_archivos if f.startswith("Malla en 2") and f.endswith(".py")), "Malla en 2 direcciones.py")
RUTINAS = [name_0, name_1, name_2, name_3]

# Inicializar
if not os.path.exists(temp_file):
    with open(temp_file, 'w') as f:
        json.dump({"rutina_idx": 0}, f)

rutina_actual_idx = 0

while True:
    try:
        with open(temp_file, 'r') as f:
            datos = json.load(f)
            nueva_rutina_idx = datos.get("rutina_idx", 0)
    except Exception:
        nueva_rutina_idx = 0
    
    if 0 <= nueva_rutina_idx < len(RUTINAS):
        nueva_rutina = RUTINAS[nueva_rutina_idx]
    else:
        nueva_rutina = RUTINAS[0]
        nueva_rutina_idx = 0
    
    script_path = os.path.join(dir_path, nueva_rutina)
    rutina_actual_idx = nueva_rutina_idx
    
    if os.path.exists(script_path):
        try:
            # Se ejecuta la lógica particular de ese .py usando runpy para compatibilidad Python 2/3
            import runpy
            namespace = {'__file__': script_path, '__name__': '__main__'}
            runpy.run_path(script_path, init_globals=namespace, run_name='__main__')
        except Exception as e:
            forms.alert("Error interno al ejecutar la rutina '{}':\n\n{}".format(nueva_rutina, str(e)))
            break
    else:
        forms.alert("No se encontró el archivo de rutina correspondiente.")
        break

    # Verificamos si el UI guardó un índice NUEVO
    try:
        with open(temp_file, 'r') as f:
            datos_actualizados = json.load(f)
            rutina_post_cierre_idx = datos_actualizados.get("rutina_idx", rutina_actual_idx)
    except Exception:
        rutina_post_cierre_idx = rutina_actual_idx

    # Si la rutina no cambió, cerramos el script principal.
    if rutina_post_cierre_idx == rutina_actual_idx:
        break
