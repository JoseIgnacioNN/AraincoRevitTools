# Numerar Columnas (portable)

Herramienta autocontenida: copie esta carpeta `32_NumerarColumnas.pushbutton` a cualquier extensión pyRevit.

## Contenido

| Archivo | Función |
|---------|---------|
| `script.py` | Entrada pyRevit |
| `bundle.yaml` | Metadatos del botón |
| `numerar_columnas.py` | Lógica de análisis y numeración |
| `numerar_columnas_ui.py` | Interfaz WPF |
| `numerar_columnas_esquema.py` | Esquemas por lote |
| `bimtools_wpf_dark_theme.py` | Estilos WPF (tema oscuro) |

## Requisitos

- Revit 2024–2026 (según `bundle.yaml`)
- pyRevit
- Parámetro de instancia **Numeracion Columna** en pilares estructurales
