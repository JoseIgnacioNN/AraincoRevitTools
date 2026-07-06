# Dependencias e inventario (`column_reinforcement`)

**Repositorio canónico**: el árbol bajo `BIMTools.extension/scripts/` es la referencia en el
repositorio. El pushbutton portable (`28_ArmadoColumnasEnterprise_Portable.pushbutton/scripts/`)
duplica estos módulos para distribución empaquetada: al cambiar código compartido, actualizar
primero aquí y luego copiar al portable (véase `ESTRUCTURA_PORTABLE.txt` en la carpeta del botón).

**En el pushbutton portable** también van copias de: `column_stirrup_creator.py`,
`bimtools_wpf_dark_theme.py`, `revit_wpf_window_position.py` (alinear con esta raíz al release).

Mapa orientativo tras reorganización. Objetivo: saber qué toca qué antes de mover archivos.

## Puntos de entrada

| Origen | Destino | Notas |
|--------|---------|--------|
| `.pushbutton/script.py` | `imp.load_source` + `runner.run_pyrevit(__revit__)` | Entrada pyRevit estándar. |
| `column_reinforcement_layout_rps.run_pyrevit` | Delega en `column_reinforcement.runner.run_pyrevit` | Misma secuencia que el pushbutton: `reload` del módulo legado, inyección `doc`/`uidoc`/`__revit__`, luego `main`. |
| `column_reinforcement_layout_rps` `if __name__ == "__main__"` | `run_pyrevit(_revit__)` o `main()` | RPS / ejecución directa según exista `__revit__`. |
| `runner.run_rps(revit_app, legacy_main)` | `_run_with_legacy_main` | Sin `reload`; el llamador pasa la función `main` ya resuelta. |

## Imports desde el motor legado hacia este paquete

Archivo: `column_reinforcement_layout_rps.py`

| Import | Uso |
|--------|-----|
| `column_reinforcement.geometry.schemes` | `is_*_scheme` (con fallback local si falla el import). |
| `column_reinforcement.ui.column_layout_wizard_window` | Flujo WPF (import diferido dentro de función). |
| `column_reinforcement.ui.troceo_scheme_window` | `TroceoSchemeOutcome` (import diferido). |

## Uso interno del paquete (flujo pyRevit actual)

- `runner.run_pyrevit` → `importlib.reload(column_reinforcement_layout_rps)` → `LegacyColumnReinforcementService(legacy.main)`.
- `services.command` + `ui.main_window` → cadena WPF enterprise cuando `show_wpf=True` (no activo desde `run_pyrevit` actual).
- `geometry.segments` / `geometry.schemes` → usados por `services.strategies` y tests; el legado usa **solo** `geometry.schemes` directamente.
- `revit.versioning.adapters` → creación de adaptador API en `runner._run_with_legacy_main`.
- `revit.api.context` → `RevitExecutionContext` se construye en el runner; el servicio legado aún no consume el contexto (solo firma de `execute`); ver `ARCHITECTURE.py`.

## Módulos sin referencias de runtime actuales (stubs / preparación)

No aparecen importados por `runner`, el legado ni el resto del árbol ejecutado hoy; se mantienen para extracción futura. Cualquier reubicación debe conservar rutas de import o reexportar.

| Módulo | Rol |
|--------|-----|
| `services/factories.py` | `StrategyFactory`; solo ensambla políticas de `strategies`. |
| `services/strategies.py` | Políticas que usan `geometry`; no referenciado por el legado. |
| `revit/api/rebar_writer.py` | Contrato de creación de barras. |
| `revit/api/model_lines.py` | `ModelLineWriter`. |
| `revit/api/selection.py` | Wrapper de selección. |
| `revit/api/transactions.py` | `TransactionRunner` (nombres `Arainco: `). |

## Pruebas fuera de Revit

- `tests/test_segments.py` → `geometry.schemes`, `geometry.segments`, `models.segments`.
