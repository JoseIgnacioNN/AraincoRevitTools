# Matriz de pruebas — Dividir barra con traslape (56)

## Offline (automatizado)

Ejecutar:

```
python BIMTools.tab/Armadura.panel/56_DividirRebarPuntoTraslape.pushbutton/scripts/_validate_geom_offline.py
```

Cubre:

- Tabla base / G25 / G35 / G45 (`traslape_mm_from_nominal_diameter_mm`)
- Solape exacto = L en corte centrado (`± L/2`)
- Rechazo de corte cerca de extremos

## En Revit 2024+ (manual)

| # | Caso | Esperado |
|---|------|----------|
| 1 | Barra recta Single ø16, corte al centro, grado Base | 2 rebars; solape medible ≈ 1140 mm; Ids nuevos; original borrado |
| 2 | Misma barra, corte a < 570 mm del extremo | Diálogo de error; modelo sin cambios |
| 3 | Cambiar combo a G35 / G45 y repetir (1) | Solape ≈ 940 mm (G35) / 820 mm (G45) para ø16 |
| 4 | Barra con gancho solo en un extremo | Gancho permanece en el tramo que conserva ese extremo; extremos de empalme sin gancho |
| 5 | Rebar free-form / NumberWithSpacing | Mensaje de no elegible; sin transacción |
| 6 | Fixed Number (N>1), corte al centro | 2 conjuntos Fixed Number ×N con mismo ArrayLength / barras incluidas |
| 7 | Maximum Spacing (N>1), corte al centro | 2 conjuntos Maximum Spacing con mismo MaxSpacing / ArrayLength / IncludeFirst/Last |
| 8 | Conjunto con Show Middle en la vista activa | Tramos con PresentationMode Middle; Detail de empalme y etiquetas ancladas a la barra media (no a la posición 0 oculta) |

Notas:

- Tags/cotas del elemento original se pierden (recreación geométrica).
- En 2025+ el resultado no es splice nativo de Revit; son dos barras independientes.
