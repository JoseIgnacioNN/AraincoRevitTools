# Produccion ofuscada — carpeta del boton

- Herramienta: Exportar\nLáminas
- Salida: `04_ExportarLaminasPDFDWG.pushbutton` (solo pushbutton; sin arbol pyRevit de extension)
- Entrada: `script.py` (parche bootstrap solo en DEST; origen intacto)
- Bootstrap: `bimtools_access_bootstrap.py`
- Inteligencia: `scripts/` ofuscados
- Modulos ofuscados: 47
- Autor: José Ignacio Núñez

## Contenido

```
04_ExportarLaminasPDFDWG.pushbutton/
  script.py
  bimtools_access_bootstrap.py
  scripts/
  bundle.yaml / icon…
```

Colocar esta carpeta en un panel/extension de cliente o copiar sobre un boton existente.

## Rebuild

  python _tools/prod_builder/ui_app.py
  o: python _tools/prod_builder/build_dist.py --tool "…\04_ExportarLaminasPDFDWG.pushbutton"
