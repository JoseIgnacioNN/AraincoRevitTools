# -*- coding: utf-8 -*-
"""
Arquitectura del paquete ``column_reinforcement`` (referencia para reorganización).

Este módulo no ejecuta lógica de negocio; solo documenta capas y límites.

## Capas

- **Dominio (puro / testeable)**: ``models/`` (datos y resultados), ``geometry/``
  (operaciones sin importar Revit API).
- **Aplicación**: ``services/`` — orquestación: ``command``, ``legacy_engine``,
  ``strategies``, ``factories``.
- **Infraestructura Revit**: ``revit/api/`` (contexto, transacciones, escritura,
  selección), ``revit/versioning/`` (diferencias de API).
- **Presentación**: ``ui/`` (WPF/XAML), ``viewmodels/``.
- **Entrada**: ``runner.py`` — fachada pyRevit/RPS; recarga el legado e inyecta
  ``doc`` / ``uidoc`` / ``__revit__`` en ``column_reinforcement_layout_rps``.
- **Motor legado (monolito)**: archivo ``column_reinforcement_layout_rps.py``
  junto a ``scripts/``; contiene ``main()`` y la mayor parte del comportamiento
  hasta una futura extracción.

## Motor legado y contexto

``LegacyColumnReinforcementService.execute`` recibe un ``RevitExecutionContext``
pero solo invoca ``legacy_main()``; el contexto queda preparado para un motor
sustituto o para desacoplar transacciones sin cambiar aún el monolito.

## División del monolito (milestone aparte — no ejecutar sin checklist de paridad)

Partir ``column_reinforcement_layout_rps.py`` en submódulos solo cuando:

1. Exista checklist de paridad documentada (ver ESTRUCTURA_PORTABLE.txt del pushbutton).
2. Los cambios sean refactors mecánicos (misma orden de efectos y mismas transacciones).
3. Cada paso se valide en Revit con el mismo .rvt de referencia.

Hasta entonces, el monolito permanece como única implementación del flujo completo.
"""
