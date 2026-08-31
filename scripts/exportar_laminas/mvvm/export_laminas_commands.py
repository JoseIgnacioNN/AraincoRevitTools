# -*- coding: utf-8 -*-
"""
Capa de comandos para la herramienta Exportar Láminas (patrón Command).

RelayCommand – implementación genérica y ligera que delega la ejecución y la
               comprobación de disponibilidad a callables externos. Desacopla la
               lógica de negocio (ViewModel) de los disparadores de la Vista.

Uso típico:
    cmd = RelayCommand(
        execute_fn=lambda: vm.execute_export(),
        can_execute_fn=lambda: not vm.is_exporting,
    )
    if cmd.can_execute():
        cmd.execute()
"""


class RelayCommand(object):
    """
    Comando ligero que envuelve un callable como acción y otro como guarda.

    Parameters
    ----------
    execute_fn : callable
        Función que se invoca al ejecutar el comando.
    can_execute_fn : callable, optional
        Función sin argumentos que devuelve bool; si es None, siempre True.
    """

    def __init__(self, execute_fn, can_execute_fn=None):
        self._execute_fn = execute_fn
        self._can_execute_fn = can_execute_fn

    def can_execute(self):
        """Devuelve True si el comando puede ejecutarse en este momento."""
        if self._can_execute_fn is None:
            return True
        try:
            return bool(self._can_execute_fn())
        except Exception:
            return True

    def execute(self, *args, **kwargs):
        """Ejecuta el comando si can_execute() es True; silencia excepciones de guarda."""
        if not self.can_execute():
            return
        try:
            self._execute_fn(*args, **kwargs)
        except Exception:
            pass

    def execute_unchecked(self, *args, **kwargs):
        """Ejecuta el comando sin verificar can_execute (para casos internos)."""
        try:
            self._execute_fn(*args, **kwargs)
        except Exception:
            pass
