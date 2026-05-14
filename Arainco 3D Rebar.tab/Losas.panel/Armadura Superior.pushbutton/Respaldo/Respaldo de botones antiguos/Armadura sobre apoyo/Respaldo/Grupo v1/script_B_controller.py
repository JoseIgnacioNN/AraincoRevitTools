# -*- coding: utf-8 -*-
from script_A_controller import ScriptAController


class ScriptBController(ScriptAController):
    """script_B: UI solo Vano 1. Por defecto usa L2 = L1."""

    def bind_to_form(self, form):
        # Reutiliza el bind original
        super(ScriptBController, self).bind_to_form(form)

        # Si estamos en modo B, garantizamos L2 coherente aunque el panel esté oculto
        if self.L1 is not None and self.L2 is None:
            self.L2 = float(self.L1)
            try:
                form.txtL2.Text = "{:.0f}".format(self.L2)
            except Exception:
                pass

    def on_vano2(self, form):
        # En B no hay Vano 2
        return

    def on_aplicar(self, form):
        # En B, si no hay L2, usamos L2 = L1
        if form.txtL1.Text and not form.txtL2.Text:
            try:
                form.txtL2.Text = str(float(form.txtL1.Text))
            except Exception:
                pass

        # Si igual quedó vacío, forzamos L2 = L1 (si L1 existe)
        if form.txtL1.Text and (not form.txtL2.Text):
            form.txtL2.Text = form.txtL1.Text

        return super(ScriptBController, self).on_aplicar(form)

