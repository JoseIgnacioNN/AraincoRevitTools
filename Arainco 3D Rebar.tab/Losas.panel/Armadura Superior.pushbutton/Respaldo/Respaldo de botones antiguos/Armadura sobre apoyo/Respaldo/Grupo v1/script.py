# -*- coding: utf-8 -*-
import os

from pyrevit import revit, forms
from System import Uri
from System.Windows import SystemParameters
from System.Windows.Media.Imaging import BitmapImage

from script_A_controller import ScriptAController
from script_B_controller import ScriptBController


class ArmaduraSuperiorDynamicForm(forms.WPFWindow):
    def __init__(self, xaml_file):
        forms.WPFWindow.__init__(self, xaml_file)

        # Logo
        ruta_logo = os.path.join(os.path.dirname(__file__), "logo.png")
        if os.path.exists(ruta_logo):
            self.imgLogo.Source = BitmapImage(Uri(ruta_logo))

        # Posición del formulario en pantalla
        self.Left = 20
        self.Top = (SystemParameters.WorkArea.Height - self.Height) / 2

        # Scripts disponibles (extensible a futuro)
        self._scripts = {
            "script_A.py": "A",
            "script_B.py": "B",
        }
        self.cmbScript.ItemsSource = sorted(self._scripts.keys())
        self.cmbScript.SelectedIndex = 0 if self.cmbScript.ItemsSource.Count > 0 else -1

        self._controller = None
        self._activate_selected_script()

    def ScriptChanged(self, sender, args):
        self._activate_selected_script()

    def _activate_selected_script(self):
        sel = self.cmbScript.SelectedItem
        if not sel:
            self._controller = None
            self.btnAplicar.IsEnabled = False
            return

        sel_s = str(sel)
        if sel_s == "script_A.py":
            self._controller = ScriptAController(revit.doc, revit.uidoc)
            self._controller.bind_to_form(self)
            self.btnAplicar.IsEnabled = True
            self._set_height_for_script("A")
            try:
                self.panelVano2.Visibility = forms.wpf.Visibility.Visible
            except Exception:
                pass
        elif sel_s == "script_B.py":
            self._controller = ScriptBController(revit.doc, revit.uidoc)
            self._controller.bind_to_form(self)
            self.btnAplicar.IsEnabled = True
            self._set_height_for_script("B")
            try:
                self.panelVano2.Visibility = forms.wpf.Visibility.Collapsed
            except Exception:
                # fallback: Collapsed via .NET enum if forms.wpf not available
                try:
                    from System.Windows import Visibility
                    self.panelVano2.Visibility = Visibility.Collapsed
                except Exception:
                    pass
        else:
            self._controller = None
            self.btnAplicar.IsEnabled = False
            try:
                from System.Windows import Visibility
                self.panelVano2.Visibility = Visibility.Visible
            except Exception:
                pass

    def _set_height_for_script(self, code):
        # Ajuste simple por módulo (si agregas B/C, define otra altura)
        if code == "A":
            self.Height = 570
        elif code == "B":
            # Un recuadro menos
            self.Height = 490
        else:
            self.Height = 300

    # ---- Eventos existentes del XAML (delegan al controlador activo) ----
    def Vano1Click(self, sender, args):
        if self._controller:
            self._controller.on_vano1(self)

    def Vano2Click(self, sender, args):
        if self._controller:
            self._controller.on_vano2(self)

    def AplicarClick(self, sender, args):
        if self._controller:
            self._controller.on_aplicar(self)


def main():
    xaml_path = os.path.join(os.path.dirname(__file__), "ArmaduraSuperiorForm.xaml")
    form = ArmaduraSuperiorDynamicForm(xaml_path)
    form.ShowDialog()


main()
