' Ejecuta LeerOutlookContactos.ps1 sin mostrar ventana
' Uso: cscript EjecutarPSTOculto.vbs "ruta\LeerOutlookContactos.ps1" "ruta\salida.json" "ruta\pst.pst"
Set args = WScript.Arguments
If args.Count < 2 Then WScript.Quit 1
psScript = args(0)
outputPath = args(1)
pstPath = ""
If args.Count >= 3 Then pstPath = args(2)
psExe = "powershell.exe"
cmd = """" & psExe & """ -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File """ & psScript & """ -OutputPath """ & outputPath & """"
If pstPath <> "" Then
  cmd = cmd & " -PstPath """ & pstPath & """ -TodosContactos"
  If args.Count >= 4 Then cmd = cmd & " -DebugLog """ & args(3) & """"
End If
Set sh = CreateObject("WScript.Shell")
result = sh.Run(cmd, 0, True)
WScript.Quit result
