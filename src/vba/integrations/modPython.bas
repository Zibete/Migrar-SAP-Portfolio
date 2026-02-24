Attribute VB_Name = "modPython"

Sub EjecutarPythonYRecibirDato()

    Dim wsh As Object, oExec As Object
    Dim sOutput As String
    Dim rutaPython As String, rutaScript As String
    Dim datoEnviar As String
    

    rutaPython = GetPythonExePath()
    rutaScript = ResolveScriptPath("retailweb.py")  'Ruta al script
    datoEnviar = "Celda_A1"  '  dato a enviar
    

    Set wsh = CreateObject("WScript.Shell")
    Set oExec = wsh.exec(rutaPython & " " & rutaScript & " " & datoEnviar)
    sOutput = oExec.StdOut.ReadAll
    
    ' Mostrar resultado
    MsgBox "Respuesta desde Python: " & sOutput
    
End Sub
