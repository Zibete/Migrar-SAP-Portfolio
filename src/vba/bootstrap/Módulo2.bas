Attribute VB_Name = "Módulo2"
'------------------------------------------------------------------------------
' File: Módulo2.bas
' Purpose: VBA module extracted from legacy Excel app (portfolio version).
' Note: Cosmetic formatting only. No behavior changes.
'------------------------------------------------------------------------------
Sub BloquearPegadoOpcional()
    ' Fuerza pegar solo valores con Ctrl+V y Shift+Insert
    Application.OnKey "^v", "PegadoSoloValores"
    Application.OnKey "+{INSERT}", "PegadoSoloValores"
    ' Si querés impedir mover datos con Cortar (Ctrl+X), descomentá la línea siguiente
    ' Application.OnKey "^x", ""
End Sub

Sub RestablecerPegadoOpcional()
    Application.OnKey "^v"
    Application.OnKey "+{INSERT}"
    Application.OnKey "^x"
End Sub

Sub PegadoSoloValores()
    On Error Resume Next
    Selection.PasteSpecial xlPasteValues
End Sub
