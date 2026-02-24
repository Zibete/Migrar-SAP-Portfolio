ARCHIVO: Hoja2.cls
RUTA: <REDACTED_PATH>\Hoja2.cls
==================================================

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'------------------------------------------------------------------------------
' File: Hoja2.cls
' Purpose: VBA module extracted from legacy Excel app (portfolio version).
' Note: Cosmetic formatting only. No behavior changes.
'------------------------------------------------------------------------------
Sub SAP()

    If Intersect(Selection, gCtx.tblDatos.DataBodyRange) Is Nothing Then
        MENSAJE = MsgBox("Seleccione los registros a contabilizar", vbExclamation, "Error")
        Exit Sub
    End If




























































==================================================
        End If
    gCtx.TARGET_SELECTION = Target.Address

    If Me.Check_Copy Then Exit Sub
    
    If gCtx.ControlarCambios Then
n            
        Application.ScreenUpdating = False
        
        If Not Intersect(Target, gCtx.tblDatos.HeaderRowRange) Is Nothing Then
n            If Selection.Count = 1 Then
n                Set flecha = ActiveSheet.Shapes("flecha")
                If flecha.Visible = False Then flecha.Visible = True
                despl = Hoja2.Columns(Target.Column).Left + (Hoja2.Columns(Target.Column).Width / 2) - 9
                If Round(flecha.Left) = Round(despl) Then
                    If flecha.LockAspectRatio Then
                        order = 2
                        flecha.LockAspectRatio = msoFalse
                        flecha.Flip msoFlipVertical
                    Else
                        flecha.LockAspectRatio = msoTrue
                        flecha.Flip msoFlipVertical
                        order = 1
                    End If
                Else
                    If flecha.LockAspectRatio = msoFalse Then
                        flecha.Flip msoFlipVertical
                        flecha.LockAspectRatio = msoTrue
                    End If
                    flecha.Left = despl
                    order = 1
                End If
                
                Call ordenar(Target.Column, order)
                
                Application.EnableEvents = False
                gCtx.tblDatos.ListColumns(Target.Column - 2).DataBodyRange.Select
                Application.EnableEvents = True
                
            End IfPrivate Sub Worksheet_SelectionChange(ByVal Target As Range)
                If celda.HasFormula And IsError(celda.Value) Then
n                    Debug.Print "Celda con error, no se evalúa: " & celda.Address
                    gCtx.ControlarCambios = False
                        Hoja2.Range(celda.Address) = ""
                    gCtx.ControlarCambios = True
                Else
                    valorCelda = UCase(celda.Value)
                    
                    If Intersect(celda, gCtx.rngComentarios_User.Range) Is Nothing And _
                       Intersect(celda, gCtx.rngEstado.Range) Is Nothing Then
                    
                        Call VerificarDatos(gCtx.tblDatos.ListRows(celda.Row - gCtx.tblDatos.HeaderRowRange.Row))
                    
                    ElseIf Not Intersect(celda, gCtx.rngComentarios_User.Range) Is Nothing Then
                    
                        If InStr(1, valorCelda, "ENDOS") > 0 Or InStr(1, valorCelda, "COMPENSA") > 0 Or valorCelda = "" Then
                            Call VerificarDatos(gCtx.tblDatos.ListRows(celda.Row - gCtx.tblDatos.HeaderRowRange.Row))
                        End If
                        gCtx.rngComentarios_User.Range.Columns.AutoFit
                        
                    End If
                End If
            Next celda
        
        End If
    End If
End Sub
        If Not Intersect(Target, gCtx.tblDatos.Range) Is Nothing Then
n            
            For Each celda In Target
    If gCtx.ControlarCambios ThenPrivate Sub Worksheet_Change(ByVal Target As Range)    Application.ScreenUpdating = False
    rangos = Array(gCtx.rngFechaNeg_SB, gCtx.rngTotalBruto_SB, gCtx.rngValorizado_SB, gCtx.rngSubtotal_SB)
    
    If Check_RetailWeb Then
        For Each rango In rangos
            rango.Range.EntireColumn.Hidden = False
        Next rango
    Else
        For Each rango In rangos
            rango.Range.EntireColumn.Hidden = True
        Next rango
    End If
    Application.ScreenUpdating = True
End SubPublic Sub Check_RetailWeb_Click()Attribute Check_RetailWeb_D.VB_ProcData.VB_Invoke_Func = "d\n14"
    
    If Check_RetailWeb Then Check_RetailWeb.Value = False Else: Check_RetailWeb.Value = True
    
End SubPublic Sub Check_RetailWeb_D()
    If Check_Copy Then Check_Copy.Value = False Else: Check_Copy.Value = True
 
End SubAttribute Check_Copy_S.VB_ProcData.VB_Invoke_Func = "s\n14"Public Sub Check_Copy_S()Attribute Check_Ajustar_A.VB_ProcData.VB_Invoke_Func = "a\n14"
    Application.ScreenUpdating = False
    If Check_Ajustar Then
        Check_Ajustar.Value = False
    Else
        Check_Ajustar.Value = True
    End If
    Application.ScreenUpdating = True
End SubPublic Sub Check_Ajustar_A()
End Sub    End If        MsgBox "Outlook no está abierto."    Else        End If            FormRutaCarpeta.Show        Else            MsgBox "No se ha seleccionado ningún correo electrónico."
    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    On Error GoTo 0
    
    If Not OutlookApp Is Nothing Then
n        Set objSeleccion = OutlookApp.ActiveExplorer.Selection
    
        If objSeleccion Is Nothing Then
    Dim OutlookApp As Object
    Dim objSeleccion As ObjectSub guardarAdjuntos()
    Call Importar_5_Finalizar(9, "Refresh")
        
    
    Application.Windows(ThisWorkbook.Name).DisplayHeadings = False
    
    ActiveSheet.Shapes("flecha").Visible = False
    
    Unload ProgressBar
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.CutCopyMode = False
    
    ProtectHoja2ForUi
    Application.Cursor = xlDefault
    
    Application.Windows(ThisWorkbook.Name).DisplayWorkbookTabs = False
    
End Sub
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    UnprotectHoja2Safe

    asignaciones
    Application.Cursor = xlWaitSub btn_Refresh()
    gCtx.reporteSB = False
        
'    If Hoja3.Range("passwordSB") = "" Then
'        Password_SB.Show
'    Else
        Call AbrirRetailWeb
'    End If
    
End SubSub gCtx.passwordSB()End Sub    form_BloqueoB.ShowSub VerificarBloqueos()End Sub    Configuración.ShowSub Config()    
End Sub    
    Paso3.Show    Next encabezado        End If            End If                Next i                    End If                        Exit Sub                        Paso2.Show                    If Hoja2.Cells(fila, encabezado.Column) <> "" Then 'Tiene datos                    fila = Selection.Cells(i, 1).Row                For i = 1 To Selection.Rows.Count            If encabezado.EntireColumn.Hidden = False Then        If (Left(encabezado.Value, 4) = "IIBB" Or Left(encabezado.Value, 4) = "Perc") Thenn    
    For Each encabezado In gCtx.tblDatos.HeaderRowRange
