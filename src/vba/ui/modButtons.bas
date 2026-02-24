Attribute VB_Name = "modButtons"

Sub btn_AbrirRetailWeb()
    Call btn_AccionBotonesSB(ACTION_ABRIR_RETAILWEB)
End Sub
Sub btn_ImprimirFactura()
    Call btn_AccionBotonesSB(ACTION_IMPRIMIR_FACTURA)
End Sub
Sub btn_CambiarEstado()
    Call btn_AccionBotonesSB(ACTION_CAMBIAR_ESTADO)
End Sub
Sub btn_PagarFactura()
    Call btn_AccionBotonesSB(ACTION_PAGAR_FACTURA)
End Sub
Sub btn_CambiarPagar()
    Call btn_AccionBotonesSB(ACTION_CAMBIAR_PAGAR)
End Sub
Sub btn_AccionBotonesSB(texto)

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Cursor = xlWait
    
    asignaciones
    gCtx.textoBtn = texto
    
    If texto <> ACTION_ABRIR_RETAILWEB Then
        If Intersect(ActiveCell, gCtx.tblDatos.DataBodyRange) Is Nothing Then
            MsgBox "Seleccione una celda dentro de la tabla"
            GoTo CleanUp
        End If
    End If
    
    For Each ventana In CreateObject("Shell.Application").Windows
        If ventana = IE_WINDOW_NAME Then
            If Left(ventana.LocationURL, Len(gCtx.dominio)) = gCtx.dominio Then
                If ventana.Visible Then botonSB = True
            End If
        End If
    Next ventana
    
    If botonSB Then
        If ActiveSheet.Shapes("LuzSB").Fill.ForeColor.RGB <> RGB(0, 255, 0) Then ActiveSheet.Shapes("LuzSB").Fill.ForeColor.RGB = RGB(0, 255, 0) 'Verde
    Else
        If ActiveSheet.Shapes("LuzSB").Fill.ForeColor.RGB <> RGB(255, 0, 0) Then ActiveSheet.Shapes("LuzSB").Fill.ForeColor.RGB = RGB(255, 0, 0) 'Rojo
'        If Hoja3.Range("passwordSB") = "" Then
'            Application.Cursor = xlDefault
'            Password_SB.Show
'            Application.Cursor = xlWait
'        End If
        Call AbrirRetailWebUser
    End If
    

    If gCtx.textoBtn = texto Then Call AbrirRecepciones
CleanUp:
    Application.Cursor = xlDefault
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub
