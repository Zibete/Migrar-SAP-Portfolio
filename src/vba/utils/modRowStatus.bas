Attribute VB_Name = "modRowStatus"

Public Sub SetRowStatus(ByVal rowIndex As Long, ByVal estado As String, ByVal mensaje As String, Optional ByVal appendMensaje As Boolean = False)

    If estado <> "" Then
        Hoja2.Cells(rowIndex, gCtx.rngEstado.Range.Column).Value = estado
    End If

    If mensaje <> "" Then
        If appendMensaje Then
            Dim actual As String
            actual = Hoja2.Cells(rowIndex, gCtx.rngMensajesSap.Range.Column).Value
            If actual = "" Then
                Hoja2.Cells(rowIndex, gCtx.rngMensajesSap.Range.Column).Value = mensaje
            ElseIf InStr(1, actual, mensaje) = 0 Then
                Hoja2.Cells(rowIndex, gCtx.rngMensajesSap.Range.Column).Value = actual & "-" & mensaje
            End If
        Else
            Hoja2.Cells(rowIndex, gCtx.rngMensajesSap.Range.Column).Value = mensaje
        End If
    End If

End Sub
