Attribute VB_Name = "modCors"

Sub asignarCORS_Manual()

    site = "6294"
    
    For Each fila In Selection.Rows
        If Not fila.EntireRow.Hidden Then
            y = fila.Row
            Call asignarCORS(y, site)
        End If
    Next fila
    
End Sub

Sub asignarCORS(y, site)
    Hoja2.Cells(y, gCtx.rngSite.Range.Column) = site
End Sub

