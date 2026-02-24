Attribute VB_Name = "modReporte"

Sub reporte()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Cursor = xlWait

    columnasRangos = Array( _
        gCtx.rngSite_SB, gCtx.rngNombreSite, gCtx.rngFechaDoc_SB, gCtx.rngRetailWeb_SB, gCtx.rngVendorProveedor_SB, gCtx.rngNombreProveedor_SB, _
        gCtx.rngRemitoRef, gCtx.rngTotalBruto_SB, gCtx.rngSubtotal_SB, gCtx.rngValorizado_SB, gCtx.rngDifCostos, gCtx.rngTieneScan_SB, gCtx.rngEstadoDelPago_SB, _
        gCtx.rngZona, gCtx.rngAN, gCtx.rngMails)
        
    ReDim encabezados(LBound(columnasRangos) To UBound(columnasRangos))
    
    For i = LBound(columnasRangos) To UBound(columnasRangos)
        encabezados(i) = columnasRangos(i).Name
    Next i
    
    Set nuevoLibro = Workbooks.Add
    Set nuevaHoja = nuevoLibro.Sheets(1)
    nuevaHoja.Name = "Reporte"
   
    For i = LBound(encabezados) To UBound(encabezados)
        nuevaHoja.Cells(1, i + 1).Value = encabezados(i)
    Next i

    totalFilas = columnasRangos(0).DataBodyRange.Rows.Count
    filaDestino = 2

    For i = 1 To totalFilas
        If gCtx.rngRetailWeb_SB.DataBodyRange.Cells(i, 1) <> "" Then
            For j = LBound(columnasRangos) To UBound(columnasRangos)
                nuevaHoja.Cells(filaDestino, j + 1) = columnasRangos(j).DataBodyRange.Cells(i, 1)
            Next j
            filaDestino = filaDestino + 1
        End If
    Next i

    Set tblReporte = nuevaHoja.ListObjects.Add(xlSrcRange, nuevaHoja.Range("A1").CurrentRegion, , xlYes)
    tblReporte.Name = "tblReporte"
    
    tblReporte.HeaderRowRange.HorizontalAlignment = xlCenter
    tblReporte.DataBodyRange.HorizontalAlignment = xlCenter

    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.HorizontalAlignment = xlLeft

    Range(Range("C2"), Range("C2").End(xlDown)).Select
    Selection.NumberFormat = "dd/mm/yyyy"
    Columns("E:E").ColumnWidth = 10.57
    Range("tblReporte[[Total" & Chr(10) & "Bruto" & Chr(10) & "(RetailWeb)]:[Diferencia" & Chr(10) & "VS" & Chr(10) & "RetailWeb]]").Select
    Selection.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "

    Range(Range("M2:P502"), Range("M2:P502").End(xlDown)).Select
    Selection.HorizontalAlignment = xlLeft
    
    ActiveSheet.ListObjects("tblReporte").Range.AutoFilter Field:=13, Criteria1 _
        :=Array("Error de Scan", "Pendiente de Nota de Crédito - Mercaderia Faltante", _
        "Pendiente de Reingreso", "Pendiente de revisar por negocio"), Operator:= _
        xlFilterValues
    
    nuevaHoja.Columns.AutoFit
    nuevaHoja.Rows.AutoFit
    
    Range("B2").Select
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Cursor = xlDefault
    
End Sub

