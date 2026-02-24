Attribute VB_Name = "modProgress"

Sub verificarWaitingPanel(Text)

    Dim startTime As Double

    startTime = Timer

    Do
        Set waitPane = gCtx.IE_NuevaVentana.Document.getElementById(SB_ID_WAITPANE)
        If Not waitPane Is Nothing Then
            Set waitingPanel = waitPane.getElementsByClassName(SB_CLASS_WAITING_PANEL)
            If waitingPanel.Length > 0 Then
                If HasTimedOut(startTime, WAIT_LONG_SECONDS) Then
                    MENSAJE = MsgBox(MSG_TIMEOUT_SB, vbCritical, MSG_TIMEOUT_TITLE)
                    gCtx.timeout = True
                    Exit Do
                End If
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
        DoEvents
    Loop

End Sub
Sub Cubo(capt1, capt2)

    SafeDeleteQuery "Cubo"
    SafeDeleteSheet "Cubo"
    On Error GoTo ErrorHandler
    
    'Progress... Paso 1
    Call incrementarProgress(capt1, capt2)
    
    PagoPendiente = GetPagoPendiente()
    
    If PagoPendiente = "SI" Then PagoPendiente = "and [Fecha de pago RW] = null"
    If PagoPendiente = "TODOS" Then PagoPendiente = ""
    If PagoPendiente = "NO" Then PagoPendiente = "and [Fecha de pago RW] <> null"

    formulaQuery = _
        "let" & vbNewLine & _
        "    Origen = AnalysisServices.Databases(""<REDACTED>"", [TypedMeasureColumns=true, Implementation=""2.0""])," & vbNewLine & _
        "    CuboVentas = Origen{[Name=""CuboVentas""]}[Data]," & vbNewLine & _
        "    Model1 = CuboVentas{[Id=""Model""]}[Data]," & vbNewLine & _
        "    Model2 = Model1{[Id=""Model""]}[Data]," & vbNewLine & _
        "    #""Elementos agregados"" = Table.SelectRows(Cube.Transform(Model2," & vbNewLine & _
        "        {" & vbNewLine & _
        "            {Cube.AddAndExpandDimensionColumn, ""[HeadDetail]"", "

    formulaQuery = formulaQuery & "{""[HeadDetail].[reference_id].[reference_id]"", " & _
        """[HeadDetail].[stock_number].[stock_number]"", " & _
        """[HeadDetail].[pay_date].[pay_date]"", " & _
        """[HeadDetail].[invoice_date].[invoice_date]"", " & _
        """[HeadDetail].[IdStore].[IdStore]"",  " & _
        """[HeadDetail].[vendor_id].[vendor_id]"", " & _
        """[HeadDetail].[reversed].[reversed]"", " & _
        """[HeadDetail].[valued_amount].[valued_amount]"", " & _
        """[HeadDetail].[TieneScan].[TieneScan]"", " & _
        """[HeadDetail].[Descripcion].[Descripcion]"", " & _
        """[HeadDetail].[pay_comment].[pay_comment]"", " & _
        """[HeadDetail].[business_date].[business_date]"", " & _
        """[HeadDetail].[total_amount].[total_amount]"", " & _
        """[HeadDetail].[total_net_amount].[total_net_amount]""}, "
        
    formulaQuery = formulaQuery & "{""Referencia"", ""RetailWeb"", ""Fecha de pago RW"", ""Fecha de documento RW"", ""Sucursal"", ""Vendor RW"", " & _
        "                ""Anulado"", ""Valorizado Documento RW"", " & _
        "                ""Tiene Scan"", ""Estado"", ""Comentario del Pago RW"",""Fecha de Negocio"",""Total RW"",""Subtotal RW""}" & vbNewLine & _
        "            }" & vbNewLine & _
        "        }), each [Vendor RW] = """ & GetVendorFilter() & """ and [Anulado] = ""False"" " & PagoPendiente & ")" & vbNewLine & _
        "in" & vbNewLine & _
        "    #""Elementos agregados"""

    ActiveWorkbook.Queries.Add Name:="Cubo", Formula:=formulaQuery

    Dim sheetCubo As Worksheet
    Set sheetCubo = ThisWorkbook.Worksheets.Add
    sheetCubo.Name = "Cubo"
    
    'Progress... Paso 2
    Call incrementarProgress(capt1, capt2)

    
    With sheetCubo.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Cubo;Extended Properties=""""" _
        , Destination:=sheetCubo.Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Cubo]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .RefreshPeriod = 0
        .ListObject.DisplayName = "Cubo"
        .Refresh
    End With
    
    'Progress... Paso 3
    Call incrementarProgress(capt1, capt2)

    ActiveWorkbook.Queries.Item("Cubo").Delete
    
    Set tblCubo = sheetCubo.ListObjects("Cubo")

    'sheetCubo.Columns("B:B").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
'''''''''''''''''''''''''''''''''''''''''
    If Not sheetCubo Is Nothing Then
    
        col = tblCubo.ListColumns("Referencia").index
    
        For Each dato In gCtx.rngTipoDoc.DataBodyRange
            
            'palabrabuscada = ""
               
            If Hoja2.Cells(dato.Row, gCtx.rngReferencia.Range.Column) = "" Then Exit For
            
            If Left(dato, 2) = "FC" Then
            
                palabrabuscada = Hoja2.Cells(dato.Row, gCtx.rngRemitoRef.Range.Column)
                
            ElseIf Left(dato, 2) = "NC" Then
            
                If Right(dato, 3) = "FAL" Then
                
                    palabrabuscada = Hoja2.Cells(dato.Row, gCtx.rngReferencia.Range.Column)
                    
                ElseIf Right(dato, 3) = "DEV" Then
                
                    fechaTbl = Hoja2.Cells(dato.Row, gCtx.rngFechaDeFactura.Range.Column)
                    siteTbl = Hoja2.Cells(dato.Row, gCtx.rngSite.Range.Column)
                    
                    For i = 1 To tblCubo.ListRows.Count
                        
                        fechaCubo = tblCubo.DataBodyRange(i, tblCubo.ListColumns("Fecha de documento RW").index)
                        siteCubo = tblCubo.DataBodyRange(i, tblCubo.ListColumns("Sucursal").index)
                        retailWeb = tblCubo.DataBodyRange(i, tblCubo.ListColumns("RetailWeb").index)
                
                        If siteCubo * 1 = siteTbl Then
                            If Format(DateValue(fechaCubo), "dd.mm.yyyy") = fechaTbl Then
                                If Left(retailWeb, 1) = "2" Then
                                    
                                    Referencia = tblCubo.DataBodyRange(i, tblCubo.ListColumns("Referencia").index)
                                    
                                    Hoja2.Cells(dato.Row, gCtx.rngRemitoRef.Range.Column) = UCase(Referencia)
                                    palabrabuscada = Referencia
                                    
                                    Exit For
                                End If
                            End If
                        End If
                        
                    Next i
                End If
            End If
            
            If Right(dato, 3) <> "INS" Then
                 'If col <> "" Then
                    'Buscamos la referencia
                    Set celdaencontrada = sheetCubo.Columns(col).Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                    
                    If Not celdaencontrada Is Nothing Then
                    
                        'Copia el RW en RW
                        retailWeb = sheetCubo.Cells(celdaencontrada.Row, tblCubo.ListColumns("RetailWeb").index)
                        Hoja2.Cells(dato.Row, gCtx.rngRetailWeb_SB.Range.Column) = retailWeb
    
                        FechaSB = sheetCubo.Cells(celdaencontrada.Row, tblCubo.ListColumns("Fecha de documento RW").index)
                        If FechaSB <> "" Then If IsDate(CDate(FechaSB)) Then FechaSB = CDate(FechaSB)
                        If FechaSB <> "" Then Hoja2.Cells(dato.Row, gCtx.rngFechaDoc_SB.Range.Column) = FechaSB
                        
                        pagado = sheetCubo.Cells(celdaencontrada.Row, tblCubo.ListColumns("Fecha de pago RW").index)
                        If pagado <> "" Then
                        Hoja2.Cells(dato.Row, gCtx.rngPagado.Range.Column) = "SI"
                        End If
                        If pagado = "" Then Hoja2.Cells(dato.Row, gCtx.rngPagado.Range.Column) = "NO"
    
                        SiteSB = sheetCubo.Cells(celdaencontrada.Row, tblCubo.ListColumns("Sucursal").index)
                        Hoja2.Cells(dato.Row, gCtx.rngSite_SB.Range.Column) = SiteSB
    
                        tieneScan = sheetCubo.Cells(celdaencontrada.Row, tblCubo.ListColumns("Tiene Scan").index)
                        Hoja2.Cells(dato.Row, gCtx.rngTieneScan_SB.Range.Column) = tieneScan
                        
                        valorizado = sheetCubo.Cells(celdaencontrada.Row, tblCubo.ListColumns("Valorizado Documento RW").index)
                        Hoja2.Cells(dato.Row, gCtx.rngValorizado_SB.Range.Column) = valorizado * 1
                        
                        totalSB = sheetCubo.Cells(celdaencontrada.Row, tblCubo.ListColumns("Total RW").index)
                        Hoja2.Cells(dato.Row, gCtx.rngTotalBruto_SB.Range.Column) = totalSB * 1
                        
                        subtotalSB = sheetCubo.Cells(celdaencontrada.Row, tblCubo.ListColumns("Subtotal RW").index)
                        Hoja2.Cells(dato.Row, gCtx.rngSubtotal_SB.Range.Column) = subtotalSB * 1
                                            
                        Estado = sheetCubo.Cells(celdaencontrada.Row, tblCubo.ListColumns("Estado").index)
                        Hoja2.Cells(dato.Row, gCtx.rngEstadoDelPago_SB.Range.Column) = Estado
                        
                        Comentarios = sheetCubo.Cells(celdaencontrada.Row, tblCubo.ListColumns("Comentario del Pago RW").index)
                        Hoja2.Cells(dato.Row, gCtx.rngObservacionesDelPago_SB.Range.Column) = Comentarios
                        
                        'If InStr(Comentarios, vbLf) > 0 Then Hoja2.Rows(dato.Row).AutoFit
    
                        fechaNegSB = sheetCubo.Cells(celdaencontrada.Row, tblCubo.ListColumns("Fecha de Negocio").index)
                        If IsDate(CDate(fechaNegSB)) Then fechaNegSB = CDate(Format(fechaNegSB, "dd/mm/yyyy"))
                        If fechaNegSB <> "" Then Hoja2.Cells(dato.Row, gCtx.rngFechaNeg_SB.Range.Column) = fechaNegSB
                        If Right(dato, 3) = "REC" Then Hoja2.Cells(dato.Row, gCtx.rngFechaBase.Range.Column) = Format(fechaNegSB, "dd.mm.yyyy")
                        
                        If CDbl(totalSB) < gCtx.montoDOA Then
                            If fechaNegSB = Date Then
                                SetRowStatus dato.Row, "", MSG_DOA_HOY
                            ElseIf (Weekday(Date) = 2 And fechaNegSB >= DateAdd("d", -3, Date)) Or fechaNegSB = DateAdd("d", -1, Date) Then
                                SetRowStatus dato.Row, "", MSG_DOA_PREFIJO & fechaNegSB & MSG_DOA_SUFIJO
                            End If
                        End If
                    End If
                 'End If
            End If
        Next dato
    End If
finProced:

    sheetCubo.Delete
    
    Exit Sub
    
ErrorHandler:

    gCtx.rngRetailWeb_SB.DataBodyRange.Formula = "Error CUBO"

    'Progress... Paso 3
    Call incrementarProgress(capt1, capt2)

GoTo finProced
    
End Sub
Public Sub incrementarProgress(capt1, capt2)

    ProgressBar.pb1.Value = ProgressBar.pb1.Value + 1
    ProgressBar.pb2.Value = ProgressBar.pb2.Value + 1
    ProgressBar.Lbl1.Caption = capt1 & " (" & Format(ProgressBar.pb1.Value / ProgressBar.pb1.Max, "0%") & ")"
    ProgressBar.Lbl2.Caption = capt2 & " (" & Format(ProgressBar.pb2.Value / ProgressBar.pb2.Max, "0%") & ")"

End Sub
Sub CuboSB(capt1, capt2)

    'Progress... Paso 1
    Call incrementarProgress(capt1, capt2)
    
    If GetMantenerDatos() = "NO" Then
        Call AbrirRetailWebCubo
    End If

    'Progress... Paso 2
    Call incrementarProgress(capt1, capt2)
    
    If gCtx.timeout = True Then GoTo ErrorHandler

    Set sheetRetailWeb = ThisWorkbook.Sheets("sheetRetailWeb")

    Set tblCuboSB = sheetRetailWeb.ListObjects("tblCuboSB")
    
    col = tblCuboSB.ListColumns("Referencia Ext.").index

    For Each dato In gCtx.rngTipoDoc.DataBodyRange 'Tipo Doc
    
        palabrabuscada = ""
        
        If Hoja2.Cells(dato.Row, gCtx.rngReferencia.Range.Column) = "" Then Exit For
        If Left(dato, 2) = "FC" Then
            palabrabuscada = Hoja2.Cells(dato.Row, gCtx.rngRemitoRef.Range.Column)
        ElseIf Left(dato, 2) = "NC" Then
            If Right(dato, 3) = "FAL" Then
                palabrabuscada = Hoja2.Cells(dato.Row, gCtx.rngReferencia.Range.Column)
            ElseIf Right(dato, 3) = "DEV" Or Right(dato, 3) = "REM" Then
            
                fechaTbl = Hoja2.Cells(dato.Row, gCtx.rngFechaDeFactura.Range.Column)
                siteTbl = Hoja2.Cells(dato.Row, gCtx.rngSite.Range.Column)
                RefTbl = Hoja2.Cells(dato.Row, gCtx.rngReferencia.Range.Column)
                
                For i = 1 To tblCuboSB.ListRows.Count
                
                    siteCubo = tblCuboSB.DataBodyRange(i, tblCuboSB.ListColumns("Negocio").index)
                    Proveedor = tblCuboSB.DataBodyRange(i, tblCuboSB.ListColumns("Proveedor").index)
                    If Proveedor = GetVendorFilter() And siteCubo = siteTbl Then
                        fechaCubo = CDate(tblCuboSB.DataBodyRange(i, tblCuboSB.ListColumns("Fecha de Factura").index))
                        If Format(DateValue(fechaCubo), "dd.mm.yyyy") = fechaTbl Then
                            retailWeb = tblCuboSB.DataBodyRange(i, tblCuboSB.ListColumns("RetailWeb #").index)
                            If Left(retailWeb, 1) = "2" Then
                                Referencia = tblCuboSB.DataBodyRange(i, tblCuboSB.ListColumns("Referencia Ext.").index)
                                If Right(dato, 3) = "REM" Or (Right(dato, 3) = "DEV" And RefTbl = Referencia) Then
                                    Hoja2.Cells(dato.Row, gCtx.rngRemitoRef.Range.Column) = UCase(Referencia)
                                    palabrabuscada = Referencia
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                    
                Next i
                
                If palabrabuscada = "" Then palabrabuscada = Hoja2.Cells(dato.Row, gCtx.rngRemitoRef.Range.Column)

            End If
        End If
        
        If Right(dato, 3) <> "INS" Then
             
            Set celdaencontrada = sheetRetailWeb.Columns(col).Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                
            If Not celdaencontrada Is Nothing Then
                If sheetRetailWeb.Cells(celdaencontrada.Row, tblCuboSB.ListColumns("Proveedor").index) <> GetVendorFilter() Then
                    Set celdaencontrada = sheetRetailWeb.Columns(col).FindNext(celdaencontrada)
                End If
            End If

            If Not celdaencontrada Is Nothing Then
            
                'Copia el RW en RW
                retailWeb = sheetRetailWeb.Cells(celdaencontrada.Row, tblCuboSB.ListColumns("RetailWeb #").index)
                Hoja2.Cells(dato.Row, gCtx.rngRetailWeb_SB.Range.Column) = retailWeb
                
                fechaDocSB = sheetRetailWeb.Cells(celdaencontrada.Row, tblCuboSB.ListColumns("Fecha de Factura").index)
                
                If fechaDocSB <> "" Then
                    If IsDate(CDate(fechaDocSB)) Then
                        fechaDocSB = CDate(fechaDocSB)
                        Hoja2.Cells(dato.Row, gCtx.rngFechaDoc_SB.Range.Column) = fechaDocSB
                    End If
                End If
                
                fechaNegSB = sheetRetailWeb.Cells(celdaencontrada.Row, tblCuboSB.ListColumns("Fecha de Negocio").index)
                
                If fechaNegSB <> "" Then
                    fechaNegSB = CDate(Int(fechaNegSB))
                    If IsDate(CDate(fechaNegSB)) Then fechaNegSB = CDate(Format(fechaNegSB, "dd/mm/yyyy"))
                    If fechaNegSB <> "" Then Hoja2.Cells(dato.Row, gCtx.rngFechaNeg_SB.Range.Column) = fechaNegSB
                End If

                If Right(dato, 3) = "REC" Then
                    If gCtx.vendorActual = "" Then gCtx.vendorActual = GetVendorFilter()
                    Set rngProveedor = gCtx.rngVendor_Prov.DataBodyRange.Find(What:=gCtx.vendorActual, LookAt:=xlWhole)
                    CONDPAGO = Hoja3.Cells(rngProveedor.Row, gCtx.rngCondPago_Prov.Range.Column)
                    If Left(CONDPAGO, 1) = "F" Then Hoja2.Cells(dato.Row, gCtx.rngFechaBase.Range.Column) = fechaDocSB
                    If Left(CONDPAGO, 1) = "Z" Then Hoja2.Cells(dato.Row, gCtx.rngFechaBase.Range.Column) = Format(fechaNegSB, "dd.mm.yyyy")
                End If
                
                pagado = sheetRetailWeb.Cells(celdaencontrada.Row, tblCuboSB.ListColumns("Fecha de Pago").index)
                If pagado <> "" Then Hoja2.Cells(dato.Row, gCtx.rngPagado.Range.Column) = "SI"
                If pagado = "" Then Hoja2.Cells(dato.Row, gCtx.rngPagado.Range.Column) = "NO"
                
                SiteSB = sheetRetailWeb.Cells(celdaencontrada.Row, tblCuboSB.ListColumns("Negocio").index)
                Hoja2.Cells(dato.Row, gCtx.rngSite_SB.Range.Column) = SiteSB

                tieneScan = sheetRetailWeb.Cells(celdaencontrada.Row, tblCuboSB.ListColumns("Tiene Factura Scaneada").index)
                Hoja2.Cells(dato.Row, gCtx.rngTieneScan_SB.Range.Column) = UCase(tieneScan)

                totalSB = sheetRetailWeb.Cells(celdaencontrada.Row, tblCuboSB.ListColumns("Total Bruto Factura").index)
                Hoja2.Cells(dato.Row, gCtx.rngTotalBruto_SB.Range.Column) = totalSB * 1
                
                subtotalSB = sheetRetailWeb.Cells(celdaencontrada.Row, tblCuboSB.ListColumns("Monto Neto Factura").index)
                If Left(retailWeb, 1) = "2" Then subtotalSB = -subtotalSB
                Hoja2.Cells(dato.Row, gCtx.rngSubtotal_SB.Range.Column) = subtotalSB * 1
                                                       
                valorizado = sheetRetailWeb.Cells(celdaencontrada.Row, tblCuboSB.ListColumns("Costo Valorizado").index)
                If Left(retailWeb, 1) = "2" Then valorizado = -valorizado
                Hoja2.Cells(dato.Row, gCtx.rngValorizado_SB.Range.Column) = valorizado * 1
                
                Estado = sheetRetailWeb.Cells(celdaencontrada.Row, tblCuboSB.ListColumns("Estado del Pago").index)
                Hoja2.Cells(dato.Row, gCtx.rngEstadoDelPago_SB.Range.Column) = Estado
                
                Comentarios = sheetRetailWeb.Cells(celdaencontrada.Row, tblCuboSB.ListColumns("Observaciones del Pago").index)
                Hoja2.Cells(dato.Row, gCtx.rngObservacionesDelPago_SB.Range.Column) = Comentarios
                
            End If
             
        End If
    Next dato
    
fin:
    'Progress... Paso 3
    Call incrementarProgress(capt1, capt2)

    Exit Sub

ErrorHandler:
    
    gCtx.rngRetailWeb_SB.DataBodyRange.Formula = "Error RW"
    GoTo fin
    
End Sub
Sub borrarArrows()

    For Each columna In Hoja2.ListObjects("tblDatos").ListColumns
        columna.DataBodyRange.AutoFilter Field:=columna.index, VisibleDropDown:=False
    Next columna


End Sub

Sub MostrarColumna(rango)
    For Each dato In rango.DataBodyRange
        If gCtx.rngReferencia.DataBodyRange.Cells(dato.Row - rango.DataBodyRange.Row + 1) = "" Then Exit For
        If dato <> "" And dato <> "0,00" And dato <> 0# Then
            rango.Range.EntireColumn.Hidden = False
            Exit For
        End If
    Next dato
End Sub



