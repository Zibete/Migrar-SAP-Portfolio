Attribute VB_Name = "modParserVendor16"

Sub ParseVendor16(hoja, y, Optional ctx As AppContext)

    Set ctx = ResolveContext(ctx)
    palabrabuscada = "Destinatario:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        cliente = celdaencontrada.Offset(2, 0)
    
        If cliente = "" Then
            
            For i = 1 To 10
                cliente = celdaencontrada.Offset(2, -i)
                If cliente <> "" Then Exit For
            Next i
            
        End If
        
        Hoja2.Cells(y, ctx.rngNuevaRuta.Range.Column) = cliente
        
        If cliente = "1880" Then
            cliente = Replace(celdaencontrada.Offset(1, 0), ")", "")
            cliente = Right(cliente, 3)
        End If
        
        If cliente = "C1416CRD" Then
            cliente = Replace(celdaencontrada.Offset(1, 0), ")", "")
        End If
        
        If cliente <> "" Then
            For Each fila In ctx.tblCORS.ListRows
                If UCase(fila.Range(ctx.tblCORS.ListColumns("Cliente Massalin").Range.Column).Value) = UCase(cliente) Then
                    site = fila.Range(ctx.tblCORS.ListColumns("Sucursal").Range.Column)
                    Call asignarCORS(y, site)
                    Exit For
                End If
            Next fila
        End If

    End If
    
    
    'Fecha
    palabrabuscada = "FECHA:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 8
            Fecha = celdaencontrada.Offset(0, i)
            If Fecha <> "" Then
                Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column) = Fecha
                Exit For
            End If
        Next i
    End If
    
    'COD FC o NC
    palabrabuscada = "A"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        'Ref
        For i = 1 To 8
            extractedRef = celdaencontrada.Offset(0, i)
            If extractedRef <> "" And IsNumeric(Left(extractedRef, 1)) Then
                extractedRef = Replace(extractedRef, "-", "A")
                Hoja2.Cells(y, ctx.rngReferencia.Range.Column).Value = extractedRef
                Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column).Value = extractedRef
                Exit For
            End If
        Next i
    
        For i = 1 To 5
            If celdaencontrada.Offset(i, 0) <> "" Then
                COD = Right(celdaencontrada.Offset(i, 0), 1)
                If COD = "1" Then
                    Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column) = "FC-REC"
                    Exit For
                End If
                If COD = "3" Then
                    Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column) = "NC-REC"
                    'REF
                    palabrabuscada = "FACTURA Nº"
                    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
                    If Not celdaencontrada Is Nothing Then
                        For j = 1 To 8
                            extractedRef = celdaencontrada.Offset(0, j)
                            If extractedRef <> "" And IsNumeric(Left(extractedRef, 1)) Then
                                extractedRef = Replace(extractedRef, "-", "A")
                                Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column).Value = extractedRef
                                Exit For
                            End If
                        Next j
                    End If
                    Exit For
                End If
            End If
        Next i
    End If

    Set celdaencontrada = hoja.UsedRange.Find(What:="Hoja 1 de 2", LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then GoTo fin
    
    'CAE
    palabrabuscada = "CAE N°: "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If celdaencontrada Is Nothing Then
        palabrabuscada = "CAEN°"
        Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    End If
    
    If Not celdaencontrada Is Nothing Then
        extractedCAE = Right(celdaencontrada, 14)
        Hoja2.Cells(y, ctx.rngCAE.Range.Column).Value = extractedCAE
        For i = 1 To 8
            If celdaencontrada.Offset(0, i) <> "" Then
                extractedVtoCAE = Right(celdaencontrada.Offset(0, i), 10)
                Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column).Value = extractedVtoCAE
                Exit For
            End If
        Next i
    End If
    
    palabrabuscada = "IMPORTE NETO GRAVADO"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 3
            For j = 0 To 2
                If celdaencontrada.Offset(i, j) <> "" Then
                    If IsNumeric(Left(celdaencontrada.Offset(i, j), 1)) Then
                        Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column) = celdaencontrada.Offset(i, j) * 1 ' SUBTOTAL
                        Exit For
                    End If
                End If
            Next j
            If Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column) <> "" Then Exit For
        Next i
    End If
    
    II = 0
    
    palabrabuscada = "Ley 24625"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 20 To 1 Step -1
            If celdaencontrada.Offset(0, i) <> "" Then
                If IsNumeric(Left(celdaencontrada.Offset(0, i), 1)) Then
                    II = celdaencontrada.Offset(0, i) * 1
                    Exit For
                End If
            End If
        Next i
    End If
    
    palabrabuscada = "Fondo Especial del Tabaco"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 20 To 1 Step -1
            If celdaencontrada.Offset(0, i) <> "" Then
                If IsNumeric(Left(celdaencontrada.Offset(0, i), 1)) Then
                    II = II + celdaencontrada.Offset(0, i) * 1
                    Exit For
                End If
            End If
        Next i
    End If
    
    palabrabuscada = "Imp.Int.Cigarrillos"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 20 To 1 Step -1
            If celdaencontrada.Offset(0, i) <> "" Then
                If IsNumeric(Left(celdaencontrada.Offset(0, i), 1)) Then
                    II = II + celdaencontrada.Offset(0, i) * 1
                    Exit For
                End If
            End If
        Next i
    End If
    
    palabrabuscada = "Imp. Int. Cigarritos"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 20 To 1 Step -1
            If celdaencontrada.Offset(0, i) <> "" Then
                If IsNumeric(Left(celdaencontrada.Offset(0, i), 1)) Then
                    II = II + celdaencontrada.Offset(0, i) * 1
                    Exit For
                End If
            End If
        Next i
    End If
    
    Hoja2.Cells(y, ctx.rngII.Range.Column) = II
    
    palabrabuscada = "IVA 21%"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 20 To 1 Step -1
            If celdaencontrada.Offset(0, i) <> "" Then
                If IsNumeric(Left(celdaencontrada.Offset(0, i), 1)) Then
                    Hoja2.Cells(y, ctx.rngIVA.Range.Column) = celdaencontrada.Offset(0, i) * 1
                    Exit For
                End If
            End If
        Next i
    End If
    
    palabrabuscada = "Per.IIBB Cap.Fed. cigarrillos"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 20 To 1 Step -1
            If celdaencontrada.Offset(0, i) <> "" Then
                If IsNumeric(Left(celdaencontrada.Offset(0, i), 1)) Then
                    Hoja2.Cells(y, ctx.rngIIBBCABA.Range.Column) = celdaencontrada.Offset(0, i) * 1
                    Exit For
                End If
            End If
        Next i
    End If
    
    palabrabuscada = "TOTAL"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then Set celdaencontrada = hoja.Cells.FindNext(celdaencontrada)
    If Not celdaencontrada Is Nothing Then
        For i = 20 To 1 Step -1
            If celdaencontrada.Offset(0, i) <> "" Then
                If IsNumeric(Left(celdaencontrada.Offset(0, i), 1)) Then
                    Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column) = celdaencontrada.Offset(0, i) * 1
                    Exit For
                End If
            End If
        Next i
    End If

fin:
End Sub
