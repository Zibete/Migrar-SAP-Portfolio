Attribute VB_Name = "modParserVendor02"

Sub ParseVendor02(hoja, y, Optional ctx As AppContext)
   
    'SITE
    Set ctx = ResolveContext(ctx)
    palabrabuscada = "Le Banana Bites"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    
    If Not celdaencontrada Is Nothing Then
    
        For Each fila In ctx.tblCORS.ListRows
            If CStr(UCase(fila.Range(ctx.tblCORS.ListColumns("Cliente BANANA'S").Range.Column).Value)) <> "" Then
                If InStr(1, CStr(UCase(celdaencontrada)), CStr(UCase(fila.Range(ctx.tblCORS.ListColumns("Cliente BANANA'S").Range.Column))), vbTextCompare) > 0 Then
                    site = fila.Range(ctx.tblCORS.ListColumns("Sucursal").Range.Column)
                    Call asignarCORS(y, site)
                    Exit For
                End If
            End If
        Next fila
        
        If Hoja2.Cells(y, ctx.rngSite.Range.Column).Value = "" Then
            For i = 1 To 20
                If celdaencontrada.Offset(0, i).Value <> "" Then
                    RES = celdaencontrada.Offset(0, i).Value
                    For Each fila In ctx.tblCORS.ListRows
                        If CStr(UCase(fila.Range(ctx.tblCORS.ListColumns("Cliente BANANA'S").Range.Column).Value)) <> "" Then
                            If InStr(1, CStr(UCase(RES)), CStr(UCase(fila.Range(ctx.tblCORS.ListColumns("Cliente BANANA'S").Range.Column).Value)), vbTextCompare) > 0 Then
                                site = fila.Range(ctx.tblCORS.ListColumns("Sucursal").Range.Column)
                                Call asignarCORS(y, site)
                                Exit For
                            End If
                        End If
                    Next fila
                End If
            Next i
        End If
        
    End If
    
    palabrabuscada = "Número:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
               
        For i = 1 To 4
            If celdaencontrada.Offset(0, i).Value <> "" Then
            
                extractedRef = celdaencontrada.Offset(0, i).Value
                PDV = Mid(extractedRef, 3, 5)
                extractedRef = PDV & "A" & Right(extractedRef, 8)
                
                'Referencia
                Hoja2.Cells(y, ctx.rngReferencia.Range.Column).Value = extractedRef
                Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column).Value = extractedRef
                
                Exit For
            End If
        Next i
        
    End If
    
    palabrabuscada = "Fecha:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then

        For i = 1 To 4
        
            extractedFECHA = celdaencontrada.Offset(0, i).Value
            
            If extractedFECHA <> "" Then
                If IsDate(extractedFECHA) Then
                
                    fechaFormateada = Format(DateValue(extractedFECHA), "dd.mm.yyyy")
                    Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column).Value = fechaFormateada
                    Exit For
                    
               End If
            End If
            
        Next i

    End If
    
    'COD FC o NC
    palabrabuscada = "A"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        For i = 1 To 4
            COD = celdaencontrada.Offset(i, 0).Value
            If COD <> "" Then
            
                If COD = "001" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "FC-REC"
                If COD = "003" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "NC-FAL"
            
                Exit For
            End If
        Next i
        
    End If
    
    
    'CAE
    palabrabuscada = "CAE:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then Hoja2.Cells(y, ctx.rngCAE.Range.Column).Value = Right(celdaencontrada, 14)

    'VTO CAE
    palabrabuscada = "Fecha Vto. CAE:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        extractedVtoCAE = Right(celdaencontrada, 10)
        If IsDate(extractedVtoCAE) Then
            extractedVtoCAE = Format(extractedVtoCAE, "dd.mm.yyyy")
            Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column).Value = extractedVtoCAE
        End If
    
    End If
    
    'Subtotal
    palabrabuscada = "Bruto:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        For i = Len(palabrabuscada) To Len(celdaencontrada)
            If IsNumeric(Mid(celdaencontrada, i, 1)) Then
                Subtotal = Mid(celdaencontrada, i)
                Exit For
            End If
        Next i
        
        If Subtotal = "" Then
            For j = 1 To 4
                RES = celdaencontrada.Offset(0, j).Value
                If RES <> "" Then
                    For i = 1 To Len(RES)
                        If IsNumeric(Mid(RES, i, 1)) Then
                            Subtotal = Mid(RES, i)
                            Exit For
                        End If
                    Next i
                    Exit For
                End If
            Next j
        End If
        
        If Subtotal <> "" Then
            Subtotal = Replace(Replace(Subtotal, ",", ""), ".", ",")
            Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column).Value = Subtotal * 1
        End If
        
    End If
    
    'IVA
    palabrabuscada = "IVA 21:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        For i = Len(palabrabuscada) To Len(celdaencontrada)
            If IsNumeric(Mid(celdaencontrada, i, 1)) Then
                IVA = Mid(celdaencontrada, i)
                Exit For
            End If
        Next i
        
        If IVA = "" Then
            For j = 1 To 4
                RES = celdaencontrada.Offset(0, j).Value
                If RES <> "" Then
                    For i = 1 To Len(RES)
                        If IsNumeric(Mid(RES, i, 1)) Then
                            IVA = Mid(RES, i)
                            Exit For
                        End If
                    Next i
                    Exit For
                End If
            Next j
        End If
        
        If IVA <> "" Then
            IVA = Replace(Replace(IVA, ",", ""), ".", ",")
            Hoja2.Cells(y, ctx.rngIVA.Range.Column).Value = IVA * 1
        End If
        
    End If
    
    'TOTAL
    palabrabuscada = "Total: $"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        For i = Len(palabrabuscada) To Len(celdaencontrada)
            If IsNumeric(Mid(celdaencontrada, i, 1)) Then
                Total = Mid(celdaencontrada, i)
                Exit For
            End If
        Next i
        
        If Total = "" Then
            For j = 1 To 4
                RES = celdaencontrada.Offset(0, j).Value
                If RES <> "" Then
                    For i = 1 To Len(RES)
                        If IsNumeric(Mid(RES, i, 1)) Then
                            Total = Mid(RES, i)
                            Exit For
                        End If
                    Next i
                    Exit For
                End If
            Next j
        End If
        
        If Total <> "" Then
            Total = Replace(Replace(Total, ",", ""), ".", ",")
            Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column).Value = Total * 1
        End If
        
    End If
    

End Sub

