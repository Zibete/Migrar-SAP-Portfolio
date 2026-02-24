Attribute VB_Name = "modParserVendor01"

Sub ParseVendor01(hoja, y, Optional ctx As AppContext)

    'INSUMOS
    Set ctx = ResolveContext(ctx)
    Set CeldaDocto = hoja.UsedRange.Find(What:="Docto: SR", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    Set CeldaINSUMO = hoja.UsedRange.Find(What:="INSUMO", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    
    If Not CeldaDocto Is Nothing Or Not CeldaINSUMO Is Nothing Then
        Hoja2.Cells(y, ctx.rngTexto.Range.Column).Value = "INSUMOS"
        Hoja2.Cells(y, ctx.rngReferencia.Range.Column).Value = "INSUMOS"
        Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "FC-INS"
        Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column).Value = "INSUMOS"
        Exit Sub
    End If

    'Cliente VENDOR01
    palabrabuscada = "AXN-"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        site = CStr(Mid(celdaencontrada, 12, 4))
        If site = "EVEN" Then site = "6300"
        Call asignarCORS(y, site)
        
    End If
    
    'Ref y fecha
    palabrabuscada = "Nro. "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then

        'Referencia VENDOR01
        extractedRef = Replace(Mid(celdaencontrada, 6, 13), "-", "A")
        Hoja2.Cells(y, ctx.rngReferencia.Range.Column).Value = extractedRef
        Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column).Value = extractedRef
        
        'Fecha VENDOR01
        For i = -4 To 8
            extractedFECHA = celdaencontrada.Offset(1, i)
            If extractedFECHA <> "" Then
                extractedFECHA = Mid(extractedFECHA, 15, 10)
                If IsDate(extractedFECHA) Then
                    fechaFormateada = Format(DateValue(extractedFECHA), "dd.mm.yyyy")
                    Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column).Value = fechaFormateada
                    Exit For
                End If
            End If
        Next i
        
    End If


    'COD FC o NC
    palabrabuscada = "No Código:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        For i = 0 To 5
            If celdaencontrada.Offset(0, i) <> "" Then
                COD = celdaencontrada.Offset(0, i).Value
                
                If COD = "01" Then
                    Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "FC-REC"
                    Exit For
                End If
                
                If COD = "03" Then
                    'Remito Ref
                    palabrabuscada = "Referencia"
                    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                    If Not celdaencontrada Is Nothing Then
                    
                        extractedRef = Replace(celdaencontrada, "Referencia: ", "")
                        If Left(extractedRef, 3) = "01A" Then
                            extractedRef = Mid(extractedRef, 4, 4) & "A" & Right(extractedRef, 8)
                            Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column).Value = extractedRef
                            Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "NC-FAL"
                        Else
                            palabrabuscada = "SAC "
                            Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                            If Not celdaencontrada Is Nothing Then
                                posFC = InStr(1, celdaencontrada, "01A")
                                If posFC > 0 Then
                                    
                                    extractedRef = Mid(celdaencontrada, posFC + 3, 12)
                                    
                                    For j = 1 To Len(extractedRef)
                                        If IsNumeric(Mid(extractedRef, j, 1)) Then
                                            Resultado = Resultado & Mid(extractedRef, j, 1)
                                        Else
                                            Exit For
                                        End If
                                    Next j
                                    
                                    Do While Len(Resultado) < 12
                                        Resultado = "0" & Resultado
                                    Loop

                                    extractedRef = Left(Resultado, 4) & "A" & Right(Resultado, 8)
                                    Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column).Value = extractedRef
                                    Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "NC-DEV"
                                End If
                            End If
                        End If
                    End If
                    Exit For
                End If
            End If
        Next i
    End If
    

    'CAE
    palabrabuscada = "C.A.E: "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        extractedCAE = Mid(celdaencontrada, 8, 14)
        Hoja2.Cells(y, ctx.rngCAE.Range.Column).Value = extractedCAE
        extractedVtoCAE = Mid(celdaencontrada, 28, 10)
        Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column).Value = Format(extractedVtoCAE, "dd.mm.yyyy")
    End If

    'IMPORTES VENDOR01
    'SUBTOTAL
    palabrabuscada = "Importe Gravado"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 20
            If celdaencontrada.Offset(0, i).Value <> "" Then
                If IsNumeric(Left(celdaencontrada.Offset(0, i).Value, 1)) Then
                    'SUBTOTAL
                    Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column).Value = celdaencontrada.Offset(0, i).Value
                    Exit For
                End If
            End If
        Next i
    End If
    
    
    'IVA
    palabrabuscada = "IVA 21.00 %"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 20
            If celdaencontrada.Offset(0, i).Value <> "" Then
                If IsNumeric(Left(celdaencontrada.Offset(0, i).Value, 1)) Then
                    'IVA
                    Hoja2.Cells(y, ctx.rngIVA.Range.Column).Value = celdaencontrada.Offset(0, i).Value
                    Exit For
                End If
            End If
        Next i
    End If
    
    'Perc IVA
    palabrabuscada = "Perc. IVA 3%"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 20
            If celdaencontrada.Offset(0, i).Value <> "" Then
                If IsNumeric(Left(celdaencontrada.Offset(0, i).Value, 1)) Then
                    'IVA
                    Hoja2.Cells(y, ctx.rngPercIVA.Range.Column).Value = celdaencontrada.Offset(0, i).Value
                    Exit For
                End If
            End If
        Next i
    End If
    
    'PERC CABA
    palabrabuscada = "Per. I.B CABA "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 20
            If celdaencontrada.Offset(0, i).Value <> "" Then
                If IsNumeric(Left(celdaencontrada.Offset(0, i).Value, 1)) Then
                    'PERC CABA
                    Hoja2.Cells(y, ctx.rngIIBBCABA.Range.Column).Value = celdaencontrada.Offset(0, i).Value
                    Exit For
                End If
            End If
        Next i
    End If
    
    'PERC SALTA
    palabrabuscada = "Per. I.B Salta"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 20
            If celdaencontrada.Offset(0, i).Value <> "" Then
                If IsNumeric(celdaencontrada.Offset(0, i).Value) Then
                    Hoja2.Cells(y, ctx.rngIIBBSalta.Range.Column).Value = celdaencontrada.Offset(0, i).Value
                    Exit For
                End If
            End If
        Next i
    End If
    
    'PERC CORRIENTES
    palabrabuscada = "Per. I.B Corrientes"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 20
            If celdaencontrada.Offset(0, i).Value <> "" Then
                If IsNumeric(celdaencontrada.Offset(0, i).Value) Then
                    Hoja2.Cells(y, ctx.rngIIBBCorrientes.Range.Column).Value = celdaencontrada.Offset(0, i).Value
                    Exit For
                End If
            End If
        Next i
    End If
    
    'PERC Mendoza
    palabrabuscada = "Per. I.B Mendoza"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 20
            If celdaencontrada.Offset(0, i).Value <> "" Then
                If IsNumeric(celdaencontrada.Offset(0, i).Value) Then
                    Hoja2.Cells(y, ctx.rngIIBBMendoza.Range.Column).Value = celdaencontrada.Offset(0, i).Value
                    Exit For
                End If
            End If
        Next i
    End If

    'TOTAL
    palabrabuscada = "TOTAL"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If celdaencontrada = "Subtotal Items:" Then Set celdaencontrada = hoja.Cells.FindNext(celdaencontrada)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 20
            If celdaencontrada.Offset(0, i).Value <> "" Then
                If IsNumeric(Left(celdaencontrada.Offset(0, i).Value, 1)) Then
                    'TOTAL
                    Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column).Value = celdaencontrada.Offset(0, i).Value
                    Exit For
                End If
            End If
        Next i
    End If

End Sub
