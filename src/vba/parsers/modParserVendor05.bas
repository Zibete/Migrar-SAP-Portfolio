Attribute VB_Name = "modParserVendor05"

Sub ParseVendor05(hoja, y, Optional ctx As AppContext)


    'RTO
    Set ctx = ResolveContext(ctx)
    palabrabuscada = "ADUANA"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=2, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        
        For i = 1 To 5
            If celdaencontrada.Offset(i, 0) <> "" Then
                site = celdaencontrada.Offset(i, 0)
                Exit For
            End If
        Next i

    End If

    'COD FC o NC
    palabrabuscada = "CODIGO Nº"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=2, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
     
        COD = Replace(celdaencontrada, palabrabuscada, "")
        
        If COD = "01" Or COD = "201" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "FC-REC"
        If COD = "03" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "NC-FAL"

    End If

    
    palabrabuscada = "Fecha"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=2, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        'Fecha
        For i = 1 To 5
            If celdaencontrada.Offset(i, 0) <> "" Then
                extractedFECHA = celdaencontrada.Offset(i, 0)
                If IsDate(extractedFECHA) Then fechaFormateada = Format(DateValue(extractedFECHA), "dd.mm.yyyy")
                Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column) = fechaFormateada
                Exit For
            End If
        Next i

        'Referencia
        For i = 1 To 5

                For j = 1 To 5
                    extractedRef = celdaencontrada.Offset(j, i)
                    If extractedRef <> "" And IsNumeric(Left(extractedRef, 1)) Then
                        extractedRef = Replace(extractedRef, "-", "A")
                        Hoja2.Cells(y, ctx.rngReferencia.Range.Column).Value = extractedRef
                        Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column).Value = extractedRef
                        Exit For
                    End If
                Next j
                If extractedRef <> "" Then Exit For

        Next i

    End If
    
    'IMPORTES
    'SUBTOTAL
    palabrabuscada = "TOTAL"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        
        For i = 1 To 5
            If celdaencontrada.Offset(i, 0) <> "" Then
            RES = Replace(Replace(celdaencontrada.Offset(i, 0), ",", ""), ".", ",")
                If IsNumeric(RES) Then
                    Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column).Value = RES * 1
                    Exit For
                End If
            End If
        Next i
    
    End If

    
    'IVA
    palabrabuscada = "I.V.A"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    

        If celdaencontrada.Offset(i, 0) <> "" Then
        RES = Replace(Replace(celdaencontrada.Offset(i, 0), ",", ""), ".", ",")
            If IsNumeric(RES) Then Hoja2.Cells(y, ctx.rngIVA.Range.Column).Value = RES * 1
        End If

        
    End If
    
    'IIBB CABA
    palabrabuscada = "P IIBB CABA"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    

        If celdaencontrada.Offset(i, 0) <> "" Then
        RES = Replace(Replace(celdaencontrada.Offset(i, 0), ",", ""), ".", ",")
            If IsNumeric(RES) Then Hoja2.Cells(y, ctx.rngIIBBCABA.Range.Column).Value = RES * 1
        End If

        
    End If

    
    'SUBTOTAL
    
    
    palabrabuscada = "SUBTOTAL"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then Set celdaencontrada = hoja.Cells.FindNext(celdaencontrada)
    If Not celdaencontrada Is Nothing Then

        
        If celdaencontrada.Offset(i, 0) <> "" Then
        RES = Replace(Replace(celdaencontrada.Offset(i, 0), ",", ""), ".", ",")
            If IsNumeric(RES) Then Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column).Value = RES * 1
        Else
        
            If celdaencontrada.Offset(i, 1) <> "" Then
            RES = Replace(Replace(celdaencontrada.Offset(i, 1), ",", ""), ".", ",")
                If IsNumeric(RES) Then Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column).Value = RES * 1
            End If
        
        End If

        
    End If

    'CAE
    palabrabuscada = "CAE"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 10
            If celdaencontrada.Offset(0, i) <> "" Then
                If IsNumeric(celdaencontrada.Offset(0, i)) Then
                    Hoja2.Cells(y, ctx.rngCAE.Range.Column).Value = celdaencontrada.Offset(0, i)
                    Exit For
                End If
            End If
        Next i
    End If
    
    
    
    'VTO CAE
    palabrabuscada = "VTO"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 10
            extractedVtoCAE = celdaencontrada.Offset(0, i)
            If extractedVtoCAE <> "" Then
                If IsDate(extractedVtoCAE) Then extractedVtoCAE = Format(DateValue(extractedVtoCAE), "dd.mm.yyyy")
                Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column).Value = extractedVtoCAE
                Exit For
            End If
        Next i
    End If

End Sub
