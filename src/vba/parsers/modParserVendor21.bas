Attribute VB_Name = "modParserVendor21"

Sub ParseVendor21(hoja, y, Optional ctx As AppContext)


    'Cliente
    Set ctx = ResolveContext(ctx)
    palabrabuscada = "e-Mail: "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then

        site = Mid(celdaencontrada, Len(palabrabuscada) + 3, 4)
        
        If site = "unoz" Then site = "6301"
        
        Hoja2.Cells(y, ctx.rngSite.Range.Column).Value = site
        
        Call asignarCORS(y, site)

    End If

    'Ref y fecha
    palabrabuscada = "FACTURA A"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If celdaencontrada Is Nothing Then
        palabrabuscada = "NOTA DE CRÉDITO A"
        Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    End If
    
    If Not celdaencontrada Is Nothing Then
        'Referencia
        extractedRef = Replace(Mid(celdaencontrada, Len(palabrabuscada) + 1), "-", "A")
        extractedRef = Left(extractedRef, 14)
                
        Hoja2.Cells(y, ctx.rngReferencia.Range.Column).Value = extractedRef
        Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column).Value = extractedRef
    End If
    
    'Fecha
    For i = 1 To 2
        extractedFECHA = Mid(celdaencontrada.Offset(i, 0), 7)
        If extractedFECHA <> "" Then
            If IsDate(extractedFECHA) Then fechaFormateada = Format(DateValue(extractedFECHA), "dd.mm.yyyy")
            Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column).Value = fechaFormateada
            Exit For
        End If
    Next i
    
    'COD FC o NC
    palabrabuscada = "CÓD "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        COD = Mid(celdaencontrada, Len(palabrabuscada) + 1)
        
        If COD = "001" Or COD = "201" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "FC-REC"
        If COD = "003" Or COD = "203" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "NC-FAL"
        
'        If COD = "201" And Len(extractedRef) = 14 Then Hoja2.Cells(y, rngReferencia.Range.Column).value = Mid(extractedRef, 2)
'        If COD = "203" And Len(extractedRef) = 14 Then Hoja2.Cells(y, rngReferencia.Range.Column).value = Mid(extractedRef, 2)
     
    End If


    'CAE
    palabrabuscada = "VTO CAE"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        extractedVtoCAE = Right(celdaencontrada, 10)
        If IsDate(extractedVtoCAE) Then extractedVtoCAE = Format(extractedVtoCAE, "dd.mm.yyyy")
        
        extractedCAE = celdaencontrada.Offset(-1, 0)
        extractedCAE = Right(extractedCAE, 14)
        
        Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column).Value = extractedVtoCAE
        Hoja2.Cells(y, ctx.rngCAE.Range.Column).Value = extractedCAE

    End If
    

    
    'CAE
    palabrabuscada = "VTO CAE"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        extractedVtoCAE = Right(celdaencontrada, 10)
        
        If IsDate(extractedVtoCAE) Then
        
            extractedVtoCAE = Format(extractedVtoCAE, "dd.mm.yyyy")
            extractedCAE = celdaencontrada.Offset(-1, 0)
            extractedCAE = Right(extractedCAE, 14)
        
        Else
        
            For i = 1 To 5
                extractedVtoCAE = celdaencontrada.Offset(0, i)
                If IsDate(extractedVtoCAE) Then
                    extractedVtoCAE = Format(extractedVtoCAE, "dd.mm.yyyy")
                    extractedCAE = celdaencontrada.Offset(-1, i)
                    Exit For
                End If
            Next i
        
        End If
        
        Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column).Value = extractedVtoCAE
        Hoja2.Cells(y, ctx.rngCAE.Range.Column).Value = extractedCAE
        

    End If


    'IMPORTES
    'SUBTOTAL
    palabrabuscada = "Subtotal"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 10
            RES = celdaencontrada.Offset(0, i).Value
            If RES <> "" Then
                RES = Replace(Replace(RES, ",", ""), ".", ",")
                If IsNumeric(RES) Then
                    Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column).Value = RES * 1
                    Exit For
                End If
            End If
        Next i
    End If


    'IVA
    palabrabuscada = "IVA Tasa General 21%"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 10
            RES = celdaencontrada.Offset(0, i).Value
            If RES <> "" Then
                RES = Replace(Replace(RES, ",", ""), ".", ",")
                If IsNumeric(RES) Then
                    Hoja2.Cells(y, ctx.rngIVA.Range.Column).Value = RES * 1
                    Exit For
                End If
            End If
        Next i
    End If

    'PERC CABA
    palabrabuscada = "AGIP Percepción IIBB (CABA)"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 10
            RES = celdaencontrada.Offset(0, i).Value
            If RES <> "" Then
                RES = Replace(Replace(RES, ",", ""), ".", ",")
                If IsNumeric(RES) Then
                    Hoja2.Cells(y, ctx.rngIIBBCABA.Range.Column).Value = RES * 1
                    Exit For
                End If
            End If
        Next i
    End If
    
   
    'TOTAL
    palabrabuscada = "Importe Total Pesos"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then If celdaencontrada = "Subtotal Items:" Then Set celdaencontrada = hoja.Cells.FindNext(celdaencontrada)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 10
            RES = celdaencontrada.Offset(0, i).Value
            If RES <> "" Then
                RES = Replace(Replace(RES, ",", ""), ".", ",")
                If IsNumeric(RES) Then
                    Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column).Value = RES * 1
                    Exit For
                End If
            End If
        Next i
    Else

        'TOTAL
        palabrabuscada = "TOTAL ········"
        
        Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
        
        If Not celdaencontrada Is Nothing Then If Left(celdaencontrada, 3) = "Sub" Then Set celdaencontrada = hoja.Cells.FindNext(celdaencontrada)
        
        If Not celdaencontrada Is Nothing Then
 
            RES = Replace(celdaencontrada, "·", "")
            RES = Replace(RES, "$", "")
            RES = Replace(RES, "TOTAL", "")
            RES = Replace(RES, ",", "")
            RES = Replace(RES, ".", ",")
            
            If IsNumeric(RES) Then Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column).Value = RES * 1

        End If

    End If
    
End Sub

