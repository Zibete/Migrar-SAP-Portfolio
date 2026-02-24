Attribute VB_Name = "modParserVendor13"

Sub ParseVendor13(hoja, y, Optional ctx As AppContext)
    
    'RTO
    Set ctx = ResolveContext(ctx)
    palabrabuscada = "Remito:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=2, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        
        For i = 1 To Len(celdaencontrada)
        
            If Mid(celdaencontrada, i, 1) Like "#" And Mid(celdaencontrada, i + 1, 1) = "-" Then
                If Mid(celdaencontrada, i + 2) Like "#####*" Then
                
                    RTO = Mid(celdaencontrada, i, 7)
                    partes = Split(RTO, "-")
                    FormatoRemito = Format(CLng(partes(0)), "00000") & "R" & Format(CLng(partes(1)), "00000000")
                    Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column).Value = Trim(FormatoRemito)
                    
                    Exit For
                End If
            End If

        Next i
        
        If RTO = "" Then
            For i = 1 To Len(celdaencontrada)
                If Mid(celdaencontrada, i, 1) Like "#" Then
                    RTO = Mid(celdaencontrada, i)
                    nroRemito = Format(CLng(RTO), "00000000")
                    FormatoRemito = "00003" & "R" & nroRemito
                    Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column).Value = Trim(FormatoRemito)
                    Exit For
                End If
            Next i
        End If
        
    End If

    'COD FC o NC
    palabrabuscada = "Código Nº:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=2, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        For i = 1 To 5
            If celdaencontrada.Offset(0, i) <> "" Then
                COD = "0" & celdaencontrada.Offset(0, i)
                Exit For
            End If
        Next i
    
        If COD = "01" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "FC-REM"
        If COD = "03" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "NC-FAL"

    End If

    'Fecha
    palabrabuscada = "Fecha:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=2, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
    
        For i = 1 To 5
            If celdaencontrada.Offset(0, i) <> "" Then
                extractedFECHA = celdaencontrada.Offset(0, i)
                Exit For
            End If
        Next i
    
        'Fecha
        If IsDate(extractedFECHA) Then fechaFormateada = Format(DateValue(extractedFECHA), "dd.mm.yyyy")
        Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column) = fechaFormateada
        'Referencia
        extractedRef = celdaencontrada.Offset(-1, 0)
        extractedRef = Right(Replace(extractedRef, "-", "A"), 14)
        Hoja2.Cells(y, ctx.rngReferencia.Range.Column).Value = Trim(extractedRef)
        
    End If

    'IMPORTES
    
    'SUBTOTAL
    palabrabuscada = "Subtotal:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=2, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 10 To 1 Step -1
            If celdaencontrada.Offset(0, i) <> "" Then
                RES = Replace(celdaencontrada.Offset(0, i), ".", "")
                If IsNumeric(RES) Then
                    Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column).Value = RES * 1
                    Exit For
                End If
            End If
        Next i
    End If

    
    'IVA
    palabrabuscada = "IVA:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 10 To 1 Step -1
            If celdaencontrada.Offset(0, i) <> "" Then
                RES = Replace(celdaencontrada.Offset(0, i), ".", "")
                If IsNumeric(RES) Then
                    Hoja2.Cells(y, ctx.rngIVA.Range.Column).Value = RES * 1
                    Exit For
                End If
            End If
        Next i
    End If
    
    
    
    'TOTAL
    palabrabuscada = "TOTAL:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 15 To 1 Step -1
            If celdaencontrada.Offset(0, i) <> "" Then
                RES = Replace(celdaencontrada.Offset(0, i), ".", "")
                If IsNumeric(RES) Then
                    Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column).Value = RES * 1
                    Exit For
                End If
            End If
        Next i
    End If

       'CAE
    palabrabuscada = "C.A.E.:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
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
    palabrabuscada = "Fecha de Vencimiento:"
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
