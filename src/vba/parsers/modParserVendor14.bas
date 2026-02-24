Attribute VB_Name = "modParserVendor14"

Sub ParseVendor14(hoja, y, Optional ctx As AppContext)

    'RTO
    Set ctx = ResolveContext(ctx)
    palabrabuscada = "Observaciones:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        
        For i = 1 To Len(celdaencontrada)
            If Mid(celdaencontrada, i, 1) Like "#" And Mid(celdaencontrada, i + 1, 1) = "-" Then
                If Mid(celdaencontrada, i + 2) Like "#####*" Then
                
                    RTO = Mid(celdaencontrada, i, 7)
                    partes = Split(RTO, "-")
                    FormatoRemito = Format(CLng(partes(0)), "0000") & "R" & Format(CLng(partes(1)), "00000000")
                    Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column).Value = Trim(FormatoRemito)
                    
                    Exit For
                End If
            End If
        Next i

    End If
    
    
    'SITE
    palabrabuscada = "SUC "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        
        i = InStr(celdaencontrada, palabrabuscada) + Len(palabrabuscada)

        Do While i <= Len(celdaencontrada)
            If IsNumeric(Mid(celdaencontrada, i, 1)) Then
                RES = RES & Mid(celdaencontrada, i, 1)
            Else
                Exit Do
            End If
            i = i + 1
        Loop

        If Len(RES) = 4 Then site = RES

    End If
    
    Call asignarCORS(y, site)

    'COD FC o NC
    palabrabuscada = "COD. "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        COD = Mid(celdaencontrada, Len(palabrabuscada) + 1, 2)

        If COD = "01" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "FC-REM"
        If COD = "03" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "NC-FAL"

    End If

    'Fecha
    palabrabuscada = "Fecha: "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        'Fecha
        extractedFECHA = Mid(celdaencontrada, Len(palabrabuscada) + 1)
        If IsDate(extractedFECHA) Then fechaFormateada = Format(DateValue(extractedFECHA), "dd.mm.yyyy")
        Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column).Value = fechaFormateada
        'Referencia
        extractedRef = celdaencontrada.Offset(-1, 0)
        extractedRef = Right(Replace(extractedRef, " - ", "A"), 14)
        Hoja2.Cells(y, ctx.rngReferencia.Range.Column).Value = Trim(extractedRef)
        
    End If

    'IMPORTES
    
    'SUBTOTAL
    palabrabuscada = "Subtotal"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 10
            If celdaencontrada.Offset(0, i) <> "" Then
                RES = Replace(Replace(celdaencontrada.Offset(0, i), ",", ""), ".", ",")
                If IsNumeric(RES) Then
                    Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column).Value = RES * 1
                    Exit For
                End If
            End If
        Next i
    End If

    'IVA
    palabrabuscada = "I.V.A."
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 10 To 1 Step -1
            If celdaencontrada.Offset(0, i) <> "" Then
                RES = Replace(Replace(celdaencontrada.Offset(0, i), ",", ""), ".", ",")
                If IsNumeric(RES) Then
                    Hoja2.Cells(y, ctx.rngIVA.Range.Column).Value = RES * 1
                    Exit For
                End If
            End If
        Next i
    End If
    
    'TOTAL
    For j = 1 To 10
        If celdaencontrada.Offset(j, 0) = "TOTAL" Then
            For i = 10 To 1 Step -1
                If celdaencontrada.Offset(j, i) <> "" Then
                    RES = Replace(Replace(celdaencontrada.Offset(j, i), ",", ""), ".", ",")
                    If IsNumeric(RES) Then
                        Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column).Value = RES * 1
                        Exit For
                    End If
                End If
            Next i
        End If
    Next j
    
    
    
    
    
       'CAE
    palabrabuscada = "CAE Nº:"
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
    palabrabuscada = "FECHA VTO:"
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

