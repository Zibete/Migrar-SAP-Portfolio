Attribute VB_Name = "modParserVendor04"

Sub ParseVendor04(hoja, y, Optional ctx As AppContext)

    'SITE
    Set ctx = ResolveContext(ctx)
    palabrabuscada = "[image]"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then Set celdaencontrada = hoja.Cells.FindNext(celdaencontrada)
    If Not celdaencontrada Is Nothing Then

        For i = 1 To 5
            RES = celdaencontrada.Offset(-i, 0)
            If RES <> "ZARATE" And RES <> "SAN ISIDRO" And RES <> "VICENTE LOPEZ - GBA" Then
                If RES <> "" Then
                    For Each fila In ctx.tblCORS.ListRows
                        If CStr(UCase(fila.Range(ctx.tblCORS.ListColumns("Cliente VENDOR04").Range.Column).Value)) <> "" Then
                            If InStr(1, CStr(UCase(RES)), CStr(UCase(fila.Range(ctx.tblCORS.ListColumns("Cliente VENDOR04").Range.Column).Value)), vbTextCompare) > 0 Then
                                Hoja2.Cells(y, ctx.rngTexto.Range.Column).Value = fila.Range(ctx.tblCORS.ListColumns("Texto").Range.Column).Value
                                Hoja2.Cells(y, ctx.rngCeBe.Range.Column).Value = fila.Range(ctx.tblCORS.ListColumns("CeBe").Range.Column).Value
                                Hoja2.Cells(y, ctx.rngNombreSite.Range.Column).Value = fila.Range(ctx.tblCORS.ListColumns("Nombre Sucursal").Range.Column).Value
                                Hoja2.Cells(y, ctx.rngSupl.Range.Column).Value = fila.Range(ctx.tblCORS.ListColumns("Supl.").Range.Column).Value
                                Hoja2.Cells(y, ctx.rngSite.Range.Column).Value = fila.Range(ctx.tblCORS.ListColumns("Sucursal").Range.Column).Value
                                Hoja2.Cells(y, ctx.rngZona.Range.Column).Value = fila.Range(ctx.tblCORS.ListColumns("Zona").Range.Column).Value
                                Hoja2.Cells(y, ctx.rngAN.Range.Column).Value = fila.Range(ctx.tblCORS.ListColumns("AN").Range.Column).Value
                                Hoja2.Cells(y, ctx.rngMails.Range.Column).Value = fila.Range(ctx.tblCORS.ListColumns("Mails").Range.Column).Value
                                Exit For
                            End If
                        End If
                    Next fila
                End If
            End If
        Next i
        
    End If

    'COD FC o NC
    palabrabuscada = "COD. "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        COD = Mid(celdaencontrada, Len(palabrabuscada) + 1, 2)

        If COD = "01" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "FC-REC"
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
        extractedRef = Right(Replace(extractedRef, "-", "A"), 14)
    End If
    
    'Referencia
    If extractedRef = "" And COD = "01" Then
        palabrabuscada = "Factura Nro "
        Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
        If Not celdaencontrada Is Nothing Then
            'Referencia
            extractedRef = Mid(celdaencontrada, Len(palabrabuscada) + 1)
            extractedRef = Replace(extractedRef, "-", "A")
        End If
    End If

    Hoja2.Cells(y, ctx.rngReferencia.Range.Column) = extractedRef
    Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column) = extractedRef
    

    'IMPORTES
    'SUBTOTAL
    palabrabuscada = "Total Neto"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 5
            If celdaencontrada.Offset(0, i) <> "" Then
                If IsNumeric(Replace(celdaencontrada.Offset(0, i), ".", "")) Then
                    'TOTAL
                    Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column).Value = Replace(celdaencontrada.Offset(0, i), ".", "") * 1
                    Exit For
                End If
            End If
        Next i
    End If
        
    'IVA
    palabrabuscada = "Total IVA"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 5
            If celdaencontrada.Offset(0, i) <> "" Then
                If IsNumeric(Replace(celdaencontrada.Offset(0, i), ".", "")) Then
                    'TOTAL
                    Hoja2.Cells(y, ctx.rngIVA.Range.Column).Value = Replace(celdaencontrada.Offset(0, i), ".", "") * 1
                    Exit For
                End If
            End If
        Next i
    End If

    'TOTAL
    palabrabuscada = "Total:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 5
            If celdaencontrada.Offset(0, i) <> "" Then
                If IsNumeric(Replace(celdaencontrada.Offset(0, i), ".", "")) Then
                    'TOTAL
                    Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column).Value = Replace(celdaencontrada.Offset(0, i), ".", "") * 1
                    Exit For
                End If
            End If
        Next i
        For i = 1 To 5
            If celdaencontrada.Offset(1, i) <> "" Then
                'CAE
                Hoja2.Cells(y, ctx.rngCAE.Range.Column).Value = celdaencontrada.Offset(1, i)
                extractedVtoCAE = celdaencontrada.Offset(2, i)
                extractedVtoCAE = Right(extractedVtoCAE, 2) & "." & Mid(extractedVtoCAE, 5, 2) & "." & Left(extractedVtoCAE, 4)
                Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column).Value = extractedVtoCAE
                Exit For
            End If
        Next i
    End If

End Sub
