Attribute VB_Name = "modParserVendor19"

Sub ParseVendor19(hoja, y, Optional ctx As AppContext)

    'Cliente
    Set ctx = ResolveContext(ctx)
    palabrabuscada = "PAN AMERICAN"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        For i = 1 To Len(celdaencontrada)
            If IsNumeric(Mid(celdaencontrada, i, 1)) Then
                cliente = cliente & Mid(celdaencontrada, i, 1)
            Else
                If cliente <> "" Then Exit For
            End If
        Next i
        
        If cliente = "" Then

            'Cliente
            palabrabuscada = "Domicilio"
            Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
            If Not celdaencontrada Is Nothing Then
                For h = 0 To 4
                    If Len(cliente) <> 6 Then
                        For i = 1 To Len(celdaencontrada.Offset(0, h))
                            If IsNumeric(Mid(celdaencontrada.Offset(0, h), i, 1)) Then
                                RES = RES & Mid(celdaencontrada.Offset(0, h), i, 1)
                            Else
                                If Len(RES) = 6 Then
                                    cliente = RES
                                    Exit For
                                Else
                                    RES = ""
                                End If
                            End If
                        Next i
                    End If
                Next h
            End If
        End If
  
        cliente = CDbl(cliente)
        For Each fila In ctx.tblCORS.ListRows
        
            uno = CStr(UCase(fila.Range(ctx.tblCORS.ListColumns("Cliente VENDOR19").Range.Column).Value))
            dos = CStr(UCase(cliente))
            If CStr(UCase(fila.Range(ctx.tblCORS.ListColumns("Cliente VENDOR19").Range.Column).Value)) = CStr(UCase(cliente)) Then
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
        Next fila


    End If

    'COD FC o NC
    palabrabuscada = "COD.AFIP:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        COD = Mid(celdaencontrada, Len(palabrabuscada) + 1, 2)

        If CInt(COD) = 1 Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "FC-REC"
        If CInt(COD) = 3 Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "NC-FAL"

    End If
    
    'Fecha
    palabrabuscada = "Fecha:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then

        'Fecha
        extractedFECHA = Mid(celdaencontrada, Len(palabrabuscada) + 1)

        If IsDate(extractedFECHA) Then
            fechaFormateada = Format(DateValue(extractedFECHA), "dd.mm.yyyy")
            Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column).Value = fechaFormateada
        End If
        
        For i = 1 To 6
            If celdaencontrada.Offset(-i, 0) <> "" Then
                'Referencia
                extractedRef = celdaencontrada.Offset(-i, 0)
                extractedRef = Replace(extractedRef, "-", "A")
                Hoja2.Cells(y, ctx.rngReferencia.Range.Column).Value = extractedRef
                Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column).Value = extractedRef
                Exit For
            End If
        Next i
        
    End If
    
    'Ref
    If COD = "03" Then
        palabrabuscada = "FC"
        Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
        
        If Not celdaencontrada Is Nothing Then
        
            posFC = InStr(1, celdaencontrada, "FC", vbTextCompare)
            
            If posFC > 0 And Len(celdaencontrada) > posFC + 1 Then
                textoPosterior = Trim(Mid(celdaencontrada, posFC + 2))
            Else
                textoPosterior = ""
            End If

            i = 1
            Do While (textoPosterior = "") And i < 10
                textoPosterior = Trim(celdaencontrada.Offset(0, i))
                i = i + 1
            Loop

            If textoPosterior <> "" Then
                refRem = textoPosterior
                cerosNecesarios = 8 - Len(refRem)
                refRem = String(cerosNecesarios, "0") & refRem
                refRem = "0001A" & refRem
                Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column).Value = Trim(refRem)
            End If
            
        End If
    End If

        
    'CAE
    palabrabuscada = "Numero CAE: "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        extractedCAE = Mid(celdaencontrada, Len(palabrabuscada))
        Hoja2.Cells(y, ctx.rngCAE.Range.Column).Value = extractedCAE
        
    End If
    
    'vto CAE
    palabrabuscada = "Vencimiento:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then

        extractedVtoCAE = Right(celdaencontrada, 8)
        extractedVtoCAE = Right(extractedVtoCAE, 2) & "." & Mid(extractedVtoCAE, 5, 2) & "." & Left(extractedVtoCAE, 4)
        Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column).Value = extractedVtoCAE
        
    End If
    
    Dim palabrasBuscadas As Variant
    Dim valoresEncontrados(6) As Variant

    
    'IMPORTES
    'SUBTOTAL
    palabrabuscada = "Subtotal"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then Set celdaencontrada = hoja.Cells.FindNext(celdaencontrada)
    If Not celdaencontrada Is Nothing Then
        
        

        For i = 1 To 5
            RES = celdaencontrada.Offset(i, 0)
            If RES <> "" Then
                If IsNumeric(RES) Then
                    'TOTAL
                    Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column).Value = RES * 1
                    Exit For
                End If
            End If
        Next i
        
        
    End If
        
    'IVA
    palabrabuscada = "IVA 21%"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        For i = 1 To 5
            RES = celdaencontrada.Offset(0, i)
            If RES <> "" Then
                If IsNumeric(RES) Then
                    'TOTAL
                    Hoja2.Cells(y, ctx.rngIVA.Range.Column).Value = RES
                    Exit For
                End If
            End If
        Next i
        
    End If

    'TOTAL
    palabrabuscada = "TOTAL"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    

    If Not celdaencontrada Is Nothing Then
        For colOffset = 0 To 5
            For i = 1 To 5
                RES = celdaencontrada.Offset(i, colOffset).Value
                If RES <> "" Then
                    If IsNumeric(RES) Then
                        Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column).Value = RES * 1
                        Exit For
                    End If
                End If
            Next i
            If Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column).Value <> "" Then Exit For
        Next colOffset
    End If


End Sub
