Attribute VB_Name = "modParserVendor12"

Sub ParseVendor12(hoja, y, Optional ctx As AppContext)


    'Cliente
    Set ctx = ResolveContext(ctx)
    palabrabuscada = "PAN AMERICAN"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 20
            If celdaencontrada.Offset(0, i) <> "" And IsNumeric(celdaencontrada.Offset(0, i)) Then
                cliente = celdaencontrada.Offset(0, i)
                Hoja2.Cells(y, ctx.rngNuevaRuta.Range.Column) = cliente
                Exit For
            End If
        Next i
        If cliente <> "" Then
            For Each fila In ctx.tblCORS.ListRows
                If CStr(fila.Range(ctx.tblCORS.ListColumns("Cliente Grupo Modo").Range.Column)) = CStr(cliente) Then
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
    End If
    
    
    palabrabuscada = "Fecha:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        'Fecha
        For i = 1 To 6
            extractedFECHA = celdaencontrada.Offset(0, i)
            If extractedFECHA <> "" Then
                If IsDate(extractedFECHA) Then
                    fechaFormateada = Format(DateValue(extractedFECHA), "dd.mm.yyyy")
                    Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column).Value = fechaFormateada
                    Exit For
                End If
            End If
        Next i
        'Referencia
        For i = -1 To 6
            extractedRef = celdaencontrada.Offset(-1, i)
            If extractedRef <> "" And IsNumeric(Right(extractedRef, 1)) Then
                extractedRef = Right(extractedRef, 12)
                extractedRef = Left(extractedRef, 4) & "A" & Right(extractedRef, 8)
                Hoja2.Cells(y, ctx.rngReferencia.Range.Column) = extractedRef
                Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column) = extractedRef
                If COD = 1 Then Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column).Value = extractedRef
                If COD = 3 Then
                    palabrabuscada = "PEDIDO"
                    Set celdaEncontrada2 = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                    If Not celdaEncontrada2 Is Nothing Then
                        If Len(celdaEncontrada2) = Len(palabrabuscada) Then
                            For j = 1 To 10
                                extractedRef = celdaEncontrada2.Offset(0, j)
                                If extractedRef <> "" And IsNumeric(Right(extractedRef, 1)) Then
                                    extractedRef = Right(extractedRef, 12)
                                    extractedRef = Left(extractedRef, 4) & "A" & Right(extractedRef, 8)
                                    Exit For
                                End If
                            Next j
                        Else
                            extractedRef = Mid(extractedRef, Len(palabrabuscada) + 1, 4) & "A" & Right(extractedRef, 8)
                        End If
                        Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column).Value = extractedRef
                    End If
                End If
                Exit For
            End If
        Next i
    End If

    'COD FC o NC
    palabrabuscada = "A"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then

        For i = 0 To 5
            For j = 0 To 5
                RES = celdaencontrada.Offset(i, j)
                If RES <> "" And RES <> palabrabuscada And IsNumeric(Left(RES, 1)) Then
                    COD = RES
                    If COD = "2" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column) = "ND-ARR"
                    If COD = "1" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column) = "FC-REC"
                    If COD = "201" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column) = "FCE-REC"
                    If COD = "3" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column) = "NC-FAL"
                    If COD = "203" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column) = "NCE-FAL"
                    Exit For
                End If
            Next j
            If COD <> "" Then Exit For
        Next i

        If COD = "3" Or COD = "203" Or COD = "2" Then
            palabrabuscada = "Pedido"
            Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
            If Not celdaencontrada Is Nothing Then
                If Len(celdaencontrada) = Len(palabrabuscada) Then
                    For k = 1 To 10
                        If celdaencontrada.Offset(0, k) <> "" Then
                            XTR = celdaencontrada.Offset(0, k)
                            Exit For
                        End If
                    Next k
                Else
                    XTR = celdaencontrada
                End If
                
                XTR = Right(XTR, 12)
                Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column) = Left(XTR, 4) & "A" & Right(XTR, 8)
            
            End If

        End If

    End If

    'CAE
    palabrabuscada = "CAE"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 10
            If celdaencontrada.Offset(0, i) <> "" Then
                extractedCAE = celdaencontrada.Offset(0, i)
                Hoja2.Cells(y, ctx.rngCAE.Range.Column).Value = extractedCAE
                Exit For
            End If
        Next i
        For i = 1 To 10
            If celdaencontrada.Offset(0, -i) <> "" Then
                extractedVtoCAE = celdaencontrada.Offset(0, -i)
                extractedVtoCAE = Trim(Replace(extractedVtoCAE, "Venc:", ""))
                If IsDate(extractedVtoCAE) Then
                    extractedVtoCAE = CDate(extractedVtoCAE)
                    extractedVtoCAE = Format(extractedVtoCAE, "dd.mm.yyyy")
                    Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column).Value = extractedVtoCAE
                    Exit For
                End If
            End If
        Next i
    End If
    

    
    Dim palabrasBuscadas As Variant
    Dim valoresEncontrados(6) As Variant

    'IMPORTES
    palabrabuscada = "Subtotal"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        
        b = 1
        For i = 1 To 50
            DATOENCONTRADO = hoja.Cells(celdaencontrada.Row + 1, i)
            If DATOENCONTRADO <> "" And IsNumeric(Right(DATOENCONTRADO, 1)) Then
                DATOENCONTRADO = Replace(Replace(Replace(DATOENCONTRADO, "$", ""), " ", ""), "-", "")

                If Left(Right(DATOENCONTRADO, 3), 1) = "." Then
                    DATOENCONTRADO = Replace(Replace(DATOENCONTRADO, ",", ""), ".", ",")
                ElseIf Left(Right(DATOENCONTRADO, 3), 1) = "," Then
                    DATOENCONTRADO = Replace(DATOENCONTRADO, ".", "")
                End If
                
                valoresEncontrados(b) = DATOENCONTRADO
                b = b + 1
                If valoresEncontrados(6) <> "" Then Exit For
                
            End If
        Next i

        Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column) = valoresEncontrados(1) * 1
        If valoresEncontrados(2) <> 0 Then Hoja2.Cells(y, ctx.rngII.Range.Column) = valoresEncontrados(2) * 1
        Hoja2.Cells(y, ctx.rngIVA.Range.Column) = valoresEncontrados(4) * 1
        Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column) = valoresEncontrados(6) * 1



    End If

End Sub

