Attribute VB_Name = "modParserVendor08"

Sub ParseVendor08(hoja, y, Optional ctx As AppContext)

    'Ref
    Set ctx = ResolveContext(ctx)
    palabrabuscada = "A"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 3
            For j = 0 To 4
                If celdaencontrada.Offset(i, j) <> "" Then
                    If IsNumeric(Left(celdaencontrada.Offset(i, j), 1)) Then
                        extracted = celdaencontrada.Offset(i, j)
                        extracted = Replace(extracted, "-", "A")
                        Hoja2.Cells(y, ctx.rngReferencia.Range.Column) = extracted
                        Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column) = extracted
                        Exit For
                    End If
                End If
            Next j
            If extracted <> "" Then Exit For
        Next i
    End If
    
    'Fecha
    palabrabuscada = "FECHA:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=2, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        Fecha = Right(celdaencontrada, 10)
        If IsDate(Fecha) Then
            Fecha = Format(Fecha, "dd.mm.yyyy")
            Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column) = Fecha
        End If
    End If

    
    'COD FC o NC
    palabrabuscada = "COD.AFIP:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=2, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        
        If Right(celdaencontrada, 3) = "201" Then
            Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column) = "FCE-REC"
        ElseIf Right(celdaencontrada, 3) = "203" Then
            Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column) = "NCE-REC"
        ElseIf Right(celdaencontrada, 1) = "3" Then
            Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column) = "NC-REC"
        ElseIf Right(celdaencontrada, 1) = "1" Then
            Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column) = "FC-REC"
        End If
        
    End If

    
    Set celdaencontrada = hoja.UsedRange.Find(What:="TOTAL", LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If celdaencontrada Is Nothing Then GoTo fin
    
    
    
    'CAE
    palabrabuscada = "CAE:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        extractedCAE = Right(celdaencontrada, 14)
        Hoja2.Cells(y, ctx.rngCAE.Range.Column).Value = extractedCAE

        For i = -1 To -6 Step -1
            If celdaencontrada.Offset(0, i) <> "" Then
                extractedVtoCAE = Right(celdaencontrada.Offset(0, i), 8)
                extractedVtoCAE = Right(extractedVtoCAE, 2) & "." & Mid(extractedVtoCAE, 5, 2) & "." & Left(extractedVtoCAE, 4)
                Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column) = extractedVtoCAE
                Exit For
            End If
        Next i
        
        If extractedVtoCAE = "" Then
            For i = 1 To 3
                If celdaencontrada.Offset(i, 0) <> "" Then
                    extractedVtoCAE = Right(celdaencontrada.Offset(i, 0), 8)
                    extractedVtoCAE = Right(extractedVtoCAE, 2) & "." & Mid(extractedVtoCAE, 5, 2) & "." & Left(extractedVtoCAE, 4)
                    Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column) = extractedVtoCAE
                    Exit For
                End If
            Next i
        End If
           
    End If

    palabrabuscada = "TOTAL"
    Set celdaTTL = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    
    If Not celdaTTL Is Nothing Then
    
        For i = 1 To 4
            For j = 0 To -2 Step -1
                If celdaTTL.Offset(i, j) <> "" Then
                    If IsNumeric(Right(celdaTTL.Offset(i, j), 1)) Then
                        XTRTTL = Replace(celdaTTL.Offset(i, j), "-", "")
                        Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column) = XTRTTL * 1
                        Exit For
                    End If
                End If
            Next j
            If XTRTTL <> "" Then Exit For
        Next i
        
        
        palabrabuscada = "SUBTOTAL"
        Set celdaencontrada = hoja.Rows(celdaTTL.Row).Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
        
        If Not celdaencontrada Is Nothing Then
            For i = 1 To 4
                For j = 0 To -2 Step -1
                    If celdaencontrada.Offset(i, j) <> "" Then
                        If IsNumeric(Right(celdaencontrada.Offset(i, j), 1)) Then
                            XTRSUB = Replace(celdaencontrada.Offset(i, j), "-", "")
                            Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column) = XTRSUB * 1 ' SUBTOTAL
                            Exit For
                        End If
                    End If
                Next j
                If XTRSUB <> "" Then Exit For
            Next i
        End If
        
        palabrabuscada = "IVA 21%"
        Set celdaencontrada = hoja.Rows(celdaTTL.Row).Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
        

        If Not celdaencontrada Is Nothing Then
            For i = 1 To 4
                For j = 2 To -2 Step -1
                    If celdaencontrada.Offset(i, j) <> "" And celdaencontrada.Offset(i, j) <> XTRTTL Then
                        If IsNumeric(Right(celdaencontrada.Offset(i, j), 1)) Then
                            XTRIVA = Replace(celdaencontrada.Offset(i, j), "-", "")
                            
                            If Left(Right(XTRIVA, 3), 1) = "." Then
                                XTRIVA = Replace(Replace(XTRIVA, ",", ""), ".", ",")
                            ElseIf Left(Right(XTRIVA, 3), 1) = "," Then
                                XTRIVA = Replace(XTRIVA, ".", "")
                            End If
                                                       
                            Hoja2.Cells(y, ctx.rngIVA.Range.Column) = XTRIVA * 1 ' IVA
                            Exit For
                        End If
                    End If
                Next j
                If XTRIVA <> "" Then Exit For
            Next i
        End If
        
        palabrabuscada = "IMP. INT."
        Set celdaencontrada = hoja.Rows(celdaTTL.Row).Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
        
        If Not celdaencontrada Is Nothing Then
            For i = 1 To 4
                For j = 1 To -1 Step -1
                    If celdaencontrada.Offset(i, j) <> "" Then
                        If Right(celdaencontrada.Offset(i, j), 1) = "," Then
                        
                            XTRII = Replace(celdaencontrada.Offset(i, j), "-", "")
                            XTRII = Replace(XTRII, "$", "")
                            XTRII = Replace(XTRII, vbLf, "")
                            
                            For c = 1 To 3
                                If celdaencontrada.Offset(i + c, j) <> "" Then
                                    XTRII = XTRII & celdaencontrada.Offset(i + c, j)
                                End If
                            Next c
                            Hoja2.Cells(y, ctx.rngII.Range.Column) = XTRII * 1 ' II
                            Exit For
                        End If
                        If IsNumeric(Right(celdaencontrada.Offset(i, j), 1)) Then
                            XTRII = Replace(celdaencontrada.Offset(i, j), "-", "")
                            Hoja2.Cells(y, ctx.rngII.Range.Column) = XTRII * 1 ' II
                            Exit For
                        End If
                    End If
                Next j
                If XTRII <> "" Then Exit For
            Next i
        End If

    End If


fin:
End Sub

