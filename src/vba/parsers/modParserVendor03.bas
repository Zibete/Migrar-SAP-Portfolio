Attribute VB_Name = "modParserVendor03"

Sub ParseVendor03(hoja, y, Optional ctx As AppContext)

    'Cliente <SUPPLIER_B>
    Set ctx = ResolveContext(ctx)
    palabrabuscada = "000000"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        Hoja2.Cells(y, ctx.rngSite.Range.Column).Value = Mid(celdaencontrada, 7, 4)
        site = CStr(Mid(celdaencontrada, 7, 4))
        Call asignarCORS(y, site)
    End If
    
    'COD FC o NC
    palabrabuscada = "CODIGO"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        COD = Mid(celdaencontrada, Len(celdaencontrada))
        If COD = "1" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "FC-REC"
        If COD = "3" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "NC-FAL"
    End If

    If COD = "1" Then Set celdaencontrada = hoja.UsedRange.Find(What:="N° ", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If COD = "3" Then Set celdaencontrada = hoja.UsedRange.Find(What:="N°", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    
    If Not celdaencontrada Is Nothing Then
        'Referencia <SUPPLIER_B>
        If COD = "1" Then extractedRef = Replace(Mid(celdaencontrada, 4, 14), "- ", "A")
        If COD = "3" Then extractedRef = celdaencontrada.Offset(0, 1) & "A" & celdaencontrada.Offset(0, 3)

        Hoja2.Cells(y, ctx.rngReferencia.Range.Column).Value = extractedRef
        Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column).Value = extractedRef
        
        'Fecha <SUPPLIER_B>
        If COD = "1" Then
        
            For i = 0 To 4
                For j = 3 To -4 Step -1
                    extractedFECHA = celdaencontrada.Offset(i, -j).Value
                    If extractedFECHA <> "" And IsDate(extractedFECHA) Then
                        extractedFECHA = Format(extractedFECHA, "dd.mm.yyyy")
                        Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column) = extractedFECHA
                        Exit For
                    End If
                Next j
                If Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column) <> "" Then Exit For
            Next i
            
        ElseIf COD = "3" Then
        
            extractedFECHA = celdaencontrada.Offset(1, 3)
            If extractedFECHA <> "" And IsDate(extractedFECHA) Then
                extractedFECHA = Format(extractedFECHA, "dd.mm.yyyy")
                Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column) = extractedFECHA
            End If
            
        End If
        
    End If
    
    'CAE <SUPPLIER_B>
    palabrabuscada = "CAE:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If celdaencontrada Is Nothing Then Set celdaencontrada = hoja.UsedRange.Find(What:="CAE", LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 10
            If celdaencontrada.Offset(0, i).Value <> "" Then
                Hoja2.Cells(y, ctx.rngCAE.Range.Column).Value = celdaencontrada.Offset(0, i).Value
                Exit For
            End If
        Next i
    End If
    
    'VTO CAE <SUPPLIER_B>
    If COD = "1" Then Set celdaencontrada = hoja.UsedRange.Find(What:="Vto CAE:", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If COD = "3" Then Set celdaencontrada = hoja.UsedRange.Find(What:="Vencimiento CAE", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    
    If Not celdaencontrada Is Nothing Then
    
        If COD = "1" Then
            For i = 1 To 10
                If celdaencontrada.Offset(0, i) <> "" Then
                    Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column).Value = Format(celdaencontrada.Offset(0, i), "dd.mm.yyyy")
                    Exit For
                End If
            Next i
            
        ElseIf COD = "3" Then
            vtoCAE = Right(celdaencontrada, 10)
            If vtoCAE <> "" And IsDate(vtoCAE) Then
                Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column) = Format(vtoCAE, "dd.mm.yyyy")
            End If
            
        End If
        
        Dim datos(3)

        b = 1
        
        For i = 1 To 30
            DATOENCONTRADO = hoja.Cells(celdaencontrada.Row - 1, i)
            If DATOENCONTRADO <> "" And IsNumeric(Left(DATOENCONTRADO, 1)) Then
                If Left(Right(DATOENCONTRADO, 3), 1) = "." Then DATOENCONTRADO = Replace(Replace(Replace(DATOENCONTRADO, ".", "#"), ",", "."), "#", ",")
                DATOENCONTRADO = Format(DATOENCONTRADO, "#,##0.00")
                If DATOENCONTRADO <> datos(b - 1) Then
                    datos(b) = DATOENCONTRADO
                    b = b + 1
                    If datos(3) <> "" Then Exit For
                End If
            End If
        Next i

        'SUBTOTAL
        Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column) = datos(1) * 1
        'IVA
        Hoja2.Cells(y, ctx.rngIVA.Range.Column) = datos(2) * 1
        'TOTAL
        Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column) = datos(3) * 1
    
    End If
    
End Sub
