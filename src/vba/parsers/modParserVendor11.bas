Attribute VB_Name = "modParserVendor11"

Sub ParseVendor11(hoja, y, Optional ctx As AppContext)

   'Cliente
    Set ctx = ResolveContext(ctx)
    palabrabuscada = "PAN AMERICAN ENERGY"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        For i = 1 To 20
            If celdaencontrada.Offset(0, i) <> "" Then
                COD = Replace(celdaencontrada.Offset(0, i), ".", "")
                If Len(COD) <> 4 Then COD = COD & celdaencontrada.Offset(0, i + 1)
                Exit For
            End If
        Next i

        For Each fila In ctx.tblCORS.ListRows
            If UCase(fila.Range(ctx.tblCORS.ListColumns("Cliente VENDOR11").Range.Column).Value) = UCase(COD) Then
                site = fila.Range(ctx.tblCORS.ListColumns("Sucursal").Range.Column)
                Call asignarCORS(y, site)
                Exit For
            End If
        Next fila
            
    End If
    
    'Referencia
    palabrabuscada = "A"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        For i = 1 To 20
            extractedRef = celdaencontrada.Offset(0, i)
            If extractedRef <> "" And IsNumeric(Right(extractedRef, 1)) Then
            
                For j = 1 To Len(extractedRef)
                    If Mid(extractedRef, j, 1) Like "[0-9]" Then Resultado = Resultado & Mid(extractedRef, j, 1)
                Next j
                
                extractedRef = Resultado
                COMP = Right(extractedRef, 8)
                PDV = Mid(extractedRef, 1, Len(extractedRef) - Len(COMP))
                extractedRef = PDV & palabrabuscada & COMP
                Hoja2.Cells(y, ctx.rngReferencia.Range.Column).Value = extractedRef
                Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column).Value = extractedRef
                Exit For
            End If
        Next i
        'COD FC o NC
        For i = 1 To 10
            COD = celdaencontrada.Offset(i, 0)
            If COD <> "" Then
                If COD = "1" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "FC-REC"
                If COD = "201" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "FCE-REC"
                If COD = "3" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "NC-FAL"
                If COD = "203" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "NCE-FAL"
                Exit For
            End If
        Next i
        
    End If
    
    'Fecha
    palabrabuscada = "Fecha:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        For i = 1 To 10
            extractedFECHA = celdaencontrada.Offset(0, i)
            If extractedFECHA <> "" Then
                extractedFECHA = CDate(extractedFECHA)
                If IsDate(extractedFECHA) Then
                    fechaFormateada = Format(DateValue(extractedFECHA), "dd.mm.yyyy")
                    Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column).Value = fechaFormateada
                End If
                Exit For
            End If
        Next i
    End If

    'CAE
    palabrabuscada = "CAE"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        For i = 1 To 5
            If celdaencontrada.Offset(0, i) <> "" Then
                extractedCAE = celdaencontrada.Offset(0, i)
                Exit For
            End If
        Next i
        
        For i = 1 To 5
            If celdaencontrada.Offset(0, -i) <> "" Then
                extractedVtoCAE = celdaencontrada.Offset(0, -i)
                extractedVtoCAE = Format(DateValue(extractedVtoCAE), "dd.mm.yyyy")
                Exit For
            End If
        Next i
        
        
        Hoja2.Cells(y, ctx.rngCAE.Range.Column).Value = extractedCAE
        Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column).Value = extractedVtoCAE
        
    End If
    
    
    
    
    
    
     'Valores
    palabrabuscada = "Subtotal"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)

    If Not celdaencontrada Is Nothing Then
    
        Dim datos(5)

        b = 1
        
        For i = 1 To 30
            DATOENCONTRADO = hoja.Cells(celdaencontrada.Row + 1, i)
            If DATOENCONTRADO <> "" And IsNumeric(Left(DATOENCONTRADO, 1)) Then
                DATOENCONTRADO = Replace(DATOENCONTRADO, ".", "")
                If DATOENCONTRADO <> datos(b - 1) Then
                    datos(b - 1) = DATOENCONTRADO
                    b = b + 1
                    If datos(5) <> "" Then Exit For
                End If
            End If
        Next i

        Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column) = datos(0) * 1 'SUBTOTAL
        If datos(1) <> 0 Then Hoja2.Cells(y, ctx.rngII.Range.Column) = datos(1) * 1 'II
        If datos(2) <> 0 Then Hoja2.Cells(y, ctx.rngIVA.Range.Column) = datos(2) * 1 'IVA
        If datos(3) <> 0 Then Hoja2.Cells(y, ctx.rngPercIVA.Range.Column) = datos(3) * 1 'PERC IVA
        If datos(4) <> 0 Then Hoja2.Cells(y, ctx.rngPercIVA.Range.Column) = datos(4) * 1 'PERCEPCION?
        Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column) = datos(5) * 1 'TOTAL

    End If
    

End Sub

