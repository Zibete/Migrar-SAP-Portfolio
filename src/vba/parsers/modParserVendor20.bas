Attribute VB_Name = "modParserVendor20"

Sub ParseVendor20(hoja, y, Optional ctx As AppContext)

    'Cliente VENDOR20
    Set ctx = ResolveContext(ctx)
    palabrabuscada = "Cliente Código:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 10
            If celdaencontrada.Offset(0, i).Value <> "" Then
                extractedData = celdaencontrada.Offset(0, i).Value
                Exit For
            End If
        Next i

        For Each fila In ctx.tblCORS.ListRows
            If fila.Range(ctx.tblCORS.ListColumns("Cliente VENDOR20").Range.Column).Value = extractedData Then
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


    palabrabuscada = "fecha:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 10
            If celdaencontrada.Offset(0, i).Value <> "" Then
                extractedFECHA = celdaencontrada.Offset(0, i).Value
                Exit For
            End If
        Next i
        'Fecha
        extractedFECHA = celdaencontrada.Offset(0, 2).Value
        extractedFECHA = Replace(extractedFECHA, ".", "/")
        
        If extractedFECHA <> "" Then fechaFormateada = Format(DateValue(extractedFECHA), "dd.mm.yyyy")
        Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column).Value = fechaFormateada
        'Referencia VENDOR20
        extractedRef = Replace(celdaencontrada.Offset(-1, 0).Value, "Nº:", "")
        Hoja2.Cells(y, ctx.rngReferencia.Range.Column).Value = Replace(Replace(extractedRef, " ", ""), "-", "A")
    End If
    
    'COD FC o NC
    palabrabuscada = "Código Nº: "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        COD = Mid(celdaencontrada, Len(celdaencontrada))
        
        If COD = "1" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "FC-REM"
        If COD = "3" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "NC-FAL"

    End If

    'Total
    palabrabuscada = "total PESOS:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 10
            If celdaencontrada.Offset(0, i).Value <> "" Then
                Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column).Value = celdaencontrada.Offset(0, i).Value
                Exit For
            End If
        Next i
    End If
    
    'Internos
    palabrabuscada = "INTERNOS:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 10
            RES = celdaencontrada.Offset(0, i).Value
            If RES <> "" And InStr(RES, "%") = 0 Then
                If RES = "0,00" Then RES = ""
                Hoja2.Cells(y, ctx.rngII.Range.Column).Value = RES
                Exit For
            End If
        Next i
    End If
    
    'IIBB BSAS
    palabrabuscada = "PERC. II.BB. BA:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 10
            RES = celdaencontrada.Offset(0, i).Value
            If RES <> "" And InStr(RES, "%") = 0 Then
                If RES = "0,00" Then RES = ""
                Hoja2.Cells(y, ctx.rngIIBBBSAS.Range.Column).Value = RES
                Exit For
            End If
        Next i
    End If
    
    'IVA
    palabrabuscada = "IVA:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 10
            If celdaencontrada.Offset(0, i).Value <> "" And InStr(celdaencontrada.Offset(0, i).Value, "%") = 0 Then
                Hoja2.Cells(y, ctx.rngIVA.Range.Column).Value = celdaencontrada.Offset(0, i).Value
                Exit For
            End If
        Next i
    End If
    
    palabrabuscada = "NETO GRAVADO:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
       'Subtotal
        For i = 1 To 10
            If celdaencontrada.Offset(0, i).Value <> "" Then
                Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column).Value = celdaencontrada.Offset(0, i).Value
                Exit For
            End If
        Next i
    End If
    
    palabrabuscada = "PERC.II.BB. C.A.B.A.:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 10
            RES = celdaencontrada.Offset(0, i).Value
            If RES <> "" And InStr(RES, "%") = 0 Then
                If RES = "0,00" Then RES = ""
                Hoja2.Cells(y, ctx.rngIIBBCABA.Range.Column).Value = RES
                Exit For
            End If
        Next i
    End If
    
    'Remito ref. VENDOR20
    palabrabuscada = "Remitos - O/C:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        texto = Replace(celdaencontrada.Offset(0, 2).Value, "R", "")
        texto = Replace(Replace(texto, "(", ""), ")", "")
        If texto <> "" Then
            texto = Trim(Left(texto, Len(texto) - 8) & "R" & Right(texto, 8))
            texto = Mid(texto, 2)
        End If
        Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column).Value = texto
    End If

    'CAE VENDOR20
    palabrabuscada = "CAE:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        Hoja2.Cells(y, ctx.rngCAE.Range.Column).Value = celdaencontrada.Offset(0, 1).Value
    End If

    palabrabuscada = "Vto. CAE:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        fechaFormateada = Format(DateValue(celdaencontrada.Offset(0, 1).Value), "dd.mm.yyyy")
        Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column).Value = fechaFormateada
    End If

End Sub

