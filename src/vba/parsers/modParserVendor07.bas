Attribute VB_Name = "modParserVendor07"

Sub ParseVendor07(hoja, y, Optional ctx As AppContext)

    'Cliente
    Set ctx = ResolveContext(ctx)
    palabrabuscada = "N° CLIENTE"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        extractedData = Right(celdaencontrada, 8)
        Hoja2.Cells(y, ctx.rngNuevaRuta.Range.Column) = extractedData
        
        For Each fila In ctx.tblCORS.ListRows
            If UCase(fila.Range(ctx.tblCORS.ListColumns("Cliente MASTELLONE").Range.Column).Value) = UCase(extractedData) Then
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
    
    'Referencia
    palabrabuscada = "N° "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        extractedData = Right(celdaencontrada, 14)
        extractedData = Replace(extractedData, " ", "A")
        Hoja2.Cells(y, ctx.rngReferencia.Range.Column) = extractedData
        Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column) = extractedData

    End If
    
    palabrabuscada = "FECHA DOCUMENTO:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        'Fecha
        extractedData = Right(celdaencontrada, 10)
        
        If IsDate(extractedData) Then
            extractedData = Format(DateValue(extractedData), "dd.mm.yyyy")
            Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column).Value = extractedData
        End If
        
    End If


    'COD FC o NC
    palabrabuscada = "COD."
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        COD = Right(celdaencontrada, 3)
    
        If COD = "001" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "FC-REC"

            
        If COD = "003" Then
            If Not hoja.UsedRange.Find(What:="RECHAZO", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False) Is Nothing Then
                Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "NC-FAL"
            ElseIf Not hoja.UsedRange.Find(What:="DEVOLUCIÓN", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False) Is Nothing Then
                Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "NC-DEV"
            End If
        End If
        
    End If
    
    
    
    'Remito Ref
    palabrabuscada = "REFERENCIA "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        
        extractedData = Trim(Split(Split(celdaencontrada.Value, "REFERENCIA ")(1), "-")(0))
            
        If extractedData <> "" Then
            
            Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column) = "0" & extractedData
                
        End If
      
    End If
    
    'RemitoRef
    palabrabuscada = "INFORME DE RECEPCION "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    
    If Not celdaencontrada Is Nothing Then
        valor = Trim(Mid(celdaencontrada.Value, InStr(1, celdaencontrada.Value, palabrabuscada, vbTextCompare) + Len(palabrabuscada)))
        
        If valor Like "R-####-########" Then
            partes = Split(valor, "-")
            remito = partes(1) & "R" & Format(partes(2), "00000000")
            
        ElseIf InStr(valor, "-") > 0 Then
            partes = Split(valor, "-")
            remito = partes(0) & "R" & Format(partes(1), "00000000")
            
        Else
            remito = "0000R" & Format(valor, "00000000")
        End If
        
        If remito <> "" Then
            Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column) = remito
        End If
    
    End If

    'CAEA
    palabrabuscada = "CAEA"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    
    If Not celdaencontrada Is Nothing Then
    
                     
        CAE = Right(celdaencontrada, 14)
        Hoja2.Cells(y, ctx.rngCAE.Range.Column) = CAE
    
        vtoCAE = Right(celdaencontrada.Offset(1, 0), 10)
        If IsDate(vtoCAE) Then
            vtoCAE = Format(DateValue(vtoCAE), "dd.mm.yyyy")
            Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column) = vtoCAE
        End If
                     
    End If
    

    'IMPORTES
    'TOTAL E IVA
    palabrabuscada = "IMPORTE TOTAL"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 15
            extractedData = celdaencontrada.Offset(0, i)
            If extractedData <> "" Then
                extractedData = Replace(extractedData, "$", "")
                Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column) = extractedData * 1
                Exit For
            End If
        Next i
    End If
    
    palabrabuscada = "IMPORTE NETO GRAVADO"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 15
            extractedData = celdaencontrada.Offset(0, i)
            If extractedData <> "" Then
                extractedData = Replace(extractedData, "$", "")
                Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column) = extractedData * 1
                Exit For
            End If
        Next i
    End If
    
    palabrabuscada = "IVA 21,00 %"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 15
            extractedData = celdaencontrada.Offset(0, i)
            If extractedData <> "" Then
                extractedData = Replace(extractedData, "$", "")
                Hoja2.Cells(y, ctx.rngIVA.Range.Column) = extractedData * 1
                Exit For
            End If
        Next i
    End If
    
    palabrabuscada = "PER.IIBB SALTA"
    
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 15 To 1 Step -1
            extractedData = celdaencontrada.Offset(0, i)
            If extractedData <> "" Then
                extractedData = Replace(extractedData, "$", "")
                Hoja2.Cells(y, ctx.rngIIBBSalta.Range.Column) = extractedData * 1
                Exit For
            End If
        Next i
    End If
    
    palabrabuscada = "PERCEPCIÓN IVA CF"
    
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 15 To 1 Step -1
            extractedData = celdaencontrada.Offset(0, i)
            If extractedData <> "" Then
                extractedData = Replace(extractedData, "$", "")
                Hoja2.Cells(y, ctx.rngPercIVA.Range.Column) = extractedData * 1
                Exit For
            End If
        Next i
    End If
    
    palabrabuscada = "PER.IIBB CABA"
    
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 15 To 1 Step -1
            extractedData = celdaencontrada.Offset(0, i)
            If extractedData <> "" Then
                extractedData = Replace(extractedData, "$", "")
                Hoja2.Cells(y, ctx.rngIIBBCABA.Range.Column) = extractedData * 1
                Exit For
            End If
        Next i
    End If
    
    palabrabuscada = "PER.IIBB CATAMARCA"
    
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 15 To 1 Step -1
            extractedData = celdaencontrada.Offset(0, i)
            If extractedData <> "" Then
                extractedData = Replace(extractedData, "$", "")
                Hoja2.Cells(y, ctx.rngIIBBCatamarca.Range.Column) = extractedData * 1
                Exit For
            End If
        Next i
    End If
    
    palabrabuscada = "PER.IIBB LA RIOJA"
    
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 15 To 1 Step -1
            extractedData = celdaencontrada.Offset(0, i)
            If extractedData <> "" Then
                extractedData = Replace(extractedData, "$", "")
                Hoja2.Cells(y, ctx.rngIIBBLaRioja.Range.Column) = extractedData * 1
                Exit For
            End If
        Next i
    End If
    
    palabrabuscada = "PER.IIBB MENDOZA"
    
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 15 To 1 Step -1
            extractedData = celdaencontrada.Offset(0, i)
            If extractedData <> "" Then
                extractedData = Replace(extractedData, "$", "")
                Hoja2.Cells(y, ctx.rngIIBBMendoza.Range.Column) = extractedData * 1
                Exit For
            End If
        Next i
    End If
    
    palabrabuscada = "PER.IIBB NEUQUEN"
    
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 15 To 1 Step -1
            extractedData = celdaencontrada.Offset(0, i)
            If extractedData <> "" Then
                extractedData = Replace(extractedData, "$", "")
                Hoja2.Cells(y, ctx.rngIIBBNeuquen.Range.Column) = extractedData * 1
                Exit For
            End If
        Next i
    End If
    
    palabrabuscada = "PER.IIBB FORMOSA"
    
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 15 To 1 Step -1
            extractedData = celdaencontrada.Offset(0, i)
            If extractedData <> "" Then
                extractedData = Replace(extractedData, "$", "")
                Hoja2.Cells(y, ctx.rngIIBBFormosa.Range.Column) = extractedData * 1
                Exit For
            End If
        Next i
    End If
    
    palabrabuscada = "IMP. MUNIC CÓRDOBA"
    
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 15 To 1 Step -1
            extractedData = celdaencontrada.Offset(0, i)
            If extractedData <> "" Then
                extractedData = Replace(extractedData, "$", "")
                Hoja2.Cells(y, ctx.rngMuniCord.Range.Column) = extractedData * 1
                Exit For
            End If
        Next i
    End If

    
End Sub
