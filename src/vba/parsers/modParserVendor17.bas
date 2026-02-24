Attribute VB_Name = "modParserVendor17"

Sub ParseVendor17(hoja, y, Optional ctx As AppContext)

    'Cliente
    Set ctx = ResolveContext(ctx)
    palabrabuscada = "O/C Cliente:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then

        nombreSite = Mid(celdaencontrada, Len(palabrabuscada) + 2)
    
        If nombreSite = "" Then
            For i = 1 To 3
                If celdaencontrada.Offset(0, i) <> "" Then
                    nombreSite = celdaencontrada.Offset(0, i)
                    Exit For
                End If
            Next i
        End If
        
        For Each fila In ctx.tblCORS.ListRows
            If UCase(fila.Range(ctx.tblCORS.ListColumns("Cliente VENDOR17").Range.Column).Value) = UCase(nombreSite) Then
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
        
        If nombreSite = "" Then
            
        End If
    
    End If
    

    'Referencia
    palabrabuscada = "N° "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then

        'Referencia
        extractedRef = Replace(Mid(celdaencontrada, Len(palabrabuscada) + 1), "-", "A")
        Hoja2.Cells(y, ctx.rngReferencia.Range.Column).Value = extractedRef
        Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column).Value = extractedRef

    End If
    
    'COD FC o NC
    palabrabuscada = "Código"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        COD = Mid(celdaencontrada, Len(palabrabuscada) + 1, 2)

        If COD = "01" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "FC-REC"
        If COD = "02" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "ND-ARR"
        If COD = "03" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "NC-FAL"
              
        If COD = "03" Or COD = "02" Then
            palabrabuscada = "Factura: "
            Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
            If Not celdaencontrada Is Nothing Then
                'RtoRef
                extractedRef = Replace(Mid(celdaencontrada, Len(palabrabuscada) + 6), "-", "A")
                Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column).Value = extractedRef
            End If
        End If
        
    End If
    
    'Fecha
    palabrabuscada = "Fecha: "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then

        'Fecha
        extractedFECHA = Mid(celdaencontrada, Len(palabrabuscada) + 1)

        If IsDate(extractedFECHA) Then fechaFormateada = Format(DateValue(extractedFECHA), "dd.mm.yyyy")
        Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column).Value = fechaFormateada

    End If

    'CAE
    palabrabuscada = "N° CAEA"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        If Len(celdaencontrada) > 14 Then
        
            extractedCAE = celdaencontrada

            If IsDate(celdaencontrada.Offset(1, 0)) Then
                extractedVtoCAE = Format(celdaencontrada.Offset(1, 0), "dd.mm.yyyy")
            End If

        Else
        
            For i = 1 To 4
            
                extractedCAE = celdaencontrada.Offset(0, i)
                If extractedCAE <> "" Then Exit For
                
            Next i
        
        End If
    
        extractedCAE = Right(extractedCAE, 14)

        If extractedVtoCAE = "" Then
            For i = 1 To 3
            
                extractedVtoCAE = celdaencontrada.Offset(1, i)
                If extractedVtoCAE <> "" Then
                    If IsDate(extractedVtoCAE) Then
                        extractedVtoCAE = Format(extractedVtoCAE, "dd.mm.yyyy")
                        Exit For
                    End If
                End If

            Next i
        End If
        
        Hoja2.Cells(y, ctx.rngCAE.Range.Column).Value = extractedCAE
        Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column).Value = extractedVtoCAE
        
    End If
    
    Dim palabrasBuscadas As Variant
    Dim valoresEncontrados(9) As Variant
    
    palabrasBuscadas = Array("Subtotal", "IVA 21 %", "IVA 10,5 %", "Percepc II.BB. Salta", "Percepc II.BB. Cap. Federal", "Percepc II.BB. La Rioja", _
                             "Percepc II.BB. Neuquén", "Percepc II.BB. Mendoza", "Percepc II.BB. Catamarca", "Total")
                                                 

    k = 0
    
    For j = LBound(palabrasBuscadas) To UBound(palabrasBuscadas)
    
        palabrabuscada = palabrasBuscadas(j)
        
        If palabrabuscada = "Total" Then Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
        If palabrabuscada <> "Total" Then Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=2, MatchCase:=False)
        
        If Not celdaencontrada Is Nothing Then
            For i = 0 To 5
                For h = 1 To 5
                    dato = celdaencontrada.Offset(i, h)
                    If dato <> "" Then
                        Dim partes() As String
                        If InStr(dato, vbLf) > 0 Then
                            dato = Replace(dato, "-", "")
                            partes = Split(dato, vbLf)
                            If IsNumeric(partes(0)) Then
                                valoresEncontrados(j) = Replace(partes(0), ".", "")
                                For k = j + 1 To UBound(palabrasBuscadas)
                                    j = j + 1
                                    If Not hoja.UsedRange.Find(What:=palabrasBuscadas(k), LookIn:=xlValues, LookAt:=2, MatchCase:=False) Is Nothing Then
                                        valoresEncontrados(k) = Replace(Replace(partes(1), ".", ""), "-", "")
                                        Exit For
                                    End If
                                Next k
                            End If
                        Else
                            If IsNumeric(dato) Then
                                valoresEncontrados(j) = Replace(Replace(dato, ".", ""), "-", "")
                                Exit For
                            End If
                        End If
                        If valoresEncontrados(j) <> "" Then Exit For
                    End If
                Next h
                If dato <> "" Then Exit For
            Next i
        End If
        
    Next j
    
    If valoresEncontrados(0) <> 0 Then If Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column) = "" Then Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column) = valoresEncontrados(0) * 1
    If valoresEncontrados(1) <> 0 Then If Hoja2.Cells(y, ctx.rngIVA.Range.Column) = "" Then Hoja2.Cells(y, ctx.rngIVA.Range.Column) = valoresEncontrados(1) * 1
    If valoresEncontrados(2) <> 0 Then If Hoja2.Cells(y, ctx.rngIVA105.Range.Column) = "" Then Hoja2.Cells(y, ctx.rngIVA105.Range.Column) = valoresEncontrados(2) * 1
    If valoresEncontrados(3) <> 0 Then If Hoja2.Cells(y, ctx.rngIIBBSalta.Range.Column) = "" Then Hoja2.Cells(y, ctx.rngIIBBSalta.Range.Column) = valoresEncontrados(3) * 1
    If valoresEncontrados(4) <> 0 Then If Hoja2.Cells(y, ctx.rngIIBBCABA.Range.Column) = "" Then Hoja2.Cells(y, ctx.rngIIBBCABA.Range.Column) = valoresEncontrados(4) * 1
    If valoresEncontrados(5) <> 0 Then If Hoja2.Cells(y, ctx.rngIIBBLaRioja.Range.Column) = "" Then Hoja2.Cells(y, ctx.rngIIBBLaRioja.Range.Column) = valoresEncontrados(5) * 1
    If valoresEncontrados(6) <> 0 Then If Hoja2.Cells(y, ctx.rngIIBBNeuquen.Range.Column) = "" Then Hoja2.Cells(y, ctx.rngIIBBNeuquen.Range.Column) = valoresEncontrados(6) * 1
    If valoresEncontrados(7) <> 0 Then If Hoja2.Cells(y, ctx.rngIIBBMendoza.Range.Column) = "" Then Hoja2.Cells(y, ctx.rngIIBBMendoza.Range.Column) = valoresEncontrados(7) * 1
    If valoresEncontrados(8) <> 0 Then If Hoja2.Cells(y, ctx.rngIIBBCatamarca.Range.Column) = "" Then Hoja2.Cells(y, ctx.rngIIBBCatamarca.Range.Column) = valoresEncontrados(8) * 1
    If valoresEncontrados(9) <> 0 Then If Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column) = "" Then Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column) = valoresEncontrados(9) * 1
    


    'Criollitos
    palabrabuscada = "CGL0198"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        For i = 15 To 5 Step -1
            dato = celdaencontrada.Offset(0, i).Value
            If dato <> "" And IsNumeric(dato) Then
                Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column).Value = valoresEncontrados(0) - Replace(Replace(dato, ".", ""), "-", "")
                Hoja2.Cells(y, ctx.rngSubtotalFactura105.Range.Column).Value = Replace(dato, "-", "")
                Exit For
            End If
        Next i
        
    End If
    
    If Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column) = 0 Then Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column) = ""

End Sub
