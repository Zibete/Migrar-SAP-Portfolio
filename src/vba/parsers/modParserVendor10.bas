Attribute VB_Name = "modParserVendor10"

Sub ParseVendor10(hoja, y, Optional ctx As AppContext)

'    'SITE
'    palabrabuscada = "[image]"
'    Set celdaEncontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
'    If Not celdaEncontrada Is Nothing Then Set celdaEncontrada = hoja.Cells.FindNext(celdaEncontrada)
'    If Not celdaEncontrada Is Nothing Then
'
'        For i = 1 To 5
'            RES = celdaEncontrada.Offset(-i, 0)
'            If RES <> "ZARATE" And RES <> "SAN ISIDRO" And RES <> "VICENTE LOPEZ - GBA" Then
'                If RES <> "" Then
'                    For Each fila In tblCORS.ListRows
'                        If CStr(UCase(fila.Range(tblCORS.ListColumns("Cliente VENDOR04").Range.Column).Value)) <> "" Then
'                            If InStr(1, CStr(UCase(RES)), CStr(UCase(fila.Range(tblCORS.ListColumns("Cliente VENDOR04").Range.Column).Value)), vbTextCompare) > 0 Then
'                                Hoja2.Cells(y, rngTexto.Range.Column).Value = fila.Range(tblCORS.ListColumns("Texto").Range.Column).Value
'                                Hoja2.Cells(y, rngCeBe.Range.Column).Value = fila.Range(tblCORS.ListColumns("CeBe").Range.Column).Value
'                                Hoja2.Cells(y, rngNombreSite.Range.Column).Value = fila.Range(tblCORS.ListColumns("Nombre Sucursal").Range.Column).Value
'                                Hoja2.Cells(y, rngSupl.Range.Column).Value = fila.Range(tblCORS.ListColumns("Supl.").Range.Column).Value
'                                Hoja2.Cells(y, rngSite.Range.Column).Value = fila.Range(tblCORS.ListColumns("Sucursal").Range.Column).Value
'                                Hoja2.Cells(y, rngZona.Range.Column).Value = fila.Range(tblCORS.ListColumns("Zona").Range.Column).Value
'                                Hoja2.Cells(y, rngAN.Range.Column).Value = fila.Range(tblCORS.ListColumns("AN").Range.Column).Value
'                                Hoja2.Cells(y, rngMails.Range.Column).Value = fila.Range(tblCORS.ListColumns("Mails").Range.Column).Value
'                                Exit For
'                            End If
'                        End If
'                    Next fila
'                End If
'            End If
'        Next i
'
'    End If

    'COD FC o NC
    Set ctx = ResolveContext(ctx)
    palabrabuscada = "Codigo:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        For i = 1 To 5
        
            If celdaencontrada.Offset(0, i) <> "" Then
            
                COD = celdaencontrada.Offset(0, i)
                
                If COD = "01" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "FC-REM"
                If COD = "02" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "ND-ARR"
                If COD = "03" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "NC-FAL"
                
                If COD = "201" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "FCE-REM"
                If COD = "202" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "NDE-ARR"
                If COD = "203" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "NCE-FAL"
                
                Exit For
            
            
            End If
        
        Next i

    End If
    
'Referencia
'Fecha
    palabrabuscada = "Nro."
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        For i = 1 To 5
        
            If celdaencontrada.Offset(0, i) <> "" Then
            
                extractedRef = celdaencontrada.Offset(0, i)
                extractedRef = Replace(extractedRef, "-", "A")
                extractedRef = Trim(Replace(extractedRef, ":", ""))
                
                Hoja2.Cells(y, ctx.rngReferencia.Range.Column) = extractedRef
                Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column) = extractedRef
            
                extractedFECHA = celdaencontrada.Offset(1, i)
                extractedFECHA = Replace(extractedFECHA, "/", ".")
                extractedFECHA = Trim(Replace(extractedFECHA, ":", ""))

                Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column) = extractedFECHA
                
                Exit For
            
            End If

        Next i
        
    End If
    
'rto
    palabrabuscada = "RTO "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)

        If Not celdaencontrada Is Nothing Then
        
    Dim texto As String
    Dim posIni As Long, posFin As Long, codigo As String
    
        texto = celdaencontrada
        posIni = InStr(1, texto, palabrabuscada, vbTextCompare) + Len(palabrabuscada)
        posFin = InStr(posIni, texto, " ")
        
        If posFin > 0 Then
            codigo = Mid(texto, posIni, posFin - posIni)
        Else
            codigo = Mid(texto, posIni)
        End If
        
    
        extractedRef = "00001R" & Format(codigo, "00000000")
        Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column) = extractedRef
    
    End If

    
    
  

'IMPORTES
'Total
    palabrabuscada = "Total"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 5
            XTR_TTL = celdaencontrada.Offset(i, 0)
            If XTR_TTL <> "" Then
                XTR_TTL = Trim(Replace(XTR_TTL, ".", ""))
                If IsNumeric(XTR_TTL) Then
                    Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column) = XTR_TTL * 1
                    Exit For
                End If
            End If
        Next i
    End If
    
'Subtotal
    palabrabuscada = "Subtotal"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 5
            XTR_SUB = celdaencontrada.Offset(i, 0)
            If XTR_SUB <> "" Then
                XTR_SUB = Trim(Replace(XTR_SUB, ".", ""))
                If IsNumeric(XTR_SUB) Then
                    Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column) = XTR_SUB * 1
                    Exit For
                End If
            End If
        Next i
    End If

'IVA
    palabrabuscada = "IVA" & Chr(10) & "Inscripto" & Chr(10) & "21,00 %"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 5
            XTR_IVA = celdaencontrada.Offset(i, 0)
            If XTR_IVA <> "" Then
                XTR_IVA = Trim(Replace(XTR_IVA, ".", ""))
                If IsNumeric(XTR_IVA) Then
                    Hoja2.Cells(y, ctx.rngIVA.Range.Column) = XTR_IVA * 1
                    Exit For
                End If
            End If
        Next i
    End If
    
    
'Percepción IIBB Capital Federal
    palabrabuscada = "Percepción IIBB Capital Federal"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 20
            XTR_PERC = celdaencontrada.Offset(0, i)
            If XTR_PERC <> "" And Right(XTR_PERC, 1) <> "%" Then
                XTR_PERC = Trim(Replace(XTR_PERC, ".", ""))
                If IsNumeric(XTR_PERC) Then
                    Hoja2.Cells(y, ctx.rngIIBBCABA.Range.Column) = XTR_PERC * 1
                    Exit For
                End If
            End If
        Next i
    End If
    
'Percepción IIBB Neuquen
    palabrabuscada = "Percepción IIBB Neuquen"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 20
            XTR_PERC = celdaencontrada.Offset(0, i)
            If XTR_PERC <> "" And Right(XTR_PERC, 1) <> "%" Then
                XTR_PERC = Trim(Replace(XTR_PERC, ".", ""))
                If IsNumeric(XTR_PERC) Then
                    Hoja2.Cells(y, ctx.rngIIBBNeuquen.Range.Column) = XTR_PERC * 1
                    Exit For
                End If
            End If
        Next i
    End If


'CAE
    palabrabuscada = "C.A.E.A. Nro: "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=2, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        XTR_CAE = Mid(celdaencontrada, Len(palabrabuscada) + 1, 14)

        VTO_CAE = Replace(Right(celdaencontrada, 10), "/", ".")
                
        Hoja2.Cells(y, ctx.rngCAE.Range.Column) = XTR_CAE

        Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column) = VTO_CAE
    
    End If


End Sub
