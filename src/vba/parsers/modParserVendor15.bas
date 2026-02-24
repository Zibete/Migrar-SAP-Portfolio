Attribute VB_Name = "modParserVendor15"

Sub ParseVendor15(hoja, y, Optional ctx As AppContext)

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
    palabrabuscada = "COD. "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then

        COD = Mid(celdaencontrada, Len(palabrabuscada) + 1)
        
        If COD = "01" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "FC-REC"
        If COD = "02" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "ND-ARR"
        If COD = "03" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "NC-FAL"
        
        If COD = "201" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "FCE-REC"
        If COD = "202" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "NDE-ARR"
        If COD = "203" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "NCE-FAL"
        
    End If

'Fecha
    palabrabuscada = "Fecha: "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        extractedFECHA = Mid(celdaencontrada, Len(palabrabuscada) + 1)
        If IsDate(extractedFECHA) Then fechaFormateada = Format(DateValue(extractedFECHA), "dd.mm.yyyy")
        Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column).Value = fechaFormateada
    End If
    
'Referencia
    palabrabuscada = "Número: "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        extractedRef = Mid(celdaencontrada, Len(palabrabuscada) + 1)
        extractedRef = Replace(extractedRef, "-", "A")
        Hoja2.Cells(y, ctx.rngReferencia.Range.Column) = extractedRef
        Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column) = extractedRef
    End If
    
'Referencia RtoRefg
    palabrabuscada = "Referencia: "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        For i = 1 To Len(celdaencontrada)
            If Mid(celdaencontrada, i, 1) Like "[0-9]" Then
                extractedRef = Mid(celdaencontrada, i, 14)
                Exit For
            End If
        Next i
 
        Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column) = Replace(extractedRef, "-", "A")
        
    End If
    

'IMPORTES
'Total
    palabrabuscada = "Total"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 8
            XTR_TTL = celdaencontrada.Offset(0, i)
            If XTR_TTL <> "" Then
                XTR_TTL = Trim(Replace(Replace(XTR_TTL, ".", ""), "$", ""))
                If IsNumeric(XTR_TTL) Then
                    
                    
                    XTR_IVA = Trim(Replace(Replace(celdaencontrada.Offset(-1, i), ".", ""), "$", ""))
                    XTR_SUB = Trim(Replace(Replace(celdaencontrada.Offset(-2, i), ".", ""), "$", ""))
                    
                    
                    Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column) = XTR_TTL * 1
                    Hoja2.Cells(y, ctx.rngIVA.Range.Column) = XTR_IVA * 1
                    Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column) = XTR_SUB * 1
                    
                    
                    Exit For
                End If
            End If
        Next i
    End If

'CAE
    palabrabuscada = "CAE: "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=2, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        XTR_CAE = Mid(celdaencontrada, Len(palabrabuscada) + 1)
        VTO_CAE = Replace(Right(celdaencontrada.Offset(1, 0), 10), "/", ".")
        
        Hoja2.Cells(y, ctx.rngCAE.Range.Column) = XTR_CAE

        Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column) = VTO_CAE
    
    End If


End Sub
