Attribute VB_Name = "modParserVendor09"

Sub ParseVendor09(hoja, y, Optional ctx As AppContext)

    Set ctx = ResolveContext(ctx)
    Dim startPos As Long
    Dim endPos As Long

    'Referencia VENDOR09
    palabrabuscada = "numero: "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        Referencia = Mid(celdaencontrada.Value, Len(palabrabuscada))
        Referencia = Replace(Replace(Referencia, " ", ""), "-", "A")
    End If

    '*** C O N T I N U A ***
    palabrabuscada = "*** C O N T I N U A ***"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then GoTo Hoja1
    
    'COD FC o NC
    palabrabuscada = "Código Nro. "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        COD = Mid(celdaencontrada, Len(celdaencontrada) - 1)

        If COD = "01" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "FC-REM"
    
        If COD = "03" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "NC-REM"

    End If
    
    'Importes: Total
    If COD = "01" Then
        palabrabuscada = "imp.total"
    ElseIf COD = "03" Then
        palabrabuscada = "importe total"
    End If
    
    'Importes: Total
    If Not hoja.UsedRange.Find(What:="imp.total", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False) Is Nothing Then palabrabuscada = "imp.total"
        
    If Not hoja.UsedRange.Find(What:="importe total", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False) Is Nothing Then palabrabuscada = "importe total"
        

    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        If COD = "01" Then
            For i = 1 To 8
                Total = celdaencontrada.Offset(0, i).Value
                If Total <> "" And Total <> "$" Then Exit For
            Next i
        ElseIf COD = "03" Then
            For i = 6 To -3 Step -1
                Total = celdaencontrada.Offset(-2, i).Value
                If Total <> "$" And Total <> "" Then Exit For
            Next i
        End If
        
        If Total = "" Then GoTo Hoja1
        
        'Total
        Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column).Value = Replace(Total, "-", "")

        If Total = "" Then GoTo Hoja1
        'II
        palabrabuscada = "I.INTERNOS"
        Set CeldaEncontradaColumna = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
        If Not CeldaEncontradaColumna Is Nothing Then
            For i = 1 To (celdaencontrada.Row - CeldaEncontradaColumna.Row)
                RES = hoja.Cells(celdaencontrada.Row - i, CeldaEncontradaColumna.Column)
                If RES <> "" And RES <> palabrabuscada Then
                    If IsNumeric(RES) Then
                        RES = Replace(RES, "-", "")
                        Hoja2.Cells(y, ctx.rngII.Range.Column).Value = RES * 1
                        Exit For
                    End If
                End If
            Next i
        End If
        'IVA
        palabrabuscada = "IVA 21%"
        Set CeldaEncontradaColumna = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
        If Not CeldaEncontradaColumna Is Nothing Then
            For i = 1 To (celdaencontrada.Row - CeldaEncontradaColumna.Row)
                RES = hoja.Cells(celdaencontrada.Row - i, CeldaEncontradaColumna.Column)
                If RES <> "" And RES <> palabrabuscada Then
                    If IsNumeric(RES) Then
                        Hoja2.Cells(y, ctx.rngIVA.Range.Column).Value = Replace(RES, "-", "")
                        Exit For
                    End If
                End If
            Next i
        End If

        'SUBTOTAL
        palabrabuscada = "SUBTOTAL"
        Set CeldaEncontradaColumna = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
        
        If Not CeldaEncontradaColumna Is Nothing Then
            For i = 1 To (celdaencontrada.Row - CeldaEncontradaColumna.Row)
                RES = hoja.Cells(celdaencontrada.Row - i, CeldaEncontradaColumna.Column)
                If RES <> "" And RES <> palabrabuscada Then
                    If IsNumeric(RES) Then
                        For j = 6 To 1 Step -1
                            buscarSubt = hoja.Cells(celdaencontrada.Row - i, CeldaEncontradaColumna.Column - j)
                            If Len(buscarSubt) = 1 Then
                                SUBT = SUBT + buscarSubt
                            End If
                        Next j
                        Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column).Value = SUBT + Replace(RES, "-", "")
                        Exit For
                    End If
                End If
            Next i
        End If
    End If

    Hoja2.Cells(y, ctx.rngReferencia.Range.Column).Value = Referencia

    'Cliente VENDOR09
    palabrabuscada = "PAN AMERICAN ENERGY"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        startPos = InStr(1, celdaencontrada.Value, "300", vbTextCompare)
        If startPos > 0 Then
            endPos = InStr(startPos, celdaencontrada.Value, " ", vbTextCompare)
            If endPos > 0 Then
                extractedData = Mid(celdaencontrada.Value, startPos, endPos - startPos)
            Else
                extractedData = Mid(celdaencontrada.Value, startPos)
            End If
            
            For Each fila In ctx.tblCORS.ListRows
                If CStr(fila.Range(ctx.tblCORS.ListColumns("Cliente VENDOR09").Range.Column)) = extractedData Then
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

    'Fecha
    palabrabuscada = "fecha:"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        extractedFECHA = Mid(celdaencontrada.Value, Len(palabrabuscada) + 1)
        extractedFECHA = Replace(extractedFECHA, ".", "/")
        fechaFormateada = Format(DateValue(extractedFECHA), "dd.mm.yyyy")
        Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column).Value = fechaFormateada
    End If
    


   

    'IIBB CABA
    palabrabuscada = "IB.CAP.FED"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 3
            RES = celdaencontrada.Offset(0, i).Value
            If RES <> "" Then
                If RES = "0,00" Then RES = ""
                Hoja2.Cells(y, ctx.rngIIBBCABA.Range.Column).Value = RES
                Exit For
            End If
        Next i
    End If

    'IIBB BSAS
    palabrabuscada = "IB.BS.AS."
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 3
            RES = celdaencontrada.Offset(0, i).Value
            If RES <> "" Then
                If RES = "0,00" Then RES = ""
                Hoja2.Cells(y, ctx.rngIIBBBSAS.Range.Column).Value = RES
                Exit For
            End If
        Next i
    End If
    
    'Remito ref.
    palabrabuscada = "Remito ref. "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        Referencia = Replace(celdaencontrada, "-", "R")
        Referencia = Mid(Referencia, 13)
        Referencia = Left(Referencia, 13)
        Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column).Value = Referencia
    End If

    'CAE VENDOR09
    palabrabuscada = "C.A.E.A. NRO."
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        textoExtraido = Mid(celdaencontrada.Value, InStr(celdaencontrada.Value, palabrabuscada) + Len(palabrabuscada), 14)
        Hoja2.Cells(y, ctx.rngCAE.Range.Column).Value = Replace(textoExtraido, " ", "")
        Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column).Value = Right(celdaencontrada, 10)
    End If
    
    
    Exit Sub
    
Hoja1:

    nombreArchivoNuevo = Referencia & "-Hoja 1.pdf"
    If nombreArchivoNuevo <> ctx.NombreArchivo Then
        If Dir(ctx.rutaCarpeta & nombreArchivoNuevo) <> "" Then nombreArchivoNuevo = Referencia & "-Hoja 2.pdf"
        If Dir(ctx.rutaCarpeta & ctx.NombreArchivo) <> "" Then
            Name ctx.rutaCarpeta & ctx.NombreArchivo As ctx.rutaCarpeta & nombreArchivoNuevo
        End If
    End If
    
End Sub
