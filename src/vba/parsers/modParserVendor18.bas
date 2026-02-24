Attribute VB_Name = "modParserVendor18"

Sub ParseVendor18(hoja, y, Optional ctx As AppContext)

    'Cliente <SUPPLIER_A>
    Set ctx = ResolveContext(ctx)
    palabrabuscada = "<REDACTED>"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        For i = 0 To 5
            extractedData = celdaencontrada.Offset(-1, -i)
            If extractedData <> "" Then
                
                For Each fila In ctx.tblCORS.ListRows
                    If CStr(fila.Range(ctx.tblCORS.ListColumns("Cliente <SUPPLIER_A>").Range.Column).Value) = extractedData Then
                        site = fila.Range(ctx.tblCORS.ListColumns("Sucursal").Range.Column)
                        Call asignarCORS(y, site)
                        Exit For
                    End If
                Next fila

                Exit For
            End If
        Next i

        For i = 1 To 6
            RES = celdaencontrada.Offset(-i, 0)
            If RES <> "" And Left(RES, 6) <> "Column" And RES <> extractedData Then
                extractedFECHA = RES
                extractedRef = celdaencontrada.Offset(-i - 1, 0)
                Exit For
            End If
            If Left(RES, 6) = "Column" Then Exit For
        Next i
        
        If extractedRef = "" And extractedFECHA = "" Then
             
            For i = 1 To 6
                RES = celdaencontrada.Offset(-i, -1)
                If RES <> "" And Left(RES, 6) <> "Column" And RES <> extractedData Then
                    extractedFECHA = RES
                    extractedRef = celdaencontrada.Offset(-i - 1, -1)
                    Exit For
                End If
                If Left(RES, 6) = "Column" Then Exit For
            Next i
        
        End If
    End If
    
    If extractedRef <> "" And extractedFECHA <> "" Then
        'Fecha
        If IsDate(extractedFECHA) Then fechaFormateada = Format(DateValue(extractedFECHA), "dd.mm.yyyy")
        Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column).Value = fechaFormateada
        'Referencia
        extractedRef = Replace(extractedRef, "A", "")
        extractedRef = Replace(extractedRef, "-", "A")
        Hoja2.Cells(y, ctx.rngReferencia.Range.Column) = extractedRef
        Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column) = extractedRef
    
    End If
    
    'REF <SUPPLIER_A>
    palabrabuscada = "REF: FAC"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        'ES NC
        palabrabuscada = "ROT"
        Set celdaEncontrada2 = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=True)
        If Not celdaEncontrada2 Is Nothing Then
            Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column) = "NC-DEV"
            extractedRef = Replace(celdaEncontrada2, "ROT", "")
            Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column) = Trim(extractedRef)
        Else
            Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column) = "NC-FAL"
            texto = Replace(celdaencontrada, "REF:", "")
            texto = Replace(texto, "FAC", "")
            texto = Replace(texto, "A", "")
            If texto <> "" Then texto = Trim(Left(texto, Len(texto) - 8) & "A" & Right(texto, 8))

            Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column) = texto
        End If
    Else
        Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column) = "FC-REC"
    End If

    Dim datos(3)
  
    'CAE
    palabrabuscada = "C.A.E. "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
    
        extractedCAE = Mid(celdaencontrada, 8)
        extractedVtoCAE = Mid(celdaencontrada.Offset(1, 0), 6)
        
    Else
    
        If Not hoja.UsedRange.Find(What:="CAE", LookIn:=xlValues, LookAt:=1, MatchCase:=False) Is Nothing Then
            Set celdaencontrada = hoja.UsedRange.Find(What:="CAE", LookIn:=xlValues, LookAt:=1, MatchCase:=False)
        ElseIf Not hoja.UsedRange.Find(What:="CAEA", LookIn:=xlValues, LookAt:=1, MatchCase:=False) Is Nothing Then
            Set celdaencontrada = hoja.UsedRange.Find(What:="CAEA", LookIn:=xlValues, LookAt:=1, MatchCase:=False)
        End If
        
        If Not celdaencontrada Is Nothing Then
        
            For i = 1 To 6
                If celdaencontrada.Offset(0, i) <> "" Then
                    extractedCAE = celdaencontrada.Offset(0, i)
                    Exit For
                End If
            Next i
            
            For i = 1 To 6
                If celdaencontrada.Offset(1, i) <> "" Then
                    extractedVtoCAE = celdaencontrada.Offset(1, i)
                    Exit For
                End If
            Next i
            
        End If

    End If
    
    If IsDate(extractedVtoCAE) Then extractedVtoCAE = Format(DateValue(extractedVtoCAE), "dd.mm.yyyy")
    
    Hoja2.Cells(y, ctx.rngCAE.Range.Column).Value = extractedCAE
    Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column).Value = extractedVtoCAE
    
    For i = 2 To 13
        If celdaencontrada.Offset(-1, i) <> "" And celdaencontrada.Offset(-1, i) <> RES Then
            RES = celdaencontrada.Offset(-1, i)
            b = b + 1
            datos(b) = celdaencontrada.Offset(-1, i)
            If datos(3) <> "" Then Exit For
        End If
    Next i
    
    If Left(Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column), 2) <> "FC" Then
        datos(1) = Replace(Replace(datos(1), ",", ""), ".", ",")
        datos(2) = Replace(Replace(datos(2), ",", ""), ".", ",")
        datos(3) = Replace(Replace(datos(3), ",", ""), ".", ",")
    End If
 
    'SUBTOTAL
    If datos(1) <> "" And datos(1) <> 0 Then Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column).Value = datos(1) * 1
    'IVA
    If datos(2) <> "" And datos(2) <> 0 Then Hoja2.Cells(y, ctx.rngIVA.Range.Column).Value = datos(2) * 1
    'TOTAL
    If datos(3) <> "" And datos(3) <> 0 Then Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column).Value = datos(3) * 1

    palabrabuscada = "AGIP RG GRUPO"

    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=2, MatchCase:=False)
    
    If Not celdaencontrada Is Nothing Then
    

        For i = 6 To 1 Step -1
            If celdaencontrada.Offset(0, i) <> "" Then
                If celdaencontrada.Offset(0, i) <> "" Then
                    Hoja2.Cells(y, ctx.rngIIBBCABA.Range.Column) = celdaencontrada.Offset(0, i) * 1
                    Exit For
                End If
            End If
        Next i
        
    End If

End Sub
