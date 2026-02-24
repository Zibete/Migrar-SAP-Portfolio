Attribute VB_Name = "modParserVendor06"

Sub ParseVendor06(hoja, y, Optional ctx As AppContext)


    'IMPORTES
    Set ctx = ResolveContext(ctx)
    palabrabuscada = "ERC.IVA"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=2, MatchCase:=True)
    If Not celdaencontrada Is Nothing Then
        'TOTAL 0
        Set segundaCeldaEncontrada = hoja.Rows(celdaencontrada.Row).Find(What:="AL", LookIn:=xlValues, LookAt:=1, MatchCase:=True)
        If Not segundaCeldaEncontrada Is Nothing Then
            If segundaCeldaEncontrada.Offset(0, -i) = "TOT" Then
                For i = 0 To 5
                    RES = segundaCeldaEncontrada.Offset(1, -i)
                    If RES <> "" And RES <> 0 And IsNumeric(RES) Then
                        RES = Replace(Replace(RES, ",", ""), ".", ",")
                        If IsNumeric(RES) Then
                            Total = RES & Total
                        End If
                    End If
                Next i
                If Total <> "" Then Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column).Value = Total * 1
            End If
        End If
        
        
        'TOTAL 1
        If Total = "" Then
            Set segundaCeldaEncontrada = hoja.Rows(celdaencontrada.Row).Find(What:="TOTAL", LookIn:=xlValues, LookAt:=1, MatchCase:=True)
            If Not segundaCeldaEncontrada Is Nothing Then
            
                textoEncontrado = segundaCeldaEncontrada.Offset(0, -2) & segundaCeldaEncontrada.Offset(0, -1) & segundaCeldaEncontrada
                
                If Trim(textoEncontrado) = "SUBTOTAL" Then Set segundaCeldaEncontrada = hoja.Cells.FindNext(segundaCeldaEncontrada)
    
                    If Not segundaCeldaEncontrada Is Nothing Then
                        For i = -3 To 2
                            RES = segundaCeldaEncontrada.Offset(1, i)
                            If RES <> "" And RES <> 0 And IsNumeric(RES) Then
                                RES = Replace(Replace(RES, ",", ""), ".", ",")
                                If IsNumeric(RES) Then

                                    Total = Total & RES
                                    
                                End If
                            End If
                        Next i
                    End If
                End If
            
            Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column).Value = Total * 1
        
        End If

        'TOTAL 2
        If Total = "" Then
            palabrabuscada = "*Otros:"
            Set Otros = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=True)
            If Not Otros Is Nothing Then
                For i = 0 To 20
                    RES = Otros.Offset(i, 0)
                    If RES <> "" And RES <> 0 And IsNumeric(RES) Then
                        Total = Replace(Replace(RES, ",", ""), ".", ",")
                        If IsNumeric(Total) Then
                            Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column).Value = Total * 1
                            Exit For
                        End If
                    End If
                Next i
            End If
        End If


        'IVA
        palabrabuscada = "INSC."
        Set segundaCeldaEncontrada = hoja.Rows(celdaencontrada.Row).Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=True)
        If Not segundaCeldaEncontrada Is Nothing Then
            For i = 1 To 5
                For j = -15 To 4
                    RES = Replace(segundaCeldaEncontrada.Offset(i, j), ".", ",")
                    If IsNumeric(RES) Then
                        texto = texto + RES
                    End If
                Next j
                If texto <> "" Then Exit For
            Next i
            
            IVA = texto
            
            If texto <> "" Then
            
                If Len(texto) - Len(Replace(texto, ",", "")) > 1 Then
                
                    posUltimaComa = InStrRev(texto, ",")
                    posAnteultimaComa = InStrRev(texto, ",", posUltimaComa - 1)
                    ultimosDosDecimales = Mid(texto, posUltimaComa + 1, 2)
                    Resultado = Mid(texto, posAnteultimaComa + 3, posUltimaComa - posAnteultimaComa - 3) & "," & ultimosDosDecimales

                    IVA = Resultado

                End If

            End If
            
            If IVA <> "" Then Hoja2.Cells(y, ctx.rngIVA.Range.Column).Value = IVA * 1
            
        End If
        

        'IVA 2
        If IVA = "" Then
            palabrabuscada = "*Otros:"
            Set Otros = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=True)
            If Not Otros Is Nothing Then
                For i = 0 To 20
                    RES = Otros.Offset(-1, -i)
                    If RES <> "" And RES <> 0 And IsNumeric(RES) Then
                        IVA = Replace(Replace(RES, ",", ""), ".", ",")
                        If IsNumeric(IVA) Then
                            Hoja2.Cells(y, ctx.rngIVA.Range.Column).Value = IVA * 1
                            Exit For
                        End If
                    End If
                Next i
            End If
        End If

        'SUBTOTAL 1
        intNoGrav = False
        
        Set intNoGravEncontrado = hoja.UsedRange.Find(What:="INT", LookIn:=xlValues, LookAt:=1, MatchCase:=True)
        
        For i = 1 To 2
        
            For Each RES In hoja.Range(hoja.Cells(segundaCeldaEncontrada.Offset(i, 0).Row, 1), hoja.Cells(segundaCeldaEncontrada.Offset(i, 0).Row, hoja.Columns.Count))
                
                RES = Replace(Replace(RES, ",", ""), ".", ",")


                If Right(II, 3) Like ",##" Then Exit For
                
                If IsNumeric(RES) And RES <> "" Then
                    If Subtotal <> "" Then

                        If RES = "0,00" And Not intNoGrav Then
                            intNoGrav = True
                        Else
                            If intNoGrav Then

                                II = II & RES
                                
                            Else
                            
                                Subtotal = Subtotal & RES

                            End If
                        End If
                    Else
                        If RES <> 0 Then Subtotal = RES
                    End If
                End If
            Next RES
        
        Next i
        

        Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column) = Subtotal * 1
        Hoja2.Cells(y, ctx.rngII.Range.Column) = II * 1
        
    End If
    
    
    palabrabuscada = "%"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    
    If Not celdaencontrada Is Nothing Then
        For h = 0 To 8
            RES = ""
            
            If celdaencontrada.Offset(h, 0) <> "%" Then Exit For
            
            For i = 20 To 1 Step -1
                RES = RES & celdaencontrada.Offset(h, -i)
            Next i

            valorIIBB = ""
            
            For i = 2 To 20
                RESIIBB = Replace(Replace(celdaencontrada.Offset(h, i), ",", ""), ".", ",")
                If RESIIBB <> "" And IsNumeric(RESIIBB) Then
                   valorIIBB = valorIIBB & RESIIBB
                End If
            Next i
            
            valorIIBB = valorIIBB * 1
             
            If RES <> "" Then
            
                If InStr(1, RES, "CABA", vbTextCompare) > 0 Then IIBBCABA = valorIIBB
                If InStr(1, RES, "Cord", vbTextCompare) > 0 Then IIBBCordoba = valorIIBB
                If InStr(1, RES, "Neuq", vbTextCompare) > 0 Then IIBBNeuquen = valorIIBB
                If InStr(1, RES, "Catam", vbTextCompare) > 0 Then IIBBCatamarca = valorIIBB
                If InStr(1, RES, "Salta", vbTextCompare) > 0 Then IIBBSalta = valorIIBB
                If InStr(1, RES, "Ctes", vbTextCompare) > 0 Then IIBBCorrientes = valorIIBB
                If InStr(1, RES, "Entre Rios", vbTextCompare) > 0 Then IIBBEntreRios = valorIIBB
                If InStr(1, RES, "Mendoza", vbTextCompare) > 0 Then IIBBMendoza = valorIIBB
                If InStr(1, RES, "Perc.Munic.", vbTextCompare) > 0 Then MuniCord = valorIIBB

            End If
        Next h
    End If

    'Referencia
    palabrabuscada = "NRO. "
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=2, MatchCase:=True)
    If Not celdaencontrada Is Nothing Then
        extractedRef = Mid(celdaencontrada, Len(palabrabuscada) + 1)
        If extractedRef <> "" And IsNumeric(Left(extractedRef, 1)) Then
            extractedRef = Replace(extractedRef, "-", "A")
        End If
    End If

    If IIBBCABA <> "" Then Hoja2.Cells(y, ctx.rngIIBBCABA.Range.Column).Value = IIBBCABA
    If IIBBCordoba <> "" Then Hoja2.Cells(y, ctx.rngIIBBCordoba.Range.Column).Value = IIBBCordoba
    If IIBBNeuquen <> "" Then Hoja2.Cells(y, ctx.rngIIBBNeuquen.Range.Column).Value = IIBBNeuquen * 1
    If MuniCord <> "" Then Hoja2.Cells(y, ctx.rngMuniCord.Range.Column).Value = MuniCord * 1
    If IIBBCatamarca <> "" Then Hoja2.Cells(y, ctx.rngIIBBCatamarca.Range.Column).Value = IIBBCatamarca * 1
    If IIBBEntreRios <> "" Then Hoja2.Cells(y, ctx.rngIIBBEntreRios.Range.Column).Value = IIBBEntreRios * 1
    If IIBBMendoza <> "" Then Hoja2.Cells(y, ctx.rngIIBBMendoza.Range.Column).Value = IIBBMendoza * 1
    If IIBBSalta <> "" Then Hoja2.Cells(y, ctx.rngIIBBSalta.Range.Column).Value = IIBBSalta * 1
    If IIBBCorrientes <> "" Then Hoja2.Cells(y, ctx.rngIIBBCorrientes.Range.Column).Value = IIBBCorrientes * 1
    
    If Total = "" Then
        'Hoja 1
        percepciones = Array( _
                            "IIBBCABA", IIBBCABA, "IIBBCordoba", IIBBCordoba, _
                            "IIBBNeuquen", IIBBNeuquen, "MuniCord", MuniCord, _
                            "IIBBCatamarca", IIBBCatamarca, "IIBBEntreRios", IIBBEntreRios, _
                            "IIBBSalta", IIBBSalta, "IIBBCorrientes", IIBBCorrientes, _
                            "IIBBMendoza", IIBBMendoza)

        If Not ctx.diccDocumentos.Exists(extractedRef) Then
            If Not IsEmpty(percepciones) And Not IsNull(percepciones) Then
                ctx.diccDocumentos.Add extractedRef, percepciones
            End If
        End If
        
        nombreArchivoNuevo = extractedRef & "-Hoja 1.pdf"
        
        If nombreArchivoNuevo <> ctx.NombreArchivo Then
        
            If Dir(ctx.rutaCarpeta & nombreArchivoNuevo) <> "" Then
            
                If GetEliminarDuplicados() = FLAG_SI Then
                
                    'eliminar
                    Kill ctx.rutaCarpeta & nombreArchivoNuevo
                
                ElseIf GetEliminarDuplicados() = "NO" Then
                
                    'cambiar nombre
                    For i = 1 To 5
                        If Dir(ctx.rutaCarpeta & extractedRef & "-Hoja 1-" & i & ".pdf") = "" Then
                            nombreArchivoNuevo = extractedRef & "-Hoja 1-" & i & ".pdf"
                            Exit For
                        End If
                    Next i
                
                End If
            
            End If
            
            If Dir(ctx.rutaCarpeta & ctx.NombreArchivo) <> "" Then
                Name ctx.rutaCarpeta & ctx.NombreArchivo As ctx.rutaCarpeta & nombreArchivoNuevo
            End If
        End If
        
        Exit Sub
        
    Else
    
        'Hoja totales
        Hoja2.Cells(y, ctx.rngReferencia.Range.Column).Value = extractedRef
        Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column).Value = extractedRef
        
    End If
    

    'SITE
    If Hoja2.Cells(y, ctx.rngSite.Range.Column).Value = "" Then
    
        palabrabuscada = "CLIENTE:"
        Do While Len(palabrabuscada) > 0
            Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
            If Not celdaencontrada Is Nothing Then Exit Do
            palabrabuscada = Left(palabrabuscada, Len(palabrabuscada) - 1)
        Loop
    
        If Not celdaencontrada Is Nothing Then
            For i = 0 To 6
                For j = 1 To 6
                    RES = celdaencontrada.Offset(i, j)
                    If RES <> "" And IsNumeric(RES) And Len(RES) = 6 Then
                        Hoja2.Cells(y, ctx.rngNuevaRuta.Range.Column) = RES
                        site = RES
                        For Each fila In ctx.tblCORS.ListRows
                            If CStr(UCase(fila.Range(ctx.tblCORS.ListColumns("Cliente VENDOR06").Range.Column).Value)) = CStr(UCase(site)) Then
                                Hoja2.Cells(y, ctx.rngTexto.Range.Column).Value = fila.Range(ctx.tblCORS.ListColumns("Texto").Range.Column)
                                Hoja2.Cells(y, ctx.rngCeBe.Range.Column).Value = fila.Range(ctx.tblCORS.ListColumns("CeBe").Range.Column)
                                Hoja2.Cells(y, ctx.rngNombreSite.Range.Column).Value = fila.Range(ctx.tblCORS.ListColumns("Nombre Sucursal").Range.Column)
                                Hoja2.Cells(y, ctx.rngSupl.Range.Column).Value = fila.Range(ctx.tblCORS.ListColumns("Supl.").Range.Column)
                                Hoja2.Cells(y, ctx.rngSite.Range.Column).Value = fila.Range(ctx.tblCORS.ListColumns("Sucursal").Range.Column)
                                Hoja2.Cells(y, ctx.rngZona.Range.Column).Value = fila.Range(ctx.tblCORS.ListColumns("Zona").Range.Column)
                                Hoja2.Cells(y, ctx.rngAN.Range.Column).Value = fila.Range(ctx.tblCORS.ListColumns("AN").Range.Column)
                                Hoja2.Cells(y, ctx.rngMails.Range.Column).Value = fila.Range(ctx.tblCORS.ListColumns("Mails").Range.Column)
                                Exit For
                            End If
                        Next fila
                        Exit For
                    End If
                Next j
                If site <> "" Then Exit For
            Next i
        End If
        
    End If



    'Fecha
    If Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column) = "" Then
        palabrabuscada = "FECHA:"
        Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=2, MatchCase:=False)
        If Not celdaencontrada Is Nothing Then
            'Fecha
            For i = 1 To 10
                If celdaencontrada.Offset(-1, i) <> "" Then
                    extractedFECHA = celdaencontrada.Offset(-1, i)
                    If IsDate(extractedFECHA) Then fechaFormateada = Format(DateValue(extractedFECHA), "dd.mm.yyyy")
                    Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column) = fechaFormateada 'Fecha
                    Exit For
                End If
            Next i
            'Fecha 2
            If fechaFormateada = "" Then
                extractedFECHA = Replace(celdaencontrada, palabrabuscada, "")
                If IsDate(extractedFECHA) Then fechaFormateada = Format(DateValue(extractedFECHA), "dd.mm.yyyy")
                Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column) = fechaFormateada 'Fecha
            End If
            
        End If
    End If
    
    'COD FC o NC
    If Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column) = "" Then
        palabrabuscada = "A"
        Set celdaCOD = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
        
        If Not celdaCOD Is Nothing Then
            For h = -4 To 4
                For i = 1 To 5
                    RES = celdaCOD.Offset(i, h)
                    If RES <> "" Then
                        COD = Right(RES, 2)
                        If COD = "01" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "FC-REC"
                        If COD = "02" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "ND-COM"
                        If COD = "03" Then
                            Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column).Value = "NC-FAL" 'COD FC o NC
                            
                            palabrabuscada = "Cte."
                            Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=2, MatchCase:=True)
                    
                            If Not celdaencontrada Is Nothing Then
                                texto = ""
                                For A = 0 To 20
                                    texto = texto & celdaencontrada.Offset(0, A)
                                Next A
                                
                                PDV = Mid(texto, InStr(texto, "-") + 1, 4)
                                nroComp = Mid(texto, InStrRev(texto, "-"))
                                nroComp = Replace(nroComp, "-", "")
                                nroComp = Right("00000000" & nroComp, 8)
                                RES = Trim(PDV & "A" & nroComp)
                                Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column).Value = RES 'RemitoRef
                                
                            End If
                        End If
                        Exit For
                    End If
                Next i
                If COD <> "" Then Exit For
            Next h
        End If
    End If
    
    
    'CAE 1
    palabrabuscada = "CAEA"
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 10
            RES = celdaencontrada.Offset(0, i)
            If RES <> "" Then
                RES = Replace(RES, ":", "")
                If IsNumeric(RES) Then
                    extractedCAE = extractedCAE & RES
                End If
            End If
        Next i
        Hoja2.Cells(y, ctx.rngCAE.Range.Column).Value = extractedCAE
    Else
        'CAE 2
        palabrabuscada = "EA"
        Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
        If Not celdaencontrada Is Nothing Then
            Set CeldaAnterior = celdaencontrada.Offset(0, -1)
            If CeldaAnterior = "CA" Then
                For i = 1 To 10
                    RES = CeldaAnterior.Offset(0, i)
                    If RES <> "" Then
                        RES = Replace(RES, ":", "")
                        If IsNumeric(RES) Then
                            extractedCAE = extractedCAE & RES
                        End If
                    End If
                Next i
                Hoja2.Cells(y, ctx.rngCAE.Range.Column).Value = extractedCAE
            End If
        End If
    End If
    
    
    'VTO.: CAE
    palabrabuscada = "Vto."
    Set celdaencontrada = hoja.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=1, MatchCase:=False)
    If Not celdaencontrada Is Nothing Then
        For i = 1 To 5
            If celdaencontrada.Offset(0, i) <> "" Then
                extractedVtoCAE = celdaencontrada.Offset(0, i)
                If IsDate(extractedVtoCAE) Then
                    extractedVtoCAE = Format(DateValue(extractedVtoCAE), "dd.mm.yyyy")
                    Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column) = extractedVtoCAE
                    Exit For
                End If
            End If
        Next i
    End If
    


End Sub


