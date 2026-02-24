Attribute VB_Name = "modValidation"

Sub probarComp()
    ComprobarEstados (Selection.Row)
End Sub

Sub ComprobarEstados(i)
   
    'asignaciones
    
    gCtx.nuevoNombre = Hoja2.Cells(i, gCtx.rngComentarios_User.Range.Column)
    
    tipoDoc = Hoja2.Cells(i, gCtx.rngTipoDoc.Range.Column)
    tieneScan = Hoja2.Cells(i, gCtx.rngTieneScan_SB.Range.Column)
    DiferenciaCostos = Hoja2.Cells(i, gCtx.rngDifCostos.Range.Column)
    EstadoDelPagoSB = Hoja2.Cells(i, gCtx.rngEstadoDelPago_SB.Range.Column)
    fechaNeg_SB = Hoja2.Cells(i, gCtx.rngFechaNeg_SB.Range.Column)
    remitoRef = Hoja2.Cells(i, gCtx.rngRemitoRef.Range.Column)
    DifConNC = Hoja2.Cells(i, gCtx.rngDifConNC.Range.Column)

    
    If Hoja2.Cells(i, gCtx.rngFechaDeFactura.Range.Column) <> "" Then FechaFC = CDate(Replace(Hoja2.Cells(i, gCtx.rngFechaDeFactura.Range.Column), ".", "/"))
    If Hoja2.Cells(i, gCtx.rngFechaDoc_SB.Range.Column) <> "" Then FechaSB = CDate(Replace(Hoja2.Cells(i, gCtx.rngFechaDoc_SB.Range.Column), ".", "/"))
    
    siteFC = CStr(Hoja2.Cells(i, gCtx.rngSite.Range.Column))
    SiteSB = CStr(Hoja2.Cells(i, gCtx.rngSite_SB.Range.Column))
    
    totalFC = Hoja2.Cells(i, gCtx.rngTotalBrutoFactura.Range.Column)
    totalSB = Hoja2.Cells(i, gCtx.rngTotalBruto_SB.Range.Column)

    subtotalFC = Round(Hoja2.Cells(i, gCtx.rngSubtotalFactura.Range.Column) * 1 + Hoja2.Cells(i, gCtx.rngII.Range.Column) * 1, 2)
    
    II = Hoja2.Cells(i, gCtx.rngII.Range.Column)
    
    subtotalSB = Hoja2.Cells(i, gCtx.rngSubtotal_SB.Range.Column)
    
    'comentario = ""
    If tieneScan = "NO" Then
        Hoja2.Cells(i, gCtx.rngEstadoDelPago.Range.Column) = ESTADO_ERROR_SCAN
        gCtx.nuevoNombre = sumarNuevoNombre("Sin Scan", gCtx.nuevoNombre)
    End If

'    If EstadoDelPagoSB <> "" Then
        If CoreShouldMarkPendienteRevisar(DiferenciaCostos, DifConNC, gCtx.montoToleranciaSB) Then
            Hoja2.Cells(i, gCtx.rngEstadoDelPago.Range.Column) = ESTADO_PENDIENTE_REVISAR
        End If
'    End If
    If CoreIsSiteMismatch(SiteSB, siteFC) Then
        gCtx.nuevoNombre = sumarNuevoNombre(CoreBuildSiteMismatchComment(tipoDoc, siteFC), gCtx.nuevoNombre)
        Hoja2.Cells(i, gCtx.rngEstadoDelPago.Range.Column) = ESTADO_PENDIENTE_REINGRESO
    End If
    
    
    If Right(tipoDoc, 3) = "REC" Then
        If FechaSB <> "" And FechaFC <> "" Then
            If FechaSB <> FechaFC Then
                Hoja2.Cells(i, gCtx.rngEstadoDelPago.Range.Column) = ESTADO_PENDIENTE_REINGRESO
                gCtx.nuevoNombre = sumarNuevoNombre("Error en fecha de " & Left(tipoDoc, 2) & " (" & FechaFC & ")", gCtx.nuevoNombre)
            End If
        End If
        If totalSB <> "" And totalFC <> "" And totalFC <> 0 Then
            If totalSB <> totalFC Then
                Hoja2.Cells(i, gCtx.rngEstadoDelPago.Range.Column) = ESTADO_PENDIENTE_REINGRESO
                gCtx.nuevoNombre = sumarNuevoNombre("Error en total de " & Left(tipoDoc, 2) & " (" & totalFC & ")", gCtx.nuevoNombre)
            End If
        End If
        If subtotalSB <> "" And subtotalFC <> "" And subtotalFC <> 0 Then
            If subtotalSB <> subtotalFC Then
                Hoja2.Cells(i, gCtx.rngEstadoDelPago.Range.Column) = ESTADO_PENDIENTE_REINGRESO
                gCtx.nuevoNombre = sumarNuevoNombre("Error en subtotal de " & Left(tipoDoc, 2) & " (" & subtotalFC & ")", gCtx.nuevoNombre)
            End If
        End If
    End If
    
    doaMessage = CoreBuildDoaMessage(fechaNeg_SB, totalSB, gCtx.montoDOA, Date)
    If doaMessage <> "" Then
        SetRowStatus i, "", doaMessage
    End If
    
    'Error en Referencia
    errorReferencia = CoreIsErrorReferencia( _
        GetVendorFilter(), _
        remitoRef, _
        tipoDoc, _
        Referencia, _
        gCtx.largoReferencia, _
        gCtx.letra _
    )
    
    If errorReferencia = True Then
        Hoja2.Cells(i, gCtx.rngEstadoDelPago.Range.Column) = ESTADO_PENDIENTE_REINGRESO
        gCtx.nuevoNombre = sumarNuevoNombre("Error en Referencia", gCtx.nuevoNombre)
    End If


    If EstadoDelPagoSB <> "" Then
        Hoja2.Cells(i, gCtx.rngEstadoDelPago.Range.Column) = EstadoDelPagoSB
    End If

    
    If Left(gCtx.nuevoNombre, 1) = "-" Then gCtx.nuevoNombre = Mid(gCtx.nuevoNombre, 2)
    Hoja2.Cells(i, gCtx.rngComentarios_User.Range.Column) = gCtx.nuevoNombre
    gCtx.rngComentarios_User.Range.Columns.AutoFit

End Sub
Sub VerificarUnoSolo()

    asignaciones

    Set fila = gCtx.tblDatos.ListRows(Selection.Row - gCtx.tblDatos.HeaderRowRange.Row)
    Call VerificarDatos(fila)

End Sub


Sub VerificarDatos(Optional filaEspecifica As ListRow = Nothing)

    Dim filas As Collection
    Dim fila As ListRow
    Set filas = New Collection

    If Not filaEspecifica Is Nothing Then
        filas.Add filaEspecifica
    Else
        For Each fila In gCtx.tblDatos.ListRows
            If fila.Range.Cells(1, gCtx.rngReferencia.index) = "" Then Exit For
            filas.Add fila
        Next fila
    End If

    gCtx.ControlarCambios = False
 
    For Each fila In filas
        With fila.Range
            estadoPago = .Cells(1, gCtx.rngEstadoDelPago.index)
            comentarios_User = .Cells(1, gCtx.rngComentarios_User.index)
            diferenciaSAP = .Cells(1, gCtx.rngDifSap.index)
            subtotalIVA21 = .Cells(1, gCtx.rngSubtotalFactura.index)
            subtotalIVA105 = .Cells(1, gCtx.rngSubtotalFactura105.index)
            CAE = .Cells(1, gCtx.rngCAE.index)
            vtoCAE = .Cells(1, gCtx.rngVTOCAE.index)
            Referencia = .Cells(1, gCtx.rngReferencia.index)
            rtoRef = .Cells(1, gCtx.rngRemitoRef.index)
            gCtx.NombreArchivo = .Cells(1, gCtx.rngNombreArchivo.index)
            site = .Cells(1, gCtx.rngSite.index)
            fechaDoc = .Cells(1, gCtx.rngFechaDeFactura.index)
            fechaNeg = .Cells(1, gCtx.rngFechaNeg_SB.index)
            fechaBase = .Cells(1, gCtx.rngFechaBase.index)
            Total = .Cells(1, gCtx.rngTotalBrutoFactura.index)
            CompensaciÃ³n = .Cells(1, gCtx.rngCompensacion.index)
            MENSAJE = .Cells(1, gCtx.rngMensajesSap.index)
            tipoDoc = .Cells(1, gCtx.rngTipoDoc.index)
            pagado = .Cells(1, gCtx.rngPagado.index)

            camposClave = True
            
            If EsVacio(site) Then camposClave = False
            If EsVacio(fechaDoc) Then camposClave = False
            If EsVacio(Referencia) Then camposClave = False
            
            If Right(tipoDoc, 3) = "REM" Then
                If Len(rtoRef) <> gCtx.largoReferencia Or InStr(1, rtoRef, gCtx.letra) = 0 Then
                    camposClave = False
                End If
            End If
            
            If GetVendorFilter() <> "Varios" Then
                If Right(tipoDoc, 3) <> "REM" Then If Len(Referencia) <> gCtx.largoReferencia Or InStr(1, Referencia, gCtx.letra) = 0 Then camposClave = False
            Else
                If Right(tipoDoc, 3) <> "REM" Then
                    If Len(Referencia) <= 12 Or (InStr(1, Referencia, "A") = 0 And InStr(1, Referencia, "C") = 0) Then camposClave = False
                Else
                    If Len(Referencia) <= 12 Or InStr(1, Referencia, "R") = 0 Then camposClave = False
                End If
            End If
            
            If EsVacio(Total) Then camposClave = False
            If EsVacio(CAE) Then camposClave = False
            If EsVacio(vtoCAE) Then camposClave = False
            If Left(tipoDoc, 2) = "FC" Then If EsVacio(fechaBase) Then camposClave = False
            
            If MENSAJE = MSG_DOA_HOY Then camposClave = False
            
            If .Cells(1, gCtx.rngII.index) <> "" And diferenciaSAP <> 0 Then
                If diferenciaSAP <= gCtx.montoToleranciaSAP And diferenciaSAP >= -gCtx.montoToleranciaSAP Then

                    mensajeNuevo = "AVISO: Se modificarÃ¡ el Impuesto Interno: " & diferenciaSAP
                    If InStr(1, MENSAJE, mensajeNuevo) = 0 Then
                        SetRowStatus fila.Range.Row, "", mensajeNuevo, True
                    End If
                    Resultado = ESTADO_VALIDAR
                    
                End If
            End If
            
            If camposClave = True And (diferenciaSAP <= gCtx.montoToleranciaSAP And diferenciaSAP >= -gCtx.montoToleranciaSAP) Then

                If EsVacio(subtotalIVA21) And EsVacio(subtotalIVA105) Then
                    If gCtx.NombreArchivo = "" Then
                        Resultado = ESTADO_COMPLETAR
                    Else
                        Resultado = ESTADO_REVISAR_DATOS
                    End If
                Else
                    If Right(tipoDoc, 3) = "REM" And estadoPago = ESTADO_REMITO Then
                        Resultado = ESTADO_OK
                    ElseIf Right(tipoDoc, 3) <> "REM" And estadoPago = ESTADO_MIGRAR_SAP Then
                        Resultado = ESTADO_OK
                    ElseIf estadoPago = "" Then
                        Resultado = ESTADO_VALIDAR
                    Else
                        If gCtx.NombreArchivo = "" Then
                            If Referencia = "" Then Exit For
                            Resultado = ESTADO_COMPLETAR
                        Else
                            If gCtx.endoso = False Then
                                Resultado = ESTADO_REVISAR_DATOS
                            Else
                                Resultado = ESTADO_OK
                            End If
                        End If
                    End If
                End If
                
            Else
                If Referencia = "" Then Exit For
                If gCtx.NombreArchivo = "" Or gCtx.NombreArchivo = "Completado por usuario (No hay PDF)" Then
                    If MENSAJE = MSG_DOA_HOY Then
                        Resultado = ESTADO_REVISAR_DATOS
                    Else
                        Resultado = ESTADO_COMPLETAR
                    End If
                Else
                    Resultado = ESTADO_REVISAR_DATOS
                End If
            End If
            
            If InStr(1, UCase(comentarios_User), "ENDOS") > 0 And Resultado <> ESTADO_COMPLETAR Then Resultado = ESTADO_OK
            If InStr(1, UCase(comentarios_User), "COMPENSA") > 0 And Resultado <> ESTADO_COMPLETAR Then Resultado = ESTADO_OK
            

            aviso = "AVISO: Se modificarÃ¡ el Impuesto Interno:"
            If Left(MENSAJE, Len(aviso)) = aviso Then
                If estadoPago = ESTADO_MIGRAR_SAP Then Resultado = ESTADO_OK
            End If
            
            If estadoPago = ESTADO_PENDIENTE_REINGRESO Then Resultado = ESTADO_REVISAR_DATOS
            If estadoPago = ESTADO_ERROR_SCAN Then Resultado = ESTADO_REVISAR_DATOS

            If pagado = "SI" Or CompensaciÃ³n <> "" Then Resultado = ESTADO_CONTABILIZADO
            
            If .Cells(1, gCtx.rngEstado.index) = ESTADO_ELIMINADO Then Resultado = ESTADO_ELIMINADO
            
            SetRowStatus fila.Range.Row, Resultado, ""
            
        End With
    Next fila
    
    gCtx.ControlarCambios = True
    
End Sub
Function EsVacio(valor As Variant) As Boolean
    EsVacio = IsEmpty(valor) Or valor = "" Or valor = 0 Or valor = 0#
End Function





