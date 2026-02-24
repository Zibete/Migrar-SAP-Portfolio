Attribute VB_Name = "modText"

Public Function truncarTXT(nuevoTxt)
    truncarTXT = CoreTruncateText(CStr(nuevoTxt), 200)
    
End Function

Public Function comentarioAutomatico(i, observaciones_SB, observaciones_User)
    
    If i = 16 Then
        i = 16
    End If
    
    estadoPago = Hoja2.Cells(i, gCtx.rngEstadoDelPago.Range.Column)
    tipoDoc = Hoja2.Cells(i, gCtx.rngTipoDoc.Range.Column)
    referencia = Hoja2.Cells(i, gCtx.rngReferencia.Range.Column)
    compensacion = Hoja2.Cells(i, gCtx.rngCompensacion.Range.Column)
    difCostos = Round(CDbl(Hoja2.Cells(i, gCtx.rngDifCostos.Range.Column)), 2)
    fechaUser = Format(Date, "dd.mm.yyyy") & "-" & Environ("USERNAME")

    comentarioAutomatico = CoreBuildAutoComment( _
        CStr(observaciones_SB), _
        CStr(observaciones_User), _
        CStr(estadoPago), _
        CStr(tipoDoc), _
        CStr(referencia), _
        CStr(compensacion), _
        difCostos, _
        fechaUser _
    )


'    'Fecha + USER
'    If InStr(1, observaciones_SB, Format(Date, "dd.mm.yyyy") & "-" & Environ("USERNAME")) = 0 Then
'        comAutomatico = Format(Date, "dd.mm.yyyy") & "-" & Environ("USERNAME")
'    End If
'    'Referencia SI o NO
'    If Right(Hoja2.Cells(i, rngTipoDoc.Range.Column), 3) = "REM" Then
'        If InStr(1, observaciones_SB, Hoja2.Cells(i, rngReferencia.Range.Column)) = 0 Then
'
'            'If Len(Hoja2.Cells(i, rngReferencia.Range.Column)) = largoReferencia And _
'            'InStr(1, Hoja2.Cells(i, rngReferencia.Range.Column), letra) = 1 Then
'                comAutomatico = comAutomatico & "-" & Hoja2.Cells(i, rngReferencia.Range.Column)
'            'End If
'
'        End If
'    End If
'    'SAP
'    If Hoja2.Cells(i, rngCompensacion.Range.Column) <> "" Then
'        If InStr(1, observaciones_SB, Hoja2.Cells(i, rngCompensacion.Range.Column)) = 0 Then
'            comAutomatico = comAutomatico & "-" & Hoja2.Cells(i, rngCompensacion.Range.Column)
'        End If
'    End If
'    'Diferencia de costos
'    difCostos = Round(CDbl(Hoja2.Cells(i, rngDifCostos.Range.Column)), 2)
'    If InStr(1, observaciones_SB, Format(difCostos, "#,##0.00")) = 0 Then
'        'If Right(Hoja2.Cells(i, rngTipoDoc.Range.Column), 3) <> "REM" Then
'            If difCostos >= montoToleranciaSB Then
'                comAutomatico = comAutomatico & "-Dif. en contra: " & Format(difCostos, "#,##0.00")
'            ElseIf difCostos <= -montoToleranciaSB Then
'                comAutomatico = comAutomatico & "-Dif. a favor: " & Format(difCostos, "#,##0.00")
'            End If
'        'End If
'    End If
'    'Observaciones pre-existentes en RW
'    If comAutomatico <> "" And observaciones_SB <> "" Then
'        If Left(comAutomatico, 10) = Format(Date, "dd.mm.yyyy") Then
'            comAutomatico = observaciones_SB & vbLf & comAutomatico
'        Else
'            comAutomatico = observaciones_SB & "-" & comAutomatico
'        End If
'    End If
'
'    If comAutomatico = "" Then comAutomatico = observaciones_SB
'
'    comentarioAutomatico = comAutomatico
'
End Function
Public Function sumarNuevoNombre(txtNuevo, txtComparar)
    sumarNuevoNombre = CoreAppendUniqueToken(CStr(txtComparar), CStr(txtNuevo), "-")
End Function

