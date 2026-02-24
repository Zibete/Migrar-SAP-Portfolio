Attribute VB_Name = "modCoreValidation"

' Pure validation helpers (no Excel/SAP/IE access).

Public Function CoreShouldMarkPendienteRevisar(ByVal difCostos As Double, ByVal difConNC As Variant, ByVal montoTolerancia As Double) As Boolean

    If difCostos >= montoTolerancia Then
        If CStr(difConNC) = "" Then
            CoreShouldMarkPendienteRevisar = True
        ElseIf IsNumeric(difConNC) Then
            If CDbl(difConNC) >= montoTolerancia Then CoreShouldMarkPendienteRevisar = True
        End If
    End If

End Function

Public Function CoreIsSiteMismatch(ByVal siteSB As String, ByVal siteFC As String) As Boolean
    CoreIsSiteMismatch = (siteSB <> "" And siteFC <> "" And siteSB <> siteFC)
End Function

Public Function CoreBuildSiteMismatchComment(ByVal tipoDoc As String, ByVal siteFC As String) As String
    If Right(tipoDoc, 3) = "REM" Then
        CoreBuildSiteMismatchComment = "Anular: Ingresan un RTO. de la Sucursal " & siteFC
    Else
        CoreBuildSiteMismatchComment = "Anular: Ingresan una " & Left(tipoDoc, 2) & " de la Sucursal " & siteFC
    End If
End Function

Public Function CoreBuildDoaMessage(ByVal fechaNegSB As Variant, ByVal totalSB As Variant, ByVal montoDOA As Double, ByVal fechaHoy As Date) As String

    Dim fechaNeg As Date

    If fechaNegSB = "" Then Exit Function

    If CDbl(totalSB) < montoDOA Then
        fechaNeg = CDate(fechaNegSB)
        If fechaNeg = fechaHoy Then
            CoreBuildDoaMessage = MSG_DOA_HOY
        ElseIf (Weekday(fechaHoy) = 2 And fechaNeg >= DateAdd("d", -3, fechaHoy)) _
            Or fechaNeg = DateAdd("d", -1, fechaHoy) Then
            CoreBuildDoaMessage = MSG_DOA_PREFIJO & fechaNeg & MSG_DOA_SUFIJO
        End If
    End If

End Function

Public Function CoreIsErrorReferencia( _
    ByVal vendorFilter As String, _
    ByVal remitoRef As String, _
    ByVal tipoDoc As String, _
    ByVal referencia As String, _
    ByVal largoReferencia As Long, _
    ByVal letra As String) As Boolean

    Dim cleanRemitoRef As String
    Dim cleanRef As String
    Dim hasInvalidLength As Boolean

    cleanRemitoRef = Trim$(remitoRef)
    cleanRef = Trim$(referencia)

    If vendorFilter <> "Varios" Then
        If largoReferencia > 0 Then
            hasInvalidLength = (Len(cleanRemitoRef) <> largoReferencia)
        Else
            hasInvalidLength = False
        End If

        If hasInvalidLength Or InStr(1, cleanRemitoRef, letra) = 0 Then
            If vendorFilter <> "<REDACTED_ID_02>" And vendorFilter <> "<REDACTED_ID_07>" Then
                CoreIsErrorReferencia = True
            End If
        End If
    Else
        If Right(tipoDoc, 3) = "REM" Then
            If largoReferencia > 0 Then
                hasInvalidLength = (Len(cleanRef) < largoReferencia)
            Else
                hasInvalidLength = False
            End If

            If hasInvalidLength Or InStr(1, cleanRef, "R") = 0 Then CoreIsErrorReferencia = True
        Else
            If largoReferencia > 0 Then
                hasInvalidLength = (Len(cleanRef) < largoReferencia)
            Else
                hasInvalidLength = False
            End If

            If hasInvalidLength Or (InStr(1, cleanRef, "A") = 0 And InStr(1, cleanRef, "C") = 0) Then CoreIsErrorReferencia = True
        End If
    End If

End Function
