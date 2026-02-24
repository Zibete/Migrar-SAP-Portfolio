Attribute VB_Name = "modCoreText"

' Pure string helpers (no Excel/SAP/IE access).

Public Function CoreAppendUniqueToken(ByVal baseText As String, ByVal token As String, Optional ByVal separator As String = "-") As String

    Dim result As String

    result = baseText

    If token = "" Then
        CoreAppendUniqueToken = result
        Exit Function
    End If

    If result = "" Then
        CoreAppendUniqueToken = token
        Exit Function
    End If

    If InStr(1, result, token) = 0 Then
        CoreAppendUniqueToken = result & separator & token
    Else
        CoreAppendUniqueToken = result
    End If

End Function

Public Function CoreTruncateText(ByVal value As String, Optional ByVal maxLen As Long = 200) As String

    Dim txt As String
    Dim i As Long

    txt = Replace(value, "--", "-")

    If Len(txt) <= maxLen Then
        CoreTruncateText = txt
        Exit Function
    End If

    Do While Len(txt) > maxLen And InStrRev(txt, " ") > 0
        i = InStrRev(txt, " ")
        txt = Left(txt, i - 1) & Mid(txt, i + 1)
    Loop

    If Len(txt) > maxLen Then
        txt = Right(txt, maxLen)
    End If

    CoreTruncateText = txt

End Function

Public Function CoreBuildAutoComment( _
    ByVal observacionesSB As String, _
    ByVal observacionesUser As String, _
    ByVal estadoPago As String, _
    ByVal tipoDoc As String, _
    ByVal referencia As String, _
    ByVal compensacion As String, _
    ByVal difCostos As Double, _
    ByVal fechaUser As String) As String

    Dim comAutomatico As String
    Dim difTexto As String

    If observacionesSB <> "" Then
        comAutomatico = observacionesSB
    End If

    If observacionesSB <> "" Then
        If InStr(1, observacionesSB, fechaUser) = 0 Then
            comAutomatico = comAutomatico & vbLf & fechaUser
        End If
    Else
        comAutomatico = fechaUser
    End If

    If compensacion <> "" Then
        If Right(tipoDoc, 3) = "REM" Then
            comAutomatico = comAutomatico & "-" & referencia
        End If
        comAutomatico = comAutomatico & "-" & compensacion
    End If

    If estadoPago <> "Pendiente de Reingreso" Then
        difTexto = Format(difCostos, "#,##0.00")
        If InStr(1, observacionesSB, difTexto) = 0 Then
            If difCostos > 0 Then
                comAutomatico = comAutomatico & "-Dif. en contra: " & difTexto
            ElseIf difCostos < 0 Then
                comAutomatico = comAutomatico & "-Dif. a favor: " & difTexto
            End If
        End If
    End If

    If observacionesUser <> "" Then
        comAutomatico = comAutomatico & "-" & observacionesUser
    End If

    CoreBuildAutoComment = comAutomatico

End Function

Public Function CoreBuildNombreBase( _
    ByVal site As String, _
    ByVal tipoDoc As String, _
    ByVal referencia As String, _
    ByVal fechaBase As String, _
    ByVal hasRetailWeb As Boolean, _
    ByVal estadoPago As String) As String

    Dim nombre As String

    If site <> "" Then nombre = CoreAppendUniqueToken(nombre, site)
    If tipoDoc <> "" Then nombre = CoreAppendUniqueToken(nombre, tipoDoc)
    If referencia <> "" Then nombre = CoreAppendUniqueToken(nombre, referencia)

    If tipoDoc = "FC-REM" Then
        If fechaBase <> "" Then nombre = CoreAppendUniqueToken(nombre, "Fecha base " & fechaBase)
    End If

    If Not hasRetailWeb Then nombre = CoreAppendUniqueToken(nombre, "Sin RW")
    If estadoPago <> "" Then nombre = CoreAppendUniqueToken(nombre, estadoPago)

    CoreBuildNombreBase = nombre

End Function
