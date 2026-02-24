Attribute VB_Name = "modRetailWebClient"

Public Function RetailWebWaitReady(ByVal timeoutSeconds As Double, ByVal context As String) As Boolean

    RetailWebWaitReady = WaitForIEReady(gCtx.IE_NuevaVentana, timeoutSeconds, context)

End Function

Public Function RetailWebSetSearchValue(ByVal retailWebId As String) As Boolean

    Dim elementoInput As Object

    For Each elementoInput In gCtx.IE_NuevaVentana.Document.getElementsByTagName("input")
        If InStr(1, elementoInput.ID, "stocknumber", vbTextCompare) > 0 Then
            elementoInput.Value = retailWebId
            RetailWebSetSearchValue = True
            Exit Function
        End If
    Next elementoInput

    RetailWebSetSearchValue = False

End Function

Public Function RetailWebClickButtonByText(ByVal className As String, ByVal textValue As String) As Boolean

    Dim elementos As Object
    Dim elemento As Object

    Set elementos = gCtx.IE_NuevaVentana.Document.getElementsByClassName(className)
    For Each elemento In elementos
        If InStr(elemento.innerText, textValue) > 0 Then
            elemento.Click
            RetailWebClickButtonByText = True
            Exit Function
        End If
    Next elemento

    RetailWebClickButtonByText = False

End Function

Public Function RetailWebHasButtonByText(ByVal className As String, ByVal textValue As String) As Boolean

    Dim elementos As Object
    Dim elemento As Object

    Set elementos = gCtx.IE_NuevaVentana.Document.getElementsByClassName(className)
    For Each elemento In elementos
        If InStr(elemento.innerText, textValue) > 0 Then
            RetailWebHasButtonByText = True
            Exit Function
        End If
    Next elemento

    RetailWebHasButtonByText = False

End Function

Public Function RetailWebTryGetElementsByClass(ByVal className As String) As Object

    On Error GoTo CleanFail
    Set RetailWebTryGetElementsByClass = gCtx.IE_NuevaVentana.Document.getElementsByClassName(className)
    Exit Function

CleanFail:
    Set RetailWebTryGetElementsByClass = Nothing
    On Error GoTo 0

End Function

Public Function RetailWebTryGetElementsByClassFromDoc(ByVal doc As Object, ByVal className As String) As Object

    On Error GoTo CleanFail
    Set RetailWebTryGetElementsByClassFromDoc = doc.getElementsByClassName(className)
    Exit Function

CleanFail:
    Set RetailWebTryGetElementsByClassFromDoc = Nothing
    On Error GoTo 0

End Function

Public Function RetailWebGetFirstRow(ByRef fila As Object, ByVal timeoutSeconds As Double) As Boolean

    Dim startTime As Double
    Dim tabla As Object

    startTime = Timer

    Do
        Set tabla = gCtx.IE_NuevaVentana.Document.getElementById(SB_ID_SCROLL_BODY)
        If Not tabla Is Nothing Then
            Set fila = tabla.getElementsByTagName("tr")(0)
            If Not fila Is Nothing Then
                RetailWebGetFirstRow = True
                Exit Function
            End If
        End If

        DoEvents
        If HasTimedOut(startTime, timeoutSeconds) Then
            RetailWebGetFirstRow = ReportTimeout(SB_TEXT_CONTROL_RECEPCIONES)
            Exit Function
        End If
    Loop

End Function

Public Function RetailWebGetPayComment() As String

    Dim elementos As Object
    Dim observacionesSB As Object

    Set elementos = gCtx.IE_NuevaVentana.Document.getElementsByClassName(SB_CLASS_INPUT_SM)
    For Each observacionesSB In elementos
        If InStr(1, observacionesSB.ID, "-payComment", vbTextCompare) > 0 Then
            RetailWebGetPayComment = CStr(observacionesSB.Value)
            Exit Function
        End If
    Next observacionesSB

    RetailWebGetPayComment = ""

End Function

Public Function RetailWebBuscarFila(ByVal retailWebId As String, ByRef fila As Object, ByVal timeoutSeconds As Double) As Boolean

    RetailWebBuscarFila = False

    If Not RetailWebSetSearchValue(retailWebId) Then Exit Function
    Call RetailWebClickButtonByText(SB_CLASS_BUTTON_DEFAULT_SM, SB_TEXT_BUSCAR)

    Application.Wait (Now + TimeValue("00:00:02"))
    verificarWaitingPanel (SB_TEXT_CONTROL_RECEPCIONES)
    If gCtx.timeout Then Exit Function

    If Not RetailWebGetFirstRow(fila, timeoutSeconds) Then Exit Function
    fila.Click

    Application.Wait (Now + TimeValue("00:00:01"))
    verificarWaitingPanel (SB_TEXT_CONTROL_RECEPCIONES)
    If gCtx.timeout Then Exit Function

    RetailWebBuscarFila = True

End Function

Public Function RetailWebImprimirFactura(ByVal timeoutSeconds As Double) As Boolean

    RetailWebImprimirFactura = False
    Call RetailWebClickButtonByText(SB_CLASS_BUTTON_DEFAULT_SM_PULL, SB_TEXT_IMPRIMIR_FACTURA)
    If Not RetailWebWaitReady(timeoutSeconds, SB_TEXT_CONTROL_RECEPCIONES) Then Exit Function

    Application.Wait (Now + TimeValue("00:00:03"))
    verificarWaitingPanel (SB_TEXT_CONTROL_RECEPCIONES)
    If gCtx.timeout Then Exit Function

    RetailWebImprimirFactura = True

End Function

Public Function RetailWebCambiarEstado(ByVal nuevoEstado As String, ByVal nuevoComentario As String, ByVal timeoutSeconds As Double) As Boolean

    Dim elementos As Object
    Dim elemento As Object
    Dim opciones As Object
    Dim opcion As Object
    Dim observacionesSB As Object
    Dim seguirEsperando As Boolean
    Dim startTime As Double

    RetailWebCambiarEstado = False

    seguirEsperando = True
    startTime = Timer
    Do While seguirEsperando
        If RetailWebClickButtonByText(SB_CLASS_BUTTON_DEFAULT_SM_PULL, SB_TEXT_CAMBIAR_ESTADO_PAGO) Then
            seguirEsperando = False
            Exit Do
        End If
        If HasTimedOut(startTime, timeoutSeconds) Then
            ReportTimeout SB_TEXT_CAMBIAR_ESTADO_PAGO
            Exit Function
        End If
        DoEvents
    Loop

    seguirEsperando = True
    startTime = Timer
    Do While seguirEsperando
        If RetailWebHasButtonByText(SB_CLASS_BUTTON_SUCCESS_SM_PULL, SB_TEXT_ACEPTAR) Then
            seguirEsperando = False
            Exit Do
        End If
        If HasTimedOut(startTime, timeoutSeconds) Then
            ReportTimeout SB_TEXT_ACEPTAR
            Exit Function
        End If
        DoEvents
    Loop

    verificarWaitingPanel (SB_TEXT_CONTROL_RECEPCIONES)
    If gCtx.timeout Then Exit Function

    Set elementos = gCtx.IE_NuevaVentana.Document.getElementsByClassName(SB_CLASS_INPUT_SM)

    For Each observacionesSB In elementos
        If InStr(1, observacionesSB.ID, "-payComment", vbTextCompare) > 0 Then
            observacionesSB.Value = nuevoComentario
            Exit For
        End If
    Next observacionesSB

    For Each elemento In elementos
        If InStr(1, elemento.ID, "-stockPayState", vbTextCompare) > 0 Then
            Set opciones = elemento.getElementsByTagName("option")
            For Each opcion In opciones
                If Trim(opcion.innerText) = nuevoEstado Then
                    opcion.Selected = True
                    Exit For
                End If
            Next opcion
        End If
    Next elemento

    Set elementos = gCtx.IE_NuevaVentana.Document.getElementsByClassName(SB_CLASS_BUTTON_SUCCESS_SM_PULL)
    For Each elemento In elementos
        If InStr(elemento.innerText, SB_TEXT_ACEPTAR) > 0 Then
            elemento.Click
            Exit For
        End If
    Next elemento

    seguirEsperando = True
    startTime = Timer
    Do While seguirEsperando
        seguirEsperando = RetailWebHasButtonByText(SB_CLASS_BUTTON_SUCCESS_SM_PULL, SB_TEXT_ACEPTAR)
        If seguirEsperando Then
            If HasTimedOut(startTime, timeoutSeconds) Then
                ReportTimeout SB_TEXT_ACEPTAR
                Exit Function
            End If
            DoEvents
        End If
    Loop

    RetailWebCambiarEstado = True

End Function

Public Function RetailWebPagar(ByVal timeoutSeconds As Double) As Boolean

    Dim tabla As Object
    Dim fila As Object
    Dim checkboxElement As Object
    Dim botones As Object
    Dim boton As Object
    Dim botones2 As Object
    Dim boton2 As Object

    RetailWebPagar = False

    Set tabla = gCtx.IE_NuevaVentana.Document.getElementById(SB_ID_SCROLL_BODY)
    If tabla Is Nothing Then Exit Function

    Set fila = tabla.getElementsByTagName("tr")(0)
    If fila Is Nothing Then Exit Function

    Set checkboxElement = fila.getElementsByTagName("td")(17).getElementsByTagName("input")(0)
    If checkboxElement Is Nothing Then Exit Function

    checkboxElement.Click
    If Not RetailWebWaitReady(timeoutSeconds, SB_TEXT_CONTROL_RECEPCIONES) Then Exit Function

    Set botones = gCtx.IE_NuevaVentana.Document.getElementsByClassName(SB_CLASS_BUTTON_DEFAULT_SM_PULL)
    For Each boton In botones
        If InStr(boton.innerText, SB_TEXT_PAGAR) > 0 Then
            boton.Click

            Application.Wait (Now + TimeValue("00:00:01"))
            verificarWaitingPanel (SB_TEXT_CONTROL_RECEPCIONES)
            If gCtx.timeout Then Exit Function

            Set botones2 = gCtx.IE_NuevaVentana.Document.getElementsByClassName(SB_CLASS_BUTTON_DEFAULT)
            For Each boton2 In botones2
                If Trim(boton2.innerText) = "Si" Then
                    boton2.Click

                    Application.Wait (Now + TimeValue("00:00:01"))
                    verificarWaitingPanel (SB_TEXT_CONTROL_RECEPCIONES)
                    If gCtx.timeout Then Exit Function

                    RetailWebPagar = True
                    Exit Function
                End If
            Next boton2
        End If
    Next boton

End Function

