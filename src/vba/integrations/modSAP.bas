Attribute VB_Name = "modSAP"

Function GetPass() As String
    Dim s As Object
    Set s = CreateObject("WScript.Shell")
    GetPass = Trim(s.exec("""" & GetPythonwExePath() & """ """ & ResolveScriptPath("Credenciales.py") & """").StdOut.ReadAll)
End Function
Sub pruebaCredenciales()
    MsgBox GetPass()
End Sub
Public Sub ABRIRSAP()

'    Dim objWMI As Object, procesos As Object, proc As Object
'    Dim sapAbierto As Boolean
'    sapAbierto = False
'
'    Set objWMI = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("Select * from Win32_Process Where Name = 'saplogon.exe'")
'    For Each proc In objWMI
'        sapAbierto = True
'        Exit For
'    Next
    
    On Error GoTo ABRIRSAP
        If Not IsObject(App) Then
            Set SapGuiAuto = GetObject("SAPGUI")
            Set App = SapGuiAuto.GetScriptingEngine
        End If
        If Not IsObject(Connection) Then
            Set Connection = App.Children(0)
        End If
        If Not IsObject(session) Then
            Set session = Connection.Children(0)
        End If
        If IsObject(WScript) Then
            WScript.ConnectObject session, "on"
            WScript.ConnectObject App, "on"
        End If
    On Error GoTo 0
    
    GoTo CleanUp
    
ABRIRSAP:

'    If sapAbierto Then Exit Sub ' SAP ya está abierto

    Dim shell As Object, exec As Object
    Dim rutaSAP As String, linea As String
    Dim pythonPath As String, scriptPath As String
    Dim cmd As String

    pythonPath = GetPythonwExePath()
    scriptPath = ResolveScriptPath("AbrirSAP.py")
    
    cmd = """" & pythonPath & """ """ & scriptPath & """"

    Set shell = CreateObject("WScript.Shell")
    Set exec = shell.exec(cmd)
    
    startTime = Timer
    Do While Not exec.StdOut.AtEndOfStream
        linea = exec.StdOut.ReadLine
        If InStr(linea, ".sap") > 0 Then
            rutaSAP = linea
            Exit Do
        End If
        If HasTimedOut(startTime, WAIT_LONG_SECONDS) Then
            ReportTimeout "Abrir SAP"
            Exit Do
        End If
    Loop

    If rutaSAP <> "" Then
        Call CreateObject("Shell.Application").ShellExecute(rutaSAP)
    Else
        MsgBox "No se pudo encontrar la ruta del archivo .sap", vbCritical
    End If

CleanUp:
    Set exec = Nothing
    Set shell = Nothing

End Sub
Sub buscarSAP()

    ABRIRSAP
    
    tareas2 = gCtx.SELECTION_USER
    tareas1 = tareas1 + tareas2
    capt1 = "Verificar si existe en SAP..."
    
    avance2 = 1
    With ProgressBar
        .Show vbModeless
        .Lbl1.Caption = capt1 & " (" & Format(avance1 + avance2 / tareas1, "0%") & ")"
        .Lbl2.Caption = "Preparando todo..." & " (" & Format(avance2 / tareas2, "0%") & ")"
        .pb1.Max = tareas1
        .pb2.Max = tareas2
        .pb1.Value = avance1 + avance2
        .pb2.Value = avance2
    End With
    
    On Error GoTo ErrorSap
    If Not IsObject(App) Then
        Set SapGuiAuto = GetObject("SAPGUI")
        Set App = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
        Set Connection = App.Children(0)
    End If
    If Not IsObject(session) Then
        Set session = Connection.Children(0)
    End If
    If IsObject(WScript) Then
        WScript.ConnectObject session, "on"
        WScript.ConnectObject App, "on"
    End If
    On Error GoTo 0
    
    session.findById("wnd[0]").resizeWorkingPane 97, 22, False
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/NFBL1N"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[16]").press
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").selectNode "         60"
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").topNode = "         55"
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").doubleClickNode "         60"
    
    

    For indice = 1 To tareas2
    
        i = Selection.Cells(indice, 1).Row
        
        Vendor = Hoja2.Cells(i, gCtx.rngVendorProveedor_SB.Range.Column)
    
        Set rngProveedor = gCtx.rngVendor_Prov.DataBodyRange.Find(What:=Vendor, LookAt:=xlWhole)
        
        esPyme = True
        If Hoja3.Cells(rngProveedor.Row, gCtx.rngEsPyme_Prov.Range.Column) = "NO" Then esPyme = False

        Referencia = Hoja2.Cells(i, gCtx.rngReferencia.Range.Column)
        Total = Hoja2.Cells(i, gCtx.rngTotalBrutoFactura.Range.Column)
        
        If esPyme And CDbl(Total) >= gCtx.montoFCE Then
            If Len(Referencia) = 14 Then Referencia = Mid(Referencia, 2)
        End If

        avance2 = indice
        With ProgressBar
            .Lbl1.Caption = capt1 & " (" & Format(avance1 + avance2 / tareas1, "0%") & ")"
            .Lbl2.Caption = "Ejecutando FBL1N en SAP: " & indice & " de " & tareas2 & " (" & Format(avance2 / tareas2, "0%") & ")"
            .pb1.Value = avance1 + avance2
            .pb2.Value = avance2
        End With

        If Referencia = "" Then Exit For
      
        session.findById("wnd[0]/usr/radX_AISEL").Select
        session.findById("wnd[0]/usr/ctxtKD_LIFNR-LOW").Text = Vendor
        session.findById("wnd[0]/usr/ctxtKD_BUKRS-LOW").Text = "<REDACTED>"
        session.findById("wnd[0]/tbar[1]/btn[16]").press
        
        If Not SapTrySetText(session, "wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/txt%%DYN015-LOW", Referencia) Then
            session.findById("wnd[0]/tbar[1]/btn[16]").press
            Call SapTrySetText(session, "wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/txt%%DYN015-LOW", Referencia)
        End If
        
        session.findById("wnd[0]").sendVKey 8
        mensajeSAP = session.findById("wnd[0]/sbar").Text

        If InStr(1, mensajeSAP, "No se ha seleccionado ninguna partida") > 0 Then
            SetRowStatus i, "", "No se encontró"
        Else
            SAP = session.findById("wnd[0]/usr/lbl[8,10]").Text
            session.findById("wnd[0]").sendVKey 12
            SetRowStatus i, "", SAP & " (" & mensajeSAP & ")"
        End If

    Next indice
    
    gCtx.rngMensajesSap.Range.Columns.AutoFit

    Unload ProgressBar
    Exit Sub

ErrorSap:

    MENSAJE = MsgBox("ERROR: SAP no está abierto", vbExclamation, "Error")
    Unload ProgressBar

End Sub

Sub buscarFCE()

    ABRIRSAP
    
    tareas2 = gCtx.SELECTION_USER
    tareas1 = tareas1 + tareas2
    capt1 = "Verificar FCE en SAP..."
    
    avance2 = 1
    With ProgressBar
        .Show vbModeless
        .Lbl1.Caption = capt1 & " (" & Format(avance1 + avance2 / tareas1, "0%") & ")"
        .Lbl2.Caption = "Preparando todo..." & " (" & Format(avance2 / tareas2, "0%") & ")"
        .pb1.Max = tareas1
        .pb2.Max = tareas2
        .pb1.Value = avance1 + avance2
        .pb2.Value = avance2
    End With
    
    On Error GoTo ErrorSap
    If Not IsObject(App) Then
        Set SapGuiAuto = GetObject("SAPGUI")
        Set App = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
        Set Connection = App.Children(0)
    End If
    If Not IsObject(session) Then
        Set session = Connection.Children(0)
    End If
    If IsObject(WScript) Then
        WScript.ConnectObject session, "on"
        WScript.ConnectObject App, "on"
    End If
    On Error GoTo 0
    
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/NZARFI_FCE_MONITOR"
    session.findById("wnd[0]").sendVKey 0
    
    gCtx.ControlarCambios = False
    
    For indice = 1 To tareas2
    
        avance2 = indice
        
        With ProgressBar
            .Lbl1.Caption = capt1 & " (" & Format(avance1 + avance2 / tareas1, "0%") & ")"
            .Lbl2.Caption = "Buscando FCE en SAP: " & indice & " de " & tareas2 & " (" & Format(avance2 / tareas2, "0%") & ")"
            .pb1.Value = avance1 + avance2
            .pb2.Value = avance2
        End With
        
        i = Selection.Cells(indice, 1).Row
        
        If Not Hoja2.Rows(i).EntireRow.Hidden Then
        
        Vendor = Hoja2.Cells(i, gCtx.rngVendorProveedor_SB.Range.Column)
        
        Set rngProveedor = gCtx.rngVendor_Prov.DataBodyRange.Find(What:=Vendor, LookAt:=xlWhole)
        
        If Hoja3.Cells(rngProveedor.Row, gCtx.rngEsPyme_Prov.Range.Column) = FLAG_SI Then
        
            CUIT = Hoja3.Cells(rngProveedor.Row, gCtx.rngCUIT_Prov.Range.Column)
            CONDPAGO = Hoja3.Cells(rngProveedor.Row, gCtx.rngCondPago_Prov.Range.Column)
     
            Set rngCondPago = gCtx.rngCod_CondPago.DataBodyRange.Find(What:=CONDPAGO, LookAt:=xlWhole)
            DESC_CONDPAGO = Hoja3.Cells(rngCondPago.Row, gCtx.rngDescripcion_CondPago.Range.Column)

            Fecha = Hoja2.Cells(i, gCtx.rngFechaDoc_SB.Range.Column)
            Referencia = Hoja2.Cells(i, gCtx.rngRemitoRef.Range.Column)
            SOC = "<REDACTED>"
            If Referencia = "" Then Exit For

            For intento = 1 To 2
    
            session.findById("wnd[0]/usr/ctxtSO_BUK2-LOW").Text = SOC
            session.findById("wnd[0]/usr/ctxtSO_CUIT-LOW").Text = CUIT
            session.findById("wnd[0]/usr/ctxtSO_EMI-LOW").Text = Replace(Fecha, "/", ".")
            session.findById("wnd[0]/usr/ctxtSO_EMI-HIGH").Text = Replace(Fecha, "/", ".")
            session.findById("wnd[0]/usr/txtSO_XBLN2-LOW").Text = Referencia
            session.findById("wnd[0]/usr/txtSO_XBLN2-HIGH").Text = Referencia
            session.findById("wnd[0]/usr/ctxtSO_LIFNR-LOW").Text = Vendor
            session.findById("wnd[0]/usr/ctxtSO_LIFNR-HIGH").Text = Vendor
            
            session.findById("wnd[0]/usr/radRB_TODOS").Select
            
            session.findById("wnd[0]").sendVKey 8
    
            Set BTN = SapTryFindById(session, "wnd[1]/tbar[0]/btn[0]")
            
            If Not BTN Is Nothing Then
                BTN.press
                If Len(Referencia) = 13 And intento = 1 Then
                    Referencia = "0" & Referencia
                Else
                    MENSAJE = MsgBox(indice & " de " & tareas2 & vbLf & vbLf & "No encontrado", vbCritical, "MONITOR FCE: " & Referencia)
                    Exit For
                End If
            Else
    
                CODIGO_CTACTE = session.findById("wnd[0]/usr/shell/shellcont/shell").GetCellValue(0, "CODIGO_CTACTE")
                ESTADOO = session.findById("wnd[0]/usr/shell/shellcont/shell").GetCellValue(0, "ESTADO")
                OPCION_TRANSFERENCIA = session.findById("wnd[0]/usr/shell/shellcont/shell").GetCellValue(0, "OPCION_TRANSFERENCIA")
                FECHA_EMISION = session.findById("wnd[0]/usr/shell/shellcont/shell").GetCellValue(0, "FECHA_EMISION")
                FECHA_VTO = session.findById("wnd[0]/usr/shell/shellcont/shell").GetCellValue(0, "FECHA_VTO")
                DOC_SAP = session.findById("wnd[0]/usr/shell/shellcont/shell").GetCellValue(0, "BELNR")
                
                diferenciaDIAS = DateDiff("d", Replace(FECHA_EMISION, ".", "/"), Replace(FECHA_VTO, ".", "/"))
     
                txtMENSAJE = indice & " de " & tareas2 & vbLf & vbLf & _
                             "Estado ------------------------> " & ESTADOO & vbLf & _
                             "Código Cta.Cte -------------> " & CODIGO_CTACTE & vbLf & _
                             "Nro. Doc. SAP ---------------> " & DOC_SAP & vbLf & vbLf & _
                             "Opción Transferencia ------> " & OPCION_TRANSFERENCIA & vbLf & vbLf & _
                             "Fecha de Emisión -----------> " & FECHA_EMISION & vbLf & _
                             "Fecha de Vencimiento -----> " & FECHA_VTO & vbLf & vbLf & _
                             "Diferencia entre días -------> " & diferenciaDIAS & vbLf & _
                             "Cond. Pago Proveedor -----> " & CONDPAGO & vbLf & _
                             "Descripción Cond. Pago ---> " & DESC_CONDPAGO
                             
                MENSAJE = MsgBox(txtMENSAJE, vbInformation, "MONITOR FCE: " & Referencia)
                
                If ESTADOO = "Rechazado" Then Hoja2.Cells(i, gCtx.rngEstadoDelPago.Range.Column) = ESTADO_VALIDACION_AFIP_RECHAZADA
                If OPCION_TRANSFERENCIA = "SCA" Then
                    Hoja2.Cells(i, gCtx.rngEstadoDelPago.Range.Column) = ESTADO_VALIDACION_AFIP_RECHAZADA
                    Hoja2.Cells(i, gCtx.rngComentarios_User.Range.Column) = sumarNuevoNombre(OPCION_TRANSFERENCIA, Hoja2.Cells(i, gCtx.rngComentarios_User.Range.Column))
                End If
                
                DIF_1 = diferenciaDIAS - CInt((Left(DESC_CONDPAGO, 2)))
                
                If DOC_SAP = "" And DIF_1 > 3 Or DIF_1 < -3 Then
                    Hoja2.Cells(i, gCtx.rngEstadoDelPago.Range.Column) = ESTADO_VALIDACION_AFIP_RECHAZADA
                    txtCom = "Vto. en ARCA (" & FECHA_VTO & " - " & diferenciaDIAS & " días) difiere en " & DIF_1 & " días de " & CONDPAGO & " (" & DESC_CONDPAGO & ")"
                    Hoja2.Cells(i, gCtx.rngComentarios_User.Range.Column) = sumarNuevoNombre(txtCom, Hoja2.Cells(i, gCtx.rngComentarios_User.Range.Column))
                End If
                
                If DOC_SAP <> "" Then Hoja2.Cells(i, gCtx.rngComentarios_User.Range.Column) = sumarNuevoNombre(DOC_SAP, Hoja2.Cells(i, gCtx.rngComentarios_User.Range.Column))
                
                session.findById("wnd[0]").sendVKey 12
                Exit For
                
            End If

            Next intento

        Else
            MENSAJE = MsgBox(indice & " de " & tareas2 & vbLf & vbLf & "No es FCE miPyme", vbCritical, "MONITOR FCE: " & Referencia)
        End If
        
End If
Next indice
    
    gCtx.ControlarCambios = True
    
    gCtx.rngMensajesSap.Range.Columns.AutoFit
    gCtx.rngComentarios_User.Range.Columns.AutoFit

    Unload ProgressBar
    Exit Sub

ErrorSap:

    MENSAJE = MsgBox("ERROR: SAP no está abierto", vbExclamation, "Error")
    Unload ProgressBar

End Sub
Sub AbrirRecepciones()

    Dim ventana As Object
    Dim startRowTime As Double
    Dim fila As Object
    Dim elementTd5 As Object
    Dim retailWebId As String

    gCtx.timeout = False

    For Each ventana In CreateObject("Shell.Application").Windows
        If Left(ventana.LocationURL, Len(gCtx.dominio)) = gCtx.dominio Then
            Set gCtx.IE_NuevaVentana = ventana
            Exit For
        End If
    Next ventana
    
    On Error GoTo errorSB

    With gCtx.IE_NuevaVentana
        .Visible = True
        .TheaterMode = False
    End With
    
    If Not RetailWebWaitReady(WAIT_LONG_SECONDS, SB_TEXT_CONTROL_RECEPCIONES) Then GoTo errorSB
    
    If gCtx.textoBtn = ACTION_IMPRIMIR_FACTURA And Hoja2.Cells(Selection.Cells(1, 1).Row, gCtx.rngEstado.Range.Column) = ESTADO_COMPLETAR Then
                
        VisualizarScan.Show
        
    Else
    
        For indice = 1 To Selection.Rows.Count

            i = Selection.Cells(indice, 1).Row

            If ShouldProcessRetailWebRow(i) Then

                retailWebId = CStr(Hoja2.Cells(i, gCtx.rngRetailWeb_SB.Range.Column).Value)

                If RetailWebBuscarFila(retailWebId, fila, WAIT_LONG_SECONDS) Then

                    startRowTime = Timer
                    Do
                        Set elementTd5 = fila.getElementsByTagName("td")(5)
                        If Not elementTd5 Is Nothing Then Exit Do
                        DoEvents
                        If HasTimedOut(startRowTime, WAIT_SHORT_SECONDS) Then
                            ReportTimeout "Validar fila RetailWeb"
                            Exit Do
                        End If
                    Loop

                    If gCtx.timeout Then GoTo errorSB

                    If Not elementTd5 Is Nothing Then
                        If elementTd5.innerText = retailWebId Then

                            Application.ScreenUpdating = True

                            If gCtx.textoBtn = ACTION_IMPRIMIR_FACTURA And Hoja2.Cells(i, gCtx.rngEstado.Range.Column) <> ESTADO_COMPLETAR And Hoja2.Cells(i, gCtx.rngTieneScan_SB.Range.Column).Value = FLAG_SI Then
                                Call ImprimirFactura2(i)
                            End If

                            If gCtx.textoBtn = ACTION_CAMBIAR_ESTADO Then Call CambiarEstado2(i)
                            If gCtx.textoBtn = ACTION_PAGAR_FACTURA Then Call PagarFactura2(i)
                            If gCtx.textoBtn = ACTION_CAMBIAR_PAGAR Then Call CambiarEstado2(i)

                            Application.ScreenUpdating = False
                        End If
                    End If

                ElseIf gCtx.timeout Then
                    GoTo errorSB
                End If

            End If

            If Selection.Rows.Count > 1 And Selection.Rows.Count <> indice And gCtx.textoBtn = ACTION_IMPRIMIR_FACTURA Then
                If MsgBox("¿Desea continuar?", vbYesNo, "Pregunta") = vbNo Then: Exit Sub
            End If

        Next indice
    
    End If
    
    Exit Sub
    
errorSB:

    If gCtx.timeout Then
        MENSAJE = MsgBox(MSG_TIMEOUT_SB, vbCritical, MSG_TIMEOUT_TITLE)
    Else
        MENSAJE = MsgBox("Error desconocido", vbCritical)
    End If

        
End Sub

Sub PagarFactura2(i)

    If RetailWebPagar(WAIT_LONG_SECONDS) Then
        Hoja2.Cells(i, gCtx.rngPagado.Range.Column).Value = FLAG_SI
    End If

End Sub
Sub CambiarEstado2(i)

    Dim comentarioActual As String
    Dim nuevoComentario As String

    comentarioActual = RetailWebGetPayComment()
    nuevoComentario = comentarioAutomatico(i, comentarioActual, Hoja2.Cells(i, gCtx.rngComentarios_User.Range.Column))
    nuevoComentario = truncarTXT(nuevoComentario)

    If Not RetailWebCambiarEstado(Hoja2.Cells(i, gCtx.rngEstadoDelPago.Range.Column), nuevoComentario, WAIT_LONG_SECONDS) Then Exit Sub

    Hoja2.Cells(i, gCtx.rngEstadoCambiado.Range.Column).Value = FLAG_SI
    Hoja2.Cells(i, gCtx.rngEstadoDelPago_SB.Range.Column).Value = Hoja2.Cells(i, gCtx.rngEstadoDelPago.Range.Column)
    Hoja2.Cells(i, gCtx.rngComentarios_SB.Range.Column).Value = nuevoComentario
    
    If GetOrigenDatos() = ORIGEN_DATOS_SB Then
        Set sheetRetailWeb = ThisWorkbook.Sheets("sheetRetailWeb")
        Set tblCuboSB = sheetRetailWeb.ListObjects("tblCuboSB")
        Set rngRetailWebCubo = tblCuboSB.ListColumns("RetailWeb #")
        Set CUBO_SB = rngRetailWebCubo.DataBodyRange.Find(What:=Hoja2.Cells(i, gCtx.rngRetailWeb_SB.Range.Column), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
        If Not CUBO_SB Is Nothing Then
            Set rngEstadoCubo = tblCuboSB.ListColumns("Estado del Pago")
            sheetRetailWeb.Cells(CUBO_SB.Row, rngEstadoCubo.Range.Column) = Hoja2.Cells(i, gCtx.rngEstadoDelPago.Range.Column)
            Set rngObservacionesCubo = tblCuboSB.ListColumns("Observaciones del Pago")
            sheetRetailWeb.Cells(CUBO_SB.Row, rngObservacionesCubo.Range.Column) = nuevoComentario
        End If
    End If

    Set DB_SB = gCtx.rngRetailWeb_DB.DataBodyRange.Find(What:=Hoja2.Cells(i, gCtx.rngRetailWeb_SB.Range.Column), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not DB_SB Is Nothing Then
        gCtx.sheetDataBase.Cells(DB_SB.Row, gCtx.rngEstado_DB.Range.Column) = Hoja2.Cells(i, gCtx.rngEstadoDelPago.Range.Column)
        'sheetDataBase.Cells(DB_SB.Row, rngComentarios_DB.Range.Column) = nuevoComentario
    End If

    If gCtx.textoBtn = ACTION_CAMBIAR_PAGAR Then
        Call PagarFactura2(i)
    End If

    Set fila = gCtx.tblDatos.ListRows(i - gCtx.tblDatos.HeaderRowRange.Row)
    Call VerificarDatos(fila)
    
        
End Sub

Sub ImprimirFactura2(i)

    If Not RetailWebImprimirFactura(WAIT_LONG_SECONDS) Then Exit Sub

'Pregunta Fecha de Sello:
    If Hoja2.Cells(i, gCtx.rngTipoDoc.Range.Column) = "FC-REC" Or Hoja2.Cells(i, gCtx.rngTipoDoc.Range.Column) = "FCE-REC" Then
    
        Set rngProveedor = gCtx.rngVendor_Prov.DataBodyRange.Find(What:=Hoja2.Cells(i, gCtx.rngVendorProveedor_SB.Range.Column), LookAt:=xlWhole)

        If Left(Hoja3.Cells(rngProveedor.Row, gCtx.rngCondPago_Prov.Range.Column), 1) = "Z" Then
        
            Fecha = CDate(Replace(Hoja2.Cells(i, gCtx.rngFechaBase.Range.Column).Value, ".", "/"))
            respuesta = MsgBox("Fecha del sello es " & Fecha & ". ¿Hay error?", vbYesNo, "Pregunta")
            If respuesta = vbYes Then
                For resta = 1 To 60
                    respuesta = MsgBox("Fecha del sello es " & Fecha - resta & ". ¿Hay error?", vbYesNo, "Pregunta")
                    If respuesta = vbNo Then
                        Fecha = Fecha - resta
                        Hoja2.Cells(i, gCtx.rngFechaBase.Range.Column).Value = Format(DateValue(Fecha), "dd.mm.yyyy")
                        Hoja2.Cells(i, gCtx.rngFechaBaseCambiada.Range.Column).Value = FLAG_SI
                        Exit For
                    End If
                Next resta
            End If
            
        End If
    Else
        pendienteReingreso = MsgBox("¿Tiene error en fecha?", vbYesNo, "Pregunta")
    End If

    ErrorScan = MsgBox("¿Tiene error de Scan?", vbYesNo, "Pregunta")
    Comentarios = Hoja2.Cells(i, gCtx.rngComentarios_User.Range.Column)
    tipoDoc = Hoja2.Cells(i, gCtx.rngTipoDoc.Range.Column).Value
    
    If ErrorScan = vbYes Then
        If MsgBox("¿Error en sello?", vbYesNo, "Pregunta") = vbYes Then
            If MsgBox("¿Falta sello?", vbYesNo, "Pregunta") = vbYes Then
                com = "FALTA SELLO DE RECIBIDO"
                If Comentarios = "" Then Comentarios = com Else: Comentarios = Comentarios & "-" & com
            Else
                com = "SELLO DE RECIBIDO: Verificar datos (Fecha, legibilidad...)"
                If Comentarios = "" Then Comentarios = com Else: Comentarios = Comentarios & "-" & com
            End If
        End If
        If MsgBox("¿Error en firma?", vbYesNo, "Pregunta") = vbYes Then
            If MsgBox("¿Falta firma?", vbYesNo, "Pregunta") = vbYes Then
                com = "FALTA FIRMA de administrador/a o supervisor/a"
                If Comentarios = "" Then Comentarios = com Else: Comentarios = Comentarios & "-" & com
            Else
                com = "FIRMA AUTORIZANTE: Verificar datos (Cargo, legibilidad...)"
                If Comentarios = "" Then Comentarios = com Else: Comentarios = Comentarios & "-" & com
            End If
        End If
    End If
    
    If ErrorScan = vbYes Then EstadoUser = ESTADO_PENDIENTE_REINGRESO
    
    If Hoja2.Cells(i, gCtx.rngEstadoDelPago.Range.Column) <> ESTADO_DIF_COSTO Then
        If Hoja2.Cells(i, gCtx.rngDifCostos.Range.Column) >= gCtx.montoToleranciaSB Then
            If tipoDoc = "FC-REM" Or tipoDoc = "NC-DEV" Then
                DC = MsgBox("Tiene diferencia de costos?", vbYesNo, "Pregunta")
            Else
                PNC = MsgBox("¿Está pendiente de Nota de Crédito?", vbYesNo, "Pregunta")
                If PNC = vbNo Then DC = MsgBox("¿Detallan diferencia de costos?", vbYesNo, "Pregunta")
            End If
        End If
    
        
        If DC = vbYes Then EstadoUser = ESTADO_DIF_COSTO
        
        If DC = vbNo Then
            
            If tipoDoc = "FC-REM" Or tipoDoc = "NC-REM" Then
                EstadoUser = ESTADO_PENDIENTE_REVISAR
                com = "Verificar artículos y cantidades"
                If Comentarios = "" Then Comentarios = com Else: Comentarios = Comentarios & "-" & com
            Else
                EstadoUser = ESTADO_PENDIENTE_REVISAR
                com = "No detallan motivo de la diferencia"
                If Comentarios = "" Then Comentarios = com Else: Comentarios = Comentarios & "-" & com
            End If
    
        End If
    Else
        EstadoUser = ESTADO_DIF_COSTO
    End If

    If ErrorScan = vbYes And DC = vbYes Then EstadoUser = "Varios motivos"
    If pendienteReingreso = vbYes Then
        EstadoUser = ESTADO_PENDIENTE_REINGRESO
        com = "Error en fecha"
        If Comentarios = "" Then Comentarios = com Else: Comentarios = Comentarios & "-" & com
    End If
    
    If PNC = vbYes Then EstadoUser = "Pendiente de Nota de Crédito - Mercaderia Faltante"
    
    If ErrorScan <> vbYes And pendienteReingreso <> vbYes And Hoja2.Cells(i, gCtx.rngDifCostos.Range.Column) < gCtx.montoToleranciaSB Then
        
        If tipoDoc = "FC-REM" Or tipoDoc = "NC-REM" Then
            EstadoUser = ESTADO_REMITO
        Else
            EstadoUser = ESTADO_MIGRAR_SAP
        End If
    End If
    
    
    Hoja2.Cells(i, gCtx.rngEstadoDelPago.Range.Column) = EstadoUser
    If Left(Comentarios, 1) = "-" Then Comentarios = Mid(Comentarios, 2)
    Hoja2.Cells(i, gCtx.rngComentarios_User.Range.Column) = Comentarios
    gCtx.rngComentarios_User.Range.Columns.AutoFit
    

    If EstadoUser <> ESTADO_MIGRAR_SAP And EstadoUser <> ESTADO_REMITO Then
        If Hoja2.Cells(i, gCtx.rngDifCostos.Range.Column) >= gCtx.montoToleranciaSB Then
            gCtx.endoso = False
            If InStr(1, UCase(Comentarios), "ENDOS") = 0 Then
                If MsgBox("¿Tiene Endoso o Nota de Crédito?", vbYesNo, "Pregunta") = vbYes Then
                    gCtx.endoso = True
                Else
                    If MsgBox("¿Cambiar estado del pago?", vbYesNo, "Pregunta") = vbYes Then CambiarEstado2 (i)
                End If
            End If
        Else
            If MsgBox("¿Cambiar estado del pago?", vbYesNo, "Pregunta") = vbYes Then CambiarEstado2 (i)
        End If
    End If
    
    Set fila = gCtx.tblDatos.ListRows(i - gCtx.tblDatos.HeaderRowRange.Row)
    Call VerificarDatos(fila)
    
    Call cerrarVentanas
  
End Sub
Public Sub cerrarVentanas()
    For Each ventana In CreateObject("Shell.Application").Windows
        If InStr(1, ventana.FullName, "iexplore.exe", vbTextCompare) > 0 Then
            If Not ventana Is gCtx.IE_NuevaVentana Then
                ventana.Quit
            End If
        End If
    Next ventana
End Sub

Private Function ShouldProcessRetailWebRow(ByVal rowIndex As Long) As Boolean

    If Hoja2.Rows(rowIndex).EntireRow.Hidden Then Exit Function
    If Hoja2.Cells(rowIndex, gCtx.rngRetailWeb_SB.Range.Column) = "" Then Exit Function
    If gCtx.textoBtn = ACTION_PAGAR_FACTURA And Hoja2.Cells(rowIndex, gCtx.rngCompensacion.Range.Column) = "" Then Exit Function
    If gCtx.textoBtn = ACTION_CAMBIAR_PAGAR And Hoja2.Cells(rowIndex, gCtx.rngCompensacion.Range.Column) = "" Then Exit Function

    ShouldProcessRetailWebRow = True

End Function















