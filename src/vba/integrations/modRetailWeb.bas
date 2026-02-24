Attribute VB_Name = "modRetailWeb"

Sub AbrirRetailWeb()
    
    If Not gCtx.reporteSB Then AbrirRetailWebUser
        
End Sub
Sub AbrirRetailWebUser()

    On Error GoTo wrongpass
 
    tareas2 = 5
    tareas1 = tareas1 + tareas2
    avance2 = 1
    
    With ProgressBar
        .Show vbModeless
        .Lbl1.Caption = "Abrir RetailWeb. Progreso..." & " (" & Format(avance1 + avance2 / tareas1, "0%") & ")"
        .Lbl2.Caption = "Abriendo el explorador..." & " (" & Format(avance2 / tareas2, "0%") & ")"
        .pb1.Max = tareas1
        .pb2.Max = tareas2
        .pb1.Value = avance1 + avance2
        .pb2.Value = avance2
    End With
    
    For Each ventana In CreateObject("Shell.Application").Windows
        If ventana = IE_WINDOW_NAME Then
            If Left(ventana.LocationURL, Len(gCtx.dominio)) = gCtx.dominio Then
                If Not ventana.Visible Then
                
                    result = shell("taskkill /F /IM iexplore.exe", vbHide)
                    
                    Call AbrirRetailWebUser
                    Exit Sub
                Else: GoTo finProced
                End If
            ElseIf ventana.LocationURL = "" Then
       
                result = shell("taskkill /F /IM iexplore.exe", vbHide)

                Call AbrirRetailWebUser
                Exit Sub
            End If
        End If
    Next ventana
   
    Set ie = CreateObject("InternetExplorer.Application")
    
    avance2 = 2
    With ProgressBar
        .Lbl1.Caption = "Abrir RetailWeb. Progreso..." & " (" & Format(avance1 + avance2 / tareas1, "0%") & ")"
        .Lbl2.Caption = "Abriendo RetailWeb..." & " (" & Format(avance2 / tareas2, "0%") & ")"
        .pb1.Value = avance1 + avance2
        .pb2.Value = avance2
    End With
    
    ie.Visible = False
    ie.Navigate gCtx.linkSB
    
    If Not WaitForIEReady(ie, WAIT_LONG_SECONDS, "RetailWeb") Then GoTo wrongpass
    
    Set elemento = ie.Document.getElementById("dgf_login_form_fd-username")
    elemento.Value = Environ("USERNAME")
    
    If gCtx.PASS = "" Then gCtx.PASS = GetPass()
            
    If Not WaitForNonEmpty(gCtx.PASS, WAIT_LONG_SECONDS, "Password RetailWeb") Then GoTo wrongpass

    Set elemento = ie.Document.getElementById("dgf_login_form_fd-password_encrypted")
    elemento.Value = gCtx.PASS

    Set elemento = ie.Document.getElementById("form.login.title")
    elemento.Click
    
    avance2 = 3
    With ProgressBar
        .Lbl1.Caption = "Abrir RetailWeb. Progreso..." & " (" & Format(avance1 + avance2 / tareas1, "0%") & ")"
        .Lbl2.Caption = "Ingresando con su usuario y contraseña..." & " (" & Format(avance2 / tareas2, "0%") & ")"
        .pb1.Value = avance1 + avance2
        .pb2.Value = avance2
    End With
    
    If Not WaitForIEReady(ie, WAIT_LONG_SECONDS, "RetailWeb") Then GoTo wrongpass

    Set elementos = ie.Document.getElementsByClassName(SB_CLASS_PULL_LEFT)
    
    For Each elemento In elementos
        If InStr(elemento.innerText, SB_TEXT_CONTROL_INVENTARIOS) > 0 Then
            elemento.Click
            Exit For
        End If
    Next elemento
    
    For Each elemento In elementos
        If InStr(elemento.innerText, SB_TEXT_CONTROL_RECEPCIONES) > 0 Then
            elemento.Click
            Exit For
        End If
    Next elemento
    
    avance2 = 4
    With ProgressBar
        .Lbl1.Caption = "Abrir RetailWeb. Progreso..." & " (" & Format(avance1 + avance2 / tareas1, "0%") & ")"
        .Lbl2.Caption = "Abriendo el control de recepciones..." & " (" & Format(avance2 / tareas2, "0%") & ")"
        .pb1.Value = avance1 + avance2
        .pb2.Value = avance2
    End With

    If Not WaitForIEReady(ie, WAIT_LONG_SECONDS, "RetailWeb") Then GoTo wrongpass
    
    avance2 = 5
    With ProgressBar
        .Lbl1.Caption = "Abrir RetailWeb. Progreso..." & " (" & Format(avance1 + avance2 / tareas1, "0%") & ")"
        .Lbl2.Caption = "Finalizando..." & " (" & Format(avance2 / tareas2, "0%") & ")"
        .pb1.Value = avance1 + avance2
        .pb2.Value = avance2
    End With
    
    startTime = Timer
    Do
        DoEvents
        cuenta = 0
        For Each ventana In CreateObject("Shell.Application").Windows
            If ventana.Name = IE_WINDOW_NAME Then cuenta = cuenta + 1
        Next ventana
        Debug.Print ("Cuenta: " & cuenta)
        If HasTimedOut(startTime, WAIT_LONG_SECONDS) Then
            ReportTimeout "Abrir RetailWeb"
            GoTo wrongpass
        End If
    Loop Until cuenta >= 2

    ie.Quit
    DoEvents

    For i = 1 To 100
    
        For Each ventana In CreateObject("Shell.Application").Windows
            If ventana = IE_WINDOW_NAME Then
                If Left(ventana.LocationURL, Len(gCtx.dominio)) = gCtx.dominio Then
                
                    Set elementos = RetailWebTryGetElementsByClassFromDoc(ventana.Document, SB_CLASS_BUTTON_DEFAULT_SM)

                    If Not elementos Is Nothing Then
                        For Each elemento In elementos
                            If InStr(elemento.innerText, SB_TEXT_BUSCAR) > 0 Then

                                Set gCtx.IE_NuevaVentana = ventana

                                gCtx.IE_NuevaVentana.TheaterMode = False
                                gCtx.IE_NuevaVentana.Visible = True
                                Debug.Print ("elemento..." & i)
                                GoTo finProced

                            End If
                        Next elemento
                    End If

                    If ventana.TheaterMode = False Then Exit For
        
                End If
            End If
        Next ventana
        
        Debug.Print ("VUELTA..." & i)

    Next i

finProced:

    On Error GoTo wrongpass
    
    Application.Cursor = xlDefault
    
    ActiveSheet.Shapes("LuzSB").Fill.ForeColor.RGB = RGB(0, 255, 0) 'Verde
    
    GoTo CleanUp
   
wrongpass:
    
    MsgBox Err.Description
'    MENSAJE = MsgBox("¿Password incorrecta?", vbCritical)
'    Hoja3.Range("passwordSB") = ""
    gCtx.textoBtn = "Error"
    Application.Cursor = xlDefault

CleanUp:
    On Error Resume Next
    Unload ProgressBar
    Set ie = Nothing
    Set elemento = Nothing
    Set elementos = Nothing
    On Error GoTo 0
    
End Sub
Sub AbrirRetailWebCubo()

    Dim proc As Object

    pythonExe = GetPythonwExePath()
    scriptPath = ResolveScriptPath("reporte-pagoPendienteSi.py")
    
    cmd = """" & pythonExe & """ """ & scriptPath & """"
    
    Set proc = CreateObject("WScript.Shell").exec(cmd)
    
    If Not WaitForProcess(proc, WAIT_LONG_SECONDS, "RetailWeb Cubo") Then Exit Sub
    
    fullPath = proc.StdOut.ReadAll
    
    If Left(fullPath, 5) <> "ERROR" Then
    
        SafeDeleteSheet "sheetRetailWeb"
    
        Dim sheetRetailWeb As Worksheet
        
        Set sheetRetailWeb = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        
        sheetRetailWeb.Name = "sheetRetailWeb"
        
        Dim libroOrigen As Workbook
        
        Set libroOrigen = Workbooks.Open(Filename:=fullPath, ReadOnly:=True)
        
        libroOrigen.Sheets(1).Cells.Copy
        
        sheetRetailWeb.Cells.PasteSpecial Paste:=xlPasteValues
        
        Application.CutCopyMode = False
        
        libroOrigen.Close SaveChanges:=False
        
        Kill fullPath
        
        sheetRetailWeb.Columns("A:C").Delete
        
        sheetRetailWeb.Range("A2").CurrentRegion.Select
        
        Dim Tabla As ListObject
        
        Set Tabla = sheetRetailWeb.ListObjects.Add(xlSrcRange, Selection, , xlYes)
        
        Tabla.Name = "tblCuboSB"
        
        sheetRetailWeb.Range("A1") = Now
    
    Else
    
        MsgBox fullPath
    
    End If
    
End Sub
