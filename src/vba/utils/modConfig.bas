Attribute VB_Name = "modConfig"

Private Const ENV_PYTHON As String = "MIGRAR_SAP_PYTHON"
Private Const ENV_PYTHONW As String = "MIGRAR_SAP_PYTHONW"
Private Const ENV_SCRIPTS As String = "MIGRAR_SAP_SCRIPTS"
Private Const ENV_WORKBOOK_PASSWORD As String = "MIGRAR_PASSWORD"

Public Const ORIGEN_DATOS_SB As String = "RW"
Public Const ORIGEN_DATOS_CUBO As String = "CUBO"
Public Const RUTA_IMPORTAR As String = "Importar"
Public Const RUTA_REFRESH As String = "Refresh"
Public Const FLAG_SI As String = "SI"

Public Function GetWorkbookUnprotectPassword() As String

    GetWorkbookUnprotectPassword = Trim$(Environ$(ENV_WORKBOOK_PASSWORD))

End Function

Public Sub UnprotectHoja2Safe()

    Dim pwd As String

    pwd = GetWorkbookUnprotectPassword()

    If Len(pwd) > 0 Then
        Hoja2.Unprotect Password:=pwd
    Else
        On Error Resume Next
        Hoja2.Unprotect
        On Error GoTo 0
    End If

End Sub

Public Sub ProtectHoja2ForUi()

    Dim pwd As String

    pwd = GetWorkbookUnprotectPassword()

    If Len(pwd) > 0 Then
        Hoja2.Protect Password:=pwd, UserInterfaceOnly:=True, AllowFiltering:=True
    Else
        Hoja2.Protect UserInterfaceOnly:=True, AllowFiltering:=True
    End If

End Sub

Public Function GetPythonExePath() As String

    GetPythonExePath = ResolvePythonExePath(False)

End Function

Public Function GetPythonwExePath() As String

    GetPythonwExePath = ResolvePythonExePath(True)

End Function

Public Function ResolveScriptPath(ByVal scriptName As String) As String

    Dim root As String

    root = GetScriptsRoot()

    If root = "" Then
        ResolveScriptPath = scriptName
    Else
        ResolveScriptPath = root & "\" & scriptName
    End If

End Function

Private Function GetScriptsRoot() As String

    Dim root As String
    Dim candidate As String
    Dim workbookPath As String

    root = Environ$(ENV_SCRIPTS)
    If root <> "" Then
        GetScriptsRoot = NormalizeRoot(root)
        Exit Function
    End If

    On Error Resume Next
    workbookPath = ThisWorkbook.Path
    On Error GoTo 0

    If workbookPath <> "" Then
        If Right$(workbookPath, 1) <> "\" Then
            workbookPath = workbookPath & "\"
        End If

        candidate = workbookPath & "scripts"
        If Dir$(candidate, vbDirectory) <> "" Then
            GetScriptsRoot = NormalizeRoot(candidate)
            Exit Function
        End If

        candidate = workbookPath & "..\scripts"
        If Dir$(candidate, vbDirectory) <> "" Then
            GetScriptsRoot = NormalizeRoot(candidate)
            Exit Function
        End If
    End If

    GetScriptsRoot = ""

End Function

Private Function NormalizeRoot(ByVal root As String) As String

    If Right$(root, 1) = "\" Then
        NormalizeRoot = Left$(root, Len(root) - 1)
    Else
        NormalizeRoot = root
    End If

End Function

Private Function ResolvePythonExePath(ByVal useWindowless As Boolean) As String

    Dim exeName As String
    Dim envPath As String
    Dim basePath As String

    If useWindowless Then
        exeName = "pythonw.exe"
        envPath = Environ$(ENV_PYTHONW)
    Else
        exeName = "python.exe"
        envPath = Environ$(ENV_PYTHON)
    End If

    If envPath <> "" Then
        ResolvePythonExePath = envPath
        Exit Function
    End If

    basePath = Environ$("LOCALAPPDATA") & "\Programs\Python\Python313\" & exeName
    If Len(Dir$(basePath)) > 0 Then
        ResolvePythonExePath = basePath
        Exit Function
    End If

    basePath = Environ$("LOCALAPPDATA") & "\Programs\Python\Python310\" & exeName
    If Len(Dir$(basePath)) > 0 Then
        ResolvePythonExePath = basePath
        Exit Function
    End If

    If useWindowless Then
        ResolvePythonExePath = "pythonw"
    Else
        ResolvePythonExePath = "python"
    End If

End Function
